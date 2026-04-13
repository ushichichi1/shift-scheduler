#!/usr/bin/env python3
"""
看護師勤務表自動作成 - Streamlit アプリ
Excel / Google Sheets 両対応
"""
import streamlit as st
import pandas as pd
import calendar
from datetime import date
from io import BytesIO

import jpholiday

from shift_scheduler import (
    Staff, build_and_solve,
    D, N, A, O, R, V, SHIFTS,
    TIER_A, TIER_AB, TIER_B, TIER_C, VALID_TIERS,
    _get_holidays_and_days_off, _write_one_sheet,
    SETTINGS_DEF, SETTINGS_KEYS,
    _parse_staff_list, _parse_requests, _parse_settings,
)

# ============================================================
# ページ設定
# ============================================================
st.set_page_config(
    page_title="勤務表作成ツール",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .stDataFrame { font-size: 12px; }
    div[data-testid="stMetric"] { background: #f8f9fa; padding: 10px; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 定数・ヘルパー
# ============================================================
WEEKDAY_MAP = {"月": 0, "火": 1, "水": 2, "木": 3, "金": 4, "土": 5, "日": 6}
WEEKDAY_REV = {v: k for k, v in WEEKDAY_MAP.items()}

# ============================================================
# 日本看護協会 夜勤ガイドラインチェック
# ============================================================
def check_nursing_guidelines(schedule, names, tiers, r_dedicated):
    """
    日本看護協会「夜勤・交代制勤務に関するガイドライン（2013年）」チェック
    Returns:
        violations: list of dict  — 違反（赤表示）
        warnings:   list of dict  — 注意（黄表示）
        ok_list:    list of str   — 適合スタッフ名
    """
    violations = []
    warnings = []
    ok_list = []

    for s in names:
        shifts = schedule[s]  # 0-indexed, len = num_days
        is_dedicated = r_dedicated.get(s, False)
        issues = []

        # ── 1. 月の夜勤回数（8回以内）────────────────────────────
        night_count = shifts.count(N)
        if is_dedicated:
            # 専従はガイドライン上「夜勤専従者」扱いのため対象外
            pass
        elif night_count > 8:
            issues.append({
                "rule": "夜勤回数", "level": "violation",
                "detail": f"{night_count}回（上限8回超過）"
            })
        elif night_count == 8:
            issues.append({
                "rule": "夜勤回数", "level": "warning",
                "detail": f"{night_count}回（上限8回ちょうど）"
            })

        # ── 2. 連続夜勤（2回以内）───────────────────────────────
        # N → A → (N → A ...) の繰り返しをカウント
        # 3回以上連続する夜勤サイクルを検出
        i = 0
        consec = 0
        max_consec = 0
        consec_start = -1
        while i < len(shifts):
            if shifts[i] == N:
                if consec == 0:
                    consec_start = i + 1  # 1-indexed day
                consec += 1
                max_consec = max(max_consec, consec)
                i += 2  # N の翌日は A のためスキップ
            else:
                consec = 0
                i += 1
        if max_consec > 2:
            issues.append({
                "rule": "連続夜勤", "level": "violation",
                "detail": f"最大{max_consec}連続（上限2回超過）"
            })

        # ── 3. 勤務間インターバル（11時間以上）──────────────────
        # ハード制約で A→O/N が保証されているが念のため A→D を検出
        interval_violations = []
        for d in range(len(shifts) - 1):
            if shifts[d] == A and shifts[d + 1] == D:
                interval_violations.append(f"{d + 2}日")
        if interval_violations:
            issues.append({
                "rule": "インターバル", "level": "violation",
                "detail": f"明け翌日に日勤: {', '.join(interval_violations)}（11時間未満の疑い）"
            })

        # ── 4. 夜勤後の休息（明け+休みの連続）──────────────────
        # N → A の翌日が必ず O/V/N であることを確認（A→D は上記で検出済み）
        for d in range(len(shifts) - 1):
            if shifts[d] == N and d + 1 < len(shifts) and shifts[d + 1] != A:
                issues.append({
                    "rule": "夜勤後の明け", "level": "violation",
                    "detail": f"{d + 1}日: 夜勤翌日が明けでない（{shifts[d+1]}）"
                })

        # ── 集計 ─────────────────────────────────────────────
        v_issues = [x for x in issues if x["level"] == "violation"]
        w_issues = [x for x in issues if x["level"] == "warning"]
        for item in v_issues:
            violations.append({"名前": s, "Tier": tiers[s], **item})
        for item in w_issues:
            warnings.append({"名前": s, "Tier": tiers[s], **item})
        if not issues:
            ok_list.append(s)

    return violations, warnings, ok_list

SHIFT_DISPLAY = {D: "日", N: "夜", O: "休", R: "研",
                 "夜不": "夜不", "休暇": "休暇", "明休": "明休"}
SHIFT_REVERSE = {"日": D, "夜": N, "休": O, "研": R,
                 "夜不": "夜不", "休暇": "休暇", "明休": "明休"}


def _staff_to_df(staff_list):
    """Staffリスト → DataFrame"""
    rows = []
    for s in staff_list:
        wd_str = ""
        if s.work_days:
            wd_str = "".join(WEEKDAY_REV[d] for d in sorted(s.work_days))
        rows.append({
            "名前": s.name, "Tier": s.tier,
            "夜勤専従": s.dedicated,
            "週勤務": s.weekly_days,
            "前月末": s.prev_month or "",
            "夜勤Min": s.night_min,
            "夜勤Max": s.night_max,
            "連勤Max": s.consec_max,
            "勤務曜日": wd_str,
            "祝日不可": s.no_holiday,
        })
    return pd.DataFrame(rows) if rows else _default_staff()


def _reqs_to_df(reqs_dict, staff_list, num_days):
    """希望dict → DataFrame"""
    req_rows = []
    for name, rq in reqs_dict.items():
        row = {"名前": name}
        for d in range(1, num_days + 1):
            v = rq.get(d, "")
            row[str(d)] = SHIFT_DISPLAY.get(v, v) if v else ""
        req_rows.append(row)
    existing = {r["名前"] for r in req_rows}
    for s in staff_list:
        if s.name not in existing:
            row = {"名前": s.name}
            for d in range(1, num_days + 1):
                row[str(d)] = ""
            req_rows.append(row)
    return pd.DataFrame(req_rows)


_SETTINGS_WIDGET_MAP = {
    "year": "inp_year", "month": "inp_month",
    "min_day_staff": "inp_min_day", "night_staff_count": "inp_night_count",
    "max_night_regular": "inp_max_n_reg", "pref_night_regular": "inp_pref_n_reg",
    "max_night_dedicated": "inp_max_n_ded", "pref_night_dedicated": "inp_pref_n_ded",
    "max_consecutive": "inp_max_consec", "pref_consecutive": "inp_pref_consec",
    "solver_time_limit": "inp_time_limit",
}

def _apply_settings(gs_settings):
    """読み込んだ設定を _pending_* キーに保存（ウィジェット描画前に反映させるため）"""
    for src, dst in _SETTINGS_WIDGET_MAP.items():
        v = gs_settings.get(src)
        if v is not None:
            st.session_state[f"_pending_{dst}"] = int(v)
    po = gs_settings.get("public_off_override")
    if po is not None and po != "":
        st.session_state["_pending_inp_po_mode"] = "手動指定"
        st.session_state["_pending_inp_po_val"] = int(po)
    else:
        st.session_state["_pending_inp_po_mode"] = "自動（土日祝）"


def _default_staff():
    return pd.DataFrame({
        "名前": ["スタッフA", "スタッフB", "スタッフC", "スタッフD", "スタッフE"],
        "Tier": ["A", "A", "AB", "C", "C"],
        "夜勤専従": [False, False, False, False, False],
        "週勤務": [None, None, None, None, None],
        "前月末": ["", "", "", "", ""],
        "夜勤Min": [None, None, None, None, None],
        "夜勤Max": [None, None, None, None, None],
        "連勤Max": [None, None, None, None, None],
        "勤務曜日": ["", "", "", "", ""],
        "祝日不可": [False, False, False, False, False],
    })


# ============================================================
# Excelテンプレート生成
# ============================================================
def _generate_template_excel(year, month):
    """入力用Excelテンプレートを生成"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = Workbook()
    num_days = calendar.monthrange(year, month)[1]
    thin = Side(style="thin")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="4472C4")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    sub_fill = PatternFill("solid", fgColor="D9E2F3")

    # --- Sheet 1: 設定 ---
    ws_s = wb.active
    ws_s.title = "設定"
    ws_s.column_dimensions["A"].width = 22
    ws_s.column_dimensions["B"].width = 14
    ws_s.column_dimensions["C"].width = 35
    # ヘッダー
    for c, txt in enumerate(["項目", "値", "説明"], 1):
        cell = ws_s.cell(row=3, column=c, value=txt)
        cell.fill = hdr_fill; cell.font = hdr_font; cell.border = bdr
        cell.alignment = Alignment(horizontal="center")
    ws_s.cell(row=1, column=1, value="勤務表設定").font = Font(bold=True, size=14)
    # データ
    for i, (label, default, desc) in enumerate(SETTINGS_DEF):
        r = 4 + i
        ws_s.cell(row=r, column=1, value=label).border = bdr
        ws_s.cell(row=r, column=2, value=default).border = bdr
        ws_s.cell(row=r, column=3, value=desc).border = bdr
        ws_s.cell(row=r, column=3).font = Font(color="888888", size=9)

    # --- Sheet 2: スタッフ一覧 ---
    ws_st = wb.create_sheet("スタッフ一覧")
    headers = ["名前", "Tier", "夜勤専従", "週勤務", "前月末",
               "夜勤Min", "夜勤Max", "連勤Max", "勤務曜日", "祝日不可"]
    for c, txt in enumerate(headers, 1):
        cell = ws_st.cell(row=1, column=c, value=txt)
        cell.fill = hdr_fill; cell.font = hdr_font; cell.border = bdr
        cell.alignment = Alignment(horizontal="center")
    ws_st.column_dimensions["A"].width = 14
    for col_letter in ["B", "C", "D", "E", "F", "G", "H", "I", "J"]:
        ws_st.column_dimensions[col_letter].width = 10
    # サンプル行
    samples = [
        ["山田太郎", "A", "", "", "", "", "", "", "", ""],
        ["佐藤花子", "AB", "", "", "", "", "", "", "", ""],
        ["鈴木一郎", "C", "", "3", "", "", "", "", "月水金", ""],
    ]
    for i, row_data in enumerate(samples):
        for c, val in enumerate(row_data, 1):
            cell = ws_st.cell(row=2 + i, column=c, value=val)
            cell.border = bdr
    # 空行追加（20行分）
    for i in range(len(samples), 20):
        for c in range(1, 11):
            ws_st.cell(row=2 + i, column=c).border = bdr

    # --- Sheet 3: 勤務希望 ---
    ws_r = wb.create_sheet("勤務希望")
    ws_r.cell(row=1, column=1, value="勤務希望入力").font = Font(bold=True, size=14)
    ws_r.cell(row=2, column=1, value="日/夜/休/研/夜不/休暇/明休 を入力").font = Font(color="888888", size=9)

    weekdays_jp = ["月", "火", "水", "木", "金", "土", "日"]
    first_wd = date(year, month, 1).weekday()
    holidays = {d.day for d, _ in jpholiday.month_holidays(year, month)}

    # ヘッダー行 (row 4)
    ws_r.column_dimensions["A"].width = 14
    cell = ws_r.cell(row=4, column=1, value="名前")
    cell.fill = hdr_fill; cell.font = hdr_font; cell.border = bdr
    for d in range(1, num_days + 1):
        wd_name = weekdays_jp[(first_wd + d - 1) % 7]
        cell = ws_r.cell(row=3, column=1 + d, value=wd_name)
        cell.alignment = Alignment(horizontal="center")
        cell.font = Font(size=9, color="888888")
        cell2 = ws_r.cell(row=4, column=1 + d, value=d)
        cell2.alignment = Alignment(horizontal="center")
        cell2.border = bdr
        # 土日祝の色分け
        wd_idx = (first_wd + d - 1) % 7
        if d in holidays:
            cell2.fill = PatternFill("solid", fgColor="F4CCCC")
            cell2.font = Font(bold=True, color="CC0000")
        elif wd_idx >= 5:
            cell2.fill = PatternFill("solid", fgColor="D9E2F3")
            cell2.font = Font(bold=True, color="4472C4")
        else:
            cell2.fill = hdr_fill
            cell2.font = hdr_font
    # 空行（20行分）
    for i in range(20):
        for c in range(1, num_days + 2):
            ws_r.cell(row=5 + i, column=c).border = bdr

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============================================================
# Excelアップロード解析
# ============================================================
def _parse_uploaded_excel(uploaded_file, year, month):
    """アップロードされたExcelを解析してスタッフ・希望・設定を返す"""
    from openpyxl import load_workbook
    wb = load_workbook(uploaded_file, data_only=True)
    num_days = calendar.monthrange(year, month)[1]

    # 設定シート
    gs_settings = {}
    if "設定" in wb.sheetnames:
        ws = wb["設定"]
        rows = []
        for r in range(4, 4 + len(SETTINGS_KEYS)):
            row_vals = [ws.cell(row=r, column=c).value or "" for c in range(1, 4)]
            rows.append(row_vals)
        gs_settings = _parse_settings(rows)

    # スタッフ一覧
    staff_list = []
    if "スタッフ一覧" in wb.sheetnames:
        ws = wb["スタッフ一覧"]
        staff_rows = []
        for r in range(2, ws.max_row + 1):
            row_vals = [ws.cell(row=r, column=c).value or "" for c in range(1, 11)]
            if str(row_vals[0]).strip():
                staff_rows.append(row_vals)
        staff_list = _parse_staff_list(staff_rows)

    # 勤務希望
    reqs = {}
    if "勤務希望" in wb.sheetnames:
        ws = wb["勤務希望"]
        staff_names = [s.name for s in staff_list]
        req_rows = []
        for r in range(5, ws.max_row + 1):
            row_vals = [ws.cell(row=r, column=c).value or "" for c in range(1, num_days + 2)]
            if str(row_vals[0]).strip():
                req_rows.append(row_vals)
        reqs = _parse_requests(req_rows, staff_names, num_days)

    return staff_list, reqs, gs_settings


# ============================================================
# セッション初期化
# ============================================================
if "staff_df" not in st.session_state:
    st.session_state.staff_df = _default_staff()
if "requests_df" not in st.session_state:
    st.session_state.requests_df = None
if "results" not in st.session_state:
    st.session_state.results = None
if "data_loaded" not in st.session_state:
    st.session_state.data_loaded = False

# ============================================================
# pending設定をウィジェットキーに反映（ウィジェット描画前に実行）
# ============================================================
for _src, _dst in _SETTINGS_WIDGET_MAP.items():
    _pk = f"_pending_{_dst}"
    if _pk in st.session_state:
        st.session_state[_dst] = st.session_state.pop(_pk)
for _pk, _wk in [("_pending_inp_po_mode", "inp_po_mode"), ("_pending_inp_po_val", "inp_po_val")]:
    if _pk in st.session_state:
        st.session_state[_wk] = st.session_state.pop(_pk)

# ============================================================
# サイドバー: 設定
# ============================================================
st.sidebar.title("勤務表設定")

col_y, col_m = st.sidebar.columns(2)
year = col_y.number_input("対象年", 2024, 2030, 2026, key="inp_year")
month = col_m.number_input("対象月", 1, 12, 5, key="inp_month")

num_days = calendar.monthrange(year, month)[1]
holidays_auto, weekends_auto, auto_public_off = _get_holidays_and_days_off(year, month)

st.sidebar.markdown("---")
st.sidebar.subheader("勤務条件")

public_off_mode = st.sidebar.radio("公休日数", ["自動（土日祝）", "手動指定"],
                                    horizontal=True, key="inp_po_mode")
if public_off_mode == "自動（土日祝）":
    public_off = auto_public_off
    st.sidebar.info(f"公休 {public_off}日（土日{len(weekends_auto)} + 祝{len(holidays_auto)}）")
    po_override = None
else:
    public_off = st.sidebar.number_input("公休日数", 0, num_days, auto_public_off, key="inp_po_val")
    po_override = public_off

min_day = st.sidebar.number_input("日勤最低人数", 1, 20, 5, key="inp_min_day")
night_count = st.sidebar.number_input("夜勤人数/日", 1, 5, 2, key="inp_night_count")

st.sidebar.markdown("---")
st.sidebar.subheader("夜勤設定")
col1, col2 = st.sidebar.columns(2)
max_n_reg = col1.number_input("通常 上限", 1, 15, 5, key="inp_max_n_reg")
pref_n_reg = col2.number_input("通常 推奨", 1, 15, 4, key="inp_pref_n_reg")
col3, col4 = st.sidebar.columns(2)
max_n_ded = col3.number_input("専従 上限", 1, 20, 10, key="inp_max_n_ded")
pref_n_ded = col4.number_input("専従 推奨", 1, 20, 9, key="inp_pref_n_ded")

st.sidebar.markdown("---")
st.sidebar.subheader("連勤設定")
col5, col6 = st.sidebar.columns(2)
max_consec = col5.number_input("最大連勤", 1, 10, 5, key="inp_max_consec")
pref_consec = col6.number_input("推奨連勤", 1, 10, 4, key="inp_pref_consec")

time_limit = st.sidebar.number_input("計算時間上限（秒）", 10, 600, 120, key="inp_time_limit")
num_patterns = st.sidebar.number_input("生成パターン数", 1, 5, 3, key="inp_num_patterns")

settings = {
    "year": year, "month": month,
    "public_off_override": po_override,
    "min_day_staff": min_day, "night_staff_count": night_count,
    "max_night_regular": max_n_reg, "pref_night_regular": pref_n_reg,
    "max_night_dedicated": max_n_ded, "pref_night_dedicated": pref_n_ded,
    "max_consecutive": max_consec, "pref_consecutive": pref_consec,
    "solver_time_limit": time_limit, "holidays": "",
}

# ============================================================
# メインエリア
# ============================================================
st.title("🏥 勤務表自動作成ツール")

tab0, tab1, tab2, tab3, tab4 = st.tabs(
    ["📂 データ入力", "📋 スタッフ一覧", "📝 勤務希望", "⚡ 生成", "📊 結果"])

# ============================================================
# Tab 0: データ入力（テンプレ出力 / アップロード / スプシ）
# ============================================================
with tab0:
    st.subheader("データの入力方法を選んでください")

    col_left, col_right = st.columns(2)

    # --- 左カラム: テンプレートダウンロード ---
    with col_left:
        st.markdown("### 📥 テンプレートをダウンロード")
        st.markdown("""
        1. 下のボタンでExcelテンプレートをダウンロード
        2. **設定・スタッフ一覧・勤務希望** を記入
        3. 右の「Excelアップロード」で読み込み
        """)
        template_bytes = _generate_template_excel(year, month)
        st.download_button(
            label=f"📄 テンプレートDL（{year}年{month}月）",
            data=template_bytes,
            file_name=f"勤務表テンプレート_{year}_{month:02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # --- 右カラム: データ読み込み ---
    with col_right:
        st.markdown("### 📤 データを読み込む")
        import_method = st.radio("読み込み方法", ["Excel ファイル", "Google スプレッドシート"],
                                  horizontal=True, key="import_method")

        # ローカルデフォルトファイルが存在する場合に表示
        import os as _os
        _default_excel = _os.path.join(_os.path.dirname(__file__), "勤務表_入力.xlsx")
        if _os.path.exists(_default_excel):
            if st.button("📁 ローカルファイルを読み込む（勤務表_入力.xlsx）",
                         use_container_width=True, key="btn_load_local"):
                try:
                    with st.spinner("読み込み中..."):
                        from shift_scheduler import load_input as _load_local
                        staff_list, reqs, gs_settings = _load_local()
                    if gs_settings:
                        _apply_settings(gs_settings)
                    st.session_state.staff_df = _staff_to_df(staff_list)
                    st.session_state.requests_df = _reqs_to_df(reqs, staff_list, num_days)
                    st.session_state.data_loaded = True
                    st.success(f"✅ {len(staff_list)}人のスタッフを読み込みました")
                    st.rerun()
                except Exception as e:
                    st.error(f"読み込みエラー: {e}")
            st.divider()

        if import_method == "Excel ファイル":
            uploaded = st.file_uploader("Excelファイルを選択", type=["xlsx", "xlsm"],
                                         key="excel_upload")
            if uploaded is not None:
                if st.button("📤 Excelから読み込み", type="primary", use_container_width=True,
                             key="btn_load_excel"):
                    try:
                        with st.spinner("読み込み中..."):
                            staff_list, reqs, gs_settings = _parse_uploaded_excel(
                                uploaded, year, month)
                        if gs_settings:
                            _apply_settings(gs_settings)
                        st.session_state.staff_df = _staff_to_df(staff_list)
                        st.session_state.requests_df = _reqs_to_df(reqs, staff_list, num_days)
                        st.session_state.data_loaded = True
                        st.success(f"✅ {len(staff_list)}人のスタッフと希望を読み込みました")
                        st.rerun()
                    except Exception as e:
                        st.error(f"読み込みエラー: {e}")

        else:  # Google スプレッドシート
            gsheet_id = st.text_input(
                "スプレッドシートID or URL",
                value="1mezi6NHOQZj0VR_TzmRb9UdkFGrFiWg7qJmvOrJopgA",
                key="gsheet_input")
            if st.button("📤 スプレッドシートから読み込み", type="primary",
                         use_container_width=True, key="btn_load_gsheet"):
                try:
                    with st.spinner("読み込み中..."):
                        from shift_scheduler import load_gsheet as _load_gs
                        staff_list, reqs_dict, gs_settings, _ = _load_gs(gsheet_id)
                    if gs_settings:
                        _apply_settings(gs_settings)
                    st.session_state.staff_df = _staff_to_df(staff_list)
                    st.session_state.requests_df = _reqs_to_df(reqs_dict, staff_list, num_days)
                    st.session_state.data_loaded = True
                    st.success(f"✅ {len(staff_list)}人のスタッフと希望を読み込みました")
                    st.rerun()
                except Exception as e:
                    st.error(f"読み込みエラー: {e}")

    # 読み込み状況
    st.markdown("---")
    if st.session_state.data_loaded:
        n_staff = len(st.session_state.staff_df.dropna(subset=["名前"]))
        n_staff = len([n for n in st.session_state.staff_df["名前"].dropna() if str(n).strip()])
        st.success(f"✅ データ読み込み済み（{n_staff}人）→ 「スタッフ一覧」「勤務希望」タブで確認・編集できます")
    else:
        st.info("💡 テンプレートをDLして記入 → アップロード、またはスプレッドシートから読み込んでください")

# ============================================================
# Tab 1: スタッフ一覧
# ============================================================
with tab1:
    st.subheader("スタッフ一覧")
    st.caption("行を追加・編集できます。Tierは A/AB/B/C から選択。")

    edited_staff = st.data_editor(
        st.session_state.staff_df,
        num_rows="dynamic",
        column_config={
            "名前": st.column_config.TextColumn("名前", width="medium"),
            "Tier": st.column_config.SelectboxColumn("Tier", options=["A", "AB", "B", "C"], width="small"),
            "夜勤専従": st.column_config.CheckboxColumn("夜勤専従", width="small"),
            "週勤務": st.column_config.NumberColumn("週勤務", min_value=1, max_value=7, step=1, width="small"),
            "前月末": st.column_config.SelectboxColumn("前月末", options=["", "夜", "明"], width="small"),
            "夜勤Min": st.column_config.NumberColumn("夜勤Min", min_value=0, max_value=15, step=1, width="small"),
            "夜勤Max": st.column_config.NumberColumn("夜勤Max", min_value=0, max_value=15, step=1, width="small"),
            "連勤Max": st.column_config.NumberColumn("連勤Max", min_value=1, max_value=10, step=1, width="small"),
            "勤務曜日": st.column_config.TextColumn("勤務曜日", width="small", help="例: 月火木"),
            "祝日不可": st.column_config.CheckboxColumn("祝日不可", width="small"),
        },
        use_container_width=True,
        key="staff_editor",
    )
    st.session_state.staff_df = edited_staff

    if holidays_auto:
        hol_names = {d.day: name for d, name in jpholiday.month_holidays(year, month)}
        hol_str = "、".join(f"{d}日({hol_names.get(d, '祝')})" for d in sorted(holidays_auto))
        st.info(f"🗓 {year}年{month}月の祝日: {hol_str}")

# ============================================================
# Tab 2: 勤務希望
# ============================================================
with tab2:
    st.subheader(f"勤務希望 — {year}年{month}月（{num_days}日間）")
    st.caption("日/夜/休/研/夜不/休暇/明休 を入力（空欄=希望なし）")

    staff_names = [n for n in edited_staff["名前"].dropna().tolist() if str(n).strip()]

    if st.session_state.requests_df is None or set(st.session_state.requests_df["名前"].tolist()) != set(staff_names):
        if st.session_state.requests_df is not None:
            old = st.session_state.requests_df.set_index("名前")
            rows = []
            for name in staff_names:
                if name in old.index:
                    rows.append(old.loc[name].to_dict() | {"名前": name})
                else:
                    row = {"名前": name}
                    for d in range(1, num_days + 1):
                        row[str(d)] = ""
                    rows.append(row)
            st.session_state.requests_df = pd.DataFrame(rows)
        else:
            rows = []
            for name in staff_names:
                row = {"名前": name}
                for d in range(1, num_days + 1):
                    row[str(d)] = ""
                rows.append(row)
            st.session_state.requests_df = pd.DataFrame(rows)

    wdj = ["月", "火", "水", "木", "金", "土", "日"]
    fwd = date(year, month, 1).weekday()
    header_map = {}
    for d in range(1, num_days + 1):
        wd = wdj[(fwd + d - 1) % 7]
        is_hol = d in holidays_auto
        is_we = d in weekends_auto
        suffix = "祝" if is_hol else ("★" if is_we else "")
        header_map[str(d)] = f"{d}({wd}{suffix})"

    col_config = {"名前": st.column_config.TextColumn("名前", disabled=True, width="medium")}
    shift_options = ["", "日", "夜", "休", "研", "夜不", "休暇", "明休"]
    for d in range(1, num_days + 1):
        col_config[str(d)] = st.column_config.SelectboxColumn(
            header_map[str(d)], options=shift_options, width="small")

    edited_reqs = st.data_editor(
        st.session_state.requests_df,
        column_config=col_config,
        use_container_width=True,
        key="reqs_editor",
        hide_index=True,
    )
    st.session_state.requests_df = edited_reqs

# ============================================================
# Tab 3: 生成
# ============================================================
with tab3:
    st.subheader("勤務表生成")

    col_a, col_b, col_c = st.columns(3)
    valid_staff = edited_staff.dropna(subset=["名前"])
    valid_staff = valid_staff[valid_staff["名前"].str.strip() != ""]
    ft_count = len(valid_staff[valid_staff["週勤務"].isna()])
    pt_count = len(valid_staff[valid_staff["週勤務"].notna()])
    ded_count = len(valid_staff[valid_staff["夜勤専従"] == True])

    col_a.metric("スタッフ数", f"{len(valid_staff)}人")
    col_b.metric("フルタイム / パート", f"{ft_count} / {pt_count}")
    col_c.metric("夜勤専従", f"{ded_count}人")

    st.markdown("---")

    if st.button("🚀 勤務表を生成", type="primary", use_container_width=True):
        staff_list = []
        for _, row in valid_staff.iterrows():
            name = str(row["名前"]).strip()
            tier = str(row["Tier"]).strip() if pd.notna(row["Tier"]) else "C"
            if tier not in VALID_TIERS:
                st.warning(f"⚠ {name}: Tier '{tier}' 不正 → スキップ")
                continue
            ded = bool(row.get("夜勤専従", False))
            weekly = int(row["週勤務"]) if pd.notna(row.get("週勤務")) else None
            prev = str(row.get("前月末", "")).strip()
            if prev not in ("夜", "明", ""):
                prev = ""
            n_min = int(row["夜勤Min"]) if pd.notna(row.get("夜勤Min")) else None
            n_max = int(row["夜勤Max"]) if pd.notna(row.get("夜勤Max")) else None
            c_max = int(row["連勤Max"]) if pd.notna(row.get("連勤Max")) else None
            wd_str = str(row.get("勤務曜日", "")).strip()
            work_days = None
            if wd_str:
                work_days = set()
                for ch in wd_str:
                    if ch in WEEKDAY_MAP:
                        work_days.add(WEEKDAY_MAP[ch])
                if not work_days:
                    work_days = None
            no_hol = bool(row.get("祝日不可", False))
            staff_list.append(Staff(name, tier, ded, weekly, prev,
                                     n_min, n_max, c_max, work_days, no_hol))

        if not staff_list:
            st.error("スタッフが0人です")
        else:
            reqs_dict = {}
            for _, row in edited_reqs.iterrows():
                name = str(row["名前"]).strip()
                if not name:
                    continue
                rq = {}
                for d in range(1, num_days + 1):
                    val = str(row.get(str(d), "")).strip()
                    if val and val in SHIFT_REVERSE:
                        rq[d] = SHIFT_REVERSE[val]
                    elif val and val in (D, N, O, R):
                        rq[d] = val
                if rq:
                    reqs_dict[name] = rq

            with st.spinner(f"最適化計算中... （最大{time_limit}秒 × {num_patterns}パターン）"):
                import io, contextlib
                console = io.StringIO()
                with contextlib.redirect_stdout(console):
                    results = build_and_solve(staff_list, reqs_dict, settings,
                                              num_patterns=num_patterns)

            console_output = console.getvalue()

            if results:
                st.session_state.results = results
                st.session_state.console_output = console_output
                st.success(f"✅ {len(results)}パターン生成完了！「結果」タブで確認できます。")

                buf = BytesIO()
                from openpyxl import Workbook
                wb = Workbook()
                wb.remove(wb.active)
                for res in results:
                    pat = res.get("pattern_num", 1)
                    title = f"パターン{pat}" if len(results) > 1 else "勤務表"
                    _write_one_sheet(wb, res, title)
                wb.save(buf)
                buf.seek(0)
                st.session_state.excel_bytes = buf.getvalue()
            else:
                st.error("❌ 解なし（Infeasible）。設定・希望を見直してください。")
                with st.expander("ソルバーログ"):
                    st.code(console_output)

# ============================================================
# Tab 4: 結果表示
# ============================================================
with tab4:
    if st.session_state.results is None:
        st.info("「生成」タブで勤務表を作成してください。")
    else:
        results = st.session_state.results

        if len(results) > 1:
            pat_idx = st.selectbox("パターン選択",
                                    range(len(results)),
                                    format_func=lambda i: f"パターン {i+1}")
        else:
            pat_idx = 0

        result = results[pat_idx]
        schedule = result["schedule"]
        names = result["names"]
        tiers = result["tiers"]
        r_num_days = result["num_days"]
        r_year = result["year"]
        r_month = result["month"]
        r_holidays = result.get("holidays", set())
        r_weekends = result.get("weekends", set())
        r_public_off = result.get("public_off", 13)
        r_weekly = result.get("weekly", {})
        r_dedicated = result.get("dedicated", {})
        missed = result.get("missed_requests", {})

        st.subheader(f"📊 {r_year}年{r_month}月 勤務表 — パターン {pat_idx + 1}")

        night_counts = {}
        for s in names:
            nc = schedule[s].count(N)
            if nc > 0:
                night_counts[s] = nc

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("公休日数", f"{r_public_off}日")
        if night_counts:
            nc_vals = [v for s, v in night_counts.items()
                       if r_weekly.get(s) is None and not r_dedicated.get(s, False)]
            if nc_vals:
                col2.metric("夜勤均等", f"{min(nc_vals)}〜{max(nc_vals)}回")
        col3.metric("未達希望", f"{sum(len(v) for v in missed.values())}件" if missed else "0件 ✓")
        col4.metric("パターン", f"{pat_idx+1} / {len(results)}")

        wdj = ["月", "火", "水", "木", "金", "土", "日"]
        fwd = date(r_year, r_month, 1).weekday()

        shift_colors = {
            D: "#FFFFFF", N: "#4472C4", A: "#FFF2CC",
            O: "#E2EFDA", R: "#E8D5F5", V: "#F4CCCC"
        }
        shift_text_colors = {N: "#FFFFFF"}

        day_cols = [f"{d+1}" for d in range(r_num_days)]
        table_data = []
        for s in names:
            row = {"名前": s, "Tier": tiers[s]}
            for d in range(r_num_days):
                row[day_cols[d]] = schedule[s][d]
            counts = {t: schedule[s].count(t) for t in SHIFTS}
            row["日"] = counts[D]
            row["夜"] = counts[N]
            row["明"] = counts[A]
            row["休"] = counts[O]
            row["研"] = counts[R]
            row["暇"] = counts[V]
            row["公休"] = counts[O] + counts[V]
            table_data.append(row)

        df = pd.DataFrame(table_data)

        def color_shifts(val):
            if val in shift_colors:
                bg = shift_colors[val]
                fg = shift_text_colors.get(val, "#000000")
                return f"background-color: {bg}; color: {fg}; text-align: center; font-weight: {'bold' if val == N else 'normal'}"
            return "text-align: center"

        styled = df.style.map(color_shifts, subset=day_cols)
        st.dataframe(styled, use_container_width=True, height=600, hide_index=True)

        with st.expander("🌙 日別夜勤ペア", expanded=False):
            pair_data = []
            for d in range(r_num_days):
                wd = wdj[(fwd + d) % 7]
                hol = "祝" if (d+1) in r_holidays else ("★" if (d+1) in r_weekends else "")
                nn = [s for s in names if schedule[s][d] == N]
                pair_data.append({
                    "日": f"{d+1}日({wd}{hol})",
                    "夜勤メンバー": " + ".join(f"{s}({tiers[s]})" for s in nn),
                })
            st.dataframe(pd.DataFrame(pair_data), use_container_width=True, hide_index=True)

        if missed:
            with st.expander("⚠ 未達希望", expanded=True):
                for s, ds in missed.items():
                    st.warning(f"{s}: {', '.join(f'{d}日' for d in ds)}")
        else:
            st.success("✓ 全希望達成")

        # ── 日本看護協会ガイドラインチェック ─────────────────────
        violations, warnings_gl, ok_list = check_nursing_guidelines(
            schedule, names, tiers, r_dedicated
        )
        gl_label = "✅ ガイドライン適合" if not violations else f"🚨 ガイドライン違反 {len(violations)}件"
        with st.expander(f"📋 日本看護協会 夜勤ガイドラインチェック — {gl_label}", expanded=bool(violations)):
            st.caption("根拠: 日本看護協会「夜勤・交代制勤務に関するガイドライン」(2013年)")
            gl_col1, gl_col2, gl_col3 = st.columns(3)
            gl_col1.metric("🚨 違反", f"{len(violations)}件",
                           delta=None if not violations else "要対応",
                           delta_color="inverse")
            gl_col2.metric("⚠ 注意", f"{len(warnings_gl)}件")
            gl_col3.metric("✅ 適合", f"{len(ok_list)}名")

            st.markdown("##### チェック項目")
            rules_info = [
                ("夜勤回数", "月8回以内（夜勤専従者は対象外）"),
                ("連続夜勤", "連続夜勤2サイクル以内"),
                ("インターバル", "勤務間隔11時間以上（明け翌日に日勤がないか）"),
                ("夜勤後の明け", "夜勤翌日は必ず明けシフト"),
            ]
            for rule, desc in rules_info:
                v_cnt = sum(1 for x in violations if x["rule"] == rule)
                w_cnt = sum(1 for x in warnings_gl if x["rule"] == rule)
                icon = "🚨" if v_cnt else ("⚠" if w_cnt else "✅")
                st.markdown(f"- {icon} **{rule}**: {desc}"
                            + (f" → **{v_cnt}件違反**" if v_cnt else "")
                            + (f" → {w_cnt}件注意" if w_cnt else ""))

            if violations:
                st.markdown("##### 🚨 違反一覧")
                st.dataframe(
                    pd.DataFrame(violations)[["名前", "Tier", "rule", "detail"]]
                    .rename(columns={"rule": "項目", "detail": "内容"}),
                    use_container_width=True, hide_index=True
                )

            if warnings_gl:
                st.markdown("##### ⚠ 注意一覧")
                st.dataframe(
                    pd.DataFrame(warnings_gl)[["名前", "Tier", "rule", "detail"]]
                    .rename(columns={"rule": "項目", "detail": "内容"}),
                    use_container_width=True, hide_index=True
                )

            if ok_list:
                st.markdown(f"##### ✅ 全項目適合: {', '.join(ok_list)}")

        with st.expander("📈 日別集計", expanded=False):
            summary_data = {"日付": [f"{d+1}" for d in range(r_num_days)]}
            for label, shift in [("日勤", D), ("夜勤", N), ("明け", A), ("休み", O), ("研修", R)]:
                summary_data[label] = [sum(1 for s in names if schedule[s][d] == shift)
                                        for d in range(r_num_days)]
            st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)

        with st.expander("🔧 ソルバーログ"):
            st.code(st.session_state.get("console_output", ""))

        st.markdown("---")
        if "excel_bytes" in st.session_state:
            st.download_button(
                label="📥 Excelダウンロード",
                data=st.session_state.excel_bytes,
                file_name=f"勤務表_{r_year}_{r_month:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
