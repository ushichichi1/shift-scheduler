#!/usr/bin/env python3
"""
看護師勤務表自動作成 - Streamlit アプリ
"""
import streamlit as st
import pandas as pd
import calendar
from datetime import date
from io import BytesIO

import jpholiday
import pulp

# 親ディレクトリの shift_scheduler.py からソルバーを再利用
from shift_scheduler import (
    Staff, build_and_solve,
    D, N, A, O, R, V, SHIFTS,
    TIER_A, TIER_AB, TIER_B, TIER_C, VALID_TIERS,
    _get_holidays_and_days_off, _write_one_sheet,
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

# カスタムCSS
st.markdown("""
<style>
    .shift-N { background-color: #4472C4; color: white; font-weight: bold; }
    .shift-A { background-color: #FFF2CC; }
    .shift-O { background-color: #E2EFDA; }
    .shift-R { background-color: #E8D5F5; }
    .shift-V { background-color: #F4CCCC; }
    .stDataFrame { font-size: 12px; }
    div[data-testid="stMetric"] { background: #f8f9fa; padding: 10px; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# セッション初期化
# ============================================================
WEEKDAY_MAP = {"月": 0, "火": 1, "水": 2, "木": 3, "金": 4, "土": 5, "日": 6}
WEEKDAY_REV = {v: k for k, v in WEEKDAY_MAP.items()}

def _default_staff():
    """初期スタッフリスト"""
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

if "staff_df" not in st.session_state:
    st.session_state.staff_df = _default_staff()
if "requests_df" not in st.session_state:
    st.session_state.requests_df = None
if "results" not in st.session_state:
    st.session_state.results = None
if "gsheet_loaded" not in st.session_state:
    st.session_state.gsheet_loaded = False

# ============================================================
# サイドバー: 設定
# ============================================================
st.sidebar.title("勤務表設定")

col_y, col_m = st.sidebar.columns(2)
year = col_y.number_input("対象年", 2024, 2030, 2026)
month = col_m.number_input("対象月", 1, 12, 5)

num_days = calendar.monthrange(year, month)[1]
holidays_auto, weekends_auto, auto_public_off = _get_holidays_and_days_off(year, month)

st.sidebar.markdown("---")
st.sidebar.subheader("勤務条件")

public_off_mode = st.sidebar.radio("公休日数", ["自動（土日祝）", "手動指定"], horizontal=True)
if public_off_mode == "自動（土日祝）":
    public_off = auto_public_off
    st.sidebar.info(f"公休 {public_off}日（土日{len(weekends_auto)} + 祝{len(holidays_auto)}）")
    po_override = None
else:
    public_off = st.sidebar.number_input("公休日数", 0, num_days, auto_public_off)
    po_override = public_off

min_day = st.sidebar.number_input("日勤最低人数", 1, 20, 5)
night_count = st.sidebar.number_input("夜勤人数/日", 1, 5, 2)

st.sidebar.markdown("---")
st.sidebar.subheader("夜勤設定")
col1, col2 = st.sidebar.columns(2)
max_n_reg = col1.number_input("通常 上限", 1, 15, 5)
pref_n_reg = col2.number_input("通常 推奨", 1, 15, 4)
col3, col4 = st.sidebar.columns(2)
max_n_ded = col3.number_input("専従 上限", 1, 20, 10)
pref_n_ded = col4.number_input("専従 推奨", 1, 20, 9)

st.sidebar.markdown("---")
st.sidebar.subheader("連勤設定")
col5, col6 = st.sidebar.columns(2)
max_consec = col5.number_input("最大連勤", 1, 10, 5)
pref_consec = col6.number_input("推奨連勤", 1, 10, 4)

time_limit = st.sidebar.number_input("計算時間上限（秒）", 10, 600, 120)
num_patterns = st.sidebar.number_input("生成パターン数", 1, 5, 3)

# 設定dict構築
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

tab1, tab2, tab3, tab4 = st.tabs(["📋 スタッフ一覧", "📝 勤務希望", "⚡ 生成", "📊 結果"])

# ============================================================
# Tab 1: スタッフ一覧
# ============================================================
with tab1:
    st.subheader("スタッフ一覧")
    st.caption("行を追加・編集できます。Tierは A/AB/B/C から選択。")

    # Google Sheets読み込み
    with st.expander("Google スプレッドシートから読み込み", expanded=False):
        gsheet_id = st.text_input("スプレッドシートID or URL",
                                   value="1mezi6NHOQZj0VR_TzmRb9UdkFGrFiWg7qJmvOrJopgA")
        if st.button("読み込み", key="load_gsheet"):
            try:
                with st.spinner("読み込み中..."):
                    from shift_scheduler import load_gsheet as _load_gs
                    staff_list, reqs_dict, gs_settings, _ = _load_gs(gsheet_id)

                # スタッフ → DataFrame
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
                st.session_state.staff_df = pd.DataFrame(rows)

                # 希望 → DataFrame
                req_rows = []
                for name, rq in reqs_dict.items():
                    row = {"名前": name}
                    sn = {D: "日", N: "夜", O: "休", R: "研",
                          "夜不": "夜不", "休暇": "休暇", "明休": "明休"}
                    for d in range(1, num_days + 1):
                        row[str(d)] = sn.get(rq.get(d, ""), rq.get(d, ""))
                    req_rows.append(row)
                # スタッフ一覧にいるが希望未入力の人も追加
                existing = {r["名前"] for r in req_rows}
                for s in staff_list:
                    if s.name not in existing:
                        row = {"名前": s.name}
                        for d in range(1, num_days + 1):
                            row[str(d)] = ""
                        req_rows.append(row)
                st.session_state.requests_df = pd.DataFrame(req_rows)
                st.session_state.gsheet_loaded = True
                st.success(f"✅ {len(staff_list)}人のスタッフと希望を読み込みました")
                st.rerun()
            except Exception as e:
                st.error(f"読み込みエラー: {e}")

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

    # 祝日情報表示
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

    staff_names = [n for n in edited_staff["名前"].dropna().tolist() if n.strip()]

    # 初期化 or スタッフ変更時に再構築
    if st.session_state.requests_df is None or set(st.session_state.requests_df["名前"].tolist()) != set(staff_names):
        if st.session_state.requests_df is not None:
            # 既存データを保持しつつ更新
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

    # 曜日ヘッダー表示
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

    # プレビュー
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
        # DataFrameからStaffリスト構築
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
            # 希望dict構築
            shift_map = {"日": D, "夜": N, "休": O, "研": R,
                         "夜不": "夜不", "休暇": "休暇", "明休": "明休"}
            reqs_dict = {}
            for _, row in edited_reqs.iterrows():
                name = str(row["名前"]).strip()
                if not name:
                    continue
                rq = {}
                for d in range(1, num_days + 1):
                    val = str(row.get(str(d), "")).strip()
                    if val and val in shift_map:
                        rq[d] = shift_map[val]
                    elif val and val in (D, N, O, R):
                        rq[d] = val
                if rq:
                    reqs_dict[name] = rq

            # ソルバー実行
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

                # Excelダウンロード用にメモリ内で生成
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

        # パターン選択
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

        # 夜勤カウント
        night_counts = {}
        for s in names:
            nc = schedule[s].count(N)
            if nc > 0:
                night_counts[s] = nc

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("公休日数", f"{r_public_off}日")
        if night_counts:
            # フルタイム非専従の夜勤均等
            nc_vals = [v for s, v in night_counts.items()
                       if r_weekly.get(s) is None and not r_dedicated.get(s, False)]
            if nc_vals:
                col2.metric("夜勤均等", f"{min(nc_vals)}〜{max(nc_vals)}回")
        col3.metric("未達希望", f"{sum(len(v) for v in missed.values())}件" if missed else "0件 ✓")
        col4.metric("パターン", f"{pat_idx+1} / {len(results)}")

        # 勤務表テーブル
        wdj = ["月", "火", "水", "木", "金", "土", "日"]
        fwd = date(r_year, r_month, 1).weekday()

        # シフトの色マップ（HTML用）
        shift_colors = {
            D: "#FFFFFF", N: "#4472C4", A: "#FFF2CC",
            O: "#E2EFDA", R: "#E8D5F5", V: "#F4CCCC"
        }
        shift_text_colors = {N: "#FFFFFF"}

        # DataFrame構築
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

        # スタイル適用
        def color_shifts(val):
            if val in shift_colors:
                bg = shift_colors[val]
                fg = shift_text_colors.get(val, "#000000")
                return f"background-color: {bg}; color: {fg}; text-align: center; font-weight: {'bold' if val == N else 'normal'}"
            return "text-align: center"

        styled = df.style.map(color_shifts, subset=day_cols)
        st.dataframe(styled, use_container_width=True, height=600, hide_index=True)

        # 日別夜勤ペア
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

        # 未達希望
        if missed:
            with st.expander("⚠ 未達希望", expanded=True):
                for s, ds in missed.items():
                    st.warning(f"{s}: {', '.join(f'{d}日' for d in ds)}")
        else:
            st.success("✓ 全希望達成")

        # 日勤集計
        with st.expander("📈 日別集計", expanded=False):
            summary_data = {"日付": [f"{d+1}" for d in range(r_num_days)]}
            for label, shift in [("日勤", D), ("夜勤", N), ("明け", A), ("休み", O), ("研修", R)]:
                summary_data[label] = [sum(1 for s in names if schedule[s][d] == shift)
                                        for d in range(r_num_days)]
            st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)

        # ソルバーログ
        with st.expander("🔧 ソルバーログ"):
            st.code(st.session_state.get("console_output", ""))

        # Excel ダウンロード
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
