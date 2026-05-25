"""
ワイズアシスト 現場管理アプリ (Streamlitプロトタイプ)

起動方法:
  pip install -r requirements.txt
  python extract.py              # データ抽出(初回および更新時)
  python allocate_costs.py       # 原価PL配賦
  python divergence_and_keiei_pl.py  # 乖離アラートと経営管理PL生成
  streamlit run app.py           # アプリ起動

外部公開時:
  Streamlit Cloud / Render / Fly.io にデプロイ可能。
"""
import streamlit as st
import sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime
import subprocess
import sys
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker
from matplotlib.dates import MonthLocator, DateFormatter

# macOS 日本語フォント設定
matplotlib.rcParams['font.family'] = ['Hiragino Sans', 'Hiragino Maru Gothic Pro',
                                       'AppleGothic', 'sans-serif']

# ===== 設定 =====
st.set_page_config(
    page_title="ワイズアシスト 現場管理",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

DB_PATH = 'wyse.db'
ARCHIVE_CUTOFF_YM = '2026/01'  # これ以降を通常表示、以前はアーカイブ

# ===== データロード =====
@st.cache_data(ttl=300)
def load_sites():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM sites", conn)
    conn.close()
    return df

@st.cache_data(ttl=300)
def load_sales_allocations():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("""
        SELECT a.*, s.genba_name, s.sales_type
        FROM sales_allocations a JOIN sites s ON a.genba_no=s.genba_no
    """, conn)
    conn.close()
    return df

@st.cache_data(ttl=300)
def load_cost_allocations(kind='kaikei'):
    conn = sqlite3.connect(DB_PATH)
    table = 'cost_allocations' if kind == 'kaikei' else 'cost_allocations_keiei'
    df = pd.read_sql_query(f"""
        SELECT c.*, s.genba_name, s.sales_type
        FROM {table} c JOIN sites s ON c.genba_no=s.genba_no
    """, conn)
    conn.close()
    return df

@st.cache_data(ttl=300)
def load_costs():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("""
        SELECT c.*, s.genba_name FROM costs c JOIN sites s ON c.genba_no=s.genba_no
    """, conn)
    conn.close()
    return df

@st.cache_data(ttl=300)
def load_alerts():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("""
        SELECT a.*, s.genba_name FROM alerts a LEFT JOIN sites s ON a.genba_no=s.genba_no
    """, conn)
    conn.close()
    return df

@st.cache_data(ttl=300)
def load_site_periods():
    """現場ごとの売上計上月レンジ（開始月・最終計上月・月数）"""
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("""
        SELECT s.genba_no, s.genba_name, s.sales_type, s.is_completed,
               MIN(a.year * 100 + a.month) AS first_ym,
               MAX(a.year * 100 + a.month) AS last_ym,
               COUNT(DISTINCT a.year * 100 + a.month) AS months_count
        FROM sites s
        LEFT JOIN sales_allocations a ON s.genba_no = a.genba_no
        GROUP BY s.genba_no
        ORDER BY first_ym, s.genba_no
    """, conn)
    conn.close()
    return df

def fmt_ym(ym):
    """202402 → '2024年2月'"""
    if ym is None or pd.isna(ym):
        return '-'
    ym = int(ym)
    return f"{ym // 100}年{ym % 100}月"

def ym_to_dt(ym):
    """202402 → datetime(2024, 2, 1)"""
    if ym is None or pd.isna(ym):
        return None
    ym = int(ym)
    return datetime(ym // 100, ym % 100, 1)

def yen(x):
    if x is None or pd.isna(x):
        return '-'
    return f'¥{int(x):,}'

def pct(x):
    if x is None or pd.isna(x):
        return '-'
    return f'{x*100:.1f}%'

# ===== サイドバー =====
st.sidebar.title("🏗️ ワイズアシスト")
st.sidebar.markdown("**現場管理アプリ v0.1**")
page = st.sidebar.radio("ページ", [
    "📊 概要",
    "🔍 現場ドリルダウン",
    "📈 月次PL",
    "📅 CFカレンダー",
    "🗓️ 工期ガントチャート",
    "⚠️ アラート",
    "📥 MF会計CSV出力",
    "📤 データ更新",
])

# DB存在チェック
if not Path(DB_PATH).exists():
    st.error(f"⚠️ データベースが見つかりません: {DB_PATH}\n\n以下を順に実行してから再度起動してください:\n```\npython extract.py\npython allocate_costs.py\npython divergence_and_keiei_pl.py\n```")
    st.stop()

# ===== 概要ページ =====
if page == "📊 概要":
    st.title("📊 全社サマリ")
    sites = load_sites()
    sales = load_sales_allocations()
    costs_k = load_cost_allocations('kaikei')
    costs_kk = load_cost_allocations('keiei')
    alerts = load_alerts()

    total_juchu = sites['juchu_zeikomi'].sum()
    total_sales = sales['sales_amount_zeikomi'].sum()
    total_cost_k = costs_k['amount_zeibetsu'].sum()
    total_cost_kk = costs_kk['amount_zeibetsu'].sum()
    gp_k = total_sales/1.1 - total_cost_k
    gp_kk = total_sales/1.1 - total_cost_kk
    gpr_k = gp_k / (total_sales/1.1) if total_sales else 0
    gpr_kk = gp_kk / (total_sales/1.1) if total_sales else 0

    # KPI行
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("現場総数", f"{len(sites)}件",
              f"元請{(sites['sales_type']=='motouke').sum()} / 通常解造{(sites['sales_type']=='tsujou_kaitai').sum()}")
    c2.metric("受注金額(税込)", yen(total_juchu))
    c3.metric("月別按分合計", yen(total_sales))
    c4.metric("差分", yen(total_juchu - total_sales),
              "正常" if abs(total_juchu - total_sales) < 10 else "要確認")

    st.markdown("---")

    # PL対比
    cl, cr = st.columns(2)
    with cl:
        st.subheader("📒 会計PL (実発生月ベース)")
        st.metric("原価合計", yen(total_cost_k))
        st.metric("粗利", yen(gp_k), pct(gpr_k))
    with cr:
        st.subheader("📗 経営管理PL (進行率揃え)")
        st.metric("原価合計", yen(total_cost_kk))
        st.metric("粗利", yen(gp_kk), pct(gpr_kk))

    st.markdown("---")

    # アラート集計
    st.subheader("⚠️ データ品質アラート")
    err = (alerts['severity']=='error').sum()
    warn = (alerts['severity']=='warning').sum()
    info = (alerts['severity']=='info').sum()
    c1, c2, c3 = st.columns(3)
    c1.metric("🔴 エラー", err)
    c2.metric("🟡 警告", warn)
    c3.metric("ℹ️ 情報", info)
    if err > 0:
        st.error(f"{err}件のエラーがあります。「アラート」タブで確認してください。")

    st.markdown("---")

    # 直近3ヶ月の月次PL（2026年1月以降）
    st.subheader("📈 直近3ヶ月の月次PL対比")
    sales['ym'] = sales['year'].astype(str) + '/' + sales['month'].astype(str).str.zfill(2)
    costs_k['ym'] = costs_k['year'].astype(str) + '/' + costs_k['month'].astype(str).str.zfill(2)
    costs_kk['ym'] = costs_kk['year'].astype(str) + '/' + costs_kk['month'].astype(str).str.zfill(2)
    m_sales = sales.groupby('ym')['sales_amount_zeikomi'].sum()
    m_k = costs_k.groupby('ym')['amount_zeibetsu'].sum()
    m_kk = costs_kk.groupby('ym')['amount_zeibetsu'].sum()
    df_pl = pd.DataFrame({
        '売上(税込)': m_sales,
        '会計原価': m_k,
        '経管原価': m_kk,
    }).fillna(0)
    df_pl = df_pl[df_pl.index >= ARCHIVE_CUTOFF_YM].tail(3)
    df_pl['会計粗利率'] = (df_pl['売上(税込)']/1.1 - df_pl['会計原価']) / (df_pl['売上(税込)']/1.1)
    df_pl['経管粗利率'] = (df_pl['売上(税込)']/1.1 - df_pl['経管原価']) / (df_pl['売上(税込)']/1.1)
    st.dataframe(df_pl.style.format({
        '売上(税込)': '¥{:,.0f}', '会計原価': '¥{:,.0f}', '経管原価': '¥{:,.0f}',
        '会計粗利率': '{:.1%}', '経管粗利率': '{:.1%}',
    }), use_container_width=True)

# ===== 現場ドリルダウン =====
elif page == "🔍 現場ドリルダウン":
    st.title("🔍 現場別収益性ドリルダウン")
    sites = load_sites()
    sales = load_sales_allocations()
    costs = load_costs()
    cost_alloc = load_cost_allocations('kaikei')
    cost_alloc_kk = load_cost_allocations('keiei')
    alerts = load_alerts()
    site_periods = load_site_periods()

    # フィルタ
    cf1, cf2, cf3 = st.columns(3)
    with cf1:
        type_filter = st.multiselect("売上種別", ["motouke", "tsujou_kaitai"],
                                      default=["motouke", "tsujou_kaitai"],
                                      format_func=lambda x: {"motouke":"元請","tsujou_kaitai":"通常/解体造成"}[x])
    with cf2:
        compl_filter = st.multiselect("完了状態", [0, 1], default=[0, 1],
                                       format_func=lambda x: {0:"未完了", 1:"完了"}[x])
    with cf3:
        search = st.text_input("現場No or 現場名で検索", "")

    df = sites[sites['sales_type'].isin(type_filter)]
    df = df[df['is_completed'].isin(compl_filter)]
    if search:
        mask = df['genba_no'].str.contains(search, case=False, na=False) | df['genba_name'].str.contains(search, case=False, na=False)
        df = df[mask]

    # 採用原価
    def adopted_cost(row):
        if row['sales_type'] == 'motouke':
            return row['shitauke_zeibetsu'] or 0
        if row['is_completed']:
            return row['genka_jissai'] or row['hendouhi_shinkou'] or 0
        return row['genka_yotei'] or row['hendouhi_shinkou'] or 0

    df = df.copy()
    df['採用原価'] = df.apply(adopted_cost, axis=1)
    df['粗利'] = df['juchu_zeinuki'].fillna(df['juchu_zeikomi']/1.1) - df['採用原価']
    df['粗利率'] = df['粗利'] / df['juchu_zeinuki'].fillna(df['juchu_zeikomi']/1.1)

    # 工期情報をマージ
    df = df.merge(
        site_periods[['genba_no', 'first_ym', 'last_ym', 'months_count']],
        on='genba_no', how='left'
    )
    df['開始月'] = df['first_ym'].apply(fmt_ym)
    df['最終計上月'] = df['last_ym'].apply(fmt_ym)

    # 通常 / アーカイブ 分割
    _cy, _cm = map(int, ARCHIVE_CUTOFF_YM.split('/'))
    _cutoff_int = _cy * 100 + _cm
    df_normal  = df[df['first_ym'].notna() & (df['first_ym'] >= _cutoff_int)]
    df_archive = df[df['first_ym'].isna()  | (df['first_ym'] <  _cutoff_int)]

    def render_site_list(df_tab, tab_key):
        """一覧テーブル + 選択した現場の詳細を描画"""
        st.markdown(f"**{len(df_tab)}件**")
        if df_tab.empty:
            st.info("この条件に該当する現場はありません。")
            return

        df_show = df_tab[['genba_no', 'genba_name', 'sales_type', 'is_completed',
                           '開始月', '最終計上月',
                           'juchu_zeikomi', '採用原価', '粗利', '粗利率', 'b4_status']].copy()
        df_show['sales_type'] = df_show['sales_type'].map({"motouke":"元請","tsujou_kaitai":"通常/解体造成"})
        df_show['is_completed'] = df_show['is_completed'].map({0:"未完了", 1:"完了"})
        df_show.columns = ['現場No', '現場名', '種別', '完了',
                           '開始月', '最終計上月',
                           '受注税込', '採用原価', '粗利', '粗利率', 'B4状態']

        selected = st.dataframe(
            df_show.style.format({'受注税込': '¥{:,.0f}', '採用原価': '¥{:,.0f}',
                                  '粗利': '¥{:,.0f}', '粗利率': '{:.1%}'}),
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-row",
            key=f"dd_{tab_key}",
        )

        if not (selected and selected.selection.rows):
            return

        idx = selected.selection.rows[0]
        no  = df_show.iloc[idx]['現場No']
        row = df_tab[df_tab['genba_no'] == no].iloc[0]

        st.markdown("---")
        st.subheader(f"🏢 {no} - {row['genba_name']}")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("受注金額(税込)", yen(row['juchu_zeikomi']))
        c2.metric("採用原価(税別)", yen(row['採用原価']))
        c3.metric("粗利", yen(row['粗利']))
        c4.metric("粗利率", pct(row['粗利率']))

        pr = site_periods[site_periods['genba_no'] == no]
        first_ym_ = pr['first_ym'].iloc[0] if not pr.empty else None
        last_ym_  = pr['last_ym'].iloc[0]  if not pr.empty else None
        mc_       = pr['months_count'].iloc[0] if not pr.empty else None
        cp1, cp2, cp3 = st.columns(3)
        cp1.metric("📅 開始月（初回計上）", fmt_ym(first_ym_))
        cp2.metric("📅 最終計上月", fmt_ym(last_ym_))
        cp3.metric("⏱️ 計上月数", f"{int(mc_)}ヶ月" if mc_ and not pd.isna(mc_) else '-')

        info = pd.DataFrame({
            '項目': ['発注者', '工事区分', '完了状態', '着工日', '入金予定日', 'B4売上計上月', 'B4状態', 'テンプレ'],
            '値': [row['hatchusha'] or '-', row['koji_kubun'] or '-',
                   '完了' if row['is_completed'] else '未完了',
                   row['chakkou_date'] or '-', row['nyukin_yotei_date'] or '-',
                   row['b4_raw'] or '-', row['b4_status'], row['template']]
        })
        st.dataframe(info, hide_index=True, use_container_width=False)

        s = sales[sales['genba_no'] == no][['year', 'month', 'ratio_wari',
                                             'sales_amount_zeikomi', 'shitauke_amount_zeikomi', 'source']].copy()
        ca  = cost_alloc[cost_alloc['genba_no'] == no].groupby(['year', 'month'])['amount_zeibetsu'].sum().reset_index()
        ca.columns  = ['year', 'month', '会計原価']
        ckk = cost_alloc_kk[cost_alloc_kk['genba_no'] == no].groupby(['year', 'month'])['amount_zeibetsu'].sum().reset_index()
        ckk.columns = ['year', 'month', '経管原価']
        merged = (s.merge(ca, on=['year','month'], how='outer')
                   .merge(ckk, on=['year','month'], how='outer')
                   .fillna(0).sort_values(['year','month']))
        merged['年月'] = merged['year'].astype(int).astype(str) + '/' + merged['month'].astype(int).astype(str).str.zfill(2)
        merged = merged[['年月', 'ratio_wari', 'sales_amount_zeikomi', 'shitauke_amount_zeikomi', '会計原価', '経管原価', 'source']]
        merged.columns = ['年月', '割合(割)', '売上(税込)', '下請按分', '会計原価', '経管原価', '計上ソース']
        st.markdown("**月別売上 × 原価**")
        st.dataframe(
            merged.style.format({'割合(割)': '{:.2f}', '売上(税込)': '¥{:,.0f}', '下請按分': '¥{:,.0f}',
                                 '会計原価': '¥{:,.0f}', '経管原価': '¥{:,.0f}'}),
            use_container_width=True, hide_index=True
        )

        cm = costs[costs['genba_no'] == no][['hi_moku', 'torihiki_saki', 'hassei_date',
                                              'shiharai_date', 'kingaku_zeibetsu', 'kingaku_zeikomi', 'memo']].copy()
        if not cm.empty:
            st.markdown("**原価明細**")
            cm.columns = ['費目', '取引先', '発生日', '支払日', '金額(税別)', '金額(税込)', '備考']
            st.dataframe(
                cm.style.format({'金額(税別)': '¥{:,.0f}', '金額(税込)': '¥{:,.0f}'}),
                use_container_width=True, hide_index=True
            )

        my_alerts = alerts[alerts['genba_no'] == no]
        if not my_alerts.empty:
            st.markdown("**⚠️ この現場に関連するアラート**")
            for _, a in my_alerts.iterrows():
                icon = "🔴" if a['severity']=='error' else ("🟡" if a['severity']=='warning' else "ℹ️")
                st.markdown(f"{icon} **[{a['kind']}]** {a['message']}")
                if a['detail']:
                    st.caption(a['detail'])

    tab_normal, tab_archive = st.tabs([
        f"📊 通常（{ARCHIVE_CUTOFF_YM}〜）",
        "📦 アーカイブ（〜2025/12）",
    ])
    with tab_normal:
        render_site_list(df_normal, 'normal')
    with tab_archive:
        st.caption("2025年12月以前に開始した現場、または売上計上月未設定の現場です。")
        render_site_list(df_archive, 'archive')

# ===== 月次PL =====
elif page == "📈 月次PL":
    st.title("📈 月次PL対比")
    st.caption("会計PL(実発生月) vs 経営管理PL(進行率揃え)")

    sales = load_sales_allocations()
    costs_k = load_cost_allocations('kaikei')
    costs_kk = load_cost_allocations('keiei')

    sales['ym'] = sales['year'].astype(str) + '/' + sales['month'].astype(str).str.zfill(2)
    costs_k['ym'] = costs_k['year'].astype(str) + '/' + costs_k['month'].astype(str).str.zfill(2)
    costs_kk['ym'] = costs_kk['year'].astype(str) + '/' + costs_kk['month'].astype(str).str.zfill(2)
    m_s = sales.groupby('ym')['sales_amount_zeikomi'].sum()
    m_mt = sales[sales['sales_type']=='motouke'].groupby('ym')['sales_amount_zeikomi'].sum()
    m_ts = sales[sales['sales_type']=='tsujou_kaitai'].groupby('ym')['sales_amount_zeikomi'].sum()
    m_k = costs_k.groupby('ym')['amount_zeibetsu'].sum()
    m_kk = costs_kk.groupby('ym')['amount_zeibetsu'].sum()
    df_all = pd.DataFrame({
        '元請': m_mt, '通常解造': m_ts, '売上計': m_s,
        '会計原価': m_k, '経管原価': m_kk,
    }).fillna(0).sort_index()
    df_all['会計粗利率'] = (df_all['売上計']/1.1 - df_all['会計原価']) / (df_all['売上計']/1.1)
    df_all['経管粗利率'] = (df_all['売上計']/1.1 - df_all['経管原価']) / (df_all['売上計']/1.1)
    df_all['粗利率差異'] = df_all['経管粗利率'] - df_all['会計粗利率']

    df_normal = df_all[df_all.index >= ARCHIVE_CUTOFF_YM]
    df_archive = df_all[df_all.index < ARCHIVE_CUTOFF_YM]

    PL_FORMAT = {
        '元請': '¥{:,.0f}', '通常解造': '¥{:,.0f}', '売上計': '¥{:,.0f}',
        '会計原価': '¥{:,.0f}', '経管原価': '¥{:,.0f}',
        '会計粗利率': '{:.1%}', '経管粗利率': '{:.1%}', '粗利率差異': '{:+.1%}',
    }

    # 費目別ピボット（通常 / アーカイブ）
    costs_k['ym'] = costs_k['year'].astype(str) + '/' + costs_k['month'].astype(str).str.zfill(2)
    costs_k_normal  = costs_k[costs_k['ym'] >= ARCHIVE_CUTOFF_YM]
    costs_k_archive = costs_k[costs_k['ym'] <  ARCHIVE_CUTOFF_YM]

    HI_MOKU_ORDER = ['労務費', '処分費', '外注費', '機材レンタル', 'アスベスト',
                     '下請外注', '進行管理変動費(按分)', 'その他原価(進行率按分)', '明細(日付なし)']

    def render_cost_breakdown(costs_df):
        """費目別月次原価内訳（スタックバー＋テーブル）"""
        if costs_df.empty:
            return
        pivot = costs_df.pivot_table(
            values='amount_zeibetsu', index='ym',
            columns='hi_moku', aggfunc='sum', fill_value=0
        ).sort_index()
        # 表示順を揃える（存在する列だけ）
        ordered_cols = [c for c in HI_MOKU_ORDER if c in pivot.columns]
        other_cols   = [c for c in pivot.columns if c not in HI_MOKU_ORDER]
        pivot = pivot[ordered_cols + other_cols]

        st.markdown("---")
        st.subheader("📊 月次・費目別原価内訳（会計PL）")

        # スタック棒グラフ
        PALETTE = ['#2e86c1', '#e67e22', '#27ae60', '#8e44ad',
                   '#c0392b', '#16a085', '#d4ac0d', '#7f8c8d', '#2c3e50']
        fig, ax = plt.subplots(figsize=(12, 4))
        bottoms = [0.0] * len(pivot)
        xs = range(len(pivot))
        for i, col in enumerate(pivot.columns):
            vals = pivot[col].values
            ax.bar(xs, vals, bottom=bottoms, label=col,
                   color=PALETTE[i % len(PALETTE)], alpha=0.88, width=0.6)
            bottoms = [b + v for b, v in zip(bottoms, vals)]
        ax.set_xticks(xs)
        ax.set_xticklabels(pivot.index, rotation=40, ha='right', fontsize=8)
        ax.yaxis.set_major_formatter(
            matplotlib.ticker.FuncFormatter(lambda x, _: f'¥{int(x):,}'))
        ax.legend(loc='upper left', fontsize=8, framealpha=0.7)
        ax.set_title('月次費目別原価（スタック）', fontsize=10)
        ax.grid(axis='y', linestyle='--', alpha=0.4)
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

        # 数値テーブル（expander）
        with st.expander("🔢 数値テーブルで確認"):
            fmt_dict = {c: '¥{:,.0f}' for c in pivot.columns}
            pivot_disp = pivot.copy()
            pivot_disp['合計'] = pivot_disp.sum(axis=1)
            fmt_dict['合計'] = '¥{:,.0f}'
            st.dataframe(
                pivot_disp.style.format(fmt_dict)
                    .highlight_between(subset=['労務費'] if '労務費' in pivot_disp.columns else [],
                                       color='#d6eaf8'),
                use_container_width=True
            )

    def render_pl_tab(df, costs_df, tab_key):
        if df.empty:
            st.info("この期間のデータはありません。")
            return

        st.caption("💡 行をクリックすると、その月の現場別内訳を表示します")
        selected_row = st.dataframe(
            df.style.format(PL_FORMAT).background_gradient(subset=['粗利率差異'], cmap='RdYlGn'),
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-row",
            key=f"pl_tbl_{tab_key}",
        )

        # ── 月別現場ドリルダウン ──
        if selected_row and selected_row.selection.rows:
            idx = selected_row.selection.rows[0]
            sel_ym = df.index[idx]
            sel_year, sel_month = map(int, sel_ym.split('/'))

            st.markdown("---")
            st.subheader(f"📋 {sel_ym} 現場別内訳")

            type_sel = st.radio(
                "絞り込み", ["全て", "元請のみ", "通常解造のみ"],
                horizontal=True, key=f"type_sel_{tab_key}"
            )

            _sales = load_sales_allocations()
            _sites = load_sites()

            s = _sales[(_sales['year'] == sel_year) & (_sales['month'] == sel_month)].copy()
            if type_sel == "元請のみ":
                s = s[s['sales_type'] == 'motouke']
            elif type_sel == "通常解造のみ":
                s = s[s['sales_type'] == 'tsujou_kaitai']

            s = s.merge(_sites[['genba_no', 'juchu_zeikomi', 'b4_raw']], on='genba_no', how='left')
            s['種別'] = s['sales_type'].map({"motouke": "元請", "tsujou_kaitai": "通常/解体造成"})
            s['売上計上月(B4)'] = s['b4_raw'].fillna('-').replace('', '-')

            disp = s[['genba_no', 'genba_name', '種別',
                       'juchu_zeikomi', '売上計上月(B4)', 'sales_amount_zeikomi']].copy()
            disp.columns = ['現場No', '現場名', '種別',
                            '請負額(税込)', '売上計上月(B4)', 'この月の計上売上(税込)']
            disp = disp.sort_values('この月の計上売上(税込)', ascending=False)

            st.dataframe(
                disp.style.format({
                    '請負額(税込)': '¥{:,.0f}',
                    'この月の計上売上(税込)': '¥{:,.0f}',
                }),
                use_container_width=True, hide_index=True
            )
            mc1, mc2 = st.columns(2)
            mc1.metric("計上売上合計(税込)", yen(disp['この月の計上売上(税込)'].sum()))
            mc2.metric("現場数", f"{len(disp)}件")

        st.markdown("---")
        st.subheader("月次売上推移グラフ")
        st.bar_chart(df[['元請', '通常解造']])
        st.subheader("月次粗利率推移")
        st.line_chart(df[['会計粗利率', '経管粗利率']])
        render_cost_breakdown(costs_df)

    tab_normal, tab_archive = st.tabs([
        f"📊 通常（{ARCHIVE_CUTOFF_YM}〜）",
        "📦 アーカイブ（〜2025/12）",
    ])
    with tab_normal:
        render_pl_tab(df_normal, costs_k_normal, 'normal')
    with tab_archive:
        st.caption("2025年12月以前のデータです。シート運用が安定する前の期間のため参考値としてご覧ください。")
        render_pl_tab(df_archive, costs_k_archive, 'archive')

# ===== CFカレンダー =====
elif page == "📅 CFカレンダー":
    st.title("📅 キャッシュフローカレンダー")
    sites = load_sites()
    costs = load_costs()

    tab0, tab1, tab2 = st.tabs(["💹 月次CF比較", "💰 入金予定（明細）", "💸 支払予定（明細）"])

    # --- 月次CF集計データ作成 ---
    ar_all = sites[(sites['nyukin_yotei_date'] != '') & (sites['juchu_zeikomi'] > 0)].copy()
    ar_all['ym'] = ar_all['nyukin_yotei_date'].str[:7].str.replace('-', '/')
    m_in = ar_all.groupby('ym')['juchu_zeikomi'].sum().rename('入金予定')

    ap_all = costs[costs['shiharai_date'] != ''].copy()
    ap_all['ym'] = ap_all['shiharai_date'].str[:7].str.replace('-', '/')
    m_out = ap_all.groupby('ym')['kingaku_zeikomi'].sum().rename('支払予定')

    cf = pd.DataFrame({'入金予定': m_in, '支払予定': m_out}).fillna(0).sort_index()
    cf['差引（純CF）'] = cf['入金予定'] - cf['支払予定']
    cf['累計CF'] = cf['差引（純CF）'].cumsum()

    with tab0:
        st.subheader("月次キャッシュフロー比較")
        st.caption("入金予定日（B5）基準の入金 vs 支払日基準の支払。入金予定日未入力の現場は含まれません。")

        if cf.empty:
            st.info("CF データがありません。")
        else:
            # KPI
            k1, k2, k3 = st.columns(3)
            k1.metric("入金予定 合計", yen(cf['入金予定'].sum()))
            k2.metric("支払予定 合計", yen(cf['支払予定'].sum()))
            net = cf['差引（純CF）'].sum()
            k3.metric("差引 合計", yen(net), delta=f"{'黒字' if net >= 0 else '赤字'}")

            # グラフ（棒：入金・支払、折線：差引）
            fig, ax1 = plt.subplots(figsize=(12, 5))
            xs = range(len(cf))
            bar_w = 0.35
            ax1.bar([x - bar_w/2 for x in xs], cf['入金予定'],
                    width=bar_w, label='入金予定', color='#2e86c1', alpha=0.85)
            ax1.bar([x + bar_w/2 for x in xs], cf['支払予定'],
                    width=bar_w, label='支払予定', color='#e67e22', alpha=0.85)
            ax1.set_xticks(xs)
            ax1.set_xticklabels(cf.index, rotation=40, ha='right', fontsize=9)
            ax1.yaxis.set_major_formatter(
                matplotlib.ticker.FuncFormatter(lambda v, _: f'¥{int(v):,}'))
            ax1.set_ylabel('金額（税込）', fontsize=9)
            ax1.legend(loc='upper left', fontsize=9)
            ax1.grid(axis='y', linestyle='--', alpha=0.4)

            ax2 = ax1.twinx()
            ax2.plot(xs, cf['差引（純CF）'], color='#1a7a4a', linewidth=2,
                     marker='o', markersize=5, label='差引（純CF）', zorder=5)
            ax2.axhline(0, color='grey', linewidth=0.8, linestyle=':')
            ax2.yaxis.set_major_formatter(
                matplotlib.ticker.FuncFormatter(lambda v, _: f'¥{int(v):,}'))
            ax2.set_ylabel('差引 純CF', fontsize=9)
            ax2.legend(loc='upper right', fontsize=9)

            ax1.set_title('月次キャッシュフロー（入金 vs 支払）', fontsize=11)
            plt.tight_layout()
            st.pyplot(fig)
            plt.close(fig)

            # 数値テーブル
            st.dataframe(
                cf.style.format({
                    '入金予定': '¥{:,.0f}', '支払予定': '¥{:,.0f}',
                    '差引（純CF）': '¥{:,.0f}', '累計CF': '¥{:,.0f}',
                }).applymap(
                    lambda v: 'color: #1a7a4a' if isinstance(v, (int, float)) and v >= 0
                              else 'color: #c0392b' if isinstance(v, (int, float)) and v < 0
                              else '',
                    subset=['差引（純CF）', '累計CF']
                ),
                use_container_width=True
            )

            no_date = (sites['nyukin_yotei_date'] == '') & (sites['juchu_zeikomi'] > 0)
            if no_date.sum() > 0:
                st.warning(f"⚠️ 入金予定日未入力 {no_date.sum()}件は集計に含まれていません。")

    with tab1:
        st.subheader("入金予定 明細 (現場シートB5「入金予定日」基準)")
        ar = sites[(sites['nyukin_yotei_date'] != '') & (sites['juchu_zeikomi'] > 0)].copy()
        ar = ar[['nyukin_yotei_date', 'genba_no', 'genba_name', 'sales_type', 'juchu_zeikomi']].copy()
        ar['sales_type'] = ar['sales_type'].map({"motouke":"元請","tsujou_kaitai":"通常/解体造成"})
        ar.columns = ['入金予定日', '現場No', '現場名', '種別', '金額(税込)']
        ar = ar.sort_values('入金予定日')
        st.dataframe(
            ar.style.format({'金額(税込)': '¥{:,.0f}'}),
            use_container_width=True, hide_index=True
        )
        st.metric("入金予定合計", yen(ar['金額(税込)'].sum()))

        no_date = (sites['nyukin_yotei_date'] == '') & (sites['juchu_zeikomi'] > 0)
        if no_date.sum() > 0:
            st.warning(f"⚠️ 入金予定日未入力: {no_date.sum()}件 (現場シートB5の入力を確認してください)")

    with tab2:
        st.subheader("支払予定 明細 (原価明細「支払日」基準)")
        ap = costs[costs['shiharai_date'] != ''].copy()
        ap = ap[['shiharai_date', 'genba_no', 'genba_name', 'hi_moku', 'torihiki_saki', 'kingaku_zeikomi']].copy()
        ap.columns = ['支払日', '現場No', '現場名', '費目', '取引先', '金額(税込)']
        ap = ap.sort_values('支払日')
        st.dataframe(
            ap.style.format({'金額(税込)': '¥{:,.0f}'}),
            use_container_width=True, hide_index=True
        )
        st.metric("支払予定合計", yen(ap['金額(税込)'].sum()))

# ===== 工期ガントチャート =====
elif page == "🗓️ 工期ガントチャート":
    st.title("🗓️ 工期ガントチャート")
    st.caption("売上計上月（B4）ベースの開始月〜最終計上月を横棒で表示します")

    site_periods = load_site_periods()

    # フィルタ
    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        type_filter = st.multiselect("売上種別", ["motouke", "tsujou_kaitai"],
                                      default=["motouke", "tsujou_kaitai"],
                                      format_func=lambda x: {"motouke":"元請","tsujou_kaitai":"通常/解体造成"}[x])
    with fc2:
        compl_filter = st.multiselect("完了状態", [0, 1], default=[0, 1],
                                       format_func=lambda x: {0:"進行中", 1:"完了"}[x])
    with fc3:
        sort_opt = st.selectbox("並び順", ["開始月（古い順）", "開始月（新しい順）", "現場No順"])

    df_g = site_periods[
        site_periods['sales_type'].isin(type_filter) &
        site_periods['is_completed'].isin(compl_filter) &
        site_periods['first_ym'].notna()
    ].copy()

    if sort_opt == "開始月（古い順）":
        df_g = df_g.sort_values('first_ym', ascending=True)
    elif sort_opt == "開始月（新しい順）":
        df_g = df_g.sort_values('first_ym', ascending=False)
    else:
        df_g = df_g.sort_values('genba_no')

    st.markdown(f"**{len(df_g)}件**（売上計上月なし {site_periods['first_ym'].isna().sum()}件は除外）")

    if df_g.empty:
        st.info("表示するデータがありません。")
    else:
        # 日付変換
        def next_month_dt(d):
            return datetime(d.year + 1, 1, 1) if d.month == 12 else datetime(d.year, d.month + 1, 1)

        df_g['start'] = df_g['first_ym'].apply(ym_to_dt)
        df_g['end']   = df_g['last_ym'].apply(ym_to_dt).apply(next_month_dt)

        # ガント描画（下から積む → sort逆順でループ）
        df_plot = df_g.iloc[::-1].reset_index(drop=True)
        n = len(df_plot)
        fig_h = max(5, n * 0.38)
        fig, ax = plt.subplots(figsize=(14, fig_h))

        COLORS = {
            ('motouke',        1): '#1a5276',
            ('motouke',        0): '#2e86c1',
            ('tsujou_kaitai',  1): '#6e2f0c',
            ('tsujou_kaitai',  0): '#e67e22',
        }

        for i, row in df_plot.iterrows():
            color = COLORS.get((row['sales_type'], int(row['is_completed'])), '#95a5a6')
            dur = (row['end'] - row['start']).days
            ax.barh(i, dur, left=row['start'], height=0.6,
                    color=color, alpha=0.9, edgecolor='white', linewidth=0.4)
            mid = row['start'] + pd.Timedelta(days=dur / 2)
            mc_label = f"{int(row['months_count'])}M" if not pd.isna(row['months_count']) else ''
            ax.text(mid, i, mc_label, ha='center', va='center',
                    fontsize=6.5, color='white', fontweight='bold')

        # Y軸ラベル
        ylabels = []
        for _, row in df_plot.iterrows():
            label = f"{row['genba_no']} {row['genba_name']}"
            ylabels.append(label[:20] + '…' if len(label) > 21 else label)
        ax.set_yticks(range(n))
        ax.set_yticklabels(ylabels, fontsize=7.5)

        # X軸
        ax.xaxis.set_major_locator(MonthLocator(bymonth=[1, 4, 7, 10]))
        ax.xaxis.set_major_formatter(DateFormatter('%Y/%m'))
        ax.xaxis.set_minor_locator(MonthLocator())
        plt.xticks(rotation=40, ha='right', fontsize=8)
        ax.grid(axis='x', which='major', linestyle='--', alpha=0.5)
        ax.grid(axis='x', which='minor', linestyle=':', alpha=0.25)

        # 今日の縦線
        ax.axvline(datetime.now(), color='red', linewidth=1.5,
                   linestyle='--', alpha=0.8, label='今日')

        # 凡例
        legend_items = [
            mpatches.Patch(color='#2e86c1', label='元請（進行中）'),
            mpatches.Patch(color='#1a5276', label='元請（完了）'),
            mpatches.Patch(color='#e67e22', label='通常解造（進行中）'),
            mpatches.Patch(color='#6e2f0c', label='通常解造（完了）'),
            plt.Line2D([0], [0], color='red', linewidth=1.5, linestyle='--', label='今日'),
        ]
        ax.legend(handles=legend_items, loc='lower right', fontsize=8, framealpha=0.8)
        ax.set_title('工期ガントチャート（売上計上月ベース）', fontsize=11, pad=8)
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    # 一覧テーブル（補助）
    with st.expander("📋 数値一覧で確認する"):
        df_table = site_periods.copy()
        df_table['開始月'] = df_table['first_ym'].apply(fmt_ym)
        df_table['最終計上月'] = df_table['last_ym'].apply(fmt_ym)
        df_table['計上月数'] = df_table['months_count'].apply(
            lambda x: f"{int(x)}ヶ月" if x and not pd.isna(x) else '-')
        df_table['種別'] = df_table['sales_type'].map({"motouke":"元請","tsujou_kaitai":"通常/解体造成"})
        df_table['完了'] = df_table['is_completed'].map({0:"進行中", 1:"完了"})
        st.dataframe(
            df_table[['genba_no', 'genba_name', '種別', '完了', '開始月', '最終計上月', '計上月数']].rename(
                columns={'genba_no':'現場No', 'genba_name':'現場名'}
            ),
            use_container_width=True, hide_index=True
        )

# ===== アラート =====
elif page == "⚠️ アラート":
    st.title("⚠️ データ品質アラート一覧")
    alerts = load_alerts()

    c1, c2, c3 = st.columns(3)
    c1.metric("🔴 エラー", (alerts['severity']=='error').sum())
    c2.metric("🟡 警告", (alerts['severity']=='warning').sum())
    c3.metric("ℹ️ 情報", (alerts['severity']=='info').sum())

    sev_filter = st.multiselect("重大度フィルタ", ['error', 'warning', 'info'],
                                  default=['error', 'warning'])
    kind_filter = st.multiselect("カテゴリフィルタ", alerts['kind'].unique().tolist(),
                                   default=alerts['kind'].unique().tolist())
    df = alerts[alerts['severity'].isin(sev_filter) & alerts['kind'].isin(kind_filter)].copy()
    df = df[['severity', 'kind', 'genba_no', 'genba_name', 'message', 'detail']]
    df.columns = ['重大度', 'カテゴリ', '現場No', '現場名', 'メッセージ', '詳細']

    def color_severity(s):
        return ['background-color: #f4cccc' if v == 'error'
                else 'background-color: #fff2cc' if v == 'warning'
                else 'background-color: #e2efda' for v in s]
    st.dataframe(
        df.style.apply(color_severity, subset=['重大度']),
        use_container_width=True, hide_index=True
    )

# ===== MF会計CSV出力 =====
elif page == "📥 MF会計CSV出力":
    st.title("📥 MF会計仕訳インポートCSV出力")
    st.caption("Shift-JIS、月末日基準で生成します")

    sales = load_sales_allocations()
    available_months = sorted(sales.apply(lambda r: f"{r['year']}/{r['month']:02d}", axis=1).unique())

    c1, c2 = st.columns(2)
    with c1:
        from_ym = st.selectbox("開始年月", available_months, index=0)
    with c2:
        to_ym = st.selectbox("終了年月", available_months, index=len(available_months)-1)
    from_y, from_m = map(int, from_ym.split('/'))
    to_y, to_m = map(int, to_ym.split('/'))

    if st.button("CSVを生成", type="primary"):
        import sys, os, io
        sys.path.insert(0, os.path.dirname(__file__) or '.')
        from generate_mf_csv import generate_uriage_csv, generate_shitauke_csv
        conn = sqlite3.connect(DB_PATH)

        # 売上CSV
        u_path = f'/tmp/uriage_{from_ym.replace("/","")}_{to_ym.replace("/","")}.csv'
        generate_uriage_csv(conn, from_y, from_m, to_y, to_m, output_path=u_path)
        with open(u_path, 'rb') as f:
            st.download_button("📥 売上計上CSV ダウンロード", f.read(),
                               file_name=f'MF売上_{from_ym.replace("/","")}-{to_ym.replace("/","")}.csv',
                               mime='text/csv')
        # 下請外注CSV
        s_path = f'/tmp/shitauke_{from_ym.replace("/","")}_{to_ym.replace("/","")}.csv'
        generate_shitauke_csv(conn, from_y, from_m, to_y, to_m, output_path=s_path)
        with open(s_path, 'rb') as f:
            st.download_button("📥 下請外注CSV ダウンロード", f.read(),
                               file_name=f'MF下請外注_{from_ym.replace("/","")}-{to_ym.replace("/","")}.csv',
                               mime='text/csv')
        conn.close()
        st.success("✅ CSV生成完了。上記ボタンからダウンロードしてください。")

# ===== データ更新 =====
elif page == "📤 データ更新":
    st.title("📤 データ更新")
    st.caption("genba.xlsx をアップロードするとデータを再生成し、全ユーザーに反映します（約1〜2分）。")

    if Path(DB_PATH).exists():
        mtime = datetime.fromtimestamp(Path(DB_PATH).stat().st_mtime)
        st.info(f"現在のDB最終更新: **{mtime.strftime('%Y-%m-%d %H:%M')}**")

    uploaded = st.file_uploader(
        "最新の genba.xlsx をアップロード",
        type=["xlsx"],
        help="ワイス_アシスト現場管理表の最新Excelを選択してください。"
    )

    if uploaded:
        st.success(f"✅ ファイル受信: {uploaded.name}  ({uploaded.size / 1024:.0f} KB)")

        if st.button("🔄 データ更新を実行", type="primary", use_container_width=True):
            app_dir = Path(__file__).parent
            github_ok = ("GITHUB_TOKEN" in st.secrets and "GITHUB_REPO" in st.secrets)

            # ── ① 更新前スナップショット ──
            pre_sites = pre_monthly = pre_alerts = None
            if Path(DB_PATH).exists():
                try:
                    _c = sqlite3.connect(DB_PATH)
                    pre_sites   = pd.read_sql_query(
                        "SELECT genba_no, genba_name, juchu_zeikomi, is_completed FROM sites", _c)
                    pre_monthly = pd.read_sql_query(
                        "SELECT year||'/'||printf('%02d',month) AS ym,"
                        " SUM(sales_amount_zeikomi) AS sales"
                        " FROM sales_allocations GROUP BY ym", _c)
                    pre_alerts  = pd.read_sql_query(
                        "SELECT genba_no, kind, severity FROM alerts", _c)
                    _c.close()
                except Exception:
                    pass

            # ── ② パイプライン実行 ──
            error_occurred = False
            with st.status("データを更新中...", expanded=True) as status:

                # Excel保存
                st.write("📂 Excelファイルを保存中...")
                excel_path = app_dir / "genba.xlsx"
                excel_path.write_bytes(uploaded.getvalue())

                pipeline = [
                    ("extract.py",                   "📊 Excelからデータを抽出中..."),
                    ("allocate_costs.py",             "💰 原価配賦を計算中..."),
                    ("divergence_and_keiei_pl.py",    "📈 乖離アラートと経営管理PLを生成中..."),
                ]
                for script, label in pipeline:
                    st.write(label)
                    result = subprocess.run(
                        [sys.executable, script],
                        capture_output=True, text=True, cwd=str(app_dir)
                    )
                    if result.returncode != 0:
                        st.error(f"❌ {script} でエラーが発生しました:\n```\n{result.stderr[-800:]}\n```")
                        status.update(label="❌ 更新失敗", state="error")
                        error_occurred = True
                        break

                if not error_occurred:
                    # GitHubにコミット（Streamlit Cloud 用）
                    if github_ok:
                        st.write("☁️ GitHubにデータを保存中...")
                        try:
                            from github import Github, GithubException
                            g = Github(st.secrets["GITHUB_TOKEN"])
                            repo = g.get_repo(st.secrets["GITHUB_REPO"])
                            db_bytes = (app_dir / "wyse.db").read_bytes()
                            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
                            try:
                                current = repo.get_contents("wyse.db")
                                repo.update_file(
                                    "wyse.db", f"Update wyse.db {now_str}", db_bytes, current.sha)
                            except GithubException:
                                repo.create_file("wyse.db", f"Initial wyse.db {now_str}", db_bytes)
                            st.write("✅ GitHubに保存しました。Streamlit Cloudが自動再デプロイを開始します。")
                            status.update(label="✅ 更新完了！全ユーザーへの反映まで約1〜2分です。", state="complete")
                        except Exception as e:
                            st.warning(f"⚠️ GitHubへの保存に失敗しました: {e}")
                            status.update(label="⚠️ ローカル更新のみ完了", state="complete")
                    else:
                        status.update(label="✅ ローカルDB更新完了（GitHub未連携）", state="complete")

                    # キャッシュクリア（現在のセッションに即反映）
                    st.cache_data.clear()
                    st.balloons()

            # ── ③ 差分サマリ表示 ──
            if not error_occurred and pre_sites is not None:
                st.markdown("---")
                st.subheader("📋 更新差分サマリ")
                try:
                    _c2 = sqlite3.connect(DB_PATH)
                    post_sites   = pd.read_sql_query(
                        "SELECT genba_no, genba_name, juchu_zeikomi, is_completed FROM sites", _c2)
                    post_monthly = pd.read_sql_query(
                        "SELECT year||'/'||printf('%02d',month) AS ym,"
                        " SUM(sales_amount_zeikomi) AS sales"
                        " FROM sales_allocations GROUP BY ym", _c2)
                    post_alerts  = pd.read_sql_query(
                        "SELECT genba_no, kind, severity FROM alerts", _c2)
                    _c2.close()

                    # 現場の増減
                    pre_nos     = set(pre_sites['genba_no'])
                    post_nos    = set(post_sites['genba_no'])
                    added_nos   = post_nos - pre_nos
                    removed_nos = pre_nos  - post_nos

                    # 変更（受注金額 or 完了状態）
                    merged_s = pre_sites.merge(post_sites, on='genba_no', suffixes=('_pre', '_post'))
                    changed  = merged_s[
                        (merged_s['juchu_zeikomi_pre'] != merged_s['juchu_zeikomi_post']) |
                        (merged_s['is_completed_pre']  != merged_s['is_completed_post'])
                    ]

                    dc1, dc2, dc3 = st.columns(3)
                    dc1.metric("🆕 新規現場", f"+{len(added_nos)}件")
                    dc2.metric("🗑️ 削除現場", f"-{len(removed_nos)}件")
                    dc3.metric("✏️ 変更あり",  f"{len(changed)}件")

                    if added_nos:
                        with st.expander(f"🆕 新規追加 ({len(added_nos)}件)"):
                            df_add = post_sites[post_sites['genba_no'].isin(added_nos)][
                                ['genba_no', 'genba_name', 'juchu_zeikomi']].copy()
                            df_add.columns = ['現場No', '現場名', '受注金額(税込)']
                            st.dataframe(df_add.style.format({'受注金額(税込)': '¥{:,.0f}'}),
                                         hide_index=True, use_container_width=True)

                    if removed_nos:
                        with st.expander(f"🗑️ 削除 ({len(removed_nos)}件)"):
                            df_rem = pre_sites[pre_sites['genba_no'].isin(removed_nos)][
                                ['genba_no', 'genba_name']].copy()
                            df_rem.columns = ['現場No', '現場名']
                            st.dataframe(df_rem, hide_index=True, use_container_width=True)

                    if not changed.empty:
                        with st.expander(f"✏️ 変更あり ({len(changed)}件)"):
                            rows = []
                            for _, r in changed.iterrows():
                                if r['juchu_zeikomi_pre'] != r['juchu_zeikomi_post']:
                                    rows.append({
                                        '現場No': r['genba_no'], '現場名': r['genba_name_post'],
                                        '変更項目': '受注金額',
                                        '変更前': f"¥{int(r['juchu_zeikomi_pre']):,}",
                                        '変更後': f"¥{int(r['juchu_zeikomi_post']):,}",
                                    })
                                if r['is_completed_pre'] != r['is_completed_post']:
                                    rows.append({
                                        '現場No': r['genba_no'], '現場名': r['genba_name_post'],
                                        '変更項目': '完了状態',
                                        '変更前': '完了' if r['is_completed_pre'] else '未完了',
                                        '変更後': '完了' if r['is_completed_post'] else '未完了',
                                    })
                            st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

                    # 月次売上の変化
                    pre_m  = pre_monthly.set_index('ym')['sales']
                    post_m = post_monthly.set_index('ym')['sales']
                    all_yms = pre_m.index.union(post_m.index)
                    diff_m  = pd.DataFrame({
                        '更新前': pre_m.reindex(all_yms, fill_value=0),
                        '更新後': post_m.reindex(all_yms, fill_value=0),
                    })
                    diff_m['差分'] = diff_m['更新後'] - diff_m['更新前']
                    diff_m_chg = diff_m[diff_m['差分'] != 0].sort_index()

                    if not diff_m_chg.empty:
                        with st.expander(f"📅 月次売上に変化あり ({len(diff_m_chg)}ヶ月)"):
                            st.dataframe(
                                diff_m_chg.style.format({
                                    '更新前': '¥{:,.0f}', '更新後': '¥{:,.0f}',
                                    '差分': '¥{:+,.0f}'}),
                                use_container_width=True)
                    else:
                        st.success("✅ 月次売上データに変化はありませんでした。")

                    # アラートの変化
                    pre_a_set  = set(zip(pre_alerts['genba_no'],  pre_alerts['kind'],  pre_alerts['severity']))
                    post_a_set = set(zip(post_alerts['genba_no'], post_alerts['kind'], post_alerts['severity']))
                    new_alerts   = post_a_set - pre_a_set
                    fixed_alerts = pre_a_set  - post_a_set
                    if new_alerts or fixed_alerts:
                        with st.expander(
                                f"⚠️ アラート変化（新規 {len(new_alerts)}件 / 解消 {len(fixed_alerts)}件）"):
                            if new_alerts:
                                st.markdown("**🆕 新規アラート**")
                                st.dataframe(
                                    pd.DataFrame(list(new_alerts), columns=['現場No', 'カテゴリ', '重大度']),
                                    hide_index=True, use_container_width=True)
                            if fixed_alerts:
                                st.markdown("**✅ 解消されたアラート**")
                                st.dataframe(
                                    pd.DataFrame(list(fixed_alerts), columns=['現場No', 'カテゴリ', '重大度']),
                                    hide_index=True, use_container_width=True)
                    else:
                        st.success("✅ アラートに変化はありませんでした。")

                except Exception as e:
                    st.warning(f"差分表示でエラーが発生しました: {e}")

# サイドバー: 補助情報
with st.sidebar:
    st.markdown("---")
    if Path(DB_PATH).exists():
        mtime = datetime.fromtimestamp(Path(DB_PATH).stat().st_mtime)
        st.caption(f"DB最終更新: {mtime.strftime('%Y-%m-%d %H:%M')}")
