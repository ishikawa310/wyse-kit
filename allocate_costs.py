"""
原価のPL月別配賦ロジック
- 完了現場 → 実原価ベース
- 未完了現場 → 予定原価ベース
- 明細にhassei_dateがあれば発生月へ計上、なければ進行率で按分
"""
import sqlite3
from datetime import datetime

DB_PATH = 'wyse.db'

def main():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # cost_allocations テーブル(あれば再作成)
    c.execute("DROP TABLE IF EXISTS cost_allocations")
    c.execute("""CREATE TABLE cost_allocations (
        genba_no TEXT, year INTEGER, month INTEGER,
        hi_moku TEXT,
        amount_zeibetsu REAL,
        source TEXT
    )""")

    # 全現場を処理
    sites_rows = c.execute("""
        SELECT genba_no, sales_type, is_completed, template,
               genka_yotei, genka_jissai, hendouhi_shinkou, shitauke_zeibetsu,
               juchu_zeikomi
        FROM sites
    """).fetchall()

    stats = {'meisai_proportional': 0, 'mochinashi_shinkou': 0,
             'meisai_only': 0, 'all_proportional': 0, 'skip': 0}

    for row in sites_rows:
        no, sales_type, is_completed, template, ge_y, ge_j, hendou, shitauke, juchu = row

        # 月別進行率(sales_allocations)を取得
        alloc_rows = c.execute(
            "SELECT year, month, sales_amount_zeikomi FROM sales_allocations WHERE genba_no=? ORDER BY year, month",
            (no,)).fetchall()
        if not alloc_rows:
            stats['skip'] += 1
            continue
        total_sales = sum(a[2] for a in alloc_rows)
        if total_sales == 0:
            stats['skip'] += 1
            continue

        # 月別の進行率(=その月の売上 / 全月売上合計)
        progress_map = [(y, m, amt / total_sales) for y, m, amt in alloc_rows]

        # 使用する原価ベースを決定
        # 通常/解体造成: 原価合計(労務費+機材+処分+外注+仕入+燃料)
        # 元請売上: 下請外注のみ
        # 完了現場は実原価、未完了は予定原価
        if sales_type == 'motouke':
            # 元請売上: 原価=下請外注(税抜)
            target_total = shitauke or 0
            relevant_himoku = ('下請外注',)
        else:
            # 通常/解体造成
            if is_completed:
                target_total = ge_j or hendou or 0  # 実原価合計(税別)
            else:
                target_total = ge_y or hendou or 0  # 予定原価合計(税別)
            relevant_himoku = None  # 全費目対象

        # 明細を取得(該当現場のみ)
        cost_rows = c.execute("""
            SELECT hi_moku, hassei_date, kingaku_zeibetsu FROM costs
            WHERE genba_no=? AND kingaku_zeibetsu IS NOT NULL
        """, (no,)).fetchall()
        if relevant_himoku:
            cost_rows = [r for r in cost_rows if r[0] in relevant_himoku]

        meisai_total = sum(r[2] for r in cost_rows if r[2])

        # ケース1: 明細あり(かつhassei_date付き) → 発生月で配賦
        meisai_with_date_total = 0
        for himoku, hd, amt in cost_rows:
            if not amt:
                continue
            if hd and len(hd) >= 7:  # YYYY-MM-DD
                try:
                    d = datetime.fromisoformat(hd[:10])
                    c.execute("INSERT INTO cost_allocations VALUES (?,?,?,?,?,?)",
                              (no, d.year, d.month, himoku, amt, 'meisai_date'))
                    meisai_with_date_total += amt
                except ValueError:
                    pass

        # ケース2: 明細はあるが日付なし → 進行率で按分
        meisai_nodate_total = meisai_total - meisai_with_date_total
        if meisai_nodate_total > 0:
            for y, m, ratio in progress_map:
                amt = meisai_nodate_total * ratio
                c.execute("INSERT INTO cost_allocations VALUES (?,?,?,?,?,?)",
                          (no, y, m, '明細(日付なし)', amt, 'meisai_proportional'))
            stats['meisai_proportional'] += 1

        # ケース3: 明細合計 < 目標合計 → 不足分を進行率で按分
        # 完了現場は明細が確定しているとみなし、gap埋めをしない
        gap = (target_total or 0) - meisai_total
        if gap > 1 and not is_completed:  # 1円以上のギャップ かつ 未完了現場のみ
            for y, m, ratio in progress_map:
                amt = gap * ratio
                label = 'その他原価(進行率按分)' if meisai_total > 0 else '進行管理変動費(按分)'
                source = 'mochinashi_shinkou' if meisai_total == 0 else 'meisai_proportional_residual'
                c.execute("INSERT INTO cost_allocations VALUES (?,?,?,?,?,?)",
                          (no, y, m, label, amt, source))
            if meisai_total == 0:
                stats['mochinashi_shinkou'] += 1

        if meisai_total > 0 and abs(gap) <= 1:
            stats['meisai_only'] += 1

    conn.commit()

    # 検証
    print("=" * 60)
    print("PL配賦結果サマリ")
    print("=" * 60)
    n = c.execute("SELECT COUNT(*) FROM cost_allocations").fetchone()[0]
    total = c.execute("SELECT SUM(amount_zeibetsu) FROM cost_allocations").fetchone()[0] or 0
    print(f"配賦レコード数: {n}件")
    print(f"配賦合計(税別): ¥{total:,.0f}")
    print(f"\n配賦元の内訳:")
    for st, cnt in stats.items():
        print(f"  {st}: {cnt}件")

    print("\n月別原価合計:")
    for r in c.execute("""SELECT year, month, ROUND(SUM(amount_zeibetsu), 0)
                          FROM cost_allocations GROUP BY year, month ORDER BY year, month"""):
        print(f"  {r[0]}/{r[1]:2d}: ¥{r[2]:>12,.0f}")

    conn.close()

if __name__ == '__main__':
    main()
