"""
乖離アラート検出 + 経営管理PL(進行率揃え)生成
- 発生月乖離: 売上計上月の範囲外で原価が発生
- 金額乖離: 原価率が異常
- 整合性: R15合計と明細合計の不一致
- 経営管理PL: 売上と原価を同じ進行率で揃えたPL
"""
import sqlite3

DB_PATH = 'wyse.db'

def detect_divergence_alerts():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # 既存の乖離アラートを削除(再実行用)
    c.execute("DELETE FROM alerts WHERE kind LIKE '乖離_%' OR kind LIKE '整合性_%'")

    alerts_added = 0

    # ===== 1. 発生月乖離アラート =====
    # 売上計上月の範囲外で原価明細が発生している現場(費目: 労務費のみ)
    # 機材レンタル/処分費は使用月や請求月とずれるのが正常なので対象外
    for r in c.execute("""
        SELECT s.genba_no, s.genba_name,
               MIN(a.year*100+a.month) as alloc_min,
               MAX(a.year*100+a.month) as alloc_max
        FROM sites s LEFT JOIN sales_allocations a ON s.genba_no=a.genba_no
        WHERE a.sales_amount_zeikomi > 0
        GROUP BY s.genba_no
    """):
        no, nm, alloc_min, alloc_max = r
        if alloc_min is None:
            continue
        # この現場の労務費明細で、計上月範囲外のもの
        for cost in c.execute("""
            SELECT hi_moku, hassei_date, kingaku_zeibetsu FROM costs
            WHERE genba_no=? AND hi_moku='労務費' AND hassei_date != ''
              AND kingaku_zeibetsu > 0
        """, (no,)):
            hi_moku, hd, amt = cost
            try:
                y = int(hd[:4]); m = int(hd[5:7])
                cost_ym = y * 100 + m
                if cost_ym < alloc_min - 1 or cost_ym > alloc_max + 1:  # ±1ヶ月の余裕
                    c.execute("""INSERT INTO alerts (genba_no, kind, severity, message, detail)
                                 VALUES (?,?,?,?,?)""",
                              (no, '乖離_労務費発生月',
                               'warning',
                               f'労務費の発生月({y}/{m:02d})が売上計上月範囲外',
                               f'売上計上: {alloc_min//100}/{alloc_min%100}〜{alloc_max//100}/{alloc_max%100}, 金額: ¥{amt:,.0f}'))
                    alerts_added += 1
            except (ValueError, TypeError):
                pass

    # ===== 2. 金額乖離アラート =====
    # 採用原価率が90%超 or マイナス → 異常
    for r in c.execute("""
        SELECT s.genba_no, s.genba_name, s.sales_type, s.juchu_zeinuki,
               s.shitauke_zeibetsu, s.genka_yotei, s.genka_jissai, s.hendouhi_shinkou,
               s.is_completed, s.template
        FROM sites s
        WHERE s.juchu_zeinuki > 0
    """):
        no, nm, stype, juchu_zb, shitauke, gy, gj, hendou, cmpl, template = r
        # 採用原価
        if stype == 'motouke':
            cost = shitauke or 0
        else:
            cost = (gj if cmpl else None) or hendou or 0
        if cost == 0:
            # 原価ゼロは要確認(売上があるのに原価情報なし)
            if template != 'NO_SHEET' and juchu_zb > 100000:  # 10万以上で原価なしは怪しい
                c.execute("""INSERT INTO alerts (genba_no, kind, severity, message, detail)
                             VALUES (?,?,?,?,?)""",
                          (no, '乖離_原価ゼロ',
                           'warning',
                           f'受注¥{juchu_zb:,.0f}に対して採用原価がゼロ',
                           f'現場シート未入力の可能性'))
                alerts_added += 1
            continue
        ratio = cost / juchu_zb
        if ratio > 0.9:
            c.execute("""INSERT INTO alerts (genba_no, kind, severity, message, detail)
                         VALUES (?,?,?,?,?)""",
                      (no, '乖離_高原価率',
                       'warning',
                       f'原価率{ratio*100:.1f}% (受注¥{juchu_zb:,.0f}, 原価¥{cost:,.0f})',
                       f'入力ミス or 赤字案件の可能性'))
            alerts_added += 1
        elif ratio > 1.0:
            c.execute("""INSERT INTO alerts (genba_no, kind, severity, message, detail)
                         VALUES (?,?,?,?,?)""",
                      (no, '乖離_原価超過',
                       'error',
                       f'原価が受注を超過: 原価率{ratio*100:.1f}%',
                       f'受注¥{juchu_zb:,.0f}, 原価¥{cost:,.0f}'))
            alerts_added += 1

    # ===== 3. 整合性アラート =====
    # 現場シートR15の費目別合計と、明細合計の差
    # ここでは明細合計のみチェック(R15は別途取り込みが必要)
    # 簡易版: 採用原価と明細合計の差が大きいケース
    for r in c.execute("""
        SELECT s.genba_no, s.genba_name, s.sales_type, s.shitauke_zeibetsu,
               s.genka_yotei, s.genka_jissai, s.is_completed, s.template
        FROM sites s WHERE s.template != 'NO_SHEET'
    """):
        no, nm, stype, shitauke, gy, gj, cmpl, template = r
        if stype == 'motouke':
            target = shitauke or 0
            meisai = c.execute("""SELECT SUM(kingaku_zeibetsu) FROM costs
                                  WHERE genba_no=? AND hi_moku='下請外注'""", (no,)).fetchone()[0] or 0
            if target > 0 and meisai > 0:
                diff = abs(target - meisai)
                if diff / target > 0.1 and diff > 50000:  # 10%以上の差&5万円以上
                    c.execute("""INSERT INTO alerts (genba_no, kind, severity, message, detail)
                                 VALUES (?,?,?,?,?)""",
                              (no, '整合性_下請外注',
                               'warning',
                               f'下請金額(F列)¥{target:,.0f} vs 下請外注明細合計¥{meisai:,.0f}',
                               f'差額¥{diff:,.0f}'))
                    alerts_added += 1

    conn.commit()
    print(f"乖離アラート追加: {alerts_added}件")
    conn.close()
    return alerts_added

def build_keiei_pl():
    """経営管理PL(進行率揃え)生成。
    会計PL(会計帳簿に載るもの)と並べて参照できるよう、別テーブルに保存。

    ロジック:
      - 各現場の「採用原価合計」を「売上の進行率」で月別に按分
      - 結果: cost_allocations_keiei (year, month, hi_moku='経営原価', amount_zeibetsu, source='進行率揃え')
    """
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    c.execute("DROP TABLE IF EXISTS cost_allocations_keiei")
    c.execute("""CREATE TABLE cost_allocations_keiei (
        genba_no TEXT, year INTEGER, month INTEGER,
        amount_zeibetsu REAL,
        source TEXT
    )""")

    n_inserted = 0
    sites_rows = c.execute("""
        SELECT s.genba_no, s.sales_type, s.is_completed,
               s.juchu_zeinuki, s.shitauke_zeibetsu,
               s.genka_yotei, s.genka_jissai, s.hendouhi_shinkou
        FROM sites s
    """).fetchall()
    for r in sites_rows:
        no, stype, cmpl, juchu_zb, shitauke, gy, gj, hendou = r
        # 採用原価決定
        if stype == 'motouke':
            adopted_cost = shitauke or 0
        else:
            if cmpl:
                adopted_cost = gj or hendou or 0
            else:
                adopted_cost = gy or hendou or 0
        if adopted_cost == 0:
            continue

        # この現場の月別売上合計
        alloc = c.execute("""SELECT year, month, sales_amount_zeikomi
                             FROM sales_allocations WHERE genba_no=?
                             AND sales_amount_zeikomi > 0
                             ORDER BY year, month""", (no,)).fetchall()
        if not alloc:
            continue
        total_sales = sum(a[2] for a in alloc)
        if total_sales == 0:
            continue

        # 進行率で按分
        for y, m, amt in alloc:
            ratio = amt / total_sales
            cost_for_month = adopted_cost * ratio
            c.execute("INSERT INTO cost_allocations_keiei VALUES (?,?,?,?,?)",
                      (no, y, m, cost_for_month, '採用原価_進行率按分'))
            n_inserted += 1

    conn.commit()
    print(f"経営管理PL原価配賦: {n_inserted}件挿入")

    # 月別の集計
    print("\n月別 会計PL vs 経営管理PL 比較:")
    print(f"{'年/月':<8} {'売上(税込)':>14} {'会計原価':>14} {'経管原価':>14} {'会計粗利率':>10} {'経管粗利率':>10}")
    for r in c.execute("""
        SELECT a.year, a.month,
               SUM(a.sales_amount_zeikomi) as sales,
               (SELECT SUM(amount_zeibetsu) FROM cost_allocations WHERE year=a.year AND month=a.month) as kc,
               (SELECT SUM(amount_zeibetsu) FROM cost_allocations_keiei WHERE year=a.year AND month=a.month) as kkc
        FROM sales_allocations a GROUP BY a.year, a.month ORDER BY a.year, a.month
    """):
        y, m, s, kc, kkc = r
        s = s or 0; kc = kc or 0; kkc = kkc or 0
        s_zb = s / 1.1
        gpk = ((s_zb - kc)/s_zb*100) if s_zb else 0
        gpkk = ((s_zb - kkc)/s_zb*100) if s_zb else 0
        print(f"{y}/{m:02d}    ¥{s:>12,.0f} ¥{kc:>12,.0f} ¥{kkc:>12,.0f} {gpk:>9.1f}% {gpkk:>9.1f}%")

    conn.close()

if __name__ == '__main__':
    print("=" * 60)
    print("乖離アラート検出 + 経営管理PL生成")
    print("=" * 60)
    detect_divergence_alerts()
    print()
    build_keiei_pl()
