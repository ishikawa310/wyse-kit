"""
MF会計 仕訳インポートCSV生成
- 売上計上CSV: 借方=未収収益 / 貸方=元請売上 or 解体・造成売上
- 下請外注CSV: 借方=下請外注 / 貸方=買掛金-暫定
Shift-JIS、月末日基準、取引No連番。
"""
import sqlite3
import csv
import calendar
import os
import sys

DB_PATH = 'wyse.db'
OUT_DIR = '.'

HEADER = ['取引No', '取引日', '借方勘定科目', '借方補助科目', '借方税区分', '借方金額', '借方税額',
          '貸方勘定科目', '貸方補助科目', '貸方税区分', '貸方金額', '貸方税額',
          '摘要', '仕訳メモ', 'タグ', 'MF仕訳タイプ']

def month_end(y, m):
    last = calendar.monthrange(y, m)[1]
    return f"{y}/{m:02d}/{last:02d}"

def sort_key_genba_no(no):
    """現場Noを数値部分でソート可能なキーに変換 (F2539 → 2539, 2547 → 2547, 001 → 1)"""
    import re
    m = re.search(r'\d+', str(no))
    return int(m.group()) if m else 0

def fmt_yen(amount):
    """金額を ￥1,234,567 形式に(整数切り捨て表示)"""
    return f"￥{int(amount):,}"

def build_memo_text(no, name, juchu, allocations, is_zeikomi_label=False):
    """摘要テキストを生成
    単月の場合: 「No.{No} {name}」
    多月の場合: 「No.{No} {name}　￥{juchu}{(税込)}　{月1割合%}　{月2割合%}...」
    """
    base = f"No.{no} {name}"
    # 複数月 or 単月でも100%でない場合は明細を付ける(進行管理上で部分計上)
    nonzero = [a for a in allocations if a['ratio_wari'] and a['ratio_wari'] > 0]
    if len(nonzero) <= 1:
        return base
    # 多月明細
    yen_text = fmt_yen(juchu)
    if is_zeikomi_label:
        yen_text += "(税込)"
    # 月別パーセント
    pct_parts = []
    for a in allocations:
        pct = int(round(a['ratio_wari'] * 10))
        pct_parts.append(f"{a['month']}月{pct}%")
    return f"{base}　{yen_text}　" + "　".join(pct_parts)

def generate_uriage_csv(conn, year_from=None, month_from=None, year_to=None, month_to=None, output_path=None):
    """売上計上CSV生成"""
    c = conn.cursor()

    where = ["a.sales_amount_zeikomi > 0"]
    params = []
    if year_from and month_from:
        where.append("(a.year * 100 + a.month) >= ?")
        params.append(year_from * 100 + month_from)
    if year_to and month_to:
        where.append("(a.year * 100 + a.month) <= ?")
        params.append(year_to * 100 + month_to)
    where_sql = " AND ".join(where)

    # 計上対象月の全レコード(年月昇順、現場No数値部分昇順)
    rows = c.execute(f"""
        SELECT a.genba_no, a.year, a.month, a.sales_amount_zeikomi,
               s.genba_name, s.juchu_zeikomi, s.sales_type
        FROM sales_allocations a JOIN sites s ON a.genba_no = s.genba_no
        WHERE {where_sql}
    """, params).fetchall()

    # 年/月昇順、現場No数値部分昇順でソート
    rows = sorted(rows, key=lambda r: (r[1], r[2], sort_key_genba_no(r[0])))

    output = []
    output.append(HEADER)
    n = 1
    for row in rows:
        no, y, m, amount, name, juchu, stype = row
        # その現場の全月別按分(摘要用)
        all_allocs = c.execute("""SELECT year, month, ratio_wari FROM sales_allocations
                                  WHERE genba_no=? ORDER BY year, month""", (no,)).fetchall()
        alloc_list = [{'year': r[0], 'month': r[1], 'ratio_wari': r[2]} for r in all_allocs]

        # 売上科目決定
        kaikamoku = '元請売上' if stype == 'motouke' else '解体・造成売上'
        memo = build_memo_text(no, name, juchu, alloc_list, is_zeikomi_label=False)

        record = [
            str(n),
            month_end(y, m),
            '未収収益', '', '',
            str(int(round(amount))), '',
            kaikamoku, '', '課税売上10%',
            str(int(round(amount))), '',
            memo, '', '', ''
        ]
        output.append(record)
        n += 1

    if output_path:
        with open(output_path, 'w', encoding='shift_jis', newline='', errors='replace') as f:
            w = csv.writer(f)
            for row in output:
                w.writerow(row)
        print(f"  売上CSV保存: {output_path} ({len(output)-1}件)")
    return output

def generate_shitauke_csv(conn, year_from=None, month_from=None, year_to=None, month_to=None, output_path=None):
    """下請外注CSV生成(元請売上のみ対象)"""
    c = conn.cursor()

    where = ["a.shitauke_amount_zeikomi > 0", "s.sales_type='motouke'"]
    params = []
    if year_from and month_from:
        where.append("(a.year * 100 + a.month) >= ?")
        params.append(year_from * 100 + month_from)
    if year_to and month_to:
        where.append("(a.year * 100 + a.month) <= ?")
        params.append(year_to * 100 + month_to)
    where_sql = " AND ".join(where)

    rows = c.execute(f"""
        SELECT a.genba_no, a.year, a.month, a.shitauke_amount_zeikomi,
               s.genba_name, s.shitauke_zeibetsu
        FROM sales_allocations a JOIN sites s ON a.genba_no = s.genba_no
        WHERE {where_sql}
    """, params).fetchall()
    rows = sorted(rows, key=lambda r: (r[1], r[2], sort_key_genba_no(r[0])))

    output = []
    output.append(HEADER)
    n = 1
    for row in rows:
        no, y, m, amount, name, shitauke_zeibetsu = row
        all_allocs = c.execute("""SELECT year, month, ratio_wari FROM sales_allocations
                                  WHERE genba_no=? ORDER BY year, month""", (no,)).fetchall()
        alloc_list = [{'year': r[0], 'month': r[1], 'ratio_wari': r[2]} for r in all_allocs]

        # 下請金額(税込) = 税抜 × 1.1
        juchu_zeikomi_shitauke = (shitauke_zeibetsu or 0) * 1.1
        memo = build_memo_text(no, name, juchu_zeikomi_shitauke, alloc_list, is_zeikomi_label=True)

        record = [
            str(n),
            month_end(y, m),
            '下請外注', '', '課税仕入10%',
            str(int(round(amount))), '',
            '買掛金-暫定', '', '',
            str(int(round(amount))), '',
            memo, '', '', ''
        ]
        output.append(record)
        n += 1

    if output_path:
        with open(output_path, 'w', encoding='shift_jis', newline='', errors='replace') as f:
            w = csv.writer(f)
            for row in output:
                w.writerow(row)
        print(f"  下請外注CSV保存: {output_path} ({len(output)-1}件)")
    return output

def main():
    conn = sqlite3.connect(DB_PATH)
    os.makedirs(OUT_DIR, exist_ok=True)

    # 2026/1〜2026/4 (サンプルと同じ範囲) でテスト生成
    print("=" * 60)
    print("MF会計仕訳CSV 生成")
    print("=" * 60)
    print("\n[期間: 2026/01〜2026/04 - サンプルと同範囲で再現]")
    generate_uriage_csv(conn, 2026, 1, 2026, 4,
                       output_path=f'{OUT_DIR}/MF会計仕訳インポート_売上_2025年5月期.csv')
    generate_shitauke_csv(conn, 2026, 1, 2026, 4,
                         output_path=f'{OUT_DIR}/MF会計仕訳インポート_下請外注_2025年5月期.csv')

    # 全期間版も生成
    print("\n[期間: 全期間 - 参考用]")
    generate_uriage_csv(conn, output_path=f'{OUT_DIR}/MF会計仕訳インポート_売上_全期間.csv')
    generate_shitauke_csv(conn, output_path=f'{OUT_DIR}/MF会計仕訳インポート_下請外注_全期間.csv')

    conn.close()

if __name__ == '__main__':
    main()
