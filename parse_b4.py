"""
現場シートB4「売上計上月」のテキストパース。
標準パターン、分数、省略、改行混在に対応。
年は着工日・入金予定日を基準に推定。
"""
import re
from datetime import datetime, date
from fractions import Fraction

def _normalize(text):
    """全角→半角、改行→空白、連続空白を1個に。分数の/は保護。"""
    if text is None:
        return ''
    s = str(text)
    # 全角→半角
    trans = str.maketrans({
        '％': '%', '　': ' ',
        '０':'0','１':'1','２':'2','３':'3','４':'4',
        '５':'5','６':'6','７':'7','８':'8','９':'9',
    })
    s = s.translate(trans)
    # 改行
    s = re.sub(r'[\r\n]+', ' ', s)
    # 分数(数字/数字)を一時マーカーに退避
    s = re.sub(r'(\d+)/(\d+)', r'\1__FRAC__\2', s)
    # 区切り文字を空白化(、,／ → 空白) ※半角/は分数なので扱わない
    s = re.sub(r'[、,／]', ' ', s)
    # 分数マーカーを戻す
    s = s.replace('__FRAC__', '/')
    # 連続空白
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def _infer_year_month_from_text(text):
    """テキストから年/月を推定。'2026/1末', '4月上旬', '1月', '5月末' などに対応。
    返却: (year, month) または (None, None)
    """
    if not text:
        return None, None
    s = str(text)
    # YYYY/M or YYYY/MM パターン
    m = re.search(r'(20\d{2})[/\-．.](\d{1,2})', s)
    if m:
        return int(m.group(1)), int(m.group(2))
    # YYYY年M月
    m = re.search(r'(20\d{2})年\s*(\d{1,2})月', s)
    if m:
        return int(m.group(1)), int(m.group(2))
    # 「N月」のみ → 年は不明
    m = re.search(r'(\d{1,2})月', s)
    if m:
        return None, int(m.group(1))
    return None, None

def _parse_value(token):
    """値文字列を割合(割)に変換: "50%"→5.0, "1/3"→3.333..., None→None
    100%は10割。
    """
    if not token:
        return None
    token = token.strip()
    # 分数
    m = re.match(r'(\d+)/(\d+)$', token)
    if m:
        num, den = int(m.group(1)), int(m.group(2))
        if den == 0:
            return None
        return (num / den) * 10  # 1/3 = 約3.333割

    # %
    m = re.match(r'(\d+(?:\.\d+)?)%?$', token)
    if m:
        v = float(m.group(1))
        # %は10倍で割換算 (50% = 5割)
        return v / 10
    return None

def parse_uriage_tsuki(b4_text, chakkou_date=None, nyukin_date=None,
                       chakkou_text=None, nyukin_text=None):
    """B4テキストをパース。
    chakkou_date/nyukin_date: dateオブジェクト(あれば優先)
    chakkou_text/nyukin_text: 文字列(着工日が文字列で記録されている場合のフォールバック)
    返却: dict
      - status: 'ok' / 'empty' / 'unparseable'
      - allocations: [(year, month, ratio_wari), ...] (statusがokのとき)
      - reason: 失敗理由
      - raw: 正規化後テキスト
    """
    if b4_text is None or str(b4_text).strip() == '':
        return {'status': 'empty', 'allocations': [], 'raw': ''}

    norm = _normalize(b4_text)
    if not norm:
        return {'status': 'empty', 'allocations': [], 'raw': ''}

    # 「毎月出来高」など、定型外の表現は要確認
    special_keywords = ['出来高', '未定', '応相談', '不明']
    if any(k in norm for k in special_keywords):
        return {'status': 'unparseable', 'allocations': [],
                'reason': f'特殊表記: {norm}', 'raw': norm}

    # 月単位のパース: まず月の位置を全て見つける
    # その後、各月の後ろから次の月の前までを「値」として解析
    month_pattern = re.compile(r'(\d+)月')
    month_matches = list(month_pattern.finditer(norm))
    if not month_matches:
        return {'status': 'unparseable', 'allocations': [],
                'reason': f'月パターン検出失敗: {norm}', 'raw': norm}

    month_vals = []
    for i, m in enumerate(month_matches):
        mo = int(m.group(1))
        if mo < 1 or mo > 12:
            return {'status': 'unparseable', 'allocations': [],
                    'reason': f'月の値が範囲外: {mo}', 'raw': norm}
        # 値の探索範囲: この月の末尾 ~ 次の月の先頭
        val_start = m.end()
        val_end = month_matches[i+1].start() if i + 1 < len(month_matches) else len(norm)
        val_text = norm[val_start:val_end].strip()
        # 値があれば抽出(分数 or 数値%)
        v = None
        if val_text:
            # 分数優先
            mfr = re.match(r'(\d+)/(\d+)', val_text)
            if mfr:
                v = _parse_value(mfr.group(0))
            else:
                mv = re.match(r'(\d+(?:\.\d+)?)%?', val_text)
                if mv:
                    v = _parse_value(mv.group(0))
        month_vals.append((mo, v))

    # 割合の決定
    with_vals = [(m, v) for m, v in month_vals if v is not None]
    blank_count = sum(1 for _, v in month_vals if v is None)

    if blank_count == 0:
        # 全部値あり → そのまま (ただし合計チェックは後段で)
        allocations_partial = month_vals
    elif blank_count == len(month_vals):
        # 全部空 → 均等分割
        per = 10.0 / len(month_vals)
        allocations_partial = [(m, per) for m, _ in month_vals]
    else:
        # 混在 → 残りを按分
        sum_vals = sum(v for _, v in with_vals)
        remain = 10.0 - sum_vals
        per = remain / blank_count if blank_count > 0 else 0
        allocations_partial = [(m, v if v is not None else per) for m, v in month_vals]

    # 年の特定: 着工日・入金予定日を起点に、月が連続するよう補完
    start_year = None
    start_month = None

    # 1. dateオブジェクトの着工日
    if chakkou_date and isinstance(chakkou_date, (date, datetime)):
        start_year = chakkou_date.year
        start_month = chakkou_date.month
    # 2. 文字列の着工日から年月推定
    if start_year is None and chakkou_text:
        sy, sm = _infer_year_month_from_text(chakkou_text)
        if sy and sm:
            start_year, start_month = sy, sm
        elif sm:  # 月だけ取れた場合は年を別ソースから
            start_month = sm
    # 3. 入金予定日(dateオブジェクト)から逆算
    if start_year is None and nyukin_date and isinstance(nyukin_date, (date, datetime)):
        if nyukin_date.month == 1:
            start_year = nyukin_date.year - 1
            if start_month is None:
                start_month = 12
        else:
            start_year = nyukin_date.year
            if start_month is None:
                start_month = max(nyukin_date.month - 1, 1)
    # 4. 入金予定日(文字列)から年月推定
    if start_year is None and nyukin_text:
        sy, sm = _infer_year_month_from_text(nyukin_text)
        if sy:
            start_year = sy
            if start_month is None and sm:
                # 入金予定の月の数ヶ月前を着工想定(雑だが起点として機能する)
                start_month = max(sm - 1, 1)
        elif sm and start_month is None:
            start_month = sm

    if start_year is None:
        # 年起点不明 → B4の月リストの最初の月を「直近の同月」と仮定
        # ただしユーザーにはアラートで通知する
        return {'status': 'unparseable', 'allocations': [],
                'reason': '年の起点となる着工日/入金予定日が空(または文字列で解釈不能)',
                'raw': norm}
    if start_month is None:
        start_month = month_vals[0][0]

    # 月→年の補完
    # B4の月リストを順に処理。startより前の月は翌年へ。
    # 売上計上月は通常、着工月以降なので、月がstart_monthより前なら翌年と推定。
    allocations = []
    current_year = start_year
    prev_month = None
    for mo, ratio in allocations_partial:
        if prev_month is None:
            # 最初の月
            if mo < start_month:
                # 起点月より前 → 翌年と推定
                current_year = start_year + 1
            else:
                current_year = start_year
        else:
            # 前の月より小さい → 翌年
            if mo < prev_month:
                current_year += 1
        allocations.append((current_year, mo, ratio))
        prev_month = mo

    # 合計のチェック (情報のみ、警告は呼び出し側で)
    total = sum(r for _, _, r in allocations)
    warnings = []
    if abs(total - 10) > 0.01:
        warnings.append(f'割合合計が10割でない: {total:.2f}')

    return {'status': 'ok', 'allocations': allocations, 'raw': norm,
            'warnings': warnings, 'total_ratio': total}

# ========== テスト ==========
if __name__ == '__main__':
    from datetime import date

    tests = [
        ('12月50％、1月50％', date(2025, 12, 18), None),
        ('12月50% 1月50%', date(2025, 12, 22), None),
        ('1月50% 2月50%', date(2026, 1, 21), None),
        ('1月30% 2月30% 3月40%', date(2026, 1, 14), None),
        ('12月 100%', date(2025, 12, 9), None),
        ('1月30％、2月10％、\n3月10％、4月50％', date(2026, 1, 31), None),
        ('1月100%', date(2026, 1, 1), None),
        ('1月　100%', None, None),  # 着工日なし
        ('3月100%', date(2026, 3, 9), None),
        ('4月、5月', date(2026, 4, 6), None),
        ('3月', date(2026, 3, 24), None),
        ('7月　1/3\n9月　2/3', None, None),
        ('毎月出来高', None, None),
        ('', None, None),
        ('1月10% 2月10% 3月10% 4月0% 5月70%', date(2026, 1, 1), None),
    ]
    for b4, ck, nk in tests:
        r = parse_uriage_tsuki(b4, ck, nk)
        print(f"\nB4='{b4}' 着工={ck}")
        print(f"  → status={r['status']}")
        if r['status'] == 'ok':
            for y, m, ratio in r['allocations']:
                print(f"    {y}/{m:02d}: {ratio:.2f}割 ({ratio*10:.1f}%)")
            if r.get('warnings'):
                print(f"  警告: {r['warnings']}")
        elif r['status'] == 'unparseable':
            print(f"  reason: {r['reason']}")
