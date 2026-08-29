# -*- coding: utf-8 -*-
"""日本の祝日と稼働日（暦）

稼働率の分母に使う「その月の稼働日数」を出すためのもの。
外部ライブラリもインターネットも使わず、祝日法のとおりに計算する。

  稼働日 = 平日（月〜金） − 祝日（振替休日・国民の休日を含む） − 会社の休業日

祝日法の改正はすべて織り込んである（2013年以降の稼動日報を対象にしている）。
  ・2016年〜  山の日(8/11) が加わる
  ・2019年    天皇の即位の日(5/1)・即位礼正殿の儀(10/22)。5/2 は国民の休日
  ・2019年まで 天皇誕生日 12/23 ／ 2020年から 2/23
  ・2020年から 体育の日 → スポーツの日
  ・2020・2021年 五輪特例（海の日・山の日・スポーツの日が動く）
"""
import datetime

MON, SUN = 0, 6


def nth_weekday(year, month, weekday, nth):
    """その月の第 nth 週の weekday（0=月）の日付"""
    first = datetime.date(year, month, 1)
    delta = (weekday - first.weekday()) % 7
    return first + datetime.timedelta(days=delta + 7 * (nth - 1))


def equinox(year, spring):
    """春分の日・秋分の日（1980〜2099年で使える近似式）"""
    base = 20.8431 if spring else 23.2488
    day = int(base + 0.242194 * (year - 1980) - (year - 1980) // 4)
    return datetime.date(year, 3 if spring else 9, day)


def _base(year):
    """振替休日・国民の休日を足す前の祝日 {日付: 名前}"""
    d = datetime.date
    h = {
        d(year, 1, 1):   '元日',
        d(year, 2, 11):  '建国記念の日',
        d(year, 4, 29):  '昭和の日',
        d(year, 5, 3):   '憲法記念日',
        d(year, 5, 4):   'みどりの日',
        d(year, 5, 5):   'こどもの日',
        d(year, 11, 3):  '文化の日',
        d(year, 11, 23): '勤労感謝の日',
        nth_weekday(year, 1, MON, 2): '成人の日',
        nth_weekday(year, 9, MON, 3): '敬老の日',
        equinox(year, True):  '春分の日',
        equinox(year, False): '秋分の日',
    }

    # 天皇誕生日（2019年は代替わりの年で無し）
    if year <= 2018:
        h[d(year, 12, 23)] = '天皇誕生日'
    elif year >= 2020:
        h[d(year, 2, 23)] = '天皇誕生日'

    # 海の日・山の日・スポーツの日（2020・2021年は五輪特例で動く）
    if year == 2020:
        h[d(2020, 7, 23)] = '海の日'
        h[d(2020, 7, 24)] = 'スポーツの日'
        h[d(2020, 8, 10)] = '山の日'
    elif year == 2021:
        h[d(2021, 7, 22)] = '海の日'
        h[d(2021, 7, 23)] = 'スポーツの日'
        h[d(2021, 8, 8)] = '山の日'
    else:
        h[nth_weekday(year, 7, MON, 3)] = '海の日'
        h[nth_weekday(year, 10, MON, 2)] = \
            'スポーツの日' if year >= 2020 else '体育の日'
        if year >= 2016:
            h[d(year, 8, 11)] = '山の日'

    # 代替わりの特例（2019年）
    if year == 2019:
        h[d(2019, 5, 1)] = '天皇の即位の日'
        h[d(2019, 10, 22)] = '即位礼正殿の儀の行われる日'
    return h


def holidays(year):
    """その年の祝日 {日付: 名前}（振替休日・国民の休日を含む）"""
    h = dict(_base(year))

    # 振替休日: 祝日が日曜のとき、その後の最初の祝日でない日
    for day in sorted(k for k, _v in h.items() if k.weekday() == SUN):
        nxt = day + datetime.timedelta(days=1)
        while nxt in h:
            nxt += datetime.timedelta(days=1)
        h[nxt] = '振替休日'

    # 国民の休日: 前後を祝日にはさまれた平日（2026年の9/22 など）
    for day in sorted(_base(year)):
        mid = day + datetime.timedelta(days=1)
        if (mid not in h and mid.weekday() != SUN
                and mid + datetime.timedelta(days=1) in _base(year)):
            h[mid] = '国民の休日'
    return h


def parse_closed(lines, year):
    """会社の休業日の設定を、その年の {日付: 名前} にする

    1行に1つ。次の書き方を受け付ける。
      2026-08-13        その年だけ
      08-13             毎年
      2026-08-13..15    範囲（同じ月のとき、終わりは日だけでよい）
      08-13..08-16      範囲（毎年）
      12-29..01-03      年をまたぐ範囲（年末年始）。終わりが始まりより前なら翌年
    「2026/8/13」のようにスラッシュ区切りでも、後ろに覚え書きを付けても構わない。
      08-13  夏季休業
    """
    out = {}
    for raw in lines or []:
        text = str(raw).replace('/', '-').replace('〜', '..').replace('～', '..').strip()
        if not text or text.startswith('#'):
            continue
        head = text.split()[0]                      # 覚え書きは切り落とす
        note = text[len(head):].strip() or '休業日'
        a, _sep, b = head.partition('..')
        # 前の年から始まる範囲も見る。「12-29..01-03」の1月ぶんを拾うため
        for base in (year - 1, year):
            start = _one(a, base)
            if not start:
                continue
            end = _one(b, base, month=start.month, any_year=True) if b else start
            if not end:
                end = start
            if end < start:                         # 年をまたぐ範囲
                try:
                    end = end.replace(year=end.year + 1)
                except ValueError:                  # 2月29日 → 翌年が平年
                    end = datetime.date(end.year + 1, 2, 28)
            day = start
            while day <= end and (day - start).days < 366:
                if day.year == year:
                    out[day] = note
                day += datetime.timedelta(days=1)
    return out


def _one(text, year, month=None, any_year=False):
    """「2026-08-13」「08-13」「15」を日付にする（駄目なら None）

    any_year は範囲の終わり用。「2026-12-29..2027-01-03」のように、
    始まりと違う年が書かれていても受け取れるようにする。
    """
    parts = [p for p in str(text).split('-') if p != '']
    try:
        nums = [int(p) for p in parts]
    except ValueError:
        return None
    try:
        if len(nums) == 3:
            d = datetime.date(*nums)
            return d if (any_year or nums[0] == year) else None
        if len(nums) == 2:
            return datetime.date(year, nums[0], nums[1])
        if len(nums) == 1 and month:
            return datetime.date(year, month, nums[0])
    except ValueError:
        return None
    return None


def month_days(year, month, closed=None):
    """その月の各日を「稼働日か、休みならその理由か」に分ける

    戻り値: [(日, 理由 or None), ...]  理由が None の日が稼働日
    """
    hol = holidays(year)
    off = closed or {}
    out = []
    day = datetime.date(year, month, 1)
    while day.month == month:
        if day in off:
            why = off[day]                       # 会社の休業日を先に見る
        elif day.weekday() >= 5:
            why = '土曜' if day.weekday() == 5 else '日曜'
        else:
            why = hol.get(day)
        out.append((day.day, why))
        day += datetime.timedelta(days=1)
    return out


def workdays(year, month, closed=None):
    """その月の稼働日（日の数字の一覧）"""
    return [d for d, why in month_days(year, month, closed) if why is None]
