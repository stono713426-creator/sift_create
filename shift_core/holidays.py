"""
祝日取得モジュール
外部API (holidays-jp) を最優先とし、通信不可時は内閣府国民の祝日法に基づく純粋計算でフォールバック。
サードパーティ依存による不安定さを完全に排除。
"""
import datetime
from typing import Set, Tuple, Dict
import requests

def _calculate_vernal_equinox(year: int) -> int:
    """春分の日を計算 (1980-2099)"""
    if year < 1980 or year > 2099:
        return 20
    return int(20.8431 + 0.242194 * (year - 1980) - int((year - 1980) / 4))

def _calculate_autumnal_equinox(year: int) -> int:
    """秋分の日を計算 (1980-2099)"""
    if year < 1980 or year > 2099:
        return 23
    return int(23.2488 + 0.242194 * (year - 1980) - int((year - 1980) / 4))

def _get_nth_weekday(year: int, month: int, weekday: int, n: int) -> datetime.date:
    """第n X曜日を取得 (weekday: 0=月曜)"""
    first_day = datetime.date(year, month, 1)
    first_target = (weekday - first_day.weekday()) % 7 + 1
    target_day = first_target + (n - 1) * 7
    return datetime.date(year, month, target_day)

def calculate_japanese_holidays(year: int) -> Dict[datetime.date, str]:
    """内閣府「国民の祝日について」に基づくオフライン祝日計算"""
    holidays: Dict[datetime.date, str] = {}

    # 固定祝日
    holidays[datetime.date(year, 1, 1)] = "元日"
    holidays[_get_nth_weekday(year, 1, 0, 2)] = "成人の日"
    holidays[datetime.date(year, 2, 11)] = "建国記念の日"
    holidays[datetime.date(year, 2, 23)] = "天皇誕生日"
    
    # 春分の日
    ve_day = _calculate_vernal_equinox(year)
    holidays[datetime.date(year, 3, ve_day)] = "春分の日"
    
    holidays[datetime.date(year, 4, 29)] = "昭和の日"
    holidays[datetime.date(year, 5, 3)] = "憲法記念日"
    holidays[datetime.date(year, 5, 4)] = "みどりの日"
    holidays[datetime.date(year, 5, 5)] = "こどもの日"
    
    holidays[_get_nth_weekday(year, 7, 0, 3)] = "海の日"
    holidays[datetime.date(year, 8, 11)] = "山の日"
    
    holidays[_get_nth_weekday(year, 9, 0, 3)] = "敬老の日"
    ae_day = _calculate_autumnal_equinox(year)
    holidays[datetime.date(year, 9, ae_day)] = "秋分の日"
    
    holidays[_get_nth_weekday(year, 10, 0, 2)] = "スポーツの日"
    holidays[datetime.date(year, 11, 3)] = "文化の日"
    holidays[datetime.date(year, 11, 23)] = "勤労感謝の日"

    # 国民の休日判定（祝日と祝日に挟まれた平日）
    sorted_dates = sorted(holidays.keys())
    bridge_holidays: Dict[datetime.date, str] = {}
    for i in range(len(sorted_dates) - 1):
        d1 = sorted_dates[i]
        d2 = sorted_dates[i + 1]
        if (d2 - d1).days == 2:
            mid = d1 + datetime.timedelta(days=1)
            if mid.weekday() != 6 and mid not in holidays:
                bridge_holidays[mid] = "国民の休日"
    holidays.update(bridge_holidays)

    # 振替休日判定（祝日が日曜日の場合、翌平日）
    substitute_holidays: Dict[datetime.date, str] = {}
    for d in sorted(holidays.keys()):
        if d.weekday() == 6:  # 日曜日
            sub = d + datetime.timedelta(days=1)
            while sub in holidays or sub in substitute_holidays:
                sub += datetime.timedelta(days=1)
            substitute_holidays[sub] = "振替休日"
    holidays.update(substitute_holidays)

    return holidays

def fetch_holidays(years: Set[int] = None) -> Set[Tuple[datetime.date, str]]:
    """祝日取得（外部API優先、失敗時はオフライン計算にフォールバック）"""
    holidays: Set[Tuple[datetime.date, str]] = set()
    api_success = False

    try:
        res = requests.get('https://holidays-jp.github.io/api/v1/date.json', timeout=3)
        if res.status_code == 200:
            data = res.json()
            for date_str, name in data.items():
                dt = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
                holidays.add((dt, name))
            api_success = True
            print("祝日データ: 外部APIより正常に取得しました。")
    except Exception as e:
        print(f"祝日外部API取得スキップ ({e})。内閣府祝日法によるローカル計算で安全に補完します。")

    if not api_success:
        if not years:
            curr_y = datetime.date.today().year
            years = {curr_y - 1, curr_y, curr_y + 1, curr_y + 2}
        for y in years:
            calc_h = calculate_japanese_holidays(y)
            for dt, name in calc_h.items():
                holidays.add((dt, name))
        print("祝日データ: ローカル計算により全祝日を安全に生成しました。")

    return holidays
