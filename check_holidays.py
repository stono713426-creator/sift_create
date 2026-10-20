import sys
import datetime
from shift_core.holidays import fetch_holidays

h_set = fetch_holidays()
h_dict = dict(h_set)
wd_names = ["月", "火", "水", "木", "金", "土", "日"]
for m in [10, 11, 12]:
    print(f"=== 2026年{m}月 ===")
    for d in range(1, 32):
        try:
            dt = datetime.date(2026, m, d)
            if dt in h_dict:
                print(f"{dt} ({wd_names[dt.weekday()]}): {h_dict[dt]}")
        except ValueError:
            pass
