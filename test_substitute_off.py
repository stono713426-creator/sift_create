import sys
import datetime
from shift_core import ShiftGeneratorService
from shift_core.parser import load_and_parse_data
from shift_core.solver import solve_monthly_shift

def test_holiday_substitute():
    print("=== 祝日週の固定公休振替 検証テスト ===")
    filepath = "ope勤務条件表_2026_4sheets.xlsx"
    
    (
        staff_list,
        df_rules,
        ope_dict_all,
        auto_target_months,
        holidays_jp_dict,
        init_tail_shifts,
        init_main_offs,
        config
    ) = load_and_parse_data(filepath)
    
    print(f"動的設定 substitute_holiday_off: {config.substitute_holiday_off}")
    assert config.substitute_holiday_off is True, "substitute_holiday_off should be True"

    # 2026年10月をテスト
    print("\n--- 2026年10月 シフト生成テスト ---")
    monthly_data_10 = {
        'staff_list': staff_list,
        'df_rules': df_rules,
        'ope_dict': ope_dict_all.get((2026, 10), {}),
        'holidays_jp_dict': holidays_jp_dict,
        'target_file': filepath,
        'config': config
    }
    prev_month_data = {
        'tail_shifts': init_tail_shifts,
        'main_offs': init_main_offs
    }
    best_x_10, main_offs_10, target_days_10, tail_10 = solve_monthly_shift(
        2026, 10, monthly_data_10, prev_month_data, stage=1
    )
    if best_x_10 is None:
        print("Stage 1 失敗 -> Stage 2 試行")
        best_x_10, main_offs_10, target_days_10, tail_10 = solve_monthly_shift(
            2026, 10, monthly_data_10, prev_month_data, stage=2
        )
    
    assert best_x_10 is not None, "2026年10月のシフト生成に失敗しました"
    print("2026年10月 シフト生成成功！")

    # 10/12週 (10/12月〜10/17土) の小田切さんの出勤を検証
    # 小田切さん: 2階表, 週5日, 月曜固定公休
    print("\n【小田切さん (2階表・月曜固定休) 10/12週の出勤状況】")
    dates_10_week = [
        datetime.date(2026, 10, 12), # 月(祝)
        datetime.date(2026, 10, 13), # 火
        datetime.date(2026, 10, 14), # 水
        datetime.date(2026, 10, 15), # 木
        datetime.date(2026, 10, 16), # 金
        datetime.date(2026, 10, 17), # 土
    ]
    wd_names = ["月", "火", "水", "木", "金", "土", "日"]
    week_work_count = 0
    for d in dates_10_week:
        work_flag = best_x_10.get(('小田切', d), 0)
        is_h = d in holidays_jp_dict
        h_name = f"({holidays_jp_dict[d]})" if is_h else ""
        status = "出勤(○)" if work_flag == 1 else "公休(休)"
        print(f"  {d} ({wd_names[d.weekday()]}){h_name}: {status}")
        if not is_h and d.weekday() != 6:
            week_work_count += work_flag

    print(f"  -> 10/12週の診療日出勤日数: {week_work_count}日 / 診療日5日中 (目標: 最大4日)")
    assert week_work_count <= 4, f"小田切さんの週出勤日数が4日以下になっていません: {week_work_count}"
    
    # 2階表の各日の出勤人数も表示
    sec_2f_front_staffs = [s['name'] for s in staff_list if s['section'] == '2階表']
    print("\n【10/12週 2階表の各日出勤人数】")
    for d in dates_10_week:
        if d in holidays_jp_dict:
            continue
        cnt = sum(best_x_10.get((n, d), 0) for n in sec_2f_front_staffs)
        odagiri_status = "出勤" if best_x_10.get(('小田切', d), 0) == 1 else "【振替公休】"
        print(f"  {d} ({wd_names[d.weekday()]}): 2階表記出勤 {cnt}名 | 小田切: {odagiri_status}")

    # 2026年11月もテスト
    print("\n--- 2026年11月 シフト生成テスト ---")
    monthly_data_11 = {
        'staff_list': staff_list,
        'df_rules': df_rules,
        'ope_dict': ope_dict_all.get((2026, 11), {}),
        'holidays_jp_dict': holidays_jp_dict,
        'target_file': filepath,
        'config': config
    }
    prev_month_data_11 = {
        'tail_shifts': tail_10,
        'main_offs': main_offs_10
    }
    best_x_11, main_offs_11, target_days_11, tail_11 = solve_monthly_shift(
        2026, 11, monthly_data_11, prev_month_data_11, stage=1
    )
    if best_x_11 is None:
        print("Stage 1 失敗 -> Stage 2 試行")
        best_x_11, main_offs_11, target_days_11, tail_11 = solve_monthly_shift(
            2026, 11, monthly_data_11, prev_month_data_11, stage=2
        )
    
    assert best_x_11 is not None, "2026年11月のシフト生成に失敗しました"
    print("2026年11月 シフト生成成功！")

    # 11/3週 (11/2月〜11/7土) の長谷川さんの出勤を検証
    # 長谷川さん: 2階裏, 週5日, 火曜固定公休 (11/3が文化の日:火)
    print("\n【長谷川さん (2階裏・火曜固定休) 11/3週の出勤状況】")
    dates_11_week = [
        datetime.date(2026, 11, 2), # 月
        datetime.date(2026, 11, 3), # 火(祝)
        datetime.date(2026, 11, 4), # 水
        datetime.date(2026, 11, 5), # 木
        datetime.date(2026, 11, 6), # 金
        datetime.date(2026, 11, 7), # 土
    ]
    hasegawa_work_count = 0
    for d in dates_11_week:
        work_flag = best_x_11.get(('長谷川', d), 0)
        is_h = d in holidays_jp_dict
        h_name = f"({holidays_jp_dict[d]})" if is_h else ""
        status = "出勤(○)" if work_flag == 1 else "公休(休)"
        print(f"  {d} ({wd_names[d.weekday()]}){h_name}: {status}")
        if not is_h and d.weekday() != 6:
            hasegawa_work_count += work_flag

    print(f"  -> 11/3週の診療日出勤日数: {hasegawa_work_count}日 / 診療日5日中 (目標: 最大4日)")
    assert hasegawa_work_count <= 4, f"長谷川さんの週出勤日数が4日以下になっていません: {hasegawa_work_count}"

    sec_2f_back_staffs = [s['name'] for s in staff_list if s['section'] == '2階裏']
    print("\n【11/3週 2階裏の各日出勤人数】")
    for d in dates_11_week:
        if d in holidays_jp_dict:
            continue
        cnt = sum(best_x_11.get((n, d), 0) for n in sec_2f_back_staffs)
        hasegawa_status = "出勤" if best_x_11.get(('長谷川', d), 0) == 1 else "【振替公休】"
        print(f"  {d} ({wd_names[d.weekday()]}): 2階裏出勤 {cnt}名 | 長谷川: {hasegawa_status}")

    print("\n=== 全ての検証を正常に通過しました！ ===")

if __name__ == "__main__":
    test_holiday_substitute()
