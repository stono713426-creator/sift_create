import datetime
from shift_core import ShiftGeneratorService
from shift_core.parser import load_and_parse_data
from shift_core.solver import solve_monthly_shift

def run_verification():
    print("=== shift_core v1.6 平準化改善テスト ===")
    filepath = 'ope勤務条件表_2026_4sheets.xlsx'
    
    staff_list, df_rules, ope_dict_all, auto_target_months, holidays_jp_dict, init_tail, init_mo, cfg = load_and_parse_data(filepath)
    
    print(f"動的設定 weekly_off_balance_rule: {cfg.weekly_off_balance_rule}")
    print(f"動的設定 sec_2f_back_weekly_4_min: {cfg.sec_2f_back_weekly_4_min}")
    print(f"動的設定 holiday_week_off_rule: {cfg.holiday_week_off_rule}")
    
    # --- 10月テスト ---
    print("\n------------------------------------------")
    print("【2026年10月 検証】")
    print("------------------------------------------")
    m_data_10 = {
        'staff_list': staff_list,
        'df_rules': df_rules,
        'ope_dict': ope_dict_all.get((2026, 10), {}),
        'holidays_jp_dict': holidays_jp_dict,
        'target_file': filepath,
        'config': cfg
    }
    prev_data = {'tail_shifts': init_tail, 'main_offs': init_mo}
    best_x, main_offs, target_days, tail = solve_monthly_shift(2026, 10, m_data_10, prev_data, stage=1)
    if best_x is None:
        print("Stage 1 失敗 -> Stage 2 試行")
        best_x, main_offs, target_days, tail = solve_monthly_shift(2026, 10, m_data_10, prev_data, stage=2)
    assert best_x is not None, "10月のシフト求解に失敗しました"
    
    weeks = [
        ('第1週 (10/01-10/03)', [datetime.date(2026, 10, d) for d in range(1, 4)]),
        ('第2週 (10/05-10/10)', [datetime.date(2026, 10, d) for d in range(5, 11)]),
        ('第3週 (10/12-10/17)', [datetime.date(2026, 10, d) for d in range(12, 18)]),
        ('第4週 (10/19-10/24)', [datetime.date(2026, 10, d) for d in range(19, 25)]),
        ('第5週 (10/26-10/31)', [datetime.date(2026, 10, d) for d in range(26, 32)]),
    ]
    
    s_2b = [s['name'] for s in staff_list if s['section'] == '2階裏']
    for w_name, w_dates in weeks:
        work_dates = [d for d in w_dates if d.weekday() != 6 and d not in holidays_jp_dict]
        daily_cnts = [sum(best_x[(n, d)] for n in s_2b) for d in work_dates]
        nurses_off = {n: sum(1 for d in work_dates if best_x[(n, d)] == 0) for n in ['堀', '江原', '石橋']}
        print(f"  {w_name} (診療{len(work_dates)}日): 2階裏日別={daily_cnts} | 看護師公休={nurses_off}")
    
    # 第5週の検証
    w5_work_dates = [datetime.date(2026, 10, d) for d in range(26, 32)]
    w5_cnts = [sum(best_x[(n, d)] for n in s_2b) for d in w5_work_dates]
    w5_count_4 = sum(1 for c in w5_cnts if c >= 4)
    print(f"\n=> 第5週の4名出勤日数: {w5_count_4}日")
    assert w5_count_4 >= 2, f"第5週に4名出勤の日が{w5_count_4}日しかありません（期待値2日以上）"
    print("=> 第5週が全日3名体制になる問題が正常に解決されました！")

    # 看護師の公休日数チェック
    for n in ['堀', '江原', '石橋']:
        w5_off = sum(1 for d in w5_work_dates if best_x[(n, d)] == 0)
        print(f"  {n}: 第5週公休 {w5_off}日")
        assert w5_off <= 3, f"{n}の第5週公休が{w5_off}日と過多です"
    print("=> 看護師の公休日数も正常に平準化されています！")

    # --- 11月・12月テスト ---
    print("\n------------------------------------------")
    print("【2026年11月・12月 検証】")
    print("------------------------------------------")
    m_data_11 = {
        'staff_list': staff_list,
        'df_rules': df_rules,
        'ope_dict': ope_dict_all.get((2026, 11), {}),
        'holidays_jp_dict': holidays_jp_dict,
        'target_file': filepath,
        'config': cfg
    }
    x_11, mo_11, tg_11, tail_11 = solve_monthly_shift(2026, 11, m_data_11, {'tail_shifts': tail, 'main_offs': main_offs}, stage=1)
    if x_11 is None:
        x_11, mo_11, tg_11, tail_11 = solve_monthly_shift(2026, 11, m_data_11, {'tail_shifts': tail, 'main_offs': main_offs}, stage=2)
    assert x_11 is not None, "11月のシフト求解に失敗しました"
    print("=> 11月: 求解成功！")

    m_data_12 = {
        'staff_list': staff_list,
        'df_rules': df_rules,
        'ope_dict': ope_dict_all.get((2026, 12), {}),
        'holidays_jp_dict': holidays_jp_dict,
        'target_file': filepath,
        'config': cfg
    }
    x_12, mo_12, tg_12, tail_12 = solve_monthly_shift(2026, 12, m_data_12, {'tail_shifts': tail_11, 'main_offs': mo_11}, stage=1)
    if x_12 is None:
        x_12, mo_12, tg_12, tail_12 = solve_monthly_shift(2026, 12, m_data_12, {'tail_shifts': tail_11, 'main_offs': mo_11}, stage=2)
    assert x_12 is not None, "12月のシフト求解に失敗しました"
    print("=> 12月: 求解成功！")

    print("\n全テストに合格しました！")

if __name__ == '__main__':
    run_verification()
