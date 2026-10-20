import sys
import datetime
from shift_core import ShiftGeneratorService
from shift_core.parser import load_and_parse_data
from shift_core.solver import solve_monthly_shift

def run_v15_tests():
    print("=== 勤務シフト自動生成システム v1.5 新条件テスト ===")
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
    
    print(f"動的設定 holiday_week_off_rule: {config.holiday_week_off_rule}")
    assert config.holiday_week_off_rule is True, "holiday_week_off_rule should be True"

    # --- テスト1: 2026年10月 (10/12 スポーツの日) ---
    print("\n==========================================")
    print("【テスト1】2026年10月の検証 (10/12週)")
    print("==========================================")
    monthly_data_10 = {
        'staff_list': staff_list,
        'df_rules': df_rules,
        'ope_dict': ope_dict_all.get((2026, 10), {}),
        'holidays_jp_dict': holidays_jp_dict,
        'target_file': filepath,
        'config': config
    }
    prev_data = {'tail_shifts': init_tail_shifts, 'main_offs': init_main_offs}
    best_x_10, main_offs_10, target_days_10, tail_10 = solve_monthly_shift(
        2026, 10, monthly_data_10, prev_data, stage=1
    )
    if best_x_10 is None:
        print("Stage 1 失敗 -> Stage 2 試行")
        best_x_10, main_offs_10, target_days_10, tail_10 = solve_monthly_shift(
            2026, 10, monthly_data_10, prev_data, stage=2
        )
    assert best_x_10 is not None, "2026年10月のシフト求解に失敗しました"

    # 10/12週の診療日 (10/13火〜10/17土: 5日間)
    work_dates_10 = [datetime.date(2026, 10, d) for d in range(13, 18)]
    wd_names = ["月", "火", "水", "木", "金", "土", "日"]
    
    print("【10/12週 コアスタッフ全員の診療日出勤日数一覧 (最大4日 / 最低1日公休)】")
    max_days_exceeded = []
    core_staffs = [s for s in staff_list if s.get('is_core') and s.get('pattern') != '除外']
    for s in core_staffs:
        name = s['name']
        sec = s['section']
        pat = s.get('pattern', '')
        b_type = s.get('basic_off_type', '')
        b_days = [wd_names[w] for w in s.get('basic_off_days', [])]
        cnt = sum(best_x_10.get((name, d), 0) for d in work_dates_10)
        off_days = [wd_names[d.weekday()] for d in work_dates_10 if best_x_10.get((name, d), 0) == 0]
        status = "OK (週4日以下)" if cnt <= 4 else "NG (5連勤)"
        print(f"  [{sec}] {name:6s} ({pat}, 公休:{','.join(b_days) if b_days else 'なし'}): 出勤 {cnt}日 / 5日中 | 振替公休: {','.join(off_days)} -> {status}")
        if cnt > 4:
            max_days_exceeded.append((name, cnt))
            
    assert len(max_days_exceeded) == 0, f"5連勤のスタッフが存在します: {max_days_exceeded}"
    print("=> 10/12週: コアスタッフ全員が週4日以下（公休最低1日以上）を完全に達成！")

    # セクション毎の人数推移
    print("\n【10/12週 セクション別出勤人数推移】")
    for d in [datetime.date(2026, 10, 12)] + work_dates_10:
        is_h = d in holidays_jp_dict
        h_str = f"({holidays_jp_dict[d]})" if is_h else ""
        cnt_2f_f = sum(best_x_10.get((s['name'], d), 0) for s in staff_list if s['section'] == '2階表')
        cnt_2f_b = sum(best_x_10.get((s['name'], d), 0) for s in staff_list if s['section'] == '2階裏')
        cnt_1f = sum(best_x_10.get((s['name'], d), 0) for s in staff_list if s['section'] == '1階')
        print(f"  {d} ({wd_names[d.weekday()]}){h_str:10s}: 2階表 {cnt_2f_f:2d}名 | 2階裏 {cnt_2f_b:2d}名 | 1階 {cnt_1f:2d}名")

    # --- テスト2: 2026年11月 (11/3 文化の日、11/23 勤労感謝の日) ---
    print("\n==========================================")
    print("【テスト2】2026年11月の検証 (祝日2回)")
    print("==========================================")
    monthly_data_11 = {
        'staff_list': staff_list,
        'df_rules': df_rules,
        'ope_dict': ope_dict_all.get((2026, 11), {}),
        'holidays_jp_dict': holidays_jp_dict,
        'target_file': filepath,
        'config': config
    }
    prev_data_11 = {'tail_shifts': tail_10, 'main_offs': main_offs_10}
    best_x_11, main_offs_11, target_days_11, tail_11 = solve_monthly_shift(
        2026, 11, monthly_data_11, prev_data_11, stage=1
    )
    if best_x_11 is None:
        print("Stage 1 失敗 -> Stage 2 試行")
        best_x_11, main_offs_11, target_days_11, tail_11 = solve_monthly_shift(
            2026, 11, monthly_data_11, prev_data_11, stage=2
        )
    assert best_x_11 is not None, "2026年11月のシフト求解に失敗しました"

    # 11/3週 (11/2月, 11/4水〜11/7土: 5日間)
    work_dates_11_1 = [datetime.date(2026, 11, 2)] + [datetime.date(2026, 11, d) for d in range(4, 8)]
    print("\n【11/3週 コアスタッフ出勤日数 (文化の日週)】")
    max_days_exceeded_11_1 = []
    for s in core_staffs:
        name = s['name']
        cnt = sum(best_x_11.get((name, d), 0) for d in work_dates_11_1)
        if cnt > 4:
            max_days_exceeded_11_1.append((name, cnt))
    assert len(max_days_exceeded_11_1) == 0, f"11/3週に5連勤のスタッフが存在します: {max_days_exceeded_11_1}"
    print(f"=> 11/3週: コアスタッフ{len(core_staffs)}名全員が週4日以下勤務を達成！")

    # 11/23週 (11/24火〜11/28土: 5日間)
    work_dates_11_2 = [datetime.date(2026, 11, d) for d in range(24, 29)]
    print("\n【11/23週 コアスタッフ出勤日数 (勤労感謝の日週)】")
    max_days_exceeded_11_2 = []
    for s in core_staffs:
        name = s['name']
        cnt = sum(best_x_11.get((name, d), 0) for d in work_dates_11_2)
        if cnt > 4:
            max_days_exceeded_11_2.append((name, cnt))
    assert len(max_days_exceeded_11_2) == 0, f"11/23週に5連勤のスタッフが存在します: {max_days_exceeded_11_2}"
    print(f"=> 11/23週: コアスタッフ{len(core_staffs)}名全員が週4日以下勤務を達成！")

    # --- テスト3: 無効化テスト ---
    print("\n==========================================")
    print("【テスト3】設定を『無効』にした場合の挙動検証")
    print("==========================================")
    config.holiday_week_off_rule = False
    best_x_disabled, _, _, _ = solve_monthly_shift(2026, 10, monthly_data_10, prev_data, stage=1)
    # 無効時は固定公休のないスタッフなどが5連勤になる場合があるか
    fives = []
    for s in core_staffs:
        name = s['name']
        cnt = sum(best_x_disabled.get((name, d), 0) for d in work_dates_10)
        if cnt == 5:
            fives.append(name)
    print(f"  無効化時に10/12週で5連勤となったスタッフ: {fives}")
    assert len(fives) > 0, "無効化時に5連勤が発生していません（制約のオン/オフが効いていない可能性）"
    print("=> 設定のオン/オフ切り替えが正常に機能していることを確認！")

    print("\n==========================================")
    print("★ すべてのテスト項目を完璧に通過しました！ ★")
    print("==========================================")

if __name__ == "__main__":
    run_v15_tests()
