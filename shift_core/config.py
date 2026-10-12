"""
設定および定数定義モジュール
"""
from typing import Dict, List, Any

# 制約レベルごとのペナルティ重み
LEVEL_WEIGHT_MAP: Dict[str, int] = {
    'HARD': 15000000,
    'HIGH': 12000000,
    'SOFT': 100000,
    'LOW': 20000
}

# システム最上位ペナルティ
WEIGHT_FIXED_OFF_SHORTAGE = 30000000  # 固定公休の不足防止（最重要防衛）
WEIGHT_2F_BACK_MIN_SLACK = 20000000    # 2階裏 人員不足スラック
WEIGHT_GENERAL_SEC_SLACK = 10000000    # 一般セクション 人員不足・過剰スラック
WEIGHT_NURSE_DEFICIT_SLACK = 15000000  # 2階裏 看護師不在スラック
WEIGHT_2F_BACK_5_PENALTY = 2000000     # 2階裏 5名出勤ペナルティ
WEIGHT_SEC_BALANCE_SQ = 500000         # セクション人数平準化（自乗）
WEIGHT_SEC_BALANCE_ABS = 500000        # セクション人数平準化（絶対値）
WEIGHT_SHIFT_CHANGE = 300000           # 確定シフト変更ペナルティ
WEIGHT_MAIN_OFF_MISMATCH = 300000      # 主公休不一致ペナルティ
WEIGHT_2F_BACK_3_PENALTY = 300000      # 2階裏 3名ペナルティ
WEIGHT_NURSE_LT2_PENALTY = 200000      # 看護師2名未満ペナルティ
WEIGHT_SPECIAL_EXTRA_WORK = 150000     # 週4(週5可)の超過勤務ペナルティ
WEIGHT_OFF_3CONSEC = 100000            # 3連休ペナルティ
WEIGHT_ROLLING_WINDOW_DIFF = 80000     # 7日間ローリング週休平準化ペナルティ
WEIGHT_PREV_OFF_REPEAT = 50000         # 前月主公休継続ペナルティ
WEIGHT_TAIL_OFF_SUM = 30000            # 月末公休偏りペナルティ

# セクション定義
SECTION_DEFINITIONS: List[Dict[str, Any]] = [
    {"id": "2階表", "staff_fill": "1F4E78", "cnt_fill": "1B365D", "blank": True, "cnt_width": 9.875},
    {"id": "2階裏", "staff_fill": "2F5597", "cnt_fill": "1B365D", "blank": True, "cnt_width": 9.875},
    {"id": "1階", "staff_fill": "203764", "cnt_fill": "1B365D", "blank": False, "cnt_width": 8.25}
]

# ソルバー設定
SOLVER_MAX_TIME_SECONDS = 120.0
SOLVER_NUM_SEARCH_WORKERS = 8
