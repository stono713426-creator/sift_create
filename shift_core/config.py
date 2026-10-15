"""
設定および定数定義モジュール
動的設定データクラス (ShiftConfig) とデフォルト定数
"""
from dataclasses import dataclass
from typing import Dict, List, Any

# 制約レベルごとのペナルティ重み (10段階体系)
LEVEL_WEIGHT_MAP: Dict[str, int] = {
    # HARD系 (Stage 1厳格探索、Stage 2スラック緩和順位)
    'HARD1': 20000000,
    'HARD2': 16000000,
    'HARD3': 14000000,
    'HARD':  16000000,  # 従来のHARDはHARD2互換

    # HIGH系 (高優先業務要件)
    'HIGH1': 12000000,
    'HIGH2': 9000000,
    'HIGH3': 6000000,
    'HIGH':  9000000,   # 従来のHIGHはHIGH2互換

    # SOFT系 (通常希望・平準化バランス)
    'SOFT1': 400000,
    'SOFT2': 100000,
    'SOFT3': 50000,
    'SOFT':  100000,   # 従来のSOFTはSOFT2互換

    # LOW系 (努力目標)
    'LOW':   20000,
    'LOW1':  20000,
}

# 推奨人数 (val_rec) 不足ペナルティ重み
REC_WEIGHT_MAP: Dict[str, int] = {
    'HARD1': 6000000,
    'HARD2': 5000000,
    'HARD3': 4000000,
    'HARD':  5000000,

    'HIGH1': 6000000,
    'HIGH2': 4000000,
    'HIGH3': 2000000,
    'HIGH':  4000000,

    'SOFT1': 400000,
    'SOFT2': 200000,
    'SOFT3': 100000,
    'SOFT':  200000,

    'LOW':   50000,
    'LOW1':  50000,
}

# システム最上位ペナルティ
WEIGHT_FIXED_OFF_SHORTAGE = 30000000  # 固定公休の不足防止（最重要防衛）
WEIGHT_2F_BACK_MIN_SLACK = 20000000    # 2階裏 人員不足スラック
WEIGHT_GENERAL_SEC_SLACK = 10000000    # 一般セクション 人員不足・過剰スラック
WEIGHT_NURSE_DEFICIT_SLACK = 15000000  # 2階裏 看護師不在スラック
WEIGHT_2F_BACK_5_PENALTY = 2000000     # 2階裏 5名出勤ペナルティ
WEIGHT_SEC_BALANCE_SQ = 500000         # セクション人数平準化（自乗）
WEIGHT_SEC_BALANCE_ABS = 500000        # セクション人数平準化（絶対値）
WEIGHT_MAIN_OFF_MISMATCH = 300000      # 主公休不一致ペナルティ
WEIGHT_2F_BACK_3_PENALTY = 300000      # 2階裏 3名ペナルティ
WEIGHT_NURSE_LT2_PENALTY = 200000      # 看護師2名未満ペナルティ
WEIGHT_SPECIAL_EXTRA_WORK = 150000     # 週4(週5可)の超過勤務ペナルティ
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

# レベル文字列マッピング
LEVEL_STRENGTH_MAP = {
    "強": 300000,
    "中": 100000,
    "弱": 30000,
    "なし": 0
}

EXISTING_SHIFT_WEIGHT_MAP = {
    "強": 1000000,
    "中": 300000,
    "弱": 50000
}

@dataclass
class ShiftConfig:
    """勤務条件表の『設定』シートから読み込まれる動的パラメータ（未指定時は初期値）"""
    # 1. 2階裏 基本最低人数
    sec_2f_back_base_cap: int = 3
    # 2. 2階裏 看護師最低人数 (HARD)
    nurse_min_count: int = 1
    # 3. 2階裏 看護師推奨人数 (不足時ソフトペナルティ)
    nurse_rec_count: int = 2
    # 4. 連勤上限日数 (5なら6連勤禁止)
    max_consecutive_work_days: int = 5
    # 5. 固定公休の削減許容日数 (目標-N日)
    fixed_off_shortage_allowance: int = 2
    # 6. 3連休抑制レベル ("強", "中", "弱", "なし")
    consecutive_off_level: str = "中"
    # 7. セクション人数ブレ幅 (平均±N名)
    sec_capacity_tolerance: int = 1
    # 8. 週休配分目標 (7日あたりN日)
    weekly_off_target: int = 2
    # 9. 既存シフト維持レベル ("強", "中", "弱")
    existing_shift_weight_level: str = "中"
    # 10. 2階裏 警告表示人数 (N名以下で⚠️)
    warn_2f_back_count: int = 2

    @property
    def weight_off_3consec(self) -> int:
        return LEVEL_STRENGTH_MAP.get(self.consecutive_off_level, 100000)

    @property
    def weight_shift_change(self) -> int:
        return EXISTING_SHIFT_WEIGHT_MAP.get(self.existing_shift_weight_level, 300000)
