"""
勤務シフト自動生成システム CLIエントリーポイント
"""
import sys
from shift_core import ShiftGeneratorService

def main():
    print("=== 勤務シフト自動生成システム (Core Service) ===")
    service = ShiftGeneratorService()
    
    input_file = sys.argv[1] if len(sys.argv) > 1 else None
    
    result = service.generate(
        input_file=input_file,
        on_progress=lambda p, msg: None  # CLIでは標準出力でログ表示
    )
    
    if result.success:
        print("\n【処理結果サマリー】")
        for mr in result.month_results:
            print(f"- {mr.year}年{mr.month:02d}月: {'成功' if mr.success else '失敗'} ({mr.message})")
        print(f"出力ファイル: {result.output_filepath}")
        sys.exit(0)
    else:
        print(f"\n【エラー】シフト生成に失敗しました: {result.error_message}")
        sys.exit(1)

if __name__ == "__main__":
    main()
