import os
import shutil
import zipfile

def package():
    dist_dir = os.path.join("dist", "シフト生成_v1.3")
    assert os.path.isdir(dist_dir), f"Dist dir not found: {dist_dir}"
    
    # 1. 勤務条件表_ひな形.xlsx の配置 (最新の設定11項目入り)
    src_excel = "ope勤務条件表_2026_4sheets.xlsx"
    dst_excel = os.path.join(dist_dir, "勤務条件表_ひな形.xlsx")
    shutil.copy2(src_excel, dst_excel)
    print(f"ひな形Excelをコピーしました: {dst_excel}")
    
    # 2. 使い方.txt の作成
    readme_path = os.path.join(dist_dir, "使い方.txt")
    readme_content = """【勤務シフト自動生成システム v1.3 使い方ガイド】

1. 起動方法
   「シフト生成_v1.3.exe」をダブルクリックして起動します。
   （Pythonのインストールは不要です）

2. 条件表の準備
   同じフォルダ内にある「勤務条件表_ひな形.xlsx」を編集、
   またはご自身の「*勤務条件表*.xlsx」をこのフォルダ内に配置してください。

3. 作成対象月の選択 (v1.3新機能)
   手術スケジュールシートに記載のあるすべての月が自動的に一覧表示されます。
   生成したい月にチェックを入れてください（「全選択」「全解除」ボタンも利用可能）。

4. 祝日週の固定公休振替 (v1.3追加機能)
   ・週5日勤務で固定公休（AND型）のスタッフについて、診療日が5日以上ある週に固定公休が祝日と重複した場合、その週内の平日に自動的に振替公休を設定します。
   ・振替公休日はセクションごとの勤務人数平準化ペナルティと連動し、セクション内で最も余剰のある日に自動配置されます。
   ・「設定」シートの「祝日週_固定公休振替」で「有効」「無効」を切り替えることも可能です。

5. 制約レベル（10段階の優先度指定）
   「シフト制約条件」シートの「制約レベル」列で、以下の優先度を自由に指定できます：
   ・HARD1: 最重要HARD（手術日専門スタッフ配置など、絶対に欠かせない制約）
   ・HARD2 (または HARD): 通常HARD（グループ出勤必須など）
   ・HARD3: 緩和可HARD（基本厳守だが、HARD同士衝突時に優先して不足を許容）
   ・HIGH1: 最優先HIGH（手術日の出勤人数指定など、HARD直下の強力な要件）
   ・HIGH2 (または HIGH): 通常HIGH（標準的な推奨人数指定など）
   ・HIGH3: 平準化優先HIGH（平日の出勤体制など、平準化よりは優先したい要件）
   ・SOFT1 / SOFT2 / SOFT3: 通常の希望（SOFT1=強め、SOFT2=普通、SOFT3=軽め）
   ・LOW: 努力目標（できればこうしたいレベル）

6. 各シートの役割
   ・スタッフマスタ: スタッフの氏名、所属、雇用形態、公休曜日、有休日などを管理
   ・シフト制約条件: 医療・業務ルール（上記10段階レベル対応）
   ・手術スケジュール: 手術日程（記載のある月がすべてGUIの対象月として選択可能）
   ・前月シフト実績: 前月の確定シフト表（月またぎ連動用）
   ・設定: 2階裏人数、連勤上限、各種ペナルティ、祝日週振替公休などの微調整パラメータ（11項目）

7. シフトの生成
   画面で「勤務条件表」が認識されていることを確認し、
   「シフト自動生成を開始」ボタンをクリックします。
   生成が完了すると、自動的に「YY年MM-MM月シフト.xlsx」が出力されます。
"""
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(readme_content)
    print(f"使い方.txtを作成しました: {readme_path}")
    
    # 3. ZIPアーカイブの作成
    zip_name = "シフト生成_v1.3.zip"
    if os.path.exists(zip_name):
        os.remove(zip_name)
        
    print(f"ZIPアーカイブ '{zip_name}' を作成中...")
    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(dist_dir):
            for file in files:
                abs_path = os.path.join(root, file)
                rel_path = os.path.relpath(abs_path, dist_dir)
                zipf.write(abs_path, rel_path)
                
    zip_size_mb = os.path.getsize(zip_name) / (1024 * 1024)
    print(f"★ 配布パッケージ作成完了: {zip_name} ({zip_size_mb:.1f} MB)")

if __name__ == "__main__":
    package()
