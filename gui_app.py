"""
勤務シフト自動生成システム GUIアプリケーション (CustomTkinter)
モダンなUI、非同期計算実行、進捗バー、リアルタイムログ、Excelワンクリック起動。
"""
import os
import glob
import threading
import queue
from typing import List, Tuple, Optional
import customtkinter as ctk
from tkinter import filedialog, messagebox

from shift_core import ShiftGeneratorService, inspect_excel_file, GenerationResult

# アプリケーションテーマ設定
ctk.set_appearance_mode("System")  # "Dark", "Light", "System"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"

class ShiftGeneratorApp(ctk.CTk):
    """シフト自動生成 デスクトップGUI"""

    def __init__(self):
        super().__init__()

        self.title("勤務シフト自動生成システム")
        self.geometry("880x720")
        self.minsize(780, 600)

        # 状態変数
        self.log_queue = queue.Queue()
        self.is_running = False
        self.month_checkboxes = {}
        self.current_output_file = None
        self.inspected_info = None

        self._create_widgets()
        self._auto_detect_file()

        # ログキューの定期監視 (UIスレッドセーフ)
        self.after(100, self._process_log_queue)

    def _create_widgets(self):
        # メインコンテナ
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # 1. ヘッダー
        header_frame = ctk.CTkFrame(self, corner_radius=10)
        header_frame.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="ew")
        header_frame.grid_columnconfigure(0, weight=1)

        title_lbl = ctk.CTkLabel(
            header_frame,
            text="🏥 勤務シフト自動生成システム",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_lbl.grid(row=0, column=0, padx=15, pady=(10, 2), sticky="w")

        sub_lbl = ctk.CTkLabel(
            header_frame,
            text="医療機関特有の制約・手術スケジュール・月またぎ実績を考慮した二段階最適化エンジン",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        sub_lbl.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="w")

        # 2. ファイル選択セクション
        file_frame = ctk.CTkFrame(self, corner_radius=10)
        file_frame.grid(row=1, column=0, padx=15, pady=5, sticky="ew")
        file_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(file_frame, text="勤務条件表 (Excel):", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, padx=(15, 5), pady=12, sticky="w"
        )

        self.file_entry = ctk.CTkEntry(file_frame, placeholder_text="Excelファイルパス...")
        self.file_entry.grid(row=0, column=1, padx=5, pady=12, sticky="ew")

        browse_btn = ctk.CTkButton(file_frame, text="参照...", width=80, command=self._browse_file)
        browse_btn.grid(row=0, column=2, padx=5, pady=12)

        auto_btn = ctk.CTkButton(file_frame, text="自動検出", width=80, fg_color="#3A3A3A", command=self._auto_detect_file)
        auto_btn.grid(row=0, column=3, padx=(5, 15), pady=12)

        # ファイル情報バッジエリア
        self.badge_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        self.badge_frame.grid(row=1, column=0, columnspan=4, padx=15, pady=(0, 10), sticky="ew")
        self.info_lbl = ctk.CTkLabel(self.badge_frame, text="ファイル未読み込み", text_color="gray", font=ctk.CTkFont(size=12))
        self.info_lbl.pack(anchor="w")

        # 3. 対象月選択セクション
        self.month_frame = ctk.CTkFrame(self, corner_radius=10)
        self.month_frame.grid(row=2, column=0, padx=15, pady=5, sticky="ew")
        self.month_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.month_frame, text="作成対象月:", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, padx=(15, 10), pady=10, sticky="w"
        )

        self.month_cb_container = ctk.CTkFrame(self.month_frame, fg_color="transparent")
        self.month_cb_container.grid(row=0, column=1, padx=5, pady=10, sticky="w")

        # 4. 実行コントロール ＆ 進捗バー
        ctrl_frame = ctk.CTkFrame(self, corner_radius=10)
        ctrl_frame.grid(row=3, column=0, padx=15, pady=5, sticky="nsew")
        ctrl_frame.grid_columnconfigure(0, weight=1)
        ctrl_frame.grid_rowconfigure(3, weight=1)

        # 実行ボタン
        btn_box = ctk.CTkFrame(ctrl_frame, fg_color="transparent")
        btn_box.grid(row=0, column=0, padx=15, pady=(12, 5), sticky="ew")
        btn_box.grid_columnconfigure(0, weight=1)

        self.run_btn = ctk.CTkButton(
            btn_box,
            text="🚀 シフト自動生成を開始",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=42,
            command=self._start_generation
        )
        self.run_btn.grid(row=0, column=0, sticky="ew")

        # プログレスバー & ステータス
        prog_box = ctk.CTkFrame(ctrl_frame, fg_color="transparent")
        prog_box.grid(row=1, column=0, padx=15, pady=5, sticky="ew")
        prog_box.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(prog_box)
        self.progress_bar.grid(row=0, column=0, sticky="ew", pady=(2, 4))
        self.progress_bar.set(0.0)

        self.status_lbl = ctk.CTkLabel(prog_box, text="待機中 - 条件ファイルを選択して開始してください", font=ctk.CTkFont(size=12))
        self.status_lbl.grid(row=1, column=0, sticky="w")

        # リアルタイムログ
        log_header_box = ctk.CTkFrame(ctrl_frame, fg_color="transparent")
        log_header_box.grid(row=2, column=0, padx=15, pady=(5, 2), sticky="ew")
        log_header_box.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(log_header_box, text="実行ログ:", font=ctk.CTkFont(size=12, weight="bold")).grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkButton(
            log_header_box,
            text="ログ消去",
            width=65,
            height=22,
            font=ctk.CTkFont(size=11),
            fg_color="#404040",
            command=self._clear_log
        ).grid(row=0, column=1, sticky="e")

        self.log_textbox = ctk.CTkTextbox(ctrl_frame, font=ctk.CTkFont(family="Consolas", size=11))
        self.log_textbox.grid(row=3, column=0, padx=15, pady=(2, 12), sticky="nsew")

        # 5. フッター (結果アクション)
        footer_frame = ctk.CTkFrame(self, corner_radius=10)
        footer_frame.grid(row=4, column=0, padx=15, pady=(5, 15), sticky="ew")
        footer_frame.grid_columnconfigure(1, weight=1)

        self.res_lbl = ctk.CTkLabel(footer_frame, text="出力: -", font=ctk.CTkFont(size=12))
        self.res_lbl.grid(row=0, column=0, padx=15, pady=10, sticky="w")

        self.open_excel_btn = ctk.CTkButton(
            footer_frame,
            text="📊 Excelで開く",
            width=120,
            fg_color="#2E7D32",
            state="disabled",
            command=self._open_result_excel
        )
        self.open_excel_btn.grid(row=0, column=2, padx=5, pady=10)

        self.open_dir_btn = ctk.CTkButton(
            footer_frame,
            text="📁 フォルダを開く",
            width=120,
            fg_color="#455A64",
            state="disabled",
            command=self._open_result_dir
        )
        self.open_dir_btn.grid(row=0, column=3, padx=(5, 15), pady=10)

    # ==========================
    # ファイル・メタ情報ハンドラ
    # ==========================
    def _browse_file(self):
        f = filedialog.askopenfilename(
            title="勤務条件表Excelを選択",
            filetypes=[("Excel Files", "*.xlsx *.xlsm")]
        )
        if f:
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, f)
            self._on_file_changed(f)

    def _auto_detect_file(self):
        files = glob.glob('*勤務条件表*.xlsx') + glob.glob('ope勤務条件表*.xlsx')
        files = list(set(files))
        if files:
            latest = max(files, key=os.path.getmtime)
            full_path = os.path.abspath(latest)
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, full_path)
            self._on_file_changed(full_path)
        else:
            self.info_lbl.configure(text="同フォルダに対象のExcelが見つかりません。『参照...』から選択してください。")

    def _on_file_changed(self, filepath: str):
        if not os.path.exists(filepath):
            return
        
        info = inspect_excel_file(filepath)
        self.inspected_info = info
        
        # バッジ情報の更新
        badge_parts = []
        badge_parts.append(f"スタッフ数: {info['staff_count']}名")
        badge_parts.append(f"制約ルール: {info['rules_count']}件")
        if info['has_prev_month']:
            badge_parts.append("前月実績シート: あり (連動有効)")
        else:
            badge_parts.append("前月実績シート: なし (単月モード)")

        if info.get('has_settings'):
            badge_parts.append("設定シート: あり (動的パラメータ適用)")
        else:
            badge_parts.append("設定シート: なし (標準初期値)")

        self.info_lbl.configure(
            text=" | ".join(badge_parts),
            text_color="#64B5F6" if info['has_prev_month'] or info.get('has_settings') else "gray"
        )

        # 対象月チェックボックスの再構築
        for widget in self.month_cb_container.winfo_children():
            widget.destroy()
        self.month_checkboxes.clear()

        target_months = info.get('target_months', [])
        if target_months:
            for (y, m) in target_months:
                var = ctk.BooleanVar(value=True)
                cb = ctk.CTkCheckBox(
                    self.month_cb_container,
                    text=f"{y}年{m:02d}月",
                    variable=var,
                    font=ctk.CTkFont(size=12, weight="bold")
                )
                cb.pack(side="left", padx=8)
                self.month_checkboxes[(y, m)] = var
        else:
            lbl = ctk.CTkLabel(self.month_cb_container, text="※ 作成対象フラグ(○)の付いた月が見つかりません", text_color="orange")
            lbl.pack(side="left")

    # ==========================
    # 実行制御
    # ==========================
    def _start_generation(self):
        if self.is_running:
            return

        input_path = self.file_entry.get().strip()
        if not os.path.exists(input_path):
            messagebox.showerror("エラー", "指定された勤務条件表Excelファイルが存在しません。")
            return

        # 選択された対象月
        selected_months = [m for m, var in self.month_checkboxes.items() if var.get()]
        if not selected_months:
            messagebox.showwarning("警告", "生成対象月が1つも選択されていません。")
            return

        self.is_running = True
        self.run_btn.configure(state="disabled", text="⏳ シフト生成中...")
        self.open_excel_btn.configure(state="disabled")
        self.open_dir_btn.configure(state="disabled")
        self.progress_bar.set(0.0)
        self.status_lbl.configure(text="処理開始中...")
        self._append_log("=" * 55 + "\nシフト生成処理を開始します...\n" + "=" * 55)

        # バックグラウンドスレッドで実行
        worker = threading.Thread(
            target=self._run_backend_task,
            args=(input_path, selected_months),
            daemon=True
        )
        worker.start()

    def _run_backend_task(self, input_path: str, selected_months: List[Tuple[int, int]]):
        def log_cb(msg: str):
            self.log_queue.put(("log", msg))

        def prog_cb(ratio: float, text: str):
            self.log_queue.put(("progress", (ratio, text)))

        service = ShiftGeneratorService(logger=log_cb)

        try:
            res: GenerationResult = service.generate(
                input_file=input_path,
                target_months=selected_months,
                on_progress=prog_cb
            )
            self.log_queue.put(("finish", res))
        except Exception as e:
            self.log_queue.put(("error", str(e)))

    def _process_log_queue(self):
        """UIスレッドで定期的にログ・進捗を反映"""
        while not self.log_queue.empty():
            msg_type, data = self.log_queue.get_nowait()

            if msg_type == "log":
                self._append_log(data)

            elif msg_type == "progress":
                ratio, text = data
                self.progress_bar.set(ratio)
                self.status_lbl.configure(text=text)

            elif msg_type == "finish":
                res: GenerationResult = data
                self.is_running = False
                self.run_btn.configure(state="normal", text="🚀 シフト自動生成を開始")

                if res.success:
                    self.current_output_file = os.path.abspath(res.output_filepath)
                    self.res_lbl.configure(
                        text=f"出力: {os.path.basename(self.current_output_file)}",
                        text_color="#4CAF50"
                    )
                    self.open_excel_btn.configure(state="normal")
                    self.open_dir_btn.configure(state="normal")
                    self.status_lbl.configure(text="🎉 すべての対象月のシフト生成が完了しました！")
                    messagebox.showinfo(
                        "完了",
                        f"シフト表の作成が正常に完了しました！\n\n出力ファイル:\n{self.current_output_file}"
                    )
                else:
                    self.status_lbl.configure(text=f"エラー: {res.error_message}")
                    messagebox.showerror("生成失敗", f"シフト生成でエラーが発生しました:\n{res.error_message}")

            elif msg_type == "error":
                self.is_running = False
                self.run_btn.configure(state="normal", text="🚀 シフト自動生成を開始")
                self.status_lbl.configure(text=f"例外エラー: {data}")
                messagebox.showerror("システムエラー", f"予期せぬエラーが発生しました:\n{data}")

        self.after(100, self._process_log_queue)

    def _append_log(self, text: str):
        self.log_textbox.insert("end", text + "\n")
        self.log_textbox.see("end")

    def _clear_log(self):
        self.log_textbox.delete("1.0", "end")

    # ==========================
    # ファイル・フォルダ起動
    # ==========================
    def _open_result_excel(self):
        if self.current_output_file and os.path.exists(self.current_output_file):
            try:
                os.startfile(self.current_output_file)
            except Exception as e:
                messagebox.showerror("起動エラー", f"Excelの起動に失敗しました: {e}")

    def _open_result_dir(self):
        if self.current_output_file and os.path.exists(self.current_output_file):
            try:
                folder = os.path.dirname(self.current_output_file)
                os.startfile(folder)
            except Exception as e:
                messagebox.showerror("起動エラー", f"フォルダの起動に失敗しました: {e}")

def main():
    app = ShiftGeneratorApp()
    app.mainloop()

if __name__ == "__main__":
    main()
