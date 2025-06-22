import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime
import json
import os

DATA_FILE = "data.json"
SETTINGS_FILE = "settings.json"


class YearApp:
    def __init__(self, root):
        """メインウィンドウの初期化処理。"""
        self.root = root
        self._setup_root()
        self._init_variables()
        self._load_settings()
        self._load_data_from_file()
        self._create_frames_and_widgets()
        self._show_month_sheet(self.current_month)
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _setup_root(self):
        """メインウィンドウの基本設定。"""
        self.root.title("💰 家計管理 2025")
        self.root.geometry("1400x960")  # 高さを増加して全行表示
        self.root.minsize(1200, 800)  # 最小サイズも調整
        self.root.resizable(True, True)

        # モダンなダークテーマベースの配色
        self.colors = {
            'bg_primary': '#1e1e2e',  # メイン背景
            'bg_secondary': '#313244',  # セカンダリ背景
            'bg_tertiary': '#45475a',  # サードレベル背景
            'accent': '#89b4fa',  # アクセントカラー（青）
            'accent_green': '#a6e3a1',  # 緑アクセント
            'accent_red': '#f38ba8',  # 赤アクセント
            'accent_yellow': '#f9e2af',  # 黄アクセント
            'text_primary': '#cdd6f4',  # メインテキスト
            'text_secondary': '#bac2de',  # セカンダリテキスト
            'text_muted': '#6c7086',  # ミュートテキスト
            'border': '#585b70',  # ボーダー
            'hover': '#74c0fc'  # ホバー効果
        }

        self.root.configure(bg=self.colors['bg_primary'])

        # スタイル設定
        style = ttk.Style()
        style.theme_use('clam')

        # カスタムスタイルを定義
        self._setup_modern_styles(style)

    def _init_variables(self):
        """アプリ内で使う各種変数を初期化する。"""
        now = datetime.now()
        self.current_year = now.year
        self.current_month = now.month
        self.tree = None
        self.child_data = {}
        self.parent_table_data = {}
        self.transaction_partners = set()  # 取引先の履歴

        # デフォルトの列定義
        self.default_columns = [
            "日付", "交通", "外食", "食品", "日常用品", "通販", "ゲーム課金",
            "ゲーム購入", "サービス", "家賃", "公共料金", "携帯・回線", "保険", "他"
        ]
        self.custom_columns = []

    def _load_settings(self):
        """設定ファイルから設定を読み込む。"""
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    settings = json.load(f)
                    self.custom_columns = settings.get("custom_columns", [])
                    self.transaction_partners = set(settings.get("transaction_partners", []))
            except Exception as e:
                print(f"設定ファイル読み込みエラー: {e}")

    def _save_settings(self):
        """設定をファイルに保存する。"""
        settings = {
            "custom_columns": self.custom_columns,
            "transaction_partners": list(self.transaction_partners)
        }
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"設定ファイル保存エラー: {e}")

    def _load_data_from_file(self):
        """JSONファイルからデータを読み込む（後方互換性あり）。"""
        if not os.path.exists(DATA_FILE):
            return

        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                all_data = json.load(f)

            # 新しいフォーマットか確認
            if "version" in all_data:
                # 新しいフォーマット
                self.child_data = all_data.get("child_data", {})
                self.parent_table_data = all_data.get("parent_table_data", {})
                # 古いdata.jsonからcustom_columnsとtransaction_partnersを削除（settings.jsonに移行）
            else:
                # 古いフォーマット（後方互換性）
                self.child_data = all_data.get("child_data", {})
                self.parent_table_data = all_data.get("parent_table_data", {})
                # 古いデータから取引先を抽出
                self._extract_transaction_partners_from_old_data()

        except Exception as e:
            print(f"データファイル読み込みエラー: {e}")

    def _extract_transaction_partners_from_old_data(self):
        """古いデータから取引先を抽出する。"""
        for key, data_list in self.child_data.items():
            for row in data_list:
                if len(row) > 0 and row[0].strip():
                    self.transaction_partners.add(row[0].strip())

    def _save_data_to_file(self):
        """データをJSONファイルに保存する（バージョン情報付き）。"""
        all_data = {
            "version": "2.0",
            "child_data": self.child_data,
            "parent_table_data": self.parent_table_data
        }
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(all_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"データファイル保存エラー: {e}")

    def _on_closing(self):
        """ウィンドウ終了時の処理。"""
        self._save_data_to_file()
        self._save_settings()
        self.root.destroy()

    def _setup_modern_styles(self, style):
        """モダンなスタイルを設定。"""
        # ボタンスタイル（小さく）
        style.configure('Modern.TButton',
                        background=self.colors['bg_secondary'],
                        foreground=self.colors['text_primary'],
                        borderwidth=0,
                        focuscolor='none',
                        font=('Segoe UI', 9),
                        relief='flat')

        style.map('Modern.TButton',
                  background=[('active', self.colors['hover']),
                              ('pressed', self.colors['accent'])],
                  foreground=[('active', '#ffffff')])

        # アクセントボタン
        style.configure('Accent.TButton',
                        background=self.colors['accent'],
                        foreground='#ffffff',
                        borderwidth=0,
                        focuscolor='none',
                        font=('Segoe UI', 10, 'bold'),
                        relief='flat')

        style.map('Accent.TButton',
                  background=[('active', self.colors['hover']),
                              ('pressed', '#4dabf7')])

        # 選択された月ボタン（小さく）
        style.configure('Selected.TButton',
                        background=self.colors['accent'],
                        foreground='#ffffff',
                        borderwidth=0,
                        focuscolor='none',
                        font=('Segoe UI', 9, 'bold'),
                        relief='flat')

        # 年ナビゲーションボタン（小さく）
        style.configure('Nav.TButton',
                        background=self.colors['bg_tertiary'],
                        foreground=self.colors['text_primary'],
                        borderwidth=0,
                        focuscolor='none',
                        font=('Segoe UI', 12, 'bold'),
                        relief='flat')

        style.map('Nav.TButton',
                  background=[('active', self.colors['accent']),
                              ('pressed', self.colors['hover'])])

        # 追加ボタン
        style.configure('Add.TButton',
                        background=self.colors['accent_green'],
                        foreground='#1e1e2e',
                        borderwidth=0,
                        focuscolor='none',
                        font=('Segoe UI', 11, 'bold'),
                        relief='flat')

        style.map('Add.TButton',
                  background=[('active', '#94d82d'),
                              ('pressed', '#74c0fc')])

        # 列追加ボタンスタイル
        style.configure('AddColumn.TButton',
                        background='#e8e8e8',
                        foreground='#333333',
                        borderwidth=1,
                        focuscolor='none',
                        font=('Arial', 12, 'bold'),
                        relief='raised')

        style.map('AddColumn.TButton',
                  background=[('active', '#d0d0d0'),
                              ('pressed', '#b0b0b0')])

    def _create_frames_and_widgets(self):
        """メインウィンドウのウィジェットを作成。"""
        # メインコンテナ
        main_container = tk.Frame(self.root, bg=self.colors['bg_primary'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # コンパクトなヘッダーセクション
        header_section = tk.Frame(main_container, bg=self.colors['bg_secondary'], relief='flat', bd=0)
        header_section.pack(fill=tk.X, pady=(0, 8))

        # 内側のパディング（小さく）
        header_inner = tk.Frame(header_section, bg=self.colors['bg_secondary'])
        header_inner.pack(fill=tk.X, padx=15, pady=8)

        # 年選択部分（左側、コンパクト）
        year_container = tk.Frame(header_inner, bg=self.colors['bg_secondary'])
        year_container.pack(side=tk.LEFT)

        # 年ナビゲーション（小さく）
        year_nav_frame = tk.Frame(year_container, bg=self.colors['bg_secondary'])
        year_nav_frame.pack()

        self.minus_button = ttk.Button(year_nav_frame, text="◀", width=3, style='Nav.TButton',
                                       command=self._decrease_year)
        self.minus_button.pack(side=tk.LEFT, padx=(0, 4))

        # 年表示（小さく）
        year_display = tk.Frame(year_nav_frame, bg=self.colors['bg_tertiary'], relief='flat')
        year_display.pack(side=tk.LEFT, padx=4)

        self.year_label = tk.Label(year_display, text=str(self.current_year),
                                   font=('Segoe UI', 16, 'bold'),
                                   bg=self.colors['bg_tertiary'],
                                   fg=self.colors['text_primary'],
                                   padx=12, pady=4)
        self.year_label.pack()

        self.plus_button = ttk.Button(year_nav_frame, text="▶", width=3, style='Nav.TButton',
                                      command=self._increase_year)
        self.plus_button.pack(side=tk.LEFT, padx=(4, 0))

        # 月選択ボタン（インライン、小さく）
        month_container = tk.Frame(header_inner, bg=self.colors['bg_secondary'])
        month_container.pack(side=tk.LEFT, padx=(20, 0))

        self.month_buttons = []
        for m in range(1, 13):
            btn = ttk.Button(month_container, text=f"{m:02d}", width=4, style='Modern.TButton',
                             command=lambda mo=m: self._select_month(mo))
            btn.pack(side=tk.LEFT, padx=1)
            self.month_buttons.append(btn)

        # 現在月表示（右側、小さく）
        month_info_frame = tk.Frame(header_inner, bg=self.colors['bg_secondary'])
        month_info_frame.pack(side=tk.RIGHT)

        month_label_frame = tk.Frame(month_info_frame, bg=self.colors['accent'], relief='flat')
        month_label_frame.pack()

        self.current_month_label = tk.Label(month_label_frame,
                                            text=f"📅 {self.current_month:02d}月",
                                            font=('Segoe UI', 12, 'bold'),
                                            bg=self.colors['accent'],
                                            fg='#ffffff',
                                            padx=10, pady=3)
        self.current_month_label.pack()

        # メインの家計テーブルセクション（大きく）
        tree_section = tk.Frame(main_container, bg=self.colors['bg_secondary'], relief='flat')
        tree_section.pack(fill=tk.BOTH, expand=True)

        self._create_tree_if_needed(tree_section)
        self._highlight_selected_month()

    def _create_tree_if_needed(self, parent):
        """Treeviewを作成。"""
        if self.tree is not None:
            return

        # 既存のウィジェットをクリア
        for widget in parent.winfo_children():
            widget.destroy()

        # 列を結合（デフォルト + カスタム + ＋ボタン）
        all_columns = self.default_columns + self.custom_columns + ["＋"]

        # メインのツリービューフレーム（高さを明示的に設定）
        main_tree_frame = tk.Frame(parent, bg='white', relief='solid', bd=1)
        main_tree_frame.pack(fill=tk.BOTH, expand=True)

        # グリッドの重み設定
        main_tree_frame.grid_rowconfigure(0, weight=1)
        main_tree_frame.grid_columnconfigure(0, weight=1)

        # Treeview（高さを明示的に指定）
        self.tree = ttk.Treeview(main_tree_frame, columns=all_columns, show="headings", height=25)
        self.tree.grid(row=0, column=0, sticky="nsew")

        # 各列の設定
        for i, col in enumerate(all_columns):
            self.tree.heading(col, text=col)
            if i == 0:  # 日付列
                self.tree.column(col, anchor="center", width=60, minwidth=50)
            elif col == "＋":  # ＋ボタン列
                self.tree.column(col, anchor="center", width=40, minwidth=40, stretch=False)
            else:
                self.tree.column(col, anchor="center", width=80, minwidth=60)

        # 縦スクロールバー
        v_scrollbar = ttk.Scrollbar(main_tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=v_scrollbar.set)

        # 横スクロールバー
        h_scrollbar = ttk.Scrollbar(main_tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=h_scrollbar.set)

        # スタイル設定（よりコンパクトに）
        style = ttk.Style()
        style.configure("Treeview",
                        fieldbackground="white",
                        background="white",
                        rowheight=25,
                        font=('Arial', 9))
        style.configure("Treeview.Heading",
                        background="#e8e8e8",
                        font=('Arial', 10, 'bold'),
                        relief="raised")

        # タグ設定
        self.tree.tag_configure("TOTAL", background="#fff3cd", font=('Arial', 10, 'bold'))
        self.tree.tag_configure("SUMMARY", background="#d4edda", font=('Arial', 10, 'bold'))

        # イベントバインド
        self.tree.bind("<Double-1>", self._on_parent_double_click)
        self.tree.bind("<Button-1>", self._on_single_click)  # シングルクリック追加
        self.tree.bind("<Button-3>", self._on_header_right_click)  # 右クリック

        # 右クリックメニューの作成
        self.column_context_menu = tk.Menu(self.root, tearoff=0)
        self.column_context_menu.add_command(label="列名を編集", command=self._edit_column_name)
        self.column_context_menu.add_separator()
        self.column_context_menu.add_command(label="列を削除", command=self._delete_column)

        # マウスホイールでのスクロール
        def on_mousewheel(event):
            if event.state & 0x4:  # Ctrlキーが押されている場合は横スクロール
                self.tree.xview_scroll(int(-1 * (event.delta / 120)), "units")
            else:  # 通常は縦スクロール
                self.tree.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.tree.bind("<MouseWheel>", on_mousewheel)
        self.tree.bind("<Shift-MouseWheel>", lambda e: self.tree.xview_scroll(int(-1 * (e.delta / 120)), "units"))

        print(f"Treeview作成完了 - 高さ: 25行, 列数: {len(all_columns)}")

    def _add_column(self):
        """新しい列を追加する。"""
        # ダイアログを親ウィンドウの中央に表示
        dialog = tk.Toplevel(self.root)
        dialog.title("列の追加")
        dialog.resizable(False, False)

        # ダイアログのサイズ
        dialog_width = 300
        dialog_height = 120

        # 親ウィンドウの位置とサイズを取得
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()

        # ダイアログを親ウィンドウの中央に配置
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2

        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        dialog.transient(self.root)
        dialog.grab_set()

        # ウィジェットの作成
        tk.Label(dialog, text="新しい列名を入力してください：", font=('Arial', 11)).pack(pady=10)

        entry = tk.Entry(dialog, font=('Arial', 11), width=25)
        entry.pack(pady=5)
        entry.focus_set()

        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)

        def on_ok():
            column_name = entry.get().strip()
            if column_name:
                all_existing = self.default_columns + self.custom_columns
                if column_name not in all_existing:
                    self.custom_columns.append(column_name)
                    print(f"新しい列を追加: {column_name}")
                    dialog.destroy()
                    self._recreate_tree()
                    self._show_month_sheet(self.current_month)
                else:
                    messagebox.showwarning("警告", "その列名は既に存在します。", parent=dialog)
            else:
                messagebox.showwarning("警告", "列名を入力してください。", parent=dialog)

        def on_cancel():
            dialog.destroy()

        tk.Button(button_frame, text="OK", command=on_ok, width=8).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="キャンセル", command=on_cancel, width=8).pack(side=tk.LEFT, padx=5)

        # Enterキーでも確定できるように
        entry.bind('<Return>', lambda e: on_ok())
        dialog.bind('<Escape>', lambda e: on_cancel())

    def _recreate_tree(self):
        """Treeviewを再作成する。"""
        if self.tree:
            # 現在のツリーの親を取得
            tree_parent = self.tree.master
            self.tree.destroy()
            self.tree = None

            # 同じ親に新しいツリーを作成
            self._create_tree_if_needed(tree_parent)
            return

        # 見つからない場合は、メインコンテナを再作成
        print("ツリービューコンテナが見つかりません。メインレイアウトを再作成します。")

    def _increase_year(self):
        """年を増やす。"""
        self.current_year += 1
        self.year_label.config(text=str(self.current_year))
        self._show_month_sheet(self.current_month)

    def _decrease_year(self):
        """年を減らす。"""
        self.current_year -= 1
        self.year_label.config(text=str(self.current_year))
        self._show_month_sheet(self.current_month)

    def _select_month(self, month):
        """月を選択する。"""
        self.current_month = month
        self.current_month_label.config(text=f"選択中の月：{month:02d}月")
        self._highlight_selected_month()
        self._show_month_sheet(month)

    def _highlight_selected_month(self):
        """選択中の月ボタンをハイライト。"""
        for i, btn in enumerate(self.month_buttons, start=1):
            if i == self.current_month:
                btn.configure(style='Selected.TButton')
            else:
                btn.configure(style='TButton')

        # 選択されたボタンのスタイルを設定
        style = ttk.Style()
        style.configure('Selected.TButton',
                        background='#2196f3',
                        foreground='white',
                        font=('Arial', 10, 'bold'))
        style.map('Selected.TButton',
                  background=[('active', '#1976d2')])
        style.configure('TButton', font=('Arial', 10))

    def _show_month_sheet(self, month):
        """月のシートを表示。"""
        if not self.tree:
            return

        # 既存の行を削除
        for item in self.tree.get_children():
            self.tree.delete(item)

        all_columns = self.default_columns + self.custom_columns
        cols = len(all_columns)
        days = self._get_days_in_month(month)

        # 日付行を挿入
        for day in range(1, days + 1):
            key_day = f"{self.current_year}-{month}-{day}"
            if key_day in self.parent_table_data:
                row_values = self.parent_table_data[key_day]
                # 列数が足りない場合は空文字で埋める
                while len(row_values) < cols:
                    row_values.append("")
            else:
                row_values = [str(day)] + [""] * (cols - 1)
            # ＋ボタン列は空のまま
            row_values.append("")
            self.tree.insert("", "end", values=row_values)

        # 合計行
        total_row = ["合計"] + [""] * (cols - 1) + [""]
        self.tree.insert("", "end", values=total_row, tags=("TOTAL",))

        # まとめ行
        summary_key = f"{self.current_year}-{month}-0"
        if summary_key in self.parent_table_data:
            sm_data = self.parent_table_data[summary_key]
            inc_str = sm_data[3] if len(sm_data) > 3 else ""
        else:
            inc_str = ""

        summary_row = ["まとめ", "", "収入", inc_str, "支出", ""] + [""] * (cols - 6) + [""]
        self.tree.insert("", "end", values=summary_row, tags=("SUMMARY",))

        self._recalc_total_and_summary()

    def _recalc_total_and_summary(self):
        """合計行とまとめ行を再計算。"""
        items = self.tree.get_children()
        if len(items) < 2:
            return

        total_row_id = items[-2]
        summary_row_id = items[-1]
        all_columns = self.default_columns + self.custom_columns
        cols = len(all_columns)  # ＋ボタン列は計算に含めない

        # 合計行の計算
        sums = [0] * (cols - 1)
        for row_id in items[:-2]:
            row_vals = self.tree.item(row_id, 'values')
            for i in range(1, cols):
                try:
                    val = int(row_vals[i])
                except (ValueError, TypeError, IndexError):
                    val = 0
                sums[i - 1] += val

        # 合計行を更新
        total_vals = list(self.tree.item(total_row_id, 'values'))
        for i in range(1, cols):
            if sums[i - 1] == 0:
                total_vals[i] = ""
            else:
                total_vals[i] = str(sums[i - 1])
        # ＋ボタン列は空のまま
        while len(total_vals) <= cols:
            total_vals.append("")
        self.tree.item(total_row_id, values=total_vals)

        # 総支出計算
        grand_total = 0
        for v in total_vals[1:cols]:  # ＋ボタン列は除外
            if v and str(v).strip():
                try:
                    grand_total += int(v)
                except:
                    pass

        # まとめ行更新
        summary_vals = list(self.tree.item(summary_row_id, 'values'))
        try:
            income_val = int(summary_vals[3]) if len(summary_vals) > 3 else 0
        except:
            income_val = 0

        sum_val = income_val - grand_total
        if sum_val == 0:
            summary_vals[1] = ""
        else:
            summary_vals[1] = str(sum_val)

        if grand_total == 0:
            summary_vals[5] = ""
        else:
            summary_vals[5] = str(grand_total)

        # ＋ボタン列は空のまま
        while len(summary_vals) <= cols:
            summary_vals.append("")

        self.tree.item(summary_row_id, values=summary_vals)

    def _on_single_click(self, event):
        """シングルクリックの処理（＋ボタン用）。"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            col_id = self.tree.identify_column(event.x)
            if col_id:
                col_index = int(col_id[1:]) - 1
                all_columns = self.default_columns + self.custom_columns

                # ＋ボタン列のクリック
                if col_index == len(all_columns):
                    self._add_column()

    def _on_parent_double_click(self, event):
        """親ダイアログのダブルクリック処理。"""
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return

        # ヘッダー部分のクリックかチェック
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            # ヘッダークリックの処理
            col_index = int(col_id[1:]) - 1
            all_columns = self.default_columns + self.custom_columns

            if col_index == len(all_columns):  # ＋ボタン列
                self._add_column()
                return
            elif col_index >= len(self.default_columns):  # カスタム列
                self._edit_column_name(col_index)
                return
            else:
                return  # デフォルト列は編集不可

        items = self.tree.get_children()
        if len(items) < 2:
            return

        total_row_id = items[-2]
        summary_row_id = items[-1]

        # 合計行は編集不可
        if row_id == total_row_id:
            return

        # まとめ行は収入金額のみ編集可
        if row_id == summary_row_id:
            if col_id != "#4":
                return

        # 日付列は編集不可
        if col_id == "#1":
            return

        # ＋ボタン列は編集不可
        col_index = int(col_id[1:]) - 1
        all_columns = self.default_columns + self.custom_columns
        if col_index >= len(all_columns):
            return

        row_vals = self.tree.item(row_id, 'values')
        if not row_vals:
            return

        if row_id == summary_row_id:
            day = 0
        else:
            try:
                day = int(row_vals[0])
            except:
                day = 0

        dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_index}"
        col_name = self.tree.heading(col_id, "text")

        print(f"ダブルクリック - キー: {dict_key}, 列名: {col_name}")
        print(f"現在の child_data: {self.child_data}")

        ChildDialog(self.root, self, dict_key, col_name)

    def _on_header_right_click(self, event):
        """ヘッダーの右クリック処理。"""
        region = self.tree.identify_region(event.x, event.y)
        if region != "heading":
            return

        col_id = self.tree.identify_column(event.x)
        if not col_id:
            return

        col_index = int(col_id[1:]) - 1
        all_columns = self.default_columns + self.custom_columns

        # ＋ボタン列や日付列、デフォルト列は右クリック無効
        if col_index >= len(all_columns) or col_index == 0 or col_index < len(self.default_columns):
            return

        # カスタム列のみ右クリック可能
        self.selected_column_index = col_index
        self.column_context_menu.post(event.x_root, event.y_root)

    def _edit_column_name(self, col_index=None):
        """列名を編集。"""
        if col_index is None:
            col_index = getattr(self, 'selected_column_index', None)

        if col_index is None or col_index < len(self.default_columns):
            return

        custom_index = col_index - len(self.default_columns)
        if custom_index >= len(self.custom_columns):
            return

        old_name = self.custom_columns[custom_index]

        # ダイアログを親ウィンドウの中央に表示
        dialog = tk.Toplevel(self.root)
        dialog.title("列名の編集")
        dialog.resizable(False, False)

        # ダイアログのサイズ
        dialog_width = 300
        dialog_height = 120

        # 親ウィンドウの位置とサイズを取得
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()

        # ダイアログを親ウィンドウの中央に配置
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2

        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        dialog.transient(self.root)
        dialog.grab_set()

        # ウィジェットの作成
        tk.Label(dialog, text="新しい列名を入力してください：", font=('Arial', 11)).pack(pady=10)

        entry = tk.Entry(dialog, font=('Arial', 11), width=25)
        entry.pack(pady=5)
        entry.insert(0, old_name)
        entry.select_range(0, tk.END)
        entry.focus_set()

        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)

        def on_ok():
            new_name = entry.get().strip()
            if new_name and new_name != old_name:
                # 重複チェック
                all_columns = self.default_columns + self.custom_columns
                if new_name not in all_columns:
                    self.custom_columns[custom_index] = new_name
                    dialog.destroy()
                    self._recreate_tree()
                    self._show_month_sheet(self.current_month)
                else:
                    messagebox.showwarning("警告", "その列名は既に存在します。", parent=dialog)
            else:
                dialog.destroy()

        def on_cancel():
            dialog.destroy()

        tk.Button(button_frame, text="OK", command=on_ok, width=8).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="キャンセル", command=on_cancel, width=8).pack(side=tk.LEFT, padx=5)

        # Enterキーでも確定できるように
        entry.bind('<Return>', lambda e: on_ok())
        dialog.bind('<Escape>', lambda e: on_cancel())

    def _delete_column(self):
        """列を削除。"""
        col_index = getattr(self, 'selected_column_index', None)
        if col_index is None or col_index < len(self.default_columns):
            return

        custom_index = col_index - len(self.default_columns)
        if custom_index >= len(self.custom_columns):
            return

        col_name = self.custom_columns[custom_index]
        result = messagebox.askyesno("確認", f"列 '{col_name}' を削除しますか？\n※この列のデータもすべて削除されます。")

        if result:
            # 列を削除
            deleted_col_name = self.custom_columns.pop(custom_index)

            # 関連するデータも削除
            keys_to_delete = []
            for key in list(self.child_data.keys()):
                parts = key.split("-")
                if len(parts) == 4 and int(parts[3]) == col_index:
                    keys_to_delete.append(key)

            for key in keys_to_delete:
                del self.child_data[key]

            # parent_table_dataからも該当列を削除
            for date_key in self.parent_table_data:
                row_data = self.parent_table_data[date_key]
                if len(row_data) > col_index:
                    row_data[col_index] = ""

            print(f"列 '{deleted_col_name}' を削除しました")
            self._recreate_tree()
            self._show_month_sheet(self.current_month)

    def update_parent_cell(self, dict_key_day, col_index, new_value):
        """親セルの値を更新。"""
        all_columns = self.default_columns + self.custom_columns
        cols = len(all_columns)

        # 空文字列の場合は親テーブルからも削除
        if not new_value or str(new_value).strip() == "":
            print(f"空の値のため親テーブルデータを更新: {dict_key_day}")
            if dict_key_day in self.parent_table_data:
                row_array = self.parent_table_data[dict_key_day]
                while len(row_array) < cols:
                    row_array.append("")
                row_array[col_index] = ""

                # 行全体が空になった場合は削除
                if all(not str(cell).strip() for cell in row_array[1:]):  # 日付列以外が全て空
                    print(f"行全体が空のため parent_table_data から削除: {dict_key_day}")
                    del self.parent_table_data[dict_key_day]
                else:
                    self.parent_table_data[dict_key_day] = row_array
        else:
            # 通常の値更新処理
            if dict_key_day not in self.parent_table_data:
                parts = dict_key_day.split("-")
                day_str = parts[2]
                row_array = [day_str] + [""] * (cols - 1)
                self.parent_table_data[dict_key_day] = row_array
            else:
                row_array = self.parent_table_data[dict_key_day]
                # 列数が足りない場合は拡張
                while len(row_array) < cols:
                    row_array.append("")

            row_array[col_index] = str(new_value)
            self.parent_table_data[dict_key_day] = row_array

        # 画面更新
        y, mo, d = dict_key_day.split("-")
        y, mo, d = int(y), int(mo), int(d)
        if (self.current_year == y) and (self.current_month == mo):
            items = self.tree.get_children()
            if len(items) < 2:
                return

            summary_row_id = items[-1]

            found = False
            for row_id in items[:-2]:
                row_vals = list(self.tree.item(row_id, 'values'))
                if row_vals and row_vals[0] == str(d):
                    # 列数が足りない場合は拡張
                    while len(row_vals) < cols + 1:  # +1 for ＋ button column
                        row_vals.append("")

                    # 空文字列または0の場合は空表示
                    display_value = ""
                    if new_value and str(new_value).strip() != "" and str(new_value) != "0":
                        display_value = str(new_value)

                    row_vals[col_index] = display_value
                    self.tree.item(row_id, values=row_vals)
                    found = True
                    break

            if not found and d == 0:
                sum_vals = list(self.tree.item(summary_row_id, 'values'))
                while len(sum_vals) < cols + 1:  # +1 for ＋ button column
                    sum_vals.append("")

                # まとめ行でも同様の処理
                display_value = ""
                if new_value and str(new_value).strip() != "" and str(new_value) != "0":
                    display_value = str(new_value)

                sum_vals[col_index] = display_value
                self.tree.item(summary_row_id, values=sum_vals)

            self._recalc_total_and_summary()

    def _get_days_in_month(self, month):
        """月の日数を取得。"""
        if month in [1, 3, 5, 7, 8, 10, 12]:
            return 31
        elif month in [4, 6, 9, 11]:
            return 30
        elif month == 2:
            y = self.current_year
            if (y % 4 == 0 and y % 100 != 0) or (y % 400 == 0):
                return 29
            else:
                return 28
        return 30


class ChildDialog(tk.Toplevel):
    def __init__(self, parent, parent_app, dict_key, col_name):
        """子ダイアログの初期化。"""
        super().__init__(parent)
        self.parent_app = parent_app
        self.dict_key = dict_key
        self.col_name = col_name

        print(f"ChildDialog初期化 - キー: {dict_key}")
        print(f"利用可能なchild_dataキー: {list(parent_app.child_data.keys())}")

        # ダイアログの設定
        dialog_width = 700
        dialog_height = 500
        min_width = 500
        min_height = 350

        # 親ウィンドウの中央に配置
        parent_x = self.master.winfo_x()
        parent_y = self.master.winfo_y()
        parent_w = self.master.winfo_width()
        parent_h = self.master.winfo_height()

        x = parent_x + (parent_w - dialog_width) // 2
        y = parent_y + (parent_h - dialog_height) // 2

        # 画面外に出ないように調整
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        if x + dialog_width > screen_width:
            x = screen_width - dialog_width - 20
        if x < 0:
            x = 20
        if y + dialog_height > screen_height:
            y = screen_height - dialog_height - 50
        if y < 0:
            y = 20

        self.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        self.minsize(min_width, min_height)  # 最小サイズを設定
        self.title(f"支出・収入詳細 - {col_name}")
        self.configure(bg='#f0f0f0')

        # モーダルに設定
        self.transient(parent)
        self.grab_set()

        # リサイズ可能に設定
        self.resizable(True, True)

        # キー解析
        parts = dict_key.split("-")
        print(f"キー分割結果: {parts}")
        if len(parts) == 4:
            self.year = int(parts[0])
            self.month = int(parts[1])
            self.day = int(parts[2])
            self.col_index = int(parts[3])
        else:
            self.year, self.month, self.day, self.col_index = (0, 0, 0, 0)

        print(f"解析結果 - 年:{self.year}, 月:{self.month}, 日:{self.day}, 列:{self.col_index}")

        self.child_tree = None
        self.entry_editor = None
        self.context_menu = None

        self._create_widgets()
        # _populate_data()の呼び出しを削除（_create_widgets内で呼び出すため）

    def _create_widgets(self):
        """ウィジェットを作成。"""
        # グリッドの重み設定（リサイズ対応）
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # メインフレーム（スクロール対応）
        main_canvas = tk.Canvas(self, bg='#f0f0f0', highlightthickness=0)
        main_canvas.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        # スクロールバー
        main_scrollbar = ttk.Scrollbar(self, orient="vertical", command=main_canvas.yview)
        main_scrollbar.grid(row=0, column=1, sticky="ns")
        main_canvas.configure(yscrollcommand=main_scrollbar.set)

        # スクロール可能フレーム
        self.scrollable_frame = tk.Frame(main_canvas, bg='#f0f0f0')
        canvas_frame = main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # タイトル
        title_label = tk.Label(self.scrollable_frame, text=f"{self.col_name} の詳細入力",
                               font=('Arial', 16, 'bold'), bg='#f0f0f0')
        title_label.pack(pady=(10, 15))

        # ツリービューコンテナ
        tree_container = tk.Frame(self.scrollable_frame, bg='#f0f0f0')
        tree_container.pack(fill=tk.BOTH, expand=True, padx=10)

        # ツリービューとスクロールバー
        tree_frame = tk.Frame(tree_container, bg='#f0f0f0')
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ["取引先", "金額", "詳細"]
        self.child_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")

        for col in columns:
            self.child_tree.heading(col, text=col)
            if col == "取引先":
                self.child_tree.column(col, anchor="center", width=150, minwidth=100)
            elif col == "金額":
                self.child_tree.column(col, anchor="center", width=100, minwidth=80)
            else:
                self.child_tree.column(col, anchor="center", width=200, minwidth=150)

        # ツリービュー用スクロールバー
        tree_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.child_tree.yview)
        self.child_tree.configure(yscrollcommand=tree_scrollbar.set)

        self.child_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # イベントバインド
        self.child_tree.bind("<Double-1>", self._on_double_click)
        self.child_tree.bind("<Button-3>", self._on_right_click)  # 右クリック

        # ボタンフレーム（固定位置）
        button_frame = tk.Frame(self.scrollable_frame, bg='#f0f0f0')
        button_frame.pack(fill=tk.X, pady=(15, 10), padx=10)

        # 行追加ボタン
        add_btn = tk.Button(button_frame, text="行を追加", font=('Arial', 12),
                            bg='#4caf50', fg='white', relief='raised', bd=2,
                            activebackground='#45a049', command=self._add_row)
        add_btn.pack(side=tk.LEFT, padx=(0, 10), ipady=5)

        # 右側ボタングループ
        right_button_frame = tk.Frame(button_frame, bg='#f0f0f0')
        right_button_frame.pack(side=tk.RIGHT)

        # キャンセルボタン
        cancel_btn = tk.Button(right_button_frame, text="キャンセル", font=('Arial', 12),
                               bg='#f44336', fg='white', relief='raised', bd=2,
                               activebackground='#d32f2f', command=self.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=(10, 0), ipady=5)

        # OKボタン
        ok_btn = tk.Button(right_button_frame, text="OK", font=('Arial', 12, 'bold'),
                           bg='#2196f3', fg='white', relief='raised', bd=2,
                           activebackground='#1976d2', command=self._on_ok_button)
        ok_btn.pack(side=tk.RIGHT, ipady=5)

        # 右クリックメニュー
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="行を削除", command=self._delete_row)

        # 使用方法のヒント
        hint_label = tk.Label(self.scrollable_frame,
                              text="使い方: セルをダブルクリックで編集、右クリックで行削除",
                              font=('Arial', 10), fg='#666666', bg='#f0f0f0')
        hint_label.pack(pady=(5, 10))

        # キーボードショートカット
        self.bind('<Return>', lambda e: self._on_ok_button())
        self.bind('<Escape>', lambda e: self.destroy())

        # スクロール設定のイベントバインド
        def configure_scroll(event):
            main_canvas.configure(scrollregion=main_canvas.bbox("all"))
            # キャンバス幅を調整
            canvas_width = main_canvas.winfo_width()
            main_canvas.itemconfig(canvas_frame, width=canvas_width)

        self.scrollable_frame.bind('<Configure>', configure_scroll)

        # マウスホイールでスクロール
        def on_mousewheel(event):
            main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        main_canvas.bind("<MouseWheel>", on_mousewheel)

        # ツリービューの最小高さを設定
        style = ttk.Style()
        style.configure("Treeview", rowheight=25)

        # 初期フォーカスをツリービューに
        self.after(100, lambda: self.child_tree.focus_set())

        # ウィジェット作成後にデータを読み込み
        self.after(50, self._populate_data)

        # 強制的にウィンドウを前面に表示
        self.lift()
        self.focus_force()

    def _populate_data(self):
        """データを表示。"""
        if not self.child_tree:
            print(f"警告: child_treeが初期化されていません")
            return

        # 既存のデータをクリア
        for item in self.child_tree.get_children():
            self.child_tree.delete(item)

        data_list = self.parent_app.child_data.get(self.dict_key, [])
        print(f"データ読み込み - キー: {self.dict_key}, データ数: {len(data_list)}")

        if not data_list:
            # データがない場合は空行を1つ追加
            print("データが空のため、空行を追加します")
            self.child_tree.insert("", "end", values=["", "", ""])
        else:
            # データがある場合は全て表示
            for i, row in enumerate(data_list):
                # 3要素未満の場合は空文字で埋める
                row_data = list(row) if row else ["", "", ""]
                while len(row_data) < 3:
                    row_data.append("")
                print(f"行{i}: {row_data}")
                self.child_tree.insert("", "end", values=row_data)

            # 最後に空行を追加（新規入力用）
            self.child_tree.insert("", "end", values=["", "", ""])

        print(f"表示完了 - 総行数: {len(self.child_tree.get_children())}")

    def _add_row(self):
        """空行を追加。"""
        print("空行を追加します")
        if self.child_tree:
            item_id = self.child_tree.insert("", "end", values=["", "", ""])
            print(f"追加された行ID: {item_id}")
            # 追加した行を選択状態にする
            self.child_tree.selection_set(item_id)
            self.child_tree.see(item_id)
        else:
            print("エラー: child_treeが存在しません")

    def _on_double_click(self, event):
        """セルのダブルクリック編集。"""
        # 編集中のエディターがあれば先に保存
        if self.entry_editor:
            self._cancel_edit()

        item_id = self.child_tree.identify_row(event.y)
        col_id = self.child_tree.identify_column(event.x)
        print(f"ダブルクリック - 行ID: {item_id}, 列ID: {col_id}")

        if not item_id or not col_id:
            print("行IDまたは列IDが無効です")
            return

        # Treeviewの相対座標でbboxを取得
        bbox = self.child_tree.bbox(item_id, col_id)
        print(f"bbox: {bbox}")
        if not bbox:
            print("bboxが取得できませんでした")
            return

        col_idx = int(col_id[1:]) - 1
        old_vals = list(self.child_tree.item(item_id, 'values'))
        print(f"現在の値: {old_vals}, 編集列インデックス: {col_idx}")

        # 値が足りない場合は空文字で埋める
        while len(old_vals) <= col_idx:
            old_vals.append("")

        x, y, w, h = bbox
        print(f"エディター配置位置: x={x}, y={y}, w={w}, h={h}")

        # 取引先列の場合はコンボボックス、それ以外はEntry
        if col_idx == 0:  # 取引先列
            print("取引先列 - コンボボックスを作成")
            self.entry_editor = ttk.Combobox(self.child_tree, font=("Arial", 11))
            # 取引先履歴をセット
            partner_list = sorted(list(self.parent_app.transaction_partners))
            self.entry_editor['values'] = partner_list
            self.entry_editor.set(old_vals[col_idx])
        else:
            print(f"通常列({col_idx}) - Entryを作成")
            self.entry_editor = tk.Entry(self.child_tree, font=("Arial", 11))
            self.entry_editor.insert(0, str(old_vals[col_idx]))

        # エディターを配置
        self.entry_editor.place(x=x, y=y, width=w, height=h)
        self.entry_editor.focus_set()
        self.entry_editor.select_range(0, tk.END)

        print("エディター配置完了")

        # イベントバインド - lambdaで現在の値をキャプチャ
        def save_and_cleanup():
            self._save_entry(item_id, col_idx)

        def cancel_and_cleanup():
            self._cancel_edit()

        self.entry_editor.bind("<Return>", lambda e: save_and_cleanup())
        self.entry_editor.bind("<Tab>", lambda e: save_and_cleanup())
        self.entry_editor.bind("<FocusOut>", lambda e: save_and_cleanup())
        self.entry_editor.bind("<Escape>", lambda e: cancel_and_cleanup())

        # 他の場所をクリックした時も保存
        def on_tree_click(e):
            if self.entry_editor:
                save_and_cleanup()

        self.child_tree.bind("<Button-1>", on_tree_click, add=True)

    def _cancel_edit(self):
        """編集をキャンセル。"""
        print("編集をキャンセルします")
        if self.entry_editor:
            try:
                self.entry_editor.destroy()
            except:
                pass
            self.entry_editor = None

        # 一時的なイベントバインドを削除
        try:
            self.child_tree.unbind("<Button-1>")
            self.child_tree.bind("<Button-1>", lambda e: None)  # デフォルトの動作を復元
        except:
            pass

    def _save_entry(self, item_id, col_idx):
        """編集内容を保存。"""
        print(f"編集内容を保存 - 行ID: {item_id}, 列: {col_idx}")
        if not self.entry_editor:
            print("エディターが存在しません")
            return

        try:
            new_val = self.entry_editor.get()
            print(f"新しい値: '{new_val}'")

            # エディターを削除
            self.entry_editor.destroy()
            self.entry_editor = None

            # 取引先を履歴に追加
            if col_idx == 0 and new_val.strip():
                self.parent_app.transaction_partners.add(new_val.strip())
                print(f"取引先履歴に追加: {new_val.strip()}")

            # ツリービューの値を更新
            row_vals = list(self.child_tree.item(item_id, 'values'))
            while len(row_vals) <= col_idx:
                row_vals.append("")

            old_val = row_vals[col_idx] if col_idx < len(row_vals) else ""
            row_vals[col_idx] = new_val

            print(f"行の値を更新: {old_val} -> {new_val}")
            self.child_tree.item(item_id, values=row_vals)

            # 最終行で何か入力されたら新しい行を追加
            all_items = self.child_tree.get_children()
            if item_id == all_items[-1]:
                if any(str(cell).strip() for cell in row_vals):
                    print("最終行に入力されたため新行を追加")
                    self._add_row()

            # 一時的なイベントバインドを削除
            try:
                self.child_tree.unbind("<Button-1>")
                self.child_tree.bind("<Button-1>", lambda e: None)
            except:
                pass

        except Exception as e:
            print(f"保存エラー: {e}")
            import traceback
            traceback.print_exc()
            if self.entry_editor:
                try:
                    self.entry_editor.destroy()
                except:
                    pass
                self.entry_editor = None

    def _on_right_click(self, event):
        """右クリックメニューを表示。"""
        item_id = self.child_tree.identify_row(event.y)
        if item_id:
            self.child_tree.selection_set(item_id)
            self.context_menu.post(event.x_root, event.y_root)

    def _delete_row(self):
        """選択された行を削除。"""
        selected_item = self.child_tree.selection()
        if selected_item:
            result = messagebox.askyesno("確認", "選択された行を削除しますか？")
            if result:
                self.child_tree.delete(selected_item[0])

    def _on_ok_button(self):
        """OKボタンの処理。"""
        all_rows = []
        for iid in self.child_tree.get_children():
            vals = self.child_tree.item(iid, 'values')
            # 3要素に統一
            row = list(vals)
            while len(row) < 3:
                row.append("")
            all_rows.append(tuple(row))

        # 空行を除去（全ての列が空またはスペースのみの行を除外）
        filtered_rows = []
        for row in all_rows:
            if any(str(cell).strip() for cell in row):
                filtered_rows.append(row)

        print(f"フィルター前の行数: {len(all_rows)}, フィルター後: {len(filtered_rows)}")

        # データが何もない場合の処理
        if not filtered_rows:
            print("データが空のため、親セルを空白にして保存を削除します")
            # 親セルを空白に設定
            dict_key_day = f"{self.year}-{self.month}-{self.day}"
            self.parent_app.update_parent_cell(dict_key_day, self.col_index, "")

            # child_dataからもこのキーを削除
            if self.dict_key in self.parent_app.child_data:
                del self.parent_app.child_data[self.dict_key]
                print(f"child_dataからキー {self.dict_key} を削除しました")
        else:
            # データがある場合は通常の処理
            print(f"データがあるため通常保存: {len(filtered_rows)}行")

            # データ保存
            self.parent_app.child_data[self.dict_key] = filtered_rows

            # 金額列（インデックス1）を合計
            total = 0
            for row in filtered_rows:
                try:
                    amount = str(row[1]).replace(',', '').replace('¥', '').strip()
                    if amount:
                        total += int(amount)
                except (ValueError, IndexError):
                    pass

            print(f"計算された合計金額: {total}")

            # 親セル更新（合計が0の場合は空文字列）
            dict_key_day = f"{self.year}-{self.month}-{self.day}"
            display_value = str(total) if total != 0 else ""
            self.parent_app.update_parent_cell(dict_key_day, self.col_index, display_value)

        self.destroy()


# ----------------------------------------------------------------------
# メイン実行部分
# ----------------------------------------------------------------------
if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = YearApp(root)
        root.mainloop()
    except Exception as e:
        print(f"アプリケーション実行エラー: {e}")
        import traceback

        traceback.print_exc()