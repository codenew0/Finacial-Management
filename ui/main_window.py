# ui/main_window.py
"""
メインウィンドウの実装
家計管理アプリケーションのメインUI
"""
import tkinter as tk
import json
from tkinter import ttk, messagebox
import tkinter.font as tkfont
from config import (
    WindowConfig, ColorTheme, TreeviewConfig, DefaultColumns,
    FontConfig, get_current_year, get_current_month, parse_amount
)
from models.data_manager import DataManager
from ui.tooltip import TreeviewTooltip
from ui.transaction_dialog import TransactionDialog
from ui.monthly_data_dialog import MonthlyDataDialog
from ui.search_dialog import SearchDialog
from ui.chart_dialog import ChartDialog
from utils.date_utils import get_days_in_month


class MainWindow:
    """
    家計管理アプリケーションのメインウィンドウクラス。
    
    年間の家計データを月別に管理し、項目ごとの支出・収入を
    記録・集計・分析する機能を提供する。
    """
    
    def __init__(self, root):
        """
        メインウィンドウを初期化する。
        
        Args:
            root: Tkinterのルートウィンドウ
        """
        self.root = root
        self.data_manager = DataManager()
        self.tree = None
        self.tooltip = None
        self.current_year = get_current_year()
        self.current_month = get_current_month()
        self.colors = self._get_color_theme()

        # コピペ用：選択された列のIDを保持
        self.selected_column_id = None
        
        # 月選択ボタンのリスト
        self.month_buttons = []
        self.current_month_button = None
        self.year_label = None
        
        # 初期化
        self._setup_window()
        self._load_data()
        self._create_ui()
        self._show_month(self.current_month)
        
        # ウィンドウクローズ時の処理を設定
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # グローバルキーボードショートカット
        self.root.bind('<Control-f>', lambda e: SearchDialog(self.root, self))

        # コピー＆ペーストのショートカット
        self.root.bind('<Control-c>', self._copy_cell)
        self.root.bind('<Control-v>', self._paste_cell)
    
    def _get_color_theme(self):
        """カラーテーマを取得"""
        return {
            'bg_primary': ColorTheme.BG_PRIMARY,
            'bg_secondary': ColorTheme.BG_SECONDARY,
            'bg_tertiary': ColorTheme.BG_TERTIARY,
            'accent': ColorTheme.ACCENT,
            'accent_green': ColorTheme.ACCENT_GREEN,
            'accent_red': ColorTheme.ACCENT_RED,
            'text_primary': ColorTheme.TEXT_PRIMARY,
            'text_secondary': ColorTheme.TEXT_SECONDARY,
            'border': ColorTheme.BORDER,
            'hover': ColorTheme.HOVER
        }
    
    def _setup_window(self):
        """メインウィンドウの基本設定を行う"""
        self.root.title("💰 家計管理 2025")
        
        # ウィンドウサイズと位置の設定
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        x = (screen_width - WindowConfig.WIDTH) // 2
        y = (screen_height - WindowConfig.HEIGHT) // 2
        
        self.root.geometry(f"{WindowConfig.WIDTH}x{WindowConfig.HEIGHT}+{x}+{y}")
        self.root.minsize(WindowConfig.MIN_WIDTH, WindowConfig.MIN_HEIGHT)
        self.root.resizable(*WindowConfig.RESIZABLE)
        
        self.root.configure(bg=self.colors['bg_primary'])
        self._setup_styles()
    
    def _setup_styles(self):
        """ttkウィジェットのカスタムスタイルを定義する"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 通常のボタンスタイル
        style.configure('Modern.TButton',
                        background=self.colors['bg_secondary'],
                        foreground=self.colors['text_primary'],
                        borderwidth=0,
                        focuscolor='none',
                        font=FontConfig.BUTTON,
                        relief='flat')
        
        style.map('Modern.TButton',
                  background=[('active', self.colors['hover']),
                              ('pressed', self.colors['accent'])],
                  foreground=[('active', '#ffffff')])
        
        # アクセントボタンスタイル
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
        
        # 月表示ボタンスタイル
        style.configure('Month.TButton',
                        background=self.colors['accent'],
                        foreground='#ffffff',
                        borderwidth=0,
                        focuscolor='none',
                        font=FontConfig.BUTTON_LARGE,
                        relief='flat')
        
        style.map('Month.TButton',
                  background=[('active', self.colors['hover']),
                              ('pressed', '#4dabf7')])
        
        # 選択された月ボタンスタイル
        style.configure('Selected.TButton',
                        background=self.colors['accent'],
                        foreground='#ffffff',
                        borderwidth=0,
                        focuscolor='none',
                        font=('Segoe UI', 9, 'bold'),
                        relief='flat')
        
        # ナビゲーションボタンスタイル
        style.configure('Nav.TButton',
                        background=self.colors['bg_tertiary'],
                        foreground=self.colors['text_primary'],
                        borderwidth=0,
                        focuscolor='none',
                        font=FontConfig.BUTTON_LARGE,
                        relief='flat')
        
        style.map('Nav.TButton',
                  background=[('active', self.colors['accent']),
                              ('pressed', self.colors['hover'])])
    
    def _load_data(self):
        """データと設定を読み込む"""
        self.data_manager.load_settings()
        self.data_manager.load_data()
    
    def _save_data(self):
        """データと設定を保存する"""
        self.data_manager.save_data()
        self.data_manager.save_settings()
    
    def _on_closing(self):
        """ウィンドウが閉じられる時の処理"""
        self._save_data()
        self.root.destroy()
    
    def _create_ui(self):
        """メインウィンドウのUI要素を作成する"""
        # メインコンテナ
        main_container = tk.Frame(self.root, bg=self.colors['bg_primary'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        # ヘッダーセクション
        header = tk.Frame(main_container, bg=self.colors['bg_secondary'])
        header.pack(fill=tk.X, pady=(0, 8))
        
        header_inner = tk.Frame(header, bg=self.colors['bg_secondary'])
        header_inner.pack(fill=tk.X, padx=15, pady=8)
        
        # 年選択コントロール(左側)
        self._create_year_controls(header_inner)
        
        # 月選択ボタン(1月～12月)
        self._create_month_buttons(header_inner)
        
        # 検索ボタン
        self._create_search_button(header_inner)
        
        # 図表ボタン
        self._create_chart_button(header_inner)
        
        # 現在月表示(右側、クリック可能)
        self._create_current_month_button(header_inner)
        
        # メインテーブルセクション
        tree_section = tk.Frame(main_container, bg=self.colors['bg_secondary'])
        tree_section.pack(fill=tk.BOTH, expand=True)
        
        self._create_treeview(tree_section)
        self._update_month_buttons()
    
    def _create_year_controls(self, parent):
        """年選択コントロールを作成"""
        year_container = tk.Frame(parent, bg=self.colors['bg_secondary'])
        year_container.pack(side=tk.LEFT)
        
        year_nav = tk.Frame(year_container, bg=self.colors['bg_secondary'])
        year_nav.pack()
        
        # 前年ボタン
        ttk.Button(year_nav, text="◀", width=3, style='Nav.TButton',
                   command=self._prev_year).pack(side=tk.LEFT, padx=(0, 4))
        
        # 年表示
        year_display = tk.Frame(year_nav, bg=self.colors['bg_tertiary'])
        year_display.pack(side=tk.LEFT, padx=4)
        
        self.year_label = tk.Label(year_display, text=str(self.current_year),
                                   font=FontConfig.TITLE,
                                   bg=self.colors['bg_tertiary'],
                                   fg=self.colors['text_primary'],
                                   padx=12, pady=4)
        self.year_label.pack()
        
        # 翌年ボタン
        ttk.Button(year_nav, text="▶", width=3, style='Nav.TButton',
                   command=self._next_year).pack(side=tk.LEFT, padx=(4, 0))
    
    def _create_month_buttons(self, parent):
        """月選択ボタンを作成"""
        month_container = tk.Frame(parent, bg=self.colors['bg_secondary'])
        month_container.pack(side=tk.LEFT, padx=(20, 0))
        
        self.month_buttons = []
        for m in range(1, 13):
            btn = ttk.Button(month_container, text=f"{m:02d}", width=4, style='Modern.TButton',
                             command=lambda mo=m: self.select_month(mo))
            btn.pack(side=tk.LEFT, padx=1)
            self.month_buttons.append(btn)
    
    def _create_search_button(self, parent):
        """検索ボタンを作成"""
        search_container = tk.Frame(parent, bg=self.colors['bg_secondary'])
        search_container.pack(side=tk.LEFT, padx=(20, 0))
        
        ttk.Button(search_container, text="🔍 検索 (Ctrl+F)", width=15, style='Accent.TButton',
                   command=lambda: SearchDialog(self.root, self)).pack()
    
    def _create_chart_button(self, parent):
        """図表ボタンを作成"""
        chart_container = tk.Frame(parent, bg=self.colors['bg_secondary'])
        chart_container.pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Button(chart_container, text="📊 図表", width=10, style='Accent.TButton',
                   command=lambda: ChartDialog(self.root, self)).pack()
    
    def _create_current_month_button(self, parent):
        """現在月表示ボタンを作成"""
        month_info = tk.Frame(parent, bg=self.colors['bg_secondary'])
        month_info.pack(side=tk.RIGHT)
        
        self.current_month_button = ttk.Button(month_info,
                                               text=f"📅 {self.current_month:02d}月",
                                               style='Month.TButton',
                                               command=self._open_monthly_data)
        self.current_month_button.pack()
    
    def _create_treeview(self, parent):
        """メインのTreeviewウィジェットを作成する"""
        if self.tree:
            return
        
        # 既存のウィジェットをクリア
        for widget in parent.winfo_children():
            widget.destroy()
        
        # 列の定義
        all_columns = self.get_all_columns()
        columns_with_button = all_columns + ["+"]
        
        # Treeviewを格納するフレーム
        tree_frame = tk.Frame(parent, bg='white', relief='solid', bd=1)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Treeviewの作成
        self.tree = ttk.Treeview(tree_frame, columns=columns_with_button, show="headings", height=25)
        self.tree.grid(row=0, column=0, sticky="nsew")

        # ヘッダーフォントの計測用オブジェクトを作成
        # FontConfig.HEADING の設定 ('Arial', 10, 'bold') を使用
        heading_font = tkfont.Font(root=self.root, font=FontConfig.HEADING)
        
        # 各列の設定
        self.default_column_widths = {}
        for i, col in enumerate(columns_with_button):
            self.tree.heading(col, text=col)
            
            # 幅の計算ロジックを変更
            if i == 0:  # 日付列
                width = TreeviewConfig.COL_WIDTH_DATE
                min_w = 50
                stretch_opt = True
            elif col == "+":  # 追加ボタン列
                width = TreeviewConfig.COL_WIDTH_BUTTON
                min_w = 40
                stretch_opt = False
            else:  # データ列
                # タイトルの文字幅を計測し、左右にパディング(+20px)を追加
                title_width = heading_font.measure(col) + 20
                # デフォルト幅(80px)とタイトル幅の大きい方を採用
                width = max(TreeviewConfig.COL_WIDTH_DATA, title_width)
                min_w = 60
                stretch_opt = True
            
            # 列設定を適用
            self.tree.column(col, anchor="center", width=width, minwidth=min_w, stretch=stretch_opt)
            self.default_column_widths[col] = width
        
        # スクロールバー(縦)
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        
        # スクロールバー(横)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=h_scrollbar.set)
        
        # Treeviewのスタイル設定
        self._configure_treeview_style()
        
        # 行のタグ設定
        self.tree.tag_configure(TreeviewConfig.TAG_TOTAL,
                                background=TreeviewConfig.BG_TOTAL,
                                font=FontConfig.HEADING)
        self.tree.tag_configure(TreeviewConfig.TAG_SUMMARY,
                                background=TreeviewConfig.BG_SUMMARY,
                                font=FontConfig.HEADING)
        self.tree.tag_configure(TreeviewConfig.TAG_NORMAL,
                                background=TreeviewConfig.BG_NORMAL)
        self.tree.tag_configure(TreeviewConfig.TAG_ODD,
                                background=TreeviewConfig.BG_ODD)
        
        # イベントバインド
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-1>", self._on_single_click)
        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<MouseWheel>", self._on_mousewheel)
        self.tree.bind("<Shift-MouseWheel>",
                       lambda e: self.tree.xview_scroll(int(-1 * (e.delta / 120)), "units"))
        self.tree.bind("<space>", self._on_space_key)
        
        # 右クリックメニュー(カスタム列用)
        self.column_context_menu = tk.Menu(self.root, tearoff=0)
        
        # ツールチップを初期化
        self.tooltip = TreeviewTooltip(self.tree, self)
    
    def _configure_treeview_style(self):
        """Treeviewのスタイルを設定"""
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure("Treeview",
                        fieldbackground="white",
                        background="white",
                        rowheight=TreeviewConfig.ROW_HEIGHT,
                        font=FontConfig.DEFAULT,
                        borderwidth=1,
                        relief="solid")
        
        style.configure("Treeview.Heading",
                        background="#e8e8e8",
                        font=FontConfig.HEADING,
                        relief="raised",
                        borderwidth=1)
        
        style.map("Treeview",
                  background=[('selected', '#0078d4')],
                  foreground=[('selected', 'yellow')])
    
    def _prev_year(self):
        """前年に移動する"""
        self.current_year -= 1
        self.update_year_display()
        self._show_month(self.current_month)
    
    def _next_year(self):
        """翌年に移動する"""
        self.current_year += 1
        self.update_year_display()
        self._show_month(self.current_month)
    
    def update_year_display(self):
        """年表示を更新"""
        if self.year_label:
            self.year_label.config(text=str(self.current_year))
    
    def select_month(self, month):
        """指定された月を選択する"""
        self.current_month = month
        self.current_month_button.config(text=f"📅 {month:02d}月")
        self._update_month_buttons()
        self._show_month(month)
    
    def _update_month_buttons(self):
        """月選択ボタンのハイライトを更新する"""
        for i, btn in enumerate(self.month_buttons, start=1):
            if i == self.current_month:
                btn.configure(style='Selected.TButton')
            else:
                btn.configure(style='Modern.TButton')
    
    def _open_monthly_data(self):
        """月間データ詳細ダイアログを開く"""
        MonthlyDataDialog(self.root, self, self.current_year, self.current_month)
    
    def get_all_columns(self):
        """全ての列(デフォルト + カスタム)を取得"""
        return DefaultColumns.ITEMS + self.data_manager.custom_columns
    
    def get_days_in_month(self):
        """現在の月の日数を取得"""
        return get_days_in_month(self.current_year, self.current_month)
    
    def _show_month(self, month):
        """指定された月のデータを表示する"""
        if not self.tree:
            return
        
        # 既存の表示をクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        all_columns = self.get_all_columns()
        days = self.get_days_in_month()
        
        # 各日のデータを表示
        for day in range(1, days + 1):
            row_values = self._calculate_day_totals(day)
            formatted_values = self._format_row_values(row_values)
            formatted_values.append("")  # +ボタン列
            
            # 奇数・偶数行で背景色を変える
            tag = TreeviewConfig.TAG_ODD if day % 2 == 1 else TreeviewConfig.TAG_NORMAL
            self.tree.insert("", "end", values=formatted_values, tags=(tag,))
        
        # 合計行
        total_row = [" 合計 "] + ["  "] * (len(all_columns) - 1) + [""]
        self.tree.insert("", "end", values=total_row, tags=(TreeviewConfig.TAG_TOTAL,))
        
        # まとめ行(収入・支出の表示)
        income_val = self._get_income_total()
        inc_str = f" {income_val} " if income_val != 0 else "  "
        summary_row = [" まとめ ", "  ", " 収入 ", inc_str, " 支出 ", "  "] + \
                      ["  "] * (len(all_columns) - 6) + [""]
        self.tree.insert("", "end", values=summary_row, tags=(TreeviewConfig.TAG_SUMMARY,))
        
        # 合計とまとめ行の値を更新
        self._update_totals()
    
    def _calculate_day_totals(self, day):
        """特定の日の各項目の合計金額を計算する"""
        all_columns = self.get_all_columns()
        totals = [""] * len(all_columns)
        totals[0] = str(day)  # 日付列
        
        # 各項目の合計を計算
        for col_index in range(1, len(all_columns)):
            dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_index}"
            data_list = self.data_manager.get_transaction_data(dict_key)
            if data_list:
                # 金額列(インデックス1)を合計
                total = sum(parse_amount(row[1]) for row in data_list if len(row) > 1)
                if total != 0:
                    totals[col_index] = str(total)
        
        return totals
    
    def _format_row_values(self, values):
        """行データを表示用にフォーマットする"""
        formatted = []
        for i, val in enumerate(values):
            if i == 0:  # 日付列
                formatted.append(f" {val} ")
            else:
                formatted.append(f" {val} " if val else "  ")
        return formatted
    
    def _get_income_total(self):
        """現在月の収入合計を取得する"""
        dict_key = f"{self.current_year}-{self.current_month}-0-3"
        data_list = self.data_manager.get_transaction_data(dict_key)
        if data_list:
            return sum(parse_amount(row[1]) for row in data_list if len(row) > 1)
        return 0
    
    def _update_totals(self):
        """合計行とまとめ行の値を更新する"""
        items = self.tree.get_children()
        if len(items) < 2:
            return
        
        total_row_id = items[-2]  # 合計行
        summary_row_id = items[-1]  # まとめ行
        all_columns = self.get_all_columns()
        cols = len(all_columns)
        
        # 各列の合計を計算
        sums = [0] * (cols - 1)
        for row_id in items[:-2]:  # 日付行のみ対象
            row_vals = self.tree.item(row_id, 'values')
            for i in range(1, cols):
                try:
                    val_str = str(row_vals[i]).strip() if i < len(row_vals) else ""
                    sums[i - 1] += int(val_str) if val_str else 0
                except (ValueError, TypeError, IndexError):
                    pass
        
        # 合計行を更新
        total_vals = list(self.tree.item(total_row_id, 'values'))
        for i in range(1, cols):
            total_vals[i] = f" {sums[i - 1]} " if sums[i - 1] != 0 else "  "
        
        while len(total_vals) <= cols:
            total_vals.append("")
        self.tree.item(total_row_id, values=total_vals)
        
        # 総支出を計算
        grand_total = sum(int(str(v).strip()) for v in total_vals[1:cols]
                          if v and str(v).strip() and str(v).strip().lstrip('-').isdigit())
        
        # まとめ行を更新
        summary_vals = list(self.tree.item(summary_row_id, 'values'))
        try:
            income_str = str(summary_vals[3]).strip() if len(summary_vals) > 3 else ""
            income_val = int(income_str) if income_str else 0
        except:
            income_val = 0
        
        # 収支差額と総支出を更新
        balance = income_val - grand_total
        summary_vals[1] = f" {balance} " if balance != 0 else "  "
        summary_vals[5] = f" {grand_total} " if grand_total != 0 else "  "
        
        while len(summary_vals) <= cols:
            summary_vals.append("")
        
        self.tree.item(summary_row_id, values=summary_vals)
    
    def _on_single_click(self, event):
        """シングルクリックイベントを処理する"""
        region = self.tree.identify_region(event.x, event.y)

        # クリックされた列IDを取得して保存（コピー＆ペースト用）
        col_id = self.tree.identify_column(event.x)
        if col_id:
            self.selected_column_id = col_id

        if region == "heading":
            if col_id:
                col_index = int(col_id[1:]) - 1
                all_columns = self.get_all_columns()
                
                if col_index == len(all_columns):  # +ボタン列
                    self._add_column()
    
    def _on_double_click(self, event):
        """ダブルクリックイベントを処理する"""
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        
        if not row_id or not col_id:
            return
        
        # ヘッダーのクリック処理
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            col_index = int(col_id[1:]) - 1
            all_columns = self.get_all_columns()
            
            if col_index == len(all_columns):  # +ボタン
                self._add_column()
            elif col_index >= len(DefaultColumns.ITEMS):  # カスタム列
                self._edit_column_name(col_index)
            return
        
        # 行の種類を判定
        items = self.tree.get_children()
        if len(items) < 2:
            return
        
        total_row_id = items[-2]
        summary_row_id = items[-1]
        
        # 合計行は編集不可
        if row_id == total_row_id:
            return
        
        # まとめ行は収入列のみ編集可能
        if row_id == summary_row_id and col_id != "#4":
            return
        
        # 日付列は編集不可
        if col_id == "#1":
            return
        
        # +ボタン列は編集不可
        col_index = int(col_id[1:]) - 1
        all_columns = self.get_all_columns()
        if col_index >= len(all_columns):
            return
        
        # 行データを取得
        row_vals = self.tree.item(row_id, 'values')
        if not row_vals:
            return
        
        # 日付と列名を特定
        if row_id == summary_row_id:
            day = 0
            col_index = 3  # 収入列
            col_name = "収入"
        else:
            try:
                day = int(row_vals[0])
            except:
                return
            col_name = self.tree.heading(col_id, "text")
        
        # 取引詳細ダイアログを開く
        dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_index}"
        TransactionDialog(self.root, self, dict_key, col_name)
    
    def _on_right_click(self, event):
        """右クリックイベントを処理する"""
        region = self.tree.identify_region(event.x, event.y)

        # ヘッダー以外（セル）での右クリックの場合、コピペメニューを表示
        if region != "heading":
            # クリック位置の行と列を選択状態にする
            row_id = self.tree.identify_row(event.y)
            col_id = self.tree.identify_column(event.x)
            
            if row_id and col_id:
                # 選択状態を更新
                self.tree.selection_set(row_id)
                self.tree.focus(row_id)
                self.selected_column_id = col_id
                
                # コンテキストメニュー作成
                cell_menu = tk.Menu(self.root, tearoff=0)
                cell_menu.add_command(label="コピー (Ctrl+C)", command=self._copy_cell)
                cell_menu.add_command(label="貼り付け (Ctrl+V)", command=self._paste_cell)
                cell_menu.add_separator()
                cell_menu.add_command(label="削除 (Delete)", command=self._delete_cell)
                cell_menu.post(event.x_root, event.y_root)
            return
        
        col_id = self.tree.identify_column(event.x)
        if not col_id:
            return
        
        col_index = int(col_id[1:]) - 1
        all_columns = self.get_all_columns()
        
        # 右クリックメニューを再作成
        self.column_context_menu = tk.Menu(self.root, tearoff=0)
        
        # カスタム列の場合は編集・削除オプションを追加
        if len(all_columns) > col_index >= len(DefaultColumns.ITEMS) and col_index != 0:
            self.selected_column_index = col_index
            self.column_context_menu.add_command(label="列名を編集", command=self._edit_column_name)
            self.column_context_menu.add_separator()
            self.column_context_menu.add_command(label="列を削除", command=self._delete_column)
            self.column_context_menu.add_separator()
        
        # すべての列で列幅リセットを利用可能
        self.column_context_menu.add_command(label="全ての列幅をリセット",
                                             command=self._reset_all_column_widths)
        
        self.column_context_menu.post(event.x_root, event.y_root)
    
    def _on_mousewheel(self, event):
        """マウスホイールイベントを処理する"""
        if event.state & 0x4:  # Ctrlキーが押されている
            self.tree.xview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            self.tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    def _on_space_key(self, event):
        """
        SPACEキーが押された時の処理
        選択中のセルの詳細ダイアログを開く
        """
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        row_id = selected_items[0]
        
        # フォーカスされている列を取得
        focus_item = self.tree.focus()
        if not focus_item:
            return
        
        # 現在選択されている列のインデックスを取得
        # （Treeviewは列のフォーカスを直接取得できないため、
        # 最後にクリックされた列を使用）
        if not hasattr(self, 'selected_column_id') or not self.selected_column_id:
            return
        
        col_id = self.selected_column_id
        col_index = int(col_id[1:]) - 1
        
        # 編集不可のセルをチェック
        items = self.tree.get_children()
        if len(items) < 2:
            return
        
        total_row_id = items[-2]
        summary_row_id = items[-1]
        
        # 合計行は編集不可
        if row_id == total_row_id:
            return
        
        # まとめ行は収入列のみ編集可能
        if row_id == summary_row_id and col_id != "#4":
            return
        
        # 日付列は編集不可
        if col_id == "#1":
            return
        
        # +ボタン列は編集不可
        all_columns = self.get_all_columns()
        if col_index >= len(all_columns):
            return
        
        # 行データを取得
        row_vals = self.tree.item(row_id, 'values')
        if not row_vals:
            return
        
        # 日付と列名を特定
        if row_id == summary_row_id:
            day = 0
            col_index = 3  # 収入列
            col_name = "収入"
        else:
            try:
                day = int(row_vals[0])
            except:
                return
            col_name = self.tree.heading(col_id, "text")
        
        # 取引詳細ダイアログを開く
        dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_index}"
        TransactionDialog(self.root, self, dict_key, col_name)
    
    def _reset_all_column_widths(self):
        """指定された列の幅をデフォルトにリセットする"""
        all_columns = self.get_all_columns() + ["+"]
        for i, col_name in enumerate(all_columns):
            col_id = f"#{i + 1}"
            if col_name in self.default_column_widths:
                self.tree.column(col_id, width=self.default_column_widths[col_name])

    def _add_column(self):
        """新しい列を追加する"""
        dialog = tk.Toplevel(self.root)
        dialog.title("列の追加")
        dialog.resizable(False, False)
        
        # ダイアログを中央に配置
        dialog_width = 300
        dialog_height = 120
        
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog_width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog_height) // 2
        
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="新しい列名を入力してください:", font=('Arial', 11)).pack(pady=10)
        
        entry = tk.Entry(dialog, font=('Arial', 11), width=25)
        entry.pack(pady=5)
        entry.focus_set()
        
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def on_ok():
            column_name = entry.get().strip()
            if column_name:
                all_columns = self.get_all_columns()
                if column_name not in all_columns:
                    self.data_manager.add_custom_column(column_name)
                    dialog.destroy()
                    self._recreate_treeview()
                    self._show_month(self.current_month)
                else:
                    messagebox.showwarning("警告", "その列名は既に存在します。", parent=dialog)
            else:
                messagebox.showwarning("警告", "列名を入力してください。", parent=dialog)
        
        tk.Button(button_frame, text="OK", command=on_ok, width=8).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="キャンセル", command=dialog.destroy, width=8).pack(side=tk.LEFT, padx=5)
        
        entry.bind('<Return>', lambda e: on_ok())
        dialog.bind('<Escape>', lambda e: dialog.destroy())

    def _recreate_treeview(self):
        """Treeviewを再作成する"""
        if self.tree:
            tree_parent = self.tree.master
            self.tree.destroy()
            self.tree = None
            self._create_treeview(tree_parent)

    def _edit_column_name(self, col_index=None):
        """カスタム列の名前を編集する"""
        if col_index is None:
            col_index = getattr(self, 'selected_column_index', None)
        
        if col_index is None or col_index < len(DefaultColumns.ITEMS):
            return
        
        custom_index = col_index - len(DefaultColumns.ITEMS)
        if custom_index >= len(self.data_manager.custom_columns):
            return
        
        old_name = self.data_manager.custom_columns[custom_index]
        
        # 編集ダイアログを表示
        dialog = tk.Toplevel(self.root)
        dialog.title("列名の編集")
        dialog.resizable(False, False)
        
        dialog_width = 300
        dialog_height = 120
        
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog_width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog_height) // 2
        
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="新しい列名を入力してください:", font=('Arial', 11)).pack(pady=10)
        
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
                all_columns = self.get_all_columns()
                if new_name not in all_columns:
                    self.data_manager.edit_custom_column(old_name, new_name)
                    dialog.destroy()
                    self._recreate_treeview()
                    self._show_month(self.current_month)
                else:
                    messagebox.showwarning("警告", "その列名は既に存在します。", parent=dialog)
            else:
                dialog.destroy()
        
        tk.Button(button_frame, text="OK", command=on_ok, width=8).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="キャンセル", command=dialog.destroy, width=8).pack(side=tk.LEFT, padx=5)
        
        entry.bind('<Return>', lambda e: on_ok())
        dialog.bind('<Escape>', lambda e: dialog.destroy())

    def _delete_column(self):
        """カスタム列を削除する"""
        col_index = getattr(self, 'selected_column_index', None)
        if col_index is None or col_index < len(DefaultColumns.ITEMS):
            return
        
        custom_index = col_index - len(DefaultColumns.ITEMS)
        if custom_index >= len(self.data_manager.custom_columns):
            return
        
        col_name = self.data_manager.custom_columns[custom_index]
        
        # 削除確認ダイアログ
        if messagebox.askyesno("確認", f"列 '{col_name}' を削除しますか?\n※この列のデータもすべて削除されます。"):
            # 列をリストから削除
            self.data_manager.delete_custom_column(col_name)
            
            # 関連するデータを削除
            self.data_manager.delete_column_data(col_index)
            
            # Treeviewを再作成して変更を反映
            self._recreate_treeview()
            self._show_month(self.current_month)

    def update_parent_cell(self, dict_key_day, col_index, new_value):
        """親画面のセル表示を更新する"""
        # キーから年月日を抽出
        y, mo, d = dict_key_day.split("-")
        y, mo, d = int(y), int(mo), int(d)
        
        # 現在表示中の年月と一致する場合のみ更新
        if (self.current_year == y) and (self.current_month == mo):
            items = self.tree.get_children()
            if len(items) < 2:
                return
            
            summary_row_id = items[-1]  # まとめ行
            
            # 該当する日付の行を検索
            for row_id in items[:-2]:  # 日付行のみ対象
                row_vals = list(self.tree.item(row_id, 'values'))
                if row_vals and str(row_vals[0]).strip() == str(d):
                    # 列数を確認して必要に応じて拡張
                    all_columns = self.get_all_columns()
                    while len(row_vals) < len(all_columns) + 1:
                        row_vals.append("")
                    
                    # 表示値をフォーマット(パディング付き)
                    display_value = "  "
                    if new_value and str(new_value).strip() != "" and str(new_value) != "0":
                        display_value = f" {new_value} "
                    
                    # 値を更新
                    row_vals[col_index] = display_value
                    self.tree.item(row_id, values=row_vals)
                    break
            
            # まとめ行(収入)の更新
            if d == 0:
                sum_vals = list(self.tree.item(summary_row_id, 'values'))
                all_columns = self.get_all_columns()
                while len(sum_vals) < len(all_columns) + 1:
                    sum_vals.append("")
                
                display_value = "  "
                if new_value and str(new_value).strip() != "" and str(new_value) != "0":
                    display_value = f" {new_value} "
                
                sum_vals[col_index] = display_value
                self.tree.item(summary_row_id, values=sum_vals)
            
            # 合計とまとめ行を再計算
            self._update_totals()

    def _copy_cell(self, event=None):
        """
        選択されているセルの詳細データをJSON形式でクリップボードにコピー
        データがない場合は表示上の値をコピー
        """
        selected_item = self.tree.selection()
        if not selected_item or not self.selected_column_id:
            return
            
        col_idx = int(self.selected_column_id[1:]) - 1
        all_columns = self.get_all_columns()
        
        # 範囲チェック
        if col_idx <= 0 or col_idx >= len(all_columns):
            return

        row_id = selected_item[0]
        items = self.tree.get_children()
        summary_row_id = items[-1] # まとめ行

        # 日付(day)を特定
        row_vals = self.tree.item(row_id, 'values')
        day = 0
        if row_id != summary_row_id:
            try:
                day = int(str(row_vals[0]).strip())
            except ValueError:
                return # 日付が取得できない行は無視

        # 内部データを取得
        dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_idx}"
        data_list = self.data_manager.get_transaction_data(dict_key)

        self.root.clipboard_clear()

        if data_list:
            # 詳細データがある場合、JSON文字列としてコピー
            json_str = json.dumps(data_list, ensure_ascii=False)
            self.root.clipboard_append(json_str)
        else:
            # データはないが表示値がある場合（稀なケース）、テキストのみコピー
            val = str(row_vals[col_idx]).strip()
            if val:
                self.root.clipboard_append(val)
        
        self.root.update()

    def _paste_cell(self, event=None):
        """
        クリップボードの値をセルに貼り付け
        JSON形式なら詳細ごと復元、数値なら新規取引として追加
        """
        selected_item = self.tree.selection()
        if not selected_item or not self.selected_column_id:
            return
        
        try:
            clipboard_text = self.root.clipboard_get()
        except tk.TclError:
            return

        # ターゲット位置の特定
        row_id = selected_item[0]
        col_idx = int(self.selected_column_id[1:]) - 1
        all_columns = self.get_all_columns()
        
        items = self.tree.get_children()
        total_row_id = items[-2]
        summary_row_id = items[-1]

        # 禁止エリア判定
        if row_id == total_row_id: return
        if row_id == summary_row_id and col_idx != 3: return
        if col_idx <= 0 or col_idx >= len(all_columns): return

        # 日付(day)を特定
        row_vals = self.tree.item(row_id, 'values')
        day = 0
        if row_id != summary_row_id:
            try:
                day = int(str(row_vals[0]).strip())
            except ValueError:
                return

        # 上書き確認
        current_val = str(row_vals[col_idx]).strip() if col_idx < len(row_vals) else ""
        if current_val and current_val != "0":
             if not messagebox.askyesno("確認", "既存のデータが存在します。\n上書きして貼り付けますか？"):
                 return

        dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_idx}"
        new_data_list = []

        # 1. JSON（詳細データ）として解析を試みる
        try:
            parsed_data = json.loads(clipboard_text)
            if isinstance(parsed_data, list):
                # データの形式チェック（リストのリストであることを期待）
                new_data_list = parsed_data
        except json.JSONDecodeError:
            pass

        # 2. JSONでなければ、単一の数値として解析（Excel等からのコピペ用）
        if not new_data_list:
            amount = parse_amount(clipboard_text)
            if amount != 0 or "0" in clipboard_text:
                new_data_list = [("貼付入力", str(amount), "")]
        
        # データがあれば保存して反映
        if new_data_list:
            self.data_manager.set_transaction_data(dict_key, new_data_list)
            
            # 合計金額を計算してUI更新
            total = sum(parse_amount(row[1]) for row in new_data_list if len(row) > 1)
            self.update_parent_cell(f"{self.current_year}-{self.current_month}-{day}", col_idx, str(total))
            
    def _delete_cell(self):
        """選択されたセルのデータを削除する"""
        selected_item = self.tree.selection()
        if not selected_item or not self.selected_column_id:
            return

        row_id = selected_item[0]
        col_idx = int(self.selected_column_id[1:]) - 1
        all_columns = self.get_all_columns()
        
        # 編集不可エリアのチェック
        items = self.tree.get_children()
        total_row_id = items[-2]
        summary_row_id = items[-1]
        
        if row_id == total_row_id: return
        if row_id == summary_row_id and col_idx != 3: return
        if col_idx <= 0 or col_idx >= len(all_columns): return

        # 確認ダイアログ
        if not messagebox.askyesno("確認", "選択されたセルのデータを削除しますか？"):
            return

        # 日付を取得
        row_vals = self.tree.item(row_id, 'values')
        day = 0
        if row_id != summary_row_id:
            try:
                day = int(str(row_vals[0]).strip())
            except ValueError:
                return

        # データを削除
        dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_idx}"
        self.data_manager.delete_transaction_data(dict_key)
        
        # UI更新（空文字にする）
        self.update_parent_cell(f"{self.current_year}-{self.current_month}-{day}", col_idx, "")
        