import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import json
import os
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime, date
import calendar
import matplotlib.font_manager as fm
import warnings


DATA_FILE = "data.json"
SETTINGS_FILE = "settings.json"

# フォント警告を無効化
warnings.filterwarnings('ignore', category=UserWarning, module='matplotlib.font_manager')


# 利用可能な日本語フォントを自動検出して設定
def setup_japanese_font():
    """日本語フォントを自動検出して設定（修正版）"""
    try:
        available_fonts = [f.name for f in fm.fontManager.ttflist]
        japanese_fonts = [
            'Yu Gothic', 'Meiryo', 'MS Gothic', 'MS PGothic', 'MS UI Gothic'
        ]

        found_font = None
        for font in japanese_fonts:
            if font in available_fonts:
                found_font = font
                break

        if found_font:
            plt.rcParams['font.family'] = found_font
            print(f"日本語フォント設定完了: {found_font}")
        else:
            plt.rcParams['font.family'] = 'DejaVu Sans'
            print("日本語フォントが見つからないため、DejaVu Sansを使用")

    except Exception as e:
        print(f"フォント設定エラー: {e}")
        plt.rcParams['font.family'] = 'DejaVu Sans'


class MonthlyDataDialog(tk.Toplevel):
    def __init__(self, parent, parent_app, year, month):
        """月間データダイアログの初期化。

        このクラスは、指定された年月の詳細データを表示するための
        モーダルダイアログウィンドウを作成します。家計簿アプリの
        メイン画面から呼び出され、その月に記録されたすべての取引
        データを一覧表示する機能を提供します。
        """
        super().__init__(parent)
        self.parent_app = parent_app  # メインアプリケーションへの参照
        self.year = year  # 表示対象の年
        self.month = month  # 表示対象の月
        self.monthly_data = []  # 月間データを格納するリスト

        # ダイアログの基本サイズを設定
        # これらの値は、データを快適に閲覧できる適切なサイズです
        dialog_width = 800
        dialog_height = 600

        # 親ウィンドウの中央に配置するための計算
        # この処理により、ダイアログが画面の中央に表示されます
        parent_x = self.master.winfo_x()
        parent_y = self.master.winfo_y()
        parent_w = self.master.winfo_width()
        parent_h = self.master.winfo_height()

        x = parent_x + (parent_w - dialog_width) // 2
        y = parent_y + (parent_h - dialog_height) // 2

        # ダイアログウィンドウの設定を適用
        self.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        self.minsize(600, 400)  # 最小サイズを設定して使いやすさを確保
        self.title(f"月間データ詳細 - {year}年{month:02d}月")
        self.configure(bg='#f0f0f0')  # 統一感のある背景色

        # モーダルダイアログとして設定
        # これにより、このダイアログが開いている間は親ウィンドウの操作が制限されます
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)  # ユーザーがサイズ調整できるように設定

        # ウィジェットを作成してデータを読み込み
        self._create_widgets()
        self._load_monthly_data()

        # ダイアログを前面に表示
        # これにより確実にユーザーの注意を引くことができます
        self.lift()
        self.focus_force()

    def _create_widgets(self):
        """ダイアログ内のウィジェットを作成。

        このメソッドは、月間データを表示するための全てのUI要素を
        構築します。グリッドレイアウトを使用して、レスポンシブな
        デザインを実現しています。
        """
        # グリッドの重み設定
        # row=1に重みを設定することで、データ表示エリアが伸縮可能になります
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # タイトル部分の作成
        # ユーザーが現在閲覧している期間を明確に示します
        title_frame = tk.Frame(self, bg='#f0f0f0')
        title_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        title_label = tk.Label(title_frame, text=f"{self.year}年{self.month:02d}月の詳細データ",
                               font=('Arial', 16, 'bold'), bg='#f0f0f0')
        title_label.pack()

        # 結果表示部分の作成
        # メインのデータ表示エリアで、リサイズ可能な設計です
        result_frame = tk.Frame(self, bg='#f0f0f0')
        result_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        result_frame.grid_rowconfigure(0, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)

        # データ表示用のTreeviewウィジェット
        # 表形式でデータを整理して表示するため、5つの列を定義します
        columns = ["年月日", "項目", "取引先", "金額", "詳細"]
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=15)

        # 各列のヘッダーと幅を設定
        # 列幅は内容に応じて最適化されています
        self.result_tree.heading("年月日", text="年月日")
        self.result_tree.heading("項目", text="項目")
        self.result_tree.heading("取引先", text="取引先")
        self.result_tree.heading("金額", text="金額")
        self.result_tree.heading("詳細", text="詳細")

        # 列の幅設定 - データの特性に合わせて調整
        self.result_tree.column("年月日", anchor="center", width=100, minwidth=80)
        self.result_tree.column("項目", anchor="center", width=120, minwidth=100)
        self.result_tree.column("取引先", anchor="center", width=150, minwidth=120)
        self.result_tree.column("金額", anchor="center", width=100, minwidth=80)
        self.result_tree.column("詳細", anchor="center", width=200, minwidth=150)

        # Treeviewをグリッドに配置
        self.result_tree.grid(row=0, column=0, sticky="nsew")

        # 縦スクロールバーの追加
        # 大量のデータでも快適にスクロールできるようになります
        result_scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        result_scrollbar.grid(row=0, column=1, sticky="ns")
        self.result_tree.configure(yscrollcommand=result_scrollbar.set)

        # 横スクロールバーの追加
        # 列が多い場合や画面が小さい場合に対応します
        h_scrollbar = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.result_tree.configure(xscrollcommand=h_scrollbar.set)

        # 統計情報表示エリア
        # 月間の合計金額や平均などの統計を表示します
        stats_frame = tk.Frame(self, bg='#f0f0f0')
        stats_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=5)

        self.stats_label = tk.Label(stats_frame, text="", font=('Arial', 12, 'bold'), bg='#f0f0f0')
        self.stats_label.pack()

        # 結果カウント表示
        # ユーザーが現在のデータ量を把握できるようにします
        self.result_label = tk.Label(self, text="データ: 0 件", font=('Arial', 10), bg='#f0f0f0', fg='#666666')
        self.result_label.grid(row=3, column=0, sticky="w", padx=10, pady=(5, 10))

        # 閉じるボタン
        # ダイアログを終了するためのボタンです
        close_btn = tk.Button(self, text="閉じる", font=('Arial', 12),
                              bg='#f44336', fg='white', relief='raised', bd=2,
                              activebackground='#d32f2f', command=self.destroy)
        close_btn.grid(row=4, column=0, pady=(0, 10), ipady=5)

        # キーボードショートカットの設定
        # Escapeキーでダイアログを閉じることができます
        self.bind('<Escape>', lambda e: self.destroy())

        # マウスホイールでのスクロール機能
        # より直感的な操作を可能にします
        def on_mousewheel(event):
            self.result_tree.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.result_tree.bind("<MouseWheel>", on_mousewheel)

    def _load_monthly_data(self):
        """月間データを読み込んで表示。

        このメソッドは、parent_appのchild_dataから指定された年月の
        データを抽出し、整理してTreeviewに表示します。また、統計
        情報も同時に計算して表示します。
        """
        # 既存の表示内容をクリア
        # 新しいデータを表示する前に、前回の内容を削除します
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # データ格納用変数の初期化
        self.monthly_data = []
        total_amount = 0  # 月間合計金額
        total_count = 0  # 取引件数

        # child_dataから該当月のデータを検索
        # parent_appに格納されている全データから必要な月のデータのみを抽出します
        for dict_key, data_list in self.parent_app.child_data.items():
            try:
                # キーを解析して年月日と項目インデックスを取得
                # キーの形式: "年-月-日-列インデックス"
                parts = dict_key.split("-")
                if len(parts) == 4:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])
                    col_index = int(parts[3])

                    # 指定された年月と一致するデータのみを処理
                    if year == self.year and month == self.month:
                        # 項目名を取得
                        # 列インデックスから実際の項目名を特定します
                        all_columns = self.parent_app.default_columns + self.parent_app.custom_columns
                        if col_index < len(all_columns):
                            column_name = all_columns[col_index]
                        else:
                            column_name = f"列{col_index}"  # 範囲外の場合の安全な処理

                        # 日付文字列の作成
                        date_str = f"{year}/{month:02d}/{day:02d}"

                        # 各取引データを処理
                        # data_listには複数の取引が含まれている可能性があります
                        for row in data_list:
                            if len(row) >= 3:  # 最低限必要なデータが揃っている場合のみ処理
                                # 各フィールドのデータを安全に取得
                                partner = str(row[0]).strip() if row[0] else ""
                                amount_str = str(row[1]).strip() if row[1] else ""
                                detail = str(row[2]).strip() if row[2] else ""

                                # 金額の数値変換処理
                                # 文字列から数値を抽出して統計計算に使用します
                                amount_value = 0
                                if amount_str:
                                    try:
                                        # カンマや通貨記号を除去して数値に変換
                                        clean_amount = amount_str.replace(',', '').replace('¥', '').strip()
                                        if clean_amount:
                                            amount_value = int(clean_amount)
                                    except ValueError:
                                        # 変換できない場合は0として処理
                                        amount_value = 0

                                # 結果データの構造化
                                # 表示と統計計算の両方に使用するデータ構造を作成します
                                result = {
                                    'date': date_str,
                                    'column': column_name,
                                    'partner': partner,
                                    'amount': amount_str,
                                    'detail': detail,
                                    'amount_value': amount_value,
                                    'sort_key': (year, month, day, col_index)  # ソート用キー
                                }
                                self.monthly_data.append(result)
                                total_amount += amount_value
                                total_count += 1

            except (ValueError, IndexError) as e:
                # データ解析でエラーが発生した場合のログ出力
                print(f"キー解析エラー: {dict_key}, エラー: {e}")
                continue

        # 結果を日付順にソート
        # ユーザーが時系列でデータを確認できるように整理します
        self.monthly_data.sort(key=lambda x: x['sort_key'])

        # 結果をTreeviewに表示
        # 整理されたデータを表形式で表示します
        for result in self.monthly_data:
            values = [
                result['date'],
                result['column'],
                result['partner'],
                result['amount'],
                result['detail']
            ]
            self.result_tree.insert("", "end", values=values)

        # 統計情報の計算と表示
        # ユーザーに有用な統計データを提供します
        if total_count > 0:
            avg_amount = total_amount / total_count
            self.stats_label.config(text=f"合計金額: ¥{total_amount:,} | 平均金額: ¥{avg_amount:.0f}")
        else:
            self.stats_label.config(text="")

        # 結果カウントの更新
        # データ件数をユーザーに表示します
        self.result_label.config(text=f"データ: {len(self.monthly_data)} 件")

        # 処理完了のログ出力
        print(f"月間データ読み込み完了: {len(self.monthly_data)} 件の結果")


class SearchDialog(tk.Toplevel):
    def __init__(self, parent, parent_app):
        """検索ダイアログの初期化。

        このクラスは家計簿アプリケーションの検索機能を担当する重要なコンポーネントです。
        ユーザーが大量の家計データの中から特定の情報を素早く見つけるための
        直感的なインターフェースを提供します。検索は全文検索方式で、
        取引先名、金額、詳細情報のすべてを対象として実行されます。
        """
        super().__init__(parent)
        self.parent_app = parent_app  # メインアプリケーションへの参照を保持
        self.search_results = []  # 検索結果を格納するリスト

        # ダイアログウィンドウのサイズ設定
        # 検索結果を十分に表示できる適切なサイズを設定しています
        dialog_width = 800
        dialog_height = 600

        # 親ウィンドウの中央にダイアログを配置する計算
        # この処理により、ユーザーの注意を適切に検索ダイアログに向けることができます
        parent_x = self.master.winfo_x()
        parent_y = self.master.winfo_y()
        parent_w = self.master.winfo_width()
        parent_h = self.master.winfo_height()

        x = parent_x + (parent_w - dialog_width) // 2
        y = parent_y + (parent_h - dialog_height) // 2

        # ダイアログの基本設定を適用
        self.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        self.minsize(600, 400)  # 最小サイズの制限により使いやすさを確保
        self.title("検索")
        self.configure(bg='#f0f0f0')

        # モーダルダイアログとして設定
        # これにより検索作業中は他の操作がブロックされ、集中して検索できます
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)

        # ユーザーインターフェースの構築
        self._create_widgets()

        # ダイアログを前面に表示して即座に使用可能な状態にします
        self.lift()
        self.focus_force()

    def _create_widgets(self):
        """検索ダイアログのユーザーインターフェースを構築。

        このメソッドは検索機能に必要な全てのUI要素を作成します。
        検索入力エリア、結果表示エリア、操作ボタンを論理的に配置し、
        ユーザーが直感的に操作できるインターフェースを実現しています。
        """
        # グリッドレイアウトの重み設定
        # row=1（結果表示エリア）に重みを設定することで、
        # ウィンドウサイズが変更されても検索結果が適切に表示されます
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 検索入力エリアの構築
        # ユーザーが検索キーワードを入力し、検索を実行するための領域です
        search_frame = tk.Frame(self, bg='#f0f0f0')
        search_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        search_frame.grid_columnconfigure(1, weight=1)  # 入力フィールドを伸縮可能に

        # 検索フィールドのラベル
        # ユーザーに何を入力すべきかを明確に示します
        tk.Label(search_frame, text="検索文字列:", font=('Arial', 12), bg='#f0f0f0').grid(
            row=0, column=0, padx=(0, 10), sticky="w")

        # 検索文字列入力フィールド
        # ユーザーが検索したいキーワードを入力する主要な要素です
        self.search_entry = tk.Entry(search_frame, font=('Arial', 12), width=30)
        self.search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        self.search_entry.focus_set()  # ダイアログ開設時に即座に入力可能な状態に

        # 検索実行ボタン
        # 明確な視覚的フィードバックを提供するスタイリングを施しています
        search_btn = tk.Button(search_frame, text="検索", font=('Arial', 12),
                               bg='#2196f3', fg='white', relief='raised', bd=2,
                               activebackground='#1976d2', command=self._search)
        search_btn.grid(row=0, column=2, padx=(0, 10), ipady=3)

        # 結果クリアボタン
        # 前回の検索結果をリセットして新しい検索を開始できます
        clear_btn = tk.Button(search_frame, text="クリア", font=('Arial', 12),
                              bg='#ff9800', fg='white', relief='raised', bd=2,
                              activebackground='#f57c00', command=self._clear_results)
        clear_btn.grid(row=0, column=3, ipady=3)

        # 検索結果表示エリアの構築
        # 検索で見つかったデータを整理して表示するための領域です
        result_frame = tk.Frame(self, bg='#f0f0f0')
        result_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        result_frame.grid_rowconfigure(0, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)

        # 検索結果表示用のTreeview
        # 家計データの構造に合わせて5つの列を定義しています
        columns = ["年月日", "項目", "取引先", "金額", "詳細"]
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=15)

        # 各列のヘッダー設定
        # ユーザーがデータの内容を理解しやすいように明確なラベルを設定
        self.result_tree.heading("年月日", text="年月日")
        self.result_tree.heading("項目", text="項目")
        self.result_tree.heading("取引先", text="取引先")
        self.result_tree.heading("金額", text="金額")
        self.result_tree.heading("詳細", text="詳細")

        # 列幅の最適化設定
        # 各データタイプの特性に応じて読みやすい幅を設定しています
        self.result_tree.column("年月日", anchor="center", width=100, minwidth=80)
        self.result_tree.column("項目", anchor="center", width=120, minwidth=100)
        self.result_tree.column("取引先", anchor="center", width=150, minwidth=120)
        self.result_tree.column("金額", anchor="center", width=100, minwidth=80)
        self.result_tree.column("詳細", anchor="center", width=200, minwidth=150)

        # Treeviewをレイアウトに配置
        self.result_tree.grid(row=0, column=0, sticky="nsew")

        # 縦スクロールバーの追加
        # 多数の検索結果でも快適にナビゲートできるようにします
        result_scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        result_scrollbar.grid(row=0, column=1, sticky="ns")
        self.result_tree.configure(yscrollcommand=result_scrollbar.set)

        # 横スクロールバーの追加
        # 列が多い場合や小さな画面での表示に対応します
        h_scrollbar = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.result_tree.configure(xscrollcommand=h_scrollbar.set)

        # 検索結果カウンター
        # ユーザーが検索の成果を数値的に把握できるようにします
        self.result_label = tk.Label(self, text="検索結果: 0 件", font=('Arial', 10), bg='#f0f0f0', fg='#666666')
        self.result_label.grid(row=2, column=0, sticky="w", padx=10, pady=(5, 10))

        # ダイアログ終了ボタン
        # ユーザーが検索作業を完了した後にダイアログを閉じることができます
        close_btn = tk.Button(self, text="閉じる", font=('Arial', 12),
                              bg='#f44336', fg='white', relief='raised', bd=2,
                              activebackground='#d32f2f', command=self.destroy)
        close_btn.grid(row=3, column=0, pady=(0, 10), ipady=5)

        # キーボードショートカットの設定
        # より効率的な操作を可能にするキーバインドを追加しています
        self.search_entry.bind('<Return>', lambda e: self._search())  # Enterキーで検索実行
        self.bind('<Escape>', lambda e: self.destroy())  # Escapeキーで終了
        self.bind('<Control-f>', lambda e: self.search_entry.focus_set())  # Ctrl+Fで検索フィールドにフォーカス

        # マウスホイールによるスクロール機能
        # より直感的な操作環境を提供します
        def on_mousewheel(event):
            self.result_tree.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.result_tree.bind("<MouseWheel>", on_mousewheel)

    def _search(self):
        """検索処理の実行。

        この関数は家計簿データ全体に対して全文検索を実行します。
        検索アルゴリズムは大文字小文字を区別せず、部分一致による
        柔軟なマッチングを行います。検索対象は取引先名、金額、詳細情報の
        すべてのフィールドに及び、ユーザーが求める情報を効率的に発見できます。
        """
        # 検索文字列の取得と検証
        search_text = self.search_entry.get().strip()
        if not search_text:
            messagebox.showwarning("警告", "検索文字列を入力してください。")
            return

        print(f"検索開始: '{search_text}'")

        # 前回の検索結果をクリア
        # 新しい検索結果のみを表示するために既存の表示をリセットします
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 検索結果格納用変数の初期化
        self.search_results = []
        search_text_lower = search_text.lower()  # 大文字小文字を区別しない検索のため

        # 全データに対する検索処理
        # parent_appのchild_dataから該当するデータを抽出します
        for dict_key, data_list in self.parent_app.child_data.items():
            try:
                # データキーの解析
                # キー形式: "年-月-日-列インデックス" からメタデータを抽出
                parts = dict_key.split("-")
                if len(parts) == 4:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])
                    col_index = int(parts[3])

                    # 項目名の特定
                    # 列インデックスから実際の項目名を取得します
                    all_columns = self.parent_app.default_columns + self.parent_app.custom_columns
                    if col_index < len(all_columns):
                        column_name = all_columns[col_index]
                    else:
                        column_name = f"列{col_index}"  # 範囲外の場合の安全な処理

                    # 日付文字列の生成
                    date_str = f"{year}/{month:02d}/{day:02d}"

                    # 各取引データの検索処理
                    # data_listに含まれる個別の取引に対して検索を実行
                    for row in data_list:
                        if len(row) >= 3:  # 必要なデータフィールドが存在する場合のみ処理
                            # 各フィールドのデータを安全に取得
                            partner = str(row[0]).strip() if row[0] else ""
                            amount = str(row[1]).strip() if row[1] else ""
                            detail = str(row[2]).strip() if row[2] else ""

                            # 検索条件のマッチング判定
                            # 全てのフィールドに対して部分一致検索を実行
                            if (search_text_lower in partner.lower() or
                                    search_text_lower in amount.lower() or
                                    search_text_lower in detail.lower()):
                                # マッチしたデータの構造化
                                # 表示用の統一されたデータ構造を作成
                                result = {
                                    'date': date_str,
                                    'column': column_name,
                                    'partner': partner,
                                    'amount': amount,
                                    'detail': detail,
                                    'sort_key': (year, month, day, col_index)  # ソート用のキー
                                }
                                self.search_results.append(result)

            except (ValueError, IndexError) as e:
                # データ解析エラーの処理
                # 不正なデータ形式による例外をログに記録し、処理を継続
                print(f"キー解析エラー: {dict_key}, エラー: {e}")
                continue

        # 検索結果の時系列ソート
        # ユーザーが結果を時間軸で理解しやすいように整理
        self.search_results.sort(key=lambda x: x['sort_key'])

        # 検索結果の表示処理
        # 整理された結果をTreeviewに表示
        for result in self.search_results:
            values = [
                result['date'],
                result['column'],
                result['partner'],
                result['amount'],
                result['detail']
            ]
            self.result_tree.insert("", "end", values=values)

        # 結果カウンターの更新
        # ユーザーに検索の成果を数値で示します
        self.result_label.config(text=f"検索結果: {len(self.search_results)} 件")

        # 処理完了のログ出力
        print(f"検索完了: {len(self.search_results)} 件の結果")

    def _clear_results(self):
        """検索結果とフィールドのクリア処理。

        この関数は新しい検索を開始する前に、前回の検索状態を
        完全にリセットします。ユーザーが混乱することなく、
        新鮮な状態で次の検索を開始できるようにします。
        """
        # 表示されている検索結果の削除
        # Treeview内の全てのアイテムを削除してクリーンな状態にします
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 内部データ構造のリセット
        self.search_results = []

        # ユーザーインターフェースの状態リセット
        self.result_label.config(text="検索結果: 0 件")  # カウンター表示をリセット
        self.search_entry.delete(0, tk.END)  # 入力フィールドをクリア
        self.search_entry.focus_set()  # 入力フィールドにフォーカスを設定


class ChartDialog(tk.Toplevel):
    def __init__(self, parent, parent_app):
        """図表ダイアログの初期化。"""
        super().__init__(parent)
        self.parent_app = parent_app

        # ダイアログの設定（サイズを小さく修正）
        dialog_width = 800  # 元: 1000
        dialog_height = 600  # 元: 700

        # 親ウィンドウの中央に配置
        parent_x = self.master.winfo_x()
        parent_y = self.master.winfo_y()
        parent_w = self.master.winfo_width()
        parent_h = self.master.winfo_height()

        x = parent_x + (parent_w - dialog_width) // 2
        y = parent_y + (parent_h - dialog_height) // 2

        self.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        self.minsize(600, 450)  # 元: minsize(800, 600)
        self.title("項目別月間推移グラフ")
        self.configure(bg='#f0f0f0')

        # モーダルに設定
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)

        # 初期表示を総支出に変更（-1は総支出、-2は総収入を表す特別な値）
        self.current_column_index = -1  # 元: 0
        self.figure = None
        self.canvas = None

        self._create_widgets()
        self._update_chart()

        # 強制的にウィンドウを前面に表示
        self.lift()
        self.focus_force()

    def _create_widgets(self):
        """ウィジェットを作成。"""
        # グリッドの重み設定
        self.grid_rowconfigure(2, weight=1)  # チャートエリアの行番号を2に変更
        self.grid_columnconfigure(0, weight=1)

        # タイトル部分
        title_frame = tk.Frame(self, bg='#f0f0f0')
        title_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        title_label = tk.Label(title_frame, text="項目別月間推移グラフ",
                               font=('Arial', 16, 'bold'), bg='#f0f0f0')
        title_label.pack()

        # 総支出・総収入ボタンフレーム（新規追加）
        summary_frame = tk.Frame(self, bg='#f0f0f0')
        summary_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)

        # 総支出ボタン
        total_expense_btn = tk.Button(summary_frame, text="総支出",
                                      font=('Arial', 12, 'bold'),
                                      bg='#f44336', fg='white',
                                      relief='raised', bd=2,
                                      activebackground='#d32f2f',
                                      command=lambda: self._select_summary_tab(-1))
        total_expense_btn.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.Y)

        # 総収入ボタン
        total_income_btn = tk.Button(summary_frame, text="総収入",
                                     font=('Arial', 12, 'bold'),
                                     bg='#4caf50', fg='white',
                                     relief='raised', bd=2,
                                     activebackground='#45a049',
                                     command=lambda: self._select_summary_tab(-2))
        total_income_btn.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.Y)

        # ボタンの参照を保存（ハイライト用）
        self.total_expense_btn = total_expense_btn
        self.total_income_btn = total_income_btn

        # タブフレーム
        tab_frame = tk.Frame(self, bg='#f0f0f0')
        tab_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 5))  # 行番号を2に変更

        # タブボタンを作成
        all_columns = self.parent_app.default_columns + self.parent_app.custom_columns
        self.tab_buttons = []

        # スクロール可能なタブフレーム
        tab_canvas = tk.Canvas(tab_frame, height=40, bg='#f0f0f0', highlightthickness=0)
        tab_canvas.pack(fill=tk.X)

        tab_inner_frame = tk.Frame(tab_canvas, bg='#f0f0f0')
        tab_canvas.create_window((0, 0), window=tab_inner_frame, anchor="nw")

        for i, col_name in enumerate(all_columns[1:], start=1):  # 日付列をスキップ
            btn = tk.Button(tab_inner_frame, text=col_name,
                            font=('Arial', 10),
                            bg='#e0e0e0', fg='black',
                            relief='raised', bd=2,
                            command=lambda idx=i: self._select_tab(idx))
            btn.pack(side=tk.LEFT, padx=2, pady=5, fill=tk.Y)
            self.tab_buttons.append(btn)

        # 初期選択を総支出に変更（ボタンの色を更新）
        self._update_button_colors()

        # タブフレームの幅を更新
        tab_inner_frame.update_idletasks()
        tab_canvas.configure(scrollregion=tab_canvas.bbox("all"))

        # グラフ表示エリア（行番号を3に変更）
        chart_frame = tk.Frame(self, bg='#f0f0f0')
        chart_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=5)
        chart_frame.grid_rowconfigure(0, weight=1)
        chart_frame.grid_columnconfigure(0, weight=1)

        # matplotlib設定
        plt.style.use('default')

        # 日本語フォント設定（修正版）
        try:
            # 利用可能なフォントを取得
            available_fonts = [f.name for f in fm.fontManager.ttflist]

            # 日本語フォントの候補（優先順位順）
            japanese_fonts = [
                'Yu Gothic', 'Meiryo', 'MS Gothic', 'MS PGothic', 'MS UI Gothic',
                'Hiragino Sans', 'Hiragino Kaku Gothic Pro', 'Hiragino Kaku Gothic ProN',
                'Takao', 'IPAexGothic', 'IPAPGothic', 'VL PGothic', 'Noto Sans CJK JP'
            ]

            # 利用可能な日本語フォントを検索
            found_font = None
            for font in japanese_fonts:
                if font in available_fonts:
                    found_font = font
                    break

            if found_font:
                plt.rcParams['font.family'] = found_font
                print(f"日本語フォントを設定: {found_font}")
            else:
                # 日本語フォントが見つからない場合はデフォルトを使用
                plt.rcParams['font.family'] = 'DejaVu Sans'
                print("日本語フォントが見つからないため、DejaVu Sansを使用")

        except Exception as e:
            print(f"フォント設定エラー: {e}")
            plt.rcParams['font.family'] = 'DejaVu Sans'

        # 図とキャンバスを作成（サイズを小さく修正）
        self.figure = plt.Figure(figsize=(10, 5), dpi=80)  # 元: figsize=(12, 6), dpi=100
        self.canvas = FigureCanvasTkAgg(self.figure, chart_frame)
        self.canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew")

        # 閉じるボタン（行番号を4に変更）
        close_btn = tk.Button(self, text="閉じる", font=('Arial', 12),
                              bg='#f44336', fg='white', relief='raised', bd=2,
                              activebackground='#d32f2f', command=self.destroy)
        close_btn.grid(row=4, column=0, pady=10, ipady=5)

        # キーボードショートカット
        self.bind('<Escape>', lambda e: self.destroy())

    def _select_summary_tab(self, summary_type):
        """総支出・総収入タブを選択。"""
        self.current_column_index = summary_type
        self._update_button_colors()
        self._update_chart()

    def _select_tab(self, column_index):
        """通常タブを選択。"""
        self.current_column_index = column_index
        self._update_button_colors()
        self._update_chart()

    def _update_button_colors(self):
        """ボタンの色を更新。"""
        # 総支出・総収入ボタンの色をリセット
        self.total_expense_btn.config(bg='#f44336', fg='white')
        self.total_income_btn.config(bg='#4caf50', fg='white')

        # タブボタンの色をリセット
        for btn in self.tab_buttons:
            btn.config(bg='#e0e0e0', fg='black')

        # 現在選択されているものをハイライト
        if self.current_column_index == -1:  # 総支出
            self.total_expense_btn.config(bg='#d32f2f', fg='white')
        elif self.current_column_index == -2:  # 総収入
            self.total_income_btn.config(bg='#45a049', fg='white')
        elif self.current_column_index > 0 and self.current_column_index - 1 < len(self.tab_buttons):
            self.tab_buttons[self.current_column_index - 1].config(bg='#2196f3', fg='white')

    def _update_chart(self):
        """グラフを更新。"""
        if not self.figure:
            return

        # 図をクリア
        self.figure.clear()

        # データを収集
        if self.current_column_index == -1:  # 総支出
            monthly_data = self._collect_total_expense_data()
            title = "月間総支出の推移"
            ylabel = "支出額 (円)"
            color = '#f44336'
        elif self.current_column_index == -2:  # 総収入
            monthly_data = self._collect_total_income_data()
            title = "月間総収入の推移"
            ylabel = "収入額 (円)"
            color = '#4caf50'
        else:
            monthly_data = self._collect_monthly_data()
            all_columns = self.parent_app.default_columns + self.parent_app.custom_columns
            column_name = all_columns[self.current_column_index] if self.current_column_index < len(
                all_columns) else "不明"
            title = f'{column_name} の月間推移'
            ylabel = "金額 (円)"
            color = '#2196f3'

        if not monthly_data:
            # データがない場合
            ax = self.figure.add_subplot(111)
            ax.text(0.5, 0.5, 'データがありません',
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax.transAxes, fontsize=16)
            ax.set_title('データなし')
            self.canvas.draw()
            return

        # グラフを描画
        ax = self.figure.add_subplot(111)

        # データを日付順にソート
        sorted_data = sorted(monthly_data.items())
        dates = [item[0] for item in sorted_data]
        amounts = [item[1] for item in sorted_data]

        # 折れ線グラフを描画
        ax.plot(dates, amounts, marker='o', linewidth=2, markersize=6, color=color)

        # グラフの装飾
        ax.set_title(title, fontsize=14, fontweight='bold')
        ax.set_xlabel('月', fontsize=12)
        ax.set_ylabel(ylabel, fontsize=12)

        # Y軸のフォーマット（カンマ区切り）
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{int(x):,}'))

        # X軸のフォーマット
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m'))
        ax.xaxis.set_major_locator(mdates.MonthLocator())

        # グリッドを表示
        ax.grid(True, alpha=0.3)

        # 日付ラベルを斜めに
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=45)

        # レイアウトを調整
        self.figure.tight_layout()

        # キャンバスを更新
        self.canvas.draw()

    def _collect_total_expense_data(self):
        """月間総支出データを収集。"""
        monthly_totals = {}

        # parent_table_dataから各月のデータを収集
        for date_key, row_data in self.parent_app.parent_table_data.items():
            try:
                parts = date_key.split("-")
                if len(parts) >= 3:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])

                    # まとめ行（day=0）はスキップ
                    if day == 0:
                        continue

                    month_key = date(year, month, 1)
                    if month_key not in monthly_totals:
                        monthly_totals[month_key] = 0

                    # 支出項目（日付列以外）を合計
                    all_columns = self.parent_app.default_columns + self.parent_app.custom_columns
                    for col_idx in range(1, len(all_columns)):  # 日付列をスキップ
                        if len(row_data) > col_idx:
                            amount_str = str(row_data[col_idx]).strip()
                            if amount_str:
                                try:
                                    amount = int(amount_str.replace(',', '').replace('¥', ''))
                                    monthly_totals[month_key] += amount
                                except ValueError:
                                    pass
            except (ValueError, IndexError):
                continue

        return monthly_totals

    def _collect_total_income_data(self):
        """月間総収入データを収集。"""
        monthly_totals = {}

        # parent_table_dataから各月のまとめ行の収入データを収集
        for date_key, row_data in self.parent_app.parent_table_data.items():
            try:
                parts = date_key.split("-")
                if len(parts) >= 3:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])

                    # まとめ行（day=0）のみ対象
                    if day == 0:
                        month_key = date(year, month, 1)

                        # 収入は列インデックス3（0ベース）に格納
                        if len(row_data) > 3:
                            income_str = str(row_data[3]).strip()
                            if income_str:
                                try:
                                    income = int(income_str.replace(',', '').replace('¥', ''))
                                    monthly_totals[month_key] = income
                                except ValueError:
                                    pass
            except (ValueError, IndexError):
                continue

        return monthly_totals

    def _collect_monthly_data(self):
        """月間データを収集（既存メソッド）。"""
        monthly_totals = {}

        # parent_table_dataから各月のデータを収集
        for date_key, row_data in self.parent_app.parent_table_data.items():
            try:
                parts = date_key.split("-")
                if len(parts) >= 3:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])

                    # まとめ行（day=0）はスキップ
                    if day == 0:
                        continue

                    # 指定された列のデータを取得
                    if len(row_data) > self.current_column_index:
                        amount_str = str(row_data[self.current_column_index]).strip()
                        if amount_str:
                            try:
                                amount = int(amount_str.replace(',', '').replace('¥', ''))
                                month_key = date(year, month, 1)
                                if month_key not in monthly_totals:
                                    monthly_totals[month_key] = 0
                                monthly_totals[month_key] += amount
                            except ValueError:
                                pass
            except (ValueError, IndexError):
                continue

        return monthly_totals


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

        # 検索のキーボードショートカットを追加
        self.root.bind('<Control-f>', self._open_search_dialog)

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

    def _open_search_dialog(self, event=None):
        """検索ダイアログを開く。"""
        SearchDialog(self.root, self)

    def _open_chart_dialog(self, event=None):
        """図表ダイアログを開く。"""
        ChartDialog(self.root, self)

    def _open_monthly_data_dialog(self, event=None):
        """月間データダイアログを開く。"""
        MonthlyDataDialog(self.root, self, self.current_year, self.current_month)

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

        # 月ボタン（クリック可能）
        style.configure('Month.TButton',
                        background=self.colors['accent'],
                        foreground='#ffffff',
                        borderwidth=0,
                        focuscolor='none',
                        font=('Segoe UI', 12, 'bold'),
                        relief='flat')

        style.map('Month.TButton',
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

        # 検索ボタンを追加（月選択ボタンの右側）
        search_container = tk.Frame(header_inner, bg=self.colors['bg_secondary'])
        search_container.pack(side=tk.LEFT, padx=(20, 0))

        search_btn = ttk.Button(search_container, text="🔍 検索 (Ctrl+F)", width=15, style='Accent.TButton',
                                command=self._open_search_dialog)
        search_btn.pack()

        # 図表ボタンを追加（検索ボタンの右側）
        chart_container = tk.Frame(header_inner, bg=self.colors['bg_secondary'])
        chart_container.pack(side=tk.LEFT, padx=(10, 0))

        chart_btn = ttk.Button(chart_container, text="📊 図表", width=10, style='Accent.TButton',
                               command=self._open_chart_dialog)
        chart_btn.pack()

        # 現在月表示（右側、クリック可能なボタンに変更）
        month_info_frame = tk.Frame(header_inner, bg=self.colors['bg_secondary'])
        month_info_frame.pack(side=tk.RIGHT)

        # 月ボタン（クリック可能）
        self.current_month_button = ttk.Button(month_info_frame,
                                               text=f"📅 {self.current_month:02d}月",
                                               style='Month.TButton',
                                               command=self._open_monthly_data_dialog)
        self.current_month_button.pack()

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

        # 各列の設定（境界線効果を強化）
        for i, col in enumerate(all_columns):
            # ヘッダー設定
            self.tree.heading(col, text=col)

            # 列設定（境界線効果のためのパディングと幅調整）
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

        # スタイル設定（列の境界線を強化）
        style = ttk.Style()

        # テーマを設定して境界線を表示
        style.theme_use('clam')

        # Treeviewの基本スタイル（グリッド線効果）
        style.configure("Treeview",
                        fieldbackground="white",
                        background="white",
                        rowheight=27,
                        font=('Arial', 9),
                        borderwidth=1,
                        relief="solid")

        # ヘッダーのスタイル（境界線強化）
        style.configure("Treeview.Heading",
                        background="#e8e8e8",
                        font=('Arial', 10, 'bold'),
                        relief="raised",
                        borderwidth=1)

        # 選択時の色を調整
        style.map("Treeview",
                  background=[('selected', '#0078d4')],
                  foreground=[('selected', 'white')])

        # 列セパレーターのスタイルを追加
        style.configure("Treeview.Separator",
                        background="#d0d0d0",
                        borderwidth=1)

        # タグ設定（背景色のみで区別）
        self.tree.tag_configure("TOTAL",
                                background="#fff3cd",
                                font=('Arial', 10, 'bold'))

        self.tree.tag_configure("SUMMARY",
                                background="#d4edda",
                                font=('Arial', 10, 'bold'))

        # 通常行のタグ（白背景）
        self.tree.tag_configure("normal_row",
                                background="white")

        # 奇数行用タグ（薄いグレー背景でグリッド効果）
        self.tree.tag_configure("odd_row",
                                background="#f8f9fa")

        # 列の境界線効果を強化するためのバインド
        self._setup_column_separators()

        # イベントバインド
        self.tree.bind("<Double-1>", self._on_parent_double_click)
        self.tree.bind("<Button-1>", self._on_single_click)
        self.tree.bind("<Button-3>", self._on_header_right_click)

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

    def _setup_column_separators(self):
        """列の境界線を設定する。"""
        try:
            # tkinterの内部スタイルを使用して列境界線を強化
            style = ttk.Style()

            # カスタムスタイルの作成
            style.element_create("Custom.Treeheading.border", "from", "default")
            style.layout("Custom.Treeview.Heading", [
                ("Custom.Treeheading.cell", {'sticky': 'nswe'}),
                ("Custom.Treeheading.border", {'children': [
                    ("Custom.Treeheading.padding", {'children': [
                        ("Custom.Treeheading.text", {'sticky': 'we'})
                    ]})
                ], 'sticky': 'nswe'})
            ])

            # セルの境界線を強化
            style.configure("Custom.Treeview.Heading",
                            relief="solid",
                            borderwidth=1,
                            background="#e8e8e8")

            # TreeviewのセルにもCustomスタイルを適用
            style.configure("Treeview.Cell",
                            relief="solid",
                            borderwidth=1)

        except Exception as e:
            print(f"列セパレーター設定エラー: {e}")
            # エラーが発生した場合は代替方法を使用
            self._apply_alternative_grid_style()

    def _apply_alternative_grid_style(self):
        """代替のグリッドスタイルを適用。"""

        # Canvasを使用してグリッド線を描画する方法
        def draw_grid_lines(event=None):
            # この関数は描画後に呼び出される
            try:
                # TreeviewのCanvasウィジェットにアクセス
                canvas = None
                for child in self.tree.winfo_children():
                    if isinstance(child, tk.Canvas):
                        canvas = child
                        break

                if canvas:
                    # 垂直線を描画
                    canvas_width = canvas.winfo_width()
                    canvas_height = canvas.winfo_height()

                    # 列の境界位置を計算
                    col_positions = []
                    x_pos = 0
                    for col in self.tree['columns']:
                        col_width = self.tree.column(col, 'width')
                        x_pos += col_width
                        col_positions.append(x_pos)

                    # 垂直線を描画
                    for x in col_positions[:-1]:  # 最後の線は除く
                        if x < canvas_width:
                            canvas.create_line(x, 0, x, canvas_height,
                                               fill="#d0d0d0", width=1, tags="grid_line")

            except Exception as e:
                print(f"グリッド線描画エラー: {e}")

        # 描画イベントをバインド
        self.tree.bind('<Configure>', draw_grid_lines)
        self.tree.bind('<Map>', draw_grid_lines)

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
        self.current_month_button.config(text=f"📅 {month:02d}月")
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
                while len(row_values) < cols:
                    row_values.append("")
            else:
                row_values = [str(day)] + [""] * (cols - 1)

            # セルの値をフォーマット（パディング追加）- ここで直接実装
            formatted_values = []
            for i, val in enumerate(row_values):
                if i == 0:  # 日付列
                    formatted_values.append(f" {val} ")
                else:
                    formatted_values.append(f" {val} " if val else "  ")

            # ＋ボタン列は空のまま
            formatted_values.append("")

            # 奇数・偶数行でタグを分ける
            tag = "odd_row" if day % 2 == 1 else "normal_row"
            self.tree.insert("", "end", values=formatted_values, tags=(tag,))

        # 合計行
        total_row = [" 合計 "] + ["  "] * (cols - 1) + [""]
        self.tree.insert("", "end", values=total_row, tags=("TOTAL",))

        # まとめ行
        summary_key = f"{self.current_year}-{month}-0"
        if summary_key in self.parent_table_data:
            sm_data = self.parent_table_data[summary_key]
            inc_str = f" {sm_data[3]} " if len(sm_data) > 3 and sm_data[3] else "  "
        else:
            inc_str = "  "

        summary_row = [" まとめ ", "  ", " 収入 ", inc_str, " 支出 ", "  "] + ["  "] * (cols - 6) + [""]
        self.tree.insert("", "end", values=summary_row, tags=("SUMMARY",))

        self._recalc_total_and_summary()

    def _setup_enhanced_grid_style(self):
        """強化されたグリッド線スタイルを設定（オプション）"""
        style = ttk.Style()

        # テーマをカスタマイズしてよりグリッド線らしく見せる
        style.theme_use('clam')

        # Treeviewのセル間隔を調整
        style.configure("Treeview",
                        fieldbackground="white",
                        background="white",
                        rowheight=28,
                        font=('Arial', 9),
                        borderwidth=1,
                        insertwidth=1)

        # より明確な境界線効果のための色設定
        self.tree.tag_configure("even_row", background="#ffffff")
        self.tree.tag_configure("odd_row", background="#f8f9fa")
        self.tree.tag_configure("separator", background="#dee2e6")  # より濃いグレー

    def _recalc_total_and_summary(self):
        """合計行とまとめ行を再計算。"""
        items = self.tree.get_children()
        if len(items) < 2:
            return

        total_row_id = items[-2]
        summary_row_id = items[-1]
        all_columns = self.default_columns + self.custom_columns
        cols = len(all_columns)

        # 合計行の計算
        sums = [0] * (cols - 1)
        for row_id in items[:-2]:
            row_vals = self.tree.item(row_id, 'values')
            for i in range(1, cols):
                try:
                    # パディングを除去して数値に変換
                    val_str = str(row_vals[i]).strip() if i < len(row_vals) else ""
                    val = int(val_str) if val_str else 0
                except (ValueError, TypeError, IndexError):
                    val = 0
                sums[i - 1] += val

        # 合計行を更新（パディング付き）
        total_vals = list(self.tree.item(total_row_id, 'values'))
        for i in range(1, cols):
            if sums[i - 1] == 0:
                total_vals[i] = "  "  # 空の場合はパディングのみ
            else:
                total_vals[i] = f" {sums[i - 1]} "  # 値をパディング

        while len(total_vals) <= cols:
            total_vals.append("")
        self.tree.item(total_row_id, values=total_vals)

        # 総支出計算
        grand_total = 0
        for v in total_vals[1:cols]:
            if v and str(v).strip():
                try:
                    grand_total += int(str(v).strip())
                except:
                    pass

        # まとめ行更新（パディング付き）
        summary_vals = list(self.tree.item(summary_row_id, 'values'))
        try:
            income_str = str(summary_vals[3]).strip() if len(summary_vals) > 3 else ""
            income_val = int(income_str) if income_str else 0
        except:
            income_val = 0

        sum_val = income_val - grand_total
        if sum_val == 0:
            summary_vals[1] = "  "
        else:
            summary_vals[1] = f" {sum_val} "

        if grand_total == 0:
            summary_vals[5] = "  "
        else:
            summary_vals[5] = f" {grand_total} "

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
            # まとめ行の場合、収入は列インデックス3（0ベース）
            col_index = 3
            col_name = "収入"
        else:
            try:
                day = int(row_vals[0])
            except:
                day = 0
            col_name = self.tree.heading(col_id, "text")

        dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_index}"

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
                if row_vals and str(row_vals[0]).strip() == str(d):
                    # 列数が足りない場合は拡張
                    while len(row_vals) < cols + 1:  # +1 for ＋ button column
                        row_vals.append("")

                    # 空文字列または0の場合は空表示（パディング付き）
                    display_value = "  "
                    if new_value and str(new_value).strip() != "" and str(new_value) != "0":
                        display_value = f" {new_value} "

                    row_vals[col_index] = display_value
                    self.tree.item(row_id, values=row_vals)
                    found = True
                    break

            if not found and d == 0:
                sum_vals = list(self.tree.item(summary_row_id, 'values'))
                while len(sum_vals) < cols + 1:  # +1 for ＋ button column
                    sum_vals.append("")

                # まとめ行でも同様の処理（パディング付き）
                display_value = "  "
                if new_value and str(new_value).strip() != "" and str(new_value) != "0":
                    display_value = f" {new_value} "

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

    def _create_widgets(self):
        """ウィジェットを作成。"""
        # グリッドの重み設定（リサイズ対応）
        self.grid_rowconfigure(1, weight=1)  # 修正: ツリービューエリアの行を1に変更
        self.grid_columnconfigure(0, weight=1)

        # タイトル
        title_label = tk.Label(self, text=f"{self.col_name} の詳細入力",
                               font=('Arial', 16, 'bold'), bg='#f0f0f0')
        title_label.grid(row=0, column=0, pady=(10, 15), sticky="ew")

        # ツリービューコンテナ（修正: スクロールバー重複を解決）
        tree_container = tk.Frame(self, bg='#f0f0f0')
        tree_container.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        # ツリービューとスクロールバー
        columns = ["取引先", "金額", "詳細"]
        self.child_tree = ttk.Treeview(tree_container, columns=columns, show="headings")

        for col in columns:
            self.child_tree.heading(col, text=col)
            if col == "取引先":
                self.child_tree.column(col, anchor="center", width=150, minwidth=100)
            elif col == "金額":
                self.child_tree.column(col, anchor="center", width=100, minwidth=80)
            else:
                self.child_tree.column(col, anchor="center", width=200, minwidth=150)

        # ツリービューを配置
        self.child_tree.grid(row=0, column=0, sticky="nsew")

        # ツリービュー用スクロールバー（縦のみ）
        tree_scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.child_tree.yview)
        tree_scrollbar.grid(row=0, column=1, sticky="ns")
        self.child_tree.configure(yscrollcommand=tree_scrollbar.set)

        # イベントバインド
        self.child_tree.bind("<Double-1>", self._on_double_click)
        self.child_tree.bind("<Button-3>", self._on_right_click)  # 右クリック

        # ボタンフレーム（修正: 行番号を2に変更）
        button_frame = tk.Frame(self, bg='#f0f0f0')
        button_frame.grid(row=2, column=0, sticky="ew", pady=(15, 10), padx=10)

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
        hint_label = tk.Label(self,
                              text="使い方: セルをダブルクリックで編集、右クリックで行削除",
                              font=('Arial', 10), fg='#666666', bg='#f0f0f0')
        hint_label.grid(row=3, column=0, pady=(5, 10))

        # キーボードショートカット
        self.bind('<Return>', lambda e: self._on_ok_button())
        self.bind('<Escape>', lambda e: self.destroy())

        # マウスホイールでスクロール
        def on_mousewheel(event):
            self.child_tree.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.child_tree.bind("<MouseWheel>", on_mousewheel)

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
        """修正されたOKボタンの処理。"""
        # ツリービューから全ての行を収集
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
    setup_japanese_font()
    try:
        root = tk.Tk()
        app = YearApp(root)
        root.mainloop()
    except Exception as e:
        print(f"アプリケーション実行エラー: {e}")
        import traceback
        traceback.print_exc()