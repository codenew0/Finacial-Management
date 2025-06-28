import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime, date
import matplotlib.font_manager as fm
import warnings

# アプリケーション設定
DATA_FILE = "data.json"  # 家計データを保存するJSONファイル
SETTINGS_FILE = "settings.json"  # アプリケーション設定を保存するファイル

# matplotlibのフォント警告を抑制（日本語フォント設定時の警告を防ぐ）
warnings.filterwarnings('ignore', category=UserWarning, module='matplotlib.font_manager')


def setup_japanese_font():
    """
    matplotlib用の日本語フォントを設定する。

    Windowsシステムに存在する日本語フォントを優先順位順に検索し、
    最初に見つかったフォントをmatplotlibのデフォルトフォントとして設定する。
    日本語フォントが見つからない場合は、DejaVu Sansを使用する。
    """
    try:
        # システムにインストールされている全フォントのリストを取得
        available_fonts = [f.name for f in fm.fontManager.ttflist]

        # Windows環境で一般的な日本語フォントのリスト（優先順位順）
        japanese_fonts = ['Yu Gothic', 'Meiryo', 'MS Gothic', 'MS PGothic', 'MS UI Gothic']

        # 利用可能な日本語フォントを検索
        for font in japanese_fonts:
            if font in available_fonts:
                plt.rcParams['font.family'] = font
                return

        # 日本語フォントが見つからない場合のフォールバック
        plt.rcParams['font.family'] = 'DejaVu Sans'
    except Exception:
        # エラーが発生した場合もフォールバック
        plt.rcParams['font.family'] = 'DejaVu Sans'


class TreeviewTooltip:
    """
    Treeviewのセルにマウスオーバーした時に詳細情報を表示するツールチップ機能。

    各セルの金額の内訳（支払先、金額、詳細）を表示し、
    合計行では日別の内訳、まとめ行では収入・支出の詳細を表示する。
    """

    def __init__(self, treeview, parent_app):
        """
        ツールチップの初期化。

        Args:
            treeview: ツールチップを適用するTreeviewウィジェット
            parent_app: メインアプリケーションのインスタンス（データアクセス用）
        """
        self.treeview = treeview
        self.parent_app = parent_app
        self.tooltip_window = None  # ツールチップウィンドウの参照
        self.current_item = None  # 現在ホバーしている行
        self.current_column = None  # 現在ホバーしている列

        # マウスイベントをバインド
        self.treeview.bind('<Motion>', self._on_mouse_motion)  # マウス移動時
        self.treeview.bind('<Leave>', self._on_mouse_leave)  # マウスがTreeviewから離れた時

    def _on_mouse_motion(self, event):
        """
        マウスがTreeview上で移動した時の処理。

        マウスの位置から現在どのセルの上にいるかを判定し、
        適切なツールチップを表示する。
        """
        # マウス位置から行と列を特定
        item = self.treeview.identify_row(event.y)
        column = self.treeview.identify_column(event.x)

        # 無効な位置の場合はツールチップを非表示
        if not item or not column:
            self._hide_tooltip()
            return

        # 同じセルにいる場合は何もしない（ちらつき防止）
        if item == self.current_item and column == self.current_column:
            return

        # 新しいセルに移動した
        self.current_item = item
        self.current_column = column

        # 列インデックスを取得（#1, #2 などから数値を抽出）
        col_index = int(column[1:]) - 1
        all_columns = self.parent_app.default_columns + self.parent_app.custom_columns

        # 日付列や＋ボタン列の場合はツールチップを表示しない
        if col_index == 0 or col_index >= len(all_columns):
            self._hide_tooltip()
            return

        # 行の値を取得
        row_values = self.treeview.item(item, 'values')
        if not row_values:
            self._hide_tooltip()
            return

        # セルの値を確認（空または0の場合はツールチップを表示しない）
        cell_value = str(row_values[col_index]).strip() if col_index < len(row_values) else ""
        if not cell_value or cell_value == "0":
            self._hide_tooltip()
            return

        # 特殊行（合計行、まとめ行）の判定
        items = self.treeview.get_children()
        if len(items) < 2:
            self._hide_tooltip()
            return

        total_row_id = items[-2]  # 最後から2番目が合計行
        summary_row_id = items[-1]  # 最後がまとめ行

        # 行の種類に応じて適切なツールチップを表示
        if item == total_row_id:
            self._show_total_tooltip(event, col_index)
        elif item == summary_row_id:
            if col_index == 3:  # 収入列
                self._show_income_tooltip(event)
            elif col_index == 5:  # 支出列
                self._show_expense_tooltip(event)
            else:
                self._hide_tooltip()
        else:
            # 通常の日付行
            try:
                day = int(str(row_values[0]).strip())
                self._show_detail_tooltip(event, day, col_index)
            except ValueError:
                self._hide_tooltip()

    def _show_detail_tooltip(self, event, day, col_index):
        """
        通常セルの詳細情報をツールチップで表示。

        指定された日付と項目の全取引内容（支払先、金額、詳細）を
        リスト形式で表示し、複数取引がある場合は合計も表示する。
        """
        # child_dataから該当するデータを取得
        dict_key = f"{self.parent_app.current_year}-{self.parent_app.current_month}-{day}-{col_index}"
        data_list = self.parent_app.data.get(dict_key, [])

        if not data_list:
            self._hide_tooltip()
            return

        # ツールチップの内容を構築
        lines = []
        total = 0

        for row in data_list:
            if len(row) >= 3:
                partner = str(row[0]).strip() if row[0] else "（未入力）"
                amount_str = str(row[1]).strip() if row[1] else "0"
                detail = str(row[2]).strip() if row[2] else ""

                # 金額をパースして合計を計算
                try:
                    amount = int(amount_str.replace(',', '').replace('¥', ''))
                    total += amount
                    amount_display = f"¥{amount:,}"
                except ValueError:
                    amount_display = amount_str

                # 1行の情報を作成
                line = f"• {partner}: {amount_display}"
                if detail:
                    line += f" ({detail})"
                lines.append(line)

        # 複数取引がある場合は合計を表示
        if len(data_list) > 1:
            lines.append("─" * 30)
            lines.append(f"合計: ¥{total:,}")

        # ツールチップを表示
        self._show_tooltip(event, "\n".join(lines))

    def _show_total_tooltip(self, event, col_index):
        """
        合計行のツールチップを表示。

        指定された項目の月間の日別内訳を表示する。
        データが多い場合は最初の10日分のみ表示し、
        残りは省略表記する。
        """
        # 列名を取得
        all_columns = self.parent_app.default_columns + self.parent_app.custom_columns
        column_name = all_columns[col_index] if col_index < len(all_columns) else "不明"

        days_with_data = []
        total = 0
        days_in_month = self.parent_app.get_days_in_month(self.parent_app.current_month)

        # 月の全日について集計
        for day in range(1, days_in_month + 1):
            dict_key = f"{self.parent_app.current_year}-{self.parent_app.current_month}-{day}-{col_index}"
            if dict_key in self.parent_app.data:
                data_list = self.parent_app.data[dict_key]
                day_total = sum(self._parse_amount(row[1]) for row in data_list if len(row) > 1)

                if day_total > 0:
                    days_with_data.append(f"{day}日: ¥{day_total:,}")
                    total += day_total

        if days_with_data:
            lines = [f"【{column_name}の内訳】"]
            lines.extend(days_with_data[:10])  # 最大10日分まで表示
            if len(days_with_data) > 10:
                lines.append(f"... 他{len(days_with_data) - 10}日分")
            lines.append("─" * 30)
            lines.append(f"合計: ¥{total:,}")
            self._show_tooltip(event, "\n".join(lines))
        else:
            self._hide_tooltip()

    def _show_income_tooltip(self, event):
        """
        収入セルのツールチップを表示。

        まとめ行の収入欄にマウスオーバーした時に、
        その月の全収入の内訳（収入源、金額、詳細）を表示する。
        """
        dict_key = f"{self.parent_app.current_year}-{self.parent_app.current_month}-0-3"
        data_list = self.parent_app.data.get(dict_key, [])

        if not data_list:
            self._hide_tooltip()
            return

        lines = ["【収入の内訳】"]
        total = 0

        for row in data_list:
            if len(row) >= 3:
                source = str(row[0]).strip() if row[0] else "（未入力）"
                amount = self._parse_amount(row[1])
                detail = str(row[2]).strip() if row[2] else ""

                total += amount
                line = f"• {source}: ¥{amount:,}"
                if detail:
                    line += f" ({detail})"
                lines.append(line)

        if len(data_list) > 1:
            lines.append("─" * 30)
            lines.append(f"合計: ¥{total:,}")

        self._show_tooltip(event, "\n".join(lines))

    def _show_expense_tooltip(self, event):
        """
        支出合計のツールチップを表示。

        まとめ行の支出欄にマウスオーバーした時に、
        その月の全項目の支出内訳を項目別に集計して表示する。
        """
        lines = ["【支出の内訳】"]
        all_columns = self.parent_app.default_columns + self.parent_app.custom_columns
        grand_total = 0

        # 各項目（列）ごとに集計
        for col_index in range(1, len(all_columns)):  # 日付列を除く
            column_total = 0
            column_name = all_columns[col_index]
            days_in_month = self.parent_app.get_days_in_month(self.parent_app.current_month)

            # その項目の月間合計を計算
            for day in range(1, days_in_month + 1):
                dict_key = f"{self.parent_app.current_year}-{self.parent_app.current_month}-{day}-{col_index}"
                if dict_key in self.parent_app.data:
                    for row in self.parent_app.data[dict_key]:
                        if len(row) > 1:
                            column_total += self._parse_amount(row[1])

            # 0円でない項目のみ表示
            if column_total > 0:
                lines.append(f"• {column_name}: ¥{column_total:,}")
                grand_total += column_total

        lines.append("─" * 30)
        lines.append(f"合計: ¥{grand_total:,}")
        self._show_tooltip(event, "\n".join(lines))

    def _parse_amount(self, amount_str):
        """
        金額文字列を整数に変換する。

        Args:
            amount_str: 金額を表す文字列（カンマや円記号を含む可能性あり）

        Returns:
            int: パースされた金額（失敗時は0）
        """
        if not amount_str:
            return 0
        try:
            # カンマと円記号を除去してから数値に変換
            clean_amount = str(amount_str).replace(',', '').replace('¥', '').strip()
            return int(clean_amount) if clean_amount else 0
        except ValueError:
            return 0

    def _show_tooltip(self, event, text):
        """
        ツールチップウィンドウを表示する。

        Args:
            event: マウスイベント（位置情報を含む）
            text: 表示するテキスト
        """
        # 既存のツールチップがあれば削除
        self._hide_tooltip()

        # 新しいツールチップウィンドウを作成
        self.tooltip_window = tk.Toplevel(self.treeview)
        self.tooltip_window.wm_overrideredirect(True)  # ウィンドウ装飾を非表示

        # マウスカーソルの少し右下に配置
        self.tooltip_window.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")

        # ツールチップの内容を設定
        label = tk.Label(self.tooltip_window,
                         text=text,
                         justify=tk.LEFT,
                         background="#ffffcc",  # 薄い黄色の背景
                         relief=tk.SOLID,
                         borderwidth=1,
                         font=("Arial", 9))
        label.pack()

    def _hide_tooltip(self):
        """ツールチップを非表示にする。"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
        self.current_item = None
        self.current_column = None

    def _on_mouse_leave(self, event):
        """マウスがTreeviewから離れた時の処理。"""
        self._hide_tooltip()


class BaseDialog(tk.Toplevel):
    """
    すべてのダイアログの基底クラス。

    共通の初期化処理（中央配置、モーダル設定など）を提供し、
    各ダイアログクラスはこのクラスを継承することで、
    一貫したUIと動作を実現する。
    """

    def __init__(self, parent, title, width=800, height=600):
        """
        ダイアログの基本設定を行う。

        Args:
            parent: 親ウィンドウ
            title: ダイアログのタイトル
            width: ダイアログの幅（デフォルト: 800）
            height: ダイアログの高さ（デフォルト: 600）
        """
        super().__init__(parent)
        self.title(title)
        self.configure(bg='#f0f0f0')  # 統一された背景色

        # ダイアログを親ウィンドウの中央に配置
        self._center_on_parent(width, height)

        # モーダルダイアログとして設定
        self.transient(parent)  # 親ウィンドウに従属
        self.grab_set()  # このダイアログ以外の操作を無効化
        self.resizable(True, True)  # サイズ変更可能

        # ダイアログを前面に表示
        self.lift()
        self.focus_force()

    def _center_on_parent(self, width, height):
        """
        ダイアログを親ウィンドウの中央に配置する。

        画面外にはみ出さないように位置を調整し、
        ユーザーが見やすい位置に表示する。
        """
        # 親ウィンドウの位置とサイズを取得
        parent_x = self.master.winfo_x()
        parent_y = self.master.winfo_y()
        parent_w = self.master.winfo_width()
        parent_h = self.master.winfo_height()

        # 中央配置の計算
        x = parent_x + (parent_w - width) // 2
        y = parent_y + (parent_h - height) // 2

        # ダイアログの位置とサイズを設定
        self.geometry(f"{width}x{height}+{x}+{y}")

        # 最小サイズを設定（元のサイズの75%）
        self.minsize(int(width * 0.75), int(height * 0.67))


class MonthlyDataDialog(BaseDialog):
    """
    月間データの詳細を表示するダイアログ。

    指定された年月のすべての取引データを一覧表示し、
    項目ごとにソート可能な機能を提供する。
    合計金額や平均金額などの統計情報も表示する。
    """

    def __init__(self, parent, parent_app, year, month):
        """
        月間データダイアログを初期化する。

        Args:
            parent: 親ウィンドウ
            parent_app: メインアプリケーションのインスタンス
            year: 表示する年
            month: 表示する月
        """
        self.parent_app = parent_app
        self.year = year
        self.month = month
        self.monthly_data = []  # 月間データを格納するリスト
        self.sort_column = None  # 現在ソートされている列
        self.sort_reverse = False  # ソートの方向（False=昇順、True=降順）

        # 基底クラスの初期化（タイトルに年月を含める）
        super().__init__(parent, f"月間データ詳細 - {year}年{month:02d}月")

        # UI要素を作成してデータを読み込む
        self._create_widgets()
        self._load_monthly_data()

    def _create_widgets(self):
        """
        ダイアログ内のUI要素を作成する。

        タイトル、データ表示用のTreeview、統計情報表示エリア、
        閉じるボタンなどを配置する。
        """
        # グリッドレイアウトの設定（データ表示エリアを伸縮可能にする）
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # タイトルセクション
        title_frame = tk.Frame(self, bg='#f0f0f0')
        title_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        title_label = tk.Label(title_frame,
                               text=f"{self.year}年{self.month:02d}月の詳細データ",
                               font=('Arial', 16, 'bold'),
                               bg='#f0f0f0')
        title_label.pack()

        # データ表示セクション
        result_frame = tk.Frame(self, bg='#f0f0f0')
        result_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        result_frame.grid_rowconfigure(0, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)

        # Treeview（データ表示用のテーブル）
        columns = ["年月日", "項目", "支払先", "金額（円）", "メモ"]
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=15)

        # 重複データ用のタグを設定（薄い赤色の背景）
        self.result_tree.tag_configure("duplicate", background="#ffcccc")

        # 各列のヘッダー設定（クリックでソート可能）
        for col in columns:
            self.result_tree.heading(col, text=col, command=lambda c=col: self._sort_by_column(c))

        # 列幅の設定とデフォルト値の保存
        self.default_column_widths = {}
        widths = {
            "年月日": 100,
            "項目": 120,
            "支払先": 150,
            "金額（円）": 100,
            "メモ": 200
        }

        for col, width in widths.items():
            self.result_tree.column(col, anchor="center", width=width, minwidth=int(width*0.8))
            self.default_column_widths[col] = width

        self.result_tree.grid(row=0, column=0, sticky="nsew")

        # スクロールバー（縦・横）
        v_scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.result_tree.configure(yscrollcommand=v_scrollbar.set)

        h_scrollbar = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.result_tree.configure(xscrollcommand=h_scrollbar.set)

        # 統計情報表示エリア
        stats_frame = tk.Frame(self, bg='#f0f0f0')
        stats_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=5)

        self.stats_label = tk.Label(stats_frame, text="", font=('Arial', 12, 'bold'), bg='#f0f0f0')
        self.stats_label.pack()

        # データ件数表示
        self.result_label = tk.Label(self, text="データ: 0 件",
                                     font=('Arial', 10), bg='#f0f0f0', fg='#666666')
        self.result_label.grid(row=3, column=0, sticky="w", padx=10, pady=(5, 10))

        # 閉じるボタン
        close_btn = tk.Button(self, text="閉じる", font=('Arial', 12),
                              bg='#f44336', fg='white', relief='raised', bd=2,
                              activebackground='#d32f2f', command=self.destroy)
        close_btn.grid(row=4, column=0, pady=(0, 10), ipady=5)

        # キーボードショートカット
        self.bind('<Escape>', lambda e: self.destroy())  # Escapeキーで閉じる
        self.result_tree.bind("<MouseWheel>", self._on_mousewheel)  # マウスホイールでスクロール

        # ダブルクリックイベントを追加
        self.result_tree.bind("<Double-1>", self._on_double_click)

        # 右クリックイベントを追加
        self.result_tree.bind("<Button-3>", self._on_header_right_click)

    def _on_header_right_click(self, event):
        """ヘッダーの右クリックイベントを処理する。"""
        region = self.result_tree.identify_region(event.x, event.y)
        if region != "heading":
            return

        col_id = self.result_tree.identify_column(event.x)
        if not col_id:
            return

        # 右クリックメニューを作成
        context_menu = tk.Menu(self, tearoff=0)
        context_menu.add_command(label="全ての列幅をリセット",
                                 command=self._reset_all_column_widths)

        context_menu.post(event.x_root, event.y_root)

    def _reset_all_column_widths(self):
        """全ての列の幅をデフォルトにリセットする。"""
        columns = ["年月日", "項目", "支払先", "金額（円）", "メモ"]

        for i, col_name in enumerate(columns):
            col_id = f"#{i + 1}"
            if col_name in self.default_column_widths:
                self.result_tree.column(col_id, width=self.default_column_widths[col_name])

    def _on_double_click(self, event):
        """
        行をダブルクリックした時の処理。

        選択された行のデータから日付と項目を特定し、
        メインウィンドウを該当するセルに移動してハイライトする。
        """
        # クリックされた行を取得
        item = self.result_tree.selection()
        if not item:
            return

        # 行のインデックスを取得
        row_index = self.result_tree.index(item[0])
        if row_index >= len(self.monthly_data):
            return

        # 該当するデータを取得
        data = self.monthly_data[row_index]

        # 日付と列インデックスを抽出
        year, month, day, col_index = data['sort_key']

        # 親アプリケーションの年月を変更（必要な場合）
        if self.parent_app.current_year != year or self.parent_app.current_month != month:
            self.parent_app.current_year = year
            self.parent_app.current_month = month
            self.parent_app.year_label.config(text=str(year))
            self.parent_app._select_month(month)

        # 該当する行とセルに移動
        self._navigate_to_cell(day, col_index)

    def _on_mousewheel(self, event):
        """マウスホイールによるスクロール処理。"""
        self.result_tree.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _navigate_to_cell(self, day, col_index):
        """
        メインウィンドウの指定されたセルに移動してハイライトする。

        Args:
            day: 日（1-31、0はまとめ行）
            col_index: 列インデックス
        """
        if not self.parent_app.tree:
            return

        # Treeviewのアイテムを取得
        items = self.parent_app.tree.get_children()
        if not items:
            return

        # 該当する行を検索
        target_item = None

        if day == 0:  # まとめ行の場合
            target_item = items[-1]  # 最後の行
        else:  # 通常の日付行
            for item in items[:-2]:  # 合計行とまとめ行を除く
                values = self.parent_app.tree.item(item, 'values')
                if values and str(values[0]).strip() == str(day):
                    target_item = item
                    break

        if target_item:
            # 該当する行を選択
            self.parent_app.tree.selection_set(target_item)

            # 行を表示範囲内にスクロール
            self.parent_app.tree.see(target_item)

            # フォーカスを設定
            self.parent_app.tree.focus(target_item)

    def _sort_by_column(self, column):
        """
        指定された列でデータをソートする。

        同じ列を再度クリックすると、ソート順が反転する。
        金額は数値として、その他は文字列としてソートする。

        Args:
            column: ソートする列名
        """
        # 同じ列をクリックした場合はソート順を反転
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            # 新しい列の場合は昇順から開始
            self.sort_column = column
            self.sort_reverse = False

        # ソートキーを決定する関数のマッピング
        sort_key_map = {
            "年月日": lambda x: x['sort_key'],  # タプル形式の日付でソート
            "項目": lambda x: x['column'],
            "支払先": lambda x: x['partner'],
            "金額（円）": lambda x: x['amount_value'],  # 数値でソート
            "メモ": lambda x: x['detail']
        }

        # データをソート
        self.monthly_data.sort(key=sort_key_map.get(column, lambda x: ""),
                               reverse=self.sort_reverse)

        # Treeviewを再表示
        self._refresh_treeview()
        self._update_column_headers()

    def _refresh_treeview(self):
        """
        ソート後のデータでTreeviewを再表示する。

        既存の表示をクリアし、ソート済みのデータを
        新しい順序で挿入する。
        """
        # 既存の表示をクリア
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # ソート済みデータを表示
        for result in self.monthly_data:
            values = [result['date'], result['column'], result['partner'],
                      result['amount'], result['detail']]
            self.result_tree.insert("", "end", values=values)

        # 重複データをチェックしてハイライト（この部分を追加）
        self._highlight_duplicates()

        # データ件数を更新
        self.result_label.config(text=f"データ: {len(self.monthly_data)} 件")

    def _update_column_headers(self):
        """
        ソート状態を示すため、列ヘッダーに矢印を表示する。

        昇順の場合は▲、降順の場合は▼を列名の後ろに付ける。
        """
        columns = ["年月日", "項目", "支払先", "金額（円）", "メモ"]

        for col in columns:
            if col == self.sort_column:
                # ソート中の列に矢印を付ける
                arrow = " ▼" if self.sort_reverse else " ▲"
                self.result_tree.heading(col, text=f"{col}{arrow}")
            else:
                # その他の列は通常表示
                self.result_tree.heading(col, text=col)

    def _highlight_duplicates(self):
        """
        重複データを検出して薄い赤色でハイライトする。

        すべてのフィールド（年月日、項目、支払先、金額、詳細）が
        完全に一致する行を重複と判定し、該当する行に
        "duplicate"タグを適用する。
        """
        # Treeviewのすべてのアイテムを取得
        items = self.result_tree.get_children()

        # 各行のデータを収集（重複チェック用）
        row_data_map = {}  # {item_id: (date, column, partner, amount, detail)}

        for item in items:
            values = self.result_tree.item(item, 'values')
            if values:
                # データをタプルとして保存（ハッシュ可能にするため）
                data_tuple = tuple(str(v) for v in values[:5])  # 5つのフィールドすべて
                row_data_map[item] = data_tuple

        # 重複を検出
        seen_data = {}  # {data_tuple: [item_id1, item_id2, ...]}

        for item_id, data_tuple in row_data_map.items():
            if data_tuple in seen_data:
                seen_data[data_tuple].append(item_id)
            else:
                seen_data[data_tuple] = [item_id]

        # 重複している行にタグを適用
        for data_tuple, item_ids in seen_data.items():
            if len(item_ids) > 1:  # 2つ以上ある場合は重複
                for item_id in item_ids:
                    # 既存のタグを保持しつつ、duplicateタグを追加
                    current_tags = list(self.result_tree.item(item_id, 'tags'))

                    # タグをリストに変換（タプルや文字列の場合があるため）
                    if isinstance(current_tags, str):
                        tag_list = [current_tags] if current_tags else []
                    else:
                        tag_list = list(current_tags)

                    # duplicateタグを追加
                    if "duplicate" not in tag_list:
                        tag_list.append("duplicate")

                    # タグをタプルとして設定（Treeviewが期待する形式）
                    self.result_tree.item(item_id, tags=tuple(tag_list))

    def _load_monthly_data(self):
        """
        指定された年月のデータをchild_dataから読み込む。

        データを整形して表示用のリストに格納し、
        統計情報（合計金額、平均金額）も計算する。
        収入データ（まとめ行）は除外する。
        """
        # 既存の表示をクリア
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        self.monthly_data = []
        total_amount = 0  # 月間合計金額
        total_count = 0  # 取引件数

        # child_dataから該当月のデータを検索
        for dict_key, data_list in self.parent_app.data.items():
            try:
                # キーを解析（形式: "年-月-日-列インデックス"）
                parts = dict_key.split("-")
                if len(parts) == 4:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])
                    col_index = int(parts[3])

                    # まとめ行（day=0）は収入データなので除外
                    if day == 0:
                        continue

                    # 指定された年月のデータのみ処理
                    if year == self.year and month == self.month:
                        # 項目名を取得
                        all_columns = self.parent_app.default_columns + self.parent_app.custom_columns
                        column_name = all_columns[col_index] if col_index < len(all_columns) else f"列{col_index}"
                        date_str = f"{year}/{month:02d}/{day:02d}"

                        # 各取引データを処理
                        for row in data_list:
                            if len(row) >= 3:  # 必要な要素（支払先、金額、詳細）が揃っているか確認
                                # 各フィールドのデータを安全に取得
                                partner = str(row[0]).strip() if row[0] else ""
                                amount_str = str(row[1]).strip() if row[1] else ""
                                detail = str(row[2]).strip() if row[2] else ""

                                # 金額を数値に変換（統計計算用）
                                amount_value = self._parse_amount(amount_str)

                                # 結果データを構造化
                                result = {
                                    'date': date_str,
                                    'column': column_name,
                                    'partner': partner,
                                    'amount': amount_str,
                                    'detail': detail,
                                    'amount_value': amount_value,
                                    'sort_key': (year, month, day, col_index)  # ソート用のタプル
                                }
                                self.monthly_data.append(result)
                                total_amount += amount_value
                                total_count += 1
            except (ValueError, IndexError):
                # データ解析エラーは無視して続行
                continue

        # デフォルトで日付順にソート
        self.monthly_data.sort(key=lambda x: x['sort_key'])
        self.sort_column = "年月日"
        self.sort_reverse = False

        # データをTreeviewに表示
        for result in self.monthly_data:
            values = [result['date'], result['column'], result['partner'],
                      result['amount'], result['detail']]
            self.result_tree.insert("", "end", values=values)

        # 重複データをチェックしてハイライト
        self._highlight_duplicates()

        # 列ヘッダーを更新（ソート状態を表示）
        self._update_column_headers()

        # 統計情報を計算して表示
        if total_count > 0:
            avg_amount = total_amount / total_count
            self.stats_label.config(text=f"合計金額: ¥{total_amount:,} | 平均金額: ¥{avg_amount:.0f}")
        else:
            self.stats_label.config(text="")

        # データ件数を表示
        self.result_label.config(text=f"データ: {len(self.monthly_data)} 件")

    def _parse_amount(self, amount_str):
        """
        金額文字列を整数に変換する。

        カンマや円記号を除去し、数値に変換する。
        変換できない場合は0を返す。

        Args:
            amount_str: 金額を表す文字列

        Returns:
            int: パースされた金額
        """
        if not amount_str:
            return 0
        try:
            clean_amount = amount_str.replace(',', '').replace('¥', '').strip()
            return int(clean_amount) if clean_amount else 0
        except ValueError:
            return 0


class SearchDialog(BaseDialog):
    """
    取引データを検索するためのダイアログ。

    全期間のデータから、支払先名、金額、詳細のいずれかに
    検索文字列が含まれる取引を抽出して表示する。
    大文字小文字を区別しない部分一致検索を行う。
    """

    def __init__(self, parent, parent_app):
        """
        検索ダイアログを初期化する。

        Args:
            parent: 親ウィンドウ
            parent_app: メインアプリケーションのインスタンス
        """
        self.parent_app = parent_app
        self.search_results = []  # 検索結果を格納するリスト

        super().__init__(parent, "検索")

        self._create_widgets()

    def _create_widgets(self):
        # グリッドレイアウトの設定
        self.grid_rowconfigure(1, weight=1)  # 結果表示エリアを伸縮可能に
        self.grid_columnconfigure(0, weight=1)

        # 検索入力セクション
        search_frame = tk.Frame(self, bg='#f0f0f0')
        search_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        search_frame.grid_columnconfigure(1, weight=1)  # 入力フィールドを伸縮可能に

        # 検索ラベル
        tk.Label(search_frame, text="検索文字列:", font=('Arial', 12), bg='#f0f0f0').grid(
            row=0, column=0, padx=(0, 10), sticky="w")

        # 検索入力フィールド
        self.search_entry = tk.Entry(search_frame, font=('Arial', 12), width=30)
        self.search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        self.search_entry.focus_set()  # 初期フォーカスを設定

        # 検索ボタン
        search_btn = tk.Button(search_frame, text="検索", font=('Arial', 12),
                               bg='#2196f3', fg='white', relief='raised', bd=2,
                               activebackground='#1976d2', command=self._search)
        search_btn.grid(row=0, column=2, padx=(0, 10), ipady=3)

        # クリアボタン
        clear_btn = tk.Button(search_frame, text="クリア", font=('Arial', 12),
                              bg='#ff9800', fg='white', relief='raised', bd=2,
                              activebackground='#f57c00', command=self._clear_results)
        clear_btn.grid(row=0, column=3, ipady=3)

        # 結果表示セクション
        result_frame = tk.Frame(self, bg='#f0f0f0')
        result_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        result_frame.grid_rowconfigure(0, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)

        # 結果表示用Treeview
        columns = ["年月日", "項目", "支払先", "金額（円）", "メモ"]
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=15)

        # 列の設定
        self.default_column_widths = {}  # デフォルト列幅を保存
        for col in columns:
            self.result_tree.heading(col, text=col)

        widths = {
            "年月日": 100,
            "項目": 120,
            "支払先": 150,
            "金額（円）": 100,
            "メモ": 200
        }

        for col, width in widths.items():
            self.result_tree.column(col, anchor="center", width=width, minwidth=int(width * 0.8))
            self.default_column_widths[col] = width

        self.result_tree.grid(row=0, column=0, sticky="nsew")

        # スクロールバー
        v_scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.result_tree.configure(yscrollcommand=v_scrollbar.set)

        h_scrollbar = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.result_tree.configure(xscrollcommand=h_scrollbar.set)

        # 結果カウンター
        self.result_label = tk.Label(self, text="検索結果: 0 件",
                                     font=('Arial', 10), bg='#f0f0f0', fg='#666666')
        self.result_label.grid(row=2, column=0, sticky="w", padx=10, pady=(5, 10))

        # 閉じるボタン
        close_btn = tk.Button(self, text="閉じる", font=('Arial', 12),
                              bg='#f44336', fg='white', relief='raised', bd=2,
                              activebackground='#d32f2f', command=self.destroy)
        close_btn.grid(row=3, column=0, pady=(0, 10), ipady=5)

        # キーボードショートカット
        self.search_entry.bind('<Return>', lambda e: self._search())  # Enterキーで検索
        self.bind('<Escape>', lambda e: self.destroy())  # Escapeキーで閉じる
        self.bind('<Control-f>', lambda e: self.search_entry.focus_set())  # Ctrl+Fで検索フィールドにフォーカス
        self.result_tree.bind("<MouseWheel>", self._on_mousewheel)

        # ダブルクリックイベントを追加
        self.result_tree.bind("<Double-1>", self._on_double_click)

        # 右クリックイベントを追加
        self.result_tree.bind("<Button-3>", self._on_header_right_click)

    def _on_header_right_click(self, event):
        """ヘッダーの右クリックイベントを処理する。"""
        region = self.result_tree.identify_region(event.x, event.y)
        if region != "heading":
            return

        col_id = self.result_tree.identify_column(event.x)
        if not col_id:
            return

        # 右クリックメニューを作成
        context_menu = tk.Menu(self, tearoff=0)
        context_menu.add_command(label="全ての列幅をリセット",
                                 command=self._reset_all_column_widths)

        context_menu.post(event.x_root, event.y_root)

    def _reset_all_column_widths(self):
        """全ての列の幅をデフォルトにリセットする。"""
        columns = ["年月日", "項目", "支払先", "金額（円）", "メモ"]

        for i, col_name in enumerate(columns):
            col_id = f"#{i + 1}"
            if col_name in self.default_column_widths:
                self.result_tree.column(col_id, width=self.default_column_widths[col_name])

    def _on_mousewheel(self, event):
        """マウスホイールによるスクロール処理。"""
        self.result_tree.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_double_click(self, event):
        """
        検索結果の行をダブルクリックした時の処理。

        選択された行のデータから日付と項目を特定し、
        メインウィンドウを該当するセルに移動してハイライトする。
        """
        # クリックされた行を取得
        item = self.result_tree.selection()
        if not item:
            return

        # 行のインデックスを取得
        try:
            # Treeviewの全アイテムから選択されたアイテムのインデックスを探す
            all_items = self.result_tree.get_children()
            row_index = all_items.index(item[0])
        except (ValueError, IndexError):
            return

        if row_index >= len(self.search_results):
            return

        # 該当するデータを取得
        data = self.search_results[row_index]

        # 日付と列インデックスを抽出
        year, month, day, col_index = data['sort_key']

        # 親アプリケーションの年月を変更（必要な場合）
        if self.parent_app.current_year != year or self.parent_app.current_month != month:
            self.parent_app.current_year = year
            self.parent_app.current_month = month
            self.parent_app.year_label.config(text=str(year))
            self.parent_app._select_month(month)

        # 該当する行とセルに移動
        self._navigate_to_cell(day, col_index)

    def _navigate_to_cell(self, day, col_index):
        """
        メインウィンドウの指定されたセルに移動してハイライトする。

        Args:
            day: 日（1-31、0はまとめ行）
            col_index: 列インデックス
        """
        if not self.parent_app.tree:
            return

        # Treeviewのアイテムを取得
        items = self.parent_app.tree.get_children()
        if not items:
            return

        # 該当する行を検索
        target_item = None

        if day == 0:  # まとめ行の場合
            target_item = items[-1]  # 最後の行
        else:  # 通常の日付行
            for item in items[:-2]:  # 合計行とまとめ行を除く
                values = self.parent_app.tree.item(item, 'values')
                if values and str(values[0]).strip() == str(day):
                    target_item = item
                    break

        if target_item:
            # 該当する行を選択
            self.parent_app.tree.selection_set(target_item)

            # 行を表示範囲内にスクロール
            self.parent_app.tree.see(target_item)

            # フォーカスを設定
            self.parent_app.tree.focus(target_item)

    def _search(self):
        """
        検索を実行する。

        入力された検索文字列を使用して、全期間のデータから
        該当する取引を検索し、結果を表示する。
        検索は大文字小文字を区別せず、部分一致で行う。
        """
        # 検索文字列を取得
        search_text = self.search_entry.get().strip()
        if not search_text:
            messagebox.showwarning("警告", "検索文字列を入力してください。")
            return

        # 前回の検索結果をクリア
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        self.search_results = []
        search_text_lower = search_text.lower()  # 大文字小文字を区別しない検索のため

        # 全データから検索
        for dict_key, data_list in self.parent_app.data.items():
            try:
                # キーを解析
                parts = dict_key.split("-")
                if len(parts) == 4:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])
                    col_index = int(parts[3])

                    # 項目名を取得
                    all_columns = self.parent_app.default_columns + self.parent_app.custom_columns
                    column_name = all_columns[col_index] if col_index < len(all_columns) else f"列{col_index}"
                    date_str = f"{year}/{month:02d}/{day:02d}"

                    # 各取引データを検索
                    for row in data_list:
                        if len(row) >= 3:
                            partner = str(row[0]).strip() if row[0] else ""
                            amount = str(row[1]).strip() if row[1] else ""
                            detail = str(row[2]).strip() if row[2] else ""

                            # いずれかのフィールドに検索文字列が含まれているかチェック
                            if (search_text_lower in partner.lower() or
                                    search_text_lower in amount.lower() or
                                    search_text_lower in detail.lower()):
                                # 検索結果を構造化
                                result = {
                                    'date': date_str,
                                    'column': column_name,
                                    'partner': partner,
                                    'amount': amount,
                                    'detail': detail,
                                    'sort_key': (year, month, day, col_index)
                                }
                                self.search_results.append(result)
            except (ValueError, IndexError):
                continue

        # 結果を日付順にソート
        self.search_results.sort(key=lambda x: x['sort_key'])

        # 結果を表示
        for result in self.search_results:
            values = [result['date'], result['column'], result['partner'],
                      result['amount'], result['detail']]
            self.result_tree.insert("", "end", values=values)

        # 結果カウンターを更新
        self.result_label.config(text=f"検索結果: {len(self.search_results)} 件")

    def _clear_results(self):
        """
        検索結果と入力フィールドをクリアする。

        新しい検索を開始する前に、前回の検索状態を
        リセットする。
        """
        # 表示されている検索結果を削除
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 内部データをリセット
        self.search_results = []

        # UIをリセット
        self.result_label.config(text="検索結果: 0 件")
        self.search_entry.delete(0, tk.END)
        self.search_entry.focus_set()


class ChartDialog(BaseDialog):
    """
    項目別の月間推移グラフを表示するダイアログ。

    総支出、総収入、各項目の月別推移を折れ線グラフで表示し、
    各データポイントに金額を表示する。
    matplotlibを使用してグラフを描画する。
    """

    def __init__(self, parent, parent_app):
        """
        グラフダイアログを初期化する。

        Args:
            parent: 親ウィンドウ
            parent_app: メインアプリケーションのインスタンス
        """
        self.parent_app = parent_app
        self.current_column_index = -1  # -1: 総支出、-2: 総収入、1以上: 各項目
        self.figure = None  # matplotlibのFigureオブジェクト
        self.canvas = None  # TkinterとmatplotlibをつなぐCanvas

        super().__init__(parent, "項目別月間推移グラフ", 900, 600)

        self._create_widgets()
        self._update_chart()

    def _create_widgets(self):
        """
        グラフダイアログのUI要素を作成する。

        タイトル、タブボタン（総支出、総収入、各項目）、
        グラフ表示エリア、閉じるボタンを配置する。
        """
        # グリッドレイアウトの設定
        self.grid_rowconfigure(2, weight=1)  # グラフエリアを伸縮可能に
        self.grid_columnconfigure(0, weight=1)

        # タイトルセクション
        title_frame = tk.Frame(self, bg='#f0f0f0')
        title_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        title_label = tk.Label(title_frame, text="項目別月間推移グラフ",
                               font=('Arial', 16, 'bold'), bg='#f0f0f0')
        title_label.pack()

        # タブセクション
        tab_frame = tk.Frame(self, bg='#f0f0f0')
        tab_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))

        # 総支出・総収入ボタンコンテナ
        summary_button_frame = tk.Frame(tab_frame, bg='#f0f0f0')
        summary_button_frame.pack(side=tk.LEFT, padx=(0, 15))

        # 総支出ボタン
        self.total_expense_btn = tk.Button(summary_button_frame, text="総支出",
                                           font=('Arial', 11, 'bold'),
                                           bg='#f44336', fg='white',
                                           relief='raised', bd=2,
                                           activebackground='#d32f2f',
                                           command=lambda: self._select_tab(-1))
        self.total_expense_btn.pack(side=tk.LEFT, padx=(0, 5))

        # 総収入ボタン
        self.total_income_btn = tk.Button(summary_button_frame, text="総収入",
                                          font=('Arial', 11, 'bold'),
                                          bg='#4caf50', fg='white',
                                          relief='raised', bd=2,
                                          activebackground='#45a049',
                                          command=lambda: self._select_tab(-2))
        self.total_income_btn.pack(side=tk.LEFT)

        # 区切り線
        separator = tk.Frame(tab_frame, bg='#cccccc', width=2)
        separator.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        # 項目別タブボタン（スクロール可能）
        all_columns = self.parent_app.default_columns + self.parent_app.custom_columns
        self.tab_buttons = []

        # スクロール可能なタブコンテナ
        tab_canvas = tk.Canvas(tab_frame, height=35, bg='#f0f0f0', highlightthickness=0)
        tab_canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tab_inner_frame = tk.Frame(tab_canvas, bg='#f0f0f0')
        tab_canvas.create_window((0, 0), window=tab_inner_frame, anchor="nw")

        # 各項目のタブボタンを作成（日付列を除く）
        for i, col_name in enumerate(all_columns[1:], start=1):
            btn = tk.Button(tab_inner_frame, text=col_name,
                            font=('Arial', 10),
                            bg='#e0e0e0', fg='black',
                            relief='raised', bd=2,
                            command=lambda idx=i: self._select_tab(idx))
            btn.pack(side=tk.LEFT, padx=2, pady=2, fill=tk.Y)
            self.tab_buttons.append(btn)

        # 初期選択状態を反映
        self._update_button_colors()

        # タブフレームのスクロール領域を設定
        tab_inner_frame.update_idletasks()
        tab_canvas.configure(scrollregion=tab_canvas.bbox("all"))

        # グラフ表示セクション
        chart_frame = tk.Frame(self, bg='#f0f0f0')
        chart_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        chart_frame.grid_rowconfigure(0, weight=1)
        chart_frame.grid_columnconfigure(0, weight=1)

        # matplotlib設定
        plt.style.use('default')
        setup_japanese_font()

        # グラフの作成
        self.figure = plt.Figure(figsize=(10, 5), dpi=80)
        self.canvas = FigureCanvasTkAgg(self.figure, chart_frame)
        self.canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew")

        # 閉じるボタン
        close_btn = tk.Button(self, text="閉じる", font=('Arial', 12),
                              bg='#f44336', fg='white', relief='raised', bd=2,
                              activebackground='#d32f2f', command=self.destroy)
        close_btn.grid(row=3, column=0, pady=10, ipady=5)

        # キーボードショートカット
        self.bind('<Escape>', lambda e: self.destroy())

    def _select_tab(self, index):
        """
        タブを選択してグラフを更新する。

        Args:
            index: -1=総支出、-2=総収入、1以上=各項目のインデックス
        """
        self.current_column_index = index
        self._update_button_colors()
        self._update_chart()

    def _update_button_colors(self):
        """
        選択されているタブのボタンの色を更新する。

        選択中のタブは強調表示され、その他のタブは
        通常の色に戻る。
        """
        # すべてのボタンの色をリセット
        self.total_expense_btn.config(bg='#f44336', fg='white')
        self.total_income_btn.config(bg='#4caf50', fg='white')

        for btn in self.tab_buttons:
            btn.config(bg='#e0e0e0', fg='black')

        # 選択されているボタンを強調
        if self.current_column_index == -1:  # 総支出
            self.total_expense_btn.config(bg='#d32f2f', fg='white')
        elif self.current_column_index == -2:  # 総収入
            self.total_income_btn.config(bg='#45a049', fg='white')
        elif 0 < self.current_column_index <= len(self.tab_buttons):
            self.tab_buttons[self.current_column_index - 1].config(bg='#2196f3', fg='white')

    def _update_chart(self):
        """
        現在選択されているタブに応じてグラフを更新する。

        データを収集し、折れ線グラフを描画する。
        各データポイントには金額ラベルを表示する。
        """
        if not self.figure:
            return

        # 既存のグラフをクリア
        self.figure.clear()

        # 選択されたタブに応じてデータを取得
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
        else:  # 各項目
            monthly_data = self._collect_category_data()
            all_columns = self.parent_app.default_columns + self.parent_app.custom_columns
            column_name = all_columns[self.current_column_index] if self.current_column_index < len(
                all_columns) else "不明"
            title = f'{column_name} の月間推移'
            ylabel = "金額 (円)"
            color = '#2196f3'

        # データがない場合の処理
        if not monthly_data:
            ax = self.figure.add_subplot(111)
            ax.text(0.5, 0.5, 'データがありません',
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax.transAxes, fontsize=16)
            ax.set_title('データなし')
            self.canvas.draw()
            return

        # グラフの描画
        ax = self.figure.add_subplot(111)

        # データを日付順にソート
        sorted_data = sorted(monthly_data.items())
        dates = [item[0] for item in sorted_data]
        amounts = [item[1] for item in sorted_data]

        # 折れ線グラフを描画
        ax.plot(dates, amounts, marker='o', linewidth=2, markersize=8, color=color)

        # 各データポイントに金額ラベルを追加
        for date, amount in zip(dates, amounts):
            ax.annotate(f'¥{amount:,}',  # 3桁ごとにカンマ区切り
                        xy=(date, amount),
                        xytext=(0, 10),  # ポイントの10ピクセル上に配置
                        textcoords='offset points',
                        ha='center',  # 水平方向の中央揃え
                        va='bottom',  # 垂直方向の下揃え
                        fontsize=9,
                        bbox=dict(boxstyle='round,pad=0.3',  # 角丸の背景
                                  facecolor='white',
                                  edgecolor=color,
                                  alpha=0.8))

        # グラフの装飾
        ax.set_title(title, fontsize=14, fontweight='bold')
        ax.set_xlabel('月', fontsize=12)
        ax.set_ylabel(ylabel, fontsize=12)

        # Y軸のフォーマット（カンマ区切り）
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{int(x):,}'))

        # X軸のフォーマット（年/月形式）
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m'))
        ax.xaxis.set_major_locator(mdates.MonthLocator())

        # グリッドを表示
        ax.grid(True, alpha=0.3)

        # X軸のラベルを45度回転
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=45)

        # Y軸の範囲を調整（上部のラベルが切れないように）
        y_min, y_max = ax.get_ylim()
        y_range = y_max - y_min
        ax.set_ylim(y_min - y_range * 0.05, y_max + y_range * 0.15)

        # レイアウトを調整
        self.figure.tight_layout()

        # グラフを更新
        self.canvas.draw()

    def _collect_total_expense_data(self):
        """
        月間総支出データを収集する。

        全期間の各月の総支出額を計算し、
        月をキー、金額を値とする辞書を返す。

        Returns:
            dict: {date(年,月,1): 金額} の形式の辞書
        """
        monthly_totals = {}

        for dict_key, data_list in self.parent_app.data.items():
            try:
                parts = dict_key.split("-")
                if len(parts) == 4:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])

                    # まとめ行（収入）は除外
                    if day == 0:
                        continue

                    # 月のキーを作成
                    month_key = date(year, month, 1)
                    if month_key not in monthly_totals:
                        monthly_totals[month_key] = 0

                    # 各取引の金額を合計
                    for row in data_list:
                        if len(row) > 1:
                            amount = self._parse_amount(row[1])
                            monthly_totals[month_key] += amount
            except (ValueError, IndexError):
                continue

        return monthly_totals

    def _collect_total_income_data(self):
        """
        月間総収入データを収集する。

        各月のまとめ行（day=0）の収入列（col_index=3）から
        収入データを抽出し、月別に集計する。

        Returns:
            dict: {date(年,月,1): 金額} の形式の辞書
        """
        monthly_totals = {}

        for dict_key, data_list in self.parent_app.data.items():
            try:
                parts = dict_key.split("-")
                if len(parts) == 4:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])
                    col_index = int(parts[3])

                    # まとめ行の収入列のみ対象
                    if day == 0 and col_index == 3:
                        month_key = date(year, month, 1)

                        # 収入データを合計
                        total_income = sum(self._parse_amount(row[1]) for row in data_list if len(row) > 1)

                        if total_income > 0:
                            monthly_totals[month_key] = total_income
            except (ValueError, IndexError):
                continue

        return monthly_totals

    def _collect_category_data(self):
        """
        特定項目の月間データを収集する。

        現在選択されている項目（current_column_index）の
        月別合計を計算し、辞書形式で返す。

        Returns:
            dict: {date(年,月,1): 金額} の形式の辞書
        """
        monthly_totals = {}

        for dict_key, data_list in self.parent_app.data.items():
            try:
                parts = dict_key.split("-")
                if len(parts) == 4:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])
                    col_index = int(parts[3])

                    # まとめ行は除外
                    if day == 0:
                        continue

                    # 選択された列のデータのみ処理
                    if col_index == self.current_column_index:
                        month_key = date(year, month, 1)
                        if month_key not in monthly_totals:
                            monthly_totals[month_key] = 0

                        # 各取引の金額を合計
                        for row in data_list:
                            if len(row) > 1:
                                amount = self._parse_amount(row[1])
                                monthly_totals[month_key] += amount
            except (ValueError, IndexError):
                continue

        return monthly_totals

    def _parse_amount(self, amount_str):
        """
        金額文字列を整数に変換する。

        Args:
            amount_str: 金額を表す文字列

        Returns:
            int: パースされた金額（失敗時は0）
        """
        if not amount_str:
            return 0
        try:
            clean_amount = str(amount_str).replace(',', '').replace('¥', '').strip()
            return int(clean_amount) if clean_amount else 0
        except ValueError:
            return 0

class TransactionDialog(BaseDialog):
    """
    取引詳細を入力・編集するダイアログ。

    特定の日付・項目の取引データ（支払先、金額、詳細）を
    テーブル形式で入力・編集できる。
    支払先の入力時には過去の履歴から候補を表示する。
    """

    def __init__(self, parent, parent_app, dict_key, col_name):
        """
        取引詳細ダイアログを初期化する。

        Args:
            parent: 親ウィンドウ
            parent_app: メインアプリケーションのインスタンス
            dict_key: データのキー（年-月-日-列インデックス）
            col_name: 項目名（表示用）
        """
        self.parent_app = parent_app
        self.dict_key = dict_key
        self.col_name = col_name
        self.entry_editor = None  # セル編集用のエディター

        # キーを解析して年月日と列インデックスを取得
        parts = dict_key.split("-")
        if len(parts) == 4:
            self.year = int(parts[0])
            self.month = int(parts[1])
            self.day = int(parts[2])
            self.col_index = int(parts[3])
        else:
            self.year = self.month = self.day = self.col_index = 0

        super().__init__(parent, f"支出・収入詳細 - {col_name}", 700, 500)

        self._create_widgets()

    def _create_widgets(self):
        """
        取引詳細ダイアログのUI要素を作成する。

        タイトル、取引データ編集用のTreeview、
        行追加ボタン、OK/キャンセルボタンなどを配置する。
        """
        # グリッドレイアウトの設定
        self.grid_rowconfigure(1, weight=1)  # Treeviewエリアを伸縮可能に
        self.grid_columnconfigure(0, weight=1)

        # タイトル
        title_label = tk.Label(self, text=f"{self.col_name} の詳細入力",
                               font=('Arial', 16, 'bold'), bg='#f0f0f0')
        title_label.grid(row=0, column=0, pady=(10, 15), sticky="ew")

        # Treeviewコンテナ
        tree_container = tk.Frame(self, bg='#f0f0f0')
        tree_container.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        # 取引データ編集用Treeview
        columns = ["支払先", "金額（円）", "メモ"]
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings")

        # 列の設定
        self.tree.heading("支払先", text="支払先")
        self.tree.heading("金額（円）", text="金額（円）")
        self.tree.heading("メモ", text="メモ")

        self.tree.column("支払先", anchor="center", width=150, minwidth=100)
        self.tree.column("金額（円）", anchor="center", width=100, minwidth=80)
        self.tree.column("メモ", anchor="center", width=200, minwidth=150)

        self.tree.grid(row=0, column=0, sticky="nsew")

        # スクロールバー
        scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        # イベントバインド
        self.tree.bind("<Double-1>", self._on_double_click)  # ダブルクリックで編集
        self.tree.bind("<Button-3>", self._on_right_click)  # 右クリックでメニュー

        # ボタンコンテナ
        button_frame = tk.Frame(self, bg='#f0f0f0')
        button_frame.grid(row=2, column=0, sticky="ew", pady=(15, 10), padx=10)

        # 行追加ボタン
        add_btn = tk.Button(button_frame, text="行を追加", font=('Arial', 12),
                            bg='#4caf50', fg='white', relief='raised', bd=2,
                            activebackground='#45a049', command=self._add_row)
        add_btn.pack(side=tk.LEFT, padx=(0, 10), ipady=5)

        # 右側のボタングループ
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
                           activebackground='#1976d2', command=self._on_ok)
        ok_btn.pack(side=tk.RIGHT, ipady=5)

        # 右クリックメニュー
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="行を削除", command=self._delete_row)

        # 使用方法のヒント
        hint_label = tk.Label(self, text="使い方: セルをダブルクリックで編集、右クリックで行削除",
                              font=('Arial', 10), fg='#666666', bg='#f0f0f0')
        hint_label.grid(row=3, column=0, pady=(5, 10))

        # キーボードショートカット
        self.bind('<Return>', lambda e: self._on_ok())
        self.bind('<Escape>', lambda e: self.destroy())
        self.tree.bind("<MouseWheel>", self._on_mousewheel)

        # データを読み込む（少し遅延させて確実に表示）
        self.after(50, self._load_data)

    def _on_mousewheel(self, event):
        """マウスホイールによるスクロール処理。"""
        self.tree.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _load_data(self):
        """
        既存のデータを読み込んでTreeviewに表示する。

        child_dataから該当するデータを取得し、
        各行をTreeviewに挿入する。
        データがない場合は空行を1つ追加する。
        """
        # 既存の表示をクリア
        for item in self.tree.get_children():
            self.tree.delete(item)

        # child_dataからデータを取得
        data_list = self.parent_app.data.get(self.dict_key, [])

        if not data_list:
            # データがない場合は空行を1つ追加
            self.tree.insert("", "end", values=["", "", ""])
        else:
            # 既存データを表示
            for row in data_list:
                # 3要素に統一（不足分は空文字で補完）
                row_data = list(row) if row else ["", "", ""]
                while len(row_data) < 3:
                    row_data.append("")
                self.tree.insert("", "end", values=row_data)

            # 最後に空行を追加（新規入力用）
            self.tree.insert("", "end", values=["", "", ""])

    def _add_row(self):
        """
        新しい空行を追加する。

        Treeviewの最後に空行を追加し、
        その行を選択状態にして見えるようにスクロールする。
        """
        item_id = self.tree.insert("", "end", values=["", "", ""])
        self.tree.selection_set(item_id)
        self.tree.see(item_id)

    def _on_double_click(self, event):
        """
        セルのダブルクリックイベントを処理する。

        クリックされたセルの位置にエディターを表示し、
        インライン編集を可能にする。
        支払先列の場合はコンボボックス、その他はEntryを使用。
        """
        # 既存のエディターがあれば保存してから削除
        if self.entry_editor:
            self._cancel_edit()

        # クリック位置から行と列を特定
        item_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)

        if not item_id or not col_id:
            return

        # セルの境界ボックスを取得
        bbox = self.tree.bbox(item_id, col_id)
        if not bbox:
            return

        # 列インデックスと現在の値を取得
        col_idx = int(col_id[1:]) - 1
        values = list(self.tree.item(item_id, 'values'))

        # 値が不足している場合は空文字で補完
        while len(values) <= col_idx:
            values.append("")

        x, y, w, h = bbox

        if col_idx == 0:  # 支払先列の場合
            # コンボボックスを作成（過去の支払先履歴を候補として表示）
            self.entry_editor = ttk.Combobox(self.tree, font=("Arial", 11))
            partner_list = sorted(list(self.parent_app.transaction_partners))
            self.entry_editor['values'] = partner_list
            self.entry_editor.set(values[col_idx])
        else:
            # その他の列はEntryを作成
            self.entry_editor = tk.Entry(self.tree, font=("Arial", 11))
            self.entry_editor.insert(0, str(values[col_idx]))

        # エディターを配置
        self.entry_editor.place(x=x, y=y, width=w, height=h)
        self.entry_editor.focus_set()
        self.entry_editor.select_range(0, tk.END)

        # エディターのイベントバインド
        self.entry_editor.bind("<Return>", lambda e: self._save_edit(item_id, col_idx))
        self.entry_editor.bind("<Tab>", lambda e: self._save_edit(item_id, col_idx))
        self.entry_editor.bind("<FocusOut>", lambda e: self._save_edit(item_id, col_idx))
        self.entry_editor.bind("<Escape>", lambda e: self._cancel_edit())

        # Treeviewの他の場所をクリックしたら保存
        self.tree.bind("<Button-1>", lambda e: self._save_edit(item_id, col_idx) if self.entry_editor else None,
                       add=True)

    def _save_edit(self, item_id, col_idx):
        """
        エディターの内容を保存する。

        Args:
            item_id: 編集中の行のID
            col_idx: 編集中の列のインデックス
        """
        if not self.entry_editor:
            return

        try:
            # エディターから値を取得
            new_value = self.entry_editor.get()

            # エディターを削除
            self.entry_editor.destroy()
            self.entry_editor = None

            # 支払先の場合は履歴に追加
            if col_idx == 0 and new_value.strip():
                self.parent_app.transaction_partners.add(new_value.strip())

            # Treeviewの値を更新
            values = list(self.tree.item(item_id, 'values'))
            while len(values) <= col_idx:
                values.append("")

            values[col_idx] = new_value
            self.tree.item(item_id, values=values)

            # 最終行に入力があったら新しい空行を追加
            all_items = self.tree.get_children()
            if item_id == all_items[-1] and any(str(cell).strip() for cell in values):
                self._add_row()

            # 一時的なイベントバインドを解除
            self.tree.unbind("<Button-1>")
        except Exception:
            # エラーが発生してもエディターは削除
            if self.entry_editor:
                self.entry_editor.destroy()
                self.entry_editor = None

    def _cancel_edit(self):
        """
        編集をキャンセルしてエディターを削除する。
        """
        if self.entry_editor:
            try:
                self.entry_editor.destroy()
            except:
                pass
            self.entry_editor = None

        # 一時的なイベントバインドを解除
        try:
            self.tree.unbind("<Button-1>")
        except:
            pass

    def _on_right_click(self, event):
        """
        右クリックイベントを処理してコンテキストメニューを表示する。

        クリックされた行を選択状態にしてから、
        マウス位置にメニューを表示する。
        """
        item_id = self.tree.identify_row(event.y)
        if item_id:
            self.tree.selection_set(item_id)
            self.context_menu.post(event.x_root, event.y_root)

    def _delete_row(self):
        """
        選択された行を削除する。

        確認ダイアログを表示してから削除を実行する。
        """
        selected_item = self.tree.selection()
        if selected_item and messagebox.askyesno("確認", "選択された行を削除しますか？"):
            self.tree.delete(selected_item[0])

    def _on_ok(self):
        """
        OKボタンの処理。

        Treeviewのデータを収集し、空行を除去してから
        child_dataに保存する。データがない場合は
        該当するキーを削除する。
        """
        # すべての行データを収集
        all_rows = []
        for item_id in self.tree.get_children():
            values = self.tree.item(item_id, 'values')
            # 3要素に統一
            row = list(values)
            while len(row) < 3:
                row.append("")
            all_rows.append(tuple(row))

        # 空行を除去（すべての列が空またはスペースのみの行を除外）
        filtered_rows = [row for row in all_rows if any(str(cell).strip() for cell in row)]

        if not filtered_rows:
            # データが空の場合
            # 親セルを空白に更新
            dict_key_day = f"{self.year}-{self.month}-{self.day}"
            self.parent_app.update_parent_cell(dict_key_day, self.col_index, "")

            # child_dataからキーを削除
            if self.dict_key in self.parent_app.data:
                del self.parent_app.data[self.dict_key]
        else:
            # データがある場合
            # child_dataに保存
            self.parent_app.data[self.dict_key] = filtered_rows

            # 金額列（インデックス1）を合計
            total = sum(self._parse_amount(row[1]) for row in filtered_rows if len(row) > 1)

            # 親セルを更新（合計が0の場合は空文字列）
            dict_key_day = f"{self.year}-{self.month}-{self.day}"
            display_value = str(total) if total != 0 else ""
            self.parent_app.update_parent_cell(dict_key_day, self.col_index, display_value)

        self.destroy()

    def _parse_amount(self, amount_str):
        """
        金額文字列を整数に変換する。

        Args:
            amount_str: 金額を表す文字列

        Returns:
            int: パースされた金額（失敗時は0）
        """
        if not amount_str:
            return 0
        try:
            clean_amount = str(amount_str).replace(',', '').replace('¥', '').strip()
            return int(clean_amount) if clean_amount else 0
        except ValueError:
            return 0


class HouseholdApp:
    """
    家計管理アプリケーションのメインクラス。

    年間の家計データを月別に管理し、項目ごとの支出・収入を
    記録・集計・分析する機能を提供する。
    データはJSONファイルに保存され、カスタム項目の追加や
    支払先の履歴管理もサポートする。
    """

    def __init__(self, root):
        """
        アプリケーションを初期化する。

        Args:
            root: TkinterのルートウィンドウInternet Explorer
        """
        self.root = root
        self._setup_window()
        self._initialize_data()
        self._load_settings()
        self._load_data()
        self._create_ui()
        self._show_month(self.current_month)

        # ウィンドウクローズ時の処理を設定
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

        # グローバルキーボードショートカット
        self.root.bind('<Control-f>', lambda e: SearchDialog(self.root, self))

    def _setup_window(self):
        """
        メインウィンドウの基本設定を行う。

        ウィンドウのタイトル、サイズ、位置、配色などを設定し、
        ttkスタイルを定義する。
        """
        self.root.title("💰 家計管理 2025")

        # ウィンドウサイズと位置の設定
        window_width = 1400
        window_height = 1020

        # 画面の中央に配置
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(1200, 800)
        self.root.resizable(True, True)

        # カラーテーマの定義（ダークテーマベース）
        self.colors = {
            'bg_primary': '#1e1e2e',  # メイン背景色
            'bg_secondary': '#313244',  # セカンダリ背景色
            'bg_tertiary': '#45475a',  # 第三背景色
            'accent': '#89b4fa',  # アクセントカラー（青）
            'accent_green': '#a6e3a1',  # 緑のアクセント
            'accent_red': '#f38ba8',  # 赤のアクセント
            'text_primary': '#cdd6f4',  # プライマリテキスト
            'text_secondary': '#bac2de',  # セカンダリテキスト
            'border': '#585b70',  # ボーダー色
            'hover': '#74c0fc'  # ホバー時の色
        }

        self.root.configure(bg=self.colors['bg_primary'])
        self._setup_styles()

    def _setup_styles(self):
        """
        ttkウィジェットのカスタムスタイルを定義する。

        ボタン、ラベルなどの見た目を統一し、
        モダンなUIを実現する。
        """
        style = ttk.Style()
        style.theme_use('clam')

        # 通常のボタンスタイル
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
                        font=('Segoe UI', 12, 'bold'),
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
                        font=('Segoe UI', 12, 'bold'),
                        relief='flat')

        style.map('Nav.TButton',
                  background=[('active', self.colors['accent']),
                              ('pressed', self.colors['hover'])])

    def _initialize_data(self):
        """
        アプリケーションのデータ構造を初期化する。

        現在の日付、データ格納用の辞書、デフォルトの項目などを設定する。
        """
        now = datetime.now()
        self.current_year = now.year
        self.current_month = now.month
        self.tree = None  # メインのTreeviewウィジェット
        self.data = {}  # 詳細データを格納する辞書
        self.transaction_partners = set()  # 支払先の履歴

        # より適切な支出項目名に変更
        self.default_columns = [
            "日付",
            "交通費",
            "外食費",
            "食料品費",
            "日用品費",
            "通販ネット",
            "娯楽費",
            "教育費",
            "サブスク",
            "住居費",
            "光熱水費",
            "通信費",
            "保険料",
            "その他"
        ]
        self.custom_columns = []  # ユーザー定義の項目

    def _load_settings(self):
        """
        設定ファイルから設定を読み込む。

        カスタム項目と支払先履歴を復元する。
        ファイルが存在しない場合や読み込みエラーの場合は
        デフォルト値を使用する。
        """
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    settings = json.load(f)
                    self.custom_columns = settings.get("custom_columns", [])
                    self.transaction_partners = set(settings.get("transaction_partners", []))
            except Exception:
                # エラーが発生しても続行（デフォルト値を使用）
                pass

    def _save_settings(self):
        """
        設定をファイルに保存する。

        カスタム項目と支払先履歴をJSONファイルに保存する。
        """
        settings = {
            "custom_columns": self.custom_columns,
            "transaction_partners": list(self.transaction_partners)
        }
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception:
            # 保存エラーは無視（次回起動時にデフォルト値が使用される）
            pass

    def _load_data(self):
        """
        データファイルから家計データを読み込む。

        JSONファイルからchild_dataを復元する。
        旧フォーマットとの互換性も維持する。
        """
        if not os.path.exists(DATA_FILE):
            return

        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                all_data = json.load(f)

            if "version" in all_data:
                # 新フォーマット
                self.data = all_data.get("data", {})
            else:
                # 旧フォーマット（後方互換性）
                self.data = all_data.get("data", {})
                self._extract_partners_from_data()
        except Exception:
            # 読み込みエラーは無視（空のデータで開始）
            pass

    def _extract_partners_from_data(self):
        """
        既存のデータから支払先を抽出して履歴に追加する。

        旧フォーマットのデータから支払先情報を
        抽出するための互換性処理。
        """
        for data_list in self.data.values():
            for row in data_list:
                if len(row) > 0 and row[0].strip():
                    self.transaction_partners.add(row[0].strip())

    def _save_data(self):
        """
        家計データをファイルに保存する。

        child_dataをJSONファイルに保存する。
        バージョン情報を含めて新フォーマットで保存。
        """
        all_data = {
            "version": "2.0",
            "data": self.data
        }
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(all_data, f, ensure_ascii=False, indent=2)
        except Exception:
            # 保存エラーは無視（データ損失を防ぐため例外を再発生させない）
            pass

    def _on_closing(self):
        """
        ウィンドウが閉じられる時の処理。

        データと設定を保存してからアプリケーションを終了する。
        """
        self._save_data()
        self._save_settings()
        self.root.destroy()

    def _create_ui(self):
        """
        メインウィンドウのUI要素を作成する。

        ヘッダー（年選択、月選択、ボタン類）と
        メインのデータ表示エリア（Treeview）を配置する。
        """
        # メインコンテナ（全体を包含するフレーム）
        main_container = tk.Frame(self.root, bg=self.colors['bg_primary'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # ヘッダーセクション
        header = tk.Frame(main_container, bg=self.colors['bg_secondary'])
        header.pack(fill=tk.X, pady=(0, 8))

        header_inner = tk.Frame(header, bg=self.colors['bg_secondary'])
        header_inner.pack(fill=tk.X, padx=15, pady=8)

        # 年選択コントロール（左側）
        year_container = tk.Frame(header_inner, bg=self.colors['bg_secondary'])
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
                                   font=('Segoe UI', 16, 'bold'),
                                   bg=self.colors['bg_tertiary'],
                                   fg=self.colors['text_primary'],
                                   padx=12, pady=4)
        self.year_label.pack()

        # 翌年ボタン
        ttk.Button(year_nav, text="▶", width=3, style='Nav.TButton',
                   command=self._next_year).pack(side=tk.LEFT, padx=(4, 0))

        # 月選択ボタン（1月〜12月）
        month_container = tk.Frame(header_inner, bg=self.colors['bg_secondary'])
        month_container.pack(side=tk.LEFT, padx=(20, 0))

        self.month_buttons = []
        for m in range(1, 13):
            btn = ttk.Button(month_container, text=f"{m:02d}", width=4, style='Modern.TButton',
                             command=lambda mo=m: self._select_month(mo))
            btn.pack(side=tk.LEFT, padx=1)
            self.month_buttons.append(btn)

        # 検索ボタン
        search_container = tk.Frame(header_inner, bg=self.colors['bg_secondary'])
        search_container.pack(side=tk.LEFT, padx=(20, 0))

        ttk.Button(search_container, text="🔍 検索 (Ctrl+F)", width=15, style='Accent.TButton',
                   command=lambda: SearchDialog(self.root, self)).pack()

        # 図表ボタン
        chart_container = tk.Frame(header_inner, bg=self.colors['bg_secondary'])
        chart_container.pack(side=tk.LEFT, padx=(10, 0))

        ttk.Button(chart_container, text="📊 図表", width=10, style='Accent.TButton',
                   command=lambda: ChartDialog(self.root, self)).pack()

        # 現在月表示（右側、クリック可能）
        month_info = tk.Frame(header_inner, bg=self.colors['bg_secondary'])
        month_info.pack(side=tk.RIGHT)

        self.current_month_button = ttk.Button(month_info,
                                               text=f"📅 {self.current_month:02d}月",
                                               style='Month.TButton',
                                               command=self._open_monthly_data)
        self.current_month_button.pack()

        # メインテーブルセクション
        tree_section = tk.Frame(main_container, bg=self.colors['bg_secondary'])
        tree_section.pack(fill=tk.BOTH, expand=True)

        self._create_treeview(tree_section)
        self._update_month_buttons()

    def _create_treeview(self, parent):
        """
        メインのTreeviewウィジェットを作成する。

        月間の家計データを表形式で表示するための
        Treeviewを作成し、各種イベントをバインドする。

        Args:
            parent: Treeviewを配置する親ウィジェット
        """
        if self.tree:
            return

        # 既存のウィジェットをクリア
        for widget in parent.winfo_children():
            widget.destroy()

        # 列の定義（デフォルト + カスタム + 追加ボタン）
        all_columns = self.default_columns + self.custom_columns + ["＋"]

        # Treeviewを格納するフレーム
        tree_frame = tk.Frame(parent, bg='white', relief='solid', bd=1)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Treeviewの作成
        self.tree = ttk.Treeview(tree_frame, columns=all_columns, show="headings", height=25)
        self.tree.grid(row=0, column=0, sticky="nsew")

        # 各列の設定
        self.default_column_widths = {}  # デフォルト列幅を保存
        for i, col in enumerate(all_columns):
            self.tree.heading(col, text=col)

            if i == 0:  # 日付列
                width = 60
                self.tree.column(col, anchor="center", width=width, minwidth=50)
            elif col == "＋":  # 追加ボタン列
                width = 40
                self.tree.column(col, anchor="center", width=width, minwidth=40, stretch=False)
            else:  # データ列
                width = 80
                self.tree.column(col, anchor="center", width=width, minwidth=60)

            self.default_column_widths[col] = width  # デフォルト幅を保存

        # スクロールバー（縦）
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=v_scrollbar.set)

        # スクロールバー（横）
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=h_scrollbar.set)

        # Treeviewのスタイル設定
        style = ttk.Style()
        style.theme_use('clam')

        style.configure("Treeview",
                        fieldbackground="white",
                        background="white",
                        rowheight=27,
                        font=('Arial', 9),
                        borderwidth=1,
                        relief="solid")

        style.configure("Treeview.Heading",
                        background="#e8e8e8",
                        font=('Arial', 10, 'bold'),
                        relief="raised",
                        borderwidth=1)

        style.map("Treeview",
                  background=[('selected', '#0078d4')],
                  foreground=[('selected', 'white')])

        # 行のタグ設定（背景色による区別）
        self.tree.tag_configure("TOTAL", background="#fff3cd", font=('Arial', 10, 'bold'))  # 合計行
        self.tree.tag_configure("SUMMARY", background="#d4edda", font=('Arial', 10, 'bold'))  # まとめ行
        self.tree.tag_configure("normal_row", background="white")  # 通常行（偶数）
        self.tree.tag_configure("odd_row", background="#f8f9fa")  # 通常行（奇数）

        # イベントバインド
        self.tree.bind("<Double-1>", self._on_double_click)  # ダブルクリック
        self.tree.bind("<Button-1>", self._on_single_click)  # シングルクリック
        self.tree.bind("<Button-3>", self._on_right_click)  # 右クリック
        self.tree.bind("<MouseWheel>", self._on_mousewheel)  # マウスホイール
        self.tree.bind("<Shift-MouseWheel>", lambda e: self.tree.xview_scroll(int(-1 * (e.delta / 120)), "units"))

        # 右クリックメニュー（カスタム列用）
        self.column_context_menu = tk.Menu(self.root, tearoff=0)
        self.column_context_menu.add_command(label="列名を編集", command=self._edit_column_name)
        self.column_context_menu.add_separator()
        self.column_context_menu.add_command(label="列を削除", command=self._delete_column)

        # ツールチップを初期化
        self.tooltip = TreeviewTooltip(self.tree, self)

    def _on_mousewheel(self, event):
        """
        マウスホイールイベントを処理する。

        通常は縦スクロール、Ctrlキーを押しながらの場合は
        横スクロールを行う。
        """
        if event.state & 0x4:  # Ctrlキーが押されている
            self.tree.xview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            self.tree.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _prev_year(self):
        """前年に移動する。"""
        self.current_year -= 1
        self.year_label.config(text=str(self.current_year))
        self._show_month(self.current_month)

    def _next_year(self):
        """翌年に移動する。"""
        self.current_year += 1
        self.year_label.config(text=str(self.current_year))
        self._show_month(self.current_month)

    def _select_month(self, month):
        """
        指定された月を選択する。

        Args:
            month: 選択する月（1-12）
        """
        self.current_month = month
        self.current_month_button.config(text=f"📅 {month:02d}月")
        self._update_month_buttons()
        self._show_month(month)

    def _update_month_buttons(self):
        """
        月選択ボタンのハイライトを更新する。

        現在選択されている月のボタンを強調表示する。
        """
        style = ttk.Style()
        for i, btn in enumerate(self.month_buttons, start=1):
            if i == self.current_month:
                btn.configure(style='Selected.TButton')
            else:
                btn.configure(style='Modern.TButton')

    def _open_monthly_data(self):
        """月間データ詳細ダイアログを開く。"""
        MonthlyDataDialog(self.root, self, self.current_year, self.current_month)

    def _show_month(self, month):
        """
        指定された月のデータを表示する。

        child_dataから該当月のデータを集計し、
        Treeviewに表示する。合計行とまとめ行も追加する。

        Args:
            month: 表示する月（1-12）
        """
        if not self.tree:
            return

        # 既存の表示をクリア
        for item in self.tree.get_children():
            self.tree.delete(item)

        all_columns = self.default_columns + self.custom_columns
        days = self.get_days_in_month(month)

        # 各日のデータを表示
        for day in range(1, days + 1):
            row_values = self._calculate_day_totals(day)
            formatted_values = self._format_row_values(row_values)
            formatted_values.append("")  # ＋ボタン列

            # 奇数・偶数行で背景色を変える
            tag = "odd_row" if day % 2 == 1 else "normal_row"
            self.tree.insert("", "end", values=formatted_values, tags=(tag,))

        # 合計行
        total_row = [" 合計 "] + ["  "] * (len(all_columns) - 1) + [""]
        self.tree.insert("", "end", values=total_row, tags=("TOTAL",))

        # まとめ行（収入・支出の表示）
        income_val = self._get_income_total()
        inc_str = f" {income_val} " if income_val != 0 else "  "
        summary_row = [" まとめ ", "  ", " 収入 ", inc_str, " 支出 ", "  "] + ["  "] * (len(all_columns) - 6) + [""]
        self.tree.insert("", "end", values=summary_row, tags=("SUMMARY",))

        # 合計とまとめ行の値を更新
        self._update_totals()

    def _calculate_day_totals(self, day):
        """
        特定の日の各項目の合計金額を計算する。

        child_dataから該当日のデータを取得し、
        各項目の合計を計算する。

        Args:
            day: 計算する日（1-31）

        Returns:
            list: 各列の合計値のリスト
        """
        all_columns = self.default_columns + self.custom_columns
        totals = [""] * len(all_columns)
        totals[0] = str(day)  # 日付列

        # 各項目の合計を計算
        for col_index in range(1, len(all_columns)):
            dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_index}"
            if dict_key in self.data:
                data_list = self.data[dict_key]
                # 金額列（インデックス1）を合計
                total = sum(self._parse_amount(row[1]) for row in data_list if len(row) > 1)
                if total != 0:
                    totals[col_index] = str(total)

        return totals

    def _format_row_values(self, values):
        """
        行データを表示用にフォーマットする。

        各セルの値にパディングを追加して、
        見やすい表示にする。

        Args:
            values: フォーマット前の値のリスト

        Returns:
            list: フォーマット済みの値のリスト
        """
        formatted = []
        for i, val in enumerate(values):
            if i == 0:  # 日付列
                formatted.append(f" {val} ")
            else:
                formatted.append(f" {val} " if val else "  ")
        return formatted

    def _get_income_total(self):
        """
        現在月の収入合計を取得する。

        まとめ行（day=0）の収入列（col_index=3）から
        収入の合計を計算する。

        Returns:
            int: 収入の合計金額
        """
        dict_key = f"{self.current_year}-{self.current_month}-0-3"
        if dict_key in self.data:
            data_list = self.data[dict_key]
            return sum(self._parse_amount(row[1]) for row in data_list if len(row) > 1)
        return 0

    def _update_totals(self):
        """
        合計行とまとめ行の値を更新する。

        表示されている各日のデータから月間合計を計算し、
        合計行とまとめ行（収支差額、総支出）を更新する。
        """
        items = self.tree.get_children()
        if len(items) < 2:
            return

        total_row_id = items[-2]  # 合計行
        summary_row_id = items[-1]  # まとめ行
        all_columns = self.default_columns + self.custom_columns
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
                          if v and str(v).strip() and str(v).strip().isdigit())

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
        """
        シングルクリックイベントを処理する。

        ヘッダーの＋ボタンがクリックされた場合、
        新しい列を追加する。
        """
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            col_id = self.tree.identify_column(event.x)
            if col_id:
                col_index = int(col_id[1:]) - 1
                all_columns = self.default_columns + self.custom_columns

                if col_index == len(all_columns):  # ＋ボタン列
                    self._add_column()

    def _on_double_click(self, event):
        """
        ダブルクリックイベントを処理する。

        セルがダブルクリックされた場合、該当する
        取引詳細ダイアログを開く。
        ヘッダーの場合は列の編集や追加を行う。
        """
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)

        if not row_id or not col_id:
            return

        # ヘッダーのクリック処理
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            col_index = int(col_id[1:]) - 1
            all_columns = self.default_columns + self.custom_columns

            if col_index == len(all_columns):  # ＋ボタン
                self._add_column()
            elif col_index >= len(self.default_columns):  # カスタム列
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

        # ＋ボタン列は編集不可
        col_index = int(col_id[1:]) - 1
        all_columns = self.default_columns + self.custom_columns
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
        """右クリックイベントを処理する。"""
        region = self.tree.identify_region(event.x, event.y)
        if region != "heading":
            return

        col_id = self.tree.identify_column(event.x)
        if not col_id:
            return

        col_index = int(col_id[1:]) - 1
        all_columns = self.default_columns + self.custom_columns

        # 右クリックメニューを再作成
        self.column_context_menu = tk.Menu(self.root, tearoff=0)

        # カスタム列の場合は編集・削除オプションを追加
        if len(all_columns) > col_index >= len(self.default_columns) and col_index != 0:
            self.selected_column_index = col_index
            self.column_context_menu.add_command(label="列名を編集", command=self._edit_column_name)
            self.column_context_menu.add_separator()
            self.column_context_menu.add_command(label="列を削除", command=self._delete_column)
            self.column_context_menu.add_separator()

        # すべての列で列幅リセットを利用可能
        self.column_context_menu.add_command(label="全ての列幅をリセット",
                                             command=self._reset_all_column_widths)

        self.column_context_menu.post(event.x_root, event.y_root)

    def _reset_all_column_widths(self):
        """指定された列の幅をデフォルトにリセットする。"""
        all_columns = self.default_columns + self.custom_columns + ["＋"]

        for i, col_name in enumerate(all_columns):
            col_id = f"#{i + 1}"
            if col_name in self.default_column_widths:
                self.tree.column(col_id, width=self.default_column_widths[col_name])

    def _add_column(self):
        """
        新しい列を追加する。

        ダイアログを表示して列名を入力し、
        カスタム列として追加する。
        """
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

        tk.Label(dialog, text="新しい列名を入力してください：", font=('Arial', 11)).pack(pady=10)

        entry = tk.Entry(dialog, font=('Arial', 11), width=25)
        entry.pack(pady=5)
        entry.focus_set()

        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)

        def on_ok():
            column_name = entry.get().strip()
            if column_name:
                all_columns = self.default_columns + self.custom_columns
                if column_name not in all_columns:
                    self.custom_columns.append(column_name)
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
        """
        Treeviewを再作成する。

        列の追加・削除・編集後に、Treeviewを
        新しい列構成で再作成する。
        """
        if self.tree:
            tree_parent = self.tree.master
            self.tree.destroy()
            self.tree = None
            self._create_treeview(tree_parent)

    def _edit_column_name(self, col_index=None):
        """
        カスタム列の名前を編集する。

        Args:
            col_index: 編集する列のインデックス
        """
        if col_index is None:
            col_index = getattr(self, 'selected_column_index', None)

        if col_index is None or col_index < len(self.default_columns):
            return

        custom_index = col_index - len(self.default_columns)
        if custom_index >= len(self.custom_columns):
            return

        old_name = self.custom_columns[custom_index]

        # 編集ダイアログを表示
        dialog = tk.Toplevel(self.root)
        dialog.title("列名の編集")
        dialog.resizable(False, False)

        # ダイアログを中央に配置
        dialog_width = 300
        dialog_height = 120

        x = self.root.winfo_x() + (self.root.winfo_width() - dialog_width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog_height) // 2

        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        dialog.transient(self.root)
        dialog.grab_set()

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
                all_columns = self.default_columns + self.custom_columns
                if new_name not in all_columns:
                    self.custom_columns[custom_index] = new_name
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
        """
        カスタム列を削除する。

        選択された列とその列に関連するすべてのデータを削除する。
        削除前に確認ダイアログを表示する。
        """
        col_index = getattr(self, 'selected_column_index', None)
        if col_index is None or col_index < len(self.default_columns):
            return

        custom_index = col_index - len(self.default_columns)
        if custom_index >= len(self.custom_columns):
            return

        col_name = self.custom_columns[custom_index]

        # 削除確認ダイアログ
        if messagebox.askyesno("確認", f"列 '{col_name}' を削除しますか？\n※この列のデータもすべて削除されます。"):
            # 列をリストから削除
            self.custom_columns.pop(custom_index)

            # 関連するデータを削除
            # child_dataから該当列のデータを持つキーを特定
            keys_to_delete = [key for key in self.data.keys()
                              if key.split("-")[3] == str(col_index)]

            # データを削除
            for key in keys_to_delete:
                del self.data[key]

            # Treeviewを再作成して変更を反映
            self._recreate_treeview()
            self._show_month(self.current_month)

    def update_parent_cell(self, dict_key_day, col_index, new_value):
        """
        親画面のセル表示を更新する。

        取引詳細ダイアログから呼ばれ、合計金額を
        メイン画面のTreeviewに反映する。

        Args:
            dict_key_day: 日付キー（年-月-日）
            col_index: 列インデックス
            new_value: 新しい値（合計金額または空文字）
        """
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
                    all_columns = self.default_columns + self.custom_columns
                    while len(row_vals) < len(all_columns) + 1:
                        row_vals.append("")

                    # 表示値をフォーマット（パディング付き）
                    display_value = "  "
                    if new_value and str(new_value).strip() != "" and str(new_value) != "0":
                        display_value = f" {new_value} "

                    # 値を更新
                    row_vals[col_index] = display_value
                    self.tree.item(row_id, values=row_vals)
                    break

            # まとめ行（収入）の更新
            if d == 0:
                sum_vals = list(self.tree.item(summary_row_id, 'values'))
                all_columns = self.default_columns + self.custom_columns
                while len(sum_vals) < len(all_columns) + 1:
                    sum_vals.append("")

                display_value = "  "
                if new_value and str(new_value).strip() != "" and str(new_value) != "0":
                    display_value = f" {new_value} "

                sum_vals[col_index] = display_value
                self.tree.item(summary_row_id, values=sum_vals)

            # 合計とまとめ行を再計算
            self._update_totals()

    def get_days_in_month(self, month):
        """
        指定された月の日数を取得する。

        うるう年を考慮して正確な日数を返す。

        Args:
            month: 月（1-12）

        Returns:
            int: その月の日数
        """
        if month in [1, 3, 5, 7, 8, 10, 12]:
            return 31
        elif month in [4, 6, 9, 11]:
            return 30
        elif month == 2:
            y = self.current_year
            # うるう年の判定
            return 29 if (y % 4 == 0 and y % 100 != 0) or (y % 400 == 0) else 28
        return 30

    def _parse_amount(self, amount_str):
        """
        金額文字列を整数に変換する。

        カンマや円記号を除去し、数値に変換する。
        変換できない場合は0を返す。

        Args:
            amount_str: 金額を表す文字列

        Returns:
            int: パースされた金額
        """
        if not amount_str:
            return 0
        try:
            clean_amount = str(amount_str).replace(',', '').replace('¥', '').strip()
            return int(clean_amount) if clean_amount else 0
        except ValueError:
            return 0


# ============================================================================
# メイン実行部
# ============================================================================
if __name__ == "__main__":
    # 日本語フォントを設定
    setup_japanese_font()

    try:
        # Tkinterのルートウィンドウを作成
        root = tk.Tk()

        # アプリケーションを初期化
        app = HouseholdApp(root)

        # メインループを開始
        root.mainloop()

    except Exception as e:
        # エラーが発生した場合は詳細を表示
        print(f"アプリケーション実行エラー: {e}")
        import traceback
        traceback.print_exc()