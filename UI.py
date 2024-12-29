import tkinter as tk  # Tkinter(標準ライブラリ)をインポート
from tkinter import ttk  # ツリービューやスタイル等に便利
from datetime import datetime  # 現在日時の取得や年月日操作
import json  # JSONファイルの読み書き
import os  # ファイルの存在チェック等に使用

DATA_FILE = "data.json"  # 保存・読み込みに使用するJSONファイル名


class YearApp:
    def __init__(self, root):
        """メインウィンドウの初期化処理。"""
        self.root = root  # ルートウィンドウを保持
        self._setup_root()  # ウィンドウの設定(タイトル・サイズ等)
        self._init_variables()  # 必要な変数を初期化
        self._load_data_from_file()  # JSONファイルからデータ読み込み

        self._create_frames_and_widgets()  # GUIパーツ(フレーム,ボタン,Treeview等)を作成
        self._show_month_sheet(self.current_month)  # 初回: 現在の月のシートを表示

        # ウィンドウを閉じる(×ボタン)操作時、データを保存して終了する
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    # -------------------------------------------------------------
    # ウィンドウ設定
    # -------------------------------------------------------------
    def _setup_root(self):
        """メインウィンドウ(root)の基本設定(タイトル,サイズ等)。"""
        self.root.title("家計管理2025")
        self.root.geometry("1300x1000")  # 幅1300, 高さ1000ピクセル
        self.root.resizable(False, False)  # ユーザによるリサイズを不可に

    # -------------------------------------------------------------
    # 変数初期化
    # -------------------------------------------------------------
    def _init_variables(self):
        """アプリ内で使う各種変数を初期化する。"""
        now = datetime.now()  # 現在日時を取得
        self.current_year = now.year  # 現在の年
        self.current_month = now.month  # 現在の月

        self.tree = None  # 親ダイアログのTreeview (まだ作っていないので None)
        self.child_data = {}  # 子ダイアログ明細用 dict
        self.parent_table_data = {}  # 親ダイアログ側の各セルデータ(dict)

    # -------------------------------------------------------------
    # JSONファイル 読み書き
    # -------------------------------------------------------------
    def _load_data_from_file(self):
        """data.json があれば読み込んで child_data / parent_table_data を復元。"""
        if not os.path.exists(DATA_FILE):  # ファイルがなければ何もしない
            return
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            all_data = json.load(f)  # JSONをパース
        # "child_data" と "parent_table_data" を取得
        self.child_data = all_data.get("child_data", {})
        self.parent_table_data = all_data.get("parent_table_data", {})

    def _save_data_to_file(self):
        """child_data と parent_table_data を JSONファイルに書き出す。"""
        all_data = {
            "child_data": self.child_data,
            "parent_table_data": self.parent_table_data
        }
        # indent=2 で見やすく整形
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(all_data, f, ensure_ascii=False, indent=2)

    def _on_closing(self):
        """ウィンドウ終了時のハンドラ。ファイル保存後に終了。"""
        self._save_data_to_file()  # 終了前に保存
        self.root.destroy()  # メインループを止める

    # -------------------------------------------------------------
    # ウィジェット作成 (フレーム,ボタン,Treeview等)
    # -------------------------------------------------------------
    def _create_frames_and_widgets(self):
        """メインウィンドウ内にフレーム、ボタン、ラベル、Treeviewエリア等を配置する。"""
        # (1) 上部: 年ボタン・年表示ラベル・現在月ラベル
        top_frame = tk.Frame(self.root)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        # 年を減らすボタン(◀)
        self.minus_button = tk.Button(
            top_frame, text="◀", font=("Arial", 20),
            command=self._decrease_year
        )
        self.minus_button.pack(side=tk.LEFT, padx=(0, 10))

        # 現在の年を表示
        self.year_label = tk.Label(
            top_frame, text=str(self.current_year), font=("Arial", 24)
        )
        self.year_label.pack(side=tk.LEFT, padx=10)

        # 年を増やすボタン(▶)
        self.plus_button = tk.Button(
            top_frame, text="▶", font=("Arial", 20),
            command=self._increase_year
        )
        self.plus_button.pack(side=tk.LEFT, padx=(10, 0))

        # 選択中の月(2桁)を表示するラベル
        self.current_month_label = tk.Label(
            top_frame, text=f"選択中の月：{self.current_month:02d}月", font=("Arial", 16)
        )
        self.current_month_label.pack(side=tk.LEFT, padx=30)

        # (2) 月ボタン(1~12)
        month_frame = tk.Frame(self.root)
        month_frame.pack(side=tk.TOP, fill=tk.X, padx=10)

        self.month_buttons = []
        for m in range(1, 13):
            btn = tk.Button(
                month_frame, text=f"{m:02d}", font=("Arial", 14),
                command=lambda mo=m: self._select_month(mo)
            )
            btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2, pady=5)
            self.month_buttons.append(btn)

        # (3) Treeview(メイン表)を配置するフレーム
        tree_frame = tk.Frame(self.root)
        tree_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Treeview生成(まだ無ければ)
        self._create_tree_if_needed(tree_frame)

        # 月ボタンのハイライト (最初の月ボタン表示)
        self._highlight_selected_month()

    def _create_tree_if_needed(self, parent):
        """親ダイアログのTreeviewを初回のみ作成し、スタイル等を設定。"""
        if self.tree is not None:
            return  # 既に作られている場合はスキップ

        # 14列(「日付」+13列)を定義
        columns = [
            "日付", "交通", "外食", "食品", "日常用品", "通販", "ゲーム課金",
            "ゲーム購入", "サービス", "家賃", "公共料金", "携帯・回線", "保険", "他"
        ]
        self.tree = ttk.Treeview(parent, columns=columns, show="headings")

        # 各列のヘッダー設定
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=len(col) * 10)

        # packでTreeview配置
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # スクロールバー(縦方向)
        sb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.tree.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        # Treeview側へ紐付け
        self.tree.configure(yscrollcommand=sb.set)

        # スタイル設定
        style = ttk.Style()
        style.configure("Treeview",
                        fieldbackground="white",
                        background="white",
                        rowheight=25,
                        borderwidth=1)
        style.configure("Treeview.Heading",
                        background="lightgray",
                        font=("Arial", 12, "bold"))
        style.map("Treeview", background=[('selected', 'lightblue')])

        # 合計行用タグ(TOTAL)
        self.tree.tag_configure("TOTAL", background="lightyellow", font=("Arial", 12, "bold"))
        # まとめ行用タグ(SUMMARY)
        self.tree.tag_configure("SUMMARY", background="lightgreen", font=("Arial", 12, "bold"))

        # ダブルクリックでセル編集へ
        self.tree.bind("<Double-1>", self._on_parent_double_click)

    # -------------------------------------------------------------
    # 年/月ボタン押下
    # -------------------------------------------------------------
    def _increase_year(self):
        """年+1して年ラベル更新 & 現在の月を再描画"""
        self.current_year += 1
        self.year_label.config(text=str(self.current_year))
        self._show_month_sheet(self.current_month)

    def _decrease_year(self):
        """年-1して年ラベル更新 & 現在の月を再描画"""
        self.current_year -= 1
        self.year_label.config(text=str(self.current_year))
        self._show_month_sheet(self.current_month)

    def _select_month(self, month):
        """月ボタンを押されたら、月を切り替えて再描画"""
        self.current_month = month
        self.current_month_label.config(text=f"選択中の月：{month:02d}月")
        self._highlight_selected_month()
        self._show_month_sheet(month)

    def _highlight_selected_month(self):
        """月ボタンの背景色を一旦リセットし、現在の月のみ青ハイライト"""
        for i, btn in enumerate(self.month_buttons, start=1):
            btn.config(bg="SystemButtonFace")
        # 選択中の月をハイライト
        self.month_buttons[self.current_month - 1].config(bg="lightblue")

    # -------------------------------------------------------------
    # 月のシートを表示 (日付行 + 合計行 + まとめ行)
    # -------------------------------------------------------------
    def _show_month_sheet(self, month):
        """Treeviewをクリアし、このmonthの日数行+合計行+まとめ行を挿入→再計算"""
        if not self.tree:
            return

        # 既存の行をすべて削除
        for item in self.tree.get_children():
            self.tree.delete(item)

        cols = len(self.tree["columns"])  # 列数(14)
        days = self._get_days_in_month(month)  # 当月の日数

        # (1) 日付行を挿入
        for day in range(1, days + 1):
            key_day = f"{self.current_year}-{month}-{day}"  # "2024-3-15"等
            if key_day in self.parent_table_data:
                row_values = self.parent_table_data[key_day]
            else:
                row_values = [str(day)] + [""] * (cols - 1)
            # 挿入
            self.tree.insert("", "end", values=row_values)

        # (2) 合計行(#0="合計")
        total_row = ["合計"] + [""] * (cols - 1)
        # タグ="TOTAL"で色や太字が付く
        self.tree.insert("", "end", values=total_row, tags=("TOTAL",))

        # (3) まとめ行( day=0 として管理 )
        summary_key = f"{self.current_year}-{month}-0"
        if summary_key in self.parent_table_data:
            sm_data = self.parent_table_data[summary_key]
            # "収入金額" は col #3 を使う
            inc_str = sm_data[3] if len(sm_data) > 3 else ""
        else:
            inc_str = ""

        # まとめ行:
        # #0="まとめ"
        # #1=数値sum(収入-支出)
        # #2="収入"
        # #3=収入金額(編集可)
        # #4="支出"
        # #5=総支出
        # #6..#13=空
        summary_row = [
                          "まとめ",
                          "",
                          "収入",
                          inc_str,  # 収入金額
                          "支出",
                          "",
                      ] + [""] * (cols - 6)
        self.tree.insert("", "end", values=summary_row, tags=("SUMMARY",))

        # 再計算
        self._recalc_total_and_summary()

    def _recalc_total_and_summary(self):
        """
        合計行(#1..#13)に日付行の合計を入れ、
        さらに合計行(#1..#13)を合算して「総支出」を出し、
        まとめ行(#3=収入金額, #1=差額, #5=総支出)を更新
        """
        items = self.tree.get_children()
        if len(items) < 2:
            return

        total_row_id = items[-2]  # 合計行
        summary_row_id = items[-1]  # まとめ行
        cols = len(self.tree["columns"])

        # (1) 合計行の計算
        sums = [0] * (cols - 1)  # #1..#13
        for row_id in items[:-2]:  # 日付行のみ合算
            row_vals = self.tree.item(row_id, 'values')
            for i in range(1, cols):
                try:
                    val = int(row_vals[i])
                except (ValueError, TypeError):
                    val = 0
                sums[i - 1] += val

        # 合計行を更新
        total_vals = list(self.tree.item(total_row_id, 'values'))
        for i in range(1, cols):
            if sums[i - 1] == 0:
                total_vals[i] = ""
            else:
                total_vals[i] = str(sums[i - 1])
        self.tree.item(total_row_id, values=total_vals)

        # (2) 合計行(#1..#13) を合算 => 総支出 grand_total
        grand_total = 0
        for v in total_vals[1:]:
            if v.strip():
                try:
                    grand_total += int(v)
                except:
                    pass

        # (3) まとめ行 (#3=収入, #1=差分, #5=総支出)
        summary_vals = list(self.tree.item(summary_row_id, 'values'))
        # 収入金額
        try:
            income_val = int(summary_vals[3])
        except:
            income_val = 0
        # 数値sum = 収入 - 総支出
        sum_val = income_val - grand_total
        if sum_val == 0:
            summary_vals[1] = ""
        else:
            summary_vals[1] = str(sum_val)
        # #5=総支出
        if grand_total == 0:
            summary_vals[5] = ""
        else:
            summary_vals[5] = str(grand_total)

        # 反映
        self.tree.item(summary_row_id, values=summary_vals)

    # -------------------------------------------------------------
    # ダブルクリックイベント (親セル編集)
    # -------------------------------------------------------------
    def _on_parent_double_click(self, event):
        """
        - 合計行(最後-1)は編集禁止
        - まとめ行(最後)はcol #3(=収入金額)のみ編集可
        - 日付列(#1)は編集不可
        - 上記を除けば、子ダイアログに飛ばして編集
        """
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return

        items = self.tree.get_children()
        if len(items) < 2:
            return

        total_row_id = items[-2]  # 合計行
        summary_row_id = items[-1]  # まとめ行

        # 合計行は編集しない
        if row_id == total_row_id:
            return
        # まとめ行 かつ col_id != "#4" (収入金額以外) なら編集禁止
        if row_id == summary_row_id:
            if col_id != "#4":
                return
        # 日付列"#1"は編集スキップ
        if col_id == "#1":
            return

        row_vals = self.tree.item(row_id, 'values')
        if not row_vals:
            return

        # day=0 => まとめ行
        if row_id == summary_row_id:
            day = 0
        else:
            # 通常の日付
            try:
                day = int(row_vals[0])
            except:
                day = 0

        col_index = int(col_id[1:]) - 1
        dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_index}"
        col_name = self.tree.heading(col_id, "text")  # heading("text")で列タイトル


        # 子ダイアログ生成
        ChildDialog(self.root, self, dict_key, col_name)

    # -------------------------------------------------------------
    # 子ダイアログからの値更新を受け取り、parent_table_data + 画面に反映
    # -------------------------------------------------------------
    def update_parent_cell(self, dict_key_day, col_index, new_value):
        """
        - dict_key_day = "year-month-day"
        - day=0 => まとめ行(収入金額)
        - それ以外 => 通常の日付行
        """
        cols = len(self.tree["columns"])
        # 新規キーなら初期化
        if dict_key_day not in self.parent_table_data:
            parts = dict_key_day.split("-")
            day_str = parts[2]
            row_array = [day_str] + [""] * (cols - 1)
            self.parent_table_data[dict_key_day] = row_array
        else:
            row_array = self.parent_table_data[dict_key_day]

        # 指定col_indexに new_value を代入
        row_array[col_index] = str(new_value)
        self.parent_table_data[dict_key_day] = row_array

        # もし表示中の年/月なら、Treeviewにも反映
        y, mo, d = dict_key_day.split("-")
        y = int(y);
        mo = int(mo);
        d = int(d)
        if (self.current_year == y) and (self.current_month == mo):
            items = self.tree.get_children()
            if len(items) < 2:
                return
            total_row_id = items[-2]
            summary_row_id = items[-1]

            found = False
            # 日付行を探す
            for row_id in items[:-2]:
                row_vals = list(self.tree.item(row_id, 'values'))
                if row_vals and row_vals[0] == str(d):
                    row_vals[col_index] = str(new_value)
                    self.tree.item(row_id, values=row_vals)
                    found = True
                    break

            # 見つからず day=0 => まとめ行
            if not found and d == 0:
                sum_vals = list(self.tree.item(summary_row_id, 'values'))
                sum_vals[col_index] = str(new_value)
                self.tree.item(summary_row_id, values=sum_vals)

            # 再計算
            self._recalc_total_and_summary()

    # -------------------------------------------------------------
    # 日数計算(うるう年対応)
    # -------------------------------------------------------------
    def _get_days_in_month(self, month):
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


# ----------------------------------------------------------------------
# 子ダイアログ: 親ウィンドウの中央に配置
# ----------------------------------------------------------------------
class ChildDialog(tk.Toplevel):
    def __init__(self, parent, parent_app, dict_key, col_name):
        """子ダイアログを生成し、親ウィンドウの中央に表示。"""
        super().__init__(parent)
        self.parent_app = parent_app
        self.dict_key = dict_key

        # ダイアログの幅・高さ
        dialog_width = 600
        dialog_height = 300

        # 親ウィンドウの位置とサイズを取得
        parent_x = self.master.winfo_x()
        parent_y = self.master.winfo_y()
        parent_w = self.master.winfo_width()
        parent_h = self.master.winfo_height()

        # (x,y) を「親の中央」になるように計算
        x = parent_x + (parent_w - dialog_width) // 2
        y = parent_y + (parent_h - dialog_height) // 2

        # geometryで子ダイアログのサイズと位置を決定
        self.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        self.title(f"支出・収入詳細({col_name})")

        # dict_key="year-month-day-colIndex" => 分割
        parts = dict_key.split("-")
        if len(parts) == 4:
            self.year = int(parts[0])
            self.month = int(parts[1])
            self.day = int(parts[2])
            self.col_index = int(parts[3])
        else:
            self.year, self.month, self.day, self.col_index = (0, 0, 0, 0)

        self.child_tree = None
        self.entry_editor = None

        # Gridレイアウト
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # ウィジェット生成
        self._create_widgets()
        # データ反映
        self._populate_data()

    def _create_widgets(self):
        """子ダイアログのツリー表示 + OKボタン"""
        columns = ["取引先", "金額", "詳細"]
        self.child_tree = ttk.Treeview(self, columns=columns, show="headings")
        # row=0に配置
        self.child_tree.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        for col in columns:
            self.child_tree.heading(col, text=col)
            self.child_tree.column(col, anchor="center", width=150)

        # スクロールバー(縦)
        sb = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.child_tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.child_tree.configure(yscrollcommand=sb.set)

        self.child_tree.bind("<Double-1>", self._on_double_click)

        # OKボタン(row=1)
        ok_btn = tk.Button(self, text="OK", command=self._on_ok_button)
        ok_btn.grid(row=1, column=0, columnspan=2, pady=10)

    def _populate_data(self):
        """child_data[...] から行リストを取得。空なら空行1行"""
        data_list = self.parent_app.child_data.get(self.dict_key, [])
        if not data_list:
            data_list = [("", "", "")]
        for row in data_list:
            self.child_tree.insert("", "end", values=row)

    def _on_double_click(self, event):
        """ツリーセルをダブルクリック -> Entry配置で直接編集"""
        item_id = self.child_tree.identify_row(event.y)
        col_id = self.child_tree.identify_column(event.x)
        if not item_id or not col_id:
            return

        bbox = self.child_tree.bbox(item_id, col_id)
        if not bbox:
            return

        col_idx = int(col_id[1:]) - 1
        old_vals = list(self.child_tree.item(item_id, 'values'))

        # 他に編集中のEntryがあれば破棄
        if self.entry_editor:
            self.entry_editor.destroy()

        x, y, w, h = bbox
        # EntryをTreeview上にPlace
        self.entry_editor = tk.Entry(self.child_tree, font=("Arial", 12))
        self.entry_editor.place(x=x, y=y, width=w, height=h)
        self.entry_editor.insert(0, old_vals[col_idx])
        self.entry_editor.focus()

        # Enter or フォーカスアウトで確定
        self.entry_editor.bind("<Return>", lambda e: self._save_entry(item_id, col_idx))
        self.entry_editor.bind("<FocusOut>", lambda e: self._save_entry(item_id, col_idx))

    def _save_entry(self, item_id, col_idx):
        """編集が確定 -> Treeviewに反映。最終行が埋まったら新行を追加"""
        if not self.entry_editor:
            return
        new_val = self.entry_editor.get()
        self.entry_editor.destroy()
        self.entry_editor = None

        row_vals = list(self.child_tree.item(item_id, 'values'))
        row_vals[col_idx] = new_val
        self.child_tree.item(item_id, values=row_vals)

        # 最終行なら、かつ空じゃなくなったら新行追加
        all_items = self.child_tree.get_children()
        if item_id == all_items[-1]:
            if any(cell.strip() for cell in row_vals):
                self._add_empty_row()

    def _add_empty_row(self):
        """空行を一つ追加(ユーザが下へ追加入力可能に)"""
        self.child_tree.insert("", "end", values=["", "", ""])

    def _on_ok_button(self):
        """OKボタン: 入力行をすべて child_data[...]に保存 & 金額合算 -> 親セル更新 -> 閉じる"""
        all_rows = []
        for iid in self.child_tree.get_children():
            vals = self.child_tree.item(iid, 'values')
            all_rows.append(vals)

        # データ保存
        self.parent_app.child_data[self.dict_key] = all_rows

        # "金額"(col #1) を合計
        total = 0
        for row in all_rows:
            try:
                total += int(row[1])
            except ValueError:
                pass

        # parentへ通知: "year-month-day"
        dict_key_day = f"{self.year}-{self.month}-{self.day}"
        self.parent_app.update_parent_cell(dict_key_day, self.col_index, total)

        self.destroy()


# ----------------------------------------------------------------------
# メイン実行ブロック
# ----------------------------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()  # ルートウィンドウ生成
    app = YearApp(root)  # メインアプリ起動
    root.mainloop()  # イベントループ開始
