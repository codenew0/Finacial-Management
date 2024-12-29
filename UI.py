import tkinter as tk
from tkinter import ttk
from datetime import datetime
import json
import os

DATA_FILE = "data.json"


class YearApp:
    def __init__(self, root):
        """メインウィンドウ初期化"""
        self.root = root
        self._setup_root()
        self._init_variables()
        self._load_data_from_file()

        self._create_frames_and_widgets()
        self._show_month_sheet(self.current_month)

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    # -------------------------------------------------------------
    # ウィンドウ設定
    # -------------------------------------------------------------
    def _setup_root(self):
        self.root.title("年数を変更するアプリ (合計 + まとめ行 + 収入金額保存)")
        self.root.geometry("1300x1000")
        self.root.resizable(False, False)

    def _init_variables(self):
        """変数初期化"""
        now = datetime.now()
        self.current_year = now.year
        self.current_month = now.month

        # 親Treeview
        self.tree = None

        # 子ダイアログの明細用 (dict_key="year-month-day-colIndex")
        self.child_data = {}

        # 親ダイアログの各セルデータを保存
        #  - 日付行  => key="year-month-day" (day>=1)
        #  - まとめ => key="year-month-0"   (day=0)
        self.parent_table_data = {}

    # -------------------------------------------------------------
    # JSON 入出力
    # -------------------------------------------------------------
    def _load_data_from_file(self):
        """ファイルがあれば読み込む"""
        if not os.path.exists(DATA_FILE):
            return
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            all_data = json.load(f)
        self.child_data = all_data.get("child_data", {})
        self.parent_table_data = all_data.get("parent_table_data", {})

    def _save_data_to_file(self):
        """child_data, parent_table_data を保存"""
        all_data = {
            "child_data": self.child_data,
            "parent_table_data": self.parent_table_data
        }
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(all_data, f, ensure_ascii=False, indent=2)

    def _on_closing(self):
        """ウィンドウ終了時ハンドラ"""
        self._save_data_to_file()
        self.root.destroy()

    # -------------------------------------------------------------
    # ウィジェット作成
    # -------------------------------------------------------------
    def _create_frames_and_widgets(self):
        top_frame = tk.Frame(self.root)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        self.minus_button = tk.Button(
            top_frame, text="◀", font=("Arial", 20),
            command=self._decrease_year
        )
        self.minus_button.pack(side=tk.LEFT, padx=(0, 10))

        self.year_label = tk.Label(
            top_frame, text=str(self.current_year), font=("Arial", 24)
        )
        self.year_label.pack(side=tk.LEFT, padx=10)

        self.plus_button = tk.Button(
            top_frame, text="▶", font=("Arial", 20),
            command=self._increase_year
        )
        self.plus_button.pack(side=tk.LEFT, padx=(10, 0))

        self.current_month_label = tk.Label(
            top_frame, text=f"選択中の月：{self.current_month:02d}月", font=("Arial", 16)
        )
        self.current_month_label.pack(side=tk.LEFT, padx=30)

        # 月ボタンフレーム
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

        # Treeviewエリア
        tree_frame = tk.Frame(self.root)
        tree_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self._create_tree_if_needed(tree_frame)

        # 月ボタンのハイライト初期化
        self._highlight_selected_month()

    def _create_tree_if_needed(self, parent):
        """親のTreeview (self.tree) を初回のみ作成"""
        if self.tree is not None:
            return

        columns = [
            "日付", "交通", "外食", "食品", "日常用品", "通販", "ゲーム課金",
            "ゲーム購入", "サービス", "家賃", "公共料金", "携帯・回線", "保険", "他"
        ]
        self.tree = ttk.Treeview(parent, columns=columns, show="headings")

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=len(col) * 10)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)

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

        # 合計行タグ
        self.tree.tag_configure(
            "TOTAL",
            background="lightyellow",
            font=("Arial", 12, "bold")
        )
        # まとめ行タグ
        self.tree.tag_configure(
            "SUMMARY",
            background="lightgreen",
            font=("Arial", 12, "bold")
        )

        # ダブルクリックで編集フロー
        self.tree.bind("<Double-1>", self._on_parent_double_click)

    # -------------------------------------------------------------
    # 年と月の変更
    # -------------------------------------------------------------
    def _increase_year(self):
        self.current_year += 1
        self.year_label.config(text=str(self.current_year))
        self._show_month_sheet(self.current_month)

    def _decrease_year(self):
        self.current_year -= 1
        self.year_label.config(text=str(self.current_year))
        self._show_month_sheet(self.current_month)

    def _select_month(self, month):
        self.current_month = month
        self.current_month_label.config(text=f"選択中の月：{month:02d}月")
        self._highlight_selected_month()
        self._show_month_sheet(month)

    def _highlight_selected_month(self):
        """月ボタンの背景色をリセットし、選択中のみ色付け"""
        for i, btn in enumerate(self.month_buttons, start=1):
            btn.config(bg="SystemButtonFace")
        self.month_buttons[self.current_month - 1].config(bg="lightblue")

    # -------------------------------------------------------------
    # 月のシート表示
    # -------------------------------------------------------------
    def _show_month_sheet(self, month):
        """Treeviewをクリアして日付行,合計行,まとめ行を表示"""
        if not self.tree:
            return

        # 全削除
        for item in self.tree.get_children():
            self.tree.delete(item)

        cols = len(self.tree["columns"])
        days_in_month = self._get_days_in_month(month)

        # 日付行 (#1...#13)
        for day in range(1, days_in_month + 1):
            key_day = f"{self.current_year}-{month}-{day}"
            if key_day in self.parent_table_data:
                row_values = self.parent_table_data[key_day]
            else:
                row_values = [str(day)] + [""] * (cols - 1)
            self.tree.insert("", "end", values=row_values)

        # 合計行 (#0="合計")
        total_row = ["合計"] + [""] * (cols - 1)
        self.tree.insert("", "end", values=total_row, tags=("TOTAL",))

        # まとめ行 (day=0) から「収入金額」を読み込み→表示
        # 事前に parent_table_data["year-month-0"] があれば使う
        summary_key = f"{self.current_year}-{month}-0"
        if summary_key in self.parent_table_data:
            summary_data = self.parent_table_data[summary_key]
            # summary_data は [ "0", "", "", 収入金額, ... ] の想定
            # 収入金額は まとめ行で col #3 に入れたい
            income_str = summary_data[3] if len(summary_data) > 3 else ""
        else:
            # なければ空
            income_str = ""

        # まとめ行 #0="まとめ" #1=数値sum #2="収入" #3=収入金額 #4="支出" #5=合計(総支出) ...
        summary_row = [
                          "まとめ",  # col #0
                          "",  # col #1 => 数値sum(あとで計算)
                          "収入",  # col #2
                          income_str,  # col #3 => 収入金額
                          "支出",  # col #4
                          "",  # col #5 => 合計値(あとで計算)
                      ] + [""] * (cols - 6)
        self.tree.insert("", "end", values=summary_row, tags=("SUMMARY",))

        # 再計算
        self._recalc_total_and_summary()

    # -------------------------------------------------------------
    # 合計行 + まとめ行 を再計算
    # -------------------------------------------------------------
    def _recalc_total_and_summary(self):
        """
        1) 合計行 (#1..#13) = 全日付行の合計
        2) 合計行(#1..#13)をさらに合算して「総支出」とする
        3) まとめ行の 収入金額(#3) との差を #1(数値sum) に、 #5(支出) に総支出
        """
        items = self.tree.get_children()
        if len(items) < 2:
            return

        # 合計行は items[-2], まとめ行は items[-1]
        total_row_id = items[-2]
        summary_row_id = items[-1]

        cols = len(self.tree["columns"])

        # 1) 合計行
        sums = [0] * (cols - 1)
        for row_id in items[:-2]:
            row_vals = self.tree.item(row_id, 'values')
            for i in range(1, cols):
                try:
                    val = int(row_vals[i])
                except (ValueError, TypeError):
                    val = 0
                sums[i - 1] += val

        total_vals = list(self.tree.item(total_row_id, 'values'))
        for i in range(1, cols):
            if sums[i - 1] == 0:
                total_vals[i] = ""
            else:
                total_vals[i] = str(sums[i - 1])
        self.tree.item(total_row_id, values=total_vals)

        # 2) 合計行(#1..#13)を合計 → 総支出 grand_total
        grand_total = 0
        for valstr in total_vals[1:]:
            if valstr.strip():
                try:
                    grand_total += int(valstr)
                except ValueError:
                    pass

        # 3) まとめ行
        summary_vals = list(self.tree.item(summary_row_id, 'values'))
        # 収入金額(#3)
        income_str = summary_vals[3]
        try:
            income_val = int(income_str)
        except ValueError:
            income_val = 0

        # 数値sum(#1) = income_val - grand_total
        sum_val = income_val - grand_total
        if sum_val == 0:
            summary_vals[1] = ""
        else:
            summary_vals[1] = str(sum_val)

        # 支出(#5) = grand_total
        if grand_total == 0:
            summary_vals[5] = ""
        else:
            summary_vals[5] = str(grand_total)

        self.tree.item(summary_row_id, values=summary_vals)

    # -------------------------------------------------------------
    # ダブルクリック -> 親ダイアログ編集
    # -------------------------------------------------------------
    def _on_parent_double_click(self, event):
        """
        - 合計行(最後から2番目)は編集しない
        - まとめ行(最後)は col #3(=収入金額)だけ編集可
        - 日付列(#1)も編集しない
        - それ以外の日付行を編集 → 子ダイアログ
        """
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return

        items = self.tree.get_children()
        if len(items) < 2:
            return

        total_row_id = items[-2]
        summary_row_id = items[-1]

        # 合計行の場合
        if row_id == total_row_id:
            return

        # まとめ行の場合、col #3(= "#4")だけ編集許可
        if row_id == summary_row_id:
            if col_id != "#4":
                return

        # 日付列(#1)は編集しない
        if col_id == "#1":
            return

        # 通常処理
        row_vals = self.tree.item(row_id, 'values')
        if not row_vals:
            return

        # "まとめ行"だと day=0 として扱う
        if row_id == summary_row_id:
            day = 0
        else:
            # 通常の日付行
            try:
                day = int(row_vals[0])
            except ValueError:
                day = 0

        col_index = int(col_id[1:]) - 1
        dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_index}"

        ChildDialog(self.root, self, dict_key)

    # -------------------------------------------------------------
    # 子ダイアログからの更新
    # -------------------------------------------------------------
    def update_parent_cell(self, dict_key_day, col_index, new_value):
        """
        dict_key_day = "year-month-day"
        もし day=0 => まとめ行
        それ以外 => 日付行
        """
        cols = len(self.tree["columns"])
        if dict_key_day not in self.parent_table_data:
            # 新規
            parts = dict_key_day.split("-")
            day_str = parts[2]
            # まとめ行の場合 day=0 => row_array[0]="0" としておくが実際表示されない
            row_array = [day_str] + [""] * (cols - 1)
            self.parent_table_data[dict_key_day] = row_array
        else:
            row_array = self.parent_table_data[dict_key_day]

        row_array[col_index] = str(new_value)
        self.parent_table_data[dict_key_day] = row_array

        # 表示中なら画面に反映
        y, mo, d = dict_key_day.split("-")
        y = int(y)
        mo = int(mo)
        d = int(d)
        if (self.current_year == y) and (self.current_month == mo):
            items = self.tree.get_children()
            if len(items) < 2:
                return
            total_row_id = items[-2]
            summary_row_id = items[-1]

            # 日付行探し
            found = False
            for row_id in items[:-2]:
                row_vals = list(self.tree.item(row_id, 'values'))
                if row_vals and row_vals[0] == str(d):
                    row_vals[col_index] = str(new_value)
                    self.tree.item(row_id, values=row_vals)
                    found = True
                    break

            if not found and d == 0:
                # まとめ行
                summary_vals = list(self.tree.item(summary_row_id, 'values'))
                summary_vals[col_index] = str(new_value)
                self.tree.item(summary_row_id, values=summary_vals)

            # 最後に再計算
            self._recalc_total_and_summary()

    # -------------------------------------------------------------
    # 日数計算
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
# 子ダイアログ (収入 or 各日付行のセル編集)
# ----------------------------------------------------------------------
class ChildDialog(tk.Toplevel):
    def __init__(self, parent, parent_app, dict_key):
        """子ダイアログ (場所, 金額, 詳細) を複数行追加できる"""
        super().__init__(parent)
        self.title("子ダイアログ(まとめ行 or 日付行)")
        self.geometry("600x300")

        self.parent_app = parent_app
        parts = dict_key.split("-")
        # 例: "2024-3-0-3" => 2024年3月 day=0 col_index=3 => 収入金額 など
        if len(parts) == 4:
            self.year = int(parts[0])
            self.month = int(parts[1])
            self.day = int(parts[2])
            self.col_index = int(parts[3])
        else:
            self.year, self.month, self.day, self.col_index = (0, 0, 0, 0)

        self.dict_key = dict_key
        self.child_tree = None
        self.entry_editor = None

        # Gridレイアウト
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self._create_widgets()

    def _create_widgets(self):
        columns = ["場所", "金額", "詳細"]
        self.child_tree = ttk.Treeview(self, columns=columns, show="headings")
        self.child_tree.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        for col in columns:
            self.child_tree.heading(col, text=col)
            self.child_tree.column(col, anchor="center", width=150)

        # スクロールバー
        sb = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.child_tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.child_tree.configure(yscrollcommand=sb.set)

        # ダブルクリックでセル編集
        self.child_tree.bind("<Double-1>", self._on_double_click)

        # OKボタン
        ok_btn = tk.Button(self, text="OK", command=self._on_ok_button)
        ok_btn.grid(row=1, column=0, columnspan=2, pady=10)

        # データ反映
        self._populate_data()

    def _populate_data(self):
        """child_data から読み込み。空なら1行だけ"""
        data_list = self.parent_app.child_data.get(self.dict_key, [])
        if not data_list:
            data_list = [("", "", "")]
        for row in data_list:
            self.child_tree.insert("", "end", values=row)

    def _on_double_click(self, event):
        """Treeviewセルをダブルクリック -> Entryで編集"""
        item_id = self.child_tree.identify_row(event.y)
        col_id = self.child_tree.identify_column(event.x)
        if not item_id or not col_id:
            return

        bbox = self.child_tree.bbox(item_id, col_id)
        if not bbox:
            return

        col_idx = int(col_id[1:]) - 1
        old_vals = list(self.child_tree.item(item_id, 'values'))

        if self.entry_editor:
            self.entry_editor.destroy()

        x, y, w, h = bbox
        self.entry_editor = tk.Entry(self.child_tree, font=("Arial", 12))
        self.entry_editor.place(x=x, y=y, width=w, height=h)
        self.entry_editor.insert(0, old_vals[col_idx])
        self.entry_editor.focus()

        self.entry_editor.bind("<Return>", lambda e: self._save_entry(item_id, col_idx))
        self.entry_editor.bind("<FocusOut>", lambda e: self._save_entry(item_id, col_idx))

    def _save_entry(self, item_id, col_idx):
        """Entry終了 -> ツリーにセット。もし最後の行が埋まったら新行を追加"""
        if not self.entry_editor:
            return
        new_val = self.entry_editor.get()
        self.entry_editor.destroy()
        self.entry_editor = None

        row_vals = list(self.child_tree.item(item_id, 'values'))
        row_vals[col_idx] = new_val
        self.child_tree.item(item_id, values=row_vals)

        all_items = self.child_tree.get_children()
        if item_id == all_items[-1]:
            # 空じゃなければ追加
            if any(cell.strip() for cell in row_vals):
                self._add_empty_row()

    def _add_empty_row(self):
        """空行追加"""
        self.child_tree.insert("", "end", values=["", "", ""])

    def _on_ok_button(self):
        """OK押下 -> 子ツリー全行→child_data[...] に保存→金額合計→親セル更新→閉じる"""
        all_rows = []
        for iid in self.child_tree.get_children():
            vals = self.child_tree.item(iid, 'values')
            all_rows.append(vals)

        self.parent_app.child_data[self.dict_key] = all_rows

        # "金額"列(#1)を合算
        total = 0
        for row in all_rows:
            try:
                total += int(row[1])
            except ValueError:
                pass

        # dict_key_day = "year-month-day"
        dict_key_day = f"{self.year}-{self.month}-{self.day}"
        self.parent_app.update_parent_cell(dict_key_day, self.col_index, total)

        self.destroy()


# ----------------------------------------------------------------------
# 実行
# ----------------------------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = YearApp(root)
    root.mainloop()
