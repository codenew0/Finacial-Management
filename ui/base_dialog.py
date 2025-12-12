# ui/base_dialog.py
"""
ダイアログの基底クラス
すべてのダイアログで共通の機能を提供
"""
import tkinter as tk
from config import DialogConfig


class BaseDialog(tk.Toplevel):
    """
    すべてのダイアログの基底クラス。
    
    共通の初期化処理(中央配置、モーダル設定など)を提供し、
    各ダイアログクラスはこのクラスを継承することで、
    一貫したUIと動作を実現する。
    """
    
    def __init__(self, parent, title, width=None, height=None):
        """
        ダイアログの基本設定を行う。
        
        Args:
            parent: 親ウィンドウ
            title: ダイアログのタイトル
            width: ダイアログの幅(デフォルト: DialogConfig.DEFAULT_WIDTH)
            height: ダイアログの高さ(デフォルト: DialogConfig.DEFAULT_HEIGHT)
        """
        super().__init__(parent)
        self.title(title)
        self.configure(bg='#f0f0f0')
        
        # デフォルトサイズの設定
        if width is None:
            width = DialogConfig.DEFAULT_WIDTH
        if height is None:
            height = DialogConfig.DEFAULT_HEIGHT
        
        # ダイアログを親ウィンドウの中央に配置
        self._center_on_parent(width, height)
        
        # モーダルダイアログとして設定
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)
        
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
        
        # 最小サイズを設定(元のサイズの75%)
        min_width = int(width * DialogConfig.MIN_SIZE_RATIO)
        min_height = int(height * 0.67)
        self.minsize(min_width, min_height)