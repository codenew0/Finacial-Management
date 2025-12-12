# main.py
"""
家計管理アプリケーションのエントリーポイント
"""
import tkinter as tk
from tkinter import messagebox
import socket
import sys
from ui.main_window import MainWindow
from utils.font_utils import setup_japanese_font

def check_single_instance():
    """
    アプリケーションの二重起動を防止する。
    特定のポートにバインドを試み、失敗した場合は既に起動していると判断する。
    """
    # 任意のポート番号（他と被りにくい番号を使用）
    PORT = 54321
    
    try:
        # ソケットを作成
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        # ポートにバインド（この処理が成功している間は、他からのバインドは失敗する）
        sock.bind(('127.0.0.1', PORT))
        return sock
    except socket.error:
        return None

def main():
    """アプリケーションのメインエントリーポイント"""
    
    # 二重起動チェック
    instance_lock = check_single_instance()
    if not instance_lock:
        # ルートウィンドウを一時的に作成してメッセージを表示
        root = tk.Tk()
        root.withdraw()  # ウィンドウ本体は表示しない
        messagebox.showwarning("警告", "アプリケーションは既に起動しています。")
        root.destroy()
        sys.exit(0)

    # 日本語フォントを設定
    setup_japanese_font()
    
    try:
        # Tkinterのルートウィンドウを作成
        root = tk.Tk()
        
        # メインウィンドウを初期化
        app = MainWindow(root)
        
        # メインループを開始
        root.mainloop()
        
    except Exception as e:
        # エラーが発生した場合は詳細を表示
        print(f"アプリケーション実行エラー: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # アプリケーション終了時にソケットを閉じる
        if instance_lock:
            instance_lock.close()

if __name__ == "__main__":
    main()