# utils/font_utils.py
"""
フォント設定ユーティリティ
"""
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import warnings
from config import FontConfig

# matplotlibのフォント警告を抑制
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
        
        # 利用可能な日本語フォントを検索
        for font in FontConfig.JAPANESE_FONTS:
            if font in available_fonts:
                plt.rcParams['font.family'] = font
                return
        
        # 日本語フォントが見つからない場合のフォールバック
        plt.rcParams['font.family'] = FontConfig.FALLBACK_FONT
    except Exception:
        # エラーが発生した場合もフォールバック
        plt.rcParams['font.family'] = FontConfig.FALLBACK_FONT