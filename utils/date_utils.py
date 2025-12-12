# utils/date_utils.py
"""
日付関連のユーティリティ関数
"""


def get_days_in_month(year, month):
    """
    指定された年月の日数を取得する。
    
    うるう年を考慮して正確な日数を返す。
    
    Args:
        year (int): 年
        month (int): 月(1-12)
        
    Returns:
        int: その月の日数
    """
    if month in [1, 3, 5, 7, 8, 10, 12]:
        return 31
    elif month in [4, 6, 9, 11]:
        return 30
    elif month == 2:
        # うるう年の判定
        return 29 if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0) else 28
    return 30


def is_leap_year(year):
    """
    うるう年かどうかを判定する。
    
    Args:
        year (int): 年
        
    Returns:
        bool: うるう年の場合True
    """
    return (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0)