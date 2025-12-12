# models/data_manager.py
"""
データ管理クラス
家計データの読み込み、保存、検索などを担当
"""
import json
import os
import shutil
from datetime import datetime
from config import DATA_FILE, SETTINGS_FILE, APP_VERSION


class DataManager:
    """家計データの管理を担当するクラス"""
    
    def __init__(self):
        """データマネージャーの初期化"""
        self.data = {}  # 詳細データを格納する辞書 {key: [[partner, amount, detail], ...]}
        self.custom_columns = []  # カスタム項目リスト
        self.transaction_partners = set()  # 支払先の履歴

        # 初期化時にデータフォルダの存在を確認・作成する
        self._ensure_data_directory()

    def _ensure_data_directory(self):
        """データ保存先のディレクトリが存在しない場合、作成する"""
        # DATA_FILEのパスからディレクトリ部分だけを取り出す
        directory = os.path.dirname(DATA_FILE)
        
        # ディレクトリパスがあり、かつ実際には存在しない場合
        if directory and not os.path.exists(directory):
            try:
                # フォルダを作成 (parents=Trueで親フォルダも作成、exist_ok=Trueで競合エラー防止)
                os.makedirs(directory, exist_ok=True)
            except OSError as e:
                print(f"データフォルダ作成エラー: {e}")

    def _create_backup(self):
        """
        data.jsonのバックアップを作成する。
        起動時に毎回実行される。
        """
        if not os.path.exists(DATA_FILE):
            return

        # バックアップディレクトリのパス (household_json/backups)
        data_dir = os.path.dirname(DATA_FILE)
        backup_dir = os.path.join(data_dir, os.path.dirname(DATA_FILE), "backups")
        
        if not os.path.exists(backup_dir):
            try:
                os.makedirs(backup_dir)
            except OSError as e:
                print(f"バックアップフォルダ作成エラー: {e}")
                return

        # 日時付きのファイル名を作成 (例: data_20251207_123000.json)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"data_{timestamp}.json"
        backup_path = os.path.join(backup_dir, backup_filename)

        try:
            shutil.copy2(DATA_FILE, backup_path)
            # 古いバックアップの削除（オプション：最新30件だけ残すなど）
            self._cleanup_old_backups(backup_dir)
        except Exception as e:
            print(f"バックアップ作成エラー: {e}")

    def _cleanup_old_backups(self, backup_dir):
        """古いバックアップを削除して容量を節約（最新50件を残す）"""
        try:
            files = [os.path.join(backup_dir, f) for f in os.listdir(backup_dir) if f.startswith("data_") and f.endswith(".json")]
            files.sort(key=os.path.getmtime) # 古い順にソート
            
            # 50件を超えている場合、古いものを削除
            max_backups = 50
            if len(files) > max_backups:
                for f in files[:-max_backups]:
                    os.remove(f)
        except Exception:
            pass # クリーンアップ失敗は無視

    def _cleanup_old_backups(self, backup_dir):
        """古いバックアップを削除して容量を節約（最新50件を残す）"""
        try:
            files = [os.path.join(backup_dir, f) for f in os.listdir(backup_dir) if f.startswith("data_") and f.endswith(".json")]
            files.sort(key=os.path.getmtime) # 古い順にソート
            
            # 50件を超えている場合、古いものを削除
            max_backups = 50
            if len(files) > max_backups:
                for f in files[:-max_backups]:
                    os.remove(f)
        except Exception:
            pass # クリーンアップ失敗は無視    
    
    def load_data(self):
        """
        データファイルから家計データを読み込む。
        読み込み前に自動バックアップを行う。
        """
        # データのバックアップを作成
        self._create_backup()

        if not os.path.exists(DATA_FILE):
            return
        
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                all_data = json.load(f)
            
            if "version" in all_data:
                # 新フォーマット
                self.data = all_data.get("data", {})
            else:
                # 旧フォーマット(後方互換性)
                self.data = all_data.get("data", {})
                self._extract_partners_from_data()
        except Exception as e:
            print(f"データ読み込みエラー: {e}")
            # 読み込みエラーは無視(空のデータで開始)
    
    def save_data(self):
        """
        家計データをファイルに保存する。
        
        child_dataをJSONファイルに保存する。
        バージョン情報を含めて新フォーマットで保存。
        """
        all_data = {
            "version": APP_VERSION,
            "data": self.data
        }
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(all_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"データ保存エラー: {e}")
    
    def load_settings(self):
        """
        設定ファイルから設定を読み込む。
        
        カスタム項目と支払先履歴を復元する。
        """
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    settings = json.load(f)
                    self.custom_columns = settings.get("custom_columns", [])
                    self.transaction_partners = set(settings.get("transaction_partners", []))
            except Exception as e:
                print(f"設定読み込みエラー: {e}")
    
    def save_settings(self):
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
        except Exception as e:
            print(f"設定保存エラー: {e}")
    
    def get_transaction_data(self, dict_key):
        """
        指定されたキーの取引データを取得する。
        
        Args:
            dict_key (str): データのキー(年-月-日-列インデックス)
            
        Returns:
            list: 取引データのリスト
        """
        return self.data.get(dict_key, [])
    
    def set_transaction_data(self, dict_key, data_list):
        """
        指定されたキーに取引データを設定する。
        
        Args:
            dict_key (str): データのキー(年-月-日-列インデックス)
            data_list (list): 設定する取引データのリスト
        """
        if data_list:
            self.data[dict_key] = data_list
        elif dict_key in self.data:
            del self.data[dict_key]
    
    def delete_transaction_data(self, dict_key):
        """
        指定されたキーの取引データを削除する。
        
        Args:
            dict_key (str): データのキー
        """
        if dict_key in self.data:
            del self.data[dict_key]
    
    def add_transaction_partner(self, partner):
        """
        支払先を履歴に追加する。
        
        Args:
            partner (str): 支払先名
        """
        if partner and partner.strip():
            self.transaction_partners.add(partner.strip())
    
    def get_transaction_partners_list(self):
        """
        支払先の履歴をソート済みリストで取得する。
        
        Returns:
            list: ソート済み支払先リスト
        """
        return sorted(list(self.transaction_partners))
    
    def add_custom_column(self, column_name):
        """
        カスタム項目を追加する。
        
        Args:
            column_name (str): 追加する項目名
            
        Returns:
            bool: 追加成功時True、既存の場合False
        """
        if column_name and column_name not in self.custom_columns:
            self.custom_columns.append(column_name)
            return True
        return False
    
    def edit_custom_column(self, old_name, new_name):
        """
        カスタム項目名を編集する。
        
        Args:
            old_name (str): 旧項目名
            new_name (str): 新項目名
            
        Returns:
            bool: 編集成功時True
        """
        if old_name in self.custom_columns and new_name not in self.custom_columns:
            index = self.custom_columns.index(old_name)
            self.custom_columns[index] = new_name
            return True
        return False
    
    def delete_custom_column(self, column_name):
        """
        カスタム項目を削除する。
        
        Args:
            column_name (str): 削除する項目名
            
        Returns:
            bool: 削除成功時True
        """
        if column_name in self.custom_columns:
            self.custom_columns.remove(column_name)
            return True
        return False
    
    def delete_column_data(self, col_index):
        """
        指定された列の全データを削除する。
        
        Args:
            col_index (int): 列インデックス
        """
        keys_to_delete = [key for key in self.data.keys()
                          if key.split("-")[3] == str(col_index)]
        for key in keys_to_delete:
            del self.data[key]
    
    def search_transactions(self, search_text):
        """
        取引データを検索する。
        
        Args:
            search_text (str): 検索文字列
            
        Returns:
            list: 検索結果のリスト
        """
        results = []
        search_text_lower = search_text.lower()
        
        for dict_key, data_list in self.data.items():
            try:
                parts = dict_key.split("-")
                if len(parts) == 4:
                    year, month, day, col_index = map(int, parts)
                    
                    for row in data_list:
                        if len(row) >= 3:
                            partner = str(row[0]).strip() if row[0] else ""
                            amount = str(row[1]).strip() if row[1] else ""
                            detail = str(row[2]).strip() if row[2] else ""
                            
                            if (search_text_lower in partner.lower() or
                                search_text_lower in amount.lower() or
                                search_text_lower in detail.lower()):
                                results.append({
                                    'year': year,
                                    'month': month,
                                    'day': day,
                                    'col_index': col_index,
                                    'partner': partner,
                                    'amount': amount,
                                    'detail': detail
                                })
            except (ValueError, IndexError):
                continue
        
        return results
    
    def _extract_partners_from_data(self):
        """
        既存のデータから支払先を抽出して履歴に追加する。
        
        旧フォーマットのデータから支払先情報を抽出するための互換性処理。
        """
        for data_list in self.data.values():
            for row in data_list:
                if len(row) > 0 and row[0] and str(row[0]).strip():
                    self.transaction_partners.add(str(row[0]).strip())