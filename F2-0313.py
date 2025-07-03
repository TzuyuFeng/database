import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
import pyodbc
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import json
from openpyxl import Workbook
import numpy as np

class FactoryComparison:
    def __init__(self):
        self.factory1_data = {}  # 彰化廠
        self.factory2_data = {}  # 台南廠
        self.conn = None
        self.config_file = 'database_config.json'
        self.excel_config_file = 'excel_config.json'
        self.db_path = self.load_db_path()
        self.excel_path = self.load_excel_path()
        self.date_ranges = {}
        self.estimated_orders = {}  # 儲存預估訂單數據
        # 添加預設比例設定
        self.ratio_settings = self.load_ratio_settings()
        self.factory1_estimated_capacity = 1000  # 彰化廠每週預估材數
        self.factory2_estimated_capacity = 1200  # 台南廠每週預估材數
        self.main_data_df = None

    def load_ratio_settings(self):
        """載入比例設定"""
        try:
            if os.path.exists('ratio_settings.json'):
                with open('ratio_settings.json', 'r') as f:
                    settings = json.load(f)
                    # 若有最大產能設定則帶入
                    if 'factory1_max_capacity' in settings:
                        self.factory1_max_capacity = settings['factory1_max_capacity']
                    if 'factory2_max_capacity' in settings:
                        self.factory2_max_capacity = settings['factory2_max_capacity']
                    return settings
        except:
            pass
        return {'upper': 2.2, 'lower': 1.8}  # 預設值1

    def load_db_path(self):
        """從配置檔案載入資料庫路徑"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    return config.get('db_path', r'\\Windows-cqa6dgu\拆單軟體\eiffel.mdb')
        except:
            pass
        return r'\\Windows-cqa6dgu\拆單軟體\eiffel.mdb'

    def load_excel_path(self):
        """從配置檔案載入預估訂單資料庫路徑（ACCDB/MDB）"""
        try:
            if os.path.exists(self.excel_config_file):
                with open(self.excel_config_file, 'r') as f:
                    config = json.load(f)
                    return config.get('excel_path', r'C:\eiffelAccess\未拆預估訂單.accdb')
        except:
            pass
        return r'C:\eiffelAccess\未拆預估訂單.accdb'

    def save_db_path(self):
        """保存資料庫路徑到配置檔案"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump({'db_path': self.db_path}, f)
        except Exception as e:
            print(f"保存資料庫路徑時出錯：{str(e)}")

    def save_excel_path(self):
        """保存預估訂單資料庫路徑到配置檔案"""
        try:
            with open(self.excel_config_file, 'w') as f:
                json.dump({'excel_path': self.excel_path}, f)
        except Exception as e:
            print(f"保存預估訂單資料庫路徑時出錯：{str(e)}")

    def save_ratio_settings(self):
        """保存比例設定（含最大產能）"""
        try:
            # 一併存最大產能
            self.ratio_settings['factory1_max_capacity'] = getattr(self, 'factory1_max_capacity', None)
            self.ratio_settings['factory2_max_capacity'] = getattr(self, 'factory2_max_capacity', None)
            with open('ratio_settings.json', 'w') as f:
                json.dump(self.ratio_settings, f)
            print("比例設定已保存")
        except Exception as e:
            print(f"保存比例設定時出錯：{str(e)}")

    def select_database(self):
        """選擇資料庫檔案"""
        try:
            root = tk.Tk()
            root.attributes('-topmost', True)
            root.withdraw()
            
            initial_dir = os.path.dirname(self.db_path) if self.db_path else os.getcwd()
            
            file_path = filedialog.askopenfilename(
                title='選擇資料庫檔案',
                initialdir=initial_dir,
                filetypes=[('Access Database', '*.mdb'), ('All files', '*.*')],
                parent=root
            )
            
            root.destroy()
            
            if file_path:
                self.db_path = file_path
                self.save_db_path()
                return True
            return False
            
        except Exception as e:
            print(f"選擇資料庫時發生錯誤：{str(e)}")
            return False

    def select_excel_file(self):
        """選擇預估訂單ACCDB檔案"""
        try:
            root = tk.Tk()
            root.attributes('-topmost', True)
            root.withdraw()
            initial_dir = os.path.dirname(self.excel_path) if self.excel_path else os.getcwd()
            file_path = filedialog.askopenfilename(
                title='選擇預估訂單ACCDB檔案',
                initialdir=initial_dir,
                filetypes=[('Access Database', '*.accdb;*.mdb'), ('All files', '*.*')],
                parent=root
            )
            root.destroy()
            if file_path:
                self.excel_path = file_path
                self.save_excel_path()
                return True
            return False
        except Exception as e:
            print(f"選擇預估訂單資料庫檔案時發生錯誤：{str(e)}")
            return False

    def connect_to_database(self):
        """連接到 Access 資料庫"""
        try:
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={self.db_path};'
            )
            self.conn = pyodbc.connect(conn_str)
            print("成功連接到資料庫！")
            return True
        except Exception as e:
            print(f"連接資料庫時出錯：{str(e)}")
            retry = input("是否要重新選擇資料庫路徑？(y/n): ")
            if retry.lower() == 'y':
                if self.select_database():
                    return self.connect_to_database()
            return False

    def load_data_from_database(self):
        """從資料庫載入數據"""
        try:
            cursor = self.conn.cursor()
            query = """
            SELECT 
                ev1020_03 as 出貨日期,
                ev1020_88 as 廠別,
                ev1020_07 as 材數,
                ev1020_13 as 生產性質,
                ev1020_20 as 門市,
                ev1020_11 as 圖號,
                ev1020_12 as 色號,
                ev1020_19 as 客戶,
                ev1020_06 as 拆單人員,
                ev1020_09 as 重量,
                ev1020_05 as 門市代號
            FROM ev1020
            WHERE ev1020_13 LIKE '%生產%'
            """
            
            cursor.execute(query)
            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()
            
            df = pd.DataFrame.from_records(rows, columns=columns)
            df['出貨日期'] = pd.to_datetime(df['出貨日期'])
            
            # 獲取當前日期
            current_date = pd.Timestamp.now()
            
            # 過濾只保留當週以後的數據
            df = df[df['出貨日期'] >= current_date.normalize() - pd.Timedelta(days=current_date.dayofweek)]
            
            if df.empty:
                print("警告：沒有找到當週以後的數據")
                return
            
            # 計算每筆資料所屬的週起始日和週結束日
            df['週起始日'] = df['出貨日期'].dt.to_period('W').dt.start_time
            df['週結束日'] = df['出貨日期'].dt.to_period('W').dt.end_time
            df['日期區間'] = df['週起始日'].dt.strftime('%Y/%m/%d') + '-' + df['週結束日'].dt.strftime('%Y/%m/%d')
            
            # 依據廠別和日期區間分組計算總材數
            factory1_data = df[df['廠別'] == '001'].groupby(['日期區間', '週起始日'])['材數'].sum()
            factory2_data = df[df['廠別'] == '002'].groupby(['日期區間', '週起始日'])['材數'].sum()
            
            # 轉換為字典格式，使用日期區間作為鍵
            self.factory1_data = {idx[0]: weight for idx, weight in factory1_data.items()}
            self.factory2_data = {idx[0]: weight for idx, weight in factory2_data.items()}
            
            # 儲存所有日期區間的週起始日，用於後續排序
            self.date_ranges = {idx[0]: idx[1] for idx in factory1_data.index.union(factory2_data.index)}
            
            print("成功從資料庫載入數據！")
            cursor.close()
            
        except Exception as e:
            print(f"載入數據時出錯：{str(e)}")

    def load_estimated_orders(self):
        """（保留空函式，避免主程式報錯）"""
        print("本系統僅支援ACCDB預估訂單數據，請用功能7直接載入。")
        return False

    def load_estimated_orders_from_accdb(self):
        """從 ACCDB/MDB 資料庫讀取預估訂單數據（彰化查詢、台南查詢）"""
        try:
            if not os.path.exists(self.excel_path):
                print(f"未找到預估訂單ACCDB檔案：{self.excel_path}")
                return False
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={self.excel_path};'
            )
            conn = pyodbc.connect(conn_str)
            orders = []
            for table, factory_name in [('彰化查詢', '彰化廠'), ('台南查詢', '台南廠')]:
                query = f"SELECT * FROM [{table}]"
                df = pd.read_sql(query, conn)
                for _, row in df.iterrows():
                    orders.append({
                        '日期': row['預計出貨日'],
                        '門市': row['門市'],  # 保留原本內容
                        '門市代號': row['門市代號'] if '門市代號' in row else '',
                        '預估材數': row['預估材數'],
                        '備註': row['備註'] if '備註' in row else '',
                        '工廠': factory_name
                    })
            conn.close()
            self.estimated_orders = orders
            print(f"成功從預估訂單ACCDB載入{len(self.estimated_orders)}筆預估訂單數據")
            return True
        except Exception as e:
            print(f"從預估訂單ACCDB載入預估訂單數據時發生錯誤：{str(e)}")
            return False

    def set_ratio_settings(self):
        """設定比例"""
        try:
            print("\n=== 目前的比例設定 ===")
            print(f"材數比例 > {self.ratio_settings['upper']} 時，建議分配給台南廠")
            print(f"材數比例 < {self.ratio_settings['lower']} 時，建議分配給彰化廠")
            
            # 輸入新的上限值
            while True:
                try:
                    upper = float(input("\n請輸入新的上限值（建議分配給台南廠的比例）："))
                    if upper <= 0:
                        print("比例必須大於0")
                        continue
                    break
                except ValueError:
                    print("請輸入有效的數字")
            
            # 輸入新的下限值
            while True:
                try:
                    lower = float(input("請輸入新的下限值（建議分配給彰化廠的比例）："))
                    if lower <= 0:
                        print("比例必須大於0")
                        continue
                    if lower >= upper:
                        print("下限值必須小於上限值")
                        continue
                    break
                except ValueError:
                    print("請輸入有效的數字")
            
            self.ratio_settings['upper'] = upper
            self.ratio_settings['lower'] = lower
            self.save_ratio_settings()
            
            print("\n=== 新的比例設定 ===")
            print(f"材數比例 > {upper} 時，建議分配給台南廠")
            print(f"材數比例 < {lower} 時，建議分配給彰化廠")
            
            return True
            
        except Exception as e:
            print(f"設定比例時出錯：{str(e)}")
            return False

    def set_max_capacity(self):
        """設定彰化廠、台南廠每週最大材數，並持久化"""
        try:
            print("\n=== 設定每週最大材數 ===")
            c1 = input(f"請輸入彰化廠每週最大材數 (目前: {getattr(self, 'factory1_max_capacity', '未設定')}): ")
            c2 = input(f"請輸入台南廠每週最大材數 (目前: {getattr(self, 'factory2_max_capacity', '未設定')}): ")
            self.factory1_max_capacity = int(c1) if c1.strip() else getattr(self, 'factory1_max_capacity', 0)
            self.factory2_max_capacity = int(c2) if c2.strip() else getattr(self, 'factory2_max_capacity', 0)
            print(f"彰化廠最大產能: {self.factory1_max_capacity}, 台南廠最大產能: {self.factory2_max_capacity}")
            self.save_ratio_settings()  # 設定後自動保存
        except Exception as e:
            print(f"設定最大產能時發生錯誤：{str(e)}")

    def generate_report(self):
        """生成比較報告"""
        try:
            # 獲取日期排序
            date_ranges_with_dates = []
            for date_range in set(self.factory1_data.keys()) | set(self.factory2_data.keys()):
                start_date = datetime.strptime(date_range.split('-')[0], '%Y/%m/%d')
                date_ranges_with_dates.append((date_range, start_date))
            
            # 按日期降序排序
            all_date_ranges = [x[0] for x in sorted(date_ranges_with_dates, 
                                                  key=lambda x: x[1], 
                                                  reverse=False)]
            
            # 創建基礎DataFrame
            df = pd.DataFrame({
                '日期區間': all_date_ranges,
                '彰化廠材數': [self.factory1_data.get(dr, 0) for dr in all_date_ranges],
                '台南廠材數': [self.factory2_data.get(dr, 0) for dr in all_date_ranges]
            })
            
            # 計算預估材數
            def get_estimated_orders(start_date_str, end_date_str):
                if not self.estimated_orders:
                    return {'彰化': 0, '台南': 0}
                start_date = datetime.strptime(start_date_str, '%Y/%m/%d')
                end_date = datetime.strptime(end_date_str, '%Y/%m/%d')
                factory1_sum = 0  # 彰化廠
                factory2_sum = 0  # 台南廠
                for order in self.estimated_orders:
                    order_date = pd.to_datetime(order['日期']).to_pydatetime()
                    if start_date <= order_date <= end_date:
                        # 根據工廠欄位判斷
                        if '工廠' in order and order['工廠'] == '彰化廠':
                            factory1_sum += order['預估材數']
                        elif '工廠' in order and order['工廠'] == '台南廠':
                            factory2_sum += order['預估材數']
                return {'彰化': factory1_sum, '台南': factory2_sum}
            
            # 添加預估材數列
            df['彰化廠預估材數'] = df['日期區間'].apply(
                lambda x: get_estimated_orders(x.split('-')[0], x.split('-')[1])['彰化'])
            df['台南廠預估材數'] = df['日期區間'].apply(
                lambda x: get_estimated_orders(x.split('-')[0], x.split('-')[1])['台南'])
            
            # 計算合計材數
            df['彰化廠合計材數'] = df['彰化廠材數'] + df['彰化廠預估材數']
            df['台南廠合計材數'] = df['台南廠材數'] + df['台南廠預估材數']
            
            # 計算材數差異和比例（使用合計材數）
            df['合計材數差異'] = df['彰化廠合計材數'] - df['台南廠合計材數']
            df['合計材數比例'] = df['彰化廠合計材數'] / df['台南廠合計材數'].replace(0, float('nan'))
            
            # 根據合計材數判斷訂單分配建議
            def get_suggestion(row):
                if pd.isna(row['合計材數比例']):
                    return '無法計算'
                elif row['合計材數比例'] > self.ratio_settings['upper']:
                    return '建議分配給台南廠'
                elif row['合計材數比例'] < self.ratio_settings['lower']:
                    return '建議分配給彰化廠'
                else:
                    return '訂單分配正常'
            
            df['訂單分配建議'] = df.apply(get_suggestion, axis=1)

            # 新增建議分配量欄位（依照新公式）
            def get_suggested_amount(row):
                try:
                    c1 = float(row['彰化廠合計材數']) if isinstance(row['彰化廠合計材數'], (int, float)) else float(str(row['彰化廠合計材數']).replace(',', ''))
                    c2 = float(row['台南廠合計材數']) if isinstance(row['台南廠合計材數'], (int, float)) else float(str(row['台南廠合計材數']).replace(',', ''))
                    upper = self.ratio_settings['upper']
                    lower = self.ratio_settings['lower']
                    total = c1 + c2
                    if row['訂單分配建議'] == '建議分配給台南廠':
                        # 上限(2.2)：(A+B)/(1+2.2) - B
                        up = total / (1 + upper) - c2
                        # 下限(1.8)：(A+B)/(1+1.8) - B
                        low = total / (1 + lower) - c2
                        return f"建議分配到台南廠：{low:,.0f} ~ {up:,.0f} 材數"
                    elif row['訂單分配建議'] == '建議分配給彰化廠':
                        # 上限(2.2)：((A+B)/(1+2.2))*2.2 - A
                        up = (total / (1 + upper)) * upper - c1
                        # 下限(1.8)：((A+B)/(1+1.8))*1.8 - A
                        low = (total / (1 + lower)) * lower - c1
                        return f"建議分配到彰化廠：{low:,.0f} ~ {up:,.0f} 材數"
                    elif row['訂單分配建議'] == '訂單分配正常':
                        return '維持現有分配'
                    else:
                        return '-'
                except Exception:
                    return '-'
            df['建議分配量'] = df.apply(get_suggested_amount, axis=1)
            
            # 設置日期區間為索引
            df.set_index('日期區間', inplace=True)
            
            # 格式化數值列
            numeric_columns = ['彰化廠材數', '台南廠材數', 
                             '彰化廠預估材數', '台南廠預估材數',
                             '彰化廠合計材數', '台南廠合計材數',
                             '合計材數差異']
            
            for col in numeric_columns:
                df[col] = df[col].map('{:,.0f}'.format)
            
            df['合計材數比例'] = df['合計材數比例'].map(lambda x: '{:.2f}'.format(x) if pd.notnull(x) else '-')
            
            # 若日期區間為index，重設為欄位並移到最左側
            if '日期區間' not in df.columns and df.index.name == '日期區間':
                df = df.reset_index()
            # 強制欄位順序，日期區間在最左
            col_order = ['日期區間'] + [col for col in df.columns if col != '日期區間']
            df = df[col_order]
            
            return df
            
        except Exception as e:
            print(f"生成報告時發生錯誤：{str(e)}")
            return pd.DataFrame()

    def plot_comparison(self):
        """繪製比較圖表"""
        try:
            fig, (ax1, ax2, ax3, ax4) = plt.subplots(4, 1, figsize=(15, 24))
            
            # 設置中文字型
            plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
            plt.rcParams['axes.unicode_minus'] = False
            
            df = self.generate_report().copy()
            
            # 將字符串格式的數值轉回浮點數
            numeric_columns = ['彰化廠材數', '台南廠材數', 
                             '彰化廠預估材數', '台南廠預估材數',
                             '彰化廠合計材數', '台南廠合計材數']
            
            for col in numeric_columns:
                df[col] = df[col].str.replace(',', '').astype(float)
            
            # 第一個子圖：實際材數和預估材數
            ax1.plot(df['日期區間'], df['彰化廠材數'], 'b-', marker='o', label='彰化廠實際材數')
            ax1.plot(df['日期區間'], df['台南廠材數'], 'r-', marker='o', label='台南廠實際材數')
            ax1.plot(df['日期區間'], df['彰化廠預估材數'], 'b--', marker='^', label='彰化廠預估材數')
            ax1.plot(df['日期區間'], df['台南廠預估材數'], 'r--', marker='^', label='台南廠預估材數')
            
            # 在plot_comparison兩個子圖中畫出最大產能橫線
            if hasattr(self, 'factory1_max_capacity') and self.factory1_max_capacity:
                ax1.axhline(y=self.factory1_max_capacity, color='b', linestyle=':', label='彰化廠最大產能')
                ax2.axhline(y=self.factory1_max_capacity, color='b', linestyle=':', label='彰化廠最大產能')
            if hasattr(self, 'factory2_max_capacity') and self.factory2_max_capacity:
                ax1.axhline(y=self.factory2_max_capacity, color='r', linestyle=':', label='台南廠最大產能')
                ax2.axhline(y=self.factory2_max_capacity, color='r', linestyle=':', label='台南廠最大產能')
            
            # 添加數據標籤
            for x, y1, y2, y3, y4 in zip(df['日期區間'], 
                                        df['彰化廠材數'], 
                                        df['台南廠材數'],
                                        df['彰化廠預估材數'],
                                        df['台南廠預估材數']):
                ax1.annotate(f'{int(y1):,}', (x, y1), textcoords="offset points", 
                            xytext=(0,10), ha='center', fontsize=8)
                ax1.annotate(f'{int(y2):,}', (x, y2), textcoords="offset points", 
                            xytext=(0,-15), ha='center', fontsize=8)
                ax1.annotate(f'{int(y3):,}', (x, y3), textcoords="offset points", 
                            xytext=(0,25), ha='center', fontsize=8)
                ax1.annotate(f'{int(y4):,}', (x, y4), textcoords="offset points", 
                            xytext=(0,-30), ha='center', fontsize=8)
            
            ax1.set_title('實際材數與預估材數比較', fontsize=14, pad=20)
            ax1.set_xlabel('日期區間', fontsize=12)
            ax1.set_ylabel('材數', fontsize=12)
            ax1.legend(fontsize=10)
            ax1.grid(True)
            
            # 不要再呼叫set_xticks/set_xticklabels，直接用tick_params旋轉
            ax1.tick_params(axis='x', rotation=45, labelsize=10)
            
            # 第二個子圖：合計材數比較
            ax2.plot(df['日期區間'], df['彰化廠合計材數'], 'b-', marker='o', label='彰化廠合計材數')
            ax2.plot(df['日期區間'], df['台南廠合計材數'], 'r-', marker='o', label='台南廠合計材數')
            
            # 添加合計材數標籤
            for x, y1, y2 in zip(df['日期區間'], df['彰化廠合計材數'], df['台南廠合計材數']):
                ax2.annotate(f'{int(y1):,}', (x, y1), textcoords="offset points", 
                            xytext=(0,10), ha='center', fontsize=8)
                ax2.annotate(f'{int(y2):,}', (x, y2), textcoords="offset points", 
                            xytext=(0,-15), ha='center', fontsize=8)
            
            # 添加比例參考線
            for ratio in [self.ratio_settings['lower'], self.ratio_settings['upper']]:
                ax2.plot(df['日期區間'], df['台南廠合計材數'] * ratio, '--', 
                         alpha=0.5, label=f'理想比例 {ratio}')
            
            ax2.set_title('合計材數比較', fontsize=14, pad=20)
            ax2.set_xlabel('日期區間', fontsize=12)
            ax2.set_ylabel('材數', fontsize=12)
            ax2.legend(fontsize=10)
            ax2.grid(True)
            
            # 不要再呼叫set_xticks/set_xticklabels，直接用tick_params旋轉
            ax2.tick_params(axis='x', rotation=45, labelsize=10)
            
            # ====== ax3：合計門市類別材數（週） ======
            # 維持原本週區間邏輯
            all_date_ranges = df['日期區間'].tolist()
            main_data = getattr(self, 'main_data_df', None)
            if main_data is None:
                cursor = self.conn.cursor()
                query = """
                SELECT 
                    ev1020_03 as 出貨日期,
                    ev1020_88 as 廠別,
                    ev1020_07 as 材數,
                    ev1020_13 as 生產性質,
                    ev1020_20 as 門市,
                    ev1020_11 as 圖號,
                    ev1020_12 as 色號,
                    ev1020_19 as 客戶,
                    ev1020_06 as 拆單人員,
                    ev1020_09 as 重量,
                    ev1020_05 as 門市代號
                FROM ev1020
                WHERE ev1020_13 LIKE '%生產%'
                """
                cursor.execute(query)
                columns = [column[0] for column in cursor.description]
                rows = cursor.fetchall()
                main_data = pd.DataFrame.from_records(rows, columns=columns)
                main_data['出貨日期'] = pd.to_datetime(main_data['出貨日期'])
                self.main_data_df = main_data
            else:
                main_data = self.main_data_df
            main_data['週起始日'] = main_data['出貨日期'].dt.to_period('W').dt.start_time
            main_data['週結束日'] = main_data['出貨日期'].dt.to_period('W').dt.end_time
            main_data['日期區間'] = main_data['週起始日'].dt.strftime('%Y/%m/%d') + '-' + main_data['週結束日'].dt.strftime('%Y/%m/%d')
            def classify_store(code):
                if isinstance(code, str):
                    if code.startswith('S'):
                        return '專案'
                    elif code.startswith('P'):
                        return '代工'
                return '零售'
            main_data['門市類別'] = main_data['門市代號'].apply(classify_store)
            main_data['廠別名稱'] = main_data['廠別'].map({'001': '彰化', '002': '台南'})
            est_orders = pd.DataFrame(self.estimated_orders) if self.estimated_orders else pd.DataFrame()
            if not est_orders.empty:
                est_orders['日期'] = pd.to_datetime(est_orders['日期'])
                est_orders['週起始日'] = est_orders['日期'].dt.to_period('W').dt.start_time
                est_orders['週結束日'] = est_orders['日期'].dt.to_period('W').dt.end_time
                est_orders['日期區間'] = est_orders['週起始日'].dt.strftime('%Y/%m/%d') + '-' + est_orders['週結束日'].dt.strftime('%Y/%m/%d')
                est_orders['門市類別'] = est_orders['門市代號'].apply(classify_store)
                est_orders['廠別名稱'] = est_orders['工廠'].map({'彰化廠': '彰化', '台南廠': '台南'})
                est_orders['材數'] = est_orders['預估材數']
            all_data = pd.concat([main_data[['日期區間','廠別名稱','門市類別','材數']],
                                  est_orders[['日期區間','廠別名稱','門市類別','材數']] if not est_orders.empty else None],
                                 ignore_index=True)
            all_data = all_data[all_data['日期區間'].isin(all_date_ranges)]
            pivot = all_data.groupby(['日期區間','廠別名稱','門市類別'])['材數'].sum().reset_index()
            group_keys = [
                ('彰化','零售'), ('彰化','專案'), ('彰化','代工'),
                ('台南','零售'), ('台南','專案'), ('台南','代工')
            ]
            colors = {
                ('彰化','零售'): '#008000', # 綠
                ('彰化','專案'): '#0000FF', # 藍
                ('彰化','代工'): '#800080', # 紫
                ('台南','零售'): '#FF0000', # 紅
                ('台南','專案'): '#FFA500', # 橘
                ('台南','代工'): '#FFFF00', # 黃
            }
            bar_width = 0.12
            x = np.arange(len(all_date_ranges))
            offset = np.linspace(-0.3, 0.3, 6)
            for idx, key in enumerate(group_keys):
                y = []
                for dr in all_date_ranges:
                    val = pivot[(pivot['日期區間']==dr)&(pivot['廠別名稱']==key[0])&(pivot['門市類別']==key[1])]['材數']
                    y.append(val.values[0] if not val.empty else 0)
                y = np.array(y)
                mask = y > 0
                if mask.any():
                    ax3.bar(x[mask]+offset[idx], y[mask], width=bar_width, color=colors[key], label=f"{key[0]}{key[1]}")
                    for xi, yi in zip(x[mask]+offset[idx], y[mask]):
                        ax3.text(xi, yi, f'{int(yi):,}', ha='center', va='bottom', fontsize=8)
            ax3.set_xticks(x)
            ax3.set_xticklabels(all_date_ranges, rotation=45, ha='right', fontsize=10)
            ax3.set_ylabel('材數', fontsize=12)
            ax3.set_title('合計門市類別材數', fontsize=14, pad=20)
            handles, labels = ax3.get_legend_handles_labels()
            by_label = dict(zip(labels, handles))
            ax3.legend(by_label.values(), by_label.keys(), fontsize=10)
            ax3.grid(True, axis='y', linestyle='--', alpha=0.5)

            # ====== ax4：合計門市類別材數（月，固定三個月） ======
            main_data['月份'] = main_data['出貨日期'].dt.strftime('%Y/%m')
            if not est_orders.empty:
                est_orders['月份'] = est_orders['日期'].dt.strftime('%Y/%m')
            all_data_m = pd.concat([
                main_data[['月份','廠別名稱','門市類別','材數']],
                est_orders[['月份','廠別名稱','門市類別','材數']] if not est_orders.empty else None
            ], ignore_index=True)
            today = pd.Timestamp.now().replace(day=1)
            month_list = [(today + pd.DateOffset(months=i)).strftime('%Y/%m') for i in range(3)]
            pivot_m = all_data_m.groupby(['月份','廠別名稱','門市類別'])['材數'].sum().reset_index()
            x4 = np.arange(len(month_list))
            for idx, key in enumerate(group_keys):
                y = []
                for m in month_list:
                    val = pivot_m[(pivot_m['月份']==m)&(pivot_m['廠別名稱']==key[0])&(pivot_m['門市類別']==key[1])]['材數']
                    y.append(val.values[0] if not val.empty else 0)
                ax4.bar(x4+offset[idx], y, width=bar_width, color=colors[key], label=f"{key[0]}{key[1]}")
                for xi, yi in zip(x4+offset[idx], y):
                    if yi > 0:
                        ax4.text(xi, yi, f'{int(yi):,}', ha='center', va='bottom', fontsize=8)
            ax4.set_xticks(x4)
            ax4.set_xticklabels(month_list, rotation=45, ha='right', fontsize=10)
            ax4.set_ylabel('材數', fontsize=12)
            ax4.set_title('合計門市類別材數（月）', fontsize=14, pad=20)
            handles4, labels4 = ax4.get_legend_handles_labels()
            by_label4 = dict(zip(labels4, handles4))
            ax4.legend(by_label4.values(), by_label4.keys(), fontsize=10)
            ax4.grid(True, axis='y', linestyle='--', alpha=0.5)

            # 調整整體布局
            plt.tight_layout()
            plt.savefig(f'Factory_Comparison_{datetime.now().strftime("%Y%m%d")}.png', dpi=300, bbox_inches='tight')
            plt.close()
            print(f"圖表已保存為: Factory_Comparison_{datetime.now().strftime('%Y%m%d')}.png")
        except Exception as e:
            print(f"生成圖表時發生錯誤：{str(e)}")

    def export_to_excel(self):
        """將比較報告匯出成 Excel 檔案（數字格式+合計列）"""
        try:
            if not self.factory1_data or not self.factory2_data:
                print("沒有資料可供匯出，請先載入數據（選項1）")
                return False
                
            # 取得報表資料
            df = self.generate_report()

            # 將數字欄位轉回數字格式
            numeric_columns = ['彰化廠材數', '台南廠材數', '彰化廠預估材數', '台南廠預估材數', '彰化廠合計材數', '台南廠合計材數', '合計材數差異']
            for col in numeric_columns:
                df[col] = df[col].astype(str).str.replace(',', '').replace('-', '0').astype(float)

            # 合計列資料
            total_row = {}
            total_row['日期區間'] = '合計'
            for col in numeric_columns:
                total_row[col] = df[col].sum()
            # 比例欄位
            try:
                total_ratio = total_row['彰化廠合計材數'] / total_row['台南廠合計材數'] if total_row['台南廠合計材數'] else 0
                total_row['合計材數比例'] = f"{total_ratio:.2f}"
            except Exception:
                total_row['合計材數比例'] = '-'
            # 其他欄位
            for col in df.columns:
                if col not in total_row:
                    total_row[col] = '' if col not in ['訂單分配建議', '建議分配量'] else '-'
            # 將合計列加到DataFrame
            df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

            # 生成檔案名稱（包含日期）
            filename = f'Factory_Comparison_{datetime.now().strftime("%Y%m%d")}.xlsx'

            # 寫入Excel
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='比較報告', index=False)
                worksheet = writer.sheets['比較報告']
                # 自動調整欄寬
                for idx, col in enumerate(df.columns):
                    max_length = max(df[col].astype(str).apply(len).max(), len(col))
                    worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2
            print(f"報表已匯出為: {filename}")
            return True
        except Exception as e:
            print(f"匯出 Excel 時發生錯誤：{str(e)}")
            return False

def main():
    try:
        print("=== 工廠材數比較系統啟動 ===")
        comparison = FactoryComparison()
        print("正在連接資料庫...")
        
        if not comparison.connect_to_database():
            print("無法連接到預設資料庫，請選擇資料庫檔案位置。")
            if not comparison.select_database() or not comparison.connect_to_database():
                print("無法連接到資料庫，程式將結束。")
                input("\n按 Enter 鍵結束程式...")
                return
        
        while True:
            print("\n=== 工廠材數比較系統 ===")
            print("1. 載入資料庫數據（含預估訂單數據）")
            print("2. 查看比較報告")
            print("3. 生成比較圖表")
            print("4. 匯出報表至 Excel")
            print("5. 查看預估訂單數據")
            print("6. 設定比例範圍")
            print("7. 設定每週最大材數")
            print("8. 更改預估訂單數據路徑")
            print("9. 更改資料庫位置")
            print("10. 退出")
            
            choice = input("請選擇操作 (1-10): ")
            
            if choice == '1':
                comparison.load_data_from_database()
                comparison.load_estimated_orders_from_accdb()
            elif choice == '2':
                if not comparison.factory1_data or not comparison.factory2_data:
                    print("請先載入數據（選項1）")
                    continue
                report = comparison.generate_report()
                print("\n=== 比較報告 ===")
                if not report.empty:
                    # 日期區間只顯示mm/dd-mm/dd
                    def short_date_range(date_range):
                        try:
                            start, end = date_range.split('-')
                            start = start.strip()
                            end = end.strip()
                            start_md = '/'.join(start.split('/')[1:])
                            end_md = '/'.join(end.split('/')[1:])
                            return f"{start_md}-{end_md}"
                        except Exception:
                            return date_range
                    report = report.copy()
                    report['日期區間'] = report['日期區間'].apply(short_date_range)
                    # 欄寬根據最大內容自動決定，所有欄名與資料都靠左，欄與欄之間4個空格
                    col_widths = [max(len(str(x)) for x in report[col].astype(str)) for col in report.columns]
                    col_widths = [max(w, len(col)) for w, col in zip(col_widths, report.columns)]
                    # 印欄位名稱（靠左）
                    header = (' ' * 4).join([str(col).ljust(width) for col, width in zip(report.columns, col_widths)])
                    print(header)
                    # 印每一列（靠左）
                    for _, row in report.iterrows():
                        line = (' ' * 4).join([str(row[col]).ljust(width) for col, width in zip(report.columns, col_widths)])
                        print(line)
                else:
                    print(report)
            elif choice == '3':
                if not comparison.factory1_data or not comparison.factory2_data:
                    print("請先載入數據（選項1）")
                    continue
                comparison.plot_comparison()
            elif choice == '4':
                comparison.export_to_excel()
            elif choice == '5':
                if comparison.estimated_orders:
                    print("\n=== 預估訂單數據 ===")
                    df = pd.DataFrame(comparison.estimated_orders)
                    # 日期轉型並排序
                    df['日期'] = pd.to_datetime(df['日期'])
                    df = df.sort_values(by=['工廠', '日期'], ascending=[True, True])
                    for _, order in df.iterrows():
                        print(f"日期: {order['日期'].strftime('%Y-%m-%d')}, 工廠: {order['工廠']}, 門市代號: {order.get('門市代號','')}, 門市: {order['門市']}, 預估材數: {order['預估材數']}, 備註: {order['備註']}")
                else:
                    print("尚未載入預估訂單數據，請先執行選項1")
            elif choice == '6':
                comparison.set_ratio_settings()
            elif choice == '7':
                comparison.set_max_capacity()
            elif choice == '8':
                if comparison.select_excel_file():
                    print("成功更改預估訂單數據路徑！")
                else:
                    print("取消更改預估訂單數據路徑。")
            elif choice == '9':
                if comparison.select_database():
                    if comparison.connect_to_database():
                        print("成功更改資料庫位置！")
                    else:
                        print("連接新資料庫失敗！")
                else:
                    print("取消更改資料庫位置。")
            elif choice == '10':
                if comparison.conn:
                    comparison.conn.close()
                print("感謝使用！")
                break
            else:
                print("無效的選擇，請重試。")
    
    except Exception as e:
        print(f"程式啟動時發生錯誤：{str(e)}")
    
    finally:
        input("\n按 Enter 鍵結束程式...")

if __name__ == "__main__":
    main()   
