import os
from datetime import datetime, timedelta
import pandas as pd

def save_leave_requests_to_excel(leave_requests, leave_date):
    try:
        save_directory = os.path.expanduser("/mnt/data")
        if not os.path.exists(save_directory):
            os.makedirs(save_directory)
        current_year = datetime.now().year
        # 格式化请假日期

        if '～' in leave_date:
# 解析 start date 和 end date
            start_date_str, end_date_str = leave_date.split('～')
            start_date = datetime.strptime(start_date_str.strip(), "%m/%d")
            end_date = datetime.strptime(end_date_str.strip(), "%m/%d")
            # 生成日期列表
            current_date = start_date
            while current_date <= end_date:
                print(current_date.strftime("%m-%d"))  # 打印 MM/DD 格式的日期
                formatted_leave_date = datetime.strptime(f"{current_year}-{current_date}", "%Y-%m-%d").strftime("%Y-%m-%d")
                excel_filename = os.path.join(save_directory, f'{formatted_leave_date}_請假單.xlsx')
                excel_filename2 = os.path.join(save_directory, f'All請假單.xlsx')
                # 将 formatted_leave_date 更新到 leave_requests 字典中
                for user in leave_requests:
                    leave_requests[user]['請假日期'] = formatted_leave_date

                # 如果文件存在，讀取現有數據
                if os.path.exists(excel_filename): #(seperate)
                    df_existing = pd.read_excel(excel_filename, index_col=0)
                    df_new = pd.DataFrame.from_dict(leave_requests, orient='index')
                    df_combined = pd.concat([df_existing, df_new])
                else:
                    df_combined = pd.DataFrame.from_dict(leave_requests, orient='index')
                # 將合併後的數據寫入
                df_combined.to_excel(excel_filename, index_label='姓名')
                print(f"已保存請假信息到 {excel_filename}")

                if os.path.exists(excel_filename2): #(ALL)
                    df_existing = pd.read_excel(excel_filename2, index_col=0)
                    df_new = pd.DataFrame.from_dict(leave_requests, orient='index')
                    df_combined = pd.concat([df_existing, df_new])
                else:
                    df_combined = pd.DataFrame.from_dict(leave_requests, orient='index')
                df_combined.to_excel(excel_filename2, index_label='姓名')
                print(f"已保存請假信息到 {excel_filename2}")
                current_date += timedelta(days=1)
        # else:
        try:
            formatted_leave_date = datetime.strptime(f"{current_year}-{leave_date}", "%Y-%m-%d").strftime("%Y-%m-%d")
        except ValueError:
            formatted_leave_date = datetime.strptime(leave_date, "%Y-%m-%d").strftime("%Y-%m-%d")

            excel_filename = os.path.join(save_directory, f'{formatted_leave_date}_請假單.xlsx')
            excel_filename2 = os.path.join(save_directory, f'All請假單.xlsx')
            # 将 formatted_leave_date 更新到 leave_requests 字典中
            for user in leave_requests:
                leave_requests[user]['請假日期'] = formatted_leave_date

            # 如果文件存在，讀取現有數據
            if os.path.exists(excel_filename): #(seperate)
                df_existing = pd.read_excel(excel_filename, index_col=0)
                df_new = pd.DataFrame.from_dict(leave_requests, orient='index')
                df_combined = pd.concat([df_existing, df_new])
            else:
                df_combined = pd.DataFrame.from_dict(leave_requests, orient='index')
            # 將合併後的數據寫入
            df_combined.to_excel(excel_filename, index_label='姓名')
            print(f"已保存請假信息到 {excel_filename}")

            if os.path.exists(excel_filename2): #(ALL)
                df_existing = pd.read_excel(excel_filename2, index_col=0)
                df_new = pd.DataFrame.from_dict(leave_requests, orient='index')
                df_combined = pd.concat([df_existing, df_new])
            else:
                df_combined = pd.DataFrame.from_dict(leave_requests, orient='index')
            df_combined.to_excel(excel_filename2, index_label='姓名')
            print(f"已保存請假信息到 {excel_filename2}")

    except Exception as e:
        print(f"保存 Excel 文件時出错: {e}")

def read_excel_to_dict(file_path):
    try:
        
        df = pd.read_excel(file_path, index_col=0)
        
        grouped = df.groupby(df.index)
        result = {}
        
        for name, group in grouped:
            result[name] = {
                '請假日期': group['請假日期'].tolist(),  # 將所有的請假日期放入列表中
                '請假理由': group['請假理由'].tolist()   # 將所有的請假理由放入列表中
            }
        
        return result
    except Exception as e:
        print(f"讀取 Excel 文件時出錯: {e}")
        return {}
    
def save_overtime_requests_to_excel(user_id, user_name, today):
    try:
        
        save_directory = os.path.expanduser("/mnt/data")
        if not os.path.exists(save_directory):
            os.makedirs(save_directory)
        excel_filename = os.path.join(save_directory, f'{today}_加班名單.xlsx')
        df_new = pd.DataFrame({'姓名': [user_name]})
        if os.path.exists(excel_filename):
            df_existing = pd.read_excel(excel_filename)
            if '姓名' in df_existing.columns:
                df_existing = df_existing[['姓名']]  # 保留「姓名」列
                # 刪除第一行並合併新數據
                df_existing = df_existing.iloc[1:]  # 刪除第一行
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            else:
                print("現有檔案中不包含 '姓名' 列")
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)  # 處理沒有「姓名」列的情況
        else:
            df_combined = df_new
        # 將合併後的數據寫入檔案
        df_combined.to_excel(excel_filename, index=False)  # 由於「姓名」作為列名，索引標籤不需要
        print(f"已儲存加班信息到 {excel_filename}")
        
    except Exception as e:
        print(f"儲存 Excel 檔案時出錯: {e}")