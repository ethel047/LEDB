from excel_w_r import read_excel_to_dict
import schedule
import time
from linebot.models import TextSendMessage
from id import id_mapping
from datetime import datetime
from sheets_handler import check_today_entry


today = datetime.now().strftime('%Y/%m/%d')   
def record_bell(line_bot_api):
    try:
        file_leave=read_excel_to_dict(f"/mnt/data/All_請假單.xlsx")
        # leave_date=[datetime.strptime(date, '%m/%d') for date in file_leave[user_name]['請假日期']] #已經請的
        file_overtime=read_excel_to_dict(f"/mnt/data/All_加班名單.xlsx")
        # overtime_date=[datetime.strptime(date, '%m/%d') for date in file_leave[user_name]['請假日期']] #已經請的

        # print(overtime_person.get('姓名', []))
        for user_name in id_mapping:
            # if leave_person or overtime_person:
            overtime_date=[datetime.strptime(date, '%m/%d') for date in file_leave[user_name]['請假日期']] #已經請的
            leave_date=[datetime.strptime(date, '%m/%d') for date in file_overtime[user_name]['加班日期']] #已經請的
            if today not in leave_date:
                if today not in overtime_date :
                    print('check start')
                    user_id = id_mapping.get(user_name, "該名字沒有對應的用戶ID")
                    if check_today_entry(user_id) == False:
                        print(user_name)
                        line_bot_api.push_message('C169b23c827c28e4c5d3c7ddbfb5aa6b9', TextSendMessage(text=f'{user_name}，尚未填紀錄')) #群組id
                        print("自動訊息發送成功")
            else:
                    line_bot_api.push_message('C169b23c827c28e4c5d3c7ddbfb5aa6b9', TextSendMessage(text=f'大家都填完了很棒喔')) #群組id
          
    except Exception as e:
        print(f"發送訊息時出錯: {e}")


