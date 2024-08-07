from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage as TextMessageContent, TextSendMessage as TextMessage

from api_handler import call_prediction_api
from sheets_handler import check_today_entry, update_google_sheet
from leave_overtime import leave_talking, work_overtime
from record_bell import record_bell

import threading
from datetime import datetime
import time
import os
import schedule
import logging
# 設定日誌記錄
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
ACCESS_TOKEN = os.getenv('ACCESS_TOKEN')
WEBHOOK_SECRET = os.getenv('WEBHOOK_SECRET')
line_bot_api = LineBotApi(ACCESS_TOKEN)
handler = WebhookHandler(WEBHOOK_SECRET)

user_state = {}
additional = False
msg = 0

leave_requests={}

id_mapping = {
    "筮修": 'U02c370807baaf7ae3f6064d7705a8638',
    "姵蓁": 'U1afd46e95a1eac5a28fbf9fb889a8d5e',
    "政憲": 'U33bccf9dadb498e7d31b5a4cea9ae297',
    "致嘉": 'U864716975bfc7e1c5b93975470810bcc',
    "玟君": 'U0de51624c010c4a2ec439e10fbd67b1f',
    "蕎安": 'U5234be59bcac751016f6d0248fd3cd1b',
    "欣妍": 'U771ef1814950547f3dc72681f305c6e5',
    "又瑋": 'U42d032c70f43518578b14d5d25df9d1f'
}

# 關鍵字列表
KEYWORDS = {
    "左邊那台的帳號密碼": " ",
    "虛擬人那台的帳號密碼": " ",
    "右邊那台的帳號密碼": " ",
    "我要遠端連線!": " ",
    "我要常用連結!": " ",
    "我要上傳工作紀錄!": " ",
    "我想要問問題!": " ",
    "我要SeaDeep的密碼!":" ",
    "我要NAS的密碼!":" "
    # "我要請假!":" "
}
@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature')
    body = request.get_data(as_text=True)
    logger.info("Request body: %s", body)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        print("Invalid signature. Please check your channel access token/channel secret.")
        abort(400)

    return 'OK'

# 啟動催紀錄排程線程
def schedule_jobs():
    schedule.every().day.at("21:30").do(lambda:record_bell(line_bot_api))
    logger.info('Scheduler started')
    while True:
        schedule.run_pending()
        time.sleep(60)  # 每分鐘檢查一次排程
schedule_thread = threading.Thread(target=schedule_jobs)
schedule_thread.daemon = True
schedule_thread.start()

state = 0
@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    global additional, msg, user_state
    global leave_requests

    group_id = event.source.group_id if event.source.type == 'group' else None

    if group_id == "C169b23c827c28e4c5d3c7ddbfb5aa6b9":
        return

    user_id = event.source.user_id
    text = event.message.text
    user_name = "未知用户"  
    answer = None
    # 初始化 response_message
    response_message = "未能識別的指令，請重新輸入。"

    if text in KEYWORDS:
        # 如果訊息在關鍵字列表中，不做回應
        return
    # 檢查訊息中的名字

    name_in_text = next((name for name in id_mapping if name == text), None)
    # print("u_s",user_state.get(user_id))
    # 如果名字存在且ID不匹配，回應"請不要輸入別人的名稱"
    if name_in_text and id_mapping[name_in_text] != user_id:
        response_message = "請不要輸入別人的名稱"
    # 如果名字存在或資料為空"
    elif name_in_text or (user_state.get(user_id) is not None):
        if user_state.get(user_id) is None:
            if check_today_entry(user_id):
                response_message = "今天的工作已經填寫過了! 修改請至 https://docs.google.com/spreadsheets/d/1wFM0u_7YxVSURxSEq0BcIc9cNVYgRCQJhq619u93VBk/edit?gid=1003752271#gid=1003752271。"
            else:
                response_message = update_google_sheet(line_bot_api,user_id, user_state, text)
        else:
            response_message = update_google_sheet(line_bot_api,user_id, user_state, text)
    
    elif user_id in id_mapping.values():
        user_name = next(key for key, value in id_mapping.items() if value == user_id)
        response_message=work_overtime(user_id,text,user_name)
        # if response_message is None:
        #     response_message=leave_talking(leave_requests,line_bot_api, user_id, text, user_name)
        if response_message is None:
            response_message = call_prediction_api(text)
    

    reply_message = TextMessage(text=response_message)
    line_bot_api.reply_message(event.reply_token, reply_message)

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
