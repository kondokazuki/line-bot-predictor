from flask import Flask, request, abort
from linebot.v3.webhook import WebhookHandler
from linebot.v3.messaging import MessagingApiClient, TextSendMessage, TextMessage, MessageEvent
from linebot.v3.exceptions import InvalidSignatureError
import pandas as pd
import os



# あなたのチャネル情報をここにセット
CHANNEL_ACCESS_TOKEN = '0lqYCmSUQjUGdpwk77aNZ8cEXe75Rlz509cftBA2F1EaJDSLXLLBBF9W4unatBKQJlPIDm02YOWaxpZaFU1qOolz99MTzRzrtT2p1PDEr+E/jYM5tMYpox5i/pbxTvwhcdsgDiQUq55+aJwpp0EkTwdB04t89/1O/w1cDnyilFU='
CHANNEL_SECRET = '13f4af9e18a2f1bf4606703527353227'

line_bot_api = MessagingApiClient(CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

# Flaskアプリ
app = Flask(__name__)

# カードデータ読み込み
arrays = {}
card_info = None

def load_card_data(file_path):
    global arrays, card_info
    df = pd.read_excel(file_path)

    # 0列目～11列目まで読み込む（つまりB列～M列）
    arrays.clear()
    for i in range(1, 13):  # ←ここポイント！（1列目から12列目まで）
        column_data = df.iloc[:, i].dropna().tolist()
        arrays[i-1] = column_data  # 0スタートで管理

    # 14〜16列目（O列～Q列）をカード情報として保存
    card_info = df.iloc[:, [14, 15, 16]]

load_card_data('cards.xlsx')

def find_current_position(arrays, recent_cards):
    for array_index, card_list in arrays.items():
        for idx in range(len(card_list) - len(recent_cards) + 1):
            if card_list[idx:idx+len(recent_cards)] == recent_cards:
                return array_index, idx + len(recent_cards)
    return None, None

def predict_up_to_end(arrays, card_info, array_index, start_idx):
    predictions = []
    array = arrays[array_index]
    for offset, idx in enumerate(range(start_idx, len(array))):
        card_no = array[idx]
        card_row = card_info[card_info.iloc[:,0] == card_no].iloc[0]
        predictions.append({
            'cards_later': offset,
            'rarity': card_row.iloc[1],
            'name': card_row.iloc[2]
        })
    return predictions

def format_predictions(predictions):
    highlight_rarities = ['U', 'P', 'SEC']
    result_lines = []
    for pred in predictions:
        if pred['rarity'] in highlight_rarities:
            line = f"✅ {pred['cards_later']}枚後: {pred['rarity']} - {pred['name']}"
        else:
            line = f"{pred['cards_later']}枚後: {pred['rarity']} - {pred['name']}"
        result_lines.append(line)
    return "\n".join(result_lines)

def predict_from_input(input_text):
    try:
        # 「、」「,」「.」すべて区切り対象にする
        import re
        split_text = re.split('[、,.]', input_text)
        recent_cards = [int(x.strip()) for x in split_text if x.strip()]
    except ValueError:
        return "入力形式が正しくありません。例: '10、18、40' または '10.18.40' のように入力してください。"

    matches = find_current_positions(arrays, recent_cards)
    if matches:
        outputs = []
        for array_index, start_idx in matches:
            predictions = predict_up_to_end(arrays, card_info, array_index, start_idx)
            outputs.append(f"配列{array_index + 1}番の予測:\n" + format_predictions(predictions))
        return "\n\n".join(outputs)
    else:
        return "一致する配列が見つかりませんでした。"

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_input = event.message.text
    result = predict_from_input(user_input)
    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text=result)
    )

import os

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port)

