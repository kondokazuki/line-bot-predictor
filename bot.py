from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
import pandas as pd
import os
import re

# LINEチャネル情報
CHANNEL_ACCESS_TOKEN = '0lqYCmSUQjUGdpwk77aNZ8cEXe75Rlz509cftBA2F1EaJDSLXLLBBF9W4unatBKQJlPIDm02YOWaxpZaFU1qOolz99MTzRzrtT2p1PDEr+E/jYM5tMYpox5i/pbxTvwhcdsgDiQUq55+aJwpp0EkTwdB04t89/1O/w1cDnyilFU='
CHANNEL_SECRET = '13f4af9e18a2f1bf4606703527353227'

line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

app = Flask(__name__)

# 配列データ保存用
arrays_normal = {}
arrays_m = {}
card_info = None

# データ読み込み
def load_card_data(file_path):
    global arrays_normal, arrays_m, card_info
    df_normal = pd.read_excel(file_path, sheet_name=0)
    df_m = pd.read_excel(file_path, sheet_name='Mシリ')

    arrays_normal.clear()
    arrays_m.clear()
    for i in range(1, 13):
        arrays_normal[i-1] = df_normal.iloc[:, i].dropna().tolist()
        arrays_m[i-1] = df_m.iloc[:, i].dropna().tolist()

    card_info = df_normal.iloc[:, [14, 15, 16]]  # O列, P列, Q列

load_card_data('cards.xlsx')

# マッチ判定（通常）
def match_normal(target, value):
    if isinstance(target, int):
        return str(target) == str(value)
    return str(target) == str(value)

# マッチ判定（Mシリ：複数値対応）
def match_m(target, value):
    if isinstance(target, str) and '/' in target:
        options = target.split('/')
        return any(str(opt) == str(value) for opt in options)
    return str(target) == str(value)

# 特別入力（★やSP）判定
def special_match(card_no, special_keyword):
    row = card_info[card_info.iloc[:,0] == int(card_no)]
    if row.empty:
        return False
    o_value = str(row.iloc[0, 0])
    if special_keyword.endswith('*'):
        return f"★{special_keyword[:-1]}" in o_value
    elif special_keyword == 'SP':
        return 'SP' in o_value
    return False

# 現在位置探索
def find_current_positions(arrays, recent_cards, is_m=False):
    matches = []
    for array_index, card_list in arrays.items():
        for idx in range(len(card_list) - len(recent_cards) + 1):
            ok = True
            for offset, rc in enumerate(recent_cards):
                target = card_list[idx + offset]
                if rc.endswith('*') or rc == 'SP':
                    if not special_match(target, rc):
                        ok = False
                        break
                else:
                    if is_m:
                        if not match_m(target, rc):
                            ok = False
                            break
                    else:
                        if not match_normal(target, rc):
                            ok = False
                            break
            if ok:
                matches.append((array_index, idx + len(recent_cards)))
    return matches

# 未来予測
def predict_up_to_end(arrays, card_info, array_index, start_idx):
    predictions = []
    array = arrays[array_index]
    for offset, idx in enumerate(range(start_idx, len(array))):
        card_no = array[idx]
        card_row = card_info[card_info.iloc[:,0] == int(card_no)].iloc[0]
        predictions.append({
            'cards_later': offset + 1,
            'rarity': card_row.iloc[1],
            'name': card_row.iloc[2]
        })
    return predictions

# 出力整形
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

# 入力テキストから予測実行
def predict_from_input(input_text):
    if input_text.startswith('M'):
        arrays = arrays_m
        is_m = True
        input_text = input_text[1:]  # 先頭Mを除去
    elif input_text.startswith('通常'):
        arrays = arrays_normal
        is_m = False
        input_text = input_text[2:]  # 先頭通常を除去
    else:
        return "最初に'M'または'通常'をつけてください。"

    try:
        split_text = re.split('[、,.]', input_text)
        recent_cards = [x.strip() for x in split_text if x.strip()]
    except ValueError:
        return "入力形式が正しくありません。例: '通常10,18,40' または 'M10,18,40' のように入力してください。"

    matches = find_current_positions(arrays, recent_cards, is_m=is_m)
    if matches:
        outputs = []
        for array_index, start_idx in matches:
            predictions = predict_up_to_end(arrays, card_info, array_index, start_idx)
            outputs.append(f"配列{array_index + 1}番の予測:\n" + format_predictions(predictions))
        return "\n\n".join(outputs)
    else:
        return "一致する配列が見つかりませんでした。"

# Flaskのコールバック設定
@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

# LINEメッセージ受信時の処理
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_input = event.message.text
    result = predict_from_input(user_input)

    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text=result)
    )

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port)
