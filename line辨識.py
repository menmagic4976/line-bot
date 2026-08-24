import os, sys, requests, time, json, base64, threading, re
from datetime import datetime
from pathlib import Path
from flask import Flask, request, abort

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment
    from linebot.v3 import WebhookHandler
    from linebot.v3.exceptions import InvalidSignatureError
    from linebot.v3.messaging import (
        Configuration, ApiClient, MessagingApi, MessagingApiBlob,
        ReplyMessageRequest, TextMessage
    )
    from linebot.v3.webhooks import MessageEvent, ImageMessageContent, TextMessageContent
    import cv2
    import numpy as np
except Exception as e:
    print(f"\n❌ 缺少必要套件: {e}"); sys.exit()

LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
LINE_CHANNEL_SECRET       = os.environ.get("LINE_CHANNEL_SECRET", "")
DATALAB_API_KEY  = os.environ.get("DATALAB_API_KEY", "")
DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "")
CLAUDE_API_KEY   = os.environ.get("CLAUDE_API_KEY", "")
EXCEL_FILE = Path(__file__).parent / "records.xlsx"
MISSING = "未找到"
KEYS = ["工單 (Part No)", "型號 (Model)", "數量 (Quantity)", "儲位 (Location)", "業單 (Sales Order)"]
ID_CHARS = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"  # 排除易混淆字元 I,O,0,1
VALID_LOC = ([f"A{i}" for i in range(1, 13)] + [f"B{i}" for i in range(1, 7)] + ["0S08", "3F", "NG", "B2", "B3"])
LOC_FIX = {"BZ":"B2","82":"B2","B 2":"B2","86":"B6","BG":"B6","81":"B1","83":"B3","84":"B4","85":"B5","ALL":"A11","A1L":"A11","A|1":"A11","OS08":"0S08","0SO8":"0S08","OSO8":"0S08","0508":"0S08"}

app = Flask(__name__)
handler = WebhookHandler(LINE_CHANNEL_SECRET)
configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
print(f"🔑 SECRET載入: {'OK' if LINE_CHANNEL_SECRET else 'EMPTY'} ({len(LINE_CHANNEL_SECRET)} chars)")
print(f"🔍 所有LINE相關變數: { {k:v[:4]+'...' for k,v in os.environ.items() if 'LINE' in k} }")

def _parse_location(v):
    v = v.upper().replace(" ", "").strip()
    return LOC_FIX[v] if v in LOC_FIX else (v if v in VALID_LOC else v)

def init_excel():
    if EXCEL_FILE.exists(): return
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "辨識紀錄"
    hd = ["辨識ID", "時間"] + KEYS
    for col, h in enumerate(hd, 1):
        cell = ws.cell(1, col, h); cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[cell.column_letter].width = 22
    wb.save(EXCEL_FILE)

def get_today_excel():
    """匯出當日資料為獨立 Excel 檔"""
    today = datetime.now().strftime("%m%d")
    try:
        wb_all = openpyxl.load_workbook(EXCEL_FILE)
        ws_all = wb_all.active
        wb_today = openpyxl.Workbook(); ws_today = wb_today.active; ws_today.title = "當日紀錄"
        hd = ["辨識ID", "時間"] + KEYS
        for col, h in enumerate(hd, 1):
            cell = ws_today.cell(1, col, h); cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")
            ws_today.column_dimensions[cell.column_letter].width = 22
        for row in ws_all.iter_rows(min_row=2, values_only=True):
            if row[1] and str(row[1]).startswith(today):  # 改成檢查第2欄（時間）
                ws_today.append(list(row))
        import io
        buf = io.BytesIO(); wb_today.save(buf); buf.seek(0)
        return buf.read()
    except Exception as e:
        print(f"當日匯出錯誤: {e}"); return None

def generate_record_id():
    """生成唯一辨識 ID：0824-1430-A3F5"""
    import random
    now = datetime.now()
    date_part = now.strftime("%m%d")
    time_part = now.strftime("%H%M")
    random_part = ''.join(random.choices(ID_CHARS, k=4))
    return f"{date_part}-{time_part}-{random_part}"

def append_excel_multi(results):
    init_excel()
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE); ws = wb.active
        record_ids = []
        for r in results:
            # 工單和業單都沒抓到就不寫入
            if r.get("工單 (Part No)") == MISSING and r.get("業單 (Sales Order)") == MISSING:
                continue
            if r.get("數量 (Quantity)") and r["數量 (Quantity)"] != MISSING:
                r["數量 (Quantity)"] = r["數量 (Quantity)"].lstrip(":：").strip()

            record_id = generate_record_id()
            record_ids.append(record_id)
            ws.append([record_id, datetime.now().strftime("%m%d-%H:%M")] + [r.get(k, MISSING) for k in KEYS])
        wb.save(EXCEL_FILE)
        return len(record_ids), record_ids
    except Exception as e:
        print(f"Excel寫入錯誤: {e}")
        return 0, []

def undo_last_record():
    """刪除最後 1 筆紀錄"""
    init_excel()
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE); ws = wb.active
        last_row = ws.max_row
        if last_row <= 1:
            return 0
        ws.delete_rows(last_row)
        wb.save(EXCEL_FILE)
        return 1
    except Exception as e:
        print(f"撤銷錯誤: {e}")
        return 0

def get_recent_records(n=10):
    """查詢最近 N 筆紀錄，回傳 [(row_index, record_id, time, part_no, model, qty, loc, sales)]"""
    init_excel()
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE); ws = wb.active
        last_row = ws.max_row
        if last_row <= 1:
            return []
        records = []
        for row_idx in range(max(2, last_row - n + 1), last_row + 1):
            row_data = [ws.cell(row_idx, col).value for col in range(1, 8)]  # 改成8欄（含ID）
            records.append((row_idx, *row_data))
        return list(reversed(records))  # 最新的在最前面
    except Exception as e:
        print(f"查詢錯誤: {e}")
        return []

def delete_record_by_row(row_idx):
    """刪除指定列（實際 Excel row number）"""
    init_excel()
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE); ws = wb.active
        if row_idx < 2 or row_idx > ws.max_row:
            return False, "超出範圍"
        row_data = [ws.cell(row_idx, col).value for col in range(1, 8)]  # 改成8欄
        ws.delete_rows(row_idx)
        wb.save(EXCEL_FILE)
        return True, row_data
    except Exception as e:
        print(f"刪除錯誤: {e}")
        return False, str(e)

def delete_record_by_id(record_id):
    """根據辨識 ID 刪除（從最新往回找第一筆）"""
    init_excel()
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE); ws = wb.active
        for row_idx in range(ws.max_row, 1, -1):
            cell_value = ws.cell(row_idx, 1).value  # 第1欄是辨識ID
            if cell_value and str(cell_value).strip() == record_id:
                row_data = [ws.cell(row_idx, col).value for col in range(1, 8)]
                ws.delete_rows(row_idx)
                wb.save(EXCEL_FILE)
                return True, row_data
        return False, "找不到該 ID"
    except Exception as e:
        print(f"刪除錯誤: {e}")
        return False, str(e)

def delete_record_by_part_no(part_no):
    """根據工單號碼刪除（從最新往回找第一筆）"""
    init_excel()
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE); ws = wb.active
        for row_idx in range(ws.max_row, 1, -1):
            cell_value = ws.cell(row_idx, 3).value  # 第3欄是工單（改成3因為前面加了ID欄）
            if cell_value and str(cell_value).strip() == part_no:
                row_data = [ws.cell(row_idx, col).value for col in range(1, 8)]
                ws.delete_rows(row_idx)
                wb.save(EXCEL_FILE)
                return True, row_data
        return False, "找不到該工單"
    except Exception as e:
        print(f"刪除錯誤: {e}")
        return False, str(e)

def vision_get_fields(img_b, missing_fields):
    """對圖片做視覺辨識，補齊missing_fields中指定的欄位"""
    try:
        field_prompts = []
        if "工單 (Part No)" in missing_fields:
            field_prompts.append("- 工單號碼(Part No)：12碼英數字，不以Q開頭，回傳格式如「工單:105239501A01」")
        if "型號 (Model)" in missing_fields:
            field_prompts.append("- 型號(Model)：3字元以上英數���合，回傳格式如「型號:CSM0005A」")
        if "儲位 (Location)" in missing_fields:
            field_prompts.append("- 儲位(Location)：合法值為A1~A12/B1~B6/0S08/3F/NG，手寫字B易誤判為8或6，回傳格式如「儲位:B6」")

        prompt = (
            "這是倉管單據圖片，請辨識以下欄位：\n"
            + "\n".join(field_prompts)
            + "\n\n找不到的欄位不要輸出。每行只輸出一個欄位，格式如上。"
        )
        resp = requests.post(
            "https://shaco.chat/api/v1/messages",
            headers={"x-api-key": CLAUDE_API_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json"},
            json={"model": "claude-haiku-4-5-20251001", "max_tokens": 1024,
                  "messages": [{"role": "user", "content": [
                      {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": base64.b64encode(img_b).decode()}},
                      {"type": "text", "text": prompt}
                  ]}]},
            timeout=15
        ).json()
        raw = next((item["text"] for item in resp["content"] if item["type"] == "text"), "")
        raw = re.sub(r'<INTERNAL_THINKING>.*?</INTERNAL_THINKING>', '', raw, flags=re.DOTALL)
        found = {}
        for line in raw.splitlines():
            line = line.strip()
            if line.startswith("工單:") and "工單 (Part No)" in missing_fields:
                val = line.split(":", 1)[1].strip()
                if len(val) == 12: found["工單 (Part No)"] = val
            elif line.startswith("型號:") and "型號 (Model)" in missing_fields:
                val = line.split(":", 1)[1].strip()
                if len(val) >= 3: found["型號 (Model)"] = val
            elif line.startswith("儲位:") and "儲位 (Location)" in missing_fields:
                val = line.split(":", 1)[1].strip()
                found["儲位 (Location)"] = _parse_location(val)
        return found
    except Exception as e:
        print(f"視覺辨識錯誤: {e}"); return {}

def llm_parse(raw_text):
    EMPTY = {k: MISSING for k in KEYS}
    try:
        prompt = (
            "你是倉管單據解析助手。以下是OCR文字，請提取明細。\n"
            "若含多筆明細（多個業單號），每筆獨立成一個物件。\n\n"
            "欄位規則：\n"
            "1. 工單(Part No)：剛好12碼英數，不以Q開頭，對應「工單號碼」欄位\n"
            "2. 型號(Model)：至少3字元英數組合\n"
            "3. 數量(Quantity)：數字或分數格式如120或120/2\n"
            "4. 儲位(Location)：英數組合如A9/B6/0S08/3F，多為手寫，找不到填「未找到」\n"
            "5. 業單(Sales Order)：剛好7碼純數字、以5開頭，不符填「未找到」\n\n"
            "找不到的欄位一律填「未找到」。只回傳JSON：\n"
            "{\"items\":[{\"工單 (Part No)\":\"...\",\"型號 (Model)\":\"...\",\"數量 (Quantity)\":\"...\",\"儲位 (Location)\":\"...\",\"業單 (Sales Order)\":\"...\"}]}\n\n"
            f"OCR文字：\n{raw_text}"
        )
        resp = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers={"Authorization": f"Bearer {DEEPSEEK_API_KEY}", "Content-Type": "application/json"},
            json={"model": "deepseek-chat", "messages": [{"role": "user", "content": prompt}], "response_format": {"type": "json_object"}},
            timeout=30
        ).json()
        items = json.loads(resp["choices"][0]["message"]["content"]).get("items", [])
        results = []
        for item in items:
            r = {k: str(item.get(k, MISSING)).strip() or MISSING for k in KEYS}
            if r["工單 (Part No)"] != MISSING and len(r["工單 (Part No)"]) != 12:
                r["工單 (Part No)"] = MISSING
            if r["儲位 (Location)"] != MISSING:
                r["儲位 (Location)"] = _parse_location(r["儲位 (Location)"])
            if r["業單 (Sales Order)"] != MISSING:
                if r["儲位 (Location)"] == MISSING:
                    r["儲位 (Location)"] = "B2"
                r["工單 (Part No)"] = MISSING
            results.append(r)
        return results or [EMPTY]
    except Exception as e:
        print(f"LLM解析錯誤: {e}"); return [EMPTY]

def preprocess_image(img_b):
    """圖片前處理：增強對比、去噪、銳化"""
    try:
        nparr = np.frombuffer(img_b, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

        # 1. 轉灰階
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # 2. 自適應對比增強（CLAHE）
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        enhanced = clahe.apply(gray)

        # 3. 去噪
        denoised = cv2.fastNlMeansDenoising(enhanced, None, 10, 7, 21)

        # 4. 銳化
        kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]])
        sharpened = cv2.filter2D(denoised, -1, kernel)

        # 5. 二值化（讓文字更清晰）
        _, binary = cv2.threshold(sharpened, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

        # 轉回 JPEG bytes
        _, buffer = cv2.imencode('.jpg', binary, [cv2.IMWRITE_JPEG_QUALITY, 95])
        return buffer.tobytes()
    except Exception as e:
        print(f"圖片前處理失敗，使用原圖: {e}")
        return img_b

def cloud_ocr_process(img_b):
    try:
        # 前處理圖片
        processed_img = preprocess_image(img_b)

        resp = requests.post("https://www.datalab.to/api/v1/convert", headers={"X-API-Key": DATALAB_API_KEY}, files={"file": ("image.jpg", processed_img, "image/jpeg")}, data={"output_format": "markdown", "use_llm": "true", "engine": "chandra-3"}, timeout=15)
        if resp.status_code != 200:
            return [{k: MISSING for k in KEYS}]
        check_url = resp.json().get("request_check_url")
        raw_text = ""
        for _ in range(15):
            time.sleep(2)
            r = requests.get(check_url, headers={"X-API-Key": DATALAB_API_KEY}, timeout=10).json()
            if r.get("status") == "complete":
                raw_text = r.get("markdown", ""); break
        if not raw_text:
            return [{k: MISSING for k in KEYS}]
        print(f"\n--- 🛠️ 偵錯模式 --- \n{raw_text}\n----------------")
        results = llm_parse(raw_text)
        for r in results:
            missing = [k for k in ["工單 (Part No)", "型號 (Model)", "儲位 (Location)"] if r[k] == MISSING]
            if missing:
                found = vision_get_fields(img_b, missing)
                for k, v in found.items():
                    r[k] = v
        return results
    except Exception as e:
        print(f"連線異常: {e}"); return [{k: MISSING for k in KEYS}]

def reply_text(reply_token, text):
    with ApiClient(configuration) as api_client:
        MessagingApi(api_client).reply_message(
            ReplyMessageRequest(reply_token=reply_token, messages=[TextMessage(text=text)])
        )

def process_image_task(reply_token, img_b):
    results = cloud_ocr_process(img_b)
    written_count, record_ids = append_excel_multi(results)
    reply = "📋 AI 辨識結果\n"
    reply += f"已自動寫入 {written_count} 筆物料明細\n"
    if written_count > 0:
        reply += "💡 辨識ID：" + "、".join([rid.split('-')[2] for rid in record_ids]) + "\n"
        reply += "輸入「查詢」可查看完整紀錄\n"
    reply += "────────────────\n"
    for i, res in enumerate(results, 1):
        if len(results) > 1: reply += f"📦 第 {i} 筆明細：\n"
        for k, v in res.items():
            reply += f" {'✅' if v != MISSING else '❌'} {k}: {v}\n"
        if len(results) > 1: reply += "──────────\n"
    reply_text(reply_token, reply)

@app.route("/webhook", methods=['POST'])
def webhook():
    signature = request.headers.get('X-Line-Signature', '')
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=ImageMessageContent)
def handle_image(event):
    reply_token = event.reply_token
    with ApiClient(configuration) as api_client:
        blob_api = MessagingApiBlob(api_client)
        img_b = bytes(blob_api.get_message_content(event.message.id))
    threading.Thread(target=process_image_task, args=(reply_token, img_b), daemon=True).start()

@handler.add(MessageEvent, message=TextMessageContent)
def handle_text(event):
    target_id = getattr(event.source, 'group_id', None) or getattr(event.source, 'room_id', None) or event.source.user_id
    text = event.message.text.strip()
    text_lower = text.lower()
    with ApiClient(configuration) as api_client:
        line_bot_api = MessagingApi(api_client)

        # 查詢最近紀錄
        if text_lower in ["查詢", "query", "list"]:
            records = get_recent_records(10)
            if not records:
                reply = "❌ 目前沒有紀錄"
            else:
                reply = "📋 最近 10 筆紀錄：\n"
                for idx, (row_idx, record_id, time, part_no, model, qty, loc, sales) in enumerate(records, 1):
                    short_id = record_id.split('-')[2] if record_id and '-' in str(record_id) else "?"
                    part_display = part_no if part_no and part_no != MISSING else sales if sales and sales != MISSING else "無"
                    loc_display = loc if loc and loc != MISSING else "?"
                    reply += f"[{idx}] {short_id} | {time} | {part_display} | {loc_display}\n"
                reply += "\n💡 輸入「撤銷 A3F5」或「撤銷 1」刪除指定紀錄"

        # 撤銷指定 ID 或編號
        elif text_lower.startswith("撤銷 ") or text_lower.startswith("undo ") or text_lower.startswith("刪除 "):
            target = text.split(None, 1)[1].strip().upper()

            # 檢查是否為數字（編號）
            if target.isdigit():
                records = get_recent_records(10)
                idx = int(target)
                if 1 <= idx <= len(records):
                    row_idx = records[idx - 1][0]
                    success, data = delete_record_by_row(row_idx)
                    if success:
                        reply = f"✅ 已刪除第 {idx} 筆：{data[2]} ({data[6]})"
                    else:
                        reply = f"❌ 刪除失敗：{data}"
                else:
                    reply = f"❌ 編號超出範圍（1-{len(records)}）"

            # 檢查是否為 4 碼 ID（如 A3F5）
            elif len(target) == 4 and all(c in ID_CHARS for c in target):
                records = get_recent_records(10)
                full_id = None
                for rec in records:
                    if rec[1] and str(rec[1]).endswith(target):
                        full_id = rec[1]
                        break
                if full_id:
                    success, data = delete_record_by_id(full_id)
                    if success:
                        reply = f"✅ 已刪除 ID {target}：{data[2]} ({data[6]})"
                    else:
                        reply = f"❌ 刪除失敗：{data}"
                else:
                    reply = f"❌ 找不到 ID：{target}"

            # 檢查是否為 12 碼工單號
            elif len(target) == 12:
                success, data = delete_record_by_part_no(target)
                if success:
                    reply = f"✅ 已刪除工單 {target}"
                else:
                    reply = f"❌ {data}"

            else:
                reply = "❌ 格式錯誤\n請輸入：\n• 撤銷 1（編號）\n• 撤銷 A3F5（ID）\n• 刪除 105239501A01（工單號）"

        # 下載 Excel
        elif text_lower in ["excel", "dl", "下載"]:
            today_data = get_today_excel()
            if today_data:
                import uuid
                temp_url = f"https://line-bot-production-b2a9.up.railway.app/download/{uuid.uuid4().hex}"
                app.excel_temp = today_data
                reply = f"📊 當日 Excel 已產生\n點此下載：{temp_url}\n（連結10分鐘內有效）"
            else:
                reply = "❌ 今日尚無紀錄。"

        # 清空所有紀錄
        elif text_lower == "clear":
            if EXCEL_FILE.exists(): EXCEL_FILE.unlink()
            init_excel()
            reply = "🗑 紀錄已清空"

        else:
            return

        line_bot_api.reply_message(ReplyMessageRequest(
            reply_token=event.reply_token,
            messages=[TextMessage(text=reply)]
        ))


@app.route("/download/<file_id>")
def download_excel(file_id):
    if hasattr(app, 'excel_temp') and app.excel_temp:
        from flask import send_file
        import io
        return send_file(
            io.BytesIO(app.excel_temp),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='倉管辨識紀錄.xlsx'
        )
    abort(404)

if __name__ == "__main__":
    init_excel()
    print("🤖 LINE 倉管機器人啟動中...")
    print("📌 請用 ngrok 或部署到 Railway 取得 HTTPS webhook URL")
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)

