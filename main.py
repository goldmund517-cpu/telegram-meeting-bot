import os
import json
import tempfile
import requests
import subprocess
import re
from flask import Flask, request
import google.generativeai as genai
from openai import OpenAI
import httpx

app = Flask(__name__)

# 環境變數
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")

# 初始化
genai.configure(api_key=GEMINI_API_KEY)
openai_client = OpenAI(api_key=OPENAI_API_KEY, http_client=httpx.Client())

TELEGRAM_API = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}"

# 暫存對話狀態
user_sessions = {}

# ─────────────────────────────────────────
# Webhook 進入點
# ─────────────────────────────────────────
@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json()
    if not data:
        return "OK"

    message = data.get("message") or data.get("edited_message")
    if not message:
        return "OK"

    chat_id = message["chat"]["id"]
    user_id = str(message["from"]["id"])

    # 語音/音訊檔案
    audio = message.get("voice") or message.get("audio") or message.get("document")
    text = message.get("text", "").strip()

    if audio:
        handle_audio(chat_id, user_id, audio, message)
    elif text:
        handle_text(chat_id, user_id, text)

    return "OK"

@app.route("/", methods=["GET"])
def health():
    return "Telegram Bot is running!"

# ─────────────────────────────────────────
# 傳送訊息
# ─────────────────────────────────────────
def send_message(chat_id, text):
    requests.post(f"{TELEGRAM_API}/sendMessage", json={
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "HTML"
    })

def send_document(chat_id, file_path, filename):
    with open(file_path, "rb") as f:
        requests.post(f"{TELEGRAM_API}/sendDocument", 
            data={"chat_id": chat_id},
            files={"document": (filename, f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
        )

# ─────────────────────────────────────────
# 處理音訊
# ─────────────────────────────────────────
def handle_audio(chat_id, user_id, audio, message):
    send_message(chat_id, "⏳ 收到音訊，處理中請稍候...\n（較長的錄音可能需要1~2分鐘）")

    try:
        # 取得檔案
        file_id = audio.get("file_id")
        file_info = requests.get(f"{TELEGRAM_API}/getFile?file_id={file_id}").json()
        file_path = file_info["result"]["file_path"]
        file_url = f"https://api.telegram.org/file/bot{TELEGRAM_BOT_TOKEN}/{file_path}"

        # 下載檔案
        ext = file_path.split(".")[-1] if "." in file_path else "ogg"
        with tempfile.NamedTemporaryFile(suffix=f".{ext}", delete=False) as f:
            content = requests.get(file_url).content
            f.write(content)
            audio_path = f.name

        # Gemini 逐字稿
        transcript, speaker_samples, meeting_info = transcribe_with_gemini(audio_path)
        os.unlink(audio_path)

        # 儲存 session
        user_sessions[user_id] = {
            "state": "waiting_speaker_confirm" if meeting_info["confirmed"] else "waiting_meeting_info",
            "transcript": transcript,
            "speaker_samples": speaker_samples,
            "meeting_info": meeting_info,
        }

        if not meeting_info["confirmed"]:
            send_message(chat_id, 
                "📋 語音中未偵測到會議資訊，請提供：\n\n"
                "請回覆格式：\n"
                "<code>日期=2024/01/15，名稱=Q3預算會議</code>"
            )
        else:
            msg = build_speaker_confirm_message(speaker_samples, meeting_info)
            send_message(chat_id, msg)

    except Exception as e:
        send_message(chat_id, f"❌ 處理音訊時發生錯誤：{str(e)}\n請重新傳送。")

# ─────────────────────────────────────────
# 處理文字
# ─────────────────────────────────────────
def handle_text(chat_id, user_id, text):
    session = user_sessions.get(user_id)

    if not session:
        send_message(chat_id, 
            "👋 歡迎使用會議記錄助理！\n\n"
            "請直接傳送會議錄音檔，支援格式：\n"
            "• 語音訊息\n• mp3、m4a、wav、ogg、mp4 等"
        )
        return

    state = session.get("state")

    # 等待會議資訊
    if state == "waiting_meeting_info":
        meeting_info = parse_meeting_info(text)
        if not meeting_info:
            send_message(chat_id, 
                "格式不正確，請使用：\n"
                "<code>日期=2024/01/15，名稱=Q3預算會議</code>"
            )
            return
        session["meeting_info"].update(meeting_info)
        session["meeting_info"]["confirmed"] = True
        session["state"] = "waiting_speaker_confirm"
        user_sessions[user_id] = session
        msg = build_speaker_confirm_message(session["speaker_samples"], session["meeting_info"])
        send_message(chat_id, msg)

    # 等待語者確認
    elif state == "waiting_speaker_confirm":
        speaker_map = parse_speaker_map(text)
        if not speaker_map:
            send_message(chat_id,
                "格式不正確，請使用：\n"
                "<code>語者1=王經理，語者2=我，語者3=陳會計</code>"
            )
            return

        session["speaker_map"] = speaker_map
        session["state"] = "generating"
        user_sessions[user_id] = session

        send_message(chat_id, "✅ 收到！正在生成會議記錄 Word 檔，請稍候...")

        try:
            word_path, filename = generate_meeting_word(
                session["transcript"],
                session["speaker_map"],
                session["meeting_info"]
            )
            send_message(chat_id, "📄 會議記錄已完成，傳送檔案中...")
            send_document(chat_id, word_path, filename)
            os.unlink(word_path)
            del user_sessions[user_id]
        except Exception as e:
            send_message(chat_id, f"❌ 生成 Word 時發生錯誤：{str(e)}")

# ─────────────────────────────────────────
# Gemini 逐字稿
# ─────────────────────────────────────────
def transcribe_with_gemini(audio_path):
    model = genai.GenerativeModel("gemini-1.5-pro")

    ext = audio_path.split(".")[-1].lower()
    mime_map = {
        "ogg": "audio/ogg", "mp3": "audio/mpeg", "m4a": "audio/mp4",
        "wav": "audio/wav", "mp4": "audio/mp4", "aac": "audio/aac",
        "flac": "audio/flac", "webm": "audio/webm"
    }
    mime_type = mime_map.get(ext, "audio/mpeg")

    with open(audio_path, "rb") as f:
        audio_data = f.read()

    prompt = """請分析這段音訊，完成以下任務：

1. 產生完整逐字稿，格式：[語者N][時間戳] 發言內容
2. 嘗試從對話判斷會議日期與名稱

只回傳 JSON，不要其他文字：
{
  "transcript": "[語者1][00:00] 內容\\n[語者2][00:30] 內容",
  "speakers": ["語者1", "語者2"],
  "meeting_date": "2024/01/15 或 null",
  "meeting_name": "名稱 或 null",
  "speaker_first_appearance": {"語者1": "00:00", "語者2": "00:30"},
  "speaker_samples": {"語者1": ["句子1", "句子2"], "語者2": ["句子1"]},
  "total_duration_seconds": 3600
}

語者取樣規則：
- 每位語者取前2~3句
- 若第一次出現超過總時長1/3，取5句"""

    response = model.generate_content([
        {"mime_type": mime_type, "data": audio_data},
        prompt
    ])

    raw = re.sub(r"```json|```", "", response.text.strip()).strip()
    data = json.loads(raw)

    transcript = data["transcript"]
    total_seconds = data.get("total_duration_seconds", 3600)
    first_appearance = data.get("speaker_first_appearance", {})
    late_threshold = total_seconds / 3

    speaker_samples = {}
    for speaker, samples in data.get("speaker_samples", {}).items():
        ts = first_appearance.get(speaker, "00:00")
        seconds = timestamp_to_seconds(ts)
        speaker_samples[speaker] = {
            "samples": samples,
            "late": seconds > late_threshold,
            "first_ts": ts
        }

    meeting_info = {
        "date": data.get("meeting_date"),
        "name": data.get("meeting_name"),
        "confirmed": bool(data.get("meeting_date") and data.get("meeting_name"))
    }

    return transcript, speaker_samples, meeting_info

def timestamp_to_seconds(ts):
    try:
        parts = ts.split(":")
        if len(parts) == 2:
            return int(parts[0]) * 60 + int(parts[1])
        elif len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
    except:
        return 0
    return 0

# ─────────────────────────────────────────
# 建立語者確認訊息
# ─────────────────────────────────────────
def build_speaker_confirm_message(speaker_samples, meeting_info):
    lines = ["🎙️ 請確認以下語者身份：\n"]

    for speaker, info in speaker_samples.items():
        if info["late"]:
            lines.append(f"⚠️【{speaker}】（{info['first_ts']} 才出現，提供較多內容）")
        else:
            lines.append(f"【{speaker}】（{info['first_ts']} 出現）")
        for s in info["samples"]:
            lines.append(f"  「{s}」")
        lines.append("")

    if meeting_info.get("date") and meeting_info.get("name"):
        lines.append(f"📋 偵測到會議資訊：")
        lines.append(f"日期：{meeting_info['date']}")
        lines.append(f"名稱：{meeting_info['name']}")
        lines.append("")

    lines.append("請回覆語者對應，例如：")
    lines.append("<code>語者1=王經理，語者2=我，語者3=陳會計</code>")

    return "\n".join(lines)

# ─────────────────────────────────────────
# 解析回覆
# ─────────────────────────────────────────
def parse_speaker_map(text):
    matches = re.findall(r"語者(\d+)\s*=\s*([^\s，,、]+)", text)
    if not matches:
        return None
    return {f"語者{num}": name for num, name in matches}

def parse_meeting_info(text):
    date_match = re.search(r"日期\s*=\s*([^\s，,、]+)", text)
    name_match = re.search(r"名稱\s*=\s*(.+?)(?:[，,]|$)", text)
    if not date_match or not name_match:
        return None
    return {"date": date_match.group(1).strip(), "name": name_match.group(1).strip()}

# ─────────────────────────────────────────
# GPT 生成內容
# ─────────────────────────────────────────
def generate_meeting_content(transcript, speaker_map, meeting_info):
    replaced = transcript
    for key, name in speaker_map.items():
        replaced = replaced.replace(key, name)

    prompt = f"""你是專業會議記錄助理，請整理以下逐字稿為正式會議記錄。

會議日期：{meeting_info['date']}
會議名稱：{meeting_info['name']}
出席人員：{', '.join(speaker_map.values())}

逐字稿：
{replaced}

只回傳 JSON，不要其他文字：
{{
  "meeting_date": "yyyy/mm/dd",
  "meeting_name": "會議名稱",
  "attendees": ["人員1", "人員2"],
  "location": "",
  "recorder": "",
  "topics": [
    {{"title": "主題標題", "points": ["重點1", "重點2"]}}
  ],
  "action_items": [
    {{"category": "分類", "content": "內容", "owner": "負責人", "due_date": "時間", "notes": "備註"}}
  ],
  "pending_items": ["未決事項1"],
  "remarks": ["備註1"]
}}"""

    response = openai_client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3
    )
    raw = re.sub(r"```json|```", "", response.choices[0].message.content.strip()).strip()
    return json.loads(raw)

# ─────────────────────────────────────────
# 生成 Word
# ─────────────────────────────────────────
def generate_meeting_word(transcript, speaker_map, meeting_info):
    content = generate_meeting_content(transcript, speaker_map, meeting_info)

    date_str = content["meeting_date"].replace("/", "")
    name_str = content["meeting_name"]
    filename = f"{date_str}_{name_str}.docx"
    output_path = f"/tmp/{filename}"

    js_script = build_docx_js(content, output_path)
    js_path = "/tmp/gen_doc.mjs"
    with open(js_path, "w", encoding="utf-8") as f:
        f.write(js_script)

    result = subprocess.run(["node", js_path], capture_output=True, text=True)
    if result.returncode != 0:
        raise Exception(f"Node error: {result.stderr}")

    return output_path, filename

def build_docx_js(content, output_path):
    topics_js = ""
    for topic in content.get("topics", []):
        points_js = "\n".join([
            f'new Paragraph({{ numbering: {{ reference: "bullets", level: 0 }}, children: [new TextRun({{ text: {json.dumps(p)}, font: "Microsoft JhengHei", size: 24 }})] }}),'
            for p in topic["points"]
        ])
        topics_js += f"""
new Paragraph({{ spacing: {{ before: 120, after: 60 }}, children: [new TextRun({{ text: {json.dumps(topic["title"])}, bold: true, size: 28, font: "Microsoft JhengHei" }})] }}),
{points_js}"""

    def make_cell(text, width):
        return f"""new TableCell({{
  width: {{ size: {width}, type: WidthType.DXA }},
  borders: {{ top: bb, bottom: bb, left: bb, right: bb }},
  margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
  children: [new Paragraph({{ children: [new TextRun({{ text: {json.dumps(str(text))}, font: "Microsoft JhengHei", size: 22 }})] }})]
}})"""

    action_rows = ""
    for item in content.get("action_items", []):
        action_rows += f"""new TableRow({{ children: [
  {make_cell(item.get("category",""), 1200)},
  {make_cell(item.get("content",""), 2800)},
  {make_cell(item.get("owner",""), 1400)},
  {make_cell(item.get("due_date",""), 1800)},
  {make_cell(item.get("notes",""), 1800)},
] }}),"""

    pending_js = "\n".join([
        f'new Paragraph({{ numbering: {{ reference: "bullets", level: 0 }}, children: [new TextRun({{ text: {json.dumps(p)}, font: "Microsoft JhengHei", size: 24 }})] }}),'
        for p in content.get("pending_items", []) + content.get("remarks", [])
    ])

    attendees = "、".join(content.get("attendees", []))
    title_text = f"{content['meeting_date']} {content['meeting_name']}紀錄"

    return f"""
import {{ Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
         AlignmentType, LevelFormat, BorderStyle, WidthType }} from 'docx';
import fs from 'fs';

const bb = {{ style: BorderStyle.SINGLE, size: 4, color: "000000" }};

const doc = new Document({{
  numbering: {{ config: [{{ reference: "bullets", levels: [{{
    level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
    style: {{ paragraph: {{ indent: {{ left: 720, hanging: 360 }} }} }}
  }}] }}] }},
  sections: [{{
    properties: {{ page: {{ size: {{ width: 11906, height: 16838 }}, margin: {{ top: 1440, right: 1440, bottom: 1440, left: 1440 }} }} }},
    children: [
      new Paragraph({{ alignment: AlignmentType.CENTER, spacing: {{ after: 240 }},
        children: [new TextRun({{ text: {json.dumps(title_text)}, bold: true, size: 48, font: "Microsoft JhengHei" }})] }}),

      new Paragraph({{ spacing: {{ before: 240, after: 120 }},
        children: [new TextRun({{ text: "一、會議基本資訊", bold: true, size: 32, font: "Microsoft JhengHei" }})] }}),
      new Paragraph({{ children: [new TextRun({{ text: "1. 會議名稱：{content['meeting_name']}", font: "Microsoft JhengHei", size: 24 }})] }}),
      new Paragraph({{ children: [new TextRun({{ text: "2. 會議日期：{content['meeting_date']}", font: "Microsoft JhengHei", size: 24 }})] }}),
      new Paragraph({{ children: [new TextRun({{ text: "3. 出席人員：{attendees}", font: "Microsoft JhengHei", size: 24 }})] }}),
      new Paragraph({{ children: [new TextRun({{ text: "4. 地點：{content.get('location','')}", font: "Microsoft JhengHei", size: 24 }})] }}),
      new Paragraph({{ children: [new TextRun({{ text: "5. 記錄人：{content.get('recorder','')}", font: "Microsoft JhengHei", size: 24 }})] }}),

      new Paragraph({{ spacing: {{ before: 240, after: 120 }},
        children: [new TextRun({{ text: "二、討論主題摘要與各方意見", bold: true, size: 32, font: "Microsoft JhengHei" }})] }}),
      {topics_js}

      new Paragraph({{ spacing: {{ before: 240, after: 120 }},
        children: [new TextRun({{ text: "三、決議事項與行動項目", bold: true, size: 32, font: "Microsoft JhengHei" }})] }}),
      new Table({{
        width: {{ size: 9000, type: WidthType.DXA }},
        columnWidths: [1200, 2800, 1400, 1800, 1800],
        rows: [
          new TableRow({{ tableHeader: true, children: [
            new TableCell({{ width: {{ size: 1200, type: WidthType.DXA }}, borders: {{ top: bb, bottom: bb, left: bb, right: bb }}, margins: {{ top: 80, bottom: 80, left: 120, right: 120 }}, shading: {{ fill: "D9D9D9" }}, children: [new Paragraph({{ children: [new TextRun({{ text: "分類", bold: true, font: "Microsoft JhengHei", size: 22 }})] }})] }}),
            new TableCell({{ width: {{ size: 2800, type: WidthType.DXA }}, borders: {{ top: bb, bottom: bb, left: bb, right: bb }}, margins: {{ top: 80, bottom: 80, left: 120, right: 120 }}, shading: {{ fill: "D9D9D9" }}, children: [new Paragraph({{ children: [new TextRun({{ text: "行動項目內容", bold: true, font: "Microsoft JhengHei", size: 22 }})] }})] }}),
            new TableCell({{ width: {{ size: 1400, type: WidthType.DXA }}, borders: {{ top: bb, bottom: bb, left: bb, right: bb }}, margins: {{ top: 80, bottom: 80, left: 120, right: 120 }}, shading: {{ fill: "D9D9D9" }}, children: [new Paragraph({{ children: [new TextRun({{ text: "負責人", bold: true, font: "Microsoft JhengHei", size: 22 }})] }})] }}),
            new TableCell({{ width: {{ size: 1800, type: WidthType.DXA }}, borders: {{ top: bb, bottom: bb, left: bb, right: bb }}, margins: {{ top: 80, bottom: 80, left: 120, right: 120 }}, shading: {{ fill: "D9D9D9" }}, children: [new Paragraph({{ children: [new TextRun({{ text: "預計完成時間", bold: true, font: "Microsoft JhengHei", size: 22 }})] }})] }}),
            new TableCell({{ width: {{ size: 1800, type: WidthType.DXA }}, borders: {{ top: bb, bottom: bb, left: bb, right: bb }}, margins: {{ top: 80, bottom: 80, left: 120, right: 120 }}, shading: {{ fill: "D9D9D9" }}, children: [new Paragraph({{ children: [new TextRun({{ text: "備註", bold: true, font: "Microsoft JhengHei", size: 22 }})] }})] }}),
          ] }}),
          {action_rows}
        ]
      }}),

      new Paragraph({{ spacing: {{ before: 240, after: 120 }},
        children: [new TextRun({{ text: "四、補充備註／未決事項追蹤", bold: true, size: 32, font: "Microsoft JhengHei" }})] }}),
      {pending_js}
    ]
  }}]
}});

Packer.toBuffer(doc).then(buf => {{ fs.writeFileSync({json.dumps(output_path)}, buf); console.log("done"); }});
"""

# ─────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
