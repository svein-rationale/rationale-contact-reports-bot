"""
Claude Contact Reports Bot for Slack
Generates professional contact reports from meeting recordings or notes using Claude AI
"""

import os
import json
import requests
import tempfile
from datetime import datetime
from flask import Flask, request
from slack_bolt import App
from slack_bolt.adapter.flask import SlackRequestHandler
from anthropic import Anthropic
from google.oauth2.service_account import Credentials
from googleapiclient import discovery
from docx import Document
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = App(
    token=os.environ.get("SLACK_BOT_TOKEN"),
    signing_secret=os.environ.get("SLACK_SIGNING_SECRET")
)

flask_app = Flask(__name__)
handler = SlackRequestHandler(app)
anthropic_client = Anthropic(api_key=os.environ.get("CLAUDE_API_KEY"))

TEAM_DATA = {
    "account_managers": [
        "Svein Clouston",
        "Rowan Morrison",
        "Ursula Fairlie",
        "Gemma Greig",
        "Matthew Collins",
        "Carron Easdon",
    ],
    "senior_oversight": [
        "Svein Clouston",
        "Rowan Morrison",
        "Helen Davidson",
        "Ursula Fairlie",
    ],
    "client_mapping": {
        "hargreaves-lansdown": "Hargreaves Lansdown",
        "bupa-global": "Bupa Global",
        "innovate-uk": "Innovate UK",
        "compare-the-market": "Compare The Market",
        "innova-nanojet": "Innova Nanojet",
    }
}

active_workflows = {}

class ContactReportWorkflow:
    def __init__(self, channel, channel_name):
        self.channel = channel
        self.channel_name = channel_name
        self.client_name = self._extract_client_name()
        self.state = "awaiting_project_name"
        self.project_name = None
        self.attendees = []
        self.project_am = None
        self.senior_oversight = None
        self.context = None

    def _extract_client_name(self):
        channel_clean = self.channel_name.lower().replace("#", "").replace("-", " ").title()
        for key, value in TEAM_DATA["client_mapping"].items():
            if key == self.channel_name.lower().replace("#", ""):
                return value
        return channel_clean

def get_google_drive_service():
    try:
        credentials_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
        if not credentials_json:
            raise Exception("No Google credentials")
        credentials_dict = json.loads(credentials_json)
        credentials = Credentials.from_service_account_info(
            credentials_dict,
            scopes=['https://www.googleapis.com/auth/drive']
        )
        service = discovery.build('drive', 'v3', credentials=credentials)
        return service
    except Exception as e:
        logger.error(f"Drive service error: {str(e)}")
        raise

def extract_text_from_docx(file_path):
    try:
        doc = Document(file_path)
        text = ""
        for para in doc.paragraphs:
            if para.text.strip():
                text += para.text + "\n"
        return text.strip() if text else "[No text found]"
    except Exception as e:
        logger.error(f"DOCX extraction error: {str(e)}")
        raise

def generate_report_with_claude(content, workflow):
    try:
        prompt = f"""You are a professional business analyst. Extract the following from this meeting content:

1. Background: Client situation, challenges, context
2. The Ask: What they're requesting, timeline, budget, decision process
3. Actions: Next steps (as bullet points)
4. Key Points: Important discussion topics (as bullet points)

Content:
{content}

Respond in JSON format:
{{"background": "...", "the_ask": "...", "actions": [...], "key_points": [...]}}"""

        message = anthropic_client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}]
        )
        
        response_text = message.content[0].text
        json_start = response_text.find('{')
        json_end = response_text.rfind('}') + 1
        if json_start >= 0 and json_end > json_start:
            json_str = response_text[json_start:json_end]
            return json.loads(json_str)
        
        return {
            "background": "Meeting analysis",
            "the_ask": "See notes",
            "actions": ["Review notes", "Follow up"],
            "key_points": []
        }
    except Exception as e:
        logger.error(f"Claude error: {str(e)}")
        raise

def create_word_document(workflow, report_data):
    try:
        doc = Document()
        
        header_table = doc.add_table(rows=2, cols=4)
        header_table.style = 'Light Grid Accent 1'
        
        cells = header_table.rows[0].cells
        cells[0].text = "Client"
        cells[1].text = workflow.client_name
        cells[2].text = "Project Name"
        cells[3].text = workflow.project_name
        
        cells = header_table.rows[1].cells
        cells[0].text = "Meeting"
        meeting_date = datetime.now().strftime("%Y-%m-%d %H:%M")
        attendees_str = ", ".join(workflow.attendees) if workflow.attendees else "N/A"
        cells[1].text = f"{meeting_date} | {attendees_str}"
        cells[2].text = "Project AM"
        cells[3].text = workflow.project_am or "N/A"
        
        doc.add_paragraph()
        doc.add_heading("Meeting Notes", level=2)
        doc.add_paragraph("[Meeting notes would appear here]")
        
        doc.add_heading("Background", level=2)
        bg = report_data.get("background", "")
        doc.add_paragraph(bg if bg else "Client background")
        
        doc.add_heading("The Ask", level=2)
        ask = report_data.get("the_ask", "")
        doc.add_paragraph(ask if ask else "Client request details")
        
        doc.add_heading("Actions", level=2)
        actions = report_data.get("actions", [])
        if actions:
            for action in actions:
                doc.add_paragraph(action, style='List Bullet')
        else:
            doc.add_paragraph("Action items TBD", style='List Bullet')
        
        doc.add_heading("AOB", level=2)
        doc.add_paragraph("Additional notes")
        
        temp_dir = tempfile.gettempdir()
        filename = f"{datetime.now().strftime('%Y-%m-%d')}_ContactReport_{workflow.client_name.replace(' ', '_')}.docx"
        doc_path = os.path.join(temp_dir, filename)
        doc.save(doc_path)
        
        return doc_path
    except Exception as e:
        logger.error(f"Document creation error: {str(e)}")
        raise

def upload_to_google_drive(doc_path, workflow):
    try:
        service = get_google_drive_service()
        
        results = service.files().list(
            q="name='Contact Reports' and mimeType='application/vnd.google-apps.folder'",
            spaces='drive',
            fields='files(id)',
            pageSize=1
        ).execute()
        
        folders = results.get('files', [])
        if not folders:
            raise Exception("Contact Reports folder not found")
        
        folder_id = folders[0]['id']
        
        file_metadata = {
            'name': os.path.basename(doc_path),
            'parents': [folder_id]
        }
        
        media = discovery.MediaFileUpload(doc_path, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink'
        ).execute()
        
        return file.get('webViewLink')
    except Exception as e:
        logger.error(f"Drive upload error: {str(e)}")
        raise

@app.command("/contact-report")
def handle_contact_report_command(ack, body, say):
    ack()
    
    channel = body["channel_id"]
    channel_name = body.get("channel_name", "general")
    user_id = body["user_id"]
    
    workflow = ContactReportWorkflow(channel, channel_name)
    workflow_key = f"{user_id}_{channel}"
    active_workflows[workflow_key] = workflow
    
    say(f"📋 Starting Contact Report for: *{workflow.client_name}*\n\nStep 1/5: What's the project name?")

@app.event("message")
def handle_message_events(body, say, logger):
    event = body.get("event", {})
    
    if event.get("bot_id"):
        return
    
    if event.get("subtype") == "file_share":
        say("File uploads coming soon!")
        return
    
    user_id = event.get("user")
    channel = event.get("channel")
    text = event.get("text", "").strip()
    
    if not text:
        return
    
    workflow_key = f"{user_id}_{channel}"
    if workflow_key not in active_workflows:
        return
    
    workflow = active_workflows[workflow_key]
    
    if workflow.state == "awaiting_project_name":
        workflow.project_name = text
        workflow.state = "awaiting_attendees"
        say("Step 2/5: Who attended the meeting? (names, comma-separated)")
    
    elif workflow.state == "awaiting_attendees":
        workflow.attendees = [name.strip() for name in text.split(",")]
        workflow.state = "awaiting_am"
        am_list = "\n".join([f"• {am}" for am in TEAM_DATA["account_managers"]])
        say(f"Step 3/5: Who's the Project AM?\n{am_list}")
    
    elif workflow.state == "awaiting_am":
        workflow.project_am = text
        workflow.state = "awaiting_oversight"
        oversight_list = "\n".join([f"• {so}" for so in TEAM_DATA["senior_oversight"]])
        say(f"Step 4/5: Who's Senior Oversight?\n{oversight_list}")
    
    elif workflow.state == "awaiting_oversight":
        workflow.senior_oversight = text
        workflow.state = "awaiting_context"
        say("Step 5/5: Any other context? (Type 'none' if nothing to add)")
    
    elif workflow.state == "awaiting_context":
        workflow.context = text if text.lower() not in ["none", "no"] else ""
        workflow.state = "awaiting_file"
        say("✅ Got it! Upload meeting file:\n📁 MP3/MP4 (transcribe)\n📄 DOCX (extract text)")

@flask_app.route("/slack/events", methods=["POST"])
def slack_events():
    return handler.handle(request)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3000))
    flask_app.run(host="0.0.0.0", port=port, debug=False)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3000))
    flask_app.run(host="0.0.0.0", port=port, debug=False)
