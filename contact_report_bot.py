"""
Claude Contact Reports Bot for Slack
Generates professional contact reports from meeting recordings or notes using Claude AI
"""

import os
import json
import requests
import tempfile
from datetime import datetime
from pathlib import Path
from flask import Flask, request
from slack_bolt import App
from slack_bolt.adapter.flask import SlackRequestHandler
from anthropic import Anthropic
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request
from googleapiclient import discovery
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import logging

# Initialize logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize Slack app
app = App(
    token=os.environ.get("SLACK_BOT_TOKEN"),
    signing_secret=os.environ.get("SLACK_SIGNING_SECRET")
)

# Initialize Flask app
flask_app = Flask(__name__)
handler = SlackRequestHandler(app)

# Initialize Anthropic client
anthropic_client = Anthropic(api_key=os.environ.get("CLAUDE_API_KEY"))

# Team data
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

# Store active workflows by user+channel
active_workflows = {}

class ContactReportWorkflow:
    """Manages the state of a contact report workflow"""
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
        self.file_path = None
        self.file_content = None

    def _extract_client_name(self):
        """Extract client name from channel name"""
        channel_clean = self.channel_name.lower().replace("#", "").replace("-", " ").title()
        # Check if it's in the mapping
        for key, value in TEAM_DATA["client_mapping"].items():
            if key == self.channel_name.lower().replace("#", ""):
                return value
        return channel_clean

def download_slack_file(file_id, file_name, token):
    """Download a file from Slack"""
    try:
        # Get file info to get download URL
        url = f"https://slack.com/api/files.info?file={file_id}"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        file_info = response.json()
        if not file_info.get("ok"):
            raise Exception(f"Slack API error: {file_info.get('error')}")
        
        file_obj = file_info.get("file", {})
        download_url = file_obj.get("url_private")
        
        if not download_url:
            raise Exception("No download URL found")
        
        # Download the file
        file_response = requests.get(download_url, headers=headers)
        file_response.raise_for_status()
        
        # Save to temp file
        temp_dir = tempfile.gettempdir()
        file_path = os.path.join(temp_dir, file_name)
        
        with open(file_path, "wb") as f:
            f.write(file_response.content)
        
        return file_path
    except Exception as e:
        logger.error(f"Error downloading file: {str(e)}")
        raise

def transcribe_audio(file_path):
    """Transcribe audio file (MP3/MP4) using Google Cloud Speech-to-Text"""
    try:
        from google.cloud import speech
        import io
        
        # Initialize Speech-to-Text client using service account
        credentials_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
        if not credentials_json:
            return "[Audio transcription skipped - no Google credentials]"
        
        credentials_dict = json.loads(credentials_json)
        credentials = Credentials.from_service_account_info(credentials_dict)
        
        client = speech.SpeechClient(credentials=credentials)
        
        # Read audio file
        with io.open(file_path, "rb") as audio_file:
            content = audio_file.read()
        
        # Prepare audio config
        audio = speech.RecognitionAudio(content=content)
        config = speech.RecognitionConfig(
            encoding=speech.RecognitionConfig.AudioEncoding.MP3,
            sample_rate_hertz=16000,
            language_code="en-US",
            enable_automatic_punctuation=True,
        )
        
        # Transcribe
        response = client.recognize(config=config, audio=audio)
        
        # Extract transcript
        transcript = ""
        for result in response.results:
            for alternative in result.alternatives:
                transcript += alternative.transcript + " "
        
        return transcript.strip() if transcript else "[No speech detected in audio]"
    except Exception as e:
        logger.error(f"Error transcribing audio: {str(e)}")
        return f"[Transcription error: {str(e)}]"

def extract_text_from_docx(file_path):
    """Extract text from DOCX file"""
    try:
        doc = Document(file_path)
        text = ""
        for para in doc.paragraphs:
            if para.text.strip():
                text += para.text + "\n"
        return text.strip() if text else "[No text found in document]"
    except Exception as e:
        logger.error(f"Error extracting from DOCX: {str(e)}")
        raise

def generate_report_with_claude(content, workflow):
    """Use Claude to extract meeting insights and generate report"""
    try:
        prompt = f"""You are a professional business analyst. Extract the following information from this meeting transcript or notes:

1. Background: What is the client's current situation, challenges, and context?
2. The Ask: What specific request or project is the client asking for? Include timeline, budget, and decision process if mentioned.
3. Actions: What are the next steps and action items? List as bullet points.
4. Key Discussion Points: Any important topics discussed (as bullet points).

Meeting Content:
{content}

Provide the response in this exact JSON format:
{{
  "background": "...",
  "the_ask": "...",
  "actions": ["action 1", "action 2", ...],
  "key_points": ["point 1", "point 2", ...]
}}"""

        message = anthropic_client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        
        response_text = message.content[0].text
        
        # Try to parse JSON from response
        try:
            # Find JSON in response
            json_start = response_text.find('{')
            json_end = response_text.rfind('}') + 1
            if json_start >= 0 and json_end > json_start:
                json_str = response_text[json_start:json_end]
                return json.loads(json_str)
        except:
            pass
        
        # Fallback if JSON parsing fails
        return {
            "background": "Meeting analysis in progress",
            "the_ask": "See transcript for details",
            "actions": ["Review meeting notes", "Follow up with client"],
            "key_points": []
        }
    except Exception as e:
        logger.error(f"Error generating report with Claude: {str(e)}")
        raise

def get_google_drive_service():
    """Get authorized Google Drive service"""
    try:
        credentials_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
        if not credentials_json:
            raise Exception("No Google service account credentials")
        
        credentials_dict = json.loads(credentials_json)
        credentials = Credentials.from_service_account_info(
            credentials_dict,
            scopes=['https://www.googleapis.com/auth/drive']
        )
        
        service = discovery.build('drive', 'v3', credentials=credentials)
        return service
    except Exception as e:
        logger.error(f"Error creating Drive service: {str(e)}")
        raise

def upload_to_google_drive(doc_path, workflow):
    """Upload document to Google Drive"""
    try:
        service = get_google_drive_service()
        
        # Get Contact Reports folder ID
        results = service.files().list(
            q="name='Contact Reports' and mimeType='application/vnd.google-apps.folder'",
            spaces='drive',
            fields='files(id)',
            pageSize=1
        ).execute()
        
        folders = results.get('files', [])
        if not folders:
            raise Exception("Contact Reports folder not found in Google Drive")
        
        folder_id = folders[0]['id']
        
        # Upload file to folder
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
        logger.error(f"Error uploading to Drive: {str(e)}")
        raise

def create_word_document(workflow, report_data):
    """Create branded Word document with contact report"""
    try:
        doc = Document()
        
        # Add logo/header section
        header_table = doc.add_table(rows=2, cols=4)
        header_table.style = 'Light Grid Accent 1'
        
        # Header row 1 - Client info
        cells = header_table.rows[0].cells
        cells[0].text = "Client"
        cells[1].text = workflow.client_name
        cells[2].text = "Project Name"
        cells[3].text = workflow.project_name
        
        # Header row 2 - Meeting info
        cells = header_table.rows[1].cells
        cells[0].text = "Meeting"
        meeting_date = datetime.now().strftime("%Y-%m-%d %H:%M")
        attendees_str = ", ".join(workflow.attendees) if workflow.attendees else "N/A"
        cells[1].text = f"{meeting_date} | {attendees_str}"
        cells[2].text = "Project AM"
        cells[3].text = workflow.project_am or "N/A"
        
        # Add some space
        doc.add_paragraph()
        
        # Meeting notes section
        doc.add_heading("Meeting Notes", level=2)
        doc.add_paragraph("[Meeting notes and transcript would appear here if provided]")
        
        # Background section
        doc.add_heading("Background", level=2)
        bg = report_data.get("background", "")
        doc.add_paragraph(bg if bg else "Client background and context to be filled in.")
        
        # The Ask section
        doc.add_heading("The Ask", level=2)
        ask = report_data.get("the_ask", "")
        doc.add_paragraph(ask if ask else "Client request details to be filled in.")
        
        # Actions section
        doc.add_heading("Actions", level=2)
        actions = report_data.get("actions", [])
        if actions:
            for action in actions:
                doc.add_paragraph(action, style='List Bullet')
        else:
            doc.add_paragraph("Action items to be determined", style='List Bullet')
        
        # AOB section
        doc.add_heading("AOB (Any Other Business)", level=2)
        doc.add_paragraph("Additional notes or follow-up items")
        
        # Save document
        temp_dir = tempfile.gettempdir()
        filename = f"{datetime.now().strftime('%Y-%m-%d')}_ContactReport_{workflow.client_name.replace(' ', '_')}.docx"
        doc_path = os.path.join(temp_dir, filename)
        doc.save(doc_path)
        
        return doc_path
    except Exception as e:
        logger.error(f"Error creating Word document: {str(e)}")
        raise

# ============ SLACK HANDLERS ============

@app.command("/contact-report")
def handle_contact_report_command(ack, body, say):
    """Handle /contact-report slash command"""
    ack()
    
    channel = body["channel_id"]
    channel_name = body.get("channel_name", "general")
    user_id = body["user_id"]
    
    # Create workflow
    workflow = ContactReportWorkflow(channel, channel_name)
    workflow_key = f"{user_id}_{channel}"
    active_workflows[workflow_key] = workflow
    
    # Start the conversation
    say(f"📋 Starting Contact Report for: *{workflow.client_name}*\n\n"
        f"Step 1/5: What's the project name?")

     @app.message(re=r".*")
    
    # Skip bot messages
    if message.get("bot_id"):
        return
    
    # Handle file uploads
    if message.get("subtype") == "file_share":
        handle_file_upload(message, say, logger)
        return
    
    # Handle text responses in workflow
    user_id = message.get("user")
    channel = message.get("channel")
    text = message.get("text", "").strip()
    
    workflow_key = f"{user_id}_{channel}"
    if workflow_key not in active_workflows:
        return
    
    workflow = active_workflows[workflow_key]
    
    # State machine for workflow
    if workflow.state == "awaiting_project_name":
        workflow.project_name = text
        workflow.state = "awaiting_attendees"
        say("Step 2/5: Who attended the meeting? (names, comma-separated)")
    
    elif workflow.state == "awaiting_attendees":
        workflow.attendees = [name.strip() for name in text.split(",")]
        workflow.state = "awaiting_am"
        
        # Show AM options
        am_list = "\n".join([f"• {am}" for am in TEAM_DATA["account_managers"]])
        say(f"Step 3/5: Who's the Project AM?\n{am_list}")
    
    elif workflow.state == "awaiting_am":
        workflow.project_am = text
        workflow.state = "awaiting_oversight"
        
        # Show senior oversight options
        oversight_list = "\n".join([f"• {so}" for so in TEAM_DATA["senior_oversight"]])
        say(f"Step 4/5: Who's Senior Oversight?\n{oversight_list}")
    
    elif workflow.state == "awaiting_oversight":
        workflow.senior_oversight = text
        workflow.state = "awaiting_context"
        say("Step 5/5: Any other context? (e.g., budget, timeline, key concerns)\n"
            "Type 'none' or 'no' if there's nothing to add.")
    
    elif workflow.state == "awaiting_context":
        workflow.context = text if text.lower() not in ["none", "no", "nope"] else ""
        workflow.state = "awaiting_file"
        say("✅ Got it! Now upload the meeting file:\n"
            "📁 MP3 / MP4 (I'll transcribe it)\n"
            "📄 DOCX (I'll extract the text)\n\n"
            "Just drag & drop or attach the file.")

def handle_file_upload(message, say, logger):
    """Handle file uploads for active workflows"""
    try:
        files = message.get("files", [])
        if not files:
            return
        
        file_info = files[0]
        file_id = file_info.get("id")
        file_name = file_info.get("name")
        
        user_id = message.get("user")
        channel = message.get("channel")
        workflow_key = f"{user_id}_{channel}"
        
        if workflow_key not in active_workflows:
            return
        
        workflow = active_workflows[workflow_key]
        
        if workflow.state != "awaiting_file":
            return
        
        token = os.environ.get("SLACK_BOT_TOKEN")
        say(f"📥 Received: {file_name}\n🤖 Processing...")
        
        # Download file
        file_path = download_slack_file(file_id, file_name, token)
        
        # Extract content based on file type
        if file_name.lower().endswith(('.mp3', '.mp4', '.wav', '.m4a')):
            say("🎙️ Transcribing audio (this may take a moment)...")
            content = transcribe_audio(file_path)
        elif file_name.lower().endswith('.docx'):
            say("📄 Extracting text...")
            content = extract_text_from_docx(file_path)
        else:
            say("❌ Unsupported file type. Please upload MP3, MP4, or DOCX.")
            return
        
        say("✍️ Generating report with Claude...")
        
        # Generate report
        report_data = generate_report_with_claude(content, workflow)
        
        # Create Word document
        say("📝 Creating document...")
        doc_path = create_word_document(workflow, report_data)
        
        # Upload to Drive
        say("☁️ Uploading to Google Drive...")
        drive_link = upload_to_google_drive(doc_path, workflow)
        
        # Success message
        say(f"✅ Contact Report Generated\n\n"
            f"📋 {workflow.attendees[0] if workflow.attendees else 'Meeting'} | {workflow.client_name}\n"
            f"📅 {workflow.project_name}\n\n"
            f"<{drive_link}|📁 View Report in Google Drive>")
        
        # Clean up
        del active_workflows[workflow_key]
        if os.path.exists(file_path):
            os.remove(file_path)
        if os.path.exists(doc_path):
            os.remove(doc_path)
    
    except Exception as e:
        logger.error(f"Error processing file: {str(e)}")
        say(f"❌ Error: {str(e)}")

# ============ FLASK ROUTES ============

@flask_app.route("/slack/events", methods=["POST"])
def slack_events():
    return handler.handle(request)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3000))
    flask_app.run(host="0.0.0.0", port=port, debug=False)
