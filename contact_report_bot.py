"""
Rationale Contact Report Bot
Generates branded contact reports from meeting recordings/transcripts
"""

import os
import json
from datetime import datetime
from slack_bolt import App
from slack_bolt.adapter.flask import SlackRequestHandler
from flask import Flask, request
import anthropic
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import mimetypes

# ============================================================================
# SETUP: Initialize Flask, Slack, Claude, and Google Drive
# ============================================================================

# Flask app for deployment
flask_app = Flask(__name__)

# Initialize Slack app
slack_app = App(
    token=os.environ["SLACK_BOT_TOKEN"],
    signing_secret=os.environ["SLACK_SIGNING_SECRET"]
)

# Initialize Claude client
claude_client = anthropic.Anthropic(api_key=os.environ["CLAUDE_API_KEY"])

# Initialize Google Drive client
def get_google_drive_service():
    """Authenticate with Google Drive using service account credentials"""
    creds = Credentials.from_service_account_file(
        'google_service_account.json',
        scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds)

# ============================================================================
# TEAM DATA: Names and roles for dropdowns
# ============================================================================

TEAM_DATA = {
    "account_managers": [
        "Svein Clouston (Co-MD)",
        "Rowan Morrison (Co-MD)",
        "Ursula Fairlie (Senior Account Director)",
        "Gemma Greig (Account Director)",
        "Matthew Collins (Account Manager)",
        "Carron Easdon (Senior Account Manager)",
    ],
    "senior_oversight": [
        "Svein Clouston (Co-MD)",
        "Rowan Morrison (Co-MD)",
        "Helen Davidson (Strategy Director)",
        "Ursula Fairlie (Senior Account Director)",
    ],
    "client_mapping": {
        # Maps Slack channel names to formal client names
        "hargreaves-lansdown": "Hargreaves Lansdown",
        "hl": "Hargreaves Lansdown",
        "canaccord": "Canaccord",
        "bupa": "Bupa",
        "mddus": "MDDUS",
        "mattioli-woods": "Mattioli Woods",
        "innovate-uk": "Innovate UK",
        # Add more as needed
    }
}

# ============================================================================
# CONVERSATION STATE: Track multi-step workflows per user
# ============================================================================

user_workflows = {}

class ContactReportWorkflow:
    """Tracks the workflow state for generating a contact report"""
    
    def __init__(self, user_id, channel_id, channel_name):
        self.user_id = user_id
        self.channel_id = channel_id
        self.channel_name = channel_name
        self.client_name = self._extract_client_name()
        self.data = {
            "client": self.client_name,
            "project_name": None,
            "attendees": None,
            "project_am": None,
            "senior_oversight": None,
            "context": None,
            "file_path": None,
            "file_type": None,
        }
        self.step = "project_name"  # Current step in workflow
    
    def _extract_client_name(self):
        """Extract client name from channel name"""
        channel_clean = self.channel_name.lower().replace("#", "")
        return TEAM_DATA["client_mapping"].get(channel_clean, channel_clean.replace("-", " ").title())

# ============================================================================
# SLACK EVENT HANDLERS
# ============================================================================

@slack_app.command("/contact-report")
def handle_contact_report_command(ack, body, say):
    """
    Slash command: /contact-report
    Initiates the contact report workflow
    """
    ack()
    
    user_id = body["user_id"]
    channel_id = body["channel_id"]
    channel_name = body["channel_name"]
    
    # Create new workflow for this user
    workflow_id = f"{user_id}_{channel_id}"
    user_workflows[workflow_id] = ContactReportWorkflow(user_id, channel_id, channel_name)
    
    workflow = user_workflows[workflow_id]
    
    # Start the workflow: ask for project name
    say(
    f"📋 Starting Contact Report for: *{workflow.client_name}*\n\n"
    f"Step 1/5: What's the project name?"
)

@slack_app.message()
def handle_messages(message, say):
    """
    Handle regular messages (responses to workflow questions)
    This processes each step of the contact report workflow
    """
    
    # Ignore bot messages
    if message.get("bot_id"):
        return
    
    user_id = message["user"]
    channel_id = message["channel"]
    workflow_id = f"{user_id}_{channel_id}"
    
    # Check if user has an active workflow
    if workflow_id not in user_workflows:
        return
    
    workflow = user_workflows[workflow_id]
    user_input = message.get("text", "").strip()
    
    # ========================================================================
    # STEP 1: Get Project Name
    # ========================================================================
    if workflow.step == "project_name":
        workflow.data["project_name"] = user_input
        workflow.step = "attendees"
        say(
            f"✅ Project: *{user_input}*\n\n"
            f"Step 2/5: Who attended the meeting? (names, comma-separated)"
        )
    
    # ========================================================================
    # STEP 2: Get Attendees
    # ========================================================================
    elif workflow.step == "attendees":
        workflow.data["attendees"] = user_input
        workflow.step = "project_am"
        
        # Show formatted list of AMs for selection
        am_list = "\n".join([f"• {am}" for am in TEAM_DATA["account_managers"]])
        say(
            f"✅ Attendees: *{user_input}*\n\n"
            f"Step 3/5: Who's the Project AM?\n{am_list}"
        )
    
    # ========================================================================
    # STEP 3: Get Project AM
    # ========================================================================
    elif workflow.step == "project_am":
        workflow.data["project_am"] = user_input
        workflow.step = "senior_oversight"
        
        # Show formatted list of oversight options
        oversight_list = "\n".join([f"• {name}" for name in TEAM_DATA["senior_oversight"]])
        say(
            f"✅ Project AM: *{user_input}*\n\n"
            f"Step 4/5: Who's Senior Oversight?\n{oversight_list}"
        )
    
    # ========================================================================
    # STEP 4: Get Senior Oversight
    # ========================================================================
    elif workflow.step == "senior_oversight":
        workflow.data["senior_oversight"] = user_input
        workflow.step = "context"
        say(
            f"✅ Senior Oversight: *{user_input}*\n\n"
            f"Step 5/5: Is there any other context we should know about this contact or meeting? "
            f"(e.g., previous conversations, budget parameters, decision timeline, known sensitivities, etc.)"
        )
    
    # ========================================================================
    # STEP 5: Get Additional Context
    # ========================================================================
    elif workflow.step == "context":
        workflow.data["context"] = user_input if user_input.lower() not in ["no", "none", "skip"] else None
        workflow.step = "file_upload"
        say(
            f"✅ Context noted.\n\n"
            f"📁 Now upload the meeting file:\n"
            f"• MP3/MP4 (will be transcribed)\n"
            f"• DOCX (meeting notes/transcript)\n\n"
            f"Just drag & drop or attach the file and I'll process it."
        )

@slack_app.event("file_shared")
def handle_file_upload(event, say):
    """
    Handle file uploads (MP3, MP4, DOCX)
    Processes the file and triggers contact report generation
    """
    user_id = event["user_id"]
    channel_id = event["channel_id"]
    file_id = event["file_id"]
    workflow_id = f"{user_id}_{channel_id}"
    
    if workflow_id not in user_workflows:
        say("❌ Error: No active workflow. Please run `/contact-report` first.")
        return
    
    workflow = user_workflows[workflow_id]
    
    # Get file info from Slack
    client = slack_app.client
    file_info = client.files_info(file=file_id)["file"]
    
    file_name = file_info["name"]
    file_type = file_info["mimetype"]
    download_url = file_info["url_private"]
    
    # Download file from Slack
    say(f"📥 Received: *{file_name}*\nProcessing...")
    
    try:
        # Download the file
        headers = {"Authorization": f"Bearer {os.environ['SLACK_BOT_TOKEN']}"}
        import requests
        response = requests.get(download_url, headers=headers)
        
        temp_file_path = f"/tmp/{file_id}_{file_name}"
        with open(temp_file_path, "wb") as f:
            f.write(response.content)
        
        workflow.data["file_path"] = temp_file_path
        workflow.data["file_type"] = file_type
        
        # Process file and generate report
        generate_contact_report(workflow, say)
        
        # Clean up workflow
        del user_workflows[workflow_id]
        
    except Exception as e:
        say(f"❌ Error processing file: {str(e)}")

# ============================================================================
# CONTACT REPORT GENERATION
# ============================================================================

def generate_contact_report(workflow, say):
    """
    Main function to generate the contact report
    1. Transcribe/read the file
    2. Extract meeting info using Claude
    3. Generate Word document
    4. Save to Google Drive
    5. Post Slack notification with link
    """
    
    file_path = workflow.data["file_path"]
    file_type = workflow.data["file_type"]
    
    say("🤖 Analyzing meeting content with Claude...")
    
    # ========================================================================
    # Step 1: Process File (Transcribe Audio or Read DOCX)
    # ========================================================================
    if file_type in ["audio/mpeg", "audio/mp4", "video/mp4"]:
        meeting_content = transcribe_audio(file_path)
    elif file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        meeting_content = extract_docx_text(file_path)
    else:
        say(f"❌ Unsupported file type: {file_type}")
        return
    
    # ========================================================================
    # Step 2: Extract Meeting Info Using Claude
    # ========================================================================
    extracted_info = extract_meeting_info(meeting_content, workflow)
    
    # ========================================================================
    # Step 3: Generate Word Document
    # ========================================================================
    say("📄 Generating report document...")
    word_doc_path = generate_word_document(workflow, extracted_info)
    
    # ========================================================================
    # Step 4: Save to Google Drive
    # ========================================================================
    say("☁️ Saving to Google Drive...")
    drive_link = upload_to_google_drive(word_doc_path, workflow)
    
    # ========================================================================
    # Step 5: Post Slack Notification
    # ========================================================================
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    contact_name = extracted_info.get("contact_name", "Unknown Contact")
    
    notification = (
        f"✅ *Contact Report Generated*\n\n"
        f"📋 *{contact_name}* | {workflow.client_name}\n"
        f"📅 {timestamp}\n\n"
        f"📁 <{drive_link}|View Report in Google Drive>"
    )
    
    say(notification)

def transcribe_audio(file_path):
    """
    Transcribe audio file (MP3/MP4) using Google Cloud Speech-to-Text
    """
    from google.cloud import speech
    import io
    
    # Initialize Speech-to-Text client
    client = speech.SpeechClient()
    
    # Read audio file
    with io.open(file_path, "rb") as audio_file:
        content = audio_file.read()
    
    # Prepare audio config
    audio = speech.RecognitionAudio(content=content)
    config = speech.RecognitionConfig(
        encoding=speech.RecognitionConfig.AudioEncoding.MP3,
        sample_rate_hertz=16000,
        language_code="en-US",
    )
    
    # Transcribe
    response = client.recognize(config=config, audio=audio)
    
    # Extract transcript
    transcript = ""
    for result in response.results:
        transcript += result.alternatives[0].transcript + " "
    
    return transcript.strip() if transcript else "No speech detected in audio"

def extract_docx_text(file_path):
    """Extract text from DOCX file"""
    from docx import Document
    doc = Document(file_path)
    text = "\n".join([para.text for para in doc.paragraphs])
    return text

def extract_meeting_info(meeting_content, workflow):
    """
    Use Claude to extract key information from meeting transcript/notes
    Returns structured data for the contact report
    """
    
    prompt = f"""
    You are analyzing meeting notes from a client meeting at Rationale.
    
    Client: {workflow.client_name}
    Project: {workflow.data['project_name']}
    Attendees: {workflow.data['attendees']}
    Additional Context: {workflow.data['context'] or 'None provided'}
    
    MEETING CONTENT:
    {meeting_content}
    
    Please extract and structure the following information:
    
    1. CONTACT_NAME: The primary contact's name from the meeting
    2. BACKGROUND: A clear summary of:
       - Client's current situation
       - What they're trying to solve
       - Key challenges mentioned
       - Relevant history with Rationale
    3. THE_ASK: What they're specifically requesting:
       - Scope of work
       - Timeline
       - Budget indication (if mentioned)
       - Decision process
       - Next steps
    4. ACTIONS: Key action items from the meeting (bullet points)
    5. AOB: Any other business/notes (bullet points)
    
    Format your response as JSON with these exact keys:
    {{
        "contact_name": "string",
        "background": "string (2-3 paragraphs)",
        "the_ask": "string (2-3 paragraphs)",
        "actions": ["action 1", "action 2"],
        "aob": ["note 1", "note 2"]
    }}
    
    Be thorough, professional, and focus on actionable information.
    """
    
    message = claude_client.messages.create(
        model="claude-sonnet-5",
        max_tokens=2000,
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ]
    )
    
    response_text = message.content[0].text
    
    # Extract JSON from response
    try:
        # Try to find JSON in the response
        import re
        json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
        if json_match:
            extracted_info = json.loads(json_match.group())
        else:
            # Fallback if no JSON found
            extracted_info = {
                "contact_name": "Unknown",
                "background": "See meeting notes",
                "the_ask": "See meeting notes",
                "actions": [],
                "aob": []
            }
    except json.JSONDecodeError:
        extracted_info = {
            "contact_name": "Unknown",
            "background": response_text,
            "the_ask": response_text,
            "actions": [],
            "aob": []
        }
    
    return extracted_info

def generate_word_document(workflow, extracted_info):
    """
    Generate a branded Word document matching Rationale's contact report template
    """
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    doc = Document()
    
    # ========================================================================
    # HEADER: Add Rationale Logo and Branding
    # ========================================================================
    # Add logo (you'll need to provide the logo file)
    try:
        logo_path = "rationale_logo.png"
        doc.add_picture(logo_path, width=Inches(2))
    except:
        # Fallback: just add text logo
        title = doc.add_paragraph()
        title_run = title.add_run("Rationale")
        title_run.font.size = Pt(36)
        title_run.font.bold = True
        title_run.font.color.rgb = RGBColor(255, 0, 102)  # Rationale pink
    
    doc.add_paragraph()  # Spacing
    
    # ========================================================================
    # INFO TABLE: Client, Project, Meeting Details
    # ========================================================================
    table = doc.add_table(rows=2, cols=3)
    table.style = 'Light Grid Accent 1'
    
    # Header row
    cells = table.rows[0].cells
    cells[0].text = "Client:"
    cells[1].text = "Project Name:"
    cells[2].text = "Meeting:"
    
    # Data row
    cells = table.rows[1].cells
    cells[0].text = workflow.client_name
    cells[1].text = workflow.data["project_name"]
    cells[2].text = f"[{extracted_info.get('contact_name', 'Contact')}]\n[Date & Attendees: {workflow.data['attendees']}]"
    
    # Add AM/Oversight info
    table2 = doc.add_table(rows=1, cols=2)
    cells = table2.rows[0].cells
    cells[0].text = f"Project AM: {workflow.data['project_am']}"
    cells[1].text = f"Senior Oversight: {workflow.data['senior_oversight']}"
    
    doc.add_paragraph()  # Spacing
    
    # ========================================================================
    # MEETING NOTES SECTION
    # ========================================================================
    heading = doc.add_heading("Meeting notes", level=2)
    
    # Background
    doc.add_heading("Background", level=3)
    doc.add_paragraph(extracted_info.get("background", ""))
    
    # The Ask
    doc.add_heading("The Ask", level=3)
    doc.add_paragraph(extracted_info.get("the_ask", ""))
    
    # ========================================================================
    # ACTIONS SECTION
    # ========================================================================
    doc.add_heading("Actions", level=2)
    for action in extracted_info.get("actions", []):
        doc.add_paragraph(action, style='List Bullet')
    
    doc.add_paragraph()  # Spacing
    
    # ========================================================================
    # AOB SECTION
    # ========================================================================
    doc.add_heading("AOB", level=2)
    for note in extracted_info.get("aob", []):
        doc.add_paragraph(note, style='List Bullet')
    
    # ========================================================================
    # SAVE DOCUMENT
    # ========================================================================
    timestamp = datetime.now().strftime("%Y-%m-%d")
    contact_name = extracted_info.get("contact_name", "Contact").replace(" ", "_")
    filename = f"{timestamp}_{contact_name}_ContactReport.docx"
    
    filepath = f"/tmp/{filename}"
    doc.save(filepath)
    
    return filepath

def upload_to_google_drive(file_path, workflow):
    """
    Upload the generated Word document to Google Drive
    in the Contact Reports folder
    """
    service = get_google_drive_service()
    
    # Find or create Contact Reports folder
    contact_reports_folder_id = find_or_create_folder(service, "Contact Reports")
    
    # Upload file
    file_metadata = {
        'name': os.path.basename(file_path),
        'parents': [contact_reports_folder_id]
    }
    
    media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
    
    return file.get('webViewLink')

def find_or_create_folder(service, folder_name):
    """
    Find a folder by name in Google Drive, or create it if it doesn't exist
    """
    # Search for existing folder
    results = service.files().list(
        q=f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields='files(id)',
        pageSize=1
    ).execute()
    
    files = results.get('files', [])
    
    if files:
        return files[0]['id']
    else:
        # Create new folder
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        folder = service.files().create(body=file_metadata, fields='id').execute()
        return folder.get('id')

# ============================================================================
# FLASK ROUTES FOR DEPLOYMENT
# ============================================================================

handler = SlackRequestHandler(slack_app)

@flask_app.route("/slack/events", methods=["POST"])
def slack_events():
    return handler.handle(request)

@flask_app.route("/health", methods=["GET"])
def health_check():
    return {"status": "healthy"}, 200

# ============================================================================
# MAIN
# ============================================================================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3000))
    flask_app.run(host="0.0.0.0", port=port, debug=False)
