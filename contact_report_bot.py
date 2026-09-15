"""
Claude Contact Reports Bot for Slack
Generates professional contact reports and posts them directly to Slack
Supports: DOCX, PDF, TXT files
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
from docx import Document
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize apps
app = App(
    token=os.environ.get("SLACK_BOT_TOKEN"),
    signing_secret=os.environ.get("SLACK_SIGNING_SECRET")
)

flask_app = Flask(__name__)
handler = SlackRequestHandler(app)
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

active_workflows = {}

class ContactReportWorkflow:
    """Manages workflow state for a contact report"""
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


def download_slack_file(file_id, file_name, token):
    """Download file from Slack"""
    try:
        url = f"https://slack.com/api/files.info?file={file_id}"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        file_info = response.json()
        if not file_info.get("ok"):
            raise Exception(f"Slack API error: {file_info.get('error')}")
        
        download_url = file_info.get("file", {}).get("url_private")
        if not download_url:
            raise Exception("Could not find download URL for file")
        
        file_response = requests.get(download_url, headers=headers)
        file_response.raise_for_status()
        
        temp_dir = tempfile.gettempdir()
        file_path = os.path.join(temp_dir, file_name)
        
        with open(file_path, "wb") as f:
            f.write(file_response.content)
        
        return file_path
    except Exception as e:
        logger.error(f"Download error: {str(e)}")
        raise


def extract_text_from_docx(file_path):
    """Extract text from DOCX/DOC file"""
    try:
        doc = Document(file_path)
        text = ""
        for para in doc.paragraphs:
            if para.text.strip():
                text += para.text + "\n"
        return text.strip() if text else "[No text found in document]"
    except Exception as e:
        logger.error(f"DOCX extraction error: {str(e)}")
        raise Exception(f"Failed to extract from DOCX: {str(e)}")


def extract_text_from_pdf(file_path):
    """Extract text from PDF file"""
    try:
        import PyPDF2
        text = ""
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        return text.strip() if text else "[No text found in PDF]"
    except ImportError:
        raise Exception("PyPDF2 library not installed")
    except Exception as e:
        logger.error(f"PDF extraction error: {str(e)}")
        raise Exception(f"Failed to extract from PDF: {str(e)}")


def extract_text_from_txt(file_path):
    """Extract text from plain text file"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()
        return text.strip() if text else "[No text found in file]"
    except Exception as e:
        logger.error(f"Text extraction error: {str(e)}")
        raise Exception(f"Failed to read text file: {str(e)}")


def generate_report_with_claude(content, workflow):
    """Use Claude to generate meeting report"""
    try:
        prompt = f"""You are a professional business analyst. Extract the following from this meeting content:

1. Background: Client's current situation, challenges, and context
2. The Ask: What they're requesting, timeline, budget, decision process if mentioned
3. Actions: Next steps and action items (list as bullet points)
4. Key Points: Important discussion topics (list as bullet points)

Meeting Content:
{content}

Respond ONLY with valid JSON in this format:
{{"background": "...", "the_ask": "...", "actions": ["action1", "action2"], "key_points": ["point1", "point2"]}}"""

        message = anthropic_client.messages.create(
            model="claude-sonnet-5",
            max_tokens=4000,
            messages=[
                {
                    "role": "user",
                    "content": prompt
                }
            ]
        )
        
        if not message.content:
            raise Exception("Empty response from Claude")
        
        # Find the text block (skip thinking blocks)
        response_text = None
        for block in message.content:
            if hasattr(block, 'type') and block.type == 'text' and hasattr(block, 'text'):
                response_text = block.text
                break
        
        if not response_text:
            raise Exception("No text block found in Claude response")
        
        # Parse JSON from response
        json_start = response_text.find('{')
        json_end = response_text.rfind('}') + 1
        
        if json_start < 0 or json_end <= json_start:
            raise Exception("Could not find JSON in response")
        
        json_str = response_text[json_start:json_end]
        return json.loads(json_str)
        
    except Exception as e:
        logger.error(f"Claude error: {str(e)}")
        raise


def create_word_document(workflow, report_data):
    """Create branded Word document matching template layout"""
    try:
        doc = Document()
        
        # Add some spacing
        doc.add_paragraph()
        doc.add_paragraph()
        
        # TABLE 1: Header info (2 rows, 3 cols)
        header_table = doc.add_table(rows=2, cols=3)
        header_table.style = 'Table Grid'
        
        # Row 1: Client, Project Name, Meeting
        row1_cells = header_table.rows[0].cells
        row1_cells[0].text = f"Client:\n{workflow.client_name}"
        row1_cells[1].text = f"Project Name:\n{workflow.project_name or 'N/A'}"
        meeting_date = datetime.now().strftime("%d/%m/%Y")
        attendees_str = ", ".join(workflow.attendees) if workflow.attendees else "N/A"
        row1_cells[2].text = f"Meeting:\n[{meeting_date} & {attendees_str}]"
        
        # Row 2: Project AM, Senior Oversight, Meeting
        row2_cells = header_table.rows[1].cells
        row2_cells[0].text = f"Project AM:\n{workflow.project_am or 'N/A'}"
        row2_cells[1].text = f"Senior Oversight:\n{workflow.senior_oversight or 'N/A'}"
        row2_cells[2].text = f"Meeting:\n[Date & Attendees]"
        
        doc.add_paragraph()
        
        # TABLE 2: Meeting Notes & Background & The Ask (2 rows, 1 col)
        notes_table = doc.add_table(rows=2, cols=1)
        notes_table.style = 'Table Grid'
        
        notes_table.rows[0].cells[0].text = "Meeting notes"
        
        # Build content for row 2
        bg = report_data.get("background", "")
        ask = report_data.get("the_ask", "")
        notes_content = f"Background\n\n{bg}\n\nThe Ask\n\n{ask}"
        notes_table.rows[1].cells[0].text = notes_content
        
        doc.add_paragraph()
        
        # TABLE 3: Actions (2 rows, 1 col)
        actions_table = doc.add_table(rows=2, cols=1)
        actions_table.style = 'Table Grid'
        
        actions_table.rows[0].cells[0].text = "Actions"
        
        # Build actions content
        actions = report_data.get("actions", [])
        if actions and isinstance(actions, list):
            actions_content = "\n".join([f"• {str(action)}" for action in actions])
        else:
            actions_content = "• Action items to be determined"
        actions_table.rows[1].cells[0].text = actions_content
        
        doc.add_paragraph()
        
        # TABLE 4: Key Points (2 rows, 1 col)
        key_points_table = doc.add_table(rows=2, cols=1)
        key_points_table.style = 'Table Grid'
        
        key_points_table.rows[0].cells[0].text = "Key Points"
        
        # Build key points content
        key_points = report_data.get("key_points", [])
        if key_points and isinstance(key_points, list):
            key_points_content = "\n".join([f"• {str(point)}" for point in key_points])
        else:
            key_points_content = "• Key discussion points"
        key_points_table.rows[1].cells[0].text = key_points_content
        
        doc.add_paragraph()
        
        # TABLE 5: AOB (2 rows, 1 col)
        aob_table = doc.add_table(rows=2, cols=1)
        aob_table.style = 'Table Grid'
        
        aob_table.rows[0].cells[0].text = "AOB"
        aob_table.rows[1].cells[0].text = "[Additional notes or follow-up items]"
        
        # Save document
        temp_dir = tempfile.gettempdir()
        filename = f"{datetime.now().strftime('%Y-%m-%d')}_ContactReport_{workflow.client_name.replace(' ', '_')}.docx"
        doc_path = os.path.join(temp_dir, filename)
        doc.save(doc_path)
        
        return doc_path, filename
    except Exception as e:
        logger.error(f"Document creation error: {str(e)}")
        raise


@app.command("/contact-report")
def handle_contact_report_command(ack, body, say):
    """Handle /contact-report slash command"""
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
    """Handle all message events"""
    event = body.get("event", {})
    
    # Skip bot messages
    if event.get("bot_id"):
        return
    
    user_id = event.get("user")
    channel = event.get("channel")
    
    workflow_key = f"{user_id}_{channel}"
    if workflow_key not in active_workflows:
        return
    
    workflow = active_workflows[workflow_key]
    
    # Handle file uploads
    if event.get("subtype") == "file_share":
        try:
            files = event.get("files", [])
            if not files:
                return
            
            file_info = files[0]
            file_id = file_info.get("id")
            file_name = file_info.get("name")
            
            # Only process if waiting for file
            if workflow.state != "awaiting_file":
                return
            
            token = os.environ.get("SLACK_BOT_TOKEN")
            say(f"📥 Received: {file_name}\n🤖 Processing...")
            
            # Download file
            file_path = download_slack_file(file_id, file_name, token)
            
            # Extract text based on file type
            file_lower = file_name.lower()
            if file_lower.endswith(('.docx', '.doc')):
                say("📄 Extracting text from document...")
                content = extract_text_from_docx(file_path)
            elif file_lower.endswith('.pdf'):
                say("📄 Extracting text from PDF...")
                content = extract_text_from_pdf(file_path)
            elif file_lower.endswith('.txt'):
                say("📄 Reading text file...")
                content = extract_text_from_txt(file_path)
            else:
                say("❌ Unsupported file type. Please upload DOCX, PDF, or TXT.")
                return
            
            say("✍️ Generating report with Claude...")
            report_data = generate_report_with_claude(content, workflow)
            
            say("📝 Creating Word document...")
            doc_path, filename = create_word_document(workflow, report_data)
            
            # Upload to Slack
            say("📤 Uploading report to Slack...")
            with open(doc_path, 'rb') as f:
                app.client.files_upload_v2(
                    channel=channel,
                    file=f,
                    filename=filename,
                    title=f"Contact Report: {workflow.client_name}",
                    initial_comment=f"✅ Contact Report Generated\n\n📋 {workflow.attendees[0] if workflow.attendees else 'Meeting'} | {workflow.client_name}\n📅 {workflow.project_name}"
                )
            
            # Cleanup
            del active_workflows[workflow_key]
            if os.path.exists(file_path):
                os.remove(file_path)
            if os.path.exists(doc_path):
                os.remove(doc_path)
        
        except Exception as e:
            logger.error(f"File processing error: {str(e)}")
            say(f"❌ Error: {str(e)}")
        
        return
    
    # Handle text responses
    text = event.get("text", "").strip()
    if not text:
        return
    
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
        say("Step 5/5: Any other context? (budget, timeline, concerns)\nType 'none' if nothing to add.")
    
    elif workflow.state == "awaiting_context":
        workflow.context = text if text.lower() not in ["none", "no", "nope"] else ""
        workflow.state = "awaiting_file"
        say("✅ Got it! Now upload the meeting file:\n📄 DOCX, PDF, or TXT\n\nJust drag & drop or attach.")


@flask_app.route("/slack/events", methods=["POST"])
def slack_events():
    return handler.handle(request)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3000))
    flask_app.run(host="0.0.0.0", port=port, debug=False)
