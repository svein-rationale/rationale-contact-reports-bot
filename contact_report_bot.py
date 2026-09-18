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

# Embedded Rationale logo as base64
RATIONALE_LOGO_BASE64 = """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"""

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
    """Create branded Word document with embedded logo and improved formatting"""
    import base64
    from docx.shared import Pt
    try:
        doc = Document()
        
        # Add embedded logo
        try:
            logo_bytes = base64.b64decode(RATIONALE_LOGO_BASE64)
            temp_logo = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
            temp_logo.write(logo_bytes)
            temp_logo.close()
            
            logo_paragraph = doc.add_paragraph()
            logo_run = logo_paragraph.add_run()
            logo_run.add_picture(temp_logo.name, width=1500000)  # 1.5 inches
            logo_paragraph.alignment = 0  # Left align
            logo_paragraph.paragraph_format.space_after = Pt(12)
            
            # Cleanup temp file
            os.unlink(temp_logo.name)
            logger.info("Logo added successfully")
        except Exception as e:
            logger.warning(f"Could not add logo: {str(e)}")
        
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
        
        # Build formatted content for Background and The Ask
        bg = report_data.get("background", "")
        ask = report_data.get("the_ask", "")
        
        # Build the content with better spacing
        notes_cell = notes_table.rows[1].cells[0]
        notes_cell.text = ""  # Clear default text
        
        # Add Background
        bg_para = notes_cell.paragraphs[0]
        bg_run = bg_para.add_run("Background")
        bg_run.bold = True
        bg_para.paragraph_format.space_after = Pt(6)
        
        # Add background text with proper breaks and spacing
        for line in bg.split('. '):
            if line.strip():
                para = notes_cell.add_paragraph(line.strip() + '.', style='Normal')
                para.paragraph_format.space_after = Pt(6)
        
        # Add spacing before The Ask
        notes_cell.add_paragraph()
        
        # Add The Ask
        ask_para = notes_cell.add_paragraph()
        ask_run = ask_para.add_run("The Ask")
        ask_run.bold = True
        ask_para.paragraph_format.space_after = Pt(6)
        
        # Add ask text with proper breaks and spacing
        for line in ask.split('. '):
            if line.strip():
                para = notes_cell.add_paragraph(line.strip() + '.', style='Normal')
                para.paragraph_format.space_after = Pt(6)
        
        doc.add_paragraph()
        
        # TABLE 3: Actions (2 rows, 1 col)
        actions_table = doc.add_table(rows=2, cols=1)
        actions_table.style = 'Table Grid'
        
        actions_table.rows[0].cells[0].text = "Actions"
        
        # Build actions content
        actions_cell = actions_table.rows[1].cells[0]
        actions_cell.text = ""  # Clear default text
        
        actions = report_data.get("actions", [])
        if actions and isinstance(actions, list):
            for action in actions:
                action_text = str(action).strip()
                if action_text:
                    para = actions_cell.add_paragraph(action_text, style='List Bullet')
                    para.paragraph_format.space_after = Pt(8)
        else:
            para = actions_cell.add_paragraph("Action items to be determined", style='List Bullet')
            para.paragraph_format.space_after = Pt(8)
        
        doc.add_paragraph()
        
        # TABLE 4: Key Points (2 rows, 1 col)
        key_points_table = doc.add_table(rows=2, cols=1)
        key_points_table.style = 'Table Grid'
        
        key_points_table.rows[0].cells[0].text = "Key Points"
        
        # Build key points content
        key_points_cell = key_points_table.rows[1].cells[0]
        key_points_cell.text = ""  # Clear default text
        
        key_points = report_data.get("key_points", [])
        if key_points and isinstance(key_points, list):
            for point in key_points:
                point_text = str(point).strip()
                if point_text:
                    para = key_points_cell.add_paragraph(point_text, style='List Bullet')
                    para.paragraph_format.space_after = Pt(8)
        else:
            para = key_points_cell.add_paragraph("Key discussion points", style='List Bullet')
            para.paragraph_format.space_after = Pt(8)
        
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
        
        # Build formatted content for Background and The Ask
        bg = report_data.get("background", "")
        ask = report_data.get("the_ask", "")
        
        # Build the content with better spacing
        notes_cell = notes_table.rows[1].cells[0]
        notes_cell.text = ""  # Clear default text
        
        # Add Background
        bg_para = notes_cell.paragraphs[0]
        bg_run = bg_para.add_run("Background")
        bg_run.bold = True
        bg_para.paragraph_format.space_after = Pt(6)
        
        # Add background text with proper breaks and spacing
        for line in bg.split('. '):
            if line.strip():
                para = notes_cell.add_paragraph(line.strip() + '.', style='Normal')
                para.paragraph_format.space_after = Pt(6)
        
        # Add spacing before The Ask
        notes_cell.add_paragraph()
        
        # Add The Ask
        ask_para = notes_cell.add_paragraph()
        ask_run = ask_para.add_run("The Ask")
        ask_run.bold = True
        ask_para.paragraph_format.space_after = Pt(6)
        
        # Add ask text with proper breaks and spacing
        for line in ask.split('. '):
            if line.strip():
                para = notes_cell.add_paragraph(line.strip() + '.', style='Normal')
                para.paragraph_format.space_after = Pt(6)
        
        doc.add_paragraph()
        
        # TABLE 3: Actions (2 rows, 1 col)
        actions_table = doc.add_table(rows=2, cols=1)
        actions_table.style = 'Table Grid'
        
        actions_table.rows[0].cells[0].text = "Actions"
        
        # Build actions content
        actions_cell = actions_table.rows[1].cells[0]
        actions_cell.text = ""  # Clear default text
        
        actions = report_data.get("actions", [])
        if actions and isinstance(actions, list):
            for action in actions:
                action_text = str(action).strip()
                if action_text:
                    para = actions_cell.add_paragraph(action_text, style='List Bullet')
                    para.paragraph_format.space_after = Pt(8)
        else:
            para = actions_cell.add_paragraph("Action items to be determined", style='List Bullet')
            para.paragraph_format.space_after = Pt(8)
        
        doc.add_paragraph()
        
        # TABLE 4: Key Points (2 rows, 1 col)
        key_points_table = doc.add_table(rows=2, cols=1)
        key_points_table.style = 'Table Grid'
        
        key_points_table.rows[0].cells[0].text = "Key Points"
        
        # Build key points content
        key_points_cell = key_points_table.rows[1].cells[0]
        key_points_cell.text = ""  # Clear default text
        
        key_points = report_data.get("key_points", [])
        if key_points and isinstance(key_points, list):
            for point in key_points:
                point_text = str(point).strip()
                if point_text:
                    para = key_points_cell.add_paragraph(point_text, style='List Bullet')
                    para.paragraph_format.space_after = Pt(8)
        else:
            para = key_points_cell.add_paragraph("Key discussion points", style='List Bullet')
            para.paragraph_format.space_after = Pt(8)
        
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
