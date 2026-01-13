import pandas as pd
import os
import smtplib
import time
import json
import random
import requests
from dotenv import load_dotenv
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formataddr
load_dotenv()


# ==================================================
# CONFIGURATION
# ==================================================

# Files
EXCEL_FILE = os.getenv("EXCEL_FILE")
RESUME_FILE = os.getenv("RESUME_FILE")
STATE_FILE = os.getenv("STATE_FILE")
IS_EXCEL_URL = bool(os.getenv("IS_EXCEL_URL"))

# Sender details
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_NAME = os.getenv("SENDER_NAME")
SENDER_PHONE = os.getenv("SENDER_PHONE")
APP_PASSWORD = os.getenv("APP_PASSWORD")

# Email
SUBJECT = os.getenv("SUBJECT")

# Links
LINKEDIN_URL = os.getenv("LINKEDIN_URL")
RESUME_LINK = os.getenv("RESUME_LINK")

# Limits
BATCH_SIZE = int(os.getenv("BATCH_SIZE"))
BATCH_SLEEP = int(os.getenv("BATCH_SLEEP"))
MIN_DELAY = int(os.getenv("MIN_DELAY"))
MAX_DELAY = int(os.getenv("MAX_DELAY"))
DAILY_LIMIT = int(os.getenv("DAILY_LIMIT"))

SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = os.getenv("SMTP_PORT")

LOGGER_FILE = f"log_{time.strftime('%Y%m%d_%H%M%S')}.txt"


# ==================================================
# BOOKMARKING and LOGGING LAST RUN
# ==================================================

def load_state():
    if not Path(STATE_FILE).exists():
        return {"last_row": 0}
    with open(STATE_FILE, "r") as f:
        return json.load(f)

def save_state(state):
    with open(STATE_FILE, "w") as f:
        json.dump(state, f, indent=2)

def logger(message):
    # check if exists and if not create logs directory
    if not Path("./logs").exists():
        os.makedirs("./logs")
    with open(f"./logs/{LOGGER_FILE}", "a") as f:
        f.write(f"{message}\n")

# ==================================================
# SMTP CONNECTION
# ==================================================

def create_smtp_connection():
    server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(SENDER_EMAIL, APP_PASSWORD)
    return server


# ==================================================
# EMAIL CREATION
# ==================================================

def create_message(recipient_name, recipient_email, recipient_company):
    body = f"""Hi {recipient_name},<br><br>

        <p>I am reaching out to express my interest in joining <b>{recipient_company}</b> as a <b>Java Developer / Software Engineer</b>.</p>

        <p>With around 4 years of focused experience in backend development, I am confident that my technical expertise and collaborative approach can contribute effectively to your team’s ongoing projects and innovations.</p>

        <p>In my current role as a <b>Digital Specialist Engineer</b> at <b>Infosys</b>, I have designed and developed high-performing backend solutions using <b>Java, Spring Boot, Neo4j, Elasticsearch, Kafka, Python, .NET Core and Node.js</b>. My work includes building microservices that automate cross-portal data flows, integrating large-scale content from third-party platforms like <i>LinkedIn and Coursera</i>, and implementing event-driven architectures for multi-tenant SaaS product serving clients such as <i>Siemens, FIFA, Kia </i> etc. I also implemented a <b>sentiment analysis</b> of contents in Python using a fine-tuned <i>RoBERTa model</i>, showcasing my ability to blend backend engineering with <b>AI-driven</b> insights.</p> 

        <p>Beyond technical skills, I believe my soft skills would be a great addition for <b>{recipient_company}</b>:</p>

        <ul>
        <li>As a <b>proactive communicator</b>, I've collaborated with cross-functional, international teams spanning diverse tenants such as NPCI, IMF, Sitrain, DeutscheBahn etc. to deliver seamless solutions.</li>
        <li><b>My adaptability</b> enables me to quickly embrace new technologies; when given ownership of modules outside my core stack, I rapidly upskilled and delivered with impact.</li>
        <li>With a <b>solution-oriented mindset</b>, I streamlined onboarding processes by automating tenant setup, reducing manual effort and errors.</li>
        <li>I take <b>end-to-end ownership</b> of every project, ensuring quality and clear communication of dependencies across stakeholders.</li>
        </ul>

        <p>I've attached my resume for your reference. I'd be grateful if you could review it and consider me for any suitable opportunities at <b>{recipient_company}</b>. I would love to contribute to impactful backend systems and continuous improvement initiatives.</p>

        <p>Thank you for your time and consideration. I look forward to the opportunity to discuss how my experience can contribute to <b>{recipient_company}</b>’s engineering goals.</p>

        <br>
        Best regards,<br>
        {SENDER_NAME}<br>
        📞 (+91) {SENDER_PHONE}<br>
        ✉️ <a href="mailto:{SENDER_EMAIL}">{SENDER_EMAIL}</a><br>
        🔗 <a href="{LINKEDIN_URL}">LinkedIn</a> | <a href="{RESUME_LINK}">Resume</a>

    """
    msg = MIMEMultipart()
    msg["From"] = formataddr((SENDER_NAME, SENDER_EMAIL))
    msg["To"] = recipient_email
    msg["Subject"] = SUBJECT

    # Important headers (deliverability)
    msg["Reply-To"] = SENDER_EMAIL
    msg["List-Unsubscribe"] = f"<mailto:{SENDER_EMAIL}?subject=unsubscribe>"
    msg["Precedence"] = "bulk"
    msg["X-Mailer"] = "Python SMTP"

    msg.attach(MIMEText(body, "html"))

    # Dont attach resume to avoid spam filters
    # with open(RESUME_FILE, "rb") as f:
    #     attach = MIMEApplication(f.read(), _subtype="pdf")
    #     attach.add_header(
    #         "Content-Disposition",
    #         "attachment",
    #         filename=RESUME_FILE
    #     )
    #     msg.attach(attach)

    return msg


# ==================================================
# READING EXCEL FILE
# ==================================================

def read_file(file_path, is_url=False):
    if is_url:
        r = requests.head(file_path, allow_redirects=True)
        r.raise_for_status()

        df = pd.read_csv(file_path)
        # df = pd.read_excel(file_path, engine='openpyxl')
    else:
        df = pd.read_excel(file_path)
    return df

# ==================================================
# PUBLIC STATIC MAIN VOID ;p
# ==================================================

def main():
    # Read recipients
    df = read_file(EXCEL_FILE, IS_EXCEL_URL)
    required_cols = {"Name", "Email", "Company"}
    if not required_cols.issubset(df.columns):
        logger(f"Excel must contain columns: {required_cols}")
        raise ValueError(f"Excel must contain columns: {required_cols}")
    
    total = len(df)
    
    state = load_state()
    start_row = state["last_row"]
    
    logger(f"Total recipients: {total}")
    logger(f"Resuming from row: {start_row + 1}")

    sent_today = 0
    smtp_conn = create_smtp_connection()

    try :
        for idx in range(start_row, len(df)):
            if sent_today >= DAILY_LIMIT:
                logger("Daily Email limit reached")
                break
            row = df.iloc[idx]
            name = str(row["Name"]).strip()
            email = str(row["Email"]).strip()
            company = str(row["Company"]).strip()

            try : 
                msg = create_message(name, email, company)
                smtp_conn.send_message(msg)

                sent_today += 1
                logger(f"Email sent to {email} ")

                # Update state
                state["last_row"] = idx
                save_state(state)

                # random sleep to avoid spam filters
                time.sleep(random.randint(MIN_DELAY, MAX_DELAY))

                # Batch control
                if sent_today % BATCH_SIZE == 0 :
                    logger(f"Cooling down for {BATCH_SLEEP}s before next batch...")
                    time.sleep(BATCH_SLEEP)

            except Exception as e:
                sent_today -= 1
                logger(f"Failed for {email}: {e}")

    finally :
        smtp_conn.quit()

    logger("Run Completed Today")


if __name__ == "__main__":
    main()
