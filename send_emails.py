import smtplib
import ssl
import sys
import os
import logging
import time
import threading
import schedule
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# Get the application directory
if getattr(sys, 'frozen', False):
    # Running as a bundled app
    if sys.platform == 'darwin':
        # For macOS app bundle, we need to use the MacOS directory
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(sys.executable)
else:
    # Running as a script
    application_path = os.path.dirname(os.path.abspath(__file__))

# Create necessary directories
data_dir = os.path.join(application_path, 'data')
logs_dir = os.path.join(application_path, 'logs')
os.makedirs(data_dir, exist_ok=True)
os.makedirs(logs_dir, exist_ok=True)

# Set up logging
log_file = os.path.join(logs_dir, 'email_log.txt')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# Function to generate email signature
def get_email_signature(use_html=False):
    if use_html:
        return """
        <br><br>
        <div style="font-family: Arial, sans-serif; font-size: 14px;">
            <p>---</p>
            <p>Pagarbiai, Best regards,</p>
            <p><strong>Lukas Gudavičius</strong><br>
            Administratorius</p>
            <p>E. lukas@bilger.lt<br>
            Office: Pramonės g. 2A (antras aukštas),<br>
            LT-72328, Tauragė<br>
            <a href="https://bilger.lt">bilger.lt</a></p>
        </div>
        """
    else:
        return """
---
Pagarbiai, Best regards,

Lukas Gudavičius
Administratorius

E. lukas@bilger.lt
Office: Pramonės g. 2A (antras aukštas),
LT-72328, Tauragė
bilger.lt
        """

# Function to send one email
def send_email(sender_email, recipient_email, subject, body, smtp_server, smtp_port, username, password, use_html=True):
    try:
        logging.info(f"Ruošiamas laiškas adresatui: {recipient_email}")
        
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = recipient_email

        # Add HTML signature to the body
        full_body = body + get_email_signature(use_html=True)

        # Attach both HTML and plain text versions
        msg.attach(MIMEText(full_body, 'html', 'utf-8'))
        msg.attach(MIMEText(full_body, 'plain', 'utf-8'))

        logging.info(f"Jungiamasi prie SMTP serverio: {smtp_server}:{smtp_port}")
        
        # Send via SMTP
        context = ssl.create_default_context()
        with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
            logging.info("Pradedama TLS sesija...")
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            
            logging.info("Bandoma prisijungti prie serverio...")
            server.login(username, password)
            
            logging.info(f"Siunčiamas laiškas į {recipient_email}")
            server.send_message(msg)
            logging.info(f"✅ Laiškas išsiųstas: {recipient_email}")

    except smtplib.SMTPAuthenticationError:
        logging.error("❌ Autentifikavimo klaida – patikrink vartotojo vardą/slaptažodį.")
        raise
    except smtplib.SMTPException as e:
        logging.error(f"❌ SMTP klaida ({recipient_email}): {str(e)}")
        raise
    except Exception as e:
        logging.error(f"❌ Kita klaida ({recipient_email}): {str(e)}")
        raise

def create_ssl_context():
    """Create a secure SSL context with appropriate settings."""
    context = ssl.create_default_context()
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE
    return context

def connect_to_smtp(server, port, username, password, max_retries=3, retry_delay=5):
    """Attempt to connect to SMTP server with retry mechanism."""
    context = create_ssl_context()
    
    for attempt in range(max_retries):
        try:
            logging.info(f"Attempting to connect to {server} on port {port} (attempt {attempt + 1}/{max_retries})")
            
            if port == 465:
                # SSL connection
                server = smtplib.SMTP_SSL(server, port, context=context, timeout=60)
            else:
                # STARTTLS connection
                server = smtplib.SMTP(server, port, timeout=60)
                server.set_debuglevel(1)  # Enable debug logging
                server.ehlo()
                if port == 587:  # Special handling for webmail on port 587
                    server.starttls(context=context)
                    server.ehlo()
                server.login(username, password)
            
            logging.info(f"Successfully connected to {server} on port {port}")
            return server
            
        except (smtplib.SMTPException, ssl.SSLError, ConnectionError) as e:
            logging.error(f"Connection attempt {attempt + 1} failed: {str(e)}")
            if attempt < max_retries - 1:
                logging.info(f"Waiting {retry_delay} seconds before retrying...")
                time.sleep(retry_delay)
            else:
                raise Exception(f"Failed to connect to SMTP server after {max_retries} attempts: {str(e)}")

def send_emails_to_clients(subject, body):
    # Email configuration
    sender_email = "seo@bilger.lt"
    password = "K^]4L,,5@n[&"
    smtp_server = "webmail.bilger.lt"
    smtp_port = 465  # Using SSL
    
    logger.info(f"Bandoma prisijungti prie {smtp_server}:{smtp_port}")
    
    # Create SSL context
    context = ssl.create_default_context()
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE
    
    server = None
    try:
        # Read email addresses from file
        email_list_path = os.path.join(data_dir, 'email_list.txt')
        with open(email_list_path, 'r', encoding='utf-8') as f:
            email_addresses = [line.strip() for line in f if line.strip()]
        
        logger.info(f"Rasta {len(email_addresses)} el. pašto adresų")
        
        # Connect to SMTP server using SSL
        try:
            logger.info("Bandoma prisijungti prie SMTP serverio...")
            server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=context, timeout=30)
            logger.info("SMTP_SSL prisijungimas sėkmingas")
            
            logger.info("Bandoma prisijungti su vartotojo duomenimis...")
            server.login(sender_email, password)
            logger.info("Prisijungimas sėkmingas")
        except smtplib.SMTPAuthenticationError as auth_error:
            logger.error(f"Autentifikacijos klaida: {str(auth_error)}")
            raise
        except Exception as e:
            logger.error(f"Klaida prisijungiant prie SMTP serverio: {str(e)}")
            raise
        
        # Send emails
        successful_sends = 0
        emails_sent_in_group = 0
        
        for email in email_addresses:
            try:
                # Jei išsiųsta 10 laiškų, darome 10 minučių pauzę ir atnaujiname ryšį
                if emails_sent_in_group >= 10:
                    logger.info("Išsiųsta 10 laiškų. Daroma 10 minučių pauzė...")
                    time.sleep(600)  # 600 sekundžių = 10 minučių
                    emails_sent_in_group = 0  # Atstatome skaitliuką
                    
                    # Atnaujiname ryšį su SMTP serveriu
                    logger.info("Atnaujinamas ryšys su SMTP serveriu...")
                    if server:
                        try:
                            server.quit()
                        except:
                            pass
                    server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=context)
                    server.login(sender_email, password)
                    logger.info("Sėkmingai atnaujintas ryšys su SMTP serveriu")
                
                # Create message
                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = email
                msg['Subject'] = subject
                
                # Add signature to the body
                full_body = body + get_email_signature(use_html=True)
                msg.attach(MIMEText(full_body, 'html'))
                
                # Send email
                server.send_message(msg)
                successful_sends += 1
                emails_sent_in_group += 1
                logger.info(f"Laiškas sėkmingai išsiųstas į {email}")
                
                # Jei dar nepasiekėme 10 laiškų, darome 2 minučių pauzę
                if emails_sent_in_group < 10:
                    time.sleep(120)  # 120 sekundžių (2 minutės) tarp laiškų siuntimo
                
            except Exception as e:
                logger.error(f"Klaida siunčiant laišką į {email}: {str(e)}")
                # Bandome atnaujinti ryšį jei įvyko klaida
                try:
                    if server:
                        server.quit()
                    server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=context)
                    server.login(sender_email, password)
                    logger.info("Sėkmingai atnaujintas ryšys po klaidos")
                except Exception as reconnect_error:
                    logger.error(f"Klaida bandant atnaujinti ryšį: {str(reconnect_error)}")
                continue
        
        logger.info(f"Baigtas laiškų siuntimas. Sėkmingai išsiųsta: {successful_sends} iš {len(email_addresses)}")
            
    except FileNotFoundError:
        logger.error(f"Nerastas el. pašto adresų sąrašo failas: {email_list_path}")
        raise
    except Exception as e:
        logger.error(f"Klaida siunčiant laiškus: {str(e)}")
        raise
    finally:
        # Uždarome SMTP serverio prisijungimą
        if server:
            try:
                server.quit()
                logger.info("SMTP serverio prisijungimas sėkmingai uždaromas")
            except Exception as e:
                logger.error(f"Klaida uždarant SMTP serverio prisijungimą: {str(e)}")

def send_log_email():
    try:
        # Read the log file
        with open(log_file, 'r', encoding='utf-8') as f:
            log_content = f.read()
        
        # Get the last 50 lines of the log
        log_lines = log_content.split('\n')
        last_50_lines = '\n'.join(log_lines[-50:]) if len(log_lines) > 50 else log_content
        
        # Create email
        msg = MIMEMultipart('alternative')
        msg['Subject'] = f'SEO Straipsnių siuntimo log - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
        msg['From'] = "seo@bilger.lt"
        msg['To'] = "seo@bilger.lt"  # Siunčiame į tą patį paštą

        # Create HTML version
        html_content = f"""
        <html>
        <body style="font-family: Arial, sans-serif; font-size: 14px;">
            <h2>Paskutiniai išsiųsti laiškai:</h2>
            <pre style="background-color: #f5f5f5; padding: 15px; border-radius: 5px;">
            {last_50_lines}
            </pre>
        </body>
        </html>
        """
        
        # Create plain text version
        text_content = f"Paskutiniai išsiųsti laiškai:\n\n{last_50_lines}"
        
        msg.attach(MIMEText(text_content, 'plain', 'utf-8'))
        msg.attach(MIMEText(html_content, 'html', 'utf-8'))
        
        # Send email
        context = ssl.create_default_context()
        with smtplib.SMTP("webmail.bilger.lt", 465) as server:
            server.starttls(context=context)
            server.login("seo@bilger.lt", "K^]4L,,5@n[&")
            server.send_message(msg)
            
        logger.info("Log informacija sėkmingai išsiųsta")
        
    except Exception as e:
        logger.error(f"Klaida siunčiant log informaciją: {str(e)}")

def start_log_scheduler():
    def run_scheduler():
        while True:
            schedule.run_pending()
            time.sleep(1)
    
    # Nustatome kas 30 minučių siuntimą
    schedule.every(30).minutes.do(send_log_email)
    
    # Paleidžiame scheduler'į atskirame thread'e
    scheduler_thread = threading.Thread(target=run_scheduler, daemon=True)
    scheduler_thread.start()
    logger.info("Log siuntimo planavimas paleistas")

# For running from terminal
if __name__ == "__main__":
    start_log_scheduler()  # Paleidžiame log siuntimo planavimą
    subject = "Test from terminal"
    body = """<p>Hello,</p>
    <p>This is a test email.</p>
    <p>Thank you for reading our newsletter!</p>"""
    send_emails_to_clients(subject, body)
