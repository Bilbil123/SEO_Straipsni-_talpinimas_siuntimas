import sys
import os
import time
import json
from PyQt5.QtWidgets import (QApplication, QWidget, QMessageBox, QPushButton, 
                           QLineEdit, QTextEdit, QVBoxLayout, QHBoxLayout, QLabel)
from PyQt5.QtCore import QTimer, QThread, pyqtSignal, QDateTime, Qt
from email_sender_ui import EmailSenderUI
import send_emails
import logging

class EmailSenderThread(QThread):
    finished = pyqtSignal()
    error = pyqtSignal(str)
    log_update = pyqtSignal(str)

    def __init__(self, subject, body):
        super().__init__()
        self.subject = subject
        self.body = body

    def run(self):
        try:
            # Create a custom logging handler that emits signals
            class SignalHandler(logging.Handler):
                def __init__(self, signal):
                    super().__init__()
                    self.signal = signal

                def emit(self, record):
                    msg = self.format(record)
                    self.signal.emit(msg)

            # Set up logging to emit signals
            handler = SignalHandler(self.log_update)
            handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            
            # Get the logger from send_emails
            logger = logging.getLogger()
            logger.addHandler(handler)
            
            # Send emails
            send_emails.send_emails_to_clients(self.subject, self.body)
            
            # Remove our custom handler
            logger.removeHandler(handler)
            
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))

class EmailSenderApp(EmailSenderUI):
    def __init__(self):
        # Set up the data directory first
        if getattr(sys, 'frozen', False):
            # Running as a bundled app
            if sys.platform == 'darwin':
                # For macOS app bundle
                self.data_dir = os.path.join(os.path.dirname(sys.executable), 'data')
            else:
                self.data_dir = os.path.join(os.path.dirname(sys.executable), 'data')
        else:
            # Running as a script
            self.data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')

        # Create necessary directories
        os.makedirs(self.data_dir, exist_ok=True)
        
        # Create log file if it doesn't exist
        self.log_file = os.path.join(self.data_dir, 'email_log.txt')
        if not os.path.exists(self.log_file):
            with open(self.log_file, 'w', encoding='utf-8') as f:
                f.write('')  # Create empty file
        
        # Initialize parent class with data_dir
        super().__init__(data_dir=self.data_dir)
        
        # Set up log update timer
        self.log_timer = QTimer(self)
        self.log_timer.timeout.connect(self.update_log_display)
        self.log_timer.start(1000)  # Update every second

    def update_log_display(self):
        try:
            if os.path.exists(self.log_file):
                with open(self.log_file, 'r', encoding='utf-8') as f:
                    log_content = f.read()
                if log_content:  # Only update if there's content
                    self.log_text.setPlainText(log_content)
                    # Scroll to bottom
                    scrollbar = self.log_text.verticalScrollBar()
                    scrollbar.setValue(scrollbar.maximum())
        except Exception as e:
            print(f"Error updating log display: {str(e)}")

    def send_emails(self):
        subject = self.subject_input.text().strip()
        body = self.text_edit.toHtml()  # Get HTML content

        if not subject or not body:
            QMessageBox.warning(self, "Klaida", "Prašome užpildyti laiško antraštę ir tekstą!")
            return

        # Update log immediately
        self.log_text.append(f"[{QDateTime.currentDateTime().toString('yyyy-MM-dd HH:mm:ss')}] Pradedamas laiškų siuntimas...")
        self.send_button.setEnabled(False)

        # Sukuriame naują thread
        self.email_thread = EmailSenderThread(subject, body)
        self.email_thread.finished.connect(self.on_sending_finished)
        self.email_thread.error.connect(self.on_sending_error)
        self.email_thread.log_update.connect(self.on_log_update)
        self.email_thread.start()

    def on_log_update(self, message):
        self.log_text.append(message)
        # Force scroll to bottom
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def on_sending_finished(self):
        self.send_button.setEnabled(True)
        self.log_text.append(f"[{QDateTime.currentDateTime().toString('yyyy-MM-dd HH:mm:ss')}] Laiškai sėkmingai išsiųsti!")
        QMessageBox.information(self, "Sėkmė", "Laiškai sėkmingai išsiųsti!")
        # Uždarome thread
        if hasattr(self, 'email_thread'):
            self.email_thread.quit()
            self.email_thread.wait()

    def on_sending_error(self, error_message):
        self.send_button.setEnabled(True)
        self.log_text.append(f"[{QDateTime.currentDateTime().toString('yyyy-MM-dd HH:mm:ss')}] Klaida: {error_message}")
        QMessageBox.critical(self, "Klaida", f"Klaida siunčiant laiškus: {error_message}")
        # Uždarome thread
        if hasattr(self, 'email_thread'):
            self.email_thread.quit()
            self.email_thread.wait()

    def load_emoji_data(self):
        try:
            emoji_path = os.path.join(self.data_dir, 'emoji_data.json')
            if not os.path.exists(emoji_path):
                print(f"Emoji data file not found at: {emoji_path}")  # Debug print
                return {}
            with open(emoji_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('categories', {})
        except Exception as e:
            print(f'Error loading emoji data: {str(e)}')  # Debug print
            return {}

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Paleidžiame log siuntimo planavimą
    try:
        from send_emails import start_log_scheduler
        start_log_scheduler()
    except Exception as e:
        print(f"Klaida paleidžiant log siuntimo planavimą: {str(e)}")
    
    # Patikriname ar yra reikalingi failai
    data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
    required_files = ['emoji_data.json', 'holiday_templates.json', 'email_list.txt']
    
    for file in required_files:
        file_path = os.path.join(data_dir, file)
        if not os.path.exists(file_path):
            print(f"Trūksta reikalingo failo: {file_path}")
            # Sukuriame tuščius failus, jei jų nėra
            if file == 'email_list.txt':
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write('')
            elif file == 'emoji_data.json':
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump({"categories": {}}, f)
            elif file == 'holiday_templates.json':
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump({"promotions": {}}, f)
    
    window = EmailSenderApp()
    window.show()
    sys.exit(app.exec_())
