from flask import Flask, request
import win32com.client
import pythoncom
import os
import time
import win32gui, win32con
import win32gui
import win32con
app = Flask(__name__)

def ensure_outlook_running():
    """מנסה להתחבר ל-Outlook, ואם לא רץ – מפעיל אותו"""
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
    except:
        # אם לא רץ, פותח את Outlook
        os.startfile("outlook")
        # נותן לו כמה שניות לעלות
        time.sleep(5)
        outlook = win32com.client.Dispatch('Outlook.Application')
    return outlook

# Route לשליחת מייל עם אפשרות לצירוף קובץ
@app.route('/sendMail', methods=['POST'])
def send_mail():
    # CoInitialize ל-thread של Flask
    pythoncom.CoInitialize()

    subject = request.form.get('subject', '')
    body = request.form.get('body', '')
    recipients = request.form.get('recipients', '').split(',')
    file = request.files.get('attachment')

    # שמירה זמנית של הקובץ
    attachment_path = None
    if file and file.filename != '':
        attachment_path = os.path.join(os.getcwd(), file.filename)
        file.save(attachment_path)

    outlook = ensure_outlook_running()

    for recipient in recipients:
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = recipient.strip()
        mail.Subject = subject
        mail.Body = body
        hwnd = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE) # או SW_SHOWNORMAL
        if attachment_path:
            mail.Attachments.Add(attachment_path)
        
        # 1. קודם כל מציגים את החלון - זה גורם לו לקפוץ למסך
        mail.Display()
        
        # 2. שומרים בטיוטות (אופציונלי אם כבר עשית Display)
        mail.Save()


    return 'OK'

@app.route('/')
def home():
    return 'Outlook Service is running!'

if __name__ == '__main__':
    app.run(port=5000)
