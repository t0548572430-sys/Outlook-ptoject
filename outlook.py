from flask import Flask, request
import os
import pythoncom
import win32com.client as win32

app = Flask(__name__)

# ----- דף HTML עם הטופס והעיצוב -----
@app.route('/')
def index():
    return '''
    <!DOCTYPE html>
    <html lang="he">
    <head>
        <meta charset="UTF-8">
        <title>שליחת טיוטות ב-Outlook</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                background-color: #f9f9f9;
                margin: 50px;
                direction: rtl;
            }
            h2 {
                color: #333;
            }
            form {
                background-color: #fff;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 0 10px rgba(0,0,0,0.1);
                max-width: 500px;
            }
            input[type=text], textarea {
                width: 100%;
                padding: 8px;
                margin: 6px 0 12px 0;
                border: 1px solid #ccc;
                border-radius: 4px;
                box-sizing: border-box;
            }
            input[type=file] {
                margin: 6px 0 12px 0;
            }
            input[type=submit] {
                background-color: #4CAF50;
                color: white;
                padding: 10px 20px;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 16px;
            }
            input[type=submit]:hover {
                background-color: #45a049;
            }
            .container {
                display: flex;
                justify-content: center;
            }
        </style>
    </head>
    <body>
        <div class="container">
        <form method="POST" action="/send" enctype="multipart/form-data">
            <h2>טופס לשליחת טיוטות ב-Outlook</h2>
            <label>כתובת נמענים (מופרדות בפסיק):</label><br>
            <input type="text" name="emails" required><br>
            <label>נושא:</label><br>
            <input type="text" name="subject" required><br>
            <label>גוף ההודעה:</label><br>
            <textarea name="body" rows="5" required></textarea><br>
            <label>קובץ מצורף:</label><br>
            <input type="file" name="attachment"><br>
            <input type="submit" value="שלח">
        </form>
        </div>
    </body>
    </html>
    '''

# ----- קבלת הנתונים ויצירת טיוטות Outlook -----
@app.route('/send', methods=['POST'])
def send():
    emails = request.form['emails'].split(',')
    subject = request.form['subject']
    body = request.form['body']
    file = request.files.get('attachment')

    attachment_path = ""
    if file:
        attachment_path = os.path.join(os.getcwd(), file.filename)
        file.save(attachment_path)
    
    # אתחול COM לפני שימוש ב-Outlook
    pythoncom.CoInitialize()
    outlook = win32.Dispatch('outlook.application')

    for email in emails:
        email = email.strip()
        mail = outlook.CreateItem(0)  # olMailItem
        mail.To = email
        mail.Subject = subject
        mail.Body = body
        if attachment_path and os.path.exists(attachment_path):
            mail.Attachments.Add(attachment_path)
        mail.Display()  # פותח טיוטה ב-Outlook

    return '''
    <h3>כל הטיוטות נפתחו ב-Outlook!</h3>
    <p><a href="/">חזור לטופס</a></p>
    '''

# ----- הרצת השרת -----
if __name__ == '__main__':
    app.run(debug=True)