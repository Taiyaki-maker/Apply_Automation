import pandas as pd
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import my_gmail_account as gmail  # Gmailアカウント情報を格納したファイルをインポート

def send_email(to_email, hotel_name, subject, resume_path):
    # Gmail SMTPサーバー情報
    smtp_server = "smtp.gmail.com"
    port = 465

    # メールメッセージの作成
    msg = MIMEMultipart()
    msg["From"] = gmail.account
    msg["To"] = to_email
    msg["Subject"] = subject

    # HTML形式の本文をフォーマットして添付
    html_body = f"""
    <html>
    <body>
        <p>Dear {hotel_name} Team,</p>
    
        <p>My name is Rena Yamada, and I am applying for the waitress position at {hotel_name}. With six years of experience in cafes, casual dining, and Japanese pubs, I have honed my customer service, communication, and time management skills.</p>
    
        <p>I am currently in Australia on a student visa and eager to contribute to your team while enhancing my English skills. I take pride in creating a welcoming environment for customers and have been trusted to train staff and handle independent responsibilities in my previous roles.</p>
    
        <p>Thank you for considering my application. Please find my resume attached, and I look forward to the opportunity to join your team.</p>
    
        <p>Warm regards,<br>
        Rena Yamada<br>
        Phone: 123 456 789<br>
        Email: <a href="mailto:1234567@gmail.com">1234567@gmail.com</a></p>
    </body>
    </html>
    """
    msg.attach(MIMEText(html_body, "html"))

    # PDFファイルを添付
    try:
        with open(resume_path, "rb") as attachment:
            part = MIMEBase("application", "pdf")
            part.set_payload(attachment.read())
        
        # 添付ファイルをエンコード
        encoders.encode_base64(part)
        
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={resume_path.split('/')[-1]}",
        )

        # メッセージに添付ファイルを追加
        msg.attach(part)
    except Exception as e:
        print(f"Failed to attach resume: {e}")

    # GmailのSMTPサーバーにSSLを使って接続し、メールを送信
    try:
        with smtplib.SMTP_SSL(smtp_server, port, context=ssl.create_default_context()) as server:
            server.login(gmail.account, gmail.password)
            server.send_message(msg)
            print(f"Email sent to {to_email} for {hotel_name}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

def send_applications_from_excel(filename, subject="Application for Receptionist Position", resume_path="Documents/Resume/resume.pdf"):
    try:
        # Excelファイルを読み込み、実行フラグがFalseの行を取得
        df = pd.read_excel(filename)
        df_to_send = df[(df["execution_flag"] == False) & df["email"].notnull()]

        # メールを送信し、送信後にフラグを更新
        for index, row in df_to_send.iterrows():
            hotel_name = row["name"]  # `Name`列のホテル名を取得
            to_email = row["email"]
            send_email(to_email, hotel_name, subject, resume_path)
            df.at[index, "execution_flag"] = True  # フラグをTrueに更新

        # 実行フラグが更新されたExcelファイルを上書き保存
        df.to_excel(filename, index=False)
        print("All emails sent successfully and execution flags updated.")

    except Exception as e:
        print(f"Failed to read or process the Excel file: {e}")

# 使用例
filename = "Resume/places_data_real.xlsx"  # Excelファイルのパスを変更
subject = "Application for Waitress Position"
resume_path = "Resume/resume_gf.pdf"  # レジュメのファイルパスを変更

send_applications_from_excel(filename, subject, resume_path)