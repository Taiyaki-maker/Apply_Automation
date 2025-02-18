import pandas as pd
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import my_gmail_account as gmail  # Gmailアカウント情報を格納したファイルをインポート

def send_email(to_email, cafe_name, subject, resume_path):
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
        <p>Dear Hiring Manager,</p>
    
        <p>I am excited to apply for the Barista position at {cafe_name}. Currently working as a barista at Brunetti Oro in Melbourne, I bring two years of barista experience from Japan, along with strong customer service and coffee-making skills. My roles have prepared me to excel in fast-paced, customer-focused environments.</p>
    
        <p>Passionate about coffee culture, I am dedicated to creating excellent customer experiences and maintaining a welcoming atmosphere. I am available to work flexible hours, including weekends and holidays.</p>
    
        <p>Please find my resume attached. I look forward to the opportunity to contribute my skills and enthusiasm to {cafe_name}.</p>
    
        <p>Thank you for your time and consideration.</p>
    
        <p>Warm regards,<br>
        Taiki Ogura<br>
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
            print(f"Email sent to {to_email} for {cafe_name}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

def send_applications_from_excel(filename, subject, resume_path):
    try:
        # Excelファイルを読み込み、実行フラグがFalseの行を取得
        df = pd.read_excel(filename)
        
        # execution_flag列の位置を動的に取得
        execution_flag_column = df.columns.get_loc("execution_flag")

        # 実行フラグがFalseの行を取得
        df_to_send = df[(df.iloc[:, execution_flag_column] == False) & df["email"].notnull()]
        
        #df_to_send = df[(df["execution_flag"] == False) & df["email"].notnull()]

        # メールを送信し、送信後にフラグを更新
        for index, row in df_to_send.iterrows():
            cafe_name = row["name"]  # `Name`列のホテル名を取得
            to_email = row["email"]
            send_email(to_email, cafe_name, subject, resume_path)
            #df.at[index, "execution_flag"] = True  # フラグをTrueに更新
            df.iloc[index, execution_flag_column] = True  # フラグをTrueに更新

        # 実行フラグが更新されたExcelファイルを上書き保存
        df.to_excel(filename, index=False)
        print("All emails sent successfully and execution flags updated.")

    except Exception as e:
        print(f"Failed to read or process the Excel file: {e}")

# 使用例
filename = "Resume/places_data.xlsx"
subject = "Application for Barista Position"
resume_path = "Resume/resume_cafe.pdf"  # レジュメのファイルパスを変更

send_applications_from_excel(filename, subject, resume_path)
