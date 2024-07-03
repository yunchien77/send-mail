import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from cfp import EmailSender
from email.mime.image import MIMEImage

sender = EmailSender("config.ini")

# 讀取 Excel 文件
df = pd.read_excel("evaselect.xlsx")

# 設定 SMTP 伺服器
server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
server.login(sender.sender_email, sender.password)

for index, row in df.iterrows():
    receiver_email = row["Email"]
    receiver_name = row["Name"]

    # 建立郵件物件
    message = MIMEMultipart("alternative")
    message["Subject"] = "2024 EVA Select Webinar 線上研討會後續意見調查"
    message["From"] = sender.sender_email
    message["To"] = receiver_email

    # 郵件正文
    html = f"""
<html>

<body>
  <p>
    親愛的{receiver_name}
  </p>
  <p>
    感謝您報名參加我們舉辦的線上研討會。
  </p>
  <p>
  感謝您撥允參加2024 EVA Select Webinar，
  有您的參加便讓我們有了繼續舉辦研討會的動力。
  為了研討會能辦得越來越好，
  能否再請您撥個2分鐘給我們一些建議呢？感謝您！
  </p>
  <p>
  意見調查連結: 
  <a href="https://docs.google.com/forms/d/e/1FAIpQLSeaKxxXXL14YSO7GhVtlcqWGHn7OYnRjHDoOQMF68L3MKH-uw/viewform">
  點擊</a>
  </p>
  <p>
    聯絡人：陳玟廷 電話：7752-2586 分機 104
    Email：anastasia.chen@cancerfree.io
  </p>
</body>

</html>
    """

    # 加入郵件正文

    message.attach(MIMEText(html, "html"))

    with open("schedule.png", "rb") as f:
        mime = MIMEImage(f.read())
        mime.add_header("Content-Disposition", "attachment", filename="schedule.png")
        mime.add_header("X-Attachment-Id", "0")
        mime.add_header("Content-ID", "<0>")
        message.attach(mime)

    # 發送郵件
    server.sendmail(sender.sender_email, receiver_email, message.as_string())
    print(f"{receiver_name},done")

# 關閉 SMTP 伺服器連接
server.quit()
