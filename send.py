# -*- coding: utf-8 -*-
import smtplib
from email.utils import make_msgid
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from openpyxl import load_workbook
global count
count = 1
excel_name = 'BU_stu_email.xlsx' #將excel文件放在與這個py文件的同一個文件夾，並在括號内填入名稱加格式
poster_name = 'poster.png' #將poster照片放在與這個py文件的同一個文件夾，並在括號内填入名稱加格式
workbook = load_workbook(excel_name)
sheets = workbook.get_sheet_names()         #从名称获取sheet
booksheet = workbook.get_sheet_by_name(sheets[0])
rows = booksheet.rows
emails = []
for row in rows:
    emails.append([col.value for col in row][0])


def send_email(email):
    mail_username = '16202406@life.hkbu.edu.hk'
    mail_password = 'Handsomeguy0!'
    from_addr = mail_username
    to_addrs = (email)
    content = "曾幾何時你發現能夠講秘密嘅朋友越來越少?\n\n\
    最好嘅朋友又因為某各種原因變得疏離?\n\
    用寫信嘅方式開始認識你喺大學生涯裏面嘅摰友,\n\n\
    把心中所想、所見、所聞、所憂、所慮, 甚至所煩都盡訴諸名為筆友嘅樹洞\n\
    讓他們為你分擔, 並成為你的支持者, 你前進的動力\n\n\
    千言萬語, 訴於筆下,\n\n\
    今天,\n\
    就把信  寄 ‧ 筆友,\n\
    我們,\n\
    筆見不散!\n"
    # HOST & PORT
    HOST = 'smtp.gmail.com'
    PORT = 25
    # Create SMTP Object
    smtp = smtplib.SMTP()
    print('connecting ...')
    # show the debug log
    smtp.set_debuglevel(1)
    # connet
    try:
        smtp.connect(HOST, PORT)
    except:
        print('CONNECT ERROR ****')
    # gmail uses ssl
    smtp.starttls()
    # login with username & password
    try:
        print('loginning ...')
        smtp.login(mail_username, mail_password)
    except:
        print('LOGIN ERROR ****')
    # fill content with MIMEText's object
    file = open(poster_name, 'rb')
    img_data = file.read()
    file.close()
    img = MIMEImage(img_data)
    img.add_header('Content-ID', 'dns_config')
    msgText = MIMEText(content)
    msg = MIMEMultipart('related')
    msg.attach(msgText)
    msg['Message-ID'] = make_msgid()
    msg['From'] = from_addr
    msg['To'] = 'to_addrs'
    msg['Subject'] = '寄 ‧ 筆友'
    msg.attach(img)
    print("%d sent successfully" %count)
    smtp.sendmail(from_addr, to_addrs, msg.as_string())
    smtp.quit()


if __name__ == '__main__':
    for email in emails:
        if count <=123:
            send_email(email)
            count = count +1
        else:
            print("500 emails have sent. See you tomorrow.")
            break