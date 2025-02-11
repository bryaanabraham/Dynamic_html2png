import pandas as pd
from datetime import datetime
import random
import smtplib
from email.message import EmailMessage
import mimetypes
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from PIL import Image
import time


class Config:
    def __init__(self, csv_path, html_path, screenshot_path, output_image, dimensions, smtp_server, sender_email, sender_password, recipient_email):
        self.csv_path = csv_path
        self.html_path = html_path
        self.screenshot_path = screenshot_path
        self.output_image = output_image
        self.dimensions = dimensions
        self.smtp_server = smtp_server
        self.sender_email = sender_email
        self.sender_password = sender_password
        self.recipient_email = recipient_email

def find_birthdays(config):
    today = datetime.today().strftime('%d-%b').upper()
    
    df = pd.read_csv(config.csv_path, encoding='utf-8')
    df['DOB'] = df['DOB'].str.upper()    
    names = df[df['DOB'] == today]['Name'].str.split().str[0].tolist()
    
    return names if names else None



def update_html(config, first_name, replacer):
    with open(config.html_path, 'r', encoding='utf-8') as file:
        html_content = file.read()

    if first_name:
        html_content = html_content.replace(replacer, first_name)
        with open(config.html_path, 'w', encoding='utf-8') as file:
            file.write(html_content)
        print(f"Updated HTML {replacer} -> {first_name}, saved as {config.html_path}")
        return config.html_path


def html_to_image(config, first_name):
    try:
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=2560x2160")
        driver = webdriver.Chrome(options=chrome_options)
        driver.get(f"file:///{config.html_path}")

        time.sleep(2)  # Allow page elements to load

        driver.save_screenshot(config.screenshot_path)
        driver.quit()

        img = Image.open(config.screenshot_path)
        cropped_img = img.crop(config.dimensions)
        cropped_img.save(config.output_image)

        print(f"Image saved as {config.output_image}")

    except Exception as e:
        print(f"Error converting HTML to image: {e}")
    
    update_html(config, '{name}', first_name)

    return cropped_img

import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

def send_email(config, first_name, bcc_emails):
    msg = MIMEMultipart()
    msg["Subject"] = "test -> Happy Birthday!"
    msg["From"] = config.sender_email
    msg["To"] = config.recipient_email
    if bcc_emails:
        msg["Bcc"] = ", ".join(bcc_emails)

    # HTML content to include in the body
    html_content = """
    <html>
        <head>
            <style>
                body {
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    height: 100vh;
                    margin: 0;
                }
            </style>
        </head>
        <body>
            <img src="cid:birthday_image" alt="Birthday Image"/>
        </body>
    </html>

    """
    msg.attach(MIMEText(html_content, "html"))
    try:
        with open(config.output_image, "rb") as img_file:
            img_data = img_file.read()
            mime_type, _ = mimetypes.guess_type(config.output_image)
            main_type, sub_type = mime_type.split("/") if mime_type else ("application", "octet-stream")
            
            # Create the MIMEImage object with Content-ID
            image_attachment = MIMEImage(img_data, maintype=main_type, subtype=sub_type)
            image_attachment.add_header('Content-ID', '<birthday_image>')
            msg.attach(image_attachment)
    except Exception as e:
        print(f"⚠ Could not embed image: {e}")
        return

    # Send the email
    try:
        with smtplib.SMTP(config.smtp_server, 587) as server:
            server.starttls()
            server.login(config.sender_email, config.sender_password)
            server.send_message(msg)
            print(f"✅ Email with embedded image sent successfully! \nSender: {config.sender_email}\nRecipient: {config.recipient_email}")
    except Exception as e:
        print(f"❌ Error: {e}")



if __name__ == "__main__":
    html_paths = [
        "D:/Mahindra AI/Projects/Birthday Card/Frame1/Frame1.html",
        "D:/Mahindra AI/Projects/Birthday Card/Frame2/Frame2.html"
    ]
    dims = [
        (343, 0, 951, 874),
        (262, 187, 1012, 687)
    ]

    # ind = random.randint(0, 1)
    ind = 0
    print(f"Sending Frame{ind + 1}")

    config = Config(
        csv_path="../birthdays.csv",
        html_path=html_paths[ind],
        screenshot_path="../Frame2/send.png",
        output_image="send.png",
        dimensions=dims[ind],
        smtp_server = "smtp.gmail.com",
        # smtp_server = "smtp.office365.com",
        sender_email="bryaanabraham@gmail.com",
        sender_password="nojm tezm qsld ddpx",
        recipient_email="50010135@mahindra.com"
    )

    first_names = find_birthdays(config)
    if first_names:
        for name in first_names:
            updated_html = update_html(config, name, '{name}')
            if updated_html:
                img = html_to_image(config, name)
                bcc_emails = []#"25013529@mahindra.com"]
                send_email(config, name, bcc_emails)
    else:
        print("No birthdays today.")
