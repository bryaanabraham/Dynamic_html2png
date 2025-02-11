import csv
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from PIL import Image
import time

def find_birthday(csv_file):
    today = datetime.today().strftime('%d-%b').upper()
    with open(csv_file, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['DOB'].upper() == today:
                return row['Name'].split()[0]
    return None


def update_html(html_file, first_name, replacer='{name}'):
    with open(html_file, 'r', encoding='utf-8') as file:
        html_content = file.read()

    if first_name:
        html_content = html_content.replace(replacer, first_name)

        with open(html_file, 'w', encoding='utf-8') as file:
            file.write(html_content)

        print(f"Updated HTML {replacer} -> {first_name}, saved as {html_file}")
        return html_file
    else:
        print("No birthday found for today.")
        return None

def html_to_image(html_file):
    try:
        chrome_options = Options()
        chrome_options.add_argument("--headless")  # Run in headless mode
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=2560x2160")  # Set high resolution for better quality
        driver = webdriver.Chrome(options=chrome_options)
        driver.get("file:///D:/Mahindra%20AI/Projects/Birthday%20Card/Frame2/Frame2.html")

        time.sleep(2)  # Wait for page elements to load properly

        # Capture full-page screenshot
        screenshot_path = "../Frame2/send.png"
        driver.save_screenshot(screenshot_path)

        # Close the browser
        driver.quit()

        # Crop the image to desired size
        img = Image.open(screenshot_path)
        cropped_img = img.crop((262, 187, 1012, 687))
        cropped_img.save("send.png")

        print("Image saved as send.png")

    except Exception as e:
        print(f"Error converting HTML to image: {e}")

    update_html(html_file, '{name}',first_name)
csv_file = '../birthdays.csv'
html_file = 'Frame2.html'
first_name = find_birthday(csv_file)

if first_name:
    updated_html = update_html(html_file, first_name)
    if updated_html:
        output_image = f"{first_name}.png"
        html_to_image(updated_html)
