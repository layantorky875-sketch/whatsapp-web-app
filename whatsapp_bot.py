import time
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ========= SETTINGS =========
MAX_PER_HOUR = 15        # أقصى عدد رسائل في الساعة
DELAY_MIN = 120          # أقل انتظار (ثواني)
DELAY_MAX = 240          # أقصى انتظار (ثواني)

MESSAGES = [
    {"phone": "201001234567", "text": "Hello Ahmed"},
    {"phone": "966501234567", "text": "Hello Mohammed"}
]
# ============================

def start_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--user-data-dir=chrome-data")  # حفظ الجلسة
    options.add_argument("--profile-directory=Default")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    return driver


driver = start_driver()
driver.get("https://web.whatsapp.com")

print("🔹 اعمل Login مرة واحدة لو لسه")
time.sleep(25)

sent_count = 0
start_time = time.time()

for item in MESSAGES:

    if sent_count >= MAX_PER_HOUR:
        print("⛔ وصلت للحد الأقصى في الساعة")
        break

    phone = item["phone"]
    text = item["text"]

    try:
        # صندوق البحث
        search = driver.find_element(
            By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'
        )
        search.clear()
        search.send_keys(phone)
        time.sleep(3)
        search.send_keys(Keys.ENTER)

        time.sleep(4)

        # صندوق الرسالة
        box = driver.find_element(
            By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'
        )
        box.send_keys(text)
        box.send_keys(Keys.ENTER)

        sent_count += 1
        print(f"✅ Sent to {phone}")

        wait_time = random.randint(DELAY_MIN, DELAY_MAX)
        time.sleep(wait_time)

    except Exception as e:
        print(f"❌ Error with {phone}: {e}")

print("🎉 Finished")
