import time
import os
import tempfile
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


# ========= إعداد Chrome (ثابت بدون كراش) =========
profile_path = os.path.join(tempfile.gettempdir(), "wa_engine_profile")

options = webdriver.ChromeOptions()
options.add_argument(f"--user-data-dir={profile_path}")
options.add_argument("--profile-directory=Default")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

# ========= فتح واتساب =========
driver.get("https://web.whatsapp.com")
print("🟢 لو أول مرة: اعمل Scan QR")
time.sleep(25)

# ========= قراءة الإكسيل =========
file_path = "WhatsApp Business.xlsm"
sheet_name = "Send"

df = pd.read_excel(file_path, sheet_name=sheet_name)

# ========= إرسال الرسائل =========
for i, row in df.iterrows():

    phone = str(row["Phone"]).strip()
    name = str(row["Name"]).strip()
    message = str(row["Message"]).strip()
    sent = str(row["Sent"]).strip()

    if sent.lower() == "sent":
        continue

    if phone == "" or message == "":
        continue

    message = message.replace("{{name}}", name)

    driver.get(f"https://web.whatsapp.com/send?phone={phone}")
    time.sleep(10)

    try:
        box = driver.find_element(
            By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'
        )
        box.click()
        box.send_keys(message)
        box.send_keys(Keys.ENTER)

        print(f"✅ Sent to {phone}")
        df.at[i, "Sent"] = "Sent"

        time.sleep(7)  # أمان

    except Exception as e:
        print(f"❌ Failed: {phone}")
        continue

# ========= حفظ التحديث =========
df.to_excel(file_path, sheet_name=sheet_name, index=False)

print("🎉 انتهى الإرسال")
time.sleep(5)
driver.quit()
