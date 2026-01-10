import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ========= قراءة الاكسيل =========
file_path = "WhatsApp Business.xlsm"
sheet_name = "Send"

df = pd.read_excel(file_path, sheet_name=sheet_name)

# ========= إعداد Chrome =========
options = webdriver.ChromeOptions()
options.add_argument("--user-data-dir=chrome-data")  # يحفظ تسجيل الدخول

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

driver.get("https://web.whatsapp.com")
print("📱 اعمل Scan QR أول مرة فقط")
time.sleep(20)

# ========= إرسال الرسائل =========
for i, row in df.iterrows():

    phone = str(row["Phone"]).strip()
    name = str(row["Name"]).strip()
    message = str(row["Message"]).strip()
    sent = str(row["Sent"]).strip()

    # لو الرسالة اتبعت قبل كده
    if sent.lower() == "sent":
        continue

    if phone == "" or message == "":
        continue

    # استبدال الاسم
    message = message.replace("{{name}}", name)

    url = f"https://web.whatsapp.com/send?phone={phone}"
    driver.get(url)

    time.sleep(10)

    try:
        msg_box = driver.find_element(
            By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'
        )
        msg_box.click()
        msg_box.send_keys(message)
        msg_box.send_keys(Keys.ENTER)

        print(f"✅ تم الإرسال إلى {phone}")

        # نعلّم Sent في الاكسيل
        df.at[i, "Sent"] = "Sent"

        time.sleep(6)  # أمان من الحظر

    except Exception:
        print(f"❌ فشل الإرسال إلى {phone}")
        continue

# ========= حفظ التحديث =========
df.to_excel(file_path, sheet_name=sheet_name, index=False)

print("🎉 تم الانتهاء")

