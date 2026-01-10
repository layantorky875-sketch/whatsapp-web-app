import time
import os
import tempfile
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


def log(msg):
    print(msg)


try:
    # ========= إعداد Chrome =========
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

    wait = WebDriverWait(driver, 60)

    # ========= فتح واتساب =========
    driver.get("https://web.whatsapp.com")
    log("🟢 افتح واتساب... لو أول مرة اعمل Scan QR")

    # ✅ استنى لحد ما الحساب يفتح فعليًا (قائمة الشات)
    wait.until(
        EC.presence_of_element_located(
            (By.ID, "pane-side")
        )
    )

    log("✅ واتساب جاهز والدردشات اتحملت")

    # ========= قراءة الإكسيل =========
    file_path = "WhatsApp Business.xlsm"
    sheet_name = "Send"

    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # ========= إرسال الرسائل =========
    for i, row in df.iterrows():

        phone = str(row.iloc[0]).strip()
        name = str(row.iloc[1]).strip()
        message = str(row.iloc[2]).strip()
        sent = str(row.iloc[3]).strip()

        if sent.lower() == "sent":
            continue

        if phone == "" or message == "":
            continue

        message = message.replace("{{name}}", name)

        log(f"➡️ فتح شات {phone}")
        driver.get(f"https://web.whatsapp.com/send?phone={phone}")

        # ✅ استنى صندوق الرسالة الحقيقي
        box = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//div[@contenteditable="true" and @role="textbox"]')
            )
        )

        box.click()
        time.sleep(1)
        box.send_keys(message)
        time.sleep(1)
        box.send_keys(Keys.ENTER)

        log(f"✅ اتبعتت لـ {phone}")
        df.at[i, df.columns[3]] = "Sent"

        time.sleep(8)  # أمان

    # ========= حفظ =========
    df.to_excel(file_path, sheet_name=sheet_name, index=False)
    log("🎉 خلص الإرسال كله")

except Exception as e:
    log("❌ حصل خطأ قاتل:")
    log(str(e))

finally:
    time.sleep(5)
    try:
        driver.quit()
    except:
        pass
