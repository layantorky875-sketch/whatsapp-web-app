import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# 1) فتح Chrome
options = webdriver.ChromeOptions()
options.add_argument("--user-data-dir=chrome-data")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

# 2) فتح واتساب ويب
driver.get("https://web.whatsapp.com")
print("افتح واتساب واعمل QR Login لو لسه")

time.sleep(25)  # استنى التحميل

# 3) الرقم والرسالة (تغييرهم لاحقًا)
phone = "201001234567"
message = "مرحبا، هذه رسالة تجريبية"

# 4) البحث عن الرقم
search = driver.find_element(By.XPATH, '//div[@contenteditable="true"]')
search.send_keys(phone)
time.sleep(3)
search.send_keys(Keys.ENTER)

time.sleep(4)

# 5) كتابة الرسالة
box = driver.find_element(By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')
box.send_keys(message)
box.send_keys(Keys.ENTER)

print("تم الإرسال")
