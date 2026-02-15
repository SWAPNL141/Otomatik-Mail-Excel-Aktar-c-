import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

# -------------------------
# Chrome debug bağlantısı
# -------------------------
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 10)

# Outlook açılıyor
print("Outlook açılıyor...")
driver.get("https://outlook.office.com/mail/inbox")

# Oturum açılana kadar bekle
while True:
    try:
        # Mail listesi yüklenmiş mi kontrol et
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='row']")))
        print("Giriş yapıldı, inbox yüklendi.")
        break
    except TimeoutException:
        print("Henüz giriş yapılmadı. Lütfen Outlook hesabınızla oturum açın...")
        time.sleep(3)

time.sleep(1)

mail_data = []

# Ana pencereyi sakla
main_window = driver.current_window_handle

# Inbox satırlarını al
rows = driver.find_elements(By.CSS_SELECTOR, "div[role='row']")
print(f"{len(rows)} mail bulundu.")

for i in range(len(rows)):
    try:
        rows = driver.find_elements(By.CSS_SELECTOR, "div[role='row']")
        row = rows[i]

        # Mail linkini al (Outlook mail row genelde data-id içerir)
        try:
            mail_url = row.get_attribute("data-id")
            if not mail_url:
                # Direkt click ile aç
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", row)
                driver.execute_script("arguments[0].click();", row)
                time.sleep(2)
            else:
                # Yeni sekmede aç
                driver.execute_script(f"window.open('https://outlook.office.com/mail/deeplink/{mail_url}', '_blank');")
                driver.switch_to.window(driver.window_handles[-1])
        except:
            driver.execute_script("arguments[0].click();", row)
            time.sleep(2)

        # Mail yüklenene kadar bekle
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[aria-label='Message body']")))

        # Başlık
        try:
            subject = driver.find_element(By.CSS_SELECTOR, "div[role='heading']").text
        except:
            subject = ""

        # Gönderen
        try:
            sender = driver.find_element(By.CSS_SELECTOR, "span[title]").get_attribute("title")
        except:
            sender = ""

        # Tarih
        try:
            date = driver.find_element(By.CSS_SELECTOR, "div[aria-label*='Received']").text
        except:
            date = ""

        # Body
        try:
            body_elem = driver.find_element(By.CSS_SELECTOR, "div[aria-label='Message body']")
            body = body_elem.text
        except:
            body = ""

        mail_data.append({
            "Gönderen": sender,
            "Başlık": subject,
            "Tarih": date,
            "İçerik": body[:3000]
        })

        print(f"{i+1} - {sender}")

        # Sekmeyi kapat ve ana pencereye dön
        if len(driver.window_handles) > 1:
            driver.close()
            driver.switch_to.window(main_window)
        time.sleep(1)

    except StaleElementReferenceException:
        driver.switch_to.window(main_window)
        continue
    except Exception as e:
        print("Hata:", e)
        driver.switch_to.window(main_window)
        continue

# Excel yaz
df = pd.DataFrame(mail_data)
df.to_excel("outlook_tüm_mailler.xlsx", index=False)
print("\nBitti: outlook_tüm_mailler.xlsx")
input("Enter ile çık.")
