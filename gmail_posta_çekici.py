import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException

chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driver = webdriver.Chrome(options=chrome_options)

wait = WebDriverWait(driver, 10)

print("Gmail açılıyor...")
driver.get("https://mail.google.com/mail/u/0/#inbox")

# Oturum açılana kadar bekleme
while True:
    try:
        # Inbox satırlarını kontrol et
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tr.zA")))
        print("Giriş yapıldı, inbox yüklendi.")
        break
    except TimeoutException:
        print("Henüz giriş yapılmadı. Lütfen Gmail hesabınızla oturum açın...")
        time.sleep(3)  # Bekle, sonra tekrar dene

time.sleep(1)  # Inbox render için kısa bekleme

mail_data = []

rows = driver.find_elements(By.CSS_SELECTOR, "tr.zA")
print(f"{len(rows)} mail bulundu.")

for i in range(len(rows)):
    try:
        rows = driver.find_elements(By.CSS_SELECTOR, "tr.zA")
        row = rows[i]

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", row)
        driver.execute_script("arguments[0].click();", row)

        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.adn")))

        # Subject
        try:
            subject = driver.find_element(By.CSS_SELECTOR, "h2.hP").text
        except:
            subject = ""

        # Sender
        try:
            sender_elem = driver.find_element(By.CSS_SELECTOR, "span.gD")
            sender = sender_elem.get_attribute("email") or sender_elem.text
        except:
            sender = ""

        # Date
        try:
            date = driver.find_element(By.CSS_SELECTOR, "span.g3").get_attribute("title")
        except:
            date = ""

        # Body
        try:
            body_elem = driver.find_element(By.CSS_SELECTOR, "div.adn div.a3s")
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

        driver.execute_script("window.history.go(-1)")
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tr.zA")))

    except StaleElementReferenceException:
        driver.execute_script("window.history.go(-1)")
        continue
    except Exception as e:
        print("Hata:", e)
        driver.execute_script("window.history.go(-1)")
        continue

df = pd.DataFrame(mail_data)
df.to_excel("gmail_tüm_mailler.xlsx", index=False)
print("\nBitti: gmail_tüm_mailler.xlsx")
input("Enter ile çık.")
