# telkomcare_login.py
import os
import warnings
import logging

os.environ["KMP_DUPLICATE_LIB_OK"] = "True"
warnings.filterwarnings("ignore", message=".*pin_memory.*")
warnings.filterwarnings("ignore", message=".*CUDA not available.*")

import time
import pyotp
import easyocr
import cv2
import numpy as np

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, TimeoutException

from telkomcare_session import save_session_cookie_from_driver  # <=== penting

# ==================== LOGGING & ENV ====================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler("telkomcare_login.log"), logging.StreamHandler()],
)

USERNAME = os.getenv("TELKOM_USERNAME", "19950145")
PASSWORD = os.getenv("TELKOM_PASSWORD", "Sla2020")
TOTP_SECRET = os.getenv("TELKOM_TOTP_SECRET", "2VGL6TXQ2IJ42YPU")

# EasyOCR untuk bahasa Indonesia
reader = easyocr.Reader(["id"], gpu=False)


# ==================== DRIVER SETUP ====================


def setup_driver():
    chrome_prefs = {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "profile.password_manager_leak_detection": False,
    }
    options = Options()
    options.add_experimental_option("prefs", chrome_prefs)
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--window-size=1280,800")
    # options.add_argument("--headless")  # aktifkan jika mau headless

    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


# ==================== CAPTCHA OCR ====================


def solve_captcha_ai(element):
    """Baca CAPTCHA Telkomcare sebagai lowercase+angka (murni OCR)."""
    img_bytes = element.screenshot_as_png
    nparr = np.frombuffer(img_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_GRAYSCALE)

    img = cv2.resize(img, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_LINEAR)
    img = cv2.GaussianBlur(img, (3, 3), 0)

    candidates = []

    # Config 1: lowercase + digit
    r1 = reader.readtext(img, detail=0, allowlist="0123456789abcdefghijklmnopqrstuvwxyz")
    if r1:
        candidates.append("".join(r1))

    # Config 2: digit only
    r2 = reader.readtext(img, detail=0, allowlist="0123456789")
    if r2:
        candidates.append("".join(r2))

    # Config 3: full bebas
    r3 = reader.readtext(img, detail=0)
    if r3:
        candidates.append("".join(r3))

    best = ""
    for raw in candidates:
        teks = raw.lower().replace(" ", "")
        teks = "".join(ch for ch in teks if ch.isalnum())
        if len(teks) > len(best):
            best = teks

    if best:
        logging.info(f"CAPTCHA OCR -> '{best}'")
        return best

    logging.warning("CAPTCHA OCR: no text detected")
    return ""


# ==================== UTILITAS UI ====================


def close_password_manager_popup(driver):
    """Tutup popup password manager Chrome jika muncul."""
    try:
        buttons = driver.find_elements(By.TAG_NAME, "button")
        for btn in buttons:
            txt = (btn.text or "").strip().lower()
            if any(word in txt for word in ("ok", "got it", "oke", "dismiss")):
                driver.execute_script("arguments[0].click();", btn)
                logging.info(f"Closed password manager popup: {btn.text}")
                time.sleep(1)
                return
    except Exception:
        pass


def check_login_success(driver, wait: WebDriverWait):
    """Cek apakah sudah masuk dashboard / home Telkomcare."""
    try:
        wait.until(
            EC.any_of(
                EC.presence_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        ".dashboard-header, .user-profile, [data-user], .navbar-menu, .main-content",
                    )
                ),
                EC.url_contains("dashboard"),
                EC.url_contains("/homenew/home"),
            )
        )
        logging.info("Dashboard detected -> login success")
        return True
    except TimeoutException:
        return False


# ==================== MAIN LOGIN FLOW ====================


def login_otomatis():
    """
    Login Telkomcare full otomatis:
    - Buka Chrome
    - Isi username/password
    - OCR CAPTCHA (loop sampai tembus)
    - Isi OTP (TOTP) dan handle tombol Verify OTP
    - Simpan cookie session ke cookies.env
    - Mengembalikan driver yang sudah login (atau None jika gagal total)
    """
    driver = None

    try:
        driver = setup_driver()
        wait = WebDriverWait(driver, 20)
        logging.info("=== Telkomcare – FULL AUTO LOGIN START ===")

        driver.get("https://telkomcare.telkom.co.id")

        # ---------- LOGIN + CAPTCHA ----------
        attempt = 0
        while True:
            attempt += 1
            logging.info(f"[CAPTCHA LOOP] Attempt #{attempt}")

            wait.until(EC.presence_of_element_located((By.ID, "uname")))

            uname_el = driver.find_element(By.ID, "uname")
            pass_el = driver.find_element(By.ID, "passw")
            uname_el.clear()
            uname_el.send_keys(USERNAME)
            pass_el.clear()
            pass_el.send_keys(PASSWORD)

            captcha_img = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#captcha-element img"))
            )

            teks_captcha = solve_captcha_ai(captcha_img)
            if not teks_captcha:
                logging.warning("OCR empty, refresh page & retry")
                driver.refresh()
                time.sleep(2)
                continue

            captcha_input = driver.find_element(By.ID, "captcha-input")
            captcha_input.clear()
            captcha_input.send_keys(teks_captcha)

            agree_cb = driver.find_element(By.ID, "agree")
            driver.execute_script("arguments[0].click();", agree_cb)

            submit_btn = driver.find_element(By.ID, "submit")
            driver.execute_script("arguments[0].click();", submit_btn)

            time.sleep(3)

            # Kalau masih ada field captcha-input -> gagal
            if driver.find_elements(By.ID, "captcha-input"):
                logging.warning(f"Server reject CAPTCHA '{teks_captcha}', retry loop")
                time.sleep(1.5)
                continue

            logging.info("CAPTCHA accepted! Proceed to OTP stage.")
            break

        # ---------- OTP ----------
        logging.info("Waiting OTP page...")
        time.sleep(3)
        close_password_manager_popup(driver)

        # Coba ambil field OTP; jika tidak ada tapi sudah di dashboard -> sukses
        try:
            otp_boxes = wait.until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "div.otp-field input[name='otp[]']")
                )
            )
        except TimeoutException:
            if check_login_success(driver, wait):
                logging.info("Tidak menemukan field OTP, tapi sudah di dashboard. Login sukses.")
                save_session_cookie_from_driver(driver)
                return driver
            else:
                raise

        if len(otp_boxes) < 6:
            if check_login_success(driver, wait):
                logging.info("OTP fields < 6 tapi sudah di dashboard, anggap login sukses.")
                save_session_cookie_from_driver(driver)
                return driver
            raise Exception(f"OTP fields < 6: {len(otp_boxes)}")

        otp_first = otp_boxes[0]

        def find_otp_button():
            """Cari tombol Verify OTP secara fleksibel; jika tidak ada tapi sudah login, return None."""
            buttons = driver.find_elements(By.TAG_NAME, "button")
            cand = [b for b in buttons if b.is_displayed()]

            for b in cand:
                txt = (b.text or "").strip().lower()
                if "verify" in txt and "otp" in txt:
                    return b

            if cand:
                return cand[-1]

            inputs = driver.find_elements(
                By.CSS_SELECTOR, "input[type='submit'], input[type='button']"
            )
            cand2 = [i for i in inputs if i.is_displayed()]
            if cand2:
                return cand2[-1]

            if check_login_success(driver, wait):
                logging.info("Tidak menemukan tombol OTP tapi sudah di dashboard, anggap login sukses.")
                return None

            raise NoSuchElementException("Tidak menemukan tombol Verify OTP")

        def wait_otp_button_enabled():
            btn = find_otp_button()
            if btn is None:
                return None
            WebDriverWait(driver, 10).until(
                lambda d: btn.is_enabled() and btn.is_displayed()
            )
            return btn

        max_otp_retries = 5
        for retry in range(1, max_otp_retries + 1):
            driver.execute_script("arguments[0].focus();", otp_first)
            time.sleep(0.5)

            token = pyotp.TOTP(TOTP_SECRET).now()[-6:]
            logging.info(f"[OTP] Attempt {retry}: {token}")

            for i, digit in enumerate(token):
                box = otp_boxes[i]
                driver.execute_script("arguments[0].scrollIntoView(true);", box)
                box.clear()
                box.send_keys(digit)
                time.sleep(0.1)

            otp_btn = wait_otp_button_enabled()
            if otp_btn is None:
                logging.info("OTP button None -> dianggap sudah login, return driver.")
                save_session_cookie_from_driver(driver)
                return driver

            driver.execute_script("arguments[0].scrollIntoView(true);", otp_btn)
            driver.execute_script("arguments[0].click();", otp_btn)

            time.sleep(5)
            if check_login_success(driver, wait):
                logging.info("=== FULL LOGIN SUCCESS (CAPTCHA + OTP AUTO) ===")
                save_session_cookie_from_driver(driver)
                return driver

            logging.warning("OTP rejected / page not changed, wait next TOTP window")
            time.sleep(25)

        raise Exception("OTP failed after auto retries")

    except Exception as e:
        logging.error(f"Automation error: {e}")
        if driver:
            try:
                wait = WebDriverWait(driver, 5)
                if check_login_success(driver, wait):
                    logging.info("Error terjadi tetapi sudah di dashboard, kembalikan driver.")
                    save_session_cookie_from_driver(driver)
                    return driver
            except Exception:
                pass
            driver.quit()
        return None
