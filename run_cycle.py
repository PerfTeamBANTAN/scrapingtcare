# run_cycle.py
import time
from datetime import datetime
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from telkomcare_login import login_otomatis
from telkomcare_downloads import (
    download_report_hsi,
    download_report_datin,
    download_ttr_datin,
    download_ttr_indibiz,
    download_ttr_reseller,
    download_wecare_gaul,
)
from telkomcare_session import ensure_logged_in

BASE_DIR = Path(__file__).resolve().parent
COOKIES_ENV_PATH = BASE_DIR / "cookies.env"
DOWNLOADS_FOLDER = str(Path.home() / "Downloads")


def create_driver():
    chrome_options = Options()
    prefs = {
        "download.default_directory": DOWNLOADS_FOLDER,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--headless=new")  # penting di CI
    chrome_options.add_argument("--window-size=1280,800")

    driver = webdriver.Chrome(options=chrome_options)
    return driver


def need_fresh_login():
    if not COOKIES_ENV_PATH.exists():
        return True
    content = COOKIES_ENV_PATH.read_text(encoding="utf-8")
    for line in content.splitlines():
        if line.startswith("TC_SESSION_VALUE="):
            return line.strip() == "TC_SESSION_VALUE="
    return True


def run_import(cmd, label):
    print(f"\n=== RUN: {cmd} ({label}) ===")
    result = subprocess.run(cmd, shell=True)
    if result.returncode != 0:
        print(f"❌ Gagal menjalankan {cmd}")
        return False
    return True


def main():
    from subprocess import run as _run  # hindari import global di atas
    import subprocess

    driver = None
    first_cycle = True
    try:
        print("\n" + "=" * 70)
        print(f"⏱  START CYCLE - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 70)

        if need_fresh_login():
            print("🔑 cookies.env kosong / belum ada → login OTP otomatis...")
            driver = login_otomatis()
            if driver is None:
                raise Exception("login_otomatis gagal, driver None")
        else:
            driver = create_driver()
            print("🔑 cookies.env ada → pakai ensure_logged_in (cookie)...")
            driver = ensure_logged_in(
                driver, login_func=login_otomatis, first_cycle=first_cycle
            )

        # 2. HSI24 → WECARE HSI
        download_report_hsi(driver)
        _ = subprocess.run("py import_telkomcare_download.py", shell=True)

        # 2b. HSI24 GAUL → WECARE GAUL
        download_wecare_gaul(driver)
        _ = subprocess.run("py import_telkomcare_wecare_gaul.py", shell=True)

        # 3. DATIN24 → WECARE DATIN
        download_report_datin(driver)
        _ = subprocess.run("py import_telkomcare_wecare_datin.py", shell=True)

        # 4. TTR DATIN
        download_ttr_datin(driver)
        _ = subprocess.run("py import_telkomcare_ttr_datin.py", shell=True)

        # 5. TTR INDIBIZ
        download_ttr_indibiz(driver)
        _ = subprocess.run("py import_telkomcare_ttr_indibiz.py", shell=True)

        # 6. TTR RESELLER
        download_ttr_reseller(driver)
        _ = subprocess.run("py import_telkomcare_ttr_reseller.py", shell=True)

    finally:
        print("🧹 Menutup browser Selenium...")
        if driver is not None:
            driver.quit()


if __name__ == "__main__":
    main()
