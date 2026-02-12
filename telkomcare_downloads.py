# telkomcare_downloads.py

import os
import sys
import time
from pathlib import Path
from datetime import date

from dotenv import load_dotenv

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ===================== LOAD ENV & KONSTAN =====================

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

dotenv_path = BASE_DIR / "cookies.env"
print("DEBUG dotenv_path:", dotenv_path)
print("EXISTS:", dotenv_path.exists())

load_dotenv(dotenv_path)

print("DEBUG ENV RAW:")
print("TC_SESSION_NAME =", os.environ.get("TC_SESSION_NAME"))
print("TC_BASE_DOMAIN  =", os.environ.get("TC_BASE_DOMAIN"))
print("TC_SESSION_VALUE=", os.environ.get("TC_SESSION_VALUE"))

DOWNLOADS_FOLDER = str(Path.home() / "Downloads")

SESSION_COOKIE_NAME = os.getenv("TC_SESSION_NAME", "newtelkomcareapache")
SESSION_COOKIE_VALUE = os.getenv("TC_SESSION_VALUE", "")
BASE_DOMAIN = os.getenv("TC_BASE_DOMAIN", "telkomcare.telkom.co.id")

print("DEBUG ENV:")
print("TC_SESSION_NAME =", SESSION_COOKIE_NAME)
print("TC_BASE_DOMAIN  =", BASE_DOMAIN)
print("TC_SESSION_VALUE length =", len(SESSION_COOKIE_VALUE))


# ===================== COOKIE SESSION TELKOMCARE =====================

def inject_session_cookie(driver):
    """
    Set cookie session TelkomCare berdasarkan cookies.env.
    """
    print("🔐 Inject session cookie TelkomCare...")
    print(f"   COOKIE NAME   = {SESSION_COOKIE_NAME}")
    print(f"   BASE DOMAIN   = {BASE_DOMAIN}")
    print(f"   COOKIE VALUE length = {len(SESSION_COOKIE_VALUE)}")

    if not SESSION_COOKIE_VALUE:
        raise RuntimeError(
            "TC_SESSION_VALUE kosong (env tidak terbaca / belum di-set). "
            "Cek lokasi cookies.env & isinya."
        )

    driver.get("https://" + BASE_DOMAIN)
    driver.delete_all_cookies()

    driver.add_cookie(
        {
            "name": SESSION_COOKIE_NAME,
            "value": SESSION_COOKIE_VALUE,
            "path": "/",
            "secure": True,
        }
    )

    print("   -> COOKIE DI DRIVER:", driver.get_cookie(SESSION_COOKIE_NAME))


def ensure_logged_in(driver, first_cycle: bool):
    """
    Login via cookie. Dipanggil dari run_all.py setiap cycle.
    """
    inject_session_cookie(driver)

    test_url = "https://telkomcare.telkom.co.id/assurance/dashboard/alertresponse"
    print(f"🔗 Test akses URL: {test_url}")
    driver.get(test_url)

    time.sleep(3)
    current = driver.current_url
    print(f"🔍 current_url: {current}")
    if "login" in current.lower():
        raise Exception("Cookie session tidak valid, masih di halaman login.")
    else:
        print("✅ Session TelkomCare valid (sudah login).")


# ===================== HELPER DOWNLOAD =====================

def wait_for_new_download(before_files, timeout=180):
    print("⏳ Menunggu file download baru (.xls/.xlsx)...")
    start = time.time()
    while time.time() - start < timeout:
        now_files = {
            f
            for f in os.listdir(DOWNLOADS_FOLDER)
            if f.lower().endswith((".xls", ".xlsx"))
        }
        new_files = now_files - before_files
        if new_files:
            latest = max(
                new_files,
                key=lambda f: os.path.getmtime(os.path.join(DOWNLOADS_FOLDER, f)),
            )
            full_path = os.path.join(DOWNLOADS_FOLDER, latest)
            print(f"✓ File baru terdeteksi: {full_path}")
            return full_path
        time.sleep(2)
    raise TimeoutError("Timeout menunggu file download TelkomCare")


def wait_download_complete(path: Path = Path.home() / "Downloads", timeout: int = 300):
    print("⏳ Menunggu proses download selesai (cek *.crdownload)...")
    start = time.time()
    while time.time() - start < timeout:
        crs = list(path.glob("*.crdownload"))
        if not crs:
            print("✅ Tidak ada .crdownload, download selesai.")
            return
        time.sleep(1)
    raise TimeoutError("Download belum selesai setelah timeout.")


# ===================== HSI24: STEP-BY-STEP (WECARESUGAR26) =====================

def download_report_hsi(driver):
    """
    HSI24 - Flow step-by-step:
    - Buka wecaresugar26?sumber=HSI24
    - Pilih TELKOMBARU
    - Klik SUBMIT
    - Tunggu tabel HSI muncul
    - Ambil link detailsugar25 sumber=HSI24, REGIONAL2, BANTEN, kategori=grand_total
    - Buka detail GRAND TOTAL (tanpa xls=1)
    - Di halaman detail, klik link download (detailsugar25 ... xls=1)
    """

    print("\n" + "=" * 70)
    print("⬇️ DOWNLOAD HSI24 (Step-by-step, detailsugar25)")
    print("=" * 70)

    before_files = {
        f
        for f in os.listdir(DOWNLOADS_FOLDER)
        if f.lower().endswith((".xls", ".xlsx"))
    }

    # 1. Buka halaman WECARE HSI
    wecaresugar_url = (
        "https://telkomcare.telkom.co.id/assurance/lapebis26/wecaresugar26?sumber=HSI24"
    )
    print(f"\n1️⃣ Buka halaman: {wecaresugar_url}")
    driver.get(wecaresugar_url)

    # 2. Pilih teritori TELKOM BARU
    print("2️⃣ Select teritori TELKOM BARU...")
    try:
        teritori_select = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.ID, "param_teritory"))
        )
        teritori_select.click()

        option_telkombaru = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//option[@value='TELKOMBARU']"))
        )
        option_telkombaru.click()
        print("   ✓ TELKOM BARU dipilih")
    except Exception as e:
        print(f"   ❌ Error select teritori: {e}")
        return None

    # 3. Klik SUBMIT
    print("3️⃣ Klik tombol SUBMIT (tanggal default hari ini)...")
    try:
        submit_btn = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'SUBMIT')]")
            )
        )
        submit_btn.click()
        print("   ✓ SUBMIT diklik, menunggu data tabel HSI...")
    except Exception as e:
        print(f"   ❌ Error klik SUBMIT: {e}")
        return None

    # 4. Tunggu tabel HSI muncul
    print("4️⃣ Tunggu tabel HSI muncul...")
    try:
        WebDriverWait(driver, 300).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "table tbody tr")
            )
        )
        print("   ✓ Minimal 1 baris data HSI terdeteksi.")
    except Exception as e:
        print(f"   ❌ Tabel HSI tidak muncul: {e}")
        return None

    # 5. Ambil link detailsugar25 GRAND TOTAL (tanpa xls=1)
    print("5️⃣ Ambil link GRAND TOTAL (detailsugar25, sumber=HSI24, REGIONAL2 BANTEN)...")
    try:
        data_link = WebDriverWait(driver, 180).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//a["
                    "contains(@href, 'detailsugar25') "
                    "and contains(@href, 'sumber=HSI24') "
                    "and contains(@href, 'regional=REGIONAL2') "
                    "and contains(@href, 'witel=BANTEN') "
                    "and contains(@href, 'kategori=grand_total')"
                    "]",
                )
            )
        )
        href = data_link.get_attribute("href") or ""
        print(f"   ✓ Link href GRAND TOTAL ditemukan: {href}")

        if href.startswith("/"):
            href = "https://telkomcare.telkom.co.id" + href
            print(f"   ✓ URL absolute GRAND TOTAL: {href}")
    except Exception as e:
        print(f"   ❌ Error ambil link GRAND TOTAL HSI: {e}")
        print("   🔍 Debug semua link detailsugar25 HSI24 di halaman...")
        try:
            all_links = driver.find_elements(By.XPATH, "//a[@href]")
            for idx, link in enumerate(all_links):
                href_attr = link.get_attribute("href") or ""
                if "detailsugar25" in href_attr and "HSI24" in href_attr:
                    print(f"      Link {idx}: {href_attr}")
        except Exception:
            pass
        return None

    # 6. Buka halaman detail GRAND TOTAL (tanpa xls=1)
    print("6️⃣ Buka halaman detail GRAND TOTAL HSI (tanpa xls=1)...")
    if "xls=1" in href:
        href_no_xls = href.split("xls=1")[0].rstrip("&?")
    else:
        href_no_xls = href
    print(f"   🔗 URL detail GRAND TOTAL HSI: {href_no_xls}")
    driver.get(href_no_xls)

    try:
        WebDriverWait(driver, 300).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "table tbody tr")
            )
        )
        print("   ✓ Tabel detail GRAND TOTAL HSI sudah terisi.")
    except Exception as e:
        print(f"   ❌ Tabel detail GRAND TOTAL HSI tidak muncul: {e}")
        return None

    # 7. Cari link download XLS di halaman detail
    print("7️⃣ Cari link download Excel HSI (href mengandung xls=1)...")
    try:
        download_link = WebDriverWait(driver, 120).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//a["
                    "contains(@href, 'detailsugar25') "
                    "and contains(@href, 'xls=1') "
                    "and contains(@href, 'sumber=HSI24') "
                    "and contains(@href, 'kategori=grand_total')"
                    "]",
                )
            )
        )
        dl_href = download_link.get_attribute("href") or ""
        print(f"   ✓ Link download HSI ditemukan: {dl_href}")

        if dl_href.startswith("/"):
            dl_href = "https://telkomcare.telkom.co.id" + dl_href
            print(f"   ✓ URL absolute download HSI: {dl_href}")
    except Exception as e:
        print(f"   ❌ Tidak menemukan link download HSI: {e}")
        return None

    # 8. Trigger download
    print("8️⃣ Trigger download HSI24 GRAND TOTAL...")
    driver.get(dl_href)

    downloaded_file = wait_for_new_download(before_files)
    wait_download_complete(Path(DOWNLOADS_FOLDER))
    print("✅ Download HSI24 selesai!")
    return downloaded_file

# ===================== WECARE GAUL (FOLLOW HSI PAGE) =====================

def download_wecare_gaul(driver):
    """
    WECARE GAUL (HSI24) - direct URL:
    - Tidak lagi cari link di halaman.
    - Langsung ke detailsugar25?xls=1 dengan param:
      sumber=HSI24&regional=REGIONAL2&witel=BANTEN&kategori=gaul&param_teritory=TELKOMBARU&enddate=YYYY-MM-DD
    """
    print("\n" + "=" * 70)
    print("⬇️ DOWNLOAD WECARE GAUL (direct URL xls=1)")
    print("=" * 70)

    before_files = {
        f
        for f in os.listdir(DOWNLOADS_FOLDER)
        if f.lower().endswith((".xls", ".xlsx"))
    }

    # Tanggal hari ini dalam format YYYY-MM-DD
    today_str = date.today().strftime("%Y-%m-%d")
    base_url = "https://telkomcare.telkom.co.id/assurance/lapebis25/detailsugar25"

    download_url = (
        f"{base_url}"
        f"?xls=1"
        f"&read=all"
        f"&param_teritory=TELKOMBARU"
        f"&enddate={today_str}"
        f"&tahun="
        f"&bulan="
        f"&sumber=HSI24"
        f"&tiket="
        f"&regional=REGIONAL2"
        f"&witel=BANTEN"
        f"&kategori=gaul"
    )

    print(f"1️⃣ Download URL GAUL: {download_url}")
    print("2️⃣ Trigger download GAUL...")
    driver.get(download_url)

    downloaded_file = wait_for_new_download(before_files)
    wait_download_complete(Path(DOWNLOADS_FOLDER))
    print(f"✅ Download WECARE GAUL selesai: {downloaded_file}")
    return downloaded_file

# ===================== DOWNLOAD DATIN (Direct URL, sudah OK) =====================

def download_report_datin(driver):
    """
    DATIN24 - Flow step-by-step:
    Mencari link dengan parameter:
    sumber=DATIN24&tiket=&regional=REGIONAL2&witel=BANTEN&kategori=grand_total
    """

    print("\n" + "=" * 70)
    print("⬇️ DOWNLOAD DATIN24 (Step-by-step)")
    print("=" * 70)

    before_files = {
        f
        for f in os.listdir(DOWNLOADS_FOLDER)
        if f.lower().endswith(".xls") or f.lower().endswith(".xlsx")
    }

    wecaresugar_url = (
        "https://telkomcare.telkom.co.id/assurance/lapebis25/wecaresugar25?sumber=DATIN24"
    )
    print(f"\n1️⃣ Buka halaman: {wecaresugar_url}")
    driver.get(wecaresugar_url)
    time.sleep(4)

    print("2️⃣ Select teritori TELKOM BARU...")
    try:
        teritori_select = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "param_teritory"))
        )
        teritori_select.click()
        time.sleep(1)

        option_telkombaru = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//option[@value='TELKOMBARU']"))
        )
        option_telkombaru.click()
        print("   ✓ TELKOM BARU dipilih")
    except Exception as e:
        print(f"   ❌ Error select teritori: {e}")
        raise

    time.sleep(2)

    print("3️⃣ Klik tombol SUBMIT (tanggal default hari ini)...")
    try:
        submit_btn = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'SUBMIT')]")
            )
        )
        submit_btn.click()
        print("   ✓ SUBMIT diklik, menunggu data tabel DATIN...")
    except Exception as e:
        print(f"   ❌ Error klik SUBMIT: {e}")
        raise

    print("4️⃣ Tunggu tabel muncul & ambil link data href...")
    time.sleep(8)  # boleh Anda ganti dengan WebDriverWait table tbody tr
    try:
        data_link = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//a["
                    "contains(@href, 'detailsugar25') "
                    "and contains(@href, 'sumber=DATIN24') "
                    "and contains(@href, 'regional=REGIONAL2') "
                    "and contains(@href, 'witel=BANTEN') "
                    "and contains(@href, 'kategori=grand_total')"
                    "]",
                )
            )
        )
        href = data_link.get_attribute("href")
        print(f"   ✓ Link href ditemukan: {href}")

        if href.startswith("/"):
            href = "https://telkomcare.telkom.co.id" + href
            print(f"   ✓ URL absolute: {href}")
    except Exception as e:
        print(f"   ❌ Error ambil link data: {e}")
        print(f"   🔍 Debug: Mencari semua link yang ada...")
        try:
            all_links = driver.find_elements(By.XPATH, "//a[@href]")
            print(f"   📋 Total link ditemukan: {len(all_links)}")
            for idx, link in enumerate(all_links):
                href_attr = link.get_attribute("href")
                if "detailsugar25" in href_attr:
                    print(f"      Link {idx}: {href_attr}")
        except Exception:
            pass
        raise

    time.sleep(1)

    print("5️⃣ Navigate ke detail & download xls=1...")
    if "xls=1" not in href:
        download_url = href + ("&" if "?" in href else "?") + "xls=1"
    else:
        download_url = href

    print(f"   📥 Download URL: {download_url}")
    driver.get(download_url)

    wait_for_new_download(before_files)
    wait_download_complete(Path(DOWNLOADS_FOLDER))
    print("✅ Download DATIN24 selesai!")

# ===================== DOWNLOAD TTR (DETAILRESCOMP25) =====================

def _prepare_before_files():
    return {
        f
        for f in os.listdir(DOWNLOADS_FOLDER)
        if f.lower().endswith((".xls", ".xlsx"))
    }


def _get_start_end_today():
    """
    Helper untuk dapatkan startdate = tanggal 1 bulan ini,
    enddate = hari ini (YYYY-MM-DD).
    """
    today = date.today()
    startdate = today.replace(day=1).strftime("%Y-%m-%d")
    enddate = today.strftime("%Y-%m-%d")
    return startdate, enddate


def download_ttr_datin(driver):
    """
    Download TTR DATIN (detailrescomp25, sumber=DATIN24, tiket=TELKOMGAMAS).
    Periode: dari tanggal 1 bulan ini sampai hari ini.
    """
    print("\n" + "=" * 70)
    print("⬇️ DOWNLOAD TTR DATIN (detailrescomp25)")
    print("=" * 70)

    before_files = _prepare_before_files()
    startdate, enddate = _get_start_end_today()

    download_url = (
        "https://telkomcare.telkom.co.id/assurance/lapebis25/detailrescomp25"
        "?xls=1"
        "&read=all"
        "&param_teritory=TELKOMBARU"
        "&tahun="
        "&bulan="
        "&sumber=DATIN24"
        "&tiket=TELKOMGAMAS"
        f"&startdate={startdate}"
        f"&enddate={enddate}"
        "&custpending="
        "&regional=REGIONAL%202"
        "&kategori="
        "&tcomp="
    )
    print(f"   📥 Download URL: {download_url}")
    driver.get(download_url)

    downloaded_file = wait_for_new_download(before_files)
    wait_download_complete(Path(DOWNLOADS_FOLDER))
    print("✅ Download TTR DATIN selesai!")
    return downloaded_file


def download_ttr_indibiz(driver):
    """
    Download TTR INDIBIZ (detailrescomp25, sumber=INDIBIZ, tiket=TELKOMGAMAS).
    Periode: dari tanggal 1 bulan ini sampai hari ini.
    """
    print("\n" + "=" * 70)
    print("⬇️ DOWNLOAD TTR INDIBIZ (detailrescomp25)")
    print("=" * 70)

    before_files = _prepare_before_files()
    startdate, enddate = _get_start_end_today()

    download_url = (
        "https://telkomcare.telkom.co.id/assurance/lapebis25/detailrescomp25"
        "?xls=1"
        "&read=all"
        "&param_teritory=TELKOMBARU"
        "&tahun="
        "&bulan="
        "&sumber=INDIBIZ"
        "&tiket=TELKOMGAMAS"
        f"&startdate={startdate}"
        f"&enddate={enddate}"
        "&custpending="
        "&regional=REGIONAL%202"
        "&kategori="
        "&tcomp="
    )
    print(f"   📥 Download URL: {download_url}")
    driver.get(download_url)

    downloaded_file = wait_for_new_download(before_files)
    wait_download_complete(Path(DOWNLOADS_FOLDER))
    print("✅ Download TTR INDIBIZ selesai!")
    return downloaded_file


def download_ttr_reseller(driver):
    """
    Download TTR RESELLER (detailrescomp25, sumber=RESELLER, tiket=TELKOMGAMAS).
    Periode: dari tanggal 1 bulan ini sampai hari ini.
    """
    print("\n" + "=" * 70)
    print("⬇️ DOWNLOAD TTR RESELLER (detailrescomp25)")
    print("=" * 70)

    before_files = _prepare_before_files()
    startdate, enddate = _get_start_end_today()

    download_url = (
        "https://telkomcare.telkom.co.id/assurance/lapebis25/detailrescomp25"
        "?xls=1"
        "&read=all"
        "&param_teritory=TELKOMBARU"
        "&tahun="
        "&bulan="
        "&sumber=RESELLER"
        "&tiket=TELKOMGAMAS"
        f"&startdate={startdate}"
        f"&enddate={enddate}"
        "&custpending="
        "&regional=REGIONAL%202"
        "&kategori="
        "&tcomp="
    )
    print(f"   📥 Download URL: {download_url}")
    driver.get(download_url)

    downloaded_file = wait_for_new_download(before_files)
    wait_download_complete(Path(DOWNLOADS_FOLDER))
    print("✅ Download TTR RESELLER selesai!")
    return downloaded_file
