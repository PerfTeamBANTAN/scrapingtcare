# telkomcare_session.py
import time
from pathlib import Path

from selenium.common.exceptions import WebDriverException

BASE_DIR = Path(__file__).resolve().parent
COOKIES_ENV_PATH = BASE_DIR / "cookies.env"

SESSION_COOKIE_NAME = "newtelkomcareapache"
BASE_DOMAIN = "telkomcare.telkom.co.id"


def save_session_cookie_from_driver(driver):
    """
    Ambil cookie session dari driver (setelah login sukses),
    lalu simpan ke cookies.env:
    TC_SESSION_NAME=...
    TC_BASE_DOMAIN=...
    TC_SESSION_VALUE=...
    """
    cookie = driver.get_cookie(SESSION_COOKIE_NAME)
    if not cookie:
        raise RuntimeError(f"Cookie '{SESSION_COOKIE_NAME}' tidak ditemukan di driver.")

    value = cookie.get("value", "")
    if not value:
        raise RuntimeError("Cookie session value kosong.")

    lines = [
        f"TC_SESSION_NAME={SESSION_COOKIE_NAME}",
        f"TC_BASE_DOMAIN={BASE_DOMAIN}",
        f"TC_SESSION_VALUE={value}",
        "",
    ]
    COOKIES_ENV_PATH.write_text("\n".join(lines), encoding="utf-8")
    print(f"💾 Session cookie disimpan ke {COOKIES_ENV_PATH}")


def load_session_from_env():
    """
    Baca cookies.env (kalau ada).
    Return: (name, domain, value) — value bisa "" kalau belum ada.
    """
    name = SESSION_COOKIE_NAME
    domain = BASE_DOMAIN
    value = ""

    if COOKIES_ENV_PATH.exists():
        for line in COOKIES_ENV_PATH.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or "=" not in line:
                continue
            k, v = line.split("=", 1)
            k = k.strip().upper()
            v = v.strip()
            if k == "TC_SESSION_NAME":
                name = v
            elif k == "TC_BASE_DOMAIN":
                domain = v
            elif k == "TC_SESSION_VALUE":
                value = v

    return name, domain, value


def ensure_logged_in(driver, login_func, first_cycle=False):
    """
    Coba pakai cookie dari cookies.env.
    Kalau masih mentok di /public/login -> jalankan login_func()
    (telkomcare_login.login_otomatis) untuk login ulang + refresh cookies.env.

    Return:
        - driver lama (kalau cookie masih valid)
        - driver baru hasil login_func() kalau cookie expired/invalid atau driver lama error
    """
    name, domain, value = load_session_from_env()

    # 1) Buka halaman login supaya domain match
    try:
        driver.get("https://telkomcare.telkom.co.id/public/login?&modules=assurance")
    except WebDriverException:
        print("⚠️ Driver lama error/mati, pakai login otomatis langsung...")
        return login_func()

    time.sleep(2)

    # 2) Kalau ada value -> inject cookie
    if value:
        driver.add_cookie(
            {
                "name": name,
                "value": value,
                "domain": domain,
                "path": "/",
                "secure": True,
                "httpOnly": False,
                "sameSite": "Lax",
            }
        )
        driver.get(
            "https://telkomcare.telkom.co.id/assurance/dashboard/alertresponse"
        )
        time.sleep(3)
    else:
        # Tidak ada value sama sekali di cookies.env → langsung login otomatis
        print("⚠️ TC_SESSION_VALUE kosong → login otomatis...")
        try:
            driver.quit()
        except Exception:
            pass
        return login_func()

    cur = (driver.current_url or "").lower()
    if "public/login" in cur or "modules=assurance" in cur:
        print("⚠️ Cookie expired/invalid → jalankan login otomatis...")
        # Tutup driver lama, karena login_otomatis bikin driver baru sendiri
        try:
            driver.quit()
        except Exception:
            pass

        new_driver = login_func()
        return new_driver
    else:
        if first_cycle:
            print("✅ Cookie valid (cycle pertama), sudah di dashboard.")
        else:
            print("✅ Cookie masih valid, sudah di dashboard.")
        return driver
