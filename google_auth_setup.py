import os
import sys
from playwright.sync_api import sync_playwright

AUTH_PATH = 'google_auth.json'

def check_session():
    if not os.path.exists(AUTH_PATH): return False
    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(storage_state=AUTH_PATH)
            page = context.new_page()
            page.goto('https://docs.google.com/spreadsheets/u/0/', wait_until='domcontentloaded', timeout=10000)
            is_logged_in = 'accounts.google.com' not in page.url
            browser.close()
            return is_logged_in
        except: return False

def save_session():
    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
            context = browser.contexts[0]
            page = context.pages[0] if context.pages else context.new_page()
            page.goto("https://www.google.com")
            context.storage_state(path=AUTH_PATH)
            browser.close()
            return True
        except Exception as e:
            print(f"Error: {e}")
            return False

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "save":
        if save_session(): sys.exit(0)
        else: sys.exit(1)
    else:
        if check_session(): sys.exit(0)
        else: sys.exit(1)