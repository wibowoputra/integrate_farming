import json, pathlib, time, sys
from playwright.sync_api import sync_playwright

BASE = pathlib.Path(r'D:\Project\integrate_planning')
OUTLINE = json.loads((BASE / 'content' / 'outline_v3.json').read_text(encoding='utf-8'))
TITLE = 'Integration Farmer - Proposal Program RT-RW'

KEYWORDS = {
    1: ['kebun sayur halaman rumah indonesia', 'urban farming pekarangan rt'],
    3: ['kompos sampah dapur', 'composting rumah tangga'],
    4: ['harga sayur pasar tradisional', 'sayur bumbu dapur'],
    5: ['panen sayur kangkung pakcoy', 'kebun pekarangan sayur'],
    6: ['gotong royong warga indonesia', 'keluarga berkebun rumah'],
    8: ['komunitas berkebun warga rt', 'kerja bakti kebun'],
    9: ['panen sayur komunitas', 'kebun bersama warga'],
    16: ['musyawarah warga rt rw', 'warga gotong royong'],
}

def run():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp('http://localhost:9222')
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
        page = ctx.new_page()
        page.goto('https://slides.google.com/create', wait_until='networkidle')
        time.sleep(4)
        # rename
        try:
            page.click('text=Presentasi tanpa judul', timeout=5000)
            page.keyboard.press('Control+A')
            page.keyboard.type(TITLE)
            page.keyboard.press('Enter')
        except:
            pass
        # delete default slides until 1? keep simple - work with current
        print('ready')
        # build step by step via JS-like clicks is complex; rely on snapshot driven
        browser.close()

if __name__ == '__main__':
    run()
