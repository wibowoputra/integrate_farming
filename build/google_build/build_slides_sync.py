import json, pathlib, time
from playwright.sync_api import sync_playwright

BASE = pathlib.Path(r'D:\Project\integrate_planning')
OUTLINE = json.loads((BASE / 'content' / 'outline_v3.json').read_text(encoding='utf-8'))
TITLE = 'Integration Farmer - Proposal Program RT-RW'

KEYWORDS = {
    1: 'kebun sayur halaman rumah indonesia',
    3: 'kompos sampah dapur',
    4: 'harga sayur pasar tradisional',
    5: 'panen sayur kangkung pakcoy',
    6: 'gotong royong warga indonesia',
    8: 'komunitas berkebun warga rt',
    9: 'panen sayur komunitas',
    16: 'musyawarah warga rt rw',
}

LAYOUTS = {
    'cover': 'Slide judul',
    'exec_summary': 'Judul dan isi',
    'content_stats': 'Judul dan isi',
    'content_image': 'Judul dan isi',
    'pillars': 'Judul dan isi',
    'photo_captions': 'Judul dan isi',
    'two_column': 'Judul dan dua kolom',
    'examples': 'Judul dan isi',
    'examples_with_banner': 'Judul dan isi',
    'timeline': 'Judul dan isi',
    'table': 'Judul dan isi',
    'sponsors': 'Judul dan isi',
    'benefits': 'Judul dan isi',
    'risks': 'Judul dan isi',
    'indicators': 'Judul dan isi',
    'closing': 'Slide judul',
}

def text_content(s, d):
    parts = []
    if d.get('headline'): parts.append(d['headline'])
    if d.get('kicker'): parts.append(d['kicker'])
    if d.get('tagline'): parts.append(d['tagline'])
    if d.get('subtitle'): parts.append(d['subtitle'])
    if d.get('bullets'):
        parts.extend(d['bullets'])
    if d.get('cards'):
        for c in d['cards']:
            if c.get('stat'): parts.append(c['stat'])
            if c.get('title'): parts.append(c['title'])
            if c.get('label'): parts.append(c['label'])
            if c.get('bullets'): parts.extend(c['bullets'])
    if d.get('captions'): parts.extend(d['captions'])
    if d.get('left_title'): parts.append(d['left_title'])
    if d.get('left_bullets'): parts.extend(d['left_bullets'])
    if d.get('right_title'): parts.append(d['right_title'])
    if d.get('right_bullets'): parts.extend(d['right_bullets'])
    if d.get('split_note'): parts.append(d['split_note'])
    if d.get('phases'):
        for ph in d['phases']:
            if ph.get('period'): parts.append(ph['period'])
            if ph.get('title'): parts.append(ph['title'])
            if ph.get('items'): parts.extend(ph['items'])
    if d.get('table'):
        t = d['table']
        parts.extend(t.get('headers', []))
        for row in t.get('rows', []):
            parts.extend(row)
        if t.get('total_label'): parts.append(t['total_label'])
        if t.get('total_value'): parts.append(t['total_value'])
        if t.get('total_note'): parts.append(t['total_note'])
    if d.get('note_line'): parts.append(d['note_line'])
    if d.get('columns'):
        for col in d['columns']:
            parts.append(col.get('title',''))
            parts.extend(col.get('items',[]))
    if d.get('cta_bullets'): parts.extend(d['cta_bullets'])
    if d.get('closing_line'): parts.append(d['closing_line'])
    if d.get('contact'): parts.append(d['contact'])
    return '\n'.join(str(x) for x in parts if x)
def set_slide_layout(page, layout_name):
    # Open layout menu
    page.locator('[role="button"][aria-label*="Tata letak"], button:has-text("Tata letak")').first().click()
    time.sleep(0.2)
    # Click layout by text
    candidates = [layout_name, layout_name.replace('dan', '&'), layout_name.replace('&', 'dan')]
    for c in candidates:
        try:
            page.getByRole('menuitem', { name: c }).click()
            time.sleep(0.2)
            return True
        except Exception:
            pass
    # fallback text
    try:
        page.locator(f'text="{layout_name}"').first().click(); time.sleep(0.2); return True
    except Exception:
        return False

def click_title(page):
    try:
        page.locator('div[role="textbox"]:has-text("Klik - tambahkan judul")').first().click()
    except:
        try:
            page.keyboard.press('Tab')
        except:
            pass

def click_body(page):
    try:
        page.locator('div[role="textbox"]:has-text("Klik - tambahkan teks")').first().click()
    except:
        try:
            page.keyboard.press('Tab'); page.keyboard.press('Tab')
        except:
            pass

def clear_and_type(page, text):
    if not text: return
    page.keyboard.press('Control+A')
    page.keyboard.type(text)

def add_image_from_search(page, keyword):
    if not keyword: return
    # Click "Stok dan web" or media
    try:
        page.locator('button[aria-label*="Stok"], button:has-text("Stok dan web")').first().click()
    except:
        pass
    time.sleep(0.5)
    # Search box
    try:
        search = page.locator('input[aria-label="Telusuri gambar"], [role="searchbox"] input').first()
        search.click(); search.fill(keyword); search.press('Enter')
    except Exception as e:
        pass
    time.sleep(1.5)
    # Insert first result
    try:
        page.locator('img[role="img"], [data-image-url]').first().click()
    except:
        pass
    try:
        page.getByRole('button', { name: 'Sisipkan', exact=False }).click()
    except:
        pass
    time.sleep(1.0)
def build():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp('http://localhost:9222')
        page = browser.contexts[0].pages[-1] if browser.contexts and browser.contexts[0].pages else None
        if not page:
            page = browser.new_page()
        # Navigate to presentation if needed
        if 'docs.google.com/presentation' not in page.url:
            page.goto('https://docs.google.com/presentation/d/155QoaQZfPkAJbxqmv2yMODrSIjpyn1zOe0MEYWo12JU/edit')
            time.sleep(2)
        # Rename title
        try:
            page.locator('textbox[aria-label="Ganti nama"], .docs-title-input').first().click()
            page.keyboard.press('Control+A')
            page.keyboard.type(TITLE)
            page.keyboard.press('Enter')
            time.sleep(0.5)
        except:
            pass
        # Ensure we have at least 1 slide, delete extras conceptually - but just build by adding
        slides = OUTLINE['slides']
        # Work with current slide 1
        for i, s in enumerate(slides):
            d = s
            lay = LAYOUTS.get(d['layout'], 'Judul dan isi')
            if i == 0:
                # set layout of first
                set_slide_layout(page, lay)
            else:
                # new slide
                page.locator('button[aria-label*="Slide baru"]').first().click()
                time.sleep(0.5)
                set_slide_layout(page, lay)
            time.sleep(0.3)
            kw = KEYWORDS.get(d['id'])
            # Fill title/headline area - usually first textbox
            try:
                # click title placeholder or first editable
                page.locator('div[aria-label*="Klik - tambahkan judul"], div[aria-label*="judul"]').first().click()
                page.keyboard.press('Control+A')
                txt = d.get('headline') or d.get('kicker') or d.get('tagline') or d.get('title') or ''
                if txt:
                    page.keyboard.type(txt)
            except:
                pass
            time.sleep(0.1)
            # Fill body/content
            try:
                page.locator('div[aria-label*="Klik - tambahkan teks"], div[aria-label*="isi"], div[aria-label*="deskripsi"]').first().click()
            except:
                pass
            page.keyboard.press('Control+A')
            body = text_content(d, d)
            if body:
                page.keyboard.type(body[:6000])  # truncate safety
            # Add image if keyword exists and not cover/closing
            if kw and d['layout'] not in ('cover', 'closing'):
                try:
                    add_image_from_search(page, kw)
                except:
                    pass
            time.sleep(0.5)
        print('done', len(slides))

if __name__ == '__main__':
    build()
