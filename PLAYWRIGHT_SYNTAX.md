# สรุป syntax Playwright Python แบบ synchronous

ใช้เปิดดูระหว่างเขียนขั้นตอนจาก codegen โดยใช้ `playwright.sync_api` และเรียก `run(playwright, case, config)` เหมือนโปรเจกต์นี้ ไม่ใช้ async หรือ pytest

- [README: ติดตั้ง uv, codegen, Excel และวิธีรัน](README.md)
- [ตัวอย่าง SauceDemo ที่ทำงานครบ](playwrighttest.py)
- [แม่แบบสำหรับเว็บอื่น](blueprint.py)

## วิธีอ่านตัวอย่าง

ยกเว้นตัวอย่างที่ระบุว่ารันได้ทั้งไฟล์ โค้ดส่วนใหญ่เป็น **ส่วนย่อยสำหรับใส่ในฟังก์ชัน** ต้องมี `page` ที่สร้างจาก `context.new_page()` แล้ว และหน้าเว็บต้องอยู่ในขั้นตอนที่ตรงกับตัวอย่าง ไม่ควรนำทุกบล็อกมาต่อกันแล้วรันทันที

ตัวอย่างที่ใช้ `case` ต้องรับข้อมูลหนึ่งแถวจาก Excel ก่อน ชื่อ role, label, selector, URL และข้อความเป็นตัวอย่าง ให้เปลี่ยนตามเว็บไซต์จริง ส่วนตัวอย่างที่ระบุ SauceDemo ใช้กับเว็บนั้นโดยเฉพาะ

หากใช้ `expect` หรือ type annotation ให้นำเข้าก่อน:

```python
from playwright.sync_api import Page, Playwright, expect, sync_playwright
```

## 1. โครงสร้างพื้นฐาน

ตัวอย่างนี้เป็นไฟล์ Python ที่รันได้หลังติดตั้ง Playwright และ Chromium ตาม README:

```python
from playwright.sync_api import sync_playwright

with sync_playwright() as playwright:
    browser = playwright.chromium.launch(headless=False)
    try:
        context = browser.new_context()
        try:
            page = context.new_page()
            page.goto("https://example.com/")  # เปลี่ยน URL
            print(page.title())
        finally:
            context.close()
    finally:
        browser.close()
```

| ส่วน | หน้าที่ |
| --- | --- |
| `sync_playwright()` | เริ่มใช้งาน Playwright แบบ synchronous |
| `chromium.launch()` | เปิดโปรแกรม browser |
| `browser.new_context()` | สร้าง session ที่แยก cookies และ storage |
| `context.new_page()` | สร้างแท็บที่ใช้ทำงาน |
| `finally` | ปิดทรัพยากรเมื่อออกจากบล็อก แม้ขั้นตอนทดสอบเกิดข้อผิดพลาด |

ในโปรเจกต์นี้ `main()` เริ่ม Playwright และ `run()` เปิด browser/context ใหม่ทุกเคส การนำคำสั่งไปเติมให้ใช้ `page` ที่ `run()` สร้างไว้ ไม่ต้องเปิด browser ซ้ำในฟังก์ชันย่อย

URL ของงานจริงอ่านด้วย `read_config()` จาก `config.yaml` แล้ว `open_website(page, config)` เรียก `page.goto(config['base_url'])` ส่วน URL ในตัวอย่าง syntax เป็นเพียงค่าประกอบการอธิบาย

ดู [Page API](https://playwright.dev/python/docs/api/class-page) และ [การจัดการหน้าเว็บ](https://playwright.dev/python/docs/pages)

## 2. เปิดและเปลี่ยนหน้า

ต้องมี `page` แล้ว:

```python
page.goto("https://example.com/")  # เปิด URL ที่กำหนด
page.reload()                     # โหลดหน้าปัจจุบันใหม่
page.go_back()                    # ย้อนประวัติ ถ้ามีหน้าก่อนหน้า
print(page.url)                    # อ่าน URL ปัจจุบัน เป็น property ไม่มี ()
```

`goto()` รอการนำทางตามเงื่อนไขของคำสั่ง แต่ข้อมูลบางส่วนอาจโหลดต่อภายหลัง ให้ตรวจองค์ประกอบที่ต้องใช้ด้วย `expect()`

## 3. ค้นหาองค์ประกอบด้วย locator

Locator เป็นตัวระบุว่าจะทำงานกับองค์ประกอบใด การสร้าง locator อย่างเดียวยังไม่ได้คลิกหรือกรอกข้อมูล

| คำสั่ง | ใช้หาอะไร |
| --- | --- |
| `get_by_role()` | บทบาท เช่น button, link, textbox กับชื่อที่เข้าถึงได้ |
| `get_by_label()` | ช่องกรอกที่มี label เชื่อมอยู่ |
| `get_by_text()` | ข้อความที่ปรากฏในองค์ประกอบ |
| `get_by_placeholder()` | ข้อความ placeholder ของช่องกรอก |
| `get_by_test_id()` | ค่า attribute สำหรับ test ID |
| `locator()` | CSS selector หรือ selector ที่ระบุ |

ต้องมี `page` ที่เปิดเว็บซึ่งมีองค์ประกอบเหล่านี้ และเปลี่ยนชื่อตัวอย่างตามเว็บ:

```python
login_button = page.get_by_role("button", name="Login", exact=True)
username = page.get_by_label("Username", exact=True)
message = page.get_by_text("บันทึกสำเร็จ", exact=True)
search = page.get_by_placeholder("ค้นหาสินค้า", exact=True)
status = page.get_by_test_id("status")

# CSS แบบที่พบใน codegen ของ SauceDemo
username_css = page.locator('[data-test="username"]')
```

`exact=True` ใช้จับคู่ชื่อหรือข้อความให้ตรงทั้งหมด ช่วยแยกคำที่คล้ายกัน ส่วน `name` ใน `get_by_role()` คือ accessible name ซึ่งอาจมาจากข้อความหรือ `aria-label` ไม่ใช่ attribute HTML `name` เสมอไป

### data-testid กับ data-test ต่างกัน

`get_by_test_id()` ใช้ `data-testid` เป็นค่าเริ่มต้น ถ้า HTML เป็น `<div data-testid="status">` ใช้ `page.get_by_test_id("status")` ได้

แต่ตัวอย่าง SauceDemo ใช้ `data-test` จึงเลือกได้สองวิธี:

```python
# วิธีที่ไฟล์ปัจจุบันใช้: ระบุ CSS โดยตรง ไม่ต้องตั้งค่าเพิ่ม
page.locator('[data-test="login-button"]').click()
```

หรือกำหนด attribute ก่อนใช้ `get_by_test_id()` ต้องมีทั้ง `playwright` และ `page`:

```python
playwright.selectors.set_test_id_attribute("data-test")
page.get_by_test_id("login-button").click()
```

ดู [Locators](https://playwright.dev/python/docs/locators)

## 4. ทำงานกับฟอร์ม

ตัวอย่างแต่ละบรรทัดใช้แยกตามชนิดองค์ประกอบบนเว็บของคุณ ต้องมี `page` และอยู่หน้าฟอร์มแล้ว:

```python
page.get_by_label("Email").fill("user@example.com")  # แทนข้อความเดิม
page.get_by_label("Email").fill("")                  # ล้างช่องกรอก
page.get_by_role("button", name="Save", exact=True).click()
page.get_by_placeholder("Search").press("Enter")    # กดปุ่มบนช่องนี้
page.get_by_label("Accept terms").check()            # ให้ checkbox ถูกเลือก
page.get_by_label("Accept terms").uncheck()          # ยกเลิก checkbox

# สำหรับ HTML <select>: เลือกด้วย value หรือข้อความ label
page.get_by_label("Country").select_option("TH")
page.get_by_label("Country").select_option(label="Thailand")
```

`fill()` รับข้อความ ถ้าค่าเป็นตัวเลขให้แปลงด้วย `str()` ก่อน ส่วน `select_option()` ใช้กับ `<select>` ถ้าเว็บทำ dropdown เองด้วย div/button ให้กดเปิดและเลือกตัวเลือกผ่าน locator ตามเว็บ

### อัปโหลดไฟล์

ต้องมีไฟล์อยู่จริง และ locator ชี้ไปที่ `<input type="file">` ตัวอย่างเป็นเพียง syntax ไม่ได้อัปโหลดกับเว็บไซต์ใดในคู่มือนี้:

```python
from pathlib import Path

file_path = Path(__file__).with_name("sample.pdf")  # เปลี่ยนชื่อไฟล์
page.locator('input[type="file"]').set_input_files(str(file_path))
```

Path แบบนี้อ้างอิงตำแหน่งคู่กับไฟล์ Python ส่วน relative path ธรรมดาอ้างอิง working directory ที่รัน

ดู [Actions](https://playwright.dev/python/docs/input)

## 5. เลือกจากหลายรายการ

ตัวอย่างนี้สมมติว่าหน้าสินค้ามีการ์ด `[data-test="inventory-item"]` เหมือน SauceDemo:

```python
items = page.locator('[data-test="inventory-item"]')

first_item = items.first           # property ไม่มี ()
second_item = items.nth(1)          # index เริ่มที่ 0 ดังนั้น 1 คือรายการที่สอง
current_count = items.count()      # จำนวนที่พบขณะเรียก ไม่รอให้ครบตามเป้าหมาย

# กรองจากข้อความที่มีอยู่ในการ์ด เป็นการหาแบบมีข้อความนี้อยู่
backpack = items.filter(has_text="Sauce Labs Backpack")

# กรองการ์ดที่มีองค์ประกอบลูกตรงกับ locator นี้
product = items.filter(
    has=page.get_by_text("Sauce Labs Backpack", exact=True)
)
product.get_by_role("button", name="Add to cart", exact=True).click()
```

`has` ใช้ locator ขององค์ประกอบภายในการ์ด ต้องอยู่ frame เดียวกัน ถ้าต้องตรวจว่ารายการโหลดครบให้ใช้ `expect(items).to_have_count(จำนวน)` แทนการอ่าน `count()` แล้วตรวจทันที

อย่าใช้ `.first` เพื่อกลบปัญหา locator กว้างโดยไม่ได้ตั้งใจเลือกรายการแรก ควรเพิ่มเงื่อนไขให้ตรงรายการที่ต้องการ

ดู [Locator API](https://playwright.dev/python/docs/api/class-locator)

## 6. อ่านข้อความและค่าจากเว็บ

ต้องมี `page` และองค์ประกอบตามตัวอย่าง เปลี่ยน locator ตามเว็บ:

```python
heading = page.get_by_role("heading", name="Summary", exact=True)
visible_text = heading.inner_text()          # อ่าน innerText
raw_text = heading.text_content()            # อ่าน textContent อาจได้ None
email_value = page.get_by_label("Email").input_value()  # ค่าใน input
href = page.get_by_role("link", name="Help", exact=True).get_attribute("href")
print(visible_text, raw_text, email_value, href)
```

`inner_text()` สะท้อนการแสดงข้อความตาม layout ส่วน `text_content()` อาจรวมข้อความที่ซ่อนอยู่ `get_attribute()` อาจคืน `None` ถ้าไม่มี attribute นั้น ค่าที่อ่านเป็นข้อมูล ณ เวลานั้น หากต้องรอให้ข้อความเปลี่ยนเป็นค่าที่ต้องการ ให้ใช้ `expect()`

ดู [Locator API](https://playwright.dev/python/docs/api/class-locator)

## 7. ตรวจผลด้วย expect

นำเข้า `expect` ตามต้นไฟล์ ต้องมี `page` ที่เปิดหน้าทดสอบแล้ว ตัวอย่างแต่ละบรรทัดเป็นคนละสถานการณ์ ให้เลือกใช้ตามงาน:

```python
expect(page.get_by_role("heading")).to_have_text("Dashboard")  # ข้อความทั้งหมด
expect(page.get_by_role("heading")).to_contain_text("Dash")   # มีข้อความบางส่วน
expect(page.get_by_text("Success", exact=True)).to_be_visible()
expect(page.get_by_role("button", name="Save", exact=True)).to_be_enabled()
expect(page.get_by_label("Email")).to_have_value("user@example.com")
expect(page.get_by_label("Accept terms")).to_be_checked()
expect(page.locator(".result-item")).to_have_count(3)
expect(page).to_have_url("https://example.com/dashboard")

# ตรวจเงื่อนไขตรงข้าม
expect(page.get_by_text("Loading...", exact=True)).not_to_be_visible()
expect(page.get_by_label("Email")).not_to_have_value("")
expect(page.get_by_label("Accept terms")).not_to_be_checked()
```

ถ้าตรวจ URL ที่มีรูปแบบเปลี่ยนได้ ใช้ regular expression และปรับให้ตรงเว็บ:

```python
import re

expect(page).to_have_url(re.compile(r"/dashboard(?:\?.*)?$"))
```

Assertions เหล่านี้ลองตรวจซ้ำจนเงื่อนไขผ่านหรือหมดเวลา ค่าเริ่มต้นของ assertion timeout คือ 5 วินาที ปรับรายคำสั่งได้ เช่น `to_be_visible(timeout=10_000)`

ดู [Assertions](https://playwright.dev/python/docs/test-assertions)

## 8. การรอและ timeout

### Auto-wait

คำสั่งอย่าง `click()` ตรวจความพร้อมขององค์ประกอบก่อนทำงาน จึงไม่ต้องใส่ sleep ก่อนทุกคลิก แต่ไม่ได้รับประกันว่ากระบวนการหลังคลิกสำเร็จแล้ว ต้องตรวจผลที่ต้องการด้วย `expect()` ดู [Auto-waiting](https://playwright.dev/python/docs/actionability)

ต้องมี `page` แล้ว:

```python
page.wait_for_url("**/dashboard")  # รอ URL ให้ตรงรูปแบบนี้
page.locator("#results").wait_for(state="visible", timeout=10_000)
page.locator("#loading").wait_for(state="hidden", timeout=10_000)
```

หน่วยเวลาเป็นมิลลิวินาที `10_000` คือ 10 วินาที การรอ element ใช้ `visible`, `hidden`, `attached` หรือ `detached` ตามสิ่งที่ต้องการ

### is_visible() กับ expect ต่างกัน

```python
# อ่านสถานะทันที เหมาะกับการตัดสินใจจากสถานะขณะนั้น
visible_now = page.locator("#success").is_visible()

# รอให้เงื่อนไขสำเร็จ เหมาะกับการตรวจผลหลังทำขั้นตอน
expect(page.locator("#success")).to_be_visible(timeout=10_000)
```

`assert locator.is_visible()` จึงไม่ได้มีพฤติกรรมรอซ้ำแบบ `expect()` และ `is_visible()` อาจเกิดข้อผิดพลาดได้หาก locator ไม่ชัดเจน เช่นตรงหลายองค์ประกอบ

### ค่า timeout แต่ละแบบ

ตั้งหลังมี `page` และนำเข้า `expect` แล้ว:

```python
page.set_default_timeout(10_000)             # คำสั่งที่ใช้ default timeout ของ page
page.set_default_navigation_timeout(20_000)  # คำสั่งนำทาง
expect.set_options(timeout=7_000)            # ค่าเริ่มต้นของ assertions
```

ค่า assertion แยกจาก timeout ของ page ถ้าหา element ผิด การเพิ่ม timeout อย่างเดียวไม่แก้สาเหตุ ใช้การรอสถานะหรือผลที่ต้องการ แทนการกำหนด `wait_for_timeout()` เป็นการรอหลักของ test

ดู [Page API](https://playwright.dev/python/docs/api/class-page) และ [Assertions](https://playwright.dev/python/docs/test-assertions)

## 9. iframe, popup, dialog และ download

### iframe

ต้องมี iframe ตาม selector บนหน้าปัจจุบัน เปลี่ยน selector และ label ตามเว็บ:

```python
form = page.frame_locator('iframe[title="Contact form"]')
form.get_by_label("Email").fill("user@example.com")
form.get_by_role("button", name="Send", exact=True).click()
```

องค์ประกอบใน iframe ต้องค้นหาผ่าน frame นั้น ดู [Frames](https://playwright.dev/python/docs/frames)

### popup หรือแท็บใหม่

ต้องมีปุ่มที่เปิดหน้าต่างใหม่ วางคำสั่งคลิกภายในบล็อกรอ event:

```python
with page.expect_popup() as popup_info:
    page.get_by_role("button", name="Open report", exact=True).click()

popup = popup_info.value
expect(popup.get_by_role("heading", name="Report", exact=True)).to_be_visible()
popup.close()
```

ใช้ `popup` ควบคุมแท็บใหม่ ไม่ใช่ `page` เดิม การตั้งรอก่อนคลิกช่วยไม่ให้พลาด event ดู [Pages](https://playwright.dev/python/docs/pages)

### JavaScript dialog

ใช้กับ `alert`, `confirm` และ `prompt` ไม่ใช่กล่อง modal HTML ทั่วไป ต้องลงทะเบียน handler ก่อนคลิก และ handler ต้อง accept หรือ dismiss dialog:

```python
# ตัวอย่าง confirm: ยืนยัน dialog ถัดไปเพียงครั้งเดียว
page.once("dialog", lambda dialog: dialog.accept())
page.get_by_role("button", name="Delete", exact=True).click()
```

ทางเลือกสำหรับ prompt หรือการยกเลิก ใช้แยกกันตามสถานการณ์:

```python
# prompt ที่ต้องกรอกข้อความ
page.once("dialog", lambda dialog: dialog.accept("ข้อความที่ต้องการ"))
page.get_by_role("button", name="Rename", exact=True).click()
```

```python
# ยกเลิก confirm
page.once("dialog", lambda dialog: dialog.dismiss())
page.get_by_role("button", name="Delete", exact=True).click()
```

หากไม่ลงทะเบียน handler Playwright จะ dismiss dialog อัตโนมัติ แต่ถ้าลงทะเบียนแล้วไม่จัดการ dialog คำสั่งที่กระตุ้นมันอาจค้างจนหมดเวลา ดู [Dialogs](https://playwright.dev/python/docs/dialogs)

### Download

ต้องมีปุ่มที่ส่งไฟล์ดาวน์โหลด โค้ดนี้เป็นตัวอย่าง syntax ไม่ได้รันดาวน์โหลดกับเว็บภายนอก:

```python
from pathlib import Path

download_dir = Path(__file__).parent / "downloads"
download_dir.mkdir(exist_ok=True)

with page.expect_download() as download_info:
    page.get_by_role("button", name="Export", exact=True).click()

download = download_info.value
download.save_as(str(download_dir / download.suggested_filename))
```

รอ event ก่อนคลิก และบันทึกด้วย `save_as()` ก่อนปิด context หากต้องเก็บไฟล์ไว้ เพราะไฟล์ชั่วคราวของการดาวน์โหลดจะถูกลบเมื่อปิด context ดู [Downloads](https://playwright.dev/python/docs/downloads)

## 10. ช่วย debug

### บันทึก screenshot

ต้องมี `page` แล้ว ตัวอย่างนี้สร้างโฟลเดอร์คู่ไฟล์ Python:

```python
from pathlib import Path

output_dir = Path(__file__).parent / "screenshots"
output_dir.mkdir(exist_ok=True)
page.screenshot(path=str(output_dir / "current-page.png"), full_page=True)
```

ชื่อเดิมจะถูกเขียนทับเมื่อบันทึกซ้ำ ถ้าต้องเก็บแยกแต่ละเคส ให้ตั้งชื่อไฟล์ที่ไม่ซ้ำและไม่มีอักขระต้องห้ามของ Windows

### หยุดดูหน้าเว็บ

เมื่อเปิด browser แบบ `headless=False` และมี `page` แล้ว ใส่คำสั่งนี้ตรงขั้นตอนที่อยากหยุด:

```python
page.pause()  # เปิด Inspector แล้วกด Resume เพื่อทำต่อ
```

### ทำให้เห็นขั้นตอนช้าลง

ใน `run()` ให้แทนบรรทัดเปิด browser เดิม ไม่ใช่เปิด browser อีกตัว:

```python
browser = playwright.chromium.launch(headless=False, slow_mo=300)
```

`slow_mo` หน่วยมิลลิวินาที ใช้ช่วยดูการทำงาน ไม่ใช่การตรวจว่าข้อมูลโหลดพร้อม ส่วน `headless=True` ใช้รันโดยไม่แสดงหน้าต่าง ดู [Debugging](https://playwright.dev/python/docs/debug)

## 11. ใช้ร่วมกับ Excel และ run()

### เปลี่ยนค่าที่กรอกตามแถว Excel

ตัวอย่างนี้ต้องมี `page` ที่หน้า Login ของ SauceDemo และ `case` ที่อ่านจาก Excel:

```python
# ก่อน: ข้อมูลตายตัวจาก codegen
page.locator('[data-test="username"]').fill("standard_user")

# หลัง: ใช้ชื่อคอลัมน์จาก Excel
page.locator('[data-test="username"]').fill(case["username"])
```

ตัวอ่าน Excel ในโปรเจกต์แปลง `None` เป็น `""` และแปลงตัวเลขเป็นข้อความให้แล้ว แต่ต้องเพิ่มชื่อคอลัมน์ในส่วนที่อ่านข้อมูลให้ถูกต้องด้วย ดูรายละเอียดใน [README](README.md)

ชื่อเคสใช้คอลัมน์ `testcase name` ส่วน username/password เป็นคนละคอลัมน์ ในตัวอย่างปัจจุบัน `check_step_result()` ตรวจ error หลัง Login/Checkout แล้วเทียบกับ `expected_result` ตรง ๆ หากตรงจะจบเคสนั้น ไม่ต้องเพิ่มคอลัมน์ Scenario และไม่ใช้ prefix ของข้อความมาตัดสินเส้นทาง

### เพิ่มฟังก์ชันใหม่

ตัวอย่างสำหรับ SauceDemo หลัง Login สร้างฟังก์ชันนี้ก่อน `run()`:

```python
def verify_inventory(page: Page, case: dict[str, str]):
    expect(page.locator('[data-test="inventory-list"]')).to_be_visible()
```

แล้วแทรกคำสั่งเรียกใน `run()` ของ `playwrighttest.py` ระหว่าง Login กับเลือกสินค้า โดยคงส่วนเปิด–ปิด browser เดิม:

```python
open_website(page, config)
log_in(page, case)
if not check_step_result(page, case, '[data-test="inventory-list"]'):
    return
verify_inventory(page, case)  # ขั้นตอนใหม่
add_items(page, case)
fill_checkout_information(page, case)
if not check_step_result(page, case, '[data-test="checkout-summary-container"]'):
    return
verify_result(page, case)
```

ทุกแถวที่ส่งเข้า `run()` จะทำตามลำดับนี้ การเพิ่มจำนวนเคสข้อมูลแบบเดิมให้เพิ่มแถวใน Excel ส่วนการเพิ่มกระบวนการต่างกันต้องแยกเส้นทางในโค้ด ไม่จำเป็นต้องสร้างฟังก์ชันใหม่ทุกครั้งที่เพิ่มแถว

## 12. ข้อผิดพลาดที่พบบ่อย

| อาการ | สิ่งที่ควรตรวจ |
| --- | --- |
| `strict mode violation` | คำสั่งที่ต้องการ element เดียวกลับพบหลายตัว เพิ่ม `exact=True` หรือจำกัดภายในการ์ด/ฟอร์ม |
| หา element ไม่เจอหรือ timeout | URL หน้าปัจจุบัน ลำดับขั้นตอน ข้อความ locator และ iframe |
| `get_by_test_id()` ไม่เจอบน SauceDemo | เว็บใช้ `data-test` ต้องตั้ง attribute หรือใช้ CSS โดยตรง |
| คลิกแล้วตรวจไม่ผ่านทันที | ใช้ `expect()` รอผล แทน `assert is_visible()` |
| `.first()` ใช้ไม่ได้ | Locator ใช้ `.first` แบบ property ไม่มีวงเล็บ |
| เลือกรายการผิด | `nth(0)` คือรายการแรก `nth(1)` คือรายการที่สอง |
| `fill()` แจ้งชนิดข้อมูลผิด | ใช้ข้อความ แปลงตัวเลขเป็น `str()` และช่องว่างเป็น `""` |
| อ่าน `count()` แล้วได้ 0 | รายการอาจยังไม่มา ใช้ `expect(locator).to_have_count(...)` เมื่อต้องรอจำนวน |
| กดปุ่ม dialog แล้วค้าง | ถ้ามี handler ต้อง accept/dismiss และลงทะเบียนก่อนคลิก |
| รอ popup/download ไม่สำเร็จ | ตั้ง `expect_popup()` / `expect_download()` ก่อนคำสั่งที่กระตุ้น event |
| เพิ่ม `def` แล้วไม่ทำงาน | ต้องเรียกฟังก์ชันใน `run()` ด้วย |
| แม่แบบขึ้น `NotImplementedError` | เติมฟังก์ชันเฉพาะเว็บ หรือลบจุดเรียกถ้าไม่ใช้ขั้นตอนนั้น |

## ตารางเปิดดูเร็ว

ตัวแปร `locator` ด้านล่างหมายถึง locator ที่สร้างและปรับให้ตรงเว็บของคุณแล้ว

| ต้องการ | Syntax |
| --- | --- |
| เปิด URL | `page.goto(url)` |
| กรอกข้อความ | `locator.fill(text)` |
| คลิก | `locator.click()` |
| กด Enter | `locator.press("Enter")` |
| เลือก checkbox | `locator.check()` |
| ยกเลิก checkbox | `locator.uncheck()` |
| เลือกค่าใน select | `locator.select_option(value)` |
| อัปโหลด | `locator.set_input_files(path)` |
| อ่านข้อความ | `locator.inner_text()` |
| อ่านค่า input | `locator.input_value()` |
| อ่าน attribute | `locator.get_attribute("href")` |
| ตรวจข้อความ | `expect(locator).to_have_text(text)` |
| ตรวจว่าแสดง | `expect(locator).to_be_visible()` |
| ตรวจว่าไม่แสดง | `expect(locator).not_to_be_visible()` |
| ตรวจจำนวน | `expect(locator).to_have_count(number)` |
| ตรวจ URL | `expect(page).to_have_url(url)` |
| รอสถานะ | `locator.wait_for(state="visible")` |
| ภาพทั้งหน้า | `page.screenshot(path="page.png", full_page=True)` |
| หยุด debug | `page.pause()` |

## ตัวอย่าง Login สั้น ๆ โดยใช้ case

ตัวอย่างสำหรับเคสที่คาดว่า Login สำเร็จเท่านั้น ต้องมี import ด้านบน รับ `page` กับ `case` จาก `run()` และเปิด URL ผ่าน `open_website(page, config)` มาก่อน โค้ดนี้เป็นฟังก์ชัน ไม่ใช่ไฟล์ที่เปิด browser เอง:

```python
def log_in_successfully(page: Page, case: dict[str, str]):
    page.locator('[data-test="username"]').fill(case["username"])
    page.locator('[data-test="password"]').fill(case["password"])
    page.locator('[data-test="login-button"]').click()
    expect(page.locator('[data-test="inventory-list"]')).to_be_visible()
```

อย่าใช้ตัวอย่างสำเร็จอย่างเดียวนี้แทน `log_in()` ของชุดที่มี negative case เพราะจะบังคับว่าทุก Login ต้องสำเร็จ โค้ดจริงใน [playwrighttest.py](playwrighttest.py) แยกกรอก Login ออกจาก `check_step_result()` เพื่อรองรับ error และหยุดเส้นทางก่อนเลือกสินค้า ส่วน [blueprint.py](blueprint.py) เป็นแม่แบบทั่วไป ดูวิธีปรับใน [README](README.md)
