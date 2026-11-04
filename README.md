# คู่มือ Playwright Python + uv + Excel

คู่มือนี้สำหรับ Windows PowerShell เริ่มตั้งแต่ติดตั้งเครื่องมือ บันทึกขั้นตอนด้วย codegen แล้วนำมาอ่านข้อมูลจาก Excel เพื่อรันทีละเคส เหมาะสำหรับคนที่ต้องการเห็นลำดับ `for case → run(playwright, case, config)` ชัดเจน

ใช้ Python แบบเรียกไฟล์โดยตรง ไม่ต้องติดตั้ง pytest สำหรับโครงสร้างนี้

## 1. เข้าใจไฟล์ในโปรเจกต์

```text
testplaywrightforfinal/
├── .venv/               Python และแพ็กเกจเฉพาะโปรเจกต์
├── playwrighttest.py    ตัวอย่างที่ทำงานครบสำหรับ SauceDemo
├── blueprint.py         แม่แบบสำหรับนำไปเติมขั้นตอนของเว็บไซต์อื่น
├── config.yaml          URL เริ่มต้นของเว็บไซต์
├── requirements.txt     รายการแพ็กเกจที่ใช้
├── test_data.xlsx       ข้อมูลแต่ละ test case
└── README.md            คู่มือนี้
```

ไฟล์ที่จะสร้างเพิ่มในขั้นตอนถัดไปคือ `recorded_flow.py` สำหรับเก็บโค้ดต้นฉบับจาก codegen และ `test_my_website.py` สำหรับงานใหม่ที่คัดลอกจากแม่แบบ

ถ้าต้องการลองโปรเจกต์ที่มีอยู่แล้วและติดตั้งเครื่องมือครบ ไปที่หัวข้อ 6 ได้เลย ส่วนหัวข้อ 2 ใช้เตรียมโปรเจกต์ใหม่ ไม่ต้องสร้าง `.venv` ซ้ำทุกครั้ง

## 2. ติดตั้งครั้งแรก

### 2.1 เตรียม Python และเปิด Terminal ในโฟลเดอร์โปรเจกต์

ต้องติดตั้ง Python ก่อน และเรียกคำสั่งนี้ได้ ตัวอย่างในคู่มือนี้ใช้ Python 3.10 ขึ้นไป:

```powershell
python --version
```

ถ้ายังไม่มี Python ให้ติดตั้งจาก [เว็บไซต์ Python](https://www.python.org/downloads/windows/) แล้วเปิด Terminal ใหม่

เปิดโฟลเดอร์โปรเจกต์ใน editor แล้วเปิด Terminal หรือเปลี่ยนโฟลเดอร์ใน PowerShell:

```powershell
Set-Location -LiteralPath 'D:\ไนพา\testing\testplaywrightforfinal'
```

ถ้าเป็นงานใหม่ ให้ใช้โฟลเดอร์งานของตัวเองแทนตำแหน่งข้างบน

### 2.2 ติดตั้ง uv

```powershell
python -m pip install uv
uv --version
```

`python -m pip install uv` ใช้ pip ของ Python ตัวที่เลือกมาติดตั้ง uv เทียบได้กับ `pip install uv` แต่ระบุชัดว่าใช้ Python ตัวไหน ส่วน uv ใช้สร้าง venv และจัดการแพ็กเกจต่อจากนี้

ผลที่ควรเห็น: คำสั่งสุดท้ายแสดงเวอร์ชัน uv

### 2.3 สร้างและเปิดใช้ venv

```powershell
uv venv
.\.venv\Scripts\Activate.ps1
```

`uv venv` สร้างโฟลเดอร์ `.venv` ส่วนคำสั่ง Activate ทำให้ `python` ใน Terminal นี้ชี้ไปยัง Python ของโปรเจกต์

ตรวจให้แน่ใจด้วย:

```powershell
python -c "import sys; print(sys.executable)"
```

ผลที่ควรเห็น: ตำแหน่งลงท้ายด้วย `.venv\Scripts\python.exe` ของโปรเจกต์นี้

### 2.4 ติดตั้งแพ็กเกจและ browser

```powershell
uv pip install playwright
uv pip install openpyxl
uv pip install pyyaml
python -m playwright install chromium
```

| คำสั่ง | หน้าที่ |
| --- | --- |
| `uv pip install playwright` | ติดตั้งไลบรารีควบคุม browser และเครื่องมือ codegen |
| `uv pip install openpyxl` | ติดตั้งไลบรารีอ่าน Excel `.xlsx` |
| `uv pip install pyyaml` | ติดตั้ง PyYAML เพื่ออ่าน `config.yaml` (import ด้วยชื่อ `yaml`) |
| `python -m playwright install chromium` | ดาวน์โหลด Chromium ที่ Playwright ใช้เปิดทดสอบ |

การติดตั้งแพ็กเกจกับการดาวน์โหลด browser เป็นคนละขั้นตอน ถ้าเพิ่งอัปเดต Playwright และแจ้งว่าไม่พบ browser ให้รันคำสั่งติดตั้ง Chromium อีกครั้ง

ตรวจแพ็กเกจ:

```powershell
python -c "import playwright, openpyxl, yaml; print('Packages ready')"
```

ถ้า PowerShell ไม่ยอมเปิด Activate สามารถใช้ Python ใน venv โดยตรงได้ ไม่ต้องเปลี่ยน execution policy:

```powershell
.\.venv\Scripts\python.exe -m playwright install chromium
```

คำสั่ง `python ...` ในหัวข้อถัดไปสามารถแทน `python` ด้วย `.\.venv\Scripts\python.exe` ได้เช่นกัน

### 2.5 ติดตั้งจากรายการแพ็กเกจและตั้งค่า URL

หากใช้ไฟล์โปรเจกต์นี้ ติดตั้งแพ็กเกจทั้งหมดพร้อมกันแทนคำสั่งแยกได้:

```powershell
uv pip install -r requirements.txt
```

ไฟล์ `config.yaml` อยู่คู่กับ Python และ Excel:

```yaml
base_url: "https://www.saucedemo.com/"
```

`main()` เรียก `read_config()` ครั้งเดียว แล้วส่ง config เข้า `run(playwright, case, config)` เพื่อให้ `open_website()` เปิด URL จาก `base_url` หากเปลี่ยนเว็บต้องปรับ locator และขั้นตอนให้ตรงด้วย

## 3. บันทึกขั้นตอนด้วย codegen

หลัง activate venv แล้ว ให้เริ่มบันทึก:

```powershell
python -m playwright codegen --target python -o recorded_flow.py https://www.saucedemo.com/
```

| ส่วนของคำสั่ง | ความหมาย |
| --- | --- |
| `codegen` | เปิด browser และ Inspector เพื่อบันทึกการใช้งาน |
| `--target python` | สร้างโค้ด Python แบบ synchronous ที่ใช้กับแม่แบบนี้ |
| `-o recorded_flow.py` | บันทึกโค้ดลงไฟล์นี้ |
| URL ท้ายคำสั่ง | เว็บไซต์ที่จะบันทึก เปลี่ยนเป็นเว็บของคุณได้ |

ทำตามขั้นตอนจริงใน browser เช่น Login เลือกสินค้า กรอกฟอร์ม และส่งข้อมูล เมื่อบันทึกเสร็จให้ปิดหน้าต่างของ codegen แล้วตรวจโค้ดใน `recorded_flow.py`

ลองรันโค้ดที่บันทึก:

```powershell
python recorded_flow.py
```

ผลที่ควรเห็น: browser ทำตามขั้นตอนที่บันทึกโดยใช้ข้อมูลชุดเดิม โค้ดยังไม่ได้อ่าน Excel

ใช้ชื่อ `recorded_flow.py` สำหรับต้นฉบับ อย่าสั่ง codegen เขียนทับ `playwrighttest.py` หรือไฟล์ที่ปรับอ่าน Excel แล้ว เพราะส่วนที่เพิ่มเองจะหายไป

## 4. เตรียม Excel

ตัวอย่างที่ใช้งานได้ในโปรเจกต์นี้อ่าน `test_data.xlsx` ชีตชื่อ `test_data` โดยแถวแรกต้องเป็นหัวคอลัมน์:

```text
testcase name | username | password | items_to_add | first_name | last_name | postal_code | expected_result
```

### ตัวอย่างข้อมูลสองแถว

| คอลัมน์ | เคสซื้อสำเร็จ | เคสเว้นชื่อจริง |
| --- | --- | --- |
| `testcase name` | `TC01_Purchase_Multiple_Items` | `TC02_Empty_First_Name` |
| `username` | `standard_user` | `standard_user` |
| `password` | `secret_sauce` | `secret_sauce` |
| `items_to_add` | `Sauce Labs Backpack, Sauce Labs Bike Light, Sauce Labs Bolt T-Shirt` | `Sauce Labs Backpack` |
| `first_name` | `Suda` | เว้นเซลล์ว่างจริง |
| `last_name` | `Meesuk` | `Jaidee` |
| `postal_code` | `123` | `10110` |
| `expected_result` | `Thank you for your order!` | `Error: First Name is required` |

ตารางข้างบนแสดงแต่ละเคสเป็นคอลัมน์เพื่อให้อ่านง่าย แต่ใน Excel ให้เก็บหนึ่งเคสต่อหนึ่งแถวข้อมูล และคอลัมน์เป็นชื่อฟิลด์ตามหัวตาราง

ชุดข้อมูลมี **2 เคสเดิม** ตามที่เลือกไว้:

- Scenario 1: TC01 ซื้อสินค้า 3 รายการ กรอกข้อมูลครบ กด Finish ใน Overview และตรวจข้อความสำเร็จ
- Scenario 2: TC02 เว้น first_name แล้วตรวจ `Error: First Name is required` โดยไม่กด Finish

โค้ดยังรองรับ error ตอน Login เช่นผู้ใช้ถูกล็อกหรือรหัสผ่านผิด หากคุณเพิ่มแถวเหล่านั้นภายหลัง แต่ไม่ได้เพิ่มแถวทดสอบเหล่านั้นในชุดปัจจุบัน ตัวอ่านข้อมูลรองรับชื่อคอลัมน์เก่า `Test Cases` เป็น alias ของ `testcase name` สำหรับไฟล์เก่าด้วย ส่วนแม่แบบใหม่ใช้ `testcase name` โดยตรง

ข้อที่ต้องเตรียมให้ตรง:

- ชื่อไฟล์ ชีต และหัวคอลัมน์ต้องตรงกับโค้ด รวมตัวพิมพ์และช่องว่าง
- หนึ่งแถวข้อมูลที่ไม่ว่างเป็นหนึ่งเคส ไม่ต้องกำหนดจำนวนเคสใน Python
- `items_to_add` ของตัวอย่างใช้จุลภาค `,` คั่นชื่อสินค้า โค้ดจะเพิ่มทีละรายการ
- ช่องที่ตั้งใจทดสอบว่าไม่กรอก ให้เว้นเซลล์ว่าง ไม่ต้องพิมพ์ `None`
- ข้อมูลประเภทเลขรหัส เช่น `00123` ให้ตั้งเซลล์เป็น Text ก่อนกรอก เพื่อเก็บศูนย์นำหน้า
- ใช้ค่าธรรมดาสำหรับข้อมูลทดสอบได้ง่ายที่สุด `data_only=True` อ่านค่าของสูตรที่บันทึกไว้ แต่ openpyxl ไม่คำนวณสูตรใหม่
- บันทึก Excel ก่อนรันทดสอบ โปรแกรมอ่านไฟล์บนดิสก์ ไม่ได้อ่านสิ่งที่ยังไม่บันทึกในหน้าต่าง Excel

## 5. เชื่อมโค้ดจาก codegen กับข้อมูล Excel

### 5.1 เปลี่ยนค่าตายตัวเป็นข้อมูลของเคส

ก่อนแก้:

```python
page.locator('[data-test="username"]').fill("standard_user")
```

หลังแก้:

```python
page.locator('[data-test="username"]').fill(case["username"])
```

`case` คือ dictionary ของหนึ่งแถว ตัวอย่าง:

```python
case = {
    "testcase name": "TC01",
    "username": "standard_user",
    "first_name": "Suda",
}
```

`read_test_cases()` อ่านทุกแถวเก็บใน list ชื่อ `cases` แล้ว `main()` ส่งทีละ `case` พร้อม config เข้า `run(playwright, case, config)` ส่วนฟังก์ชันย่อยรับ `case` เพื่อดึงค่าที่ต้องใช้

### 5.2 ย้ายเฉพาะคำสั่งทำงานกับ page

โค้ดที่ codegen สร้างมีทั้งส่วนเริ่ม Playwright เปิด browser และขั้นตอนทำงาน เมื่อนำเข้าแม่แบบ ให้ย้ายคำสั่ง `page.goto()`, `.fill()`, `.click()` และคำสั่งอื่นของหน้าเว็บไปตามฟังก์ชัน

ใช้ส่วนเปิด–ปิด browser และ `sync_playwright()` ของแม่แบบ ไม่ต้องนำส่วนเปิด browser หรือส่วนเรียก `run()` ท้ายไฟล์จาก codegen มาซ้ำอีก

ใน `playwrighttest.py` ตัวอย่างแบ่งเป็น:

```text
open_website()               เปิด URL จาก config.yaml
log_in()                     กรอก Login
check_step_result()          รอหน้าถัดไปหรือ error แล้วตรวจข้อความเมื่อเกิด error
add_items()                  เลือกสินค้าจาก Excel
fill_checkout_information()  เปิด Checkout และกรอกข้อมูล
verify_result()              ตรวจผลที่คาดหวัง
```

### 5.3 ตรวจผลด้วย expect

การทำขั้นตอนครบไม่ได้ยืนยันว่าผลถูกต้อง ต้องเพิ่มการตรวจ เช่น:

```python
expect(page.locator('[data-test="complete-header"]')).to_have_text(
    case["expected_result"]
)
```

ตัวอย่างนี้ใช้ locator ของ SauceDemo ถ้าเปลี่ยนเว็บให้เปลี่ยนตามข้อความหรือองค์ประกอบจริงบนเว็บนั้น `expect()` จะรอตาม timeout และแจ้งข้อผิดพลาดเมื่อผลยังไม่ตรง

ตัวอย่างปัจจุบันตรวจผลจริงหลัง Login และหลัง Continue ผ่าน `check_step_result()` โดยรอหน้าถัดไปหรือ error หากพบ error จะเทียบข้อความทั้งหมดกับ `expected_result` แล้วจบเคสนั้น ถ้าไม่มี error จะทำต่อจน Overview กด Finish แล้วตรวจข้อความสำเร็จ จึงรองรับข้อความ Login ที่ขึ้นต้น `Epic sadface:` และข้อความ Checkout ที่ขึ้นต้น `Error:` โดยไม่ตัดสินจาก prefix

ถ้า error ไม่ตรงกับ Excel จะเป็น FAIL หากคาดหวัง error แต่เว็บกลับซื้อสำเร็จ ข้อความสำเร็จก็จะไม่ตรง expected_result และเป็น FAIL เช่นกัน คำสั่งที่ timeout หรือหา element ไม่เจอไม่ถือเป็น validation ที่คาดหวัง

## 6. รันตัวอย่างที่มีอยู่

ถ้า activate venv แล้ว:

```powershell
python playwrighttest.py
```

หรือเรียก Python ใน venv โดยตรง:

```powershell
.\.venv\Scripts\python.exe playwrighttest.py
```

ถ้ากด Run ใน VS Code ให้ใช้ `Python: Select Interpreter` แล้วเลือก `.venv\Scripts\python.exe` ของโปรเจกต์ก่อน ปุ่ม Run Python ใช้ interpreter ของ editor ซึ่งอาจต่างจาก Terminal

ลำดับทำงาน:

```text
if __name__ == "__main__"
  → main()
  → read_config() อ่าน URL และ read_test_cases() อ่าน Excel ทั้งหมด
  → เริ่ม Playwright
  → for case in cases
      → run(playwright, case, config)
          → เปิด browser / context / page ใหม่
          → เปิด URL / Login / ตรวจ error
          → เพิ่มสินค้า / Checkout / ตรวจ error
          → ถ้าไม่มี error ให้ยืนยันใน Overview
          → ตรวจผลด้วย expect()
          → ปิด context และ browser ด้วย finally
      → แสดง PASS หรือ FAIL
  → ปิด Playwright
  → สรุป TOTAL / PASS / FAIL
```

แต่ละเคสรอเคสก่อนหน้าจบ ไม่ได้รันพร้อมกัน เคสที่เกิด error จะถูก `main()` จับไว้แล้วเริ่มเคสถัดไป ส่วนข้อมูล Excel ที่ผิดรูปแบบจะถูกแจ้งก่อนเริ่มรันเคส

ตัวอย่างผลเมื่อข้อมูลและเว็บไซต์ตรงกัน:

```text
RUN: TC01_Purchase_Multiple_Items
PASS: TC01_Purchase_Multiple_Items

RUN: TC02_Empty_First_Name
PASS: TC02_Empty_First_Name

TOTAL: 2 | PASS: 2 | FAIL: 0
```

เคสเว้นชื่อจริงเป็น PASS ได้ เพราะผลที่ต้องการคือเว็บแจ้งว่าต้องกรอกชื่อ ไม่ได้แปลว่าทุกเคสต้องสั่งซื้อสำเร็จ

สรุปนี้มาจากโค้ดใน `main()` ไม่ใช่ Terminal สรุปเอง `SystemExit(main())` ส่ง exit code เป็น 0 เมื่อผ่านทุกเคส หรือ 1 เมื่อมีเคสล้มเหลว

## 7. ใช้ blueprint กับเว็บอื่น

### 7.1 คัดลอกแม่แบบ

ใช้ชื่อใหม่ที่ยังไม่มีในโฟลเดอร์:

```powershell
Copy-Item -LiteralPath 'blueprint.py' -Destination 'test_my_website.py'
```

เปิดแก้ `test_my_website.py` เพื่อเก็บ `blueprint.py` ไว้ใช้ครั้งต่อไป ค้นหา `[แก้ตรงนี้` และ `[TODO` เพื่อไปยังส่วนที่ต้องเติม

### 7.2 แก้แหล่งข้อมูลด้านบนไฟล์

ตัวอย่างการตั้งค่า สำหรับเว็บที่ต้องใช้ username/password:

```python
EXCEL_FILE = "my_test_data.xlsx"
SHEET_NAME = "Cases"
REQUIRED_COLUMNS = {"testcase name", "username", "password", "expected_result"}
NONEMPTY_COLUMNS = {"testcase name"}
```

สร้าง Excel ให้มีชีตและหัวคอลัมน์ตามนี้:

- `REQUIRED_COLUMNS` หมายถึงคอลัมน์ที่ต้องมี ไม่ได้บังคับว่าทุกเซลล์ต้องมีค่า
- `NONEMPTY_COLUMNS` หมายถึงคอลัมน์ที่ห้ามว่างในทุกเคส และต้องมีในหัวตารางด้วย
- ถ้าจะทดสอบ username ว่าง อย่าใส่ `username` ใน `NONEMPTY_COLUMNS`
- แม่แบบอ่านทุกคอลัมน์ที่มีหัวตาราง เก็บเป็น `case['ชื่อคอลัมน์']` ได้เลย หัวคอลัมน์ต้องเป็นข้อความ ไม่ว่าง และไม่ซ้ำกัน

### 7.3 เติมฟังก์ชันเฉพาะเว็บ

| ฟังก์ชัน | สิ่งที่เติม |
| --- | --- |
| `open_website()` | อ่าน URL จาก config ให้แล้ว เปลี่ยน `base_url` ใน `config.yaml` |
| `log_in()` | locator ของหน้า Login และข้อมูลจาก `case` |
| `perform_actions()` | ขั้นตอนทำงานจาก codegen เช่นค้นหาหรือกรอกฟอร์ม |
| `verify_result()` | `expect()` ตรวจข้อความ สถานะ จำนวน หรือ URL |

ฟังก์ชัน `log_in()`, `perform_actions()` และ `verify_result()` ยังมี `raise NotImplementedError(...)` ให้แทนบรรทัดนั้นด้วยโค้ดจริง หากยังไม่เติมแล้วรัน จะรายงาน FAIL พร้อมบอกส่วนที่ยังขาด ส่วน `open_website()` อ่าน URL จาก config ให้แล้ว

ส่วนที่ใช้ร่วมกันพร้อมแล้วคือ `read_config()`, `open_website()`, `read_test_cases()` การจัดการ browser ใน `run()` และการวนเคสใน `main()` ปรับเฉพาะการตั้งค่าและลำดับขั้นตอนตามงานได้

### 7.4 ปรับลำดับใน run()

ภายใน `run()` มีลำดับเริ่มต้น:

```python
open_website(page, config)
log_in(page, case)
perform_actions(page, case)
verify_result(page, case)
```

ถ้าเว็บไม่ต้อง Login ให้เอา `log_in(page, case)` ออก ถ้าเพิ่มขั้นตอนใหม่ ให้สร้างฟังก์ชันแล้วแทรกการเรียกตามลำดับ โดยรักษาโครง `try/finally` ปิด browser ไว้

แม่แบบยังเป็นโครงสำหรับเว็บทั่วไป ต้องเติมการตรวจ error และแยกเส้นทางเองตามเว็บ ดูตัวอย่าง `check_step_result()` และการ `return` จาก `run()` ใน `playwrighttest.py` ซึ่งรองรับ error ตอน Login/Checkout แล้ว

### 7.5 รันไฟล์ใหม่

```powershell
python test_my_website.py
```

ผลที่ควรเห็นหลังเติมครบ: browser ทำตามขั้นตอนของเว็บใหม่ทีละแถว และ Terminal แสดงผลแยกตามชื่อเคส

## 8. อยากแก้อะไร ต้องแก้ตรงไหน

| สิ่งที่อยากเปลี่ยน | จุดที่ต้องแก้ |
| --- | --- |
| เพิ่มหรือลดจำนวนเคส แต่ขั้นตอนเดิม | เพิ่ม/ลบแถวใน Excel ไม่ต้องเพิ่มฟังก์ชันตามจำนวนเคส |
| เปลี่ยนชื่อ รหัส หรือค่าที่กรอก | เซลล์ของแถวนั้นใน Excel |
| เปลี่ยนไฟล์หรือชีตในแม่แบบ | `EXCEL_FILE` และ `SHEET_NAME` |
| เปลี่ยนไฟล์หรือชีตในตัวอย่างปัจจุบัน | `EXCEL_FILE` และ `SHEET_NAME` ด้านบนไฟล์ |
| เพิ่มข้อมูลคอลัมน์ใหม่ | Excel, การตรวจคอลัมน์ และ `case['ชื่อใหม่']` ที่จุดใช้งาน |
| เปลี่ยน URL ทั้งตัวอย่างและแม่แบบ | `base_url` ใน `config.yaml` |
| เปลี่ยนปุ่ม/ช่องกรอกที่หาไม่เจอ | locator ในฟังก์ชันของขั้นตอนนั้น |
| เพิ่มขั้นตอนใหม่ | สร้าง `def` แล้วเรียกใน `run()` ตามลำดับ |
| เพิ่มการตรวจในขั้นตอนเดิม | เพิ่ม `expect()` ในฟังก์ชันนั้น |
| เปลี่ยนผลที่คาดหวัง | `expected_result` ใน Excel และตรวจว่า `verify_result()` รองรับผลประเภทนั้น |
| เพิ่มเคส error ที่ Login/Checkout ในตัวอย่างปัจจุบัน | เพิ่มแถว Excel พร้อม expected_result; `check_step_result()` รองรับสองจุดนี้แล้ว |
| เพิ่มเส้นทางอื่นหรือ error ที่ขั้นตอนอื่น | เพิ่มฟังก์ชันตรวจและเงื่อนไขใน `run()` ตามผลของเว็บ |
| เปลี่ยน browser หรือซ่อนหน้าต่าง | `launch()` ใน `run()` |
| เลือกรันเฉพาะบางเคส | กรอง `cases` หลังอ่าน Excel ใน `main()` |
| เปลี่ยนสรุปผลหรือพฤติกรรมเมื่อผิดพลาด | `try/except/else` และตัวนับใน `main()` |

### ข้อแตกต่างของการเพิ่มคอลัมน์ในสองไฟล์

ทั้ง `playwrighttest.py` และ `blueprint.py` อ่านทุกคอลัมน์ที่มีหัวตาราง เพิ่มใน `REQUIRED_COLUMNS` เมื่อต้องการบังคับว่าต้องมีคอลัมน์นั้น ตัวอย่างปัจจุบันบังคับเซลล์ไม่ว่างเฉพาะชื่อเคสและ expected_result เพื่อให้ทดสอบ username/password และข้อมูล Checkout ว่างได้ ส่วนแม่แบบกำหนดการห้ามว่างผ่าน `NONEMPTY_COLUMNS`

ทั้งสองแบบต้องมีคำสั่งใช้ข้อมูล เช่น `.fill(case['email'])` ด้วย การเพิ่มคอลัมน์ใน Excel อย่างเดียวไม่ได้ทำให้ browser กรอกข้อมูลเพิ่มเอง

### ตัวอย่างเพิ่มฟังก์ชันตรวจจำนวนสินค้า

สำหรับ `playwrighttest.py` เขียนฟังก์ชันนี้เหนือ `run()`:

```python
def verify_cart_count(page: Page, case: dict[str, str]):
    items = [
        name.strip()
        for name in case["items_to_add"].split(",")
        if name.strip()
    ]
    expect(page.locator('[data-test="shopping-cart-badge"]')).to_have_text(
        str(len(items))
    )
```

จากนั้นแทรกการเรียกในลำดับเดิมของ `run()`:

```python
open_website(page, config)
log_in(page, case)
if not check_step_result(page, case, '[data-test="inventory-list"]'):
    return
add_items(page, case)
verify_cart_count(page, case)  # เพิ่มตรงนี้ หลังเลือกสินค้า
fill_checkout_information(page, case)
if not check_step_result(page, case, '[data-test="checkout-summary-container"]'):
    return
verify_result(page, case)
```

ตัวอย่างนี้สมมติว่าแต่ละชื่อสินค้าในเซลล์ไม่ซ้ำกัน การประกาศ `def` อย่างเดียวไม่ทำงาน ต้องมีคำสั่งเรียกด้วย

### ตัวอย่างเลือกบางเคส

หลัง `cases = read_test_cases()` ใน `main()` เลือกใช้แบบใดแบบหนึ่ง:

```python
# ลองเคสแรกเท่านั้น
cases = cases[:1]
```

```python
# ลองเฉพาะชื่อที่ต้องการ
cases = [
    case for case in cases
    if case["testcase name"] == "TC02_Empty_First_Name"
]
```

เอาตัวกรองออกเมื่อต้องการกลับมารันครบ ถ้ากรองชื่อไม่ตรงจะไม่มีเคสถูกรัน ให้ดูจำนวน TOTAL ประกอบด้วย

### เปลี่ยน browser

ใน `run()` เลือกปรับบรรทัดเปิด browser:

```python
# Chromium แบบไม่แสดงหน้าต่าง
browser = playwright.chromium.launch(headless=True)
```

```python
# Microsoft Edge ที่ติดตั้งไว้ในเครื่อง
browser = playwright.chromium.launch(channel="msedge", headless=False)
```

## 9. แก้ปัญหาที่พบบ่อย

| อาการ | สิ่งที่ตรวจหรือแก้ |
| --- | --- |
| `python` ไม่ทำงานหรือเปิด Microsoft Store | ตรวจการติดตั้ง Python และ PATH แล้วเปิด Terminal ใหม่ |
| หา `uv` ไม่เจอ | เปิด Terminal ใหม่ ตรวจ PATH; ถ้าติดตั้งผ่าน Python แล้วลอง `python -m uv --version` และใช้ `python -m uv` แทน `uv` ชั่วคราว |
| Activate ถูกบล็อก | ใช้ `.\.venv\Scripts\python.exe` เรียกงานโดยตรง |
| `ModuleNotFoundError` | ตรวจ `sys.executable` ว่าอยู่ใน `.venv` นี้ แล้วติดตั้งแพ็กเกจด้วย `uv pip install` ในโฟลเดอร์โปรเจกต์ |
| แจ้งไม่พบ executable ของ browser | รัน `.\.venv\Scripts\python.exe -m playwright install chromium` |
| `FileNotFoundError` ของ Excel | ตรวจชื่อไฟล์และตำแหน่งคู่กับไฟล์ Python รวม `EXCEL_FILE` หรือ `excel_path` |
| ไม่พบ worksheet | ตรวจชื่อแท็บใน Excel ให้ตรง `SHEET_NAME` หรือชื่อใน `workbook[...]` |
| `config.yaml` ไม่พบหรือ URL ไม่ถูกต้อง | วาง config คู่ไฟล์ Python และระบุ `base_url` เป็น URL http/https |
| ขาดคอลัมน์หรือ `KeyError` | ตรวจหัวตารางกับชื่อใน `case[...]` และการอ่านคอลัมน์ของแต่ละไฟล์ |
| ช่องว่างถูกปฏิเสธก่อนเปิด browser | ตรวจรายการห้ามว่าง; เคสที่ตั้งใจเว้นว่างต้องไม่ถูกบังคับกรอก |
| รหัสศูนย์นำหน้าหาย | เก็บค่าเป็น Text ใน Excel และกรอกค่าที่มีศูนย์ใหม่ การแปลง `str()` ไม่คืนศูนย์ที่หายไปแล้ว |
| `NotImplementedError` ในแม่แบบ | เติมฟังก์ชันที่ระบุ หรือเอาการเรียกออกถ้าไม่ใช้ขั้นตอนนั้น |
| Timeout หา element ไม่เจอ | ตรวจ URL หน้าปัจจุบัน ลำดับขั้นตอน และ locator จาก codegen ว่ายังตรงเว็บ |
| เพิ่มฟังก์ชันแล้วไม่ทำงาน | ตรวจว่ามีคำสั่งเรียกใน `run()` แล้วหรือยัง |
| browser เปิดแล้วปิดเร็ว | เป็นพฤติกรรมปิดเมื่อจบแต่ละเคส ให้ดู RUN/PASS/FAIL ใน Terminal |
| ไม่มีสรุป PASS/FAIL | ตรวจว่ารันไฟล์ที่ใช้ `main()` นี้; โค้ดต้นฉบับจาก codegen ไม่ได้สรุปผลให้เอง |

หากต้องการตรวจว่าใช้ Python ตัวไหน:

```powershell
.\.venv\Scripts\python.exe -c "import sys; print(sys.executable)"
```

## 10. คำสั่งสำหรับกลับมาทำงานครั้งถัดไป

เมื่อเตรียมโปรเจกต์ครบแล้ว ไม่ต้องติดตั้ง uv สร้าง venv หรือติดตั้งแพ็กเกจซ้ำทุกครั้ง:

```powershell
Set-Location -LiteralPath 'D:\ไนพา\testing\testplaywrightforfinal'
.\.venv\Scripts\Activate.ps1
python playwrighttest.py
```

ถ้าจะบันทึกขั้นตอนใหม่:

```powershell
python -m playwright codegen --target python -o recorded_flow.py https://www.saucedemo.com/
```

ถ้าทำเว็บใหม่และเติมแม่แบบแล้ว:

```powershell
python test_my_website.py
```

ออกจาก venv ใน Terminal:

```powershell
deactivate
```

## เอกสารอ้างอิง

- [ติดตั้ง uv](https://docs.astral.sh/uv/getting-started/installation/)
- [สร้างและใช้ venv ด้วย uv](https://docs.astral.sh/uv/pip/environments/)
- [ติดตั้งแพ็กเกจด้วย uv](https://docs.astral.sh/uv/pip/packages/)
- [บันทึกขั้นตอนด้วย Playwright codegen](https://playwright.dev/python/docs/codegen)
- [ติดตั้งและเลือก browser](https://playwright.dev/python/docs/browsers)
- [การตรวจผลด้วย Playwright assertions](https://playwright.dev/python/docs/test-assertions)
- [อ่าน Excel ด้วย openpyxl](https://openpyxl.readthedocs.io/en/stable/tutorial.html)
- [อ่าน YAML ด้วย PyYAML](https://pyyaml.org/wiki/PyYAMLDocumentation)
