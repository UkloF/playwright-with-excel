"""แม่แบบ Playwright + Excel สำหรับนำไปเติมขั้นตอนทดสอบเว็บอื่น

วิธีใช้:
1. คัดลอกไฟล์นี้เป็นชื่อใหม่ เช่น test_my_website.py
2. แก้ค่าตั้งต้น [แก้ตรงนี้] แล้วเติมส่วน [TODO] ตามเว็บของคุณ
3. นำคำสั่งที่ได้จาก codegen มาใส่ในฟังก์ชันของแต่ละขั้นตอน
4. เปลี่ยนค่าตายตัว เช่น fill('ชื่อเดิม') เป็น fill(case['ชื่อคอลัมน์'])
5. ปรับลำดับเรียกฟังก์ชันใน run(); main() วนรันแต่ละเคสให้แล้ว

ลำดับการทำงาน: main -> read_test_cases -> วน case -> run -> ขั้นตอน -> ตรวจผล
หนึ่งแถวข้อมูลที่ไม่ว่างใน Excel = หนึ่งเคส ไม่ต้องกำหนดจำนวนเคสเป็นตัวเลขใน Python

ส่วนใช้ร่วมกันเขียนพร้อมแล้ว: read_config(), read_test_cases(), open_website(), run(), main()
ส่วนเฉพาะเว็บที่ต้องเติม: log_in(), perform_actions(), verify_result()
ฟังก์ชันที่ยังไม่เติมจะแจ้ง NotImplementedError เพื่อไม่รายงาน PASS ทั้งที่ยังไม่ได้ทดสอบ
ตัวอย่างที่เขียนหลัง # เป็นคอมเมนต์ ต้องเอา # ออกและเติมรายละเอียดก่อนใช้งาน
ดูตัวอย่างที่ทำงานครบแล้วได้ใน playwrighttest.py
"""

from pathlib import Path

import yaml
from openpyxl import load_workbook
from playwright.sync_api import Page, Playwright, expect, sync_playwright


# [แก้ตรงนี้: Excel] ระบุชื่อไฟล์ที่อยู่คู่กับไฟล์ Python และชื่อแท็บด้านใน
EXCEL_FILE = "test_data.xlsx"
SHEET_NAME = "test_data"
# [แก้ตรงนี้: คอลัมน์] เพิ่มชื่อคอลัมน์ที่เว็บใหม่ต้องใช้ เช่น 'username', 'keyword'
# ทั้งชื่อในชุดนี้และ case['ชื่อคอลัมน์'] ต้องตรงกับหัวตารางใน Excel
REQUIRED_COLUMNS = {"testcase name", "expected_result"}
# [แก้ตรงนี้: ห้ามว่าง] ใส่เฉพาะคอลัมน์ที่ห้ามว่างทุกเคส
# ช่องที่ตั้งใจเว้นว่างเพื่อทดสอบ error ไม่ควรอยู่ในชุดนี้
NONEMPTY_COLUMNS = {"testcase name"}


def read_config():
    """ใช้ร่วมกันได้: อ่าน URL จาก config.yaml ที่อยู่คู่ไฟล์ Python"""
    with Path(__file__).with_name("config.yaml").open(encoding="utf-8") as stream:
        config = yaml.safe_load(stream)
    if not isinstance(config, dict) or not isinstance(config.get("base_url"), str):
        raise ValueError("config.yaml ต้องมี base_url เป็นข้อความ")
    return config


def read_test_cases():
    """ใช้ร่วมกันได้: อ่านข้อมูลทุกแถวเป็น list ของ dictionary"""
    excel_path = Path(__file__).with_name(EXCEL_FILE)
    # อ่านอย่างเดียว ไม่แก้ Excel; data_only อ่านค่าที่บันทึกไว้ ไม่คำนวณสูตรใหม่
    workbook = load_workbook(excel_path, read_only=True, data_only=True)
    try:
        sheet = workbook[SHEET_NAME]
        rows = sheet.iter_rows(values_only=True)
        headers = next(rows, ())
        missing = (REQUIRED_COLUMNS | NONEMPTY_COLUMNS) - set(headers)
        if missing:
            raise ValueError(f"Excel ขาดคอลัมน์: {', '.join(sorted(missing))}")
        named_headers = [name for name in headers if name is not None]
        if any(not isinstance(name, str) or not name.strip() for name in named_headers):
            raise ValueError("หัวคอลัมน์ Excel ต้องเป็นข้อความที่ไม่ว่าง")
        if len(named_headers) != len(set(named_headers)):
            raise ValueError("หัวคอลัมน์ Excel ต้องไม่ซ้ำกัน")

        cases = []
        for row_number, row in enumerate(rows, start=2):
            if all(value is None for value in row):
                continue
            # อ่านทุกคอลัมน์ที่มีหัวตาราง คอลัมน์ใหม่จึงนำไปใช้ได้โดยไม่แก้ลูปนี้
            # None -> '' สำหรับเคสช่องว่าง; ตัวเลข -> str สำหรับ fill()
            # ตั้งเซลล์รหัสที่มีศูนย์นำหน้าเป็น Text ใน Excel เพื่อเก็บศูนย์ไว้
            case = {
                name: "" if value is None else str(value)
                for name, value in zip(headers, row)
                if name is not None
            }
            for name in NONEMPTY_COLUMNS:
                if not case[name].strip():
                    raise ValueError(f"Excel แถว {row_number}: ต้องระบุ {name}")
            cases.append(case)

        if not cases:
            raise ValueError(f"ไม่พบข้อมูลทดสอบในชีต {SHEET_NAME}")
        return cases
    finally:
        workbook.close()


def open_website(page: Page, config: dict):
    """ใช้ร่วมกันได้: เปลี่ยน base_url ใน config.yaml เป็นเว็บไซต์ของคุณ"""
    page.goto(config["base_url"])


def log_in(page: Page, case: dict[str, str]):
    """[TODO 3] เข้าสู่ระบบ ถ้าเว็บไม่ต้อง Login สามารถลบฟังก์ชันนี้ได้"""
    # วางคำสั่งกรอก username/password และกด Login จาก codegen
    # เปลี่ยนข้อความที่กรอกเป็น case['username'] และ case['password']
    # ถ้าเว็บไม่มี Login ไม่ต้องเรียกฟังก์ชันนี้ใน run()
    raise NotImplementedError("เติม log_in() หรือเอาคำสั่งเรียกออกจาก run() ถ้าไม่ใช้")


def perform_actions(page: Page, case: dict[str, str]):
    """[TODO 4] ทำขั้นตอนหลัก เช่นค้นหา กรอกฟอร์ม หรือส่งคำขอ"""
    # วาง locator.fill(), click(), select_option() จาก codegen ตามลำดับ
    # เปลี่ยนชื่อฟังก์ชันให้สื่อถึงงานจริงได้ และเปลี่ยนจุดเรียกใน run() ให้ตรงกัน
    # ถ้าขั้นตอนยาว ให้แยกเป็น def ย่อยเพิ่มเติม แล้วเรียกใน run()
    raise NotImplementedError("เติมขั้นตอนเฉพาะเว็บใน perform_actions()")


def verify_result(page: Page, case: dict[str, str]):
    """[TODO 5] ตรวจผลจริงบนเว็บเทียบกับผลที่คาดหวังจาก Excel"""
    # เพิ่ม expect(locator).to_have_text(case['expected_result'])
    # หรือใช้ to_be_visible(), to_have_count(), to_have_url() ตามสิ่งที่ต้องตรวจ
    # ต้องมีการตรวจจริงเพื่อยืนยันผล การคลิกครบทุกปุ่มไม่ได้แปลว่าผลถูกต้อง
    # เคสสำเร็จและเคสผิดพลาดอาจต้องตรวจคนละตำแหน่ง ให้แยกตามประเภทเคส
    raise NotImplementedError("เติม expect เพื่อตรวจผลใน verify_result()")


# [TODO เพิ่มขั้นตอน] เขียน def ชื่อใหม่(page, case) แถวนี้
# แล้วเพิ่มคำสั่งเรียกชื่อใหม่(page, case) ใน run() ตามตำแหน่งที่ต้องการ
# ถ้าเพียงประกาศ def แต่ไม่เรียก ฟังก์ชันนั้นจะยังไม่ทำงาน


def run(playwright: Playwright, case: dict[str, str], config: dict):
    """ใช้ร่วมกันได้: เปิด/ปิด browser และแยก session ให้แต่ละเคส"""
    # [แก้ตรงนี้: browser] headless=False แสดงหน้าต่าง; Edge เพิ่ม channel='msedge'
    browser = playwright.chromium.launch(headless=False)
    try:
        context = browser.new_context()
        try:
            page = context.new_page()
            # [แก้ตรงนี้: ลำดับขั้นตอน] เพิ่ม/ลบ/สลับฟังก์ชันตามเว็บของคุณ
            # ไม่ต้อง Login ให้ลบ log_in(page, case) ออก
            # ทุกเคสจะทำลำดับนี้ หากขั้นตอนต่างกันให้เพิ่ม if แยกตามประเภทเคส
            open_website(page, config)
            log_in(page, case)
            perform_actions(page, case)
            verify_result(page, case)
        finally:
            context.close()
    finally:
        browser.close()


def main():
    """ใช้ร่วมกันได้: วนรันเคสและสรุปผล ไม่ต้องแก้เมื่อเพิ่มแถวใน Excel"""
    config = read_config()
    cases = read_test_cases()
    # [แก้ตรงนี้: เลือกเคส] ลองเฉพาะเคสแรกด้วย cases = cases[:1] ก่อนวน for
    passed = 0
    failed = 0
    with sync_playwright() as playwright:
        for case in cases:
            case_name = case["testcase name"]
            print(f"\nRUN: {case_name}", flush=True)
            try:
                run(playwright, case, config)
            except Exception as error:
                # เคสผิดพลาดยังทำเคสถัดไป; จะแสดงสาเหตุไว้ใน Terminal
                failed += 1
                print(f"FAIL: {case_name}\n{type(error).__name__}: {error}", flush=True)
            else:
                passed += 1
                print(f"PASS: {case_name}", flush=True)
    print(f"\nTOTAL: {len(cases)} | PASS: {passed} | FAIL: {failed}", flush=True)
    # main() ไม่จำเป็น สามารถย้ายโค้ดเริ่มต้นไปใต้ if ด้านล่างได้เช่นกัน
    # exit code 1 เมื่อมีเคสไม่ผ่าน; 0 เมื่อทุกเคสผ่าน
    return 1 if failed else 0


if __name__ == '__main__':
    # เริ่มเมื่อกด Run/รันไฟล์โดยตรง ไม่เริ่มเมื่อไฟล์อื่น import เข้ามา
    raise SystemExit(main())
