"""Scenario 1: ซื้อสำเร็จ / Scenario 2: Login หรือ Checkout แจ้ง validation error

รัน: .venv/Scripts/python.exe playwrighttest.py
ลำดับ: main -> อ่าน config/Excel -> for case -> run -> Login -> Cart -> Checkout
หนึ่งแถว Excel เป็นหนึ่งเคส ใช้ browser/context ใหม่เพื่อแยก session ทุกครั้ง
"""

from pathlib import Path
from urllib.parse import urlparse

import yaml
from openpyxl import load_workbook
from playwright.sync_api import Page, Playwright, expect, sync_playwright


PROJECT_DIR = Path(__file__).resolve().parent
# [แก้ตรงนี้: Excel] เปลี่ยนชื่อไฟล์และชีตให้ตรงข้อมูลของคุณ
EXCEL_FILE = "test_data.xlsx"
SHEET_NAME = "test_data"
CASE_NAME = "testcase name"
REQUIRED_COLUMNS = {
    CASE_NAME, "username", "password", "items_to_add",
    "first_name", "last_name", "postal_code", "expected_result",
}


def read_config():
    """อ่าน URL จาก config.yaml ครั้งเดียวก่อนเริ่มทดสอบ"""
    # [แก้ตรงนี้: URL] เปลี่ยน base_url ใน config.yaml ไม่ต้องแก้ URL ในฟังก์ชัน Login
    with (PROJECT_DIR / "config.yaml").open(encoding="utf-8") as stream:
        config = yaml.safe_load(stream)
    if not isinstance(config, dict):
        raise ValueError("config.yaml ต้องเป็น mapping ที่มี base_url")
    base_url = config.get("base_url")
    if not isinstance(base_url, str):
        raise ValueError("ต้องระบุ base_url เป็นข้อความใน config.yaml")
    parsed = urlparse(base_url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("base_url ต้องเป็น URL http/https ที่ถูกต้อง")
    return config


def read_test_cases():
    """อ่านทุกคอลัมน์ที่มีหัวตาราง แปลงแต่ละแถวเป็น dictionary"""
    workbook = load_workbook(PROJECT_DIR / EXCEL_FILE, read_only=True, data_only=True)
    try:
        rows = workbook[SHEET_NAME].iter_rows(values_only=True)
        headers = next(rows, ())
        # รองรับไฟล์เก่าที่หัวช่องชื่อเคสยังเป็น Test Cases
        headers = tuple(CASE_NAME if name == "Test Cases" else name for name in headers)
        names = [name for name in headers if name is not None]
        if any(not isinstance(name, str) or not name.strip() for name in names):
            raise ValueError("หัวคอลัมน์ต้องเป็นข้อความที่ไม่ว่าง")
        if len(names) != len(set(names)):
            raise ValueError("ชื่อคอลัมน์ซ้ำกัน")
        missing = REQUIRED_COLUMNS - set(headers)
        if missing:
            raise ValueError(f"Excel ขาดคอลัมน์: {', '.join(sorted(missing))}")

        cases = []
        for row_number, row in enumerate(rows, start=2):
            if all(value is None for value in row):
                continue
            # None -> '' เพื่อทดสอบช่องว่าง; เลข -> str เพื่อส่งเข้า fill()
            # รหัสที่มีศูนย์นำหน้าต้องเก็บเป็น Text ใน Excel ก่อน เช่น 00123
            case = {
                name: "" if value is None else str(value)
                for name, value in zip(headers, row)
                if name is not None
            }
            # [แก้ตรงนี้: ห้ามว่าง] ไม่บังคับข้อมูล Login/Checkout เพราะมี negative case
            for name in (CASE_NAME, "expected_result"):
                if not case[name].strip():
                    raise ValueError(f"Excel แถว {row_number}: ต้องระบุ {name}")
            cases.append(case)
        if not cases:
            raise ValueError("ไม่พบข้อมูลทดสอบใน Excel")
        if len({case[CASE_NAME] for case in cases}) != len(cases):
            raise ValueError("ชื่อ test case ต้องไม่ซ้ำกัน")
        return cases
    finally:
        workbook.close()


def open_website(page: Page, config: dict):
    page.goto(config["base_url"])


def log_in(page: Page, case: dict[str, str]):
    """กรอก Login เท่านั้น การตัดสินผลทำใน check_step_result()"""
    # [แก้ตรงนี้: locator] หากเปลี่ยนเว็บไซต์ ให้แก้ช่องกรอกและปุ่มตาม codegen
    page.locator('[data-test="username"]').fill(case["username"])
    page.locator('[data-test="password"]').fill(case["password"])
    page.locator('[data-test="login-button"]').click()


def check_step_result(page: Page, case: dict[str, str], next_selector: str):
    """รอหน้าถัดไปหรือ error: ถ้า error ตรง expected_result ให้จบเคสนั้น"""
    error = page.locator('[data-test="error"]')
    next_step = page.locator(next_selector)
    # รอผลจริงก่อน ไม่ใช้ is_visible() ทันทีหลังคลิกจนพลาดข้อความที่มาช้า
    expect(next_step.or_(error)).to_be_visible()
    if error.is_visible():
        expect(error).to_have_text(case["expected_result"])
        return False  # validation ตรงตามคาด: ไม่ทำขั้นตอนซื้อสินค้าต่อ
    return True


def add_items(page: Page, case: dict[str, str]):
    """แยกรายการด้วย comma และค้นหาการ์ดจากชื่อสินค้าเต็ม"""
    items = [name.strip() for name in case["items_to_add"].split(",") if name.strip()]
    if not items:
        raise ValueError("เคสที่ Login สำเร็จต้องระบุ items_to_add อย่างน้อยหนึ่งรายการ")
    if len(items) != len(set(items)):
        raise ValueError("items_to_add ต้องไม่ระบุชื่อสินค้าซ้ำ")
    for item_name in items:
        product = page.locator('[data-test="inventory-item"]').filter(
            has=page.get_by_text(item_name, exact=True)
        )
        product.get_by_role("button", name="Add to cart", exact=True).click()


def fill_checkout_information(page: Page, case: dict[str, str]):
    """ไป Cart กด Checkout กรอกข้อมูล และกด Continue"""
    page.locator('[data-test="shopping-cart-link"]').click()
    page.locator('[data-test="checkout"]').click()
    # [แก้ตรงนี้: ข้อมูลเพิ่ม] เพิ่มคอลัมน์ Excel แล้วใช้ case['ชื่อใหม่'] ที่นี่
    page.locator('[data-test="firstName"]').fill(case["first_name"])
    page.locator('[data-test="lastName"]').fill(case["last_name"])
    page.locator('[data-test="postalCode"]').fill(case["postal_code"])
    page.locator('[data-test="continue"]').click()


def verify_result(page: Page, case: dict[str, str]):
    """ยืนยันสั่งซื้อจาก Overview และตรวจข้อความสำเร็จตาม Excel"""
    page.locator('[data-test="finish"]').click()
    expect(page.locator('[data-test="complete-header"]')).to_have_text(
        case["expected_result"]
    )


# [แก้ตรงนี้: ฟังก์ชันใหม่] เขียน def ใหม่เหนือ run() แล้วเรียกตามลำดับภายใน run()
def run(playwright: Playwright, case: dict[str, str], config: dict):
    """หนึ่งครั้ง = หนึ่งเคส; error ที่คาดไว้จบเคสได้โดยไม่ไปขั้นตอนถัดไป"""
    # [แก้ตรงนี้: browser] headless=True ซ่อนหน้าต่าง; Edge เพิ่ม channel='msedge'
    browser = playwright.chromium.launch(headless=False)
    try:
        context = browser.new_context()
        try:
            page = context.new_page()
            open_website(page, config)
            log_in(page, case)
            if not check_step_result(page, case, '[data-test="inventory-list"]'):
                return  # เช่น locked_out_user หรือรหัสผ่านผิด: ตรวจ error แล้วจบเคส
            add_items(page, case)
            fill_checkout_information(page, case)
            if not check_step_result(page, case, '[data-test="checkout-summary-container"]'):
                return  # เช่นเว้นชื่อ: ตรวจ validation แล้วไม่กด Finish
            verify_result(page, case)
        finally:
            context.close()  # ทำเสมอ แม้ return จาก negative case หรือเกิด error
    finally:
        browser.close()


def main():
    config = read_config()
    cases = read_test_cases()
    # [แก้ตรงนี้: จำนวนเคส] เพิ่มแถว Excel; ลองเคสแรกชั่วคราวใช้ cases = cases[:1]
    passed = 0
    failed = 0
    with sync_playwright() as playwright:
        for case in cases:
            print(f"\nRUN: {case[CASE_NAME]}", flush=True)
            try:
                run(playwright, case, config)
            except Exception as error:
                failed += 1
                print(f"FAIL: {case[CASE_NAME]}\n{type(error).__name__}: {error}", flush=True)
            else:
                passed += 1
                print(f"PASS: {case[CASE_NAME]}", flush=True)
    print(f"\nTOTAL: {len(cases)} | PASS: {passed} | FAIL: {failed}", flush=True)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
