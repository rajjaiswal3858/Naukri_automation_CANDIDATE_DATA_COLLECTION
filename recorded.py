import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page1 = context.new_page()
    page1.goto("https://www.naukri.com/")
    page1.get_by_role("button", name="For employers").click()
    page1.get_by_role("button", name="For employers").click()
    page1.get_by_role("link", name="Employer Login").click()
    page1.get_by_role("tab", name="Register/Log in").click()
    page1.get_by_role("textbox", name="Enter registered email ID").click()
    page1.get_by_role("textbox", name="Enter registered email ID").click()
    page1.get_by_role("textbox", name="Enter registered email ID").fill("manvi")
    page1.get_by_role("textbox", name="Enter registered email ID").click()
    page1.get_by_role("textbox", name="Enter registered email ID").fill("manvi@head-field.com")
    page1.get_by_role("textbox", name="Enter registered email ID").press("Tab")
    with page1.expect_download() as download_info:
        page1.get_by_role("link", name="Download").click()
    download = download_info.value
    page2 = context.new_page()
    page2.goto("chrome://downloads/")
    page2.get_by_text("7dc51389-aacc-4a82-8aa1-f643dcfa31497dc51389-aacc-4a82-8aa1-f643dcfa3149 From").click()
    page2.get_by_text("7dc51389-aacc-4a82-8aa1-f643dcfa31497dc51389-aacc-4a82-8aa1-f643dcfa3149 From").click(click_count=4)
    page2.locator("#downloadsList").click()
    page2.get_by_role("button", name="Clear all").click()
    page2.close()
    page1.close()
    page.close()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
