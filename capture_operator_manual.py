#!/usr/bin/env python3
"""Capture read-only screenshots for the operator manual.

Credentials are read from OPERATOR_USER and OPERATOR_PASSWORD.  The script only
opens pages and dialogs; it never starts calls, conferences, deletes records, or
submits configuration changes.
"""

from __future__ import annotations

import os
from pathlib import Path

from playwright.sync_api import Page, sync_playwright


BASE_URL = "https://36.141.81.170:8091/webuser/"
ASSET_DIR = Path(__file__).resolve().parent / "manual_assets"
CHROME_PATH = "/usr/local/bin/google-chrome"


def screenshot(page: Page, name: str, *, locator=None, masks=None) -> Path:
    path = ASSET_DIR / name
    effective_masks = list(masks or [])
    for sensitive_text in ("091000", "张涛"):
        candidate = page.get_by_text(sensitive_text, exact=False)
        if candidate.count():
            effective_masks.append(candidate)
    options = {
        "path": str(path),
        "type": "png",
        "animations": "disabled",
        "caret": "hide",
    }
    if effective_masks:
        options["mask"] = effective_masks
        options["mask_color"] = "#D9E2F3"
    if locator is None:
        page.screenshot(**options)
    else:
        locator.screenshot(**options)
    return path


def wait_ready(page: Page, route: str) -> None:
    page.goto(
        f"{BASE_URL}#/console/{route}",
        wait_until="domcontentloaded",
        timeout=60_000,
    )
    page.wait_for_timeout(2_500)


def close_dialog(page: Page) -> None:
    close_button = page.get_by_label("el.dialog.close")
    if close_button.count():
        close_button.last.click()
        page.wait_for_timeout(300)


def capture_conference_dialog(page: Page, mode: str, prefix: str) -> None:
    if mode == "添加":
        page.get_by_role("button", name="添加", exact=True).click()
    else:
        page.get_by_title("查看详情").click()
    page.wait_for_timeout(500)
    dialog = page.locator(".el-dialog:visible")
    tab_specs = [
        ("基本设置", "基本设置"),
        ("视频设置", "视频设置"),
        ("录音录像", "录音录像"),
        ("会议成员", "会议成员"),
    ]
    for index, (tab_name, file_label) in enumerate(tab_specs, start=1):
        dialog.get_by_text(tab_name, exact=True).click()
        page.wait_for_timeout(350)
        masks = []
        if tab_name == "会议成员":
            # The selector can contain personal contact details.  Keep the
            # workflow visible while masking the actual directory contents.
            candidates = dialog.locator(
                ".member-list, .contact-list, .el-tree, .virtual-list, "
                ".el-scrollbar__view"
            )
            if candidates.count():
                masks = [candidates.first]
        screenshot(
            page,
            f"{prefix}-{index:02d}-{file_label}.png",
            locator=dialog,
            masks=masks,
        )
    close_dialog(page)


def main() -> None:
    username = os.environ.get("OPERATOR_USER")
    password = os.environ.get("OPERATOR_PASSWORD")
    if not username or not password:
        raise SystemExit(
            "Set OPERATOR_USER and OPERATOR_PASSWORD before running this script."
        )

    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=True,
            executable_path=CHROME_PATH,
            args=["--no-sandbox", "--ignore-certificate-errors"],
        )
        page = browser.new_page(
            viewport={"width": 1920, "height": 1080},
            ignore_https_errors=True,
        )

        page.goto(
            f"{BASE_URL}#/console/login?redirect=/console/conference",
            wait_until="networkidle",
            timeout=60_000,
        )
        screenshot(page, "010-系统登录页.png")
        page.locator("input").nth(0).fill(username)
        page.locator("input").nth(1).fill(password)
        page.get_by_role("button", name="登录").click()
        page.wait_for_timeout(5_000)

        wait_ready(page, "call")
        screenshot(page, "020-语音调度工作台.png")
        page.get_by_role("button", name="拨号键盘").click()
        page.wait_for_timeout(300)
        screenshot(page, "021-语音调度拨号键盘.png")
        page.get_by_role("button", name="拨号键盘").click()

        wait_ready(page, "video")
        screenshot(page, "030-视频会议工作台.png")
        page.get_by_role("button", name="更多").click()
        page.wait_for_timeout(300)
        screenshot(page, "031-视频会议更多控制.png")

        wait_ready(page, "conference")
        screenshot(page, "040-会议会商列表.png")
        capture_conference_dialog(page, "查看详情", "041-会议详情")
        capture_conference_dialog(page, "添加", "042-添加会议")

        wait_ready(page, "conference")
        page.get_by_role("tab", name="视频窗口").click()
        page.wait_for_timeout(500)
        screenshot(page, "043-会议会商视频窗口.png")

        wait_ready(page, "cdr")
        table_body = page.locator(".el-table__body-wrapper:visible")
        masks = [table_body] if table_body.count() else None
        screenshot(page, "050-录音录像查询页.png", masks=masks)

        play_button = page.locator(
            ".el-table:visible tbody tr .el-button--success:not(.is-disabled)"
        )
        if play_button.count():
            play_button.first.click()
            page.wait_for_timeout(1_000)
            player_dialog = page.locator(".el-dialog:visible")
            if player_dialog.count():
                screenshot(
                    page,
                    "053-录音录像播放窗口.png",
                    locator=player_dialog,
                )
                close_dialog(page)

        type_select = page.locator("input[placeholder='请选择类型']")
        if type_select.count():
            type_select.click()
            page.wait_for_timeout(300)
            screenshot(page, "051-录音录像类型筛选.png", masks=masks)
            page.keyboard.press("Escape")

        date_input = page.locator("input[placeholder='请选择录音日期']")
        if date_input.count():
            date_input.click()
            page.wait_for_timeout(300)
            screenshot(page, "052-录音录像日期筛选.png", masks=masks)
            page.keyboard.press("Escape")

        browser.close()


if __name__ == "__main__":
    main()
