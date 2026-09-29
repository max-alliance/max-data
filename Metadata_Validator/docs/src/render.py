"""README 그림 다시 그리기: docs/src/*.html → docs/images/*.png

    pip install playwright && python -m playwright install chromium
    python docs/src/render.py

각 HTML 의 id="fig" 요소를 2배 해상도로 캡처한다. (한글 폰트: 맑은 고딕 또는 Noto Sans CJK KR)
"""
import glob
import os
import sys

from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(os.path.dirname(HERE), "images")
PAGES = {
    "max_submission_flow.html": "max_submission_flow.png",
    "app_guide.html": "app_guide.png",
}


def main():
    os.makedirs(OUT, exist_ok=True)
    exe = os.environ.get("CHROMIUM_PATH") or next(
        iter(glob.glob("/opt/pw-browsers/chromium-*/chrome-linux/chrome")), None)
    only = set(sys.argv[1:])
    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=exe) if exe else p.chromium.launch()
        page = browser.new_page(device_scale_factor=2, viewport={"width": 2400, "height": 1600})
        for src, dst in PAGES.items():
            if only and src not in only:
                continue
            path = os.path.join(HERE, src)
            if not os.path.isfile(path):
                continue
            page.goto("file://" + path.replace(os.sep, "/"))
            page.wait_for_load_state("networkidle")
            page.evaluate("document.fonts.ready")
            page.locator("#fig").screenshot(path=os.path.join(OUT, dst))
            print("→", os.path.join(OUT, dst))
        browser.close()


if __name__ == "__main__":
    main()
