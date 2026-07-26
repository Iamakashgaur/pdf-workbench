"""Render the LinkedIn carousel slides to PNG at 1200x1200."""
import glob
import os
import sys

from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = sys.argv[1]
os.makedirs(OUT, exist_ok=True)

with sync_playwright() as pw:
    browser = pw.chromium.launch(channel="chrome", headless=True)
    page = browser.new_page(viewport={"width": 1200, "height": 1200},
                            device_scale_factor=2, color_scheme="light")
    for src in sorted(glob.glob(os.path.join(HERE, "*.html"))):
        name = os.path.splitext(os.path.basename(src))[0]
        page.goto("file:///" + src.replace("\\", "/"), wait_until="networkidle")
        page.wait_for_timeout(350)
        dest = os.path.join(OUT, f"{name}.png")
        page.screenshot(path=dest, clip={"x": 0, "y": 0, "width": 1200, "height": 1200})
        print(f"  {name}.png")
    browser.close()
