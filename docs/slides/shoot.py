"""Render each slide to PNG at whatever size its own <body> declares.

Sizing is taken from the document rather than fixed here, so a square carousel
slide and a landscape README diagram can live in the same folder and be built
by the same command.
"""
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

        size = page.evaluate(
            "() => ({w: document.body.offsetWidth, h: document.body.offsetHeight})")
        page.set_viewport_size({"width": size["w"], "height": size["h"]})
        page.wait_for_timeout(350)

        dest = os.path.join(OUT, f"{name}.png")
        page.screenshot(path=dest, clip={"x": 0, "y": 0,
                                         "width": size["w"], "height": size["h"]})
        print(f"  {name}.png  {size['w']}x{size['h']}")
    browser.close()
