#!/usr/bin/env python3
"""Baut aus src/deck.html + src/fonts/ eine einzelne, offline lauffähige HTML-Datei
und exportiert sie (mit Playwright) als PDF. Aufruf: python3 build.py"""
import base64, pathlib, subprocess, sys

root = pathlib.Path(__file__).parent
fonts = [
    ("Anton", "normal", 400, "anton-latin-400-normal.woff2"),
    ("Instrument Serif", "normal", 400, "instrument-serif-latin-400-normal.woff2"),
    ("Instrument Serif", "italic", 400, "instrument-serif-latin-400-italic.woff2"),
    ("JetBrains Mono", "normal", 400, "jetbrains-mono-latin-400-normal.woff2"),
    ("JetBrains Mono", "normal", 700, "jetbrains-mono-latin-700-normal.woff2"),
]
css = ""
for fam, style, weight, fn in fonts:
    b64 = base64.b64encode((root / "src/fonts" / fn).read_bytes()).decode()
    css += (f"@font-face{{font-family:'{fam}';font-style:{style};font-weight:{weight};"
            f"src:url(data:font/woff2;base64,{b64}) format('woff2')}}\n")

html = (root / "src/deck.html").read_text(encoding="utf-8").replace("/*FONTS*/", css)
out = root / "Lachs-aus-dem-Kaefig.html"
out.write_text(html, encoding="utf-8")
print("HTML:", out, round(out.stat().st_size / 1024), "KB")

if "--no-pdf" not in sys.argv:
    js = f"""
const {{chromium}}=require('playwright');
(async()=>{{
  const b=await chromium.launch();const p=await b.newPage({{viewport:{{width:1920,height:1080}}}});
  await p.goto('file://{out}');await p.emulateMedia({{media:'print'}});await p.evaluate(()=>document.fonts.ready);
  await p.pdf({{path:'{root}/Lachs-aus-dem-Kaefig.pdf',width:'1920px',height:'1080px',printBackground:true}});
  await b.close();
}})();"""
    subprocess.run(["node", "-e", js], check=True)
    print("PDF: ", root / "Lachs-aus-dem-Kaefig.pdf")
