"""Build store-upload zips into dist/: Firefox (manifest as-is) and Chrome
(Firefox-specific keys stripped, description trimmed to the CWS 132-char cap)."""
import json
import os
import zipfile

ROOT = os.path.dirname(os.path.abspath(__file__))
EXT = os.path.join(ROOT, "extension")
DIST = os.path.join(ROOT, "dist")
os.makedirs(DIST, exist_ok=True)

FILES = ["manifest.json", "background.js", "app.html", "app.css", "app.js",
         "emdash.js", "zip.js",
         "icons/icon16.png", "icons/icon48.png", "icons/icon128.png"]

with open(os.path.join(EXT, "manifest.json"), encoding="utf-8") as f:
    manifest = json.load(f)
version = manifest["version"]

ff = os.path.join(DIST, f"emdash-filter-{version}-firefox.zip")
with zipfile.ZipFile(ff, "w", zipfile.ZIP_DEFLATED) as z:
    for name in FILES:
        z.write(os.path.join(EXT, *name.split("/")), name)
print("built", ff)

chrome_manifest = dict(manifest)
chrome_manifest.pop("browser_specific_settings", None)
chrome_manifest["background"] = {"service_worker": "background.js"}
chrome_manifest["description"] = ("Find em dashes in a document, review smart "
                                  "punctuation fixes, and download a cleaned copy. "
                                  "Runs entirely on your computer.")
assert len(chrome_manifest["description"]) <= 132

cr = os.path.join(DIST, f"emdash-filter-{version}-chrome.zip")
with zipfile.ZipFile(cr, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("manifest.json", json.dumps(chrome_manifest, indent=2))
    for name in FILES[1:]:
        z.write(os.path.join(EXT, *name.split("/")), name)
print("built", cr)
