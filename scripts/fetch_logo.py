"""현대캔버스 사이트에서 로고 URL 탐색."""
import re
import urllib.request

req = urllib.request.Request(
    "https://hyundaicanvas.com/",
    headers={"User-Agent": "Mozilla/5.0"},
)
html = urllib.request.urlopen(req, timeout=30).read().decode("utf-8", errors="ignore")

all_imgs = re.findall(r'<img[^>]+src="([^"]+)"', html, re.I)
print("=== ALL IMGS (first 20) ===")
for u in all_imgs[:20]:
    print(u)

logo_candidates = [
    u for u in all_imgs
    if any(k in u.lower() for k in ["logo", "top", "hd", "banner", "title"])
]
print("\n=== LOGO CANDIDATES ===")
for u in logo_candidates:
    print(u)

# og:image
og = re.findall(r'property="og:image"[^>]+content="([^"]+)"', html, re.I)
print("\n=== OG IMAGE ===", og)

# cafe24 typical path
for u in all_imgs:
    if "upload" in u or "skin" in u:
        print("upload/skin:", u)
