#!/usr/bin/env python3
"""
car_tracker.py
--------------
Monitoruje oferty samochodów z OLX i Otomoto zapisane w pliku offers.csv.
Dla każdej oferty:
  - pobiera aktualną cenę, przebieg i lokalizację
  - porównuje z poprzednim odczytem (historia zmian cen)
  - jeśli ogłoszenie zniknęło / jest nieaktualne -> oznacza jako "SPRZEDANE"
  - zapisuje wszystko do tracker.xlsx (kolorowanie, arkusz z historią)

Uruchamiaj cyklicznie (np. co godzinę) przez Harmonogram zadań (Windows)
lub cron (Mac/Linux) - patrz README.md.

WAŻNE: OLX i Otomoto regularnie zmieniają strukturę HTML swoich stron
oraz mogą blokować automatyczne zapytania (captcha, blokada IP przy zbyt
częstym odpytywaniu). Skrypt korzysta w pierwszej kolejności z danych
strukturalnych (JSON-LD), które są najbardziej odporne na zmiany
wizualne strony, a dopiero w drugiej kolejności z selektorów CSS jako
fallback. Jeśli po jakimś czasie przestanie działać - to najpewniej
strona zmieniła strukturę i selektory (SELECTORS_* poniżej) trzeba
będzie zaktualizować.
"""

import csv
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Brakuje openpyxl. Zainstaluj: pip install openpyxl")
    sys.exit(1)

BASE_DIR = Path(__file__).resolve().parent
OFFERS_CSV = BASE_DIR / "offers.csv"
OUTPUT_XLSX = BASE_DIR / "tracker.xlsx"
LOG_FILE = BASE_DIR / "tracker.log"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pl-PL,pl;q=0.9,en;q=0.8",
}

SOLD_PHRASES = [
    "ogłoszenie zakończone", "ogłoszenie nieaktualne", "ogłoszenie wygasło",
    "to ogłoszenie jest nieaktualne", "oferta niedostępna", "not found",
    "strona nie została znaleziona",
]


def log(msg: str):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def load_offers():
    if not OFFERS_CSV.exists():
        log(f"Nie znaleziono {OFFERS_CSV}. Utwórz plik offers.csv wg wzoru z README.")
        sys.exit(1)
    with open(OFFERS_CSV, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def extract_json_ld(soup):
    """Szuka danych structured data (schema.org) - najbardziej stabilne źródło ceny."""
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(tag.string or "{}")
        except (json.JSONDecodeError, TypeError):
            continue
        candidates = data if isinstance(data, list) else [data]
        for item in candidates:
            offers = item.get("offers") if isinstance(item, dict) else None
            if offers:
                offers = offers if isinstance(offers, list) else [offers]
                for off in offers:
                    price = off.get("price")
                    if price:
                        mileage = (
                            item.get("mileageFromOdometer", {}).get("value")
                            if isinstance(item.get("mileageFromOdometer"), dict)
                            else item.get("mileageFromOdometer")
                        )
                        image = item.get("image")
                        if isinstance(image, list):
                            image = image[0] if image else None
                        elif isinstance(image, dict):
                            image = image.get("url")
                        return {
                            "price": str(price),
                            "currency": off.get("priceCurrency", "PLN"),
                            "title": item.get("name"),
                            "mileage": str(mileage) if mileage else None,
                            "image": image,
                        }
    return None


def extract_image_fallback(soup):
    """Plan B na zdjęcie: tag og:image (meta tag do udostępniania w social media,
    prawie zawsze obecny i wskazuje na pierwsze/główne zdjęcie oferty)."""
    tag = soup.find("meta", property="og:image")
    if tag and tag.get("content"):
        return tag["content"]
    img = soup.find("img")
    if img and img.get("src"):
        return img["src"]
    return None


def parse_price_fallback(soup, site):
    """Selektory CSS jako plan B, gdyby JSON-LD nie było dostępne.
    UWAGA: te selektory trzeba czasem poprawić ręcznie, bo strony się zmieniają."""
    text = soup.get_text(" ", strip=True)
    match = re.search(r"(\d[\d\s]{2,10})\s*zł", text)
    price = match.group(1).replace(" ", "") if match else None

    mileage_match = re.search(r"(\d[\d\s]{2,7})\s*km\b", text)
    mileage = mileage_match.group(1).replace(" ", "") if mileage_match else None

    location = None
    if site == "olx":
        loc_el = soup.select_one('[data-testid="location-date"]')
        if loc_el:
            location = loc_el.get_text(strip=True).split(" - ")[0]
    elif site == "otomoto":
        loc_el = soup.select_one('[data-testid="location"]') or soup.select_one(".offer-meta__location")
        if loc_el:
            location = loc_el.get_text(strip=True)

    return price, location, mileage


def detect_site(url: str) -> str:
    if "olx.pl" in url:
        return "olx"
    if "otomoto.pl" in url:
        return "otomoto"
    return "inne"


def check_offer(url: str):
    """Zwraca dict: status, price, currency, location, title."""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
    except requests.RequestException as e:
        return {"status": "BŁĄD POŁĄCZENIA", "price": None, "currency": None,
                "location": None, "title": None, "note": str(e)}

    if resp.status_code == 404:
        return {"status": "SPRZEDANE / USUNIĘTE", "price": None, "currency": None,
                "location": None, "title": None, "mileage": None, "image": None, "note": "HTTP 404"}

    lower_text = resp.text.lower()
    if any(phrase in lower_text for phrase in SOLD_PHRASES):
        return {"status": "SPRZEDANE / NIEAKTUALNE", "price": None, "currency": None,
                "location": None, "title": None, "mileage": None, "image": None,
                "note": "wykryto frazę o nieaktualności"}

    soup = BeautifulSoup(resp.text, "html.parser")
    site = detect_site(url)

    structured = extract_json_ld(soup)
    _, location, fallback_mileage = parse_price_fallback(soup, site)
    image = extract_image_fallback(soup)
    if structured:
        price = structured["price"]
        currency = structured["currency"]
        title = structured.get("title")
        mileage = structured.get("mileage") or fallback_mileage
        image = structured.get("image") or image
    else:
        price, location, mileage = parse_price_fallback(soup, site)
        currency = "PLN"
        title_el = soup.find("h1")
        title = title_el.get_text(strip=True) if title_el else None

    if not price:
        return {"status": "NIE UDAŁO SIĘ ODCZYTAĆ", "price": None, "currency": None,
                "location": location, "title": title, "mileage": mileage, "image": image,
                "note": "strona odpowiedziała, ale nie znaleziono ceny - sprawdź selektory"}

    return {"status": "AKTYWNE", "price": price, "currency": currency,
            "location": location, "title": title, "mileage": mileage, "image": image, "note": ""}


def load_previous_data():
    """Wczytuje poprzedni stan z tracker.xlsx (jeśli istnieje), żeby wykryć zmiany cen."""
    if not OUTPUT_XLSX.exists():
        return {}
    wb = load_workbook(OUTPUT_XLSX)
    if "Oferty" not in wb.sheetnames:
        return {}
    ws = wb["Oferty"]
    prev = {}
    headers = [c.value for c in ws[1]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        row_dict = dict(zip(headers, row))
        if row_dict.get("URL"):
            prev[row_dict["URL"]] = row_dict
    return prev


def write_xlsx(results, history_entries):
    if OUTPUT_XLSX.exists():
        wb = load_workbook(OUTPUT_XLSX)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    # --- Arkusz "Oferty" (aktualny stan) ---
    if "Oferty" in wb.sheetnames:
        wb.remove(wb["Oferty"])
    ws = wb.create_sheet("Oferty", 0)

    columns = ["Nazwa", "URL", "Zdjęcie", "Status", "Cena", "Waluta", "Przebieg (km)", "Lokalizacja",
               "Ostatnia zmiana ceny", "Ostatnie sprawdzenie", "Uwagi"]
    ws.append(columns)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

    for r in results:
        ws.append([
            r["title"] or "", r["url"], r.get("image") or "", r["status"], r["price"] or "",
            r["currency"] or "", r.get("mileage") or "", r["location"] or "",
            r["price_change"], r["checked_at"], r["note"],
        ])
        row_idx = ws.max_row
        status = r["status"]
        fill = green if status == "AKTYWNE" else (
            red if "SPRZEDANE" in status else yellow
        )
        for col in range(1, len(columns) + 1):
            ws.cell(row=row_idx, column=col).fill = fill

    for i, col in enumerate(columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(col) + 4)

    # --- Arkusz "Historia cen" ---
    if "Historia cen" in wb.sheetnames:
        hist_ws = wb["Historia cen"]
    else:
        hist_ws = wb.create_sheet("Historia cen")
        hist_ws.append(["Data", "URL", "Poprzednia cena", "Nowa cena", "Zmiana"])
        for cell in hist_ws[1]:
            cell.font = Font(bold=True)

    for entry in history_entries:
        hist_ws.append(entry)

    wb.save(OUTPUT_XLSX)


def status_dot_html(status: str) -> str:
    if status == "AKTYWNE":
        return '<span class="dot dot-active"></span>Aktywna'
    if "SPRZEDANE" in status:
        return '<span class="dot dot-sold"></span>Sprzedane / nieaktualne'
    return f'<span class="dot dot-unknown"></span>{status}'


def generate_html(results):
    """Tworzy dashboard.html - statyczną stronę z wynikami, publikowaną przez
    GitHub Pages. Dane pobiera skrypt (uruchamiany przez GitHub Actions),
    więc strona działa bez Twojego komputera."""

    active_count = sum(1 for r in results if r["status"] == "AKTYWNE")
    sold_count = sum(1 for r in results if "SPRZEDANE" in r["status"])
    prices = [float(r["price"]) for r in results if r["price"]]
    avg_price = f'{sum(prices)/len(prices):,.0f} zł'.replace(",", " ") if prices else "—"

    rows_html = []
    for r in results:
        price_txt = f'{float(r["price"]):,.0f} zł'.replace(",", " ") if r["price"] else "—"
        mileage_txt = f'{int(float(r["mileage"])):,} km'.replace(",", " ") if r.get("mileage") else "—"
        change_txt = f'<span class="delta">{r["price_change"]}</span>' if r["price_change"] else ""
        row_class = "row-sold" if "SPRZEDANE" in r["status"] else ""
        status_key = "active" if r["status"] == "AKTYWNE" else ("sold" if "SPRZEDANE" in r["status"] else "unknown")
        thumb_html = (
            f'<img class="thumb" src="{r["image"]}" alt="" loading="lazy" '
            f'onerror="this.style.display=\'none\'">'
            if r.get("image") else '<div class="thumb thumb-empty">brak zdjęcia</div>'
        )
        rows_html.append(f"""
        <tr class="{row_class}" data-status="{status_key}">
          <td>
            <div class="offer-cell">
              {thumb_html}
              <div>
                <div class="offer-name">{r['title'] or 'Bez nazwy'}</div>
                <a href="{r['url']}" target="_blank" rel="noopener">zobacz ogłoszenie ↗</a>
              </div>
            </div>
          </td>
          <td class="mono">{price_txt}{change_txt}</td>
          <td class="mono dim">{mileage_txt}</td>
          <td class="dim">{r['location'] or '—'}</td>
          <td>{status_dot_html(r['status'])}</td>
          <td class="dim">{r['checked_at']}</td>
        </tr>""")

    html = f"""<!DOCTYPE html>
<html lang="pl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="refresh" content="1200">
<title>Monitor ofert aut</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500;700&display=swap');
  :root{{
    --bg:#1b1e24; --panel:#23262d; --line:#33373f;
    --text:#eef0f3; --dim:#8b93a1;
    --amber:#f0a93b;
    --green:#5cc296; --green-bg:rgba(92,194,150,.12);
    --red:#e2694e; --red-bg:rgba(226,105,78,.12);
    --yellow:#d9b45a; --yellow-bg:rgba(217,180,90,.12);
  }}
  *{{box-sizing:border-box;}}
  body{{
    margin:0; background:var(--bg); color:var(--text);
    font-family:'Inter',system-ui,sans-serif;
    padding:36px 24px 70px;
  }}
  .wrap{{max-width:980px;margin:0 auto;}}
  header{{display:flex;justify-content:space-between;align-items:flex-end;
    flex-wrap:wrap;gap:12px;margin-bottom:28px;border-bottom:1px solid var(--line);padding-bottom:20px;}}
  h1{{font-family:'Space Grotesk';font-weight:700;font-size:30px;margin:0;letter-spacing:0.2px;}}
  .updated{{color:var(--dim);font-size:12.5px;}}
  .updated b{{color:var(--amber);font-weight:600;}}

  .gauges{{display:flex;gap:0;margin-bottom:28px;border:1px solid var(--line);border-radius:8px;overflow:hidden;}}
  .gauge{{flex:1;padding:16px 20px;border-right:1px solid var(--line);}}
  .gauge:last-child{{border-right:none;}}
  .gauge-value{{font-family:'JetBrains Mono';font-size:26px;font-weight:700;color:var(--amber);}}
  .gauge-label{{font-size:11.5px;color:var(--dim);margin-top:2px;}}

  .filters{{display:flex;gap:8px;margin-bottom:14px;}}
  .filter-btn{{
    background:transparent;border:1px solid var(--line);color:var(--dim);
    font-family:'Inter';font-size:12.5px;font-weight:600;padding:6px 13px;
    border-radius:20px;cursor:pointer;
  }}
  .filter-btn.active{{background:var(--amber);border-color:var(--amber);color:#1b1e24;}}

  table{{width:100%;border-collapse:collapse;}}
  thead th{{text-align:left;font-size:11.5px;letter-spacing:.3px;color:var(--dim);
    font-weight:600;padding:8px 10px;border-bottom:1px solid var(--line);}}
  tbody tr{{border-bottom:1px solid var(--line);}}
  tbody tr:hover{{background:var(--panel);}}
  tbody tr.row-sold{{opacity:.55;}}
  td{{padding:14px 10px;font-size:14px;vertical-align:top;}}
  .offer-cell{{display:flex;gap:12px;align-items:center;}}
  .thumb{{width:64px;height:48px;border-radius:5px;object-fit:cover;flex-shrink:0;background:var(--panel);border:1px solid var(--line);}}
  .thumb-empty{{display:flex;align-items:center;justify-content:center;font-size:9px;color:var(--dim);text-align:center;line-height:1.2;padding:2px;}}
  .offer-name{{font-weight:600;margin-bottom:2px;}}
  a{{color:var(--dim);font-size:12px;text-decoration:none;}}
  a:hover{{color:var(--amber);}}
  .mono{{font-family:'JetBrains Mono';font-size:14.5px;}}
  .dim{{color:var(--dim);font-size:13px;}}
  .delta{{font-size:11px;color:var(--amber);margin-left:8px;}}

  .dot{{display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:7px;}}
  .dot-active{{background:var(--green);box-shadow:0 0 0 3px var(--green-bg);}}
  .dot-sold{{background:var(--red);box-shadow:0 0 0 3px var(--red-bg);}}
  .dot-unknown{{background:var(--yellow);box-shadow:0 0 0 3px var(--yellow-bg);}}

  .empty{{text-align:center;padding:60px 20px;color:var(--dim);}}
  footer{{margin-top:30px;color:var(--dim);font-size:11.5px;text-align:center;}}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>Monitor ofert aut</h1>
    <div class="updated">Ostatnie sprawdzenie: <b>{datetime.now().strftime('%Y-%m-%d %H:%M')}</b> &middot; odświeża się co godzinę</div>
  </header>

  <div class="gauges">
    <div class="gauge"><div class="gauge-value">{len(results)}</div><div class="gauge-label">śledzonych ofert</div></div>
    <div class="gauge"><div class="gauge-value">{active_count}</div><div class="gauge-label">aktywnych</div></div>
    <div class="gauge"><div class="gauge-value">{sold_count}</div><div class="gauge-label">sprzedanych / zniknęło</div></div>
    <div class="gauge"><div class="gauge-value">{avg_price}</div><div class="gauge-label">średnia cena aktywnych</div></div>
  </div>

  <div class="filters">
    <button class="filter-btn active" onclick="filterRows('all', this)">Wszystkie</button>
    <button class="filter-btn" onclick="filterRows('active', this)">Aktywne</button>
    <button class="filter-btn" onclick="filterRows('sold', this)">Sprzedane</button>
  </div>

  {"<table><thead><tr><th>Oferta</th><th>Cena</th><th>Przebieg</th><th>Lokalizacja</th><th>Status</th><th>Sprawdzono</th></tr></thead><tbody>" + ''.join(rows_html) + "</tbody></table>" if results else '<div class="empty">Brak ofert w offers.csv — dodaj linki, żeby zaczęło się wypełniać.</div>'}

  <footer>Dane pobierane automatycznie przez GitHub Actions co godzinę, niezależnie od tego czy masz włączony komputer.</footer>
</div>
<script>
function filterRows(status, btn){{
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  document.querySelectorAll('tbody tr').forEach(row => {{
    row.style.display = (status === 'all' || row.dataset.status === status) ? '' : 'none';
  }});
}}
</script>
</body>
</html>"""

    out_path = BASE_DIR / "dashboard.html"
    out_path.write_text(html, encoding="utf-8")
    return out_path


def main():
    offers = load_offers()
    prev_data = load_previous_data()
    results = []
    history_entries = []
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    for offer in offers:
        url = offer["url"].strip()
        if not url:
            continue
        log(f"Sprawdzam: {url}")
        info = check_offer(url)

        prev = prev_data.get(url)
        price_change = ""
        if prev and info["price"]:
            try:
                old_price = float(str(prev.get("Cena", "")).replace(" ", "").replace(",", "."))
                new_price = float(str(info["price"]).replace(" ", "").replace(",", "."))
                if old_price and new_price and old_price != new_price:
                    diff = new_price - old_price
                    price_change = f"{'+' if diff > 0 else ''}{diff:.0f} zł"
                    history_entries.append([now, url, old_price, new_price, price_change])
                    log(f"  Zmiana ceny: {old_price} -> {new_price} ({price_change})")
            except (ValueError, TypeError):
                pass

        title = info["title"] or (offer.get("nazwa") or "")
        results.append({
            "title": title,
            "url": url,
            "status": info["status"],
            "price": info["price"],
            "currency": info["currency"],
            "location": info["location"],
            "mileage": info.get("mileage"),
            "image": info.get("image"),
            "price_change": price_change,
            "checked_at": now,
            "note": info["note"],
        })

    write_xlsx(results, history_entries)
    dashboard_path = generate_html(results)
    log(f"Gotowe. Zapisano {len(results)} ofert do {OUTPUT_XLSX.name} i {dashboard_path.name}")


if __name__ == "__main__":
    main()
