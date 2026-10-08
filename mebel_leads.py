"""Бесплатный сбор телефонов мебельщиков по городу РФ -> Excel.

Источник: OpenStreetMap (без ключей и оплаты). Если у компании в OSM
нет телефона, но есть сайт, телефон берётся со страниц сайта.

Пример:
    python mebel_leads.py Казань
    python mebel_leads.py "Нижний Новгород" --no-sites
"""
import argparse
import re
import sys
import time
from urllib.parse import urljoin

import requests
from openpyxl import Workbook

OVERPASS = [
    "https://maps.mail.ru/osm/tools/overpass/api/interpreter",
    "https://overpass-api.de/api/interpreter",
]
QUERY = """[out:json][timeout:120];
area["name"="{city}"]["place"~"city|town"]->.a;
(nwr["shop"~"furniture|kitchen"](area.a);
 nwr["craft"~"carpenter|cabinet_maker|joiner|upholsterer"](area.a););
out tags;"""
PHONE_RE = re.compile(r"(?:\+7|8)[\s\-(]*\d{3}[\s\-)]*\d{3}[\s\-]*\d{2}[\s\-]*\d{2}")
UA = {"User-Agent": "Mozilla/5.0"}


def normalize_phone(raw):
    digits = re.sub(r"\D", "", raw or "")
    if len(digits) == 11 and digits[0] in "78":
        return "+7" + digits[1:]
    if len(digits) == 10:
        return "+7" + digits
    return None


def fetch_osm(city):
    for attempt in range(4):
        for url in OVERPASS:
            try:
                r = requests.get(url, params={"data": QUERY.format(city=city)}, timeout=150)
                if r.ok:
                    return r.json()["elements"]
            except (requests.RequestException, ValueError):
                pass
        time.sleep(10 * (attempt + 1))
    sys.exit("OpenStreetMap не ответил, попробуйте позже")


def phones_from_site(site):
    if not site.startswith("http"):
        site = "http://" + site
    found = []
    for path in ("", "/contacts", "/kontakty", "/contact"):
        try:
            html = requests.get(urljoin(site, path), headers=UA, timeout=10).text
        except requests.RequestException:
            continue
        tel_links = re.findall(r'href="tel:([^"]+)"', html)
        found += tel_links or PHONE_RE.findall(html)
        if found:
            break
    return found


def main():
    ap = argparse.ArgumentParser(description="Телефоны мебельщиков города -> Excel")
    ap.add_argument("city", help="Город, например Казань")
    ap.add_argument("--no-sites", action="store_true", help="Не заходить на сайты компаний")
    ap.add_argument("--out", help="Имя файла .xlsx")
    args = ap.parse_args()

    elements = fetch_osm(args.city)
    print(f"Найдено в OpenStreetMap: {len(elements)}")
    rows, seen = [], set()
    for e in elements:
        t = e.get("tags", {})
        site = t.get("website") or t.get("contact:website")
        raw = ";".join(filter(None, [t.get("phone"), t.get("contact:phone"),
                                     t.get("contact:mobile")]))
        phones, source = re.split(r"[;,]", raw) if raw else [], "OSM"
        if not phones and site and not args.no_sites:
            phones, source = phones_from_site(site), "сайт"
        for p in phones:
            phone = normalize_phone(p)
            if phone and phone not in seen:
                seen.add(phone)
                addr = ", ".join(filter(None, [t.get("addr:street"), t.get("addr:housenumber")]))
                rows.append([t.get("name"), phone, addr, site,
                             t.get("shop") or t.get("craft"), source])

    wb = Workbook()
    ws = wb.active
    ws.title = "Мебельщики"
    ws.append(["Название", "Телефон", "Адрес", "Сайт", "Тип", "Откуда телефон"])
    for r in rows:
        ws.append(r)
    for col, width in zip("ABCDEF", (40, 16, 35, 35, 15, 15)):
        ws.column_dimensions[col].width = width
    out = args.out or f"mebel_{args.city.replace(' ', '_')}.xlsx"
    wb.save(out)
    print(f"Готово: {len(rows)} номеров -> {out}")


if __name__ == "__main__":
    main()
