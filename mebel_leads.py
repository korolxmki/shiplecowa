"""Сбор телефонов мебельщиков по городу РФ из 2ГИС или Google Maps в Excel.

Примеры:
    python mebel_leads.py Казань --source 2gis --key КЛЮЧ_2ГИС
    python mebel_leads.py "Нижний Новгород" --source google --key КЛЮЧ_GOOGLE
"""
import argparse
import re
import sys
import time

import requests
from openpyxl import Workbook

QUERIES = [
    "мебель на заказ",
    "мебельная фабрика",
    "мебельный цех",
    "корпусная мебель",
    "кухни на заказ",
    "шкафы-купе на заказ",
    "мягкая мебель производство",
    "офисная мебель",
]


def normalize_phone(raw):
    digits = re.sub(r"\D", "", raw or "")
    if len(digits) == 11 and digits[0] in "78":
        return "+7" + digits[1:]
    if len(digits) == 10:
        return "+7" + digits
    return None


def search_2gis(city, query, key):
    url = "https://catalog.api.2gis.com/3.0/items"
    page, page_size = 1, 10
    while True:
        params = {
            "q": f"{query} {city}",
            "type": "branch",
            "fields": "items.contact_groups,items.address,items.full_address_name",
            "page": page,
            "page_size": page_size,
            "key": key,
        }
        data = requests.get(url, params=params, timeout=30).json()
        code = data.get("meta", {}).get("code")
        if code == 404:  # результаты закончились
            return
        if code != 200:
            sys.exit(f"2ГИС ошибка: {data.get('meta', {}).get('error')}")
        items = data["result"]["items"]
        for it in items:
            phones, site = [], None
            for group in it.get("contact_groups", []):
                for c in group.get("contacts", []):
                    if c.get("type") == "phone":
                        phones.append(c.get("value") or c.get("text"))
                    elif c.get("type") == "website" and not site:
                        site = c.get("text") or c.get("value")
            yield {
                "name": it.get("name"),
                "phones": phones,
                "address": it.get("full_address_name") or it.get("address_name"),
                "site": site,
            }
        if page * page_size >= data["result"].get("total", 0) or not items:
            return
        page += 1
        time.sleep(0.3)


def search_google(city, query, key):
    url = "https://places.googleapis.com/v1/places:searchText"
    headers = {
        "X-Goog-Api-Key": key,
        "X-Goog-FieldMask": "places.displayName,places.nationalPhoneNumber,"
                            "places.internationalPhoneNumber,places.formattedAddress,"
                            "places.websiteUri,nextPageToken",
    }
    body = {"textQuery": f"{query} {city}", "languageCode": "ru",
            "regionCode": "RU", "pageSize": 20}
    while True:
        data = requests.post(url, json=body, headers=headers, timeout=30).json()
        if "error" in data:
            sys.exit(f"Google ошибка: {data['error'].get('message')}")
        for p in data.get("places", []):
            yield {
                "name": p.get("displayName", {}).get("text"),
                "phones": [p.get("internationalPhoneNumber") or p.get("nationalPhoneNumber")],
                "address": p.get("formattedAddress"),
                "site": p.get("websiteUri"),
            }
        token = data.get("nextPageToken")
        if not token:
            return
        body["pageToken"] = token
        time.sleep(2)


def main():
    ap = argparse.ArgumentParser(description="Телефоны мебельщиков города -> Excel")
    ap.add_argument("city", help="Город, например Казань")
    ap.add_argument("--source", choices=["2gis", "google"], default="2gis")
    ap.add_argument("--key", required=True, help="API-ключ 2ГИС или Google")
    ap.add_argument("--out", help="Имя файла .xlsx")
    args = ap.parse_args()

    search = search_2gis if args.source == "2gis" else search_google
    rows, seen = [], set()
    for q in QUERIES:
        found = 0
        for org in search(args.city, q, args.key):
            for raw in org["phones"]:
                phone = normalize_phone(raw)
                if phone and phone not in seen:
                    seen.add(phone)
                    rows.append([org["name"], phone, org["address"], org["site"], q])
                    found += 1
        print(f"{q}: +{found} новых номеров")

    wb = Workbook()
    ws = wb.active
    ws.title = "Мебельщики"
    ws.append(["Название", "Телефон", "Адрес", "Сайт", "Запрос"])
    for r in rows:
        ws.append(r)
    for col, width in zip("ABCDE", (40, 16, 50, 35, 28)):
        ws.column_dimensions[col].width = width
    out = args.out or f"mebel_{args.city.replace(' ', '_')}_{args.source}.xlsx"
    wb.save(out)
    print(f"Готово: {len(rows)} номеров -> {out}")


if __name__ == "__main__":
    main()
