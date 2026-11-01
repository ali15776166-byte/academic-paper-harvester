import os
import re
import time
import requests
import pandas as pd
from datetime import datetime, timedelta

SEARCH_QUERY = "modernity development sociology of organizations bureaucracy"
MAX_ARTICLES = 40
OUTPUT_DIR = "downloaded_pdfs"
EXCEL_REPORT = "download_report.xlsx"

os.makedirs(OUTPUT_DIR, exist_ok=True)

HEADERS = {
    "User-Agent": "AcademicHarvester/2.0 (mailto:your_email@example.com)"
}

def clean_filename(text, max_len=60):
    clean = re.sub(r'[\\/*?:"<>|]', "", text)
    clean = re.sub(r"\s+", "_", clean).strip()
    return clean[:max_len]

def validate_and_download_pdf(url, save_path, timeout=15):
    if not url:
        return False, "لینک وجود ندارد"
    try:
        response = requests.get(url, headers=HEADERS, stream=True, timeout=timeout, allow_redirects=True)
        if response.status_code != 200:
            return False, f"خطای HTTP {response.status_code}"

        content_type = response.headers.get("Content-Type", "").lower()
        is_pdf_header = "application/pdf" in content_type or "octet-stream" in content_type or url.lower().endswith(".pdf")

        first_chunk = next(response.iter_content(chunk_size=200), None)
        if not first_chunk or not (b"%PDF" in first_chunk or is_pdf_header):
            return False, "محتوای لینک HTML است نه PDF"

        with open(save_path, "wb") as f:
            f.write(first_chunk)
            for chunk in response.iter_content(chunk_size=1024 * 64):
                if chunk:
                    f.write(chunk)

        if os.path.getsize(save_path) < 10240:
            os.remove(save_path)
            return False, "فایل دانلودشده خراب یا ناقص است"

        return True, "دانلود موفق"
    except Exception as e:
        return False, f"خطا: {str(e)[:30]}"

def run_pipeline():
    seven_days_ago = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
    api_url = "https://api.openalex.org/works"
    
    params = {
        "search": SEARCH_QUERY,
        "filter": f"has_oa_accepted_or_published_version:true,from_publication_date:{seven_days_ago}",
        "per-page": min(MAX_ARTICLES, 100),
        "sort": "publication_date:desc"
    }

    print("[*] در حال اتصال به OpenAlex...")
    res = requests.get(api_url, params=params, headers=HEADERS)
    if res.status_code != 200:
        print(f"[!] خطا: {res.status_code}")
        return

    results = res.json().get("results", [])
    print(f"[*] تعداد {len(results)} مقاله جدید یافت شد.")

    records = []
    success_count = 0

    for idx, item in enumerate(results, 1):
        title = item.get("title") or "Untitled"
        pub_year = item.get("publication_year") or "Undated"
        doi = item.get("doi") or "No-DOI"

        authorships = item.get("authorships", [])
        author_name = authorships[0].get("author", {}).get("display_name", "Unknown").split()[-1] if authorships else "Unknown"

        oa_info = item.get("open_access", {})
        target_url = oa_info.get("oa_url")
        if not target_url:
            primary_loc = item.get("primary_location") or {}
            target_url = primary_loc.get("pdf_url") or primary_loc.get("landing_page_url")

        file_name = f"{clean_filename(author_name, 15)}_{pub_year}_{clean_filename(title)}.pdf"
        file_path = os.path.join(OUTPUT_DIR, file_name)

        is_downloaded, status_msg = validate_and_download_pdf(target_url, file_path)
        if is_downloaded:
            success_count += 1
            final_file_name = file_name
        else:
            final_file_name = "-"

        records.append({
            "شناسه": idx,
            "عنوان مقاله": title,
            "نویسنده": author_name,
            "سال": pub_year,
            "وضعیت": status_msg,
            "لینک": target_url,
            "DOI": doi,
            "فایل": final_file_name
        })
        time.sleep(0.5)

    df = pd.DataFrame(records)
    df.to_excel(EXCEL_REPORT, index=False)
    print(f"[✔] {success_count} فایل PDF دریافت شد.")

if __name__ == "__main__":
    run_pipeline()
