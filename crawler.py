import os
import re
import time
import socket
import requests
import pandas as pd

# تنظیم مهلت اتصال به سرورها (حداکثر ۱۰ ثانیه)
socket.setdefaulttimeout(10)

OUTPUT_DIR = "downloaded_pdfs"
EXCEL_REPORT = "download_report.xlsx"
os.makedirs(OUTPUT_DIR, exist_ok=True)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AcademicHarvester/3.0"
}

# دسته‌بندی موضوعی به تفکیک حوزه‌ها و تلاقی آن‌ها
TOPIC_CONFIGS = [
    {
        "category_fa": "تلاقی موضوعات",
        "category_en": "Intersection",
        "query": "modernity development \"sociology of organizations\" bureaucracy",
        "limit": 15
    },
    {
        "category_fa": "مدرنیته",
        "category_en": "Modernity",
        "query": "modernity rationalization \"classical sociology\" Weberian",
        "limit": 12
    },
    {
        "category_fa": "توسعه",
        "category_en": "Development",
        "query": "\"development theory\" \"institutional development\" modernization",
        "limit": 12
    },
    {
        "category_fa": "جامعه‌شناسی سازمان‌ها",
        "category_en": "Org_Sociology",
        "query": "\"sociology of organizations\" \"organizational sociology\" bureaucracy \"neo-institutionalism\"",
        "limit": 12
    }
]

# معادل‌سازی انواع اسناد علمی در OpenAlex
DOCUMENT_TYPES = {
    "article": "مقاله",
    "book": "کتاب",
    "dissertation": "پایان‌نامه-رساله",
    "book-chapter": "فصل کتاب",
    "report": "گزارش علمی",
    "other": "سند علمی"
}

def clean_filename(text, max_len=40):
    clean = re.sub(r'[\\/*?:"<>|]', "", str(text))
    clean = re.sub(r"\s+", "_", clean).strip()
    return clean[:max_len]

def validate_and_download_pdf(url, save_path):
    if not url:
        return False, "فاقد لینک معتبر"
    try:
        response = requests.get(url, headers=HEADERS, stream=True, timeout=(5, 10), allow_redirects=True)
        if response.status_code != 200:
            return False, f"HTTP {response.status_code}"

        content_type = response.headers.get("Content-Type", "").lower()
        is_pdf_header = "application/pdf" in content_type or "octet-stream" in content_type or url.lower().endswith(".pdf")

        first_chunk = next(response.iter_content(chunk_size=200), None)
        if not first_chunk or not (b"%PDF" in first_chunk or is_pdf_header):
            return False, "فایل PDF نیست (صفحه وب است)"

        with open(save_path, "wb") as f:
            f.write(first_chunk)
            downloaded = len(first_chunk)
            for chunk in response.iter_content(chunk_size=1024 * 64):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if downloaded > 15 * 1024 * 1024:  # سقف ۱۵ مگابایت برای هر سند
                        break

        if os.path.getsize(save_path) < 10240:
            if os.path.exists(save_path):
                os.remove(save_path)
            return False, "فایل ناقص یا خراب"

        return True, "دانلود موفق"
    except requests.exceptions.Timeout:
        return False, "تایم‌اوت سرور"
    except Exception as e:
        return False, f"خطا: {str(e)[:20]}"

def run_pipeline():
    api_url = "https://api.openalex.org/works"
    all_records = []
    seen_ids = set()
    total_downloaded = 0

    print("[*] آغاز فرآیند استخراج چندگانه بدون محدودیت زمانی...", flush=True)

    for conf in TOPIC_CONFIGS:
        category_fa = conf["category_fa"]
        category_en = conf["category_en"]
        query = conf["query"]
        limit = conf["limit"]

        print(f"\n==================================================", flush=True)
        print(f"[*] کاوش حوزه: {category_fa} ({category_en})", flush=True)
        print(f"==================================================", flush=True)

        params = {
            "search": query,
            # استخراج انواع اسناد: مقالات، کتاب‌ها، رساله‌ها و فصول کتاب با دسترسی آزاد
            "filter": "has_oa_accepted_or_published_version:true,type:article|book|dissertation|book-chapter",
            "per-page": limit,
            "sort": "cited_by_count:desc"  # اولویت قطعی با معتبرترین و پراستنادترین اسناد
        }

        try:
            res = requests.get(api_url, params=params, headers=HEADERS, timeout=12)
        except Exception as e:
            print(f"[!] خطا در ارتباط با پایگاه داده برای حوزه {category_fa}: {e}", flush=True)
            continue

        if res.status_code != 200:
            print(f"[!] خطا از سرور: {res.status_code}", flush=True)
            continue

        results = res.json().get("results", [])
        print(f"[*] {len(results)} اثر معتبر یافت شد. آغاز پردازش...", flush=True)

        for idx, item in enumerate(results, 1):
            work_id = item.get("id")
            if work_id in seen_ids:
                continue
            seen_ids.add(work_id)

            title = item.get("title") or "Untitled"
            pub_year = item.get("publication_year") or "Undated"
            doc_type_raw = item.get("type") or "other"
            doc_type_fa = DOCUMENT_TYPES.get(doc_type_raw, "سند علمی")
            citations = item.get("cited_by_count", 0)

            authorships = item.get("authorships", [])
            author_name = authorships[0].get("author", {}).get("display_name", "Unknown").split()[-1] if authorships else "Unknown"

            oa_info = item.get("open_access", {})
            target_url = oa_info.get("oa_url")
            if not target_url:
                primary_loc = item.get("primary_location") or {}
                target_url = primary_loc.get("pdf_url") or primary_loc.get("landing_page_url")

            # نام‌گذاری مشخص برای تمایز اسناد در NotebookLM:
            # مثال: [Modernity]_[کتاب]_Weber_1922_Economy_and_Society.pdf
            safe_author = clean_filename(author_name, 15)
            safe_title = clean_filename(title, 35)
            file_name = f"[{category_en}]_[{doc_type_fa}]_{safe_author}_{pub_year}_{safe_title}.pdf"
            file_path = os.path.join(OUTPUT_DIR, file_name)

            print(f"[{idx}] {doc_type_fa} | {safe_title}...", end=" ", flush=True)
            is_downloaded, status_msg = validate_and_download_pdf(target_url, file_path)
            print(f"-> {status_msg}", flush=True)

            if is_downloaded:
                total_downloaded += 1
                final_file = file_name
            else:
                final_file = "-"

            all_records.append({
                "حوزه موضوعی": category_fa,
                "نوع سند": doc_type_fa,
                "عنوان اثر": title,
                "نویسنده اصلی": author_name,
                "سال انتشار": pub_year,
                "تعداد ارجاعات": citations,
                "وضعیت دسترسی": status_msg,
                "لینک دانلود": target_url,
                "شناسه سند": work_id,
                "فایل ذخیره‌شده": final_file
            })

    # خروجی نهایی اکسل با تمام تفکیک‌ها
    df = pd.DataFrame(all_records)
    df.to_excel(EXCEL_REPORT, index=False)
    print(f"\n[✔] فرآیند تکمیل شد. در مجموع {total_downloaded} فایل معتبر دانلود و در اکسل ثبت شد.", flush=True)

if __name__ == "__main__":
    run_pipeline()
