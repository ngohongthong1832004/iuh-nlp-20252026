import json
import logging
import os
import time
import uuid
from datetime import datetime, timedelta

import psycopg2
import psycopg2.extras
import requests
from airflow import DAG
from airflow.models import Variable
from airflow.operators.python import PythonOperator

# ── Cấu hình ──────────────────────────────────────────────────────────────────
DB_CONN = {
    "host":     os.getenv("DB_HOST", "postgres"),
    "database": os.getenv("DB_NAME", "sentiment_db"),
    "user":     os.getenv("DB_USER", "postgres"),
    "password": os.getenv("DB_PASSWORD", "postgres"),
    "port":     int(os.getenv("DB_PORT", 5432)),
}

SENTIMENT_API_URL = os.getenv("SENTIMENT_API_URL", "http://sentiment-api:8013")

TIKI_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "vi-VN,vi;q=0.9",
    "Referer": "https://tiki.vn/",
}

# Từ khóa mặc định — có thể override qua Airflow Variable "tiki_keywords"
DEFAULT_KEYWORDS = ["sách", "điện thoại", "máy tính bảng", "tai nghe", "đồ gia dụng"]


# ── Helpers ───────────────────────────────────────────────────────────────────
def tiki_get(url, params=None, retries=3):
    for attempt in range(retries):
        try:
            r = requests.get(url, headers=TIKI_HEADERS, params=params, timeout=15)
            if r.status_code == 200:
                return r.json()
            logging.warning(f"Tiki API status {r.status_code}: {url}")
        except Exception as e:
            logging.warning(f"Tiki API lỗi (lần {attempt+1}): {e}")
            time.sleep(2 ** attempt)
    return None


def ensure_products_table(conn):
    with conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS products (
                product_id    VARCHAR(50) PRIMARY KEY,
                name          TEXT,
                price         BIGINT,
                thumbnail_url TEXT,
                tiki_url      TEXT,
                fetched_at    TIMESTAMP DEFAULT NOW()
            )
        """)
    conn.commit()


# ── Task 1: Tìm sản phẩm qua API ─────────────────────────────────────────────
def discover_products(**context):
    """
    Tìm sản phẩm từ Tiki search API theo danh sách từ khóa.
    Từ khóa đọc từ:
      1. dag_run.conf["keywords"]  (khi trigger thủ công)
      2. Airflow Variable "tiki_keywords" (JSON array)
      3. DEFAULT_KEYWORDS
    """
    conf = context.get("dag_run").conf or {}

    # Lấy keywords
    if "keywords" in conf:
        keywords = conf["keywords"] if isinstance(conf["keywords"], list) else [conf["keywords"]]
    else:
        try:
            raw = Variable.get("tiki_keywords", default_var=None)
            keywords = json.loads(raw) if raw else DEFAULT_KEYWORDS
        except Exception:
            keywords = DEFAULT_KEYWORDS

    products_per_keyword = int(conf.get("products_per_keyword", 20))
    trackity_id = str(uuid.uuid4())

    conn = psycopg2.connect(**DB_CONN)
    ensure_products_table(conn)

    all_product_ids = []

    for keyword in keywords:
        logging.info(f"Tìm sản phẩm cho từ khóa: '{keyword}'")
        data = tiki_get(
            "https://tiki.vn/api/v2/products",
            params={
                "q": keyword,
                "limit": products_per_keyword,
                "page": 1,
                "trackity_id": trackity_id,
                "include": "advertisement",
                "aggregations": "2",
            },
        )
        if not data:
            logging.warning(f"Không lấy được sản phẩm cho '{keyword}'")
            continue

        products = data.get("data", [])
        logging.info(f"Tìm thấy {len(products)} sản phẩm cho '{keyword}'")

        with conn.cursor() as cur:
            for p in products:
                pid = str(p.get("id", ""))
                if not pid:
                    continue
                all_product_ids.append(pid)
                try:
                    cur.execute("""
                        INSERT INTO products (product_id, name, price, thumbnail_url, tiki_url)
                        VALUES (%s, %s, %s, %s, %s)
                        ON CONFLICT (product_id) DO UPDATE
                        SET name=EXCLUDED.name, price=EXCLUDED.price,
                            thumbnail_url=EXCLUDED.thumbnail_url, fetched_at=NOW()
                    """, (
                        pid,
                        p.get("name", ""),
                        p.get("price", 0),
                        p.get("thumbnail_url", ""),
                        f"https://tiki.vn/{p.get('url_key', '')}",
                    ))
                except Exception as e:
                    logging.error(f"Lỗi lưu product {pid}: {e}")
        conn.commit()
        time.sleep(1)

    conn.close()

    # Loại trùng và đẩy qua XCom cho task tiếp theo
    unique_ids = list(dict.fromkeys(all_product_ids))
    logging.info(f"Tổng {len(unique_ids)} sản phẩm duy nhất sẽ được cào.")
    context["ti"].xcom_push(key="product_ids", value=unique_ids)
    return unique_ids


# ── Task 2: Cào review ────────────────────────────────────────────────────────
def scrape_tiki_reviews(**context):
    """Cào review từ danh sách product_id lấy từ task discover_products."""
    product_ids = context["ti"].xcom_pull(task_ids="discover_products", key="product_ids") or []
    conf = context.get("dag_run").conf or {}
    max_pages = int(conf.get("max_pages", 5))

    if not product_ids:
        logging.warning("Không có product_id nào để cào.")
        return 0

    conn = psycopg2.connect(**DB_CONN)
    cur = conn.cursor()
    total_saved = 0

    for product_id in product_ids:
        logging.info(f"Cào review product_id={product_id}")
        for page in range(1, max_pages + 1):
            data = tiki_get(
                "https://tiki.vn/api/v2/reviews",
                params={
                    "product_id": product_id,
                    "page": page,
                    "limit": 20,
                    "sort": "score|desc,id|desc,stars|all",
                    "include": "comments,photos",
                },
            )
            if not data:
                break

            reviews = data.get("data", [])
            if not reviews:
                break

            for r in reviews:
                content = (r.get("content") or "").strip()
                if not content:
                    continue
                try:
                    cur.execute(
                        """
                        INSERT INTO raw_reviews
                            (tiki_id, product_id, content, rating, reviewer_name,
                             review_title, created_at, thank_count, has_photo)
                        VALUES (%s, %s, %s, %s, %s, %s, to_timestamp(%s), %s, %s)
                        ON CONFLICT (tiki_id) DO NOTHING
                        """,
                        (
                            r.get("id"),
                            product_id,
                            content,
                            r.get("rating", 3),
                            (r.get("created_by") or {}).get("name", ""),
                            r.get("title", ""),
                            r.get("created_at"),
                            r.get("thank_count", 0),
                            len(r.get("images", [])) > 0,
                        ),
                    )
                    total_saved += 1
                except Exception as e:
                    logging.error(f"Lỗi insert review: {e}")
                    conn.rollback()

            conn.commit()
            time.sleep(1)

        time.sleep(0.5)

    cur.close()
    conn.close()
    logging.info(f"Tổng review mới lưu: {total_saved}")
    return total_saved


# ── Task 3: Phân tích sentiment ───────────────────────────────────────────────
def run_sentiment_analysis(**context):
    """Phân tích sentiment cho các review chưa được xử lý."""
    conn = psycopg2.connect(**DB_CONN)
    cur = conn.cursor()

    cur.execute("""
        SELECT r.id, r.content
        FROM raw_reviews r
        LEFT JOIN sentiment_results sr ON sr.review_id = r.id
        WHERE sr.id IS NULL
        LIMIT 500
    """)
    reviews = cur.fetchall()
    logging.info(f"Phân tích sentiment cho {len(reviews)} review chưa xử lý.")

    for review_id, content in reviews:
        try:
            resp = requests.post(
                f"{SENTIMENT_API_URL}/predict",
                json={"text": content},
                timeout=15,
            )
            if resp.status_code == 200:
                result = resp.json()
                cur.execute(
                    """
                    INSERT INTO sentiment_results (review_id, emotion, confidence, all_scores)
                    VALUES (%s, %s, %s, %s)
                    ON CONFLICT DO NOTHING
                    """,
                    (
                        review_id,
                        result["emotion"],
                        result["confidence"],
                        psycopg2.extras.Json(result.get("all_scores", {})),
                    ),
                )
                conn.commit()
            else:
                logging.warning(f"Sentiment API lỗi {resp.status_code} cho review {review_id}")
        except Exception as e:
            logging.error(f"Lỗi sentiment review {review_id}: {e}")

    cur.close()
    conn.close()


# ── DAG definition ────────────────────────────────────────────────────────────
default_args = {
    "owner": "airflow",
    "depends_on_past": False,
    "start_date": datetime(2024, 1, 1),
    "retries": 2,
    "retry_delay": timedelta(minutes=5),
}

with DAG(
    "tiki_sentiment_pipeline",
    default_args=default_args,
    description="Tìm sản phẩm Tiki qua API → Cào review → Phân tích sentiment",
    schedule_interval="@daily",
    catchup=False,
    tags=["tiki", "sentiment", "nlp"],
    params={
        "keywords": ["sách", "điện thoại"],
        "products_per_keyword": 20,
        "max_pages": 5,
    },
) as dag:
    t_discover  = PythonOperator(task_id="discover_products",      python_callable=discover_products)
    t_scrape    = PythonOperator(task_id="scrape_tiki_reviews",    python_callable=scrape_tiki_reviews)
    t_sentiment = PythonOperator(task_id="run_sentiment_analysis", python_callable=run_sentiment_analysis)

    t_discover >> t_scrape >> t_sentiment
