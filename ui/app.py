import json
import os
import re
import time
import uuid

import pandas as pd
import plotly.express as px
import psycopg2
import requests
import streamlit as st

st.set_page_config(page_title="Tiki Sentiment Dashboard", layout="wide")

DB_CONFIG = {
    "host":     os.getenv("DB_HOST", "postgres"),
    "database": os.getenv("DB_NAME", "sentiment_db"),
    "user":     os.getenv("DB_USER", "postgres"),
    "password": os.getenv("DB_PASSWORD", "postgres"),
    "port":     int(os.getenv("DB_PORT", 5432)),
}

SENTIMENT_API = os.getenv("SENTIMENT_API_URL", "http://sentiment-api:8013")

TIKI_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/123.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
}
TRACKITY_ID = str(uuid.uuid4())

EMOTION_COLORS = {
    "happiness": "#2ECC71",
    "surprise":  "#F39C12",
    "disgust":   "#8E44AD",
    "sadness":   "#3498DB",
    "anger":     "#E74C3C",
    "fear":      "#95A5A6",
}


# ── Helpers DB ────────────────────────────────────────────────────────────────
def get_conn():
    return psycopg2.connect(**DB_CONFIG)


def ensure_tables():
    try:
        conn = get_conn()
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
        conn.close()
    except Exception as e:
        st.error(f"Lỗi tạo bảng: {e}")


def query(sql, params=None):
    try:
        conn = get_conn()
        with conn.cursor() as cur:
            cur.execute(sql, params)
            rows = cur.fetchall()
            cols = [desc[0] for desc in cur.description]
        conn.close()
        return pd.DataFrame(rows, columns=cols)
    except Exception as e:
        st.error(f"Lỗi DB: {e}")
        return pd.DataFrame()


def execute(sql, params=None):
    conn = get_conn()
    with conn.cursor() as cur:
        cur.execute(sql, params)
        conn.commit()
    conn.close()


ensure_tables()


# ── Helpers Tiki API ──────────────────────────────────────────────────────────
def tiki_get(url, params=None):
    try:
        r = requests.get(url, headers=TIKI_HEADERS, params=params, timeout=10)
        if r.status_code == 200:
            return r.json()
    except Exception:
        pass
    return None


def get_suggestions(keyword):
    data = tiki_get(
        "https://tiki.vn/api/v2/search/suggestion",
        params={"trackity_id": TRACKITY_ID, "q": keyword, "clear": "1"},
    )
    if data and "data" in data:
        return [item["keyword"] for item in data["data"] if item.get("type") == "keyword"]
    return []


def search_products(keyword, page=1, limit=40):
    data = tiki_get(
        "https://tiki.vn/api/v2/products",
        params={
            "trackity_id": TRACKITY_ID,
            "q": keyword,
            "limit": limit,
            "page": page,
            "include": "advertisement",
            "aggregations": "2",
        },
    )
    if not data:
        return [], 0
    products = []
    for p in data.get("data", []):
        products.append({
            "product_id":    str(p.get("id", "")),
            "name":          p.get("name", ""),
            "price":         p.get("price", 0),
            "thumbnail_url": p.get("thumbnail_url", ""),
            "url_key":       p.get("url_key", ""),
            "rating":        p.get("rating_average", 0),
            "review_count":  p.get("review_count", 0),
        })
    total = data.get("paging", {}).get("total", 0)
    return products, total


def fetch_product_info(product_id):
    data = tiki_get(f"https://tiki.vn/api/v2/products/{product_id}")
    if data:
        return {
            "name":          data.get("name", ""),
            "price":         data.get("price", 0),
            "thumbnail_url": data.get("thumbnail_url", ""),
        }
    return None


def extract_ids(url):
    pid  = re.search(r"-p(\d+)", url)
    spid = re.search(r"spid=(\d+)", url)
    return (pid.group(1) if pid else None, spid.group(1) if spid else None)


def fetch_review_stats(product_id):
    """Lấy tổng số đánh giá, rating trung bình và phân bố sao từ Tiki."""
    data = tiki_get(
        "https://tiki.vn/api/v2/reviews",
        params={
            "product_id": product_id, "limit": 1, "page": 1,
            "include": "comments,contribute_info,attribute_vote_summary",
            "sort": "score|desc,id|desc,stars|all",
        },
    )
    if not data:
        return None
    return {
        "total":          data.get("reviews_count") or data.get("paging", {}).get("total", 0),
        "rating_average": data.get("rating_average", 0),
        "stars":          data.get("stars", {}),
    }


def crawl_reviews(pid, use_spid=False, max_pages=10):
    url_type = "spid" if use_spid else "product_id"
    reviews = []
    for page in range(1, max_pages + 1):
        data = tiki_get(
            "https://tiki.vn/api/v2/reviews",
            params={url_type: pid, "limit": 20, "page": page, "include": "comments,photos"},
        )
        if not data:
            break
        items = data.get("data", [])
        if not items:
            break
        for item in items:
            reviews.append({
                "tiki_id":       item.get("id"),
                "rating":        item.get("rating"),
                "review_title":  item.get("title", ""),
                "content":       item.get("content", ""),
                "reviewer_name": (item.get("created_by") or {}).get("name", ""),
                "created_at":    item.get("created_at"),
                "thank_count":   item.get("thank_count", 0),
                "has_photo":     len(item.get("images", [])) > 0,
            })
        time.sleep(0.3)
    return reviews


def call_sentiment(text):
    try:
        r = requests.post(f"{SENTIMENT_API}/predict", json={"text": text}, timeout=30)
        if r.status_code == 200:
            return r.json()
    except Exception:
        pass
    return None


def crawl_and_analyze(product_id, product_name, max_pages, progress_prefix=""):
    """Cào review + phân tích sentiment + lưu DB cho một sản phẩm."""
    reviews = crawl_reviews(product_id, use_spid=False, max_pages=max_pages)
    if not reviews:
        # thử với spid nếu không có kết quả
        reviews = crawl_reviews(product_id, use_spid=True, max_pages=max_pages)

    if not reviews:
        return 0, 0

    bar = st.progress(0, text=f"{progress_prefix}Đang phân tích {product_name}...")
    saved, skipped = 0, 0
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            for i, rv in enumerate(reviews):
                sentiment = call_sentiment(rv["content"]) if rv["content"] else None
                try:
                    cur.execute("""
                        INSERT INTO raw_reviews
                            (tiki_id, product_id, content, rating, reviewer_name,
                             review_title, thank_count, has_photo)
                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
                        ON CONFLICT (tiki_id) DO NOTHING
                        RETURNING id
                    """, (rv["tiki_id"], product_id, rv["content"], rv["rating"],
                          rv["reviewer_name"], rv["review_title"],
                          rv["thank_count"], rv["has_photo"]))
                    row = cur.fetchone()
                    if row and sentiment:
                        cur.execute("""
                            INSERT INTO sentiment_results (review_id, emotion, confidence, all_scores)
                            VALUES (%s,%s,%s,%s)
                        """, (row[0], sentiment["emotion"], sentiment["confidence"],
                              json.dumps(sentiment["all_scores"])))
                        saved += 1
                    else:
                        skipped += 1
                except Exception:
                    skipped += 1
                bar.progress((i + 1) / len(reviews),
                              text=f"{progress_prefix}{product_name}: {i+1}/{len(reviews)}")
        conn.commit()
    finally:
        conn.close()
    bar.empty()
    return saved, skipped


# ═══════════════════════════════════════════════════════════════════════════════
st.title("Tiki Sentiment Dashboard")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Tổng quan",
    "Danh sách sản phẩm",
    "Tìm kiếm & Cào",
    "Cào từ URL",
    "Thử phân tích",
])


# ═══════════════════════════════════════════════════════════════════
# TAB 1 — Tổng quan
# ═══════════════════════════════════════════════════════════════════
with tab1:
    stats = query("""
        SELECT COUNT(*) AS total_reviews, COUNT(sr.id) AS analyzed,
               COUNT(*) - COUNT(sr.id) AS pending
        FROM raw_reviews r
        LEFT JOIN sentiment_results sr ON sr.review_id = r.id
    """)
    if not stats.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("Tổng đánh giá",  int(stats["total_reviews"].iloc[0]))
        c2.metric("Đã phân tích",   int(stats["analyzed"].iloc[0]))
        c3.metric("Chờ xử lý",      int(stats["pending"].iloc[0]))

    st.divider()
    col_l, col_r = st.columns(2)

    with col_l:
        st.subheader("Phân bố cảm xúc")
        dist = query("""
            SELECT emotion, COUNT(*) AS cnt
            FROM sentiment_results GROUP BY emotion ORDER BY cnt DESC
        """)
        if not dist.empty:
            fig = px.pie(dist, values="cnt", names="emotion",
                         color="emotion", color_discrete_map=EMOTION_COLORS, hole=0.35)
            fig.update_traces(textinfo="percent+label")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Chưa có dữ liệu sentiment.")

    with col_r:
        st.subheader("Cảm xúc theo ngày (7 ngày gần nhất)")
        daily = query("""
            SELECT DATE(sr.predicted_at) AS date, sr.emotion, COUNT(*) AS cnt
            FROM sentiment_results sr
            WHERE sr.predicted_at >= NOW() - INTERVAL '7 days'
            GROUP BY DATE(sr.predicted_at), sr.emotion ORDER BY date
        """)
        if not daily.empty:
            fig = px.bar(daily, x="date", y="cnt", color="emotion",
                         color_discrete_map=EMOTION_COLORS, barmode="stack")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Chưa có dữ liệu 7 ngày gần nhất.")

    st.divider()
    conf = query("""
        SELECT emotion, ROUND(AVG(confidence)::NUMERIC, 3) AS avg_conf
        FROM sentiment_results GROUP BY emotion ORDER BY avg_conf DESC
    """)
    if not conf.empty:
        st.subheader("Độ tin cậy trung bình theo cảm xúc")
        fig = px.bar(conf, x="emotion", y="avg_conf", color="emotion",
                     color_discrete_map=EMOTION_COLORS, text_auto=".1%")
        fig.update_layout(yaxis_tickformat=".0%", showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.subheader("Đánh giá gần đây")
    all_emotions = list(EMOTION_COLORS.keys())
    selected = st.multiselect("Lọc cảm xúc:", all_emotions, default=all_emotions, key="dash_filter")
    if selected:
        ph = ",".join(["%s"] * len(selected))
        recent = query(
            f"""
            SELECT r.reviewer_name AS "Người dùng", r.rating AS "Sao",
                   sr.emotion AS "Cảm xúc",
                   ROUND(sr.confidence::NUMERIC, 3) AS "Tin cậy",
                   r.content AS "Nội dung",
                   r.scraped_at AS "Thời gian"
            FROM raw_reviews r
            JOIN sentiment_results sr ON sr.review_id = r.id
            WHERE sr.emotion IN ({ph})
            ORDER BY r.scraped_at DESC LIMIT 100
            """,
            tuple(selected),
        )
        if not recent.empty:
            st.dataframe(recent.style.format({"Tin cậy": "{:.1%}"}),
                         use_container_width=True, hide_index=True)
        else:
            st.info("Không có đánh giá nào.")


# ═══════════════════════════════════════════════════════════════════
# TAB 2 — Danh sách sản phẩm
# ═══════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Danh sách sản phẩm đã cào")

    products_df = query("""
        SELECT
            r.product_id,
            COALESCE(p.name, 'Không rõ') AS name,
            p.thumbnail_url,
            p.price,
            COUNT(r.id)                   AS total_reviews,
            ROUND(AVG(r.rating), 1)       AS avg_rating,
            MODE() WITHIN GROUP (ORDER BY sr.emotion) AS top_emotion
        FROM raw_reviews r
        LEFT JOIN products p ON p.product_id = r.product_id
        LEFT JOIN sentiment_results sr ON sr.review_id = r.id
        GROUP BY r.product_id, p.name, p.thumbnail_url, p.price
        ORDER BY total_reviews DESC
    """)

    if products_df.empty:
        st.info("Chưa có sản phẩm nào trong cơ sở dữ liệu.")
    else:
        st.caption(f"Tổng {len(products_df)} sản phẩm")
        search = st.text_input("Tìm kiếm tên sản phẩm:")
        if search:
            products_df = products_df[
                products_df["name"].str.contains(search, case=False, na=False)
            ]

        for _, row in products_df.iterrows():
            with st.container(border=True):
                col_img, col_info, col_stats = st.columns([1, 3, 2])
                with col_img:
                    if row.get("thumbnail_url"):
                        st.image(row["thumbnail_url"], width=100)
                with col_info:
                    st.markdown(f"**{row['name']}**")
                    st.caption(f"Product ID: {row['product_id']}")
                    if row.get("price"):
                        st.markdown(f"{int(row['price']):,} VNĐ")
                with col_stats:
                    st.metric("Đánh giá", int(row["total_reviews"]))
                    if row.get("avg_rating"):
                        st.markdown(f"Rating: {row['avg_rating']}/5")
                    if row.get("top_emotion"):
                        color = EMOTION_COLORS.get(row["top_emotion"], "#999")
                        st.markdown(
                            f"<span style='color:{color}; font-weight:bold'>"
                            f"{row['top_emotion']}</span>",
                            unsafe_allow_html=True,
                        )

            with st.expander(f"Xem chi tiết — {row['name']}"):
                detail = query("""
                    SELECT sr.emotion, COUNT(*) AS cnt,
                           ROUND(AVG(sr.confidence)::NUMERIC, 2) AS avg_conf
                    FROM raw_reviews r
                    JOIN sentiment_results sr ON sr.review_id = r.id
                    WHERE r.product_id = %s
                    GROUP BY sr.emotion ORDER BY cnt DESC
                """, (row["product_id"],))
                if not detail.empty:
                    fig = px.bar(detail, x="emotion", y="cnt", color="emotion",
                                 color_discrete_map=EMOTION_COLORS, text_auto=True)
                    fig.update_layout(showlegend=False, height=280)
                    st.plotly_chart(fig, use_container_width=True)

                reviews_detail = query("""
                    SELECT r.reviewer_name AS "Người dùng", r.rating AS "Sao",
                           sr.emotion AS "Cảm xúc",
                           ROUND(sr.confidence::NUMERIC, 2) AS "Tin cậy",
                           r.content AS "Nội dung"
                    FROM raw_reviews r
                    JOIN sentiment_results sr ON sr.review_id = r.id
                    WHERE r.product_id = %s
                    ORDER BY r.scraped_at DESC LIMIT 50
                """, (row["product_id"],))
                if not reviews_detail.empty:
                    st.dataframe(
                        reviews_detail.style.format({"Tin cậy": "{:.0%}"}),
                        use_container_width=True, hide_index=True,
                    )


# ═══════════════════════════════════════════════════════════════════
# TAB 3 — Tìm kiếm & Cào
# ═══════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("Tìm kiếm sản phẩm và cào đánh giá")

    # ── Nhập từ khóa ─────────────────────────────────────────────
    col_kw, col_sug, col_num, col_search = st.columns([3, 1, 1, 1])
    with col_kw:
        keyword_input = st.text_input(
            "Từ khóa tìm kiếm:",
            placeholder="Ví dụ: sách, điện thoại...",
        )
    with col_sug:
        st.markdown("<div style='margin-top:28px'>", unsafe_allow_html=True)
        btn_suggest = st.button("Gợi ý", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    with col_num:
        num_products = st.selectbox("Số sản phẩm:", [20, 40, 60, 100], index=1)
    with col_search:
        st.markdown("<div style='margin-top:28px'>", unsafe_allow_html=True)
        btn_search = st.button("Tìm kiếm", type="primary", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # Gợi ý từ khóa
    if btn_suggest and keyword_input.strip():
        with st.spinner("Đang lấy gợi ý..."):
            suggestions = get_suggestions(keyword_input.strip())
        st.session_state["t3_suggestions"] = suggestions if suggestions else []

    if st.session_state.get("t3_suggestions"):
        chosen = st.selectbox(
            "Chọn gợi ý:",
            options=["-- Dùng từ khóa đã nhập --"] + st.session_state["t3_suggestions"],
            key="t3_chosen",
        )
        final_keyword = (
            keyword_input.strip()
            if chosen == "-- Dùng từ khóa đã nhập --"
            else chosen
        )
    else:
        final_keyword = keyword_input.strip()

    # Tìm sản phẩm
    if btn_search and final_keyword:
        with st.spinner(f"Đang tìm '{final_keyword}'..."):
            products, total = search_products(final_keyword, limit=num_products)
        if products:
            st.session_state["t3_products"] = products
            st.session_state["t3_keyword"] = final_keyword
            st.session_state.pop("t3_crawl_target", None)
            st.session_state.pop("t3_crawl_result", None)
        else:
            st.warning("Không tìm thấy sản phẩm nào.")

    # ── Danh sách sản phẩm ───────────────────────────────────────
    if st.session_state.get("t3_products"):
        products = st.session_state["t3_products"]
        max_pages_s = st.slider("Số trang review tối đa:", 1, 20, 5, key="t3_max_pages")

        st.caption(
            f"Tìm thấy {len(products)} sản phẩm cho "
            f"'{st.session_state.get('t3_keyword', '')}'"
        )
        st.divider()

        for idx, p in enumerate(products):
            c_img, c_info, c_link, c_btn = st.columns([1, 6, 1, 1])
            with c_img:
                if p["thumbnail_url"]:
                    st.image(p["thumbnail_url"], width=70)
            with c_info:
                st.markdown(f"**{p['name']}**")
                parts = []
                if p["price"]:
                    parts.append(f"<span style='color:#E74C3C'>{int(p['price']):,} VNĐ</span>")
                if p["rating"]:
                    parts.append(f"Rating: {p['rating']}")
                if p["review_count"]:
                    parts.append(f"{p['review_count']} đánh giá")
                if parts:
                    st.markdown(" &nbsp;|&nbsp; ".join(parts), unsafe_allow_html=True)
                st.caption(f"ID: {p['product_id']}")
            with c_link:
                tiki_url = f"https://tiki.vn/{p['url_key']}" if p.get("url_key") else f"https://tiki.vn/search?q={p['product_id']}"
                st.link_button("Tiki", tiki_url, use_container_width=True)
            with c_btn:
                if st.button("Phân tích", key=f"t3_crawl_{idx}_{p['product_id']}", use_container_width=True):
                    st.session_state["t3_crawl_target"] = p
                    st.session_state.pop("t3_crawl_result", None)

            # Nếu sản phẩm này đang được chọn → hiện kết quả ngay bên dưới
            target = st.session_state.get("t3_crawl_target", {})
            if target.get("product_id") == p["product_id"]:
                with st.container():
                    st.markdown(
                        f"<div style='background:#f0f2f6; border-radius:8px; padding:12px; margin:4px 0'>",
                        unsafe_allow_html=True,
                    )

                    # Chạy cào nếu chưa có kết quả
                    if "t3_crawl_result" not in st.session_state:
                        # Lấy stats tổng từ Tiki
                        stats_tiki = fetch_review_stats(p["product_id"])

                        # Lưu thông tin sản phẩm vào DB
                        try:
                            execute("""
                                INSERT INTO products (product_id, name, price, thumbnail_url, tiki_url)
                                VALUES (%s,%s,%s,%s,%s)
                                ON CONFLICT (product_id) DO UPDATE
                                SET name=EXCLUDED.name, price=EXCLUDED.price,
                                    thumbnail_url=EXCLUDED.thumbnail_url, fetched_at=NOW()
                            """, (p["product_id"], p["name"], p.get("price", 0),
                                  p.get("thumbnail_url", ""),
                                  f"https://tiki.vn/{p.get('url_key', '')}"))
                        except Exception:
                            pass

                        saved, skipped = crawl_and_analyze(
                            p["product_id"], p["name"], max_pages_s
                        )
                        st.session_state["t3_crawl_result"] = {
                            "product_id": p["product_id"],
                            "saved": saved,
                            "skipped": skipped,
                            "stats_tiki": stats_tiki,
                        }
                        st.rerun()

                    result = st.session_state.get("t3_crawl_result", {})
                    if result.get("product_id") == p["product_id"]:
                        total_in_db = result["saved"] + result["skipped"]
                        st.success(
                            f"Hiển thị **{total_in_db}** đánh giá "
                            f"({result['saved']} mới thêm, {result['skipped']} đã có sẵn trong database)."
                        )

                        # Thống kê tổng từ Tiki
                        stats_tiki = result.get("stats_tiki")
                        if stats_tiki:
                            st.markdown("**Thống kê đánh giá trên Tiki**")
                            m1, m2, m3 = st.columns(3)
                            m1.metric("Tổng đánh giá", f"{stats_tiki['total']:,}")
                            m2.metric("Rating trung bình", f"{stats_tiki['rating_average']:.1f} / 5")
                            stars = stats_tiki.get("stars", {})
                            if stars:
                                stars_df = pd.DataFrame([
                                    {"Sao": f"{k} sao", "Số lượng": v["count"], "Tỉ lệ": v["percent"]}
                                    for k, v in sorted(stars.items(), key=lambda x: -int(x[0]))
                                ])
                                with m3:
                                    st.dataframe(stars_df, hide_index=True, use_container_width=True)
                            st.divider()

                        # Biểu đồ sentiment
                        detail = query("""
                            SELECT sr.emotion, COUNT(*) AS cnt
                            FROM raw_reviews r
                            JOIN sentiment_results sr ON sr.review_id = r.id
                            WHERE r.product_id = %s
                            GROUP BY sr.emotion ORDER BY cnt DESC
                        """, (p["product_id"],))
                        if not detail.empty:
                            col_pie, col_bar = st.columns(2)
                            with col_pie:
                                fig = px.pie(
                                    detail, values="cnt", names="emotion",
                                    color="emotion", color_discrete_map=EMOTION_COLORS,
                                    hole=0.35, height=280,
                                )
                                fig.update_traces(textinfo="percent+label")
                                fig.update_layout(showlegend=False, margin=dict(t=20, b=0))
                                st.plotly_chart(fig, use_container_width=True)
                            with col_bar:
                                fig2 = px.bar(
                                    detail, x="emotion", y="cnt", color="emotion",
                                    color_discrete_map=EMOTION_COLORS,
                                    text_auto=True, height=280,
                                )
                                fig2.update_layout(showlegend=False, margin=dict(t=20, b=0))
                                st.plotly_chart(fig2, use_container_width=True)

                        # Bảng đánh giá (tất cả, kể cả trùng)
                        reviews_tbl = query("""
                            SELECT r.reviewer_name AS "Người dùng",
                                   r.rating AS "Sao",
                                   sr.emotion AS "Cảm xúc",
                                   ROUND(sr.confidence::NUMERIC, 2) AS "Tin cậy",
                                   r.content AS "Nội dung"
                            FROM raw_reviews r
                            JOIN sentiment_results sr ON sr.review_id = r.id
                            WHERE r.product_id = %s
                            ORDER BY r.scraped_at DESC
                        """, (p["product_id"],))
                        if not reviews_tbl.empty:
                            st.caption(f"Tổng {len(reviews_tbl)} đánh giá trong database")
                            st.dataframe(
                                reviews_tbl.style.format({"Tin cậy": "{:.0%}"}),
                                use_container_width=True, hide_index=True,
                            )

                    st.markdown("</div>", unsafe_allow_html=True)

            st.divider()


# ═══════════════════════════════════════════════════════════════════
# TAB 4 — Cào từ URL
# ═══════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("Cào đánh giá từ URL sản phẩm Tiki")

    url_input = st.text_input(
        "Nhập URL sản phẩm Tiki:",
        placeholder="https://tiki.vn/san-pham-abc-p123456.html?spid=654321",
    )
    max_pages_url = st.slider("Số trang tối đa:", 1, 20, 5, key="max_pages_url")

    if st.button("Bắt đầu cào", type="primary") and url_input.strip():
        product_id, spid = extract_ids(url_input.strip())

        if not product_id and not spid:
            st.error("Không tìm được product_id hoặc spid từ URL.")
            st.stop()

        # Lấy thông tin sản phẩm
        product_info = None
        if product_id:
            with st.spinner("Đang lấy thông tin sản phẩm..."):
                product_info = fetch_product_info(product_id)

        if product_info:
            st.divider()
            ci, cn = st.columns([1, 4])
            with ci:
                if product_info["thumbnail_url"]:
                    st.image(product_info["thumbnail_url"], width=120)
            with cn:
                st.markdown(f"### {product_info['name']}")
                if product_info["price"]:
                    st.markdown(
                        f"<span style='color:#E74C3C; font-size:1.1rem'>"
                        f"**{int(product_info['price']):,} VNĐ**</span>",
                        unsafe_allow_html=True,
                    )
                st.caption(f"Product ID: {product_id}")
            st.divider()

            try:
                execute("""
                    INSERT INTO products (product_id, name, price, thumbnail_url, tiki_url)
                    VALUES (%s,%s,%s,%s,%s)
                    ON CONFLICT (product_id) DO UPDATE
                    SET name=EXCLUDED.name, price=EXCLUDED.price,
                        thumbnail_url=EXCLUDED.thumbnail_url, fetched_at=NOW()
                """, (product_id, product_info["name"], product_info["price"],
                      product_info["thumbnail_url"], url_input.strip()))
            except Exception:
                pass

        # Xác định id và loại
        use_spid = False
        pid_to_use = product_id
        if product_id:
            data = tiki_get("https://tiki.vn/api/v2/reviews",
                            params={"product_id": product_id, "limit": 1, "page": 1})
            if not data or not data.get("data"):
                pid_to_use = spid
                use_spid = True

        if not pid_to_use:
            st.error("Không tìm được review cho URL này.")
            st.stop()

        pname = product_info["name"] if product_info else (product_id or spid)
        saved, skipped = crawl_and_analyze(pid_to_use, pname, max_pages_url)

        st.success(f"Hoàn tất! Đã lưu **{saved}** đánh giá mới, bỏ qua **{skipped}** trùng.")

        # Hiển thị kết quả từ DB
        st.divider()
        st.subheader("Kết quả phân tích sentiment")
        result_data = query("""
            SELECT sr.emotion, COUNT(*) AS cnt,
                   ROUND(AVG(sr.confidence)::NUMERIC, 2) AS avg_conf
            FROM raw_reviews r
            JOIN sentiment_results sr ON sr.review_id = r.id
            WHERE r.product_id = %s
            GROUP BY sr.emotion ORDER BY cnt DESC
        """, (product_id or spid,))
        if not result_data.empty:
            col_c, col_t = st.columns(2)
            with col_c:
                fig = px.pie(result_data, values="cnt", names="emotion",
                             color="emotion", color_discrete_map=EMOTION_COLORS,
                             hole=0.35)
                st.plotly_chart(fig, use_container_width=True)
            with col_t:
                st.dataframe(result_data, use_container_width=True, hide_index=True)

        reviews_show = query("""
            SELECT r.reviewer_name AS "Người dùng", r.rating AS "Sao",
                   sr.emotion AS "Cảm xúc",
                   ROUND(sr.confidence::NUMERIC, 2) AS "Tin cậy",
                   r.content AS "Nội dung"
            FROM raw_reviews r
            JOIN sentiment_results sr ON sr.review_id = r.id
            WHERE r.product_id = %s
            ORDER BY r.scraped_at DESC LIMIT 100
        """, (product_id or spid,))
        if not reviews_show.empty:
            st.dataframe(
                reviews_show.style.format({"Tin cậy": "{:.0%}"}),
                use_container_width=True, hide_index=True,
            )


# ═══════════════════════════════════════════════════════════════════
# TAB 5 — Thử phân tích
# ═══════════════════════════════════════════════════════════════════
with tab5:
    st.subheader("Thử phân tích cảm xúc")
    st.caption("Nhập một đoạn văn bản để kiểm tra mô hình sentiment.")

    manual_text = st.text_area(
        "Nội dung đánh giá:", height=150,
        placeholder="Sản phẩm chất lượng tốt, giao hàng nhanh...",
    )
    if st.button("Phân tích", type="primary") and manual_text.strip():
        with st.spinner("Đang phân tích..."):
            res = call_sentiment(manual_text)
        if res:
            emotion = res["emotion"]
            color = EMOTION_COLORS.get(emotion, "#999")
            st.markdown(
                f"<div style='background:{color}22; border-left:4px solid {color}; "
                f"padding:12px; border-radius:4px; margin-bottom:12px;'>"
                f"<b style='color:{color}; font-size:1.3rem'>{emotion}</b>"
                f" &nbsp;&nbsp; {res['confidence']:.1%}</div>",
                unsafe_allow_html=True,
            )
            scores_df = (
                pd.DataFrame(list(res["all_scores"].items()), columns=["Cảm xúc", "Xác suất"])
                .sort_values("Xác suất", ascending=False)
            )
            fig = px.bar(scores_df, x="Cảm xúc", y="Xác suất",
                         color="Cảm xúc", color_discrete_map=EMOTION_COLORS,
                         text_auto=".1%")
            fig.update_layout(showlegend=False, yaxis_tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(scores_df.style.format({"Xác suất": "{:.1%}"}),
                         hide_index=True, use_container_width=True)
        else:
            st.error("Không kết nối được API sentiment.")
