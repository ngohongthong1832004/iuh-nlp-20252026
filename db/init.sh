#!/bin/bash
set -e

psql -v ON_ERROR_STOP=1 -U "$POSTGRES_USER" -c "CREATE DATABASE airflow;" || echo "airflow db already exists"

psql -v ON_ERROR_STOP=1 -U "$POSTGRES_USER" -d sentiment_db <<-EOSQL
    CREATE TABLE IF NOT EXISTS products (
        product_id    VARCHAR(50) PRIMARY KEY,
        name          TEXT,
        price         BIGINT,
        thumbnail_url TEXT,
        tiki_url      TEXT,
        fetched_at    TIMESTAMP DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS raw_reviews (
        id            SERIAL PRIMARY KEY,
        tiki_id       BIGINT UNIQUE,
        product_id    VARCHAR(50),
        content       TEXT NOT NULL,
        rating        INTEGER CHECK (rating BETWEEN 1 AND 5),
        reviewer_name VARCHAR(200),
        review_title  VARCHAR(500),
        created_at    TIMESTAMP,
        thank_count   INTEGER DEFAULT 0,
        has_photo     BOOLEAN DEFAULT FALSE,
        scraped_at    TIMESTAMP DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS sentiment_results (
        id           SERIAL PRIMARY KEY,
        review_id    INTEGER REFERENCES raw_reviews(id) ON DELETE CASCADE,
        emotion      VARCHAR(50) NOT NULL,
        confidence   FLOAT NOT NULL,
        all_scores   JSONB,
        predicted_at TIMESTAMP DEFAULT NOW()
    );

    CREATE INDEX IF NOT EXISTS idx_reviews_product  ON raw_reviews(product_id);
    CREATE INDEX IF NOT EXISTS idx_reviews_scraped  ON raw_reviews(scraped_at);
    CREATE INDEX IF NOT EXISTS idx_sentiment_emotion ON sentiment_results(emotion);
    CREATE INDEX IF NOT EXISTS idx_sentiment_review  ON sentiment_results(review_id);
EOSQL
