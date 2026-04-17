import logging
import os
import pickle
import re

import numpy as np

_model = None
_tokenizer = None
_label_encoder = None
_max_len = 100

OUTPUT_DIR = os.getenv("OUTPUT_DIR", "/app/output")

TEENCODE = {
    "ko": "không", "k": "không", "kh": "không", "kg": "không", "khong": "không",
    "dc": "được", "đc": "được", "dk": "được",
    "vs": "với", "mk": "mình", "mn": "mọi người", "mik": "mình",
    "cx": "cũng", "cg": "cũng", "nx": "nữa",
    "bh": "bao giờ", "h": "giờ", "ntn": "như thế nào", "trc": "trước",
    "sp": "sản phẩm", "sdt": "số điện thoại", "dt": "điện thoại",
    "ok": "ổn", "oke": "ổn", "okk": "ổn", "oce": "ổn",
    "tks": "cảm ơn", "thks": "cảm ơn", "thanks": "cảm ơn",
    "gd": "giao diện", "hàng": "hàng", "ship": "giao hàng",
    "nhanh": "nhanh", "chậm": "chậm", "tốt": "tốt", "xấu": "xấu",
}


def load_artifacts():
    global _model, _tokenizer, _label_encoder, _max_len
    if _model is not None:
        return

    logging.info("Loading model artifacts from %s", OUTPUT_DIR)
    from keras.models import load_model

    _model = load_model(os.path.join(OUTPUT_DIR, "sentiment_model_best.keras"))

    with open(os.path.join(OUTPUT_DIR, "tokenizer.pkl"), "rb") as f:
        _tokenizer = pickle.load(f)

    with open(os.path.join(OUTPUT_DIR, "label_encoder.pkl"), "rb") as f:
        _label_encoder = pickle.load(f)

    try:
        with open(os.path.join(OUTPUT_DIR, "preprocess_utils.pkl"), "rb") as f:
            utils = pickle.load(f)
            if isinstance(utils, dict):
                _max_len = utils.get("max_len", 100)
    except Exception:
        pass

    logging.info("Model ready. Classes: %s", list(_label_encoder.classes_))


def _clean_text(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = text.lower().strip()
    text = re.sub(r"https?://\S+|www\.\S+", " ", text)
    text = re.sub(r"\S+@\S+", " ", text)
    text = re.sub(r"\b\d{9,11}\b", " ", text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"(.)\1{2,}", r"\1\1", text)
    text = re.sub(r"[^\w\s\u00C0-\u024F\u1E00-\u1EFF]", " ", text)
    words = [TEENCODE.get(w, w) for w in text.split()]
    return re.sub(r"\s+", " ", " ".join(words)).strip()


def predict(text: str) -> dict:
    load_artifacts()
    from keras.utils import pad_sequences

    cleaned = _clean_text(text)
    if not cleaned:
        return {"emotion": "unknown", "confidence": 0.0, "all_scores": {}}

    try:
        from underthesea import word_tokenize
        processed = " ".join(word_tokenize(cleaned))
    except Exception:
        processed = cleaned

    seq = _tokenizer.texts_to_sequences([processed])
    padded = pad_sequences(seq, maxlen=_max_len, padding="post", truncating="post")

    probs = _model.predict(padded, verbose=0)[0]
    class_idx = int(np.argmax(probs))
    emotion = _label_encoder.inverse_transform([class_idx])[0]
    confidence = float(probs[class_idx])
    all_scores = {cls: float(p) for cls, p in zip(_label_encoder.classes_, probs)}

    return {"emotion": emotion, "confidence": confidence, "all_scores": all_scores}
