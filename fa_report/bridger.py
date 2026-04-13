import json
import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

import cv2


IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}
DEFAULT_CONFIG_FILENAME = "user_settings.json"
CONFIG_KEY_IMAGE_DIR = "classification_image_dir"


def default_config_path() -> str:
    return str(Path(__file__).resolve().with_name(DEFAULT_CONFIG_FILENAME))


def load_config(path: str | None = None) -> dict:
    cfg_path = Path(path or default_config_path())
    if not cfg_path.is_file():
        return {}
    try:
        return json.loads(cfg_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def save_config(data: dict, path: str | None = None) -> None:
    cfg_path = Path(path or default_config_path())
    cfg_path.parent.mkdir(parents=True, exist_ok=True)
    cfg_path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def get_image_dir_from_config(path: str | None = None) -> str:
    cfg = load_config(path)
    value = cfg.get(CONFIG_KEY_IMAGE_DIR, "")
    return str(value).strip() if isinstance(value, str) else ""


def set_image_dir_in_config(image_dir: str, path: str | None = None) -> None:
    cfg = load_config(path)
    cfg[CONFIG_KEY_IMAGE_DIR] = (image_dir or "").strip()
    save_config(cfg, path)


def normalize_vit(vit_id: str) -> str:
    text = (vit_id or "").strip().upper()
    if text.isdigit():
        return f"VIT-{text}"
    if text.startswith("VIT") and not text.startswith("VIT-"):
        return f"VIT-{text[3:]}"
    return text


def _vit_digits(vit_id: str) -> str:
    m = re.search(r"VIT[^0-9]*([0-9]+)", vit_id, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r"([0-9]+)", vit_id)
    return m.group(1) if m else ""


def _filename_matches_vit(filename_stem: str, vit_norm: str) -> bool:
    digits = _vit_digits(vit_norm)
    if not digits:
        return False
    patterns = [
        rf"^VIT\s*-?\s*{re.escape(digits)}(?:[_\-\s]|$)",
        rf"^VI?I?T\s*-?\s*{re.escape(digits)}(?:[_\-\s]|$)",
        rf"^{re.escape(digits)}(?:[_\-\s]|$)",
        rf"VIT\s*-?\s*{re.escape(digits)}(?:[_\-\s]|$)",
    ]
    return any(re.search(p, filename_stem, flags=re.IGNORECASE) for p in patterns)


def find_images_for_vit(image_dir: str, vit_id: str) -> list[str]:
    base = Path((image_dir or "").strip())
    if not base.is_dir():
        return []
    vit_norm = normalize_vit(vit_id)
    matches: list[str] = []
    for path in base.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in IMAGE_EXTS:
            continue
        if _filename_matches_vit(path.stem, vit_norm):
            matches.append(str(path))
    return sorted(matches)


def _load_main_classifier_module():
    try:
        import fa_report.image_classifier as main_mod
        return main_mod
    except Exception:
        return None


def _map_category_to_camh(category: str) -> str:
    normalized = (category or "").strip().lower()
    mapping = {
        "whole segment": "Whole Segment",
        "white segment": "White Segment",
        "die segment": "<=2 segment",
        "line between": "Line Between Segment",
        "pending": "",
    }
    return mapping.get(normalized, "")


@dataclass
class FallbackResult:
    camh_value: str
    matched_images: int
    classified_images: int
    pending_votes: int


def classify_camh_from_vit(vit_id: str, image_dir: str) -> FallbackResult:
    image_paths = find_images_for_vit(image_dir, vit_id)
    if not image_paths:
        return FallbackResult(camh_value="", matched_images=0, classified_images=0, pending_votes=0)

    main_mod = _load_main_classifier_module()
    if main_mod is None or not hasattr(main_mod, "classify_image"):
        return FallbackResult(
            camh_value="",
            matched_images=len(image_paths),
            classified_images=0,
            pending_votes=0,
        )

    camh_votes: list[str] = []
    pending_votes = 0
    for path in image_paths:
        img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
        if img is None:
            continue
        try:
            result = main_mod.classify_image(img)
            cat = (result or {}).get("category", "")
            
            # Fallback: if model is not loaded or returned pending, try rule-based check directly
            if cat.lower() == "pending" or result.get("method") == "No Model":
                if hasattr(main_mod, "run_rule_based_check"):
                    rule_cat = main_mod.run_rule_based_check(img)
                    if rule_cat and rule_cat.lower() != "pending":
                        cat = rule_cat
                        
        except Exception:
            continue
            
        cat_norm = (cat or "").strip().lower()
        if cat_norm == "pending":
            pending_votes += 1
            continue
        camh = _map_category_to_camh(cat)
        if camh:
            camh_votes.append(camh)

    if not camh_votes:
        return FallbackResult(
            camh_value="",
            matched_images=len(image_paths),
            classified_images=0,
            pending_votes=pending_votes,
        )

    counts = Counter(camh_votes)
    value = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
    return FallbackResult(
        camh_value=value,
        matched_images=len(image_paths),
        classified_images=len(camh_votes),
        pending_votes=pending_votes,
    )
