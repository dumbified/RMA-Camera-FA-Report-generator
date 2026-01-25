import os
import csv
from pathlib import Path

import cv2
import numpy as np


# --- CONFIGURATION ---
INPUT_FOLDER = Path("samples")
OUTPUT_FOLDER = Path("classified_visualizations")
DEBUG_CSV_PATH = Path("segment_debug.csv")

LABEL_TO_FOLDER = {
    "PERFECT": "Perfect",
    "WHOLE SEGMENT": "Whole Segment",
    "DIE SEGMENT": "Die Segment",
    "WHITE SEGMENT": "White Segment",
}

NUM_SEGMENTS = 10
DEFECT_RATIO_LIMIT = 0.05  # 5% of pixels in a segment
IMAGE_EXTS = (".jpg", ".jpeg", ".png")

# TUNING PARAMETERS
MIN_VALID_PEAK = 50       # Threshold separating "Black Background" from "Grey Object"
RIGHT_SIGMA = 1.9         # Statistical multiplier for White Segment detection
PEAK_SENSITIVITY = 0.05   # The Grey Peak must be at least 5% height of the Black Peak to be valid

# Mean-based detection (from main.py style)
ABS_DIE_MEAN = 55
ABS_WHITE_MEAN = 160
WHITE_OUTLIER_SIGMA = 2.5
WHITE_OUTLIER_MIN_DIFF = 15

COLOR_OK = (0, 200, 0)
COLOR_DIE = (0, 0, 255)
COLOR_WHITE = (0, 200, 255)
COLOR_TEXT_BG = (0, 0, 0)


def is_image_file(path: Path) -> bool:
    return path.suffix.lower() in IMAGE_EXTS


def get_safe_zone_thresholds(img):
    """
    Calculates thresholds using Priority Peak Logic:
    1. SEARCH: Look for a significant peak in the Valid Zone (>50).
    2. FALLBACK: If no valid peak exists, select the Background Peak (<50).
    3. VALIDATE: If selected peak is < 50, flag as WHOLE SEGMENT DIE.
    """
    # 1. Histogram Calculation & Smoothing
    hist = cv2.calcHist([img], [0], None, [256], [0, 256]).flatten()

    kernel_size = 9
    sigma = 2.0
    x = np.linspace(-kernel_size // 2, kernel_size // 2, kernel_size)
    kernel = np.exp(-0.5 * (x / sigma) ** 2)
    kernel /= kernel.sum()
    hist_smooth = np.convolve(hist, kernel, mode="same")

    # 2. Identify Potential Peaks
    # Global Max (likely the Black Background if image is mostly dark)
    global_max_idx = np.argmax(hist_smooth)
    global_max_height = hist_smooth[global_max_idx]

    # Signal Max (The tallest point in the "Grey Area" > 50)
    # We slice [MIN_VALID_PEAK:] to ignore the dark background
    signal_region = hist_smooth[MIN_VALID_PEAK:]
    
    if signal_region.size > 0:
        rel_signal_idx = np.argmax(signal_region)
        signal_idx = rel_signal_idx + MIN_VALID_PEAK
        signal_height = hist_smooth[signal_idx]
    else:
        signal_height = 0
        signal_idx = 0

    # 3. Peak Selection Logic
    # We prefer the 'signal_idx' (Grey Peak) IF it is significant enough.
    # It must be at least X% height of the global max (to ensure it's not just flat noise).
    if signal_height > (PEAK_SENSITIVITY * global_max_height):
        # We found a valid Grey Mountain! Use it.
        main_peak_idx = signal_idx
        peak_height = signal_height
        print(f"  -> Found Grey Peak at {main_peak_idx} (Height: {signal_height:.0f})")
    else:
        # No significant grey object found. The Black Spike is the only feature.
        main_peak_idx = global_max_idx
        peak_height = global_max_height
        print(f"  -> No Grey Object. Locked to Background Peak at {main_peak_idx}")

    # 4. Critical Check: Is the selected peak valid?
    if main_peak_idx < MIN_VALID_PEAK:
        print(f"  -> Peak {main_peak_idx} is too dark (<{MIN_VALID_PEAK}). Flagging as DIE.")
        return 255, 255

    # 5. Calculate Left Threshold (Foot of Mountain Logic)
    # Walk left from peak until we hit the noise floor (< 10% of peak height)
    left_thresh = 0
    for i in range(main_peak_idx, 1, -1):
        if hist_smooth[i] < (0.1 * peak_height):
            left_thresh = i
            break

    # 6. Calculate Right Threshold (Symmetric 10% peak cutoff)
    right_thresh = 255
    for i in range(main_peak_idx, 255):
        if hist_smooth[i] < (0.1 * peak_height):
            right_thresh = i
            break

    return left_thresh, right_thresh


def segment_bounds(width: int, num_segments: int):
    edges = np.linspace(0, width, num_segments + 1, dtype=int)
    return [(int(edges[i]), int(edges[i + 1])) for i in range(num_segments)]


def classify_segment(
    seg_low,
    seg_high,
    seg_gray,
    defect_ratio_limit: float,
    mid_gray: float,
    seg_mean: float,
    dynamic_white_thresh: float,
):
    total_pixels = seg_low.size
    if total_pixels == 0:
        return "OK", 0.0, 0.0

    count_dark = cv2.countNonZero(seg_low)
    count_white = cv2.countNonZero(seg_high)
    ratio_dark = count_dark / total_pixels
    ratio_white = count_white / total_pixels

    # Mean-based hard rules (quick wins)
    if seg_mean < ABS_DIE_MEAN:
        return "DIE", ratio_dark, ratio_white
    if seg_mean > ABS_WHITE_MEAN or seg_mean > dynamic_white_thresh:
        return "WHITE", ratio_dark, ratio_white

    # Mask-based rules (sensitive to localized defects)
    if ratio_dark > defect_ratio_limit and ratio_white > defect_ratio_limit:
        if seg_mean < mid_gray:
            return "DIE", ratio_dark, ratio_white
        return "WHITE", ratio_dark, ratio_white
    if ratio_dark > defect_ratio_limit:
        return "DIE", ratio_dark, ratio_white
    if ratio_white > defect_ratio_limit:
        return "WHITE", ratio_dark, ratio_white
    return "OK", ratio_dark, ratio_white


def classify_image(segment_labels):
    if all(label == "DIE" for label in segment_labels):
        return "WHOLE SEGMENT"
    if "DIE" in segment_labels:
        return "DIE SEGMENT"
    if "WHITE" in segment_labels:
        return "WHITE SEGMENT"
    return "PERFECT"


def draw_label(img, text, x, y, color):
    font = cv2.FONT_HERSHEY_SIMPLEX
    scale = 0.5
    thickness = 1
    text_size, _ = cv2.getTextSize(text, font, scale, thickness)
    text_w, text_h = text_size

    pad = 3
    box_x1 = max(0, x - pad)
    box_y1 = max(0, y - text_h - pad * 2)
    box_x2 = min(img.shape[1] - 1, x + text_w + pad)
    box_y2 = min(img.shape[0] - 1, y + pad)

    cv2.rectangle(img, (box_x1, box_y1), (box_x2, box_y2), COLOR_TEXT_BG, -1)
    cv2.putText(img, text, (x, y - pad), font, scale, color, thickness, cv2.LINE_AA)


def annotate_image(img_color, segment_labels, bounds, image_label):
    for i, (x1, x2) in enumerate(bounds):
        label = segment_labels[i]
        if label == "DIE":
            color = COLOR_DIE
        elif label == "WHITE":
            color = COLOR_WHITE
        else:
            color = COLOR_OK

        cv2.rectangle(img_color, (x1, 0), (x2, img_color.shape[0]), color, 2)
        draw_label(img_color, label, x1 + 5, 20, color)

    draw_label(img_color, image_label, 10, img_color.shape[0] - 5, (255, 255, 255))
    return img_color


def process_image(img_gray):
    low_cut, high_cut = get_safe_zone_thresholds(img_gray)
    mid_gray = (low_cut + high_cut) / 2.0

    _, mask_low = cv2.threshold(img_gray, low_cut, 255, cv2.THRESH_BINARY_INV)
    _, mask_high = cv2.threshold(img_gray, high_cut, 255, cv2.THRESH_BINARY)

    height, width = img_gray.shape
    bounds = segment_bounds(width, NUM_SEGMENTS)

    segment_means = []
    for x1, x2 in bounds:
        seg_gray = img_gray[:, x1:x2]
        segment_means.append(float(np.mean(seg_gray)))

    median_bg = float(np.median(segment_means))
    std_bg = float(np.std(segment_means))
    dynamic_white_thresh = median_bg + max(WHITE_OUTLIER_SIGMA * std_bg, WHITE_OUTLIER_MIN_DIFF)

    segment_labels = []
    debug_rows = []
    for i, (x1, x2) in enumerate(bounds):
        seg_low = mask_low[:, x1:x2]
        seg_high = mask_high[:, x1:x2]
        seg_gray = img_gray[:, x1:x2]
        seg_mean = segment_means[i]
        # Edge segments: use thresholded mean for a stable decision
        thresh_mean = ""
        if i == 0 or i == NUM_SEGMENTS - 1:
            valid = (seg_gray >= low_cut) & (seg_gray <= high_cut)
            if np.any(valid):
                thresh_mean = float(np.mean(seg_gray[valid]))
            else:
                thresh_mean = seg_mean
            # Edge override: DIE/WHITE first, then OK band, else fallback
            if seg_mean < ABS_DIE_MEAN:
                label = "DIE"
                ratio_dark = cv2.countNonZero(seg_low) / max(seg_low.size, 1)
                ratio_white = cv2.countNonZero(seg_high) / max(seg_high.size, 1)
            elif seg_mean > ABS_WHITE_MEAN or seg_mean > dynamic_white_thresh:
                label = "WHITE"
                ratio_dark = cv2.countNonZero(seg_low) / max(seg_low.size, 1)
                ratio_white = cv2.countNonZero(seg_high) / max(seg_high.size, 1)
            # Treat around 69 as OK (tolerance +/- 5)
            elif 64.0 <= thresh_mean <= 115.0:
                label = "OK"
                ratio_dark = cv2.countNonZero(seg_low) / max(seg_low.size, 1)
                ratio_white = cv2.countNonZero(seg_high) / max(seg_high.size, 1)
            else:
                label, ratio_dark, ratio_white = classify_segment(
                    seg_low,
                    seg_high,
                    seg_gray,
                    DEFECT_RATIO_LIMIT,
                    mid_gray,
                    seg_mean,
                    dynamic_white_thresh,
                )
        else:
            label, ratio_dark, ratio_white = classify_segment(
                seg_low,
                seg_high,
                seg_gray,
                DEFECT_RATIO_LIMIT,
                mid_gray,
                seg_mean,
                dynamic_white_thresh,
            )
        segment_labels.append(label)
        debug_rows.append(
            {
                "segment": i + 1,
                "mean": float(seg_mean),
                "thresh_mean": "" if thresh_mean == "" else float(thresh_mean),
                "ratio_dark": float(ratio_dark),
                "ratio_white": float(ratio_white),
                "low_cut": int(low_cut),
                "high_cut": int(high_cut),
                "dyn_white": float(dynamic_white_thresh),
                "label": label,
            }
        )

    image_label = classify_image(segment_labels)
    return image_label, bounds, segment_labels, low_cut, high_cut, debug_rows


def iter_images(root: Path):
    for path in root.rglob("*"):
        if path.is_file() and is_image_file(path):
            yield path


def main():
    if not INPUT_FOLDER.exists():
        raise SystemExit(f"Input folder not found: {INPUT_FOLDER}")

    images = sorted(iter_images(INPUT_FOLDER))
    if not images:
        raise SystemExit(f"No images found in: {INPUT_FOLDER}")

    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    debug_all_rows = []

    print(f"Processing {len(images)} images...")

    for img_path in images:
        img_color = cv2.imread(str(img_path))
        if img_color is None:
            print(f"Skipping unreadable image: {img_path}")
            continue

        img_gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)

        print(f"Analyzing: {img_path.name}")
        image_label, bounds, segment_labels, low_cut, high_cut, debug_rows = process_image(img_gray)
        annotated = annotate_image(img_color, segment_labels, bounds, image_label)

        rel_path = img_path.relative_to(INPUT_FOLDER)
        label_folder = LABEL_TO_FOLDER.get(image_label, "Unknown")
        output_dir = OUTPUT_FOLDER / label_folder / rel_path.parent
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / img_path.name
        cv2.imwrite(str(output_path), annotated)

        print(
            f"   -> Result: {image_label} (Cutoffs: {low_cut}, {high_cut})\n"
        )
        for row in debug_rows:
            row["file"] = str(img_path.relative_to(INPUT_FOLDER))
            row["image_label"] = image_label
            debug_all_rows.append(row)

    if debug_all_rows:
        with open(DEBUG_CSV_PATH, "w", newline="") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=[
                    "file",
                    "image_label",
                    "segment",
                    "mean",
                    "thresh_mean",
                    "ratio_dark",
                    "ratio_white",
                    "low_cut",
                    "high_cut",
                    "dyn_white",
                    "label",
                ],
            )
            writer.writeheader()
            writer.writerows(debug_all_rows)

    print(f"Done. Output saved to: {OUTPUT_FOLDER}")
    print(f"Debug CSV: {DEBUG_CSV_PATH}")


if __name__ == "__main__":
    main()