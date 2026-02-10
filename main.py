import os
import re
import cv2
import numpy as np
import torch
import timm
from PIL import Image, ImageOps
from torchvision import transforms

# --- CONFIGURATION (from main) ---
DEFAULT_INPUT_FOLDER = "samples"
DEFAULT_OUTPUT_FOLDER = "results"
NUM_SEGMENTS = 10 
CONFIDENCE_THRESHOLD = 0.75  # If model confidence below this, fallback to rules

# --- ADVANCED LOGIC PARAMS (from adv_classification.py) ---
MIN_VALID_PEAK = 50       # Threshold separating "Black Background" from "Grey Object"
PEAK_SENSITIVITY = 0.05   # The Grey Peak must be at least 5% height of the Black Peak
DEFECT_RATIO_LIMIT = 0.05  # 5% of pixels in a segment
ABS_DIE_MEAN = 35
ABS_WHITE_MEAN = 160
WHITE_OUTLIER_SIGMA = 2.5
WHITE_OUTLIER_MIN_DIFF = 15
WHOLE_SEGMENT_MEAN_MAX = 30.0

# 1. SquarePad
class SquarePad:
    def __call__(self, img):
        w, h = img.size
        max_side = max(w, h)
        pad_w, pad_h = (max_side - w) // 2, (max_side - h) // 2
        return ImageOps.expand(img, (pad_w, pad_h, pad_w, pad_h), fill=0)

# Setup Inference
DEVICE = torch.device("cuda" if torch.cuda.is_available() else "cpu")
# Try to find the model file
MODEL_PATH = 'mobilevitv2_noweights2.pth'
if not os.path.exists(MODEL_PATH):
    MODEL_PATH = 'defect_model_mobilevit_v2.pth'

# Load Model
model = None
preprocess = None
CATEGORIES = ["Die Segment", "Line Between", "White Segment", "Whole Segment"] # Default

try:
    if os.path.exists(MODEL_PATH):
        print(f"Loading model from {MODEL_PATH}...")
        checkpoint = torch.load(MODEL_PATH, map_location=DEVICE)
        
        # Check if 'class_names' key exists, otherwise try to infer or use default
        if 'class_names' in checkpoint:
            CATEGORIES = checkpoint['class_names']
        
        # Reconstruct Model Architecture
        model = timm.create_model('mobilevitv2_100', pretrained=False, num_classes=len(CATEGORIES))
        
        # Handle state dict loading
        if 'model_state_dict' in checkpoint:
            model.load_state_dict(checkpoint['model_state_dict'])
        else:
            # Maybe the checkpoint is the state dict itself?
            model.load_state_dict(checkpoint)
            
        model.to(DEVICE)
        model.eval()

        preprocess = transforms.Compose([
            SquarePad(),
            transforms.Resize((256, 256)),
            transforms.ToTensor(),
            transforms.Normalize(mean=[0.485, 0.456, 0.406], std=[0.229, 0.224, 0.225])
        ])
        print(f"Model loaded successfully. Categories: {CATEGORIES}")
    else:
        print(f"Model file not found: {MODEL_PATH}")

except Exception as e:
    print(f"Error loading model: {e}")

# --- NEW RULE BASED LOGIC (from adv_classification.py) ---

def get_safe_zone_thresholds(img):
    """
    Calculates thresholds using Priority Peak Logic from adv_classification.py
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
    global_max_idx = np.argmax(hist_smooth)
    global_max_height = hist_smooth[global_max_idx]

    signal_region = hist_smooth[MIN_VALID_PEAK:]
    
    if signal_region.size > 0:
        rel_signal_idx = np.argmax(signal_region)
        signal_idx = rel_signal_idx + MIN_VALID_PEAK
        signal_height = hist_smooth[signal_idx]
    else:
        signal_height = 0
        signal_idx = 0

    # 3. Peak Selection Logic
    if signal_height > (PEAK_SENSITIVITY * global_max_height):
        main_peak_idx = signal_idx
        peak_height = signal_height
    else:
        main_peak_idx = global_max_idx
        peak_height = global_max_height

    # 4. Critical Check
    if main_peak_idx < MIN_VALID_PEAK:
        return 255, 255

    # 5. Calculate Left Threshold
    left_thresh = 0
    for i in range(main_peak_idx, 1, -1):
        if hist_smooth[i] < (0.1 * peak_height):
            left_thresh = i
            break

    # 6. Calculate Right Threshold
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

    # High-priority mask-based rule:
    # - If more than 50% of pixels are dark  -> DIE
    # - Else if more than 50% are white     -> WHITE
    if ratio_dark > 0.5:
        return "DIE", ratio_dark, ratio_white
    if ratio_white > 0.5:
        return "WHITE", ratio_dark, ratio_white

    # Mean-based hard rules (fallback when neither dark nor white > 50%)
    if seg_mean < ABS_DIE_MEAN:
        return "DIE", ratio_dark, ratio_white
    if seg_mean > ABS_WHITE_MEAN or seg_mean > dynamic_white_thresh:
        return "WHITE", ratio_dark, ratio_white

    # Otherwise, treat segment as OK
    return "OK", ratio_dark, ratio_white

def determine_final_label(segment_labels, segment_means, overall_mean):
    """
    Consolidates per-segment labels into a single image label.
    """
    # Additional whole-segment rule: if overall image mean is very low,
    # treat it as Whole Segment regardless of per-segment labels.
    if overall_mean <= 35.0:
        return "Whole Segment"

    # Priority check
    if "DIE" in segment_labels:
        return "Die Segment"
    if "WHITE" in segment_labels:
        return "White Segment"
    
    # If mostly OK (PERFECT), return Pending
    return "Pending"

def run_rule_based_check(img_gray, num_segments=10):
    """
    Replaces old simple logic with the advanced histogram/threshold logic from adv_classification.py
    """
    low_cut, high_cut = get_safe_zone_thresholds(img_gray)
    mid_gray = (low_cut + high_cut) / 2.0

    _, mask_low = cv2.threshold(img_gray, low_cut, 255, cv2.THRESH_BINARY_INV)
    _, mask_high = cv2.threshold(img_gray, high_cut, 255, cv2.THRESH_BINARY)
    
    height, width = img_gray.shape
    bounds = segment_bounds(width, num_segments)

    # 1. Calculate means first (needed for dynamic threshold)
    segment_means = []
    for x1, x2 in bounds:
        seg_gray = img_gray[:, x1:x2]
        segment_means.append(float(np.mean(seg_gray)))

    median_bg = float(np.median(segment_means))
    std_bg = float(np.std(segment_means))
    dynamic_white_thresh = median_bg + max(WHITE_OUTLIER_SIGMA * std_bg, WHITE_OUTLIER_MIN_DIFF)

    segment_labels = []

    for i, (x1, x2) in enumerate(bounds):
        seg_low = mask_low[:, x1:x2]
        seg_high = mask_high[:, x1:x2]
        seg_gray = img_gray[:, x1:x2]
        seg_mean = segment_means[i]

        # Edge handling logic from adv_classification
        thresh_mean = ""
        label = "OK" 
        ratio_dark = 0.0
        ratio_white = 0.0

        if i == 0 or i == num_segments - 1:
            valid = (seg_gray >= low_cut) & (seg_gray <= high_cut)
            if np.any(valid):
                thresh_mean = float(np.mean(seg_gray[valid]))
            else:
                thresh_mean = seg_mean
            
            # Edge override logic
            if seg_mean < ABS_DIE_MEAN:
                label = "DIE"
            elif seg_mean > ABS_WHITE_MEAN or seg_mean > dynamic_white_thresh:
                label = "WHITE"
            elif 64.0 <= thresh_mean <= 115.0:
                label = "OK"
            else:
                label, ratio_dark, ratio_white = classify_segment(
                    seg_low, seg_high, seg_gray, DEFECT_RATIO_LIMIT,
                    mid_gray, seg_mean, dynamic_white_thresh
                )
        else:
             label, ratio_dark, ratio_white = classify_segment(
                seg_low, seg_high, seg_gray, DEFECT_RATIO_LIMIT,
                mid_gray, seg_mean, dynamic_white_thresh
            )
        
        segment_labels.append(label)

    overall_mean = float(np.mean(img_gray))
    final_label = determine_final_label(segment_labels, segment_means, overall_mean)
    return final_label

def ensure_output_folders(base_output_folder):
    if not os.path.exists(base_output_folder):
        os.makedirs(base_output_folder)
    for cat in CATEGORIES:
        cat_path = os.path.join(base_output_folder, cat)
        if not os.path.exists(cat_path):
            os.makedirs(cat_path)

def classify_image(img, num_segments=NUM_SEGMENTS):
    """
    Classify an image using the loaded model.
    Fallback to rule-based check if confidence is low.
    """
    if model is None:
        return {"category": "Pending", "confidence": 0.0, "method": "No Model"}
    
    try:
        # Prapare for Model
        if len(img.shape) == 2:
            img_pil = Image.fromarray(img).convert('RGB')
            img_gray = img
        else:
            img_pil = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
            img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # Preprocess
        input_tensor = preprocess(img_pil).unsqueeze(0).to(DEVICE)
        
        # Inference
        with torch.no_grad():
            outputs = model(input_tensor)
            probabilities = torch.nn.functional.softmax(outputs, dim=1)
            confidence, preds = torch.max(probabilities, 1) # Use probabilities max
            predicted_idx = preds.item()
            model_category = CATEGORIES[predicted_idx]
            conf_score = confidence.item()
            
        # Decision Logic
        final_category = model_category
        method = "Model"

        if conf_score <= CONFIDENCE_THRESHOLD:
            print(f"Low confidence ({conf_score:.2f}) for detection. Running rule-based check...")
            
            # 1. Set to Pending because conf is low
            final_category = "Pending" 
            
            # 2. Check rules (OLD logic replaced by NEW logic from adv_classification.py)
            rule_category = run_rule_based_check(img_gray, num_segments)
            
            # 3. If rules found a specific defect, use it.
            if rule_category != "Pending":
                final_category = rule_category
                
            method = f"Rule-Based (Model Conf: {conf_score:.2f})"

        return {
            "category": final_category,
            "confidence": conf_score,
            "method": method
        }

    except Exception as e:
        print(f"Inference error for image: {e}")
        return {"category": "Error", "confidence": 0.0, "method": "Error"}

def collect_review_items(input_folder, num_segments=NUM_SEGMENTS):
    if not os.path.exists(input_folder):
        print(f"Input folder not found: {input_folder}")
        return []

    files = [
        f for f in os.listdir(input_folder)
        if f.lower().endswith((".jpg", ".png", ".jpeg", ".bmp", ".tif", ".tiff"))
    ]

    print(f"Processing {len(files)} images from {input_folder}...")
    items = []

    for filename in files:
        path = os.path.join(input_folder, filename)
        img = cv2.imread(path, cv2.IMREAD_GRAYSCALE) 
        if img is None:
            print(f"Failed to load {filename}")
            continue

        result = classify_image(img, num_segments)
        # result now includes 'method' key which might be useful for debug
        items.append({"path": path, "filename": filename, "result": result})

    print("Done processing.")
    return items

def save_image_to_category(image_path, base_output_folder, category):
    img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    if img is None:
        return False
    
    cat_path = os.path.join(base_output_folder, category)
    if not os.path.exists(cat_path):
        os.makedirs(cat_path)
        
    save_path = os.path.join(cat_path, os.path.basename(image_path))
    return cv2.imwrite(save_path, img)

def extract_vit_id(filename: str) -> str:
    name = os.path.splitext(os.path.basename(filename))[0]
    if not name:
        return ""
    match = re.search(r"VIT[^0-9]*([0-9]+)", name)
    if match:
        return f"VIT-{match.group(1)}"
    first_part = name.split("_", 1)[0]
    digits = "".join(ch for ch in first_part if ch.isdigit())
    return f"VIT-{digits}" if digits else ""
