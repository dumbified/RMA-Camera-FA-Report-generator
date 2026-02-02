import cv2
import numpy as np
import os

# --- SETTINGS ---
INPUT_FOLDER = "samples"
BASE_OUTPUT_FOLDER = "results"
NUM_SEGMENTS = 10

# Defined Categories (Perfect changed to Pending)
CATEGORIES = ["Pending", "White Segment", "Whole Segment", "Die Segment"]

# Create output folder structure
if not os.path.exists(BASE_OUTPUT_FOLDER):
    os.makedirs(BASE_OUTPUT_FOLDER)

for cat in CATEGORIES:
    cat_path = os.path.join(BASE_OUTPUT_FOLDER, cat)
    if not os.path.exists(cat_path):
        os.makedirs(cat_path)

def process_images():
    files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith(('.jpg', '.png', '.jpeg'))]

    print(f"Processing {len(files)} images...")

    for filename in files:
        path = os.path.join(INPUT_FOLDER, filename)
        img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
        if img is None: continue

        height, width = img.shape
        seg_width = width // NUM_SEGMENTS
        
        # 1. Global Statistics (For Whole Segment Check)
        global_mean = np.mean(img)
        global_std = np.std(img)

        # 2. Collect Segment Statistics
        segment_avgs = []
        for i in range(NUM_SEGMENTS):
            x1 = i * seg_width
            x2 = (i + 1) * seg_width
            # Handle edge case for last segment rounding
            if i == NUM_SEGMENTS - 1: 
                x2 = width
                
            segment = img[:, x1:x2]
            avg = np.mean(segment)
            segment_avgs.append(avg)

        # 3. Compute Dynamic Thresholds for White Segments
        # Using Median + StdDev to find outliers relative to the specific image
        median_bg = np.median(segment_avgs)
        std_bg = np.std(segment_avgs)
        
        # Threshold: Median + 2.5 sigma (with a minimum floor of 15 to avoid noise in flat gray images)
        dynamic_white_thresh = median_bg + max(2.5 * std_bg, 15)

        # 4. Identify Defective Segments
        dark_indices = []
        white_indices = []

        for i, avg in enumerate(segment_avgs):
            # DIE SEGMENT LOGIC: Absolute dark threshold
            if avg < 55:
                dark_indices.append(i)
            
            # WHITE SEGMENT LOGIC: Absolute bright OR Relative Outlier
            elif avg > 160 or avg > dynamic_white_thresh:
                white_indices.append(i)

        # 5. Final Categorization Logic
        category = "Pending" # Default to Pending instead of Perfect

        # Priority 1: White Segments (Distinct outliers)
        if len(white_indices) > 0:
            category = "White Segment"

        # Priority 2: Whole Segment 
        # STRICTER CHECK: Global mean + std must be low AND all segments must be DIE.
        # This prevents "Dim OK" images from being classified as Whole Segment.
        elif len(dark_indices) == NUM_SEGMENTS and global_mean < 45 and global_std < 20:
            category = "Whole Segment"

        # Priority 3: Die Segment (Partial defects)
        elif len(dark_indices) > 0:
            category = "Die Segment"
        
        # Else: Remains "Pending" (Was Perfect)

        # 6. Save Raw Image to Category Folder (No Annotation)
        save_path = os.path.join(BASE_OUTPUT_FOLDER, category, filename)
        cv2.imwrite(save_path, img)

    print(f"Done. Images sorted into '{BASE_OUTPUT_FOLDER}'.")

if __name__ == "__main__":
    process_images()