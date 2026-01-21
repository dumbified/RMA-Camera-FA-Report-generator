import cv2
import numpy as np
import os
import csv
import shutil
import re

# --- SETTINGS ---
INPUT_FOLDER = "samples2"          
BASE_OUTPUT_FOLDER = "results2"    
CSV_NAME = "defect_report2.csv"
NUM_SEGMENTS = 10

# Define categories
CATEGORIES = ["Perfect", "White Segment", "Whole Segment", "Die Segment"]

# Create folder structure
if not os.path.exists(BASE_OUTPUT_FOLDER):
    os.makedirs(BASE_OUTPUT_FOLDER)

for cat in CATEGORIES:
    cat_path = os.path.join(BASE_OUTPUT_FOLDER, cat)
    if not os.path.exists(cat_path):
        os.makedirs(cat_path)

def extract_vit_id(filename: str) -> str:
    m = re.search(r'VIT(\d+)', filename, flags=re.IGNORECASE)
    return m.group(1) if m else ""

def process_batch():
    all_results = []
    files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith(('.jpg', '.png', '.jpeg'))]

    for filename in files:
        path = os.path.join(INPUT_FOLDER, filename)
        img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
        if img is None: continue
        vit_id = extract_vit_id(filename)

        height, width = img.shape
        seg_width = width // NUM_SEGMENTS
        vis = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
        
        # 1. Collect Segment Statistics
        segment_avgs = []
        for i in range(NUM_SEGMENTS):
            x1 = i * seg_width
            x2 = (i + 1) * seg_width
            segment = img[:, x1:x2]
            avg = np.mean(segment)
            segment_avgs.append(avg)

        # 2. Compute Dynamic Thresholds
        # Use median to be robust against outliers (like the one white segment)
        median_bg = np.median(segment_avgs)
        std_bg = np.std(segment_avgs)

        # Dynamic Threshold Calculation:
        # - We use Median + K * StdDev to find statistical outliers.
        # - We impose a 'min_diff' (e.g. 15) so that uniform gray images don't trigger.
        # - 2.5 sigma is a standard outlier bound.
        dynamic_white_thresh = median_bg + max(2.5 * std_bg, 15)

        dark_indices = []
        white_indices = []

        # 3. Classify and Visualize
        for i in range(NUM_SEGMENTS):
            avg = segment_avgs[i]
            x1 = i * seg_width
            x2 = (i + 1) * seg_width
            
            # Identification Logic
            if avg < 55:
                # Keep absolute threshold for "DIE" (very dark/missing)
                seg_type = "DIE"
                dark_indices.append(i)
                color = (0, 0, 255) # Red
            elif avg > 160 or avg > dynamic_white_thresh:
                # Trigger if absolute white OR relative outlier
                seg_type = "WHITE"
                white_indices.append(i)
                color = (0, 255, 255) # Yellow
            else:
                seg_type = "OK"
                color = (0, 255, 0) # Green

            # Draw Box & Text
            cv2.rectangle(vis, (x1 + 2, 2), (x2 - 2, height - 2), color, 1)
            cv2.putText(vis, f"{int(avg)}", (x1 + 5, 50), 
                    cv2.FONT_HERSHEY_SIMPLEX, 0.6, color, 2)
            cv2.putText(vis, f"{seg_type}", (x1 + 5, 80), 
                    cv2.FONT_HERSHEY_SIMPLEX, 0.6, color, 2)

        # 2. Final Categorization Logic
        total_defects = len(dark_indices) + len(white_indices)
        if total_defects == 0:
            category = "Perfect"
        elif len(dark_indices) == NUM_SEGMENTS:
            category = "Whole Segment"
        elif len(white_indices) > 0:
            category = "White Segment"
        else:
            category = "Die Segment"

        # 3. Save to Category Folder
        save_path = os.path.join(BASE_OUTPUT_FOLDER, category, filename)
        cv2.imwrite(save_path, vis)

        # 4. Data for CSV
        all_results.append({
            "VIT": vit_id,
            "Filename": filename,
            "Category": category,
            "Defect_Indices": ",".join(map(str, sorted(dark_indices + white_indices))),
            "Total_Defects": total_defects
        })

    # Write CSV
    with open(CSV_NAME, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=["VIT", "Filename", "Category", "Defect_Indices", "Total_Defects"])
        writer.writeheader()
        writer.writerows(all_results)

    print(f"Batch processed successfully! Check the subfolders in '{BASE_OUTPUT_FOLDER}'.")

if __name__ == "__main__":
    process_batch()