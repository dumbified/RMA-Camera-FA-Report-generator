import base64
import io
import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import cv2
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

from main import (
    CATEGORIES,
    DEFAULT_INPUT_FOLDER,
    DEFAULT_OUTPUT_FOLDER,
    NUM_SEGMENTS,
    collect_review_items,
    save_image_to_category,
    extract_vit_id,
)

CATEGORIES_FILE = os.path.join(os.path.dirname(__file__), "categories.json")


def load_categories():
    default_categories = list(CATEGORIES)
    if not os.path.exists(CATEGORIES_FILE):
        return default_categories
    try:
        with open(CATEGORIES_FILE, "r") as handle:
            data = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return default_categories
    if not isinstance(data, list):
        return default_categories
    merged = []
    for name in data + default_categories:
        if isinstance(name, str):
            clean = name.strip()
            if clean and clean not in merged:
                merged.append(clean)
    return merged


def save_categories():
    try:
        with open(CATEGORIES_FILE, "w") as handle:
            json.dump(CATEGORIES, handle, indent=2)
    except OSError:
        pass


def get_hotkey_categories():
    preferred = ["Whole Segment", "Die Segment", "White Segment"]
    ordered = []
    for name in preferred:
        if name in CATEGORIES and name not in ordered:
            ordered.append(name)
    for name in CATEGORIES:
        if name not in ordered and name != "Pending":
            ordered.append(name)
    if "Pending" in CATEGORIES:
        ordered.append("Pending")
    return ordered





def to_photo_image(gray_img, max_width, max_height):
    if gray_img is None:
        return None
    height, width = gray_img.shape
    scale = min(max_width / width, max_height / height, 1.0)
    if scale < 1.0:
        new_size = (int(width * scale), int(height * scale))
        gray_img = cv2.resize(gray_img, new_size, interpolation=cv2.INTER_AREA)
    rgb = cv2.cvtColor(gray_img, cv2.COLOR_GRAY2RGB)
    ok, buf = cv2.imencode(".png", rgb)
    if not ok:
        return None
    b64 = base64.b64encode(buf).decode("ascii")
    return tk.PhotoImage(data=b64)


def profile_plot_photo(gray_img, max_width, max_height):
    if gray_img is None:
        return None
    profile = np.mean(gray_img, axis=0)
    width_px = max_width
    height_px = max_height
    fig = plt.figure(figsize=(width_px / 100, height_px / 100), dpi=100)
    ax = fig.add_subplot(111)
    ax.plot(profile, color="#1f77b4", linewidth=1)
    ax.set_title("Avg Intensity Profile")
    ax.set_xlabel("Width (Pixels)")
    ax.set_ylabel("Avg Intensity (0-255)")
    ax.set_ylim(0, 255)
    ax.set_yticks([0, 85, 170, 255])
    ax.grid(True, linestyle=":", alpha=0.6)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=100)
    plt.close(fig)
    b64 = base64.b64encode(buf.getvalue()).decode("ascii")
    return tk.PhotoImage(data=b64)


def build_gui():
    loaded = load_categories()
    CATEGORIES.clear()
    CATEGORIES.extend(loaded)

    root = tk.Tk()
    root.title("Image classification")
    root.geometry("1080x960")

    input_var = tk.StringVar(value=DEFAULT_INPUT_FOLDER)
    output_var = tk.StringVar(value=DEFAULT_OUTPUT_FOLDER)
    csv_name_var = tk.StringVar(value="review_results.csv")

    def browse_input():
        path = filedialog.askdirectory(title="Select Input Folder")
        if path:
            input_var.set(path)

    def browse_output():
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            output_var.set(path)

    all_review_items = []
    review_items = []
    review_index = 0
    current_photo = None
    current_plot = None
    decisions = {}

    def refresh_category_widgets():
        review_filter_box.config(values=CATEGORIES)
        update_hotkey_label()

    def apply_review_filter():
        nonlocal review_index
        selected = review_filter_var.get()
        review_items.clear()
        review_items.extend(
            [item for item in all_review_items if item["result"]["category"] == selected]
        )
        review_index = 0
        show_review_item()

    def show_review_item():
        nonlocal review_index, current_photo, current_plot
        if not review_items:
            review_status_var.set("No items to review.")
            preview_label.config(image="", text="No image loaded")
            plot_label.config(image="")
            review_text.config(state="normal")
            review_text.delete("1.0", tk.END)
            review_text.config(state="disabled")
            return

        if review_index < 0:
            review_index = 0
        if review_index >= len(review_items):
            review_index = len(review_items) - 1

        item = review_items[review_index]
        image_path = item["path"]
        result = item["result"]

        img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
        current_photo = to_photo_image(img, 520, 320)
        if current_photo is not None:
            preview_label.config(image=current_photo, text="")
            preview_label.image = current_photo
        else:
            preview_label.config(image="", text="No image loaded")

        current_plot = profile_plot_photo(img, 700, 320)
        if current_plot is not None:
            plot_label.config(image=current_plot)
            plot_label.image = current_plot
        else:
            plot_label.config(image="")

        review_status_var.set(f"Reviewing {review_index + 1}/{len(review_items)}")
        category_var.set(result["category"])

        info_lines = [
            f"Image: {os.path.basename(image_path)}",
        ]
        if "confidence" in result:
            info_lines.append(f"Confidence: {result['confidence']:.2%}")
        if "method" in result:
            info_lines.append(f"Method: {result['method']}")

        review_text.config(state="normal")
        review_text.delete("1.0", tk.END)
        review_text.insert(tk.END, "\n".join(info_lines))
        review_text.config(state="disabled")

    def run_batch():
        input_folder = input_var.get().strip()
        output_folder = output_var.get().strip()
        if not input_folder or not os.path.isdir(input_folder):
            messagebox.showerror("Invalid Input", "Input folder is not valid.")
            return
        if not output_folder:
            messagebox.showerror("Invalid Input", "Output folder is required.")
            return
        try:
            all_review_items.clear()
            all_review_items.extend(collect_review_items(input_folder, NUM_SEGMENTS))
            if not all_review_items:
                messagebox.showinfo("Done", "No images found to review.")
            else:
                messagebox.showinfo("Done", "Batch processing completed. Review results below.")
            decisions.clear()
            for item in all_review_items:
                decisions[item["path"]] = item["result"]["category"]
            review_filter_var.set(CATEGORIES[0] if CATEGORIES else "")
            apply_review_filter()
        except Exception as exc:
            messagebox.showerror("Error", f"Batch processing failed: {exc}")

    def update_current_and_next():
        if not review_items:
            return

        item = review_items[review_index]
        chosen_category = category_var.get()
        if chosen_category not in CATEGORIES:
            messagebox.showerror("Invalid Input", "Please select a valid category.")
            return

        decisions[item["path"]] = chosen_category
        next_item()

    def skip_current():
        next_item()

    def next_item():
        nonlocal review_index
        if not review_items:
            return
        review_index += 1
        if review_index >= len(review_items):
            finalize_review()
            return
        show_review_item()

    def prev_item():
        nonlocal review_index
        if not review_items:
            return
        review_index -= 1
        if review_index < 0:
            review_index = 0
        show_review_item()

    def finalize_review():
        output_folder = output_var.get().strip()
        if not output_folder:
            messagebox.showerror("Invalid Input", "Output folder is required.")
            return False
        if not os.path.isdir(output_folder):
            messagebox.showerror("Invalid Input", "Output folder is not valid.")
            return False

        if not decisions:
            for item in all_review_items:
                decisions[item["path"]] = item["result"]["category"]

        if not decisions:
            review_status_var.set("Review complete. No updates to save.")
        else:
            csv_name = csv_name_var.get().strip()
            if not csv_name:
                messagebox.showerror("Invalid Input", "CSV filename is required.")
                return False
            if not csv_name.lower().endswith(".csv"):
                csv_name += ".csv"
            csv_path = os.path.join(output_folder, csv_name)
            try:
                with open(csv_path, "w", newline="") as csv_file:
                    csv_file.write("VIT,filename,category\n")
                    for image_path, category in decisions.items():
                        vit_id = extract_vit_id(os.path.basename(image_path))
                        csv_file.write(f"{vit_id},{os.path.basename(image_path)},{category}\n")
            except OSError as exc:
                messagebox.showerror("Error", f"Failed to write CSV: {exc}")
                return False

            for image_path, category in decisions.items():
                save_image_to_category(image_path, output_folder, category)

            review_status_var.set(f"Review complete. Saved CSV to {csv_path}.")

        preview_label.config(image="", text="No image loaded")
        plot_label.config(image="")
        review_text.config(state="normal")
        review_text.delete("1.0", tk.END)
        review_text.config(state="disabled")
        return True

    def save_and_quit():
        if finalize_review():
            messagebox.showinfo("Saved", "Result successfully saved.")
            root.destroy()

    def update_hotkey_label():
        ordered = get_hotkey_categories()
        lines = ["Hotkeys"]
        for idx, name in enumerate(ordered, start=1):
            if idx > 9:
                break
            lines.append(f"{idx}  {name}")
        if len(lines) == 1:
            lines.append("None")
        hotkey_var.set("\n".join(lines))

    def handle_keypress(event):
        widget_class = event.widget.winfo_class()
        if widget_class in ("Entry", "TEntry", "Combobox", "TCombobox"):
            return
        if not review_items:
            return
        if event.char.isdigit() and event.char != "0":
            index = int(event.char) - 1
            ordered = get_hotkey_categories()
            if index < len(ordered):
                category_var.set(ordered[index])
                update_current_and_next()

    main_frame = ttk.Frame(root, padding=12)
    main_frame.pack(fill="both", expand=True)

    ttk.Label(main_frame, text="Batch Processing", font=("Helvetica", 12, "bold")).pack(anchor="w")

    input_row = ttk.Frame(main_frame)
    input_row.pack(fill="x", pady=6)
    ttk.Label(input_row, text="Input folder:", width=14).pack(side="left")
    ttk.Entry(input_row, textvariable=input_var).pack(side="left", fill="x", expand=True, padx=6)
    ttk.Button(input_row, text="Browse", command=browse_input).pack(side="left")

    output_row = ttk.Frame(main_frame)
    output_row.pack(fill="x", pady=6)
    ttk.Label(output_row, text="Output folder:", width=14).pack(side="left")
    ttk.Entry(output_row, textvariable=output_var).pack(side="left", fill="x", expand=True, padx=6)
    ttk.Button(output_row, text="Browse", command=browse_output).pack(side="left")

    run_row = ttk.Frame(main_frame)
    run_row.pack(fill="x", pady=6)
    ttk.Button(run_row, text="Run Batch", command=run_batch).pack(side="left")
    ttk.Label(run_row, text="CSV name:").pack(side="left", padx=(16, 6))
    ttk.Entry(run_row, textvariable=csv_name_var, width=24).pack(side="left")

    ttk.Separator(main_frame).pack(fill="x", pady=12)
    ttk.Label(main_frame, text="Review Batch Results", font=("Helvetica", 12, "bold")).pack(anchor="w")

    review_status_var = tk.StringVar(value="No batch results yet.")
    ttk.Label(main_frame, textvariable=review_status_var).pack(anchor="w", pady=2)

    preview_frame = ttk.Frame(main_frame)
    preview_frame.pack(fill="x", pady=6)
    preview_container = tk.Frame(preview_frame, width=520, height=320)
    preview_container.pack_propagate(False)
    preview_container.pack(side="left")
    preview_label = tk.Label(preview_container, text="No image loaded")
    preview_label.pack(fill="both", expand=True)

    review_side = ttk.Frame(preview_frame)
    review_side.pack(side="left", fill="both", expand=True, padx=12)

    category_var = tk.StringVar(value="Pending")
    review_filter_var = tk.StringVar(value=CATEGORIES[0] if CATEGORIES else "")

    ttk.Label(review_side, text="Review category:").pack(anchor="w")
    review_filter_box = ttk.Combobox(
        review_side,
        textvariable=review_filter_var,
        values=CATEGORIES,
        state="readonly",
        width=18,
    )
    review_filter_box.pack(anchor="w", pady=(0, 8))
    review_filter_box.bind("<<ComboboxSelected>>", lambda _e: apply_review_filter())

    add_row = ttk.Frame(review_side)
    add_row.pack(anchor="w", pady=(8, 2))
    ttk.Label(add_row, text="Add category:").pack(side="left")
    new_category_var = tk.StringVar(value="")
    new_category_entry = ttk.Entry(add_row, textvariable=new_category_var, width=16)
    new_category_entry.pack(side="left", padx=6)
    ttk.Button(add_row, text="Add", command=lambda: add_category(new_category_var.get().strip())).pack(side="left")

    hotkey_var = tk.StringVar(value="")
    ttk.Label(review_side, textvariable=hotkey_var, wraplength=220).pack(anchor="w", pady=(6, 2))

    def add_category(name):
        if not name:
            messagebox.showerror("Invalid Input", "Category name cannot be empty.")
            return
        if name in CATEGORIES:
            messagebox.showerror("Invalid Input", "Category already exists.")
            return
        CATEGORIES.append(name)
        save_categories()
        refresh_category_widgets()
        category_var.set(name)
        new_category_var.set("")

    def add_category_from_entry(_event=None):
        add_category(new_category_var.get().strip())

    new_category_entry.bind("<Return>", add_category_from_entry)

    review_buttons = ttk.Frame(review_side)
    review_buttons.pack(anchor="w", pady=8)
    ttk.Button(review_buttons, text="Previous", command=prev_item).pack(side="left")
    ttk.Button(review_buttons, text="Update & Next", command=update_current_and_next).pack(side="left", padx=6)
    ttk.Button(review_buttons, text="Skip", command=skip_current).pack(side="left")
    ttk.Button(review_buttons, text="Save & Quit", command=save_and_quit).pack(side="left", padx=6)

    review_text = tk.Text(main_frame, height=6, state="disabled")
    review_text.pack(fill="x", expand=False, pady=6)

    plot_label = ttk.Label(main_frame)
    plot_label.pack(fill="x", pady=4)

    update_hotkey_label()
    root.bind("<Key>", handle_keypress)
    root.mainloop()


if __name__ == "__main__":
    build_gui()
