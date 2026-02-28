import csv
import json
import os
import tkinter as tk
from tkinter import messagebox, ttk


def _bool_to_str(value: bool) -> str:
    return "Yes" if value else "No"


def _ensure_directory(path: str) -> None:
    directory = os.path.dirname(path)
    if directory and not os.path.isdir(directory):
        os.makedirs(directory, exist_ok=True)


def _normalize_newlines(s: str) -> str:
    """Turn literal \\n in strings (e.g. from JSON) into real newlines for display."""
    return (s or "").replace("\\n", "\n")


def _checkbox_row(parent, text: str, variable: tk.BooleanVar, row: int) -> None:
    """Label on left, checkbox (small box) on right."""
    f = ttk.Frame(parent)
    f.grid(row=row, column=0, columnspan=2, sticky="w", pady=2)
    ttk.Label(f, text=text).pack(side="left")
    ttk.Checkbutton(f, variable=variable).pack(side="right", padx=(8, 0))


def _dropdown(parent, variable: tk.StringVar, values: list, row: int, label: str, col: int = 0, width: int = 20) -> ttk.Combobox:
    """Single-row = dropdown. Returns combobox."""
    ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", padx=(0, 8), pady=2)
    w = ttk.Combobox(parent, textvariable=variable, values=values, state="readonly", width=width)
    w.grid(row=row, column=col + 1, sticky="ew", pady=2)
    return w


def _textbox(parent, label: str, height: int, row: int, col: int = 0) -> tk.Text:
    """Multi-row = textbox. Returns Text widget."""
    ttk.Label(parent, text=label).grid(row=row, column=col, sticky="nw", padx=(0, 8), pady=2)
    w = tk.Text(parent, height=height, wrap="word")
    w.grid(row=row, column=col + 1, sticky="ew", pady=2)
    return w


def build_failure_form_gui() -> None:
    root = tk.Tk()
    root.title("Camera Failure Analysis Form")
    root.geometry("900x820")

    main = ttk.Frame(root, padding=12)
    main.pack(fill="both", expand=True)

    canvas = tk.Canvas(main, highlightthickness=0)
    scrollbar = ttk.Scrollbar(main, orient="vertical", command=canvas.yview)
    form_frame = ttk.Frame(canvas)

    form_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=form_frame, anchor="nw")

    def _on_mousewheel(event):
        # macOS: event.delta 1/-1; Windows: event.delta 120/-120; Linux: event.num 4/5
        if getattr(event, "num", None) == 5:
            canvas.yview_scroll(1, "units")
        elif getattr(event, "num", None) == 4:
            canvas.yview_scroll(-1, "units")
        elif hasattr(event, "delta"):
            d = event.delta
            if abs(d) >= 120:
                d = d // 120
            canvas.yview_scroll(-d, "units")

    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    for i in range(2):
        form_frame.columnconfigure(i, weight=1)

    # --- Variables ---
    npf_all_reworked_var = tk.BooleanVar(value=False)
    customer_request_var = tk.BooleanVar(value=False)
    camera_model_var = tk.StringVar(value="3.3G")
    burnt_var = tk.BooleanVar(value=False)
    power_on_unit_var = tk.StringVar(value="")
    remark_missing_burnt_var = tk.StringVar(value="")  # stored from textbox
    pcba_var = tk.StringVar(value="")
    camh_var = tk.StringVar(value="")
    scintillator_var = tk.StringVar(value="")
    ecc_rework_var = tk.StringVar(value="")
    good_camh_var = tk.StringVar(value="")
    bad_camh_var = tk.StringVar(value="")
    which_dvm_fail_var = tk.StringVar(value="")  # from textbox
    # Can-repair states now come from dropdowns:
    # "able" / "unable" (and "Scrap" for PCBA/FT).
    can_repair_bad_camh_var = tk.StringVar(value="")
    component_cause_var = tk.StringVar(value="")  # from textbox
    pcba_ats_result_var = tk.StringVar(value="")
    ats_result_if_failed_var = tk.StringVar(value="")  # from textbox
    can_repair_ats_var = tk.StringVar(value="")
    component_cause_ats_var = tk.StringVar(value="")  # from textbox
    component_category_var = tk.StringVar(value="")
    ft_result_var = tk.StringVar(value="")
    ft_fail_component_change_var = tk.StringVar(value="")  # from textbox
    can_repair_ft_var = tk.StringVar(value="")
    ft_pass_which_camh_var = tk.StringVar(value="")
    npf_final_var = tk.BooleanVar(value=False)
    camh_final_var = tk.StringVar(value="")
    pcba_final_var = tk.StringVar(value="")
    bad_camh_assemble_bubble_var = tk.BooleanVar(value=False)

    row = 0

    # --- If NPF and all reworked? [checkbox right] ---
    _checkbox_row(form_frame, "If NPF and all reworked?", npf_all_reworked_var, row)
    row += 1

    # --- Customer Request? [checkbox right] ---
    _checkbox_row(form_frame, "Customer Request?", customer_request_var, row)
    row += 1

    ttk.Separator(form_frame, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=8)
    row += 1

    # --- Camera Model [dropdown] ---
    _dropdown(form_frame, camera_model_var, ["3G", "3.1G Old", "3.1G New", "3.3G", "3.4G"], row, "Camera Model", width=10)
    row += 1

    ttk.Separator(form_frame, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=8)
    row += 1

    # --- Visual Check ---
    ttk.Label(form_frame, text="Visual Check", font=("", 10, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", pady=(4, 2))
    row += 1

    # Burnt? [checkbox right]  and  Power on unit [dropdown]  (same row: left side and right side)
    left = ttk.Frame(form_frame)
    left.grid(row=row, column=0, sticky="w", pady=2)
    ttk.Label(left, text="Burnt?").pack(side="left")
    ttk.Checkbutton(left, variable=burnt_var).pack(side="right", padx=(8, 0))
    right = ttk.Frame(form_frame)
    right.grid(row=row, column=1, sticky="w", pady=2)
    ttk.Label(right, text="Power on unit").pack(side="left", padx=(0, 8))
    cb_power = ttk.Combobox(right, textvariable=power_on_unit_var, values=["Can Ping", "Cannot Ping"], state="readonly", width=18)
    cb_power.pack(side="left", fill="x", expand=True)
    row += 1

    # Remark if missing or burnt [multiline textbox]
    remark_missing_burnt_text = _textbox(form_frame, "Remark if missing or burnt", 3, row)
    row += 1

    ttk.Separator(form_frame, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=8)
    row += 1

    # --- Screening result ---
    ttk.Label(form_frame, text="Screening result", font=("", 10, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", pady=(4, 2))
    row += 1

    # PCBA, CAMH, Scintillator [three dropdowns one row]
    f3 = ttk.Frame(form_frame)
    f3.grid(row=row, column=0, columnspan=2, sticky="ew", pady=2)
    f3.columnconfigure(1, weight=1)
    f3.columnconfigure(3, weight=1)
    f3.columnconfigure(5, weight=1)
    ttk.Label(f3, text="PCBA").grid(row=0, column=0, sticky="w", padx=(0, 4))
    ttk.Combobox(f3, textvariable=pcba_var, values=["Perfect", "Line between segment", "Cannot Ping"], state="readonly", width=18).grid(row=0, column=1, sticky="ew", padx=2)
    ttk.Label(f3, text="CAMH").grid(row=0, column=2, sticky="w", padx=(8, 4))
    ttk.Combobox(f3, textvariable=camh_var, values=["Whole Segment", "Perfect", "<= 2 segment", "Line Between Segment", "White Segment", "Multiple Segment Die"], state="readonly", width=18).grid(row=0, column=3, sticky="ew", padx=2)
    ttk.Label(f3, text="Scintillator").grid(row=0, column=4, sticky="w", padx=(8, 4))
    ttk.Combobox(f3, textvariable=scintillator_var, values=["Good", "Aging", "Gap", "Bubble", "Bent"], state="readonly", width=12).grid(row=0, column=5, sticky="ew", padx=2)
    row += 1

    ttk.Separator(form_frame, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=8)
    row += 1

    # --- ECC Rework [dropdown] ---
    _dropdown(form_frame, ecc_rework_var, ["NA", "3G rework", "3.1G old Rework", "3.1G new Rework"], row, "ECC Rework", width=18)
    row += 1

    # Good CAMH? [dropdown] (IF good camh)
    _dropdown(form_frame, good_camh_var, ["NA", "All Pass", "Got Bubble", "HCTE Fail", "Got bubble and HCTE Fail", "All PASS + Rework"], row, "Good CAMH?", width=22)
    row += 1

    # Bad CAMH? [dropdown]
    _dropdown(form_frame, bad_camh_var, ["NA", "Whole Segment", "Line between segment", "Die segment", "Multiple segment DIE", "White segment"], row, "Bad CAMH?", width=20)
    row += 1

    # Which DVM show it fail [multiline textbox]
    which_dvm_fail_text = _textbox(form_frame, "Which DVM show it fail", 3, row)
    row += 1

    # Can repair? [dropdown]
    _dropdown(form_frame, can_repair_bad_camh_var, ["", "able", "unable"], row, "Can repair?", width=10)
    row += 1

    # Component Cause: [multiline textbox]
    component_cause_text = _textbox(form_frame, "Component Cause:", 2, row)
    row += 1

    ttk.Separator(form_frame, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=8)
    row += 1

    # --- PCBA ATS result [dropdown] ---
    _dropdown(form_frame, pcba_ats_result_var, ["Pass", "Fail"], row, "PCBA ATS result")
    row += 1

    # If failed, state the ATS result [multiline textbox]
    ats_result_if_failed_text = _textbox(form_frame, "If failed, state the ATS result", 2, row)
    row += 1

    # Can repair? [dropdown]
    _dropdown(form_frame, can_repair_ats_var, ["", "able", "unable", "Scrap"], row, "Can repair?", width=10)
    row += 1

    # Component cause: [multiline textbox]
    component_cause_ats_text = _textbox(form_frame, "Component cause:", 2, row)
    row += 1

    # Component Category: [dropdown - single row] (PCBA component categories)
    _dropdown(form_frame, component_category_var, ["", "IC", "Capacitor", "Inductor"], row, "Component Category:")
    row += 1

    ttk.Separator(form_frame, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=8)
    row += 1

    # --- FT ---
    ttk.Label(form_frame, text="FT", font=("", 10, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", pady=(4, 2))
    row += 1

    # Result: [dropdown] (FT results)
    _dropdown(form_frame, ft_result_var, ["Pass with new camh", "Pass with ori camh", "Pass with swap camh", "Pass with repaired camh", "Fail"], row, "Result:", width=22)
    row += 1

    # IF fail, component change? [multiline textbox]
    ft_fail_component_change_text = _textbox(form_frame, "IF fail, component change?", 2, row)
    row += 1

    # Can repair? [dropdown]
    _dropdown(form_frame, can_repair_ft_var, ["", "able", "unable", "Scrap"], row, "Can repair?", width=10)
    row += 1

    # IF pass, which CAMH to use? [dropdown]
    _dropdown(form_frame, ft_pass_which_camh_var, ["", "Ori camh", "New camh", "Swap camh", "Repaired camh"], row, "IF pass, which CAMH to use?")
    row += 1

    ttk.Separator(form_frame, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=8)
    row += 1

    # --- FA ---
    ttk.Label(form_frame, text="FA", font=("", 10, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", pady=(4, 2))
    row += 1

    # NPF? [checkbox right]
    _checkbox_row(form_frame, "NPF?", npf_final_var, row)
    row += 1

    # CAMH [dropdown] (FA Camh)
    _dropdown(
        form_frame,
        camh_final_var,
        [
            "",
            "Whole Segment (5V Cap)",
            "<=2 segment and 89504-0004",
            "<=2 segment and 89504-0004 to 89504-00085",
            "HCTE Fail",
            "Old CCD fail",
            "Repairing Camh",
            "Line between segment",
            "Whole segment",
            "White segment",
            "perfect but got bubble",
            "<=2 segment",
        ],
        row,
        "CAMH",
        width=35,
    )
    row += 1

    # PCBA [dropdown] (FA PCBA)   Scrap? why? [dropdown]
    scrap_why_var = tk.StringVar(value="")
    fa_row = ttk.Frame(form_frame)
    fa_row.grid(row=row, column=0, columnspan=2, sticky="ew", pady=2)
    fa_row.columnconfigure(1, weight=1)
    fa_row.columnconfigure(3, weight=1)
    ttk.Label(fa_row, text="PCBA").grid(row=0, column=0, sticky="w", padx=(0, 8))
    ttk.Combobox(fa_row, textvariable=pcba_final_var, values=["", "Component knock off/burnt", "Component short/malfunction", "Scrap"], state="readonly", width=28).grid(row=0, column=1, sticky="ew", padx=2)
    ttk.Label(fa_row, text="Scrap? why?").grid(row=0, column=2, sticky="w", padx=(16, 8))
    ttk.Combobox(fa_row, textvariable=scrap_why_var, values=["", "Board Aging", "Open Pad", "Unable to capture image", "Short to ground"], state="readonly", width=22).grid(row=0, column=3, sticky="ew", padx=2)
    row += 1

    # Bad CAMH did assemble? got bubble? [checkbox right]
    _checkbox_row(form_frame, "Bad CAMH did assemble? got bubble?", bad_camh_assemble_bubble_var, row)
    row += 1

    # --- Buttons ---
    btn_frame = ttk.Frame(form_frame)
    btn_frame.grid(row=row, column=0, columnspan=2, sticky="ew", pady=16)
    btn_frame.columnconfigure(0, weight=1)
    btn_frame.columnconfigure(1, weight=1)
    btn_frame.columnconfigure(2, weight=1)
    btn_frame.columnconfigure(3, weight=1)

    def get_text(w: tk.Text) -> str:
        return w.get("1.0", "end").strip()

    def set_text(w: tk.Text, s: str) -> None:
        w.delete("1.0", "end")
        w.insert("1.0", s)

    def collect_data() -> dict:
        pcba_can_repair_val = (can_repair_ats_var.get() or "").strip().lower()
        ft_can_repair_val = (can_repair_ft_var.get() or "").strip().lower()
        return {
            "npf_all_reworked": _bool_to_str(npf_all_reworked_var.get()),
            "customer_request": _bool_to_str(customer_request_var.get()),
            "camera_model": camera_model_var.get(),
            "burnt": _bool_to_str(burnt_var.get()),
            # Visual check status used by report_rules.json
            "visual_check": "Got missing/burnt" if burnt_var.get() else "No missing/No Burnt",
            "power_on_unit": power_on_unit_var.get(),
            "remark_missing_burnt": get_text(remark_missing_burnt_text),
            "pcba": pcba_var.get(),
            "camh": camh_var.get(),
            "scintillator": scintillator_var.get(),
            "ecc_rework": ecc_rework_var.get(),
            "good_camh": good_camh_var.get(),
            "bad_camh": bad_camh_var.get(),
            "which_dvm_fail": get_text(which_dvm_fail_text),
            "can_repair_bad_camh": can_repair_bad_camh_var.get(),
            "component_cause": get_text(component_cause_text),
            "pcba_ats_result": pcba_ats_result_var.get(),
            "ats_result_if_failed": get_text(ats_result_if_failed_text),
            "can_repair_ats": can_repair_ats_var.get(),
            "component_cause_ats": get_text(component_cause_ats_text),
            "component_category": component_category_var.get(),
            "ft_result": ft_result_var.get(),
            "ft_fail_component_change": get_text(ft_fail_component_change_text),
            "can_repair_ft": can_repair_ft_var.get(),
            "ft_pass_which_camh": ft_pass_which_camh_var.get(),
            # booleans kept as real True/False where JSON rules expect them
            "npf_final": npf_final_var.get(),
            "camh_final": camh_final_var.get(),
            "pcba_final": pcba_final_var.get(),
            "scrap_why": scrap_why_var.get(),
            "bad_camh_assemble_bubble": _bool_to_str(bad_camh_assemble_bubble_var.get()),
            # Derived/alias fields used by report_rules.json
            "ats_fail_mode": get_text(ats_result_if_failed_text),
            "pcba_can_repair": pcba_can_repair_val,
            "pcba_component_category": component_category_var.get(),
            "ft_can_repair": ft_can_repair_val,
            "ft_fail_component": get_text(ft_fail_component_change_text),
            "ft_pass_camh": ft_pass_which_camh_var.get(),
            "fa_camh": camh_final_var.get(),
            "fa_pcba": pcba_final_var.get(),
            "pcba_scrap_why": scrap_why_var.get(),
            "pcba_scrap_component": component_category_var.get(),
            "pcba_component_name": component_category_var.get(),
            "bad_camh_condition": bad_camh_var.get(),
            "dvm_location": get_text(which_dvm_fail_text),
            "dvm_result": get_text(component_cause_text),
            # AXI matches: IF(OR(D18="scrap",D20="scrap"),"", "Perform testing AXI machine --> PASS")
            "axi_performed": pcba_can_repair_val != "scrap" and ft_can_repair_val != "scrap",
        }

    def clear_form() -> None:
        npf_all_reworked_var.set(False)
        customer_request_var.set(False)
        camera_model_var.set("3.3G")
        burnt_var.set(False)
        power_on_unit_var.set("")
        set_text(remark_missing_burnt_text, "")
        pcba_var.set("")
        camh_var.set("")
        scintillator_var.set("")
        ecc_rework_var.set("")
        good_camh_var.set("")
        bad_camh_var.set("")
        set_text(which_dvm_fail_text, "")
        can_repair_bad_camh_var.set("")
        set_text(component_cause_text, "")
        pcba_ats_result_var.set("")
        set_text(ats_result_if_failed_text, "")
        can_repair_ats_var.set("")
        set_text(component_cause_ats_text, "")
        component_category_var.set("")
        ft_result_var.set("")
        set_text(ft_fail_component_change_text, "")
        can_repair_ft_var.set("")
        ft_pass_which_camh_var.set("")
        npf_final_var.set(False)
        camh_final_var.set("")
        pcba_final_var.set("")
        scrap_why_var.set("")
        bad_camh_assemble_bubble_var.set(False)

    def _load_report_rules():
        """Load report_rules.json from the project directory."""
        rules_path = os.path.join(os.path.dirname(__file__), "report_rules.json")
        try:
            with open(rules_path, "r", encoding="utf-8") as handle:
                return json.load(handle)
        except (OSError, json.JSONDecodeError) as exc:
            messagebox.showerror("Report rules", f"Could not load report_rules.json:\n{exc}")
            return None

    def _match_conditions(form: dict, conditions: dict) -> bool:
        """Return True if form data satisfies all key/value pairs in conditions."""
        for key, expected in conditions.items():
            actual = form.get(key)
            if expected == "*":
                # wildcard: require a non-empty value
                if actual is None:
                    return False
                if isinstance(actual, str) and not actual.strip():
                    return False
                if isinstance(actual, bool) and not actual:
                    return False
                continue
            if isinstance(expected, bool):
                # normalise common string representations of booleans
                if isinstance(actual, str):
                    lower = actual.strip().lower()
                    if lower in ("yes", "true", "1"):
                        actual_val = True
                    elif lower in ("no", "false", "0", ""):
                        actual_val = False
                    else:
                        return False
                else:
                    actual_val = bool(actual)
                if actual_val is not expected:
                    return False
            else:
                if actual != expected:
                    return False
        return True

    def build_report_text(form: dict) -> str:
        rules = _load_report_rules()
        if not rules:
            return "No report rules available."

        sections_order = rules.get("order", [])
        sections = rules.get("sections", {})
        paragraphs: list[str] = []

        for section_name in sections_order:
            for rule in sections.get(section_name, []):
                cond = rule.get("when", {})
                if _match_conditions(form, cond):
                    text = rule.get("text", "")
                    try:
                        text = text.format(**form)
                    except Exception:
                        pass
                    text = _normalize_newlines(text).strip()
                    if text:
                        paragraphs.append(text)

        if not paragraphs:
            return "No matching report content for current form."

        # Add dynamic numbering 1), 2), 3) ... for step-style blocks only.
        # Do not renumber the intro or any explicit closing "Based on failure analysis..." text,
        # and once we reach any "Based on (the) failure analysis" line, stop numbering entirely
        # so that the subsequent FA paragraphs keep their own internal numbering.
        numbered_lines: list[str] = []
        step = 1
        stopped_numbering = False
        for para in paragraphs:
            lines = para.splitlines()
            if not lines:
                continue
            first_line = lines[0].strip()
            is_based_on = first_line.startswith("Based on failure analysis") or first_line.startswith(
                "Based on the failure analysis"
            )
            if is_based_on:
                stopped_numbering = True
            skip_number = (
                stopped_numbering
                or first_line.startswith("After received RMA")
                or is_based_on
            )
            if skip_number:
                numbered_lines.extend(lines)
            else:
                numbered_lines.append(f"{step}) {lines[0]}")
                numbered_lines.extend(lines[1:])
                step += 1
        return "\n".join(numbered_lines)

    def build_summary_fields(form: dict, failure_body: str) -> str:
        """Build key/value summary block similar to Sheet20 header."""
        # Helper flags
        npf_final = bool(form.get("npf_final"))
        customer_req = str(form.get("customer_request", "")).lower() == "yes"
        camh_fa = form.get("fa_camh") or ""
        pcba_fa = form.get("fa_pcba") or ""
        scint = form.get("scintillator") or ""
        comp_cat = form.get("component_category") or ""
        pcba_final = form.get("pcba_final") or ""
        ft_result = form.get("ft_result") or ""
        pcba_can_repair = form.get("pcba_can_repair") or ""
        pcba_scrap_why = form.get("pcba_scrap_why") or ""
        ats_fail_mode = form.get("ats_fail_mode") or ""

        # 1. Camera Main Issue – use only derived head_issue / pcba_issue (A45/A46), not NPF checkbox.
        # P/F status for key areas (matches database N42:N48).
        visual_status = "P" if not form.get("burnt") or str(form.get("burnt")).lower() in ("no", "false", "0") else "F"
        power_status = "P" if power_on_unit_var.get() == "Can Ping" else ("F" if power_on_unit_var.get() == "Cannot Ping" else "")
        pcba_status = "P" if pcba_var.get() == "Perfect" else ("F" if pcba_var.get() else "")
        camh_status = "" if not camh_var.get() else ("P" if camh_var.get() == "Perfect" else "F")
        scint_status = "" if not scint else ("P" if scint == "Good" else "F")
        ats_status = "" if not form.get("pcba_ats_result") else ("P" if form.get("pcba_ats_result") == "Pass" else "F")
        ft_status = "" if not ft_result else ("F" if ft_result == "Fail" else "P")

        # 2. Camera - Head (mirrors A51:A57 table via Good CAMH + CAMH selection)
        good_camh = (form.get("good_camh") or "").lower()
        camh_value = (form.get("camh") or "").lower()

        head_flags = []
        # Wafer trace bubble: any 'bubble' in good CAMH description
        head_flags.append(("Wafer trace bubble", "bubble" in good_camh))
        # HCTE Test Fail: contains 'hcte'
        head_flags.append(("HCTE Test Fail", "hcte" in good_camh))
        # CCD segment classifications from CAMH dropdown (these should take priority over NPF)
        head_flags.append(("CCD Segment Die-<=2 segment", "<= 2 segment" in camh_value))
        head_flags.append(("CCD Segment Die-Line between segment", "line between segment" in camh_value))
        head_flags.append(("CCD Segment Die-White segment", "white segment" in camh_value))
        head_flags.append(("CCD Segment Die-Whole segment", "whole segment" in camh_value))
        # No Problem Found (head): ALL PASS
        head_flags.append(("No Problem Found", good_camh.strip() in ("all pass", "all pass + rework", "all pass+rework")))

        camera_head = ""
        for label, flag in head_flags:
            if flag:
                camera_head = label
                break

        # Camera Main Issue: four-row priority from A43:B46 (no NPF checkbox).
        # Head issue must come from true defect rows only (A51:A52,A54:A57), not the head-NPF row.
        head_issue = any(flag for label, flag in head_flags if label != "No Problem Found")
        pcba_issue = (
            visual_status == "F"
            or power_status == "F"
            or pcba_status == "F"
            or ats_status == "F"
            or ft_status == "F"
        )
        if not head_issue and not pcba_issue:
            camera_main_issue = "No Problem Found"
        elif head_issue and pcba_issue:
            camera_main_issue = "Camera - Head,Camera - PCBA"
        elif head_issue:
            camera_main_issue = "Camera - Head"
        elif pcba_issue:
            camera_main_issue = "Camera - PCBA"
        else:
            camera_main_issue = "No Problem Found"

        # 3. Camera - PCBA (mirrors A59:A68 table approximately)
        camera_pcba = ""
        pcba_component_ic = pcba_status == "F" and comp_cat == "IC"
        pcba_component_cap = pcba_status == "F" and comp_cat == "Capacitor"
        pcba_component_ind = pcba_status == "F" and comp_cat == "Inductor"
        firmware_flag = pcba_scrap_why.lower() == "unable to capture image"
        # Camera - PCBA "No Problem Found" follows the Excel rule:
        # AND(N42="P",N43="P",N44="P",N47="P",N48="P") i.e. all PCBA-related P/F flags are P.
        pcba_npf_status = (
            visual_status == "P"
            and power_status == "P"
            and pcba_status == "P"
            and ats_status == "P"
            and ft_status == "P"
        )

        pcba_flags = [
            ("Customer Request", customer_req),
            ("Handling-Customer", visual_status == "F"),
            ("Component-IC", pcba_component_ic),
            ("Component-Capacitor", pcba_component_cap),
            ("Component-Inductor", pcba_component_ind),
            ("Firmware", firmware_flag),
            ("No Problem Found", pcba_npf_status),
        ]

        for label, flag in pcba_flags:
            if flag:
                camera_pcba = label
                break

        # 4. Scintillator – report table label (A72:B76), not raw dropdown value.
        _scint_label = {
            "Good": "No Problem Found",
            "Aging": "Aging",
            "Gap": "Gap",
            "Bubble": "Bubble",
            "Bent": "Bent",
        }
        scintillator_field = _scint_label.get(scint, scint) if scint else ""

        # 5. Root Cause Categories
        root_cause_flags = [
            ("Customer - Handling", visual_status == "F"),
            # Customer - Retest (no direct mapping with current GUI)
            ("Customer - Upgrade Hardware", customer_req),
            (
                "Material Quality",
                any(flag for _, flag in head_flags if _ != "No Problem Found"),
            ),
            (
                "Material - component",
                pcba_component_ic or pcba_component_cap or pcba_component_ind,
            ),
            (
                "No Problem Found",
                not head_issue and not pcba_issue,
            ),
        ]
        root_cause_cat = ""
        for label, flag in root_cause_flags:
            if flag:
                root_cause_cat = label
                break

        # 6. Failure Analysis / Root Cause
        failure_root_cause = failure_body

        # 7–8. Disposition + Disposition on RMA unitRequired (rows 219–223)
        ft_can_repair = form.get("ft_can_repair") or ""
        scrap_flag = pcba_can_repair == "scrap" or ft_can_repair == "scrap"
        unable_flag = (
            pcba_can_repair == "unable"
            or ft_can_repair == "unable"
            or ft_result in ("Pass with new camh", "Pass with swap camh")
        )

        # Follow Excel's VLOOKUP table order: Scrap, Unable to Repair, Repair, No problem found.
        if scrap_flag:
            disposition = "Scrap"
        elif npf_final:
            disposition = "No problem found"
        elif unable_flag:
            disposition = "Unable to Repair"
        else:
            disposition = "Repair"
        disposition_rma = disposition

        # 9. Reason to Scrap – if scrapped, show only the final FA paragraph starting from "Based on (the) failure analysis"
        if disposition == "Scrap" and failure_root_cause:
            text = failure_root_cause
            # Try to find the last occurrence of the FA closing prefix
            idx = max(
                text.rfind("Based on failure analysis"),
                text.rfind("Based on the failure analysis"),
                text.rfind("Based on the Failure Analysis"),
            )
            if idx != -1:
                reason_to_scrap = text[idx:].lstrip()
            else:
                reason_to_scrap = text
        else:
            reason_to_scrap = "N/A"

        # 10–11. Need Replacement + Replaced By – single value from first TRUE row (A95:B100).
        ft_lower = ft_result.strip().lower() if ft_result else ""
        unable = pcba_can_repair == "unable" or ft_can_repair == "unable"
        e20_new_camh = (form.get("ft_pass_which_camh") or "").strip().lower() == "new camh"
        replaced_by = "N/A"
        if ft_lower == "pass with new camh" and unable:
            replaced_by = "PASS with new camh"
        elif ft_lower == "pass with swap camh" and unable:
            replaced_by = "Swap Camera Head and New PCBA"
        elif ft_lower == "pass with new camh" or e20_new_camh:
            replaced_by = "New Camera Head"
        elif ft_lower == "pass with swap camh":
            replaced_by = "Swap Camera Head"
        elif unable:
            replaced_by = "New PCBA"
        need_replacement = "Yes" if replaced_by != "N/A" else "No"

        # 12. Action Taken on Repairing – CONCATENATE(B104, CHAR(10), B105, CHAR(10), B106) with exact wording.
        ecc_rework = (form.get("ecc_rework") or "").strip()
        camera_model = (form.get("camera_model") or "").strip()
        b104 = ""
        if ecc_rework == "3G rework":
            b104 = (
                "- R89, R91, R97, R98 change to 10kohm resistor --> 2220-0057 x4\n"
                "- R93, R95, R99, R100 change to 100kohm resistor --> 2220-0060 x4\n"
                "- C112, C113, C114, C116 added 10nF capacitor --> 2230-0014 x4 \n"
            )
        elif ecc_rework == "3.1G old Rework":
            b104 = (
                "- R89, R91, R97, R98 change to 10kohm resistor --> 2220-0057 x4\n"
                "- R93, R95, R99, R100 change to 100kohm resistor --> 2220-0060 x4\n"
                "- C112, C113, C114, C116 added 10nF capacitor --> 2230-0014 x4 \n"
                "- C60  change to 10nF capacitor --> 2230-0185 x1\n"
                "- C65,replace Zener diodes --> 2250-0057 x1\n"
                "- Scratch the trace and add on 4 resistor in Power IC --> 2220-0399 x4\n"
                "- R451 change to 4.64K ohm resistor -->        2220-0054 x1\n"
                "- R487, R524 change to 7.5k ohm resistor --> 2220-0056 x2"
            )
        elif ecc_rework == "3.1G new Rework":
            b104 = (
                "- R89, R91, R97, R98 change to 10kohm resistor --> 2220-0057 x4\n"
                "- R93, R95, R99, R100 change to 100kohm resistor --> 2220-0060 x4\n"
                "- C112, C113, C114, C116 added 10nF capacitor --> 2230-0014 x4 \n"
                "- C60  change to 10nF capacitor --> 2230-0185 x1\n"
                "- C65,replace Zener diodes --> 2250-0057 x1\n"
                "- R15, R18, R19, R20, added 4.7K ohm resistor --> 2220-0279 x4\n"
                "- R451 change to 4.64K ohm resistor -->        2220-0054 x1"
            )
        if ft_lower == "pass with swap camh":
            b105 = "- Swap camera head ----> 89504-0008 x1"
        elif ft_lower == "pass with new camh" or e20_new_camh:
            b105 = "- New camera head ----> 89504-0008 x1"
        else:
            b105 = ""
        if unable and camera_model == "3.1G New":
            b106 = "- PCBA ---> 89504-0007 x1"
        elif unable and camera_model in ("3.3G", "3.4G"):
            b106 = "- PCBA ---> 89504-0010 x1"
        else:
            b106 = ""
        action_parts = [p for p in (b104, b105, b106) if p]
        action_taken_str = "\n".join(action_parts) if action_parts else "N/A"

        # 13. Countermeasure – VLOOKUP(B24, A111:B119); exact keys and paragraph text from database.
        _fa_camh_to_db_key = {
            "<=2 segment and 89504-0004": "<=2 segment and 89504-0004",
            "<=2 segment and 89504-0004 to 89504-00085": "<=2 segment and 89504-0004 to 89504-00085",
            "HCTE Fail": "HCTE Fail",
            "Old CCD fail": "Old CCD fail",
            "Line between segment": "Line Between Segment",
            "Whole segment": "Whole Segment",
            "White segment": "White segment",
            "<=2 segment": "<=2 segment",
            "Whole Segment (5V Cap)": "Whole Segment (5V Cap)",
        }
        db_key = _fa_camh_to_db_key.get(camh_fa, camh_fa)
        countermeasure_map = {
            "<=2 segment and 89504-0004": (
                "1.Reduce speed of dispenser during wirebond encapsulation process to minimize any risk/stress on the wirebond\n"
                "Result: Speed had been reduced from 4mm/sec to 1mm/sec.\n\n"
                "2.Reduce pressure of ionizing blowing to minimize the pressure given to the wirebond during ionizing blowing\n"
                "Result: Ionizing pressure had been reduced from 0.03 to 0.02mPa\n\n"
                "3.  Improve encapsulation process by adding adhesion on the wirebond pad\n"
                "Result: Adhesion process implemented on Jan'21 and start to ship on Feb'21\n\n"
                "4. Electromigration fix implemented and ship on Jun'22 and and CCD S/ N start by 89504-00085XXX. "
            ),
            "<=2 segment and 89504-0004 to 89504-00085": (
                "1. change the capacitor (C60) to 10uF (Part ID: 2230-0185)\n"
                "2. change the resistors (R89, R91, R97, R98) to 10kohm (Part ID: 2220-0057)\n"
                "3. change the resistors (R93, R95, R99, R100) to 100kohm (Part ID: 2220-0060)\n"
                "4. add capacitors with 10nF (Part ID: 2230-0014) to empty footprint C112, C113, C114, C116\n"
                "able to reduce the switching effect of the VDDF and VDDR to the +20V. Thus, improve the stability of +20V voltage and the current drawn by the CCD's VOD bus.\n\n"
                "5. Electromigration fix implemented and ship on Jun'22 and and CCD S/ N start by 89504-00085XXX. "
            ),
            "HCTE Fail": (
                " 1. Using additional  HTCE measurement methods at Functional Test to filter those CCD fall below 0.9. Our main goal is  to improve CCD quality. "
            ),
            "Old CCD fail": (
                "1. Will disassemble al the cover to check wafer trace if the CAMH is under good condition.\n"
                "2. Using additional  HTCE measurement methods at Functional Test to filter those CCD fall below 0.9. Our main goal is  to improve CCD quality. "
            ),
            "Line Between Segment": "1. Collect data of IC damaged/degradation for within warranty unit and feedback to supplier and RnD team",
            "Whole Segment": (
                "1. Implemented the CCD Capacitor (C443 and C444) to the high voltage rating at 17 Oct 2023 and the serial number start by  9704-0111769, 3.3G CAMH start by 89504-0008 9318\n\n"
                "2. Proceed CD&A test aging test to 6 hour\n"
                "3. Implemented the CCD Capacitor to the high voltage rating and reheat the capacitor before rework\n"
                "4. capacitor pre-heating up to 150degC before mount/solder to the CCD + solder cleaning staging to allow CCD/cap back to normal temperature to minimize thermal shock impact start by 7 Dec 23\n"
                "5. The CCD that implement and build by supplier 3.1G camera head S/N start by 89504-00089360, 3.3G camera head s/n by 89504-00133802 and camera card start by 9704-0113401(5 March 2024)."
            ),
            "<=2 segment": "1. Collect data of IC damaged/degradation for within warranty unit and feedback to supplier and RnD team",
            "White segment": "1. Collect data of IC damaged/degradation for within warranty unit and feedback to supplier and RnD team",
            "Whole Segment (5V Cap)": (
                "1. fulling resecreening all the capacitor on the CCD board.\n"
                "2. 5V CCD cap change to flexi cap - start by 10 Feb 25, 3.3G: 9704-0115871, and for refurbish camera 3.1G CCD: 89504-0008C081 and 3.3G CCD: 89504-00136910"
            ),
        }
        countermeasure = countermeasure_map.get(db_key, "N/A")

        rows = [
            ("Camera Main Issue", camera_main_issue or "N/A"),
            ("Camera - Head", camera_head or "N/A"),
            ("Camera - PCBA", camera_pcba or "N/A"),
            ("Scintillator", scintillator_field or "N/A"),
            ("Root Cause Categories", root_cause_cat or "N/A"),
            ("Failure Analysis / Root Cause", _normalize_newlines(failure_root_cause or "N/A")),
            ("Disposition", disposition),
            ("Disposition on RMA unitRequired", disposition_rma),
            ("Reason to Scrap", _normalize_newlines(reason_to_scrap or "N/A")),
            ("Need Replacement", need_replacement),
            ("Replaced By", replaced_by),
            ("Action Taken on Repairing", _normalize_newlines(action_taken_str or "N/A")),
            ("Countermeasure", _normalize_newlines(countermeasure or "N/A")),
        ]
        return rows

    def show_report_window() -> None:
        data = collect_data()
        body_text = build_report_text(data)
        rows = build_summary_fields(data, body_text)

        win = tk.Toplevel(root)
        win.title("Generated JIRA Report")
        win.geometry("920x620")

        def _copy_report():
            names = [name for name, _ in rows]
            values = [(value or "").strip().replace("\n", " ") for _, value in rows]
            tsv = "\t".join(names) + "\n" + "\t".join(values)
            win.clipboard_clear()
            win.clipboard_append(tsv)

        main = ttk.Frame(win, padding=8)
        main.pack(fill="both", expand=True)
        canvas = tk.Canvas(main, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main, orient="vertical", command=canvas.yview)
        content = ttk.Frame(canvas)

        def _on_content_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(canvas_window_id, width=e.width)

        content.bind("<Configure>", _on_content_configure)
        canvas_window_id = canvas.create_window((0, 0), window=content, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window_id, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        scroll_units = 4

        def _on_mousewheel(e):
            if getattr(e, "num", None) == 5:
                canvas.yview_scroll(scroll_units, "units")
            elif getattr(e, "num", None) == 4:
                canvas.yview_scroll(-scroll_units, "units")
            elif hasattr(e, "delta"):
                d = e.delta if abs(e.delta) < 120 else e.delta // 120
                canvas.yview_scroll(-d * scroll_units, "units")

        def _bind_wheel(w):
            w.bind("<MouseWheel>", _on_mousewheel)
            w.bind("<Button-4>", _on_mousewheel)
            w.bind("<Button-5>", _on_mousewheel)
            for c in w.winfo_children():
                _bind_wheel(c)

        btn_row = ttk.Frame(content)
        btn_row.pack(fill="x", pady=(0, 8))
        ttk.Button(btn_row, text="Copy report to clipboard", command=_copy_report).pack(side="left")

        table_frame = ttk.Frame(content)
        table_frame.pack(fill="both", expand=True)
        _bind_wheel(canvas)
        _bind_wheel(content)

        table_frame.columnconfigure(1, weight=1)
        value_wraplength = 680
        for r, (field_name, value) in enumerate(rows):
            ttk.Label(table_frame, text=field_name + ":", font=("", 10, "bold")).grid(
                row=r, column=0, sticky="nw", padx=(0, 12), pady=4
            )
            lbl = ttk.Label(
                table_frame,
                text=value or "",
                wraplength=value_wraplength,
                anchor="w",
                justify="left",
            )
            lbl.grid(row=r, column=1, sticky="nw", pady=4)

        win.focus_set()

    def save_to_csv() -> None:
        data = collect_data()
        csv_path = os.path.join(os.getcwd(), "failure_form_entries.csv")
        _ensure_directory(csv_path)
        file_exists = os.path.isfile(csv_path)
        try:
            with open(csv_path, "a", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=list(data.keys()))
                if not file_exists:
                    writer.writeheader()
                writer.writerow(data)
        except OSError as exc:
            messagebox.showerror("Save failed", str(exc))
            return
        messagebox.showinfo("Saved", f"Form saved to\n{csv_path}")

    ttk.Button(btn_frame, text="Save", command=save_to_csv).grid(row=0, column=0, padx=4)
    ttk.Button(btn_frame, text="Generate Report", command=show_report_window).grid(row=0, column=1, padx=4)
    ttk.Button(btn_frame, text="Clear", command=clear_form).grid(row=0, column=2, padx=4)
    ttk.Button(btn_frame, text="Close", command=root.destroy).grid(row=0, column=3, padx=4)

    for w in (root, canvas, form_frame):
        w.bind("<MouseWheel>", _on_mousewheel)
        w.bind("<Button-4>", _on_mousewheel)  # Linux scroll up
        w.bind("<Button-5>", _on_mousewheel)  # Linux scroll down

    def _bind_mousewheel(widget):
        widget.bind("<MouseWheel>", _on_mousewheel)
        widget.bind("<Button-4>", _on_mousewheel)
        widget.bind("<Button-5>", _on_mousewheel)
        for child in widget.winfo_children():
            _bind_mousewheel(child)
    _bind_mousewheel(form_frame)

    root.mainloop()


if __name__ == "__main__":
    build_failure_form_gui()
