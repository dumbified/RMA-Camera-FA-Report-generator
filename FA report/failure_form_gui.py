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


def _load_31g_component_lookup() -> dict[str, tuple[str, str, str]]:
    """
    Load 3.1G component CSV (MfgComment, PartID, Description).
    Return mapping: component_ref_lower -> (PartID, Description, display_ref).
    Keys are lowercased for case-insensitive lookup; display_ref is the ref as in CSV.
    """
    path = os.path.join(os.path.dirname(__file__), "3.1G component.csv")
    ref_to_part: dict[str, tuple[str, str, str]] = {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                mfg = (row.get("MfgComment") or "").strip()
                part_id = (row.get("PartID") or "").strip()
                desc = (row.get("Description") or "").strip()
                if not mfg or not part_id:
                    continue
                for ref in (r.strip() for r in mfg.split(",")):
                    ref = ref.strip()
                    if ref:
                        key = ref.lower()
                        if key not in ref_to_part:
                            ref_to_part[key] = (part_id, desc, ref)
    except (OSError, csv.Error):
        pass
    return ref_to_part


def _build_component_action_lines(
    e18: str, c20: str, ref_to_part: dict[str, tuple[str, str, str]]
) -> str:
    """
    Build Action Taken lines from E18 (PCBA ATS component cause) and C20 (FT If fail, component change).
    Format per line: {refs} {Description} {PartID}----->{count}
    e.g. "F1 FUSE,SMD,SLOW-BLOW,2A,125V,OMNI-BLOK 2215-0026----->1"
    Lookup is case-insensitive.
    """
    combined = f"{e18},{c20}".replace("\n", " ").replace("\r", " ")
    refs = [r.strip() for r in combined.split(",") if r.strip()]
    if not refs or not ref_to_part:
        return ""
    # group by (part_id, desc) -> list of display_refs
    group_key_to_refs: dict[tuple[str, str], list[str]] = {}
    for ref in refs:
        key = ref.lower()
        if key in ref_to_part:
            part_id, desc, display_ref = ref_to_part[key]
            gk = (part_id, desc)
            group_key_to_refs.setdefault(gk, []).append(display_ref)
    if not group_key_to_refs:
        return ""
    lines = []
    for (part_id, desc), ref_list in sorted(group_key_to_refs.items()):
        ref_list_sorted = sorted(set(ref_list), key=lambda r: (r[0].upper(), len(r), r))
        refs_str = ", ".join(ref_list_sorted)
        count = len(ref_list)
        lines.append(f"- {refs_str} {desc} {part_id}----->{count}")
    return "\n".join(lines)


def _checkbox_row(parent, text: str, variable: tk.BooleanVar, row: int) -> ttk.Checkbutton:
    """Label on left, checkbox (small box) on right."""
    f = ttk.Frame(parent)
    f.grid(row=row, column=0, columnspan=2, sticky="w", pady=2)
    ttk.Label(f, text=text).pack(side="left")
    cb = ttk.Checkbutton(f, variable=variable)
    cb.pack(side="right", padx=(8, 0))
    return cb


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
    vit_id_var = tk.StringVar(value="")
    npf_all_reworked_var = tk.BooleanVar(value=False)
    customer_request_var = tk.BooleanVar(value=False)
    customer_request_type_var = tk.StringVar(value="")
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

    # --- VIT ID [single-line entry] ---
    ttk.Label(form_frame, text="VIT ID").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=2)
    vit_id_entry = ttk.Entry(form_frame, textvariable=vit_id_var, width=20)
    vit_id_entry.grid(row=row, column=1, sticky="ew", pady=2)
    row += 1

    # --- If NPF and all reworked? [checkbox right] ---
    npf_all_reworked_cb = _checkbox_row(form_frame, "If NPF and all reworked?", npf_all_reworked_var, row)
    row += 1

    # --- Customer Request? [checkbox + reason dropdown] ---
    customer_req_row = ttk.Frame(form_frame)
    customer_req_row.grid(row=row, column=0, columnspan=2, sticky="w", pady=2)
    ttk.Label(customer_req_row, text="Customer Request?").pack(side="left")
    customer_request_cb = ttk.Checkbutton(customer_req_row, variable=customer_request_var)
    customer_request_cb.pack(side="left", padx=(8, 0))
    customer_request_type_cb = ttk.Combobox(
        customer_req_row,
        textvariable=customer_request_type_var,
        values=[
            "",
            "Rework",
            "Rework + CAMH",
            "DONE REWORK + CAMH",
            "Rework (DONE REWORKED)",
            "Rework ( <= 2 segment die)",
            "Rework (HCTE Fail)",
        ],
        state="readonly",
        width=30,
    )
    customer_request_type_cb.pack(side="left", padx=(16, 0))
    row += 1

    ttk.Separator(form_frame, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=8)
    row += 1

    # --- Camera Model [dropdown] ---
    camera_model_cb = _dropdown(form_frame, camera_model_var, ["3G", "3.1G Old", "3.1G New", "3.3G", "3.4G"], row, "Camera Model", width=10)
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
    camh_screening_cb = ttk.Combobox(
        f3,
        textvariable=camh_var,
        values=["Whole Segment", "Perfect", "<= 2 segment", "Line Between Segment", "White Segment", "Multiple Segment Die"],
        state="readonly",
        width=18,
    )
    camh_screening_cb.grid(row=0, column=3, sticky="ew", padx=2)
    ttk.Label(f3, text="Scintillator").grid(row=0, column=4, sticky="w", padx=(8, 4))
    ttk.Combobox(f3, textvariable=scintillator_var, values=["Good", "Aging", "Gap", "Bubble", "Bent"], state="readonly", width=12).grid(row=0, column=5, sticky="ew", padx=2)
    row += 1

    ttk.Separator(form_frame, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=8)
    row += 1

    # --- ECC Rework [dropdown] ---
    ecc_rework_cb = _dropdown(form_frame, ecc_rework_var, ["NA", "3G rework", "3.1G old Rework", "3.1G new Rework"], row, "ECC Rework", width=18)
    row += 1

    # Good CAMH? [dropdown] (IF good camh)
    good_camh_cb = _dropdown(form_frame, good_camh_var, ["NA", "All Pass", "Got Bubble", "HCTE Fail", "Got bubble and HCTE Fail", "All PASS + Rework"], row, "Good CAMH?", width=22)
    row += 1

    # Bad CAMH? [dropdown]
    bad_camh_cb = _dropdown(form_frame, bad_camh_var, ["NA", "Whole Segment", "Line between segment", "Die segment", "Multiple segment DIE", "White segment"], row, "Bad CAMH?", width=20)
    row += 1

    # Which DVM show it fail [multiline textbox]
    which_dvm_fail_text = _textbox(form_frame, "Which DVM show it fail", 3, row)
    row += 1

    # Can repair? [dropdown]
    can_repair_bad_camh_cb = _dropdown(form_frame, can_repair_bad_camh_var, ["", "able", "unable"], row, "Can repair?", width=10)
    row += 1

    # Component Cause: [multiline textbox]
    component_cause_text = _textbox(form_frame, "Component Cause:", 2, row)
    row += 1

    ttk.Separator(form_frame, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=8)
    row += 1

    # --- PCBA ATS result [dropdown] ---
    pcba_ats_result_cb = _dropdown(form_frame, pcba_ats_result_var, ["Pass", "Fail"], row, "PCBA ATS result")
    row += 1

    # If failed, state the ATS result [multiline textbox]
    ats_result_if_failed_text = _textbox(form_frame, "If failed, state the ATS result", 2, row)
    row += 1

    # Can repair? [dropdown]
    can_repair_ats_cb = _dropdown(form_frame, can_repair_ats_var, ["", "able", "unable", "Scrap"], row, "Can repair?", width=10)
    row += 1

    # Component cause: [multiline textbox]
    component_cause_ats_text = _textbox(form_frame, "Component cause:", 2, row)
    row += 1

    # Component Category: [dropdown - single row] (PCBA component categories)
    component_category_cb = _dropdown(form_frame, component_category_var, ["", "IC", "Capacitor", "Inductor"], row, "Component Category:")
    row += 1

    def _clear_textbox(w: tk.Text) -> None:
        prev_state = str(w.cget("state"))
        if prev_state == "disabled":
            w.config(state="normal")
        w.delete("1.0", "end")
        if prev_state == "disabled":
            w.config(state="disabled")

    def _set_widget_var(widget: tk.Widget, option: str, value) -> None:
        var_name = str(widget.cget(option) or "")
        if var_name:
            root.setvar(var_name, value)

    def update_screening_camh_fields() -> None:
        camh_val = (camh_var.get() or "").strip()
        if camh_val == "" or camh_val == "Perfect":
            ecc_rework_var.set("")
            good_camh_var.set("")
            bad_camh_var.set("")
            can_repair_bad_camh_var.set("")
            _clear_textbox(which_dvm_fail_text)
            _clear_textbox(component_cause_text)
            ecc_rework_cb.config(state="disabled")
            good_camh_cb.config(state="disabled")
            bad_camh_cb.config(state="disabled")
            can_repair_bad_camh_cb.config(state="disabled")
            which_dvm_fail_text.config(state="disabled")
            component_cause_text.config(state="disabled")
        else:
            good_camh_var.set("")
            good_camh_cb.config(state="disabled")
            ecc_rework_cb.config(state="readonly")
            bad_camh_cb.config(state="readonly")
            can_repair_bad_camh_cb.config(state="readonly")
            which_dvm_fail_text.config(state="normal")
            component_cause_text.config(state="normal")

    camh_screening_cb.bind("<<ComboboxSelected>>", lambda e: refresh_interaction_states())

    def update_pcba_ats_fail_fields() -> None:
        if pcba_ats_result_var.get() == "Fail":
            ats_result_if_failed_text.config(state="normal")
            can_repair_ats_cb.config(state="readonly")
            component_cause_ats_text.config(state="normal")
            component_category_cb.config(state="readonly")
        else:
            _clear_textbox(ats_result_if_failed_text)
            _set_widget_var(can_repair_ats_cb, "textvariable", "")
            _clear_textbox(component_cause_ats_text)
            _set_widget_var(component_category_cb, "textvariable", "")
            ats_result_if_failed_text.config(state="disabled")
            can_repair_ats_cb.config(state="disabled")
            component_cause_ats_text.config(state="disabled")
            component_category_cb.config(state="disabled")

    pcba_ats_result_cb.bind("<<ComboboxSelected>>", lambda e: refresh_interaction_states())

    ttk.Separator(form_frame, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=8)
    row += 1

    # --- FT ---
    ttk.Label(form_frame, text="FT", font=("", 10, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", pady=(4, 2))
    row += 1

    # Result: [dropdown] (FT results)
    ft_result_cb = _dropdown(
        form_frame,
        ft_result_var,
        ["Pass with new camh", "Pass with ori camh", "Pass with swap camh", "Pass with repaired camh", "Fail"],
        row,
        "Result:",
        width=22,
    )
    row += 1

    # IF fail, component change? [multiline textbox]
    ft_fail_component_change_text = _textbox(form_frame, "IF fail, component change?", 2, row)
    row += 1

    # Can repair? [dropdown]
    can_repair_ft_cb = _dropdown(form_frame, can_repair_ft_var, ["", "able", "unable", "Scrap"], row, "Can repair?", width=10)
    row += 1

    # IF pass, which CAMH to use? [dropdown]
    ft_pass_which_camh_cb = _dropdown(
        form_frame,
        ft_pass_which_camh_var,
        ["", "Ori camh", "New camh", "Swap camh", "Repaired camh"],
        row,
        "IF pass, which CAMH to use?",
    )
    row += 1

    def update_ft_fail_fields() -> None:
        r = ft_result_var.get() or ""
        if r == "Fail":
            ft_fail_component_change_text.config(state="normal")
            can_repair_ft_cb.config(state="readonly")
            _set_widget_var(ft_pass_which_camh_cb, "textvariable", "")
            ft_pass_which_camh_cb.config(state="disabled")
        elif r.startswith("Pass"):
            _clear_textbox(ft_fail_component_change_text)
            _set_widget_var(can_repair_ft_cb, "textvariable", "")
            ft_fail_component_change_text.config(state="disabled")
            can_repair_ft_cb.config(state="disabled")
            ft_pass_which_camh_cb.config(state="readonly")
        else:
            _clear_textbox(ft_fail_component_change_text)
            _set_widget_var(can_repair_ft_cb, "textvariable", "")
            _set_widget_var(ft_pass_which_camh_cb, "textvariable", "")
            ft_fail_component_change_text.config(state="disabled")
            can_repair_ft_cb.config(state="disabled")
            ft_pass_which_camh_cb.config(state="disabled")

    ft_result_cb.bind("<<ComboboxSelected>>", lambda e: refresh_interaction_states())

    ttk.Separator(form_frame, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=8)
    row += 1

    # --- FA ---
    ttk.Label(form_frame, text="FA", font=("", 10, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", pady=(4, 2))
    row += 1

    # NPF? [checkbox right]
    npf_final_cb = _checkbox_row(form_frame, "NPF?", npf_final_var, row)
    row += 1

    # CAMH [dropdown] (FA Camh)
    camh_final_cb = _dropdown(
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
    pcba_final_cb = ttk.Combobox(
        fa_row,
        textvariable=pcba_final_var,
        values=["", "Component knock off/burnt", "Component short/malfunction", "Scrap"],
        state="readonly",
        width=28,
    )
    pcba_final_cb.grid(row=0, column=1, sticky="ew", padx=2)
    ttk.Label(fa_row, text="Scrap? why?").grid(row=0, column=2, sticky="w", padx=(16, 8))
    scrap_why_cb = ttk.Combobox(
        fa_row,
        textvariable=scrap_why_var,
        values=["", "Board Aging", "Open Pad", "Unable to capture image", "Short to ground"],
        state="readonly",
        width=22,
    )
    scrap_why_cb.grid(row=0, column=3, sticky="ew", padx=2)
    row += 1

    # Bad CAMH did assemble? got bubble? [checkbox right]
    bad_camh_assemble_bubble_cb = _checkbox_row(form_frame, "Bad CAMH did assemble? got bubble?", bad_camh_assemble_bubble_var, row)
    row += 1

    def update_fa_npf_fields() -> None:
        disabled = bool(npf_final_var.get())
        if disabled:
            _set_widget_var(camh_final_cb, "textvariable", "")
            _set_widget_var(pcba_final_cb, "textvariable", "")
            _set_widget_var(scrap_why_cb, "textvariable", "")
            _set_widget_var(bad_camh_assemble_bubble_cb, "variable", False)
            camh_final_cb.config(state="disabled")
            pcba_final_cb.config(state="disabled")
            scrap_why_cb.config(state="disabled")
            bad_camh_assemble_bubble_cb.config(state="disabled")
        else:
            camh_final_cb.config(state="readonly")
            pcba_final_cb.config(state="readonly")
            scrap_why_cb.config(state="readonly")
            bad_camh_assemble_bubble_cb.config(state="normal")

    def _set_form_inputs_locked(parent: tk.Widget, locked: bool, skip: set | None = None) -> None:
        for child in parent.winfo_children():
            if skip and child in skip:
                _set_form_inputs_locked(child, locked, skip)
                continue
            if isinstance(child, tk.Text):
                if locked:
                    _clear_textbox(child)
                child.config(state="disabled" if locked else "normal")
            elif isinstance(child, ttk.Combobox):
                if locked:
                    _set_widget_var(child, "textvariable", "")
                child.config(state="disabled" if locked else "readonly")
            elif isinstance(child, ttk.Checkbutton):
                if locked:
                    _set_widget_var(child, "variable", False)
                child.config(state="disabled" if locked else "normal")
            elif isinstance(child, (tk.Entry, ttk.Entry)):
                if locked:
                    var_name = str(child.cget("textvariable") or "")
                    if var_name:
                        root.setvar(var_name, "")
                    else:
                        prev_state = str(child.cget("state"))
                        if prev_state == "disabled":
                            child.config(state="normal")
                        child.delete(0, "end")
                child.config(state="disabled" if locked else "normal")
            _set_form_inputs_locked(child, locked, skip)

    def refresh_interaction_states() -> None:
        if npf_all_reworked_var.get():
            _set_form_inputs_locked(form_frame, True, skip={npf_all_reworked_cb})
            return
        _set_form_inputs_locked(form_frame, False)
        update_screening_camh_fields()
        update_pcba_ats_fail_fields()
        update_ft_fail_fields()
        update_fa_npf_fields()

        if customer_request_var.get():
            always_enabled = {
                vit_id_entry,
                npf_all_reworked_cb,
                customer_request_cb,
                customer_request_type_cb,
                camera_model_cb,
            }
            _set_form_inputs_locked(form_frame, True, skip=always_enabled)
            customer_request_type_cb.config(state="readonly")
        else:
            customer_request_type_var.set("")
            customer_request_type_cb.config(state="disabled")

    npf_all_reworked_var.trace_add("write", lambda *_: refresh_interaction_states())
    npf_final_var.trace_add("write", lambda *_: refresh_interaction_states())
    customer_request_var.trace_add("write", lambda *_: refresh_interaction_states())
    refresh_interaction_states()

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
        data = {
            "vit_id": vit_id_var.get().strip(),
            "npf_all_reworked": _bool_to_str(npf_all_reworked_var.get()),
            "customer_request": _bool_to_str(customer_request_var.get()),
            "customer_request_type": customer_request_type_var.get(),
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
        camh_val = (data.get("camh") or "").strip()
        if camh_val == "" or camh_val == "Perfect":
            data["ecc_rework"] = ""
            data["good_camh"] = ""
            data["bad_camh"] = ""
            data["which_dvm_fail"] = ""
            data["can_repair_bad_camh"] = ""
            data["component_cause"] = ""
            data["bad_camh_condition"] = ""
            data["dvm_result"] = ""
        if (data.get("pcba_ats_result") or "").strip() != "Fail":
            data["ats_result_if_failed"] = ""
            data["can_repair_ats"] = ""
            data["component_cause_ats"] = ""
            data["component_category"] = ""
            data["ats_fail_mode"] = ""
            data["pcba_can_repair"] = ""
            data["pcba_component_category"] = ""
            data["pcba_scrap_component"] = ""
            data["pcba_component_name"] = ""
        ft_res = (data.get("ft_result") or "").strip()
        if ft_res == "Fail":
            data["ft_pass_which_camh"] = ""
            data["ft_pass_camh"] = ""
        else:
            data["ft_fail_component_change"] = ""
            data["can_repair_ft"] = ""
            data["ft_can_repair"] = ""
            data["ft_fail_component"] = ""
        if bool(data.get("npf_final")):
            data["camh_final"] = ""
            data["pcba_final"] = ""
            data["scrap_why"] = ""
            data["fa_camh"] = ""
            data["fa_pcba"] = ""
            data["pcba_scrap_why"] = ""
            data["bad_camh_assemble_bubble"] = "No"
        return data

    def clear_form() -> None:
        vit_id_var.set("")
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

    def _load_customer_request_templates() -> dict | None:
        """Load customer_request_templates.json (customer request summary presets)."""
        # Prefer local file if present; otherwise fall back to advanced FA report folder.
        here = os.path.dirname(__file__)
        local_path = os.path.join(here, "customer_request_templates.json")
        adv_path = os.path.abspath(os.path.join(here, "..", "advanced FA report", "customer_request_templates.json"))
        path = local_path if os.path.isfile(local_path) else adv_path
        try:
            with open(path, "r", encoding="utf-8") as handle:
                return json.load(handle)
        except (OSError, json.JSONDecodeError) as exc:
            messagebox.showerror("Customer request", f"Could not load customer_request_templates.json:\n{exc}")
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
        customer_tpl = None

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
        rules = _load_report_rules()
        ctx = (rules or {}).get("report_context", {})
        scint_labels = ctx.get("scintillator_labels", {})
        scintillator_field = scint_labels.get(scint, scint) if scint else ""

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

        # Customer request templates override (from customer_request_templates.json)
        templates = _load_customer_request_templates()
        cust_req_type = (form.get("customer_request_type") or "").strip()
        label_to_key = {
            "Rework": "customer request rework",
            "Rework + CAMH": "customer request rework+camh",
            "DONE REWORK + CAMH": "customer request DONE REWORK+camh",
            "Rework (DONE REWORKED)": "customer request rework (DONE REWORKED)",
            "Rework ( <= 2 segment die)": "customer request rework(<= 2 segement die)",
            "Rework (HCTE Fail)": "customer request rework (HCTE Fail)",
        }
        if customer_req and templates and cust_req_type in label_to_key:
            customer_tpl = templates.get(label_to_key[cust_req_type])
            if customer_tpl:
                camera_main_issue = customer_tpl.get("camera_main_issue", camera_main_issue)
                camera_head = customer_tpl.get("camera_head", camera_head)
                camera_pcba = customer_tpl.get("camera_pcba", camera_pcba)
                scintillator_field = customer_tpl.get("scintillator", scintillator_field)
                root_cause_cat = customer_tpl.get("root_cause_categories", root_cause_cat)
                failure_root_cause = customer_tpl.get("failure_root_cause", failure_root_cause)
                disposition = customer_tpl.get("disposition", "Unable to Repair")
                disposition_rma = customer_tpl.get("disposition_rma", disposition)
                reason_to_scrap = customer_tpl.get("reason_to_scrap", "N/A")
                need_replacement = customer_tpl.get("need_replacement", "Yes")
                replaced_by = customer_tpl.get("replaced_by", "New Camera Head")
                countermeasure = customer_tpl.get("countermeasure", "N/A")

        # 7–8. Disposition + Disposition on RMA unitRequired (rows 219–223)
        ft_can_repair = form.get("ft_can_repair") or ""
        scrap_flag = pcba_can_repair == "scrap" or ft_can_repair == "scrap"
        unable_flag = (
            pcba_can_repair == "unable"
            or ft_can_repair == "unable"
            or ft_result in ("Pass with new camh", "Pass with swap camh")
        )

        if customer_tpl is None:
            # Follow Excel's VLOOKUP table order: Scrap, Unable to Repair, Repair, No problem found.
            # Special case: if Camera Main Issue is "No Problem Found", Disposition is "No problem found"
            # and Disposition on RMA unitRequired is "Not going to repair".
            if (camera_main_issue or "").strip() == "No Problem Found":
                disposition = "No problem found"
                disposition_rma = "Not going to repair"
            elif scrap_flag:
                disposition = "Scrap"
                disposition_rma = disposition
            elif unable_flag:
                disposition = "Unable to Repair"
                disposition_rma = disposition
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
        customer_req = str(form.get("customer_request", "")).lower() == "yes"
        customer_req_type = (form.get("customer_request_type") or "").strip()

        if customer_tpl is not None:
            # Use template action_taken as base, but tweak part number by camera model.
            base_action = (customer_tpl.get("action_taken") or "").strip()
            if base_action:
                override = ctx.get("action_taken", {}).get("part_number_override", {})
                if camera_model in override.get("camera_models", []):
                    base_action = base_action.replace(
                        override.get("from", "89504-0008 x1"),
                        override.get("to", "89504-0010 x1"),
                    )
            action_taken_str = base_action or "N/A"
        else:
            # Default action taken from ECC rework / FT logic (text from report_context)
            action_ctx = ctx.get("action_taken", {})
            ecc_texts = action_ctx.get("ecc_rework", {})
            ft_camh_texts = action_ctx.get("ft_camh", {})
            pcba_texts = action_ctx.get("pcba_replacement", {})

            b104 = ecc_texts.get(ecc_rework, "")
            b105 = ""
            if ft_lower == "pass with swap camh":
                b105 = ft_camh_texts.get("pass with swap camh", "")
            elif ft_lower == "pass with new camh" or e20_new_camh:
                b105 = ft_camh_texts.get("pass with new camh", "")

            b106 = ""
            if unable:
                b106 = pcba_texts.get(camera_model, "")

            action_parts = [p for p in (b104, b105, b106) if p]
            # E18 = PCBA ATS component cause, C20 = FT If fail, component change → component lines from 3.1G lookup
            e18 = (form.get("component_cause_ats") or "").strip()
            c20 = (form.get("ft_fail_component_change") or "").strip()
            if e18 or c20:
                ref_to_part = _load_31g_component_lookup()
                component_lines = _build_component_action_lines(e18, c20, ref_to_part)
                if component_lines:
                    action_parts.append(component_lines)
            action_taken_str = "\n".join(action_parts) if action_parts else "N/A"

        # 13. Countermeasure – from report_context
        cm_ctx = ctx.get("countermeasure", {})
        fa_camh_key_mapping = cm_ctx.get("fa_camh_key_mapping", {})
        countermeasure_texts = cm_ctx.get("texts", {})
        db_key = fa_camh_key_mapping.get(camh_fa, camh_fa)
        countermeasure = countermeasure_texts.get(db_key, "N/A")

        summary_labels = ctx.get("summary_field_labels", [
            "Camera Main Issue", "Camera - Head", "Camera - PCBA", "Scintillator",
            "Root Cause Categories", "Failure Analysis / Root Cause", "Disposition",
            "Disposition on RMA unitRequired", "Reason to Scrap", "Need Replacement",
            "Replaced By", "Action Taken on Repairing", "Countermeasure",
        ])
        summary_values = [
            camera_main_issue or "N/A",
            camera_head or "N/A",
            camera_pcba or "N/A",
            scintillator_field or "N/A",
            root_cause_cat or "N/A",
            _normalize_newlines(failure_root_cause or "N/A"),
            disposition,
            disposition_rma,
            _normalize_newlines(reason_to_scrap or "N/A"),
            need_replacement,
            replaced_by,
            _normalize_newlines(action_taken_str or "N/A"),
            _normalize_newlines(countermeasure or "N/A"),
        ]
        rows = list(zip(summary_labels, summary_values))
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
