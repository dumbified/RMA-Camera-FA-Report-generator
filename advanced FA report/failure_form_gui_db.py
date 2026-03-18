import csv
import json
import os
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from functools import lru_cache

from db_config_fa import get_fa_mysql_config


def _bool_to_str(value: bool) -> str:
    return "Yes" if value else "No"


def _ensure_directory(path: str) -> None:
    directory = os.path.dirname(path)
    if directory and not os.path.isdir(directory):
        os.makedirs(directory, exist_ok=True)


def _normalize_newlines(s: str) -> str:
    """Turn literal \\n in strings (e.g. from JSON/DB) into real newlines for display."""
    return (s or "").replace("\\n", "\n")


def _checkbox_row(parent, text: str, variable: tk.BooleanVar, row: int) -> ttk.Checkbutton:
    """Label on left, checkbox (small box) on right. Returns the Checkbutton for state control."""
    f = ttk.Frame(parent)
    f.grid(row=row, column=0, columnspan=2, sticky="w", pady=2)
    ttk.Label(f, text=text).pack(side="left")
    cb = ttk.Checkbutton(f, variable=variable)
    cb.pack(side="right", padx=(8, 0))
    return cb


@lru_cache(maxsize=None)
def _load_31g_component_from_db() -> dict[str, tuple[str, str, str]]:
    """
    Load 3.1G component lookup from MySQL (Component, partID, Description).
    Return mapping: component_ref_lower -> (part_id, desc, display_ref).
    """
    try:
        import mysql.connector  # type: ignore[import]
    except ImportError:
        return {}

    cfg = get_fa_mysql_config()
    try:
        conn = mysql.connector.connect(
            host=cfg.host,
            port=cfg.port,
            user=cfg.user,
            password=cfg.password,
            database=cfg.database,
        )
    except mysql.connector.Error:
        return {}

    ref_to_part: dict[str, tuple[str, str, str]] = {}
    try:
        cursor = conn.cursor()
        cursor.execute(f"SELECT Component, partID, Description FROM `{cfg.component_31g_table}`")
        for component, part_id, desc in cursor.fetchall():
            component = (component or "").strip()
            part_id = (part_id or "").strip()
            desc = (desc or "").strip()
            if not component or not part_id:
                continue
            key = component.lower()
            if key not in ref_to_part:
                ref_to_part[key] = (part_id, desc, component)
    except mysql.connector.Error:
        pass
    finally:
        conn.close()
    return ref_to_part


def _build_component_action_lines(
    e18: str, c20: str, ref_to_part: dict[str, tuple[str, str, str]]
) -> str:
    """
    Build Action Taken lines from E18 (PCBA ATS component cause) and C20 (FT If fail, component change).
    Format per line: {refs} {Description} {PartID}----->{count}
    """
    combined = f"{e18},{c20}".replace("\n", " ").replace("\r", " ")
    refs = [r.strip() for r in combined.split(",") if r.strip()]
    if not refs or not ref_to_part:
        return ""
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


def _fetch_screening_for_vit(vit_id: str) -> tuple[str, str, str] | None:
    """
    Fetch SCREENING PCBA, SCREENING CAMH, SCREENING SCINTILLATOR for a VIT from rma_cam.
    Maps DB values to form values:
      - PCBA bad -> "Cannot Ping"
      - CAMH bad -> "White Segment"
      - Scintillator bad -> "Aging"
    Returns (pcba_val, camh_val, scintillator_val) or None if not found.
    """
    vit_id = (vit_id or "").strip()
    if not vit_id:
        return None
    try:
        import mysql.connector  # type: ignore[import]
    except ImportError:
        return None
    cfg = get_fa_mysql_config()
    try:
        conn = mysql.connector.connect(
            host=cfg.host,
            port=cfg.port,
            user=cfg.user,
            password=cfg.password,
            database=cfg.database,
        )
    except mysql.connector.Error:
        return None
    col_sets = [
        ("`SCREENING PCBA`", "`SCREENING CAMH`", "`SCREENING SCINTILLATOR`"),
        ("SCREENING_PCBA", "SCREENING_CAMH", "SCREENING_SCINTILLATOR"),
        ("screening_pcba", "screening_camh", "screening_scintillator"),
    ]
    row = None
    try:
        cursor = conn.cursor()
        for pcba_col, camh_col, scint_col in col_sets:
            try:
                cursor.execute(
                    f"SELECT {pcba_col}, {camh_col}, {scint_col} "
                    f"FROM `{cfg.rma_cam_table}` WHERE VIT = %s LIMIT 1",
                    (vit_id,),
                )
                row = cursor.fetchone()
                break
            except mysql.connector.Error:
                continue
    except mysql.connector.Error:
        return None
    finally:
        conn.close()
    if not row:
        return None
    pcba_raw = (row[0] or "").strip().lower()
    camh_raw = (row[1] or "").strip().lower()
    scint_raw = (row[2] or "").strip().lower()

    PCBA_VALUES = ["Perfect", "Line between segment", "Cannot Ping"]
    CAMH_VALUES = ["Whole Segment", "Perfect", "<= 2 segment", "Line Between Segment", "White Segment", "Multiple Segment Die"]
    SCINT_VALUES = ["Good", "Aging", "Gap", "Bubble", "Bent"]

    def _map_pcba(v: str) -> str:
        if v in ("bad", "fail", "failed"):
            return "Cannot Ping"
        for opt in PCBA_VALUES:
            if opt.lower() == v or v in opt.lower():
                return opt
        return PCBA_VALUES[0]

    def _map_camh(v: str) -> str:
        if v in ("bad", "fail", "failed"):
            return "White Segment"
        for opt in CAMH_VALUES:
            if opt.lower() == v or v in opt.lower():
                return opt
        return CAMH_VALUES[1]

    def _map_scint(v: str) -> str:
        if v in ("bad", "fail", "failed"):
            return "Aging"
        for opt in SCINT_VALUES:
            if opt.lower() == v or v in opt.lower():
                return opt
        return SCINT_VALUES[0]

    return (_map_pcba(pcba_raw), _map_camh(camh_raw), _map_scint(scint_raw))


def _fetch_ft_result_for_vit(vit_id: str) -> str | None:
    """
    Fetch RESERVATIONS CAMH for a VIT from rma_cam and map it to FT Result value:

      - "Receive"       -> "Pass with new camh"
      - "Receive swap"  -> "Pass with swap camh"
      - "Repaired"      -> "Pass with repaired camh"
      - "N/A" (or empty)-> "Pass with ori camh"
    """
    vit_id = (vit_id or "").strip()
    if not vit_id:
        return None
    try:
        import mysql.connector  # type: ignore[import]
    except ImportError:
        return None

    cfg = get_fa_mysql_config()
    try:
        conn = mysql.connector.connect(
            host=cfg.host,
            port=cfg.port,
            user=cfg.user,
            password=cfg.password,
            database=cfg.database,
        )
    except mysql.connector.Error:
        return None

    col_candidates = [
        "`RESERVATIONS CAMH`",
        "RESERVATIONS_CAMH",
        "reservations_camh",
    ]
    row = None
    try:
        cursor = conn.cursor()
        for col in col_candidates:
            try:
                cursor.execute(
                    f"SELECT {col} FROM `{cfg.rma_cam_table}` WHERE VIT = %s LIMIT 1",
                    (vit_id,),
                )
                row = cursor.fetchone()
                break
            except mysql.connector.Error:
                continue
    except mysql.connector.Error:
        return None
    finally:
        conn.close()

    if not row or not row[0]:
        return None

    raw = str(row[0]).strip().lower()
    if not raw or raw in ("n/a", "na"):
        return "Pass with ori camh"
    if raw.startswith("receive swap"):
        return "Pass with swap camh"
    if raw.startswith("receive"):
        return "Pass with new camh"
    if raw.startswith("repaired"):
        return "Pass with repaired camh"

    return None


def _fetch_reservations_camh_raw(vit_id: str) -> str | None:
    """Fetch raw RESERVATIONS CAMH S/N for VIT from rma_cam (e.g. '87004', '8B811') for last-4 suffix. Returns None if not found."""
    vit_id = (vit_id or "").strip()
    if not vit_id:
        return None
    try:
        import mysql.connector  # type: ignore[import]
    except ImportError:
        return None
    cfg = get_fa_mysql_config()
    try:
        conn = mysql.connector.connect(
            host=cfg.host,
            port=cfg.port,
            user=cfg.user,
            password=cfg.password,
            database=cfg.database,
        )
    except mysql.connector.Error:
        return None
    col_candidates = [
        "`RESERVATIONS CAMH S/N`",
        "RESERVATIONS_CAMH_SN",
        "reservations_camh_sn",
    ]
    row = None
    try:
        cursor = conn.cursor()
        for col in col_candidates:
            try:
                cursor.execute(
                    f"SELECT {col} FROM `{cfg.rma_cam_table}` WHERE VIT = %s LIMIT 1",
                    (vit_id,),
                )
                row = cursor.fetchone()
                break
            except mysql.connector.Error:
                continue
    except mysql.connector.Error:
        return None
    finally:
        conn.close()
    if not row or row[0] is None:
        return None
    return str(row[0]).strip()


def _infer_camera_model_from_db(vit_id: str) -> str | None:
    """
    Infer camera model (3G, 3.1G Old, 3.1G New, 3.3G, 3.4G) from UNIT S/N and CAMH S/N in rma_cam.

    Rules:
      - UNIT S/N starts with 9504         -> 3G
      - UNIT S/N starts with 9704-007/009:
          CAMH S/N starts with 89504-0005 -> 3.1G Old
          CAMH S/N starts with 89404-0007 -> 3.1G New
      - UNIT S/N starts with 9704-011     -> 3.3G
      - UNIT S/N starts with 9704-013     -> 3.4G
    """
    vit_id = (vit_id or "").strip()
    if not vit_id:
        return None
    try:
        import mysql.connector  # type: ignore[import]
    except ImportError:
        return None

    cfg = get_fa_mysql_config()
    try:
        conn = mysql.connector.connect(
            host=cfg.host,
            port=cfg.port,
            user=cfg.user,
            password=cfg.password,
            database=cfg.database,
        )
    except mysql.connector.Error:
        return None

    unit_cols = ["`UNIT S/N`", "UNIT_SN", "unit_sn"]
    camh_cols = ["`CAMH S/N`", "CAMH_SN", "camh_sn"]
    row = None
    try:
        cursor = conn.cursor()
        for u_col in unit_cols:
            for c_col in camh_cols:
                try:
                    cursor.execute(
                        f"SELECT {u_col}, {c_col} FROM `{cfg.rma_cam_table}` WHERE VIT = %s LIMIT 1",
                        (vit_id,),
                    )
                    row = cursor.fetchone()
                    if row:
                        break
                except mysql.connector.Error:
                    continue
            if row:
                break
    except mysql.connector.Error:
        return None
    finally:
        conn.close()

    if not row or not row[0]:
        return None

    unit_sn = str(row[0]).strip()
    camh_sn = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""

    if unit_sn.startswith("9504"):
        return "3G"
    if unit_sn.startswith("9704-007") or unit_sn.startswith("9704-009"):
        if camh_sn.startswith("89504-0005"):
            return "3.1G Old"
        if camh_sn.startswith("89404-0007"):
            return "3.1G New"
        return "3.1G Old"
    if unit_sn.startswith("9704-011"):
        return "3.3G"
    if unit_sn.startswith("9704-013"):
        return "3.4G"
    return None


def _fetch_main_issue_fields_for_vit(vit_id: str) -> tuple[str, str, str] | None:
    """
    Fetch (Main Issue, Main Issue Category 1, Main Issue Category 2) for VIT from rma_cam.
    Returns tuple of strings, or None if not found.
    """
    vit_id = (vit_id or "").strip()
    if not vit_id:
        return None
    try:
        import mysql.connector  # type: ignore[import]
    except ImportError:
        return None

    cfg = get_fa_mysql_config()
    try:
        conn = mysql.connector.connect(
            host=cfg.host,
            port=cfg.port,
            user=cfg.user,
            password=cfg.password,
            database=cfg.database,
        )
    except mysql.connector.Error:
        return None

    main_issue_cols = ["`Main Issue`", "MAIN_ISSUE", "main_issue"]
    cat1_cols = ["`Main Issue Category 1`", "MAIN_ISSUE_CATEGORY_1", "main_issue_category_1"]
    cat2_cols = ["`Main Issue Category 2`", "MAIN_ISSUE_CATEGORY_2", "main_issue_category_2"]

    row = None
    try:
        cursor = conn.cursor()
        for mi in main_issue_cols:
            for c1 in cat1_cols:
                for c2 in cat2_cols:
                    try:
                        cursor.execute(
                            f"SELECT {mi}, {c1}, {c2} FROM `{cfg.rma_cam_table}` WHERE VIT = %s LIMIT 1",
                            (vit_id,),
                        )
                        row = cursor.fetchone()
                        if row:
                            break
                    except mysql.connector.Error:
                        continue
                if row:
                    break
            if row:
                break
    except mysql.connector.Error:
        return None
    finally:
        conn.close()

    if not row:
        return None

    mi = (row[0] or "").strip()
    c1 = (row[1] or "").strip()
    c2 = (row[2] or "").strip()
    return (mi, c1, c2)


def _map_main_issue_to_camh_screening(main_issue: str, cat1: str, cat2: str) -> str:
    """
    Map DB main-issue fields into CAMH Screening dropdown value.

    Rules:
      - If Main Issue is CAMH and Category 1 is CCD Segment Die -> use Category 2 mapped to CAMH values.
      - If they are NPF -> "Perfect"
      - If blank -> placeholder (currently empty string)
    """
    mi = (main_issue or "").strip().lower()
    c1n = (cat1 or "").strip().lower()
    c2n = (cat2 or "").strip()

    if not mi and not c1n and not c2n:
        return ""  # placeholder for now

    if mi == "npf" or c1n == "npf" or (cat2 or "").strip().lower() == "npf":
        return "Perfect"

    if mi == "camh" and c1n == "ccd segment die":
        v = (cat2 or "").strip().lower()
        # map common variants to the existing CAMH dropdown values
        mapping = {
            "white segment": "White Segment",
            "white": "White Segment",
            "whole segment": "Whole Segment",
            "whole": "Whole Segment",
            "line between segment": "Line Between Segment",
            "line between": "Line Between Segment",
            "line": "Line Between Segment",
            "<=2 segment": "<= 2 segment",
            "<= 2 segment": "<= 2 segment",
            "die segment": "Multiple Segment Die",
            "multiple segment die": "Multiple Segment Die",
            "multiple": "Multiple Segment Die",
        }
        for key, out in mapping.items():
            if v == key or key in v:
                return out
        return c2n  # last resort: put raw value (may not match dropdown; user can correct)

    return ""  # placeholder for other combinations for now


@lru_cache(maxsize=None)
def _load_camh_base_sn_from_db() -> dict[str, str]:
    """Load camera_group -> base_sn from fa_camh_base_sn. Fallback to defaults if table missing."""
    try:
        import mysql.connector  # type: ignore[import]
    except ImportError:
        return {"3G_31G": "89504-0008", "33G_34G": "89504-0010"}
    cfg = get_fa_mysql_config()
    try:
        conn = mysql.connector.connect(
            host=cfg.host,
            port=cfg.port,
            user=cfg.user,
            password=cfg.password,
            database=cfg.database,
        )
    except mysql.connector.Error:
        return {"3G_31G": "89504-0008", "33G_34G": "89504-0010"}
    result: dict[str, str] = {}
    try:
        cursor = conn.cursor()
        cursor.execute(f"SELECT camera_group, base_sn FROM `{cfg.camh_base_sn_table}`")
        for group, base_sn in cursor.fetchall():
            if group and base_sn:
                result[str(group).strip()] = str(base_sn).strip()
    except mysql.connector.Error:
        pass
    finally:
        conn.close()
    if not result:
        return {"3G_31G": "89504-0008", "33G_34G": "89504-0010"}
    return result


def _get_dynamic_camh_sn(camera_model: str, ft_result: str, vit_id: str) -> tuple[str, str]:
    """
    Return (sn_31g, sn_33g) for use in Action Taken.
    Base S/N loaded from fa_camh_base_sn (3G_31G, 33G_34G). If FT result is Pass with new camh or
    Pass with swap camh, append last 4 chars of RESERVATIONS CAMH S/N to each base.
    """
    bases = _load_camh_base_sn_from_db()
    base_31g = bases.get("3G_31G", "89504-0008")
    base_33g = bases.get("33G_34G", "89504-0010")
    ft_lower = (ft_result or "").strip().lower()
    vit_id = (vit_id or "").strip()
    if ft_lower in ("pass with new camh", "pass with swap camh") and vit_id:
        raw = _fetch_reservations_camh_raw(vit_id)
        if raw:
            suffix = raw[-4:] if len(raw) >= 4 else raw
            return (base_31g + suffix, base_33g + suffix)
    return (base_31g, base_33g)


def _replace_dynamic_sn_in_action_text(text: str, sn_31g: str, sn_33g: str) -> str:
    """Replace fixed 89504-0008 and 89504-0010 in action taken text with dynamic S/N."""
    if not text:
        return text
    return text.replace("89504-0008", sn_31g).replace("89504-0010", sn_33g)


@lru_cache(maxsize=None)
def _load_report_context_from_db() -> dict | None:
    """Load report_context from MySQL. Returns the context dict or None."""
    try:
        import mysql.connector  # type: ignore[import]
    except ImportError:
        return None

    cfg = get_fa_mysql_config()
    try:
        conn = mysql.connector.connect(
            host=cfg.host,
            port=cfg.port,
            user=cfg.user,
            password=cfg.password,
            database=cfg.database,
        )
    except mysql.connector.Error:
        return None

    try:
        cursor = conn.cursor()
        cursor.execute(
            f"SELECT context_data FROM `{cfg.report_data_table}` "
            f"WHERE item_type = 'context' AND context_key = 'report_context'"
        )
        row = cursor.fetchone()
    except mysql.connector.Error:
        return None
    finally:
        conn.close()

    if not row:
        return None
    data = row[0]
    if isinstance(data, str):
        return json.loads(data)
    return data


@lru_cache(maxsize=None)
def _load_customer_request_templates_from_db() -> dict | None:
    """Load customer request templates from MySQL. Returns dict of template_key -> template_data or None."""
    try:
        import mysql.connector  # type: ignore[import]
    except ImportError:
        return None

    cfg = get_fa_mysql_config()
    try:
        conn = mysql.connector.connect(
            host=cfg.host,
            port=cfg.port,
            user=cfg.user,
            password=cfg.password,
            database=cfg.database,
        )
    except mysql.connector.Error:
        return None

    templates: dict = {}
    try:
        cursor = conn.cursor()
        cursor.execute(
            f"SELECT template_key, template_data FROM `{cfg.customer_request_templates_table}`"
        )
        for tpl_key, tpl_data in cursor.fetchall():
            if isinstance(tpl_data, str):
                tpl_data = json.loads(tpl_data)
            templates[tpl_key or ""] = tpl_data or {}
    except mysql.connector.Error:
        return None
    finally:
        conn.close()
    return templates if templates else None


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


@lru_cache(maxsize=None)
def _load_report_rules_from_db():
    """
    Load FA report rules from MySQL and return the same structure as report_rules.json:

    {
      "order": [...],
      "sections": {
        "section_name": [
          {"id": "...", "when": {...}, "text": "..."},
          ...
        ],
        ...
      },
    }
    """
    try:
        import mysql.connector  # type: ignore[import]
    except ImportError as exc:  # pragma: no cover - environment-specific
        messagebox.showerror(
            "Report rules",
            "mysql-connector-python is not installed.\nRun: pip install mysql-connector-python",
        )
        return None

    cfg = get_fa_mysql_config()
    try:
        conn = mysql.connector.connect(
            host=cfg.host,
            port=cfg.port,
            user=cfg.user,
            password=cfg.password,
            database=cfg.database,
        )
    except mysql.connector.Error as exc:  # pragma: no cover - runtime connectivity
        messagebox.showerror("Report rules", f"Could not connect to MySQL:\n{exc}")
        return None

    sections_sql = f"""
        SELECT id, name
        FROM `{cfg.sections_table}`
        ORDER BY display_order ASC, id ASC
    """
    rules_sql = f"""
        SELECT section_id, rule_key, conditions, text, rule_order
        FROM `{cfg.report_data_table}`
        WHERE item_type = 'rule'
        ORDER BY section_id ASC, rule_order ASC, id ASC
    """

    try:
        cursor = conn.cursor()
        cursor.execute(sections_sql)
        sections_rows = cursor.fetchall()

        cursor.execute(rules_sql)
        rules_rows = cursor.fetchall()
    except mysql.connector.Error as exc:
        conn.close()
        messagebox.showerror("Report rules", f"Failed to query rules tables:\n{exc}")
        return None

    conn.close()

    # Build name -> id mapping and ordered section list
    order = []
    section_id_to_name = {}
    for section_id, name in sections_rows:
        section_id_to_name[section_id] = name
        order.append(name)

    sections: dict[str, list[dict]] = {name: [] for name in order}

    for section_id, rule_key, conditions_json, text, rule_order in rules_rows:
        name = section_id_to_name.get(section_id)
        if not name:
            continue
        try:
            # conditions_json is a JSON string or native JSON type depending on connector
            if isinstance(conditions_json, str):
                conditions = json.loads(conditions_json)
            else:
                conditions = conditions_json
        except Exception:
            conditions = {}
        sections.setdefault(name, []).append(
            {"id": rule_key, "when": conditions or {}, "text": text or ""}
        )

    return {"order": order, "sections": sections}


def build_failure_form_gui() -> None:
    root = tk.Tk()
    root.title("Camera Failure Analysis Form (Advanced / DB rules)")
    root.geometry("1150x820")

    main_paned = ttk.PanedWindow(root, orient="horizontal")
    main_paned.pack(fill="both", expand=True, padx=12, pady=12)

    # --- Left Panel: Batch Queue ---
    left_panel = ttk.Frame(main_paned, width=220)
    main_paned.add(left_panel, weight=0)

    ttk.Label(left_panel, text="Batch Queue", font=("", 10, "bold")).pack(anchor="w", pady=(0, 4))
    ttk.Label(left_panel, text="Paste VIT IDs (one per line):").pack(anchor="w")
    batch_input_text = tk.Text(left_panel, height=8, width=25)
    batch_input_text.pack(fill="x", pady=4)

    batch_vits = []
    batch_states = {}
    current_batch_idx = -1

    def add_to_queue():
        text = batch_input_text.get("1.0", tk.END).strip()
        if not text:
            return
        raw_vits = [v.strip() for v in text.split("\n") if v.strip()]
        for raw_v in raw_vits:
            v_upper = raw_v.upper()
            if v_upper.isdigit():
                v = f"VIT-{v_upper}"
            elif v_upper.startswith("VIT") and not v_upper.startswith("VIT-"):
                v = f"VIT-{v_upper[3:]}"
            elif v_upper.startswith("VIT-"):
                v = f"VIT-{v_upper[4:]}"
            else:
                v = raw_v
                
            if v not in batch_vits:
                batch_vits.append(v)
                batch_listbox.insert(tk.END, v)
        batch_input_text.delete("1.0", tk.END)

    def remove_from_queue():
        nonlocal current_batch_idx
        selection = batch_listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        vit_to_remove = batch_vits[idx]
        
        batch_listbox.delete(idx)
        batch_vits.pop(idx)
        
        if vit_to_remove in batch_states:
            del batch_states[vit_to_remove]
            
        if idx == current_batch_idx:
            current_batch_idx = -1
            clear_form()
        elif idx < current_batch_idx:
            current_batch_idx -= 1

    ttk.Button(left_panel, text="Add to Queue", command=add_to_queue).pack(fill="x", pady=(4, 2))
    ttk.Button(left_panel, text="Remove Selected", command=remove_from_queue).pack(fill="x", pady=(0, 4))

    batch_listbox = tk.Listbox(left_panel, height=15)
    batch_listbox.pack(fill="both", expand=True, pady=4)

    # We will define on_listbox_select and batch_export later in the function
    # after all variables are defined.

    # --- Right Panel: Form ---
    right_panel = ttk.Frame(main_paned)
    main_paned.add(right_panel, weight=1)

    canvas = tk.Canvas(right_panel, highlightthickness=0)
    scrollbar = ttk.Scrollbar(right_panel, orient="vertical", command=canvas.yview)
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

    # --- VIT ID [user types; Enter or FocusOut fetches screening from rma_cam] ---
    ttk.Label(form_frame, text="VIT ID").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=2)
    vit_id_entry = ttk.Entry(form_frame, textvariable=vit_id_var, width=20)
    vit_id_entry.grid(row=row, column=1, sticky="ew", pady=2)
    last_missing_main_issue_prompt_vit: str | None = None

    def _on_vit_confirm(_e=None, silent=False) -> None:
        nonlocal last_missing_main_issue_prompt_vit
        vit = (vit_id_var.get() or "").strip()
        if not vit:
            return
            
        # Auto-format VIT ID if user only typed numbers
        v_upper = vit.upper()
        if v_upper.isdigit():
            vit = f"VIT-{v_upper}"
        elif v_upper.startswith("VIT") and not v_upper.startswith("VIT-"):
            vit = f"VIT-{v_upper[3:]}"
        elif v_upper.startswith("VIT-"):
            vit = f"VIT-{v_upper[4:]}"
        vit_id_var.set(vit)
            
        main_issue_missing = False
        # Main Issue (from DB) -> CAMH screening (if applicable)
        mi_row = _fetch_main_issue_fields_for_vit(vit)
        if mi_row:
            mi, c1, c2 = mi_row
            # If the row exists but the main-issue fields are all blank, prompt the user once per VIT.
            if not (mi or c1 or c2):
                main_issue_missing = True
                # Keep CAMH screening blank when Main Issue is missing.
                camh_var.set("")
                if not silent and last_missing_main_issue_prompt_vit != vit:
                    last_missing_main_issue_prompt_vit = vit
                    messagebox.showinfo(
                        "Main Issue missing",
                        "CAMH Screening value is missing, please proceed to the classification system or manual identify the defect type.",
                    )
            mapped_camh = _map_main_issue_to_camh_screening(mi, c1, c2)
            if mapped_camh:
                camh_var.set(mapped_camh)

        screening = _fetch_screening_for_vit(vit)
        if screening:
            pcba_var.set(screening[0])
            # If Main Issue is missing, do NOT auto-fill CAMH (keep blank).
            if (not main_issue_missing) and (not (camh_var.get() or "").strip()):
                camh_var.set(screening[1])
            scintillator_var.set(screening[2])
        ft_result = _fetch_ft_result_for_vit(vit)
        if ft_result:
            ft_result_var.set(ft_result)
        inferred_model = _infer_camera_model_from_db(vit)
        if inferred_model:
            camera_model_var.set(inferred_model)
        refresh_interaction_states()

    vit_id_entry.bind("<Return>", _on_vit_confirm)
    vit_id_entry.bind("<FocusOut>", _on_vit_confirm)
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
        return {
            "vit_id": vit_id_var.get().strip(),
            "npf_all_reworked": _bool_to_str(npf_all_reworked_var.get()),
            "customer_request": _bool_to_str(customer_request_var.get()),
            "customer_request_type": customer_request_type_var.get(),
            "camera_model": camera_model_var.get(),
            "burnt": _bool_to_str(burnt_var.get()),
            # Visual check status used by rules
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
            # booleans kept as real True/False where JSON/DB rules expect them
            "npf_final": npf_final_var.get(),
            "camh_final": camh_final_var.get(),
            "pcba_final": pcba_final_var.get(),
            "scrap_why": scrap_why_var.get(),
            "bad_camh_assemble_bubble": _bool_to_str(bad_camh_assemble_bubble_var.get()),
            # Derived/alias fields used by rules
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
        vit_id_var.set("")
        npf_all_reworked_var.set(False)
        customer_request_var.set(False)
        customer_request_type_var.set("")
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

    def get_form_state() -> dict:
        return {
            "vit_id_var": vit_id_var.get(),
            "npf_all_reworked_var": npf_all_reworked_var.get(),
            "customer_request_var": customer_request_var.get(),
            "customer_request_type_var": customer_request_type_var.get(),
            "camera_model_var": camera_model_var.get(),
            "burnt_var": burnt_var.get(),
            "power_on_unit_var": power_on_unit_var.get(),
            "remark_missing_burnt_text": get_text(remark_missing_burnt_text),
            "pcba_var": pcba_var.get(),
            "camh_var": camh_var.get(),
            "scintillator_var": scintillator_var.get(),
            "ecc_rework_var": ecc_rework_var.get(),
            "good_camh_var": good_camh_var.get(),
            "bad_camh_var": bad_camh_var.get(),
            "which_dvm_fail_text": get_text(which_dvm_fail_text),
            "can_repair_bad_camh_var": can_repair_bad_camh_var.get(),
            "component_cause_text": get_text(component_cause_text),
            "pcba_ats_result_var": pcba_ats_result_var.get(),
            "ats_result_if_failed_text": get_text(ats_result_if_failed_text),
            "can_repair_ats_var": can_repair_ats_var.get(),
            "component_cause_ats_text": get_text(component_cause_ats_text),
            "component_category_var": component_category_var.get(),
            "ft_result_var": ft_result_var.get(),
            "ft_fail_component_change_text": get_text(ft_fail_component_change_text),
            "can_repair_ft_var": can_repair_ft_var.get(),
            "ft_pass_which_camh_var": ft_pass_which_camh_var.get(),
            "npf_final_var": npf_final_var.get(),
            "camh_final_var": camh_final_var.get(),
            "pcba_final_var": pcba_final_var.get(),
            "scrap_why_var": scrap_why_var.get(),
            "bad_camh_assemble_bubble_var": bad_camh_assemble_bubble_var.get(),
        }

    def set_form_state(state: dict) -> None:
        vit_id_var.set(state.get("vit_id_var", ""))
        npf_all_reworked_var.set(state.get("npf_all_reworked_var", False))
        customer_request_var.set(state.get("customer_request_var", False))
        customer_request_type_var.set(state.get("customer_request_type_var", ""))
        camera_model_var.set(state.get("camera_model_var", "3.3G"))
        burnt_var.set(state.get("burnt_var", False))
        power_on_unit_var.set(state.get("power_on_unit_var", ""))
        set_text(remark_missing_burnt_text, state.get("remark_missing_burnt_text", ""))
        pcba_var.set(state.get("pcba_var", ""))
        camh_var.set(state.get("camh_var", ""))
        scintillator_var.set(state.get("scintillator_var", ""))
        ecc_rework_var.set(state.get("ecc_rework_var", ""))
        good_camh_var.set(state.get("good_camh_var", ""))
        bad_camh_var.set(state.get("bad_camh_var", ""))
        set_text(which_dvm_fail_text, state.get("which_dvm_fail_text", ""))
        can_repair_bad_camh_var.set(state.get("can_repair_bad_camh_var", ""))
        set_text(component_cause_text, state.get("component_cause_text", ""))
        pcba_ats_result_var.set(state.get("pcba_ats_result_var", ""))
        set_text(ats_result_if_failed_text, state.get("ats_result_if_failed_text", ""))
        can_repair_ats_var.set(state.get("can_repair_ats_var", ""))
        set_text(component_cause_ats_text, state.get("component_cause_ats_text", ""))
        component_category_var.set(state.get("component_category_var", ""))
        ft_result_var.set(state.get("ft_result_var", ""))
        set_text(ft_fail_component_change_text, state.get("ft_fail_component_change_text", ""))
        can_repair_ft_var.set(state.get("can_repair_ft_var", ""))
        ft_pass_which_camh_var.set(state.get("ft_pass_which_camh_var", ""))
        npf_final_var.set(state.get("npf_final_var", False))
        camh_final_var.set(state.get("camh_final_var", ""))
        pcba_final_var.set(state.get("pcba_final_var", ""))
        scrap_why_var.set(state.get("scrap_why_var", ""))
        bad_camh_assemble_bubble_var.set(state.get("bad_camh_assemble_bubble_var", False))
        refresh_interaction_states()

    def on_listbox_select(event):
        nonlocal current_batch_idx
        selection = batch_listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        if idx == current_batch_idx:
            return
            
        # Save current state
        if current_batch_idx >= 0 and current_batch_idx < len(batch_vits):
            old_vit = batch_vits[current_batch_idx]
            batch_states[old_vit] = get_form_state()
            
        current_batch_idx = idx
        new_vit = batch_vits[current_batch_idx]
        
        # Load new state or fetch from DB
        if new_vit in batch_states:
            set_form_state(batch_states[new_vit])
        else:
            clear_form()
            vit_id_var.set(new_vit)
            _on_vit_confirm() # This triggers the DB fetch
            
    batch_listbox.bind("<<ListboxSelect>>", on_listbox_select)
    
    def batch_export():
        # Save current state
        if current_batch_idx >= 0 and current_batch_idx < len(batch_vits):
            old_vit = batch_vits[current_batch_idx]
            batch_states[old_vit] = get_form_state()

        if not batch_vits:
            messagebox.showinfo("Batch Export", "Queue is empty.")
            return

        success_count = 0
        all_rows_for_clipboard = []
        header_names = None

        try:
            for vit in batch_vits:
                # If user never clicked it, fetch it now
                if vit not in batch_states:
                    clear_form()
                    vit_id_var.set(vit)
                    _on_vit_confirm(silent=True)
                    batch_states[vit] = get_form_state()

                # Temporarily load the state to collect data properly
                set_form_state(batch_states[vit])
                data = collect_data()

                # Build report
                body_text = build_report_text(data)
                rows = build_summary_fields(data, body_text)
                
                # Insert VIT ID as the first column
                rows.insert(0, ("VIT ID", vit))

                if header_names is None:
                    header_names = [r[0] for r in rows]

                row_values = [r[1] for r in rows]

                # Prepare for clipboard
                values_for_clipboard = [(val or "").strip().replace("\n", " ") for val in row_values]
                all_rows_for_clipboard.append("\t".join(values_for_clipboard))

                success_count += 1

            # Restore the currently selected form state
            if current_batch_idx >= 0 and current_batch_idx < len(batch_vits):
                set_form_state(batch_states[batch_vits[current_batch_idx]])
            else:
                clear_form()

            # Copy to clipboard
            if header_names and all_rows_for_clipboard:
                tsv = "\t".join(header_names) + "\n" + "\n".join(all_rows_for_clipboard)
                root.clipboard_clear()
                root.clipboard_append(tsv)

            messagebox.showinfo("Batch Export", f"Successfully generated {success_count} reports.\n\nThe combined report data has been copied to your clipboard.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export: {e}")

    ttk.Button(left_panel, text="Batch Export (Copy to Clipboard)", command=batch_export).pack(fill="x", pady=4)

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
        rules = _load_report_rules_from_db()
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

    def build_summary_fields(form: dict, failure_body: str) -> list[tuple[str, str]]:
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
        ctx = _load_report_context_from_db() or {}

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
        scint_labels = ctx.get("scintillator_labels", {}) or {
            "Good": "No Problem Found",
            "Aging": "Aging",
            "Gap": "Gap",
            "Bubble": "Bubble",
            "Bent": "Bent",
        }
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

        # Customer request templates override (from DB)
        templates = _load_customer_request_templates_from_db()
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

        disposition = "Repair"
        disposition_rma = disposition
        reason_to_scrap = "N/A"
        need_replacement = "No"
        replaced_by = "N/A"
        action_taken_str = "N/A"
        countermeasure = "N/A"

        if customer_tpl is None:
            # Follow Excel's VLOOKUP table order: Scrap, Unable to Repair, Repair, No problem found.
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
        if customer_tpl is None:
            if disposition == "Scrap" and failure_root_cause:
                text = failure_root_cause
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

            # 10–11. Need Replacement + Replaced By
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

        # 12. Action Taken on Repairing (dynamic S/N by camera model; optional suffix from RESERVATIONS CAMH S/N)
        ecc_rework = (form.get("ecc_rework") or "").strip()
        camera_model = (form.get("camera_model") or "").strip()
        ft_lower = ft_result.strip().lower() if ft_result else ""
        unable = pcba_can_repair == "unable" or ft_can_repair == "unable"
        e20_new_camh = (form.get("ft_pass_which_camh") or "").strip().lower() == "new camh"
        vit_id = (form.get("vit_id") or "").strip()
        sn_31g, sn_33g = _get_dynamic_camh_sn(camera_model, ft_result or "", vit_id)

        if customer_tpl is not None:
            base_action = (customer_tpl.get("action_taken") or "").strip()
            if base_action:
                base_action = _replace_dynamic_sn_in_action_text(base_action, sn_31g, sn_33g)
            action_taken_str = base_action or "N/A"
        else:
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
            e18 = (form.get("component_cause_ats") or "").strip()
            c20 = (form.get("ft_fail_component_change") or "").strip()
            if e18 or c20:
                ref_to_part = _load_31g_component_from_db()
                component_lines = _build_component_action_lines(e18, c20, ref_to_part)
                if component_lines:
                    action_parts.append(component_lines)
            action_taken_str = "\n".join(action_parts) if action_parts else "N/A"
            action_taken_str = _replace_dynamic_sn_in_action_text(action_taken_str, sn_31g, sn_33g)

        # 13. Countermeasure – from report_context (DB)
        if customer_tpl is None:
            cm_ctx = ctx.get("countermeasure", {})
            fa_camh_key_mapping = cm_ctx.get("fa_camh_key_mapping") or {
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
        win.title("Generated JIRA Report (Advanced / DB rules)")
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

