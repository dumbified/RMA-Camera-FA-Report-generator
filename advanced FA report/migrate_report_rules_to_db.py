"""
Migration script: load FA report rules and related data from JSON into MySQL tables.

This reads:
  - FA report/report_rules.json -> fa_report_sections, fa_report_data (rules + context)
  - FA report/customer_request_templates.json -> fa_customer_request_templates
  - fa_camh_base_sn: base S/N by camera group (3G_31G -> 89504-0008, 33G_34G -> 89504-0010).

3.1G component data is NOT migrated; it is fetched directly from fa_3.1_cam_component in SQL Workbench.

Tables are (re)created on each run to stay in sync with the source files.
"""

from __future__ import annotations

import json
import os
from typing import Any, Dict, List

from db_config_fa import get_fa_mysql_config


def _load_json_rules() -> Dict[str, Any]:
    """Load report_rules.json from the advanced FA report folder."""
    here = os.path.dirname(__file__)
    json_path = os.path.abspath(os.path.join(here, "report_rules.json"))
    with open(json_path, "r", encoding="utf-8") as handle:
        return json.load(handle)


def _connect_mysql():
    """Return (conn, cursor) using FA MySQL configuration."""
    try:
        import mysql.connector  # type: ignore[import]
    except ImportError as exc:  # pragma: no cover - runtime environment detail
        raise SystemExit(
            "mysql-connector-python is not installed. Run: pip install mysql-connector-python"
        ) from exc

    cfg = get_fa_mysql_config()
    conn = mysql.connector.connect(
        host=cfg.host,
        port=cfg.port,
        user=cfg.user,
        password=cfg.password,
        database=cfg.database,
    )
    return conn, conn.cursor()


def _load_customer_request_templates() -> Dict[str, Any]:
    """Load customer_request_templates.json from the FA report folder."""
    here = os.path.dirname(__file__)
    path = os.path.join(here, "..", "FA report", "customer_request_templates.json")
    path = os.path.abspath(path)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _create_tables(cursor) -> None:
    """Create (or recreate) FA report tables."""
    cfg = get_fa_mysql_config()
    sections = cfg.sections_table
    data_tbl = cfg.report_data_table
    tpl_tbl = cfg.customer_request_templates_table
    camh_sn_tbl = cfg.camh_base_sn_table

    # Drop in order: tables with FKs first, then referenced tables
    cursor.execute(f"DROP TABLE IF EXISTS `{data_tbl}`")
    cursor.execute("DROP TABLE IF EXISTS `fa_report_rules`")  # legacy
    cursor.execute("DROP TABLE IF EXISTS `fa_report_context`")  # legacy
    cursor.execute(f"DROP TABLE IF EXISTS `{sections}`")
    cursor.execute(f"DROP TABLE IF EXISTS `{tpl_tbl}`")
    cursor.execute(f"DROP TABLE IF EXISTS `{camh_sn_tbl}`")

    cursor.execute(
        f"""
        CREATE TABLE `{sections}` (
            id INT AUTO_INCREMENT PRIMARY KEY,
            name VARCHAR(64) NOT NULL UNIQUE,
            display_order INT NOT NULL
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """
    )

    cursor.execute(
        f"""
        CREATE TABLE `{data_tbl}` (
            id INT AUTO_INCREMENT PRIMARY KEY,
            item_type VARCHAR(20) NOT NULL,
            section_id INT NULL,
            rule_key VARCHAR(128) NULL,
            conditions JSON NULL,
            text TEXT NULL,
            rule_order INT NULL,
            context_key VARCHAR(64) NULL,
            context_data JSON NULL,
            CONSTRAINT fk_{data_tbl}_section
                FOREIGN KEY (section_id) REFERENCES `{sections}` (id)
                ON DELETE CASCADE
            ,
            UNIQUE KEY uniq_item_context (item_type, context_key)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """
    )

    cursor.execute(
        f"""
        CREATE TABLE `{tpl_tbl}` (
            id INT AUTO_INCREMENT PRIMARY KEY,
            template_key VARCHAR(128) NOT NULL UNIQUE,
            template_data JSON NOT NULL
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """
    )

    cursor.execute(
        f"""
        CREATE TABLE `{camh_sn_tbl}` (
            camera_group VARCHAR(32) NOT NULL PRIMARY KEY,
            base_sn VARCHAR(64) NOT NULL
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """
    )


def migrate() -> None:
    """Main migration routine."""
    rules_json = _load_json_rules()
    order: List[str] = list(rules_json.get("order", []))
    sections_data: Dict[str, List[Dict[str, Any]]] = rules_json.get("sections", {})  # type: ignore[assignment]

    conn, cursor = _connect_mysql()

    try:
        _create_tables(cursor)

        cfg = get_fa_mysql_config()
        sections_table = cfg.sections_table
        data_table = cfg.report_data_table

        # Insert sections with display_order based on the JSON "order" list.
        section_ids: Dict[str, int] = {}
        for idx, name in enumerate(order):
            cursor.execute(
                f"INSERT INTO `{sections_table}` (name, display_order) VALUES (%s, %s)",
                (name, idx),
            )
            section_ids[name] = cursor.lastrowid

        # Insert rules into fa_report_data (item_type='rule')
        for section_name in order:
            section_rules = sections_data.get(section_name, [])
            section_id = section_ids[section_name]
            for rule_index, rule in enumerate(section_rules):
                rule_key = rule.get("id", f"{section_name}_{rule_index}")
                conditions = rule.get("when", {})
                text = rule.get("text", "")
                cursor.execute(
                    f"""
                    INSERT INTO `{data_table}`
                        (item_type, section_id, rule_key, conditions, text, rule_order)
                    VALUES ('rule', %s, %s, %s, %s, %s)
                    """,
                    (
                        section_id,
                        rule_key,
                        json.dumps(conditions),
                        text,
                        rule_index,
                    ),
                )

        # Insert report_context into fa_report_data (item_type='context')
        # Split into smaller rows for DB cleanliness:
        #   report_context.<top_level_key> -> each top-level key's JSON value
        report_ctx = rules_json.get("report_context", {})
        if report_ctx:
            for top_key, top_value in report_ctx.items():
                cursor.execute(
                    f"""
                    INSERT INTO `{data_table}`
                        (item_type, context_key, context_data)
                    VALUES ('context', %s, %s)
                    """,
                    (f"report_context.{top_key}", json.dumps(top_value)),
                )

        # Migrate customer_request_templates
        try:
            templates = _load_customer_request_templates()
            for tpl_key, tpl_data in templates.items():
                cursor.execute(
                    f"""
                    INSERT INTO `{cfg.customer_request_templates_table}` (template_key, template_data)
                    VALUES (%s, %s)
                    """,
                    (tpl_key, json.dumps(tpl_data)),
                )
        except (OSError, json.JSONDecodeError) as e:
            print(f"Warning: Could not load customer_request_templates.json: {e}")

        # Base S/N by camera group (used in Action Taken; append last 4 from RESERVATIONS CAMH when applicable)
        cursor.execute(
            f"INSERT INTO `{cfg.camh_base_sn_table}` (camera_group, base_sn) VALUES (%s, %s)",
            ("3G_31G", "89504-0008"),
        )
        cursor.execute(
            f"INSERT INTO `{cfg.camh_base_sn_table}` (camera_group, base_sn) VALUES (%s, %s)",
            ("33G_34G", "89504-0010"),
        )

        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    migrate()
    print("FA report rules, report context, customer request templates, and fa_camh_base_sn migrated to MySQL.")

