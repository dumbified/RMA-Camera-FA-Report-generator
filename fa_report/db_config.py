"""
MySQL configuration for the advanced FA report rule engine.

This keeps FA-report specific DB settings in one place, similar to csv_to_mysql.get_mysql_config().
Adjust the constants below to match your MySQL setup.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class FaMySQLConfig:
    host: str
    port: int
    user: str
    password: str
    database: str
    sections_table: str
    report_data_table: str  # Unified: rules + report_context
    customer_request_templates_table: str
    component_31g_table: str
    rma_cam_table: str  # VIT IDs and screening results
    camh_base_sn_table: str  # Camera group -> base S/N for Action Taken


def get_fa_mysql_config() -> FaMySQLConfig:
    """
    Return MySQL configuration for FA report rules.

    By default this reuses the same database as csv_to_mysql (\"upload_csv\"),
    but with distinct tables for FA rules.
    """
    return FaMySQLConfig(
        host="localhost",
        port=3306,
        user="root",
        password="admin123",
        database="fa_report",
        sections_table="fa_report_sections",
        report_data_table="fa_report_data",
        customer_request_templates_table="fa_customer_request_templates",
        component_31g_table="fa_3.1_cam_component",
        rma_cam_table="rma_cam",
        camh_base_sn_table="fa_camh_base_sn",
    )

