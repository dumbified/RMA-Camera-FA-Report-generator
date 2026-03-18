"""
Upload review results CSV (VIT, filename, category) to MySQL.

Set the constants below to your MySQL connection. Install: pip install mysql-connector-python
"""

import csv
import sys
from pathlib import Path

# --- MySQL config (edit these) ---
MYSQL_HOST = "localhost"
MYSQL_PORT = 3306
MYSQL_USER = "root"
MYSQL_PASSWORD = "admin123"
MYSQL_DATABASE = "fa_report"
MYSQL_TABLE = "review_results"


def get_mysql_config():
    """Return MySQL config from module constants."""
    return {
        "host": MYSQL_HOST,
        "port": MYSQL_PORT,
        "user": MYSQL_USER,
        "password": MYSQL_PASSWORD,
        "database": MYSQL_DATABASE,
        "table": MYSQL_TABLE,
    }


def upload_csv_to_mysql(csv_path, table_name=None, if_exists="append"):
    """
    Upload a review-results CSV to MySQL.

    CSV must have header: VIT,filename,category

    Args:
        csv_path: Path to the CSV file.
        table_name: MySQL table name (default: MYSQL_TABLE constant).
        if_exists: 'append' (default) or 'replace'. Replace drops and recreates the table.

    Returns:
        (success: bool, message: str)
    """
    try:
        import mysql.connector
    except ImportError:
        return False, "mysql-connector-python is not installed. Run: pip install mysql-connector-python"

    csv_path = Path(csv_path)
    if not csv_path.is_file():
        return False, f"CSV file not found: {csv_path}"

    cfg = get_mysql_config()
    if not cfg["user"] or not cfg["database"]:
        return False, "Set MYSQL_USER and MYSQL_DATABASE (and MYSQL_PASSWORD if needed) in csv_to_mysql.py."

    table = table_name or cfg["table"]

    try:
        conn = mysql.connector.connect(
            host=cfg["host"],
            port=cfg["port"],
            user=cfg["user"],
            password=cfg["password"],
            database=cfg["database"],
        )
    except mysql.connector.Error as e:
        return False, f"MySQL connection failed: {e}"

    cursor = conn.cursor()

    try:
        with open(csv_path, newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader, None)
            if not header or [h.strip() for h in header] != ["VIT", "filename", "category"]:
                return False, "CSV must have header: VIT,filename,category"
            rows = list(reader)

        if not rows:
            conn.close()
            return True, "CSV has no data rows; nothing to upload."

        if if_exists == "replace":
            cursor.execute(f"DROP TABLE IF EXISTS `{table}`")
        cursor.execute(
            f"""
            CREATE TABLE IF NOT EXISTS `{table}` (
                vit VARCHAR(255) NOT NULL,
                filename VARCHAR(512) NOT NULL,
                category VARCHAR(255) NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (vit, filename)
            )
            """
        )

        insert_sql = (
            f"INSERT INTO `{table}` (vit, filename, category) VALUES (%s, %s, %s) "
            f"ON DUPLICATE KEY UPDATE category = VALUES(category)"
        )

        # Optional: also backfill rma_cam with Main Issue fields for CAMH segmentation
        # Only writes when Main Issue, Main Issue Category 1, and Main Issue Category 2 are all blank.
        try:
            import mysql.connector  # type: ignore[import]
        except ImportError:
            mysql = None  # type: ignore[assignment]
        else:
            mysql = mysql.connector  # type: ignore[assignment]

        update_sql = None
        if mysql is not None:
            update_sql = """
                UPDATE `rma_cam`
                SET
                    `Main Issue` = %s,
                    `Main Issue Category 1` = %s,
                    `Main Issue Category 2` = %s
                WHERE VIT = %s
                  AND (`Main Issue` IS NULL OR `Main Issue` = '')
                  AND (`Main Issue Category 1` IS NULL OR `Main Issue Category 1` = '')
                  AND (`Main Issue Category 2` IS NULL OR `Main Issue Category 2` = '')
            """

        for row in rows:
            if len(row) >= 3:
                vit = row[0].strip()
                filename = row[1].strip()
                category = row[2].strip()
                cursor.execute(insert_sql, (vit, filename, category))

                # Backfill rma_cam only when connector is available and update_sql prepared
                if update_sql is not None and vit and category:
                    try:
                        cursor.execute(
                            update_sql,
                            ("CAMH", "CCD Segment Die", category, vit),
                        )
                    except Exception:
                        # Silently ignore if rma_cam or columns do not exist;
                        # review_results upload should still succeed.
                        pass

        conn.commit()
        count = len([r for r in rows if len(r) >= 3])
        conn.close()
        return True, f"Uploaded {count} row(s) to MySQL table `{table}`."
    except Exception as e:
        conn.rollback()
        conn.close()
        return False, str(e)


def fetch_review_results(table_name=None):
    """
    Fetch all rows from the review_results table.

    Returns:
        (success: bool, data: list | str)
        On success: data is a list of (vit, filename, category, created_at) tuples.
        On failure: data is the error message string.
    """
    try:
        import mysql.connector
    except ImportError:
        return False, "mysql-connector-python is not installed. Run: pip install mysql-connector-python"

    cfg = get_mysql_config()
    if not cfg["user"] or not cfg["database"]:
        return False, "Set MYSQL_USER and MYSQL_DATABASE in csv_to_mysql.py."

    table = table_name or cfg["table"]
    try:
        conn = mysql.connector.connect(
            host=cfg["host"],
            port=cfg["port"],
            user=cfg["user"],
            password=cfg["password"],
            database=cfg["database"],
        )
    except mysql.connector.Error as e:
        return False, str(e)

    cursor = conn.cursor()
    try:
        cursor.execute(
            f"SELECT vit, filename, category, created_at FROM `{table}` ORDER BY created_at DESC"
        )
        rows = cursor.fetchall()
        conn.close()
        return True, rows
    except mysql.connector.Error as e:
        conn.close()
        return False, str(e)


def main():
    if len(sys.argv) < 2:
        print("Usage: python csv_to_mysql.py <path/to/review_results.csv> [--replace]", file=sys.stderr)
        sys.exit(1)
    csv_path = sys.argv[1]
    if_exists = "replace" if "--replace" in sys.argv else "append"
    ok, msg = upload_csv_to_mysql(csv_path, if_exists=if_exists)
    if ok:
        print(msg)
    else:
        print(f"Error: {msg}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
