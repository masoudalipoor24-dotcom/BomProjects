from __future__ import annotations

import argparse
from pathlib import Path
from typing import Final

import pyodbc


KNOWN_SKIP_SNIPPETS: Final[tuple[str, ...]] = (
    "already exists",
    "already has an index named",
    "there is already a relationship named",
)


def _is_expected_already_exists(exc: Exception) -> bool:
    msg = str(exc).lower()
    return any(snippet in msg for snippet in KNOWN_SKIP_SNIPPETS)


def _is_cascade_constraint_syntax_error(stmt: str, exc: Exception) -> bool:
    msg = str(exc).lower()
    return "on delete cascade" in stmt.lower() and "syntax error in constraint clause" in msg


def _remove_on_delete_cascade(stmt: str) -> str:
    return stmt.replace("ON DELETE CASCADE", "").replace("on delete cascade", "")


def execute_schema(conn: pyodbc.Connection, schema_sql_path: Path) -> tuple[int, int, list[tuple[int, str]]]:
    sql_text = schema_sql_path.read_text(encoding="utf-8")
    statements = [stmt.strip() for stmt in sql_text.split(";") if stmt.strip()]

    cur = conn.cursor()
    ok = 0
    skipped = 0
    failures: list[tuple[int, str]] = []

    for idx, stmt in enumerate(statements, start=1):
        try:
            cur.execute(stmt)
            ok += 1
        except Exception as exc:  # noqa: BLE001
            if _is_expected_already_exists(exc):
                skipped += 1
                continue

            if _is_cascade_constraint_syntax_error(stmt, exc):
                fallback_stmt = _remove_on_delete_cascade(stmt).strip()
                try:
                    cur.execute(fallback_stmt)
                    ok += 1
                    continue
                except Exception as fallback_exc:  # noqa: BLE001
                    if _is_expected_already_exists(fallback_exc):
                        skipped += 1
                        continue
                    failures.append((idx, str(fallback_exc)))
                    continue

            failures.append((idx, str(exc)))

    return ok, skipped, failures


def seed_lookups(conn: pyodbc.Connection) -> None:
    cur = conn.cursor()

    seed_sql = [
        "INSERT INTO tblItemType (TypeCode, TypeName) VALUES ('FG','Finished Good')",
        "INSERT INTO tblItemType (TypeCode, TypeName) VALUES ('SA','Sub Assembly')",
        "INSERT INTO tblItemType (TypeCode, TypeName) VALUES ('RM','Raw Material')",
        "INSERT INTO tblItemType (TypeCode, TypeName) VALUES ('PKG','Packaging')",
        "INSERT INTO tblUOM (UOMCode, UOMName, IsActive) VALUES ('PCS','Piece',True)",
        "INSERT INTO tblUOM (UOMCode, UOMName, IsActive) VALUES ('KG','Kilogram',True)",
        "INSERT INTO tblUOM (UOMCode, UOMName, IsActive) VALUES ('M','Meter',True)",
    ]

    for stmt in seed_sql:
        try:
            cur.execute(stmt)
        except Exception:
            # Ignore duplicate insert failures on re-run.
            pass


def list_user_tables(conn: pyodbc.Connection) -> list[str]:
    cur = conn.cursor()
    names = []
    for row in cur.tables(tableType="TABLE"):
        name = row.table_name
        if not name.startswith("MSys"):
            names.append(name)
    return sorted(set(names))


def select_access_driver(explicit_driver: str | None = None) -> str:
    if explicit_driver:
        return explicit_driver

    installed = pyodbc.drivers()
    preferred = [
        "Microsoft Access Driver (*.mdb, *.accdb)",
        "Microsoft Access Driver (*.mdb)",
    ]

    for name in preferred:
        if name in installed:
            return name

    for name in installed:
        low = name.lower()
        if "access" in low and ("accdb" in low or "mdb" in low):
            return name

    # Some environments fail to enumerate drivers even when they exist.
    # Return a best-effort default and let connect() validate.
    return "Microsoft Access Driver (*.mdb, *.accdb)"


def main() -> int:
    parser = argparse.ArgumentParser(description="Populate an ACCDB with BOM schema and seed data.")
    parser.add_argument("--db", required=True, type=Path, help="Path to .accdb file")
    parser.add_argument("--schema", required=True, type=Path, help="Path to schema SQL file")
    parser.add_argument("--driver", default=None, help="Optional ODBC driver name override")
    args = parser.parse_args()

    db_path = args.db.resolve()
    schema_path = args.schema.resolve()

    if not db_path.exists():
        raise FileNotFoundError(f"DB file not found: {db_path}")
    if not schema_path.exists():
        raise FileNotFoundError(f"Schema file not found: {schema_path}")

    driver_candidates = []
    if args.driver:
        driver_candidates.append(args.driver)
    else:
        primary = select_access_driver(None)
        driver_candidates.extend(
            [
                primary,
                "Microsoft Access Driver (*.mdb, *.accdb)",
                "Microsoft Access Driver (*.mdb)",
                "Driver do Microsoft Access (*.mdb)",
            ]
        )

    conn = None
    used_driver = None
    last_exc: Exception | None = None
    for driver in dict.fromkeys(driver_candidates):
        conn_str = f"DRIVER={{{driver}}};DBQ={db_path};"
        try:
            conn = pyodbc.connect(conn_str, autocommit=True)
            used_driver = driver
            break
        except Exception as exc:  # noqa: BLE001
            last_exc = exc

    if conn is None or used_driver is None:
        raise RuntimeError(
            "Could not connect to ACCDB with available Access ODBC drivers. "
            f"Last error: {last_exc}"
        )

    try:
        ok_count, skip_count, failures = execute_schema(conn, schema_path)
        seed_lookups(conn)
        tables = list_user_tables(conn)
    finally:
        conn.close()

    print(f"DRIVER={used_driver}")
    print(f"SCHEMA_OK={ok_count}")
    print(f"SCHEMA_SKIP={skip_count}")
    print(f"SCHEMA_FAIL={len(failures)}")
    for idx, msg in failures:
        print(f"FAIL[{idx}] {msg}")
    print("TABLES=" + ",".join(tables))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
