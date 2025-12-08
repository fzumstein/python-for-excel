"""This module handles all database interactions

Note that SQLite doesn't enforce foreign keys by default. You can fix this via
SQLAlchemy as described here:
https://docs.sqlalchemy.org/en/20/dialects/sqlite.html#foreign-key-support
"""

from pathlib import Path

import pandas as pd
import sqlalchemy
from sqlalchemy import text

# We want the database file to sit next to this file.
# Here, we are turning the path into an absolute path.
this_dir = Path(__file__).resolve().parent
db_path = this_dir / "packagetracker.db"

# Database engine
engine = sqlalchemy.create_engine(f"sqlite:///{db_path}")


def get_packages() -> pd.DataFrame:
    """Get all packages as DataFrame"""
    return pd.read_sql_table("packages", con=engine, index_col="package_id")


def package_exists(package_name: str) -> bool:
    """Check if a package already exists in the database"""
    statement = text(
        "SELECT COUNT(*) as count FROM packages WHERE package_name = :package_name"
    )
    with engine.begin() as con:
        result = con.execute(statement, {"package_name": package_name})
        count = result.scalar()
    return count > 0


def store_package(package_name: str) -> None:
    """Insert a new package into the packages table"""
    statement = text("INSERT INTO packages (package_name) VALUES (:package_name)")
    with engine.begin() as con:
        con.execute(statement, {"package_name": package_name})


def get_versions(package_name: str) -> pd.DataFrame:
    """Get all versions for the package with the name package_name"""
    statement = text("""
    SELECT v.uploaded_at, v.version
    FROM packages p
    INNER JOIN package_versions v ON p.package_id = v.package_id
    WHERE p.package_name = :package_name
    ORDER BY v.uploaded_at
    """)
    return pd.read_sql_query(
        statement,
        engine,
        parse_dates=["uploaded_at"],
        params={"package_name": package_name},
        index_col=["uploaded_at"],
    )


def store_versions(df: pd.DataFrame) -> None:
    """Insert the records of the provided DataFrame into the package_versions table"""
    df.to_sql("package_versions", con=engine, if_exists="append", index=False)


def delete_versions() -> None:
    """Delete all records from the version table"""
    statement = text("DELETE FROM package_versions")
    with engine.begin() as con:
        con.execute(statement)


def create_db() -> None:
    """Run this function to create the database tables.
    In case of SQLite, this is also creating the database file.
    """
    statement_packages = text("""
    CREATE TABLE packages (
        package_id INTEGER PRIMARY KEY,
        package_name TEXT NOT NULL,
        UNIQUE(package_name)
    )
    """)

    statement_versions = text("""
    CREATE TABLE package_versions (
        package_id INTEGER,
        version TEXT,
        uploaded_at TIMESTAMP NOT NULL,
        PRIMARY KEY (package_id, version),
        FOREIGN KEY (package_id) REFERENCES packages (package_id)
    )
    """)

    with engine.begin() as con:
        for statement in [statement_packages, statement_versions]:
            con.execute(statement)

    print("Database created successfully!")


if __name__ == "__main__":
    create_db()
