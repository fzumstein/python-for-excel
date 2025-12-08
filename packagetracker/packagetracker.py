"""This module contains all functions that are either called from Excel
or manipulate Excel.
"""

import datetime as dt
from typing import Any

import pandas as pd
import requests
import xlwings as xw
from matplotlib.figure import Figure

import database

BASE_URL = "https://pypi.org/pypi"
REQUEST_TIMEOUT = 10


def fetch_package_data(package_name: str) -> dict[str, Any]:
    """Fetch package data from PyPI

    By default, Requests is waiting forever for a response, which could lead the
    application to "hang" in cases where PyPI has an issue and is responding slowly.
    That's why you should always include an explicit timeout parameter.
    """
    response = requests.get(f"{BASE_URL}/{package_name}/json", timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    return response.json()


def extract_releases(data: dict[str, Any], package_id: int) -> pd.DataFrame:
    """Extract release information from PyPI data"""
    releases_data = []
    for version, files in data["releases"].items():
        if files:
            releases_data.append(
                {
                    "uploaded_at": files[0]["upload_time"],
                    "version": version,
                    "package_id": package_id,
                }
            )

    if not releases_data:
        return pd.DataFrame()

    releases = pd.DataFrame(releases_data)
    releases["uploaded_at"] = pd.to_datetime(releases["uploaded_at"])
    releases = releases.sort_values("uploaded_at")
    return releases


def add_package() -> None:
    """Adds a new package including the version history to the database.
    Triggers an update of the dropdown on the Tracker tab.
    """
    # Excel objects
    db_sheet = xw.Book.caller().sheets["Database"]
    feedback_cell = db_sheet["new_package"].offset(column_offset=1)

    # Clear feedback cell
    feedback_cell.clear_contents()

    # Require an input
    package_name = db_sheet["new_package"].value
    if not package_name:
        feedback_cell.value = "ERROR: Please provide a name!"
        return

    # Clean up package name by removing any extra spaces
    package_name = package_name.strip()

    # Check if package already exists in database
    if database.package_exists(package_name):
        feedback_cell.value = f"ERROR: {package_name} already exists!"
        return

    # Check if package exists on PyPI
    try:
        fetch_package_data(package_name)
    except requests.HTTPError as e:
        if e.response.status_code == 404:
            # Show a specific error message when the package doesn't exist
            feedback_cell.value = f"ERROR: '{package_name}' doesn't exist on PyPI"
        else:
            # These could be internal server errors, etc.
            feedback_cell.value = f"ERROR: Couldn't access '{package_name}' on PyPI"
        return
    except requests.RequestException as e:
        # This covers all other Requests errors, e.g., connection errors, timeouts, etc.
        feedback_cell.value = f"ERROR: Couldn't contact PyPI! Details: {repr(e)}"
        return

    try:
        database.store_package(package_name)
    except Exception as e:
        feedback_cell.value = f"ERROR: {repr(e)}"
        return

    db_sheet["new_package"].clear_contents()
    feedback_cell.value = f"Added {package_name} successfully."
    update_database()
    refresh_dropdown()


def update_database() -> None:
    """Deletes all records from the versions table, fetches all
    data from PyPI and stores the versions of all packages again in the table.
    """
    # Excel objects
    sheet_db = xw.Book.caller().sheets["Database"]

    # Clear logs
    sheet_db["log"].expand().clear_contents()

    # Keeping things super simple: delete all versions for all packages
    # and repopulate the package_versions table from scratch
    database.delete_versions()
    packages = database.get_packages()
    logs = []

    for package_id, row in packages.iterrows():
        package_name = row["package_name"]

        try:
            data = fetch_package_data(package_name)
            logs.append(f"INFO: {package_name} downloaded successfully")
        except requests.RequestException as e:
            logs.append(
                f"ERROR: Could not fetch data for {package_name} ({type(e).__name__})"
            )
            continue

        releases = extract_releases(data, package_id)

        try:
            database.store_versions(releases)
            logs.append(f"INFO: {package_name} stored successfully")
        except Exception as e:
            logs.append(f"ERROR: Failed to store {package_name} ({repr(e)})")

    sheet_db["updated_at"].value = f"Last updated: {dt.datetime.now()}"
    sheet_db["log"].options(transpose=True).value = logs


def create_releases_chart(releases: pd.DataFrame, package_name: str) -> Figure:
    """Create a bar chart showing releases per year"""
    releases_yearly = releases.resample("YE").count()
    releases_yearly.index = releases_yearly.index.year
    releases_yearly.index.name = "Years"
    releases_yearly = releases_yearly.rename(columns={"version": "Number of Releases"})

    ax = releases_yearly.plot.bar(
        title=f"Releases per Year ({package_name})", legend=False
    )
    return ax.get_figure()


def show_history() -> None:
    """Shows the latest release and plots the release history
    (number of releases per year)
    """
    # Excel objects
    book = xw.Book.caller()
    tracker_sheet = book.sheets["Tracker"]
    package_name = tracker_sheet["package_selection"].value
    feedback_cell = tracker_sheet["package_selection"].offset(column_offset=1)
    picture_cell = tracker_sheet["latest_release"].offset(row_offset=2)
    latest_release_cell = tracker_sheet["latest_release"]

    # Check input
    if not package_name:
        feedback_cell.value = "ERROR: Please select a package first!"
        return

    # Clear output cells and picture
    feedback_cell.clear_contents()
    tracker_sheet["latest_release"].clear_contents()
    if "releases_per_year" in tracker_sheet.pictures:
        tracker_sheet.pictures["releases_per_year"].delete()

    # Get all versions of the package from the database
    try:
        releases = database.get_versions(package_name)
    except Exception as e:
        feedback_cell.value = f"ERROR: {e}"
        return

    if releases.empty:
        feedback_cell.value = f"ERROR: No releases found for {package_name}"
        return

    # Plot
    fig = create_releases_chart(releases, package_name)

    latest_date = releases.index.max()
    latest_version = releases.loc[latest_date, "version"]
    latest_release_cell.value = f"{latest_version} ({latest_date:%B %d, %Y})"

    tracker_sheet.pictures.add(
        fig, name="releases_per_year", top=picture_cell.top, left=picture_cell.left
    )


def refresh_dropdown() -> None:
    """Refreshes the dropdown on the Tracker tab with the content of
    the packages table.
    """
    # Excel objects
    book = xw.Book.caller()
    dropdown_sheet = book.sheets["Dropdown"]
    tracker_sheet = book.sheets["Tracker"]
    dropdown_content_cell = dropdown_sheet["dropdown_content"]

    # Clear the current value in the dropdown
    tracker_sheet["package_selection"].clear_contents()

    # Delete the content of the Excel table and repopulate it
    # with the values from the packages database table
    if dropdown_content_cell.value:
        dropdown_content_cell.delete()
    packages = database.get_packages()
    dropdown_content_cell.options(header=False, index=False).value = packages


if __name__ == "__main__":
    xw.Book("packagetracker.xlsm").set_mock_caller()
    add_package()
