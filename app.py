import io
import re
import json
import time
import zipfile
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Worldsensing Automation",
    page_icon="📡",
    layout="wide",
)

EXCEL_MAX_ROWS = 1_048_000  # Excel hard limit is 1,048,576 rows per sheet
XLSX_MIME = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


@st.cache_resource(show_spinner="Installing browser (first run only)...")
def ensure_browser():
    """Download Chromium for Playwright once per server start."""
    import subprocess
    import sys

    result = subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr[-800:] or result.stdout[-800:])
    return True


st.title("📡 Worldsensing Automation Portal")
st.caption("Sensor Management | CSV Collection | Excel Consolidation")

# --------------------------------------------------
# Helpers
# --------------------------------------------------

def safe_name(value):
    value = re.sub(r"[^\w.-]+", "_", str(value)).strip("._")
    return value[:100] or "unnamed_sensor"


def normalize_sensor_name(name, sensor_id=""):
    name = str(name).strip()
    if not name or name.lower().startswith("http"):
        name = f"Sensor_{sensor_id}" if sensor_id else "Sensor"
    return name


# Filename example:
# 177340-LSG7ACLBXLHVIB-1-readings-VIBPPV-current.csv
#   177340          -> Serial number (node ID)
#   LSG7ACLBXLHVIB  -> Model code
#   1               -> Channel
#   VIBPPV          -> Reading type
#   current         -> Period
FILENAME_PATTERN = re.compile(
    r"^(?P<serial>\d+)-(?P<model>[A-Za-z0-9]+)-(?P<channel>\d+)"
    r"-readings-(?P<reading_type>[A-Za-z0-9]+)-(?P<period>[A-Za-z0-9]+)",
    re.IGNORECASE,
)


def parse_report_filename(filename, fallback_serial=""):
    """Extract serial / model / channel / reading type from the CSV name."""
    stem = Path(str(filename)).name
    m = FILENAME_PATTERN.match(stem)
    if m:
        return {
            "Serial_Number": m.group("serial"),
            "Model": m.group("model").upper(),
            "Channel": m.group("channel"),
            "Reading_Type": m.group("reading_type"),
            "Period": m.group("period"),
        }
    return {
        "Serial_Number": fallback_serial,
        "Model": "",
        "Channel": "",
        "Reading_Type": "",
        "Period": "",
    }


def enrich_frame(df, sensor_name, sensor_id, filename):
    """Insert sensor identification columns at the start of a dataframe."""
    meta = parse_report_filename(filename, fallback_serial=sensor_id)
    current = df.copy()

    columns = {
        "Sensor_Name": sensor_name,
        "Sensor_ID": sensor_id,
        "Serial_Number": meta["Serial_Number"],
        "Model": meta["Model"],
        "Channel": meta["Channel"],
        "Reading_Type": meta["Reading_Type"],
        "Source_File": filename,
    }

    # Drop same-named columns if the CSV already has them.
    for col in columns:
        if col in current.columns:
            current = current.drop(columns=[col])

    for pos, (col, val) in enumerate(columns.items()):
        current.insert(pos, col, val)

    return current


def sort_by_time(data):
    if "Measurement-datetime" in data.columns:
        data = data.copy()
        data["_sort_time"] = pd.to_datetime(
            data["Measurement-datetime"], errors="coerce", utc=True
        )
        data = data.sort_values(
            ["Sensor_Name", "_sort_time"], na_position="last"
        ).drop(columns=["_sort_time"])
    return data


def build_summary(data):
    id_cols = [
        c for c in
        ["Sensor_Name", "Serial_Number", "Model", "Channel", "Reading_Type"]
        if c in data.columns
    ]
    return (
        data.groupby(id_cols, dropna=False)
        .size()
        .reset_index(name="Rows")
    )


def make_excel(sheets):
    """Generic writer: sheets = {sheet_name: dataframe}."""
    output = io.BytesIO()
    used = set()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            sheet = safe_name(name)[:31] or "Sheet"
            base, n = sheet, 1
            while sheet in used:
                sheet = f"{base[:28]}_{n}"
                n += 1
            used.add(sheet)
            df.to_excel(writer, sheet_name=sheet, index=False)

        if not sheets:
            pd.DataFrame({"Info": ["No readable reports found."]}).to_excel(
                writer, sheet_name="Info", index=False
            )

    return output.getvalue()


def make_sensor_workbook(frames):
    """One workbook for one sensor (frames are already enriched)."""
    frames = [f for f in frames if not f.empty]
    if not frames:
        return make_excel({})
    data = sort_by_time(pd.concat(frames, ignore_index=True, sort=False))
    return make_excel({"All_Readings": data})


def make_master_outputs(all_frames):
    """
    Build the single consolidated file for ALL sensors.
    Returns (xlsx_bytes_or_None, csv_bytes, combined_dataframe).
    """
    frames = [f for f in all_frames if not f.empty]
    if not frames:
        return None, b"", pd.DataFrame()

    data = sort_by_time(pd.concat(frames, ignore_index=True, sort=False))
    csv_bytes = data.to_csv(index=False).encode("utf-8-sig")

    xlsx_bytes = None
    if len(data) <= EXCEL_MAX_ROWS:
        xlsx_bytes = make_excel(
            {
                "All_Sensors": data,
                "Sensor_Summary": build_summary(data),
            }
        )

    return xlsx_bytes, csv_bytes, data


# --------------------------------------------------
# Sidebar: portal configuration
# --------------------------------------------------

with st.sidebar:
    st.header("Portal configuration")

    login_url = st.text_input(
        "Login URL",
        placeholder="https://your-portal/login",
    )

    sensors_url = st.text_input(
        "Network / sensors page URL",
        value="https://loadsensing.wocs3.com/33998/dataserver/",
    )

    username = st.text_input("Username", key="portal_username")
    password = st.text_input(
        "Password",
        type="password",
        key="portal_password",
    )

    st.divider()
    st.subheader("Login selectors")

    username_selector = st.text_input(
        "Username input CSS",
        value="input[name='username']",
    )
    password_selector = st.text_input(
        "Password input CSS",
        value="input[type='password']",
    )
    submit_selector = st.text_input(
        "Login button CSS",
        value="button[type='submit']",
    )

    st.divider()
    st.subheader("Sensor page selectors")

    sensor_selector = st.text_input(
        "Sensor links CSS",
        value="a[href*='/node/'], a[href*='/nodes/']",
    )

    report_selector = st.text_input(
        "Readings CSV links CSS",
        value="a[href*='.csv'], a:has-text('readings CSV')",
    )

    sensor_name_selector = st.text_input(
        "Sensor name CSS (optional)",
        value="",
    )

    delay = st.number_input(
        "Navigation delay (seconds)",
        min_value=1,
        max_value=15,
        value=2,
    )

    max_sensors = st.number_input(
        "Maximum sensors per run (0 = all)",
        min_value=0,
        max_value=1000,
        value=5,
    )

# --------------------------------------------------
# Main tabs
# --------------------------------------------------

tab_inventory, tab_collection, tab_direct = st.tabs(
    [
        "1. Sensor Management",
        "2. Automatic Collection",
        "3. Test a CSV URL",
    ]
)

# --------------------------------------------------
# Browser functions
# --------------------------------------------------

def create_authenticated_page(playwright):
    ensure_browser()

    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    if login_url:
        page.goto(login_url, wait_until="domcontentloaded", timeout=60000)

        if username and password:
            if page.locator(username_selector).count():
                page.locator(username_selector).first.fill(username)

            if page.locator(password_selector).count():
                page.locator(password_selector).first.fill(password)

            if page.locator(submit_selector).count():
                page.locator(submit_selector).first.click()
                page.wait_for_load_state("domcontentloaded", timeout=30000)

            page.wait_for_timeout(delay * 1000)

    page.goto(sensors_url, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(delay * 1000)

    return browser, context, page


def get_sensors(page):
    elements = page.locator(sensor_selector)
    sensors = []

    for i in range(elements.count()):
        el = elements.nth(i)
        href = el.get_attribute("href")

        if not href:
            continue

        full_url = urljoin(page.url, href)
        label = ""

        if sensor_name_selector:
            try:
                label = el.locator(
                    sensor_name_selector
                ).first.inner_text(timeout=1000).strip()
            except Exception:
                pass

        if not label:
            try:
                label = el.inner_text(timeout=1000).strip()
            except Exception:
                pass

        sensor_id = re.search(r"/(\d+)(?:/|$)", urlparse(full_url).path)
        sensor_id = sensor_id.group(1) if sensor_id else ""

        sensors.append({
            "Sensor_Name": normalize_sensor_name(label, sensor_id),
            "Sensor_ID": sensor_id,
            "Sensor_URL": full_url,
        })

    # Deduplicate by URL.
    unique = {}
    for sensor in sensors:
        unique[sensor["Sensor_URL"]] = sensor

    return list(unique.values())


def get_report_links(page):
    elements = page.locator(report_selector)
    results = []
    seen = set()

    for i in range(elements.count()):
        el = elements.nth(i)
        href = el.get_attribute("href")

        if href:
            full_url = urljoin(page.url, href)

            if full_url in seen:
                continue

            # Restrict downloads to CSV endpoints.
            if ".csv" in full_url.lower() or "reading/" in full_url.lower():
                seen.add(full_url)
                label = el.get_attribute("download") or ""
                results.append({
                    "url": full_url,
                    "label": label or f"report_{i + 1}.csv",
                })

    return results


def download_csv_with_context(context, url):
    response = context.request.get(url, timeout=60000)

    if not response.ok:
        raise RuntimeError(f"HTTP {response.status} while fetching the report")

    content = response.body()

    if not content:
        raise RuntimeError("The response was empty.")

    # Detect HTML login pages returned instead of CSV data.
    start = content[:500].lstrip().lower()
    if start.startswith(b"<!doctype html") or b"<html" in start:
        raise RuntimeError(
            "Received HTML instead of CSV. Authentication may have expired "
            "or the URL may require another login step."
        )

    return content


def parse_csv(content):
    errors = []

    for encoding in ["utf-8-sig", "utf-8", "cp1252", "latin1"]:
        try:
            df = pd.read_csv(
                io.BytesIO(content),
                encoding=encoding,
                sep=None,
                engine="python",
            )
            if len(df.columns) > 1:
                return df
        except Exception as exc:
            errors.append(str(exc))

    raise ValueError("Could not parse CSV: " + "; ".join(errors[-2:]))


# --------------------------------------------------
# Tab 1: Sensor Management
# --------------------------------------------------

with tab_inventory:
    st.subheader("Discover sensors")

    st.write(
        "Sign in through the configured login page, then scan the network "
        "page for links matching the sensor selector."
    )

    if st.button("Load sensor inventory", type="primary"):
        if not sensors_url:
            st.error("Enter the network/sensors page URL.")
        else:
            try:
                from playwright.sync_api import sync_playwright

                with sync_playwright() as p:
                    browser, context, page = create_authenticated_page(p)
                    inventory = get_sensors(page)
                    browser.close()

                st.session_state["sensor_inventory"] = inventory

                if inventory:
                    st.success(f"Found {len(inventory)} sensor link(s).")
                else:
                    st.warning(
                        "No sensors were found. Check the sensor selector "
                        "and confirm the page is logged in."
                    )

            except Exception as exc:
                st.error(f"Could not load inventory: {exc}")

    inventory = st.session_state.get("sensor_inventory", [])

    if inventory:
        inventory_df = pd.DataFrame(inventory)

        st.dataframe(inventory_df, use_container_width=True, hide_index=True)

        st.download_button(
            "Export inventory CSV",
            inventory_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="sensor_inventory.csv",
            mime="text/csv",
        )

# --------------------------------------------------
# Tab 2: Automatic collection
# --------------------------------------------------

with tab_collection:
    st.subheader("Collect sensor reports")

    st.write(
        "The app opens each sensor page, downloads every readings CSV, adds "
        "Sensor / Serial / Model columns, and merges everything into ONE "
        "master file (plus per-sensor workbooks in the ZIP)."
    )

    if st.button("Start report collection", type="primary"):
        try:
            from playwright.sync_api import sync_playwright

            with sync_playwright() as p:
                browser, context, page = create_authenticated_page(p)

                sensors = get_sensors(page)

                if max_sensors:
                    sensors = sensors[:int(max_sensors)]

                if not sensors:
                    browser.close()
                    raise RuntimeError(
                        "No sensors found. Configure the sensor selector first."
                    )

                downloaded = []
                errors = []
                workbooks = {}
                all_frames = []  # enriched frames from ALL sensors

                progress = st.progress(0)
                status = st.empty()

                for i, sensor in enumerate(sensors):
                    name = sensor["Sensor_Name"]
                    sensor_id = sensor["Sensor_ID"]

                    status.write(f"Sensor {i + 1}/{len(sensors)}: {name}")

                    try:
                        page.goto(
                            sensor["Sensor_URL"],
                            wait_until="domcontentloaded",
                            timeout=60000,
                        )
                        page.wait_for_timeout(delay * 1000)

                        report_links = get_report_links(page)
                        sensor_frames = []

                        for report in report_links:
                            try:
                                content = download_csv_with_context(
                                    context, report["url"]
                                )

                                filename = (
                                    Path(urlparse(report["url"]).path).name
                                    or report["label"]
                                )

                                df = parse_csv(content)
                                enriched = enrich_frame(
                                    df, name, sensor_id, filename
                                )

                                downloaded.append({
                                    "sensor": name,
                                    "sensor_id": sensor_id,
                                    "serial": (
                                        enriched["Serial_Number"].iloc[0]
                                        if len(enriched) else ""
                                    ),
                                    "model": (
                                        enriched["Model"].iloc[0]
                                        if len(enriched) else ""
                                    ),
                                    "filename": filename,
                                    "content": content,
                                    "url": report["url"],
                                    "rows": len(df),
                                })

                                sensor_frames.append(enriched)
                                all_frames.append(enriched)

                            except Exception as exc:
                                errors.append(
                                    f"{name} / {report['url']}: {exc}"
                                )

                        if sensor_frames:
                            workbooks[name] = make_sensor_workbook(
                                sensor_frames
                            )
                        else:
                            errors.append(
                                f"{name}: no usable readings CSV was downloaded."
                            )

                    except Exception as exc:
                        errors.append(f"{name}: {exc}")

                    progress.progress((i + 1) / len(sensors))

                browser.close()

            if not downloaded:
                st.error("No CSV reports were downloaded.")
                if errors:
                    st.code("\n".join(errors[:30]))
            else:
                st.success(
                    f"Downloaded {len(downloaded)} report(s) from "
                    f"{len(sensors)} sensor page(s)."
                )

                summary = pd.DataFrame([
                    {
                        "Sensor": item["sensor"],
                        "Sensor ID": item["sensor_id"],
                        "Serial": item["serial"],
                        "Model": item["model"],
                        "Report": item["filename"],
                        "Rows": item["rows"],
                    }
                    for item in downloaded
                ])

                st.dataframe(summary, use_container_width=True, hide_index=True)

                # ---------- Master consolidated file ----------
                master_xlsx, master_csv, master_df = make_master_outputs(
                    all_frames
                )

                st.subheader("Consolidated master file")
                st.caption(
                    f"{len(master_df):,} rows from {len(sensors)} sensor(s)."
                )
                st.dataframe(
                    master_df.head(50),
                    use_container_width=True,
                    hide_index=True,
                )

                col1, col2 = st.columns(2)

                with col1:
                    if master_xlsx:
                        st.download_button(
                            "Download MASTER Excel (all sensors)",
                            data=master_xlsx,
                            file_name="worldsensing_master.xlsx",
                            mime=XLSX_MIME,
                            type="primary",
                        )
                    else:
                        st.warning(
                            "Too many rows for one Excel sheet - use the CSV."
                        )

                with col2:
                    st.download_button(
                        "Download MASTER CSV (all sensors)",
                        data=master_csv,
                        file_name="worldsensing_master.csv",
                        mime="text/csv",
                    )

                # ---------- ZIP with everything ----------
                bundle = io.BytesIO()

                with zipfile.ZipFile(
                    bundle, "w", zipfile.ZIP_DEFLATED
                ) as archive:

                    for item in downloaded:
                        archive.writestr(
                            "original_reports/"
                            + safe_name(item["sensor"])
                            + "__"
                            + safe_name(item["filename"]),
                            item["content"],
                        )

                    for name, workbook in workbooks.items():
                        archive.writestr(
                            f"per_sensor_workbooks/{safe_name(name)}.xlsx",
                            workbook,
                        )

                    if master_xlsx:
                        archive.writestr(
                            "worldsensing_master.xlsx", master_xlsx
                        )
                    archive.writestr("worldsensing_master.csv", master_csv)

                    archive.writestr(
                        "sensor_inventory.csv",
                        pd.DataFrame(sensors).to_csv(index=False),
                    )
                    archive.writestr(
                        "collection_summary.csv",
                        summary.to_csv(index=False),
                    )
                    archive.writestr(
                        "manifest.json",
                        json.dumps(
                            {
                                "reports_downloaded": len(downloaded),
                                "sensors_scanned": len(sensors),
                                "master_rows": len(master_df),
                                "issues": errors,
                                "created_at_unix": int(time.time()),
                            },
                            indent=2,
                        ),
                    )

                st.download_button(
                    "Download everything (ZIP)",
                    data=bundle.getvalue(),
                    file_name="worldsensing_collection.zip",
                    mime="application/zip",
                )

                if errors:
                    st.warning(f"{len(errors)} issue(s) occurred.")
                    st.code("\n".join(errors[:50]))

        except Exception as exc:
            st.error(f"Collection failed: {exc}")

# --------------------------------------------------
# Tab 3: Direct CSV test
# --------------------------------------------------

with tab_direct:
    st.subheader("Test a direct readings CSV URL")

    st.write(
        "Use this with a sample URL to validate CSV parsing, the "
        "serial/model detection and Excel generation."
    )

    csv_url = st.text_input(
        "Readings CSV URL",
        value=(
            "https://loadsensing.wocs3.com/33998/dataserver/"
            "current/reading/177340/ppvSufre/"
            "177340-LSG7ACLBXLHVIB-1-readings-VIBPPV-current.csv"
        ),
    )

    if st.button("Test CSV URL"):
        try:
            from playwright.sync_api import sync_playwright

            with sync_playwright() as p:
                browser, context, page = create_authenticated_page(p)
                content = download_csv_with_context(context, csv_url)
                df = parse_csv(content)
                browser.close()

            filename = Path(urlparse(csv_url).path).name
            meta = parse_report_filename(filename)
            sensor_id = re.search(r"/reading/(\d+)/", csv_url)
            sensor_id = (
                sensor_id.group(1) if sensor_id else meta["Serial_Number"]
            )

            enriched = enrich_frame(
                df, f"Sensor_{sensor_id}", sensor_id, filename
            )

            st.success(
                f"CSV loaded: {len(df)} rows, {len(df.columns)} columns."
            )

            st.write("Detected from filename")
            st.json(meta)

            st.write("Detected columns")
            st.code(", ".join(map(str, enriched.columns)))

            st.dataframe(enriched.head(20), use_container_width=True)

            excel_bytes = make_excel({"All_Readings": sort_by_time(enriched)})

            st.download_button(
                "Download test Excel",
                data=excel_bytes,
                file_name=f"Sensor_{sensor_id}.xlsx",
                mime=XLSX_MIME,
            )

        except Exception as exc:
            st.error(f"CSV test failed: {exc}")
            st.caption(
                "If the URL opens only when you are signed in, the login "
                "flow or session handling must be configured correctly."
            )
