import io
import re
import json
import time
import zipfile
from collections import Counter
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Worldsensing Automation", page_icon="📡", layout="wide")

for _module, _package in (("requests", "requests"), ("openpyxl", "openpyxl")):
    try:
        __import__(_module)
    except ImportError:
        st.error(
            f"Missing Python package: **{_package}**. Add the line "
            f"`{_package}` to requirements.txt (repo root), commit, "
            "then Reboot the app."
        )
        st.stop()

import requests  # noqa: E402

EXCEL_MAX_ROWS = 1_048_000
XLSX_MIME = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
TIMEOUT = 60


class AuthError(Exception):
    pass


def secret(key, default=""):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


# --------------------------------------------------
# HTML parsers (links, forms, tables) - no extra packages
# --------------------------------------------------
class PageParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.links = []
        self.forms = []
        self._a = None
        self._form = None

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "a" and a.get("href"):
            self._a = {"href": a["href"], "text": ""}
            self.links.append(self._a)
        elif tag == "form":
            self._form = {
                "action": a.get("action") or "",
                "method": (a.get("method") or "get").lower(),
                "inputs": [],
            }
            self.forms.append(self._form)
        elif tag == "input" and self._form is not None:
            self._form["inputs"].append({
                "name": a.get("name"),
                "type": (a.get("type") or "text").lower(),
                "value": a.get("value") or "",
            })

    def handle_endtag(self, tag):
        if tag == "a":
            self._a = None
        elif tag == "form":
            self._form = None

    def handle_data(self, data):
        if self._a is not None:
            self._a["text"] += data


class TableParser(HTMLParser):
    """Collects every table row as a list of cells (text + links)."""

    def __init__(self):
        super().__init__()
        self.rows = []
        self._row = None
        self._cell = None
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in ("script", "style"):
            self._skip += 1
        elif tag == "tr":
            self._row = []
            self.rows.append(self._row)
        elif tag in ("td", "th") and self._row is not None:
            self._cell = {"text": "", "links": []}
            self._row.append(self._cell)
        elif tag == "a" and self._cell is not None and a.get("href"):
            self._cell["links"].append(a["href"])

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self._skip = max(0, self._skip - 1)
        elif tag in ("td", "th"):
            self._cell = None
        elif tag == "tr":
            self._row = None

    def handle_data(self, data):
        if self._cell is not None and not self._skip:
            self._cell["text"] += data


def parse_page(html):
    parser = PageParser()
    try:
        parser.feed(html)
    except Exception:
        pass
    return parser


def parse_tables(html):
    parser = TableParser()
    try:
        parser.feed(html)
    except Exception:
        pass
    return parser.rows


def clean_text(text):
    return " ".join(str(text).split())


def find_login_form(page):
    for form in page.forms:
        if any(i["type"] == "password" for i in form["inputs"]):
            return form
    return None


def build_login_payload(form, username, password):
    inputs = [i for i in form["inputs"] if i["name"]]
    pass_field = next((i["name"] for i in inputs if i["type"] == "password"), None)
    candidates = [i for i in inputs if i["type"] in ("text", "email")]
    user_field = next(
        (c["name"] for c in candidates
         if re.search(r"user|email|login|name", c["name"], re.I)),
        candidates[0]["name"] if candidates else None,
    )
    if not pass_field or not user_field:
        raise AuthError(
            "Login form found, but username/password fields could not be "
            "identified."
        )

    payload = {}
    submit_added = False
    for i in inputs:
        t = i["type"]
        if i["name"] in (user_field, pass_field):
            continue
        if t in ("checkbox", "radio", "button", "reset", "file", "image"):
            continue
        if t == "submit":
            if not submit_added:
                payload[i["name"]] = i["value"]
                submit_added = True
            continue
        payload[i["name"]] = i["value"]

    payload[user_field] = username
    payload[pass_field] = password
    return payload


# --------------------------------------------------
# Connection / login
# --------------------------------------------------
def connect(site_url, login_url, username, password):
    if not site_url.strip():
        raise AuthError("Enter the site link.")

    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (Worldsensing Collector)"})
    log = []
    start = login_url.strip() or site_url.strip()

    try:
        r = session.get(start, timeout=TIMEOUT)

        if r.status_code == 401:
            if not (username and password):
                raise AuthError("Site requires login: enter username and password.")
            session.auth = (username, password)
            r = session.get(start, timeout=TIMEOUT)
            if r.status_code == 401:
                raise AuthError("Username or password was rejected (HTTP 401).")
            log.append("Authenticated using HTTP basic authentication.")

        if r.status_code >= 400:
            raise RuntimeError(f"Site returned HTTP {r.status_code} for {start}")

        form = find_login_form(parse_page(r.text))

        if form:
            if not (username and password):
                raise AuthError("A login form was found: enter username and password.")

            payload = build_login_payload(form, username, password)
            action = urljoin(r.url, form["action"]) if form["action"] else r.url

            if form["method"] == "post":
                r2 = session.post(
                    action, data=payload, timeout=TIMEOUT,
                    headers={"Referer": r.url},
                )
            else:
                r2 = session.get(action, params=payload, timeout=TIMEOUT)

            if r2.status_code >= 400:
                raise AuthError(f"Login request failed (HTTP {r2.status_code}).")

            if find_login_form(parse_page(r2.text)):
                raise AuthError(
                    "Login failed: the site is still showing the login form. "
                    "Check username/password (or the Login page link)."
                )
            log.append("Login form submitted successfully.")
        else:
            no_forms = not parse_page(r.text).forms
            if no_forms and not session.auth and "login" in r.url.lower():
                raise AuthError(
                    "The login page seems to be built with JavaScript, so this "
                    "simple version cannot fill it in."
                )
            log.append("No login form detected; continuing.")

        r3 = session.get(site_url.strip(), timeout=TIMEOUT)
        if r3.status_code >= 400:
            raise RuntimeError(f"Site returned HTTP {r3.status_code} for the site link.")
        if find_login_form(parse_page(r3.text)):
            raise AuthError("Not authenticated: site link still asks for login.")

    except requests.exceptions.RequestException as exc:
        raise RuntimeError(f"Could not reach the site: {type(exc).__name__}")

    return session, r3, log


def ensure_session(site_url, login_url, username, password):
    if st.session_state.get("session") is not None:
        return st.session_state["session"]
    session, response, log = connect(site_url, login_url, username, password)
    st.session_state["session"] = session
    st.session_state["site_html"] = response.text
    st.session_state["site_final_url"] = response.url
    st.session_state["connect_log"] = log
    return session


# --------------------------------------------------
# Sensor discovery (Nodes table: Id | Name | Model | Status)
# --------------------------------------------------
def find_col(lowered, key):
    for i, t in enumerate(lowered):
        if t == key or t.startswith(key + " "):
            return i
    return None


def default_node_url(base_url, sid):
    m = re.match(r"(.*?/dataserver/)", base_url)
    prefix = m.group(1) if m else base_url.rstrip("/") + "/"
    return f"{prefix}node/view/{sid}"


def discover_nodes(base_url, html):
    """Read the Nodes table. Returns (nodes, expected_total_or_None)."""
    rows = parse_tables(html)
    header = None
    nodes = {}

    for row in rows:
        texts = [clean_text(c["text"]) for c in row]
        lowered = [t.lower() for t in texts]

        if header is None:
            i_id = find_col(lowered, "id")
            i_model = find_col(lowered, "model")
            if i_id is not None and i_model is not None:
                header = {
                    "id": i_id,
                    "model": i_model,
                    "name": find_col(lowered, "name"),
                    "status": find_col(lowered, "status"),
                }
            continue

        if len(texts) <= max(header["id"], header["model"]):
            continue

        sid = texts[header["id"]]
        if not re.fullmatch(r"\d+", sid):
            continue

        links = row[header["id"]]["links"]
        url = (
            urljoin(base_url, links[0]) if links
            else default_node_url(base_url, sid)
        )

        name = ""
        if header["name"] is not None and header["name"] < len(texts):
            name = texts[header["name"]]
        status = ""
        if header["status"] is not None and header["status"] < len(texts):
            status = texts[header["status"]]

        nodes[url] = {
            "Sensor_Name": name or f"Sensor_{sid}",
            "Sensor_ID": sid,
            "Model": texts[header["model"]],
            "Status": status,
            "Sensor_URL": url,
        }

    expected = None
    m = re.search(r"selected of\s+(\d+)", clean_text(re.sub(r"<[^>]+>", " ", html)))
    if m:
        expected = int(m.group(1))

    return list(nodes.values()), expected


def discover_by_links(base_url, html, link_filter):
    """Fallback when no Nodes table is found."""
    page = parse_page(html)
    found = {}
    for link in page.links:
        href = link["href"].strip()
        if not href or href.startswith(("#", "javascript:", "mailto:")):
            continue
        full = urljoin(base_url, href)
        if link_filter and link_filter.lower() not in full.lower():
            continue
        m = re.search(r"(\d+)/?$", urlparse(full).path)
        sid = m.group(1) if m else ""
        label = clean_text(link["text"])
        if not label or label.lower().startswith("http"):
            label = f"Sensor_{sid}" if sid else "Sensor"
        found[full] = {
            "Sensor_Name": label,
            "Sensor_ID": sid,
            "Model": "",
            "Status": "",
            "Sensor_URL": full,
        }
    return list(found.values())


def manual_sensors(text):
    sensors = []
    for line in text.splitlines():
        url = line.strip()
        if url.lower().startswith("http"):
            m = re.search(r"(\d+)/?$", urlparse(url).path)
            sid = m.group(1) if m else ""
            sensors.append({
                "Sensor_Name": f"Sensor_{sid}" if sid else "Sensor",
                "Sensor_ID": sid,
                "Model": "",
                "Status": "",
                "Sensor_URL": url,
            })
    return sensors


def get_csv_links(base_url, html):
    page = parse_page(html)
    results, seen = [], set()
    for link in page.links:
        href = link["href"].strip()
        if not href:
            continue
        text = clean_text(link["text"]).lower()
        if ".csv" in href.lower() or "readings csv" in text:
            full = urljoin(base_url, href)
            if full not in seen:
                seen.add(full)
                results.append(full)
    return results


# --------------------------------------------------
# Download / parse CSV
# --------------------------------------------------
def download_csv(session, url):
    try:
        r = session.get(url, timeout=TIMEOUT)
    except requests.exceptions.RequestException as exc:
        raise RuntimeError(f"Network error ({type(exc).__name__})")

    if r.status_code in (401, 403):
        raise AuthError(f"Access denied (HTTP {r.status_code}). Reconnect and try again.")
    if r.status_code >= 400:
        raise RuntimeError(f"HTTP {r.status_code} while downloading the report")

    content = r.content
    if not content:
        raise RuntimeError("The response was empty.")

    head = content[:600].lstrip().lower()
    if head.startswith(b"<!doctype html") or b"<html" in head:
        raise AuthError(
            "Received a web page instead of CSV. The session may have expired "
            "or the link is not a CSV file."
        )
    return content


def parse_csv(content):
    errors = []
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            df = pd.read_csv(
                io.BytesIO(content), encoding=encoding, sep=None, engine="python"
            )
            if len(df.columns) > 1:
                return df
        except Exception as exc:
            errors.append(str(exc)[:120])
    raise ValueError("Could not parse CSV: " + "; ".join(errors[-2:]))


# --------------------------------------------------
# Sensor identification + Excel
# --------------------------------------------------
FILENAME_PATTERN = re.compile(
    r"^(?P<serial>\d+)-(?P<model>[A-Za-z0-9]+)-(?P<channel>\d+)"
    r"-readings-(?P<rtype>[A-Za-z0-9]+)-(?P<period>[A-Za-z0-9]+)",
    re.IGNORECASE,
)


def safe_name(value):
    value = re.sub(r"[^\w.-]+", "_", str(value)).strip("._")
    return value[:100] or "unnamed"


def parse_filename(filename, fallback_serial=""):
    m = FILENAME_PATTERN.match(Path(str(filename)).name)
    if m:
        return {
            "Serial_Number": m.group("serial"),
            "Model": m.group("model").upper(),
            "Channel": m.group("channel"),
            "Reading_Type": m.group("rtype"),
        }
    return {
        "Serial_Number": fallback_serial,
        "Model": "",
        "Channel": "",
        "Reading_Type": "",
    }


def enrich_frame(df, sensor_name, sensor_id, filename, model=""):
    """Add identification columns. Model from the Nodes table wins."""
    meta = parse_filename(filename, fallback_serial=sensor_id)
    out = df.copy()
    out.columns = [str(c) for c in out.columns]

    columns = {
        "Sensor_Name": sensor_name,
        "Sensor_ID": sensor_id,
        "Serial_Number": meta["Serial_Number"] or sensor_id,
        "Model": model or meta["Model"],
        "Channel": meta["Channel"],
        "Reading_Type": meta["Reading_Type"],
        "Source_Report": filename,
    }
    for col in columns:
        if col in out.columns:
            out = out.drop(columns=[col])
    for pos, (col, val) in enumerate(columns.items()):
        out.insert(pos, col, val)
    return out


def sort_by_time(data):
    if "Measurement-datetime" in data.columns:
        data = data.copy()
        data["_t"] = pd.to_datetime(
            data["Measurement-datetime"], errors="coerce", utc=True
        )
        data = data.sort_values(["Sensor_Name", "_t"], na_position="last")
        data = data.drop(columns=["_t"])
    return data


def make_excel(sheets):
    output = io.BytesIO()
    used = set()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wrote = False
        for name, df in sheets.items():
            sheet = safe_name(name)[:31] or "Sheet"
            base, n = sheet, 1
            while sheet.lower() in used:
                sheet = f"{base[:27]}_{n}"
                n += 1
            used.add(sheet.lower())

            if len(df) > EXCEL_MAX_ROWS:
                pd.DataFrame({"Info": [
                    f"{len(df)} rows exceed the Excel limit - use the CSV."
                ]}).to_excel(writer, sheet_name=sheet, index=False)
            else:
                df.to_excel(writer, sheet_name=sheet, index=False)
            wrote = True

        if not wrote:
            pd.DataFrame({"Info": ["No readable reports found."]}).to_excel(
                writer, sheet_name="Info", index=False
            )
    return output.getvalue()


def sensor_workbook(frames):
    """frames: list of (filename, enriched_df)."""
    sheets = {}
    nonempty = [f for _, f in frames if not f.empty]
    if nonempty:
        sheets["All_Readings"] = sort_by_time(
            pd.concat(nonempty, ignore_index=True, sort=False)
        )
    for filename, df in frames:
        meta = parse_filename(filename)
        label = (
            f"{meta['Reading_Type']}_ch{meta['Channel']}"
            if meta["Reading_Type"] else Path(filename).stem
        )
        sheets[label] = df
    return make_excel(sheets)


def master_outputs(all_frames):
    frames = [f for f in all_frames if not f.empty]
    if not frames:
        return None, b"", pd.DataFrame()

    data = sort_by_time(pd.concat(frames, ignore_index=True, sort=False))
    csv_bytes = data.to_csv(index=False).encode("utf-8-sig")

    xlsx = None
    if len(data) <= EXCEL_MAX_ROWS:
        id_cols = [c for c in
                   ["Sensor_Name", "Serial_Number", "Model", "Channel", "Reading_Type"]
                   if c in data.columns]
        summary = data.groupby(id_cols, dropna=False).size().reset_index(name="Rows")
        xlsx = make_excel({"All_Sensors": data, "Sensor_Summary": summary})
    return xlsx, csv_bytes, data


# --------------------------------------------------
# Collection
# --------------------------------------------------
def run_collection(session, sensors, delay):
    downloaded, errors, log = [], [], []
    originals, all_frames, workbooks = [], [], {}
    seen_urls = set()

    progress = st.progress(0.0)
    status = st.empty()

    for i, sensor in enumerate(sensors):
        name, sid = sensor["Sensor_Name"], sensor["Sensor_ID"]
        url, model = sensor["Sensor_URL"], sensor.get("Model", "")
        status.write(f"Sensor {i + 1}/{len(sensors)}: **{name}** (serial {sid})")

        try:
            r = session.get(url, timeout=TIMEOUT)
            if r.status_code >= 400:
                raise RuntimeError(f"HTTP {r.status_code} on the sensor page")
            if find_login_form(parse_page(r.text)):
                raise AuthError("Session expired (login form shown). Reconnect.")

            links = get_csv_links(r.url, r.text)
            if not links:
                raise RuntimeError("No CSV links found on this sensor page.")

            frames = []
            for csv_url in links:
                if csv_url in seen_urls:
                    continue
                seen_urls.add(csv_url)
                filename = Path(urlparse(csv_url).path).name or "report.csv"
                try:
                    content = download_csv(session, csv_url)
                    df = parse_csv(content)
                    enriched = enrich_frame(df, name, sid, filename, model)

                    originals.append((
                        f"original_reports/{sid}__{safe_name(filename)}",
                        content,
                    ))
                    frames.append((filename, enriched))
                    all_frames.append(enriched)
                    downloaded.append({
                        "Sensor": name,
                        "Serial": sid,
                        "Model": enriched["Model"].iloc[0] if len(enriched) else model,
                        "Report": filename,
                        "Rows": len(df),
                        "Columns": len(df.columns),
                    })
                    log.append(f"OK   {name} / {filename}: {len(df)} rows")
                except AuthError as exc:
                    errors.append(f"{name} ({sid}) / {filename}: AUTH - {exc}")
                except ValueError as exc:
                    errors.append(f"{name} ({sid}) / {filename}: CSV - {exc}")
                except Exception as exc:
                    errors.append(f"{name} ({sid}) / {filename}: {exc}")
                time.sleep(delay)

            if frames:
                workbooks[f"{sid}_{name}"] = sensor_workbook(frames)

        except AuthError as exc:
            errors.append(f"{name} ({sid}): AUTH - {exc}")
        except Exception as exc:
            errors.append(f"{name} ({sid}): {exc}")

        progress.progress((i + 1) / len(sensors))

    status.write("Collection finished.")

    summary = pd.DataFrame(downloaded)
    m_xlsx, m_csv, m_df = master_outputs(all_frames)

    bundle = io.BytesIO()
    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as z:
        for path, content in originals:
            z.writestr(path, content)
        for key, wb in workbooks.items():
            z.writestr(f"per_sensor_workbooks/{safe_name(key)}.xlsx", wb)
        if m_xlsx:
            z.writestr("worldsensing_master.xlsx", m_xlsx)
        if len(m_csv):
            z.writestr("worldsensing_master.csv", m_csv)
        z.writestr("sensor_inventory.csv", pd.DataFrame(sensors).to_csv(index=False))
        z.writestr("collection_summary.csv", summary.to_csv(index=False))
        z.writestr("collection_log.txt", "\n".join(log + ["ERRORS:"] + errors))
        z.writestr("manifest.json", json.dumps({
            "created_at": datetime.now(timezone.utc).isoformat(),
            "sensors_selected": len(sensors),
            "reports_downloaded": len(downloaded),
            "total_rows": int(summary["Rows"].sum()) if len(summary) else 0,
            "rows_per_report": downloaded,
            "errors": errors,
        }, indent=2, default=str))

    return {
        "summary": summary,
        "errors": errors,
        "master_xlsx": m_xlsx,
        "master_csv": m_csv,
        "master_df": m_df,
        "zip": bundle.getvalue(),
        "sensors": len(sensors),
    }


# --------------------------------------------------
# UI
# --------------------------------------------------
st.title("📡 Worldsensing Automation Portal")
st.caption("Connect | Filter sensors by model | Collect every serial | One Excel file")

with st.sidebar:
    st.header("Connection")
    site_url = st.text_input(
        "Nodes page link",
        value=secret("WS_SITE_URL", "https://loadsensing.wocs3.com/33998/dataserver/"),
        help="The page that shows the Nodes table (Id, Name, Model, Status). "
             "Copy it from your browser address bar.",
    )
    username = st.text_input("Username", value=secret("WS_USERNAME", ""))
    password = st.text_input("Password", type="password", value=secret("WS_PASSWORD", ""))

    with st.expander("Advanced (optional)"):
        login_url = st.text_input(
            "Login page link",
            value=secret("WS_LOGIN_URL", ""),
            help="Only needed if the Nodes page link does not redirect to login.",
        )
        link_filter = st.text_input(
            "Fallback: sensor link contains",
            value="/node/view/",
            help="Used only if no Nodes table is found on the page.",
        )
        delay = st.number_input("Pause between downloads (seconds)", 0.0, 10.0, 1.0, 0.5)

    st.caption("Credentials are only used for this session and never saved or logged.")

tab_sensors, tab_collect, tab_test = st.tabs(
    ["1. Connect & Sensors", "2. Collect", "3. Test a CSV URL"]
)

# ---------------- Tab 1 ----------------
with tab_sensors:
    st.subheader("Connect and read the Nodes table")

    if st.button("Connect and find sensors", type="primary"):
        st.session_state.pop("session", None)
        st.session_state.pop("model_pick", None)
        st.session_state.pop("status_pick", None)
        try:
            with st.spinner("Connecting..."):
                ensure_session(site_url, login_url, username, password)

            nodes, expected = discover_nodes(
                st.session_state["site_final_url"], st.session_state["site_html"]
            )
            if not nodes:
                nodes = discover_by_links(
                    st.session_state["site_final_url"],
                    st.session_state["site_html"],
                    link_filter,
                )
                st.session_state["table_mode"] = False
            else:
                st.session_state["table_mode"] = True

            st.session_state["nodes"] = nodes
            st.session_state["expected_nodes"] = expected
            st.success("Connected.")

            if not nodes:
                st.warning(
                    "Connected, but no sensors were found on this page. Make sure "
                    "'Nodes page link' is the page that shows the Nodes table, "
                    "or paste sensor links below."
                )
            elif expected and expected != len(nodes):
                st.warning(
                    f"The page says {expected} nodes, but {len(nodes)} were read. "
                    "The table may be paginated or loaded by JavaScript."
                )
        except AuthError as exc:
            st.error(f"Authentication problem: {exc}")
        except Exception as exc:
            st.error(f"Connection problem: {exc}")

    nodes = st.session_state.get("nodes", [])
    filtered = list(nodes)

    if nodes:
        st.write(f"**{len(nodes)}** sensors found on the page.")

        model_counts = Counter(n["Model"] for n in nodes if n["Model"])
        if model_counts:
            options = {
                f"{m} ({c})": m
                for m, c in sorted(model_counts.items(), key=lambda x: -x[1])
            }
            default_label = next(iter(options))
            picked = st.multiselect(
                "Filter by model",
                list(options.keys()),
                default=[default_label],
                key="model_pick",
                help="Only sensors with the selected model(s) will be collected.",
            )
            picked_models = {options[p] for p in picked}
            filtered = [n for n in filtered if n["Model"] in picked_models]

        status_values = sorted({n["Status"] for n in nodes if n["Status"]})
        if status_values:
            picked_status = st.multiselect(
                "Filter by status",
                status_values,
                default=status_values,
                key="status_pick",
            )
            filtered = [n for n in filtered if n["Status"] in picked_status]

    manual = st.text_area(
        "Optional: paste extra sensor page links (one per line)",
        placeholder="https://loadsensing.wocs3.com/33998/dataserver/node/view/177340",
    )

    inventory = list(filtered)
    known = {s["Sensor_URL"] for s in inventory}
    for s in manual_sensors(manual):
        if s["Sensor_URL"] not in known:
            inventory.append(s)
    st.session_state["inventory"] = inventory

    if inventory:
        st.success(f"{len(inventory)} sensor(s) selected for collection.")
        inv_df = pd.DataFrame(inventory)
        st.dataframe(inv_df, use_container_width=True, hide_index=True)
        st.download_button(
            "Export inventory CSV",
            inv_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="sensor_inventory.csv",
            mime="text/csv",
        )

    if st.session_state.get("site_html"):
        with st.expander("Diagnostics"):
            for line in st.session_state.get("connect_log", []):
                st.caption(line)
            st.caption(
                "Mode: "
                + ("Nodes table" if st.session_state.get("table_mode") else "link scan")
            )
            page = parse_page(st.session_state["site_html"])
            rows = [
                {"Text": clean_text(l["text"])[:60],
                 "Link": urljoin(st.session_state["site_final_url"], l["href"])}
                for l in page.links[:60]
            ]
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ---------------- Tab 2 ----------------
with tab_collect:
    st.subheader("Collect sensor reports")

    inventory = st.session_state.get("inventory", [])
    if not inventory:
        st.info("First connect and choose sensors in tab 1.")
    else:
        labels = {
            f"{s['Sensor_ID']} - {s['Sensor_Name']}": s for s in inventory
        }
        chosen = st.multiselect(
            "Sensors to collect (all filtered sensors by default)",
            list(labels.keys()),
            default=list(labels.keys()),
        )

        c1, c2 = st.columns(2)
        with c1:
            start_at = st.number_input("Start from sensor #", 1, 5000, 1)
        with c2:
            max_sensors = st.number_input(
                "Max sensors in this run (0 = all)", 0, 5000, 5,
                help="Start small (5) to test, then raise it or run in batches.",
            )

        selected = [labels[c] for c in chosen][int(start_at) - 1:]
        if max_sensors:
            selected = selected[:int(max_sensors)]

        st.write(f"This run will collect **{len(selected)}** sensor(s).")

        if st.button("Start collection", type="primary", disabled=not selected):
            try:
                session = ensure_session(site_url, login_url, username, password)
                st.session_state["result"] = run_collection(session, selected, delay)
            except AuthError as exc:
                st.error(f"Authentication problem: {exc}")
            except Exception as exc:
                st.error(f"Collection failed: {exc}")

    res = st.session_state.get("result")
    if res:
        if res["summary"].empty:
            st.error("No CSV reports were downloaded.")
        else:
            st.success(
                f"{len(res['summary'])} report(s) from {res['sensors']} sensor(s), "
                f"{len(res['master_df']):,} rows in total."
            )
            st.dataframe(res["summary"], use_container_width=True, hide_index=True)
            st.dataframe(res["master_df"].head(50), use_container_width=True, hide_index=True)

            d1, d2, d3 = st.columns(3)
            with d1:
                if res["master_xlsx"]:
                    st.download_button(
                        "MASTER Excel (all sensors)", res["master_xlsx"],
                        file_name="worldsensing_master.xlsx", mime=XLSX_MIME,
                        type="primary",
                    )
                else:
                    st.warning("Too many rows for Excel - use the CSV.")
            with d2:
                st.download_button(
                    "MASTER CSV", res["master_csv"],
                    file_name="worldsensing_master.csv", mime="text/csv",
                )
            with d3:
                st.download_button(
                    "Everything (ZIP)", res["zip"],
                    file_name="worldsensing_collection.zip", mime="application/zip",
                )

        if res["errors"]:
            st.warning(f"{len(res['errors'])} issue(s):")
            st.code("\n".join(res["errors"][:60]))

# ---------------- Tab 3 ----------------
with tab_test:
    st.subheader("Test a direct readings CSV link")

    csv_url = st.text_input(
        "Readings CSV link",
        value=(
            "https://loadsensing.wocs3.com/33998/dataserver/current/reading/"
            "177340/ppvSufre/177340-LSG7ACLBXLHVIB-1-readings-VIBPPV-current.csv"
        ),
    )

    if st.button("Download and test"):
        try:
            session = ensure_session(site_url, login_url, username, password)
            content = download_csv(session, csv_url)
            df = parse_csv(content)

            filename = Path(urlparse(csv_url).path).name
            meta = parse_filename(filename)
            m = re.search(r"/reading/(\d+)/", csv_url)
            sid = m.group(1) if m else meta["Serial_Number"]
            enriched = enrich_frame(df, f"Sensor_{sid}", sid, filename)

            st.session_state["test_result"] = {
                "df": enriched,
                "meta": meta,
                "rows": len(df),
                "cols": len(df.columns),
                "xlsx": sensor_workbook([(filename, enriched)]),
                "sid": sid,
            }
        except AuthError as exc:
            st.session_state.pop("test_result", None)
            st.error(f"Authentication problem: {exc}")
        except ValueError as exc:
            st.session_state.pop("test_result", None)
            st.error(f"CSV problem: {exc}")
        except Exception as exc:
            st.session_state.pop("test_result", None)
            st.error(f"Test failed: {exc}")

    t = st.session_state.get("test_result")
    if t:
        st.success(f"Loaded {t['rows']} rows, {t['cols']} columns.")
        st.write("Detected from filename")
        st.json(t["meta"])
        st.write("Columns")
        st.code(", ".join(map(str, t["df"].columns)))
        st.dataframe(t["df"].head(20), use_container_width=True)
        st.download_button(
            "Download test Excel", t["xlsx"],
            file_name=f"Sensor_{t['sid']}.xlsx", mime=XLSX_MIME,
        )
