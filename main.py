import argparse
import pathlib
from typing import Dict, List, Optional
from datetime import datetime, date

import bs4 as bs
import pandas as pd
import requests
import meraki
from fpdf import FPDF

import config  # must contain api_key = "..."


EOL_URL = (
    "https://documentation.meraki.com/General_Administration/Other_Topics/"
    "Meraki_End-of-Life_(EOL)_Products_and_Dates"
)


def normalize_product_to_key(s: str) -> str:
    """
    Normalize EoL 'Product' strings from the Meraki EoL page so they line up
    with Meraki inventory 'model' values.
    """
    if not isinstance(s, str):
        return ""
    s = s.strip()
    s = s.split(",")[0].strip()        # first SKU before comma
    s = s.split(" ")[0]                # drop anything after first space
    if "-HW" in s:
        s = s.split("-HW")[0]          # strip -HW suffix
    return s.upper()


def fetch_eol_table() -> pd.DataFrame:
    """
    Fetch the Meraki EoL table, attach upgrade-path URLs,
    expand multi-SKU rows, and compute a normalized ProductKey.
    """
    dfs = pd.read_html(EOL_URL)
    if not dfs:
        raise RuntimeError("No tables found on EoL page; layout may have changed.")

    eol_df = dfs[0].copy()

    if "Product" not in eol_df.columns:
        raise RuntimeError("EoL table has no 'Product' column; parser needs updating.")

    # Scrape links (upgrade path info, as clean URLs)
    resp = requests.get(EOL_URL)
    resp.raise_for_status()
    soup = bs.BeautifulSoup(resp.text, "html.parser")
    table = soup.find("table")

    if table is None:
        raise RuntimeError("Could not find table element on EoL page; layout may have changed.")

    links: List[List[str]] = []
    for row in table.find_all("tr"):
        for td in row.find_all("td"):
            sublinks: List[str] = []
            for a in td.find_all("a"):
                href = a.get("href")
                if href:
                    sublinks.append(href)
            if sublinks:
                links.append(sublinks)

    # Attach links (pad/truncate if counts differ)
    if len(links) == len(eol_df):
        eol_df["Upgrade Path"] = links
    else:
        padded_links = links[: len(eol_df)]
        while len(padded_links) < len(eol_df):
            padded_links.append([])
        eol_df["Upgrade Path"] = padded_links

    # Expand multi-SKU rows
    expanded_rows = []
    for _, row in eol_df.iterrows():
        product_raw = str(row["Product"])
        parts = [p.strip() for p in product_raw.split(",") if p.strip()]
        if not parts:
            expanded_rows.append(row)
            continue

        for part in parts:
            new_row = row.copy()
            new_row["Product"] = part
            expanded_rows.append(new_row)

    expanded_df = pd.DataFrame(expanded_rows)

    # Normalize ProductKey
    expanded_df["ProductKey"] = expanded_df["Product"].apply(normalize_product_to_key)
    expanded_df = expanded_df[expanded_df["ProductKey"] != ""]

    # Convert Upgrade Path from list -> semicolon-separated URL string for CSV
    if "Upgrade Path" in expanded_df.columns:
        expanded_df["Upgrade Path"] = expanded_df["Upgrade Path"].apply(
            lambda v: "; ".join(v) if isinstance(v, list) else ""
        )

    if expanded_df.empty:
        raise RuntimeError("All EoL rows collapsed after normalization; parser likely broken.")

    return expanded_df


def choose_orgs_interactively(orgs: List[Dict]) -> List[Dict]:
    """
    Print orgs and prompt user to select one or more by index, or 'all'.
    """
    print("Your API Key has access to the following organizations:")
    for i, org in enumerate(orgs, start=1):
        print(f"{i} - {org['name']}")

    choice = input(
        "Type the number(s) of the org(s) to include (e.g. 1,3,5) "
        "or 'all' to include every organization: "
    ).strip()

    if not choice:
        raise SystemExit("No organizations selected; exiting.")

    if choice.lower() == "all":
        return orgs

    try:
        indices = [int(x) - 1 for x in choice.split(",")]
    except ValueError as exc:
        raise SystemExit(f"Invalid input for org selection: {choice}") from exc

    selected: List[Dict] = []
    for idx in indices:
        if idx < 0 or idx >= len(orgs):
            raise SystemExit(f"Org index {idx+1} is out of range.")
        selected.append(orgs[idx])

    return selected


def fetch_inventories(
    dashboard: meraki.DashboardAPI, orgs: List[Dict]
) -> Dict[str, List[Dict]]:
    """
    Fetch inventory devices for each selected organization.
    Returns dict: { "OrgName - OrgId": [device, ...], ... }.
    """
    inventories: Dict[str, List[Dict]] = {}
    for org in orgs:
        org_name = org["name"]
        org_id = org["id"]
        label = f"{org_name} - {org_id}"
        print(f"Fetching inventory for {label} ...")
        devices = dashboard.organizations.getOrganizationInventoryDevices(org_id)
        inventories[label] = devices
    return inventories


def build_eol_reports(
    eol_df: pd.DataFrame, inventories: Dict[str, List[Dict]]
) -> List[Dict]:
    """
    Build EoL reports for each org.
    Returns list of dicts: [{ "name": label, "report": DataFrame }, ...].

    NOTE: All orgs appear, even if report is empty (no EoL devices).
    """
    reports: List[Dict] = []

    for label, devices in inventories.items():
        inv_df = pd.DataFrame(devices)

        if inv_df.empty:
            reports.append({"name": label, "report": pd.DataFrame()})
            continue

        # Only consider devices assigned to a network (in use)
        inv_assigned = inv_df.loc[~inv_df["networkId"].isna()].copy()
        if inv_assigned.empty:
            reports.append({"name": label, "report": pd.DataFrame()})
            continue

        inv_assigned["ModelKey"] = inv_assigned["model"].astype(str).str.upper()

        counts = inv_assigned["ModelKey"].value_counts()

        final_eol = eol_df.copy()
        final_eol["Total Units"] = final_eol["ProductKey"].map(counts)

        # Keep only products where we actually own at least one unit
        final_eol = final_eol.dropna(subset=["Total Units"])

        if not final_eol.empty:
            final_eol = (
                final_eol.sort_values(by=["Total Units"], ascending=False)
                .reset_index(drop=True)
            )

        reports.append({"name": label, "report": final_eol})

    return reports


class ReportPDF(FPDF):
    pass  # using explicit layout in code; no global header/footer


def _find_logo(current_dir: pathlib.Path) -> pathlib.Path:
    """
    Try to find the Meraki logo in a couple of common locations.
    """
    candidates = [
        current_dir / "cisco-meraki-logo.png",
        current_dir / "images" / "cisco-meraki-logo.png",
    ]
    for c in candidates:
        if c.is_file():
            return c
    return candidates[0]  # default location for error message


def _parse_end_of_support(value) -> Optional[date]:
    """
    Best-effort parse for End-of-Support Date values from the EoL table.
    Returns a date or None if parsing fails.
    """
    if hasattr(value, "to_pydatetime"):
        try:
            return value.to_pydatetime().date()
        except Exception:
            pass

    if isinstance(value, date):
        return value

    if isinstance(value, datetime):
        return value.date()

    s = str(value).strip()
    if not s:
        return None

    for fmt in ("%b %d, %Y", "%B %d, %Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue

    return None


def generate_pdf(
    reports: List[Dict],
    output_pdf: pathlib.Path,
    logo_path: pathlib.Path,
) -> None:
    """
    Generate a PDF with one section per organization using fpdf2,
    mimicking the green/white Meraki table branding and layout.

    - Single initial page + auto page breaks → multiple orgs per page
    - Orgs alphabetized
    - Heading like "ORG - 123456"
    - Fixed column widths (uniform across all tables) that span full page width
    - Headers shortened: End-of-Sale, End-of-Support, Units
    - Upgrade Path rendered as blue, underlined clickable [Product] link
    - Rows highlighted:
        * RED   if End-of-Support is in the past
        * YELLOW if End-of-Support is within the next 365 days
    """
    # Sort orgs alphabetically by label (name - id)
    reports_sorted = sorted(reports, key=lambda r: r["name"].lower())
    today = date.today()

    pdf = ReportPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_margins(15, 15, 15)

    pdf.add_page()

    # Logo once at the top of first page
    if logo_path.is_file():
        try:
            pdf.image(str(logo_path), w=120)
            pdf.ln(8)
        except Exception as e:
            print(f"Warning: failed to load logo image '{logo_path}': {e}")
    else:
        print(f"Warning: logo image not found at '{logo_path}'")

    # Title & intro once – match original text more closely
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "Cisco Meraki Lifecycle Report")
    pdf.ln(10)

    pdf.set_font("Helvetica", "", 11)
    intro = (
        "This report lists all of your equipment currently in use that has an "
        "end of life announcement. They are ordered by the total units column, "
        "and the Upgrade Path column links you to the EoS announcement with "
        "recommendations on upgrade paths."
    )
    pdf.multi_cell(0, 6, intro)
    pdf.ln(4)

    # Underlying data columns and display labels
    pdf_cols = [
        ("Product", "Product"),
        ("Announcement", "Announcement"),
        ("End-of-Sale Date", "End-of-Sale"),
        ("End-of-Support Date", "End-of-Support"),
        ("ProductKey", "ProductKey"),
        ("Upgrade Path", "Upgrade Path"),
        ("Total Units", "Units"),
    ]
    display_labels = {data_col: label for data_col, label in pdf_cols}

    # Fixed relative widths – balanced version
    col_width_weights = {
        "Product": 16,
        "Announcement": 20,
        "End-of-Sale Date": 18,
        "End-of-Support Date": 24,
        "ProductKey": 12,
        "Upgrade Path": 18,
        "Total Units": 8,
    }

    for idx, item in enumerate(reports_sorted):
        org_name = item["name"]
        df = item["report"]

        # Spacing before each org, but no forced page break
        if idx > 0:
            pdf.ln(4)

        # Org heading – "ORG - 123456"
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 8, org_name)
        pdf.ln(8)

        if df is None or df.empty:
            pdf.set_font("Helvetica", "", 11)
            pdf.multi_cell(0, 6, "No EoL devices found for this organization.")
            continue

        # Choose columns in desired order if present in df
        available_cols = [col for col, _ in pdf_cols if col in df.columns]
        if not available_cols:
            pdf.set_font("Helvetica", "", 11)
            pdf.multi_cell(0, 6, "No displayable columns in EoL data.")
            continue

        # Compute fixed column widths from weights
        page_width = pdf.w - 2 * pdf.l_margin
        total_weight = sum(col_width_weights.get(c, 10) for c in available_cols)
        col_widths: Dict[str, float] = {
            c: page_width * col_width_weights.get(c, 10) / total_weight
            for c in available_cols
        }

        # Header row: green background, white text
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_fill_color(4, 170, 109)  # #04AA6D
        pdf.set_text_color(255, 255, 255)
        for col in available_cols:
            header_label = display_labels.get(col, col)
            pdf.cell(col_widths[col], 8, header_label, border=1, align="L", fill=True)
        pdf.ln(8)

        # Reset text color for body
        pdf.set_font("Helvetica", "", 8)
        pdf.set_text_color(0, 0, 0)

        # Table rows, with risk-based highlighting
        fill_toggle = False
        for _, row in df[available_cols].iterrows():
            # Determine highlight color based on End-of-Support Date
            eos_raw = row.get("End-of-Support Date", "")
            eos_date = _parse_end_of_support(eos_raw)
            highlight = None  # None, "red", "yellow"

            if eos_date is not None:
                days_left = (eos_date - today).days
                if days_left < 0:
                    highlight = "red"
                elif days_left <= 365:
                    highlight = "yellow"

            if highlight == "red":
                # Light red
                pdf.set_fill_color(255, 204, 204)
            elif highlight == "yellow":
                # Light yellow
                pdf.set_fill_color(255, 255, 204)
            else:
                # Normal alternating zebra stripes
                if fill_toggle:
                    pdf.set_fill_color(222, 220, 220)  # #dedcdc
                else:
                    pdf.set_fill_color(255, 255, 255)

            for col in available_cols:
                w = col_widths[col]

                if col == "Upgrade Path":
                    urls_str = str(row[col])
                    urls = [u.strip() for u in urls_str.split(";") if u.strip()]
                    if urls:
                        display = "[%s]" % row.get("Product", "Link")
                        # Blue, underlined clickable link
                        pdf.set_text_color(0, 0, 255)
                        pdf.set_font("Helvetica", "U", 8)
                        pdf.cell(
                            w,
                            6,
                            display,
                            border=1,
                            align="L",
                            fill=True,
                            link=urls[0],
                        )
                        # Reset font/color
                        pdf.set_text_color(0, 0, 0)
                        pdf.set_font("Helvetica", "", 8)
                    else:
                        pdf.cell(w, 6, "", border=1, align="L", fill=True)
                else:
                    text = str(row[col])
                    pdf.cell(w, 6, text, border=1, align="L", fill=True)

            pdf.ln(6)
            fill_toggle = not fill_toggle

    pdf.output(str(output_pdf))


def generate_csv(reports: List[Dict], output_csv: pathlib.Path) -> None:
    """
    Generate a single CSV combining all org reports, with an Organization column.
    Includes orgs with no EoL devices as placeholder rows.
    Sorted by Organization, then Total Units descending.
    """
    # Sort orgs alphabetically first (for predictable placeholder order)
    reports_sorted = sorted(reports, key=lambda r: r["name"].lower())

    # First collect all columns that appear in any report
    all_columns = set()
    for item in reports_sorted:
        df = item["report"]
        if df is not None and not df.empty:
            all_columns.update(df.columns)

    base_cols = list(all_columns) if all_columns else []
    # Ensure Total Units is present if we expect it
    if "Total Units" not in base_cols:
        base_cols.append("Total Units")
    # Remove Organization/Note from base; we'll add explicitly
    base_cols = [c for c in base_cols if c not in ("Organization", "Note")]

    # Final column order (Upgrade Path is already plain URLs string)
    columns = ["Organization"] + base_cols + ["Note"]

    all_rows: List[Dict] = []

    for item in reports_sorted:
        org_name = item["name"]
        df = item["report"]

        if df is None or df.empty:
            # Placeholder row for org with no EoL devices
            row = {col: "" for col in columns}
            row["Organization"] = org_name
            row["Total Units"] = 0
            row["Note"] = "No EoL devices found"
            all_rows.append(row)
        else:
            tmp = df.copy()
            tmp.insert(0, "Organization", org_name)
            if "Note" not in tmp.columns:
                tmp["Note"] = ""
            tmp = tmp.reindex(columns=columns, fill_value="")
            all_rows.extend(tmp.to_dict(orient="records"))

    if not all_rows:
        pd.DataFrame(columns=columns).to_csv(output_csv, index=False)
        return

    combined = pd.DataFrame(all_rows)

    # Sort by Organization then Total Units (descending)
    if "Total Units" in combined.columns:
        combined["Total Units"] = pd.to_numeric(
            combined["Total Units"], errors="coerce"
        ).fillna(0)
        combined = combined.sort_values(
            by=["Organization", "Total Units"],
            ascending=[True, False],
            ignore_index=True,
        )
    else:
        combined = combined.sort_values(
            by=["Organization"], ascending=True, ignore_index=True
        )

    combined.to_csv(output_csv, index=False)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate Cisco Meraki Lifecycle (EoL) report (PDF/CSV)."
    )
    parser.add_argument(
        "--no-pdf",
        action="store_true",
        help="Skip generating PDF (only CSV).",
    )
    parser.add_argument(
        "--no-csv",
        action="store_true",
        help="Skip generating CSV (only PDF).",
    )
    parser.add_argument(
        "--output-prefix",
        default="Lifecycle Report",
        help="Base name for output files (default: 'Lifecycle Report').",
    )
    args = parser.parse_args()

    if args.no_pdf and args.no_csv:
        raise SystemExit("Both --no-pdf and --no-csv were specified; nothing to do.")

    # Initialize Meraki Dashboard API client (read-only usage)
    dashboard = meraki.DashboardAPI(
        api_key=config.api_key,
        print_console=True,
    )

    # Fetch orgs and let the user choose (supports 'all')
    orgs = dashboard.organizations.getOrganizations()
    selected_orgs = choose_orgs_interactively(orgs)

    # Fetch EoL data and inventories
    eol_df = fetch_eol_table()
    inventories = fetch_inventories(dashboard, selected_orgs)

    # Build per-org reports (includes orgs with empty reports)
    reports = build_eol_reports(eol_df, inventories)

    current_dir = pathlib.Path().absolute()
    pdf_path = current_dir / f"{args.output_prefix}.pdf"
    csv_path = current_dir / f"{args.output_prefix}.csv"
    logo_path = _find_logo(current_dir)

    # PDF
    if not args.no_pdf:
        print("Generating PDF: %s" % pdf_path)
        generate_pdf(reports, pdf_path, logo_path)
        print("PDF report written to: %s" % pdf_path)

    # CSV
    if not args.no_csv:
        print("Generating CSV: %s" % csv_path)
        generate_csv(reports, csv_path)
        print("CSV report written to: %s" % csv_path)


if __name__ == "__main__":
    main()

