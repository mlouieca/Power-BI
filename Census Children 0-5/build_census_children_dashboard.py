import csv
import html
import json
import math
import os
import re
import shutil
import stat
import struct
import uuid
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent
SOURCE_XLSX = Path(r"C:\Users\MLOUI\OneDrive\Work\Census Data DDP\Demographic_Census_2021_Children0_5.xlsx")
EMAIL_FILES = [
    Path(r"C:\Users\MLOUI\Downloads\FW_ Census information for children aged 0-5.msg"),
    Path(r"C:\Users\MLOUI\Downloads\FW_ Demographic information.msg"),
]

PROJECT_NAME = "Census Children 0-5"
REPORT_DIR = ROOT / f"{PROJECT_NAME}.Report"
MODEL_DIR = ROOT / f"{PROJECT_NAME}.SemanticModel"
DATA_DIR = ROOT / "Data"
CONTEXT_DIR = ROOT / "Context"

INDICATORS = [
    ("Below LIM", "Low income", "Children aged 0 to 5 living below LIM", 1),
    ("Indigenous Identity", "Indigenous identity", "Children aged 0 to 5 with Indigenous identity", 2),
    ("Parent(s) Less Than High School", "Parent education", "Children aged 0 to 5 living with parent(s) with less than high school education", 3),
    ("Immigrant Parent(s)", "Immigrant parent(s)", "Children aged 0 to 5 living with immigrant parent(s)", 4),
    ("Lone Parent", "Lone parent", "Children aged 0 to 5 living with a lone parent", 5),
]

PROVINCE_CODES = {
    10: "Newfoundland and Labrador",
    11: "Prince Edward Island",
    12: "Nova Scotia",
    13: "New Brunswick",
    24: "Quebec",
    35: "Ontario",
    46: "Manitoba",
    47: "Saskatchewan",
    48: "Alberta",
    59: "British Columbia",
    60: "Yukon",
    61: "Northwest Territories",
    62: "Nunavut",
}

MIZ_LABELS = {
    1: "Inside CMA",
    2: "Inside tracted CA",
    3: "Inside non-tracted CA",
    4: "Strong MIZ",
    5: "Moderate MIZ",
    6: "Weak MIZ",
    7: "No MIZ",
    8: "Territories",
}


def child_population_band(children):
    if children is None:
        return "Unknown / suppressed"
    if children >= 1000:
        return "1,000+ children"
    if children >= 250:
        return "250 to 999 children"
    if children >= 50:
        return "50 to 249 children"
    return "Under 50 children"


def child_population_band_sort(children):
    if children is None:
        return 5
    if children >= 1000:
        return 1
    if children >= 250:
        return 2
    if children >= 50:
        return 3
    return 4


def miz_analysis_label(miz_id, miz_label):
    if miz_id == 8:
        return "Territories outside CA"
    return miz_label


def geography_type(miz_id):
    if miz_id == 1:
        return "CMA"
    if miz_id in (2, 3):
        return "CA"
    if miz_id in (4, 5, 6, 7):
        return "MIZ"
    if miz_id == 8:
        return "Territories outside CA"
    return "Unknown"


def clean_text(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return str(value).strip()


def number_value(value, as_int=False):
    text = clean_text(value)
    if not text:
        return None, "Missing"
    lowered = text.lower()
    if lowered == "x":
        return None, "Suppressed"
    if text == "-":
        return None, "Not applicable"
    try:
        num = float(text.replace(",", ""))
    except ValueError:
        return None, "Missing"
    if as_int:
        return int(round(num)), "Reportable"
    return num, "Reportable"


def write_csv(path, rows, fieldnames):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def existing_logical_id(path):
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8")).get("config", {}).get("logicalId") or str(uuid.uuid4())
        except (OSError, json.JSONDecodeError):
            return str(uuid.uuid4())
    return str(uuid.uuid4())


def parse_workbook():
    prov_raw = pd.read_excel(SOURCE_XLSX, sheet_name="Provinces", header=None, dtype=object)
    csd_raw = pd.read_excel(SOURCE_XLSX, sheet_name="CSD level", header=None, dtype=object)
    notes_raw = pd.read_excel(SOURCE_XLSX, sheet_name="NOTES", header=None, dtype=object)

    province_rows = []
    province_indicator_rows = []
    province_code_by_name = {name: code for code, name in PROVINCE_CODES.items()}
    province_code_by_name["Canada"] = 0

    prov_cols = {
        "Below LIM": (4, 5),
        "Indigenous Identity": (7, 8),
        "Parent(s) Less Than High School": (10, 11),
        "Immigrant Parent(s)": (13, 14),
        "Lone Parent": (16, 17),
    }
    for _, row in prov_raw.iloc[4:].iterrows():
        province_name = clean_text(row[0])
        if not province_name or province_name.startswith("Source:"):
            continue
        population, pop_status = number_value(row[1], as_int=True)
        children, child_status = number_value(row[2], as_int=True)
        province_code = province_code_by_name.get(province_name)
        province_rows.append({
            "Province Code": province_code,
            "Province Name": province_name,
            "Is Canada": "Yes" if province_name == "Canada" else "No",
            "Population in Census Families": population,
            "Children Aged 0 to 5": children,
            "Population Status": pop_status,
            "Children Status": child_status,
        })
        for indicator, _, _, sort_order in INDICATORS:
            count_col, pct_col = prov_cols[indicator]
            count, count_status = number_value(row[count_col], as_int=True)
            pct, pct_status = number_value(row[pct_col], as_int=False)
            status = count_status if count_status != "Reportable" else pct_status
            if count_status == "Reportable":
                status = "Reportable"
            province_indicator_rows.append({
                "Province Code": province_code,
                "Province Name": province_name,
                "Indicator": indicator,
                "Indicator Sort": sort_order,
                "Children Count": count,
                "Percent": pct,
                "Value Status": status,
            })

    csd_rows = []
    csd_indicator_rows = []
    csd_cols = {
        "Below LIM": (6, 7),
        "Indigenous Identity": (9, 10),
        "Parent(s) Less Than High School": (12, 13),
        "Immigrant Parent(s)": (15, 16),
        "Lone Parent": (18, 19),
    }
    for _, row in csd_raw.iloc[4:].iterrows():
        csd_number = clean_text(row[0])
        if not csd_number or csd_number.startswith("Source:"):
            continue
        csd_name = clean_text(row[1])
        province_code, province_status = number_value(row[2], as_int=True)
        miz_id, miz_status = number_value(row[3], as_int=True)
        population, pop_status = number_value(row[4], as_int=True)
        children, child_status = number_value(row[5], as_int=True)
        province_name = PROVINCE_CODES.get(province_code, f"Province {province_code}" if province_code is not None else "")
        miz_label = MIZ_LABELS.get(miz_id, f"MIZ {miz_id}" if miz_id is not None else "")
        band = child_population_band(children)
        csd_rows.append({
            "CSD Number": csd_number,
            "CSD Name": csd_name,
            "Province Code": province_code,
            "Province Name": province_name,
            "MIZ Identifier": miz_id,
            "MIZ Label": miz_label,
            "MIZ Analysis Label": miz_analysis_label(miz_id, miz_label),
            "Geography Type": geography_type(miz_id),
            "Population in Census Families": population,
            "Children Aged 0 to 5": children,
            "Child Population Band": band,
            "Child Population Band Sort": child_population_band_sort(children),
            "Public Denominator Flag": "1,000+ children" if children is not None and children >= 1000 else "Under 1,000 children",
            "Population Status": pop_status,
            "Children Status": child_status,
            "MIZ Status": miz_status,
            "Province Status": province_status,
        })
        for indicator, _, _, sort_order in INDICATORS:
            count_col, pct_col = csd_cols[indicator]
            count, count_status = number_value(row[count_col], as_int=True)
            pct, pct_status = number_value(row[pct_col], as_int=False)
            status = count_status if count_status != "Reportable" else pct_status
            if count_status == "Reportable":
                status = "Reportable"
            csd_indicator_rows.append({
                "CSD Number": csd_number,
                "Province Code": province_code,
                "Indicator": indicator,
                "Indicator Sort": sort_order,
                "Children Count": count,
                "Percent": pct,
                "Value Status": status,
            })

    indicator_rows = [
        {
            "Indicator": indicator,
            "Indicator Short Name": short_name,
            "Indicator Description": description,
            "Indicator Sort": sort_order,
        }
        for indicator, short_name, description, sort_order in INDICATORS
    ]
    notes_rows = []
    for idx, value in enumerate(notes_raw.iloc[:, 0].tolist(), start=1):
        note = clean_text(value)
        if note:
            notes_rows.append({"Note Number": idx, "Note": note})

    csd_children_by_miz = {}
    for row in csd_rows:
        key = (
            row["MIZ Identifier"],
            row["MIZ Analysis Label"],
            row["MIZ Label"],
            row["Geography Type"],
        )
        csd_children_by_miz[key] = csd_children_by_miz.get(key, 0) + (row["Children Aged 0 to 5"] or 0)

    csd_lookup = {row["CSD Number"]: row for row in csd_rows}
    miz_indicator_counts = {}
    for row in csd_indicator_rows:
        csd = csd_lookup.get(row["CSD Number"])
        if not csd:
            continue
        key = (
            csd["MIZ Identifier"],
            csd["MIZ Analysis Label"],
            csd["MIZ Label"],
            csd["Geography Type"],
            row["Indicator"],
            row["Indicator Sort"],
        )
        miz_indicator_counts[key] = miz_indicator_counts.get(key, 0) + (row["Children Count"] or 0)

    miz_summary_rows = []
    for key, children in csd_children_by_miz.items():
        miz_id, analysis_label, base_label, geo_type = key
        for indicator, short_name, _, sort_order in INDICATORS:
            count_key = key + (indicator, sort_order)
            count = miz_indicator_counts.get(count_key, 0)
            miz_summary_rows.append({
                "MIZ Identifier": miz_id,
                "MIZ Analysis Label": analysis_label,
                "MIZ Label": base_label,
                "Geography Type": geo_type,
                "Indicator": indicator,
                "Indicator Short Name": short_name,
                "Indicator Sort": sort_order,
                "Children Aged 0 to 5": children,
                "Children Count": count,
                "Percent": count / children if children else None,
            })

    return {
        "province": province_rows,
        "province_indicators": province_indicator_rows,
        "csd": csd_rows,
        "csd_indicators": csd_indicator_rows,
        "miz_summary": miz_summary_rows,
        "indicator": indicator_rows,
        "notes": notes_rows,
    }


class MsgReader:
    ENDOFCHAIN = 0xFFFFFFFE
    FREESECT = 0xFFFFFFFF

    def __init__(self, path):
        self.path = path
        self.data = path.read_bytes()
        if self.data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            raise ValueError("Not a compound file")
        self.sector_size = 1 << struct.unpack_from("<H", self.data, 30)[0]
        self.mini_sector_size = 1 << struct.unpack_from("<H", self.data, 32)[0]
        self.num_fat_sectors = struct.unpack_from("<I", self.data, 44)[0]
        self.dir_start = struct.unpack_from("<I", self.data, 48)[0]
        self.mini_cutoff = struct.unpack_from("<I", self.data, 56)[0]
        self.minifat_start = struct.unpack_from("<I", self.data, 60)[0]
        self.num_minifat_sectors = struct.unpack_from("<I", self.data, 64)[0]
        self.difat_start = struct.unpack_from("<I", self.data, 68)[0]
        self.num_difat_sectors = struct.unpack_from("<I", self.data, 72)[0]
        self.fat = self._load_fat()
        self.dir_entries = self._load_directory()
        self.root = next((e for e in self.dir_entries if e["type"] == 5), None)
        self.minifat = self._load_minifat()
        self.mini_stream = self._read_regular_stream(self.root["start"], self.root["size"]) if self.root else b""

    def _sector_offset(self, sector):
        return 512 + sector * self.sector_size

    def _sector(self, sector):
        start = self._sector_offset(sector)
        return self.data[start:start + self.sector_size]

    def _chain(self, start, fat=None):
        if start in (self.FREESECT, self.ENDOFCHAIN):
            return []
        table = self.fat if fat is None else fat
        seen = set()
        out = []
        sector = start
        while sector not in (self.FREESECT, self.ENDOFCHAIN) and sector < len(table) and sector not in seen:
            seen.add(sector)
            out.append(sector)
            sector = table[sector]
        return out

    def _load_fat(self):
        difat = list(struct.unpack_from("<109I", self.data, 76))
        sector = self.difat_start
        for _ in range(self.num_difat_sectors):
            block = self._sector(sector)
            entries = list(struct.unpack("<" + "I" * (self.sector_size // 4), block))
            difat.extend(entries[:-1])
            sector = entries[-1]
            if sector == self.ENDOFCHAIN:
                break
        difat = [s for s in difat if s not in (self.FREESECT, self.ENDOFCHAIN)][:self.num_fat_sectors]
        fat = []
        for fat_sector in difat:
            block = self._sector(fat_sector)
            fat.extend(struct.unpack("<" + "I" * (self.sector_size // 4), block))
        return fat

    def _read_regular_stream(self, start, size):
        chunks = [self._sector(s) for s in self._chain(start)]
        return b"".join(chunks)[:size]

    def _load_minifat(self):
        chunks = [self._sector(s) for s in self._chain(self.minifat_start)]
        raw = b"".join(chunks)[: self.num_minifat_sectors * self.sector_size]
        if not raw:
            return []
        return list(struct.unpack("<" + "I" * (len(raw) // 4), raw))

    def _read_mini_stream(self, start, size):
        out = []
        for sector in self._chain(start, self.minifat):
            offset = sector * self.mini_sector_size
            out.append(self.mini_stream[offset:offset + self.mini_sector_size])
        return b"".join(out)[:size]

    def _load_directory(self):
        raw = self._read_regular_stream(self.dir_start, len(self._chain(self.dir_start)) * self.sector_size)
        entries = []
        for i in range(0, len(raw), 128):
            entry = raw[i:i + 128]
            if len(entry) < 128:
                continue
            name_len = struct.unpack_from("<H", entry, 64)[0]
            if name_len >= 2:
                name = entry[:name_len - 2].decode("utf-16le", errors="ignore")
            else:
                name = ""
            entries.append({
                "name": name,
                "type": entry[66],
                "left": struct.unpack_from("<I", entry, 68)[0],
                "right": struct.unpack_from("<I", entry, 72)[0],
                "child": struct.unpack_from("<I", entry, 76)[0],
                "start": struct.unpack_from("<I", entry, 116)[0],
                "size": struct.unpack_from("<Q", entry, 120)[0],
            })
        return entries

    def _read_stream_entry(self, entry):
        if entry["size"] < self.mini_cutoff and self.minifat:
            return self._read_mini_stream(entry["start"], entry["size"])
        return self._read_regular_stream(entry["start"], entry["size"])

    def stream_values(self):
        values = {}
        for entry in self.dir_entries:
            name = entry["name"]
            if entry["type"] != 2 or not name.startswith("__substg1.0_"):
                continue
            raw = self._read_stream_entry(entry)
            prop = name.replace("__substg1.0_", "")
            if prop.endswith("001F"):
                text = raw.decode("utf-16le", errors="ignore").rstrip("\x00")
            elif prop.endswith("001E"):
                text = raw.decode("cp1252", errors="ignore").rstrip("\x00")
            else:
                continue
            values[prop] = text
        return values

    def named_stream(self, stream_name):
        for entry in self.dir_entries:
            if entry["type"] == 2 and entry["name"] == stream_name:
                return self._read_stream_entry(entry)
        return b""


def html_to_text(raw):
    if not raw:
        return ""
    head = raw[:1000].decode("ascii", errors="ignore").lower()
    charset = "utf-8"
    match = re.search(r"charset=([a-z0-9_-]+)", head)
    if match:
        charset = match.group(1)
    charset_text = raw.decode(charset, errors="ignore")
    utf8_text = raw.decode("utf-8", errors="ignore")
    text = utf8_text if utf8_text.count("Ã") < charset_text.count("Ã") else charset_text
    text = re.sub(r"(?is)<(script|style).*?</\1>", " ", text)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</p\s*>", "\n", text)
    text = re.sub(r"(?s)<[^>]+>", " ", text)
    text = html.unescape(text)
    text = re.sub(r"[ \t\r\f\v]+", " ", text)
    text = re.sub(r"\n\s+", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_email_context():
    rows = []
    for path in EMAIL_FILES:
        item = {"file": path.name, "subject": "", "sender": "", "to": "", "body": ""}
        try:
            reader = MsgReader(path)
            values = reader.stream_values()
            item["subject"] = values.get("0037001F") or values.get("0037001E", "")
            item["sender"] = values.get("0C1A001F") or values.get("0C1A001E", "")
            item["to"] = values.get("0E04001F") or values.get("0E04001E", "")
            body = values.get("1000001F") or values.get("1000001E", "")
            if not body:
                body = html_to_text(reader.named_stream("__substg1.0_10130102"))
            item["body"] = "\n".join(line.rstrip() for line in body.splitlines() if line.strip())
        except Exception as exc:
            item["body"] = f"Could not extract message body: {exc}"
        rows.append(item)
    return rows


def qname(name):
    if any(ch in name for ch in " .'-()/"):
        return "'" + name.replace("'", "''") + "'"
    return name


def tmdl_column(name, data_type, source=None, hidden=False, fmt=None, summarize="none", sort_by=None):
    source = source or name
    lines = [f"\tcolumn {qname(name)}", f"\t\tdataType: {data_type}"]
    if fmt:
        lines.append(f"\t\tformatString: {fmt}")
    if hidden:
        lines.append("\t\tisHidden")
    if sort_by:
        lines.append(f"\t\tsortByColumn: {qname(sort_by)}")
    lines.append(f"\t\tsummarizeBy: {summarize}")
    lines.append(f"\t\tsourceColumn: {source}")
    return "\n".join(lines)


def tmdl_measure(name, expression, fmt="#,##0", folder=None, description=None):
    lines = []
    if description:
        lines.append(f"\t/// {description}")
    lines.append(f"\tmeasure {qname(name)} = ```")
    for line in expression.strip().splitlines():
        lines.append(f"\t\t\t{line}")
    lines.append("\t\t\t```")
    lines.append(f"\t\tformatString: {fmt}")
    if folder:
        lines.append(f"\t\tdisplayFolder: {folder}")
    return "\n".join(lines)


def m_csv_partition(table_name, file_name, columns, types):
    type_pairs = ", ".join([f'{{"{col}", {typ}}}' for col, typ in types])
    return f"""
\tpartition {qname(table_name)} = m
\t\tmode: import
\t\tsource =
\t\t\tlet
\t\t\t    Source = Csv.Document(File.Contents(#"DataFolder" & "{file_name}"), [Delimiter=",", Columns={columns}, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
\t\t\t    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
\t\t\t    #"Blank Values As Null" = Table.ReplaceValue(#"Promoted Headers", "", null, Replacer.ReplaceValue, Table.ColumnNames(#"Promoted Headers")),
\t\t\t    #"Changed Type" = Table.TransformColumnTypes(#"Blank Values As Null", {{{type_pairs}}}, "en-CA")
\t\t\tin
\t\t\t    #"Changed Type"
""".rstrip()


def make_semantic_model():
    definition = MODEL_DIR / "definition"
    tables_dir = definition / "tables"
    cultures_dir = definition / "cultures"
    tables_dir.mkdir(parents=True, exist_ok=True)
    cultures_dir.mkdir(parents=True, exist_ok=True)

    model_platform = MODEL_DIR / ".platform"
    model_platform.write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "SemanticModel", "displayName": PROJECT_NAME},
        "config": {"version": "2.0", "logicalId": existing_logical_id(model_platform)},
    }, indent=2), encoding="utf-8")
    (MODEL_DIR / "definition.pbism").write_text(json.dumps({"version": "4.2", "settings": {"qnaEnabled": True}}, indent=2), encoding="utf-8")
    (definition / "database.tmdl").write_text("database\n\tcompatibilityLevel: 1600\n", encoding="utf-8")
    (definition / "model.tmdl").write_text(
        "\n".join([
            "model Model",
            "\tculture: en-CA",
            "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
            "\tsourceQueryCulture: en-CA",
            "\tdataAccessOptions",
            "\t\tlegacyRedirects",
            "\t\treturnErrorValuesAsNull",
            "",
            "annotation PBI_QueryOrder = [\"Province\",\"CSD\",\"Indicator\",\"Province Indicators\",\"CSD Indicators\",\"MIZ Summary\",\"Notes\"]",
            "annotation PBI_ProTooling = [\"DevMode\"]",
            "",
            "ref table Province",
            "ref table CSD",
            "ref table Indicator",
            "ref table 'Province Indicators'",
            "ref table 'CSD Indicators'",
            "ref table 'MIZ Summary'",
            "ref table Notes",
            "",
            "ref cultureInfo en-CA",
            "",
        ]),
        encoding="utf-8",
    )
    data_folder = str(DATA_DIR).replace("\\", "\\\\") + "\\\\"
    (definition / "expressions.tmdl").write_text(
        f'expression DataFolder = "{data_folder}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n',
        encoding="utf-8",
    )
    (cultures_dir / "en-CA.tmdl").write_text("cultureInfo en-CA\n", encoding="utf-8")

    province_measures = [
        tmdl_measure("Province Children 0 to 5", "SUM ( Province[Children Aged 0 to 5] )", "#,##0", "Core"),
        tmdl_measure("Province Children 0 to 5 (Excluding Canada)", 'IF ( SELECTEDVALUE ( Province[Is Canada] ) = "Yes", BLANK (), [Province Children 0 to 5] )', "#,##0", "Core"),
        tmdl_measure("Canada Children 0 to 5", 'CALCULATE ( [Province Children 0 to 5], Province[Province Name] = "Canada" )', "#,##0", "Canada"),
        tmdl_measure("Canada Census Family Population", 'CALCULATE ( SUM ( Province[Population in Census Families] ), Province[Province Name] = "Canada" )', "#,##0", "Canada"),
        tmdl_measure("Province Indicator Children", "SUM ( 'Province Indicators'[Children Count] )", "#,##0", "Indicators"),
        tmdl_measure("Province Indicator Rate", "DIVIDE ( [Province Indicator Children], [Province Children 0 to 5] )", "0.0%", "Indicators"),
        tmdl_measure("Province Indicator Rate (Excluding Canada)", 'IF ( SELECTEDVALUE ( Province[Is Canada] ) = "Yes", BLANK (), [Province Indicator Rate] )', "0.0%", "Indicators"),
        tmdl_measure("Canada Indicator Children", 'CALCULATE ( [Province Indicator Children], Province[Province Name] = "Canada" )', "#,##0", "Canada"),
        tmdl_measure("Canada Indicator Rate", 'CALCULATE ( [Province Indicator Rate], Province[Province Name] = "Canada" )', "0.0%", "Canada"),
        tmdl_measure("Canada Low Income Rate", 'CALCULATE ( [Canada Indicator Rate], Indicator[Indicator] = "Below LIM" )', "0.0%", "Canada"),
        tmdl_measure("Canada Indigenous Identity Rate", 'CALCULATE ( [Canada Indicator Rate], Indicator[Indicator] = "Indigenous Identity" )', "0.0%", "Canada"),
        tmdl_measure("Canada Lone Parent Rate", 'CALCULATE ( [Canada Indicator Rate], Indicator[Indicator] = "Lone Parent" )', "0.0%", "Canada"),
    ]
    province_cols = [
        tmdl_column("Province Code", "int64", hidden=True),
        tmdl_column("Province Name", "string"),
        tmdl_column("Is Canada", "string"),
        tmdl_column("Population in Census Families", "int64", hidden=True, fmt="#,##0", summarize="sum"),
        tmdl_column("Children Aged 0 to 5", "int64", hidden=True, fmt="#,##0", summarize="sum"),
        tmdl_column("Population Status", "string"),
        tmdl_column("Children Status", "string"),
    ]
    (tables_dir / "Province.tmdl").write_text(
        "table Province\n\n" + "\n\n".join(province_measures + province_cols) + "\n\n" +
        m_csv_partition("Province", "province.csv", 7, [
            ("Province Code", "Int64.Type"),
            ("Province Name", "type text"),
            ("Is Canada", "type text"),
            ("Population in Census Families", "Int64.Type"),
            ("Children Aged 0 to 5", "Int64.Type"),
            ("Population Status", "type text"),
            ("Children Status", "type text"),
        ]) + "\n",
        encoding="utf-8",
    )

    csd_measures = [
        tmdl_measure("CSD Children 0 to 5", "SUM ( CSD[Children Aged 0 to 5] )", "#,##0", "Core"),
        tmdl_measure("Number of CSDs", "DISTINCTCOUNT ( CSD[CSD Number] )", "#,##0", "Core"),
        tmdl_measure("CSD Indicator Children", "SUM ( 'CSD Indicators'[Children Count] )", "#,##0", "Indicators"),
        tmdl_measure("CSD Indicator Rate", "DIVIDE ( [CSD Indicator Children], [CSD Children 0 to 5] )", "0.0%", "Indicators"),
        tmdl_measure("CSD Selected Indicator Children", """
VAR SelectedIndicator = SELECTEDVALUE ( Indicator[Indicator], "Below LIM" )
RETURN
    CALCULATE ( [CSD Indicator Children], Indicator[Indicator] = SelectedIndicator )
""", "#,##0", "Indicators"),
        tmdl_measure("CSD Selected Indicator Rate", "DIVIDE ( [CSD Selected Indicator Children], [CSD Children 0 to 5] )", "0.0%", "Indicators"),
        tmdl_measure("CSD Selected Indicator Rate (1,000+ Children)", 'IF ( SELECTEDVALUE ( CSD[Public Denominator Flag] ) = "Under 1,000 children", BLANK (), [CSD Selected Indicator Rate] )', "0.0%", "Indicators"),
        tmdl_measure("CSD Indicator Rate (1,000+ Children)", 'IF ( SELECTEDVALUE ( CSD[Public Denominator Flag] ) = "Under 1,000 children", BLANK (), [CSD Indicator Rate] )', "0.0%", "Indicators"),
        tmdl_measure("Reportable CSD Indicator Rows", 'CALCULATE ( COUNTROWS ( \'CSD Indicators\' ), \'CSD Indicators\'[Value Status] = "Reportable" )', "#,##0", "Quality"),
        tmdl_measure("Suppressed CSD Indicator Rows", 'CALCULATE ( COUNTROWS ( \'CSD Indicators\' ), \'CSD Indicators\'[Value Status] = "Suppressed" )', "#,##0", "Quality"),
    ]
    csd_cols = [
        tmdl_column("CSD Number", "string"),
        tmdl_column("CSD Name", "string"),
        tmdl_column("Province Code", "int64", hidden=True),
        tmdl_column("Province Name", "string"),
        tmdl_column("MIZ Identifier", "int64", fmt="0", summarize="none"),
        tmdl_column("MIZ Label", "string"),
        tmdl_column("MIZ Analysis Label", "string", sort_by="MIZ Identifier"),
        tmdl_column("Geography Type", "string"),
        tmdl_column("Population in Census Families", "int64", hidden=True, fmt="#,##0", summarize="sum"),
        tmdl_column("Children Aged 0 to 5", "int64", hidden=True, fmt="#,##0", summarize="sum"),
        tmdl_column("Child Population Band", "string"),
        tmdl_column("Child Population Band Sort", "int64", hidden=True, fmt="0", summarize="none"),
        tmdl_column("Public Denominator Flag", "string"),
        tmdl_column("Population Status", "string"),
        tmdl_column("Children Status", "string"),
        tmdl_column("MIZ Status", "string"),
        tmdl_column("Province Status", "string"),
    ]
    (tables_dir / "CSD.tmdl").write_text(
        "table CSD\n\n" + "\n\n".join(csd_measures + csd_cols) + "\n\n" +
        m_csv_partition("CSD", "csd.csv", 17, [
            ("CSD Number", "type text"),
            ("CSD Name", "type text"),
            ("Province Code", "Int64.Type"),
            ("Province Name", "type text"),
            ("MIZ Identifier", "Int64.Type"),
            ("MIZ Label", "type text"),
            ("MIZ Analysis Label", "type text"),
            ("Geography Type", "type text"),
            ("Population in Census Families", "Int64.Type"),
            ("Children Aged 0 to 5", "Int64.Type"),
            ("Child Population Band", "type text"),
            ("Child Population Band Sort", "Int64.Type"),
            ("Public Denominator Flag", "type text"),
            ("Population Status", "type text"),
            ("Children Status", "type text"),
            ("MIZ Status", "type text"),
            ("Province Status", "type text"),
        ]) + "\n",
        encoding="utf-8",
    )

    indicator_cols = [
        tmdl_column("Indicator", "string"),
        tmdl_column("Indicator Short Name", "string"),
        tmdl_column("Indicator Description", "string"),
        tmdl_column("Indicator Sort", "int64", hidden=True, fmt="0", summarize="none"),
    ]
    (tables_dir / "Indicator.tmdl").write_text(
        "table Indicator\n\n" + "\n\n".join(indicator_cols) + "\n\n" +
        m_csv_partition("Indicator", "indicator.csv", 4, [
            ("Indicator", "type text"),
            ("Indicator Short Name", "type text"),
            ("Indicator Description", "type text"),
            ("Indicator Sort", "Int64.Type"),
        ]) + "\n",
        encoding="utf-8",
    )

    province_indicator_cols = [
        tmdl_column("Province Code", "int64", hidden=True),
        tmdl_column("Province Name", "string"),
        tmdl_column("Indicator", "string"),
        tmdl_column("Indicator Sort", "int64", hidden=True, fmt="0", summarize="none"),
        tmdl_column("Children Count", "int64", hidden=True, fmt="#,##0", summarize="sum"),
        tmdl_column("Percent", "decimal", hidden=True, fmt="0.0", summarize="none"),
        tmdl_column("Value Status", "string"),
    ]
    (tables_dir / "Province Indicators.tmdl").write_text(
        "table 'Province Indicators'\n\n" + "\n\n".join(province_indicator_cols) + "\n\n" +
        m_csv_partition("Province Indicators", "province_indicators.csv", 7, [
            ("Province Code", "Int64.Type"),
            ("Province Name", "type text"),
            ("Indicator", "type text"),
            ("Indicator Sort", "Int64.Type"),
            ("Children Count", "Int64.Type"),
            ("Percent", "type number"),
            ("Value Status", "type text"),
        ]) + "\n",
        encoding="utf-8",
    )

    csd_indicator_cols = [
        tmdl_column("CSD Number", "string", hidden=True),
        tmdl_column("Province Code", "int64", hidden=True),
        tmdl_column("Indicator", "string"),
        tmdl_column("Indicator Sort", "int64", hidden=True, fmt="0", summarize="none"),
        tmdl_column("Children Count", "int64", hidden=True, fmt="#,##0", summarize="sum"),
        tmdl_column("Percent", "decimal", hidden=True, fmt="0.0", summarize="none"),
        tmdl_column("Value Status", "string"),
    ]
    (tables_dir / "CSD Indicators.tmdl").write_text(
        "table 'CSD Indicators'\n\n" + "\n\n".join(csd_indicator_cols) + "\n\n" +
        m_csv_partition("CSD Indicators", "csd_indicators.csv", 7, [
            ("CSD Number", "type text"),
            ("Province Code", "Int64.Type"),
            ("Indicator", "type text"),
            ("Indicator Sort", "Int64.Type"),
            ("Children Count", "Int64.Type"),
            ("Percent", "type number"),
            ("Value Status", "type text"),
        ]) + "\n",
        encoding="utf-8",
    )

    miz_summary_measures = [
        tmdl_measure("MIZ Children 0 to 5", """
SUMX (
    SUMMARIZE (
        'MIZ Summary',
        'MIZ Summary'[MIZ Analysis Label],
        "Children", MAX ( 'MIZ Summary'[Children Aged 0 to 5] )
    ),
    [Children]
)
""", "#,##0", "MIZ"),
        tmdl_measure("MIZ Indicator Children", "SUM ( 'MIZ Summary'[Children Count] )", "#,##0", "MIZ"),
        tmdl_measure("MIZ Indicator Rate", "DIVIDE ( [MIZ Indicator Children], [MIZ Children 0 to 5] )", "0.0%", "MIZ"),
    ]
    miz_summary_cols = [
        tmdl_column("MIZ Identifier", "int64", hidden=True, fmt="0", summarize="none"),
        tmdl_column("MIZ Analysis Label", "string", sort_by="MIZ Identifier"),
        tmdl_column("MIZ Label", "string"),
        tmdl_column("Geography Type", "string"),
        tmdl_column("Indicator", "string"),
        tmdl_column("Indicator Short Name", "string", sort_by="Indicator Sort"),
        tmdl_column("Indicator Sort", "int64", hidden=True, fmt="0", summarize="none"),
        tmdl_column("Children Aged 0 to 5", "int64", hidden=True, fmt="#,##0", summarize="sum"),
        tmdl_column("Children Count", "int64", hidden=True, fmt="#,##0", summarize="sum"),
        tmdl_column("Percent", "decimal", hidden=True, fmt="0.0%", summarize="none"),
    ]
    (tables_dir / "MIZ Summary.tmdl").write_text(
        "table 'MIZ Summary'\n\n" + "\n\n".join(miz_summary_measures + miz_summary_cols) + "\n\n" +
        m_csv_partition("MIZ Summary", "miz_summary.csv", 10, [
            ("MIZ Identifier", "Int64.Type"),
            ("MIZ Analysis Label", "type text"),
            ("MIZ Label", "type text"),
            ("Geography Type", "type text"),
            ("Indicator", "type text"),
            ("Indicator Short Name", "type text"),
            ("Indicator Sort", "Int64.Type"),
            ("Children Aged 0 to 5", "Int64.Type"),
            ("Children Count", "Int64.Type"),
            ("Percent", "type number"),
        ]) + "\n",
        encoding="utf-8",
    )

    notes_cols = [tmdl_column("Note Number", "int64", fmt="0", summarize="none"), tmdl_column("Note", "string")]
    (tables_dir / "Notes.tmdl").write_text(
        "table Notes\n\n" + "\n\n".join(notes_cols) + "\n\n" +
        m_csv_partition("Notes", "notes.csv", 2, [("Note Number", "Int64.Type"), ("Note", "type text")]) + "\n",
        encoding="utf-8",
    )

    relationships = [
        ("CSD to Province", "CSD.'Province Code'", "Province.'Province Code'"),
        ("Province Indicators to Province", "'Province Indicators'.'Province Code'", "Province.'Province Code'"),
        ("Province Indicators to Indicator", "'Province Indicators'.Indicator", "Indicator.Indicator"),
        ("CSD Indicators to CSD", "'CSD Indicators'.'CSD Number'", "CSD.'CSD Number'"),
        ("CSD Indicators to Indicator", "'CSD Indicators'.Indicator", "Indicator.Indicator"),
        ("MIZ Summary to Indicator", "'MIZ Summary'.Indicator", "Indicator.Indicator"),
    ]
    rel_lines = []
    for name, from_col, to_col in relationships:
        rel_lines.extend([f"relationship {qname(name)}", f"\tfromColumn: {from_col}", f"\ttoColumn: {to_col}", ""])
    (definition / "relationships.tmdl").write_text("\n".join(rel_lines), encoding="utf-8")


def literal(value):
    return {"expr": {"Literal": {"Value": value}}}


def solid(color):
    return {"solid": {"color": {"expr": {"Literal": {"Value": f"'{color}'"}}}}}


def col_field(table, column):
    return {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": column}}


def measure_field(table, measure):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": table}}, "Property": measure}}


def projection(field, query_ref, native_ref, active=None):
    result = {"field": field, "queryRef": query_ref, "nativeQueryRef": native_ref}
    if active is not None:
        result["active"] = active
    return result


def visual_base(name, vtype, x, y, w, h, z):
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.10.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z},
        "visual": {"visualType": vtype},
    }


def add_title(v, title):
    v["visual"].setdefault("visualContainerObjects", {})["title"] = [{
        "properties": {
            "show": literal("true"),
            "text": literal(f"'{title}'"),
        }
    }]
    return v


class Ids:
    def __init__(self):
        self.n = 0

    def next(self):
        self.n += 1
        return f"{self.n:020x}"[-20:]


def textbox(ids, text, x, y, w, h, size="24pt", color="#1F2933"):
    v = visual_base(ids.next(), "textbox", x, y, w, h, 1000 + ids.n)
    v["visual"]["objects"] = {
        "general": [{
            "properties": {
                "paragraphs": [{
                    "textRuns": [{
                        "value": text,
                        "textStyle": {
                            "fontFamily": "Segoe UI Semibold",
                            "fontSize": size,
                            "color": color,
                        },
                    }],
                    "horizontalTextAlignment": "left",
                }]
            }
        }]
    }
    v["visual"]["visualContainerObjects"] = {
        "background": [{"properties": {"show": literal("false")}}],
        "border": [{"properties": {"show": literal("false")}}],
        "padding": [{"properties": {"top": literal("0D"), "bottom": literal("0D"), "left": literal("0D"), "right": literal("0D")}}],
    }
    return v


def card(ids, measures, x, y, w, h, title=None):
    v = visual_base(ids.next(), "cardVisual", x, y, w, h, 1000 + ids.n)
    v["visual"]["query"] = {"queryState": {"Data": {"projections": [
        projection(measure_field(table, measure), f"{table}.{measure}", measure) for table, measure in measures
    ]}}}
    v["visual"]["visualContainerObjects"] = {
        "background": [{"properties": {"show": literal("true"), "color": solid("#FFFFFF"), "transparency": literal("0D")}}],
        "border": [{"properties": {"show": literal("true"), "color": solid("#D9E2EC"), "radius": literal("6D")}}],
        "padding": [{"properties": {"top": literal("6D"), "bottom": literal("6D"), "left": literal("8D"), "right": literal("8D")}}],
    }
    if title:
        add_title(v, title)
    return v


def chart(ids, vtype, category_table, category_col, measure_table, measure, x, y, w, h, title, sort_desc=True, series=None):
    v = visual_base(ids.next(), vtype, x, y, w, h, 1000 + ids.n)
    query_state = {
        "Category": {"projections": [projection(col_field(category_table, category_col), f"{category_table}.{category_col}", category_col, True)]},
        "Y": {"projections": [projection(measure_field(measure_table, measure), f"{measure_table}.{measure}", measure)]},
    }
    if series:
        table, col = series
        query_state["Series"] = {"projections": [projection(col_field(table, col), f"{table}.{col}", col)]}
    v["visual"]["query"] = {"queryState": query_state}
    v["visual"]["visualContainerObjects"] = {
        "background": [{"properties": {"show": literal("true"), "color": solid("#FFFFFF"), "transparency": literal("0D")}}],
        "border": [{"properties": {"show": literal("true"), "color": solid("#D9E2EC"), "radius": literal("6D")}}],
        "padding": [{"properties": {"top": literal("8D"), "bottom": literal("8D"), "left": literal("8D"), "right": literal("8D")}}],
    }
    return add_title(v, title)


def scatter_chart(ids, x, y, w, h, title):
    v = visual_base(ids.next(), "scatterChart", x, y, w, h, 1000 + ids.n)
    v["visual"]["query"] = {
        "queryState": {
            "Category": {"projections": [projection(col_field("CSD", "CSD Name"), "CSD.CSD Name", "CSD Name", True)]},
            "X": {"projections": [projection(measure_field("CSD", "CSD Children 0 to 5"), "CSD.CSD Children 0 to 5", "CSD Children 0 to 5")]},
            "Y": {"projections": [projection(measure_field("CSD", "CSD Selected Indicator Rate"), "CSD.CSD Selected Indicator Rate", "CSD Selected Indicator Rate")]},
            "Size": {"projections": [projection(measure_field("CSD", "CSD Selected Indicator Children"), "CSD.CSD Selected Indicator Children", "CSD Selected Indicator Children")]},
            "Series": {"projections": [projection(col_field("CSD", "MIZ Analysis Label"), "CSD.MIZ Analysis Label", "MIZ Analysis Label")]},
        }
    }
    v["visual"]["visualContainerObjects"] = {
        "background": [{"properties": {"show": literal("true"), "color": solid("#FFFFFF"), "transparency": literal("0D")}}],
        "border": [{"properties": {"show": literal("true"), "color": solid("#D9E2EC"), "radius": literal("6D")}}],
        "padding": [{"properties": {"top": literal("8D"), "bottom": literal("8D"), "left": literal("8D"), "right": literal("8D")}}],
    }
    return add_title(v, title)


def line_chart(ids, x, y, w, h, title):
    v = visual_base(ids.next(), "lineChart", x, y, w, h, 1000 + ids.n)
    v["visual"]["query"] = {
        "queryState": {
            "Category": {"projections": [projection(col_field("MIZ Summary", "MIZ Analysis Label"), "MIZ Summary.MIZ Analysis Label", "MIZ Analysis Label", True)]},
            "Y": {"projections": [projection(measure_field("MIZ Summary", "MIZ Indicator Rate"), "MIZ Summary.MIZ Indicator Rate", "MIZ Indicator Rate")]},
            "Series": {"projections": [projection(col_field("MIZ Summary", "Indicator Short Name"), "MIZ Summary.Indicator Short Name", "Indicator Short Name")]},
            "Tooltips": {"projections": [
                projection(measure_field("MIZ Summary", "MIZ Indicator Children"), "MIZ Summary.MIZ Indicator Children", "MIZ Indicator Children"),
                projection(measure_field("MIZ Summary", "MIZ Children 0 to 5"), "MIZ Summary.MIZ Children 0 to 5", "MIZ Children 0 to 5"),
            ]},
        },
        "sortDefinition": {
            "sort": [{"field": col_field("MIZ Summary", "MIZ Analysis Label"), "direction": "Ascending"}],
            "isDefaultSort": True,
        },
    }
    v["visual"]["visualContainerObjects"] = {
        "background": [{"properties": {"show": literal("true"), "color": solid("#FFFFFF"), "transparency": literal("0D")}}],
        "border": [{"properties": {"show": literal("true"), "color": solid("#D9E2EC"), "radius": literal("6D")}}],
        "padding": [{"properties": {"top": literal("8D"), "bottom": literal("8D"), "left": literal("8D"), "right": literal("8D")}}],
    }
    return add_title(v, title)


def slicer(ids, table, column, label, x, y, w=200):
    v = visual_base(ids.next(), "slicer", x, y, w, 80, 1000 + ids.n)
    v["visual"]["query"] = {"queryState": {"Values": {"projections": [
        projection(col_field(table, column), f"{table}.{column}", column, True)
    ]}}}
    v["visual"]["objects"] = {
        "data": [{"properties": {"mode": literal("'Dropdown'")}}],
        "header": [{"properties": {"show": literal("true"), "text": literal(f"'{label}'")}}],
    }
    v["visual"]["visualContainerObjects"] = {
        "background": [{"properties": {"show": literal("true"), "color": solid("#FFFFFF"), "transparency": literal("0D")}}],
        "border": [{"properties": {"show": literal("true"), "color": solid("#D9E2EC"), "radius": literal("6D")}}],
        "padding": [{"properties": {"top": literal("8D"), "bottom": literal("8D"), "left": literal("8D"), "right": literal("8D")}}],
    }
    return v


def table_visual(ids, values, x, y, w, h, title):
    v = visual_base(ids.next(), "tableEx", x, y, w, h, 1000 + ids.n)
    projections = []
    for kind, table, name in values:
        if kind == "column":
            projections.append(projection(col_field(table, name), f"{table}.{name}", name))
        else:
            projections.append(projection(measure_field(table, name), f"{table}.{name}", name))
    v["visual"]["query"] = {"queryState": {"Values": {"projections": projections}}}
    v["visual"]["visualContainerObjects"] = {
        "stylePreset": [{"properties": {"name": literal("'None'")}}],
        "background": [{"properties": {"show": literal("true"), "color": solid("#FFFFFF"), "transparency": literal("0D")}}],
        "border": [{"properties": {"show": literal("true"), "color": solid("#D9E2EC"), "radius": literal("6D")}}],
        "padding": [{"properties": {"top": literal("8D"), "bottom": literal("8D"), "left": literal("8D"), "right": literal("8D")}}],
    }
    return add_title(v, title)


def make_page(page_name, display_name, visuals):
    page_dir = REPORT_DIR / "definition" / "pages" / page_name
    (page_dir / "visuals").mkdir(parents=True, exist_ok=True)
    page_json = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json",
        "name": page_name,
        "displayName": display_name,
        "displayOption": "FitToPage",
        "height": 720,
        "width": 1280,
    }
    (page_dir / "page.json").write_text(json.dumps(page_json, indent=2), encoding="utf-8")
    for visual in visuals:
        visual_dir = page_dir / "visuals" / visual["name"]
        visual_dir.mkdir(parents=True, exist_ok=True)
        (visual_dir / "visual.json").write_text(json.dumps(visual, indent=2), encoding="utf-8")


def make_report():
    definition = REPORT_DIR / "definition"
    pages_dir = definition / "pages"
    pages_dir.mkdir(parents=True, exist_ok=True)
    (ROOT / f"{PROJECT_NAME}.pbip").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{PROJECT_NAME}.Report"}}],
        "settings": {"enableAutoRecovery": True},
    }, indent=2), encoding="utf-8")
    report_platform = REPORT_DIR / ".platform"
    report_platform.write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "Report", "displayName": PROJECT_NAME},
        "config": {"version": "2.0", "logicalId": existing_logical_id(report_platform)},
    }, indent=2), encoding="utf-8")
    (REPORT_DIR / "definition.pbir").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0",
        "datasetReference": {"byPath": {"path": f"../{PROJECT_NAME}.SemanticModel"}},
    }, indent=2), encoding="utf-8")
    (definition / "version.json").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
        "version": "2.0.0",
    }, indent=2), encoding="utf-8")
    (definition / "report.json").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.3.0/schema.json",
        "themeCollection": {
            "baseTheme": {
                "name": "CY26SU04",
                "reportVersionAtImport": {
                    "visual": "2.8.0",
                    "report": "3.2.0",
                    "page": "2.3.1",
                },
                "type": "SharedResources",
            }
        },
        "objects": {
            "section": [{"properties": {"verticalAlignment": literal("'Top'")}}],
            "outspacePane": [{"properties": {"visible": literal("false")}}],
        },
        "resourcePackages": [{
            "name": "SharedResources",
            "type": "SharedResources",
            "items": [{
                "name": "CY26SU04",
                "path": "BaseThemes/CY26SU04.json",
                "type": "BaseTheme",
            }],
        }],
        "settings": {
            "useStylableVisualContainerHeader": True,
            "exportDataMode": "AllowSummarized",
            "defaultDrillFilterOtherVisuals": True,
            "allowChangeFilterTypes": True,
            "useEnhancedTooltips": True,
            "useDefaultAggregateDisplayName": True,
        },
    }, indent=2), encoding="utf-8")

    overview = "7b1fbceec3b84d13a101"
    explorer = "7b1fbceec3b84d13a102"
    scale = "7b1fbceec3b84d13a103"
    rural = "7b1fbceec3b84d13a104"
    gradient = "7b1fbceec3b84d13a105"
    notes = "7b1fbceec3b84d13a106"
    (pages_dir / "pages.json").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json",
        "pageOrder": [overview, scale, explorer, gradient, rural, notes],
        "activePageName": overview,
    }, indent=2), encoding="utf-8")

    ids = Ids()
    make_page(overview, "Overview", [
        textbox(ids, "Census 2021: Children Aged 0 to 5", 24, 18, 680, 44),
        textbox(ids, "Off-reserve population in census families and selected vulnerability indicators", 24, 56, 840, 28, "13pt", "#52606D"),
        card(ids, [("Province", "Canada Children 0 to 5"), ("Province", "Canada Low Income Rate"), ("Province", "Canada Indigenous Identity Rate"), ("Province", "Canada Lone Parent Rate")], 24, 96, 1232, 112),
        chart(ids, "columnChart", "Province", "Province Name", "Province", "Province Children 0 to 5 (Excluding Canada)", 24, 232, 590, 430, "Children aged 0 to 5 by province"),
        chart(ids, "clusteredBarChart", "Indicator", "Indicator Short Name", "Province", "Canada Indicator Rate", 646, 232, 610, 204, "Canada rate by indicator"),
        chart(ids, "clusteredBarChart", "Province", "Province Name", "Province", "Province Indicator Rate (Excluding Canada)", 646, 458, 610, 204, "Provincial rates by selected indicator", series=("Indicator", "Indicator Short Name")),
    ])

    make_page(scale, "Scale vs Intensity", [
        textbox(ids, "Scale vs Intensity", 24, 18, 480, 44),
        textbox(ids, "Compare child population size with selected indicator intensity; use denominator filters before interpreting small places.", 24, 56, 980, 28, "13pt", "#52606D"),
        slicer(ids, "Province", "Province Name", "Province", 24, 96, 220),
        slicer(ids, "CSD", "MIZ Analysis Label", "MIZ / CMA / CA", 262, 96, 230),
        slicer(ids, "Indicator", "Indicator Short Name", "Indicator", 510, 96, 230),
        slicer(ids, "CSD", "Public Denominator Flag", "Denominator", 758, 96, 230),
        scatter_chart(ids, 24, 196, 620, 454, "CSD scale vs selected indicator rate"),
        chart(ids, "clusteredBarChart", "CSD", "CSD Name", "CSD", "CSD Selected Indicator Rate (1,000+ Children)", 676, 196, 580, 204, "Highest selected rates among CSDs with 1,000+ children"),
        table_visual(ids, [
            ("column", "CSD", "CSD Name"),
            ("column", "CSD", "Province Name"),
            ("column", "CSD", "MIZ Analysis Label"),
            ("measure", "CSD", "CSD Children 0 to 5"),
            ("measure", "CSD", "CSD Selected Indicator Children"),
            ("measure", "CSD", "CSD Selected Indicator Rate"),
        ], 676, 422, 580, 228, "Community profile list"),
    ])

    make_page(explorer, "CSD Explorer", [
        textbox(ids, "Explore CSD-Level Indicators", 24, 18, 600, 44),
        slicer(ids, "Province", "Province Name", "Province", 24, 76, 210),
        slicer(ids, "CSD", "MIZ Analysis Label", "MIZ / CMA / CA", 252, 76, 210),
        slicer(ids, "Indicator", "Indicator Short Name", "Indicator", 480, 76, 210),
        slicer(ids, "CSD", "Public Denominator Flag", "Denominator", 708, 76, 210),
        card(ids, [("CSD", "Number of CSDs"), ("CSD", "CSD Children 0 to 5"), ("CSD", "CSD Selected Indicator Children"), ("CSD", "CSD Selected Indicator Rate")], 936, 76, 320, 112),
        chart(ids, "clusteredBarChart", "CSD", "CSD Name", "CSD", "CSD Selected Indicator Children", 24, 220, 590, 430, "CSD selected indicator count"),
        chart(ids, "clusteredBarChart", "CSD", "CSD Name", "CSD", "CSD Selected Indicator Rate", 646, 220, 610, 190, "CSD selected indicator rate"),
        table_visual(ids, [
            ("column", "CSD", "CSD Name"),
            ("column", "CSD", "Province Name"),
            ("column", "CSD", "MIZ Analysis Label"),
            ("column", "CSD", "Public Denominator Flag"),
            ("measure", "CSD", "CSD Children 0 to 5"),
            ("measure", "CSD", "CSD Selected Indicator Children"),
            ("measure", "CSD", "CSD Selected Indicator Rate"),
            ("column", "CSD Indicators", "Value Status"),
        ], 646, 432, 610, 218, "CSD detail table"),
    ])

    make_page(rural, "Rural / Urban Lens", [
        textbox(ids, "Rural / Urban Lens", 24, 18, 520, 44),
        textbox(ids, "MIZ separates metropolitan, agglomeration, rural influence, and territorial outside-CA geographies.", 24, 56, 980, 28, "13pt", "#52606D"),
        slicer(ids, "Indicator", "Indicator Short Name", "Indicator", 24, 96, 230),
        slicer(ids, "MIZ Summary", "Geography Type", "Geography type", 272, 96, 230),
        card(ids, [("CSD", "CSD Children 0 to 5"), ("CSD", "Number of CSDs"), ("CSD", "Reportable CSD Indicator Rows"), ("CSD", "Suppressed CSD Indicator Rows")], 728, 96, 528, 112),
        chart(ids, "columnChart", "MIZ Summary", "MIZ Analysis Label", "MIZ Summary", "MIZ Children 0 to 5", 24, 236, 590, 414, "Children aged 0 to 5 by MIZ / CMA / CA category"),
        chart(ids, "clusteredBarChart", "MIZ Summary", "MIZ Analysis Label", "MIZ Summary", "MIZ Indicator Rate", 646, 236, 610, 190, "Selected rate by MIZ / CMA / CA category", series=("MIZ Summary", "Indicator Short Name")),
        table_visual(ids, [
            ("column", "MIZ Summary", "MIZ Analysis Label"),
            ("column", "MIZ Summary", "Geography Type"),
            ("column", "MIZ Summary", "Indicator Short Name"),
            ("measure", "MIZ Summary", "MIZ Children 0 to 5"),
            ("measure", "MIZ Summary", "MIZ Indicator Children"),
            ("measure", "MIZ Summary", "MIZ Indicator Rate"),
        ], 646, 448, 610, 202, "MIZ summary table"),
    ])

    make_page(gradient, "Rurality Gradient", [
        textbox(ids, "The Rurality Story", 24, 18, 520, 44),
        textbox(ids, "Selected rates by MIZ / CMA / CA category from the CSD-level table. The x-axis moves from metropolitan to more remote geographies.", 24, 58, 1060, 58, "13pt", "#52606D"),
        line_chart(ids, 24, 132, 1232, 360, "Children aged 0 to 5 by geography type"),
        textbox(ids, "Percentages are calculated from non-suppressed CSD rows; small geographies may be affected by suppression and rounding.", 24, 512, 1080, 44, "12pt", "#6B7280"),
        textbox(ids, "Rural is not one thing: strong MIZ communities near metro areas look different from weak/no MIZ or territorial outside-CA communities.", 24, 572, 1120, 72, "14pt", "#1F2933"),
    ])

    make_page(notes, "Notes and Caveats", [
        textbox(ids, "Notes and Caveats", 24, 18, 600, 44),
        textbox(ids, "Suppressed cells from the source workbook are kept blank in numeric fields and labelled in status columns.", 24, 58, 920, 36, "13pt", "#52606D"),
        table_visual(ids, [
            ("column", "Notes", "Note Number"),
            ("column", "Notes", "Note"),
        ], 24, 112, 1232, 540, "Source notes from workbook"),
    ])


def clean_output_dirs():
    def on_remove_error(func, path, exc_info):
        try:
            os.chmod(path, stat.S_IWRITE)
            func(path)
        except Exception:
            raise

    for path in [REPORT_DIR, MODEL_DIR, DATA_DIR, CONTEXT_DIR]:
        if path.exists():
            shutil.rmtree(path, onerror=on_remove_error)
    ROOT.mkdir(parents=True, exist_ok=True)


def main():
    clean_output_dirs()
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    CONTEXT_DIR.mkdir(parents=True, exist_ok=True)

    parsed = parse_workbook()
    write_csv(DATA_DIR / "province.csv", parsed["province"], ["Province Code", "Province Name", "Is Canada", "Population in Census Families", "Children Aged 0 to 5", "Population Status", "Children Status"])
    write_csv(DATA_DIR / "province_indicators.csv", parsed["province_indicators"], ["Province Code", "Province Name", "Indicator", "Indicator Sort", "Children Count", "Percent", "Value Status"])
    write_csv(DATA_DIR / "csd.csv", parsed["csd"], ["CSD Number", "CSD Name", "Province Code", "Province Name", "MIZ Identifier", "MIZ Label", "MIZ Analysis Label", "Geography Type", "Population in Census Families", "Children Aged 0 to 5", "Child Population Band", "Child Population Band Sort", "Public Denominator Flag", "Population Status", "Children Status", "MIZ Status", "Province Status"])
    write_csv(DATA_DIR / "csd_indicators.csv", parsed["csd_indicators"], ["CSD Number", "Province Code", "Indicator", "Indicator Sort", "Children Count", "Percent", "Value Status"])
    write_csv(DATA_DIR / "miz_summary.csv", parsed["miz_summary"], ["MIZ Identifier", "MIZ Analysis Label", "MIZ Label", "Geography Type", "Indicator", "Indicator Short Name", "Indicator Sort", "Children Aged 0 to 5", "Children Count", "Percent"])
    write_csv(DATA_DIR / "indicator.csv", parsed["indicator"], ["Indicator", "Indicator Short Name", "Indicator Description", "Indicator Sort"])
    write_csv(DATA_DIR / "notes.csv", parsed["notes"], ["Note Number", "Note"])

    email_rows = extract_email_context()
    write_csv(CONTEXT_DIR / "email_context.csv", email_rows, ["file", "subject", "sender", "to", "body"])
    with (CONTEXT_DIR / "email_context.txt").open("w", encoding="utf-8") as handle:
        for item in email_rows:
            handle.write(f"File: {item['file']}\nSubject: {item['subject']}\nSender: {item['sender']}\nTo: {item['to']}\n\n{item['body']}\n\n{'=' * 80}\n\n")

    make_semantic_model()
    make_report()

    summary = {
        "province_rows": len(parsed["province"]),
        "province_indicator_rows": len(parsed["province_indicators"]),
        "csd_rows": len(parsed["csd"]),
        "csd_indicator_rows": len(parsed["csd_indicators"]),
        "miz_summary_rows": len(parsed["miz_summary"]),
        "note_rows": len(parsed["notes"]),
        "emails": [{"file": item["file"], "subject": item["subject"]} for item in email_rows],
    }
    (CONTEXT_DIR / "build_summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
