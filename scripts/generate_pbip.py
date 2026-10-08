"""Generates the Power BI Project (PBIP) for the SCMS supply chain report.

This created the first version of the project. Power BI Desktop has since saved it
(newer file formats, reformatted TMDL), so the files in the repo are the source of truth;
running this again overwrites them with the generated versions.

Writes a TMDL semantic model (Power Query, relationships, DAX) and a PBIR report
(4 pages) next to this folder. Open SCMS_Supply_Chain.pbip in Power BI Desktop,
then Home > Refresh to load the data.

    python scripts/generate_pbip.py
"""
import json
import re
import shutil
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
NAME = "SCMS_Supply_Chain"
SM = ROOT / f"{NAME}.SemanticModel"
RP = ROOT / f"{NAME}.Report"
DATA_FOLDER = str(ROOT / "data") + "\\"
THEME_SRC = Path(__file__).resolve().parent / "assets" / "CY20SU09.json"


def tag(*parts):
    """Stable lineage tags so regenerating doesn't churn every file."""
    return str(uuid.uuid5(uuid.NAMESPACE_URL, "scms-pbip/" + "/".join(parts)))


def q(name):
    """TMDL object name, quoted only when it needs to be."""
    return name if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name) else "'" + name.replace("'", "''") + "'"


def write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def write_json(path, obj):
    write(path, json.dumps(obj, indent=2, ensure_ascii=False) + "\n")


# --------------------------------------------------------------------------------------
# Semantic model
# --------------------------------------------------------------------------------------
CURRENCY = r"\$#,0;(\$#,0);\$#,0"
CURRENCY2 = r"\$#,0.00;(\$#,0.00);\$#,0.00"
PERCENT = "0.0%;-0.0%;0.0%"
COUNT = "#,0"
DEC1 = "#,0.0"

LINE_ITEM_COLUMNS = [  # (name, dataType, hidden)
    ("ID", "int64", False), ("Project Code", "string", False), ("PQ #", "string", False),
    ("PO / SO #", "string", False), ("ASN/DN #", "string", False), ("Country", "string", True),
    ("Managed By", "string", False), ("Fulfill Via", "string", True), ("Vendor INCO Term", "string", False),
    ("Shipment Mode", "string", True), ("PQ First Sent to Client Date", "dateTime", False),
    ("PO Sent to Vendor Date", "dateTime", False), ("Scheduled Delivery Date", "dateTime", False),
    ("Delivered to Client Date", "dateTime", True), ("Delivery Recorded Date", "dateTime", False),
    ("Product Group", "string", False), ("Sub Classification", "string", False), ("Vendor", "string", True),
    ("Item Description", "string", True), ("Molecule/Test Type", "string", False), ("Brand", "string", False),
    ("Dosage", "string", False), ("Dosage Form", "string", False), ("Unit of Measure (Per Pack)", "int64", False),
    ("Line Item Quantity", "int64", False), ("Line Item Value", "double", False), ("Pack Price", "double", False),
    ("Unit Price", "double", False), ("Manufacturing Site", "string", False), ("First Line Designation", "string", False),
    ("Weight (Kilograms)", "string", True), ("Freight Cost (USD)", "string", True),
    ("Line Item Insurance (USD)", "double", False), ("PO Date Status", "string", False),
]

SHIPMENT_COLUMNS = [
    ("ASN/DN #", "string", False), ("Country", "string", True), ("Vendor", "string", True),
    ("Fulfill Via", "string", True), ("Vendor INCO Term", "string", False), ("Shipment Mode", "string", True),
    ("Managed By", "string", False), ("PO Sent to Vendor Date", "dateTime", False),
    ("Scheduled Delivery Date", "dateTime", False), ("Delivered to Client Date", "dateTime", True),
    ("Weight kg", "double", False), ("Freight Cost USD", "double", False), ("Freight Status", "string", False),
    ("Weight Status", "string", False), ("Days vs Schedule", "int64", False), ("PO to Delivery Days", "int64", False),
    ("Delivery Status", "string", False), ("Has Freight and Weight", "boolean", False),
    ("Shipment Value USD", "double", False),
]

M_LINE_ITEMS = r'''let
    Source = Csv.Document(File.Contents(DataFolder & "SCMS_Delivery_History_Dataset.csv"), [Delimiter = ",", Columns = 33, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    #"Added PO Date Status" = Table.AddColumn(#"Promoted Headers", "PO Date Status", each if [PO Sent to Vendor Date] = "N/A - From RDC" then "Not applicable (RDC)" else if [PO Sent to Vendor Date] = "Date Not Captured" then "Not captured" else "Captured", type text),
    #"Changed Type" = Table.TransformColumnTypes(#"Added PO Date Status", {{"ID", Int64.Type}, {"Unit of Measure (Per Pack)", Int64.Type}, {"Line Item Quantity", Int64.Type}, {"Line Item Value", type number}, {"Pack Price", type number}, {"Unit Price", type number}, {"Line Item Insurance (USD)", type number}}, "en-US"),
    #"Parsed Delivery Dates" = Table.TransformColumns(#"Changed Type", {{"Scheduled Delivery Date", each Date.FromText(_, [Format = "d-MMM-yy", Culture = "en-US"]), type date}, {"Delivered to Client Date", each Date.FromText(_, [Format = "d-MMM-yy", Culture = "en-US"]), type date}, {"Delivery Recorded Date", each Date.FromText(_, [Format = "d-MMM-yy", Culture = "en-US"]), type date}}),
    #"Parsed Order Dates" = Table.TransformColumns(#"Parsed Delivery Dates", {{"PQ First Sent to Client Date", each try Date.FromText(_, [Format = "M/d/yy", Culture = "en-US"]) otherwise null, type nullable date}, {"PO Sent to Vendor Date", each try Date.FromText(_, [Format = "M/d/yy", Culture = "en-US"]) otherwise null, type nullable date}}),
    #"Replaced Errors" = Table.ReplaceErrorValues(#"Parsed Order Dates", {{"Line Item Insurance (USD)", null}}),
    #"Replaced Mode" = Table.ReplaceValue(#"Replaced Errors", "N/A", "Not recorded", Replacer.ReplaceValue, {"Shipment Mode"})
in
    #"Replaced Mode"'''

M_SHIPMENTS = r'''let
    Source = Fact_LineItems,
    #"First Lines" = Table.SelectRows(Source, each [First Line Designation] = "Yes"),
    #"Kept Columns" = Table.SelectColumns(#"First Lines", {"ASN/DN #", "Country", "Vendor", "Fulfill Via", "Vendor INCO Term", "Shipment Mode", "Managed By", "PO Sent to Vendor Date", "Scheduled Delivery Date", "Delivered to Client Date", "Weight (Kilograms)", "Freight Cost (USD)"}),
    #"Added Freight Status" = Table.AddColumn(#"Kept Columns", "Freight Status", each if Text.StartsWith([#"Freight Cost (USD)"], "Freight Included") then "Included in commodity cost" else if [#"Freight Cost (USD)"] = "Invoiced Separately" then "Invoiced separately" else "Recorded", type text),
    #"Added Weight Status" = Table.AddColumn(#"Added Freight Status", "Weight Status", each if [#"Weight (Kilograms)"] = "Weight Captured Separately" then "Captured separately" else "Recorded", type text),
    #"To Numbers" = Table.TransformColumns(#"Added Weight Status", {{"Freight Cost (USD)", each try Number.FromText(_, "en-US") otherwise null, type nullable number}, {"Weight (Kilograms)", each try Number.FromText(_, "en-US") otherwise null, type nullable number}}),
    #"Renamed Columns" = Table.RenameColumns(#"To Numbers", {{"Freight Cost (USD)", "Freight Cost USD"}, {"Weight (Kilograms)", "Weight kg"}}),
    #"Added Days vs Schedule" = Table.AddColumn(#"Renamed Columns", "Days vs Schedule", each Duration.Days([Delivered to Client Date] - [Scheduled Delivery Date]), Int64.Type),
    #"Added PO to Delivery Days" = Table.AddColumn(#"Added Days vs Schedule", "PO to Delivery Days", each if [PO Sent to Vendor Date] = null then null else Duration.Days([Delivered to Client Date] - [PO Sent to Vendor Date]), Int64.Type),
    #"Added Delivery Status" = Table.AddColumn(#"Added PO to Delivery Days", "Delivery Status", each if [Days vs Schedule] > 0 then "Late" else if [Days vs Schedule] = 0 then "On schedule" else "Early", type text),
    #"Added Has Freight and Weight" = Table.AddColumn(#"Added Delivery Status", "Has Freight and Weight", each [Freight Cost USD] <> null and [Weight kg] <> null and [Weight kg] > 0, type logical),
    #"Shipment Values" = Table.Group(Source, {"ASN/DN #"}, {{"Shipment Value USD", each List.Sum([Line Item Value]), type number}}),
    #"Merged Shipment Value" = Table.NestedJoin(#"Added Has Freight and Weight", {"ASN/DN #"}, #"Shipment Values", {"ASN/DN #"}, "SV", JoinKind.LeftOuter),
    #"Expanded Shipment Value" = Table.ExpandTableColumn(#"Merged Shipment Value", "SV", {"Shipment Value USD"}, {"Shipment Value USD"}),
    #"Typed Shipment Value" = Table.TransformColumnTypes(#"Expanded Shipment Value", {{"Shipment Value USD", type number}})
in
    #"Typed Shipment Value"'''


def m_dimension(columns):
    cols = ", ".join(f'"{c}"' for c in columns)
    return f'''let
    Source = Fact_LineItems,
    #"Kept Columns" = Table.SelectColumns(Source, {{{cols}}}),
    #"Removed Duplicates" = Table.Distinct(#"Kept Columns")
in
    #"Removed Duplicates"'''


DIMENSIONS = {
    "Dim_Country": ["Country"],
    "Dim_Vendor": ["Vendor"],
    "Dim_ShipmentMode": ["Shipment Mode"],
    "Dim_FulfillVia": ["Fulfill Via"],
    "Dim_Product": ["Item Description", "Product Group", "Sub Classification", "Molecule/Test Type", "Brand", "Dosage", "Dosage Form"],
}

LINE_ITEM_MEASURES = [  # (name, DAX, format)
    ("Line Items", "COUNTROWS ( Fact_LineItems )", COUNT),
    ("Commodity Value (USD)", "SUM ( Fact_LineItems[Line Item Value] )", CURRENCY),
    ("Packs Delivered", "SUM ( Fact_LineItems[Line Item Quantity] )", COUNT),
    ("Commodity Value Top 10 Countries",
     "VAR CountryRank =\n    RANKX ( ALLSELECTED ( Dim_Country[Country] ), [Commodity Value (USD)],, DESC )\nRETURN\n    IF ( CountryRank <= 10, [Commodity Value (USD)] )",
     CURRENCY),
]

SHIPMENT_MEASURES = [
    ("Shipments", "COUNTROWS ( Fact_Shipments )", COUNT),
    ("Late Shipments", "CALCULATE ( [Shipments], Fact_Shipments[Days vs Schedule] > 0 )", COUNT),
    ("% Late vs Schedule", "DIVIDE ( [Late Shipments] + 0, [Shipments] )", PERCENT),
    ("Avg Days Late",
     "CALCULATE (\n    AVERAGE ( Fact_Shipments[Days vs Schedule] ),\n    Fact_Shipments[Days vs Schedule] > 0\n)", DEC1),
    ("Avg PO to Delivery Days", "AVERAGE ( Fact_Shipments[PO to Delivery Days] )", DEC1),
    ("Median PO to Delivery Days", "MEDIAN ( Fact_Shipments[PO to Delivery Days] )", COUNT),
    ("Freight Cost (USD)", "SUM ( Fact_Shipments[Freight Cost USD] )", CURRENCY),
    ("Weight (kg)", "SUM ( Fact_Shipments[Weight kg] )", COUNT),
    ("Freight per kg (USD)",
     "DIVIDE (\n    CALCULATE ( SUM ( Fact_Shipments[Freight Cost USD] ), Fact_Shipments[Has Freight and Weight] = TRUE () ),\n    CALCULATE ( SUM ( Fact_Shipments[Weight kg] ), Fact_Shipments[Has Freight and Weight] = TRUE () )\n)",
     CURRENCY2),
    ("Freight % of Value",
     "DIVIDE (\n    [Freight Cost (USD)],\n    CALCULATE (\n        SUM ( Fact_Shipments[Shipment Value USD] ),\n        NOT ISBLANK ( Fact_Shipments[Freight Cost USD] )\n    )\n)",
     PERCENT),
    ("% Freight Not Recorded",
     "DIVIDE (\n    CALCULATE ( [Shipments], Fact_Shipments[Freight Status] <> \"Recorded\" ),\n    [Shipments]\n)", PERCENT),
    ("% Weight Not Recorded",
     "DIVIDE (\n    CALCULATE ( [Shipments], Fact_Shipments[Weight Status] <> \"Recorded\" ),\n    [Shipments]\n)", PERCENT),
    ("% Late (Vendors 50+ Shipments)", "IF ( [Shipments] >= 50, [% Late vs Schedule] )", PERCENT),
    ("% Late (Countries 100+ Shipments)", "IF ( CALCULATE ( [Shipments], REMOVEFILTERS ( Dim_Date ) ) >= 100, [% Late vs Schedule] )", PERCENT),
    ("Freight % of Value (Top 10 Freight Countries)",
     "VAR CountryRank =\n    RANKX ( ALLSELECTED ( Dim_Country[Country] ), [Freight Cost (USD)],, DESC )\nRETURN\n    IF ( CountryRank <= 10, [Freight % of Value] )",
     PERCENT),
]

DATE_TABLE_DAX = """ADDCOLUMNS (
    CALENDAR ( DATE ( 2006, 1, 1 ), DATE ( 2015, 12, 31 ) ),
    "Year", YEAR ( [Date] ),
    "Month Number", MONTH ( [Date] ),
    "Month", FORMAT ( [Date], "MMM" ),
    "Year-Month", FORMAT ( [Date], "YYYY-MM" )
)"""


def indent(text, tabs):
    return "\n".join(("\t" * tabs + line) if line else "" for line in text.splitlines())


def tmdl_measure(table, name, dax, fmt):
    if "\n" in dax:
        head = f"\tmeasure {q(name)} = ```\n{indent(dax, 3)}\n\t\t\t```"
    else:
        head = f"\tmeasure {q(name)} = {dax}"
    return f"{head}\n\t\tformatString: {fmt}\n\t\tlineageTag: {tag(table, 'measure', name)}\n"


def tmdl_column(table, name, dtype, hidden=False, source=None, extra=()):
    lines = [f"\tcolumn {q(name)}", f"\t\tdataType: {dtype}"]
    if hidden:
        lines.append("\t\tisHidden")
    if dtype == "dateTime":
        lines.append("\t\tformatString: Short Date")
    elif dtype == "int64":
        lines.append("\t\tformatString: 0")
    elif dtype == "double":
        lines.append("\t\tformatString: #,0.00")
    lines.append(f"\t\tlineageTag: {tag(table, 'column', name)}")
    numeric_measure = dtype == "double" and name not in ("Pack Price", "Unit Price")
    lines.append(f"\t\tsummarizeBy: {'sum' if numeric_measure else 'none'}")
    lines.extend(f"\t\t{e}" for e in extra)
    lines.append(f"\t\tsourceColumn: {source or name}")
    lines.append("")
    lines.append("\t\tannotation SummarizationSetBy = Automatic")
    if dtype == "dateTime":
        lines.append("")
        lines.append("\t\tannotation UnderlyingDateTimeDataType = Date")
    return "\n".join(lines) + "\n"


def tmdl_table(name, columns, m_query, measures=()):
    parts = [f"table {q(name)}\n\tlineageTag: {tag(name)}\n"]
    parts += [tmdl_measure(name, *m) for m in measures]
    parts += [tmdl_column(name, c, t, h) for c, t, h in columns]
    parts.append(
        f"\tpartition {q(name + '-' + tag(name, 'partition'))} = m\n\t\tmode: import\n\t\tsource =\n"
        f"{indent(m_query, 4)}\n"
    )
    parts.append("\tannotation PBI_ResultType = Table\n")
    return "\n".join(parts)


def tmdl_date_table():
    name = "Dim_Date"
    parts = [f"table {name}\n\tlineageTag: {tag(name)}\n\tdataCategory: Time\n"]
    parts.append(tmdl_column(name, "Date", "dateTime", source="[Date]", extra=("isKey", "isNameInferred")))
    parts.append(tmdl_column(name, "Year", "int64", source="[Year]", extra=("isNameInferred",)))
    parts.append(tmdl_column(name, "Month Number", "int64", hidden=True, source="[Month Number]", extra=("isNameInferred",)))
    parts.append(tmdl_column(name, "Month", "string", source="[Month]", extra=("isNameInferred", "sortByColumn: 'Month Number'")))
    parts.append(tmdl_column(name, "Year-Month", "string", source="[Year-Month]", extra=("isNameInferred",)))
    parts.append(
        f"\tpartition {q(name + '-' + tag(name, 'partition'))} = calculated\n\t\tmode: import\n\t\tsource = ```\n"
        f"{indent(DATE_TABLE_DAX, 4)}\n\t\t\t\t```\n"
    )
    return "\n".join(parts)


def col_ref(table, column):
    return f"{q(table)}.{q(column)}"


def build_semantic_model():
    if SM.exists():
        shutil.rmtree(SM)
    d = SM / "definition"
    write_json(SM / "definition.pbism", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.2",
        "settings": {},
    })
    write_json(SM / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "SemanticModel", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": tag("semanticmodel")},
    })
    write(d / "database.tmdl", "database\n\tcompatibilityLevel: 1567\n\n")

    tables = ["Fact_LineItems", "Fact_Shipments", *DIMENSIONS, "Dim_Date"]
    query_order = json.dumps(["DataFolder", "Fact_LineItems", "Fact_Shipments", *DIMENSIONS])
    write(d / "model.tmdl", (
        "model Model\n\tculture: en-US\n\tdefaultPowerBIDataSourceVersion: powerBI_V3\n\tsourceQueryCulture: en-US\n"
        "\tdataAccessOptions\n\t\tlegacyRedirects\n\t\treturnErrorValuesAsNull\n\n"
        f"annotation PBI_QueryOrder = {query_order}\n\n"
        "annotation __PBI_TimeIntelligenceEnabled = 0\n\n"
        "annotation PBI_ProTooling = [\"DevMode\"]\n\n"
        + "".join(f"ref table {q(t)}\n" for t in tables)
        + "\nref cultureInfo en-US\n\n"
    ))
    write(d / "cultures" / "en-US.tmdl", (
        'cultureInfo en-US\n\n\tlinguisticMetadata =\n\t\t\t{\n\t\t\t  "Version": "1.0.0",\n\t\t\t  "Language": "en-US"\n\t\t\t}\n'
        "\t\tcontentType: json\n\n"
    ))
    write(d / "expressions.tmdl", (
        f'/// Folder holding SCMS_Delivery_History_Dataset.csv. Change it if you move or clone the project.\n'
        f'expression DataFolder = "{DATA_FOLDER}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n'
        f"\tlineageTag: {tag('DataFolder')}\n\n\tannotation PBI_NavigationStepName = Navigation\n\n\tannotation PBI_ResultType = Text\n\n"
    ))

    write(d / "tables" / "Fact_LineItems.tmdl", tmdl_table("Fact_LineItems", LINE_ITEM_COLUMNS, M_LINE_ITEMS, LINE_ITEM_MEASURES))
    write(d / "tables" / "Fact_Shipments.tmdl", tmdl_table("Fact_Shipments", SHIPMENT_COLUMNS, M_SHIPMENTS, SHIPMENT_MEASURES))
    for dim, cols in DIMENSIONS.items():
        write(d / "tables" / f"{dim}.tmdl", tmdl_table(dim, [(c, "string", False) for c in cols], m_dimension(cols)))
    write(d / "tables" / "Dim_Date.tmdl", tmdl_date_table())

    rels = []
    for fact in ("Fact_LineItems", "Fact_Shipments"):
        for dim, key in (("Dim_Country", "Country"), ("Dim_Vendor", "Vendor"), ("Dim_ShipmentMode", "Shipment Mode"),
                         ("Dim_FulfillVia", "Fulfill Via"), ("Dim_Date", "Delivered to Client Date")):
            to_col = "Date" if dim == "Dim_Date" else key
            rels.append((fact, key, dim, to_col))
    rels.append(("Fact_LineItems", "Item Description", "Dim_Product", "Item Description"))
    write(d / "relationships.tmdl", "".join(
        f"relationship {tag('rel', f, fc, t)}\n\tfromColumn: {col_ref(f, fc)}\n\ttoColumn: {col_ref(t, tc)}\n\n"
        for f, fc, t, tc in rels
    ))


# --------------------------------------------------------------------------------------
# Report (PBIR)
# --------------------------------------------------------------------------------------
VISUAL_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.2.0/schema.json"
LATE_COLOUR = "#D9534F"
MAIN_COLOUR = "#2F5D8A"


def lit(value):
    return {"expr": {"Literal": {"Value": value}}}


def text_lit(s):
    return lit("'" + s.replace("'", "''") + "'")


def colour(hex_code):
    return {"solid": {"color": text_lit(hex_code)}}


def column(entity, prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def measure(entity, prop):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def proj(field, active=False, display=None):
    kind = "Column" if "Column" in field else "Measure"
    p = {"field": field, "queryRef": f"{field[kind]['Expression']['SourceRef']['Entity']}.{field[kind]['Property']}"}
    if active:
        p["active"] = True
    if display:
        p["displayName"] = display
    return p


def container(title):
    return {
        "title": [{"properties": {"show": lit("true"), "text": text_lit(title)}}],
        "visualHeader": [{"properties": {"show": lit("false")}}],
    }


def visual(name, vtype, pos, query_state=None, title=None, objects=None, sort=None):
    x, y, w, h = pos
    v = {"visualType": vtype}
    if query_state is not None:
        v["query"] = {"queryState": {role: {"projections": ps} for role, ps in query_state.items()}}
        if sort:
            field, direction = sort
            v["query"]["sortDefinition"] = {"sort": [{"field": field, "direction": direction}], "isDefaultSort": False}
    if objects:
        v["objects"] = objects
    if title:
        v["visualContainerObjects"] = container(title)
    v["drillFilterOtherVisuals"] = True
    return name, {"$schema": VISUAL_SCHEMA, "name": name,
                  "position": {"x": x, "y": y, "z": 0, "height": h, "width": w, "tabOrder": 0}, "visual": v}


def card(name, pos, entity, m, title, label_props=None):
    return visual(name, "card", pos, {"Values": [proj(measure(entity, m))]}, title,
                  objects={"labels": [{"properties": {"fontSize": lit("20D"), **(label_props or {})}}],
                           "categoryLabels": [{"properties": {"show": lit("false")}}]})


def bar(name, pos, cat, val, title, vtype="clusteredBarChart", bar_colour=MAIN_COLOUR, sort_desc=True):
    objects = {"labels": [{"properties": {"show": lit("true")}}],
               "dataPoint": [{"properties": {"fill": colour(bar_colour)}}]}
    sort = (val, "Descending") if sort_desc else (cat, "Ascending")
    return visual(name, vtype, pos, {"Category": [proj(cat, active=True)], "Y": [proj(val)]}, title, objects, sort)


def table_visual(name, pos, fields, title, labels=None):
    labels = labels or [None] * len(fields)
    return visual(name, "tableEx", pos, {"Values": [proj(f, display=l) for f, l in zip(fields, labels)]}, title)


def slicer(name, pos, field, title):
    return visual(name, "slicer", pos, {"Values": [proj(field, active=True)]}, title,
                  objects={"data": [{"properties": {"mode": text_lit("Dropdown")}}]})


def textbox(name, pos, paragraphs):
    """paragraphs: list of lists of (text, size_pt, bold)."""
    paras = []
    for runs in paragraphs:
        text_runs = []
        for text, size, bold in runs:
            style = {"fontSize": f"{size}pt"}
            if bold:
                style["fontWeight"] = "bold"
            text_runs.append({"value": text, "textStyle": style})
        paras.append({"textRuns": text_runs})
    return visual(name, "textbox", pos, objects={"general": [{"properties": {"paragraphs": paras}}]})


LI, SH = "Fact_LineItems", "Fact_Shipments"
YEAR = column("Dim_Date", "Year")
COUNTRY = column("Dim_Country", "Country")
VENDOR = column("Dim_Vendor", "Vendor")
MODE = column("Dim_ShipmentMode", "Shipment Mode")
FULFIL = column("Dim_FulfillVia", "Fulfill Via")


def title_box(name, heading, sub):
    return textbox(name, (20, 4, 1240, 64), [[(heading, 16, True)], [(sub, 9, False)]])


PAGES = [
    ("overview", "Overview", [
        title_box("ov_title", "Global Health Supply Chain Performance",
                  "USAID / PEPFAR SCMS shipments, 2006-2015. Late = delivered after the scheduled date; per the data dictionary the scheduled date is not the client-promised date."),
        card("ov_card_value", (20, 72, 236, 100), LI, "Commodity Value (USD)", "Commodity value", {"labelPrecision": lit("2L")}),
        card("ov_card_shipments", (268, 72, 236, 100), SH, "Shipments", "Shipments"),
        card("ov_card_late", (516, 72, 236, 100), SH, "% Late vs Schedule", "% late vs schedule"),
        card("ov_card_lead", (764, 72, 236, 100), SH, "Avg PO to Delivery Days", "Avg PO-to-delivery days (Direct Drop)"),
        card("ov_card_freight", (1012, 72, 248, 100), SH, "Freight Cost (USD)", "Freight cost (recorded)", {"labelPrecision": lit("2L")}),
        visual("ov_combo_year", "lineClusteredColumnComboChart", (20, 184, 760, 330),
               {"Category": [proj(YEAR, active=True)], "Y": [proj(measure(SH, "Shipments"))],
                "Y2": [proj(measure(SH, "% Late vs Schedule"))]},
               "Shipments and % late vs schedule by delivery year",
               objects={"dataPoint": [{"properties": {"fill": colour(MAIN_COLOUR)}}],
                        "labels": [{"properties": {"show": lit("false")}}]},
               sort=(YEAR, "Ascending")),
        bar("ov_bar_country", (800, 184, 460, 526), COUNTRY, measure(LI, "Commodity Value Top 10 Countries"),
            "Top 10 countries by commodity value"),
        slicer("ov_slicer_year", (20, 530, 245, 80), YEAR, "Delivery year"),
        slicer("ov_slicer_mode", (277, 530, 245, 80), MODE, "Shipment mode"),
        slicer("ov_slicer_fulfil", (534, 530, 246, 80), FULFIL, "Fulfilment route"),
        textbox("ov_note", (20, 622, 760, 88), [
            [("How to read this page", 11, True)],
            [("Volume and lateness are counted per shipment (7,030), value per line item (10,324). Freight and weight are recorded once per shipment, so they come from the shipment table.", 10, False)],
        ]),
    ]),
    ("delivery", "Delivery performance", [
        title_box("dl_title", "Delivery performance",
                  "Which modes, vendors, routes and countries deliver later than scheduled."),
        bar("dl_bar_mode", (20, 72, 400, 300), MODE, measure(SH, "% Late vs Schedule"),
            "% late vs schedule by shipment mode", bar_colour=LATE_COLOUR),
        bar("dl_bar_vendor", (440, 72, 820, 300), VENDOR, measure(SH, "% Late (Vendors 50+ Shipments)"),
            "% late vs schedule by vendor (vendors with 50+ shipments)", bar_colour=LATE_COLOUR),
        table_visual("dl_table_route", (20, 388, 400, 150),
                     [FULFIL, measure(SH, "Shipments"), measure(SH, "% Late vs Schedule"), measure(SH, "Avg PO to Delivery Days")],
                     "Direct Drop vs From RDC", labels=["Route", "Shipments", "% late", "Avg lead time (days)"]),
        bar("dl_col_status", (20, 552, 400, 158), column(SH, "Delivery Status"), measure(SH, "Shipments"),
            "Shipments by delivery status", vtype="clusteredColumnChart"),
        visual("dl_matrix_country_year", "pivotTable", (440, 388, 820, 322),
               {"Rows": [proj(COUNTRY, active=True)], "Columns": [proj(YEAR, active=True)],
                "Values": [proj(measure(SH, "% Late (Countries 100+ Shipments)"))]},
               "% late vs schedule by country and year (countries with 100+ shipments)"),
    ]),
    ("freight", "Freight and cost", [
        title_box("fr_title", "Freight and cost",
                  "Freight per kg and freight as a share of commodity value, using only shipments where both numbers were recorded."),
        card("fr_card_freight", (20, 72, 400, 100), SH, "Freight Cost (USD)", "Freight cost (recorded)", {"labelPrecision": lit("2L")}),
        card("fr_card_perkg", (440, 72, 400, 100), SH, "Freight per kg (USD)", "Freight per kg"),
        card("fr_card_share", (860, 72, 400, 100), SH, "Freight % of Value", "Freight as % of commodity value"),
        bar("fr_bar_perkg_mode", (20, 188, 610, 270), MODE, measure(SH, "Freight per kg (USD)"),
            "Freight cost per kg by shipment mode"),
        bar("fr_bar_share_country", (650, 188, 610, 270), COUNTRY, measure(SH, "Freight % of Value (Top 10 Freight Countries)"),
            "Freight as % of value, top 10 countries by freight spend"),
        table_visual("fr_table_tradeoff", (20, 474, 1240, 236),
                     [MODE, measure(SH, "Shipments"), measure(SH, "Freight per kg (USD)"), measure(SH, "Freight % of Value"),
                      measure(SH, "% Late vs Schedule"), measure(SH, "Avg PO to Delivery Days")],
                     "Cost vs reliability by shipment mode"),
    ]),
    ("dataquality", "Data quality", [
        title_box("dq_title", "Data quality",
                  "What is missing or unreliable in the source data, and how it was handled."),
        card("dq_card_freight", (20, 72, 300, 100), SH, "% Freight Not Recorded", "Shipments without separate freight"),
        card("dq_card_weight", (340, 72, 300, 100), SH, "% Weight Not Recorded", "Shipments without weight"),
        card("dq_card_lines", (660, 72, 300, 100), LI, "Line Items", "Line items", {"labelDisplayUnits": lit("1D")}),
        card("dq_card_shipments", (980, 72, 280, 100), SH, "Shipments", "Shipments"),
        bar("dq_bar_freight", (20, 188, 610, 260), column(SH, "Freight Status"), measure(SH, "Shipments"),
            "Shipments by freight status"),
        bar("dq_bar_weight", (650, 188, 610, 260), column(SH, "Weight Status"), measure(SH, "Shipments"),
            "Shipments by weight status"),
        table_visual("dq_table_mode", (20, 464, 610, 246),
                     [MODE, measure(SH, "Shipments"), measure(SH, "Avg PO to Delivery Days"), measure(SH, "Median PO to Delivery Days")],
                     "Lead time by shipment mode: check 'Not recorded'"),
        textbox("dq_notes", (650, 464, 610, 246), [
            [("Cleaning steps", 12, True)],
            [("1. Imported as UTF-8 so names like Cote d'Ivoire stay intact.", 10, False)],
            [("2. Parsed two date formats (28-Sep-06 and 9/11/14) with an en-US locale.", 10, False)],
            [("3. Kept placeholder text (Date Not Captured, Freight Included in Commodity Cost) in status columns, then set the values to blank.", 10, False)],
            [("4. Split line items (10,324) from shipments (7,030): weight and freight are stored once per shipment.", 10, False)],
            [("5. Flag: shipments with no recorded mode average about 2 days from PO to delivery, which is not plausible. Treat that group with caution.", 10, False)],
        ]),
    ]),
]


def build_report():
    if RP.exists():
        shutil.rmtree(RP)
    write_json(RP / "definition.pbir", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0",
        "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}},
    })
    write_json(RP / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "Report", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": tag("report")},
    })
    d = RP / "definition"
    write_json(d / "version.json", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
        "version": "2.0.0",
    })
    write_json(d / "report.json", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.0.0/schema.json",
        "themeCollection": {"baseTheme": {
            "name": "CY20SU09",
            "reportVersionAtImport": {"visual": "1.8.53", "report": "2.0.53", "page": "1.3.53"},
            "type": "SharedResources"}},
        "objects": {"outspacePane": [{"properties": {"expanded": lit("false")}}]},
        "resourcePackages": [{"name": "SharedResources", "type": "SharedResources",
                              "items": [{"name": "CY20SU09", "path": "BaseThemes/CY20SU09.json", "type": "BaseTheme"}]}],
        "settings": {"useStylableVisualContainerHeader": True, "exportDataMode": "AllowSummarized",
                     "defaultDrillFilterOtherVisuals": True, "allowChangeFilterTypes": True, "useEnhancedTooltips": True},
    })
    theme_dst = RP / "StaticResources" / "SharedResources" / "BaseThemes" / "CY20SU09.json"
    theme_dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(THEME_SRC, theme_dst)

    write_json(d / "pages" / "pages.json", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
        "pageOrder": [p[0] for p in PAGES],
        "activePageName": PAGES[0][0],
    })
    for page_name, display, visuals in PAGES:
        write_json(d / "pages" / page_name / "page.json", {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json",
            "name": page_name, "displayName": display, "displayOption": "FitToPage", "height": 720, "width": 1280,
        })
        for z, (vname, vjson) in enumerate(visuals):
            vjson["position"]["z"] = z * 1000
            vjson["position"]["tabOrder"] = z * 1000
            write_json(d / "pages" / page_name / "visuals" / vname / "visual.json", vjson)


def build_pbip():
    write_json(ROOT / f"{NAME}.pbip", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
        "settings": {"enableAutoRecovery": True},
    })
    write(ROOT / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")


if __name__ == "__main__":
    build_semantic_model()
    build_report()
    build_pbip()
    print("Wrote", ROOT / f"{NAME}.pbip")
