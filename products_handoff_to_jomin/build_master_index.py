"""Generate _MASTER_PRODUCT_INDEX.xlsx — LIVE products only with per-product sheets + SharePoint hyperlinks.

Sources (all internal SharePoint, read via Microsoft 365 MCP):
- WELLNESS EXTRACT/DEPARTMENTS/COMPLIANCE/NPN'S/List of NPN and Product Development.xlsx
- WELLNESS EXTRACT/DEPARTMENTS/MARKETPLACE/AMAZON/AMAZON COMPLIANCE/Compliance Issues for all Amazon Marketplaces.xlsx
- wellness_extract_brand_brief.docx
- Wellness_Extract_Ecosystem_Combinations.xlsx
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SP = "https://wellnessextractca.sharepoint.com/sites/GoogleDrive/WELLNESS%20EXTRACT"

# Live products (one row per product line)
LIVE = [
    {
        "name": "Eannatto Tocotrienols (DeltaGold)",
        "brand": "Eannatto / Wellness Extract",
        "ingredient": "Annatto-derived Tocotrienol (DeltaGold)",
        "doses": "50 mg / 125 mg / 300 mg",
        "form": "Softgel (30 / 60 ct, packs 1/3/6)",
        "npn": "80121197 (50mg) · 80090692 (125mg) · 80121207 (300mg)",
        "category": "Antioxidant — Vitamin E",
        "channels": "Shopify WE+Eannatto · Amazon US/CA/UK/AU · DE-FBM",
        "asins": "B0B986W9VP, B09RG32WVQ, B07PY8RF69, B08KZJT3WH, B08N4CYMV5, B097DBVL8W, B0B9LHKYCQ, B0BG634KRK, B097DD7RWL, B0B9LJ9DB3, B08C76CGPS, B0F9XG4F97, B0F67F1QNY",
        "docs": [
            ("Tocotrienol + 150 Box - Canada - French.docx (label content)", f"{SP}/DEPARTMENTS/CONTENT/ACTIVE%20PRODUCTS/Tocotrienols+150/Box/"),
            ("Health_Canada_NNHPD_claim_rules.md", f"{SP}/PRODUCTS/ANNATTO%20PRODUCTS/New%20Purpose%20structure/10_compliance_and_claims/"),
            ("_OVERVIEW_regulatory_status.md", f"{SP}/PRODUCTS/ANNATTO%20PRODUCTS/New%20Purpose%20structure/05_monographs_regulatory/"),
            ("Google_Ads_Healthcare_Medicines.md", f"{SP}/PRODUCTS/ANNATTO%20PRODUCTS/New%20Purpose%20structure/11_ad_platform_policies/"),
            ("Annatto Products folder (root)", f"{SP}/PRODUCTS/ANNATTO%20PRODUCTS/"),
            ("Tocotrienol research library", "https://wellnessextractca-my.sharepoint.com/personal/viren_wellnessextract_com/Documents/Documents/Claude/Projects/tocotrienol%20re-search/paper/"),
            ("List of NPN and Product Development.xlsx", "https://wellnessextractca.sharepoint.com/sites/WETeam/Shared%20Documents/Wellness%20Extract/Departments/Compliance-Naresh/List%20of%20NPN%20and%20Product%20Development.xlsx"),
            ("Amazon Compliance Issues sheet", f"{SP}/DEPARTMENTS/MARKETPLACE/AMAZON/AMAZON%20COMPLIANCE/Compliance%20Issues%20for%20all%20Amazon%20Marketplaces.xlsx"),
            ("Amazon Influencer Outreach SOP v1.0 (Tocotrienol)", "https://wellnessextractca-my.sharepoint.com/personal/sarika_wellnessextract_com/Documents/Microsoft%20Teams%20Chat%20Files/Wellness_Extract_Amazon_Influencer_Outreach_SOP_v1.0.docx"),
            ("Objection handling Q&A — Eannatto DeltaGold", "https://wellnessextractca-my.sharepoint.com/personal/viren_wellnessextract_com/Documents/Documents/Claude/Projects/tocotrienol%20re-search/paper/12_marketing_toolkit/objection_handling_Q&A.md"),
        ],
    },
    {
        "name": "Eannatto 150+ (Tocotrienol + GG)",
        "brand": "Eannatto / Wellness Extract",
        "ingredient": "Tocotrienol 150 mg + Geranylgeraniol (GG-Gold) 75 mg",
        "doses": "150 mg Toco + 75 mg GG",
        "form": "Softgel (60 ct)",
        "npn": "80121195",
        "category": "Antioxidant + Cellular Energy",
        "channels": "Shopify WE+Eannatto · Amazon US/CA/UK/AU · DE-FBM · NL-FBM",
        "asins": "B0CQ569TQ7 · B0D79CFLMZ (CA bundle) · B0FDRDV69Z (DE FBM)",
        "docs": [
            ("Tocotrienol + 150 Box - Canada - French.docx", f"{SP}/DEPARTMENTS/CONTENT/ACTIVE%20PRODUCTS/Tocotrienols+150/Box/"),
            ("Tocotrienols+150 folder", f"{SP}/DEPARTMENTS/CONTENT/ACTIVE%20PRODUCTS/Tocotrienols+150/"),
            ("List of NPN and Product Development.xlsx", "https://wellnessextractca.sharepoint.com/sites/WETeam/Shared%20Documents/Wellness%20Extract/Departments/Compliance-Naresh/List%20of%20NPN%20and%20Product%20Development.xlsx"),
            ("Amazon Compliance Issues sheet", f"{SP}/DEPARTMENTS/MARKETPLACE/AMAZON/AMAZON%20COMPLIANCE/Compliance%20Issues%20for%20all%20Amazon%20Marketplaces.xlsx"),
        ],
    },
    {
        "name": "GG Essential — Geranylgeraniol",
        "brand": "Wellness Extract",
        "ingredient": "Geranylgeraniol (GG-Gold)",
        "doses": "150 mg",
        "form": "Softgel (60 ct)",
        "npn": "PENDING — flagged in NPN sheet (Dr Naresh follow-up)",
        "category": "CoQ10 / Cellular Energy support",
        "channels": "Amazon US/UK/AU · NL-FBM (GGE150-60-FBM-NL-V1)",
        "asins": "B0953TSGT8",
        "docs": [
            ("Amazon Compliance Issues sheet", f"{SP}/DEPARTMENTS/MARKETPLACE/AMAZON/AMAZON%20COMPLIANCE/Compliance%20Issues%20for%20all%20Amazon%20Marketplaces.xlsx"),
            ("List of NPN and Product Development.xlsx (GG row, NPN pending)", "https://wellnessextractca.sharepoint.com/sites/WETeam/Shared%20Documents/Wellness%20Extract/Departments/Compliance-Naresh/List%20of%20NPN%20and%20Product%20Development.xlsx"),
        ],
    },
    {
        "name": "Astaxanthin",
        "brand": "Wellness Extract",
        "ingredient": "Astaxanthin (variants with Toco + GG)",
        "doses": "9.7 mg / 12 mg",
        "form": "Softgel (60 ct)",
        "npn": "80120929 (9.7 mg) · 80137328 (12 mg)",
        "category": "Antioxidant — Carotenoid",
        "channels": "DE-FBM (AST97-60-FBM-DE-V1) · verify other markets",
        "asins": "B0FF5JGNGX (DE FBM)",
        "docs": [
            ("List of NPN and Product Development.xlsx", "https://wellnessextractca.sharepoint.com/sites/WETeam/Shared%20Documents/Wellness%20Extract/Departments/Compliance-Naresh/List%20of%20NPN%20and%20Product%20Development.xlsx"),
            ("Amazon Compliance Issues sheet (DE/NL)", f"{SP}/DEPARTMENTS/MARKETPLACE/AMAZON/AMAZON%20COMPLIANCE/Compliance%20Issues%20for%20all%20Amazon%20Marketplaces.xlsx"),
        ],
    },
    {
        "name": "Bio-Qunol (CoQ10 / Ubiquinol)",
        "brand": "Wellness Extract",
        "ingredient": "Ubiquinol (CoQ10)",
        "doses": "TBD",
        "form": "Softgel",
        "npn": "80113107",
        "category": "Antioxidant — CoQ10",
        "channels": "Verify storefront listings",
        "asins": "TBD",
        "docs": [
            ("List of NPN and Product Development.xlsx", "https://wellnessextractca.sharepoint.com/sites/WETeam/Shared%20Documents/Wellness%20Extract/Departments/Compliance-Naresh/List%20of%20NPN%20and%20Product%20Development.xlsx"),
        ],
    },
    {
        "name": "Bovine Colostrum Powder (Grass-Fed Grade A)",
        "brand": "Wellness Extract",
        "ingredient": "Bovine Colostrum (6-hour extracted, supplier ImmuneTree)",
        "doses": "60 g · 150 g",
        "form": "Powder",
        "npn": "80091452",
        "category": "Gut Health / Immune",
        "channels": "Shopify · Amazon US/CA/UK/AU · NL-FBM",
        "asins": "B0BBPL1WTJ · B0FB4TLKY9 (DE FBM)",
        "docs": [
            ("List of NPN and Product Development.xlsx (Colostrum row)", "https://wellnessextractca.sharepoint.com/sites/WETeam/Shared%20Documents/Wellness%20Extract/Departments/Compliance-Naresh/List%20of%20NPN%20and%20Product%20Development.xlsx"),
            ("Amazon Compliance — Colostrum negative reviews appeal (CLP150g)", f"{SP}/DEPARTMENTS/MARKETPLACE/AMAZON/AMAZON%20COMPLIANCE/Compliance%20Issues%20for%20all%20Amazon%20Marketplaces.xlsx"),
            ("Customer service email drafts", f"{SP}/DEPARTMENTS/CUSTOMER%20SERVICE/EMAIL%20DRAFTS/"),
        ],
    },
    {
        "name": "Molecular Hydrogen Tablets (H2)",
        "brand": "Wellness Extract",
        "ingredient": "Molecular Hydrogen (H2)",
        "doses": "30 / 60 tablets · Unflavoured + Raspberry",
        "form": "Effervescent tablet",
        "npn": "PENDING — flagged in NPN sheet",
        "category": "Hydration / Antioxidant",
        "channels": "Shopify (H2O_UFL_60_WEB_US_P2) · Amazon US (B0DC6QYHK5) · NL-FBM",
        "asins": "B0DC6QYHK5 (Raspberry 60ct)",
        "docs": [
            ("Amazon Compliance — H2 GMP cert appeal", f"{SP}/DEPARTMENTS/MARKETPLACE/AMAZON/AMAZON%20COMPLIANCE/Compliance%20Issues%20for%20all%20Amazon%20Marketplaces.xlsx"),
            ("Listing Compliance Amazon.docx", f"{SP}/DEPARTMENTS/MARKETPLACE/AMAZON/AMAZON%20COMPLIANCE/Listing%20Compliance%20Amazon.docx"),
        ],
    },
    {
        "name": "Cell D-Tox (CDX)",
        "brand": "Wellness Extract",
        "ingredient": "Humic-fulvic-zeolite (supplier: Sharpa Venture)",
        "doses": "2 oz · 4 oz · 3 g sachet (kids)",
        "form": "Liquid drops + sachet",
        "npn": "PENDING — flagged (Sharpa Venture supplier docs)",
        "category": "Body Detox",
        "channels": "Shopify · Amazon US (B0GF8DKF8W) · Amazon AU FBM",
        "asins": "B0GF8DKF8W",
        "docs": [
            ("Amazon Compliance — CDX GMP/Therapeutic Goods", f"{SP}/DEPARTMENTS/MARKETPLACE/AMAZON/AMAZON%20COMPLIANCE/Compliance%20Issues%20for%20all%20Amazon%20Marketplaces.xlsx"),
            ("List of NPN and Product Development.xlsx (Cell Dtox row, NPN to apply)", "https://wellnessextractca.sharepoint.com/sites/WETeam/Shared%20Documents/Wellness%20Extract/Departments/Compliance-Naresh/List%20of%20NPN%20and%20Product%20Development.xlsx"),
        ],
    },
]

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HYPER_FONT = Font(color="0563C1", underline="single")
THIN = Side(border_style="thin", color="B4B4B4")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header(ws):
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border = BORDER


def safe_sheet_name(name):
    bad = '[]:*?/\\'
    out = "".join(c for c in name if c not in bad)
    return out[:31]


wb = Workbook()
master = wb.active
master.title = "Master Product Index"

cols = ["product_name", "brand", "primary_ingredient", "doses", "form_factor",
        "NPN_number", "category", "channels", "ASINs", "details_sheet"]
master.append(cols)
style_header(master)

for p in LIVE:
    sheet_name = safe_sheet_name(p["name"])
    master.append([
        p["name"], p["brand"], p["ingredient"], p["doses"], p["form"],
        p["npn"], p["category"], p["channels"], p["asins"], sheet_name,
    ])
    last = master.cell(row=master.max_row, column=10)
    last.hyperlink = f"#'{sheet_name}'!A1"
    last.font = HYPER_FONT

widths = [34, 28, 38, 28, 26, 40, 30, 50, 48, 30]
for i, w in enumerate(widths, 1):
    master.column_dimensions[chr(64 + i)].width = w
master.row_dimensions[1].height = 24

# Per-product sheets
for p in LIVE:
    ws = wb.create_sheet(safe_sheet_name(p["name"]))
    ws.append([p["name"]])
    ws["A1"].font = Font(bold=True, size=14, color="1F3864")
    ws.append([])
    meta = [
        ("Brand", p["brand"]),
        ("Primary ingredient", p["ingredient"]),
        ("Doses", p["doses"]),
        ("Form factor", p["form"]),
        ("NPN", p["npn"]),
        ("Category", p["category"]),
        ("Channels", p["channels"]),
        ("ASINs", p["asins"]),
    ]
    for k, v in meta:
        ws.append([k, v])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    ws.append([])
    ws.append(["Document title", "SharePoint link"])
    style_header_row = ws.max_row
    for cell in ws[style_header_row]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
    for title, url in p["docs"]:
        ws.append([title, url])
        link_cell = ws.cell(row=ws.max_row, column=2)
        link_cell.hyperlink = url
        link_cell.font = HYPER_FONT
    ws.append([])
    back = ws.cell(row=ws.max_row + 1, column=1)
    back.value = "← Back to Master Index"
    back.hyperlink = "#'Master Product Index'!A1"
    back.font = HYPER_FONT
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 110

# Sources sheet
src = wb.create_sheet("Sources")
src.append(["source", "sharepoint_path"])
style_header(src)
src.append(["List of NPN and Product Development.xlsx (AUTHORITATIVE)",
            "https://wellnessextractca.sharepoint.com/sites/WETeam/Shared%20Documents/Wellness%20Extract/Departments/Compliance-Naresh/List%20of%20NPN%20and%20Product%20Development.xlsx"])
src.append(["Amazon Compliance Issues sheet",
            f"{SP}/DEPARTMENTS/MARKETPLACE/AMAZON/AMAZON%20COMPLIANCE/Compliance%20Issues%20for%20all%20Amazon%20Marketplaces.xlsx"])
src.append(["Wellness_Extract_Ecosystem_Combinations.xlsx",
            f"{SP}/DEPARTMENTS/PRODUCT%20DEVELOPMENT/Wellness_Extract_Ecosystem_Combinations.xlsx"])
src.append(["Wellness Extract Brand Brief (CEO direction, April 2026)",
            "https://wellnessextractca-my.sharepoint.com/personal/viren_wellnessextract_com/Documents/Microsoft%20Teams%20Chat%20Files/wellness_extract_brand_brief.docx"])
src.append(["Annatto Products folder (root, Tocotrienol research library)",
            f"{SP}/PRODUCTS/ANNATTO%20PRODUCTS/"])
src.append(["Tocotrienol research library (Viren / Claude Projects)",
            "https://wellnessextractca-my.sharepoint.com/personal/viren_wellnessextract_com/Documents/Documents/Claude/Projects/tocotrienol%20re-search/paper/"])
src.append(["Customer Service email drafts",
            f"{SP}/DEPARTMENTS/CUSTOMER%20SERVICE/EMAIL%20DRAFTS/"])
for col, w in [("A", 60), ("B", 130)]:
    src.column_dimensions[col].width = w
for row in src.iter_rows(min_row=2):
    row[1].font = HYPER_FONT
    row[1].hyperlink = row[1].value

out = "/home/user/Project-Product-files-document/products_handoff_to_jomin/_MASTER_PRODUCT_INDEX.xlsx"
wb.save(out)
print(f"Wrote {out}")
print(f"Live products: {len(LIVE)}")
print(f"Sheets: {len(wb.sheetnames)} ({wb.sheetnames})")
