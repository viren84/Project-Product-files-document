"""Generate _MASTER_PRODUCT_INDEX.xlsx from reconciled internal sources.

Sources used:
- WELLNESS EXTRACT/DEPARTMENTS/COMPLIANCE/NPN'S/List of NPN and Product Development.xlsx (authoritative NPN registry)
- WELLNESS EXTRACT/DEPARTMENTS/MARKETPLACE/AMAZON/AMAZON COMPLIANCE/Compliance Issues for all Amazon Marketplaces.xlsx (active SKUs/ASINs across US/CA/UK/DE/AU/NL)
- wellness_extract_brand_brief.docx (product portfolio)
- Wellness_Extract_Ecosystem_Combinations.xlsx (ingredient-benefit map)
- Web search of amazon.com / .ca / .co.uk / .com.au storefronts (storefront URLs verified)
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

wb = Workbook()
ws = wb.active
ws.title = "Master Product Index"

HEADERS = [
    "product_name", "brand", "primary_ingredient", "dose", "form_factor",
    "NPN_number", "status", "category", "amazon_US_ASIN", "amazon_CA",
    "amazon_UK", "amazon_AU", "amazon_DE_ASIN", "shopify_we", "shopify_eannatto",
    "amazon_NL_status", "source_refs",
]

# Active products with confirmed NPN + visible on at least one storefront
# brand: WE = Wellness Extract, EA = Eannatto sub-brand (annatto-derived line)
ROWS = [
    # ====== EANNATTO (TOCOTRIENOL) LINE — confirmed live on Shopify + Amazon US/CA/UK/AU ======
    ["Eannatto Tocotrienols 50 mg",  "Eannatto / Wellness Extract", "Tocotrienol (Annatto-DeltaGold)", "50 mg",  "Softgel (30 ct)",
     "80121197", "Active", "Antioxidants — Vitamin E",
     "B0B986W9VP", "yes", "—", "—", "—", "yes", "yes", "—",
     "NPN list; Amazon US listing"],

    ["Eannatto Tocotrienols 125 mg (30 ct)", "Eannatto / Wellness Extract", "Tocotrienol (Annatto-DeltaGold)", "125 mg", "Softgel (30 ct)",
     "80090692", "Active", "Antioxidants — Vitamin E",
     "B09RG32WVQ", "B09RG32WVQ", "B09RG32WVQ", "—", "B0F9XG4F97 (FBM)", "yes", "yes", "needs label NL",
     "NPN list; Compliance sheet; Amazon US/CA/UK"],

    ["Eannatto Tocotrienols 125 mg (60 ct)", "Eannatto / Wellness Extract", "Tocotrienol (Annatto-DeltaGold)", "125 mg", "Softgel (60 ct)",
     "80090692", "Active", "Antioxidants — Vitamin E",
     "B07PY8RF69", "B07PY8RF69", "B07PY8RF69", "B08C76CGPS", "—", "yes", "yes", "ETO125-60-FBA-NL-V1+1",
     "NPN list; Compliance sheet; Amazon US/CA/UK/AU"],

    ["Eannatto Tocotrienols 125 mg (60 ct, Pack of 3)", "Eannatto / Wellness Extract", "Tocotrienol (Annatto-DeltaGold)", "125 mg", "Softgel (60 ct × 3)",
     "80090692", "Active", "Antioxidants — Vitamin E",
     "B08KZJT3WH", "—", "—", "—", "—", "yes", "yes", "—",
     "Amazon US listing"],

    ["Eannatto Tocotrienols 125 mg (60 ct, Pack of 6)", "Eannatto / Wellness Extract", "Tocotrienol (Annatto-DeltaGold)", "125 mg", "Softgel (60 ct × 6)",
     "80090692", "Active", "Antioxidants — Vitamin E",
     "B08N4CYMV5", "—", "—", "—", "—", "yes", "yes", "—",
     "Amazon US listing"],

    ["Eannatto Tocotrienols 300 mg (30 ct)", "Eannatto / Wellness Extract", "Tocotrienol (Annatto-DeltaGold)", "300 mg", "Softgel (30 ct)",
     "80121207", "Active", "Antioxidants — Vitamin E",
     "B097DBVL8W", "B0B9LHKYCQ", "B0BG634KRK", "B0BG634KRK", "B0F67F1QNY (FBM)", "yes", "yes", "—",
     "NPN list; Compliance sheet; Amazon US/CA/UK/AU/DE"],

    ["Eannatto Tocotrienols 300 mg (60 ct)", "Eannatto / Wellness Extract", "Tocotrienol (Annatto-DeltaGold)", "300 mg", "Softgel (60 ct)",
     "80121207", "Active", "Antioxidants — Vitamin E",
     "B097DD7RWL", "—", "B097DD7RWL", "B0B9LJ9DB3", "—", "yes", "yes", "—",
     "Amazon US/UK/AU"],

    ["Eannatto 150+ Tocotrienols + GG", "Eannatto / Wellness Extract", "Tocotrienol 150 mg + Geranylgeraniol (GG-Gold) 75 mg", "150 mg Toco + 75 mg GG", "Softgel (60 ct)",
     "80121195", "Active", "Antioxidants — Vitamin E + GG",
     "B0CQ569TQ7", "B0CQ569TQ7", "B0CQ569TQ7", "B0CQ569TQ7", "B0FDRDV69Z (FBM)", "yes", "yes", "needs label NL",
     "NPN list; Compliance sheet; Brand brief Q2; All 4 Amazon storefronts"],

    ["Eannatto 150+ Vitamin E Tocotrienols Supplement (CA bundle)", "Eannatto / Wellness Extract", "Tocotrienol + GG", "150 mg / 75 mg", "Softgel (60 ct)",
     "80121195", "Active", "Antioxidants — Vitamin E + GG",
     "—", "B0D79CFLMZ", "—", "—", "—", "yes", "yes", "—",
     "Amazon CA bundle SKU"],

    # ====== GG ESSENTIAL ======
    ["GG Essential — Geranylgeraniol", "Wellness Extract", "Geranylgeraniol (GG-Gold)", "150 mg", "Softgel (60 ct)",
     "PENDING (need to apply — flagged in NPN sheet)", "Active (Amazon live; NPN follow-up Dr Naresh)", "Antioxidants — CoQ10 support",
     "B0953TSGT8", "—", "B0953TSGT8", "B0953TSGT8", "GGE150-60-FBM-NL-V1", "yes", "—", "GGE150-60-FBM-NL-V1",
     "Compliance sheet (NL); Amazon US/UK/AU"],

    # ====== ASTAXANTHIN ======
    ["Astaxanthin", "Wellness Extract", "Astaxanthin (with Toco + GG combo)", "9.7 mg / 12 mg variants", "Softgel (60 ct)",
     "80120929 (9.7 mg) / 80137328 (12 mg)", "Active", "Antioxidants — Carotenoid",
     "TBD (verify ASIN)", "TBD", "TBD", "TBD", "B0FF5JGNGX (FBM)", "yes", "—", "AST97-60-FBM-NL-V1",
     "NPN list; Compliance sheet (DE/NL)"],

    # ====== COQ10 ======
    ["Bio-Qunol (CoQ10)", "Wellness Extract", "Ubiquinol (CoQ10)", "TBD", "Softgel",
     "80113107", "Active (Health Canada licensed; verify storefront listing)", "Antioxidants — CoQ10",
     "TBD", "TBD", "TBD", "TBD", "TBD", "TBD", "—", "—",
     "NPN list"],

    # ====== COLOSTRUM ======
    ["Bovine Colostrum Powder (Grass-Fed Grade A)", "Wellness Extract", "Bovine Colostrum (6-hour)", "150 g", "Powder",
     "80091452", "Active", "Gut Health / Immune",
     "B0BBPL1WTJ", "yes", "yes", "yes", "B0FB4TLKY9 (FBM)", "yes", "—", "CLP150-FBA-NL-V1+1",
     "NPN list; Compliance sheet (US negative review issue resolved); Brand brief Q7"],

    ["Bovine Colostrum Powder — 60 g", "Wellness Extract", "Bovine Colostrum", "60 g", "Powder",
     "80091452", "Active", "Gut Health / Immune",
     "TBD (CLP-60g-FBA-US-P1)", "TBD", "TBD", "TBD", "TBD", "yes", "—", "—",
     "Compliance sheet SKU CLP-60g-FBA-US-P1"],

    # ====== H2 / MOLECULAR HYDROGEN ======
    ["Molecular Hydrogen Tablets — Unflavoured", "Wellness Extract", "Molecular hydrogen (H2)", "30 / 60 tablets", "Tablet",
     "PENDING (need to apply — flagged in NPN sheet)", "Active (Amazon US live; NPN follow-up)", "Hydration / Antioxidant",
     "TBD (H2-UFL SKUs)", "TBD", "TBD", "TBD", "TBD", "yes", "—", "H2-UFL-30/60-FBM-NL-V1",
     "Compliance sheet (NL); Outlook order SKU H2O_UFL_60_WEB_US_P2"],

    ["Molecular Hydrogen Tablets — Raspberry", "Wellness Extract", "Molecular hydrogen (H2)", "60 tablets", "Tablet",
     "PENDING", "Active US (compliance issue ongoing — GMP cert)", "Hydration / Antioxidant",
     "B0DC6QYHK5", "—", "—", "—", "—", "yes", "—", "—",
     "Compliance sheet SKU H2-RSP-60-FBA-US-P1+1"],

    # ====== CELL D-TOX ======
    ["Cell D-Tox 4 oz", "Wellness Extract", "Humic-fulvic-zeolite", "4 oz", "Liquid (drops)",
     "PENDING (Sharpa Venture supplier)", "Active (Amazon US/AU live; NPN follow-up)", "Body Detox",
     "B0GF8DKF8W", "—", "—", "B0GF8DKF8W (FBM)", "—", "yes", "—", "—",
     "Compliance sheet SKU CDX-4OZ-FBA-US-P1, CDX-4OZ-FBA-AUS-P1"],

    ["Cell D-Tox 2 oz", "Wellness Extract", "Humic-fulvic-zeolite", "2 oz", "Liquid (drops)",
     "PENDING", "Active US", "Body Detox",
     "TBD (CDX2oz-FBA-US-P1+1)", "—", "—", "—", "—", "yes", "—", "—",
     "Compliance sheet SKU CDX2oz-FBA-US-P1+1"],

    # ====== FLAGGED — review status ======
    ["Premium Coffee Enema Kit", "Wellness Extract", "Organic coffee enema (Class II medical device CA)", "—", "Kit",
     "—", "FLAGGED — listing removed Amazon.ca Jan-2026 (medical device licence required); status unclear",
     "Lifestyle / Detox",
     "TBD", "REMOVED Jan-2026", "TBD", "TBD", "TBD", "TBD", "—", "—",
     "Compliance sheet (CA Jan-2026 violation)"],
]

# Pipeline / not yet active (NPN unfiled or product not launched) — separate sheet
PIPELINE = [
    ["WE B12", "80120932", "Capsule"],
    ["Elderberry", "80120933", "Capsule"],
    ["Elderberry 2", "80120935", "Capsule"],
    ["Elderberry 3 (64:1)", "80120936", "Capsule"],
    ["WE Biotin", "80122907", "Capsule"],
    ["WE Curcumin", "80122196", "Capsule"],
    ["WE Folic Acid", "80122904", "Capsule"],
    ["WE Niacin", "80122191", "Capsule"],
    ["WE Niacin (alt)", "80122905", "Capsule"],
    ["WE Quercetin", "80120938", "Capsule"],
    ["WE Vitamin A", "80122910", "Capsule"],
    ["WE Vitamin B5 (Pantothenic Acid)", "80122909", "Capsule"],
    ["WE Reishi Ganoderma", "80122886", "Capsule"],
    ["L-Theanine", "80120930", "Capsule"],
    ["WE Lion's Mane", "80122891", "Capsule"],
    ["WE Maitake", "80122887", "Capsule"],
    ["WE Cordyceps", "80122888", "Capsule"],
    ["WE Psyllium — Powder", "80122189", "Powder"],
    ["WE Psyllium — Capsule", "80122186", "Capsule"],
    ["Bromelain", "80120937", "Capsule"],
    ["WE Alpha Lipoic Acid 600 mg", "80122192", "Capsule"],
    ["Liquid Magnesium", "80119181", "Liquid"],
    ["WE Magnesium Glycinate", "80120931", "Capsule"],
    ["Selenium WE 200", "80116403", "Tablet"],
    ["WE Zinc Chelated", "80119169", "Tablet"],
]

PIPELINE_NEW = [
    ["Manuka Honey", "Need to apply", "Liquid", "Supplier: DownUnder Honey (NZ)"],
    ["Human Lactoferrin (Eferra™)", "Need to apply", "Capsule", "Supplier: Helaina"],
    ["Liposomal Vitamin C / B / D / Glutathione / CoQ10 / Curcumin", "No NPN", "Liquid/capsule", "Supplier: Suzhou Leadwell Biotech"],
    ["Omegia™ (vegan Omega 3-6-7-9 from sea buckthorn)", "No NPN", "Softgel", "Supplier: Puredia"],
    ["CyanthOx™", "No NPN", "—", "Supplier: Osage Food"],
    ["Colostrum Kids 60 g sachet", "Applied (Sharpa Venture)", "Sachet", "Pet/kids extension"],
    ["Colostrum Pets 150 g jar", "Applied", "Jar", "Pet line"],
    ["WE ProBiome (range)", "—", "Capsule", "Mentioned in brand brief Q7 as launching"],
]


def styled_header(ws, row=1):
    fill = PatternFill("solid", fgColor="1F3864")
    font = Font(bold=True, color="FFFFFF")
    for cell in ws[row]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


ws.append(HEADERS)
styled_header(ws)
for row in ROWS:
    ws.append(row)
for col_idx in range(1, len(HEADERS) + 1):
    ws.column_dimensions[chr(64 + col_idx)].width = 22

ws2 = wb.create_sheet("Pipeline — NPN Filed (not on storefronts yet)")
ws2.append(["product_name", "NPN_number", "form"])
styled_header(ws2)
for r in PIPELINE:
    ws2.append(r)
for col_idx in range(1, 4):
    ws2.column_dimensions[chr(64 + col_idx)].width = 32

ws3 = wb.create_sheet("Pipeline — In Development")
ws3.append(["product_name", "NPN_status", "form", "notes"])
styled_header(ws3)
for r in PIPELINE_NEW:
    ws3.append(r)
for col_idx in range(1, 5):
    ws3.column_dimensions[chr(64 + col_idx)].width = 32

ws4 = wb.create_sheet("Sources")
ws4.append(["source", "sharepoint_path"])
styled_header(ws4)
sources = [
    ("List of NPN and Product Development.xlsx (AUTHORITATIVE)",
     "/sites/WETeam/Shared Documents/Wellness Extract/Departments/Compliance-Naresh/"),
    ("Compliance Issues for all Amazon Marketplaces.xlsx",
     "/sites/GoogleDrive/WELLNESS EXTRACT/DEPARTMENTS/MARKETPLACE/AMAZON/AMAZON COMPLIANCE/"),
    ("wellness_extract_brand_brief.docx",
     "viren_wellnessextract_com / Microsoft Teams Chat Files/"),
    ("Wellness_Extract_Ecosystem_Combinations.xlsx",
     "/sites/GoogleDrive/WELLNESS EXTRACT/DEPARTMENTS/PRODUCT DEVELOPMENT/"),
    ("Master Listing Sheet.xlsm (could not read — macro-enabled .xlsm)",
     "anmol_sangwan_wellnessextract_com /Documents/"),
    ("Amazon storefronts (verified live)",
     "amazon.com / .ca / .co.uk / .com.au /stores/WellnessExtract"),
    ("Shopify storefronts (could not auto-fetch — Shopify blocks bot)",
     "wellnessextract.com / wellnessextract.ca / eannatto.com"),
]
for s in sources:
    ws4.append(s)
ws4.column_dimensions["A"].width = 60
ws4.column_dimensions["B"].width = 80

out = "/home/user/Project-Product-files-document/products_handoff_to_jomin/_MASTER_PRODUCT_INDEX.xlsx"
wb.save(out)
print(f"Wrote {out}")
print(f"Active rows: {len(ROWS)}")
print(f"Pipeline (NPN filed): {len(PIPELINE)}")
print(f"Pipeline (in dev): {len(PIPELINE_NEW)}")
