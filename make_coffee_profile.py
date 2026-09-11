"""Build 'Coffee Product Profile' - a full narrative product profile mirroring
the 'Tea Product Profile template.docx' (13-chapter structure) but for coffee,
populated with researched facts from credible sources and the ITC Trade Map
matrices in the Coffee\\ folder.

Outputs (both in the 'product profile/' directory next to the tea template):
  - product profile/Coffee Product Profile.docx
  - product profile/Coffee Product Profile TABLES.xlsx

Run:  python make_coffee_profile.py
"""

import json
import os
import sys

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from generate_product_profile import (  # noqa: E402
    ProfileBuilder, ProfileData, _ranked_rows, short_label, display,
    usd_phrase, write_excel_deliverable)
from country_names import fix_label  # noqa: E402

CFG = json.load(open(os.path.join(BASE_DIR, "config",
                                  "product_profile_coffee.json"),
                     encoding="utf-8"))
DATA_DIR = os.path.join(BASE_DIR, CFG.get("data_dir", "Coffee"))
OUT_DOC = os.path.join(BASE_DIR, "product profile", "Coffee Product Profile.docx")
OUT_XLS = os.path.join(BASE_DIR, "product profile",
                       "Coffee Product Profile TABLES.xlsx")


def _field(doc, code, placeholder="Update this field (Ctrl+A, F9)."):
    p = doc.add_paragraph()
    r = p.add_run()
    f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), "begin"); r._r.append(f)
    r = p.add_run()
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
    it.text = code
    r._r.append(it)
    r = p.add_run()
    s = OxmlElement("w:fldChar"); s.set(qn("w:fldCharType"), "separate")
    r._r.append(s)
    r.add_text(placeholder)
    r = p.add_run()
    e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end"); r._r.append(e)
    return p


def _hyperlink(paragraph, text, url):
    """Insert a clickable hyperlink into ``paragraph`` under ``text``."""
    part = paragraph.part
    r_id = part.relate_to(
        url, "http://schemas.openxmlformats.org/officeDocument/2006/"
             "relationships/hyperlink", is_external=True)
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    style = OxmlElement("w:rStyle"); style.set(qn("w:val"), "Hyperlink")
    rPr.append(style)
    new_run.append(rPr)
    t = OxmlElement("w:t"); t.text = text
    new_run.append(t)
    hl.append(new_run)
    paragraph._p.append(hl)
    return hl


def _static_table(b, rows, widths=None, header=True):
    """Plain bordered table for narrative-only content (players, institutions,
    incentives, programmes, interventions). 'rows' = list of row lists."""
    tbl = b.doc.add_table(rows=len(rows), cols=len(rows[0]),
                          style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = tbl.rows[ri].cells[ci]
            cell.text = "" if val is None else str(val)
            for p in cell.paragraphs:
                p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if ci == 0
                               and header else WD_ALIGN_PARAGRAPH.LEFT)
                for run in p.runs:
                    if ri == 0:
                        run.font.bold = True
    return tbl


def _cap(b, code, *args):
    return b._next_table(code, CFG.get("source"))


def main():
    data = ProfileData(DATA_DIR, CFG.get("include_codes"),
                       CFG.get("family_title"))
    years = data.years
    rev = data.review_year

    b = ProfileBuilder(CFG)
    doc = b.doc

    kenya_2025 = sum(r["years"].get(rev) or 0.0 for r in data.members)
    dest = data.destinations()
    dest_top = dest[:3]
    dest_total = sum(r["years"].get(rev) or 0.0 for r in dest)
    green = next((r for r in data.members if str(r["code"]).startswith("090111")),
                 None)

    # ------------------------------------------------------------------ cover
    for _ in range(3):
        b.doc.add_paragraph()
    p = b.doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("COFFEE PRODUCT PROFILE"); r.bold = True
    r.font.size = Pt(30)
    p = b.doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("KENYA'S COFFEE SECTOR"); r.bold = True; r.font.size = Pt(18)
    p = b.doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(str(rev)).bold = True
    b.doc.add_paragraph()
    p = b.doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Kenya Export Promotion and Branding Agency (KEPROBA)")
    p = b.doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Research and Innovation Directorate (RI)")
    b.doc.add_page_break()

    # -------------------------------------------------------------- disclaimer
    b.doc.add_heading("Disclaimer", level=1)
    b.add_para(CFG["intro"][0])
    for para in (
        "The information presented in this product profile has been compiled "
        "from various secondary sources, including the International Trade "
        "Centre (ITC) Trade Map, the United Nations Food and Agriculture "
        "Organization, the United Nations Industrial Development Organization "
        "(UNIDO), the United States Department of Agriculture (USDA) Foreign "
        "Agricultural Service, the International Coffee Organization (ICO), "
        "and Kenya's Ministry of Agriculture and Livestock Development, and is "
        "intended for general reference purposes only. While efforts have been "
        "made to ensure the accuracy and reliability of the data, KEPROBA "
        "assumes no responsibility for any errors, omissions, or "
        "discrepancies.",
        "DISCLAIMER: The information contained in this product profile is for "
        "general information only and is not intended to provide technical or "
        "legal advice. This profile is prepared for the general information of "
        "the public and other interested readers. The information contained in "
        "this coffee profile should not be treated as a guide for decision "
        "making and planning but general information about the sector. We do "
        "not accept responsibility or liability to users or any third parties "
        "in relation to use of this product profile or its contents.",
        "For any further information or clarification on this profile, please "
        "contact the Kenya Export Promotion and Branding Agency through "
        "chiefexe@brand.ke",
    ):
        b.add_para(para)
    b.doc.add_page_break()

    # ---------------------------------------------------------------- foreword
    b.doc.add_page_break()
    b.doc.add_heading("FOREWORD BY THE CHIEF EXECUTIVE OFFICER, KENYA EXPORT "
                      "PROMOTION AND BRANDING AGENCY (KEPROBA)", level=1)
    for para in (
        "Kenya is globally celebrated for the quality of its coffee. Grown on "
        "smallholder farms and estates in the volcanic highlands, Kenyan coffee "
        "is prized the world over for its bright acidity, full body and "
        "distinctive fruity and floral character, and consistently attracts "
        "premium prices at auction, including the highest prices paid for "
        "African coffees in most seasons.",
        "At the Kenya Export Promotion and Branding Agency (KEPROBA), we are "
        "proud to support and promote Kenya's coffee industry, ensuring that "
        "our coffee reaches new markets while maintaining its high standards. "
        "Through strategic partnerships, investments in value addition, and "
        "adherence to global sustainability and traceability practices, we are "
        "positioning Kenya to capture more value from the world's fastest "
        "growing premium beverage market.",
        "This product profile highlights the status, performance and "
        "opportunities of Kenya's coffee sector. By building on our strengths "
        "and addressing the challenges we face, we can restore Kenya's "
        "coffee sector to the leading position it occupied in the 1980s and "
        "secure greater prosperity for the more than 800,000 smallholder "
        "families whose livelihoods depend on the crop.",
        "Together, we will continue to advance Kenya's coffee industry, "
        "creating sustainable prosperity for our people while delighting "
        "coffee lovers around the world.",
        "Thank you for your continued support.",
        "Warm regards,",
        "FLOICE MUKABANA",
        "CHIEF EXECUTIVE OFFICER",
        "KENYA EXPORT PROMOTION AND BRANDING AGENCY (KEPROBA)",
    ):
        b.add_para(para)
    b.doc.add_page_break()

    # --------------------------------------------------------------- copyright
    b.add_para("\u00a9 [" + str(rev) + "] Coffee Product Profile. All rights "
               "reserved.")
    b.add_para("This Coffee Product Profile, including all articles, images, "
               "graphics, and other content, is protected by copyright laws. "
               "No part of this publication may be reproduced, distributed, "
               "transmitted, or stored in any form or by any means, electronic "
               "or mechanical, including photocopying, recording, or by any "
               "information storage and retrieval system, without prior "
               "written permission from the Kenya Export Promotion and "
               "Branding Agency (KEPROBA).")
    b.add_para("Unauthorized use of any content from this publication is "
               "strictly prohibited and may result in legal action. For "
               "permission requests, please contact the Kenya Export Promotion "
               "and Branding Agency at chiefexe@brand.ke")

    b.doc.add_heading("EDITORIAL TEAM", level=2)
    b.add_para("Editor in Chief\u2002\u2002Maureen Mambo")
    b.add_para("Editors\u2002\u2002\u2002\u2002\u2002\u2002Rebecca R. "
               "Mpaayei-Saruni; David S. Yamina")
    b.add_para("Contributors\u2002\u2002Rebecca Mpaayei-Saruni, David S. "
               "Yamina, Peris Isaboke, Lynda Koske, Sammy Mutwiri, Anthony "
               "Gathambo, Sarah Wandia")
    b.add_para("Design and Layout\u2002Sam Njaaga, Kevin Walter")
    b.doc.add_page_break()

    # ------------------------------------------------------------------- front
    b.doc.add_heading("Table of Contents", level=1)
    _field(doc, 'TOC \\o "1-2" \\h \\z \\u')
    b.doc.add_heading("List of Tables", level=1)
    _field(doc, 'TOC \\t "Caption Table;1" \\h \\z')
    b.doc.add_page_break()

    # ===================================================================== 1.0
    b.doc.add_heading("1.0\tEXECUTIVE SUMMARY", level=1)
    b.add_para(
        "Kenya is one of the world's most respected producers of high-quality "
        "arabica coffee and a well-known origin in the global specialty coffee "
        "market. Coffee is cultivated in 33 counties by more than 800,000 "
        "smallholder farmers and about 3,000 estates, on roughly 110,000 "
        "hectares, and arabica accounts for over 99 percent of national "
        "production (Ministry of Agriculture and Livestock Development, "
        "Coffee Development and Marketing Strategy 2024\u20132029; UNIDO/FAO, "
        "Kenya Coffee Value Chain Analysis 2025). Smallholders contribute over "
        "70 percent of production, and about 80 percent of Kenya's coffee is "
        "marketed through producer co-operatives (USDA FAS, Kenya Coffee "
        "Annual 2025).")
    b.add_para(
        "Coffee is a leading agricultural export earner. In %d, Kenya exported "
        "coffee (HS 0901) worth about USD %.0f million (International Trade "
        "Centre, Trade Map). The sector contributes substantially to "
        "agricultural export revenues and supports the livelihoods of around "
        "six million people across the value chain (FAO; UNIDO 2025). Kenya "
        "currently accounts for under one percent of world coffee exports by "
        "volume, but its auction prices are consistently among the highest in "
        "the world, rewarding quality and offering a clear path to value "
        "capture through specialty marketing and value addition." % (rev,
                                                                       display(kenya_2025)))
    b.add_para(
        "The sector's main challenges include low and declining yields "
        "(about 475 kg per hectare against 970 kg/ha achieved in 1963), "
        "climate variability and disease pressure, high input and financing "
        "costs, delayed farmer payments, and an evolving regulatory "
        "environment (UNIDO 2025; USDA FAS). Opportunities lie in the ongoing "
        "government expansion programme, productivity recovery on existing "
        "trees, compliance-driven market access (including the EU "
        "Deforestation Regulation), direct trade and certifications, and "
        "domestic roasting and branding.")
    b.add_para(
        "The future of the coffee sector is positive but requires sustained "
        "coordination among the Ministry of Agriculture and Livestock "
        "Development, the re-established Coffee Board of Kenya, co-operative "
        "societies, millers, marketers and exporters, research institutions "
        "and county governments, alongside investment in productivity, "
        "resilience and value addition.")

    # ===================================================================== 2.0
    b.doc.add_heading("2.0\tOVERVIEW OF THE COFFEE SECTOR", level=1)
    b.doc.add_heading("Introduction", level=2)
    b.add_para(
        "Kenya's coffee sector is a well-structured value chain that spans "
        "production, wet and dry processing, marketing and export. The chain "
        "is anchored by smallholder farmers organised in co-operatives, "
        "alongside estates, and it supports the livelihoods of more than six "
        "million people directly and indirectly. Coffee is one of the priority "
        "crops identified in Kenya's Bottom-Up Economic Transformation Agenda "
        "(BETA) to support export growth and revenue generation "
        "(FAO Investment Roundtable on Coffee Value Chains, 2025).")
    b.add_para(
        "Kenyan coffee is grown on deep, well-drained red volcanic soils at "
        "elevations of between 1,400 and 2,200 metres above sea level, with "
        "temperatures of 15\u201324\u00b0C. Nearly all production is rain-fed. "
        "About 90 percent of Kenya's coffee is wet-processed (fully washed) at "
        "washing stations owned by co-operative societies and estates, while "
        "the remaining 10 percent is dry-processed into mbuni (unwashed, dried "
        "cherry). Wet processing is the basis of Kenya's bright, "
        "clean-cupping 'mild arabica' reputation (UNIDO/FAO, Kenya Coffee "
        "Value Chain Analysis 2025).")
    b.add_para(
        "The high value of Kenyan coffee is anchored in its varieties and "
        "cupping character. Kenya is one of the few origins where run-of-crop "
        "coffee is auctioned by grade, with the large-bean grades "
        "(AA, AB, PB peaberry) attracting the highest prices. Home-grown "
        "varieties such as SL28, SL34 and K7, together with the newer Ruiru 11 "
        "and Batian hybrids, deliver the floral, bright acidity and berry and "
        "citrus notes for which Kenyan coffees are known.")
    b.doc.add_heading("Export market performance", level=2)
    b.add_para(
        "Kenya's coffee exports over the %d\u2013%d review period (ITC Trade "
        "Map) show sustained demand from established European and Asian "
        "roasting markets. The European Union absorbs over half of Kenya's "
        "coffee exports, making continued alignment with EU market "
        "regulations \u2013 including the EU Regulation on Deforestation-free "
        "Products (EUDR) \u2013 essential to maintaining market access "
        "(UNIDO/FAO 2025). Within Kenya, the Nairobi Coffee Exchange (NCE) "
        "remains the primary price-discovery platform, complemented by a "
        "growing share of direct-sale and specialty contracts." % (data.start_year,
                                                                   rev))
    b.doc.add_page_break()

    # ===================================================================== 3.0
    b.doc.add_heading("3.0\tKEY PLAYERS IN THE COFFEE VALUE CHAIN", level=1)
    b.add_para(
        "The Kenyan coffee value chain involves several key players, each "
        "contributing to the process from cultivation to consumption. The main "
        "actors are:")
    _cap(b, "Key Players in the Coffee Value Chain")
    _static_table(b, [
        ["No.", "Coffee Key Players", "Function"],
        ["1", "Smallholder Farmers", "Cultivate coffee on small plots (mostly "
               "under two hectares) clustered in co-operative societies; "
               "contribute over 70% of national production."],
        ["2", "Large-scale Coffee Estates", "Owned by plantations and "
               "companies; control production, wet milling and marketing, and "
               "lead in mechanisation."],
        ["3", "Co-operative Societies", "Organise smallholders, operate wet "
               "mills (factory co-operative committees), and market parchment "
               "on behalf of farmers."],
        ["4", "Wet Mills (FFCs)", "Process ripe cherry into parchment through "
               "pulping, fermentation and drying at factory level."],
        ["5", "Dry Millers", "Process parchment and mbuni into clean, graded "
               "green coffee for the export market."],
        ["6", "Marketers / Brokers", "Prepare, grade, value and market clean "
               "coffee at the auction and for direct sales, operating under "
               "licence."],
        ["7", "Nairobi Coffee Exchange", "The central auction for Kenyan "
               "coffee; runs the electronic trading platform and publishes "
               "market prices."],
        ["8", "Buyers / Exporters", "Purchase coffee at auction or through "
               "direct trade and export it to global roasters and markets."],
        ["9", "Coffee Board of Kenya / AFA", "Regulatory bodies overseeing "
               "licensing, marketing and quality (under the new Coffee Act)."],
        ["10", "Research Institutions", "The Coffee Research Institute (CRI) "
                "develops improved varieties, disease control and agronomy."],
        ["11", "Transporters", "Move cherry, parchment and green coffee "
                "through the chain from farm to port."],
        ["12", "Packaging & Warehouse Operators", "Provide quality packages "
                "and designated warehousing for auction and export lots."],
        ["13", "Financial Institutions", "Provide crop financing and operate "
                "the Direct Settlement System that pays farmers."],
        ["14", "Input Suppliers", "Supply certified seedlings, fertilizers "
                "and crop-protection inputs."],
        ["15", "International Buyers / Roasters", "Purchase Kenyan coffee for "
                "blending, roasting and specialty retail worldwide."],
    ])
    b.doc.add_heading("Capacity", level=2)
    b.add_para(
        "Kenya produces an average of 800,000\u2013850,000 60-kg bags of clean "
        "coffee per year in the 2023/24 and 2025/26 marketing years, with "
        "production forecast to climb to about 950,000 bags by 2026/27 on the "
        "back of higher prices and the government's expansion programme "
        "(USDA FAS, Kenya Coffee Annual 2025 and 2026). Area planted is "
        "estimated at 105,000\u2013110,000 hectares, bearing roughly 170 "
        "million trees. About 90% of production is processed through "
        "co-operative and estate wet mills; the remainder is dry-processed.")

    # ===================================================================== 4.0
    b.doc.add_heading("4.0\tSOURCES OF COFFEE PRODUCTION AREAS IN KENYA",
                      level=1)
    b.add_para(
        "Coffee is grown in two main geographical blocks: the East of the "
        "Rift and the West of the Rift.")
    b.add_para(
        "The East of the Rift block is predominantly mountainous, comprising "
        "the Mt Kenya massif, the Aberdares ranges and the Nyambene hills. The "
        "principal producing counties are Nyeri, Kirinyaga, Embu, Tharaka-"
        "Nithi, Meru, Murang'a and Kiambu. This bloc has historically "
        "accounted for about 60 percent of national coffee auction volumes.")
    b.add_para(
        "The West of the Rift block is characterised by plains and undulating "
        "hills and includes Kericho, Nandi, Bomet, Narok, Bungoma, Kakamega, "
        "Kisii, Nyamira and Trans-Nzoia. Western Kenya is the fastest-growing "
        "production region: in the 2025/26 season, Nandi auction volumes rose "
        "92 percent and Bungoma 97 percent, and the combined western bloc "
        "(with Kisii and Nyamira) lifted its share of national auction volume "
        "from 19 to 34 percent. Kericho emerged as Kenya's leading coffEE-"
        "producing county by auction volume in 2025/26 (Nairobi Coffee "
        "Exchange report via Kenya News Agency, 2026).")
    b.doc.add_heading("Coffees produced in Kenya", level=2)
    b.add_para(
        "Arabica coffee accounts for more than 99 percent of Kenya's output; "
        "robusta production is minimal. Kenya's arabicas are classified by "
        "bean size and preparation, with the principal grades being E "
        "(elephant beans), PB (peaberry), AA, AB and C. The high-grown "
        "washed arabicas from Nyeri, Kirinyaga and Kiambu (the 'Kenya AA' "
        "profile) are among the most sought-after coffees in the specialty "
        "market.")
    b.add_para(
        "Specialty categories \u2013 single-estate lots, micro-lots, high-"
        "scoring auction lots, organic and UTZ/Rainforest Alliance-certified "
        "coffees, and roasted-and-ground retail packs \u2013 are a fast-growing "
        "segment. Premium producing counties command visibly higher prices: "
        "in 2025/26, Kirinyaga coffee averaged about USD 7.25/kg against a "
        "national average of USD 6.81/kg at the NCE (Kenya News Agency, 2026).")

    # ===================================================================== 5.0
    b.doc.add_heading("5.0\tCOFFEE SECTOR DRIVERS", level=1)
    b.add_para("Coffee sector institutions")
    _cap(b, "Coffee Sector Institutions")
    _static_table(b, [
        ["No.", "INSTITUTION / BODY", "RESPONSIBILITY"],
        ["1.", "Ministry of Agriculture and Livestock Development", "Policy "
               "development, strategy (Coffee Development and Marketing "
               "Strategy 2024\u20132029) and sector coordination."],
        ["2.", "Coffee Board of Kenya", "Re-established regulator overseeing "
               "licensing, marketing, quality and promotion of coffee under "
               "the new Coffee Act."],
        ["3.", "Agriculture and Food Authority (AFA)", "Houses the Coffee "
               "Directorate, which regulates production and licensing."],
        ["4.", "Coffee Research Institute (CRI)", "Breeding, agronomy, "
               "disease research and supply of certified planting material."],
        ["5.", "New Kenya Planters Co-operative Union (NKPCU)", "Implements "
               "the government coffee expansion revolving fund: saplings, "
               "fertilizers and extension to farmers."],
        ["6.", "Kenya Coffee Platform (KCP)", "A public-private forum "
               "coordinating sector reforms and value-chain stakeholders."],
        ["7.", "Nairobi Coffee Exchange (NCE)", "The auction and electronic "
               "trading platform where most Kenyan coffee is priced and "
               "sold."],
        ["8.", "Co-operative Societies & Farmers", "Cultivation, wet "
               "processing and primary marketing of cherry and parchment."],
        ["9.", "County Governments", "Devolved extension services, and "
               "production and marketing support under Schedule IV of the "
               "Constitution."],
        ["10.", "Financial Institutions", "Crop financing and operation of "
                "the Direct Settlement System (DSS) for farmer payments."],
    ])

    # ===================================================================== 6.0
    b.doc.add_heading("6.0\tREGULATION OF THE SECTOR", level=1)
    b.add_para(
        "The coffee sub-sector is governed by the Crops Act, 2013 (which "
        "vests licensing and regulation in the Agriculture and Food "
        "Authority), supported by the Coffee (General) Regulations. Kenya has "
        "recently enacted a new Coffee Act (2026) that re-establishes the "
        "Coffee Board of Kenya as the dedicated regulator, shifting "
        "oversight of licensing, marketing, quality and export promotion of "
        "coffee to the Board (USDA FAS, Kenya Coffee Annual 2026).")
    b.add_para(
        "In July 2023, the Government introduced marketing reforms that "
        "abolished marketing agents, prohibited any single business from "
        "holding multiple licences across the value chain, and established the "
        "Direct Settlement System (DSS) \u2013 operated through the Co-operative "
        "Bank of Kenya \u2013 so that proceeds of coffee sales are paid "
        "directly to farmers' bank accounts rather than through millers or "
        "marketers (USDA FAS, Kenya Coffee Annual 2024).")
    b.add_para(
        "The Nairobi Coffee Exchange operates an electronic trading platform "
        "for the auction of coffee, and most Kenya coffee (other than "
        "orthodox, specialty and directly-contracted lots) is offered at the "
        "auction floor. Sales at the NCE are transacted in United States "
        "dollars, giving farmers exposure to hard-currency prices.")
    b.add_para(
        "Complementary instruments include the Coffee Development and "
        "Marketing Strategy 2024\u20132029, which sets out eight strategic "
        "pillars \u2013 productivity, quality enhancement, value addition, "
        "market expansion, farmer empowerment, sustainability, coordination "
        "and financing \u2013 and the requirement that coffee is traceable from "
        "farm to cup, a growing prerequisite for EU and specialty buyers.")

    # ===================================================================== 7.0
    b.doc.add_heading("7.0\tPRODUCTION CAPACITY", level=1)
    b.add_para(
        "Coffee is a major foreign-exchange earner for Kenya. In %d, Kenya's "
        "coffee exports (HS 0901) were valued at about USD %.0f million, with "
        "green (unroasted, not decaffeinated) coffee \u2013 HS 090111 \u2013 "
        "dominating the value chain (see Tables 1\u20132; ITC Trade Map)." %
        (rev, display(kenya_2025)))
    b.doc.add_heading("Production capacity", level=2)
    b.add_para(
        "Kenya produces an estimated 800,000 \u2013 850,000 60-kg bags a year "
        "(USDA FAS, Coffee Annuals), with output projected to reach roughly "
        "950,000 bags in 2026/27. Area under coffee is about 105,000\u2013"
        "110,000 hectares and the total tree population is estimated at 170 "
        "million bearing trees. Average productivity of about 475 kg/ha is "
        "well below the 970 kg/ha achieved at independence, indicating "
        "substantial headroom for yield gains (UNIDO/FAO 2025).")
    b.doc.add_heading("Manufacturing infrastructure", level=2)
    b.add_para(
        "The manufacturing structure of Kenyan coffee involves several steps, "
        "from harvesting to export.")
    _cap(b, "Coffee Processing Flow")
    _static_table(b, [
        ["No.", "Coffee Processing Flow", "Function"],
        ["1", "Harvesting", "Ripe cherry is hand-picked; estates are "
               "introducing selective mechanical harvesting."],
        ["2", "Transportation", "Cherry is delivered to the wet mill (factory "
               "co-operative committee) within hours of picking."],
        ["3", "Pulping", "Outer skin and pulp are removed, typically using "
               "disc pulpers; mbuni (dry-process) skips pulping."],
        ["4", "Fermentation", "Mucilage is broken down over 12\u201336 hours "
               "to develop the clean cup profile."],
        ["5", "Washing & Soaking", "Parchment coffee is washed and soaked in "
               "clean channels or tanks."],
        ["6", "Drying", "Parchment is sun-dried on raised beds or in driers "
               "to about 10\u201312% moisture."],
        ["7", "Dry Milling", "Parchment is hulled to green coffee; mbuni is "
               "hulled and polished."],
        ["8", "Grading & Sorting", "Green coffee is size-graded (E, PB, AA, "
               "AB, C, others), density-triaged and hand-/machine-sorted."],
        ["9", "Quality Control", "Samples are liquored (cupped) and graded "
               "before auction or direct sale."],
        ["10", "Packaging & Warehousing", "Clean coffee is packed into 60-kg "
                "export bags and stored in designated warehouses before "
                "delivery."],
    ])
    b.doc.add_heading("Technological advancement", level=2)
    b.add_para(
        "Mechanisation is gradually spreading. Selective mechanical "
        "harvesting, mostly on large estates, raises efficiency and cuts "
        "labour costs, while precision grading uses density and "
        "colour-sorting machines. On the research side, the Coffee Research "
        "Institute is expanding certified sapling production to support the "
        "government's replanting and expansion programme, which is "
        "implemented through the New Kenya Planters Co-operative Union "
        "revolving fund (USDA FAS 2025 and 2026). Digital integration of the "
        "Direct Settlement System and the NCE electronic platform is "
        "modernising payment and price discovery.")

    # ===================================================================== 8.0
    b.doc.add_heading("8.0\tEXPORT PERFORMANCE & CATEGORIZATION", level=1)
    b.doc.add_heading("Export volume & value", level=2)
    b.add_para(
        "According to the International Trade Centre Trade Map (%d), Kenya's "
        "coffee exports (HS 0901) were valued at USD %.0f million in %d, "
        "reflecting firm demand and elevated global arabica prices. Over the "
        "review period 2016\u2013%d, the value and share of Kenya's coffee by "
        "product heading is summarised in Table 1." % (rev, display(kenya_2025), rev, rev))

    # --- Table 1: Kenya exports by product (HS6 under 0901) ----------------
    _cap(b, "Trend on Coffee: Kenya's Exports by Product")
    b.add_value_table(
        "Product", sorted(data.members, key=lambda r: r["code"]), years,
        "Share in %d" % rev, "Kenya's Coffee Exports by Product",
        CFG.get("source"), total_label="Total")
    b.add_para(
        "Green, unroasted, non-decaffeinated coffee (HS 090111) dominates "
        "Kenya's exports, accounting for roughly %.0f%% of export value in %d, "
        "with smaller shares for roasted coffee (HS 090121) and coffee husks "
        "and skins (HS 090190) \u2013 confirming that most value is still "
        "captured in raw green-bean form and that roasting and grinding for "
        "export is a priority opportunity." % (
            (green["years"].get(rev) or 0.0) / kenya_2025 * 100, rev))

    # ===================================================================== 9.0
    b.doc.add_heading("9.0\tLATEST KENYA'S EXPORT VALUE & MARKETS", level=1)
    b.add_para("Destination markets for Kenya's Coffee")
    _cap(b, "Kenya's Coffee Export Markets by Destination")
    b.add_value_table(
        "Destination market",
        _ranked_rows(dest, 10, years, residual="All other markets"),
        years, "Share in %d" % rev, "Destinations for Kenya's Coffee",
        CFG.get("source"), total_label="Total")
    if dest_top:
        lead, second, third = dest_top[0], dest_top[1], dest_top[2]
        b.add_para(
            "In %d, Kenya's coffee exports were valued at USD %.0f million "
            "(ITC Trade Map). %s was the leading destination with USD %.0f "
            "million (%.1f%%), followed by %s (USD %.0f million) and %s "
            "(USD %.0f million). The concentration of exports into a small "
            "number of roasting markets underscores both Kenya's strong "
            "position in premium European and Asian roasteries and the need "
            "to diversify into emerging specialty markets." % (
                rev, dest_total, lead["label"], lead["years"].get(rev) or 0,
                (lead["years"].get(rev) or 0) / dest_total * 100,
                second["label"], second["years"].get(rev) or 0,
                third["label"], third["years"].get(rev) or 0))
    b.doc.add_page_break()

    # ==================================================================== 10.0
    b.doc.add_heading("10.0\tLEADING COMPETITORS", level=1)
    b.add_para(
        "The table below summarises the competitive position of Kenya's "
        "coffee exports in the world market and among its leading African "
        "peers.")
    _cap(b, "World Exports of Coffee by Economy")
    b.add_value_table(
        "Exporting economy",
        _ranked_rows(data.exporters(), 10, years, residual="All other economies"),
        years, "Share in %d" % rev, "Countries Exporting Coffee",
        CFG.get("source"), total_label="Total", rank=True)
    b.add_para(
        "Brazil, Vietnam, Colombia, Indonesia and Ethiopia are the dominant "
        "global producers, with Brazil alone producing about 37 percent of "
        "the world crop in 2024/25 (about 65 million 60-kg bags) and Vietnam "
        "a further 29 million bags (USDA FAS, World Coffee Production). World "
        "coffee production reached a record 175 million bags in 2024/25, and "
        "the ICO composite indicator price averaged about 348 US cents/lb in "
        "March 2025 (International Coffee Organization, Coffee Market Report, "
        "March 2025). Kenya, by contrast, accounts for under one percent of "
        "world exports by volume but commands the world's most consistent "
        "premiums for high-grown washed arabica.")
    peers = data.african_peers(5)
    if peers:
        _cap(b, "Kenya vs Leading African Exporters of %s" % (
            _anchor(b, data)))
        b.add_value_table("Exporting economy", peers, years,
                          "Share in %d" % rev, "African Coffee Exporters",
                          CFG.get("source"), total_label="Total")
        older = [p for p in peers if p.get("src_year") and p["src_year"] != rev]
        if older:
            b.add_para(
                "Note: %s did not report a %d figure in the source download; "
                "the value shown is for their most recent available year."
                % (" and ".join(p["label"] for p in older), rev))
        b.add_para(
            "Among African producers, Kenya competes directly with Ethiopia "
            "and Uganda in the premium mild-arabica segment. Uganda has grown "
            "strongly on increased production and exports, while Ethiopia \u2013 "
            "the continent's largest coffee producer \u2013 competes on volume. "
            "Kenya's differentiation rests on quality, traceability and "
            "specialty premiums rather than scale.")

    # ==================================================================== 11.0
    b.doc.add_heading("11.0\tGOVERNMENT SUPPORT AND POLICIES", level=1)
    b.doc.add_heading("Trade agreements", level=2)
    b.add_para(
        "Trade agreements can be complex and may change over time. An "
        "overview of the trade arrangements that typically involve Kenyan "
        "coffee:")
    b.add_para("11.1.1\u2002African Continental Free Trade Area (AfCFTA): "
               "creates a single continental market for goods and services, "
               "opening up African markets for Kenyan roasted and instant "
               "coffee.")
    b.add_para("11.1.2\u2002East African Community (EAC): provides "
               "preferential trade among EAC members and supports intra-regional "
               "coffee trade.")
    b.add_para("11.1.3\u2002COMESA: allows reduced tariffs on coffee within "
               "member states.")
    b.add_para("11.1.4\u2002AGOA: provides duty-free access to the U.S. "
               "market for certain products, including roasted coffee.")
    b.add_para("11.1.5\u2002European Union EPA (and UK\u2013Kenya EPA): offers "
               "preferential quota-free market access for Kenyan coffee to the "
               "EU and the UK \u2013 the destination for over half of Kenya's "
               "coffee.")
    b.add_para("11.1.6\u2002Bilateral agreements: Kenya maintains bilateral "
               "arrangements with key buyers such as the United States, United "
               "Kingdom, and several Gulf and Asian markets that support "
               "coffee trade.")
    b.doc.add_heading("Incentives for exporters", level=2)
    b.add_para(
        "These incentives aim to enhance the competitiveness of Kenyan coffee "
        "and support the growth of the export sector. Their availability "
        "varies over time with government policy.")
    _cap(b, "Coffee Export Incentives")
    _static_table(b, [
        ["No", "Item", "Incentive"],
        ["1", "Export Processing Zones", "Tax holidays and reduced duties for "
               "value-adding coffee processors located in EPZs."],
        ["2", "Export Promotion Programmes", "Marketing support and trade-fair "
               "participation through KEPROBA and the Coffee Board."],
        ["3", "Duty Drawback Scheme", "Refund of import duties on inputs used "
               "in producing exported coffee."],
        ["4", "VAT Refund System", "Refund of VAT paid on goods and services "
               "used in export production."],
        ["5", "Export Credit Guarantee", "Government-backed insurance to de-"
               "risk export financing for millers and exporters."],
        ["6", "Infrastructure Development", "Investment in rural roads and "
               "market infrastructure for coffee-growing counties."],
        ["7", "Trade Agreements", "Preferential access to key markets under "
               "EPAs, AGOA and AfCFTA."],
        ["8", "Research Support", "Funding for coffee research through the "
               "Coffee Research Institute."],
        ["9", "Quality Certification Assistance", "Support for farms and "
               "factories to obtain international sustainability and "
               "traceability certification."],
    ])
    b.doc.add_heading("Industry development programmes", level=2)
    b.add_para(
        "Coffee industry development programmes in Kenya are initiatives "
        "aimed at improving and sustaining the coffee sector.")
    _cap(b, "Coffee Development Programmes")
    _static_table(b, [
        ["No", "ITEM", "Programme"],
        ["1", "Coffee Expansion Programme", "Government revolving fund "
               "through NKPCU supplying improved saplings and fertilizers to "
               "expand area under coffee in Central, Eastern and Rift Valley "
               "regions."],
        ["2", "Research and Innovation", "CRI expansion of certified planting "
               "material and breeding of resilient, high-yielding varieties."],
        ["3", "Productivity Recovery", "Replanting of old/senescent trees and "
               "improved agronomy to lift yields toward historic averages."],
        ["4", "Market Reforms", "Direct Settlement System, removal of "
               "marketing agents, and modernised NCE electronic auction."],
        ["5", "Market Development", "Diversification into specialty, "
               "single-origin and direct-trade markets."],
        ["6", "Quality Improvement", "Grading, liquoring and cupping standards "
               "to protect the Kenya brand."],
        ["7", "EUDR Readiness", "Farm-to-cup traceability systems and "
               "geolocation mapping to comply with the EU Deforestation "
               "Regulation."],
        ["8", "Technology Adoption", "Mechanisation, improved drying, and "
               "digital payment and traceability platforms."],
        ["9", "Sustainable Practices", "Climate-smart agronomy, water "
               "conservation and certification promotion."],
        ["10", "Capacity Building", "Training for farmers, co-operatives and "
                "millers on quality and financial management."],
        ["11", "Financial Support", "Access to affordable credit for "
                "smallholders, co-operatives and value-adding processors."],
    ])

    # ==================================================================== 12.0
    b.doc.add_heading("12.0\tCHALLENGES FACING THE SECTOR", level=1)
    b.add_para(
        "Despite Kenya's competitive advantage as a premium coffee origin, "
        "several challenges threaten sustained growth. The following "
        "challenges have been identified:")
    for item in (
        "Low and declining productivity: average yields of about 475 kg/ha "
        "against 970 kg/ha in 1963, partly due to ageing trees and limited "
        "replanting of improved varieties (UNIDO/FAO 2025).",
        "Climate variability: dependence on rain-fed farming makes production "
        "vulnerable to erratic rainfall, drought and rising temperatures, "
        "intensifying disease (coffee berry disease and leaf rust) pressure.",
        "High cost and shortage of inputs: fertilizer prices have risen "
        "sharply, and the Coffee Research Institute remains the sole supplier "
        "of certified planting material, constraining expansion (USDA FAS).",
        "Access to finance: smallholders frequently lack capital and "
        "collateral, limiting investment in inputs, wet-mill upgrades and "
        "value addition (UNIDO/FAO 2025).",
        "Payment delays: the transition to the Direct Settlement System has "
        "been accompanied by inordinate delays in farmer payments as the "
        "platform matures (USDA FAS 2024).",
        "Regulatory churn: seven reforms in seven years have at times "
        "disrupted licensing, milling and marketing, with several millers "
        "ceasing operations under the 2023 licensing requirements (USDA FAS; "
        "UNIDO 2025).",
        "Global and EU compliance: over half of exports go to Europe, and "
        "compliance with the EU Deforestation Regulation and sustainability "
        "schemes adds real cost for smallholders.",
        "Value-capture gap: over 97% of export value is sold as raw green "
        "coffee; roasting, grinding and brand value largely accrue offshore.",
        "Youth and labour: a shrinking, ageing farmer base and competition "
        "from real-estate conversion of peri-urban coffee land.",
        "Low domestic consumption: Kenyan per-capita coffee consumption "
        "remains far below major producing countries, limiting the domestic "
        "market as an outlet.",
    ):
        b.add_bullet(item)
    b.doc.add_page_break()

    # ==================================================================== 13.0
    b.doc.add_heading("13.0\tPROPOSED INTERVENTIONS", level=1)
    b.add_para(
        "For Kenya to maintain and grow its position in the premium coffee "
        "market, strategic interventions are needed. A tabulated overview of "
        "key initiatives follows:")
    _cap(b, "Sector Interventions")
    _static_table(b, [
        ["No", "ITEM", "INITIATIVE"],
        ["A", "Environmental Sustainability", ""],
        ["1", "Climate Resilience", "Water harvesting at wet mills, "
               "agroforestry, drought- and disease-tolerant varieties, and "
               "regenerative soil management."],
        ["2", "Conservation", "Protection of riparian zones and biodiversity "
               "on coffee farms and estates."],
        ["B", "Social Sustainability", ""],
        ["1", "Farmer Payments", "Fully operationalise the Direct Settlement "
               "System to guarantee timely, transparent farmer payments."],
        ["2", "Empowerment", "Strengthen co-operatives and extend input "
               "credit, extension and training, including to women and "
               "youth."],
        ["C", "Economic Sustainability", ""],
        ["1", "Productivity", "Replant old trees, expand certified-sapling "
               "supply and deliver agronomy packages to lift yields toward "
               "historical averages."],
        ["2", "Value Addition", "Incentivise local roasting, grinding and "
               "packaging; support EPZ-based roasters to capture retail "
               "margins."],
        ["3", "Market Diversification", "Deepen specialty and direct-trade "
               "markets and expand domestic consumption."],
        ["D", "Compliance & Traceability", "Farm-to-cup digital traceability "
               "and EUDR readiness (geolocation, deforestation-free mapping) "
               "to protect market access."],
        ["E", "Finance", "Expand affordable crop financing, guarantee schemes "
               "and co-operative credit for productivity and value-addition "
               "investment."],
        ["F", "Policy and Governance", "Full implementation of the Coffee Act "
               "and the Coffee Development and Marketing Strategy 2024\u2013"
               "2029, with coordinated county-level support."],
    ])
    b.add_para(
        "For Kenya to secure its future as a leading global coffee origin, "
        "strategic interventions are needed, including:")
    b.add_bullet("Investing in climate resilience and sustainable production "
                 "practices for coffee cultivation.")
    b.add_bullet("Raising productivity through replanting, better agronomy "
                 "and access to improved, disease-resistant varieties.")
    b.add_bullet("Enhancing value addition in the coffee supply chain, "
                 "especially the growing segments of roasted, ground and "
                 "ready-to-drink coffee for premium export and domestic "
                 "markets.")
    b.add_bullet("Strengthening traceability and EUDR compliance to secure "
                 "and grow the European market share.")
    b.add_bullet("Expanding market diversification into specialty, "
                 "single-origin and direct-trade segments, alongside "
                 "domestic consumption development.")

    # ---------------------------------------------------------------- save
    os.makedirs(os.path.dirname(OUT_DOC), exist_ok=True)
    doc.save(OUT_DOC)
    print("Report saved to: %s" % OUT_DOC)
    try:
        write_excel_deliverable(CFG, data, OUT_XLS)
        print("Tables saved to: %s" % OUT_XLS)
    except Exception as e:  # pragma: no cover
        print("Excel deliverable skipped: %s" % e)

    print("Kenya coffee exports %d: USD %.1f million" % (rev, display(kenya_2025)))
    print("Top destinations %d: %s" % (rev, ", ".join(
        "%s (%.1f)" % (r["label"], r["years"].get(rev) or 0)
        for r in dest_top)))


def _anchor(b, data):
    label = data.anchor_label or CFG.get("family_title", "")
    return label.split(",")[0]


if __name__ == "__main__":
    main()