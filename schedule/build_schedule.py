"""Neumann Residence (1567 Steveles Cres) - construction schedule builder.

Builds a linked, calendar-driven schedule with MPXJ, schedules it with
MPXJ's Microsoft Project-compatible CPM scheduler, and writes Microsoft
Project XML (MSPDI). Open the .xml in MS Project and Save As .mpp.

Sources (SharePoint: Neumann Admin): Steveles proj stat Oct 8.xlsx,
Steveles Internal Meeting notes Oct 8, Neumann RFI Log.xlsx (28-Aug
meeting), Budget REV3 with AA, CVC Permit 26/255 summary, bid results
(Demolition, Foundation & Underpinning Rev), client emails to Oct 8.

Usage: python3 build_schedule.py <output.xml>
"""
import sys
from datetime import date, datetime

import jpype
import jpype.imports
import mpxj  # noqa: F401  (puts MPXJ jars on the classpath)

jpype.startJVM()

from java.time import LocalDate, LocalDateTime  # noqa: E402
from org.mpxj import (  # noqa: E402
    ConstraintType, Duration, ProjectFile, Relation, RelationType,
    TaskField, TimeUnit,
)
from org.mpxj.cpm import MicrosoftScheduler  # noqa: E402
from org.mpxj.mspdi import MSPDIWriter  # noqa: E402

PROJECT_START = datetime(2026, 10, 1, 8, 0)
STATUS_DATE = datetime(2026, 10, 13, 8, 0)

# Ontario statutory holidays plus the Christmas shutdown.
HOLIDAYS = [
    (date(2026, 12, 24), date(2027, 1, 1), "Christmas shutdown"),
    (date(2027, 2, 15), date(2027, 2, 15), "Family Day"),
    (date(2027, 3, 26), date(2027, 3, 26), "Good Friday"),
    (date(2027, 5, 24), date(2027, 5, 24), "Victoria Day"),
    (date(2027, 7, 1), date(2027, 7, 1), "Canada Day"),
    (date(2027, 8, 2), date(2027, 8, 2), "Civic Holiday"),
    (date(2027, 9, 6), date(2027, 9, 6), "Labour Day"),
]

# Row format: (level, key, name, duration_days, predecessors, responsible, notes, options)
#   level 1 = phase summary (duration/preds ignored), level 2 = task.
#   duration 0 = milestone.
#   predecessors: "key", "key+5" (FS lag, working days), "key:SS", "key:SS+3", "key:FF".
#   options: snet=date (start no earlier than), deadline=date, done=True (100% complete).
S = None
ROWS = [
    # ------------------------------------------------------------------
    (1, "PH_PRE", "Preconstruction, Design & Approvals", S, "", "", "", {}),
    (2, "cvc", "CVC Permit 26/255 issued (expires Oct 2, 2028)", 0, "", "Darren (Structured Creations)",
     "Credit Valley Conservation, O.Reg 41/24. Conditions: 48-hr notice before start and on completion; "
     "ESC installed before construction; restore disturbed areas. Any change in regulated area needs CVC amendment.",
     {"snet": date(2026, 10, 2), "done": True}),
    (2, "zoning", "Zoning review - City of Mississauga", 8, "", "Darren (Structured Creations)",
     "Outstanding as of Oct 8; holding up final architectural drawings.", {"snet": date(2026, 10, 13)}),
    (2, "a100", "Correct A100 coverage table and CVC permit owner name", 3, "", "Darren / Gosia",
     "A100 'Total Coverage' shows 70.80 m2, should be 290.80 m2. Permit lists Stephen & Robert Neumann; file has Stephen & Rachael.",
     {"snet": date(2026, 10, 13)}),
    (2, "wd", "Working drawings - architectural & structural (incl. basement program)", 13, "", "Darren (Structured Creations)",
     "Target end of October for permit submission. Basement mechanical/layout questions (RFI-002) to feed in.",
     {"snet": date(2026, 10, 13)}),
    (2, "permit_sub", "Building permit application submitted", 0, "wd, zoning, a100", "Darren / Andrew", "", {}),
    (2, "permit_rev", "City permit review incl. one comment cycle (assumed 25 working days)", 25, "permit_sub",
     "City of Mississauga",
     "Assumption - adjust this duration once the City confirms review timing. Drives all structural/exterior work.", {}),
    (2, "permit", "Building permit issued", 0, "permit_rev", "City of Mississauga",
     "Status report target was Oct 30; meeting notes now show submission end of October.", {}),
    (2, "budget4", "Budget Revision 4 - anchor budget to current scope (incl. loft / 3rd floor)", 6, "",
     "Dan / Gosia", "Rev 3 omits loft/3rd floor. Demo bid already shows $9,982.94 scope change (asbestos + loft). Carry 10-15% contingency.",
     {"snet": date(2026, 10, 13)}),
    (2, "id_dwgs", "Interior design drawings & ID budget (Vicky)", 15, "", "Vicky (ID) / Gosia", "",
     {"snet": date(2026, 10, 13)}),
    (2, "nz_price", "Net Zero palette - incremental pricing by item", 10, "", "Gosia / Building Knowledge",
     "Stephen to decide item by item. Wood-burning fireplace flagged as a potential Net Zero deal breaker.",
     {"snet": date(2026, 10, 13)}),
    (2, "nz_dec", "Client decision - Net Zero direction, fireplace, window spec", 5, "nz_price", "Stephen (Owner)",
     "Drives window SHGC/glazing, heat pump vs. existing furnaces, insulation levels, solar.", {}),
    (2, "slab_price", "Price basement slab insulation options A / B incl. height loss", 10, "", "Gosia",
     "A: remove slab, insulate under, re-pour (preferred if new plumbing). B: insulate over slab, double-board.",
     {"snet": date(2026, 10, 13)}),
    (2, "slab_dec", "Client decision - basement slab option", 5, "slab_price", "Stephen (Owner)", "", {}),
    (2, "ug_price", "Underground hydro service - pricing from City / Alectra", 15, "", "Gosia",
     "Needed before the 2nd-floor bump-out wall comes down.", {"snet": date(2026, 10, 13)}),
    (2, "ug_dec", "Client decision - underground hydro service", 5, "ug_price", "Stephen (Owner)", "", {}),
    (2, "gen_dec", "Generator sizing, pricing & client decision", 15, "", "Dan / Stephen",
     "RFI-008. Overhead line through ravine - tree strike risk. Not in budget.", {"snet": date(2026, 10, 19)}),
    (2, "insurance", "Confirm builder's risk / wrap-up insurance & issue PO", 2, "", "David",
     "Overdue (due Oct 9). Required before site possession.", {"snet": date(2026, 10, 13)}),

    # ------------------------------------------------------------------
    (1, "PH_PROC", "Tendering, Trade Awards & Long-Lead Procurement", S, "", "", "", {}),
    (2, "aw_demo", "Award demolition (recommend Manpower - $49,025 incl. asbestos & loft)", 2, "", "Andrew / Dan",
     "Budget $30,000. Scope change $9,982.94 required (asbestos $4,775 + loft $5,207.94).", {"snet": date(2026, 10, 13)}),
    (2, "aw_fence", "PO site hoarding, fence & silt control", 2, "", "David", "", {"snet": date(2026, 10, 13)}),
    (2, "aw_hydro_book", "Book Alectra overhead disconnect & temp pole", 15, "", "Gosia",
     "Overhead line runs through 2nd-floor bump-out (bathroom area).", {"snet": date(2026, 10, 13)}),
    (2, "aw_exc", "Excavation quotes - Kruno / Grizzly; all-in-one vs. separate", 8, "", "Gosia",
     "Excavation + slab est. $50-60K. All-in-one may save ~$10K and reduce coordination.", {"snet": date(2026, 10, 13)}),
    (2, "aw_fdn", "Award excavation, foundation & underpinning (recommend EECF)", 5, "aw_exc", "Dan / Gosia",
     "EECF $44,310 levelled incl. winter heat. Confirm in writing underpinning is included (p.3 boilerplate excludes it).", {}),
    (2, "aw_frame", "Tender & award framing labour (RamPeel, King, Murray, Sunvalley, VAV)", 15, "", "Gosia",
     "RamPeel $62,650 vs. $62,500 labour budget.", {"snet": date(2026, 10, 13)}),
    (2, "aw_lumber", "PO lumber & roof framing; delivery", 10, "aw_frame, wd", "Gosia", "", {}),
    (2, "aw_steel", "Steel beams - confirm scope, shop drawings, fabricate", 20, "wd", "Dan",
     "Not carried in budget (carport steel removed). Confirm with structural.", {}),
    (2, "win_meas", "Measure brick openings; issue window / mill schedule", 3, "possession", "Andrew / Stephen",
     "No dimensions on drawings. Suppliers to measure Day 1 or 2.", {}),
    (2, "win_tender", "Window tender - double vs. triple pane, heritage profile", 10, "win_meas", "Andrew",
     "Budget $98K (49 units). Wellington Windows, Mullins. Price all-windows and outlined-only options.", {}),
    (2, "win_po", "Window, exterior door & skylight PO", 3, "win_tender, nz_dec", "Andrew / Stephen", "", {}),
    (2, "win_fab", "Window, door & skylight fabrication & delivery (~10 wks)", 50, "win_po", "Supplier",
     "Lead time assumption - confirm with awarded supplier.", {}),
    (2, "hvac_des", "HVAC design / heat loss - heat pump, HRV/ERV per Net Zero direction", 15, "hvac_walk, nz_dec",
     "Dan / Andrew / Gosia", "Ask Darren who he uses; McCallum preferred.", {}),
    (2, "aw_hvac", "Award HVAC / mechanical", 5, "hvac_des", "Gosia", "", {}),
    (2, "aw_plumb", "Award plumbing", 10, "", "Gosia", "", {"snet": date(2026, 12, 1)}),
    (2, "aw_elec", "Award electrical (confirm EV charging, solar & AV scope)", 10, "", "Gosia",
     "Common scope gaps: EV charging, solar, AV/automation - none carried in Rev 3.", {"snet": date(2026, 12, 1)}),
    (2, "aw_drywall", "Award insulation & drywall", 10, "", "Gosia", "", {"snet": date(2027, 1, 4)}),
    (2, "aw_finish", "Award finishing trades - tile, flooring, trim, paint", 15, "", "Gosia", "",
     {"snet": date(2027, 3, 1)}),
    (2, "cab_design", "Kitchen & cabinetry design approval", 20, "id_dwgs", "Vicky (ID) / Stephen",
     "Budget $176,400.", {"snet": date(2027, 1, 4)}),
    (2, "cab_shops", "Cabinetry shop drawings & sign-off", 15, "cab_design", "Cabinetmaker / Vicky", "", {}),
    (2, "cab_fab", "Cabinetry fabrication", 45, "cab_shops", "Cabinetmaker", "", {}),
    (2, "stair_shops", "Interior stair & railing shop drawings & fabrication", 30, "fr_review", "Stair supplier", "", {}),

    # ------------------------------------------------------------------
    (1, "PH_SEL", "Owner & Interior Design Selections", S, "", "", "", {}),
    (2, "sel_elec", "Electrical, lighting & AV / automation layouts", 10, "id_dwgs", "Vicky / Stephen", "", {}),
    (2, "sel_plumb", "Plumbing fixtures & bathroom layouts (curbless showers ensuite 1 & 2)", 10, "id_dwgs",
     "Vicky / Stephen", "RFI-015: add pot filler at dog station.", {}),
    (2, "sel_fp", "Fireplace / logsets & mantels", 5, "nz_dec", "Stephen", "Existing fireplace to remain; reclad with salvaged brick if possible.", {}),
    (2, "sel_ext", "Exterior - cedar shake, metal panel, colours, railings, garage door", 10, "", "Stephen / Darren", "",
     {"snet": date(2027, 1, 11)}),
    (2, "sel_floor", "Hardwood, floor tile & wall tile", 10, "", "Vicky / Stephen", "", {"snet": date(2027, 2, 1)}),
    (2, "sel_trim", "Trim, interior doors, plaster mouldings, hardware", 10, "", "Vicky / Stephen", "",
     {"snet": date(2027, 2, 15)}),
    (2, "sel_appl", "Appliances", 5, "cab_design", "Stephen", "Budget $50,000.", {}),
    (2, "sel_counter", "Countertop slabs, shower glass & bath hardware", 10, "", "Vicky / Stephen", "",
     {"snet": date(2027, 3, 15)}),
    (2, "sel_paint", "Paint finish schedule", 5, "", "Vicky / Stephen", "", {"snet": date(2027, 4, 1)}),

    # ------------------------------------------------------------------
    (1, "PH_MOB", "Site Mobilization & Pre-Demolition", S, "", "", "", {}),
    (2, "possession", "Site possession - family moved out (confirm Oct 15)", 0, "insurance", "Stephen (Owner)",
     "Confirm whether Oct 15 is move-out day or first day on site. Stephen's Oct 8 email references Nov 15 - clarify.",
     {"snet": date(2026, 10, 15)}),
    (2, "bk_audit", "Building Knowledge as-is audit & baseline air test", 1, "possession", "Andrew / Building Knowledge",
     "Must precede demo to establish baseline.", {}),
    (2, "hst_doc", "Pre-demolition condition photos (HST rebate documentation)", 1, "possession", "Andrew / Stephen",
     "RFI-018. Substantial reconstruction rebate est. $50-60K.", {}),
    (2, "dss", "Designated substance / hazardous materials confirmation", 1, "possession", "Andrew", "", {}),
    (2, "sign", "Job site sign & Notice of Project posted", 1, "possession", "Dillon", "", {}),
    (2, "fence", "Site fencing, tree protection & erosion / sediment controls", 2, "possession, aw_fence", "Site super",
     "CVC condition: ESC installed before construction and maintained until stabilized.", {}),
    (2, "temp_fac", "Temp facilities - porta potty, bins, security, internet", 1, "possession", "Site super", "", {}),
    (2, "salvage", "Remove, label & store salvage items", 2, "bk_audit, hst_doc", "Site super",
     "Pocket doors (hall/sitting), stained glass (sitting & powder), foyer glass door, intercom panel. Protect fireplace.", {}),

    # ------------------------------------------------------------------
    (1, "PH_DEMO", "Interior Demolition (non-structural, pre-permit)", S, "", "", "", {}),
    (2, "abate", "Asbestos abatement - 2nd floor SW & SE bedroom ceilings", 3, "salvage, dss, aw_demo", "Demo trade", "", {}),
    (2, "demo_loft", "Strip-out - loft / 3rd floor", 4, "salvage, aw_demo", "Demo trade", "", {}),
    (2, "demo_2nd", "Strip-out - second floor", 7, "abate, demo_loft", "Demo trade", "", {}),
    (2, "demo_main", "Strip-out - main floor", 7, "demo_2nd", "Demo trade", "", {}),
    (2, "demo_bsmt", "Strip-out - basement", 5, "demo_main", "Demo trade", "", {}),
    (2, "hvac_walk", "HVAC designer as-is walk (ducting exposed)", 1, "demo_2nd", "Andrew", "Before mid-November.", {}),
    (2, "ri_assess", "Assess existing rough-ins - keep / replace decisions", 3, "demo_main", "Dan / Andrew / trades",
     "Wiring, HVAC ducting. Focus on Stephen's worst-performing rooms (RFI-007).", {}),
    (2, "demo_done", "Interior demolition complete", 0, "demo_bsmt, ri_assess", "", "", {"deadline": date(2026, 11, 20)}),

    # ------------------------------------------------------------------
    (1, "PH_UTIL", "Utilities & Temporary Services", S, "", "", "", {}),
    (2, "temp_hydro", "Disconnect overhead service; temp pole & panel", 2, "aw_hydro_book, demo_done", "Electrician / Alectra",
     "Interim plan until permanent service.", {}),
    (2, "ug_install", "Permanent hydro service (underground if approved)", 5, "ug_dec, backfill", "Alectra / Electrician", "", {}),
    (2, "gas_rr", "Gas service remove & reinstall as required (Enbridge)", 3, "backfill", "Gas fitter", "", {}),

    # ------------------------------------------------------------------
    (1, "PH_SITE", "Site Work, Excavation & Foundations (rear addition)", S, "", "", "", {}),
    (2, "cvc_notice", "CVC 48-hr notice of commencement; permit card posted", 1, "permit", "Andrew", "Diarize.", {}),
    (2, "stakeout", "Survey stake-out of addition", 1, "cvc_notice", "Surveyor", "", {}),
    (2, "ext_demo", "Structural demo - rear deck / porch, openings, 2nd-floor bump-out wall", 4,
     "cvc_notice, temp_hydro, demo_done", "Demo trade", "", {}),
    (2, "topsoil", "Strip & stockpile topsoil", 1, "stakeout", "Excavator", "", {}),
    (2, "excavate", "Excavate rear addition", 3, "topsoil, aw_fdn, ext_demo", "Excavator", "Normal soil assumed.", {}),
    (2, "underpin", "Underpinning - 8 sections, sequenced", 10, "excavate", "Foundation trade", "", {}),
    (2, "ftg_form", "Layout, pin & form footings", 2, "excavate", "Foundation trade", "", {}),
    (2, "ftg_insp", "Footing inspection", 1, "ftg_form", "City inspector", "", {}),
    (2, "ftg_pour", "Pour footings (winter heat)", 1, "ftg_insp", "Foundation trade", "", {}),
    (2, "wall_form", "Form foundation walls & install rebar", 3, "ftg_pour", "Foundation trade", "", {}),
    (2, "eng_insp", "Engineer's rebar inspection", 1, "wall_form", "Structural engineer", "", {}),
    (2, "wall_pour", "Pour foundation walls", 1, "eng_insp, underpin", "Foundation trade", "", {}),
    (2, "wall_strip", "Strip forms; check heights & anchor bolts", 2, "wall_pour", "Foundation trade", "", {}),
    (2, "wall_wp", "Membrane, 2\" exterior insulation, weepers, stone & sump", 3, "wall_strip", "Foundation trade", "", {}),
    (2, "asbuilt", "Foundation as-built survey", 1, "wall_strip", "Surveyor", "", {}),
    (2, "bf_insp", "Backfill inspection", 1, "wall_wp", "City inspector", "", {}),
    (2, "backfill", "Backfill & rough grade", 2, "bf_insp, asbuilt", "Excavator", "", {}),

    # ------------------------------------------------------------------
    (1, "PH_FRAME", "Framing & Structure", S, "", "", "", {}),
    (2, "int_frame", "Interior structural alterations - openings, beams, bulkheads", 10,
     "permit, demo_done, aw_lumber, aw_steel", "Framer", "Can proceed in parallel with foundations.", {}),
    (2, "steel", "Install steel beams / posts", 2, "backfill, aw_steel", "Framer / welder", "", {}),
    (2, "fr_deck", "Frame main floor deck - rear addition", 4, "steel, aw_lumber", "Framer", "", {}),
    (2, "fr_walls", "Frame addition walls & flat roof", 6, "fr_deck", "Framer", "", {}),
    (2, "fr_bays", "Cantilevered box bays, 2nd-floor infill & balcony", 8, "fr_walls, ext_demo", "Framer", "", {}),
    (2, "sheath", "Roof & wall sheathing; Zip-R outsulation", 6, "fr_bays", "Framer", "", {}),
    (2, "fr_review", "Framing review - point loads, engineer review, deficiencies", 3, "sheath, int_frame", "Framer / engineer", "", {}),

    # ------------------------------------------------------------------
    (1, "PH_ENV", "Building Envelope - Weathertight", S, "", "", "", {}),
    (2, "temp_encl", "Temporary enclosure, tarping & winter heat", 3, "fr_walls", "Site super", "", {}),
    (2, "flat_roof", "Flat roof membrane & flashings (addition, balcony)", 3, "sheath", "Roofer", "", {}),
    (2, "shingle", "Shingle roof tie-ins & repairs", 3, "flat_roof", "Roofer", "", {}),
    (2, "win_exist", "Replace designated windows & doors in existing openings", 8, "win_fab, demo_done", "Window installer",
     "Front windows to remain per budget; removal per demolition drawings.", {}),
    (2, "win_add", "Install windows & doors - rear addition & box bays", 4, "win_fab, sheath", "Window installer", "", {}),
    (2, "skylight", "Install architectural skylight", 2, "win_fab, flat_roof", "Roofer", "", {}),
    (2, "watertight", "Watertight", 0, "flat_roof, shingle, win_exist, win_add, skylight", "", "",
     {"deadline": date(2027, 1, 15)}),

    # ------------------------------------------------------------------
    (1, "PH_SLAB", "Basement Slab & Underground Services", S, "", "", "", {}),
    (2, "ug_plumb", "Subgrade drains & underground plumbing", 4, "backfill, slab_dec, aw_plumb, temp_encl", "Plumber", "", {}),
    (2, "drain_insp", "Underground drain inspection", 1, "ug_plumb", "City inspector", "", {}),
    (2, "slab_ins", "Under-slab insulation, radon / vapour barrier", 2, "drain_insp", "Concrete trade", "", {}),
    (2, "hydronic", "In-floor hydronic tubing (if carried in mechanical design)", 2, "slab_ins, aw_hvac", "Mechanical", "", {}),
    (2, "slab_pour", "Pour basement slab", 1, "hydronic", "Concrete trade", "", {}),

    # ------------------------------------------------------------------
    (1, "PH_RI", "Mechanical, Electrical & Low-Voltage Rough-In", S, "", "", "", {}),
    (2, "ri_plumb", "Rough-in plumbing - existing house", 10, "int_frame, temp_hydro, aw_plumb, sel_plumb", "Plumber", "", {}),
    (2, "ri_hvac", "Rough-in HVAC - duct rework, heat pump / HRV-ERV - existing house", 12, "int_frame, temp_hydro, aw_hvac",
     "HVAC", "Reuse 2 Enercare units unless Net Zero heat pump route chosen.", {}),
    (2, "ri_gas", "Rough-in gas - fireplace, range, BBQ manifold", 3, "ri_hvac:SS+5, sel_fp", "Gas fitter", "", {}),
    (2, "ri_elec", "Rough-in electrical incl. EV charger & solar conduit - existing house", 12,
     "ri_plumb:SS+3, aw_elec, sel_elec", "Electrician", "", {}),
    (2, "ri_lv", "Rough-in low voltage - data, security, AV, intercom reinstatement", 5, "ri_elec:SS+5", "LV contractor", "", {}),
    (2, "ri_walk", "Rough-in walkthrough with client & ID; changes", 2, "ri_plumb, ri_hvac, ri_gas, ri_elec, ri_lv",
     "Andrew / Stephen / Vicky", "", {}),
    (2, "ri_insp", "Rough-in inspections - existing house (plumbing, HVAC, framing, ESA)", 3, "ri_walk", "City / ESA", "", {}),
    (2, "ri_done", "Rough-in inspections passed - existing house", 0, "ri_insp", "", "", {"deadline": date(2027, 2, 5)}),
    (2, "ri_add", "Rough-in all trades - rear addition & box bays", 8, "watertight, fr_review, ri_walk, slab_pour",
     "All MEP trades", "", {}),
    (2, "ri_insp_add", "Rough-in inspections - rear addition", 2, "ri_add", "City / ESA", "", {}),

    # ------------------------------------------------------------------
    (1, "PH_DW", "Insulation, Air Sealing & Drywall", S, "", "", "", {}),
    (2, "aeroseal", "AeroSeal duct sealing (reused ductwork)", 2, "ri_hvac", "HVAC", "~$2,500/unit; not in Rev 3.", {}),
    (2, "insul", "Insulation & VB - existing house, sound insulation", 6, "ri_done, aw_drywall", "Insulation trade",
     "Older front section: brick, minimal insulation - priority for comfort.", {}),
    (2, "insul_add", "Insulation & VB - addition, ceiling spray foam", 3, "ri_insp_add, insul", "Insulation trade", "", {}),
    (2, "insul_insp", "Insulation / VB inspection", 1, "insul_add", "City inspector", "", {}),
    (2, "bd_mid", "Mid-construction blower door test", 1, "insul_insp, win_exist, win_add", "Building Knowledge", "", {}),
    (2, "board", "Board drywall", 12, "bd_mid", "Drywall trade", "", {}),
    (2, "tape", "Tape & finish drywall", 12, "board:SS+5", "Drywall trade", "", {}),
    (2, "prime", "Prime drywall (tinted)", 3, "tape", "Painter", "", {}),
    (2, "plaster", "Plaster mouldings", 6, "prime, sel_trim", "Plaster trade", "", {}),
    (2, "hvac_on", "Interior environment controlled (permanent HVAC running)", 0, "prime, ug_install", "HVAC", "", {}),

    # ------------------------------------------------------------------
    (1, "PH_FIN", "Interior Finishes", S, "", "", "", {}),
    (2, "tile_floor", "Floor tile & Schluter waterproofing", 12, "prime, aw_finish, sel_floor+15", "Tile setter", "", {}),
    (2, "hardwood", "Hardwood & LVP", 15, "tile_floor:SS+5, hvac_on, sel_floor+20", "Flooring trade", "", {}),
    (2, "stairs", "Install interior stairs & railings", 5, "hardwood:SS+5, stair_shops", "Stair supplier", "", {}),
    (2, "trim", "Doors, trim, panelling, closets & shelving", 25, "plaster, aw_finish, sel_trim+20", "Trim carpenter", "", {}),
    (2, "cab_inst", "Install cabinetry - kitchen, island, servery, vanities, laundry", 10, "cab_fab, prime", "Cabinetmaker", "", {}),
    (2, "counter", "Template, fabricate & install countertops", 12, "cab_inst, sel_counter", "Countertop fabricator", "", {}),
    (2, "wall_tile", "Wall tile, showers & backsplashes", 15, "counter, tile_floor", "Tile setter", "", {}),
    (2, "paint", "Interior painting", 25, "trim:SS+10, sel_paint", "Painter", "", {}),
    (2, "heritage", "Reinstall salvaged items; fireplace brick cladding & mantel", 5, "trim", "Trim carpenter / mason", "", {}),
    (2, "glass", "Shower glass - template, fabricate, install", 12, "wall_tile", "Glass supplier", "", {}),
    (2, "appl", "Appliances delivery & install", 2, "counter, sel_appl+30", "Appliance supplier", "", {}),
    (2, "plumb_fin", "Plumbing finishing - fixtures", 5, "wall_tile, counter", "Plumber", "", {}),
    (2, "elec_fin", "Electrical finishing - devices, pot lights, fixtures", 6, "paint:SS+10", "Electrician", "", {}),
    (2, "hvac_fin", "HVAC finishing - grilles, humidifier, HRV balancing", 3, "paint", "HVAC", "", {}),
    (2, "hardware", "Door hardware, bath accessories, mirrors", 4, "paint, glass", "Trim carpenter", "", {}),
    (2, "floor_fin", "Sand & finish hardwood; stain stairs", 5, "paint, stairs", "Flooring trade", "", {}),
    (2, "touchup", "Final trade & paint touch-ups", 5, "floor_fin, hardware, elec_fin, plumb_fin, hvac_fin, appl, heritage",
     "All trades", "", {}),

    # ------------------------------------------------------------------
    (1, "PH_EXT", "Exterior Finishes & Landscaping", S, "", "", "", {}),
    (2, "siding", "Cedar shake siding", 10, "watertight, sel_ext", "Siding trade", "", {"snet": date(2027, 4, 5)}),
    (2, "ext_trim", "Exterior trim, columns, fascia & soffits", 15, "siding:SS+5", "Exterior carpenter", "", {}),
    (2, "masonry", "Masonry repointing & flagstone front porch", 8, "watertight", "Mason", "", {"snet": date(2027, 4, 19)}),
    (2, "porch_ceil", "Porch ceilings - cedar T&G", 4, "ext_trim", "Exterior carpenter", "", {}),
    (2, "troughs", "Eavestroughs & downspouts", 3, "ext_trim", "Eavestrough trade", "", {}),
    (2, "solar", "Solar PV (required for Net Zero - not in budget)", 5, "flat_roof, nz_dec", "Solar contractor",
     "$30-50K. Scope gap - confirm with Net Zero decision.", {"snet": date(2027, 5, 3)}),
    (2, "deck", "Rear deck (Trex), balcony & exterior railings", 8, "ext_trim, sel_ext", "Deck builder", "", {}),
    (2, "ext_paint", "Exterior painting & caulking", 12, "ext_trim, porch_ceil", "Painter", "", {}),
    (2, "garage_dr", "Garage door", 2, "ext_paint", "Garage door supplier", "", {}),
    (2, "hardscape", "Hard landscaping - interlock, walls, fences", 15, "deck, ext_paint, masonry", "Landscaper",
     "Landscaping allowance $100,000.", {}),
    (2, "softscape", "Soft landscaping & restoration of disturbed areas", 8, "hardscape", "Landscaper",
     "CVC condition: restore to existing condition or better.", {}),
    (2, "cvc_close", "CVC notice of completion", 0, "softscape", "Andrew", "", {}),

    # ------------------------------------------------------------------
    (1, "PH_CLOSE", "Commissioning, Inspections & Handover", S, "", "", "", {}),
    (2, "bd_final", "Final blower door & EnerGuide / Net Zero labelling test", 2, "touchup", "Building Knowledge", "", {}),
    (2, "duct_clean", "Duct cleaning", 1, "touchup", "Duct cleaner", "", {}),
    (2, "final_insp", "Final building inspection", 3, "touchup, ext_paint", "City inspector", "", {}),
    (2, "clean", "Final cleaning, windows & power wash", 4, "final_insp, duct_clean", "Cleaners", "", {}),
    (2, "pdi", "Client walkthrough & deficiency list", 2, "clean", "Andrew / Stephen", "", {}),
    (2, "occupancy", "Occupancy / substantial completion", 0, "clean, bd_final, cvc_close", "", "",
     {"deadline": date(2027, 8, 30)}),
    (2, "movein", "Client move-in", 0, "occupancy, pdi", "Stephen (Owner)", "", {"deadline": date(2027, 9, 1)}),
    (2, "defs", "Deficiency completion", 10, "pdi", "All trades", "", {}),
]


def ldt(d, hour=8):
    return LocalDateTime.of(d.year, d.month, d.day, hour, 0)


def parse_pred(token):
    token = token.strip()
    lag = 0
    rtype = RelationType.FINISH_START
    for sign in ("+", "-"):
        if sign in token:
            token, lag_s = token.split(sign, 1)
            lag = int(lag_s) * (1 if sign == "+" else -1)
            break
    if ":" in token:
        token, kind = token.split(":")
        rtype = {"SS": RelationType.START_START, "FF": RelationType.FINISH_FINISH,
                 "SF": RelationType.START_FINISH, "FS": RelationType.FINISH_START}[kind]
    return token, rtype, lag


def build(out_path):
    project = ProjectFile()
    props = project.getProjectProperties()
    props.setProjectTitle("Neumann Residence - Construction Schedule")
    props.setName("Neumann Residence - 1567 Steveles Cres, Mississauga")
    props.setCompany("Chatsworth Fine Homes")
    props.setManager("Andrew Ducker")
    props.setAuthor("Chatsworth Fine Homes")
    props.setStartDate(ldt(PROJECT_START))
    props.setStatusDate(ldt(STATUS_DATE))
    props.setCurrentDate(ldt(STATUS_DATE))
    props.setComments("Renovation & rear addition under CCDC 5A. Built from Neumann Admin project records to Oct 8, 2026.")

    cal = project.addDefaultBaseCalendar()
    cal.setName("Chatsworth Standard")
    for start, finish, name in HOLIDAYS:
        exc = cal.addCalendarException(LocalDate.of(start.year, start.month, start.day),
                                       LocalDate.of(finish.year, finish.month, finish.day))
        exc.setName(name)
    project.setDefaultCalendar(cal)

    project.getCustomFields().getOrCreate(TaskField.TEXT1).setAlias("Responsible")

    tasks = {}
    phase = None
    for level, key, name, dur, preds, resp, notes, opts in ROWS:
        if level == 1:
            phase = project.addTask()
            phase.setName(name)
            phase.setSummary(True)
            tasks[key] = phase
            continue
        t = phase.addTask()
        t.setName(name)
        t.setDuration(Duration.getInstance(float(dur), TimeUnit.DAYS))
        t.setActualDuration(Duration.getInstance(0.0, TimeUnit.DAYS))
        t.setRemainingDuration(Duration.getInstance(float(dur), TimeUnit.DAYS))
        t.setPercentageComplete(0.0)
        t.setMilestone(dur == 0)
        if resp:
            t.setText(1, resp)
        if notes:
            t.setNotes(notes)
        if "snet" in opts:
            t.setConstraintType(ConstraintType.START_NO_EARLIER_THAN)
            t.setConstraintDate(ldt(opts["snet"]))
        if "deadline" in opts:
            t.setDeadline(ldt(opts["deadline"], 17))
        tasks[key] = t

    for level, key, name, dur, preds, resp, notes, opts in ROWS:
        if level == 1 or not preds:
            continue
        succ = tasks[key]
        for tok in preds.split(","):
            pkey, rtype, lag = parse_pred(tok)
            if pkey not in tasks:
                raise KeyError(f"{key}: unknown predecessor {pkey}")
            succ.addPredecessor(Relation.Builder()
                                .predecessorTask(tasks[pkey])
                                .type(rtype)
                                .lag(Duration.getInstance(float(lag), TimeUnit.DAYS)))

    MicrosoftScheduler().schedule(project, ldt(PROJECT_START))

    for level, key, name, dur, preds, resp, notes, opts in ROWS:
        t = tasks[key]
        if opts.get("done"):
            t.setPercentageComplete(100.0)
            t.setActualStart(t.getStart())
            t.setActualFinish(t.getFinish())
        if level == 2:
            t.setBaselineStart(t.getStart())
            t.setBaselineFinish(t.getFinish())
            t.setBaselineDuration(t.getDuration())

    MSPDIWriter().write(project, out_path)
    return project, tasks


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "Neumann_Residence_Schedule.xml"
    project, tasks = build(out)
    for level, key, name, dur, preds, resp, notes, opts in ROWS:
        t = tasks[key]
        flag = ""
        if "deadline" in opts and t.getFinish() is not None:
            fin = t.getFinish().toLocalDate()
            dl = opts["deadline"]
            flag = f"  <-- target {dl}" + ("  LATE" if str(fin) > str(dl) else "  ok")
        print(f"{'  ' * (level - 1)}{key:12s} {str(t.getStart())[:10]} -> {str(t.getFinish())[:10]}  {name[:60]}{flag}")
