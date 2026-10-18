import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList

def generate_lead_tracking_system():
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # BRANDING & COLOR PALETTE
    # (Tailored for AuthBridge B2B Corporate Identity)
    # -------------------------------------------------------------
    FONT_FAMILY = "Segoe UI"
    
    # Core Colors
    NAVY_DARK = "0F2537"       # Deep Executive Navy
    NAVY_BLUE = "1B365D"       # Corporate Primary Header
    SLATE_MED = "2C5E8A"       # Subheaders & Table Headers
    ICE_BLUE = "EBF3FA"        # KPI Card Fill
    WHITE = "FFFFFF"
    GRAY_LIGHT = "F8F9FA"      # Background Zebra
    GRAY_BORDER = "D3D8DE"     # Subtle Elegant Borders
    GRAY_TEXT = "5A6578"       # Secondary Labels
    DARK_TEXT = "1C2530"       # Primary Content Text
    
    # Alert & Status Accents
    GREEN_FILL = "E2EFDA"      # Won / Converted Soft Green
    GREEN_TEXT = "276A3C"
    
    RED_FILL = "FCE8E6"        # Overdue / Critical Soft Coral
    RED_TEXT = "B3261E"
    RED_BORDER = "F5A39E"
    
    GRAY_PILL_FILL = "F2F2F2"  # Lost / Inactive Neutral
    GRAY_PILL_TEXT = "7F7F7F"
    
    # Fonts
    font_hero_title = Font(name=FONT_FAMILY, size=15, bold=True, color=WHITE)
    font_hero_sub = Font(name=FONT_FAMILY, size=9.5, italic=True, color="D0E1F9")
    font_sec_hdr = Font(name=FONT_FAMILY, size=11, bold=True, color=NAVY_DARK)
    font_table_hdr = Font(name=FONT_FAMILY, size=10, bold=True, color=WHITE)
    font_data = Font(name=FONT_FAMILY, size=10, color=DARK_TEXT)
    font_data_bold = Font(name=FONT_FAMILY, size=10, bold=True, color=DARK_TEXT)
    font_kpi_lbl = Font(name=FONT_FAMILY, size=8.5, bold=True, color=GRAY_TEXT)
    font_kpi_val = Font(name=FONT_FAMILY, size=18, bold=True, color=NAVY_DARK)
    font_kpi_val_alert = Font(name=FONT_FAMILY, size=18, bold=True, color=RED_TEXT)
    font_kpi_sub = Font(name=FONT_FAMILY, size=8, italic=True, color=GRAY_TEXT)
    font_mini_lbl = Font(name=FONT_FAMILY, size=7.5, bold=True, color=GRAY_TEXT)
    font_mini_val = Font(name=FONT_FAMILY, size=13, bold=True, color=NAVY_DARK)

    # Borders
    thin_border_side = Side(border_style="thin", color=GRAY_BORDER)
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    card_top_side = Side(border_style="medium", color=SLATE_MED)
    card_border = Border(left=thin_border_side, right=thin_border_side, top=card_top_side, bottom=thin_border_side)
    
    tot_top_side = Side(border_style="thin", color=NAVY_BLUE)
    tot_bot_side = Side(border_style="double", color=NAVY_BLUE)
    total_border = Border(top=tot_top_side, bottom=tot_bot_side, left=thin_border_side, right=thin_border_side)

    # Fills
    fill_navy_hero = PatternFill(start_color=NAVY_DARK, end_color=NAVY_DARK, fill_type="solid")
    fill_navy_sub = PatternFill(start_color=NAVY_BLUE, end_color=NAVY_BLUE, fill_type="solid")
    fill_slate_hdr = PatternFill(start_color=SLATE_MED, end_color=SLATE_MED, fill_type="solid")
    fill_card = PatternFill(start_color=ICE_BLUE, end_color=ICE_BLUE, fill_type="solid")
    fill_card_alert = PatternFill(start_color="FFF0F0", end_color="FFF0F0", fill_type="solid")
    fill_zebra = PatternFill(start_color="F5F8FB", end_color="F5F8FB", fill_type="solid")
    fill_total = PatternFill(start_color="E7ECF2", end_color="E7ECF2", fill_type="solid")

    # Alignments
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # Number Formats
    CURRENCY_FORMAT = '₹#,##,##0'
    PERCENT_FORMAT = '0.0%'
    DATE_FORMAT = 'yyyy-mm-dd'
    INT_FORMAT = '#,##0'

    # =============================================================
    # 1. WORKSHEET: LEAD TRACKER
    # =============================================================
    ws_leads = wb.active
    ws_leads.title = "Lead Tracker"
    ws_leads.views.sheetView[0].showGridLines = True

    lead_headers = [
        "Lead ID",
        "Customer Name",
        "Company Name",
        "Contact Number",
        "Lead Source",
        "Interested Service",
        "Estimated Budget (₹)",
        "Lead Status",
        "Last Contact Date",
        "Next Follow-up Date",
        "Remarks"
    ]

    # 20 Fictional B2B Records tailored to AuthBridge BGV & KYC services
    # Reference Date: Mid October 2026
    sample_leads = [
        ("L-1001", "Rajesh Sharma", "Flipkart Internet Pvt Ltd", "+91 98201 45678", "LinkedIn", "Employee Background Verification", 1450000, "Proposal Sent", datetime.date(2026, 10, 12), datetime.date(2026, 10, 15), "Sent commercial proposal for 4,500 annual white-collar candidate checks; awaiting CFO sign-off."),
        ("L-1002", "Priya Nair", "Swiggy (Bundl Technologies)", "+91 98112 34567", "Website", "Digital Address Verification", 2200000, "Qualified", datetime.date(2026, 10, 14), datetime.date(2026, 10, 21), "High interest in contactless GPS-geotagged address verification for 15,000 pan-India fleet partners."),
        ("L-1003", "Amit Verma", "Razorpay Payments Ltd", "+91 98334 11223", "Referral", "Instant KYC & Digital ID Suite", 1850000, "Converted", datetime.date(2026, 10, 5), None, "Contract executed for automated merchant onboarding API, Aadhaar OCR, and bank account validation."),
        ("L-1004", "Sneha Patel", "Infosys Technologies Ltd", "+91 98450 99887", "Email", "Employee Background Verification", 3500000, "Meeting Scheduled", datetime.date(2026, 10, 16), datetime.date(2026, 10, 22), "Product demo booked with Talent Acquisition VP and Compliance Director regarding fast-track BGV API."),
        ("L-1005", "Vikram Singhania", "Tata 1mg Healthcare", "+91 98101 66554", "Cold Call", "Vendor Due Diligence", 650000, "Contacted", datetime.date(2026, 10, 13), datetime.date(2026, 10, 20), "Spoke with Head of Procurement; shared technical brochure on supply chain and pharmacy vendor risk screening."),
        ("L-1006", "Ananya Iyer", "Zomato Media Pvt Ltd", "+91 98710 44332", "LinkedIn", "Workforce & Blue-Collar Verification", 1950000, "Proposal Sent", datetime.date(2026, 10, 10), datetime.date(2026, 10, 14), "Follow-up required on RFP submission for criminal and identity checks across delivery personnel."),
        ("L-1007", "Rohan Kapoor", "Delhivery Logistics", "+91 99100 88776", "Website", "Criminal & Court Records Search", 1200000, "Interested", datetime.date(2026, 10, 15), datetime.date(2026, 10, 23), "Requested technical architecture specs on automated e-Courts search and nationwide police record checks."),
        ("L-1008", "Neha Gupta", "Urban Company", "+91 98220 33445", "Referral", "Instant KYC & Digital ID Suite", 850000, "Converted", datetime.date(2026, 9, 28), None, "Closed annual agreement for service professional biometric KYC and document verification."),
        ("L-1009", "Siddharth Mehta", "One97 Communications (Paytm)", "+91 98188 77665", "Cold Call", "AML & Financial Sanctions Screening", 1600000, "Qualified", datetime.date(2026, 10, 11), datetime.date(2026, 10, 24), "Passed BANT qualification; requirements mapped to AuthBridge PEP, Interpol, and global sanctions API."),
        ("L-1010", "Kavita Reddy", "Pine Labs Pvt Ltd", "+91 98480 22334", "LinkedIn", "API Onboarding & OCR Suite", 950000, "Interested", datetime.date(2026, 10, 8), datetime.date(2026, 10, 16), "Evaluating optical character recognition (OCR) speed and accuracy benchmark against incumbent provider."),
        ("L-1011", "Arpit Bansal", "Unacademy EdTech", "+91 99300 55667", "Email", "Education & Credential Check", 450000, "Lost", datetime.date(2026, 10, 2), None, "Paused all external hiring screening integrations due to internal organization budget freeze."),
        ("L-1012", "Pooja Joshi", "Groww (Nextbillion Tech)", "+91 98205 11998", "Website", "Instant KYC & Digital ID Suite", 2100000, "Meeting Scheduled", datetime.date(2026, 10, 17), datetime.date(2026, 10, 25), "Demo scheduled with VP of Customer Onboarding for SEBI/KRA compliant digital verification suite."),
        ("L-1013", "Manish Agarwal", "Zepto (KiranaKart)", "+91 98115 66778", "Cold Call", "Workforce & Blue-Collar Verification", 1750000, "New Lead", datetime.date(2026, 10, 16), datetime.date(2026, 10, 19), "Lead generated from dark-store network expansion news; introductory BDR outreach drafted."),
        ("L-1014", "Tarun Batra", "CRED (Dreamplug Tech)", "+91 98199 44556", "Referral", "Executive Due Diligence", 780000, "Qualified", datetime.date(2026, 10, 9), datetime.date(2026, 10, 21), "VP Risk approved technical scope for C-suite and leadership background due diligence."),
        ("L-1015", "Divya Sundaram", "Shadowfax Technologies", "+91 98455 33221", "LinkedIn", "Digital Address Verification", 1300000, "Contacted", datetime.date(2026, 10, 11), datetime.date(2026, 10, 18), "Spoke with Logistics Partner Operations Manager; sent comparative turnaround-time (TAT) performance report."),
        ("L-1016", "Nitin Saxena", "Ather Energy", "+91 98212 99001", "Email", "Vendor Due Diligence", 520000, "Lost", datetime.date(2026, 9, 25), None, "Decided to renew with incumbent local vendor citing multi-year volume commitment lock-in."),
        ("L-1017", "Swati Deshmukh", "Meesho Online Marketplace", "+91 98104 12345", "Website", "Instant KYC & Digital ID Suite", 2400000, "Interested", datetime.date(2026, 10, 14), datetime.date(2026, 10, 22), "High volume seller onboarding verification inquiry; requested developer sandbox environment access."),
        ("L-1018", "Gaurav Chopra", "Nykaa E-Retail Ltd", "+91 98765 43210", "LinkedIn", "Employee Background Verification", 1150000, "New Lead", datetime.date(2026, 10, 17), datetime.date(2026, 10, 20), "Connected with Assistant Director HR on LinkedIn; preparing customized corporate presentation deck."),
        ("L-1019", "Meghna Sen", "Spinny Car Hub", "+91 99201 88223", "Cold Call", "Criminal & Court Records Search", 600000, "Contacted", datetime.date(2026, 10, 12), datetime.date(2026, 10, 16), "Follow-up call scheduled regarding driver and vehicle inspector background checks."),
        ("L-1020", "Rahul Kulkarni", "PolicyBazaar Insurance", "+91 98180 55443", "Email", "AML & Financial Sanctions Screening", 1400000, "New Lead", datetime.date(2026, 10, 18), datetime.date(2026, 10, 23), "Outbound BDR sequence initiated focusing on IRDAI anti-fraud compliance and agent onboarding rules.")
    ]

    # Write Headers
    ws_leads.append(lead_headers)
    ws_leads.row_dimensions[1].height = 26
    for col_num in range(1, len(lead_headers) + 1):
        cell = ws_leads.cell(row=1, column=col_num)
        cell.font = font_table_hdr
        cell.fill = fill_navy_sub
        cell.alignment = align_center

    # Write Rows
    for r_idx, row_data in enumerate(sample_leads, start=2):
        ws_leads.append(list(row_data))
        ws_leads.row_dimensions[r_idx].height = 20
        for c_idx in range(1, len(row_data) + 1):
            cell = ws_leads.cell(row=r_idx, column=c_idx)
            cell.font = font_data
            cell.border = thin_border
            
            # Alignments & formats
            if c_idx == 1: # Lead ID
                cell.alignment = align_center
                cell.font = font_data_bold
            elif c_idx in [4, 5, 8]: # Contact, Source, Status
                cell.alignment = align_center
            elif c_idx == 7: # Budget
                cell.number_format = CURRENCY_FORMAT
                cell.alignment = align_right
            elif c_idx in [9, 10]: # Dates
                cell.number_format = DATE_FORMAT
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    # Freeze Header Row
    ws_leads.freeze_panes = "A2"

    # Excel Table
    tab_ref = f"A1:K{len(sample_leads) + 1}"
    lead_table = Table(displayName="LeadTrackerTable", ref=tab_ref)
    table_style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )
    lead_table.tableStyleInfo = table_style
    ws_leads.add_table(lead_table)

    # Data Validations (Dropdowns)
    source_list = '"LinkedIn,Email,Cold Call,Referral,Website"'
    status_list = '"New Lead,Contacted,Interested,Qualified,Meeting Scheduled,Proposal Sent,Converted,Lost"'

    dv_source = DataValidation(type="list", formula1=source_list, allow_blank=True)
    dv_source.error = "Please select a valid Lead Source from the dropdown list."
    dv_source.errorTitle = "Invalid Lead Source"
    dv_source.prompt = "Choose the lead acquisition channel."
    dv_source.promptTitle = "Lead Source"
    dv_source.showInputMessage = True
    dv_source.showErrorMessage = True
    ws_leads.add_data_validation(dv_source)
    dv_source.add("E2:E100")

    dv_status = DataValidation(type="list", formula1=status_list, allow_blank=True)
    dv_status.error = "Please select a valid Lead Status from the dropdown list."
    dv_status.errorTitle = "Invalid Lead Status"
    dv_status.prompt = "Select current sales pipeline stage."
    dv_status.promptTitle = "Lead Status"
    dv_status.showInputMessage = True
    dv_status.showErrorMessage = True
    ws_leads.add_data_validation(dv_status)
    dv_status.add("H2:H100")

    # Conditional Formatting Rules
    # 1. Converted Leads (Soft Green Fill on Status Column H)
    fill_converted = PatternFill(start_color=GREEN_FILL, end_color=GREEN_FILL, fill_type="solid")
    font_converted = Font(name=FONT_FAMILY, size=10, bold=True, color=GREEN_TEXT)
    rule_converted = CellIsRule(operator="equal", formula=['"Converted"'], stopIfTrue=True, fill=fill_converted, font=font_converted)
    ws_leads.conditional_formatting.add("H2:H100", rule_converted)

    # 2. Lost Leads (Soft Neutral Gray Fill on Status Column H)
    fill_lost = PatternFill(start_color=GRAY_PILL_FILL, end_color=GRAY_PILL_FILL, fill_type="solid")
    font_lost = Font(name=FONT_FAMILY, size=10, bold=True, color=GRAY_PILL_TEXT)
    rule_lost = CellIsRule(operator="equal", formula=['"Lost"'], stopIfTrue=True, fill=fill_lost, font=font_lost)
    ws_leads.conditional_formatting.add("H2:H100", rule_lost)

    # 3. Overdue Follow-ups (Next Follow-up Date < TODAY() and Status not Converted or Lost)
    fill_overdue = PatternFill(start_color=RED_FILL, end_color=RED_FILL, fill_type="solid")
    font_overdue = Font(name=FONT_FAMILY, size=10, bold=True, color=RED_TEXT)
    border_overdue = Border(
        left=Side(style='thin', color=RED_BORDER),
        right=Side(style='thin', color=RED_BORDER),
        top=Side(style='thin', color=RED_BORDER),
        bottom=Side(style='thin', color=RED_BORDER)
    )
    rule_overdue = FormulaRule(
        formula=['AND(ISNUMBER(J2), J2<TODAY(), H2<>"Converted", H2<>"Lost")'],
        stopIfTrue=False, fill=fill_overdue, font=font_overdue, border=border_overdue
    )
    ws_leads.conditional_formatting.add("J2:J100", rule_overdue)

    # Column Widths for Lead Tracker
    column_widths_tracker = {
        "A": 12,  # Lead ID
        "B": 20,  # Customer Name
        "C": 31,  # Company Name
        "D": 18,  # Contact Number
        "E": 14,  # Lead Source
        "F": 35,  # Interested Service
        "G": 22,  # Estimated Budget
        "H": 20,  # Lead Status
        "I": 18,  # Last Contact Date
        "J": 20,  # Next Follow-up Date
        "K": 55   # Remarks
    }
    for col, width in column_widths_tracker.items():
        ws_leads.column_dimensions[col].width = width

    # =============================================================
    # 2. WORKSHEET: DASHBOARD
    # =============================================================
    ws_dash = wb.create_sheet(title="Dashboard", index=0)
    ws_dash.views.sheetView[0].showGridLines = True

    # Title Banner Block
    ws_dash.merge_cells("B2:M2")
    ws_dash["B2"] = "PROSPECTIVE CUSTOMER & SALES LEAD TRACKING SYSTEM"
    ws_dash["B2"].font = font_hero_title
    ws_dash["B2"].fill = fill_navy_hero
    ws_dash["B2"].alignment = Alignment(horizontal="center", vertical="center")
    ws_dash.row_dimensions[2].height = 36

    ws_dash.merge_cells("B3:M3")
    ws_dash["B3"] = "BDR Sales Operations & Executive Pipeline Intelligence | Portfolio Project for AuthBridge Research Services"
    ws_dash["B3"].font = font_hero_sub
    ws_dash["B3"].fill = fill_navy_sub
    ws_dash["B3"].alignment = Alignment(horizontal="center", vertical="center")
    ws_dash.row_dimensions[3].height = 20

    # Helper function to build top KPI cards
    def build_kpi_card(ws, start_col, end_col, top_row, label, formula_val, num_format, subtext, alert=False):
        start_cell = f"{start_col}{top_row}"
        end_cell = f"{end_col}{top_row}"
        ws.merge_cells(f"{start_cell}:{end_cell}")
        ws[start_cell] = label
        ws[start_cell].font = font_kpi_lbl
        ws[start_cell].alignment = align_center
        ws[start_cell].fill = fill_card_alert if alert else fill_card
        
        val_start = f"{start_col}{top_row+1}"
        val_end = f"{end_col}{top_row+1}"
        ws.merge_cells(f"{val_start}:{val_end}")
        ws[val_start] = formula_val
        ws[val_start].alignment = align_center
        ws[val_start].fill = fill_card_alert if alert else fill_card
        ws[val_start].font = font_kpi_val_alert if alert else font_kpi_val
        if num_format:
            ws[val_start].number_format = num_format
            
        sub_start = f"{start_col}{top_row+2}"
        sub_end = f"{end_col}{top_row+2}"
        ws.merge_cells(f"{sub_start}:{sub_end}")
        ws[sub_start] = subtext
        ws[sub_start].font = font_kpi_sub
        ws[sub_start].alignment = align_center
        ws[sub_start].fill = fill_card_alert if alert else fill_card
        
        cols = [chr(c) for c in range(ord(start_col), ord(end_col) + 1)]
        for c in cols:
            for r in range(top_row, top_row + 3):
                cell = ws[f"{c}{r}"]
                cell.border = thin_border
                if r == top_row:
                    cell.border = Border(top=card_top_side, left=thin_border_side, right=thin_border_side, bottom=thin_border_side)

    # Row 5-7: 5 Primary Executive KPI Cards
    build_kpi_card(ws_dash, "B", "C", 5, "TOTAL LEADS", "=COUNTA('Lead Tracker'!A2:A21)", INT_FORMAT, "Total Tracked Prospects")
    build_kpi_card(ws_dash, "D", "F", 5, "TOTAL POTENTIAL PIPELINE VALUE", "=SUMIF('Lead Tracker'!H2:H21, \"<>Lost\", 'Lead Tracker'!G2:G21)", CURRENCY_FORMAT, "Active Pipeline (Excl. Lost)")
    build_kpi_card(ws_dash, "G", "H", 5, "CONVERTED REVENUE", "=SUMIF('Lead Tracker'!H2:H21, \"Converted\", 'Lead Tracker'!G2:G21)", CURRENCY_FORMAT, "Won Deals Closed Value")
    build_kpi_card(ws_dash, "I", "J", 5, "CONVERSION RATE", "=COUNTIF('Lead Tracker'!H2:H21, \"Converted\")/COUNTA('Lead Tracker'!A2:A21)", PERCENT_FORMAT, "Converted / Total Leads")
    build_kpi_card(ws_dash, "K", "M", 5, "OVERDUE FOLLOW-UPS", "=COUNTIFS('Lead Tracker'!J2:J21, \"<\"&TODAY(), 'Lead Tracker'!J2:J21, \">0\", 'Lead Tracker'!H2:H21, \"<>Converted\", 'Lead Tracker'!H2:H21, \"<>Lost\")", INT_FORMAT, "Requires Immediate Outreach", alert=True)

    ws_dash.row_dimensions[5].height = 18
    ws_dash.row_dimensions[6].height = 28
    ws_dash.row_dimensions[7].height = 16

    # Row 9-11: Detailed Stage Breakdown Mini-Cards
    # Satisfying all prompt requirements: New Leads, Interested Leads, Qualified Leads, Meetings Scheduled, Proposals Sent, Converted Leads, Lost Leads
    ws_dash.merge_cells("B9:M9")
    ws_dash["B9"] = "SALES PIPELINE FUNNEL & STAGE VOLUME BREAKDOWN"
    ws_dash["B9"].font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=NAVY_DARK)
    ws_dash["B9"].alignment = Alignment(horizontal="left", vertical="center")
    ws_dash.row_dimensions[9].height = 20

    stage_mini_cards = [
        ("B", "B", "NEW LEADS", "=COUNTIF('Lead Tracker'!H2:H21, \"New Lead\")"),
        ("C", "C", "CONTACTED", "=COUNTIF('Lead Tracker'!H2:H21, \"Contacted\")"),
        ("D", "E", "INTERESTED LEADS", "=COUNTIF('Lead Tracker'!H2:H21, \"Interested\")"),
        ("F", "G", "QUALIFIED LEADS", "=COUNTIF('Lead Tracker'!H2:H21, \"Qualified\")"),
        ("H", "I", "MEETINGS SCHEDULED", "=COUNTIF('Lead Tracker'!H2:H21, \"Meeting Scheduled\")"),
        ("J", "K", "PROPOSALS SENT", "=COUNTIF('Lead Tracker'!H2:H21, \"Proposal Sent\")"),
        ("L", "L", "CONVERTED", "=COUNTIF('Lead Tracker'!H2:H21, \"Converted\")"),
        ("M", "M", "LOST LEADS", "=COUNTIF('Lead Tracker'!H2:H21, \"Lost\")")
    ]

    for start_c, end_c, title, form in stage_mini_cards:
        # Header Row 10
        ws_dash.merge_cells(f"{start_c}10:{end_c}10")
        ws_dash[f"{start_c}10"] = title
        ws_dash[f"{start_c}10"].font = font_mini_lbl
        ws_dash[f"{start_c}10"].alignment = align_center
        ws_dash[f"{start_c}10"].fill = fill_zebra
        
        # Value Row 11
        ws_dash.merge_cells(f"{start_c}11:{end_c}11")
        ws_dash[f"{start_c}11"] = form
        ws_dash[f"{start_c}11"].font = font_mini_val
        ws_dash[f"{start_c}11"].alignment = align_center
        ws_dash[f"{start_c}11"].fill = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
        
        cols = [chr(c) for c in range(ord(start_c), ord(end_c) + 1)]
        for c in cols:
            for r in [10, 11]:
                ws_dash[f"{c}{r}"].border = thin_border

    ws_dash.row_dimensions[10].height = 16
    ws_dash.row_dimensions[11].height = 22

    # Section Headers for Summary Tables
    ws_dash["B13"] = "SALES PIPELINE STAGE BREAKDOWN"
    ws_dash["B13"].font = font_sec_hdr

    ws_dash["I13"] = "LEAD SOURCE PERFORMANCE & CHANNELS"
    ws_dash["I13"].font = font_sec_hdr

    # -------------------------------------------------------------
    # Table 1: Stage Breakdown (B14:F23)
    # -------------------------------------------------------------
    stage_headers = ["Lead Status / Stage", "Leads", "Pipeline Value (₹)", "Avg Deal Size (₹)", "Stage Share %"]
    for idx, h in enumerate(stage_headers, start=2): # col B is 2
        col_letter = chr(ord("B") + idx - 2)
        cell = ws_dash[f"{col_letter}14"]
        cell.value = h
        cell.font = font_table_hdr
        cell.fill = fill_slate_hdr
        cell.alignment = align_center if idx != 2 else align_left
        cell.border = thin_border
    ws_dash.row_dimensions[14].height = 22

    stages = [
        "New Lead",
        "Contacted",
        "Interested",
        "Qualified",
        "Meeting Scheduled",
        "Proposal Sent",
        "Converted",
        "Lost"
    ]

    for idx, stage in enumerate(stages, start=15):
        ws_dash.row_dimensions[idx].height = 19
        
        # Stage name
        ws_dash[f"B{idx}"] = stage
        ws_dash[f"B{idx}"].font = font_data_bold
        ws_dash[f"B{idx}"].alignment = align_left
        ws_dash[f"B{idx}"].border = thin_border
        
        # Leads count: COUNTIF
        ws_dash[f"C{idx}"] = f'=COUNTIF(\'Lead Tracker\'!$H$2:$H$21, "{stage}")'
        ws_dash[f"C{idx}"].font = font_data
        ws_dash[f"C{idx}"].alignment = align_center
        ws_dash[f"C{idx}"].border = thin_border
        
        # Pipeline Value: SUMIF
        ws_dash[f"D{idx}"] = f'=SUMIF(\'Lead Tracker\'!$H$2:$H$21, "{stage}", \'Lead Tracker\'!$G$2:$G$21)'
        ws_dash[f"D{idx}"].font = font_data
        ws_dash[f"D{idx}"].alignment = align_right
        ws_dash[f"D{idx}"].number_format = CURRENCY_FORMAT
        ws_dash[f"D{idx}"].border = thin_border
        
        # Avg Deal Size: IFERROR(Value/Leads, 0)
        ws_dash[f"E{idx}"] = f'=IFERROR(D{idx}/C{idx}, 0)'
        ws_dash[f"E{idx}"].font = font_data
        ws_dash[f"E{idx}"].alignment = align_right
        ws_dash[f"E{idx}"].number_format = CURRENCY_FORMAT
        ws_dash[f"E{idx}"].border = thin_border
        
        # Stage Share %
        ws_dash[f"F{idx}"] = f'=IFERROR(C{idx}/$C$23, 0)'
        ws_dash[f"F{idx}"].font = font_data
        ws_dash[f"F{idx}"].alignment = align_right
        ws_dash[f"F{idx}"].number_format = PERCENT_FORMAT
        ws_dash[f"F{idx}"].border = thin_border

    # Table 1 Totals Row (Row 23)
    ws_dash.row_dimensions[23].height = 22
    ws_dash["B23"] = "Total / Portfolio Summary"
    ws_dash["B23"].font = font_data_bold
    ws_dash["B23"].fill = fill_total
    ws_dash["B23"].border = total_border
    
    ws_dash["C23"] = "=SUM(C15:C22)"
    ws_dash["C23"].font = font_data_bold
    ws_dash["C23"].alignment = align_center
    ws_dash["C23"].fill = fill_total
    ws_dash["C23"].border = total_border

    ws_dash["D23"] = "=SUM(D15:D22)"
    ws_dash["D23"].font = font_data_bold
    ws_dash["D23"].alignment = align_right
    ws_dash["D23"].number_format = CURRENCY_FORMAT
    ws_dash["D23"].fill = fill_total
    ws_dash["D23"].border = total_border

    ws_dash["E23"] = "=IFERROR(D23/C23, 0)"
    ws_dash["E23"].font = font_data_bold
    ws_dash["E23"].alignment = align_right
    ws_dash["E23"].number_format = CURRENCY_FORMAT
    ws_dash["E23"].fill = fill_total
    ws_dash["E23"].border = total_border

    ws_dash["F23"] = "=SUM(F15:F22)"
    ws_dash["F23"].font = font_data_bold
    ws_dash["F23"].alignment = align_right
    ws_dash["F23"].number_format = PERCENT_FORMAT
    ws_dash["F23"].fill = fill_total
    ws_dash["F23"].border = total_border

    # -------------------------------------------------------------
    # Table 2: Source Breakdown (I14:M20)
    # -------------------------------------------------------------
    source_headers = ["Lead Source", "Total Leads", "Won Leads", "Win Rate %", "Total Budget (₹)"]
    for idx, h in enumerate(source_headers):
        col_letter = chr(ord("I") + idx)
        cell = ws_dash[f"{col_letter}14"]
        cell.value = h
        cell.font = font_table_hdr
        cell.fill = fill_slate_hdr
        cell.alignment = align_center if idx != 0 else align_left
        cell.border = thin_border

    sources = ["LinkedIn", "Email", "Cold Call", "Referral", "Website"]
    for idx, src in enumerate(sources, start=15):
        ws_dash.row_dimensions[idx].height = 19
        
        # Source name
        ws_dash[f"I{idx}"] = src
        ws_dash[f"I{idx}"].font = font_data_bold
        ws_dash[f"I{idx}"].alignment = align_left
        ws_dash[f"I{idx}"].border = thin_border

        # Total Leads: COUNTIF
        ws_dash[f"J{idx}"] = f'=COUNTIF(\'Lead Tracker\'!$E$2:$E$21, "{src}")'
        ws_dash[f"J{idx}"].font = font_data
        ws_dash[f"J{idx}"].alignment = align_center
        ws_dash[f"J{idx}"].border = thin_border

        # Converted Leads: COUNTIFS
        ws_dash[f"K{idx}"] = f'=COUNTIFS(\'Lead Tracker\'!$E$2:$E$21, "{src}", \'Lead Tracker\'!$H$2:$H$21, "Converted")'
        ws_dash[f"K{idx}"].font = font_data
        ws_dash[f"K{idx}"].alignment = align_center
        ws_dash[f"K{idx}"].border = thin_border

        # Win Rate %
        ws_dash[f"L{idx}"] = f'=IFERROR(K{idx}/J{idx}, 0)'
        ws_dash[f"L{idx}"].font = font_data
        ws_dash[f"L{idx}"].alignment = align_right
        ws_dash[f"L{idx}"].number_format = PERCENT_FORMAT
        ws_dash[f"L{idx}"].border = thin_border

        # Total Budget
        ws_dash[f"M{idx}"] = f'=SUMIF(\'Lead Tracker\'!$E$2:$E$21, "{src}", \'Lead Tracker\'!$G$2:$G$21)'
        ws_dash[f"M{idx}"].font = font_data
        ws_dash[f"M{idx}"].alignment = align_right
        ws_dash[f"M{idx}"].number_format = CURRENCY_FORMAT
        ws_dash[f"M{idx}"].border = thin_border

    # Table 2 Totals Row (Row 20)
    ws_dash.row_dimensions[20].height = 22
    ws_dash["I20"] = "Total Channels"
    ws_dash["I20"].font = font_data_bold
    ws_dash["I20"].fill = fill_total
    ws_dash["I20"].border = total_border

    ws_dash["J20"] = "=SUM(J15:J19)"
    ws_dash["J20"].font = font_data_bold
    ws_dash["J20"].alignment = align_center
    ws_dash["J20"].fill = fill_total
    ws_dash["J20"].border = total_border

    ws_dash["K20"] = "=SUM(K15:K19)"
    ws_dash["K20"].font = font_data_bold
    ws_dash["K20"].alignment = align_center
    ws_dash["K20"].fill = fill_total
    ws_dash["K20"].border = total_border

    ws_dash["L20"] = "=IFERROR(K20/J20, 0)"
    ws_dash["L20"].font = font_data_bold
    ws_dash["L20"].alignment = align_right
    ws_dash["L20"].number_format = PERCENT_FORMAT
    ws_dash["L20"].fill = fill_total
    ws_dash["L20"].border = total_border

    ws_dash["M20"] = "=SUM(M15:M19)"
    ws_dash["M20"].font = font_data_bold
    ws_dash["M20"].alignment = align_right
    ws_dash["M20"].number_format = CURRENCY_FORMAT
    ws_dash["M20"].fill = fill_total
    ws_dash["M20"].border = total_border

    # -------------------------------------------------------------
    # CHARTS ON DASHBOARD
    # -------------------------------------------------------------
    # Chart 1: Lead Distribution by Status (Doughnut Chart)
    chart_status = DoughnutChart()
    chart_status.title = "Lead Distribution by Status"
    chart_status.style = 10
    
    data_status = Reference(ws_dash, min_col=3, min_row=14, max_row=22)
    labels_status = Reference(ws_dash, min_col=2, min_row=15, max_row=22)
    chart_status.add_data(data_status, titles_from_data=True)
    chart_status.set_categories(labels_status)
    chart_status.dataLabels = DataLabelList()
    chart_status.dataLabels.showVal = True
    chart_status.width = 15
    chart_status.height = 11
    ws_dash.add_chart(chart_status, "B25")

    # Chart 2: Sales Pipeline Value by Stage (Column Chart)
    chart_pipeline = BarChart()
    chart_pipeline.type = "col"
    chart_pipeline.style = 11
    chart_pipeline.title = "Sales Pipeline Value by Stage (₹)"
    chart_pipeline.y_axis.title = "Pipeline Value (INR)"
    chart_pipeline.x_axis.title = "Sales Pipeline Stage"
    chart_pipeline.legend = None # Single series
    
    data_pipeline = Reference(ws_dash, min_col=4, min_row=14, max_row=22)
    labels_pipeline = Reference(ws_dash, min_col=2, min_row=15, max_row=22)
    chart_pipeline.add_data(data_pipeline, titles_from_data=True)
    chart_pipeline.set_categories(labels_pipeline)
    chart_pipeline.dataLabels = DataLabelList()
    chart_pipeline.dataLabels.showVal = True
    chart_pipeline.width = 17
    chart_pipeline.height = 11
    ws_dash.add_chart(chart_pipeline, "H25")

    # Column Widths for Dashboard
    dash_widths = {
        "A": 3,
        "B": 22,
        "C": 12,
        "D": 22,
        "E": 20,
        "F": 16,
        "G": 18,
        "H": 18,
        "I": 18,
        "J": 14,
        "K": 14,
        "L": 14,
        "M": 22,
        "N": 3
    }
    for col, width in dash_widths.items():
        ws_dash.column_dimensions[col].width = width

    # =============================================================
    # 3. WORKSHEET: INSTRUCTIONS
    # =============================================================
    ws_inst = wb.create_sheet(title="Instructions", index=2)
    ws_inst.views.sheetView[0].showGridLines = True

    # Title Banner Block
    ws_inst.merge_cells("B2:K2")
    ws_inst["B2"] = "SYSTEM DOCUMENTATION & BDR USER MANUAL"
    ws_inst["B2"].font = font_hero_title
    ws_inst["B2"].fill = fill_navy_hero
    ws_inst["B2"].alignment = Alignment(horizontal="center", vertical="center")
    ws_inst.row_dimensions[2].height = 36

    ws_inst.merge_cells("B3:K3")
    ws_inst["B3"] = "Operating Manual for Prospective Account Tracking, Qualification, Pipeline Hygiene, and Analytics"
    ws_inst["B3"].font = font_hero_sub
    ws_inst["B3"].fill = fill_navy_sub
    ws_inst["B3"].alignment = Alignment(horizontal="center", vertical="center")
    ws_inst.row_dimensions[3].height = 20

    def write_instruction_section(ws, start_row, sec_num, sec_title, items):
        # Section Header
        ws.merge_cells(f"B{start_row}:K{start_row}")
        header_cell = ws[f"B{start_row}"]
        header_cell.value = f"{sec_num}. {sec_title}"
        header_cell.font = Font(name=FONT_FAMILY, size=11, bold=True, color=WHITE)
        header_cell.fill = fill_slate_hdr
        header_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[start_row].height = 24
        
        curr_row = start_row + 1
        for label, desc in items:
            ws.merge_cells(f"B{curr_row}:C{curr_row}")
            lbl_cell = ws[f"B{curr_row}"]
            lbl_cell.value = label
            lbl_cell.font = font_data_bold
            lbl_cell.alignment = Alignment(horizontal="left", vertical="top")
            lbl_cell.border = thin_border
            lbl_cell.fill = fill_card
            ws[f"C{curr_row}"].border = thin_border
            
            ws.merge_cells(f"D{curr_row}:K{curr_row}")
            desc_cell = ws[f"D{curr_row}"]
            desc_cell.value = desc
            desc_cell.font = font_data
            desc_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            desc_cell.border = thin_border
            
            for col_c in ["E", "F", "G", "H", "I", "J", "K"]:
                ws[f"{col_c}{curr_row}"].border = thin_border
                
            ws.row_dimensions[curr_row].height = 28
            curr_row += 1
            
        return curr_row + 1

    row_cursor = 5
    
    # Section 1
    sec1_items = [
        ("Project Objective & Purpose", "Developed as an executive sales operations portfolio project for the Business Development Representative (BDR) Intern application at AuthBridge Research Services Pvt. Ltd. Demonstrates full-cycle prospecting management, pipeline analytics, and data hygiene."),
        ("B2B Verification Domain Fit", "Tailored to AuthBridge's product suite: Employee Background Verification (BGV), Instant KYC & Digital ID, Workforce/Blue-Collar Screening, Vendor Due Diligence, Criminal & Court Records Search, and API Integrations."),
        ("Architecture & Three Sheets", "• 'Dashboard': Executive summary KPIs, stage funnel mini-cards, distribution tables, and dynamic charts.\n• 'Lead Tracker': Transactional Excel table featuring data validation, conditional formatting, and autofilters.\n• 'Instructions': Standard operating procedures, BANT framework, and KPI glossary.")
    ]
    row_cursor = write_instruction_section(ws_inst, row_cursor, 1, "PURPOSE & WORKBOOK ARCHITECTURE", sec1_items)

    # Section 2
    sec2_items = [
        ("Step 1: Adding a Prospective Lead", "Open 'Lead Tracker'. Add a new row to the table. Assign a sequential Lead ID (e.g. L-1021), enter Customer Name, Company Name, and Contact Number. The Excel table automatically expands formatting and formulas."),
        ("Step 2: Assigning Lead Source & Service", "Select the channel from the dropdown (LinkedIn, Email, Cold Call, Referral, Website). Enter the specific AuthBridge solution and estimated contract budget in INR (₹)."),
        ("Step 3: Managing Pipeline Status", "Update Lead Status as discussions advance: New Lead -> Contacted -> Interested -> Qualified -> Meeting Scheduled -> Proposal Sent -> Converted / Lost. Always select from the dropdown to ensure dashboard accuracy."),
        ("Step 4: Scheduling Follow-ups", "Enter 'Last Contact Date' and 'Next Follow-up Date' (format: YYYY-MM-DD). If a follow-up date passes without an update, the cell triggers an automated visual alert and increments the Overdue KPI on the Dashboard.")
    ]
    row_cursor = write_instruction_section(ws_inst, row_cursor, 2, "HOW TO USE THE LEAD TRACKER (DAILY BDR WORKFLOW)", sec2_items)

    # Section 3
    sec3_items = [
        ("Lead Source Dropdown", "• LinkedIn: Targeted InMail & social selling to HR/Risk heads.\n• Email: Cold outbound email campaigns & sequences.\n• Cold Call: Direct telephone outreach to key decision makers.\n• Referral: Inbound introductions from clients or business partners.\n• Website: Inbound demo requests and enterprise contact forms."),
        ("Lead Status Progression", "• New Lead: Uncontacted prospect identified from industry research.\n• Contacted: Initial touchpoint made; awaiting prospect response.\n• Interested: Prospect expressed interest in AuthBridge capabilities.\n• Qualified: Passed BANT criteria (Budget, Authority, Need, Timeline).\n• Meeting Scheduled: Technical discovery or product demo booked.\n• Proposal Sent: Commercial quotation and SLA proposal shared.\n• Converted: Contract signed / closed-won customer account.\n• Lost: Disqualified, deal lost to competitor, or cancelled project."),
        ("Input Protection", "In-cell data validation restricts values to authorized options, eliminating spelling errors that could distort analytics.")
    ]
    row_cursor = write_instruction_section(ws_inst, row_cursor, 3, "DATA VALIDATION & DROPDOWN GLOSSARY", sec3_items)

    # Section 4
    sec4_items = [
        ("Overdue Follow-up Alert (Red)", "Highlights Next Follow-up Date in soft coral with bold red text if: Next Follow-up Date < TODAY() AND Status is not 'Converted' or 'Lost'. Serves as an immediate prompt to contact the client."),
        ("Converted Lead Highlight (Green)", "Highlights the Lead Status cell in soft mint green when marked 'Converted', providing instant visual feedback for closed-won revenue."),
        ("Lost Lead Highlight (Gray)", "Formats the Lead Status cell in soft neutral gray when marked 'Lost' to indicate inactive deals without cluttering the active pipeline."),
        ("Handling Closed Leads", "For Converted or Lost accounts, leave the 'Next Follow-up Date' empty. The alert formula ignores blank cells.")
    ]
    row_cursor = write_instruction_section(ws_inst, row_cursor, 4, "CONDITIONAL FORMATTING LOGIC & VISUAL ALERTS", sec4_items)

    # Section 5
    sec5_items = [
        ("Total Leads", "=COUNTA('Lead Tracker'!A2:A21). Total number of prospective accounts logged in the tracker."),
        ("Total Potential Pipeline Value", "=SUMIF('Lead Tracker'!H2:H21, \"<>Lost\", 'Lead Tracker'!G2:G21). Total active monetary pipeline excluding closed-lost deals."),
        ("Conversion Rate %", "=COUNTIF('Lead Tracker'!H2:H21, \"Converted\")/COUNTA('Lead Tracker'!A2:A21). Closed-won deal percentage."),
        ("Overdue Follow-ups Count", "=COUNTIFS('Lead Tracker'!J2:J21, \"<\"&TODAY(), 'Lead Tracker'!J2:J21, \">0\", 'Lead Tracker'!H2:H21, \"<>Converted\", 'Lead Tracker'!H2:H21, \"<>Lost\"). Real-time count of leads requiring urgent outreach today."),
        ("Interactive Charts", "• Chart 1 (Doughnut): Proportional distribution of leads across all 8 pipeline stages.\n• Chart 2 (Column): Total pipeline monetary value (INR) mapped across each stage.")
    ]
    row_cursor = write_instruction_section(ws_inst, row_cursor, 5, "DASHBOARD FORMULA & ANALYTICS REFERENCE", sec5_items)

    # Section 6
    sec6_items = [
        ("Budget (B)", "Validate annual spend capability. (AuthBridge enterprise verification packages typically range from ₹3,00,000 to ₹35,00,000+ depending on volume)."),
        ("Authority (A)", "Engage primary decision-makers: VP/Director of Talent Acquisition, Chief Risk Officer (CRO), Compliance Head, or Chief Technology Officer (CTO)."),
        ("Need (N)", "Uncover core verification challenges: slow turnaround time (TAT), manual verification bottlenecks, regulatory compliance (SEBI/RBI/IRDAI), or high gig-worker fraud risk."),
        ("Timeline (T)", "Identify hiring surge windows, incumbent vendor renewal deadlines, or target go-live dates for API integration.")
    ]
    row_cursor = write_instruction_section(ws_inst, row_cursor, 6, "BDR SALES PLAYBOOK & BANT QUALIFICATION (AUTHBRIDGE)", sec6_items)

    # Column Widths for Instructions
    inst_widths = {
        "A": 3,
        "B": 26,
        "C": 10,
        "D": 14,
        "E": 14,
        "F": 14,
        "G": 14,
        "H": 14,
        "I": 14,
        "J": 14,
        "K": 20,
        "L": 3
    }
    for col, width in inst_widths.items():
        ws_inst.column_dimensions[col].width = width

    # Save Workbook
    filename = "Sales_Lead_Tracking_System.xlsx"
    wb.save(filename)
    print(f"Workbook successfully generated and saved as: {filename}")

if __name__ == "__main__":
    generate_lead_tracking_system()
