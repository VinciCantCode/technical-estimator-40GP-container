"""
Build Version 2 (Linked Architecture) of the Precise Pipe 40GP Container Model:
  File: Precise_Pipe_40GP_Container_Model_V2_Linked.xlsx
  Architecture:
    Tab 3 (3_ASME_Spec_DB): Raw ASME B36.10M / B36.19M dimensions & densities
    Tab 2 (2_Selectable_Master_Table): Single calculation engine for all 707 pipe specs
    Tab 1 (1_Container_Calculator): Front-end Interactive Lookup Dashboard that directly
          matches & retrieves pre-calculated results from Tab 2 via Lookup Key (Col V).
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from pipe_data import ASME_PIPE_SPECS, SCHEDULES, CARBON_SCHEDULES, STAINLESS_SCHEDULES
from build_step1_db import create_step1_workbook, INCH_TO_MM


def build_v2_linked_model(filename="Precise_Pipe_40GP_Container_Model_V2_Linked.xlsx"):
    # 1. Create Base Workbook with Tab 3 (3_ASME_Spec_DB)
    create_step1_workbook(filename)
    wb = openpyxl.load_workbook(filename)

    # 2. Build Tab 2: 2_Selectable_Master_Table (Calculation Engine)
    ws2 = wb.create_sheet(title="2_Selectable_Master_Table", index=0)
    ws2.views.sheetView[0].showGridLines = True

    # Shared styles
    NAVY_HEADER = PatternFill(start_color="0F2C59", end_color="0F2C59", fill_type="solid")
    SLATE_SUBHEADER = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    PARAM_HDR_FILL = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    PARAM_LBL_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    INPUT_CELL_FILL = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
    CALC_CELL_FILL = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    TBL_HDR_MAIN = PatternFill(start_color="0F2C59", end_color="0F2C59", fill_type="solid")
    TBL_HDR_CALC = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    TBL_HDR_OUT = PatternFill(start_color="047857", end_color="047857", fill_type="solid")
    TBL_HDR_KEY = PatternFill(start_color="475569", end_color="475569", fill_type="solid")

    ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    OUTPUT_COL_FILL = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")

    FONT_TITLE = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    FONT_SUBTITLE = Font(name="Calibri", size=10, italic=True, color="E2E8F0")
    FONT_PARAM_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    FONT_PARAM_LBL = Font(name="Calibri", size=10, bold=True, color="1E293B")
    FONT_PARAM_INP = Font(name="Calibri", size=10, bold=True, color="B45309")
    FONT_PARAM_VAL = Font(name="Calibri", size=10, bold=True, color="0F2C59")
    FONT_TBL_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    FONT_REGULAR = Font(name="Calibri", size=10, bold=False, color="1E293B")
    FONT_BOLD = Font(name="Calibri", size=10, bold=True, color="0F172A")
    FONT_OUTPUT = Font(name="Calibri", size=10, bold=True, color="065F46")
    FONT_CS_TAG = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
    FONT_SS_TAG = Font(name="Calibri", size=10, bold=True, color="0F766E")
    FONT_KEY = Font(name="Calibri", size=9, italic=True, color="64748B")

    THIN_BORDER_SIDE = Side(border_style="thin", color="CBD5E1")
    MED_NAVY_SIDE = Side(border_style="medium", color="0F2C59")
    INPUT_BORDER_SIDE = Side(border_style="thin", color="F59E0B")

    BORDER_DATA = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
    BORDER_INPUT = Border(left=INPUT_BORDER_SIDE, right=INPUT_BORDER_SIDE, top=INPUT_BORDER_SIDE, bottom=INPUT_BORDER_SIDE)
    BORDER_HDR = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=MED_NAVY_SIDE, bottom=MED_NAVY_SIDE)

    # Title Banner on Sheet 2
    ws2.merge_cells("A1:V1")
    ws2["A1"] = "PRECISE PIPE — 40GP CONTAINER LOADING MASTER SELECTABLE TABLE (CORE CALCULATION ENGINE)"
    ws2["A1"].font = FONT_TITLE
    ws2["A1"].fill = NAVY_HEADER
    ws2["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws2.row_dimensions[1].height = 30

    ws2.merge_cells("A2:V2")
    ws2["A2"] = "Core Calculation Engine (1,176 Combinations — 100% of Tab 3 Specs) | Container parameters (Rows 5-8) are synchronized with Tab 1. Col V provides the Lookup Key."
    ws2["A2"].font = FONT_SUBTITLE
    ws2["A2"].fill = SLATE_SUBHEADER
    ws2["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws2.row_dimensions[2].height = 20

    # Top Control Block on Sheet 2 (Synchronized with Tab 1 inputs!)
    ws2.merge_cells("A4:D4")
    ws2["A4"] = "40GP CONTAINER DIMENSIONS & WEIGHT LIMIT"
    ws2["A4"].font = FONT_PARAM_HDR
    ws2["A4"].fill = PARAM_HDR_FILL
    ws2["A4"].alignment = Alignment(horizontal="center", vertical="center")

    ws2.merge_cells("E4:H4")
    ws2["E4"] = "EFFECTIVE LOADING SPACE & PIPE LENGTH (SYNCED WITH TAB 1)"
    ws2["E4"].font = FONT_PARAM_HDR
    ws2["E4"].fill = PARAM_HDR_FILL
    ws2["E4"].alignment = Alignment(horizontal="center", vertical="center")

    ws2.merge_cells("I4:L4")
    ws2["I4"] = "PACKING METHOD & MATERIAL FACTORS (FROM TAB 3)"
    ws2["I4"].font = FONT_PARAM_HDR
    ws2["I4"].fill = PARAM_HDR_FILL
    ws2["I4"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[4].height = 22

    param_rows_s2 = [
        # Row 5
        [
            ("A5:B5", "40GP Internal Length (m)", PARAM_LBL_FILL, FONT_PARAM_LBL, BORDER_DATA, None),
            ("C5:D5", "='1_Container_Calculator'!D12", CALC_CELL_FILL, FONT_PARAM_VAL, BORDER_DATA, "0.000"),
            ("E5:F5", "Selected Pipe Length (m) [From Tab 1]", PARAM_LBL_FILL, FONT_PARAM_LBL, BORDER_DATA, None),
            ("G5:H5", "='1_Container_Calculator'!D8", CALC_CELL_FILL, FONT_PARAM_VAL, BORDER_DATA, "0.00"),
            ("I5:J5", "Lengthwise Tiers in 40GP", PARAM_LBL_FILL, FONT_PARAM_LBL, BORDER_DATA, None),
            ("K5:L5", "=INT(C5/G5)", CALC_CELL_FILL, FONT_PARAM_VAL, BORDER_DATA, "0"),
        ],
        # Row 6
        [
            ("A6:B6", "40GP Internal Width (mm)", PARAM_LBL_FILL, FONT_PARAM_LBL, BORDER_DATA, None),
            ("C6:D6", 2352, CALC_CELL_FILL, FONT_PARAM_VAL, BORDER_DATA, "#,##0"),
            ("E6:F6", "Effective Loading Width (mm) [From Tab 1]", PARAM_LBL_FILL, FONT_PARAM_LBL, BORDER_DATA, None),
            ("G6:H6", "='1_Container_Calculator'!D13", CALC_CELL_FILL, FONT_PARAM_VAL, BORDER_DATA, "#,##0"),
            ("I6:J6", "Cross-Section Stacking [From Tab 1]", PARAM_LBL_FILL, FONT_PARAM_LBL, BORDER_DATA, None),
            ("K6:L6", "='1_Container_Calculator'!D16", CALC_CELL_FILL, FONT_PARAM_VAL, BORDER_DATA, None),
        ],
        # Row 7
        [
            ("A7:B7", "40GP Internal Height (mm)", PARAM_LBL_FILL, FONT_PARAM_LBL, BORDER_DATA, None),
            ("C7:D7", 2393, CALC_CELL_FILL, FONT_PARAM_VAL, BORDER_DATA, "#,##0"),
            ("E7:F7", "Effective Loading Height (mm) [From Tab 1]", PARAM_LBL_FILL, FONT_PARAM_LBL, BORDER_DATA, None),
            ("G7:H7", "='1_Container_Calculator'!D14", CALC_CELL_FILL, FONT_PARAM_VAL, BORDER_DATA, "#,##0"),
            ("I7:J7", "Carbon Steel Factor C (kg/m)", PARAM_LBL_FILL, FONT_PARAM_LBL, BORDER_DATA, None),
            ("K7:L7", "='3_ASME_Spec_DB'!D6", CALC_CELL_FILL, FONT_PARAM_VAL, BORDER_DATA, "0.0000000"),
        ],
        # Row 8
        [
            ("A8:B8", "Max Container Payload (MT) [From Tab 1]", PARAM_LBL_FILL, FONT_PARAM_LBL, BORDER_DATA, None),
            ("C8:D8", "='1_Container_Calculator'!D15", CALC_CELL_FILL, FONT_PARAM_VAL, BORDER_DATA, "0.00"),
            ("E8:F8", "Max Container Payload (kg)", PARAM_LBL_FILL, FONT_PARAM_LBL, BORDER_DATA, None),
            ("G8:H8", "=C8*1000", CALC_CELL_FILL, FONT_PARAM_VAL, BORDER_DATA, "#,##0"),
            ("I8:J8", "SS304 / SS316 Factor C (kg/m)", PARAM_LBL_FILL, FONT_PARAM_LBL, BORDER_DATA, None),
            ("K8", "='3_ASME_Spec_DB'!D7", CALC_CELL_FILL, FONT_PARAM_VAL, BORDER_DATA, "0.0000000"),
            ("L8", "='3_ASME_Spec_DB'!D8", CALC_CELL_FILL, FONT_PARAM_VAL, BORDER_DATA, "0.0000000"),
        ],
    ]

    for r_idx, row_items in enumerate(param_rows_s2, start=5):
        ws2.row_dimensions[r_idx].height = 20
        for rng, val, fill, font, border, num_fmt in row_items:
            if ":" in rng:
                ws2.merge_cells(rng)
                top_left_coord = rng.split(":")[0]
                cell_coords = rng.split(":")
            else:
                top_left_coord = rng
                cell_coords = [rng]
            cell = ws2[top_left_coord]
            cell.value = val
            cell.font = font
            cell.fill = fill
            is_label = isinstance(val, str) and not val.startswith("=") and top_left_coord[0] in ("A", "E", "I")
            cell.alignment = Alignment(horizontal="left" if is_label else "center", vertical="center", indent=1 if is_label else 0)
            if num_fmt:
                cell.number_format = num_fmt
            for col_let in cell_coords:
                ws2[col_let].border = border
                ws2[col_let].fill = fill

    # Master Table Headers (Row 10, Cols A to V)
    headers_s2 = [
        ("ID", TBL_HDR_MAIN),
        ("Material", TBL_HDR_MAIN),
        ("Standard", TBL_HDR_MAIN),
        ("NPS (in)", TBL_HDR_MAIN),
        ("DN (mm)", TBL_HDR_MAIN),
        ("Schedule", TBL_HDR_MAIN),
        ("OD (in)", TBL_HDR_MAIN),
        ("OD (mm)", TBL_HDR_MAIN),
        ("WT (in)", TBL_HDR_MAIN),
        ("WT (mm)", TBL_HDR_MAIN),
        ("Unit Weight\n(kg/m)", TBL_HDR_CALC),
        ("Piece Weight\n(kg/pc)", TBL_HDR_CALC),
        ("Horiz. Fit\n(pcs/row)", TBL_HDR_CALC),
        ("Vert. Layers\n(rows)", TBL_HDR_CALC),
        ("Max Pieces\nby Volume", TBL_HDR_CALC),
        ("Max Pieces\nby Weight", TBL_HDR_CALC),
        ("40GP Actual\nPieces (pcs)", TBL_HDR_OUT),
        ("40GP Actual\nMetric Tons (MT)", TBL_HDR_OUT),
        ("Payload\nUtil. (%)", TBL_HDR_CALC),
        ("Space Slot\nUtil. (%)", TBL_HDR_CALC),
        ("Limiting Factor\n(Bottleneck)", TBL_HDR_OUT),
        ("Lookup Key\n(Material|NPS|Sch)", TBL_HDR_KEY),
    ]

    hdr_row = 10
    ws2.row_dimensions[hdr_row].height = 30
    for c_idx, (h_text, h_fill) in enumerate(headers_s2, start=1):
        c = ws2.cell(row=hdr_row, column=c_idx, value=h_text)
        c.font = FONT_TBL_HDR
        c.fill = h_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER_HDR

    # Build all valid combinations present in Tab 3 (3_ASME_Spec_DB) so 100% of Tab 3 specs exist in Tab 2 & Tab 1
    master_records = []
    for nps_str, dn_val, od_in, sch_dict in ASME_PIPE_SPECS:
        for sch in SCHEDULES:
            if sch in sch_dict:
                master_records.append(("Carbon Steel", "ASME B36.10M / B36.19M", nps_str, dn_val, sch, od_in, sch_dict[sch]))
    for nps_str, dn_val, od_in, sch_dict in ASME_PIPE_SPECS:
        for sch in SCHEDULES:
            if sch in sch_dict:
                master_records.append(("Stainless Steel 304", "ASME B36.19M / B36.10M", nps_str, dn_val, sch, od_in, sch_dict[sch]))
    for nps_str, dn_val, od_in, sch_dict in ASME_PIPE_SPECS:
        for sch in SCHEDULES:
            if sch in sch_dict:
                master_records.append(("Stainless Steel 316", "ASME B36.19M / B36.10M", nps_str, dn_val, sch, od_in, sch_dict[sch]))

    start_data_row = 11
    for idx, (mat_name, std_name, nps_str, dn_val, sch_code, od_in, wt_in) in enumerate(master_records):
        r = start_data_row + idx
        ws2.row_dimensions[r].height = 19
        row_fill = ZEBRA_FILL if idx % 2 == 1 else WHITE_FILL

        od_mm = round(od_in * INCH_TO_MM, 2)
        wt_mm = round(wt_in * INCH_TO_MM, 2)

        f_unit_wt = f'=ROUND((H{r}-J{r})*J{r}*IF(B{r}="Carbon Steel",$K$7,IF(B{r}="Stainless Steel 304",$K$8,$L$8)), 3)'
        f_pc_wt = f'=ROUND(K{r}*$G$5, 2)'
        f_horiz = f'=INT($G$6/H{r})'
        f_vert = (
            f'=IF(M{r}=0, 0, IF(OR($K$6="Square Grid", M{r}=1), INT($G$7/H{r}), '
            f'IF($G$7<H{r}, 0, 1+INT(($G$7-H{r})/(0.8660254*H{r})))))'
        )
        f_max_vol = (
            f'=IF(OR(M{r}=0,N{r}=0), 0, IF(OR($K$6="Square Grid", M{r}=1), M{r}*N{r}*$K$5, '
            f'(INT((N{r}+1)/2)*M{r} + INT(N{r}/2)*IF(($G$6-M{r}*H{r})>=0.5*H{r}, M{r}, M{r}-1))*$K$5))'
        )
        f_max_wt = f'=IF(L{r}<=0, 0, INT($G$8/L{r}))'
        f_act_pcs = f'=MIN(O{r}, P{r})'
        f_act_mt = f'=ROUND((Q{r}*L{r})/1000, 2)'
        f_wt_util = f'=IF($C$8>0, R{r}/$C$8, 0)'
        f_vol_util = f'=IF(O{r}>0, Q{r}/O{r}, 0)'
        f_limit = f'=IF(Q{r}=0, "Exceeds Limit", IF(P{r}<=O{r}, "Weight Limited", "Volume Limited"))'
        f_key = f'=B{r}&"|"&D{r}&"|"&F{r}'

        row_values = [
            idx + 1, mat_name, std_name, nps_str, dn_val, sch_code,
            od_in, od_mm, wt_in, wt_mm,
            f_unit_wt, f_pc_wt, f_horiz, f_vert, f_max_vol, f_max_wt,
            f_act_pcs, f_act_mt, f_wt_util, f_vol_util, f_limit, f_key
        ]

        for c_idx, val in enumerate(row_values, start=1):
            cell = ws2.cell(row=r, column=c_idx, value=val)
            cell.border = BORDER_DATA
            if c_idx in (17, 18):
                cell.fill = OUTPUT_COL_FILL
                cell.font = FONT_OUTPUT
            elif c_idx == 22:
                cell.fill = row_fill
                cell.font = FONT_KEY
            else:
                cell.fill = row_fill
                if c_idx == 2:
                    cell.font = FONT_CS_TAG if mat_name == "Carbon Steel" else FONT_SS_TAG
                elif c_idx in (4, 6, 21):
                    cell.font = FONT_BOLD
                else:
                    cell.font = FONT_REGULAR

            if c_idx in (1, 3, 4, 5, 6, 13, 14, 21):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in (2, 22):
                cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")

            if c_idx in (1, 5, 13, 14, 15, 16, 17):
                cell.number_format = "#,##0"
            elif c_idx in (7, 9, 11):
                cell.number_format = "0.000"
            elif c_idx in (8, 10, 12, 18):
                cell.number_format = "#,##0.00"
            elif c_idx in (19, 20):
                cell.number_format = "0.0%"

    last_s2_row = start_data_row + len(master_records) - 1

    wt_lim_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    wt_lim_font = Font(name="Calibri", size=10, bold=True, color="B45309")
    vol_lim_fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    vol_lim_font = Font(name="Calibri", size=10, bold=True, color="1D4ED8")

    ws2.conditional_formatting.add(
        f"U11:U{last_s2_row}",
        CellIsRule(operator="equal", formula=['"Weight Limited"'], fill=wt_lim_fill, font=wt_lim_font)
    )
    ws2.conditional_formatting.add(
        f"U11:U{last_s2_row}",
        CellIsRule(operator="equal", formula=['"Volume Limited"'], fill=vol_lim_fill, font=vol_lim_font)
    )

    ws2.auto_filter.ref = f"A10:V{last_s2_row}"
    ws2.freeze_panes = "A11"  # Freeze top 10 header rows only (no vertical line between F and G!)

    col_widths_s2 = {
        1: 7, 2: 20, 3: 15, 4: 11, 5: 10, 6: 11, 7: 10, 8: 11, 9: 10, 10: 11,
        11: 14, 12: 14, 13: 12, 14: 12, 15: 14, 16: 14, 17: 16, 18: 18, 19: 13, 20: 13, 21: 18, 22: 28
    }
    for c_idx, width in col_widths_s2.items():
        ws2.column_dimensions[get_column_letter(c_idx)].width = width

    # =========================================================================
    # 3. Build Tab 1: 1_Container_Calculator (Pure Lookup Dashboard from Tab 2)
    # =========================================================================
    ws1 = wb.create_sheet(title="1_Container_Calculator", index=0)
    ws1.views.sheetView[0].showGridLines = True

    NAVY_DARK = PatternFill(start_color="0F2C59", end_color="0F2C59", fill_type="solid")
    NAVY_SUB = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    SEC_HDR_FILL = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    SUB_HDR_FILL = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    EMERALD_HDR = PatternFill(start_color="047857", end_color="047857", fill_type="solid")
    INDIGO_HDR = PatternFill(start_color="3730A3", end_color="3730A3", fill_type="solid")

    INPUT_FILL = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
    LABEL_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    VALUE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    KPI_GREEN_BG = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    KPI_BLUE_BG = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    KPI_AMBER_BG = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
    INSIGHT_BG = PatternFill(start_color="F0FDFA", end_color="F0FDFA", fill_type="solid")

    FONT_TITLE_1 = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    FONT_SEC_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    FONT_LBL = Font(name="Calibri", size=10, bold=True, color="1E293B")
    FONT_INPUT = Font(name="Calibri", size=11, bold=True, color="B45309")
    FONT_VAL_BOLD = Font(name="Calibri", size=10, bold=True, color="0F2C59")
    FONT_VAL_REG = Font(name="Calibri", size=10, bold=False, color="1E293B")
    FONT_KPI_BIG_GRN = Font(name="Calibri", size=22, bold=True, color="047857")
    FONT_KPI_BIG_BLU = Font(name="Calibri", size=22, bold=True, color="1D4ED8")
    FONT_KPI_BIG_AMB = Font(name="Calibri", size=15, bold=True, color="B45309")
    FONT_KPI_SUB = Font(name="Calibri", size=9, italic=True, color="475569")
    FONT_INSIGHT = Font(name="Calibri", size=10, bold=True, color="0F766E")

    MED_EMERALD = Side(border_style="medium", color="047857")
    MED_AMBER = Side(border_style="medium", color="D97706")
    BORDER_KPI_GRN = Border(left=MED_EMERALD, right=MED_EMERALD, top=MED_EMERALD, bottom=MED_EMERALD)
    BORDER_KPI_BLU = Border(left=MED_NAVY_SIDE, right=MED_NAVY_SIDE, top=MED_NAVY_SIDE, bottom=MED_NAVY_SIDE)
    BORDER_KPI_AMB = Border(left=MED_AMBER, right=MED_AMBER, top=MED_AMBER, bottom=MED_AMBER)

    def style_merged(ws_obj, cell_range, value=None, font=None, fill=None, border=None, alignment=None, num_format=None):
        ws_obj.merge_cells(cell_range)
        start_col, start_row, end_col, end_row = openpyxl.utils.range_boundaries(cell_range)
        top_left = ws_obj.cell(row=start_row, column=start_col)
        if value is not None:
            top_left.value = value
        if font is not None:
            top_left.font = font
        if alignment is not None:
            top_left.alignment = alignment
        if num_format is not None:
            top_left.number_format = num_format
        for r_i in range(start_row, end_row + 1):
            for c_i in range(start_col, end_col + 1):
                cell = ws_obj.cell(row=r_i, column=c_i)
                if fill is not None:
                    cell.fill = fill
                if border is not None:
                    cell.border = border

    # Title Banner on Tab 1
    style_merged(
        ws1, "A1:N1",
        "PRECISE PIPE — 40GP CONTAINER LOADING & COMMERCIAL ESTIMATOR MODEL (V2 LINKED)",
        font=FONT_TITLE_1, fill=NAVY_DARK,
        alignment=Alignment(horizontal="left", vertical="center", indent=1)
    )
    ws1.row_dimensions[1].height = 32

    style_merged(
        ws1, "A2:N2",
        "Front-End Lookup Dashboard | Directly retrieves pre-calculated results from Tab 2 (2_Selectable_Master_Table) based on your dropdown selections (Col D).",
        font=FONT_SUBTITLE, fill=NAVY_SUB,
        alignment=Alignment(horizontal="left", vertical="center", indent=1)
    )
    ws1.row_dimensions[2].height = 20
    ws1.row_dimensions[3].height = 10

    # Left Panel: Section 1 (Pipe Specification Inputs)
    style_merged(
        ws1, "A4:E4",
        "1. PIPE SPECIFICATION INPUTS (SELECT DROPDOWNS)",
        font=FONT_SEC_HDR, fill=SEC_HDR_FILL, border=BORDER_HDR,
        alignment=Alignment(horizontal="left", vertical="center", indent=1)
    )
    ws1.row_dimensions[4].height = 24

    # D9 stores the exact matched row index (1 to 707) in 2_Selectable_Master_Table!
    match_formula_d9 = f'=IFERROR(MATCH(D5&"|"&D6&"|"&D7, \'2_Selectable_Master_Table\'!$V$11:$V${last_s2_row}, 0), "SPEC N/A (SEE TABLE BELOW)")'

    pipe_inputs = [
        (5, "Material Grade [Select]", "Carbon Steel", "@", True),
        (6, "Nominal Pipe Size - NPS [Select]", "6''", "@", True),
        (7, "Pipe Schedule [Select]", "40", "@", True),
        (8, "Single Pipe Length (m) [Select]", 11.80, "0.00", True),
        (9, "Matched ID in Tab 2 (Master Table)", match_formula_d9, "#,##0", False),
    ]

    for r_idx, lbl_text, default_val, num_fmt, is_input in pipe_inputs:
        ws1.row_dimensions[r_idx].height = 22
        style_merged(
            ws1, f"A{r_idx}:C{r_idx}", lbl_text,
            font=FONT_LBL, fill=LABEL_FILL, border=BORDER_DATA,
            alignment=Alignment(horizontal="left", vertical="center", indent=1)
        )
        style_merged(
            ws1, f"D{r_idx}:E{r_idx}", default_val,
            font=FONT_INPUT if is_input else FONT_VAL_BOLD,
            fill=INPUT_FILL if is_input else KPI_BLUE_BG,
            border=BORDER_INPUT if is_input else BORDER_DATA,
            alignment=Alignment(horizontal="center", vertical="center"),
            num_format=num_fmt
        )

    # Left Panel: Section 2 (40GP Container & Commercial Parameters)
    ws1.row_dimensions[10].height = 8
    style_merged(
        ws1, "A11:E11",
        "2. 40GP CONTAINER & COMMERCIAL PARAMETERS",
        font=FONT_SEC_HDR, fill=SEC_HDR_FILL, border=BORDER_HDR,
        alignment=Alignment(horizontal="left", vertical="center", indent=1)
    )
    ws1.row_dimensions[11].height = 24

    container_inputs = [
        (12, "40GP Internal Length (m)", 12.032, "0.000", False),
        (13, "Effective Loading Width (mm) [Input]", 2300, "#,##0", True),
        (14, "Effective Loading Height (mm) [Input]", 2280, "#,##0", True),
        (15, "Max Container Payload (MT) [Input]", 26.00, "0.00", True),
        (16, "Cross-Section Packing Mode [Select]", "Staggered (Hexagonal)", None, True),
        (17, "Est. Ocean Freight / 40GP (USD) [Input]", 3500, "$#,##0", True),
        (18, "Nested Pipe Clearance Margin (mm)", 25, "#,##0", True),
    ]

    for r_idx, lbl_text, default_val, num_fmt, is_input in container_inputs:
        ws1.row_dimensions[r_idx].height = 21
        style_merged(
            ws1, f"A{r_idx}:C{r_idx}", lbl_text,
            font=FONT_LBL, fill=LABEL_FILL, border=BORDER_DATA,
            alignment=Alignment(horizontal="left", vertical="center", indent=1)
        )
        style_merged(
            ws1, f"D{r_idx}:E{r_idx}", default_val,
            font=FONT_INPUT if is_input else FONT_VAL_REG,
            fill=INPUT_FILL if is_input else VALUE_FILL,
            border=BORDER_INPUT if is_input else BORDER_DATA,
            alignment=Alignment(horizontal="center", vertical="center"),
            num_format=num_fmt
        )

    # Right Panel Top: Executive KPI Cards (Directly referencing M16, M17, and Tab 2 Col U!)
    style_merged(
        ws1, "G4:I4", "OUTPUT 1: 40GP TOTAL PIECES",
        font=FONT_SEC_HDR, fill=EMERALD_HDR, border=BORDER_HDR,
        alignment=Alignment(horizontal="center", vertical="center")
    )
    style_merged(
        ws1, "G5:I7", "=M16",
        font=FONT_KPI_BIG_GRN, fill=KPI_GREEN_BG, border=BORDER_KPI_GRN,
        alignment=Alignment(horizontal="center", vertical="center"),
        num_format='#,##0" pcs"'
    )
    style_merged(
        ws1, "G8:I8", '="Total Length: " & TEXT(G5*D8, "#,##0.0") & " m (" & TEXT(G5*D8*3.28084, "#,##0") & " ft)"',
        font=FONT_KPI_SUB, fill=KPI_GREEN_BG, border=BORDER_DATA,
        alignment=Alignment(horizontal="center", vertical="center")
    )

    style_merged(
        ws1, "J4:L4", "OUTPUT 2: 40GP METRIC TONS (MT)",
        font=FONT_SEC_HDR, fill=NAVY_SUB, border=BORDER_HDR,
        alignment=Alignment(horizontal="center", vertical="center")
    )
    style_merged(
        ws1, "J5:L7", "=M17",
        font=FONT_KPI_BIG_BLU, fill=KPI_BLUE_BG, border=BORDER_KPI_BLU,
        alignment=Alignment(horizontal="center", vertical="center"),
        num_format='#,##0.00" MT"'
    )
    style_merged(
        ws1, "J8:L8", '="Net Weight: " & TEXT(J5*1000, "#,##0") & " kg (" & TEXT(J5*2204.62, "#,##0") & " lbs)"',
        font=FONT_KPI_SUB, fill=KPI_BLUE_BG, border=BORDER_DATA,
        alignment=Alignment(horizontal="center", vertical="center")
    )

    style_merged(
        ws1, "M4:N4", "BOTTLENECK & FREIGHT",
        font=FONT_SEC_HDR, fill=INDIGO_HDR, border=BORDER_HDR,
        alignment=Alignment(horizontal="center", vertical="center")
    )
    style_merged(
        ws1, "M5:N6", f'=IF(ISNUMBER(D9), UPPER(INDEX(\'2_Selectable_Master_Table\'!$U$11:$U${last_s2_row}, D9)), "N/A")',
        font=FONT_KPI_BIG_AMB, fill=KPI_AMBER_BG, border=BORDER_KPI_AMB,
        alignment=Alignment(horizontal="center", vertical="center")
    )
    style_merged(
        ws1, "M7:N7", '=IF(J5>0, "Freight: $" & TEXT(D17/J5, "#,##0.00") & " / MT", "Freight: N/A")',
        font=FONT_VAL_BOLD, fill=KPI_AMBER_BG, border=BORDER_DATA,
        alignment=Alignment(horizontal="center", vertical="center")
    )
    style_merged(
        ws1, "M8:N8", '=IF(G5>0, "Unit Freight: $" & TEXT(D17/(G5*D8), "0.00") & " / m", "Unit Freight: N/A")',
        font=FONT_KPI_SUB, fill=KPI_AMBER_BG, border=BORDER_DATA,
        alignment=Alignment(horizontal="center", vertical="center")
    )

    # Right Panel Middle: Pure Lookups from Tab 2 (2_Selectable_Master_Table) using row index D9!
    style_merged(
        ws1, "G10:J10", "3. PIPE SPECS (LOOKED UP FROM TAB 2)",
        font=FONT_SEC_HDR, fill=SUB_HDR_FILL, border=BORDER_HDR,
        alignment=Alignment(horizontal="left", vertical="center", indent=1)
    )
    style_merged(
        ws1, "K10:N10", "4. 40GP PACKING RESULTS (LOOKED UP FROM TAB 2)",
        font=FONT_SEC_HDR, fill=SUB_HDR_FILL, border=BORDER_HDR,
        alignment=Alignment(horizontal="left", vertical="center", indent=1)
    )
    ws1.row_dimensions[10].height = 22

    eng_specs_lookup_rows = [
        (11, "Outer Diameter - OD (mm) [Tab2 Col H]", f'=IF(ISNUMBER(D9), INDEX(\'2_Selectable_Master_Table\'!$H$11:$H${last_s2_row}, D9), 0)', "0.00"),
        (12, "Wall Thickness - WT (mm) [Tab2 Col J]", f'=IF(ISNUMBER(D9), INDEX(\'2_Selectable_Master_Table\'!$J$11:$J${last_s2_row}, D9), "N/A")', "0.00"),
        (13, "Outer / Wall (Inches) [Tab2 Col G,I]", f'=IF(ISNUMBER(D9), TEXT(INDEX(\'2_Selectable_Master_Table\'!$G$11:$G${last_s2_row}, D9),"0.000") & " in OD x " & TEXT(INDEX(\'2_Selectable_Master_Table\'!$I$11:$I${last_s2_row}, D9),"0.000") & " in WT", "N/A")', None),
        (14, "Inner Diameter - ID (mm)", '=IF(ISNUMBER(D9), ROUND(I11 - 2*I12, 2), "N/A")', "0.00"),
        (15, "Material Weight Factor C [Tab2 Row 7/8]", '=IF(D5="Carbon Steel", \'2_Selectable_Master_Table\'!K7, IF(D5="Stainless Steel 304", \'2_Selectable_Master_Table\'!K8, \'2_Selectable_Master_Table\'!L8))', "0.0000000"),
        (16, "Unit Weight (kg/m) [Tab2 Col K]", f'=IF(ISNUMBER(D9), INDEX(\'2_Selectable_Master_Table\'!$K$11:$K${last_s2_row}, D9), 0)', "#,##0.000"),
        (17, "Unit Weight (lb/ft)", '=IF(ISNUMBER(D9), ROUND(I16 * 0.671969, 3), 0)', "#,##0.000"),
        (18, "Piece Weight (kg/pc) [Tab2 Col L]", f'=IF(ISNUMBER(D9), INDEX(\'2_Selectable_Master_Table\'!$L$11:$L${last_s2_row}, D9), 0)', "#,##0.00"),
    ]

    packing_lookup_rows = [
        (11, "Lengthwise Tiers [Tab2 Cell K5]", "='2_Selectable_Master_Table'!K5", "#,##0"),
        (12, "Horiz. Fit (pcs/row) [Tab2 Col M]", f'=IF(ISNUMBER(D9), INDEX(\'2_Selectable_Master_Table\'!$M$11:$M${last_s2_row}, D9), 0)', "#,##0"),
        (13, "Vert. Layers (rows) [Tab2 Col N]", f'=IF(ISNUMBER(D9), INDEX(\'2_Selectable_Master_Table\'!$N$11:$N${last_s2_row}, D9), 0)', "#,##0"),
        (14, "Max Pieces by Volume [Tab2 Col O]", f'=IF(ISNUMBER(D9), INDEX(\'2_Selectable_Master_Table\'!$O$11:$O${last_s2_row}, D9), 0)', "#,##0"),
        (15, "Max Pieces by Weight [Tab2 Col P]", f'=IF(ISNUMBER(D9), INDEX(\'2_Selectable_Master_Table\'!$P$11:$P${last_s2_row}, D9), 0)', "#,##0"),
        (16, "40GP Actual Pieces [Tab2 Col Q]", f'=IF(ISNUMBER(D9), INDEX(\'2_Selectable_Master_Table\'!$Q$11:$Q${last_s2_row}, D9), 0)', "#,##0"),
        (17, "40GP Actual Net Weight [Tab2 Col R]", f'=IF(ISNUMBER(D9), INDEX(\'2_Selectable_Master_Table\'!$R$11:$R${last_s2_row}, D9), 0)', "#,##0.00"),
        (18, "Payload & Space Util. [Tab2 Col S,T]", f'=IF(ISNUMBER(D9), TEXT(INDEX(\'2_Selectable_Master_Table\'!$S$11:$S${last_s2_row}, D9), "0.0%") & " Wt  |  " & TEXT(INDEX(\'2_Selectable_Master_Table\'!$T$11:$T${last_s2_row}, D9), "0.0%") & " Space", "N/A")', None),
    ]

    for (r_idx, lbl_l, form_l, fmt_l), (_, lbl_r, form_r, fmt_r) in zip(eng_specs_lookup_rows, packing_lookup_rows):
        ws1.row_dimensions[r_idx].height = 21
        style_merged(
            ws1, f"G{r_idx}:H{r_idx}", lbl_l,
            font=FONT_LBL, fill=LABEL_FILL, border=BORDER_DATA,
            alignment=Alignment(horizontal="left", vertical="center", indent=1)
        )
        style_merged(
            ws1, f"I{r_idx}:J{r_idx}", form_l,
            font=FONT_VAL_BOLD if r_idx in (11, 12, 16, 18) else FONT_VAL_REG,
            fill=VALUE_FILL, border=BORDER_DATA,
            alignment=Alignment(horizontal="right" if fmt_l else "center", vertical="center"),
            num_format=fmt_l
        )
        style_merged(
            ws1, f"K{r_idx}:L{r_idx}", lbl_r,
            font=FONT_LBL, fill=LABEL_FILL, border=BORDER_DATA,
            alignment=Alignment(horizontal="left", vertical="center", indent=1)
        )
        is_key_out = r_idx in (16, 17)
        style_merged(
            ws1, f"M{r_idx}:N{r_idx}", form_r,
            font=FONT_VAL_BOLD if is_key_out else FONT_VAL_REG,
            fill=KPI_GREEN_BG if is_key_out else VALUE_FILL,
            border=BORDER_DATA,
            alignment=Alignment(horizontal="right" if fmt_r else "center", vertical="center"),
            num_format=fmt_r
        )

    # Row 20: Commercial Advisor Banner
    ws1.row_dimensions[19].height = 8
    ws1.row_dimensions[20].height = 26
    insight_formula = (
        '=IF(NOT(ISNUMBER(D9)), "COMMERCIAL ADVISOR: Schedule " & D7 & " is not applicable to " & D5 & " NPS " & D6 & " in Tab 2. See the live lookup table below for all valid schedules.", '
        'IF(M15<=M14, "COMMERCIAL ADVISOR [WEIGHT LIMITED]: Container reaches Max Payload (" & TEXT(M17,"0.00") & " MT) at " & TEXT(M16/M14,"0.0%") & " space occupancy.", '
        'IF(I14>80, "COMMERCIAL ADVISOR [VOLUME LIMITED — PIPE NESTING OPPORTUNITY]: Space is 100% full at " & TEXT(M17,"0.00") & " MT (" & TEXT(D15-M17,"0.00") & " MT unused payload!). Inner Diameter is " & TEXT(I14,"0.0") & " mm -> Telescope smaller pipes (OD <= " & TEXT(I14-D18,"0") & " mm) inside to ship extra tonnage for FREE ocean freight!", '
        '"COMMERCIAL ADVISOR [VOLUME LIMITED]: Container space is 100% utilized (" & TEXT(M16,"#,##0") & " pcs / " & TEXT(M17,"0.00") & " MT).")))'
    )
    style_merged(
        ws1, "A20:N20", insight_formula,
        font=FONT_INSIGHT, fill=INSIGHT_BG, border=BORDER_KPI_GRN,
        alignment=Alignment(horizontal="left", vertical="center", indent=1)
    )

    # Bottom Section (Rows 22 to 41): All Schedules Lookup Table from Tab 2!
    ws1.row_dimensions[21].height = 8
    style_merged(
        ws1, "A22:N22",
        '="5. ALL SCHEDULES LOOKUP MATRIX FROM TAB 2 FOR: " & D5 & " | NPS " & D6 & " (Pipe Length: " & TEXT(D8,"0.00") & "m)"',
        font=FONT_SEC_HDR, fill=NAVY_DARK, border=BORDER_HDR,
        alignment=Alignment(horizontal="left", vertical="center", indent=1)
    )
    ws1.row_dimensions[22].height = 24

    comp_headers = [
        ("Schedule", NAVY_DARK),
        ("Tab 2 Row ID", NAVY_DARK),
        ("OD (mm)\n[Tab2 Col H]", SUB_HDR_FILL),
        ("WT (in)\n[Tab2 Col I]", SUB_HDR_FILL),
        ("WT (mm)\n[Tab2 Col J]", SUB_HDR_FILL),
        ("ID (mm)", SUB_HDR_FILL),
        ("Unit Wt\n[Tab2 Col K]", SUB_HDR_FILL),
        ("Piece Wt\n[Tab2 Col L]", SUB_HDR_FILL),
        ("Max Vol\n[Tab2 Col O]", SUB_HDR_FILL),
        ("Max Wt\n[Tab2 Col P]", SUB_HDR_FILL),
        ("40GP Pieces\n[Tab2 Col Q]", EMERALD_HDR),
        ("40GP Tons\n[Tab2 Col R]", EMERALD_HDR),
        ("Payload Util\n[Tab2 Col S]", INDIGO_HDR),
        ("Bottleneck\n[Tab2 Col U]", INDIGO_HDR),
    ]

    ws1.row_dimensions[23].height = 28
    for c_i, (h_title, h_fill) in enumerate(comp_headers, start=1):
        c = ws1.cell(row=23, column=c_i, value=h_title)
        c.font = FONT_SEC_HDR
        c.fill = h_fill
        c.border = BORDER_HDR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    start_comp_row = 24
    for idx, sch_code in enumerate(SCHEDULES):
        r = start_comp_row + idx
        ws1.row_dimensions[r].height = 19
        row_fill = ZEBRA_FILL if idx % 2 == 1 else VALUE_FILL

        # Col B (Tab 2 Row ID) finds the exact row in 2_Selectable_Master_Table for ($D$5 | $D$6 | A{r})
        f_row_id = f'=IFERROR(MATCH($D$5&"|"&$D$6&"|"&A{r}, \'2_Selectable_Master_Table\'!$V$11:$V${last_s2_row}, 0), "N/A")'
        # Every other column simply pulls from 2_Selectable_Master_Table using B{r}!
        f_od_mm = f'=IF(ISNUMBER(B{r}), INDEX(\'2_Selectable_Master_Table\'!$H$11:$H${last_s2_row}, B{r}), "-")'
        f_wt_in = f'=IF(ISNUMBER(B{r}), INDEX(\'2_Selectable_Master_Table\'!$I$11:$I${last_s2_row}, B{r}), "-")'
        f_wt_mm = f'=IF(ISNUMBER(B{r}), INDEX(\'2_Selectable_Master_Table\'!$J$11:$J${last_s2_row}, B{r}), "-")'
        f_id_mm = f'=IF(ISNUMBER(B{r}), ROUND(C{r}-2*E{r}, 2), "-")'
        f_unit_wt = f'=IF(ISNUMBER(B{r}), INDEX(\'2_Selectable_Master_Table\'!$K$11:$K${last_s2_row}, B{r}), "-")'
        f_pc_wt = f'=IF(ISNUMBER(B{r}), INDEX(\'2_Selectable_Master_Table\'!$L$11:$L${last_s2_row}, B{r}), "-")'
        f_max_vol = f'=IF(ISNUMBER(B{r}), INDEX(\'2_Selectable_Master_Table\'!$O$11:$O${last_s2_row}, B{r}), "-")'
        f_max_wt = f'=IF(ISNUMBER(B{r}), INDEX(\'2_Selectable_Master_Table\'!$P$11:$P${last_s2_row}, B{r}), "-")'
        f_act_pcs = f'=IF(ISNUMBER(B{r}), INDEX(\'2_Selectable_Master_Table\'!$Q$11:$Q${last_s2_row}, B{r}), "-")'
        f_act_mt = f'=IF(ISNUMBER(B{r}), INDEX(\'2_Selectable_Master_Table\'!$R$11:$R${last_s2_row}, B{r}), "-")'
        f_wt_util = f'=IF(ISNUMBER(B{r}), INDEX(\'2_Selectable_Master_Table\'!$S$11:$S${last_s2_row}, B{r}), "-")'
        f_bottleneck = f'=IF(ISNUMBER(B{r}), INDEX(\'2_Selectable_Master_Table\'!$U$11:$U${last_s2_row}, B{r}), "-")'

        row_vals = [
            sch_code, f_row_id, f_od_mm, f_wt_in, f_wt_mm, f_id_mm,
            f_unit_wt, f_pc_wt, f_max_vol, f_max_wt, f_act_pcs, f_act_mt,
            f_wt_util, f_bottleneck
        ]

        for c_i, val in enumerate(row_vals, start=1):
            cell = ws1.cell(row=r, column=c_i, value=val)
            cell.border = BORDER_DATA
            if c_i in (11, 12):
                cell.fill = KPI_GREEN_BG
                cell.font = FONT_VAL_BOLD
            else:
                cell.fill = row_fill
                cell.font = FONT_VAL_BOLD if c_i in (1, 2, 14) else FONT_VAL_REG

            if c_i in (1, 2, 14):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")

            if c_i == 2:
                cell.number_format = '"Row #"#,##0'
            elif c_i in (3, 5, 6, 8, 12):
                cell.number_format = "#,##0.00"
            elif c_i in (4, 7):
                cell.number_format = "#,##0.000"
            elif c_i in (9, 10, 11):
                cell.number_format = "#,##0"
            elif c_i == 13:
                cell.number_format = "0.0%"

    last_comp_row = start_comp_row + len(SCHEDULES) - 1

    na_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    na_font = Font(name="Calibri", size=9, italic=True, color="94A3B8")
    avail_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    avail_font = Font(name="Calibri", size=10, bold=True, color="15803D")

    ws1.conditional_formatting.add(
        f"B24:B{last_comp_row}",
        CellIsRule(operator="equal", formula=['"N/A"'], fill=na_fill, font=na_font)
    )
    ws1.conditional_formatting.add(
        f"B24:B{last_comp_row}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=avail_fill, font=avail_font)
    )
    ws1.conditional_formatting.add(
        f"N24:N{last_comp_row}",
        CellIsRule(operator="equal", formula=['"Weight Limited"'], fill=wt_lim_fill, font=wt_lim_font)
    )
    ws1.conditional_formatting.add(
        f"N24:N{last_comp_row}",
        CellIsRule(operator="equal", formula=['"Volume Limited"'], fill=vol_lim_fill, font=vol_lim_font)
    )

    # Dropdown Data Validations on Tab 1
    dv_mat = DataValidation(
        type="list",
        formula1='"Carbon Steel,Stainless Steel 304,Stainless Steel 316"',
        allow_blank=False
    )
    ws1.add_data_validation(dv_mat)
    dv_mat.add("D5:E5")

    nps_literal_list = '"' + ",".join([spec[0] for spec in ASME_PIPE_SPECS]) + '"'
    dv_nps = DataValidation(type="list", formula1=nps_literal_list, allow_blank=False)
    ws1.add_data_validation(dv_nps)
    dv_nps.add("D6:E6")

    sch_literal_list = '"' + ",".join(SCHEDULES) + '"'
    dv_sch = DataValidation(type="list", formula1=sch_literal_list, allow_blank=False)
    ws1.add_data_validation(dv_sch)
    dv_sch.add("D7:E7")

    dv_len = DataValidation(type="list", formula1='"5.80,6.00,11.80,12.00"', allow_blank=False)
    ws1.add_data_validation(dv_len)
    dv_len.add("D8:E8")

    dv_pack = DataValidation(type="list", formula1='"Staggered (Hexagonal),Square Grid"', allow_blank=False)
    ws1.add_data_validation(dv_pack)
    dv_pack.add("D16:E16")

    col_widths_s1 = {
        1: 12, 2: 14, 3: 12, 4: 13, 5: 13, 6: 11, 7: 14, 8: 14,
        9: 13, 10: 13, 11: 15, 12: 16, 13: 13, 14: 16
    }
    for c_i, w in col_widths_s1.items():
        ws1.column_dimensions[get_column_letter(c_i)].width = w

    wb.active = 0
    wb.save(filename)

    # Validate formulas
    errors = []
    formula_count = 0
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formula_count += 1
                    if cell.value.count("(") != cell.value.count(")") or cell.value.count('"') % 2 != 0:
                        errors.append((sheet.title, cell.coordinate, cell.value))

    print(f"Successfully generated {filename}!")
    print(f"Validated {formula_count} dynamic Excel formulas across {len(wb.worksheets)} sheets. Syntax errors: {len(errors)}")


if __name__ == "__main__":
    build_v2_linked_model()
