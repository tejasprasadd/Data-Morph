import os
import pdfplumber
import pandas as pd

# Prepare containers for each category of data
intake_records = []
strength_records = []
placement_records = []
phd_records = []
finance_capital = []
finance_operational = []
sponsored_records = []
consultancy_records = []
facilities_records = []
faculty_records = []
# Change the folder path to All-Pdfs
pdf_folder = "All-Pdfs"
# Get the first PDF file in the folder
pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]
pdf_files.sort()  # Sort to ensure we get the first one (001-...)

if pdf_files:
    fname = pdf_files[0]  # Get the first PDF file
    print(f"Processing PDF file: {fname}")
    path = os.path.join(pdf_folder, fname)
    pdf = pdfplumber.open(path)
    # Assume first page has institute name and intake table
    page0 = pdf.pages[0]
    
    # Extract text and print first few lines to debug
    text_lines = page0.extract_text().splitlines()
    for i, line in enumerate(text_lines[:10]):
        print(f"Line {i}: {line}")
    
    # More robust institute name extraction
    inst_name = "Unknown Institute"
    for line in text_lines[:15]:  # Check first 15 lines
        if "Institute Name:" in line:
            inst_name = line.split("Institute Name:")[1].strip()
            break
        # Alternative patterns that might appear in the PDF
        elif "Name of Institution:" in line:
            inst_name = line.split("Name of Institution:")[1].strip()
            break
        elif "Institution:" in line:
            inst_name = line.split("Institution:")[1].strip()
            break
    
    print(f"Extracted institute name: {inst_name}")
    
    # --- Sanctioned Intake --- 
    tables = page0.find_tables()
    # The first table on page0 is typically the intake table
    intake_table = tables[0].extract()
    years = intake_table[0][1:]  # e.g. ['2022-23','2021-22',...]
    for row in intake_table[1:]:
        program = row[0]
        for i, year in enumerate(years):
            intake = row[i+1]
            # Skip if data is missing or '-'
            if intake is None or intake.strip() == '-':
                continue
            intake_records.append({
                "Institute": inst_name,
                "Program": program,
                "Year": year,
                "ApprovedIntake": int(intake)
            })
    
    # --- Student Strength / Demographics ---
    # The second table on page0 is student strength breakdown
    student_table = tables[1].extract()
    # The header spans multiple lines, but columns align with keys:
    columns = ["Male", "Female", "Total", "WithinState", "OutsideState", 
               "Abroad", "EconomicallyBackward", "SociallyChallenged", 
               "FeeReimb_State", "FeeReimb_Inst", "FeeReimb_Private", 
               "NoReimbursement"]
    for row in student_table[1:]:
        program = row[0]
        values = row[1:]
        rec = {"Institute": inst_name, "Program": program}
        for col, val in zip(columns, values):
            # Improved handling of non-numeric values
            if val is not None and val.strip() and val.strip() != '-' and val.strip().isdigit():
                rec[col] = int(val.strip())
            else:
                rec[col] = None
        strength_records.append(rec)
    
    # --- Placement & Higher Studies (UG/PG) ---
    # Page0 table2 is UG 4-year placement; page1 tables for UG5Y, PG2Y, PG3Y
    placement_pages = [pdf.pages[0], pdf.pages[1]]
    # Identify and parse each placement table by checking header text
    for page in placement_pages:
        tbls = page.find_tables()
        for tbl in tbls:
            data = tbl.extract()
            if not data or len(data) < 1:  # Check if data exists
                continue
                
            header = data[0]
            if len(header) > 0 and header[0].startswith("Academic Year"):
                # Determine program type by context in the PDF
                # For example, UG-4Y has caption on page0, UG-5Y on page1, etc.
                # Here we guess by number of intake columns or existing data.
                # For brevity, assume orders: UG4 (page0), UG5 (page1 first), PG2 (page1 second), PG3 (page1 third).
                # Use loop index or content to assign program.
                # Example for UG5:
                for row in data[1:]:
                    if not row:  # Skip empty rows
                        continue
                        
                    rec = {"Institute": inst_name}
                    rec["AcademicYear"] = row[0] if len(row) > 0 else None
                    
                    # Safely access columns with length checks
                    rec["FirstYearIntake"] = int(row[1]) if len(row) > 1 and row[1] and row[1].strip() != '-' and row[1].strip().isdigit() else None
                    rec["FirstYearAdmitted"] = int(row[2]) if len(row) > 2 and row[2] and row[2].strip() != '-' and row[2].strip().isdigit() else None
                    rec["GraduatingYear"] = row[5] if len(row) > 5 else None  # 2nd Academic Year in row
                    rec["GraduatingStudents"] = int(row[6]) if len(row) > 6 and row[6] and row[6].strip() != '-' and row[6].strip().isdigit() else None
                    rec["Placed"] = int(row[7]) if len(row) > 7 and row[7] and row[7].strip() != '-' and row[7].strip().isdigit() else None
                    
                    # Handle median salary which might have text in parentheses
                    if len(row) > 8 and row[8] and row[8].strip() != '-':
                        salary_parts = row[8].split("(")
                        if salary_parts[0].strip().isdigit():
                            rec["MedianSalary"] = int(salary_parts[0].strip())
                        else:
                            rec["MedianSalary"] = None
                    else:
                        rec["MedianSalary"] = None
                        
                    rec["HigherStudies"] = int(row[9]) if len(row) > 9 and row[9] and row[9].strip() != '-' and row[9].strip().isdigit() else None
                    placement_records.append(rec)
                # (Repeat similarly for PG2Y, PG3Y using their specific table indices)
    
    # --- Ph.D. Data ---
    # Page1 contains Ph.D student counts and graduations
    page1 = pdf.pages[1]
    phd_tables = [tbl.extract() for tbl in page1.find_tables() if tbl.extract()[0][0].startswith("Ph.D")]
    if phd_tables:  # Make sure we found a Ph.D table
        phd_table = phd_tables[0]
        # Parse total students - with error handling
        if len(phd_table) > 3 and len(phd_table[2]) > 2 and phd_table[2][2] and phd_table[2][2].strip().isdigit():
            total_full = int(phd_table[2][2])
            phd_records.append({"Institute": inst_name, "Type": "FullTime_Total", "Count": total_full})
        
        if len(phd_table) > 3 and len(phd_table[3]) > 2 and phd_table[3][2] and phd_table[3][2].strip().isdigit():
            total_part = int(phd_table[3][2])
            phd_records.append({"Institute": inst_name, "Type": "PartTime_Total", "Count": total_part})
        
        # Parse graduates per year - with error handling
        if len(phd_table) > 5:
            year_header = phd_table[5]  # e.g., ['', '2022-23', '2021-22', '2020-21']
            for row in phd_table[6:]:
                if len(row) > 0:
                    mode = row[0]  # 'Full Time' or 'Part Time'
                    for j, year in enumerate(year_header[1:], start=1):
                        if j < len(row) and row[j] and row[j].strip() != '-':
                            try:
                                count = int(row[j])
                                phd_records.append({"Institute": inst_name, "Type": mode, "Year": year, "Graduated": count})
                            except ValueError:
                                # Skip if conversion fails
                                pass
    
    # --- Financial Resources ---
    # Page2 tables: capital, operational, sponsored, consultancy
    page2 = pdf.pages[2]
    for tbl in page2.find_tables():
        data = tbl.extract()
        if len(data) > 2 and data[0][0].startswith("Financial Year") and data[2][0].startswith("Annual Capital"):
            # Capital expenditures
            categories = [row[0] for row in data[3:]]
            amounts = data[3:]
            years = data[0][1:]
            for i, cat in enumerate(categories):
                for col_idx, year in enumerate(years, start=1):
                    if 3+i < len(data) and col_idx < len(data[3+i]):
                        amt = data[3+i][col_idx]
                        if amt and amt.strip() != '-' and amt.strip().isdigit():
                            finance_capital.append({
                                "Institute": inst_name,
                                "Category": cat.strip(),
                                "Year": year,
                                "Amount": int(amt)
                            })
        # Operational similar logic
        if len(data) > 2 and data[0][0].startswith("Financial Year") and data[2][0].startswith("Annual Operational"):
            cats = [row[0] for row in data[3:]]
            years = data[0][1:]
            for i, cat in enumerate(cats):
                for col_idx, year in enumerate(years, start=1):
                    if 3+i < len(data) and col_idx < len(data[3+i]):
                        amt = data[3+i][col_idx]
                        if amt and amt.strip() != '-' and amt.strip().isdigit():
                            finance_operational.append({
                                "Institute": inst_name,
                                "Category": cat.strip(),
                                "Year": year,
                                "Amount": int(amt)
                            })
        # Sponsored Projects
        if len(data) > 1 and data[0][0] == "Financial Year" and "Sponsored Projects" in data[1][0]:
            years = data[0][1:]
            for row in data[1:4]:
                if len(row) > 0:
                    key = row[0]
                    for i, year in enumerate(years, start=1):
                        if i < len(row):
                            val = row[i]
                            if val and val.strip() != '-' and val.strip().isdigit():
                                rec = {"Institute": inst_name, "Type": key, "Year": year, "Value": int(val.strip())}
                                sponsored_records.append(rec)
                            else:
                                rec = {"Institute": inst_name, "Type": key, "Year": year, "Value": None}
                                sponsored_records.append(rec)
        # Consultancy Projects
        if len(data) > 1 and data[0][0] == "Financial Year" and "Consultancy Projects" in data[1][0]:
            years = data[0][1:]
            for row in data[1:4]:
                if len(row) > 0:
                    key = row[0]
                    for i, year in enumerate(years, start=1):
                        if i < len(row):
                            val = row[i]
                            if val and val.strip() != '-' and val.strip().isdigit():
                                consultancy_records.append({
                                    "Institute": inst_name, "Type": key, "Year": year, "Value": int(val.strip())
                                })
                            else:
                                consultancy_records.append({
                                    "Institute": inst_name, "Type": key, "Year": year, "Value": None
                                })
    
    # --- Facilities for Physically Challenged ---
    # Page3 may contain Q&A table (here simplified parse)
    page3 = pdf.pages[3]
    for tbl in page3.find_tables():
        cells = tbl.extract()
        # e.g. if first row is a question about lifts/ramps:
        if len(cells) > 0 and len(cells[0]) > 0 and cells[0][0].startswith("1. Do your institution buildings"):
            if len(cells[0]) > 1:
                has_lifts = cells[0][1]
                facilities_records.append({
                    "Institute": inst_name,
                    "Feature": "Lifts/Ramps in Buildings",
                    "Available": has_lifts
                })
        if len(cells) > 0 and len(cells[0]) > 2 and cells[0][2].startswith("2. Do you offer any separate cell"):
            if len(cells[0]) > 3:
                has_cell = cells[0][3]
                facilities_records.append({
                    "Institute": inst_name,
                    "Feature": "Special Facilities for Challenged",
                    "Available": has_cell
                })
    # --- Faculty Details ---
    # Last table on page3: faculty count
    for tbl in page3.find_tables():
        cells = tbl.extract()
        if len(cells) > 0 and len(cells[0]) > 0 and cells[0][0].startswith("Number of faculty"):
            if len(cells[0]) > 1 and cells[0][1] and cells[0][1].strip().isdigit():
                num_faculty = int(cells[0][1])
                faculty_records.append({"Institute": inst_name, "TotalFaculty": num_faculty})
    
    pdf.close()

    # Convert lists to DataFrames
    df_intake = pd.DataFrame(intake_records)
    df_strength = pd.DataFrame(strength_records)
    df_placement = pd.DataFrame(placement_records)
    df_phd = pd.DataFrame(phd_records)
    df_fin_cap = pd.DataFrame(finance_capital)
    df_fin_op = pd.DataFrame(finance_operational)
    df_sp = pd.DataFrame(sponsored_records)
    df_consult = pd.DataFrame(consultancy_records)
    df_facil = pd.DataFrame(facilities_records)
    df_fac = pd.DataFrame(faculty_records)

    # Create a new Excel file name based on the PDF file name
    excel_filename = f"NIRF_Data_{fname.replace('.pdf', '')}.xlsx"
    
    # Write to Excel with separate sheets
    with pd.ExcelWriter(excel_filename, engine="openpyxl") as writer:
        df_intake.to_excel(writer, sheet_name="SanctionedIntake", index=False)
        df_strength.to_excel(writer, sheet_name="StudentStrength", index=False)
        df_placement.to_excel(writer, sheet_name="PlacementData", index=False)
        df_phd.to_excel(writer, sheet_name="PhDData", index=False)
        df_fin_cap.to_excel(writer, sheet_name="CapitalExpenditure", index=False)
        df_fin_op.to_excel(writer, sheet_name="OpExpenditure", index=False)
        df_sp.to_excel(writer, sheet_name="SponsoredProjects", index=False)
        df_consult.to_excel(writer, sheet_name="ConsultancyProjects", index=False)
        df_facil.to_excel(writer, sheet_name="Facilities", index=False)
        df_fac.to_excel(writer, sheet_name="FacultyCount", index=False)
    
    print(f"Excel file created: {excel_filename}")
else:
    print("No PDF files found in the All-Pdfs folder.")
