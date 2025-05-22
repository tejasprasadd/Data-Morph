import os
import pdfplumber
import xml.etree.ElementTree as ET
from xml.dom import minidom

# Function to prettify XML output
def prettify(elem):
    """Return a pretty-printed XML string for the Element."""
    rough_string = ET.tostring(elem, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")

# Define folder paths
pdf_folder = "All-Pdfs"
xml_folder = "All-XML"

# Ensure the XML folder exists
if not os.path.exists(xml_folder):
    os.makedirs(xml_folder)
    print(f"Created output directory: {xml_folder}")

# Get all PDF files in the folder
pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]
pdf_files.sort()  # Sort to ensure consistent processing order

if not pdf_files:
    print("No PDF files found in the All-Pdfs folder.")
    exit()

print(f"Found {len(pdf_files)} PDF files to process")

# Process each PDF file
for pdf_index, fname in enumerate(pdf_files):
    print(f"\nProcessing PDF file {pdf_index+1}/{len(pdf_files)}: {fname}")
    
    # Reset data containers for each file
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
    
    try:
        path = os.path.join(pdf_folder, fname)
        pdf = pdfplumber.open(path)
        
        # Assume first page has institute name and intake table
        page0 = pdf.pages[0]
        
        # Extract text and print first few lines to debug
        text_lines = page0.extract_text().splitlines()
        
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
        if tables and len(tables) > 0:
            # The first table on page0 is typically the intake table
            intake_table = tables[0].extract()
            years = intake_table[0][1:]  # e.g. ['2022-23','2021-22',...]
            for row in intake_table[1:]:
                program = row[0]
                for i, year in enumerate(years):
                    if i+1 < len(row):
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
        if tables and len(tables) > 1:
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
        placement_pages = [pdf.pages[0], pdf.pages[1]] if len(pdf.pages) > 1 else [pdf.pages[0]]
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
        if len(pdf.pages) > 1:
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
        if len(pdf.pages) > 2:
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
        if len(pdf.pages) > 3:
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

        # Create XML structure
        root = ET.Element("NIRF_Data")
        
        # Add institute information
        institute = ET.SubElement(root, "Institute")
        ET.SubElement(institute, "Name").text = inst_name
        ET.SubElement(institute, "SourceFile").text = fname
        
        # Add sanctioned intake data
        intake = ET.SubElement(root, "SanctionedIntake")
        for record in intake_records:
            entry = ET.SubElement(intake, "Entry")
            for key, value in record.items():
                if value is not None:  # Skip None values
                    ET.SubElement(entry, key).text = str(value)
        
        # Add student strength data
        strength = ET.SubElement(root, "StudentStrength")
        for record in strength_records:
            entry = ET.SubElement(strength, "Entry")
            for key, value in record.items():
                if value is not None:  # Skip None values
                    ET.SubElement(entry, key).text = str(value)
        
        # Add placement data
        placement = ET.SubElement(root, "PlacementData")
        for record in placement_records:
            entry = ET.SubElement(placement, "Entry")
            for key, value in record.items():
                if value is not None:  # Skip None values
                    ET.SubElement(entry, key).text = str(value)
        
        # Add PhD data
        phd = ET.SubElement(root, "PhDData")
        for record in phd_records:
            entry = ET.SubElement(phd, "Entry")
            for key, value in record.items():
                if value is not None:  # Skip None values
                    ET.SubElement(entry, key).text = str(value)
        
        # Add financial data - capital expenditure
        finance_cap = ET.SubElement(root, "CapitalExpenditure")
        for record in finance_capital:
            entry = ET.SubElement(finance_cap, "Entry")
            for key, value in record.items():
                if value is not None:  # Skip None values
                    ET.SubElement(entry, key).text = str(value)
        
        # Add financial data - operational expenditure
        finance_op = ET.SubElement(root, "OperationalExpenditure")
        for record in finance_operational:
            entry = ET.SubElement(finance_op, "Entry")
            for key, value in record.items():
                if value is not None:  # Skip None values
                    ET.SubElement(entry, key).text = str(value)
        
        # Add sponsored projects data
        sponsored = ET.SubElement(root, "SponsoredProjects")
        for record in sponsored_records:
            entry = ET.SubElement(sponsored, "Entry")
            for key, value in record.items():
                if value is not None:  # Skip None values
                    ET.SubElement(entry, key).text = str(value)
        
        # Add consultancy projects data
        consultancy = ET.SubElement(root, "ConsultancyProjects")
        for record in consultancy_records:
            entry = ET.SubElement(consultancy, "Entry")
            for key, value in record.items():
                if value is not None:  # Skip None values
                    ET.SubElement(entry, key).text = str(value)
        
        # Add facilities data
        facilities = ET.SubElement(root, "Facilities")
        for record in facilities_records:
            entry = ET.SubElement(facilities, "Entry")
            for key, value in record.items():
                if value is not None:  # Skip None values
                    ET.SubElement(entry, key).text = str(value)
        
        # Add faculty data
        faculty = ET.SubElement(root, "FacultyCount")
        for record in faculty_records:
            entry = ET.SubElement(faculty, "Entry")
            for key, value in record.items():
                if value is not None:  # Skip None values
                    ET.SubElement(entry, key).text = str(value)
        
        # Create a new XML file name based on the PDF file name (without extension)
        base_name = os.path.splitext(fname)[0]
        xml_filename = os.path.join(xml_folder, f"{base_name}.xml")
        
        # Write to XML file with pretty formatting
        with open(xml_filename, "w", encoding="utf-8") as f:
            f.write(prettify(root))
        
        print(f"XML file created: {xml_filename}")
        
    except Exception as e:
        print(f"Error processing {fname}: {str(e)}")

# print("\nAll PDF files have been processed and saved to the All-XML folder.")

# # Comparison of PDF-to-XML vs Excel-to-XML conversion approaches
# print("\nComparison of PDF-to-XML vs Excel-to-XML conversion approaches:")
# print("\nPDF-to-XML Benefits:")
# print("1. Direct conversion eliminates intermediate steps, reducing potential for data loss")
# print("2. Maintains original data structure from the source document")
# print("3. Avoids Excel-specific formatting or calculation issues")
# print("4. More efficient for batch processing of multiple documents")
# print("5. Better for preserving text-based content like descriptions and notes")

# print("\nExcel-to-XML Benefits:")
# print("1. Excel data is already structured in rows and columns, simplifying conversion")
# print("2. Excel files may have already undergone data cleaning and normalization")
# print("3. Better for data that has been manually verified or enhanced")
# print("4. Easier to handle complex calculations or derived values")
# print("5. May preserve formatting and styling information better")

# print("\nRecommendation:")
# print("For NIRF data extraction, PDF-to-XML is generally preferable because:")
# print("- It preserves the original data structure directly from the source")
# print("- Eliminates potential data loss or transformation errors in the Excel intermediate step")
# print("- Creates a more standardized output format for downstream processing")
# print("- XML provides better semantic structure for hierarchical NIRF data")
# print("- More suitable for automation in data processing pipelines")