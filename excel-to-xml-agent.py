import os
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom

# Directory containing the Excel files
EXCEL_DIR = 'All-Excels'

def prettify(elem):
    """Return a pretty-printed XML string for the Element."""
    rough_string = ET.tostring(elem, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")

def get_sorted_excel_files(directory):
    files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
    # Sort by the numeric prefix
    files.sort(key=lambda x: int(x.split('-')[0]))
    return files

def extract_sheet_name(filename):
    # Remove the numeric prefix and dash
    match = filename.split('-', 1)
    if len(match) > 1:
        return match[1].replace('.xlsx', '')
    return filename.replace('.xlsx', '')

def convert_excel_to_xml(excel_file):
    print(f"Processing Excel file: {excel_file}")
    file_path = os.path.join(EXCEL_DIR, excel_file)
    
    # Extract institute name from filename
    institute_name = extract_sheet_name(excel_file)
    
    # Create root XML element
    root = ET.Element("NIRF_Data")
    
    # Add institute information
    institute = ET.SubElement(root, "Institute")
    ET.SubElement(institute, "Name").text = institute_name
    ET.SubElement(institute, "SourceFile").text = excel_file
    
    # Read Excel file with all sheets
    excel = pd.ExcelFile(file_path)
    
    # Process each sheet in the Excel file
    for sheet_name in excel.sheet_names:
        # Read the sheet into a DataFrame
        df = pd.read_excel(excel, sheet_name=sheet_name)
        
        # Skip empty sheets
        if df.empty:
            continue
        
        # Create a section in XML for this sheet
        section = ET.SubElement(root, sheet_name)
        
        # Convert each row to an XML entry
        for _, row in df.iterrows():
            entry = ET.SubElement(section, "Entry")
            
            # Add each column as an element
            for column, value in row.items():
                # Skip NaN values
                if pd.notna(value):
                    # Convert to appropriate string representation
                    if isinstance(value, (int, float)):
                        value_str = str(int(value) if value.is_integer() else value)
                    else:
                        value_str = str(value)
                    
                    # Create element with column name and value
                    ET.SubElement(entry, str(column)).text = value_str
    
    # Create XML filename
    xml_filename = f"NIRF_Data_{excel_file.replace('.xlsx', '')}.xml"
    
    # Write to XML file with pretty formatting
    with open(xml_filename, "w", encoding="utf-8") as f:
        f.write(prettify(root))
    
    print(f"XML file created: {xml_filename}")
    return xml_filename

def main():
    print("Starting Excel to XML conversion...")
    
    # Get list of Excel files
    excel_files = get_sorted_excel_files(EXCEL_DIR)
    
    if not excel_files:
        print("No Excel files found in the All-Excels directory.")
        return
    
    print(f"Found {len(excel_files)} Excel files")
    
    # Process the first Excel file
    first_file = excel_files[0]
    xml_file = convert_excel_to_xml(first_file)
    
    print("\nComparison of PDF-to-XML vs Excel-to-XML conversion approaches:")
    print("\nPDF-to-XML Benefits:")
    print("1. Direct conversion eliminates intermediate steps, reducing potential for data loss")
    print("2. Maintains original data structure from the source document")
    print("3. Avoids Excel-specific formatting or calculation issues")
    print("4. More efficient for batch processing of multiple documents")
    print("5. Better for preserving text-based content like descriptions and notes")

    print("\nExcel-to-XML Benefits:")
    print("1. Excel data is already structured in rows and columns, simplifying conversion")
    print("2. Excel files may have already undergone data cleaning and normalization")
    print("3. Better for data that has been manually verified or enhanced")
    print("4. Easier to handle complex calculations or derived values")
    print("5. May preserve formatting and styling information better")

    print("\nRecommendation:")
    print("For NIRF data extraction, the best approach depends on your specific needs:")
    print("- Use PDF-to-XML when working directly with source documents and prioritizing original structure")
    print("- Use Excel-to-XML when working with already processed data that has been cleaned or enhanced")
    print("- Consider your data pipeline requirements and which format better serves downstream processing")

if __name__ == "__main__":
    main()