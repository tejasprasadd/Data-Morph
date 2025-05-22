import xml.etree.ElementTree as ET
import pymongo
import sys
import os
import glob
from datetime import datetime
from tqdm import tqdm


def parse_xml_to_dict(xml_file):
    """
    Parse XML file and convert it to a Python dictionary with standardized structure
    """
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        
        # Create main document
        nirf_data = {}
        
        # Extract Institute information
        institute_elem = root.find('Institute')
        if institute_elem is not None:
            nirf_data['institute'] = {
                'name': institute_elem.findtext('Name'),
                'source_file': institute_elem.findtext('SourceFile')
            }
            
            # Extract college ID from name if available
            name = institute_elem.findtext('Name')
            if name:
                # Try to extract ID from format like "College Name [ID]"
                if '[' in name and ']' in name:
                    college_id = name.split('[')[1].split(']')[0]
                    nirf_data['institute']['college_id'] = college_id
                # Try to extract ID from filename if available
                elif os.path.basename(xml_file).startswith(tuple([str(i).zfill(3) for i in range(1, 101)])):
                    file_id = os.path.basename(xml_file).split('-')[0]
                    nirf_data['institute']['college_id'] = file_id
        
        # Process each section with standardized field names
        sections = ['SanctionedIntake', 'StudentStrength', 'PlacementData', 
                   'PhDData', 'CapitalExpenditure', 'OperationalExpenditure', 
                   'SponsoredProjects', 'ConsultancyProjects', 'Facilities', 'FacultyCount']
        
        # Define field name mappings for standardization
        field_mappings = {
            # Common misspellings or variations
            'acadmic_year': 'academic_year',
            'academicyear': 'academic_year',
            'acad_year': 'academic_year',
            'year': 'academic_year',
            'program': 'program_name',
            'programname': 'program_name',
            'prog': 'program_name',
            'dept': 'department',
            'department': 'department',
            'male': 'male_count',
            'female': 'female_count',
            'total': 'total_count',
            'approved_intake': 'approved_intake',
            'approvedintake': 'approved_intake',
            'sanctioned_intake': 'approved_intake',
            'median_salary': 'median_salary',
            'mediansalary': 'median_salary',
            'salary': 'median_salary'
        }
        
        for section in sections:
            section_elem = root.find(section)
            if section_elem is not None:
                entries = []
                for entry in section_elem.findall('Entry'):
                    entry_data = {}
                    for child in entry:
                        # Standardize field name
                        field_name = child.tag.lower().strip()
                        # Apply mapping if available
                        field_name = field_mappings.get(field_name, field_name)
                        
                        # Try to convert numeric values
                        try:
                            # Handle different number formats
                            if child.text:
                                text = child.text.strip()
                                # Remove commas, currency symbols, etc.
                                cleaned_text = text.replace(',', '').replace('₹', '').replace('$', '').strip()
                                
                                # Check if it's a number
                                if cleaned_text and cleaned_text.replace('-', '').replace('.', '').isdigit():
                                    # Convert to int or float as appropriate
                                    if '.' in cleaned_text:
                                        entry_data[field_name] = float(cleaned_text)
                                    else:
                                        entry_data[field_name] = int(cleaned_text)
                                else:
                                    entry_data[field_name] = text
                            else:
                                entry_data[field_name] = None
                        except (ValueError, TypeError):
                            entry_data[field_name] = child.text
                    
                    # Only add entries with actual data
                    if entry_data:
                        entries.append(entry_data)
                
                # Use standardized section names (lowercase)
                section_key = section.lower()
                nirf_data[section_key] = entries
        
        return nirf_data
    
    except ET.ParseError as e:
        print(f"Error parsing XML file: {e}")
        return None
    except Exception as e:
        print(f"Unexpected error: {e}")
        return None


def connect_to_mongodb(host='localhost', port=27017, db_name='nirf_database'):
    """
    Connect to MongoDB and return database object
    """
    try:
        # Use MongoDB connection string format
        connection_string = f"mongodb://{host}:{port}/"
        client = pymongo.MongoClient(connection_string)
        # Test connection
        client.admin.command('ping')
        print(f"Connected successfully to MongoDB at {connection_string}")
        
        # Get or create database
        db = client[db_name]
        return db, client
    
    except pymongo.errors.ConnectionFailure as e:
        print(f"Could not connect to MongoDB: {e}")
        return None, None
    except Exception as e:
        print(f"Unexpected error: {e}")
        return None, None


def insert_individual_data(db, data, collection_name='individuals'):
    """
    Insert individual college data into MongoDB collection
    """
    try:
        # Get or create collection
        collection = db[collection_name]
        
        # Add metadata
        data['metadata'] = {
            'imported_at': datetime.now(),
            'source_file': data.get('institute', {}).get('source_file', 'unknown')
        }
        
        # Extract college ID for document ID
        college_id = data.get('institute', {}).get('college_id')
        if college_id:
            # Use college_id as part of the document ID for better organization
            data['_id'] = f"individual_{college_id}_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        
        # Insert document
        result = collection.insert_one(data)
        return result.inserted_id
    
    except pymongo.errors.PyMongoError as e:
        print(f"MongoDB error: {e}")
        return None
    except Exception as e:
        print(f"Unexpected error: {e}")
        return None


def create_master_database(db, source_collection='individuals', master_collection='master_database'):
    """
    Create a structured master database by organizing data by college with proper headers
    """
    try:
        # Get source collection
        source = db[source_collection]
        
        # Create or get master collection
        master = db[master_collection]
        
        # Create index on college_id for faster lookups
        master.create_index('college_id', unique=True)
        
        # Get all documents from source collection
        all_docs = list(source.find({}))
        print(f"Creating structured master database from {len(all_docs)} documents...")
        
        # Process each document
        for doc in tqdm(all_docs, desc="Processing colleges"):
            college_id = doc.get('institute', {}).get('college_id')
            college_name = doc.get('institute', {}).get('name')
            
            if not college_id or not college_name:
                # Generate a unique ID if college_id is missing
                if not college_id and college_name:
                    college_id = f"unnamed_{hash(college_name) % 10000}"
                    print(f"Generated ID {college_id} for college: {college_name}")
                else:
                    print(f"Warning: Missing college ID and name in document {doc.get('_id')}")
                    continue
            
            # Check if college already exists in master collection
            existing = master.find_one({'college_id': college_id})
            
            if existing:
                # Update existing college document
                update_data = {}
                
                # Update each section with proper structure
                for section in doc:
                    if section not in ['_id', 'metadata', 'institute']:
                        # Ensure consistent header names and structure
                        structured_data = []
                        for entry in doc[section]:
                            # Clean and standardize field names
                            cleaned_entry = {}
                            for key, value in entry.items():
                                # Convert keys to standard format (camelCase)
                                clean_key = key.strip()
                                # Keep the value as is
                                cleaned_entry[clean_key] = value
                            structured_data.append(cleaned_entry)
                        
                        update_data[section] = structured_data
                
                # Update metadata
                update_data['metadata.last_updated'] = datetime.now()
                update_data['metadata.source_files'] = existing.get('metadata', {}).get('source_files', [])
                source_file = doc.get('metadata', {}).get('source_file')
                if source_file and source_file not in update_data['metadata.source_files']:
                    update_data['metadata.source_files'].append(source_file)
                
                # Update the document
                master.update_one(
                    {'college_id': college_id},
                    {'$set': update_data}
                )
            else:
                # Create new college document with structured data
                master_doc = {
                    'college_id': college_id,
                    'college_name': college_name,
                    'metadata': {
                        'created_at': datetime.now(),
                        'last_updated': datetime.now(),
                        'source_files': [doc.get('metadata', {}).get('source_file', 'unknown')]
                    }
                }
                
                # Add each section with proper structure
                for section in doc:
                    if section not in ['_id', 'metadata', 'institute']:
                        # Ensure consistent header names and structure
                        structured_data = []
                        for entry in doc[section]:
                            # Clean and standardize field names
                            cleaned_entry = {}
                            for key, value in entry.items():
                                # Convert keys to standard format
                                clean_key = key.strip()
                                # Keep the value as is
                                cleaned_entry[clean_key] = value
                            structured_data.append(cleaned_entry)
                        
                        master_doc[section] = structured_data
                
                # Insert the document
                master.insert_one(master_doc)
        
        print(f"Master database created successfully with {master.count_documents({})} colleges")
        return True
    
    except pymongo.errors.PyMongoError as e:
        print(f"MongoDB error: {e}")
        return False
    except Exception as e:
        print(f"Unexpected error: {e}")
        return False


def process_all_xml_files(xml_dir, host='localhost', port=27017, db_name='nirf_database', 
                         individual_collection='individuals', master_collection='master_database'):
    """
    Process all XML files in a directory and create a master database
    """
    # Get all XML files in the directory
    xml_files = glob.glob(os.path.join(xml_dir, '*.xml'))
    
    if not xml_files:
        print(f"No XML files found in directory: {xml_dir}")
        return False
    
    print(f"Found {len(xml_files)} XML files to process")
    
    # Connect to MongoDB
    print(f"Connecting to MongoDB at {host}:{port}")
    db, client = connect_to_mongodb(host, port, db_name)
    if db is None:
        print("Failed to connect to MongoDB.")
        return False
    
    try:
        # Process each XML file
        successful_imports = 0
        failed_imports = 0
        
        for xml_file in tqdm(xml_files, desc="Processing XML files"):
            # Parse XML to dictionary
            data = parse_xml_to_dict(xml_file)
            if data is None:
                print(f"Failed to parse XML file: {xml_file}")
                failed_imports += 1
                continue
            
            # Insert data to MongoDB individuals collection
            result = insert_individual_data(db, data, individual_collection)
            if result:
                successful_imports += 1
            else:
                failed_imports += 1
        
        print(f"XML import completed. Successful: {successful_imports}, Failed: {failed_imports}")
        
        # Create master database
        if successful_imports > 0:
            create_master_database(db, individual_collection, master_collection)
        
        return True
    
    finally:
        # Close MongoDB connection
        if client:
            client.close()
            print("MongoDB connection closed.")


def main():
    # Check command line arguments
    if len(sys.argv) > 1 and sys.argv[1] == '--help':
        print("Usage: python xml_to_mongodb.py [xml_dir] [host] [port] [db_name] [individual_collection] [master_collection]")
        print("Example: python xml_to_mongodb.py ./All-XML localhost 27017 nirf_database individuals master_database")
        return
    
    # Get parameters
    xml_dir = sys.argv[1] if len(sys.argv) > 1 else './All-XML'
    host = sys.argv[2] if len(sys.argv) > 2 else 'localhost'
    port = int(sys.argv[3]) if len(sys.argv) > 3 else 27017  # Default MongoDB port
    db_name = sys.argv[4] if len(sys.argv) > 4 else 'nirf_database'
    individual_collection = sys.argv[5] if len(sys.argv) > 5 else 'individuals'
    master_collection = sys.argv[6] if len(sys.argv) > 6 else 'master_database'
    
    # Check if directory exists
    if not os.path.isdir(xml_dir):
        print(f"Error: Directory '{xml_dir}' does not exist.")
        return
    
    # Process all XML files
    process_all_xml_files(xml_dir, host, port, db_name, individual_collection, master_collection)


if __name__ == "__main__":
    main()