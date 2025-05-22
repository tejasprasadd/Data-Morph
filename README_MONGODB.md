# NIRF Data MongoDB Import Guide

This guide explains how to import NIRF XML data into MongoDB and create a structured master database using the provided script.

## Prerequisites

1. **MongoDB Installation**
   - Download and install MongoDB Community Edition from [MongoDB Download Center](https://www.mongodb.com/try/download/community)
   - Follow the installation instructions for your operating system
   - Start the MongoDB service (ensure it's running on the default port 27017)

2. **Python Dependencies**
   - Install the required Python packages:
     ```
     pip install pymongo tqdm
     ```
     or
     ```
     pip install -r requirements.txt
     ```

3. **XML Files**
   - Ensure all 100 XML files are available in the `All-XML` directory

## Importing XML Data to MongoDB

### Basic Usage

Run the script without arguments to process all XML files in the All-XML directory:

```
python xml_to_mongodb.py
```

This will:
- Process all XML files in the `./All-XML` directory
- Connect to MongoDB at `mongodb://localhost:27017/`
- Create a database named `nirf_database`
- Create two collections: `individuals` and `master_database`

### Advanced Usage

You can specify custom parameters for batch processing:

```
python xml_to_mongodb.py [xml_dir] [host] [port] [db_name] [individual_collection] [master_collection]
```

Example:
```
python xml_to_mongodb.py ./All-XML localhost 27017 nirf_database individuals master_database
```

## Database Structure

### Individual Collection
Each document in the `individuals` collection represents data from a single XML file with the following structure:

```
{
  "_id": "individual_[college_id]_[timestamp]",
  "institute": {
    "name": "College Name",
    "source_file": "filename.xml",
    "college_id": "unique_id"
  },
  "sanctionedintake": [...],
  "studentstrength": [...],
  "placementdata": [...],
  "phddata": [...],
  "capitalexpenditure": [...],
  "operationalexpenditure": [...],
  "sponsoredprojects": [...],
  "consultancyprojects": [...],
  "facilities": [...],
  "facultycount": [...],
  "metadata": {
    "imported_at": "timestamp",
    "source_file": "filename.xml"
  }
}
```

### Master Database Collection
The `master_database` collection organizes data by college with standardized field names and proper structure:

```
{
  "college_id": "unique_id",
  "college_name": "College Name",
  "metadata": {
    "created_at": "timestamp",
    "last_updated": "timestamp",
    "source_files": ["file1.xml", "file2.xml"]
  },
  "sanctionedintake": [
    {
      "program_name": "B.Tech",
      "academic_year": "2022-23",
      "approved_intake": 120
    },
    ...
  ],
  "studentstrength": [...],
  "placementdata": [...],
  "phddata": [...],
  "capitalexpenditure": [...],
  "operationalexpenditure": [...],
  "sponsoredprojects": [...],
  "consultancyprojects": [...],
  "facilities": [...],
  "facultycount": [...]
}
```

## Features

1. **Standardized Field Names**: The script normalizes field names across different XML files to ensure consistency.

2. **Structured Data Organization**: Data is organized by college with proper relationships between different sections.

3. **Automatic ID Generation**: If a college ID is not found in the XML, the script attempts to extract it from the filename or generates a unique ID.

4. **Data Cleaning**: The script cleans and standardizes data values, converting numeric strings to proper number types.

5. **Incremental Updates**: If new XML files are added, the script will update the master database without duplicating existing data.

## Viewing the Data

You can use MongoDB Compass (a GUI tool) to view and query the imported data:

1. Download and install [MongoDB Compass](https://www.mongodb.com/try/download/compass)
2. Connect to your MongoDB server
3. Navigate to the `nirf_database` database and explore both collections:
   - `individuals`: Raw data from each XML file
   - `master_database`: Structured data organized by college

## Querying the Database

Example MongoDB queries to retrieve data:

```python
# Connect to MongoDB
from pymongo import MongoClient
client = MongoClient('mongodb://localhost:27017/')
db = client['nirf_database']

# Get all colleges
colleges = list(db.master_database.find({}, {"college_id": 1, "college_name": 1}))

# Get placement data for a specific college
placement_data = db.master_database.find_one(
    {"college_id": "001"}, 
    {"placementdata": 1}
)

# Find colleges with high placement rates
high_placement_colleges = db.master_database.find({
    "placementdata.placed": {"$gt": 100}
})
```

## Troubleshooting

- **Connection Issues**: Ensure MongoDB service is running on port 27017
- **Import Errors**: Check that the XML files are valid and properly formatted
- **Missing Data**: Verify that the XML files contain the expected sections and data
- **Duplicate Key Errors**: If you encounter duplicate key errors, it may indicate that the same college ID exists in multiple XML files