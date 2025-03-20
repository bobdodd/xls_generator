#!/usr/bin/env python3
"""
Script to generate an accessibility report that includes documentation,
with support for specifying which MongoDB database to use.
"""
import os
import sys
import argparse
from datetime import datetime
from pymongo import MongoClient
from bson.objectid import ObjectId

# Add the source directory to path
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

# Import the generator and documentation
from xls_generator import AccessibilityReportGenerator, AccessibilityDB, DEFAULT_DB_NAME

def main(db_name=None, create_doc_run=False):
    try:
        # Use the specified database name or default
        if db_name is None:
            db_name = DEFAULT_DB_NAME
            print(f"Warning: No database name specified. Using default database '{DEFAULT_DB_NAME}'.")
        
        print(f"Connecting to MongoDB database: '{db_name}'")
        
        # Connect to MongoDB with the specified database
        client = MongoClient('mongodb://localhost:27017/')
        db = client[db_name]
        
        # Find our documentation test run
        test_run = db.test_runs.find_one({"name": "Documentation Test Run"})
        
        # If no documentation test run exists and create_doc_run is True, create one
        if not test_run and create_doc_run:
            print("Documentation Test Run not found. Creating a new one...")
            
            # Get the most recent test run to use as a template
            latest_run = db.test_runs.find_one(sort=[('timestamp_start', -1)])
            
            if latest_run:
                # Create a new documentation test run based on the latest run
                doc_run = {
                    "name": "Documentation Test Run",
                    "timestamp_start": datetime.now(),
                    "timestamp_end": datetime.now(),
                    "urls": latest_run.get("urls", []),
                    "tests": latest_run.get("tests", {}),
                    "documentation": {}
                }
                
                # Add documentation for standard test types
                from add_direct_documentation import (
                    ANIMATIONS_DOCUMENTATION,
                    COLORS_DOCUMENTATION,
                    FORMS_DOCUMENTATION
                )
                doc_run["documentation"]["animations"] = ANIMATIONS_DOCUMENTATION
                doc_run["documentation"]["colors"] = COLORS_DOCUMENTATION
                doc_run["documentation"]["forms"] = FORMS_DOCUMENTATION
                
                # Insert the documentation test run
                result = db.test_runs.insert_one(doc_run)
                test_run_id = result.inserted_id
                print(f"Created Documentation Test Run with ID: {test_run_id}")
            else:
                print("No existing test runs found to use as a template.")
                print("Cannot create Documentation Test Run without a template.")
                return
        elif not test_run:
            print("Documentation Test Run not found in database.")
            print("Run with --create-doc-run flag to create a documentation test run.")
            return
        else:
            test_run_id = test_run['_id']
            print(f"Found Documentation Test Run with ID: {test_run_id}")
        
        # Find the most recent regular test run to merge results
        latest_run = db.test_runs.find_one(
            {"name": {"$ne": "Documentation Test Run"}},
            sort=[('timestamp_start', -1)]
        )
        
        if not latest_run:
            print("No regular test run found. Using only documentation test run.")
            test_run_ids = [test_run_id]
        else:
            latest_run_id = latest_run['_id']
            print(f"Found latest regular test run with ID: {latest_run_id}")
            test_run_ids = [test_run_id, latest_run_id]
        
        # Generate report with combined results
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f"documentation_report_{timestamp}.xlsx"
        
        print(f"Generating combined report for {len(test_run_ids)} test runs")
        
        # Initialize DB wrapper with the specified database name
        db_wrapper = AccessibilityDB(db_name=db_name)
        generator = AccessibilityReportGenerator(db_wrapper)
        
        # Generate the report with database name prepended to the filename
        generator.generate_excel_report(test_run_ids, output_file, db_name=db_name)
        
        # Determine the actual filename after db_name prepending
        if db_name and db_name != DEFAULT_DB_NAME:
            base, ext = output_file.rsplit('.', 1)
            actual_file = f"{db_name}_{base}.{ext}"
        else:
            actual_file = output_file
            
        print(f"Documentation report generated successfully: {actual_file}")

    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Generate Excel report from MongoDB accessibility test results with documentation')
    parser.add_argument('--database', '-db', help=f'MongoDB database name to use (default: {DEFAULT_DB_NAME})')
    parser.add_argument('--create-doc-run', '-c', action='store_true', help='Create a Documentation Test Run if one does not exist')
    args = parser.parse_args()
    
    main(db_name=args.database, create_doc_run=args.create_doc_run)