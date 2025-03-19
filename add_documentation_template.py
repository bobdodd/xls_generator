#!/usr/bin/env python3
"""
Helper script to add TEST_DOCUMENTATION template to test files.

Usage: 
1. Run the script with a test file path:
   python add_documentation_template.py /path/to/test_file.py

2. The script will:
   - Create a backup of the original file
   - Add the TEST_DOCUMENTATION template at the top of the file
   - The template can then be filled in with specific test documentation
"""

import sys
import os
import re
from datetime import datetime

# Documentation template to insert
DOCUMENTATION_TEMPLATE = """# Test metadata for documentation and reporting
TEST_DOCUMENTATION = {
    "testName": "INSERT_TEST_NAME_HERE",
    "description": "INSERT_GENERAL_DESCRIPTION_HERE",
    "version": "1.0.0",
    "date": "YYYY-MM-DD",
    "dataSchema": {
        "timestamp": "ISO timestamp when the test was run",
        "url": "The URL of the page being analyzed"
        # Add other top-level fields in test results
    },
    "tests": [
        {
            "id": "INSERT_TEST_ID_HERE",
            "name": "INSERT_TEST_NAME_HERE",
            "description": "INSERT_TEST_DESCRIPTION_HERE",
            "impact": "high|medium|low",
            "wcagCriteria": ["1.1.1", "1.3.1"],  # List applicable WCAG criteria
            "howToFix": "INSERT_INSTRUCTIONS_HERE",
            "resultsFields": {
                "pageFlags.hasSomeFlag": "Description of this flag",
                "keyElements.primaryElement": "Description of this element"
                # Add other result fields specific to this test
            }
        }
        # Add more test objects for each subtest or check
    ]
}
"""

def add_documentation_to_file(file_path):
    """
    Add the TEST_DOCUMENTATION template to a test file.
    
    Args:
        file_path: Path to the test file to modify
    """
    if not os.path.exists(file_path) or not file_path.endswith('.py'):
        print(f"Error: {file_path} is not a valid Python file.")
        return False
    
    # Extract the test name from the file name (remove "test_" prefix and ".py" suffix)
    filename = os.path.basename(file_path)
    test_name = filename[5:-3] if filename.startswith('test_') else filename[:-3]
    
    # Format the test name for the documentation
    formatted_test_name = ' '.join(word.capitalize() for word in test_name.split('_')) + ' Analysis'
    
    # Load the file content
    with open(file_path, 'r') as f:
        content = f.read()
    
    # Check if TEST_DOCUMENTATION already exists
    if 'TEST_DOCUMENTATION' in content:
        print(f"Warning: TEST_DOCUMENTATION already exists in {file_path}. Skipping.")
        return False
    
    # Create a backup of the original file
    backup_path = file_path + '.bak'
    with open(backup_path, 'w') as f:
        f.write(content)
    
    # Prepare the documentation template with the test name filled in
    template = DOCUMENTATION_TEMPLATE.replace('INSERT_TEST_NAME_HERE', formatted_test_name)
    template = template.replace('YYYY-MM-DD', datetime.now().strftime('%Y-%m-%d'))
    
    # Find the right place to insert the documentation
    # Typically after imports and docstring, before the first function
    docstring_pattern = r'^""".*?"""'
    import_pattern = r'^(?:import|from)\s+.*$'
    
    # Check if the file has a module docstring
    docstring_match = re.search(docstring_pattern, content, re.DOTALL | re.MULTILINE)
    
    # Find the last import statement
    import_matches = list(re.finditer(import_pattern, content, re.MULTILINE))
    
    # Determine insertion point
    if docstring_match:
        # Insert after the docstring
        insertion_point = docstring_match.end()
        modified_content = content[:insertion_point] + '\n\n' + template + '\n\n' + content[insertion_point:]
    elif import_matches:
        # Insert after the last import
        last_import = import_matches[-1]
        insertion_point = last_import.end()
        modified_content = content[:insertion_point] + '\n\n' + template + '\n\n' + content[insertion_point:]
    else:
        # Insert at the top of the file
        modified_content = template + '\n\n' + content
    
    # Write the modified content back to the file
    with open(file_path, 'w') as f:
        f.write(modified_content)
    
    print(f"Added TEST_DOCUMENTATION template to {file_path}")
    print(f"Backup saved to {backup_path}")
    return True

def main():
    if len(sys.argv) < 2:
        print("Usage: python add_documentation_template.py /path/to/test_file.py")
        return
    
    file_path = sys.argv[1]
    add_documentation_to_file(file_path)

if __name__ == "__main__":
    main()