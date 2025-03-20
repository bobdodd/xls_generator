import pandas as pd
from jinja2 import Environment, FileSystemLoader
from datetime import datetime
from pymongo import MongoClient
import openpyxl
from openpyxl.styles import Alignment
from urllib.parse import urlparse
from collections import defaultdict
import json

class TemplateAnalyzer:
    def __init__(self, db):
        self.db = db
        self.structures = defaultdict(set)
        self.examples = {}

    def analyze_test_structures(self):
        """Analyze all test structures in the database"""
        results = self.db.page_results.find()
        
        for result in results:
            if 'results' in result and 'accessibility' in result['results']:
                tests = result['results']['accessibility'].get('tests', {})
                
                # Analyze each test type
                for test_name, test_data in tests.items():
                    if isinstance(test_data, dict):
                        self._record_structure(self.structures[test_name], test_data)
                        
                        # Store first example if we don't have one yet
                        if test_name not in self.examples:
                            self.examples[test_name] = test_data
        
        return self.structures, self.examples

    def _record_structure(self, structure_set, data, path=""):
        """Recursively record the structure of a dictionary"""
        if isinstance(data, dict):
            for key, value in data.items():
                new_path = f"{path}.{key}" if path else key
                structure_set.add(f"{new_path} ({type(value).__name__})")
                self._record_structure(structure_set, value, new_path)
        elif isinstance(data, list) and data:
            # Record structure of first item in list as example
            if data and isinstance(data[0], dict):
                self._record_structure(structure_set, data[0], f"{path}[]")

    def print_analysis(self):
        """Print a summary of all test types and their structures"""
        print("\nTest Types Analysis Summary:")
        print("=" * 50)
        
        for test_name in sorted(self.structures.keys()):
            print(f"\n{test_name}:")
            print("-" * len(test_name))
            for path in sorted(self.structures[test_name]):
                print(f"  {path}")

DEFAULT_DB_NAME = 'accessibility_tests'

class AccessibilityDB:
    def __init__(self, db_name=None):
        try:
            self.client = MongoClient('mongodb://localhost:27017/')
            
            # Use the specified database name or default
            if db_name is None:
                db_name = DEFAULT_DB_NAME
                print(f"Warning: No database name specified for XLS generator. Using default database '{DEFAULT_DB_NAME}'.")
            
            self.db_name = db_name
            self.db = self.client[db_name]
            self.test_runs = self.db['test_runs']
            self.page_results = self.db['page_results']
            
            print(f"XLS Generator connected to database: '{db_name}'")
        except Exception as e:
            print(f"Failed to connect to MongoDB: {e}")
            raise

    def get_page_results(self, test_run_ids=None):
        """Get all page results for specific test runs"""
        query = {}
        if test_run_ids:
            if isinstance(test_run_ids, list):
                query['test_run_id'] = {'$in': test_run_ids}
            else:
                query['test_run_id'] = test_run_ids
        return list(self.page_results.find(query))

    def get_most_recent_test_run_id(self):
        """Get the most recent test run ID"""
        latest_run = self.test_runs.find_one(
            sort=[('timestamp_start', -1)]
        )
        return str(latest_run['_id']) if latest_run else None

    def get_all_test_run_ids(self):
        """Get all test run IDs"""
        test_runs = self.test_runs.find({}, {'_id': 1})
        return [str(run['_id']) for run in test_runs]
        
    def get_test_run_by_name(self, name):
        """Get a test run by name"""
        test_run = self.test_runs.find_one({"name": name})
        return test_run

    def __del__(self):
        if hasattr(self, 'client'):
            self.client.close()

class AccessibilityReportGenerator:
    def __init__(self, db, structures=None, examples=None):
        self.db = db
        self.env = Environment(loader=FileSystemLoader('templates'))
        self.structures = structures or {}
        self.examples = examples or {}
        self.test_documentation = {}  # Store test documentation from all tests

    def collect_test_documentation(self, results):
        """
        Extract and collect test documentation from test results and test run metadata
        
        Args:
            results: List of page result documents from MongoDB
            
        Returns:
            Dictionary with test documentation organized by test type
        """
        total_docs_found = 0
        
        # First, check if documentation exists in the test_runs collection
        print("Checking test_runs collection for documentation...")
        try:
            # Get all test runs
            test_runs = list(self.db.test_runs.find({}))
            for test_run in test_runs:
                # Check if this test run has documentation
                if 'documentation' in test_run:
                    documentation = test_run.get('documentation', {})
                    if isinstance(documentation, dict):
                        for test_name, doc_data in documentation.items():
                            if test_name not in self.test_documentation:
                                print(f"Found documentation for {test_name} in test_run {test_run.get('_id')}")
                                self.test_documentation[test_name] = doc_data
                                total_docs_found += 1
                
                # Also check for a nested tests collection within the test run
                if 'tests' in test_run:
                    tests = test_run.get('tests', {})
                    for test_name, test_data in tests.items():
                        if isinstance(test_data, dict) and 'documentation' in test_data:
                            if test_name not in self.test_documentation:
                                print(f"Found documentation for {test_name} in test_run.tests")
                                self.test_documentation[test_name] = test_data['documentation']
                                total_docs_found += 1
        except Exception as e:
            print(f"Error checking test_runs collection: {str(e)}")
            
        # Then continue with checking page results
        for result in results:
            print(f"Processing result for URL: {result.get('url', 'unknown')}")
            
            if 'results' in result and 'accessibility' in result['results']:
                tests = result['results']['accessibility'].get('tests', {})
                print(f"Found {len(tests)} test types in this result")
                
                for test_name, test_data in tests.items():
                    print(f"  Checking {test_name} test data...")
                    
                    if isinstance(test_data, dict):
                        # Check for documentation directly in the test result data (new structure)
                        print(f"  Keys in {test_name} test data: {list(test_data.keys())}")
                        if 'documentation' in test_data:
                            print(f"  Found documentation directly in {test_name} -> documentation")
                            doc_data = test_data['documentation']
                            print(f"  Documentation content: {doc_data.get('testName', 'N/A')}, tests: {len(doc_data.get('tests', []))}")
                            if test_name in self.test_documentation:
                                print(f"  Already have documentation for {test_name}, skipping duplicate")
                            else:
                                self.test_documentation[test_name] = test_data['documentation']
                                total_docs_found += 1
                                print(f"  Added documentation for {test_name} with {len(test_data['documentation'].get('tests', []))} individual tests")
                            continue
                            
                        # First, check for documentation in nested test_data if it's nested
                        if test_name in test_data and 'documentation' in test_data[test_name]:
                            print(f"  Found nested documentation in {test_name} -> {test_name} -> documentation")
                            if test_name in self.test_documentation:
                                print(f"  Already have documentation for {test_name}, skipping duplicate")
                            else:
                                self.test_documentation[test_name] = test_data[test_name]['documentation']
                                total_docs_found += 1
                                print(f"  Added documentation for {test_name} with {len(test_data[test_name]['documentation'].get('tests', []))} individual tests")
                        else:
                            # Extra debugging for example.com tests
                            if result.get('url') == 'https://example.com':
                                print(f"  No documentation found in {test_name} test data. Keys at this level: {list(test_data.keys())}")
                                # If there's a nested structure with the same name, inspect it
                                if test_name in test_data:
                                    print(f"  Found nested {test_name} object. Keys: {list(test_data[test_name].keys())}")
                                    if 'documentation' in test_data[test_name]:
                                        print(f"  Found documentation in nested structure!")
                                        if test_name in self.test_documentation:
                                            print(f"  Already have documentation for {test_name}, skipping duplicate")
                                        else:
                                            self.test_documentation[test_name] = test_data[test_name]['documentation']
                                            total_docs_found += 1
                                            print(f"  Added documentation for {test_name} with {len(test_data[test_name]['documentation'].get('tests', []))} individual tests")
        
        # Add a separate step to directly check for documentation in specific locations
        # This addresses the different ways tests might structure their output
        print("\nAttempting comprehensive documentation scan with multiple strategies...")
        
        # Map of possible test output field names based on test names
        test_output_fields = {
            'images': 'images',
            'tables': 'tables', 
            'headings': 'headings',
            'focus_management': 'focus_management',
            'page_structure': 'page_structure',
            'accessible_names': 'accessible_names',
            'landmarks': 'landmarks',
            'forms': 'forms',
            'colors': 'colors',
            'html_structure': 'html_structure',
            'animations': 'animations',
            'videos': 'videos'
            # Add more mappings as needed
        }
        
        print("Scanning database for test documentation using multiple approaches...")
        for result in results:
            if 'results' in result and 'accessibility' in result['results']:
                tests = result['results']['accessibility'].get('tests', {})
                
                for test_name, test_data in tests.items():
                    # Skip if we already found documentation for this test
                    if test_name in self.test_documentation:
                        continue
                    
                    # 1. Check if the test itself contains documentation field
                    if isinstance(test_data, dict) and 'documentation' in test_data:
                        print(f"Found direct documentation in test data for {test_name}")
                        self.test_documentation[test_name] = test_data['documentation']
                        total_docs_found += 1
                        continue
                    
                    # 2. Check if test has a corresponding field with documentation
                    output_field = test_output_fields.get(test_name, test_name)
                    if isinstance(test_data, dict) and output_field in test_data:
                        output_data = test_data[output_field]
                        if isinstance(output_data, dict) and 'documentation' in output_data:
                            print(f"Found documentation in {test_name} -> {output_field} field")
                            self.test_documentation[test_name] = output_data['documentation']
                            total_docs_found += 1
                            continue
                    
                    # 3. For specific multi-word test names, try direct matching of test field
                    if '_' in test_name:
                        # Try with underscores
                        if test_name in test_data:
                            field_data = test_data[test_name]
                            if isinstance(field_data, dict) and 'documentation' in field_data:
                                print(f"Found documentation using underscore name match: {test_name}")
                                self.test_documentation[test_name] = field_data['documentation']
                                total_docs_found += 1
                                continue
                    
                    # 4. Look inside the main result for top-level fields that might contain documentation
                    # This checks for potential custom field names not in our mapping
                    for field_name, field_data in test_data.items():
                        if isinstance(field_data, dict) and 'documentation' in field_data:
                            print(f"Found documentation in {test_name} -> {field_name} field")
                            self.test_documentation[test_name] = field_data['documentation']
                            total_docs_found += 1
                            break
        
        print(f"Collected documentation for {len(self.test_documentation)} test types (found {total_docs_found} new)")
        return self.test_documentation
    
    def format_issue_name(self, test_name, flag_name):
        """Format issue name consistently with documentation if available"""
        # Remove 'has' prefix
        flag_text = flag_name[3:] if flag_name.startswith('has') else flag_name
        
        # Split camel case into words
        words = []
        current_word = ''
        for char in flag_text:
            if char.isupper() and current_word:
                words.append(current_word)
                current_word = char
            else:
                current_word += char
        words.append(current_word)
        
        # Format issue type and test name
        issue_type = ' '.join(words).title()
        formatted_test_name = ' '.join(
            word.capitalize() 
            for word in test_name.split('_')
        )
        
        # Try multiple variants of the test name to match documentation
        test_name_variants = [
            test_name,                     # Original name
            test_name.replace('-', '_'),   # Replace hyphens with underscores
            test_name.replace('_', '-'),   # Replace underscores with hyphens
            test_name.lower(),             # Lowercase
            test_name.replace('-', '').replace('_', '')  # No separators
        ]
        
        # Check if we have documentation for any variant of this test name
        for variant in test_name_variants:
            if variant in self.test_documentation:
                docs = self.test_documentation[variant]
                doc_test_name = docs.get('testName', formatted_test_name)
                
                # Try to find a specific test that matches this flag
                for test in docs.get('tests', []):
                    results_fields = test.get('resultsFields', {})
                    
                    # Look for various ways the flag could be referenced
                    flag_variants = [
                        f"pageFlags.{flag_name}",
                        f"pageFlags.has{flag_text}",
                        f"details.{flag_name}",
                        flag_name
                    ]
                    
                    # Look for any matching result field
                    for field in results_fields:
                        if any(field.endswith(fv) or field == fv for fv in flag_variants):
                            # Use the documented test name instead
                            return f"{doc_test_name} - {test.get('name', issue_type)}"
                            
                # If no specific test found, use the general test name from documentation
                return f"{doc_test_name}: {issue_type}"
        
        # Fallback to formatted name if no documentation found
        return f"{formatted_test_name}: {issue_type}"

    def format_json_as_table(self, data, indent=0):
        """Convert JSON data to a readable table format"""
        if not isinstance(data, (dict, list)):
            return str(data)

        lines = []
        indent_str = "  " * indent

        if isinstance(data, dict):
            for key, value in data.items():
                if isinstance(value, (dict, list)):
                    lines.append(f"{indent_str}{key}:")
                    lines.append(self.format_json_as_table(value, indent + 1))
                else:
                    lines.append(f"{indent_str}{key}: {value}")
        elif isinstance(data, list):
            for item in data:
                if isinstance(item, (dict, list)):
                    lines.append(self.format_json_as_table(item, indent))
                else:
                    lines.append(f"{indent_str}- {item}")

        return "\n".join(lines)

    def calculate_summary(self, test_run_ids=None):
        """Calculate summary statistics across all test runs"""
        # If no test_run_ids provided, get all of them
        if test_run_ids is None:
            test_run_ids = self.db.get_all_test_run_ids()
        
        # Build a query that includes all the specified test runs
        query = {}
        if test_run_ids:
            if isinstance(test_run_ids, list):
                query['test_run_id'] = {'$in': test_run_ids}
            else:
                query['test_run_id'] = test_run_ids
        
        results = list(self.db.page_results.find(query))
        
        summary = {
            'total_pages': len(results),
            'pages_with_issues': 0,
            'total_issues': 0,
            'issues_by_type': {},
            'completion_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }

        for result in results:
            if 'results' in result and 'accessibility' in result['results']:
                tests = result['results']['accessibility'].get('tests', {})
                
                for test_name, test_data in tests.items():
                    if isinstance(test_data, dict):
                        test_obj = test_data.get(test_name, {})
                        page_flags = test_obj.get('pageFlags', {})
                        
                        if page_flags:
                            for flag_name, flag_value in page_flags.items():
                                if isinstance(flag_value, bool) and flag_value and flag_name.startswith('has'):
                                    full_issue_type = self.format_issue_name(test_name, flag_name)
                                    
                                    details = page_flags.get('details', {})
                                    count = 1
                                    
                                    for detail_key, detail_value in details.items():
                                        if isinstance(detail_value, (list, int)):
                                            if isinstance(detail_value, list):
                                                count = len(detail_value)
                                            else:
                                                count = detail_value
                                            
                                            if count > 0:
                                                summary['total_issues'] += count
                                                summary['issues_by_type'][full_issue_type] = \
                                                    summary['issues_by_type'].get(full_issue_type, 0) + count

        summary['pages_with_issues'] = len([
            r for r in results 
            if 'results' in r and 'accessibility' in r['results'] and r['results']['accessibility'].get('tests', {})
        ])

        return summary

    def format_detailed_results(self, results):
        """Format detailed results for Excel with improved JSON formatting"""
        formatted_data = {}
        
        def process_dict(d, prefix=''):
            items = {}
            for key, value in d.items():
                full_key = f"{prefix}.{key}" if prefix else key
                if isinstance(value, dict):
                    items.update(process_dict(value, full_key))
                elif isinstance(value, list):
                    items[full_key] = json.dumps(value)
                else:
                    items[full_key] = str(value)
            return items

        for result in results:
            url = result.get('url', 'Unknown URL')
            accessibility = result.get('results', {}).get('accessibility', {})
            
            if accessibility and 'tests' in accessibility:
                tests = accessibility['tests']
                
                for test_name, test_results in tests.items():
                    if isinstance(test_results, dict):
                        # Process the dictionary to get flat structure with full paths
                        flat_data = process_dict(test_results, test_name)
                        for key, value in flat_data.items():
                            if key not in formatted_data:
                                formatted_data[key] = {}
                            formatted_data[key][url] = value
                    else:
                        if test_name not in formatted_data:
                            formatted_data[test_name] = {}
                        formatted_data[test_name][url] = str(test_results)
        
        # Convert to DataFrame
        df = pd.DataFrame.from_dict(formatted_data, orient='index')
        df = df.sort_index()
        
        return df

    def generate_excel_report(self, test_run_ids=None, output_file='accessibility_report.xlsx', db_name=None):
        """Generate Excel report with multiple sheets using all test runs"""
        # If no test_run_ids provided, get all of them
        if test_run_ids is None:
            test_run_ids = self.db.get_all_test_run_ids()
            
        # Prepend database name to output file if provided
        if db_name:
            # Extract the file extension and base name
            if '.' in output_file:
                base, ext = output_file.rsplit('.', 1)
                output_file = f"{db_name}_{base}.{ext}"
            else:
                output_file = f"{db_name}_{output_file}"
        
        # Build a query that includes all the specified test runs
        query = {}
        if test_run_ids:
            if isinstance(test_run_ids, list):
                query['test_run_id'] = {'$in': test_run_ids}
            else:
                query['test_run_id'] = test_run_ids
        
        results = list(self.db.page_results.find(query))
        
        # First, explicitly fetch test documentation from test_runs collection
        # This ensures we get documentation for all tests, even those not in the current results
        print("\n=== DETAILED DOCUMENTATION COLLECTION DEBUG ===")
        print("Explicitly fetching documentation from test_runs collection...")
        try:
            # Get all test runs to look for documentation
            all_test_runs = list(self.db.test_runs.find({}))
            print(f"Found {len(all_test_runs)} test runs in the database")
            
            # Look for specific tests we want to document
            target_tests = ['animations', 'colors', 'forms']
            print(f"Specifically looking for documentation for: {', '.join(target_tests)}")
            
            for run in all_test_runs:
                print(f"\nInspecting test run: {run.get('_id')} ({run.get('name', 'unnamed run')})")
                
                # Debug the test run
                if 'documentation' in run:
                    doc_count = len(run['documentation'])
                    print(f"  Found documentation with {doc_count} test types")
                    print(f"  Documentation keys: {list(run['documentation'].keys())}")
                    
                    # Check for our specific targets
                    for target in target_tests:
                        if target in run['documentation']:
                            print(f"  ✓ Found '{target}' documentation in this test run")
                        elif any(k.startswith(target) or k.endswith(target) for k in run['documentation'].keys()):
                            matching_keys = [k for k in run['documentation'].keys() if k.startswith(target) or k.endswith(target)]
                            print(f"  ⚠ Found similar keys to '{target}': {matching_keys}")
                        else:
                            print(f"  ✗ Did not find '{target}' documentation in this test run")
                    
                    # Process all documentation
                    for test_name, doc in run['documentation'].items():
                        # Store multiple variants of the name to improve matching
                        test_keys = [
                            test_name.lower(),
                            test_name.replace('-', '_').lower(),
                            test_name.replace('_', '-').lower()
                        ]
                        
                        # Use the best key (original first)
                        for test_key in test_keys:
                            if test_key not in self.test_documentation:
                                print(f"  Adding documentation for {test_name} with key {test_key}")
                                self.test_documentation[test_key] = doc
                                # Also add variant keys for improved matching
                                if test_key != test_name.lower():
                                    self.test_documentation[test_name.lower()] = doc
                                break
                            
                # Check for inline documentation in test results
                if 'tests' in run:
                    print(f"  This run has a 'tests' field with {len(run.get('tests', {}))} items")
                    tests = run.get('tests', {})
                    for test_name, test_data in tests.items():
                        if isinstance(test_data, dict) and 'documentation' in test_data:
                            print(f"  Found inline documentation for {test_name} in tests field")
                            self.test_documentation[test_name.lower()] = test_data['documentation']
                            
        except Exception as e:
            print(f"Error fetching documentation from test_runs: {e}")
            import traceback
            traceback.print_exc()
        
        # Then, collect test documentation from results
        self.collect_test_documentation(results)
        
        # For debugging: print all test documentation found
        print("\n=== FINAL DOCUMENTATION SUMMARY ===")
        print(f"Total test types with documentation: {len(self.test_documentation)}")
        
        # Check specifically for our target tests
        target_tests = ['animations', 'colors', 'forms']
        for target in target_tests:
            found = False
            for key in self.test_documentation.keys():
                if key == target or key.startswith(target) or key.endswith(target):
                    found = True
                    print(f"✓ '{target}' documentation found as key '{key}'")
                    break
            if not found:
                print(f"✗ '{target}' documentation NOT FOUND in any key")
        
        print("\nAll documented test types:")
        for test_name in sorted(self.test_documentation.keys()):
            test_obj = self.test_documentation[test_name]
            display_name = test_obj.get('testName', test_name)
            tests_count = len(test_obj.get('tests', []))
            print(f"  - {test_name} → {display_name} ({tests_count} subtests)")
        
        summary = self.calculate_summary(test_run_ids)
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Summary sheet
            summary_df = pd.DataFrame({
                'Metric': [
                    'Total Pages',
                    'Pages with Issues',
                    'Total Issues',
                    'Completion Time'
                ],
                'Value': [
                    summary['total_pages'],
                    summary['pages_with_issues'],
                    summary['total_issues'],
                    summary['completion_time']
                ]
            })
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Issues by type
            if summary['issues_by_type']:
                issues_df = pd.DataFrame(
                    summary['issues_by_type'].items(),
                    columns=['Issue Type', 'Count']
                )
                issues_df.to_excel(writer, sheet_name='Issues by Type', index=False)
            
            # Issues by URL
            issues_by_url = {}
            for result in results:
                url = result.get('url', 'Unknown URL')
                if 'results' in result and 'accessibility' in result['results']:
                    tests = result['results']['accessibility']['tests']
                    url_issues = []
                    
                    for test_name, test_data in tests.items():
                        if isinstance(test_data, dict):
                            test_obj = test_data.get(test_name, {})
                            page_flags = test_obj.get('pageFlags', {})
                            
                            for flag_name, flag_value in page_flags.items():
                                if isinstance(flag_value, bool) and flag_value and flag_name.startswith('has'):
                                    full_issue_type = self.format_issue_name(test_name, flag_name)
                                    
                                    details = page_flags.get('details', {})
                                    count = 1
                                    for detail_key, detail_value in details.items():
                                        if isinstance(detail_value, (list, int)):
                                            if isinstance(detail_value, list):
                                                count = len(detail_value)
                                            else:
                                                count = detail_value
                                            if count > 0:
                                                url_issues.append({
                                                    'Issue Type': full_issue_type,
                                                    'Count': count,
                                                    'Details': f"Found {count} issue(s)"
                                                })
                    
                    if url_issues:
                        issues_by_url[url] = url_issues
            
            if issues_by_url:
                url_issues_data = []
                for url, issues in issues_by_url.items():
                    for issue in issues:
                        url_issues_data.append({
                            'URL': url,
                            'Issue Type': issue['Issue Type'],
                            'Count': issue['Count'],
                            'Details': issue['Details']
                        })
                
                url_issues_df = pd.DataFrame(url_issues_data)
                url_issues_df.to_excel(writer, sheet_name='Issues by URL', index=False)
            
            # Issues by Site
            issues_by_site = {}
            for result in results:
                full_url = result.get('url', 'Unknown URL')
                try:
                    parsed_url = urlparse(full_url)
                    site_url = f"{parsed_url.scheme}://{parsed_url.netloc}"
                except:
                    site_url = 'Unknown Site'

                if 'results' in result and 'accessibility' in result['results']:
                    tests = result['results']['accessibility']['tests']
                    
                    if site_url not in issues_by_site:
                        issues_by_site[site_url] = {}
                    
                    for test_name, test_data in tests.items():
                        if isinstance(test_data, dict):
                            test_obj = test_data.get(test_name, {})
                            page_flags = test_obj.get('pageFlags', {})
                            
                            for flag_name, flag_value in page_flags.items():
                                if isinstance(flag_value, bool) and flag_value and flag_name.startswith('has'):
                                    full_issue_type = self.format_issue_name(test_name, flag_name)
                                    
                                    details = page_flags.get('details', {})
                                    count = 1
                                    for detail_key, detail_value in details.items():
                                        if isinstance(detail_value, (list, int)):
                                            if isinstance(detail_value, list):
                                                count = len(detail_value)
                                            else:
                                                count = detail_value
                                            if count > 0:
                                                if full_issue_type not in issues_by_site[site_url]:
                                                    issues_by_site[site_url][full_issue_type] = {
                                                        'Count': 0,
                                                        'Pages Affected': 0
                                                    }
                                                issues_by_site[site_url][full_issue_type]['Count'] += count
                                                issues_by_site[site_url][full_issue_type]['Pages Affected'] += 1

            if issues_by_site:
                site_issues_data = []
                for site_url, issues in issues_by_site.items():
                    for issue_type, data in issues.items():
                        site_issues_data.append({
                            'Site': site_url,
                            'Issue Type': issue_type,
                            'Total Count': data['Count'],
                            'Pages Affected': data['Pages Affected']
                        })
                
                site_issues_df = pd.DataFrame(site_issues_data)
                site_issues_df.to_excel(writer, sheet_name='Issues by Site', index=False)
            
            # Documentation sheet
            if self.test_documentation:
                print("\nPreparing documentation sheet...")
                docs_data = []
                for test_name, doc in self.test_documentation.items():
                    # Get a clean test name for display
                    display_test_name = doc.get('testName', test_name.replace('_', ' ').title())
                    
                    print(f"Adding documentation for {display_test_name}")
                    
                    # Add test general documentation
                    docs_data.append({
                        'Test Name': display_test_name,
                        'Type': 'Test',
                        'Description': doc.get('description', ''),
                        'Version': doc.get('version', ''),
                        'Date': doc.get('date', ''),
                        'WCAG Criteria': '',
                        'Impact': '',
                        'How to Fix': ''
                    })
                    
                    # Add individual checks documentation
                    for test in doc.get('tests', []):
                        subtest_name = test.get('name', '')
                        if subtest_name:  # Only add subtests that have names
                            full_name = f"{display_test_name} - {subtest_name}"
                            print(f"  Adding subtest: {subtest_name}")
                            
                            docs_data.append({
                                'Test Name': full_name,
                                'Type': 'Check',
                                'Description': test.get('description', ''),
                                'Version': doc.get('version', ''),
                                'Date': doc.get('date', ''),
                                'WCAG Criteria': ', '.join(test.get('wcagCriteria', [])),
                                'Impact': test.get('impact', ''),
                                'How to Fix': test.get('howToFix', '')
                            })
                
                if docs_data:
                    # Sort by Test Name and make Type a secondary sort key (Test first, then Check)
                    # This ensures each test is followed by its checks
                    print(f"Creating documentation sheet with {len(docs_data)} entries")
                    docs_df = pd.DataFrame(docs_data)
                    
                    # Create a custom sorter for Type column
                    type_order = {'Test': 0, 'Check': 1}
                    docs_df['type_order'] = docs_df['Type'].map(type_order)
                    
                    # Sort first by Test Name, then by type_order
                    docs_df = docs_df.sort_values(['Test Name', 'type_order'])
                    
                    # Remove the temporary sorting column
                    docs_df = docs_df.drop(columns=['type_order'])
                    
                    # Write to Excel
                    docs_df.to_excel(writer, sheet_name='Test Documentation', index=False)
                    
                    print("Documentation sheet created successfully")
                else:
                    print("No documentation data found to include in report")
            
            # Detailed results
            detailed_df = self.format_detailed_results(results)
            detailed_df.to_excel(writer, sheet_name='Detailed Results')
            
            # Apply formatting to all sheets
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                
                # Format header row
                for cell in worksheet[1]:
                    cell.font = openpyxl.styles.Font(size=16)
                    cell.alignment = openpyxl.styles.Alignment(
                        wrap_text=True,
                        vertical='center',
                        horizontal='center'
                    )
                worksheet.row_dimensions[1].height = 40
                
                # Format data rows
                for row in worksheet.iter_rows(min_row=2):
                    for cell in row:
                        cell.font = openpyxl.styles.Font(size=16)
                        
                        # Check if content is a table-like structure
                        if isinstance(cell.value, str) and '\n' in cell.value:
                            cell.alignment = openpyxl.styles.Alignment(
                                wrap_text=True,
                                vertical='top',
                                horizontal='left'
                            )
                            # Calculate row height based on content
                            line_count = cell.value.count('\n') + 1
                            min_height = 20 * line_count  # 20 pixels per line
                            current_height = worksheet.row_dimensions[cell.row].height or 15
                            worksheet.row_dimensions[cell.row].height = max(min_height, current_height)
                        else:
                            # Regular cell formatting
                            if cell.column == 1:  # Column A
                                cell.alignment = openpyxl.styles.Alignment(
                                    wrap_text=True,
                                    vertical='top',
                                    horizontal='right'
                                )
                            else:
                                cell.alignment = openpyxl.styles.Alignment(
                                    wrap_text=True,
                                    vertical='top',
                                    horizontal='left'
                                )
                
                # Adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    for cell in column:
                        try:
                            if cell.value:
                                # For table-like content, consider the longest line
                                if isinstance(cell.value, str) and '\n' in cell.value:
                                    lines = cell.value.split('\n')
                                    max_length = max(len(line) for line in lines)
                                else:
                                    max_length = max(max_length, len(str(cell.value)))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
                    worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

def main(db_name=None):
    try:
        # Initialize database connection
        db = AccessibilityDB(db_name=db_name)
        
        # Analyze database structure
        print("Analyzing database structure...")
        analyzer = TemplateAnalyzer(db)
        structures, examples = analyzer.analyze_test_structures()
        analyzer.print_analysis()
        
        # Generate report
        print("\nGenerating accessibility report...")
        generator = AccessibilityReportGenerator(db, structures, examples)
        
        # Get regular test run IDs
        test_run_ids = db.get_all_test_run_ids()
        
        # Also add documentation test run if it exists
        doc_test_run = db.get_test_run_by_name("Documentation Test Run")
        if doc_test_run:
            doc_test_run_id = str(doc_test_run['_id'])
            print(f"Including Documentation Test Run: {doc_test_run_id}")
            # Add documentation test run ID if not already in list
            if doc_test_run_id not in test_run_ids:
                test_run_ids.append(doc_test_run_id)
                
            # Directly get the example.com test result and process its documentation
            example_result = db.db.page_results.find_one({"url": "https://example.com"})
            if example_result:
                print("Found example.com test result - processing documentation directly")
                # Process the documentation in advance
                accessibility_tests = example_result.get('results', {}).get('accessibility', {}).get('tests', {})
                if 'page_structure' in accessibility_tests and 'page_structure' in accessibility_tests['page_structure']:
                    doc = accessibility_tests['page_structure']['page_structure'].get('documentation')
                    if doc:
                        print(f"Pre-loading Page Structure documentation with {len(doc.get('tests', []))} individual tests")
                        generator.test_documentation['page_structure'] = doc
                
                if 'accessible_names' in accessibility_tests and 'accessible_names' in accessibility_tests['accessible_names']:
                    doc = accessibility_tests['accessible_names']['accessible_names'].get('documentation')
                    if doc:
                        print(f"Pre-loading Accessible Names documentation with {len(doc.get('tests', []))} individual tests")
                        generator.test_documentation['accessible_names'] = doc
                
                if 'focus_management' in accessibility_tests:
                    # Debug output to inspect the focus_management data structure
                    print("  Found focus_management in tests, examining structure...")
                    fm_data = accessibility_tests['focus_management']
                    
                    # Option 1: Nested structure with focus_management.focus_management.documentation
                    if isinstance(fm_data, dict) and 'focus_management' in fm_data and isinstance(fm_data['focus_management'], dict):
                        print("  Checking for nested focus_management > focus_management > documentation structure")
                        if 'documentation' in fm_data['focus_management']:
                            doc = fm_data['focus_management']['documentation']
                            print(f"  Found nested documentation with {len(doc.get('tests', []))} individual tests")
                            generator.test_documentation['focus_management'] = doc
                    
                    # Option 2: Direct documentation in focus_management
                    elif isinstance(fm_data, dict) and 'documentation' in fm_data:
                        print("  Checking for direct focus_management > documentation structure")
                        doc = fm_data['documentation']
                        print(f"  Found direct documentation with {len(doc.get('tests', []))} individual tests")
                        generator.test_documentation['focus_management'] = doc
                    
                    # Option 3: Focus management as direct document  
                    elif isinstance(fm_data, dict):
                        print("  Dumping focus_management keys for diagnosis: " + str(list(fm_data.keys())))
        
        if test_run_ids:
            print(f"Generating report for {len(test_run_ids)} test runs")
            output_file = 'accessibility_report.xlsx'
            generator.generate_excel_report(test_run_ids, output_file, db_name=db.db_name)
            
            # Determine the actual filename after possible db_name prepending
            if db.db_name != DEFAULT_DB_NAME:
                base, ext = output_file.rsplit('.', 1)
                actual_file = f"{db.db_name}_{base}.{ext}"
            else:
                actual_file = output_file
                
            print(f"Report generated successfully: {actual_file}")
        else:
            print("No test runs found in the database")

    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Generate Excel report from MongoDB accessibility test results')
    parser.add_argument('--database', '-db', help='MongoDB database name to use (default: accessibility_tests)')
    args = parser.parse_args()
    
    main(db_name=args.database)