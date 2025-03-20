#!/usr/bin/env python3
"""
Script to add documentation directly to MongoDB without running full tests.
This script will:
1. Connect to the MongoDB database
2. Get the most recent test run
3. Add documentation for animations, colors, and forms directly to that test run
4. Save the changes to the database
"""
from pymongo import MongoClient
from bson.objectid import ObjectId
import sys
import os
import json

# Documentation for animations
ANIMATIONS_DOCUMENTATION = {
    "testName": "CSS Animations Analysis",
    "description": "Evaluates CSS animations on the page for accessibility considerations, focusing on prefers-reduced-motion media query support, infinite animations, and animation duration. This test helps ensure content is accessible to users who may experience motion sickness, vestibular disorders, or other conditions affected by movement on screen.",
    "version": "1.0.0",
    "date": "2025-03-19",
    "dataSchema": {
        "pageFlags": "Boolean flags indicating key animation issues",
        "details": "Full animation data including keyframes, animated elements, and media queries",
        "timestamp": "ISO timestamp when the test was run"
    },
    "tests": [
        {
            "id": "animations-reduced-motion",
            "name": "Reduced Motion Support",
            "description": "Checks if the page provides prefers-reduced-motion media query support for users who have indicated they prefer reduced motion in their system settings.",
            "impact": "high",
            "wcagCriteria": ["2.3.3"],
            "howToFix": "Add a prefers-reduced-motion media query in your CSS to disable or reduce animations for users who prefer reduced motion:\n@media (prefers-reduced-motion: reduce) {\n  * {\n    animation-duration: 0.001ms !important;\n    animation-iteration-count: 1 !important;\n    transition-duration: 0.001ms !important;\n  }\n}",
            "resultsFields": {
                "pageFlags.lacksReducedMotionSupport": "Indicates if animations are present without prefers-reduced-motion support",
                "pageFlags.details.hasReducedMotionSupport": "Indicates if prefers-reduced-motion media query is present",
                "details.mediaQueries": "List of media queries related to reduced motion"
            }
        },
        {
            "id": "animations-infinite",
            "name": "Infinite Animations",
            "description": "Identifies animations set to run indefinitely (animation-iteration-count: infinite). These can cause significant accessibility issues for users with vestibular disorders or attention-related disabilities.",
            "impact": "high",
            "wcagCriteria": ["2.2.2", "2.3.3"],
            "howToFix": "Modify infinite animations to either:\n1. Have a defined, short duration (less than 5 seconds)\n2. Only play once (animation-iteration-count: 1)\n3. Add controls to pause animation\n4. Respect prefers-reduced-motion media query",
            "resultsFields": {
                "pageFlags.hasInfiniteAnimations": "Indicates if page contains infinite animations",
                "pageFlags.details.infiniteAnimations": "Count of infinite animations on the page",
                "details.violations": "List of elements with infinite animations"
            }
        },
        {
            "id": "animations-duration",
            "name": "Long Duration Animations",
            "description": "Identifies animations that run for an extended period (over 5 seconds), which can be distracting or disorienting for some users.",
            "impact": "medium",
            "wcagCriteria": ["2.2.2"],
            "howToFix": "Reduce the duration of animations to be under 5 seconds or provide a mechanism to pause, stop, or hide the animation.",
            "resultsFields": {
                "pageFlags.hasLongAnimations": "Indicates if page contains long-running animations",
                "pageFlags.details.longDurationAnimations": "Count of animations exceeding 5 seconds",
                "details.animatedElements": "List of animated elements with their durations"
            }
        }
    ]
}

# Documentation for colors
COLORS_DOCUMENTATION = {
    "testName": "Color and Contrast Analysis",
    "description": "Evaluates color usage and contrast ratios on the page to ensure content is perceivable by users with low vision, color vision deficiencies, or those who use high contrast modes. This test examines text contrast, color-only distinctions, non-text contrast, color references, and adjacent element contrast.",
    "version": "1.0.0",
    "date": "2025-03-19",
    "dataSchema": {
        "pageFlags": "Boolean flags indicating key color and contrast issues",
        "details": "Full color data including contrast measurements and violations",
        "timestamp": "ISO timestamp when the test was run"
    },
    "tests": [
        {
            "id": "color-text-contrast",
            "name": "Text Contrast Ratio",
            "description": "Evaluates the contrast ratio between text color and its background to ensure readability. Normal text should have a contrast ratio of at least 4.5:1, while large text should have a ratio of at least 3:1.",
            "impact": "high",
            "wcagCriteria": ["1.4.3", "1.4.6"],
            "howToFix": "Increase the contrast between text and background colors. Options include:\n1. Darkening the text color (for light backgrounds)\n2. Lightening the text color (for dark backgrounds)\n3. Changing the background color to increase contrast\n4. Using a contrast checker tool to verify ratios meet WCAG standards",
            "resultsFields": {
                "pageFlags.hasContrastIssues": "Indicates if any text has insufficient contrast",
                "pageFlags.details.contrastViolations": "Count of text elements with contrast issues",
                "details.textContrast.violations": "Detailed information about each contrast violation"
            }
        },
        {
            "id": "color-only-distinction",
            "name": "Color-Only Distinctions",
            "description": "Identifies cases where color alone is used to convey information, particularly for links that are distinguished only by color without additional visual cues.",
            "impact": "high",
            "wcagCriteria": ["1.4.1"],
            "howToFix": "Add non-color distinctions to links and interactive elements:\n1. Add underlines to links\n2. Use icons or symbols alongside color\n3. Apply text styles like bold or italic\n4. Add borders or background changes on hover/focus",
            "resultsFields": {
                "pageFlags.hasColorOnlyLinks": "Indicates if links are distinguished only by color",
                "pageFlags.details.colorOnlyLinks": "Count of links with color-only distinction",
                "details.links.violations": "List of links that rely solely on color"
            }
        },
        {
            "id": "color-non-text-contrast",
            "name": "Non-Text Contrast",
            "description": "Evaluates contrast for UI components and graphical objects to ensure they're perceivable by users with low vision.",
            "impact": "medium",
            "wcagCriteria": ["1.4.11"],
            "howToFix": "Ensure UI components (buttons, form controls, focus indicators) and graphics required for understanding have a contrast ratio of at least 3:1 against adjacent colors.",
            "resultsFields": {
                "pageFlags.hasNonTextContrastIssues": "Indicates if non-text elements have contrast issues",
                "pageFlags.details.nonTextContrastViolations": "Count of non-text contrast violations",
                "details.nonText.violations": "List of non-text elements with contrast issues"
            }
        },
        {
            "id": "color-references",
            "name": "Color References in Content",
            "description": "Identifies text that refers to color as the only means of conveying information, which can be problematic for users with color vision deficiencies.",
            "impact": "medium",
            "wcagCriteria": ["1.4.1"],
            "howToFix": "Supplement color references with additional descriptors:\n1. Use patterns, shapes, or labels in addition to color\n2. Add text that doesn't rely on perceiving color\n3. Use 'located at [position]' instead of 'in red'",
            "resultsFields": {
                "pageFlags.hasColorReferences": "Indicates if content references color as an identifier",
                "pageFlags.details.colorReferences": "Count of color references in text",
                "details.colorReferences.instances": "Text fragments containing color references"
            }
        },
        {
            "id": "color-adjacent-contrast",
            "name": "Adjacent Element Contrast",
            "description": "Examines contrast between adjacent UI elements to ensure boundaries are perceivable.",
            "impact": "medium",
            "wcagCriteria": ["1.4.11"],
            "howToFix": "Increase contrast between adjacent elements by:\n1. Adding borders between sections\n2. Increasing the color difference between adjacent elements\n3. Adding visual separators like lines or spacing",
            "resultsFields": {
                "pageFlags.hasAdjacentContrastIssues": "Indicates if adjacent elements lack sufficient contrast",
                "pageFlags.details.adjacentContrastViolations": "Count of adjacent contrast violations",
                "details.adjacentDivs.violations": "List of adjacent elements with insufficient contrast"
            }
        },
        {
            "id": "color-media-queries",
            "name": "Contrast and Color Scheme Preferences",
            "description": "Checks if the page respects user preferences for increased contrast and color schemes through media queries.",
            "impact": "medium",
            "wcagCriteria": ["1.4.12"],
            "howToFix": "Implement media queries to support user preferences:\n@media (prefers-contrast: more) { /* high contrast styles */ }\n@media (prefers-color-scheme: dark) { /* dark mode styles */ }",
            "resultsFields": {
                "pageFlags.supportsContrastPreferences": "Indicates if prefers-contrast media query is used",
                "pageFlags.supportsColorSchemePreferences": "Indicates if prefers-color-scheme media query is used",
                "details.mediaQueries": "Details about detected media queries"
            }
        }
    ]
}

# Documentation for forms
FORMS_DOCUMENTATION = {
    "testName": "Form Accessibility Analysis",
    "description": "Evaluates web forms for accessibility requirements including proper labeling, structure, layout, and contrast. This test helps ensure that forms can be completed by all users, regardless of their abilities or assistive technologies.",
    "version": "1.0.0",
    "date": "2025-03-19",
    "dataSchema": {
        "pageFlags": "Boolean flags indicating key form accessibility issues",
        "details": "Full form data including structure, inputs, and violations",
        "timestamp": "ISO timestamp when the test was run"
    },
    "tests": [
        {
            "id": "form-landmark-context",
            "name": "Form Landmark Context",
            "description": "Checks if forms are properly placed within appropriate landmark regions. Forms should typically be within main content areas and not directly in header or footer unless they are simple search or subscription forms.",
            "impact": "medium",
            "wcagCriteria": ["1.3.1", "2.4.1"],
            "howToFix": "Place forms within the appropriate landmark regions:\n1. Use <main> for primary content forms\n2. Use appropriate ARIA landmarks for specialized forms\n3. Only place search forms in <header> or forms directly related to footer content in <footer>",
            "resultsFields": {
                "pageFlags.hasFormsOutsideLandmarks": "Indicates if any forms are outside appropriate landmarks",
                "pageFlags.details.formsOutsideLandmarks": "Count of forms outside landmarks",
                "details.forms[].location": "Details of each form's landmark context"
            }
        },
        {
            "id": "form-headings",
            "name": "Form Headings and Identification",
            "description": "Checks if forms are properly identified with headings and ARIA labeling to make their purpose clear to all users.",
            "impact": "high",
            "wcagCriteria": ["1.3.1", "2.4.6"],
            "howToFix": "Add proper headings to forms:\n1. Include a heading (h2-h6) that describes the form's purpose\n2. Connect the heading to the form using aria-labelledby\n3. Ensure the heading text clearly describes the form's purpose",
            "resultsFields": {
                "pageFlags.hasFormsWithoutHeadings": "Indicates if forms lack proper headings",
                "pageFlags.details.formsWithoutHeadings": "Count of forms without proper headings",
                "details.forms[].heading": "Details of each form's heading information"
            }
        },
        {
            "id": "form-input-labels",
            "name": "Input Field Labeling",
            "description": "Verifies that all form inputs have properly associated labels that clearly describe the purpose of each field.",
            "impact": "critical",
            "wcagCriteria": ["1.3.1", "3.3.2", "4.1.2"],
            "howToFix": "Implement proper labels for all form fields:\n1. Use <label> elements with a 'for' attribute that matches the input's 'id'\n2. Ensure labels are visible and descriptive\n3. For complex controls, use aria-labelledby or aria-label if needed\n4. Never rely on placeholders alone for labeling",
            "resultsFields": {
                "pageFlags.hasInputsWithoutLabels": "Indicates if any inputs lack proper labels",
                "pageFlags.details.inputsWithoutLabels": "Count of inputs without labels",
                "details.forms[].inputs[].label": "Details of each input's label"
            }
        },
        {
            "id": "form-placeholder-misuse",
            "name": "Placeholder Text Misuse",
            "description": "Identifies form fields that use placeholder text as a substitute for proper labels, which creates accessibility barriers.",
            "impact": "high",
            "wcagCriteria": ["1.3.1", "3.3.2"],
            "howToFix": "Address placeholder-only fields:\n1. Add proper labels for all inputs\n2. Use placeholders only for format examples, not as instruction text\n3. Ensure placeholder text has sufficient contrast\n4. Keep placeholder text concise and helpful",
            "resultsFields": {
                "pageFlags.hasPlaceholderOnlyInputs": "Indicates if any inputs use placeholders as labels",
                "pageFlags.details.inputsWithPlaceholderOnly": "Count of inputs with placeholder-only labeling",
                "details.violations": "List of inputs with placeholder misuse"
            }
        },
        {
            "id": "form-layout-issues",
            "name": "Form Layout and Structure",
            "description": "Examines the layout of form elements to ensure fields are properly positioned for logical completion and navigation.",
            "impact": "medium",
            "wcagCriteria": ["1.3.1", "1.3.2", "3.3.2"],
            "howToFix": "Improve form layout:\n1. Position labels above or to the left of their fields\n2. Left-align labels with their fields\n3. Avoid multiple fields on the same line (except for related fields like city/state/zip)\n4. Group related fields with fieldset and legend elements\n5. Arrange fields in a logical sequence",
            "resultsFields": {
                "pageFlags.hasLayoutIssues": "Indicates if forms have layout problems",
                "pageFlags.details.inputsWithLayoutIssues": "Count of inputs with layout issues",
                "details.forms[].layoutIssues": "Details of layout problems in each form"
            }
        },
        {
            "id": "form-input-contrast",
            "name": "Form Control Contrast",
            "description": "Checks that form controls, labels, and placeholder text have sufficient contrast against their backgrounds.",
            "impact": "high",
            "wcagCriteria": ["1.4.3", "1.4.11"],
            "howToFix": "Improve form control contrast:\n1. Ensure form input text has at least 4.5:1 contrast ratio\n2. Ensure form borders have at least 3:1 contrast ratio\n3. Ensure placeholder text has at least 4.5:1 contrast ratio\n4. Use visual indicators for focus states with 3:1 contrast ratio",
            "resultsFields": {
                "pageFlags.hasContrastIssues": "Indicates if form controls have contrast issues",
                "pageFlags.details.inputsWithContrastIssues": "Count of inputs with contrast problems",
                "details.forms[].inputs[].contrast": "Contrast measurements for each input"
            }
        }
    ]
}

def main(db_name=None):
    """Main function to add documentation to the database"""
    print("Connecting to MongoDB...")
    client = MongoClient('mongodb://localhost:27017/')
    
    # Use the specified database name or default
    if db_name is None:
        db_name = 'accessibility_tests'
        print(f"Warning: No database name specified. Using default database '{db_name}'.")
    
    print(f"Using database: {db_name}")
    db = client[db_name]
    
    # Get the most recent test run
    latest_run = db.test_runs.find_one(sort=[('timestamp_start', -1)])
    
    if not latest_run:
        print("No test runs found in the database.")
        return
    
    test_run_id = latest_run['_id']
    print(f"Found latest test run with ID: {test_run_id}")
    
    # Check if documentation already exists
    if 'documentation' in latest_run:
        print("Test run already has documentation field.")
        documentation = latest_run['documentation']
    else:
        print("Creating new documentation field.")
        documentation = {}
    
    # Add the documentation
    documentation['animations'] = ANIMATIONS_DOCUMENTATION
    documentation['colors'] = COLORS_DOCUMENTATION
    documentation['forms'] = FORMS_DOCUMENTATION
    
    # Update the test run
    result = db.test_runs.update_one(
        {'_id': test_run_id},
        {'$set': {'documentation': documentation}}
    )
    
    if result.modified_count > 0:
        print("Documentation added successfully!")
        print("Updated documentation keys:", list(documentation.keys()))
    else:
        print("No changes made to the database.")
    
    # Now run the XLS generator to create the report
    print("\nGenerating XLS report...")
    if db_name:
        os.system(f'python3 src/xls_generator/xls_generator.py --database "{db_name}"')
    else:
        os.system('python3 src/xls_generator/xls_generator.py')
    
    print("\nDone! Check the documentation sheet in the generated Excel file.")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='Add documentation directly to a MongoDB database')
    parser.add_argument('--database', '-db', help='MongoDB database name to use (default: accessibility_tests)')
    args = parser.parse_args()
    
    main(db_name=args.database)