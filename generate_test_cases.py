import google.generativeai as genai
import pandas as pd
import os
import json
import re
import xlsxwriter
import argparse
import sys
from datetime import datetime
# Configure Gemini API

api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    raise ValueError("❌ API key not found. Set GEMINI_API_KEY as an environment variable.")

genai.configure(api_key=api_key)

# Parse command-line arguments
parser = argparse.ArgumentParser(description="Generate test cases from a user stories file.")
parser.add_argument("input_file", help="Path to the user stories input file")
parser.add_argument(
    "output_file",
    nargs="?",
    default=f"Test_Cases_{datetime.now().strftime('%d%m%Y%H%M%S')}.xlsx",
    help="Path to the output Excel file (must have .xlsx extension, optional, default: Test_Cases_DDMMYYYYHHmmss.xlsx)"
)
parser.add_argument(
    "system_info_file",
    nargs="?",
    default=None,
    help="Optional system information text file to improve test case relevance."
)
args = parser.parse_args()

# Validate output file extension
if not args.output_file.lower().endswith(".xlsx"):
    print("❌ Error: Output file must have a .xlsx extension.")
    sys.exit(1)

INPUT_FILE = args.input_file
OUTPUT_FILE = args.output_file

# Define column headers
HEADERS = [
    "Name", "Objective", "Precondition", "Test Script (Step-by-Step) - Step",
    "Test Script (Step-by-Step) - Test Data", "Test Script (Step-by-Step) - Expected Result", "Coverage (Issues)", "Status"
]

# Function to read additional info from the text file
def read_system_info(system_info_file):
    """ Reads system context from a text file. """
    if system_info_file and os.path.exists(system_info_file):
        try:
            with open(system_info_file, "r", encoding="utf-8") as file:
                return file.read().strip()
        except Exception as e:
            print(f"❌ Error reading system information file: {e}")
    return "No additional system context provided."

# Function to generate test cases using Gemini API
def generate_test_cases(user_story, system_info):
    prompt = f"""Generate detailed test cases for the following user story:
{user_story}
Please return the output in **strict JSON format**, with no extra text, using the structure below:
Output the test cases in the following structured format:

{{
    'Title': 'Verify Login Functionality',
    'Objective': 'Verify that a user can log in with valid credentials',
    'Precondition': 'User must be registered and have valid credentials',
    'Test steps': [
        {{
            'Step': '1. Enter valid username and password',
            'Test Data': 'Username: testuser, Password: testpass',
            'Expected Result': 'User is logged in successfully'
        }},
        {{
            'Step': '2. Enter invalid username and password',
            'Test Data': 'Username: invaliduser, Password: invalidpass',
            'Expected Result': 'Error message is displayed'
        }}
    ],
    'Coverage': 'Requirement ID: LOGIN-001',
    'Status': 'Draft'
}}
Ensure that the response is **valid JSON**, without any markdown formatting, explanations, or additional text.
Use the following additional information as an aid to write the steps in the test cases and not the full test cases. 
Additional System Context:
{system_info}"""

    try:
        response = genai.GenerativeModel("gemini-2.0-flash-001").generate_content(prompt)
        return response.text if response else None
    except Exception as e:
        print(f"❌ API Error: {e}")
        return None

# Function to transform test case data into a structured list
def parse_test_cases(response_text, user_story_id):
    """
    Parses the JSON response from Gemini and extracts test cases.
    Handles both a list of test cases and a single test case.
    """
    if not response_text:
        print("❌ No response received from Gemini.")
        return []

    # Remove Markdown-style code blocks if they exist
    response_text = re.sub(r"```(?:json)?\n(.*?)\n```", r"\1", response_text, flags=re.S).strip()

    try:
        response_json = json.loads(response_text)  # Convert string to JSON
    except json.JSONDecodeError:
        print("❌ Failed to parse JSON response. Raw response:\n", response_text)
        return []

    test_cases = []

    # Ensure response is a list (handling single test case scenario)
    if isinstance(response_json, dict):
        response_json = [response_json]

    if not isinstance(response_json, list):
        print("⚠️ Unexpected response format. Raw response:\n", response_text)
        return []

    for case in response_json:
        title = case.get("Title", "N/A")
        objective = case.get("Objective", "N/A")
        precondition = case.get("Precondition", "N/A")
        coverage = f"{user_story_id}"  # Store User Story ID in Coverage
        status = case.get("Status", "Draft")

        steps = case.get("Test steps", [])
        first_step = True

        for step in steps:
            test_cases.append({
                "Name": title if first_step else "",
                "Objective": objective if first_step else "",
                "Precondition": precondition if first_step else "",
                "Test Script (Step-by-Step) - Step": step.get("Step", "N/A"),
                "Test Script (Step-by-Step) - Test Data": step.get("Test Data", "N/A"),
                "Test Script (Step-by-Step) - Expected Result": step.get("Expected Result", "N/A"),
                "Coverage (Issues)": coverage if first_step else "",
                "Status": status if first_step else ""
            })
            first_step = False  # Ensure only the first step gets the merged columns

    return test_cases

# Function to save test cases to an Excel file
def save_to_excel(test_cases):
    if not test_cases:
        print("⚠️ No test cases generated. Excel file was not created.")
        return

    df = pd.DataFrame(test_cases, columns=HEADERS)
    
    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Test Cases", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Test Cases"]
        
        # Adjust column widths for better readability
        for col_num, col_name in enumerate(HEADERS):
            worksheet.set_column(col_num, col_num, max(15, len(col_name) + 5))
        
        # Merge cells for columns: Title, Objective, Precondition, Coverage, Status
        merge_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 0, 'text_wrap': True})
        row = 1  # Start from row 1 (since row 0 is headers)
        while row < len(df):
            start_row = row

            # Identify the range for merging by finding consecutive empty title rows
            while row < len(df) and df.loc[row, "Name"] == "":
                row += 1

            if row - start_row > 1:  # Merge if there are multiple steps
                for col in ["Name", "Objective", "Precondition", "Coverage (Issues)", "Status"]:
                    col_idx = df.columns.get_loc(col)
                    value = df.at[start_row, col]
                    worksheet.merge_range(start_row + 1, col_idx, row, col_idx, value, merge_format)

            row += 1  # Move to the next test case
    
    print(f"✅ Test cases saved to {OUTPUT_FILE}")

def read_user_stories():

    """
    Reads user stories from an Excel file.
    Extracts 'User Story ID' and 'User Story' columns.
    Returns a list of tuples (User Story ID, User Story text).
    """
    if not os.path.exists(INPUT_FILE):
        print(f"⚠️ Input file '{INPUT_FILE}' not found.")
        return []

    try:
        df = pd.read_excel(INPUT_FILE, dtype=str)  # Read the Excel file as text
    except Exception as e:
        print(f"❌ Error reading the Excel file: {e}")
        return []

    # Validate required columns exist
    required_columns = {"User Story ID", "User Story"}
    if not required_columns.issubset(df.columns):
        print(f"❌ Error: Input file must have the columns: {required_columns}")
        return []

    # Convert DataFrame to list of tuples
    user_stories = list(df[["User Story ID", "User Story"]].itertuples(index=False, name=None))

    return user_stories

# Main function
def main():
    if not os.path.exists(INPUT_FILE):
        print(f"⚠️ Input file '{INPUT_FILE}' not found.")
        return

#    with open(INPUT_FILE, "r", encoding="utf-8") as file:
#        user_stories = file.readlines()

    all_test_cases = []
    user_stories = read_user_stories()
    system_info = read_system_info(args.system_info_file)

    for user_story_id, story_text in user_stories:
        print(f"\n🔹 Generating test cases for: {user_story_id}\n")
        response = generate_test_cases(story_text, system_info)
        #print("\nResponse from Gemini:\n", response)

        if not response:
            print("⚠️ No response received from Gemini.")
            continue

        test_cases = parse_test_cases(response, user_story_id)
        #print("\nParsed Test Cases:\n", test_cases)

        all_test_cases.extend(test_cases)

    save_to_excel(all_test_cases)

# Run the script
if __name__ == "__main__":
    main()
