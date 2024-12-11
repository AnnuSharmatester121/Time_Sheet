from playwright.sync_api import sync_playwright
import pandas as pd

# Load the Excel file
file_path = r"C:\Users\annu.sharma\Desktop\playwright\timesheet_data.xlsx"
all_sheets = pd.read_excel(file_path, sheet_name=None)  # Load all sheets into a dictionary
print("Sheets found:", all_sheets.keys())

def scrape_with_excel_data(data, sheet_name):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # Use headless=False for debugging
        context = browser.new_context(ignore_https_errors=True)  # Ignore SSL issues
        page = context.new_page()

        for index, row in data.iterrows():
            try:
                # Extract and convert data
                Username = str(row['Username']).strip()
                Password = str(row['Password']).strip()
                TaskDate = row['TaskDate'].strftime('%Y-%m-%d')  # If not pd.isna(row['TaskDate'])
                DepartmentID = str(row['DepartmentID']).strip()
                TaskID = str(row['TaskID']).strip()
                Hours = str(row['Hours']).strip()
                Description = str(row['Description']).strip()

                print(f"Processing user: {Username} in sheet: {sheet_name}")

                # Navigate to the login page
                page.goto("https://intranet.ifciltd.com/", timeout=60000)
                page.get_by_placeholder("Windows/Laptop Username").fill(Username)
                page.get_by_placeholder("Windows/Laptop Password").fill(Password)
                page.get_by_role("button", name=" Login").click()

                # Navigate to the Time Sheet page
                page.get_by_role("link", name=" Time Sheet", exact=True).click()
                page.get_by_role("button", name="Create").click()

                # Fill out the form
                page.locator("#task_date").fill(TaskDate)
                page.locator("#dept_id0").select_option(DepartmentID)
                page.locator("#task_id0").select_option(TaskID)
                page.get_by_placeholder("Enter a number").fill(Hours)
                page.get_by_role("cell", name="").get_by_placeholder("Description...").fill(Description)

                # Submit the form
                page.get_by_role("button", name=" Submit").click()
                page.get_by_role("button", name="Ok").click()

                # Optional delay to observe interactions
                page.wait_for_timeout(6000)

            except Exception as e:
                print(f"Error processing row {index}, Username: {Username}. Error: {e}")

        # Close browser context and browser
        context.close()
        browser.close()

# Process each sheet/tab in the Excel file
for sheet_name, sheet_data in all_sheets.items():
    print(f"Processing sheet: {sheet_name}")
    scrape_with_excel_data(sheet_data, sheet_name)
