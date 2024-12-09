import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from playwright.sync_api import sync_playwright

# Path to the Excel file
file_path = r"C:\Users\annu.sharma\Desktop\playwright\timesheet_data.xlsx"

# Function to handle a single user's timesheet
def process_user_timesheet(sheet_name, sheet_data, browser_type):
    with sync_playwright() as p:
        # Launch the specific browser type
        if browser_type == "chromium":
            browser = p.chromium.launch(headless=False)
        elif browser_type == "firefox":
            browser = p.firefox.launch(headless=False)
        elif browser_type == "webkit":
            browser = p.webkit.launch(headless=False)
        else:
            print(f"Unknown browser type: {browser_type}")
            return

        context = browser.new_context(ignore_https_errors=True)  # Ignore SSL issues
        page = context.new_page()

        try:
            print(f"Processing sheet: {sheet_name} using {browser_type}")

            for _, row in sheet_data.iterrows():
                # Extract user details from the row
                username = str(row['Username']).strip()
                password = str(row['Password']).strip()
                task_date = row['TaskDate'].strftime('%Y-%m-%d')
                department_id = str(row['DepartmentID']).strip()
                task_id = str(row['TaskID']).strip()
                hours = str(row['Hours']).strip()
                description = str(row['Description']).strip()

                print(f"Processing user: {username} from sheet: {sheet_name} on {browser_type}")

                # Navigate to the login page
                page.goto("https://intranet.ifciltd.com/", timeout=60000)
                page.get_by_placeholder("Windows/Laptop Username").fill(username)
                page.get_by_placeholder("Windows/Laptop Password").fill(password)
                page.get_by_role("button", name=" Login").click()

                # Navigate to the Time Sheet page
                page.get_by_role("link", name=" Time Sheet", exact=True).click()
                page.get_by_role("button", name="Create").click()

                # Fill out the timesheet form
                page.locator("#task_date").fill(task_date)
                page.locator("#dept_id0").select_option(department_id)
                page.locator("#task_id0").select_option(task_id)
                page.get_by_placeholder("Enter a number").fill(hours)
                page.get_by_role("cell", name="").get_by_placeholder("Description...").fill(description)

                # Submit the timesheet
                page.get_by_role("button", name=" Submit").click()
                page.get_by_role("button", name="Ok").click()

                # Optional delay to observe interactions
                page.wait_for_timeout(6000)

        except Exception as e:
            print(f"Error processing sheet {sheet_name} on {browser_type}: {e}")

        finally:
            # Close the browser
            context.close()
            browser.close()

# Load all sheets from the Excel file
all_sheets_data = pd.read_excel(file_path, sheet_name=None)

# List of browsers to use
browser_types = ["chromium", "firefox", "webkit"]

# Use ThreadPoolExecutor to process each tab with a different browser
with ThreadPoolExecutor(max_workers=len(browser_types)) as executor:
    for i, (sheet_name, sheet_data) in enumerate(all_sheets_data.items()):
        if sheet_data.empty:
            print(f"Skipping empty sheet: {sheet_name}")
            continue

        # Assign a browser type for each sheet in a round-robin manner
        browser_type = browser_types[i % len(browser_types)]
        executor.submit(process_user_timesheet, sheet_name, sheet_data, browser_type)

print("All timesheets processed.")

 