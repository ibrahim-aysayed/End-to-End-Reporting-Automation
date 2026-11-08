
import win32com.client
import time
import os

# ==========================================
# Dashboard Files
# ==========================================

dashboard_files = [
    r"C:\MAF_Reports_2026\DBs\DBU Dashbaord.xlsx",

    r"C:\MAF_Reports_2026\DBs\MAF Fintech Dashboard.xlsx",

    r"C:\MAF_Reports_2026\DBs\MAF Entertainment Dashboard.xlsx"
]

# ==========================================
# Check Files
# ==========================================

for file_path in dashboard_files:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

# ==========================================
# Open Excel
# ==========================================

excel = win32com.client.DispatchEx("Excel.Application")

excel.Visible = False
excel.DisplayAlerts = False
excel.AskToUpdateLinks = False

successful_files = []
failed_files = []

try:

    total_files = len(dashboard_files)

    # ==========================================
    # Refresh Dashboards One by One
    # ==========================================

    for index, file_path in enumerate(dashboard_files, start=1):

        file_name = os.path.basename(file_path)

        print("\n" + "=" * 55)
        print(f"Starting Dashboard {index} of {total_files}")
        print(f"File: {file_name}")
        print("=" * 55)

        workbook = None

        try:

            # Open Dashboard
            workbook = excel.Workbooks.Open(file_path)

            print(f"[{index}/{total_files}] Dashboard opened successfully.")

            # Refresh All Connections
            workbook.RefreshAll()

            print("Refresh started... Please wait.")

            # Wait for Queries to Finish
            excel.CalculateUntilAsyncQueriesDone()

            # Additional Waiting Time
            time.sleep(10)

            # Full Calculation
            excel.CalculateFull()

            # Save Dashboard
            workbook.Save()

            print(f"Dashboard {index} completed successfully!")

            successful_files.append(file_name)

        except Exception as e:

            print(f"ERROR in Dashboard {index}: {e}")

            failed_files.append(file_name)

        finally:

            # Close Current Dashboard
            if workbook is not None:

                workbook.Close(SaveChanges=True)
                print(f"Dashboard {index} closed.")

        # Move to Next Dashboard
        if index < total_files:

            print("\nWaiting to start the next dashboard...")
            time.sleep(3)

    # ==========================================
    # Final Summary
    # ==========================================

    print("\n" + "=" * 55)
    print("REFRESH PROCESS COMPLETED")
    print("=" * 55)

    print(f"\nSuccessfully Refreshed: {len(successful_files)}")

    for file in successful_files:
        print(f"[DONE] {file}")

    print(f"\nFailed: {len(failed_files)}")

    for file in failed_files:
        print(f"[FAILED] {file}")

finally:

    excel.Quit()

    print("\nExcel closed successfully.")
    print("All processes finished.")
