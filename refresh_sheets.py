import win32com.client as win32
import os

# حط مسارات الملفات الكاملة هنا
files = [
    r"C:\MAF_Reports_2026\data_sheets\SCH & Forecasting -All BUs.xlsx",
    r"C:\MAF_Reports_2026\data_sheets\Genyses Raw Data..xlsx",
]

excel = win32.DispatchEx("Excel.Application")
excel.Visible = False          # خليها True لو عاوز تشوف إكسيل وهو شغال
excel.DisplayAlerts = False

try:
    for path in files:
        path = os.path.abspath(path)
        print(f"فتح: {path}")
        wb = excel.Workbooks.Open(path)

        # لازم نقفل الـ Background Refresh عشان الكود يستنى الريفرش يخلص
        for conn in wb.Connections:
            try:
                conn.OLEDBConnection.BackgroundQuery = False
            except Exception:
                pass
            try:
                conn.ODBCConnection.BackgroundQuery = False
            except Exception:
                pass

        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()  # يستنى لحد ما كل الكويريز تخلص

        wb.Save()
        wb.Close(SaveChanges=True)
        print(f"تم التحديث والحفظ: {path}")
finally:
    excel.Quit()

print("✅ خلصنا")