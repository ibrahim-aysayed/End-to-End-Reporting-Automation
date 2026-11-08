import win32com.client
import os
import re
from datetime import date

# ==================================================
# 1. SETTINGS
# ==================================================

EMAIL_ACCOUNT = "Ibrahim_AySayed@rayacx.com"

REPORTS = {
    "LEC Ticket Report - Daily": (
        r"C:\MAF_Reports_2026\data_sourse",
        "MAFLEC"
    ),

    "MAF-E : FCR Ticket Report - Daily": (
        r"C:\MAF_Reports_2026\data_sourse",
        "FCR"
    ),

    "Fintech Ticket Report - Daily": (
        r"C:\MAF_Reports_2026\data_sourse",
        "FTK"
    )
}

ALLOWED_EXTENSIONS = (".xlsx", ".xls", ".csv")

# ==================================================
# 2. NORMALIZE SUBJECT
# ==================================================

def normalize_subject(subject):
    """
    Remove FW:/Fwd: prefixes and normalize spaces.
    This supports subjects with or without forwarding prefixes.
    """

    subject = subject.strip()

    # Remove one or more FW/Fwd prefixes
    subject = re.sub(
        r"^(?:(?:FW|FWD)\s*:\s*)+",
        "",
        subject,
        flags=re.IGNORECASE
    )

    # Normalize repeated spaces
    subject = re.sub(r"\s+", " ", subject)

    return subject.strip().casefold()


# ==================================================
# 3. CONNECT TO OUTLOOK
# ==================================================

try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    print("Connected to Outlook successfully.")

except Exception as error:
    print(f"Could not connect to Outlook: {error}")
    raise SystemExit(1)


# ==================================================
# 4. FIND ME FOLDER
# ==================================================

me_folder = None

for store in namespace.Stores:

    if store.DisplayName.strip().lower() == EMAIL_ACCOUNT.lower():

        root_folder = store.GetRootFolder()

        for folder in root_folder.Folders:

            if folder.Name.strip().lower() == "me":
                me_folder = folder
                break

        break

if me_folder is None:
    print("ERROR: Could not find the ME folder.")
    raise SystemExit(1)

print(f"Connected to: {me_folder.FolderPath}")


# ==================================================
# 5. PREPARE REPORT SUBJECTS
# ==================================================

subject_lookup = {
    normalize_subject(subject): (subject, path, filename)
    for subject, (path, filename) in REPORTS.items()
}

today = date.today()
date_text = today.strftime("%Y-%m-%d")

downloaded_count = 0
matched_count = 0


# ==================================================
# 6. READ EMAILS
# ==================================================

messages = me_folder.Items
messages.Sort("[ReceivedTime]", True)

for message in messages:

    try:

        # Only process Outlook email messages
        if message.Class != 43:
            continue

        received_date = message.ReceivedTime.date()

        # Process today's emails only
        if received_date < today:
            break

        original_subject = message.Subject or ""
        cleaned_subject = normalize_subject(original_subject)

        # Check whether the email matches one of our 3 reports
        if cleaned_subject not in subject_lookup:
            continue

        report_subject, save_folder, filename_prefix = (
            subject_lookup[cleaned_subject]
        )

        matched_count += 1

        print("\n----------------------------------------")
        print("Matched report:", report_subject)
        print("Email subject:", original_subject)

        # Create destination folder
        os.makedirs(save_folder, exist_ok=True)

        # Process attachments
        for attachment in message.Attachments:

            original_filename = attachment.FileName
            extension = os.path.splitext(original_filename)[1].lower()

            if extension not in ALLOWED_EXTENSIONS:
                print("Skipped unsupported attachment:", original_filename)
                continue

            # Fixed filename = short name + extension (no date)
            new_filename = f"{filename_prefix}{extension}"
            save_path = os.path.join(save_folder, new_filename)

            # Remove any old version of this report (any allowed extension)
            file_locked = False
            for old_ext in ALLOWED_EXTENSIONS:
                old_path = os.path.join(save_folder, f"{filename_prefix}{old_ext}")

                if os.path.exists(old_path):
                    try:
                        os.remove(old_path)
                        print("Removed old file:", old_path)
                    except PermissionError:
                        print("ERROR: File is open, close it first:", old_path)
                        file_locked = True

            if file_locked:
                continue

            attachment.SaveAsFile(save_path)

            downloaded_count += 1

            print("Downloaded:", save_path)

    except Exception as error:
        print("Error processing email:", error)


# ==================================================
# 7. SUMMARY
# ==================================================

print("\n========================================")
print("PROCESS COMPLETED")
print("Date:", date_text)
print("Matched emails:", matched_count)
print("Attachments downloaded:", downloaded_count)
print("========================================")