import win32com.client
import os
import re
from datetime import date

EMAIL_ACCOUNT = "Ibrahim_AySayed@rayacx.com"

SAVE_FOLDER = r"C:\MAF_Reports_2026\data_sourse"

REPORTS = {
    "data_nps": {"mode": "by_size", "names": ("NPSE", "NPSF")},
    "data":     {"mode": "original"},
    "Sales":    {"mode": "single", "name": "Sales"},
}

ALLOWED_EXTENSIONS = (".xlsx", ".xls", ".csv")


def normalize_subject(subject):
    """Remove FW:/Fwd: prefixes and normalize spaces."""
    subject = subject.strip()
    subject = re.sub(
        r"^(?:(?:FW|FWD)\s*:\s*)+",
        "",
        subject,
        flags=re.IGNORECASE
    )
    subject = re.sub(r"\s+", " ", subject)
    return subject.strip().casefold()


def remove_old_versions(base_name):
    """Delete any old file with the same base name (any allowed extension).
    Returns False if a file is locked (open)."""
    ok = True
    for old_ext in ALLOWED_EXTENSIONS:
        old_path = os.path.join(SAVE_FOLDER, f"{base_name}{old_ext}")
        if os.path.exists(old_path):
            try:
                os.remove(old_path)
                print("Removed old file:", old_path)
            except PermissionError:
                print("ERROR: File is open, close it first:", old_path)
                ok = False
    return ok


def save_attachment(attachment, base_name):
    """Save attachment as base_name + its original extension."""
    extension = os.path.splitext(attachment.FileName)[1].lower()
    save_path = os.path.join(SAVE_FOLDER, f"{base_name}{extension}")

    if not remove_old_versions(base_name):
        return False

    attachment.SaveAsFile(save_path)
    print("Downloaded:", save_path)
    return True


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
# 5. PREPARE
# ==================================================

subject_lookup = {
    normalize_subject(subject): (subject, config)
    for subject, config in REPORTS.items()
}

today = date.today()
date_text = today.strftime("%Y-%m-%d")

downloaded_count = 0
matched_count = 0

os.makedirs(SAVE_FOLDER, exist_ok=True)


# ==================================================
# 6. READ EMAILS
# ==================================================

messages = me_folder.Items
messages.Sort("[ReceivedTime]", True)

for message in messages:

    try:
        if message.Class != 43:
            continue

        received_date = message.ReceivedTime.date()

        # Process today's emails only
        if received_date < today:
            break

        original_subject = message.Subject or ""
        cleaned_subject = normalize_subject(original_subject)

        if cleaned_subject not in subject_lookup:
            continue

        report_subject, config = subject_lookup[cleaned_subject]
        mode = config["mode"]

        matched_count += 1

        print("\n----------------------------------------")
        print("Matched report:", report_subject)
        print("Email subject:", original_subject)

        # ---- اجمع المرفقات المسموحة فقط ----
        valid_attachments = []
        for attachment in message.Attachments:
            ext = os.path.splitext(attachment.FileName)[1].lower()
            if ext in ALLOWED_EXTENSIONS:
                valid_attachments.append(attachment)
            else:
                print("Skipped unsupported attachment:", attachment.FileName)

        if not valid_attachments:
            print("No valid attachments found.")
            continue

        # ---------------- MODE: by_size (data_nps) ----------------
        if mode == "by_size":
            big_name, small_name = config["names"]

            if len(valid_attachments) < 2:
                print(f"WARNING: expected 2 attachments, found {len(valid_attachments)}")

            # ترتيب من الأكبر للأصغر حسب الحجم
            valid_attachments.sort(key=lambda a: a.Size, reverse=True)

            names = [big_name, small_name]
            for attachment, base_name in zip(valid_attachments, names):
                print(f"{attachment.FileName} ({attachment.Size / 1024 / 1024:.2f} MB) -> {base_name}")
                if save_attachment(attachment, base_name):
                    downloaded_count += 1

        # ---------------- MODE: original (data) ----------------
        elif mode == "original":
            for attachment in valid_attachments:
                base_name = os.path.splitext(attachment.FileName)[0]
                if save_attachment(attachment, base_name):
                    downloaded_count += 1

        # ---------------- MODE: single (Sales) ----------------
        elif mode == "single":
            if len(valid_attachments) > 1:
                print("WARNING: more than one attachment, using the first one only.")

            attachment = valid_attachments[0]
            if save_attachment(attachment, config["name"]):
                downloaded_count += 1

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