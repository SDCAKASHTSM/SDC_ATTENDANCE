from pathlib import Path
import shutil

BASE = Path(r"C:\Users\DELL\SDC_ATTENDANCE")
BACKUP = Path(r"C:\Users\DELL\SDC_ATTENDANCE_BACKUP")

files = [
    BASE / "templates" / "attendance_register.html",
    BASE / "templates" / "selected_attendance.html",
]

print("=" * 60)
print("SDC ATTENDANCE - FL UI CLEANUP")
print("=" * 60)

for file in files:

    backup = BACKUP / f"{file.stem}_before_fl_ui{file.suffix}"

    if not backup.exists():
        shutil.copy2(file, backup)

    text = file.read_text(encoding="utf-8")
    original = text

    # CSS class names
    text = text.replace("pl-all", "fl-all")

    # Button text
    text = text.replace("PL ALL", "FL ALL")

    # Visible option text
    text = text.replace(">PL</option>", ">FL</option>")

    # Any remaining PL paid-leave wording
    text = text.replace("PL — PAID LEAVE", "FL — FULL LEAVE")
    text = text.replace("PL ΓÇö PAID LEAVE", "FL ΓÇö FULL LEAVE")

    if text != original:
        file.write_text(text, encoding="utf-8")
        print("[UPDATED]", file)
    else:
        print("[NO CHANGE]", file)

print()
print("-" * 60)
print("FL UI CLEANUP SUCCESSFUL")
print("-" * 60)
print("[OK] PL ALL -> FL ALL")
print("[OK] pl-all -> fl-all")
print("[OK] PL option -> FL")
print("=" * 60)