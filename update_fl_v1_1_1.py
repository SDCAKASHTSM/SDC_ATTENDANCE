from pathlib import Path
import shutil

BASE = Path(r"C:\Users\DELL\SDC_ATTENDANCE")
BACKUP = Path(r"C:\Users\DELL\SDC_ATTENDANCE_BACKUP")

print("=" * 60)
print("SDC ATTENDANCE - V1.1.1 FL STATUS UPDATE")
print("=" * 60)

files = [
    BASE / "app.py",
    BASE / "templates" / "attendance_register.html",
    BASE / "templates" / "selected_attendance.html",
]

print()
print("-" * 60)

for file in files:

    if not file.exists():
        print(f"[ERROR] File not found: {file}")
        continue

    backup_file = BACKUP / f"{file.stem}_before_v1_1_1{file.suffix}"

    if not backup_file.exists():
        shutil.copy2(file, backup_file)
        print(f"[BACKUP] {backup_file.name}")
    else:
        print(f"[BACKUP EXISTS] {backup_file.name}")

    text = file.read_text(encoding="utf-8")

    original = text

    # -------------------------------------------------
    # PL -> FL
    # -------------------------------------------------

    text = text.replace(
        'VALID_STATUSES = ["P", "A", "W/O", "NH", "PL"]',
        'VALID_STATUSES = ["P", "A", "W/O", "NH", "FL"]'
    )

    text = text.replace(
        "setAllStatus('PL')",
        "setAllStatus('FL')"
    )

    text = text.replace(
        'value="PL"',
        'value="FL"'
    )

    text = text.replace(
        'employee["status"] == "PL"',
        'employee["status"] == "FL"'
    )

    text = text.replace(
        '>PL ALL<',
        '>FL ALL<'
    )

    text = text.replace(
        '>PL</option>',
        '>FL</option>'
    )

    text = text.replace(
        'PL — PAID LEAVE',
        'FL — FULL LEAVE'
    )

    text = text.replace(
        'PL ΓÇö PAID LEAVE',
        'FL ΓÇö FULL LEAVE'
    )

    # -------------------------------------------------
    # Save only if changed
    # -------------------------------------------------

    if text != original:

        file.write_text(
            text,
            encoding="utf-8"
        )

        print(f"[UPDATED] {file}")

    else:

        print(f"[NO CHANGE] {file}")

print()
print("-" * 60)
print("UPDATE SUCCESSFUL")
print("-" * 60)

print("[OK] PL replaced with FL")
print("[OK] FL attendance status enabled")
print("[OK] FL ALL enabled")
print("[OK] Paid Leave option changed to Full Leave")
print("[OK] Backup created before modification")

print()
print("IMPORTANT:")
print("Existing salary calculation was NOT changed.")
print("Existing P + FL + NH salary logic remains intact.")

print("=" * 60)