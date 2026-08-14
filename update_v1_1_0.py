from pathlib import Path
import shutil
import sys

APP = Path("app.py")
BACKUP = Path(r"C:\Users\DELL\SDC_ATTENDANCE_BACKUP\app_before_v1_1_0.py")

print("=" * 60)
print("SDC ATTENDANCE - V1.1.0 BACKEND UPDATE")
print("=" * 60)

if not APP.exists():
    print("[ERROR] app.py nahi mila.")
    sys.exit(1)

if not BACKUP.exists():
    print("[ERROR] Safety backup nahi mila.")
    print("Pehle backup create karo.")
    sys.exit(1)

original = APP.read_text(encoding="utf-8-sig")
text = original

changes = []


# =========================================================
# 1. ADD DAILY ATTENDANCE VARIABLES
# =========================================================

old = '''        days = []

        present_days = 0

        ot_hours = 0
'''

new = '''        days = []

        daily_attendance = []

        present_days = 0

        actual_present_days = 0

        fl_days = 0

        nh_days = 0

        ot_hours = 0
'''

if old not in text:
    print("[ERROR] Step 1 pattern nahi mila.")
    sys.exit(1)

text = text.replace(old, new, 1)
changes.append("Daily attendance variables")


# =========================================================
# 2. P + FL + NH PAID DAYS
# =========================================================

old = '''            days.append(status)

            if status == "P":

                present_days += 1

            ot_hours += float(
                ot or 0
            )
'''

new = '''            days.append(status)

            daily_attendance.append({
                "status": status,
                "ot": float(ot or 0)
            })

            if status == "P":

                actual_present_days += 1

            if status == "FL":

                fl_days += 1

            if status == "NH":

                nh_days += 1

            if status in ("P", "FL", "NH"):

                present_days += 1

            ot_hours += float(
                ot or 0
            )
'''

if old not in text:
    print("[ERROR] Step 2 pattern nahi mila.")
    sys.exit(1)

text = text.replace(old, new, 1)
changes.append("P + FL + NH paid-day calculation")


# =========================================================
# 3. ADD NEW DATA TO REPORT EMPLOYEE
# =========================================================

old = '''            "days":
                days,

            "present_days":
                present_days,

            "ot_hours":
'''

new = '''            "days":
                days,

            "present_days":
                present_days,

            "daily_attendance":
                daily_attendance,

            "actual_present_days":
                actual_present_days,

            "fl_days":
                fl_days,

            "nh_days":
                nh_days,

            "ot_hours":
'''

if old not in text:
    print("[ERROR] Step 3 pattern nahi mila.")
    sys.exit(1)

text = text.replace(old, new, 1)
changes.append("Daily attendance + P/FL/NH counts")


# =========================================================
# 4. SALARY SLIP - PRESENT DAYS = ONLY ACTUAL P
# =========================================================

old = '''                Paragraph(
                    "<b>Payable Days</b>",
                    normal
                ),

                Paragraph(
                    f": {employee['present_days']}",
                    normal
                )
'''

new = '''                Paragraph(
                    "<b>Present Days</b>",
                    normal
                ),

                Paragraph(
                    f": {employee['actual_present_days']}",
                    normal
                )
'''

if old not in text:
    print("[ERROR] Step 4 pattern nahi mila.")
    sys.exit(1)

text = text.replace(old, new, 1)
changes.append("Salary slip Present Days = P only")


# =========================================================
# 5. ADD FL + NH TO SALARY SLIP
# =========================================================

old = '''            [
                Paragraph(
                    "<b>Employee Branch</b>",
                    normal
                ),

                Paragraph(
                    f": {site}",
                    normal
                ),

                Paragraph(
                    "<b>OT Hours</b>",
                    normal
                ),

                Paragraph(
                    f": {employee['ot_hours']}",
                    normal
                )
            ]
'''

new = '''            [
                Paragraph(
                    "<b>FL</b>",
                    normal
                ),

                Paragraph(
                    f": {employee['fl_days']}",
                    normal
                ),

                Paragraph(
                    "<b>NH</b>",
                    normal
                ),

                Paragraph(
                    f": {employee['nh_days']}",
                    normal
                )
            ],

            [
                Paragraph(
                    "<b>Employee Branch</b>",
                    normal
                ),

                Paragraph(
                    f": {site}",
                    normal
                ),

                Paragraph(
                    "<b>OT Hours</b>",
                    normal
                ),

                Paragraph(
                    f": {employee['ot_hours']}",
                    normal
                )
            ]
'''

if old not in text:
    print("[ERROR] Step 5 pattern nahi mila.")
    sys.exit(1)

text = text.replace(old, new, 1)
changes.append("Salary slip FL + NH quantities")


# =========================================================
# SAFETY CHECK
# =========================================================

if text == original:
    print("[ERROR] Koi change nahi hua.")
    sys.exit(1)


# =========================================================
# WRITE UPDATED APP.PY
# =========================================================

APP.write_text(text, encoding="utf-8")

print()
print("-" * 60)
print("UPDATE SUCCESSFUL")
print("-" * 60)

for item in changes:
    print("[OK]", item)

print()
print("Updated file:")
print(APP)

print()
print("Safety backup:")
print(BACKUP)

print()
print("IMPORTANT:")
print("Show Attendance HTML abhi change nahi hua hai.")
print("Next step me HTML update karenge.")

print("=" * 60)