from pathlib import Path
import sys

HTML = Path("templates/show_attendance_report.html")
BACKUP = Path(
    r"C:\Users\DELL\SDC_ATTENDANCE_BACKUP\show_attendance_report_before_v1_1_0.html"
)

print("=" * 60)
print("SDC ATTENDANCE - SHOW ATTENDANCE V1.1.0 UPDATE")
print("=" * 60)

if not HTML.exists():
    print("[ERROR] show_attendance_report.html nahi mila.")
    sys.exit(1)

if not BACKUP.exists():
    print("[ERROR] HTML backup nahi mila.")
    sys.exit(1)

text = HTML.read_text(encoding="utf-8-sig")
original = text


# =========================================================
# 1. DATE HEADER
# =========================================================

old = """{% for day in calendar_days %}

<th>
{{ day }}
</th>

{% endfor %}"""

new = """{% for day in calendar_days %}

<th colspan="2" class="date-header">
{{ day }}
</th>

{% endfor %}"""

if old not in text:
    print("[ERROR] Date header pattern nahi mila.")
    print("File modify nahi ki gayi.")
    sys.exit(1)

text = text.replace(old, new, 1)


# =========================================================
# 2. ADD PRESENT + OT SUB-HEADER
# =========================================================

old = """</tr>

</thead>


<tbody>"""

new = """</tr>

<tr>

<th colspan="4" class="sub-header-space"></th>

{% for day in calendar_days %}

<th class="sub-header">
Present
</th>

<th class="sub-header">
OT
</th>

{% endfor %}

<th colspan="18" class="sub-header-space"></th>

</tr>

</thead>


<tbody>"""

if old not in text:
    print("[ERROR] Header ending pattern nahi mila.")
    print("File modify nahi ki gayi.")
    sys.exit(1)

text = text.replace(old, new, 1)


# =========================================================
# 3. EMPLOYEE DATE-WISE ATTENDANCE
# =========================================================

old = """{% for day in employee["days"] %}

<td class="attendance

{% if day == 'P' %}
present
{% elif day == 'A' %}
absent
{% elif day == 'W/O' %}
weekoff
{% elif day == 'NH' %}
holiday
{% elif day == 'FL' %}
leave
{% endif %}"

>

{{ day }}

</td>

{% endfor %}"""

new = """{% for item in employee["daily_attendance"] %}

<td class="attendance

{% if item['status'] == 'P' %}
present
{% elif item['status'] == 'A' %}
absent
{% elif item['status'] == 'W/O' %}
weekoff
{% elif item['status'] == 'NH' %}
holiday
{% elif item['status'] == 'FL' %}
leave
{% endif %}"

>

{{ item['status'] }}

</td>

<td class="attendance ot-cell">

{{ "%.1f"|format(item['ot']) }}

</td>

{% endfor %}"""

if old not in text:
    print("[ERROR] Employee attendance pattern nahi mila.")
    print("File modify nahi ki gayi.")
    sys.exit(1)

text = text.replace(old, new, 1)


# =========================================================
# 4. CSS FOR NEW PRESENT + OT COLUMNS
# =========================================================

css_marker = """</style>"""

css_add = """
/* =====================================================
   V1.1.0 - DATE WISE PRESENT + OT
   ===================================================== */

.date-header {
    text-align: center;
    min-width: 90px;
}

.sub-header {
    font-size: 10px;
    white-space: nowrap;
    text-align: center;
    padding: 6px 8px;
}

.sub-header-space {
    padding: 0;
}

.ot-cell {
    font-size: 11px;
    text-align: center;
    white-space: nowrap;
}
"""

if css_marker not in text:
    print("[ERROR] CSS ending nahi mila.")
    print("File modify nahi ki gayi.")
    sys.exit(1)

text = text.replace(
    css_marker,
    css_add + "\n" + css_marker,
    1
)


# =========================================================
# SAFETY CHECK
# =========================================================

if text == original:
    print("[ERROR] Koi change nahi hua.")
    sys.exit(1)


# =========================================================
# SAVE
# =========================================================

HTML.write_text(text, encoding="utf-8")

print()
print("-" * 60)
print("SHOW ATTENDANCE HTML UPDATE SUCCESSFUL")
print("-" * 60)

print("[OK] Date header updated")
print("[OK] Present + OT sub-columns added")
print("[OK] Date-wise OT display added")
print("[OK] Existing salary/payment sections untouched")
print("[OK] Existing HTML backup preserved")

print()
print("Updated:")
print(HTML)

print()
print("Backup:")
print(BACKUP)

print("=" * 60)