import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace occurrences with {reviewer_due_date or 'N/A'} -> Initial Review Due Date
content = re.sub(
    r'>Your Due Date</td>\s*\n\s*<td style="border:1px solid #ddd; color:#2980b9; font-weight:bold;">\{reviewer_due_date',
    r'>Initial Review Due Date</td>\n            <td style="border:1px solid #ddd; color:#2980b9; font-weight:bold;">{reviewer_due_date',
    content
)

# Replace occurrences with {qcr_due_date_email or 'N/A'} -> QC Due Date  
content = re.sub(
    r'>Your Due Date</td>\s*\n\s*<td style="border:1px solid #ddd; color:#2980b9; font-weight:bold;">\{qcr_due_date_email',
    r'>QC Due Date</td>\n            <td style="border:1px solid #ddd; color:#2980b9; font-weight:bold;">{qcr_due_date_email',
    content
)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)

print('Replacements done!')
