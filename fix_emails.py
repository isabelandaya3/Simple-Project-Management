"""Script to add Contractor row to all email templates in app.py"""
import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Count existing Contractor rows
existing = content.count("get_contractor_name(item.get('bucket'))")
print(f'Already have {existing} Contractor rows')

# Pattern 1: Expanded format with blank line between rows
pattern1 = r'''        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Type</td>
            <td style="border:1px solid #ddd;">\{item\['type'\]\}</td>
        </tr>

        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Identifier</td>'''

replacement1 = '''        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Type</td>
            <td style="border:1px solid #ddd;">{item['type']}</td>
        </tr>

        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Contractor</td>
            <td style="border:1px solid #ddd; font-weight:bold; color:#2c3e50;">{get_contractor_name(item.get('bucket'))}</td>
        </tr>

        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Identifier</td>'''

count1 = len(re.findall(pattern1, content))
print(f'Found {count1} matches for pattern 1 (expanded with blank line)')
content = re.sub(pattern1, replacement1, content)

# Pattern 2: Expanded format without blank line between rows  
pattern2 = r'''        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Type</td>
            <td style="border:1px solid #ddd;">\{item\['type'\]\}</td>
        </tr>
        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Identifier</td>'''

replacement2 = '''        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Type</td>
            <td style="border:1px solid #ddd;">{item['type']}</td>
        </tr>
        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Contractor</td>
            <td style="border:1px solid #ddd; font-weight:bold; color:#2c3e50;">{get_contractor_name(item.get('bucket'))}</td>
        </tr>
        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Identifier</td>'''

count2 = len(re.findall(pattern2, content))
print(f'Found {count2} matches for pattern 2 (expanded no blank line)')
content = re.sub(pattern2, replacement2, content)

# Pattern 2b: Expanded format with width:160px on Identifier  
pattern2b = r'''        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Type</td>
            <td style="border:1px solid #ddd;">\{item\['type'\]\}</td>
        </tr>
        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Identifier</td>'''

replacement2b = '''        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Type</td>
            <td style="border:1px solid #ddd;">{item['type']}</td>
        </tr>
        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Contractor</td>
            <td style="border:1px solid #ddd; font-weight:bold; color:#2c3e50;">{get_contractor_name(item.get('bucket'))}</td>
        </tr>
        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Identifier</td>'''

count2b = len(re.findall(pattern2b, content))
print(f'Found {count2b} matches for pattern 2b (expanded with width on Identifier)')
content = re.sub(pattern2b, replacement2b, content)

# Pattern 3: Missing </td> variant
pattern3 = r'''        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Type</td>
            <td style="border:1px solid #ddd;">\{item\['type'\]\}
        </tr>

        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Identifier</td>'''

replacement3 = '''        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Type</td>
            <td style="border:1px solid #ddd;">{item['type']}
        </tr>

        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Contractor</td>
            <td style="border:1px solid #ddd; font-weight:bold; color:#2c3e50;">{get_contractor_name(item.get('bucket'))}</td>
        </tr>

        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Identifier</td>'''

count3 = len(re.findall(pattern3, content))
print(f'Found {count3} matches for pattern 3 (missing </td>)')
content = re.sub(pattern3, replacement3, content)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)

# Final count
final = content.count("get_contractor_name(item.get('bucket'))")
print(f'Now have {final} Contractor rows (added {final - existing})')

