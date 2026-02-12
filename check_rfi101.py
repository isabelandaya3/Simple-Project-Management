import sqlite3
import re

conn = sqlite3.connect('tracker.db')
c = conn.cursor()

# Check for RFI 101
c.execute("SELECT id, identifier, type, title, status, due_date, bucket FROM item WHERE identifier LIKE '%101%'")
results = c.fetchall()
print("Items with '101' in identifier:")
for row in results:
    print(f"  ID: {row[0]}")
    print(f"  Identifier: {row[1]}")
    print(f"  Type: {row[2]}")
    print(f"  Title: {row[3]}")
    print(f"  Status: {row[4]}")
    print(f"  Due Date: {row[5]}")
    print(f"  Bucket: {row[6]}")
    print()

conn.close()

# Test the subject parsing
test_subject = "Action Required: LEB - Mortenson (NB.TypeF2.0) - RFI #101: was assigned to you for co-review"
print(f"\n=== Testing subject parsing ===")
print(f"Subject: {test_subject}")

# Test the skip regex
skip_pattern = (
    r'review response was (edited|added|created|updated)|'
    r'was forwarded|was assigned(?! to you)|workflow step|status changed|'
    r'comment was added|comment was edited|'
    r'you were mentioned|attachment was added|'
    r'A new review response'
)
match = re.search(skip_pattern, test_subject, re.IGNORECASE)
print(f"Skip regex match: {match}")
if match:
    print(f"  Matched: '{match.group()}'")

# Test item type parsing
def parse_item_type(subject):
    if re.search(r'submittal', subject, re.IGNORECASE):
        return 'Submittal'
    elif re.search(r'RFI', subject, re.IGNORECASE):
        return 'RFI'
    return None

item_type = parse_item_type(test_subject)
print(f"Item type: {item_type}")

# Test identifier parsing
def parse_identifier(subject, item_type):
    if item_type == 'RFI':
        match = re.search(r'RFI\s*[#\-]?\s*(\d+)', subject, re.IGNORECASE)
        if match:
            return f"RFI #{match.group(1).strip()}"
    return None

identifier = parse_identifier(test_subject, item_type)
print(f"Identifier: {identifier}")

# Test co-reviewer parsing with simulated email body
test_body = """Your action is required
Tim Vendel (Mortenson) assigned the following RFI to Crystal Starr (HDR Inc):
RFI #101 LEB50 - 101 - Pedestal Requirements at BF-07, BF-08, BF-09

What's changed
Status: DRAFT -> OPEN IN REVIEW
Ball in court: Tim Vendel (Mortenson) -> Crystal Starr (HDR Inc)

Details
Question: As shown on Sheets S110.1-A.A1 and S110.1-A.A2, the foundation plan and typical details show concrete pedestals under most columns, but not at BF-07, BF-08, and BF-09. Columns BF-08 and BF-09 appear to be set on a thickened slab around the shelter area. Per Emmanuel from HDR Structural Team on 1-10-2026, the original structural model had no pedestals modeled. In an effort to include the pedestals in the model, some pedestals were added to the model later and were mistakenly added at some braced frame columns. Please see attached document, 04 Structural-DOMINO-LEB50-IFC - Foundation Plan Markups 1-18-26, to view correct pedestal requirements. Please verify the final intent of columns BF-07, BF-08, and BF-09 is shown in the attached document, Foundation Plan Markups 1-18-26, and update Sheets S110.A-A1 and S110.1-A.A2 accordingly.
Status: OPEN
Ball in court: Crystal Starr (HDR Inc)
Due Date: February 19, 2025

Participants
Creator: Tim Vendel (Mortenson)
Manager: Tim Vendel (Mortenson)
Reviewers: Crystal Starr (HDR Inc)
Co-reviewers: Isabel Andaya (HDR Inc), Emmanuel Agno (HDR Inc)
Watchers: Angela Lawrence (Mortenson), Matthew Wright (Mortenson), Richard Bagdon (Mortenson), Josh Nicholls (Mortenson), Julie Lorenzo (Mortenson), Chad Owen (HDR Inc)
"""

print(f"\n=== Testing Co-reviewer parsing ===")

# From app.py - parse_rfi_reviewers
def parse_rfi_reviewers(body):
    if not body:
        return []
    reviewers = []
    patterns = [
        r'(?:^|\n)\s*Reviewers[:\s\t]+(.+?)(?=\n\s*(?:Co-reviewers|Watchers|$|\n\n))',
        r'(?:^|\n)\s*Reviewers[:\s\t]+([^\n]+)',
    ]
    for pattern in patterns:
        match = re.search(pattern, body, re.IGNORECASE | re.DOTALL)
        if match:
            reviewer_text = match.group(1).strip()
            parts = reviewer_text.split(',')
            for part in parts:
                part = part.strip()
                name_match = re.match(r'([^(]+)', part)
                if name_match:
                    name = name_match.group(1).strip()
                    if name:
                        reviewers.append(name)
            break
    return reviewers

# From app.py - parse_rfi_coReviewers (NEW)
def parse_rfi_coReviewers(body):
    if not body:
        return []
    coReviewers = []
    patterns = [
        r'(?:^|\n)\s*Co-reviewers[:\s\t]+(.+?)(?=\n\s*(?:Watchers|$|\n\n))',
        r'(?:^|\n)\s*Co-reviewers[:\s\t]+([^\n]+)',
    ]
    for pattern in patterns:
        match = re.search(pattern, body, re.IGNORECASE | re.DOTALL)
        if match:
            coReviewer_text = match.group(1).strip()
            parts = coReviewer_text.split(',')
            for part in parts:
                part = part.strip()
                name_match = re.match(r'([^(]+)', part)
                if name_match:
                    name = name_match.group(1).strip()
                    if name:
                        coReviewers.append(name)
            break
    return coReviewers

def is_user_in_rfi_reviewers(body, user_names):
    if not body or not user_names:
        return False
    reviewers = parse_rfi_reviewers(body)
    coReviewers = parse_rfi_coReviewers(body)
    all_reviewers = reviewers + coReviewers
    for user_name in user_names:
        user_name_lower = user_name.lower()
        for reviewer in all_reviewers:
            if user_name_lower in reviewer.lower():
                return True
    return False

reviewers = parse_rfi_reviewers(test_body)
print(f"Reviewers found: {reviewers}")

coReviewers = parse_rfi_coReviewers(test_body)
print(f"Co-reviewers found: {coReviewers}")

user_names = ["Isabel Andaya", "Andaya", "EoR - Structural"]
result = is_user_in_rfi_reviewers(test_body, user_names)
print(f"Is user in reviewers (with co-reviewers): {result}")

conn.close()
