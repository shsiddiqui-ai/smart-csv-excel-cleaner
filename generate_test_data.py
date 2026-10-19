
"""Generate 1,000 fictional contacts for testing."""

from pathlib import Path
import random
import pandas as pd

random.seed(42)

FIRST_NAMES = [
    "Ali", "Ahmed", "Sara", "Fatima", "Hassan",
    "Ayesha", "Bilal", "Zara", "Usman", "Hina"
]

LAST_NAMES = [
    "Khan", "Ahmed", "Siddiqui", "Malik",
    "Raza", "Shah", "Iqbal", "Hussain"
]

CITIES = [
    "Karachi", "Lahore", "Islamabad",
    "Multan", "Peshawar"
]

records = []

for number in range(1, 1001):

    first = random.choice(FIRST_NAMES)
    last = random.choice(LAST_NAMES)

    name = f"{first} {last}"
    email = f"contact{number}@example.com"
    phone = f"0300{number:07d}"
    city = random.choice(CITIES)

    # Deliberately create formatting problems.
    if number % 10 == 0:
        email = ""

    elif number % 15 == 0:
        email = "invalid-email"

    if number % 20 == 0:
        phone = "123"

    elif number % 7 == 0:
        phone = f"0300 {number:07d}"

    if number % 8 == 0:
        name = f"  {name}  "

    record = {
        "Name": name,
        "Email": email,
        "Phone": phone,
        "City": city
    }

    records.append(record)

    # Add exact duplicate records for testing.
    if number % 50 == 0:
        records.append(record.copy())

output = Path("input/test_1000_contacts.csv")
output.parent.mkdir(parents=True, exist_ok=True)

pd.DataFrame(records).to_csv(
    output,
    index=False,
    encoding="utf-8-sig"
)

print("Test dataset generated successfully!")
print(f"Total input rows: {len(records)}")
print(f"Saved to: {output}")
