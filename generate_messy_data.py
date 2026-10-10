#!/usr/bin/env python3
"""
Generate a synthetic, deliberately messy customer list.

Everything here is made up: no real people, emails or phone numbers
(phone numbers use the reserved fictional 555-01xx range). The mess is
realistic on purpose, so the cleaning pipeline has something to fix:

* duplicates (exact, re-typed, and partial)        * mixed name formats
* inconsistent capitalization and stray spaces     * invalid / typo'd emails
* five or more phone formats and invalid numbers   * province/state spelled many ways
* five date formats and invalid dates              * currency symbols and bad amounts
* placeholder text like "N/A"                      * rows with no usable contact info

Usage:
    python generate_messy_data.py                 # 500 rows -> data/messy_customers.csv
    python generate_messy_data.py --rows 1000 --seed 7
"""

import argparse
import csv
import random
from datetime import date, timedelta
from pathlib import Path

FIELDS = ["customer_id", "full_name", "email", "phone", "city",
          "province_state", "signup_date", "total_spend"]

FIRST_NAMES = [
    "James", "Mary", "Michael", "Priya", "David", "Aisha", "Chen", "Sofia",
    "Daniel", "Fatima", "Kwame", "Emily", "Omar", "Hannah", "Luis", "Grace",
    "Noah", "Amara", "Ethan", "Mei", "Lucas", "Zainab", "Ryan", "Isabella",
    "Tyrone", "Chloe", "Samuel", "Nadia", "Marcus", "Olivia", "Andre", "Leah",
    "Jean-Luc", "Maria", "Victor", "Tanya", "Ibrahim", "Sarah", "Kevin", "Anita",
]
LAST_NAMES = [
    "Smith", "Johnson", "Williams", "Patel", "Brown", "Singh", "Lee", "Garcia",
    "Martin", "Ahmed", "Mensah", "Taylor", "Khan", "Wilson", "Lopez", "Campbell",
    "O'Brien", "McDonald", "Nguyen", "Robinson", "Clarke", "Thompson", "Osei",
    "Anderson", "Wright", "Scott", "Green", "Baker", "Adams", "Hill",
    "Smith-Jones", "Mitchell", "Roberts", "Turner", "Phillips", "Evans",
]
# code: (full name, area codes, cities)
REGIONS = {
    "ON": ("Ontario", ["416", "647", "437", "905", "613", "519"],
           ["Toronto", "Scarborough", "Markham", "Mississauga", "Ottawa", "London"]),
    "QC": ("Quebec", ["514", "438"], ["Montreal", "Laval"]),
    "AB": ("Alberta", ["403", "587"], ["Calgary", "Edmonton"]),
    "BC": ("British Columbia", ["604", "778"], ["Vancouver", "Victoria"]),
    "MB": ("Manitoba", ["204"], ["Winnipeg"]),
    "NS": ("Nova Scotia", ["902"], ["Halifax"]),
    "NY": ("New York", ["212", "718"], ["New York", "Buffalo"]),
    "IL": ("Illinois", ["312"], ["Chicago"]),
    "CA": ("California", ["310", "415"], ["Los Angeles", "San Francisco"]),
    "FL": ("Florida", ["305"], ["Miami"]),
    "WA": ("Washington", ["206"], ["Seattle"]),
}
REGION_WEIGHTS = [8, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1]
DOMAINS = ["gmail.com", "yahoo.com", "outlook.com", "hotmail.com", "icloud.com"]
TYPO_DOMAINS = ["gmial.com", "gmail.con", "hotmial.com", "yaho.com", "gamil.com"]
DATE_FORMATS = ["%Y-%m-%d", "%m/%d/%Y", "%b %d, %Y", "%d-%b-%Y", "%B %d, %Y"]


# --------------------------------------------------------------------------
# Clean "truth" records
# --------------------------------------------------------------------------
def make_customer(rng: random.Random, idx: int, used_emails: set) -> dict:
    first, last = rng.choice(FIRST_NAMES), rng.choice(LAST_NAMES)
    code = rng.choices(list(REGIONS), weights=REGION_WEIGHTS)[0]
    _, areas, cities = REGIONS[code]

    local_style = rng.choice([
        f"{first}.{last}", f"{first}{last}{rng.randint(1, 99)}", f"{first[0]}{last}",
    ])
    local = "".join(ch for ch in local_style.lower() if ch.isalnum() or ch in "._")
    domain = rng.choice(DOMAINS)
    while f"{local}@{domain}" in used_emails:  # every customer gets a unique email
        local += str(rng.randint(1, 99))
    used_emails.add(f"{local}@{domain}")

    start = date(2022, 1, 1)
    span = (date(2026, 9, 30) - start).days
    return {
        "idx": idx,
        "first": first,
        "last": last,
        "email": f"{local}@{domain}",
        "area": rng.choice(areas),
        "line": f"01{rng.randint(0, 99):02d}",  # 555-01xx is reserved for fiction
        "city": rng.choice(cities),
        "region": code,
        "signup": start + timedelta(days=rng.randint(0, span)),
        "spend": round(rng.uniform(15, 1800), 2),
    }


# --------------------------------------------------------------------------
# Messifiers: each turns a clean value into a realistic dirty one
# --------------------------------------------------------------------------
def mess_name(rng, c):
    first, last, r = c["first"], c["last"], rng.random()
    if r < 0.05:
        return ""
    if r < 0.12:
        return f"{last}, {first}"
    if r < 0.30:
        return f"{first} {last}".upper()
    if r < 0.45:
        return f"{first} {last}".lower()
    if r < 0.55:
        return f"  {first}   {last} "
    if r < 0.60:
        return f"{first} {last}{rng.randint(1, 9)}"
    return f"{first} {last}"


def mess_email(rng, c):
    email, r = c["email"], rng.random()
    if r < 0.04:
        return ""
    if r < 0.07:
        return "N/A"
    if r < 0.11:
        return email.replace("@", ".")  # no @ sign -> invalid
    if r < 0.17:
        return email.split("@")[0] + "@" + rng.choice(TYPO_DOMAINS)
    if r < 0.27:
        return email.upper()
    if r < 0.35:
        return f"  {email} "
    return email


def mess_phone(rng, c):
    a, e, l, r = c["area"], "555", c["line"], rng.random()
    if r < 0.04:
        return ""
    if r < 0.06:
        return "N/A"
    if r < 0.10:
        return rng.choice([f"{e}-{l}", "123", "000-000-0000", f"{a}-{e}"])
    return rng.choice([
        f"{a}-{e}-{l}", f"({a}) {e}-{l}", f"{a}.{e}.{l}", f"{a}{e}{l}",
        f"+1 {a} {e} {l}", f"1-{a}-{e}-{l}", f"{a} {e} {l} ext {rng.randint(1, 99)}",
    ])


def mess_region(rng, c):
    code = c["region"]
    full = REGIONS[code][0]
    if rng.random() < 0.05:
        return ""
    variants = [code, code.lower(), full, full.lower(), full.upper()]
    if code == "ON":
        variants.append("Ont.")
    return rng.choice(variants)


def mess_city(rng, c):
    city, r = c["city"], rng.random()
    if r < 0.03:
        return ""
    if r < 0.18:
        return city.upper()
    if r < 0.33:
        return city.lower()
    if r < 0.38:
        return f"  {city} "
    return city


def mess_date(rng, c):
    r = rng.random()
    if r < 0.02:
        return ""
    if r < 0.05:
        return rng.choice(["not sure", "TBD", "32/13/2024"])
    return c["signup"].strftime(rng.choice(DATE_FORMATS))


def mess_spend(rng, c):
    x, r = c["spend"], rng.random()
    if r < 0.03:
        return ""
    if r < 0.05:
        return "N/A"
    if r < 0.07:
        return str(-x)
    if r < 0.20:
        return f"${x:,.2f}"
    if r < 0.30:
        return f"CAD {x}"
    return str(x)


def messify(rng, c, customer_id):
    cid = customer_id
    r = rng.random()
    if r < 0.05:
        cid = cid.lower()
    elif r < 0.08:
        cid = f" {cid} "
    return {
        "customer_id": cid,
        "full_name": mess_name(rng, c),
        "email": mess_email(rng, c),
        "phone": mess_phone(rng, c),
        "city": mess_city(rng, c),
        "province_state": mess_region(rng, c),
        "signup_date": mess_date(rng, c),
        "total_spend": mess_spend(rng, c),
    }


# --------------------------------------------------------------------------
# Build the full dataset
# --------------------------------------------------------------------------
def generate(rows: int = 500, seed: int = 42) -> list:
    rng = random.Random(seed)
    n_unique = max(1, int(rows * 0.88))
    used_emails: set = set()
    customers = [make_customer(rng, i + 1, used_emails) for i in range(n_unique)]
    messy = [messify(rng, c, f"C{c['idx']:05d}") for c in customers]

    # ~2% of customers end up with no usable contact details at all
    for row in rng.sample(messy, k=max(1, n_unique // 50)):
        row["email"], row["phone"] = "", ""

    out = list(messy)
    next_id = n_unique + 1
    for _ in range(rows - n_unique):
        i = rng.randrange(n_unique)
        kind = rng.random()
        if kind < 0.20:                       # exact copy
            out.append(dict(messy[i]))
            continue
        dup = messify(rng, customers[i], messy[i]["customer_id"])
        if rng.random() < 0.70:               # re-entered under a new ID
            dup["customer_id"] = f"C{next_id:05d}"
            next_id += 1
        if rng.random() < 0.25:
            dup["email"] = ""                 # same person, no email this time
        elif rng.random() < 0.15:
            dup["phone"] = ""
        out.append(dup)

    rng.shuffle(out)
    return out


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a messy sample customer list")
    parser.add_argument("--rows", type=int, default=500, help="total rows (default 500)")
    parser.add_argument("--seed", type=int, default=42, help="random seed (default 42)")
    parser.add_argument("--output", type=Path, default=Path("data/messy_customers.csv"))
    args = parser.parse_args()

    data = generate(args.rows, args.seed)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(data)
    print(f"Wrote {len(data)} messy rows to {args.output}")


if __name__ == "__main__":
    main()
