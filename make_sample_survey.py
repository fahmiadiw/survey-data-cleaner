"""Generate a fake, messy household health survey export (sample data only).

Mimics a KoboToolbox / ODK style export from field enumerators, with the
problems real survey data usually has.
"""
import csv
import random
from datetime import date, timedelta

random.seed(7)

DISTRICTS = ["Pringsewu", "Gadingrejo", "Ambarawa", "Pardasuka", "Pagelaran",
             "Banyumas", "Adiluwih", "Sukoharjo", "Pagelaran Utara"]
ENUMERATORS = ["ENUM01", "ENUM02", "ENUM03", "ENUM04", "ENUM05", "ENUM06"]
# ENUM04 is careless on purpose, so the per-enumerator chart shows something
SLOPPY = {"ENUM04": 0.35}
SEX = {"F": ["F", "Female", "female", "P", "Perempuan"], "M": ["M", "Male", "male", "L", "Laki-laki"]}
YES = ["Yes", "yes", "Y", "Ya", "1", "TRUE"]
NO = ["No", "no", "N", "Tidak", "0", "FALSE"]
FACILITY = ["Puskesmas", "puskesmas", "Hospital", "RS", "Midwife clinic", "Bidan", "None", ""]

LAT0, LON0 = -5.36, 104.97   # roughly central Pringsewu


def yn(v):
    return random.choice(YES if v else NO)


def fmt_date(d):
    return random.choice([d.strftime("%Y-%m-%d"), d.strftime("%d/%m/%Y"), d.strftime("%m/%d/%Y"),
                          d.strftime("%d-%b-%Y")])


rows = []
start = date(2026, 3, 2)
for i in range(1, 801):
    enum = random.choice(ENUMERATORS)
    err = SLOPPY.get(enum, 0.06)
    sex = random.choices(["F", "M"], [75, 25])[0]
    age = random.randint(17, 49) if sex == "F" else random.randint(18, 65)
    pregnant = sex == "F" and random.random() < 0.3
    anc = random.randint(1, 8) if pregnant else 0
    hh = random.randint(1, 9)
    kids = random.randint(0, min(5, hh - 1))
    lat = LAT0 + random.uniform(-0.12, 0.12)
    lon = LON0 + random.uniform(-0.15, 0.15)
    d = start + timedelta(days=random.randint(0, 45))
    row = {
        "submission_id": f"uuid-{100000 + i}",
        "enumerator_id": enum,
        "interview_date": fmt_date(d),
        "district": random.choice(DISTRICTS),
        "respondent_age": age,
        "respondent_sex": random.choice(SEX[sex]),
        "household_size": hh,
        "children_under5": kids,
        "currently_pregnant": yn(pregnant) if sex == "F" else random.choice(NO + [""]),
        "anc_visits": anc if pregnant else random.choice([0, "", 0]),
        "nearest_facility": random.choice(FACILITY[:-1]),
        "travel_time_min": random.choice([5, 10, 15, 20, 30, 45, 60, 90]),
        "has_bpjs": yn(random.random() < 0.7),
        "phone": random.choice([f"08{random.randint(11, 99)}{random.randint(10000000, 99999999)}",
                                f"+62 8{random.randint(11, 99)} {random.randint(1000, 9999)} {random.randint(1000, 9999)}",
                                f"628{random.randint(11, 99)}{random.randint(10000000, 99999999)}"]),
        "gps_lat": round(lat, 6),
        "gps_lon": round(lon, 6),
    }
    # Inject realistic errors
    if random.random() < err:
        row["respondent_age"] = random.choice([3, 0, 150, 999, ""])
    if random.random() < err:
        row["household_size"] = random.choice([0, 45, -1])
    if random.random() < err:
        row["children_under5"] = row["household_size"] if isinstance(row["household_size"], int) and row["household_size"] > 0 else 3
        row["children_under5"] = int(row["children_under5"]) + 2
    if random.random() < err and not pregnant:
        row["anc_visits"] = random.randint(2, 6)          # skip-logic violation
    if random.random() < err:
        row["gps_lat"], row["gps_lon"] = random.choice([(0, 0), (-6.2, 106.8), ("", "")])
    if random.random() < err / 2:
        row["district"] = random.choice(["pringsewu ", "GADINGREJO", "Ambarawa.", ""])
    if random.random() < err / 2:
        row["interview_date"] = random.choice(["2025-13-40", "", "2027-01-05"])
    rows.append(row)

# Duplicate submissions (form sent twice from a tablet)
for r in random.sample(rows, 25):
    d = dict(r)
    d["submission_id"] = r["submission_id"] + "-dup"
    rows.append(d)
random.shuffle(rows)

with open("sample_survey_export.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
    w.writeheader()
    w.writerows(rows)
print(f"Wrote {len(rows)} rows")
