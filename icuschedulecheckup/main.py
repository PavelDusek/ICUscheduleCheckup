#!/usr/bin/env python
# coding: utf-8

# Post mortem debuging python3 -m pdb -c continue Rozpis.py -p Bik 2024-06.xls
#TODO sluzba from different column
#TODO toml with precise dates missing
#TODO -o notation

import pandas as pd
import toml
from collections import defaultdict
from icecream import ic
import datetime
import pdb
import argparse

now = datetime.datetime.now()
next_month = now + datetime.timedelta(days=30)

parser = argparse.ArgumentParser(
    prog="Kontrola rozpisu",
    description="kontroluje rozpis v xls souboru",
    epilog="xls soubor musi mit spravnou strukturu",
)
parser.add_argument("filename")
parser.add_argument("-v", "--verbose", action="store_true", help="vypise obsazeni jednotlivych dnu")
parser.add_argument("-s", "--sluzby", action="store_true", help="vytvori ics se rozpisem sluzeb")
parser.add_argument("-k", "--kalendar", action="store_true", help="vytvori ics, kde je vypsany Dusek")
parser.add_argument("-y", "--year", type=int, default=next_month.year, help="rok")
parser.add_argument("-m", "--month", type=int, default=next_month.month, help="mesic")
parser.add_argument("-p", "--posluzbe", help="Kdo je prvni den v mesici po sluzbe, napr. Fik, Du, Ke, atd.")
parser.add_argument("-t", "--toml", help="lidi.toml file", default="lidi.toml")
args = parser.parse_args()


lidi_toml = toml.load("lidi.toml")

def is_absent(name, dow, part_of_day, lidi_toml=lidi_toml):
    # kdy lidi chybi
    dopo_absent = {
        "Rek": [0, 1, 2, 3, 4],
        "Ga": [0, 1, 2, 3, 4],
        "Kre": [3],
        "Růž": [0, 1, 2, 3],
        "For": [3, 4],
        #"Bik": [0, 1, 2, 3],
        "Slo": [4],
        "Ho": [1],
        "Sul": [3],
        "Hry": [2],
    }
    odpo_absent = {
        "Rek": [0, 1, 2, 3, 4],
        "Ga": [0, 1, 2, 3, 4],
        "Kre": [0, 3],
        "Růž": [0, 1, 2, 3],
        "For": [3, 4],
        #"Bik": [0, 1, 2, 3],
        "Slo": [4],
        "Ho": [1],
        "Sul": [0, 1, 2, 3, 4],
        "Hry": [2],
    }
    dopo_absent = defaultdict(list)
    odpo_absent = defaultdict(list)
    dny = ["po", "ut", "st", "ct", "pa", "so", "ne"]
    for clovek, rozvrh in lidi_toml.items():
        for cast_dne, present in rozvrh.items():
            den, cas = cast_dne.split("_")
            if cas == "dopo" and not present:
                dopo_absent[clovek].append(dny.index(den))
            if cas == "odpo" and not present:
                odpo_absent[clovek].append(dny.index(den))

    if part_of_day == "dopo":
        absent = dopo_absent
    elif part_of_day == "odpo":
        absent = odpo_absent
    else:
        raise ValueError("Part of Day not correctly specified.")

    if name in absent.keys():
        if dow in absent[name]:
            return True
    return False


# df = pd.read_excel('červen rozpis.xls')
# df = pd.read_excel('červenec rozpis.xls')
# df, month, year, posluzbe = pd.read_excel('rijen.xls'), 10, 2021, "Ke"
# df, month, year, posluzbe = pd.read_excel('rijen.xls'), 10, 2021, "Ke"
# df, month, year, posluzbe = pd.read_excel('listopad.xls'), 11, 2021, "Fik"
# df, month, year, posluzbe = pd.read_excel('prosinec.xls'), 12, 2021, "Kre"
# df, month, year, posluzbe = pd.read_excel('leden.xls'), 1, 2022, "Slo"
# df, month, year, posluzbe = pd.read_excel('2022-02b.xls'), 2, 2022, "Ke"
# df, month, year, posluzbe = pd.read_excel('2022-03_.xls'), 3, 2022, "Ke"
# df, month, year, posluzbe = pd.read_excel('2022-03.xls'), 3, 2022, "Ke"
# df, month, year, posluzbe = pd.read_excel('2022-04.xls'), 4, 2022, "Ke"
# df, month, year, posluzbe = pd.read_excel('2022-05.xls'), 5, 2022, "Slo"
# df, month, year, posluzbe = pd.read_excel('2022-06.xls'), 6, 2022, "Kre"
# df, month, year, posluzbe = pd.read_excel('2022-07.xls'), 7, 2022, "Fik"
# df, month, year, posluzbe = pd.read_excel('2022-08.xls'), 8, 2022, "Slo"
# df, month, year, posluzbe = pd.read_excel('2022-09.xls'), 9, 2022, "Slo"
# df, month, year, posluzbe = pd.read_excel('2022-10.xls'), 10, 2022, "Slo"
# df, month, year, posluzbe = pd.read_excel('2022-11.xls'), 11, 2022, "Fik"
# df, month, year, posluzbe = pd.read_excel('2023-01.xls'), 1, 2023, "Fik"
# filename, month, year, posluzbe = '2023-01.xls', 1, 2023, "Fik"
# filename, month, year, posluzbe = '2023-02.xls', 2, 2023, "Kre"
# filename, month, year, posluzbe = '2023-04.xls', 4, 2023, "Du"
# filename, month, year, posluzbe = '2023-12.xls', 12, 2023, "Fik"
#filename, month, year, posluzbe = "2024-02.xls", 2, 2024, "Slo"
#filename, month, year, posluzbe = "2024-03.xls", 3, 2024, "Ho"
#filename, month, year, posluzbe = "2024-04.xls", 4, 2024, "Růž"
#filename, month, year, posluzbe = "2024-05.xls", 5, 2024, "For"

filename, month, year, posluzbe = args.filename, args.month, args.year, args.posluzbe

assert posluzbe

print(f"Using {filename}, year {year}, month {month}, po sluzbe {posluzbe}.")
# lidi = ['Rek', 'Du', 'Ga', 'Ke', 'Kre', 'Fik', 'Slo', 'Růž', 'Her']
# lidi = ['Rek', 'Du', 'Ga', 'Ke', 'Kre', 'Fik', 'Slo', 'Růž', 'Bar']
# lidi = ['Rek', 'Du', 'Ga', 'Ke', 'Kre', 'Fik', 'Slo', 'Růž', 'Škr']
# lidi = ['Rek', 'Du', 'Ga', 'Ke', 'Kre', 'Fik', 'Slo', 'Růž', 'Ho', 'Dol' ]
# lidi = ['Ho', 'Du', 'Ga', 'Ke', 'Kre', 'Fik', 'Slo', 'Růž', 'Ho', 'For' ]
# lidi = ['Ho', 'Du', 'Hry', 'Ke', 'Kre', 'Fik', 'Slo', 'Růž', 'Ho', 'For' ]
# lidi = ['Ho', 'Du', 'Kre', 'Fik', 'Slo', 'Růž', 'Ho', 'For', 'Bik' ]
# lidi = ['Ho', 'Du', 'Kre', 'Fik', 'Slo', 'Růž', 'Ho', 'For', 'Bik', 'Jak' ]
# lidi = ['Ho', 'Du', 'Kre', 'Fik', 'Ku', 'Růž', 'Ho', 'For', 'Bik', 'Šar', 'Sul' ]
#lidi = ["Ho", "Du", "Kre", "Fik", "Růž", "Ho", "For", "Bik", "Jak", "Sul"]
#lidi = ["Ho", "Du", "Kre", "Fik", "Růž", "Ho", "For", "Hry", "Sul"]
lidi = ["Ho", "Du", "Kre", "Fik", "Růž", "Ho", "For", "Pil", "Sul"]
lidi = lidi_toml.keys()


def makeVEVENT(name, start, end):
    time_format = "%Y%m%dT%H%M%SZ"
    vevent = "BEGIN:VEVENT\n"
    vevent += "DTSTAMP:" + start.strftime(time_format) + "\n"
    vevent += "SUMMARY:" + name + "\n"
    vevent += "DTSTART:" + start.strftime(time_format) + "\n"
    vevent += "DTEND:" + end.strftime(time_format) + "\n"
    vevent += "END:VEVENT\n"
    return vevent

def makeEvent(year, month, day, text):
    name, typ = text.split("_")
    if name != "ne":
        if typ == "dopo":
            start = datetime.datetime(year, month, day, 7, 0)
            end = datetime.datetime(year, month, day, 11, 0)
        if typ == "odpo":
            start = datetime.datetime(year, month, day, 11, 0)
            end = datetime.datetime(year, month, day, 15, 0)
        return makeVEVENT(name, start, end)

def parse_missing(text, missing_type) -> str:
    if pd.isna(text):
        return ""
    dopo_pattern, odpo_pattern = ["dop", "d", "dopo", "do"], ["o", "od", "odp", "odpo"]
    entries = text.split(",")
    missing = []
    for entry in entries:
        clovek = entry.split("-")
        if len(clovek) > 1:
            clovek, cas = clovek
            if missing_type == 'dopo' and cas.strip() in dopo_pattern:
                missing.append(clovek)
            if missing_type == 'odpo' and cas.strip() in odpo_pattern:
                missing.append(clovek)
        else:
            missing.append(clovek[0])
    ic(missing_type, missing)
    return ",".join(missing)

df = pd.read_excel(filename)

df.rename(
    columns={
        "Unnamed: 1": "datum",
        "Unnamed: 2": "jip_dopo",
        "Unnamed: 3": "jip_odpo",
        "Unnamed: 4": "sono_dopo",
        "Unnamed: 5": "sono_odpo",
        "Unnamed: 6": "sono2_dopo",
        "Unnamed: 7": "sono2_odpo",
        "Unnamed: 8": "amb_dopo",
        "Unnamed: 9": "amb_odpo",
        "Unnamed: 10": "kons_dopo",
        "Unnamed: 11": "kons_odpo",
        "Unnamed: 12": "vyu_dopo",
        "Unnamed: 13": "vyu_odpo",
        "Unnamed: 14": "ne",
        "Unnamed: 15": "slouzi",
    },
    inplace=True,
)
df.drop(columns=["Unnamed: 0"], inplace=True)
df.dropna(subset=["datum"], inplace=True)
ic("***** missing dopo *****")
df['ne_dopo'] = df['ne'].apply(parse_missing, missing_type='dopo')
ic("***** missing odpo *****")
df['ne_odpo'] = df['ne'].apply(parse_missing, missing_type='odpo')
print(df)

dusek = defaultdict(list)

if args.verbose:
    print(df.head())

for i, rows in df.iterrows():
    datum = rows["datum"]
    den = datetime.date(year, month, int(datum)).weekday()
    dopoledne = defaultdict(int)
    odpoledne = defaultdict(int)

    if not "ne_dopo" in rows.index:
        rows["ne_dopo"] = posluzbe
    else:
        rows["ne_dopo"] = str(rows["ne_dopo"]) + ", " + posluzbe

    if not "ne_odpo" in rows.index:
        rows["ne_odpo"] = posluzbe
    else:
        rows["ne_odpo"] = str(rows["ne_odpo"]) + ", " + posluzbe

    if den < 5:  # pouze vsedni dny
        print(datum)
        for clovek in lidi:
            dopoledne[clovek] = 0
            odpoledne[clovek] = 0

        for row, value in rows.items():
            if row != "datum" and not pd.isnull(value):
                if row.endswith("_dopo"):
                    for obsazeni in value.split(","):
                        dopoledne[obsazeni.strip()] += 1
                if row.endswith("_odpo"):
                    for obsazeni in value.split(","):
                        odpoledne[obsazeni.strip()] += 1
                if "Du" in value:
                    dusek[datum].append(row)

        # kontrola
        for key, value in dopoledne.items():
            if (value != 1) and not is_absent(key, den, "dopo"):
                print("* dopo", key, value)
                if args.verbose:
                    print(dopoledne)
        for key, value in odpoledne.items():
            if (value != 1) and not is_absent(key, den, "odpo"):
                print("* odpo", key, value)
                if args.verbose:
                    print(odpoledne)
        print()
    if not pd.isnull(rows["jip_dopo"]):
        posluzbe = rows["jip_dopo"].split(",")[0].strip()

if args.kalendar:
    for i, value in dusek.items():
        print(i, value)

    with open(filename.replace(".xls", "_dusek.ics"), "w") as f:
        header = "BEGIN:VCALENDAR\nVERSION:2.0\n"
        footer = "END:VCALENDAR"

        f.write(header)
        print(header)
        for i, values in dusek.items():
            for value in values:
                try:
                    vevent = makeEvent(year, month, int(i), value)
                except Exception as e:
                    print(e)
                if vevent:
                    f.write(vevent)
                    print(vevent)
        f.write(footer)
        print(footer)

    with open(filename.replace(".xls", "_rozpis.ics"), "w") as f:
        header = "BEGIN:VCALENDAR\nVERSION:2.0\n"
        footer = "END:VCALENDAR"
        print(header)
        f.write(header)

        for i, row in df.iterrows():
            day = int(row["datum"])
            dopo = f"jip: {row['jip_dopo']}"
            odpo = f"jip: {row['jip_odpo']}"

            for pozice in ["sono_dopo", "amb_dopo", "kons_dopo", "vyu_dopo"]:
                if not pd.isnull(row[pozice]):
                    dopo = dopo + "; " + pozice.replace("_dopo", ": ") + row[pozice]
            for pozice in ["sono_odpo", "amb_odpo", "kons_odpo", "vyu_odpo"]:
                if not pd.isnull(row[pozice]):
                    odpo = odpo + "; " + pozice.replace("_odpo", ": ") + row[pozice]

            vevent = makeVEVENT(
                dopo,
                datetime.datetime(year, month, day, 7, 0),
                datetime.datetime(year, month, day, 11, 0),
            )
            print(vevent)
            f.write(vevent)

            vevent = makeVEVENT(
                odpo,
                datetime.datetime(year, month, day, 11, 0),
                datetime.datetime(year, month, day, 15, 0),
            )
            print(vevent)
            f.write(vevent)
        print(footer)
        f.write(footer)


# == Služby ==
def makeSluzbyVEVENT(name, date):
    time_format = "%Y%m%d"
    vevent = "BEGIN:VEVENT\n"
    vevent += "DTSTAMP:" + date.strftime(time_format) + "\n"
    vevent += "SUMMARY:" + name + "\n"
    vevent += "DTSTART:" + date.strftime(time_format) + "\n"
    vevent += "DTEND:" + date.strftime(time_format) + "\n"
    vevent += "END:VEVENT\n"
    return vevent

if args.sluzby:
    sluzby_filename = filename.replace(".xls", "_sluzby.xlsx")
    sluzby = pd.read_excel(sluzby_filename)
    sluzby.rename(
        columns={
            "Lékařské pohotovostní služby neurologické kliniky": "datum",
            "Unnamed: 4": "JIP",
        },
        inplace=True,
    )
    sluzby.drop(
        columns=["Unnamed: 0", "Unnamed: 2", "Unnamed: 3", "Unnamed: 5"], inplace=True
    )
    print(sluzby.head())
    sluzby.dropna(subset=["datum"], inplace=True)
    sluzby.iloc[1, 0] = sluzby.iloc[1, 0].replace("+RC:RC:R[27]C[4]", "")
    sluzby.drop(index=2, inplace=True)

    if args.verbose:
        print(sluzby.head())

    with open(sluzby_filename.replace("xlsx", "ics"), "w") as f:
        header = "BEGIN:VCALENDAR\nVERSION:2.0\n"
        footer = "END:VCALENDAR"

        f.write(header)
        print(header)
        for i, row in sluzby.iterrows():
            datum = row["datum"]
            name = row["JIP"]
            if name == "Dušek":
                name = "Dušek Pavel"

            date = datetime.datetime(year, month, int(datum.replace(".", "")))
            if not pd.isna(name):
                vevent = makeSluzbyVEVENT(name, date)
                f.write(vevent)
            print(vevent)
        f.write(footer)
        print(footer)
