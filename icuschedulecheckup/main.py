#!/usr/bin/env python
# coding: utf-8

# Post mortem debuging python3 -m pdb -c continue Rozpis.py -p Bik 2024-06.xls
#TODO sluzba from different column
#TODO toml with both precise days missing and days of week in one person
#TODO -o notation

import pandas as pd
import toml
from collections import defaultdict
import datetime
import pdb
import argparse
from icecream import ic
from rich import print

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

filename, month, year, posluzbe = args.filename, args.month, args.year, args.posluzbe
print(f"Using {filename}, year {year}, month {month}, po sluzbe {posluzbe}.")
assert posluzbe
lidi_toml = toml.load("lidi.toml")
lidi = lidi_toml.keys()

prezence_den_v_tydnu, prezence_datum = {}, {}
for clovek, rozvrh in lidi_toml.items():
    if "list" in rozvrh.keys():
        #vycet dany datumy
        prezence_datum[clovek] = rozvrh["list"]
    else:
        #vycet dany dny v tydnu
        prezence_den_v_tydnu[clovek] = rozvrh

def is_absent(name, date, part_of_day, lidi_toml=lidi_toml):
    if days_list := prezence_datum.get(name):
        #pro cloveka je vycet prezence dany datumy
        return not date.day in days_list
        
    dopo_absent = defaultdict(list)
    odpo_absent = defaultdict(list)
    dny = ["po", "ut", "st", "ct", "pa", "so", "ne"]
    for clovek, rozvrh in prezence_den_v_tydnu.items():
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
        raise ValueError("Part of Day specified incorrectly.")

    if name in absent.keys():
        if date.weekday() in absent[name]:
            return True
    return False


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
    if text != "ne":
        name, typ = text.split("_")
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
    dopo_pattern, odpo_pattern = ["dop", "d", "dopo"], ["o", "od", "odp", "odpo"]
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
    date = datetime.date(year, month, int(datum))
    den = date.weekday()
    dopoledne = defaultdict(int)
    odpoledne = defaultdict(int)

    #TODO kolonka po sluzbe z posledniho sloupecku

    if not "ne_dopo" in rows.index:
        rows["ne_dopo"] = posluzbe
    else:
        rows["ne_dopo"] = str(rows["ne_dopo"]) + ", " + posluzbe

    if not "ne_odpo" in rows.index:
        rows["ne_odpo"] = posluzbe
    else:
        rows["ne_odpo"] = str(rows["ne_odpo"]) + ", " + posluzbe

    if den < 5:  # pouze vsedni dny
        print(f"[green]{datum}[/green]")
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
            if (value != 1) and not is_absent(key, date, "dopo"):
                print("* [red]dopo[/red]", key, value)
                if args.verbose:
                    print(dopoledne)
        for key, value in odpoledne.items():
            if (value != 1) and not is_absent(key, date, "odpo"):
                print("* [cyan]odpo[/cyan]", key, value)
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
