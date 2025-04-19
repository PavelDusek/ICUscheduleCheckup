#!/usr/bin/env python
# coding: utf-8

"""checks for errors in the ICU schedule"""

# TODO toml with both precise days missing and days of week in one person
# TODO -o notation
# dicts to dataclass {'day_of_week': ..., 'date': ...}

from collections import defaultdict
from pathlib import Path
import datetime
import argparse
import logging

import pandas as pd
import toml
import rich

HEADER = "BEGIN:VCALENDAR\nVERSION:2.0\n"
FOOTER = "END:VCALENDAR"

def parse_args() -> dict:
    """Parses the commandline arguments."""
    now = datetime.datetime.now()
    next_month = now + datetime.timedelta(days=30)

    parser = argparse.ArgumentParser(
        prog="Kontrola rozpisu",
        description="kontroluje rozpis v xls souboru",
        epilog="xls soubor musi mit spravnou strukturu",
    )
    parser.add_argument("filename")
    parser.add_argument(
        "-s", "--sluzby", action="store_true", help="vytvori ics se rozpisem sluzeb"
    )
    parser.add_argument(
        "-k",
        "--kalendar",
        action="store_true",
        help="vytvori ics, kde je vypsany Dusek",
    )
    parser.add_argument(
        "-i", "--skiprows", type=int, default=0, help="vytvori ics se rozpisem sluzeb"
    )
    parser.add_argument(
        "-l", "--log", type=str, default="NOTSET", help="uroven informaci z logging"
    )
    parser.add_argument("-y", "--year", type=int, default=next_month.year, help="rok")
    parser.add_argument(
        "-m", "--month", type=int, default=next_month.month, help="mesic"
    )
    parser.add_argument(
        "-p",
        "--posluzbe",
        help="Kdo je prvni den v mesici po sluzbe, napr. Fik, Du, Ke, atd.",
    )
    parser.add_argument("-t", "--toml", help="lidi.toml file", default="lidi.toml")
    args = parser.parse_args()

    loglevels = {
        "NOTSET": logging.NOTSET,
        "DEBUG": logging.DEBUG,
        "INFO": logging.INFO,
        "WARNING": logging.WARNING,
        "ERROR": logging.ERROR,
        "CRITICAL": logging.CRITICAL,
    }
    loglevel = loglevels[args.log]
    logging.basicConfig(encoding="utf-8", level=loglevel)

    return args


def get_schedule_patterns(path: Path) -> dict:
    """Parse toml file and get the personal schedule patterns."""
    lidi_toml = toml.load(path)
    prezence_den_v_tydnu, prezence_datum = {}, {}
    for clovek, rozvrh in lidi_toml.items():
        if "list" in rozvrh.keys():
            # vycet dany datumy
            prezence_datum[clovek] = rozvrh["list"]
        else:
            # vycet dany dny v tydnu
            prezence_den_v_tydnu[clovek] = rozvrh
    return {"day_of_week": prezence_den_v_tydnu, "date": prezence_datum}


def is_absent(
    name: str, date: datetime.date, part_of_day: str, schedule_patterns: dict
) -> bool:
    """
    Checks the personal presence patterns and
    finds out if the person should be present or absent.
    """
    if days_list := schedule_patterns["date"].get(name):
        # pro cloveka je vycet prezence dany datumy
        return date.day not in days_list

    dopo_absent = defaultdict(list)
    odpo_absent = defaultdict(list)
    dny = ["po", "ut", "st", "ct", "pa", "so", "ne"]
    for clovek, rozvrh in schedule_patterns["day_of_week"].items():
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

    if name in absent.keys() and date.weekday() in absent[name]:
        return True
    return False


def get_ics_string(
    name: str, start: datetime.datetime, end: datetime.datetime, all_day: bool = False
) -> str:
    """Creates ics event string."""
    if all_day:
        time_format = "%Y%m%d"
    else:
        time_format = "%Y%m%dT%H%M%SZ"
    vevent = "BEGIN:VEVENT\n"
    vevent += "DTSTAMP:" + start.strftime(time_format) + "\n"
    vevent += "SUMMARY:" + name + "\n"
    vevent += "DTSTART:" + start.strftime(time_format) + "\n"
    vevent += "DTEND:" + end.strftime(time_format) + "\n"
    vevent += "END:VEVENT\n"
    return vevent


def get_event(year: int, month: int, day: int, text: str, event_type: str) -> str:
    """
    Runs get_ics_sring to get ics event string with proper arguments.
    event_type can be either "dopo", "odpo", or "sluzba".
    """
    if event_type == "sluzba":
        date = datetime.datetime(year, month, day, 0, 0)
        return get_ics_string(name=text, start=date, end=date, all_day=True)
    if event_type == "dopo":
        return get_ics_string(
            name = text,
            start = datetime.datetime(year, month, day, 7, 0),
            end = datetime.datetime(year, month, day, 11, 0),
            all_day = False
        )
    if event_type == "odpo":
        return get_ics_string(
            name = text,
            start = datetime.datetime(year, month, day, 11, 0),
            end = datetime.datetime(year, month, day, 15, 0),
            all_day = False
        )
    return ""


def parse_missing(text: str, part_of_day: str) -> str:
    """
    Parses missing cell, and decides if the absence includes the part_of_day.
    part_of_day can be either "dopo" or "odpo".
    """

    if pd.isna(text):
        return ""
    dopo_pattern, odpo_pattern = ["dop", "d", "dopo"], ["o", "od", "odp", "odpo"]
    entries = text.split(",")
    missing = []
    for entry in entries:
        clovek = entry.split("-")
        if len(clovek) > 1:
            clovek, cas = clovek
            if part_of_day == "dopo" and cas.strip() in dopo_pattern:
                missing.append(clovek)
            if part_of_day == "odpo" and cas.strip() in odpo_pattern:
                missing.append(clovek)
        else:
            missing.append(clovek[0])
    logging.debug("Missing part of day: %s", part_of_day)
    logging.debug("Missing: %s", missing)
    return ",".join(missing)


def get_dataframe(path: Path, args: dict) -> pd.DataFrame:
    """Loads the excel dataframe."""

    df = pd.read_excel(path)

    # Skip n rows while keeping the original excel header
    # (thats the reason for not using pd.read_excel(path, skiprows = args.skiprows)
    lines_to_skip = df.index[0 : args.skiprows]
    df.drop(lines_to_skip, inplace=True)

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
            "Unnamed: 15": "sluzba",
        },
        inplace=True,
    )
    logging.debug(df)
    df.drop(columns=["Unnamed: 0"], inplace=True)
    df.dropna(subset=["datum"], inplace=True)
    logging.debug("***** missing dopo *****")
    df["ne_dopo"] = df["ne"].apply(parse_missing, part_of_day="dopo")
    logging.debug("***** missing odpo *****")
    df["ne_odpo"] = df["ne"].apply(parse_missing, part_of_day="odpo")
    logging.debug(df)
    return df


def create_event_calendar(calendar_dict: dict, path: Path) -> None:
    """Creates ics file according to calendar_dict."""
    logging.info("Creating calendar %s", path)
    for date, event in calendar_dict.items():
        logging.debug("Calendar date: %s", date)
        logging.debug("Calendar event: %s", event)

    with open(path, "w", encoding="utf-8") as f:
        f.write(HEADER)
        logging.info(HEADER)
        for date, events in calendar_dict.items():
            for event_type, value in events.items():
                vevent = get_event(
                    year = date.year,
                    month = date.month,
                    day = date.day,
                    text = value,
                    event_type = event_type
                )
                f.write(vevent)
                logging.info(vevent)
        f.write(FOOTER)
        logging.info(FOOTER)


def calculate_allocations(row: dict, part_of_day: str) -> dict:
    """ Parses row for schedule and calculates how many allocations each person has."""
    allocations = defaultdict(int)
    for key, value in row.items():
        if key.endswith(f"_{part_of_day}") and not pd.isnull(value):
            for person in value.split(","):
                allocations[person.strip()] += 1
    return allocations

def parse_global_events(row: dict ) -> dict:
    """ Parses row and creates an entry into the icu calendar."""
    events = defaultdict(list)
    for pozice, value in row.items():
        if pozice.endswith("_dopo"):
            events["dopo"].append(f"{pozice}: {value}")
        if pozice.endswith("_odpo"):
            events["odpo"].append(f"{pozice}: {value}")
    return {"dopo": "; ".join(events["dopo"]), "odpo": "; ".join(events["odpo"])}


def parse_personal_events(row: dict, name: str) -> dict:
    """ Parses row and creates an entry into a personal calendar."""
    events = defaultdict(list)
    for pozice, value in row.items():
        if name in [person.strip() for person in str(value).split(",")]:
            if pozice.endswith("_dopo"):
                events["dopo"].append(pozice)
            if pozice.endswith("_odpo"):
                events["odpo"].append(pozice)
    return {"dopo": ", ".join(events["dopo"]), "odpo": ", ".join(events["dopo"])}


def check_allocations(
    date: datetime.date, allocations: dict, part_of_day: str, schedule_patterns: dict
) -> None:
    """
    Check if there is more allocations for one person.
    part_of_day could be either "dopo", or "odpo".
    """
    for key, value in allocations.items():
        absent = is_absent(
            name=key,
            date=date,
            part_of_day=part_of_day,
            schedule_patterns=schedule_patterns,
        )
        if (value != 1) and not absent:
            rich.print(f"* [red]{part_of_day}[/red]", key, value)
            logging.info("part_of_day if %s with following values:", part_of_day)
            logging.info(allocations)


def main() -> None:
    """Checks the icu schedule."""

    # Parse command line arguments:
    args = parse_args()
    logging.debug(args)
    posluzbe = args.posluzbe
    rich.print(
        f"Using {args.filename}, year {args.year}, month {args.month}, po sluzbe {args.posluzbe}."
    )
    assert posluzbe

    # Get schedule patterns from toml file:
    schedule_patterns = get_schedule_patterns(path=Path(args.toml))

    # Get allocations from excel file:
    df = get_dataframe(path=Path(args.filename), args=args)

    # Create dicts for calendar event storage:
    # (will be used for ics files generation)
    personal_calendar_dict = defaultdict(dict)
    icu_calendar_dict = defaultdict(dict)
    sluzby_calendar_dict = defaultdict(dict)

    # Parse allocations for each day:
    for _, row in df.iterrows():
        datum = row["datum"]
        logging.debug("Datum: %s", datum)
        date = datetime.date(args.year, args.month, int(datum))

        # Add the person after nightshift to missing
        row["ne_dopo"] = (
            f"{row['ne_dopo']},{posluzbe}" if row["ne_dopo"] else posluzbe
        )
        row["ne_odpo"] = (
            f"{row['ne_odpo']},{posluzbe}" if row["ne_odpo"] else posluzbe
        )

        # Calculate number of allocations for each person
        dopoledne = calculate_allocations(row = row, part_of_day = "dopo")
        odpoledne = calculate_allocations(row = row, part_of_day = "odpo")
        rich.print(f"[green]{datum}[/green]")

        # Fill calendar_dicts
        personal_calendar_dict[date] = parse_personal_events(row = row, name = "Du")
        icu_calendar_dict[date] = parse_global_events(row = row)

        # check allocations for working days only
        if date.weekday() < 5:
            check_allocations(
                date=date,
                allocations=dopoledne,
                part_of_day="dopo",
                schedule_patterns=schedule_patterns,
            )
            check_allocations(
                date=date,
                allocations=odpoledne,
                part_of_day="odpo",
                schedule_patterns=schedule_patterns,
            )
        rich.print()

        # Set posluzbe for the next day and fill sluzby_calendar_dict
        if not pd.isnull(row["sluzba"]):
            # pokud je vyplnena kolonka sluzba, pouzijeme tu
            sluzba = row["sluzba"].strip()
        elif not pd.isnull(row["jip_dopo"]):
            # pokud neni vyplnena kolonka sluzba, pouzijeme hlavniho lekare z dopoledne
            sluzba = row["jip_dopo"].split(",")[0].strip()
        posluzbe = sluzba
        sluzba = "Dušek Pavel" if sluzba == "Du" else sluzba
        sluzby_calendar_dict[date] = {"sluzba": sluzba}

    # Use calendar_event_dicts to get ics files
    if args.kalendar:
        create_event_calendar(
            calendar_dict=personal_calendar_dict,
            path=Path(args.filename.replace(".xls", "_dusek.ics")),
        )
        create_event_calendar(
            calendar_dict=icu_calendar_dict,
            path=Path(args.filename.replace(".xls", "_rozpis.ics")),
        )

    if args.sluzby:
        create_event_calendar(
            calendar_dict=sluzby_calendar_dict,
            path=Path(args.filename.replace(".xls", "_sluzby.ics")),
        )


if __name__ == "__main__":
    main()
