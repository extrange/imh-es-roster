"""
Export events in the generated JSON to iCalendar format
"""

from __future__ import print_function
from rich import print
import json
from collections import namedtuple
from datetime import date, datetime, time, timedelta
from pathlib import Path
from icalendar import Calendar, Event
from zoneinfo import ZoneInfo
from dateutil.relativedelta import relativedelta
import sys
from copy import deepcopy
from rich.console import Console
from rich.table import Table

# Change according to roster month
MONTH = 1

YEAR = 2022

tzinfo = ZoneInfo("Asia/Singapore")
Shift = namedtuple("Shift", ["name", "start", "delta"], defaults=(None, None))

LEGEND = {
    # If a shift is not in the dictionary, it will be added
    # as a whole day event
    "A1": Shift("A1", 8, timedelta(hours=9)),  # 8am-5pm
    "A2": Shift("A2", 9, timedelta(hours=9)),  # 9am-6pm
    "P1": Shift("P1", 12, timedelta(hours=8)),  # 12pm-8pmexi
    "P2": Shift("P2", 14, timedelta(hours=8)),  # 2pm-10pm
    "P3": Shift("P3", 17, timedelta(hours=6)),  # 5pm-11pm
    "N": Shift("Night", 22, timedelta(hours=10)),  # 10pm-8am
    "X": Shift("Post-Night"),
    "SB1": Shift("Standby"),
    "SB2": Shift("Standby"),
    "OFF": Shift("Off"),
    "SSU": Shift("SSU", 10, timedelta(hours=9)),  # 10am-7pm
    "SSUw": Shift("SSU Weekend", 8, timedelta(hours=4)),  # SSU Weekend Round
    "AL": Shift("Annual Leave"),
    "TL": Shift("Training Leave"),
}

console = Console()


def check_representation(legend):
    """
    Check that shift start times and timedeltas are formatted correctly
    """
    for name, shift in legend.items():
        if shift.start:
            start = datetime.combine(date(YEAR, MONTH, 1), time(hour=shift.start))
            print(
                f"{name}: {shift.name} {start.strftime('%-l%P')}-{(start+shift.delta).strftime('%-l%P')}"
            )
        else:
            print(f"{name}: {shift.name}")


def print_schedule_copy(schedule_copy):
    """
    Pretty-print the iCal output for user verification
    """

    table = Table(title="iCal output", header_style="bold magenta")
    table.add_column("Raw")
    table.add_column("Extracted Name")
    table.add_column("Start")
    table.add_column("End")
    for day, shift, vevent in schedule_copy:
        format = "%b %-e, %-l%P"
        name = vevent["SUMMARY"]
        start_date = vevent["DTSTART"].dt
        end_date = vevent["DTEND"].dt
        if (end_date - start_date) == timedelta(days=1):
            start_date = "[green]Whole Day[/green]"
            end_date = "[green]Whole Day[/green]"
        else:
            start_date = start_date.strftime(format)
            end_date = end_date.strftime(format)

        table.add_row(f"{day: <2} {shift: <4}", name, start_date, end_date)

    console.print(table)


def main(schedule_json: Path):

    # Read in schedule
    schedule = json.load(schedule_json.open("r", encoding="utf8"))

    # For verification against original JSON
    schedule_copy = deepcopy(schedule)

    cal = Calendar()

    # Used to determine if a date is in the current or next month
    prev_day = 0

    # Lookup dictionary, convert to valid datetimes for start and end
    for idx, (day, shift_name) in enumerate(schedule):
        shift = LEGEND.get(shift_name)

        # Move to the next month once day number resets
        month_correction = relativedelta(months=1 if day < prev_day else 0)

        event = Event()
        event.add("SUMMARY", shift.name if shift else shift_name)

        if shift and shift.start:
            start_datetime = (
                datetime(YEAR, MONTH, day, shift.start, tzinfo=tzinfo)
                + month_correction
            )
            event.add("DTSTART", start_datetime)
            event.add("DTEND", start_datetime + shift.delta)
        else:
            # If not in dictionary or no start time specified, add as whole day event
            if not shift:
                print(
                    f"[red]Unknown shift: {shift_name} for day {day}. Adding as whole day event.[/red]",
                    file=sys.stderr,
                )
            start_date = date(YEAR, MONTH, day) + month_correction
            event.add("DTSTART", start_date)
            event.add("DTEND", start_date + timedelta(days=1))

        cal.add_component(event)
        schedule_copy[idx].append(event)

        if day > prev_day:
            prev_day = day

    # Print schedule for user verification
    print_schedule_copy(schedule_copy)

    output_path = Path(schedule_json.stem + ".ical")
    with output_path.open("wb") as f:
        f.write(cal.to_ical())


if __name__ == "__main__":
    main(Path("roster.json"))
