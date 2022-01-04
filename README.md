# IMH ES Roster to iCal Converter

This utility converts ES rosters in `.xlsx` to `.ical` files which can be imported into Google Calendar or any other client supporting the [iCalendar][icalendar] format.

*From this...*

![](raw.jpg)

*...to this*

![](output.jpg)

## Usage

- In the base directory, create a `.env` file with variable `MY_NAME` set to your name as it appears in the roster
- First, extract the schedule from the `.xlsx` file to JSON with `schedule_to_json.py`
- Then, export to iCal with `json_to_ical.py`

## Todo

- Commandline interface with Typer
- Usage instructions

[icalendar]: https://icalendar.org/