#!/usr/bin/python

"""
Python script to convert Microsoft outlook calendar file (.hol) to .ics
"""

import argparse
import datetime
import re
import uuid
from pathlib import Path

ICS_DATETIME_FORMAT = "%Y%m%dT%H%M%S"


def line_to_event_tuple(line):
    """Convert a given line of a hol file to an event tuple (title, date)

    :param line: A single line from a hol file
    :return: Event tuple (title, date)
    """
    day_title, date_str = line.split(",")
    # hol does not specify timezone, I think, so we default to the local timezone of the user
    begin_datetime = datetime.datetime.strptime(
        date_str.strip(), "%Y/%m/%d"
    )  # .astimezone()
    # print(day_title, date_str, begin_datetime, begin_datetime.tzinfo)
    return day_title, begin_datetime


def write_ics_file(events, destination_filename, title):
    """Write an ics file from the given event map

    :param events: Iterable of event tuples
    :param destination_filename: Pathlib path to destination
    :param title: String calendar title
    """
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        f"PRODID:-//{title}//EN",
    ]

    for name, start_date in events:
        creation_time = (
            datetime.datetime.now().astimezone().strftime(ICS_DATETIME_FORMAT)
        )
        start_date = start_date.strftime("%Y%m%d")
        event_lines = [
            "BEGIN:VEVENT",
            f"UID:{uuid.uuid4()}",
            f"DTSTAMP:{creation_time}",
            f"DTSTART;VALUE=DATE:{start_date}",
            f"SUMMARY:{name}",
            "END:VEVENT",
        ]
        lines += event_lines

    lines.append("END:VCALENDAR")
    # the ics format must have a CRLF at the end of each line
    # see https://icalendar.org/iCalendar-RFC-5545/3-1-content-lines.html
    ics_str = "\r\n".join(lines)

    # finally, write to the new file
    with destination_filename.open("w") as out_file:
        out_file.writelines(ics_str)


def read_hol_file(source_filename):
    """Read and unpack a hol file to an event list

    :param source_filename: Pathlib path to source file
    :return: Tuple of calendar title and associated events
    """
    # TODO: there can be several section in a HOL file - the header and the count
    #  will tell us what to do - but do this later!
    header_pattern = r"\[(?P<title>.+)\]\s*(?P<count>[0-9]+)"

    with source_filename.open("r") as handler:
        header = handler.readline()
        # print(header)
        matches = re.match(header_pattern, header)

        if matches:
            title = matches.group("title")
            count = matches.group("count")
            # print(f"title={title}, count={count}")
        else:
            raise ValueError("Unable to find header in given hol file")

        # how do add the title to the new calendar?

        print(
            "You can try to validate the resultant file using this webform:",
            "https://icalendar.org/validator.html",
        )
        return title, map(line_to_event_tuple, handler.readlines())


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="convert an outlook calendar file (.hol) to an icalendar file (.ics)"
    )

    parser.add_argument(
        "source_path", help="outlook calendar source file (.hol)", type=Path
    )

    parser.add_argument(
        "--dest",
        dest="destination_path",
        help="optional destination filename (.ics)",
        type=Path,
    )

    args = parser.parse_args()

    if args.source_path.suffix != ".hol":
        raise ValueError("Source file has to end with the hol extension")

    if args.destination_path is not None:
        if args.source_path.suffix != ".ics":
            raise ValueError("Destination file has to end with the ics extension")
        destination_path = args.destination_path
    else:
        destination_path = args.source_path.with_suffix(".ics")

    print(f"Attempting to convert {args.source_path} into {destination_path}")

    title, events = read_hol_file(args.source_path)
    write_ics_file(events, destination_path, title)
