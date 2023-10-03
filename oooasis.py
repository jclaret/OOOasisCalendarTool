#! /usr/bin/python -W ignore
"""
OOOasis: Google Calendar Management
==================================

This module provides a set of tools for managing Google Calendar events, with a
particular focus on "Out of Office" (OOO) events. It allows users to enable,
disable, and check OOO status through various classes and functions, and it
offers command-line interaction for ease of use.

The module utilizes the Google Calendar API to interact with user calendars and
manage events. It also leverages the `rich` library for console output to enhance
user experience in the command-line interface.

Classes:
    - GoogleCalendarAuth: Handles authentication processes for Google Calendar.
    - GoogleCalendarManager: Manages various operations on Google Calendar, such
      as enabling/disabling OOO, checking OOO status, etc.

Dependencies:
    - google-auth, google-auth-oauthlib, google-auth-httplib2, and google-api-python-client:
      For interacting with the Google Calendar API.
    - rich: For enhanced console output.
    - dateutil: For timezone handling.
    - configparser: For configuration file parsing.

Note:
    Only "default" and "workingLocation" event types can be created using the API
    as of the last update. Extended support for other event types may be available
    in future API releases. For more details, refer to the Google Calendar API documentation
    and issue tracker.
    - API Reference: https://developers.google.com/calendar/api/v3/reference/events/insert
    - Issue Tracker: https://issuetracker.google.com/issues/112063903
"""

import os.path
import datetime as dt
import argparse
import sys
import configparser
from datetime import datetime, timedelta
from dateutil import tz

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from rich.console import Console

SCOPES = ["https://www.googleapis.com/auth/calendar"]
console = Console()

# NOTE: Currently, only "default " and "workingLocation" events can be created using the API.
# Extended support for other event types will be made available in later releases.
# https://developers.google.com/calendar/api/v3/reference/events/insert
# https://issuetracker.google.com/issues/112063903


class GoogleCalendarAuth:
    """
    Handles authentication for Google Calendar.

    This class provides a method to authenticate the user with Google Calendar
    and return a service object that can be used to interact with the Google Calendar API.
    It uses OAuth 2.0 for authentication and saves the credentials in a local file
    to reuse them in future sessions.
    """

    @staticmethod
    def authenticate():
        """
        Authenticate the user and return a Google Calendar API service object.

        This method performs OAuth 2.0 authentication using credentials from a client
        secrets file. It saves the credentials in a token file for reuse in future sessions.
        If valid credentials are found in the token file, they are refreshed and used;
        otherwise, new credentials are obtained via OAuth 2.0.

        Returns:
            googleapiclient.discovery.Resource: A service object for the Google Calendar API.

        Raises:
            FileNotFoundError: If the client secrets file is not found.
            google.auth.exceptions.RefreshError: If the credentials refresh fails.
        """

        creds = None
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "client_secret.json", SCOPES
                )
                creds = flow.run_local_server(port=0)
            with open("token.json", "w", encoding='utf-8') as token:
                token.write(creds.to_json())

        return build("calendar", "v3", credentials=creds)


class GoogleCalendarManager:
    """
    Manages Google Calendar operations by utilizing the authenticated service.

    This class is responsible for managing operations on Google Calendar, such as
    creating, updating, and deleting events. It uses an authenticated service
    obtained through the GoogleCalendarAuth class to interact with the Google Calendar API.

    Attributes:
        service (googleapiclient.discovery.Resource): An authenticated service
            object for interacting with the Google Calendar API.
    """

    def __init__(self):
        """
        Initializes a new instance of GoogleCalendarManager.

        The constructor authenticates with Google Calendar using GoogleCalendarAuth
        and initializes the service attribute, which will be used for subsequent
        Google Calendar operations.
        """
        self.service = GoogleCalendarAuth.authenticate()

    def enable_out_of_office(self, start_date, end_date):
        """
        Enable an "Out of Office" event on the team calendar for specified dates.

        This method creates an "Out of Office" event on the team calendar for the
        specified date range. It first checks if an event for the given date range
        already exists to avoid duplicates. If no such event exists, it creates a new
        event with the specified summary and description.

        Parameters:
            start_date (str): The start date of the "Out of Office" event in 'YYYY-MM-DD' format.
            end_date (str): The end date of the "Out of Office" event in 'YYYY-MM-DD' format.

        Returns:
            None

        Raises:
            Exception: If an error occurs while creating the event in the Google Calendar API.

        Notes:
            - The method adjusts the end date by adding one day to ensure the "Out of Office"
              status for the entire duration of the end date.
            - The method uses configurations like 'default_team_calendar', 'timezone',
              'default_personal_calendar', and 'ooo_pattern' from a configuration method
              (get_default_config).
            - Warnings and errors are printed to the console if the event already exists or
              if an error occurs while creating the event.
        """
        team_calendar_name = get_default_config("default_team_calendar")
        timezone = get_default_config("timezone")
        username = get_default_config("default_personal_calendar").split("@")[0]
        ooo_pattern = get_default_config("ooo_pattern").strip()
        event_summary = f"{username} {ooo_pattern}"

        # Get the calendar ID from the calendar name
        team_calendar_id = self.get_calendar_id_by_name(team_calendar_name)

        # Check if the event already exists
        adjusted_end_date = (
            datetime.strptime(end_date, "%Y-%m-%d") + timedelta(days=1)
        ).strftime("%Y-%m-%d")
        if self.event_exists(
            team_calendar_id, start_date, adjusted_end_date, event_summary
        ):
            console.print(
                f"[yellow]Warning: OOO event for {start_date} to {end_date} already exists on the team calendar - {team_calendar_name}.  Skipping action.[/yellow]"
            )
            return

        # Create the OOO event
        event = {
            "summary": event_summary,
            "description": "Out of Office",
            "start": {
                "date": start_date,
                "timeZone": timezone,
            },
            "end": {
                "date": adjusted_end_date,
                "timeZone": timezone,
            },
            "transparency": "opaque",
            "visibility": "default",
            "status": "confirmed",
            "eventType": "default",
        }

        calendar_id = self.get_calendar_id_by_name(team_calendar_name)
        if not calendar_id:
            console.print(
                f"[red]Error:[/red] Calendar '{team_calendar_name}' not found."
            )
            return

        try:
            created_event = (
                self.service.events()
                .insert(calendarId=calendar_id, body=event)
                .execute()
            )
            console.print(
                f"[cyan]OutOfOffice event created (Id: {created_event['id']} from[/cyan] [yellow]{start_date}[/yellow] [cyan]to[/cyan] [yellow]{end_date}[/yellow][cyan] on calendar [/cyan][blue] {team_calendar_name}[/blue]"
            )
        except Exception as e:
            console.print(f"[red]Error:[/red] {e}")

    def event_exists(self, calendar_id, start_date, end_date, event_summary):
        """
        Check if an "Out of Office" event already exists for the given date range.

        Parameters:
            calendar_id (str): The ID of the calendar to check.
            start_date (str): The start date in 'YYYY-MM-DD' format.
            end_date (str): The end date in 'YYYY-MM-DD' format.
            event_summary (str): The summary text to search for in events.

        Returns:
            bool: True if an event with the specified summary exists within
                the date range, False otherwise.

        Notes:
            - The method searches for events that exactly match the provided event_summary.
            - The search is case-sensitive.
        """
        existing_events = (
            self.service.events()
            .list(
                calendarId=calendar_id,
                timeMin=start_date + "T00:00:00Z",
                timeMax=end_date + "T00:00:00Z",
                q=event_summary,  # Search for OOO events
            )
            .execute()
            .get("items", [])
        )

        return any(event.get("summary") == event_summary for event in existing_events)

    def is_ooo_today(self, team_member=None):
        """
        Check if the user or a specified team member is marked "Out of Office" today.

        Parameters:
            team_member (str, optional): The username of the team member to check.
                If None, checks the default user. Defaults to None.

        Returns:
            None

        Notes:
            - The method prints to the console whether the user/team member is out of office.
            - The method considers weekends as automatic "Out of Office" days and prints a
              message to the console if today is a weekend.
            - The method uses configurations like 'default_team_calendar',
              'default_personal_calendar', and 'ooo_pattern' from a configuration method
              (get_default_config).
        """
        team_calendar_name = get_default_config("default_team_calendar")
        team_events = self.get_upcoming_ooo_events(team_calendar_name, max_results=100)

        if team_member:
            username = team_member
        else:
            username = get_default_config("default_personal_calendar").split("@")[0]

        ooo_pattern = get_default_config("ooo_pattern").strip()
        event_summary = f"{username} {ooo_pattern}"
        today = datetime.now().date()

        # Check if today is a weekend
        if today.weekday() == 5 or today.weekday() == 6:
            console.print(
                f"[green]User {username} is Out of Office today due to the weekend.[/green]"
            )
            return

        for event in team_events:
            if event_summary in event.get("summary", ""):
                start = self.get_date_from_event(event, "start")
                end = self.get_date_from_event(event, "end")

                if start <= today <= end:
                    console.print(
                        f"[green]User {username} is Out of Office today.[/green]"
                    )
                    return

        console.print(f"[yellow]User {username} is not Out of Office today.[/yellow]")

    def check_out_of_office(self, max_results=10):
        """
        List upcoming "Out of Office" events for the next month in a compact view.

        This method retrieves and prints upcoming OOO events for the next month from
        the team calendar. It displays the event's date range, summary, event ID,
        event type, and the calendar name in a compact view in the console.

        Parameters:
            max_results (int, optional): The maximum number of upcoming events to retrieve.
                Defaults to 10.

        Returns:
            None

        Notes:
            - The method prints each event's details directly to the console.
            - If there are no upcoming OOO events, a message is printed to the console.
            - The method uses configurations like 'default_team_calendar' from a configuration
              method (get_default_config).
            - The method adjusts the end date for multi-day events to ensure accurate display.
        """
        team_calendar_name = get_default_config("default_team_calendar")
        team_events = self.get_upcoming_ooo_events(team_calendar_name, max_results)

        if not team_events:
            console.print("[yellow]No upcoming OOO events found.[/yellow]")
            return

        for event in team_events:
            start = self.get_date_from_event(event, "start")
            end = self.get_date_from_event(event, "end")
            summary = event.get("summary", "")
            event_id = event.get("id", "")
            event_type = event.get("eventType", "")

            # Adjust the end date for multi-day events
            if start != end:
                adjusted_end_date = end - timedelta(days=1)
            else:
                adjusted_end_date = end

            start_str = start.strftime("%Y-%m-%d")
            end_str = adjusted_end_date.strftime("%Y-%m-%d")

            console.print(
                f"☀️ 🏖️ 🌴 [green]{start_str} to {end_str} - {summary} (Event ID: {event_id}, Type: {event_type}) on [/green] [blue]{team_calendar_name}[/blue]"
            )

    def get_upcoming_ooo_events(self, team_calendar_name, max_results=100):
        """
        Fetch upcoming "Out of Office" events for a specified calendar within the next month.

        This method retrieves OOO events from the specified team calendar that occur
        from today until the last day of the next month. It uses the Google Calendar API
        to fetch events and returns them as a list of event objects.

        Parameters:
            team_calendar_name (str): The name of the team calendar to fetch events from.
            max_results (int, optional): The maximum number of events to retrieve.
                Defaults to 100.

        Returns:
            list: A list of upcoming OOO event objects. Each object contains details
                about an event, such as its summary, start time, and end time. Returns
                an empty list if no events are found or an error occurs.

        Raises:
            HttpError: If an HTTP error occurs while fetching events from the API.
            Exception: If an unexpected error occurs.

        Notes:
            - The method uses configurations like 'default_personal_calendar' and 'ooo_pattern'
              from a configuration method (get_default_config).
            - The method prints error messages directly to the console if the calendar is not
              found or if an error occurs while fetching events.
            - The method constructs a query to fetch events that match a specific summary pattern.
        """
        calendar_id = self.get_calendar_id_by_name(team_calendar_name)
        if not calendar_id:
            console.print(
                f"[red]Error:[/red] Calendar '{team_calendar_name}' not found."
            )
            return []

        today = dt.datetime.utcnow().date()
        first_day_next_month = (today.replace(day=1) + dt.timedelta(days=32)).replace(
            day=1
        )
        last_day_next_month = (first_day_next_month + dt.timedelta(days=32)).replace(
            day=1
        ) - dt.timedelta(days=1)

        start_date_str = f"{today}T00:00:00Z"
        end_date_str = f"{last_day_next_month}T23:59:59Z"

        username = get_default_config("default_personal_calendar").split("@")[0]
        ooo_pattern = get_default_config("ooo_pattern").strip()
        event_summary = f"{username} {ooo_pattern}"

        try:
            events_result = (
                self.service.events()
                .list(
                    calendarId=calendar_id,
                    timeMin=start_date_str,
                    timeMax=end_date_str,
                    maxResults=max_results,
                    singleEvents=True,
                    orderBy="startTime",
                    q=event_summary,
                )
                .execute()
            )
            events = events_result.get("items", [])
            return events

        except HttpError as error:
            console.print(f"[red]An error occurred:[/red] {error}")
            return []
        except Exception as e:
            console.print(f"[red]Unexpected error:[/red] {e}")
            return []

    def get_date_from_event(self, event, time_key):
        """
        Extract the date from an event's start or end time, considering the local timezone.

        This method retrieves the date from the specified time key ('start' or 'end') of
        an event object. If the time information includes a timezone, the method converts
        the date to the local timezone. If the time information does not include a specific
        time (only a date), it returns the date without conversion.

        Parameters:
            event (dict): The event object containing details about a Google Calendar event.
            time_key (str): The key to retrieve time information from the event.
                Expected values are 'start' or 'end'.

        Returns:
            datetime.date: The extracted date from the event's specified time key.

        Notes:
            - The method expects the event object to contain valid time information under
              the specified time_key.
            - The method handles two formats of time information:
                1. 'dateTime': A string in ISO 8601 format, which includes specific time and
                   optionally timezone information.
                2. 'date': A string in 'YYYY-MM-DD' format, representing a full day without
                   specific time.
            - If 'dateTime' is provided and includes timezone information, the method converts
              the date to the local timezone before returning it.
        """
        time_info = event[time_key]
        if "dateTime" in time_info:
            local_tz = tz.gettz(event[time_key]["timeZone"])
            return (
                dt.datetime.fromisoformat(time_info["dateTime"])
                .astimezone(local_tz)
                .date()
            )
        if "date" in time_info:
            return dt.datetime.strptime(time_info["date"], "%Y-%m-%d").date()

    def get_calendar_id_by_name(self, team_calendar_name):
        """
        Retrieve the calendar ID based on its name.

        This method attempts to retrieve the ID of a Google Calendar given its name.
        Initially, it tries to get the calendar directly using the provided name as an ID.
        If this fails (e.g., due to an invalid ID or lack of permissions), it iterates
        through all available calendars in the calendar list to find a calendar with a
        matching name.

        Parameters:
            team_calendar_name (str): The name of the team calendar to retrieve the ID for.

        Returns:
            str or None: The ID of the calendar if found; otherwise, None.

        Notes:
            - The method prints error messages directly to the console if the calendar is not
              found or if an error occurs while fetching calendar data.
            - Two attempts are made to retrieve the calendar ID:
                1. Direct retrieval using the name as an ID.
                2. Searching through all available calendars.
            - If both attempts fail, the method returns None and prints an error message.
        """
        try:
            calendar = (
                self.service.calendars().get(calendarId=team_calendar_name).execute()
            )
            if calendar:
                return calendar["id"]
        except Exception:
            try:
                calendar_list = self.service.calendarList().list().execute()

                for calendar in calendar_list["items"]:
                    if calendar["summary"] == team_calendar_name:
                        return calendar["id"]
            except Exception as e2:
                console.print(f"[red]Error:[/red] {e2}")

        console.print(f"[red]Error:[/red] Calendar '{team_calendar_name}' not found.")
        return None

    def disable_out_of_office(self):
        """
        Disable (delete) "Out of Office" events for the user on the team calendar.

        This method attempts to find and delete OOO events for the user on the team
        calendar. It retrieves the user's upcoming OOO events and deletes any event
        that matches the specified summary pattern. If an event is successfully deleted,
        a success message is printed to the console. If no matching events are found,
        a different message is printed.

        Parameters:
            None

        Returns:
            None

        Notes:
            - The method uses configurations like 'default_team_calendar',
              'default_personal_calendar', and 'ooo_pattern' from a configuration
              method (get_default_config).
            - The method prints messages directly to the console to inform about the success
              or failure of the operation.
            - If an error occurs while deleting an event, an error message is printed, and
              the method returns early.
            - If no events are found, a message is printed, and the method completes normally.
        """
        team_calendar_name = get_default_config("default_team_calendar")
        username = get_default_config("default_personal_calendar").split("@")[0]
        ooo_pattern = get_default_config("ooo_pattern").strip()
        event_summary = f"{username} {ooo_pattern}"
        calendar_id = self.get_calendar_id_by_name(team_calendar_name)

        # Search for the event
        events = self.get_upcoming_ooo_events(team_calendar_name)
        for event in events:
            if event.get("summary") == event_summary:
                try:
                    self.service.events().delete(
                        calendarId=calendar_id, eventId=event["id"]
                    ).execute()
                    console.print(
                        f"[green]Successfully disabled Out of Office for {username} on {team_calendar_name}.[/green]"
                    )
                except Exception as e:
                    console.print(f"[red]Error:[/red] {e}")
                    return

        if not events:
            console.print(
                f"[yellow]No Out of Office event found for {username} \
                on {team_calendar_name}.[/yellow]"
            )


def get_default_config(key):
    """
    Load a default configuration value from the config.ini file.

    Given a key, this function retrieves the corresponding value from the
    'DEFAULT' section of a configuration file named 'config.ini'. If the key
    is not found, it returns None.

    Parameters:
        key (str): The key for which to retrieve the value from the configuration file.

    Returns:
        str or None: The value corresponding to the provided key in the 'DEFAULT'
            section of the configuration file. If the key is not found, returns None.

    Notes:
        - The function uses the configparser module to parse the 'config.ini' file.
        - The function does not handle exceptions that might be raised by configparser
          (e.g., if the file does not exist or is not properly formatted). Ensure that
          the 'config.ini' file is available and correctly formatted in the same directory
          as the script.
        - If the key is not found in the 'DEFAULT' section, the function returns None
          without raising an error.
    """
    config = configparser.ConfigParser()
    config.read("config.ini")
    return config["DEFAULT"].get(key, None)


def main():
    """
    Main function to handle command-line interactions for managing Google Calendar events.

    This function sets up argument parsing for a command-line tool that allows users to
    interact with Google Calendar, particularly for managing "Out of Office" (OOO) events.
    It supports various operations like checking OOO status, enabling, and disabling OOO
    events, with optional parameters for specifying team members and date ranges.

    Supported Command-Line Arguments:
        --check-outofoffice: Check if a team member is Out of Office.
        --is-ooo-today: Check if the user is Out of Office today.
        --team-member [NAME]: Specify the team member's name to check their OOO status.
        --enable-outofoffice: Enable Out of Office for the specified dates.
        --start-date [YYYY-MM-DD]: Specify the start date for the OOO event.
        --end-date [YYYY-MM-DD]: Specify the end date for the OOO event.
        --disable-outofoffice: Disable Out of Office.

    Notes:
        - The function uses argparse to parse command-line arguments and determine the
          appropriate action to take.
        - If no arguments are provided, the function prints the help message and exits.
        - The function creates an instance of GoogleCalendarManager to interact with
          Google Calendar and perform the requested operations.
        - Error messages are printed to the console if required arguments for an operation
          are missing.
    """
    parser = argparse.ArgumentParser(description="Google Calendar Command Line Tool")
    parser.add_argument(
        "--check-outofoffice",
        action="store_true",
        help="Check if a team member is Out of Office",
    )
    parser.add_argument(
        "--is-ooo-today",
        action="store_true",
        help="Check if the user is Out of Office today",
    )
    parser.add_argument(
        "--team-member",
        type=str,
        help="Specify the team member's name to check their OOO status",
    )
    parser.add_argument(
        "--enable-outofoffice",
        action="store_true",
        help="Enable Out of Office for the specified dates",
    )
    parser.add_argument(
        "--start-date",
        type=str,
        help="Specify the start date for the OOO event (format: YYYY-MM-DD)",
    )
    parser.add_argument(
        "--end-date",
        type=str,
        help="Specify the end date for the OOO event (format: YYYY-MM-DD)",
    )
    parser.add_argument(
        "--disable-outofoffice", action="store_true", help="Disable Out of Office"
    )

    args = parser.parse_args()

    calendar = GoogleCalendarManager()

    if len(sys.argv) == 1:
        parser.print_help(sys.stderr)
        sys.exit(1)

    if args.check_outofoffice:
        calendar.check_out_of_office()

    if args.disable_outofoffice:
        calendar.disable_out_of_office()

    if args.is_ooo_today:
        calendar.is_ooo_today(args.team_member)

    if args.enable_outofoffice:
        if args.start_date and args.end_date:
            calendar.enable_out_of_office(args.start_date, args.end_date)
        else:
            console.print(
                "[red]Please specify both --start-date and --end-date for the OOO event.[/red]"
            )


if __name__ == "__main__":
    main()
