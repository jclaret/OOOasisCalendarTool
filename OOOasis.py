#! /usr/bin/python -W ignore
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
    """Handles Google Calendar Authentication."""

    @staticmethod
    def authenticate():
        """Authenticate and return the service object."""
        creds = None
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", SCOPES)
                creds = flow.run_local_server(port=0)
            with open("token.json", "w") as token:
                token.write(creds.to_json())

        return build("calendar", "v3", credentials=creds)


class GoogleCalendarManager:
    """Manages Google Calendar operations."""

    def __init__(self):
        self.service = GoogleCalendarAuth.authenticate()

    # Enable OOO
    def enable_out_of_office(self, start_date, end_date):
        """Enable Out of Office for the specified dates."""
        team_calendar_name = get_default_config('default_team_calendar')
        timezone = get_default_config('timezone')
        username = get_default_config('default_personal_calendar').split('@')[0]
        ooo_pattern = get_default_config('ooo_pattern').strip()
        event_summary = f"{username} {ooo_pattern}"

        # Get the calendar ID from the calendar name
        team_calendar_id = self.get_calendar_id_by_name(team_calendar_name)
 
        # Check if the event already exists
        adjusted_end_date = (datetime.strptime(end_date, '%Y-%m-%d') + timedelta(days=1)).strftime('%Y-%m-%d')
        if self.event_exists(team_calendar_id, start_date, adjusted_end_date, event_summary):
            console.print(f"[yellow]Warning: OOO event for {start_date} to {end_date} already exists on the team calendar - {team_calendar_name}. Skipping action.[/yellow]")
            return
        
        # Create the OOO event
        event = {
            'summary': event_summary,
            'description': 'Out of Office',
            'start': {
                'date': start_date,
                'timeZone': timezone,
            },
            'end': {
                'date': adjusted_end_date,
                'timeZone': timezone,
            },
            'transparency': 'opaque',
            'visibility': 'default',
            'status': 'confirmed',
            'eventType': 'default'
        }

        calendar_id = self.get_calendar_id_by_name(team_calendar_name)
        if not calendar_id:
            console.print(f"[red]Error:[/red] Calendar '{team_calendar_name}' not found.")
            return
    
        try:
            created_event = self.service.events().insert(calendarId=calendar_id, body=event).execute()
            console.print(f"[cyan]OutOfOffice event created (Id: {created_event['id']} from[/cyan] [yellow]{start_date}[/yellow] [cyan]to[/cyan] [yellow]{end_date}[/yellow][cyan] on calendar [/cyan][blue]{team_calendar_name}[/blue]")
        except Exception as e:
            console.print(f"[red]Error:[/red] {e}")

    def event_exists(self, calendar_id, start_date, end_date, event_summary):
        """Check if an OOO event already exists for the given date range."""
        existing_events = self.service.events().list(
            calendarId=calendar_id,
            timeMin=start_date + "T00:00:00Z",
            timeMax=end_date + "T00:00:00Z",
            q=event_summary  # Search for OOO events
        ).execute().get('items', []) 
    
        return any(event.get('summary') == event_summary for event in existing_events)

    # Check OOO today
    def is_ooo_today(self, team_member=None):
        """Check if the user or specified team member is Out of Office today."""
        team_calendar_name = get_default_config('default_team_calendar')
        team_events = self.get_upcoming_ooo_events(team_calendar_name, max_results=100)
    
        if team_member:
            username = team_member
        else:
            username = get_default_config('default_personal_calendar').split('@')[0]
    
        ooo_pattern = get_default_config('ooo_pattern').strip()
        event_summary = f"{username} {ooo_pattern}"
        today = datetime.now().date()
    
        # Check if today is a weekend
        if today.weekday() == 5 or today.weekday() == 6:
            console.print(f"[green]User {username} is Out of Office today due to the weekend.[/green]")
            return
    
        for event in team_events:
            if event_summary in event.get('summary', ''):
                start = self.get_date_from_event(event, 'start')
                end = self.get_date_from_event(event, 'end')
    
                if start <= today <= end:
                    console.print(f"[green]User {username} is Out of Office today.[/green]")
                    return
    
        console.print(f"[yellow]User {username} is not Out of Office today.[/yellow]")

    # Check OOO events
    def check_out_of_office(self, max_results=10):
        """List upcoming OOO events for the next month in a compact view for both personal and team calendars."""
        team_calendar_name = get_default_config('default_team_calendar')
        team_events = self.get_upcoming_ooo_events(team_calendar_name, max_results)

        if not team_events:
            console.print("[yellow]No upcoming OOO events found.[/yellow]")
            return

        for event in team_events:
            start = self.get_date_from_event(event, 'start')
            end = self.get_date_from_event(event, 'end')
            summary = event.get('summary', '')
            event_id = event.get('id', '')
            calendar_name = event.get('organizer', {}).get('displayName', '')
            event_type = event.get('eventType', '')
        
            # Adjust the end date for multi-day events
            if start != end:
                adjusted_end_date = end - timedelta(days=1)
            else:
                adjusted_end_date = end
        
            start_str = start.strftime('%Y-%m-%d')
            end_str = adjusted_end_date.strftime('%Y-%m-%d')
        
            console.print(f"☀️ 🏖️ 🌴 [green]{start_str} to {end_str} - {summary} (Event ID: {event_id}, Type: {event_type}) on [/green][blue]{team_calendar_name}[/blue]")

    def get_upcoming_ooo_events(self, team_calendar_name, max_results=100):
        """Fetch upcoming OOO events for a specified calendar."""
        calendar_id = self.get_calendar_id_by_name(team_calendar_name)
        if not calendar_id:
            console.print(f"[red]Error:[/red] Calendar '{team_calendar_name}' not found.")
            return []

        today = dt.datetime.utcnow().date()
        first_day_next_month = (today.replace(day=1) + dt.timedelta(days=32)).replace(day=1)
        last_day_next_month = (first_day_next_month + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)

        start_date_str = f"{today}T00:00:00Z"
        end_date_str = f"{last_day_next_month}T23:59:59Z"

        username = get_default_config('default_personal_calendar').split('@')[0]
        ooo_pattern = get_default_config('ooo_pattern').strip()
        event_summary = f"{username} {ooo_pattern}"

        try:
            events_result = self.service.events().list(
                calendarId=calendar_id, timeMin=start_date_str, timeMax=end_date_str,
                maxResults=max_results, singleEvents=True,
                orderBy='startTime', q=event_summary
            ).execute()
            events = events_result.get('items', [])
            return events

        except HttpError as error:
            console.print(f"[red]An error occurred:[/red] {error}")
            return []
        except Exception as e:
            console.print(f"[red]Unexpected error:[/red] {e}")
            return []

    def get_date_from_event(self, event, time_key):
        """Extract the date from an event's start or end time."""

    def get_date_from_event(self, event, time_key):
        """Extract the date from an event's start or end time in local timezone."""
        time_info = event[time_key]
        if 'dateTime' in time_info:
            local_tz = tz.gettz(event[time_key]['timeZone'])
            return dt.datetime.fromisoformat(time_info['dateTime']).astimezone(local_tz).date()
        elif 'date' in time_info:
            return dt.datetime.strptime(time_info['date'], '%Y-%m-%d').date()

    def get_calendar_id_by_name(self, team_calendar_name):
        """Retrieve the calendar ID based on its name."""
        try:
            calendar = self.service.calendars().get(calendarId=team_calendar_name).execute()
            if calendar:
                return calendar['id']
        except Exception:
            try:
                calendar_list = self.service.calendarList().list().execute()

                for calendar in calendar_list['items']:
                    if calendar['summary'] == team_calendar_name:
                        return calendar['id']
            except Exception as e2:
                console.print(f"[red]Error:[/red] {e2}")

        console.print(f"[red]Error:[/red] Calendar '{team_calendar_name}' not found.")
        return None

    # Disable OOO
    def disable_out_of_office(self):
        """Disable Out of Office for the user."""
        team_calendar_name = get_default_config('default_team_calendar')
        username = get_default_config('default_personal_calendar').split('@')[0]
        ooo_pattern = get_default_config('ooo_pattern').strip()
        event_summary = f"{username} {ooo_pattern}"
        calendar_id = self.get_calendar_id_by_name(team_calendar_name)

        # Search for the event
        events = self.get_upcoming_ooo_events(team_calendar_name)
        for event in events:
            if event.get('summary') == event_summary:
                try:
                    self.service.events().delete(calendarId=calendar_id, eventId=event['id']).execute()
                    console.print(f"[green]Successfully disabled Out of Office for {username} on {team_calendar_name}.[/green]")
                except Exception as e:
                    console.print(f"[red]Error:[/red] {e}")
                    return

        if not events:
            console.print(f"[yellow]No Out of Office event found for {username} on {team_calendar_name}.[/yellow]")

def get_default_config(key):
    """Load default configurations from the config.ini file."""
    config = configparser.ConfigParser()
    config.read('config.ini')
    return config['DEFAULT'].get(key, None)

def main():
    """Main function."""
    parser = argparse.ArgumentParser(description="Google Calendar Command Line Tool")
    parser.add_argument("--check-outofoffice", action="store_true", help="Check if a team member is Out of Office")
    parser.add_argument("--is-ooo-today", action="store_true", help="Check if the user is Out of Office today")
    parser.add_argument("--team-member", type=str, help="Specify the team member's name to check their OOO status")
    parser.add_argument("--enable-outofoffice", action="store_true", help="Enable Out of Office for the specified dates")
    parser.add_argument("--start-date", type=str, help="Specify the start date for the OOO event (format: YYYY-MM-DD)")
    parser.add_argument("--end-date", type=str, help="Specify the end date for the OOO event (format: YYYY-MM-DD)")
    parser.add_argument("--disable-outofoffice", action="store_true", help="Disable Out of Office")

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
            console.print("[red]Please specify both --start-date and --end-date for the OOO event.[/red]")

if __name__ == "__main__":
    main()
