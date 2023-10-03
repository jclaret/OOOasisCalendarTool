# **Overview**:
`oooasis.py` is a command-line tool designed to manage Google Calendar events, specifically focusing on Out of Office (OOO) events. The script provides functionalities to check, enable, and disable OOO events for a user or a team member.

---

## **Prerequisites**:

1. **Python**: Ensure you have Python installed on your local machine.
2. **Google API Client Library**
3. **Rich Library**

## **Functionality**:

1. **Authentication**:
   - The script uses OAuth 2.0 to authenticate with the Google Calendar API.
   - On the first run, it will prompt you to authorize access. Once authorized, it will save the token in `token.json` for subsequent runs.

2. **OOO Event Management**:
   - The script can enable, check, and disable OOO events.
   - It can check if a user or a specified team member is OOO on a given day.

3. **Event Types**:
   - Currently, only "default" and "workingLocation" events can be created using the API.
   - Note: Extended support for other event types like `outOfOffice` will be made available in later releases.

## **Setup**:

1. **Google API Credentials**:
- Visit the [Google Developer Console](https://console.cloud.google.com/cloud-resource-manager) and create a new project.
- Enable the Google Calendar API for the project.
- Create OAuth 2.0 client IDs and download the `client_secret.json` file.
- Place the `client_secret.json` file in the same directory as `oooasis.py`.

2. **Configuration File**:
- Create a `config.ini` file in the same directory as `oooasis.py`.
- Add default configurations like `default_team_calendar`, `timezone`, and `default_personal_calendar`.

3. **Install Dependencies**:
```
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### **Usage**:

Run the script using the command:
**Available Options**:
- `--check-outofoffice`: Check if a team member is Out of Office.
- `--is-ooo-today`: Check if the user is Out of Office today.
- `--team-member [NAME]`: Specify the team member's name to check their OOO status.
- `--enable-outofoffice`: Enable Out of Office for the specified dates. Requires `--start-date` and `--end-date`.
- `--start-date [YYYY-MM-DD]`: Specify the start date for the OOO event.
- `--end-date [YYYY-MM-DD]`: Specify the end date for the OOO event.
- `--disable-outofoffice`: Disable Out of Office.

### **Examples**:

- Enable Out of Office from 2023-10-09 to 2023-10-12
```
$ oooasis.py --enable-outofoffice --start-date 2023-10-09 --end-date 2023-10-12
OutOfOffice event created (Id: 541gmpjbgdjstobc4khfceoi50 from 2023-10-09 to 2023-10-12 on calendar rh-eng-telco5g-integration
```

- Check if there's Out of Office
```
$ oooasis.py --check-outofoffice
☀️ 🏖️ 🌴 2023-10-09 to 2023-10-12 - jclaretm -- PTO (Event ID: 541gmpjbgdjstobc4khfceoi50, Type: default) on rh-eng-telco5g-integration
```

- Check if a team member is Out of Office today
```
$ oooasis.py --is-ooo-today --team-member jclaretm
User jclaretm is not Out of Office today.
```

- Disable Out of Office will delete events 
```
$ oooasis.py --disable-outofoffice
Successfully disabled Out of Office for jclaretm on rh-eng-telco5g-integration.
```

### **References**:
- [Google Calendar API Documentation](https://developers.google.com/calendar/api/v3/reference/events/insert)
- [Google Issue Tracker](https://issuetracker.google.com/issues/112063903)


**Note**: Always ensure you have the latest version of the script and the required libraries for optimal functionality.
