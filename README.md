# **Overview**:
`oooasis.py` is a command-line tool designed to manage Google Calendar events, specifically focusing on Out of Office (OOO) events. The script provides functionalities to check, enable, and disable OOO events for a user or a team member.

---
## **Functionality**:

1. **Authentication**:
   - The script uses OAuth 2.0 to authenticate with the Google Calendar API.
   - On the first run, it will prompt you to authorize access. Once authorized, it will save the token in `token.json` for subsequent runs.

2. **OOO Event Management**:
   - The script can enable, check, and disable OOO events.
   - It can check if a user or a specified team member is OOO on a given day.

3. **Event Types**:
   - Currently, only "default" and "workingLocation" events can be created using the API.
   - Support for event types like `outOfOffice` will be made available in later API releases. See references.

## **Set up your environment**:

To complete this, [set up your environment](https://developers.google.com/calendar/api/quickstart/python#set_up_your_environment)

### Enable the API

Before using Google APIs, you need to turn them on in a Google Cloud project. 

[Enable the API](https://console.cloud.google.com/flows/enableapi?apiid=calendar)

NOTE: You might need to create a project if you don't already have one

### Configure the OAUTH consent screen

1.  In the Google Cloud console, go to **Menu > APIs & Services > OAuth** consent screen. [Go to OAuth consent screen](https://console.cloud.google.com/apis/credentials/consent)
   
2. Select the user type for your app, then click `Create`.
   
3. Complete the app registration form, then click `Save and Continue`.
   
4. Add "Google Calendar API" scope and click `Save and Continue`.
   
5. If you selected `External` for user type, add test users:
   - Under `Test users`, click `Add users`.
   - Enter your email address and any other authorized test users, then click `Save and Continue`.
   
6. Review your app registration summary. To make changes, click `Edit`. If the app registration looks OK, click `Back to Dashboard`.

### Authorize credentials for a desktop Application

To authenticate as an end user and access user data in your app, you need to create one OAuth 2.0 Client IDs. A client ID is used to identify a single app to Google's OAuth servers.

1. In the Google Cloud console, go to Menu > APIs & Services > Credentials. [Go to Credentials](https://console.cloud.google.com/apis/credentials)
   
2. Click `Create Credentials` > `OAuth client ID`.
   
3. Click `Application type` > `Desktop app`.
   
4. In the `Name` field, type a name for the credential. This name is only shown in the Google Cloud console.
   
5. Click `Create`. The OAuth client created screen appears, showing your new Client ID and Client secret.
   
6. Click `OK`. The newly created credential appears under `OAuth 2.0 Client IDs`.
   
7. Save the downloaded JSON file as `client_secret.json`, and move the file in the same directory as `oooasis.py`..


### Create Configuration File

- Create a `config.ini` file in the same directory as `oooasis.py`.
- Add default configurations like `default_team_calendar`, `timezone`, and `default_personal_calendar`.

### Install Dependencies

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
- [Google Issue - Create out of office event in Calendar API](https://issuetracker.google.com/issues/112063903)


**Note**: Always ensure you have the latest version of the script and the required libraries for optimal functionality.
