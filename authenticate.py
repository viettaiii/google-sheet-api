import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow



class AuthenticateGoogleApi():
    def __init__(self, creds_file, scopes) -> None:
        self.creds_file = creds_file
        self.creds = None
        self.scopes = scopes

    def authenticate(self):
        if os.path.exists("token.json"):
            self.creds = Credentials.from_authorized_user_file("token.json", self.scopes)
    # If there are no (valid) credentials available, let the user log in.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", self.scopes
                )
            self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open("token.json", "w") as token:
                token.write(self.creds.to_json())




# def main():
#     SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
#     creds_file = "credentials.json"
#     sheet_api = AuthenticateGoogleApi(creds_file, SCOPES)
#     service = sheet_api.build_service("sheets", "v4")

#     SAMPLE_SPREADSHEET_ID = "1_opfC9pqtLOhkcJrqxN-YQrD3R-TVaP-tHJyC1Zg32E"
#     SAMPLE_RANGE_NAME = "'Sheet1'!A1:C5"

#     try:
#         sheet = service.spreadsheets()
#         result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
#         values = result.get("values", [])

#         if not values:
#             print("No data found.")
#             return

#         print("Name, Major:")
#         print(values)
#     except HttpError as err:
#         print(err)


# main()
