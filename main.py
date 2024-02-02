import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def main():

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

    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("sheets", "v4", credentials=creds)


    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId="1JTMcAD1EVfWWzYnDdWo83JX2kUzdRGdikOLjpYUqHqk", range="engenharia_de_software!A4:H27")
        .execute()
    )
    values = result.get("values", [])

    status = ""
    formatGradeForFinalApproval = 0
    number = 3

    for row in values:
      test1 = float(row[3]) if row[3] else 0
      test2 = float(row[4]) if row[3] else 0
      test3 = float(row[5]) if row[3] else 0
      schoolAbsences = int(row[2])

      sum = ((test1 + test2 + test3)/3)/10

      if sum >= 5 and sum < 7:
        gradeForFinalApproval = 7 - sum
        formatGradeForFinalApproval = "{:.1f}".format(round(gradeForFinalApproval, 1))

      if schoolAbsences > 15:
        status = "Reprovado Por falta"
      elif sum < 5:
        status = "Reprovado por Nota"
      elif sum >= 5 and sum < 7:
        status = "Exame Final"
      elif sum >= 7:
        status = "Aprovado"

      put = [
        [status, formatGradeForFinalApproval]
      ]
      number = number + 1
      result = (
        sheet.values()
        .update(spreadsheetId="1JTMcAD1EVfWWzYnDdWo83JX2kUzdRGdikOLjpYUqHqk",
                range=f"engenharia_de_software!G{number}", valueInputOption="USER_ENTERED",
                body={"values": put})
        .execute()
      )


  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()