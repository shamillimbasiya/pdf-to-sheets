import os, io, re
from math import ceil

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

import PyPDF2

"""
Krav för att skriptet ska funka:
1. Första artikeln måste börja på rad A10
2. Kvitto ligger i h columnen
3. Det är 70 glas i en keg 
"""

"https://drive.google.com/file/d/16XafdDsQUjC-50SVe3yPQkIv7MBJSCGP/view?usp=drive_link"


SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.readonly"]

SPREADSHEET_ID = "1yzoMIHl9hVgLavMRW62CRCkSF7qFHJpyM7Fyb4yVLjk"

def handleCredentials():
    credentials = None
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else: 
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(credentials.to_json())
    return credentials

def downloadPDF(creds):
    try:
        url = input("Input url: ")
        file_id = url.split("/")[-2]
        drive_service = build("drive", "v3", credentials=creds)
        request = drive_service.files().get_media(fileId=file_id)
        file = io.BytesIO()
        downloader = MediaIoBaseDownload(file, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print(f"Download {int(status.progress() * 100)}.")

    except HttpError as error:
        print(f"An error occurred: {error}")
        file = None
    file.seek(0)
    with open(os.path.join("dowloaded_files", "artikel_rapport.pdf"), "wb") as f:
        f.write(file.read())
        f.close()

def extract_text(pdf_file):
    with open(pdf_file, "rb") as pdf:
        reader = PyPDF2.PdfReader(pdf, strict=False)
        pdf_text = ""

        for page in reader.pages:
            content = page.extract_text()
            pdf_text += content

        #Filter out the sought data
        pattern = re.compile(r'(\b[\w/]+(\s+\w+)*\s+\d+\s+\d+,\d+\b)')
        matches = re.findall(pattern, pdf_text)

        #Divide the data so it is accesible 
        result_dict = {}
        for match in matches:
            info = match[0].split()
            result_dict[' '.join(info[:-2]).lower()] = {"quantity" : int(info[-2]), "price": float(info[-1].replace(',', '.'))}
            # result_dict.append({
            #     'name': ' '.join(info[:-2]),  # Join the name parts
            #     'quantity': int(info[-2]),
            #     'price': float(info[-1].replace(',', '.'))  # Replace comma with dot for decimal part
            # })
        return result_dict


def main():
    credentials = handleCredentials()
    
    try:
        downloadPDF(credentials)
        sold_articles_dict = extract_text("dowloaded_files/artikel_rapport.pdf")
        #print(result_list)

        # Print or use the result_list as needed
        # for item in result_list:
        #     print(item)
        spreadsheet_service = build("sheets", "v4", credentials = credentials)
        sheets = spreadsheet_service.spreadsheets()

        result = sheets.values().get(spreadsheetId=SPREADSHEET_ID, range="Blad1!A10:A").execute()

        articles = result.get("values")
        filtered_articles : list = filter_data(articles)
        # print(result_list[0]["name"] in filtered_articles)

        spreadsheet = sheets.get(spreadsheetId=SPREADSHEET_ID).execute()
        sheet_names = [sheet['properties']['title'] for sheet in spreadsheet['sheets']]

        # ––––– Change to throw error if we don't have a sheet name –––––
        # if sheet_names:
        #     first_sheet_name = sheet_names[0]
        #     print(f"The name of the first sheet is: {first_sheet_name}")
        # else:
        #     print("No sheets found in the spreadsheet.")
        sold_articles_keys= list(sold_articles_dict.keys())
        #print(sold_articles_keys) # 9/11
        offset = 10 
        filtered_articles_length = len(filtered_articles)
        for i in range(0, filtered_articles_length):
            current_item = filtered_articles[i].lower()
            if(current_item in sold_articles_keys):
               if("fat" in current_item):
                quiantity = ceil(sold_articles_dict[current_item]["quantity"]/70)   
                sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_names[0]}!H{i+offset}", valueInputOption="USER_ENTERED", body={"values":[[f"{quiantity}"]]}).execute()
               else:   
                sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_names[0]}!H{i+offset}", valueInputOption="USER_ENTERED", body={"values":[[f"{sold_articles_dict[current_item]['quantity']}"]]}).execute()
        print("Done")
    except HttpError as error:
        print(error)

def filter_data(articles):
    filtered_data = []
    for article in articles:
        if(len(article) > 0):
            filtered_data.append(article[0])
        else:
            filtered_data.append(article)
        # if(len(article) == 2):
        #     try: 
        #         float(article[1].split(" ")[0].replace("\xa0", "").replace(",", "."))
        #     except ValueError:
        #         continue
    return filtered_data

if __name__ == "__main__":
    main()