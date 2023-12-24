import os, io, re
from math import ceil
import PyPDF2
import PySimpleGUI as psg

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload


"""
Krav för att skriptet ska funka:
1. Första artikeln måste börja på rad A10
2. Kvitto ligger i h columnen
3. Det är 70 glas i en keg 
4. Coors måste byta namn från Coors Light i kassaservern
5. Bright light måste ta bort "The" i kassaservern
"""

"""
 # if(len(article) == 2):
        #     try: 
        #         float(article[1].split(" ")[0].replace("\xa0", "").replace(",", "."))
        #     except ValueError:
        #         continue
"""


SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.readonly"]

"""
Check if we have the right credentials, otherwise we get them, if they are expired we update them.
"""
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


"""
Downloads a pdf from drive and stores the pdf in dowload_fiels folder 
"""
def downloadPDF(creds, pdf_id):
    try:
        drive_service = build("drive", "v3", credentials=creds)
        request = drive_service.files().get_media(fileId=pdf_id)
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

"""
Scrapes the text from the pdf and stores the disired result in a dictornary which has the form: {article: {quantity: int, price: float}}
"""
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
            result_dict[' '.join(info[:-2]).lower()] = {
                "quantity" : int(info[-2]), 
                "price": float(info[-1].replace(',', '.'))}
        return result_dict

"""
Filters and makes the data from the spreadsheet easier the use
"""
def filter_data(articles):
    filtered_data = []
    for article in articles:
        if(len(article) > 0):
            filtered_data.append(article[0])
        else:
            filtered_data.append(article)
    return filtered_data

"""
Add each article and it's quantity from the pdf to spreadsheet we define.
"""
def add_quantities_to_sheet(sold_articles_dict, sheets, sheet_names, sheet_id, column):
    try:
        result = sheets.values().get(spreadsheetId=sheet_id, range=f"{sheet_names[0]}!A10:A").execute()
        articles = result.get("values")
        filtered_articles : list = filter_data(articles)

        sold_articles_keys= list(sold_articles_dict.keys())
        offset = 10 
        filtered_articles_length = len(filtered_articles)

        for i in range(0, filtered_articles_length):
            if not isinstance(filtered_articles[i], str):
         # Skip iteration if the type is not a string
                continue
            current_item = filtered_articles[i].lower()

            if(current_item in sold_articles_keys):

                if("fat" in current_item):
                    quiantity = ceil(sold_articles_dict[current_item]["quantity"]/70)  
                    sheets.values().update(spreadsheetId=sheet_id, range=f"{sheet_names[0]}!{'F' if column else 'H'}{i+offset}", valueInputOption="USER_ENTERED", body={"values":[[f"{quiantity}"]]}).execute()
                else:   
                    sheets.values().update(spreadsheetId=sheet_id, range=f"{sheet_names[0]}!{'F' if column else 'H'}{i+offset}", valueInputOption="USER_ENTERED", body={"values":[[f"{sold_articles_dict[current_item]['quantity']}"]]}).execute()

    except HttpError as error:
        print(error)


def create_popup(message="", short=True):
    layout = [
        [psg.Text(f"{message}", background_color="black", font=("", 16), size=(36,4 if short else 8), justification="center")],
        [psg.Text("Du måste kryssa ned / trycka på OK för att fortsätta", background_color="black",text_color="red" ,font=("", 14))],
        [psg.Button("OK",button_color=("white", "black"), font=("", 16))]
    ]

    window = psg.Window("Popup Window", layout, modal=True, background_color="black", element_justification="center")

    while True:
        event, values = window.read()

        if event == psg.WIN_CLOSED or event == "OK":
            break

    window.close()

def createLayout():
    column = [
        [
            psg.Button("Hjälp", button_color=("white", "black"),font=("",16), key="-Help-"),
            psg.Button("Krav", button_color=("white", "black"), font=("", 16),key="-Requirements-"),
            psg.Button("Om", button_color=("white", "black"), font=("", 16),key="-About-")
            
        ],
        [
            psg.Text("", size=(1, 1), background_color="black")
        ],
        [
            psg.Text("Länk till Artikellrapporten i drive:",size=(25,1),background_color="black", font=("", 16)),
            psg.In(size=(95,1), enable_events=True, key="-PDF-", font=("",12))
        ],
        [
            psg.Text("Länk till Beställningslistan i drive:",size=(25,1),background_color="black", font=("",16)),
            psg.In(size=(95,1), enable_events=True, key="-Sheet-", font=("", 12))
        ],
        [
            psg.Text("", size=(1, 1), background_color="black")
        ],
        [
            psg.Button("Kör", font=("", 16), size=(15,1),button_color=("white","black"),key="-RUN-",),
            psg.Checkbox(
                "Krökens/Brandons beställningslista", 
                "Tangebels",
                enable_events=True,
                key="KBB", 
                background_color="black", 
                font=("", 16))
        ]
    ]
    return column

def getHelpMsg():
    msg = "Fyll i länken för respektive dokument, länken får du genom att högerklicka på filen i drive, tryck dela och sedan kopiera länk. När båda fält är ifyllda välj om mallen för beställningslistan är Kröken/Brandon eller inte. Kör sedan programmet och boom allt är klart!!!"
    return msg

def getRequirementsMsg():
    msg = """Krav för att skriptet ska funka:
            1. Första artikeln måste börja på rad A10
            2. Kvitto ligger i h columnen
            3. Det är 70 glas i en keg
            4. Coors måste byta namn från Coors Light i kassaservern
            5. Bright light måste ta bort "The" i kassaservern
        """
    return msg

def getAboutMsg():
    msg = """Kodad av Shamil Limbasiya DA 23/24, om ni av någon anledning vill kontakta mig skicka ett mail till: shamil0110@gmail.com eller ett SMS till: 070-44 07 954!!! :D
    """
    return msg
    

def gui(credentials):
    
    layout = [[psg.Column(createLayout(), background_color="black",justification="center", element_justification="center", expand_y=True,vertical_alignment="center")]]

    window = psg.Window(title="Artikelrapport -> Beställningslista", layout=layout,background_color="black")

    while True:
        event, values = window.read()
        print(values)
        if event == psg.WIN_CLOSED:
            break
        if event == "-RUN-":
            if(values["-PDF-"] is None or values["-PDF-"] == "" 
               or values["-Sheet-"] is None or values["-Sheet-"] == ""):
                create_popup("You must input values in both fields")
            else:
                try:
                    run_script(credentials, values["-PDF-"], values["-Sheet-"], values["KBB"])
                except ValueError as error:
                    create_popup(error)
        if(event == "-Help-"):
            create_popup(getHelpMsg(), False)
        
        if(event == "-Requirements-"):
            create_popup(getRequirementsMsg(), False)
        if(event == "-About-"):
            create_popup(getAboutMsg())


    window.close()

def run_script(credentials, pdf_url, sheet_url, column):
    try:
        pdf_id = pdf_url.split("/")[-2]
        sheet_id = sheet_url.split("/")[-2]
    except:
        raise ValueError("Something wrong with input strings")

    spreadsheet_service = build("sheets", "v4", credentials = credentials)
    sheets = spreadsheet_service.spreadsheets()

    try:
        spreadsheet = sheets.get(spreadsheetId=sheet_id).execute()
        sheet_names = [sheet['properties']['title'] for sheet in spreadsheet['sheets']]
    except:
        raise ValueError("Requested entity was not found")

    
    downloadPDF(credentials, pdf_id)
    sold_articles_dict = extract_text("dowloaded_files/artikel_rapport.pdf")
    add_quantities_to_sheet(sold_articles_dict, sheets, sheet_names, sheet_id, column)
    print("Done")


def main():
    try:
        credentials = handleCredentials()
        gui(credentials)
        
    except HttpError as error:
        print(error)


if __name__ == "__main__":
    main()