
from __future__ import print_function
import requests
import re
import html
from bs4 import BeautifulSoup
from gdoctableapppy import gdoctableapp
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import urllib.parse

HEADERS = {
    "Host": "www.zap.co.il",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.5672.93 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/sig",
    "Accept-Encoding": "gzip, deflate",
    # "Accept-Language": "he-IL, en-US;q=0.9, en;q=0.8"
}


class MakeRequest():

    def __init__(self) -> None:
        print("objected created!")

    def post_request(self, url, headers, json=None) -> requests.Response:
        if json == None:
            response = requests.post(url=url, headers=headers)
            return response
        response = requests.post(url=url, headers=headers, json=json)
        return response

    def get_request(self, url, headers=None) -> requests.Response:
        if headers == None:
            response = requests.get(url=url)
            return response
        response = requests.get(url=url, headers=headers)
        return response

    def get_model_id(self, response_content) -> int:
        pattern = r'data-modelId="(\d+)">'
        match = re.search(pattern, str(response_content))
        if match:
            number = match.group(1)
            return number
        else:
            #print("in second pattern")
            pattern = r'modelid=(\d+)'
            match = re.search(pattern, str(response_content))

        if match:
            number = match.group(1)
            return number
        print("Number not found.")

    def get_companies_and_their_prices(self, response_content) -> str:

        soup = BeautifulSoup(str(response_content), 'html.parser')
        matches = soup.find_all('div', class_='compare-item-row')
        name_matches, price_matches = [], []
        for match in matches:
            if match.get('data-sale-type') != None and int(match.get('data-sale-type')) == 3:  # EILAT
                continue
            price = match.get('data-total-price')
            # name=reverse_hebrew_substrings(match.get('data-site-name'))
            name = match.get('data-site-name')
            if name == None:
                continue
            name_matches.append(name)
            price_matches.append(price)
            #print(f"Total Price: {price}, Site Name: {name}")

        return price_matches, name_matches
    

    def get_stores_name_and_id(self, response_content) -> dict:
        soup = BeautifulSoup(str(response_content), 'html.parser')
        matches = soup.find_all('div', class_='compare-item-row')
        names_and_id={}
        for match in matches:
            if match.get('data-sale-type') != None and int(match.get('data-sale-type')) == 3:  # EILAT
                continue
            store_id = match.get('data-site-id')
            name = match.get('data-site-name')
            if name == None:
                continue
            names_and_id[name]=store_id
        return names_and_id
    

    def get_store_locations(self,response_content):
        soup = BeautifulSoup(str(response_content), 'html.parser')
        span_element = soup.find('span', itemprop='streetAddress')
        street_address = span_element.get_text(strip=True)
        if street_address=='לרשימת הסניפים':
            input_element = soup.find('input', {'name': 'Coords'})
            if input_element:
                value = input_element['value']
                return extract_address_with_city(value)
        return street_address


def extract_address_with_city(value):
    decoded_value = urllib.parse.unquote(value)
    addresses = []

    coordinates = decoded_value.split('|')
    for coordinate in coordinates:
        parts = coordinate.split(',')
        if len(parts) >= 3:
            address = ','.join(parts[2:]).strip()
            addresses.append(address)

    return addresses



def reverse_hebrew_substrings(text):
    pattern = r'[\u0590-\u05FF]+'
    hebrew_substrings = re.findall(pattern, text)
    reversed_text = text
    for substring in hebrew_substrings:
        reversed_substring = substring[::-1]
        reversed_text = reversed_text.replace(substring, reversed_substring)
    return reversed_text


def create_document(table_data, creds):

    docs_service = build('drive', 'v3', credentials=creds)
    file_metadata = {
        'name': 'My Document',
        'parents': ['1lBmUcDPDAiMqyHt3lBIZSvpOkPN5L8Hc'],
        'mimeType': 'application/vnd.google-apps.document'
    }
    document = docs_service.files().create(body=file_metadata).execute()
    document_id = document['id']

    # Build the Google Docs API client
    docs_service = build('docs', 'v1', credentials=creds)

    # Create the table
    requests = [{
        'insertTable': {
            'rows': 1,
            'columns': 2,
            'endOfSegmentLocation': {}
        }
    }]

    try:
        result = docs_service.documents().batchUpdate(
            documentId=document_id, body={'requests': requests}).execute()
    except HttpError as e:
        print(e)

    resource = {
        "oauth2": creds,
        "documentId": document_id,
        "tableIndex": 0,
        "values": table_data
    }

    res = gdoctableapp.SetValues(resource)
    drive_service = build('drive', 'v3', credentials=creds)

    # Set up permission metadata
    permission_metadata = {
        'role': 'reader',
        'type': 'anyone',
    }

    # Create the permission
    permission = drive_service.permissions().create(
        fileId=document_id,
        body=permission_metadata,
        fields='id'
    ).execute()

    # Get the permission ID
    permission_id = permission['id']

    # Generate the link
    link = f'https://drive.google.com/file/d/{document_id}/view?usp=sharing'
    print("Document Link: "+link)


def main():
    product_name = "מקרר ‏מקפיא עליון Sharp SJSE75DBK/SL ‏590 ‏ליטר שארפ"
    URL = "https://www.zap.co.il/search.aspx?keyword="
    URL = URL + product_name
    req_obj = MakeRequest()
    response = req_obj.get_request(URL, headers=HEADERS)
    # with open("test.txt", "w") as file1:
    #     file1.write(str(respone.content.decode('utf-8')))
    model_number = req_obj.get_model_id(response.content)
    if model_number == None:
        print("This item does not exist! Please be more specific.")
        exit()
    comparison_url = "https://www.zap.co.il/model.aspx?modelid=" + model_number
    respone = req_obj.get_request(comparison_url, HEADERS)
    # with open("test.txt", "w") as file1:
    #     file1.write(str(respone.content.decode('utf-8')))
    # pass
    print("Product URL: " + comparison_url)

    price_matches, name_matches = req_obj.get_companies_and_their_prices(respone.content.decode('utf-8'))

    table_data = []
    table_data.append([":Prices", ":Store Names"])  # table headers
    for price, name in zip(price_matches, name_matches):
        table_data.append([price, name])

    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/drive.file',
              'https://www.googleapis.com/auth/drive']
    creds = None
    # The file token.json stores the user's access and refresh tokens, and iscl
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    create_document(table_data, creds)


if __name__ == "__main__":
    main()
