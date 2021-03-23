import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import re
import time
from selenium import webdriver

SCOPES = ['https://www.googleapis.com/auth/documents']
## insert your document_id
DOCUMENT_ID = ''


def get_indexnum(doc_content, match_text):
    endIdx = 0
    for d in doc_content:
        para = d.get('paragraph')
        if para is None:
            continue
        else:
            elements = para.get('elements')
            for e in elements:
                if e.get('textRun'):
                    content = e.get('textRun').get('content')
                    if match_text == str(content).rstrip():
                        endIdx = e.get('endIndex')
    return endIdx


def insert_text(idx, text):
    req = {'insertText': {
        'location': {
            'index': idx,
        },
        'text': text
    }}
    return req


def read_paragraph_element(element):
    text_run = element.get('textRun')
    if not text_run:
        return ''
    return text_run.get('content')


def read_strucutural_elements(elements):
    text = ''
    for value in elements:
        if 'paragraph' in value:
            elements = value.get('paragraph').get('elements')
            for elem in elements:
                text += read_paragraph_element(elem)
        elif 'table' in value:
            table = value.get('table')
            for row in table.get('tableRows'):
                cells = row.get('tableCells')
                for cell in cells:
                    text += read_strucutural_elements(cell.get('content'))
        elif 'tableOfContents' in value:
            toc = value.get('tableOfContents')
            text += read_strucutural_elements(toc.get('content'))
    return text


def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('docs', 'v1', credentials=creds)

    document = service.documents().get(documentId=DOCUMENT_ID).execute()
    print('The title of the document is: {}'.format(document.get('title')))
    doc_content = document.get('body').get('content')
    readElement = read_strucutural_elements(doc_content)
    englist = []
    hangul = re.compile('[ㄱ-ㅣ가-힣]+')
    english_only = re.compile('[a-zA-Z]+')
    url_not = re.compile('http+')
    str_list = list(filter(None, readElement.split('\n')))
    for i in str_list:
        if re.search(english_only, i) is not None and re.search(hangul, i) is None \
                and re.search(url_not, i) is None:
            englist.append(i)
    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument("disable-gpu")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 6.3; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.90 "
        "Safari/537.36")
    driver = webdriver.Chrome('./chromedriver', options=options)
    driver.implicitly_wait(3)
    for i in englist:
        driver.get('https://translators.to/?source=&target=ko&text=')
        search = driver.find_element_by_css_selector(
            '#root > div > main > div > section > div > div.columns.is-desktop > '
            'div:nth-child(1) > div > div > textarea')
        search.clear()
        search.send_keys(i)
        driver.find_element_by_css_selector(
            '#root > div > main > div > section > div >'
            ' div.columns.is-desktop > div:nth-child(1) > '
            'div > div > div > div > p > button').click()
        time.sleep(4)
        google = driver.find_element_by_css_selector('#googleNmt').text
        papago = driver.find_element_by_css_selector('#papago').text
        kakao = driver.find_element_by_css_selector('#kakao').text
        document = service.documents().get(documentId=DOCUMENT_ID).execute()
        doc_content = document.get('body').get('content')
        transSentance = get_indexnum(doc_content, i)
        print(insert_text(transSentance, google))
        service.documents().batchUpdate(
            documentId=DOCUMENT_ID, body={'requests': insert_text(transSentance,
                                          'google: ' + google + '\n'
                                          'papago: ' + papago + '\n'
                                          'kakao: ' + kakao + '\n')}).execute()
    print('finish')
    driver.close()

main()
