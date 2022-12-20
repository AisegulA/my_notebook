import pandas as pd
import urllib.request
import shutil

data_file = 'data.xlsx'


def download_file(filename):
    headers = dict(
        [
            (
                "User-agent",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:103.0) Gecko/20100101 Firefox/103.0",
            ),
        ]
    )
    request = urllib.request.Request(url=link, headers=headers)
    with urllib.request.urlopen(request) as response:
        with open(filename, "wb") as file_:
            shutil.copyfileobj(response, file_)


df = pd.read_excel(data_file)
for key, link in enumerate(df['PDF Url']):
    if pd.isna(link):
        continue

    link_parts = link.split('/')
    pdf_id = link_parts[len(link_parts) - 2]
    download_file('pdfs/{0}.pdf'.format(pdf_id))
