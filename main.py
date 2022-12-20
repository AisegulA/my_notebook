import argparse
import time

import pandas as pd
from PyPDF2 import PdfReader
from textsearch import TextSearch
from tqdm import tqdm
from slugify import slugify

parser = argparse.ArgumentParser()
parser.add_argument('-l', '--lang', default='en')
args = parser.parse_args()

term_raw = input("Aranacak terimler (';' ile ayırabilirsiniz): ")
terms = [x.strip() for x in term_raw.split(';')]
if not len(terms) or (len(terms) > 0 and terms[0] == ''):
    print('En az bir terim girilmelidir.')
    exit(1)


def find_pdfs(lang='en'):
    pdf_box = []

    df = pd.read_excel('data.xlsx')
    for key, pdf_lang in enumerate(df['PDF Lang']):
        if pd.isna(pdf_lang) or pdf_lang != lang:
            continue

        pdf_url = df['PDF Url'][key]
        link_parts = pdf_url.split('/')
        pdf_id = link_parts[len(link_parts) - 2]
        pdf_box.append(pdf_id)

    return pdf_box


def pdf_fulltext(pdf_file):
    full_text = ''

    try:
        reader = PdfReader(pdf_file)
        for page in reader.pages:
            full_text += page.extractText()

        return full_text
    except (Exception,):
        return full_text


def search_pdf(text):
    search_results = []
    for term in terms:
        ts = TextSearch(case='insensitive', returns='match')
        ts.add(term)
        found = ts.findall(text)
        search_results.append(len(found))

    return search_results


results = []
pdf_lists = find_pdfs(args.lang)

for pdf_id in tqdm(pdf_lists):
    path = 'pdfs/{0}.pdf'.format(pdf_id)
    fulltext = pdf_fulltext(path)
    result = search_pdf(fulltext)

    results.append(result)

df_results = pd.DataFrame(results, index=pdf_lists, columns=terms)

total_results = []
for k, t in enumerate(terms):
    total = 0
    for counts in results:
        total = total + counts[k]

    total_results.append([total])

df_total = pd.DataFrame(total_results, index=terms, columns=['Toplam Eşleşme'])

filename = time.strftime('%Y%m%d-%H%M%S')
term_slug = slugify(term_raw)
with pd.ExcelWriter('pdf_results/{0}_{1}_{2}.xlsx'.format(filename, args.lang, term_slug)) as writer:
    df_results.to_excel(writer, sheet_name='Results')
    df_total.to_excel(writer, sheet_name='Total Results')
