import io
import os

import PyPDF2
import pytesseract
from pdf2image import convert_from_path
import pycountry
import pandas as pd

data_file = 'data.xlsx'

poppler_path = None
if os.name == 'nt':
    poppler_path = 'binaries/etc/poppler-22.11.0/Library/bin'
    pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract.exe'


def find_lang(lang_code):
    lang = pycountry.languages.get(alpha_2=lang_code)
    if not lang:
        return None

    return lang.alpha_3


def ocr_pdf(pdf_file, pdf_lang):
    images = convert_from_path(pdf_file, poppler_path=poppler_path)
    pdf_writer = PyPDF2.PdfFileWriter()
    for image in images:
        page = pytesseract.image_to_pdf_or_hocr(image, extension='pdf', lang=pdf_lang)
        pdf = PyPDF2.PdfFileReader(io.BytesIO(page))
        pdf_writer.addPage(pdf.getPage(0))

    pdf_filename = os.path.basename(pdf_file)
    with open('ocr_pdf/{0}'.format(pdf_filename), 'wb') as f:
        pdf_writer.write(f)


def is_need_ocr(pdf_file):
    is_needle = False
    reader = PyPDF2.PdfReader(pdf_file)
    for page in reader.pages:
        if len(page.extractText()) <= 0:
            is_needle = True
            break
    return is_needle


print('OCR işlemi başlatıldı.')

df = pd.read_excel(data_file)
for key, link in enumerate(df['PDF Url']):
    if pd.isna(link):
        continue

    link_parts = link.split('/')
    pdf_id = link_parts[len(link_parts) - 2]
    pdf_lcode2 = df['PDF Lang'][key]
    pdf_lcode3 = find_lang(pdf_lcode2)
    pdf_path = 'pdfs/{0}.pdf'.format(pdf_id)
    pdf_ocr_path = 'ocr_pdf/{0}.pdf'.format(pdf_id)

    try:
        if not is_need_ocr(pdf_path):
            print('{0}.pdf dosyası için OCR işlemi gerekmiyor.'.format(pdf_id))
            continue

        if not pdf_lcode3:
            print("{0}.pdf dosyası için '{1}' dil kodu bulunamadı.".format(pdf_id, pdf_lcode2))
            continue

        if os.path.exists(pdf_ocr_path):
            print('{0}.pdf dosyası zaten tarandı. Atlanıyor.'.format(pdf_id, pdf_lcode2))
            continue

        ocr_pdf(pdf_path, pdf_lcode3)
        print('{0}.pdf dosyası kaydedildi.'.format(pdf_id))
    except (Exception,):
        print('{0}.pdf dosyası OCR taramasından geçirilirken bir hata oluştu.'.format(pdf_id))

print('OCR işlemi tamamlandı.')
