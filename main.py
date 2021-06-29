import PyPDF2
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import pdfplumber
import pandas as pd

pdf_path='D:/2021/매채별디렉토리ICT/PDF 크롤링/수질,수리수문/DG2018E002_다사~왜관간 광역도로 건설/3.pdf'

# data_only=True로 해줘야 수식이 아닌 값으로 받아온다.
# load_wb = load_workbook("ME2015E003_141.xlsx", data_only=True)

def pdf2txt(pdf_file):
    fp = open(pdf_file, 'rb')
    total_pages = PyPDF2.PdfFileReader(fp).numPages

    page_text = {}
    for page_no in range(total_pages):
        rsrcmgr = PDFResourceManager()

        retstr = StringIO()
        codec = 'utf-8'
        laparams = LAParams()
        device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
        fp = open(pdf_file, 'rb')
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        password = None
        maxpages = 0
        caching = True
        pagenos = [page_no]

        for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,
                                      caching=caching, check_extractable=True):
            interpreter.process_page(page)

        page_text[page_no] = retstr.getvalue()

        fp.close()
        device.close()
        retstr.close()
    return page_text

def pdftoexcel(page_num):
    pdf = pdfplumber.open("D:/2021/매채별디렉토리ICT/PDF 크롤링/수질,수리수문/DG2018E002_다사~왜관간 광역도로 건설/3.pdf")
    p0 = pdf.pages[page_num]
    pd.set_option('display.max_rows', 20)
    pd.set_option('display.max_columns', 20)
    table2df = lambda table: pd.DataFrame(table[1:], columns=table[0])
    tables = table2df(p0.extract_table())

    tables.to_excel(excel_writer='D:/2021/매채별디렉토리ICT/PDF 크롤링/수질,수리수문/DG2018E002_다사~왜관간 광역도로 건설/DG2018E002_' + str(page_num) + '.xlsx')  # 엑셀로 저장
    return 0

page = 0
text = pdf2txt(pdf_path)
print("총 페이지 수 : ", range(len(text)))
for item in range(len(text)):
    test = text[item][:-1].replace(" ","").find('개발전·중·후홍수량증감표')
    if test != -1:
        print("있다")
        print("크롤링 페이지 :",item)
        pdftoexcel(item)
    else:
        print("없다")

