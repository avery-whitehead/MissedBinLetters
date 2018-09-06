"""
new_rounds_gen_html.py
How to run:
Install pyodbc version 4.0.22 or greater
From the command line, run `py -3 new_rounds_gen_html.py`
How it works:
Generates an HTML file and uses wkhtmltopdf to convert it to a PDF
Contains changes to a round
"""
import os
import glob
import subprocess
import datetime
import json
from PyPDF2 import PdfFileMerger, PdfFileReader
import pyodbc

class CollectionChange():
    """
    Represents a generic change in collection arrangements
    """
    def __init__(self, occup: str, addr: str, uprn: str):
        """
        Args:
            occup (str): The occupier of the property
            addr (str): The address of the property
            uprn (str): The UPRN of the property
        """
        self.occup = occup
        self.addr = addr
        # Used for display on the letter
        self.addr_str = addr.replace(', ', '<br>')
        self.uprn = uprn


def query_changes() -> list:
    """
    Queries the SQL database for addresses of properties that are
    having a change in the collections
    """
    changes = []
    with open('.\\changes_info.sql', 'r') as changes_query_f:
        changes_query = changes_query_f.read()
    cursor = CONN.cursor()
    cursor.execute(changes_query)
    results = cursor.fetchall()
    for result in results:
        changes.append(CollectionChange(
            result.occup,
            result.addr,
            result.uprn))
    return changes

def create_html(change: CollectionChange) -> str:
    """
    Creates the HTML using the change information
    Args:
        change (CollectionChange): The changes to fill the template with
    Returns:
        (str): The HTML template with the information filed in
    """
    date = datetime.datetime.now().strftime('%A %d %B %Y')
    html = '' \
        '<!DOCTYPE html>\n' \
        '<html>\n' \
        '<head>\n' \
        '<style>\n' \
        'body {\n' \
        'font-family: "Calibri", sans-serif;\n' \
        '}\n' \
        'h1 {\n' \
        'font-weight: normal;\n' \
        'text-align: center;\n' \
        'font-size: 36pt;\n' \
        'text-decoration: underline;\n' \
        '}\n' \
        'p {\n' \
        'font-size: 12pt;\n' \
        '}\n' \
        '.addr {\n' \
        'position: fixed;\n' \
        'top: 5cm;\n' \
        'height: 4cm;\n' \
        'left: 1.8cm;\n' \
        'width: 9cm;\n' \
        'font-weight: bold;\n' \
        'font-size: 12pt;\n' \
        '}\n' \
        '.content {\n' \
        'position: fixed;\n' \
        'top: 11cm;\n' \
        'left: 1.8cm;\n' \
        'width: 17cm;\n' \
        'height: 13cm;\n' \
        'font-family: "Calibri";\n' \
        'font-size: 12pt;\n' \
        '}\n' \
        '.header {\n' \
        'text-decoration: underline;\n' \
        '}\n' \
        '.signature {\n' \
        'position: fixed;\n' \
        'left: 1.8cm;\n' \
        'top: 22cm;\n' \
        '}\n' \
        '.footer {\n' \
        'position: fixed;\n' \
        'left: 1.8cm;\n' \
        'font-size: 9pt;\n' \
        'top: 27.5cm;\n' \
        'color: #808080;\n' \
        '}\n' \
        '.hdc-td-greenborder {\n' \
        'font-size: 12pt;\n' \
        '}\n' \
        '#collections-table {\n' \
        'color: #444444;\n' \
        'margin: 0 auto;\n' \
        '}\n' \
        '.table {\n' \
        'width: 450px;\n' \
        '}\n' \
        '.table-striped tbody tr:nth-of-type(odd) {\n' \
        'background-color: #FFFFFF;\n' \
        '}\n' \
        'tr {\n' \
        'border-bottom: 1px solid #C8C8C8;\n' \
        '}\n' \
        '</style>\n' \
        '</head>\n' \
        '<body>\n' \
        '<section>\n' \
        '<div class="addr">\n' \
       f'{change.occup}<br>\n' \
       f'{change.addr_str}\n' \
        '</div>\n' \
        '<br>\n' \
        '<div class="content">\n' \
        f'{date}\n' \
        '<p>\n' \
        'Dear Sir/Madam\n' \
        '</p>\n' \
        '<p class="header">\n' \
        '<strong>Waste and Recycling Collections</strong>\n' \
        '</p>\n' \
        '<p>\n' \
        'At the start of June 2018 we implemented changes to our ' \
        'collections for waste and recycling which meant changes for your ' \
        'home. Unfortunately we have identified further amendments needed ' \
        'to ensure the most efficient service deliver for our residents. ' \
        'This only affects a small number of properties - but it includes ' \
        'your home.\n' \
        '</p>\n' \
        '<p>\n' \
        'Below are the details of the new fortnightly collection ' \
        'arrangements for your property, which come into effect as ' \
        'detailed below:\n' \
        '</p>\n' \
       f'{get_html_table(change.uprn)}\n' \
        '<br>\n' \
        'Please put your containers at your collection point by 6am. I ' \
        'would like to apologise for any inconvenience caused as a result ' \
        'of these further changes.\n' \
        '</div>\n' \
        '<div class="signature">\n' \
        'Gary Brown\n' \
        '<br>\n' \
        '<span>\n' \
        'Waste & Street Scene Manager\n' \
        '</span>\n' \
        '<br> WasteandStreetScene@hambleton.gov.uk\n' \
        '</div>\n' \
        '<div class="footer">\n' \
        '<span style="color: #9EB4D0">\n' \
        'Hambleton District Council\n' \
        '</span>\n' \
        '<br> Waste and Street Scene, Bridge End House\n' \
        '<br> Darlington Road, Northallerton, North Yorkshire DL6 2PL\n' \
        '<br>\n' \
        '<span style="font-size: 7pt">\n' \
        'Some of our calls are recorded. For further information visit ' \
        'our website www.hambleton.gov.uk to view the Call Recording ' \
        'Policy\n' \
        '</span>\n' \
        '</div style="page-break-after: always;">\n' \
        '</body>\n' \
        '</html>\n'
    return html

def get_html_table(uprn: str) -> str:
    """
    Queries the SQL database to return the HTML table containing the
    details of a collection change
    Args:
        uprn (str): The UPRN of the property to get the changes for
    Returns:
        (str): The HTML table with the change details
    """
    with open('.\\changes_html_table.sql', 'r') as html_query_f:
        html_query = html_query_f.read()
    cursor = CONN.cursor()
    cursor.execute(""" SET NOCOUNT ON; """ + html_query, uprn)
    return cursor.fetchone().html

def save_html(html: str, change: CollectionChange) -> str:
    """
    Writes the HTML to file to be converted later
    Args:
        html (str): The HTML to write to file and later convert
        change (CollectionChange): The CollectionChange this HTML was
        generated from
    Returns:
        A string denoting success
    """
    dir_path = f'.\\htmls\\changes'
    file_path = f'{dir_path}\\{change.uprn}-{change.addr}.html'
    with open(file_path, 'w+') as html_f:
        html_f.write(html)
    return f'Saved {file_path}'

def convert_html() -> str:
    """
    Converts each HTML file to a PDF using wkhtmltopdf
    Returns:
        A string denoting success
    """
    exe = '"C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"'
    htmls = glob.glob('.\\htmls\\changes\\*.html')
    count = 0
    for html in htmls:
        out_f = html[16:-5]
        pdf = f'.\\pdfs\\changes\\{out_f}.pdf'
        flags = '--disable-smart-shrinking ' \
        '-B 0mm -L 0mm -R 0mm -T 0mm'
        args = f'{exe} {flags} "{html}" "{pdf}"'
        print(args)
        subprocess.call(args, shell=False)
        print(f'Converted {pdf}')
        count += 1
    return f'Converted {count}/{len(htmls)} HTMLs to PDFs'

def merge_pdfs(sys_date: str) -> str:
    """
    Merges each output PDF page into a single document
    Args:
        sys_date (str): A string of the date in YYYYmmddHHMM format, used as
        the file name for the output file
    Returns:
        A string denoting success
    """
    pdfs = glob.glob('.\\pdfs\\changes\\*.pdf')
    merger = PdfFileMerger()
    count = 0
    for pdf in pdfs:
        merger.append(PdfFileReader(pdf), 'rb')
        count += 1
    with open(f'.\\pdfs\\changes\\out\\{sys_date}.pdf', 'wb') as pdf_out:
        merger.write(pdf_out)
    merger.close()
    return f'Merged {count}/{len(pdfs)} PDFs'

def clean_files() -> str:
    """
    Cleans up any leftover files, leaving only system files and the final PDF
    output
    Returns:
        A string denoting success
    """
    count = 0
    htmls_to_remove = glob.glob('.\\htmls\\changes\\*.html')
    pdfs_to_remove = glob.glob('.\\pdfs\\changes\\*.pdf')
    to_remove = htmls_to_remove + pdfs_to_remove
    for remove_f in to_remove:
        os.remove(remove_f)
        count += 1
    return f'Cleaned up {count} files'

if __name__ == '__main__':
    sys_date = datetime.datetime.today().strftime('%Y%m%d%H%M')
    with open ('.\\.config_chngs', 'r') as config_f:
        config = json.load(config_f)
    CONN = pyodbc.connect(
        driver=config['driver'],
        server=config['server'],
        database=config['database'],
        uid=config['uid'],
        pwd=config['pwd'])
    changes = query_changes()
    for change in changes:
        html = create_html(change)
        print(save_html(html, change))
    print(convert_html())
    print(merge_pdfs(sys_date))
    print(clean_files())
    print(f'Done! Output at /pdfs/changes/{sys_date}.pdf')
