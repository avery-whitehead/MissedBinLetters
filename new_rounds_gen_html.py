"""
gen_html.py
How to run:
Install pyodbc version 4.0.22 or greater
From the command line, run `py -3 new_rounds_gen_html.py`
How it works:
Generates an HTML file and uses wkhtmltopdf to convert it to a PDF
Contains changes to a round
"""
import glob
import subprocess
import datetime
import json
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
        'font-family: "Arial", sans-serif;\n' \
        '}\n' \
        'h1 {\n' \
        'font-weight: normal;\n' \
        'text-align: center;\n' \
        'font-size: 36pt;\n' \
        'text-decoration: underline;\n' \
        '}\n' \
        ' p {\n' \
        'font-size: 11pt;\n' \
        '}' \
        '.addr {\n' \
        'font-weight: bold;\n' \
        'padding-top: 90px;\n' \
        'font-size: 18px;\n' \
        '}\n' \
        '.content {\n' \
        'font-family: "Calibri"\n' \
        '}\n' \
        '.header {\n' \
        'text-decoration: underline;\n' \
        '}\n' \
        '.footer {\n' \
        'font-size: 9pt;\n' \
        'color: #808080;\n' \
        '}\n' \
        'img {\n' \
        'max-width: 100%;\n' \
        'max-height: 75px;\n' \
        '}\n' \
        '.hdc-td-greenborder {\n' \
        'font-size: 12px;\n' \
        '}\n' \
        '#collections-table {\n' \
        'color: #444444;\n' \
        'margin: 0 auto;\n' \
        '}\n' \
        '.table {\n' \
        'width: 450px;\n' \
        'text-align: center;\n' \
        'border-collapse: collapse;\n' \
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
        '<p>\n' \
        f'{date}\n' \
        '</p>\n' \
        '<p>\n' \
        'Dear Sir/Madam\n' \
        '</p>\n' \
        '<p class="header">\n' \
        '<strong>Waste and Recycling Collections</strong>\n' \
        '</p>\n' \
        '<p>\n' \
        'Earlier this month we implemented changes to our collections for ' \
        'waste and recycling which meant changes for your home. ' \
        'Unfortunately we have identified further amendments needed to ' \
        'ensure the most efficient service delivery for our residents. ' \
        'This only affects a small number of properties – but it includes ' \
        'your home.\n' \
        '<br>\n' \
        'Below are the details of the new fortnightly collection ' \
        'arrangements for your property, which come into effect from ' \
        'Monday June 18:\n' \
        '</p>\n' \
       f'{get_html_table(change.uprn)}\n' \
        '<p>\n' \
        '<br>\n' \
        'Please put your containers at your collection point by 6am.\n' \
        '</p>\n' \
        '<p>\n' \
        'I would like to apologise for any inconvenience caused as a ' \
        'result of these further changes.\n' \
        '</p>\n' \
        '<br>\n' \
        '<p>\n' \
        'Gary Brown\n' \
        '<br>\n' \
        '<span>\n' \
        'Waste & Street Scene Manager\n' \
        '</span>\n' \
        '<br>\n' \
        'WasteandStreetScene@hambleton.gov.uk\n' \
        '</p>\n' \
        '<p class="footer">\n' \
        '<span style="color: #9EB4D0">\n' \
        'Hambleton District Council\n' \
        '</span>\n' \
        '<br>\n' \
        'Waste and Street Scene, Bridge End House\n' \
        '<br>\n' \
        'Darlington Road, Northallerton, North Yorkshire DL6 2PL\n' \
        '<br>\n' \
        '<span style="font-size: 7pt">\n' \
        'Some of our calls are recorded. For further information visit our ' \
        'website www.hambleton.gov.uk to view the Call Recording Policy\n' \
        '</span>\n' \
        '</p>\n' \
        '</body>\n' \
        '</html>'
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
    htmls = glob.glob('.\\htmls\\changes\\*')
    count = 0
    for html in htmls:
        out_f = html[16:-5]
        pdf = f'.\\pdfs\\changes\\{out_f}.pdf'
        flags = '--proxy 127.0.0.1:3128 ' \
        '--disable-smart-shrinking ' \
        '-B 25.4mm -L 25.4mm -R 25.4mm -T 25.4mm'
        args = f'{exe} {flags} "{html}" "{pdf}"'
        print(args)
        subprocess.call(args, shell=False)
        print(f'Converted {pdf}')
        count += 1
    return f'Saved {count}/{len(htmls)} PDFs'


if __name__ == '__main__':
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
