"""
gen_html.py
How to run:
Install pyodbc version 4.0.22 or greater
From the command line, run `py -3 new_rounds_gen_html.py`
How it works:
Generates an HTML file and uses wkhtmltopdf to convert it to a PDF
Contains changes to a round
"""
import datetime
import sys
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
        self.uprn = uprn


def query_changes() -> list:
    """
    Queries the SQL database for addresses of properties that are
    having a change in the collections
    """
    changes = []
    with open('.\\changes_info.sql', 'r') as changes_query_f:
        changes_query = changes_query_f.read()
    print(changes_query)
    cursor = CONN.cursor()
    cursor.execute(changes_query)
    results = cursor.fetchall()
    for result in results:
        changes.append(CollectionChange(
            result.occup,
            result.addr,
            result.uprn))
    return changes

def log_error(log_path: str, error: Exception):
    """
    Writes exception messages to the log file and exits the program
    Args:
        log_path (str): The path of the log file to write to
        error (Exception): The error message given by an exception
    """
    with open(log_path, 'a') as log:
        log.write(f'{SYSTIME} - {error}\n')
    sys.exit(1)

def create_html(change: CollectionChange) -> str:
    """
    Creates the HTML using the change information
    Args:
        change (CollectionChange): The changes to fill the template with
    Returns:
        (str): The HTML template with the information filed in
    """
    date = datetime.datetime.now().strftime('%d %B %Y')
    html = '' \
        '<!DOCTYPE html>\n' \
        '<html>\n' \
        '<head>\n' \
        '<link rel="stylesheet" href="./bootstrap.min.css">\n' \
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
        'font-size: 13pt;\n' \
        '}' \
        '.addr {\n' \
        'font-weight: bold;\n' \
        'padding-top: 70px;\n' \
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
        '.collections-table {\n' \
        'color: #444444;\n' \
        'line-height: 14px;\n' \
        'margin: 0 auto;\n' \
        '}\n' \
        '.table {\n' \
        'width: 66%;\n' \
        'text-align: center;\n' \
        'border-collapse: collapse;\n' \
        '}\n' \
        '.table-striped tbody tr:nth-of-type(odd) {\n' \
        'background-color: #FFFFFF;\n' \
        '}\n' \
        '</style>\n' \
        '</head>\n' \
        '<body>\n' \
        '<section>\n' \
        '<div class="addr">\n' \
       f'{change.occup}<br>\n' \
       f'{change.addr}\n' \
        '</div>\n' \
        '<br>\n' \
        '<br>\n' \
        '<br>\n' \
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
        '</body>\n' \
        '<p>\n' \
        'TABLE\n' \
        '</p>\n' \
        '<p>\n' \
        'Please put your containers at your collection point by 6am.\n' \
        '</p>\n' \
        '<p>\n' \
        'I would like to apologise for any inconvenience caused as a ' \
        'result of these further changes.\n' \
        '</p>\n' \
        '<br>\n' \
        '<br>\n' \
        '<br>\n' \
        '<br>\n' \
        '<br>\n' \
        '<br>\n' \
        '<p>\n' \
        'Gary Brown\n' \
        '<br>\n' \
        '<span class="header">\n' \
        '<strong>Waste & Street Scene Manager</strong>\n' \
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
        '</p>\n' \
        '<p class="footer" style="font-size: 8pt">\n' \
        'Some of our calls are recorded. For further information visit our ' \
        'website www.hambleton.gov.uk to view the Call Recording Policy\n' \
        '</p>\n' \
        '</body>\n' \
        '</html>'
    return html

def get_table_html():
    pass

if __name__ == '__main__':
    SYSTIME = datetime.datetime.now().strftime('%d-%b-%Y %H:%M:%S')
    print(type(SYSTIME))
    print(SYSTIME)
    with open ('.\\.config_chngs', 'r') as config_f:
        config = json.load(config_f)
    try:
        CONN = pyodbc.connect(
            driver=config['driver'],
            server=config['server'],
            database=config['database'],
            uid=config['uid'],
            pwd=config['pwd'])
    except (pyodbc.DatabaseError, pyodbc.InterfaceError) as error:
        log_error('.\\missed_bin_letters.log', error)
    changes = query_changes()
    for change in changes:
        print(change.uprn)
