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
    def __init__(self, occup: str, addr: str, case_ref: str):
        """
        Args:
            occup (str): The occupier of the property
            addr (str): The address of the property
            case_ref (str): The case reference of the request
        """
        self.occup = occup
        self.addr = addr
        self.case_ref = case_ref


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
    # TODO: Create a CollectionChange object from the results
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
        '<style>\n' \
        'body {\n' \
        'font-family: sans-serif;\n' \
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
        '<p class="header>\n' \
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
        '</html>'
    return html


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
