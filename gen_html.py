"""
gen_html.py
How to run:
Install pyodbc version 4.0.22 or greater
From the command line, run `py -3 gen_html.py`
How it works:
Generates an HTML file and uses wkhtmltopdf to convert it to a PDF
"""
import glob
import subprocess
import datetime
import sys
import json
import pyodbc

class Request():
    """
    Parent class for GardenWasteRequest and RecyclingRequest
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


class GardenWasteRequest(Request):
    """
    Represents a request for more garden waste sacks
    """
    def __init__(self, occup: str, addr: str, case_ref: str, num_subs: str):
        """
        Initialises the different attributes between GardenWasteRequest and
        Request
        Args:
            num_subs (str): The number of subscriptions the property has
        """
        super().__init__(occup, addr, case_ref)
        self.num_subs = num_subs
        self.req_type = 'gw'


class RecyclingRequest(Request):
    """
    Represents a request for more recycling sacks
    """
    def __init__(self, occup: str, addr: str, case_ref: str):
        """
        Initialises the different attributes between RecyclingRequest and
        Request
        """
        super().__init__(occup, addr, case_ref)
        self.req_type = 'rec'

def query_gw_requests() -> list:
    """
    Queries the SQL database for the addresses of people who requested
    garden waste sacks
    Args:
        conn (pyodbc.Connection): The connection to the database
    Returns:
        A list of GardenWasteRequest objects containing the information
        from the query
    """
    gw_requests = []
    with open('.\\gw_address_info.sql', 'r') as gw_query_f:
        gw_query = gw_query_f.read()
    cursor = CONN.cursor()
    cursor.execute(gw_query)
    results = cursor.fetchall()
    for result in results:
        gw_requests.append(GardenWasteRequest(
            result.occupier,
            result.address,
            result.case_ref,
            result.num_subs))
    return gw_requests

def query_rec_requests() -> list:
    """
    Queries the SQL database for the addresses of people who requested
    recycling sacks
    Args:
        conn (pyodbc.Connection): The connection to the database
    Returns:
        A list of RecyclingRequest objects containing the information from
        the query
    """
    rec_requests = []
    with open('.\\rec_address_info.sql', 'r') as rec_query_f:
        rec_query = rec_query_f.read()
    cursor = CONN.cursor()
    cursor.execute(rec_query)
    results = cursor.fetchall()
    for result in results:
        rec_requests.append(RecyclingRequest(
            result.occupier,
            result.address,
            result.case_ref))
    return rec_requests

def create_html(request: Request) -> str:
    """
    Creates the HTML using the request information
    Args:
        request (Request): The GardenWasteRequest or RecyclingRequest
        to fill the template with
    Returns:
        (str): The HTLM template with the information filed in
    """
    if isinstance(request, GardenWasteRequest):
        content = '' \
        '<h1>\n' \
        'Garden Waste Letter\n' \
        '</h1>\n' \
        '<p>\n' \
        'Further to your report of a missed garden waste bin collection. As ' \
        'you are most probably aware as of 4 June refuse, recycling and ' \
        'garden waste collections have been revised to improve the ' \
        'efficiency of the service.  These changes involve new collection ' \
        'routes with different drivers and have resulted in occasional ' \
        'missed collections.\n' \
        '</p>\n' \
        '<p>\n' \
        'Unfortunately we are unable to return to collect your bin(s), ' \
        'however please find enclosed 3 sacks for your garden waste. ' \
        'Please affix 1 of the enclosed stickers to each sack so that the ' \
        'licence can be seen by the collection crews.  Please ensure sacks ' \
        'are tied intact and not overflowing.\n' \
        '</p>\n' \
        '<p>\n' \
        'Please present filled sacks alongside your green wheeled bin by ' \
        '06:00 on your collection day. Please ensure that your green waste ' \
        'bin(s) correctly display current garden waste licences using the ' \
        'ties provided.\n' \
        '</p>\n' \
        '<p>\n' \
        'Thank you for your patience during this time.\n' \
        '</p>\n' \
        '<p>\n' \
        '<strong>Hambleton District Council</strong>\n' \
        '<br>\n' \
        '01609 779977\n' \
        '<br>\n' \
        'info@hambleton.gov.uk\n' \
        '</p>\n'
    elif isinstance(request, RecyclingRequest):
        content = '' \
        '<h1>\n' \
        'Recycling letter\n' \
        '</h1>\n' \
        '<p>\n' \
        'Further to your report of a missed garden waste bin collection. As ' \
        'you are most probably aware as of 4 June refuse, recycling and ' \
        'garden waste collections have been revised to improve the ' \
        'efficiency of the service.  These changes involve new collection ' \
        'routes with different drivers and have resulted in occasional ' \
        'missed collections.\n' \
        '</p>\n' \
        '<p>\n' \
        'Please ensure sacks are tied, intact and not overflowing.\n' \
        '</p>\n' \
        '<p>\n' \
        'Please present filled sacks alongside your recycling bin by 06:00 ' \
        'on your collection day.\n' \
        '</p>\n' \
        '<p>\n' \
        'Please note this is the only type of plastic bag which can be used ' \
        'to contain your recycling, all other forms of plastic bags such as ' \
        'shopping bags or black plastic bin bags are not recyclable on our ' \
        'kerbside collection scheme and must be kept out of your recycling ' \
        'bin.\n' \
        '</p>\n' \
        '<p>\n' \
        'Please refer to your collection day postcard or visit ' \
        'hambleton.go.uk to check collect day details and for other ' \
        'information.\n' \
        '</p>\n' \
        '<p>\n' \
        'Thank you for your patience during this time.\n' \
        '</p>\n' \
        '<p>\n' \
        '<strong>Hambleton District Council</strong>\n' \
        '<br>\n' \
        '01609 779977\n' \
        '<br>\n' \
        'info@hambleton.gov.uk\n' \
        '</p>\n'
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
        'font-size: 11pt;\n' \
        '}' \
        '.addr {\n' \
        'font-weight: bold;\n' \
        'padding-top: 70px;\n' \
        'font-size: 18px;\n' \
        'page-break-after: always;\n' \
        '}\n' \
        '.content {\n' \
        'font-family: "Calibri"\n' \
        '}\n' \
        '</style>\n' \
        '</head>\n' \
        '<body>\n' \
        '<section>\n' \
        '<div class="addr">\n' \
       f'{request.addr}\n' \
        '</div>\n' \
        '<div class="content">\n' \
       f'{content}\n' \
        '</div>\n' \
        '</section>\n' \
        '</body>\n' \
        '</html>'

    return html

def save_html(html: str, request: Request) -> None:
    """
    Writes the HTML to file, creating additional files for properties with
    more than one garden waste subscription
    Args:
        html (str): The HTML to write to file and later convert
        request (Request): The Request this HTML was generated from
    """
    try:
        dir_path = f'.\\htmls\\{request.req_type}'
        html_path = f'{dir_path}\\{request.case_ref}-1.html'
        with open(html_path, 'w+') as html_f:
            html_f.write(html)
        success_str = f'{SYSTIME} - Successfully saved {html_path}'
        print(success_str)
        with open('.\\missed_bin_letters.log', 'a') as log:
            log.write(f'{success_str}\n')
        # If there is more than one license, create multiple letters
        if isinstance(request, GardenWasteRequest):
            for i in range(1, int(request.num_subs)):
                html_path = f'{dir_path}\\{request.case_ref}-{i + 1}.html'
                with open(html_path, 'w+') as html_f:
                    html_f.write(html)
            success_str = f'{SYSTIME} - Successfully saved {html_path}'
            print(success_str)
            with open('.\\missed_bin_letters.log', 'a') as log:
                log.write(f'{success_str}\n')
    except (IOError, FileNotFoundError) as error:
        log_error('.\\missed_bin_letters.log', error)

def convert_html() -> None:
    """
    Converts each HTML file to a PDF using wkhtmltopdf
    """
    try:
        exe = '"C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"'
        html_dirs = ['.\\htmls\\gw\\*', '.\\htmls\\rec\\*']
        for html_dir in html_dirs:
            htmls = glob.glob(html_dir)
            for html in htmls:
                if html_dir == '.\\htmls\\gw\\*':
                    case_ref = html[11:-5]
                    pdf = f'.\\pdfs\\gw\\{case_ref}.pdf'
                elif html_dir == '.\\htmls\\rec\\*':
                    case_ref = html[12:-5]
                    pdf = f'.\\pdfs\\rec\\{case_ref}.pdf'
                flags = '-B 25.4mm -L 25.4mm -R 25.4mm -T 25.4mm'
                args = f'{exe} --disable-smart-shrinking {flags} {html} {pdf}'
                print(args)
                subprocess.call(args, shell=False)
    except (IOError, FileNotFoundError) as error:
        log_error('.\\missed_bin_letters.log', error)

def update_database(request: Request) -> str:
    """
    Runs the two UPDATE queries to keep track of which requests have been
    processed
    Args:
        request (Request): The request in the database to update
    Returns:
        (str): A string indicating success
    """
    try:
        with open('.\\gw_update.sql', 'r') as gw_update_f:
            gw_update = gw_update_f.read()
        with open('.\\rec_update.sql', 'r') as rec_update_f:
            rec_update = rec_update_f.read()
        cursor = CONN.cursor()
        cursor.execute(gw_update)
        cursor.execute(rec_update)
        CONN.commit()
        success_str = f'{SYSTIME} - Updated database for {request.case_ref}'
        with open('.\\missed_bin_letters.log', 'a') as log:
            log.write(f'{success_str}\n')
        return success_str
    except (pyodbc.DatabaseError, pyodbc.InterfaceError) as error:
        log_error('.\\missed_bin_letters.log', error)


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


if __name__ == '__main__':
    SYSTIME = datetime.datetime.now().strftime('%d-%b-%Y %H:%M:%S')
    with open ('.\\.config', 'r') as config_f:
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
    gw_requests = query_gw_requests()
    rec_requests = query_rec_requests()
    for request in gw_requests + rec_requests:
        html = create_html(request)
        save_html(html, request)
    convert_html()
