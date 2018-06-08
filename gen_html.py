"""
gen_html.py
How to run:
Install pyodbc version 4.0.22 or greater
From the command line, run `py -3 gen_html.py`
How it works:
Generates an HTML file and uses wkhtmltopdf to convert it to a PDF
"""
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
        Args:
            num_subs (str): The number of subscriptions the property has
        """
        super().__init__(occup, addr, case_ref)
        self.num_subs = num_subs


class RecyclingRequest(Request):
    """
    Represents a request for more recycling sacks
    """
    def __init__(self, occup: str, addr: str, case_ref: str):
        """
        No different to the parent Request class, so has no extra attributes
        """
        super().__init__(occup, addr, case_ref)

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
    html = '' \
        '<!DOCTYPE html>\n' \
        '<html>\n' \
        '<head>\n' \
        '    <style>\n' \
        '        body {\n' \
        '            font-family: sans-serif;\n' \
        '        }\n' \
        '        .addr {\n' \
        '            font-weight: bold;\n' \
        '            padding-left: 100px;\n' \
        '            padding-top: 175px;\n' \
        '            font-size: 22px;\n' \
        '        }\n' \
        '    </style>\n' \
        '</head>\n' \
        '<body>\n' \
        '    <section>\n' \
        '        <div class="addr">\n' \
       f'            {request.addr}\n' \
        '        </div>\n' \
        '    </section>\n' \
        '</body>\n' \
        '</html>'
    return html

def save_and_convert_html(html: str, request: Request) -> str:
    """
    Writes the HTML to file, calls the wkhtmltopdf process and converts
    that file to a PDF
    Args:
        html (str): The HTML to write to file and later convert
        request (Request): The Request this HTML was generated from
    Returns:
        (str): A string indicating success
    """
    try:
        wkhtml_path = '"C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"'
        html_path = f'.\\htmls\\{request.case_ref}.html'
        pdf_path = f'.\\pdfs\\{request.case_ref}.pdf'
        with open(html_path, 'w+') as html_f:
            html_f.write(html)
        args = f'{wkhtml_path} {html_path} {pdf_path}'
        subprocess.call(args, shell=False)
        success_str = f'{SYSTIME} - Successfully created {pdf_path}'
        with open('.\\missed_bin_letters.log', 'a') as log:
            log.write(f'{success_str}\n')
        return success_str
    except (IOError, FileNotFoundError) as error:
        with open('.\\missed_bin_letters.log', 'a') as log:
            log.write(f'{SYSTIME} - {error}\n')
        sys.exit(1)

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
        with open('.\\missed_bin_letters', 'a') as log:
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
        with open('.\\missed_bin_letters', 'a') as log:
            log.write(f'{SYSTIME} - {error}\n')
        sys.exit(1)
    gw_requests = query_gw_requests()
    rec_requests = query_rec_requests()
    for request in gw_requests + rec_requests:
        html = create_html(request)
        print(save_and_convert_html(html, request))
        print(update_database(request))
