"""
gen_html.py

Fills in some HTML templates with address information to be sent out to
people who have requested more waste sacks
Uses wkhtmltopdf to convert the templates to an A4 PDF
"""
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

def query_gw_requests(conn: pyodbc.Connection) -> list:
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
    with open('.\\gw_address_info.sql') as gw_query_f:
        gw_query = gw_query_f.read()
    cursor = conn.cursor()
    cursor.execute(gw_query)
    results = cursor.fetchall()
    for result in results:
        gw_requests.append(GardenWasteRequest(
            result.occupier,
            result.address,
            result.case_ref,
            result.num_subs))
    return gw_requests

def query_rec_requests(conn: pyodbc.Connection) -> list:
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
    with open('.\\rec_address_info.sql') as rec_query_f:
        rec_query = rec_query_f.read()
    cursor = conn.cursor()
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
    """
    html = '' \
        '<!DOCTYPE html>\n' \
        '<html>\n' \
        '<head>\n' \
        '    <link rel="stylesheet" href=".\\paper.css">\n' \
        '    <style>\n' \
        '        body {\n' \
        '            font-family: sans-serif;\n' \
        '        }\n' \
        '        .addr {\n' \
        '            font-weight: bold;\n' \
        '            padding-left: 100px;\n' \
        '            padding-top: 175px;\n' \
        '        }\n' \
        '        @page {\n' \
        '            size: A4;\n' \
        '        }\n' \
        '    </style>\n' \
        '</head>\n' \
        '<body class="A4">\n' \
        '    <section class="sheet padding-10mm">\n' \
        '        <div class="addr">\n' \
       f'            {request.addr}\n' \
        '        </div>\n' \
        '    </section>\n' \
        '</body>\n' \
        '</html>'
    print(html)

if __name__ == '__main__':
    SYSTIME = datetime.datetime.now().strftime('%d-%b-%Y %H:%M:%S')
    with open ('.\\.config') as config_f:
        config = json.load(config_f)
    try:
        CONN = pyodbc.connect(
            driver=config['driver'],
            server=config['server'],
            database=config['database'],
            uid=config['uid'],
            pwd=config['pwd'])
    except pyodbc.InterfaceError as error:
        with open('.\\missed_bin_letters', 'a') as log:
            log.write(f'{SYSTIME} - {error}\n')
        sys.exit(1)
    gw_requests = query_gw_requests(CONN)
    rec_requests = query_rec_requests(CONN)
    for rec_request in rec_requests:
        create_html(rec_request)
