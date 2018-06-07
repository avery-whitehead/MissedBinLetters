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

class GardenWasteRequest():
    """
    Represents a request for more garden waste sacks
    """
    def __init__(self, occup: str, addr: str, case_ref: str, num_subs: str):
        """
        Args:
            occup (str): The occupier of the property
            addr (str): The address of the property
            case_ref (str): The case reference of the request
            num_subs (str): The number of subscriptions the property has
        """
        self.occup = occup
        self.addr = addr
        self.case_ref = case_ref
        self.num_subs = num_subs

class RecyclingRequest():
    """
    Represents a request for more recycling sacks
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

def query_gw_addresses(conn: pyodbc.Connection) -> list:
    """
    Queries the SQL database for the addresses of people who requested
    garden waste sacks
    Args:
        conn (pyodbc.Connection): The connection to the database
    Returns:
        A list of GardenWasteRequest objects containing the information
        from the query
    """
    with open('.\\gw_address_info.sql') as gw_query_f:
        gw_query = gw_query_f.read()
    cursor = conn.cursor()
    cursor.execute(gw_query)
    results = cursor.fetchall()
    print('--gw--')
    print(results)

def query_rec_addresses(conn: pyodbc.Connection) -> list:
    """
    Queries the SQL database for the addresses of people who requested
    recycling sacks
    Args:
        conn (pyodbc.Connection): The connection to the database
    Returns:
        A list of RecyclingRequest objects containing the information from
        the query
    """
    with open('.\\rec_address_info.sql') as rec_query_f:
        rec_query = rec_query_f.read()
    cursor = conn.cursor()
    cursor.execute(rec_query)
    results = cursor.fetchall()
    print('--rec--')
    print(results)


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
    query_gw_addresses(CONN)
    query_rec_addresses(CONN)