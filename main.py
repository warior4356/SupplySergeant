import collections
try:
    from collections import abc
    collections.MutableMapping = abc.MutableMapping
    collections.Mapping = abc.Mapping
except:
    pass

import cfg
import items
import ships
from esipy import EsiApp
from esipy import EsiClient
from esipy import EsiSecurity
import operator
from time import sleep
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread import Cell

app = EsiApp().get_latest_swagger
# Item Id format {Name: item_id}
item_ids = dict()
location_cache = {}

security = EsiSecurity(
    redirect_uri='http://localhost:5000/callback',
    client_id=cfg.client_id,
    secret_key= cfg.secret,
    headers={'User-Agent': cfg.agent},
)

esi_client = EsiClient(
    retry_requests=True,
    headers={'User-Agent': cfg.agent},
    security=security
)

security.update_token({
    'access_token': '',  # leave this empty
    'expires_in': -1,  # seconds until expiry, so we force refresh anyway
    'refresh_token': cfg.refresh_token
})

# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
# sheet = client.open("Staging Stocks").get_worksheet(1)

# Extract and print all of the values
# list_of_hashes = sheet.get_all_records()
# print(list_of_hashes)

def get_refresh_token():
    print(security.get_auth_uri(state='1234567890', scopes=['esi-markets.structure_markets.v1',
                                                            'esi-contracts.read_corporation_contracts.v1',
                                                            'esi-universe.read_structures.v1']))
    print(security.auth(cfg.auth_code))


def _convert_swagger_dt(dt) -> datetime.datetime:
    """Converts a pyswagger timestamp.
    Args:
    dt: pyswagger timestamp
    Returns:
    Python stdlib datetime object
    """

    return datetime.datetime.strptime(dt.to_json(), '%Y-%m-%dT%H:%M:%S+00:00')


def check_location(location_id, system_id):
    if location_id in location_cache.keys():
        return location_cache[location_id] == system_id

    if 60000000 < location_id < 64000000:
        op = app.op['get_universe_stations_station_id'](
            station_id=location_id
        )
        contents = esi_client.request(op)
        station_system = contents.data['system_id']
        location_cache[location_id] = station_system
        return station_system == system_id
    else:
        op = app.op['get_universe_structures_structure_id'](
            structure_id=location_id
        )
        contents = esi_client.request(op)
        print(contents.data)
        structure_system = contents.data['solar_system_id']
        location_cache[location_id] = structure_system
        return structure_system == system_id


def get_item_ids():
    types = open('invTypes.csv', 'r', encoding='utf-8')
    lines = types.readlines()
    for line in lines:
        parts = line.split(",")
        if len(parts) > 2:
            item_ids[parts[2]] = parts[0]

    types.close()


def generate_report(file_name, station_ids, ship_list, item_list, sheet_index, corporation_id):
    staging_ships = dict()
    staging_items = dict()

    # Fetch items
    for item in item_list:
        staging_items[int(item_ids[item])] = [0, item]

    # Parse ships
    for ship_name in ship_list:
        fitting = open("ships/" + ship_name + ".txt", 'r', encoding='utf-8')
        ship_name = ship_name.strip("[]")
        lines = fitting.readlines()
        staging_ships[ship_name] = [0, 0, [int(item_ids[ship_name.split(",")[0]])]]
        for line in lines[1:]:
            if line.strip() in item_ids.keys():
                staging_ships[ship_name][2].append(int(item_ids[line.strip()]))
            elif line.strip().rsplit(',', 1)[0] in item_ids.keys():
                parts = line.strip().rsplit(',', 1)
                staging_ships[ship_name][2].append(int(item_ids[parts[0]]))
                staging_items[int(item_ids[parts[1].lstrip(' ')])] = [0, parts[1].lstrip(' ')]
            elif line.strip().rsplit(' ', 1)[0] in item_ids.keys():
                staging_items[int(item_ids[line.strip().rsplit(' ', 1)[0]])] = [0, line.strip().rsplit(' ', 1)[0]]
        fitting.close()

    # Get orders in citadel with given items
    for station_id in station_ids:
        market_results = []
        op = app.op['get_markets_structures_structure_id'](
            structure_id=station_id,
            page=1,
        )

        res = esi_client.head(op)

        if res.status == 200:
            number_of_page = res.header['X-Pages'][0]

            # now we know how many pages we want, let's prepare all the requests
            operations = []
            for page in range(1, number_of_page+1):
                operations.append(
                    app.op['get_markets_structures_structure_id'](
                        structure_id=station_id,
                        page=page,
                    )
                )

            market_results = esi_client.multi_request(operations)

        for pair in market_results:
            for result in pair[1].data:
                if not result.get("is_buy_order") and result.get("type_id") in staging_items.keys():
                    staging_items[result.get("type_id")][0] += result.get("volume_remain")

    # print(staging_items)

    # Get contracts with ships
    contract_results = []

    op = app.op['get_corporations_corporation_id_contracts'](
        corporation_id=corporation_id,
        page=1,
    )

    res = esi_client.head(op)

    if res.status == 200:
        number_of_page = res.header['X-Pages'][0]

        # now we know how many pages we want, let's prepare all the requests
        operations = []
        for page in range(1, number_of_page+1):
            operations.append(
                app.op['get_corporations_corporation_id_contracts'](
                    corporation_id=corporation_id,
                    page=page,
                )
            )

        contract_results = esi_client.multi_request(operations)

    # print(contract_results)
    for pair in contract_results:
        for result in pair[1].data:
            if result.get("status") == "outstanding" and result.get("type") == "item_exchange" \
                    and result.get("start_location_id") in station_ids and \
                    _convert_swagger_dt(result.get("date_expired")) > datetime.datetime.utcnow():
                sleep(.5)
                op = app.op['get_corporations_corporation_id_contracts_contract_id_items'](
                    contract_id=result.get("contract_id"),
                    corporation_id=corporation_id,
                )
                # print(result)
                while(True):
                    try:
                        contents = esi_client.request(op)
                    except:
                        sleep(3)
                        continue
                    break

                contract_types = list()

                for item in contents.data:
                    contract_types.append(item.get("type_id"))

                for ship in staging_ships.keys():
                    if all(item in contract_types for item in staging_ships[ship][2]):
                        staging_ships[ship][0] += 1
                    elif staging_ships[ship][2][0] in contract_types:
                        staging_ships[ship][1] += 1


    # Write to outfile and google sheets
    out_file = open("output/" + file_name, 'w')
    sheet = client.open("Staging Stocks").get_worksheet(sheet_index + 1)
    sheet.update_cell(1, 7, str(datetime.datetime.utcnow()))
    cell_list = []
    row = 2
    out_file.write("Item Name,Number Found,Number Expected,Missing\n")
    for key in sorted(staging_items.keys()):
        out_file.write(str(staging_items[key][1]) + "," + str(staging_items[key][0]) + "\n")

        cell_list.append(Cell(row=row, col=1, value=str(staging_items[key][1])))
        cell_list.append(Cell(row=row, col=2, value=int(staging_items[key][0])))
        row += 1

    sheet.update_cells(cell_list)

    sheet = client.open("Staging Stocks").get_worksheet(sheet_index)
    sheet.update_cell(1, 9, str(datetime.datetime.utcnow()))
    cell_list = []
    row = 2
    out_file.write("\n\n\nShip Name,Number Found,Hull Match Only,Number Expected,Missing\n")
    for key in staging_ships:
        out_file.write(key.strip(",") + "," + str(staging_ships[key][0]) + "," + str(staging_ships[key][1]) +
                       str(staging_ships[key][2]) + "\n")

        cell_list.append(Cell(row=row, col=1, value=key))
        cell_list.append(Cell(row=row, col=2, value=int(staging_ships[key][0])))
        cell_list.append(Cell(row=row, col=3, value=int(staging_ships[key][1])))
        row += 1

    sheet.update_cells(cell_list)

    out_file.close()


def main():
    # get_refresh_token()
    print("Connection established")
    get_item_ids()
    print("Item ids fetched")
    generate_report("elanoda.csv", [1040246076254], ships.elanoda, items.elanoda, 0, 1018389948)
    generate_report("enaluri.csv", [60015068], ships.enaluri, items.enaluri, 2, 1018389948)
    generate_report("5ZXX_K.csv", [1038708751029, 1039071618828], ships._5ZXX_K, items._5ZXX_K, 4, 1018389948)
    # print(check_location(60012580, 30002005))
    # print(check_location(1037022287754, 30002005))
    # print(check_location(1038708751029, 30002005))
    # print(location_cache)


if __name__ == "__main__":
    main()
