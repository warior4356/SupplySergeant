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
import esipy
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


def generate_report(file_name, station_ids, ship_list, item_list, sheet_index, corporation_id, region_id, contracts=True):
    staging_ships = dict()  # [match, hull match, parts, hull_id, max fittable from market, target quantity]
    staging_charges = dict()  # [name, local count, local price,  jita volume, jita price, target quantity]
    staging_parts = dict()  # [name, local count, local price, jita volume, jita price, target quantity]
    contract_owners = dict()

    # Fetch items
    for item in item_list:
        staging_charges[int(item_ids[item[0]])] = [item[0], 0, 0, 0, 0, item[1]]

    # Parse ships
    for ship in ship_list:
        fitting = open("ships/" + ship[0] + ".txt", 'r', encoding='utf-8')
        ship_name = ship[0].strip("[]")
        lines = fitting.readlines()
        ship_id = int(item_ids[ship_name.split(",")[0]])
        staging_ships[ship_name] = [0, 0, {ship_id: 1}, ship_id, -1, ship[1]]
        staging_parts[int(item_ids[ship_name.split(",")[0]])] = [ship_name.split(",")[0], 0, 0, 0, 0, 0]
        for line in lines[1:]:
            if line.strip() in item_ids.keys():
                module_id = int(item_ids[line.strip()])
                staging_ships[ship_name][2][module_id] = staging_ships[ship_name][2].get(module_id, 0) + 1
                staging_parts[module_id] = [line.strip(), 0, 0, 0, 0, 0]
            elif line.strip().rsplit(',', 1)[0] in item_ids.keys():
                parts = line.strip().rsplit(',', 1)
                module_id = int(item_ids[parts[0]])
                # charge_id = int(item_ids[parts[1].lstrip(' ')])
                # print(charge_id)
                staging_ships[ship_name][2][module_id] = staging_ships[ship_name][2].get(module_id, 0) + 1
                # staging_ships[ship_name][2][charge_id] = staging_ships[ship_name][2].get(charge_id, 0) + 1
                staging_parts[module_id] = [parts[0], 0, 0, 0, 0, 0]
                # staging_charges[charge_id] = [parts[1].lstrip(' '), 0, 0, 0, 0]
            elif line.strip().rsplit(' ', 1)[0] in item_ids.keys():
                item_name = line.strip().rsplit(' ', 1)[0]
                count = line.strip().rsplit(' ', 1)[1].strip("x")
                charge_id = int(item_ids[item_name])
                staging_charges[charge_id] = [item_name, 0, 0, 0, 0, 0]
                staging_ships[ship_name][2][charge_id] = staging_ships[ship_name][2].get(charge_id, 0) + int(count)
        fitting.close()
    # print(staging_parts)

    # Get orders in citadel with given items
    for station_id in station_ids:
        market_results = []
        op = app.op['get_markets_structures_structure_id'](
            structure_id=station_id,
            page=1,
        )

        res = esi_client.head(op)

        if res.status == 200:
            number_of_pages = res.header['X-Pages'][0]

            # now we know how many pages we want, let's prepare all the requests
            operations = []
            for page in range(1, number_of_pages+1):
                operations.append(
                    app.op['get_markets_structures_structure_id'](
                        structure_id=station_id,
                        page=page,
                    )
                )

            market_results = esi_client.multi_request(operations)

        for pair in market_results:
            for result in pair[1].data:
                if not result.get("is_buy_order") and result.get("type_id") in staging_charges.keys():
                    staging_charges[result.get("type_id")][1] += result.get("volume_remain")
                    current_price = staging_charges[result.get("type_id")][2]
                    if current_price == 0 or current_price > result.get("price"):
                        staging_charges[result.get("type_id")][2] = result.get("price")
                if not result.get("is_buy_order") and result.get("type_id") in staging_parts.keys():
                    staging_parts[result.get("type_id")][1] += result.get("volume_remain")
                    current_price = staging_parts[result.get("type_id")][2]
                    if current_price == 0 or current_price > result.get("price"):
                        staging_parts[result.get("type_id")][2] = result.get("price")


    # Get orders in stations with given items
    for station_id in station_ids:
        market_results = []
        op = app.op['get_markets_region_id_orders'](
            region_id=region_id,
            order_type="sell",
            page=1,
        )

        res = esi_client.head(op)

        if res.status == 200:
            number_of_pages = res.header['X-Pages'][0]
            # print(number_of_pages)

            # now we know how many pages we want, let's prepare all the requests
            operations = []
            for page in range(1, number_of_pages + 1):
                operations.append(
                    app.op['get_markets_region_id_orders'](
                        region_id=region_id,
                        order_type="sell",
                        page=page,
                    )
                )

            market_results = esi_client.multi_request(operations)
            # print(market_results)

        for pair in market_results:
            for result in pair[1].data:
                if result.get("location_id") == station_id:
                    # print(result)
                    if result.get("type_id") in staging_charges.keys():
                        staging_charges[result.get("type_id")][1] += result.get("volume_remain")
                        current_price = staging_charges[result.get("type_id")][2]
                        if current_price == 0 or current_price > result.get("price"):
                            staging_charges[result.get("type_id")][2] = result.get("price")
                    if result.get("type_id") in staging_parts.keys():
                        staging_parts[result.get("type_id")][1] += result.get("volume_remain")
                        current_price = staging_parts[result.get("type_id")][2]
                        if current_price == 0 or current_price > result.get("price"):
                            staging_parts[result.get("type_id")][2] = result.get("price")

    # print(staging_parts)

    # Fetch Jita prices and volumes for comparison
    market_results = []
    op = app.op['get_markets_region_id_orders'](
        region_id=10000002,
        order_type="sell",
        page=1,
    )

    res = esi_client.head(op)

    if res.status == 200:
        number_of_pages = res.header['X-Pages'][0]
        # print(number_of_pages)

        # now we know how many pages we want, let's prepare all the requests
        operations = []
        for page in range(1, number_of_pages + 1):
            operations.append(
                app.op['get_markets_region_id_orders'](
                    region_id=10000002,
                    order_type="sell",
                    page=page,
                )
            )

        market_results = esi_client.multi_request(operations)
        # print(market_results)

    for pair in market_results:
        for result in pair[1].data:
            if result.get("location_id") == 60003760:
                if result.get("type_id") in staging_charges.keys():
                    staging_charges[result.get("type_id")][3] += result.get("volume_remain")
                    current_price = staging_charges[result.get("type_id")][4]
                    if current_price == 0 or current_price > result.get("price"):
                        staging_charges[result.get("type_id")][4] = result.get("price")
                if result.get("type_id") in staging_parts.keys():
                    staging_parts[result.get("type_id")][3] += result.get("volume_remain")
                    current_price = staging_parts[result.get("type_id")][4]
                    if current_price == 0 or current_price > result.get("price"):
                        staging_parts[result.get("type_id")][4] = result.get("price")


    # Determine how many ships can be fit from what's on market
    for ship in staging_ships.keys():
        for item in staging_ships[ship][2].keys():
            if item in staging_parts.keys() and staging_parts[item][1] > 0:
                if staging_ships[ship][4] == -1 or \
                        staging_parts[item][1] / staging_ships[ship][2][item] < staging_ships[ship][4]:
                    staging_ships[ship][4] = int(staging_parts[item][1] / staging_ships[ship][2][item])
            elif item in staging_charges.keys() and staging_charges[item][1] > 0:
                if staging_ships[ship][4] == -1 or \
                        staging_charges[item][1] / staging_ships[ship][2][item] < staging_ships[ship][4]:
                    staging_ships[ship][4] = int(staging_charges[item][1] / staging_ships[ship][2][item])
            else:
                staging_ships[ship][4] = 0

    # Determine target number of modules and charges based on target quantities
    for ship in staging_ships.keys():
        for item in staging_ships[ship][2].keys():
            if item in staging_parts.keys():
                staging_parts[item][5] = staging_parts[item][5] + (staging_ships[ship][2][item] * staging_ships[ship][5])
            if item in staging_charges.keys():
                staging_charges[item][5] = staging_charges[item][5] + (staging_ships[ship][2][item] * staging_ships[ship][5])

    if contracts:
        # Get corp/alliance contracts with ships
        contract_results = []

        op = app.op['get_corporations_corporation_id_contracts'](
            corporation_id=corporation_id,
            page=1,
        )

        res = esi_client.head(op)

        if res.status == 200:
            number_of_pages = res.header['X-Pages'][0]

            # now we know how many pages we want, let's prepare all the requests
            operations = []
            for page in range(1, number_of_pages+1):
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

                    contract_owners[result.get("contract_id")] = [result.get("title"), result.get("issuer_id")]

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
                        if all(item in contract_types for item in staging_ships[ship][2].keys()):
                            staging_ships[ship][0] += 1
                        elif staging_ships[ship][3] in contract_types:
                            staging_ships[ship][1] += 1

        # Get public contracts with ships
        contract_results = []

        op = app.op['get_contracts_public_region_id'](
            region_id=region_id,
            page=1,
        )

        res = esi_client.head(op)

        if res.status == 200:
            number_of_pages = res.header['X-Pages'][0]

            # now we know how many pages we want, let's prepare all the requests
            operations = []
            for page in range(1, number_of_pages + 1):
                operations.append(
                    app.op['get_contracts_public_region_id'](
                        region_id=region_id,
                        page=page,
                    )
                )

            contract_results = esi_client.multi_request(operations)

        # print(contract_results)
        for pair in contract_results:
            for result in pair[1].data:
                if result.get("type") == "item_exchange" \
                        and result.get("start_location_id") in station_ids and \
                        _convert_swagger_dt(result.get("date_expired")) > datetime.datetime.utcnow():
                    sleep(.5)
                    op = app.op['get_contracts_public_items_contract_id'](
                        contract_id=result.get("contract_id"),
                    )
                    # print(result)

                    contract_owners[result.get("contract_id")] = [result.get("title"), result.get("issuer_id")]

                    while (True):
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
                        if all(item in contract_types for item in staging_ships[ship][2].keys()):
                            staging_ships[ship][0] += 1
                        elif staging_ships[ship][3] in contract_types:
                            staging_ships[ship][1] += 1



        # print(staging_ships)

    # Write to outfile and google sheets
    out_file = open("output/" + file_name, 'w')

    # Ships
    sheet = client.open("Staging Stocks").get_worksheet(sheet_index)
    sheet.update_cell(1, 8, str(datetime.datetime.now(datetime.UTC)))
    cell_list = []
    row = 2
    out_file.write("\n\n\nShip Name,Number Found,Hull Match Only, Fits On Market\n")
    for key in staging_ships:
        out_file.write(key.strip(",") + "," + str(staging_ships[key][0]) + "," + str(staging_ships[key][1]) +
                       str(staging_ships[key][4]) + "\n")

        cell_list.append(Cell(row=row, col=1, value=key))
        cell_list.append(Cell(row=row, col=2, value=int(staging_ships[key][0])))
        cell_list.append(Cell(row=row, col=3, value=int(staging_ships[key][1])))
        cell_list.append(Cell(row=row, col=4, value=int(staging_ships[key][4])))
        cell_list.append(Cell(row=row, col=5, value=int(staging_ships[key][5])))
        row += 1

    sheet.update_cells(cell_list)

    # Charges
    sheet = client.open("Staging Stocks").get_worksheet(sheet_index + 1)
    sheet.update_cell(1, 11, str(datetime.datetime.now(datetime.UTC)))
    cell_list = []
    row = 2
    out_file.write("Item Name,Local Volume,Local Price,Jita Volume,Jita Price\n")
    for key in sorted(staging_charges.keys()):
        out_file.write(str(staging_charges[key][0]) + "," + str(staging_charges[key][1]) + str(staging_charges[key][2])
                       + str(staging_charges[key][3]) + str(staging_charges[key][4]) + "\n")

        cell_list.append(Cell(row=row, col=1, value=str(staging_charges[key][0])))
        cell_list.append(Cell(row=row, col=2, value=int(staging_charges[key][1])))
        cell_list.append(Cell(row=row, col=3, value=int(staging_charges[key][2])))
        cell_list.append(Cell(row=row, col=4, value=int(staging_charges[key][3])))
        cell_list.append(Cell(row=row, col=5, value=int(staging_charges[key][4])))
        cell_list.append(Cell(row=row, col=7, value=int(staging_charges[key][5])))
        row += 1

    sheet.update_cells(cell_list)

    # Parts
    sheet = client.open("Staging Stocks").get_worksheet(sheet_index + 2)
    sheet.update_cell(1, 11, str(datetime.datetime.now(datetime.UTC)))
    cell_list = []
    row = 2
    out_file.write("Item Name,Local Volume,Local Price,Jita Volume,Jita Price\n")
    for key in sorted(staging_parts.keys()):
        out_file.write(str(staging_parts[key][0]) + "," + str(staging_parts[key][1]) + str(staging_parts[key][2])
                       + str(staging_parts[key][3]) + str(staging_parts[key][4]) + "\n")

        cell_list.append(Cell(row=row, col=1, value=str(staging_parts[key][0])))
        cell_list.append(Cell(row=row, col=2, value=int(staging_parts[key][1])))
        cell_list.append(Cell(row=row, col=3, value=int(staging_parts[key][2])))
        cell_list.append(Cell(row=row, col=4, value=int(staging_parts[key][3])))
        cell_list.append(Cell(row=row, col=5, value=int(staging_parts[key][4])))
        cell_list.append(Cell(row=row, col=7, value=int(staging_parts[key][5])))
        row += 1

    sheet.update_cells(cell_list)

    out_file.write("\n\n\nContract ID,Contract Name,Contract Owner\n")
    for key in contract_owners:
        try:
            out_file.write(str(key) + "," + str(contract_owners[key][0]) + "," + str(contract_owners[key][1]) + "\n")
        except:
            out_file.write(str(key) + "," + "BAD NAME" + "," + str(contract_owners[key][1]) + "\n")

    out_file.close()

    # print(contract_owners)


def main():
    # get_refresh_token()
    print("Connection established")
    get_item_ids()
    print("Item ids fetched")
    # generate_report("elanoda.csv", [1040246076254], ships.elanoda, items.elanoda, 4, 1018389948, 10000016)
    # generate_report("enaluri.csv", [60015068], ships.enaluri, items.enaluri, 2, 1018389948, 10000069)
    # generate_report("UMI_KK.csv", [1036351551330], ships.preds, items.preds, 0, 1018389948, 10000010)
    generate_report("3T7_M8.csv", [1043621617719], ships._3T7_M8, items._3T7_M8, 0, 1018389948, 10000035, contracts=False)
    generate_report("Nakah.csv", [60014068], ships.Nakah, items.Nakah, 3, 1018389948, 10000001, contracts=False)
    # generate_report("F4R2_Q.csv", [1044008398262], ships.F4R2_Q, items.F4R2_Q, 3, 1018389948, 10000014)
    # generate_report("5ZXX_K.csv", [1038708751029, 1039071618828], ships._5ZXX_K, items._5ZXX_K, 4, 1018389948, 10000023)
    # print(check_location(60012580, 30002005))
    # print(check_location(1037022287754, 30002005))
    # print(check_location(1038708751029, 30002005))
    # print(location_cache)


if __name__ == "__main__":
    main()
