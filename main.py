import cfg
from esipy import EsiApp
from esipy import EsiClient
from esipy import EsiSecurity
from time import sleep
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread.models import Cell

app = EsiApp().get_latest_swagger

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
                                                            'esi-contracts.read_corporation_contracts.v1']))
    print(security.auth(cfg.auth_code))

def _convert_swagger_dt(dt) -> datetime.datetime:
    """Converts a pyswagger timestamp.
    Args:
    dt: pyswagger timestamp
    Returns:
    Python stdlib datetime object
    """

    return datetime.datetime.strptime(dt.to_json(), '%Y-%m-%dT%H:%M:%S+00:00')


def generate_report(file_name, station, ships, items, sheet_index):
    item_counts = dict()
    ship_counts = dict()


    for key in items:
        item_counts[key] = 0

    for key in ships:
        ship_counts[key] = 0

    security.update_token({
        'access_token': '',  # leave this empty
        'expires_in': -1,  # seconds until expiry, so we force refresh anyway
        'refresh_token': cfg.refresh_token
    })

    tokens = security.refresh()

    # Get orders in citadel with given items
    market_results = []

    op = app.op['get_markets_structures_structure_id'](
        structure_id=station,
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
                    structure_id=station,
                    page=page,
                )
            )

        market_results = esi_client.multi_request(operations)

    for pair in market_results:
        for result in pair[1].data:
            if not result.get("is_buy_order") and result.get("type_id") in items.keys():
                item_counts[result.get("type_id")] += result.get("volume_remain")

    print(item_counts)

    # Get contracts with given items
    contract_results = []

    op = app.op['get_corporations_corporation_id_contracts'](
        corporation_id=1018389948,
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
                    corporation_id=1018389948,
                    page=page,
                )
            )

        contract_results = esi_client.multi_request(operations)

    # print(contract_results)
    for pair in contract_results:
        for result in pair[1].data:
            if result.get("status") == "outstanding" and result.get("type") == "item_exchange" \
                    and result.get("start_location_id") == station and \
                    _convert_swagger_dt(result.get("date_expired")) > datetime.datetime.utcnow():
                sleep(.5)
                op = app.op['get_corporations_corporation_id_contracts_contract_id_items'](
                    contract_id=result.get("contract_id"),
                    corporation_id=1018389948,
                )
                # print(result)
                while(True):
                    try:
                        contents = esi_client.request(op)
                    except:
                        continue
                    break

                for item in contents.data:
                    # if(item.get("type_id") == 17843):
                    #     print(result)
                    #     print(result.get("date_expired"))
                    #     print(datetime.datetime.utcnow())
                    #     if _convert_swagger_dt(result.get("date_expired")) < datetime.datetime.utcnow():
                    #         print("expired")
                    # print(item.get("type_id"))
                    # print(item)
                    if type(item) == str:
                        print(item)
                    else:
                        if item.get("type_id") in ships.keys():
                            ship_counts[item.get("type_id")] += 1
                            break

    print(ship_counts)

    out_file = open(file_name, 'w')
    sheet = client.open("Staging Stocks").get_worksheet(sheet_index)
    sheet.update_cell(1, 7, str(datetime.datetime.utcnow()))
    cell_list = []
    row = 2
    out_file.write("Item Name,Number Found,Number Expected,Missing\n")
    for key in item_counts:
        out_file.write(items[key][0] + "," + str(item_counts[key]) + "," + str(items[key][1])
                       + "," + str(items[key][1] - item_counts[key]) + "\n")

        cell_list.append(Cell(row=row, col=1, value=str(items[key][0])))
        cell_list.append(Cell(row=row, col=2, value=item_counts[key]))
        cell_list.append(Cell(row=row, col=3, value=items[key][1]))
        cell_list.append(Cell(row=row, col=4, value=items[key][1] - item_counts[key]))
        row += 1

    sheet.update_cells(cell_list)

    sheet = client.open("Staging Stocks").get_worksheet(sheet_index + 1)
    sheet.update_cell(1, 7, str(datetime.datetime.utcnow()))
    cell_list = []
    row = 2
    out_file.write("\n\n\nShip Name,Number Found,Number Expected,Missing\n")
    for key in ship_counts:
        out_file.write(ships[key][0] + "," + str(ship_counts[key]) + "," + str(ships[key][1])
                       + "," + str(ships[key][1] - ship_counts[key]) + "\n")

        cell_list.append(Cell(row=row, col=1, value=str(ships[key][0])))
        cell_list.append(Cell(row=row, col=2, value=ship_counts[key]))
        cell_list.append(Cell(row=row, col=3, value=ships[key][1]))
        cell_list.append(Cell(row=row, col=4, value=ships[key][1] - ship_counts[key]))
        row += 1

    sheet.update_cells(cell_list)

    out_file.close()


def main():
    # get_refresh_token()

    # generate_report("7R5_7R.csv", 1032766218625, cfg.ships_7R5_7R, cfg.items_7R5_7R, 0)
    # generate_report("C0O6_K.csv", 1032819384255, cfg.ships_C0O6_K, cfg.items_C0O6_K, 2)
    # generate_report("FAT_6P.csv", 1033753242053, cfg.ships_FAT_6P, cfg.items_FAT_6P, 0)
    # generate_report("D_PNP9.csv", 1024004680659, cfg.ships_D_PNP9, cfg.items_D_PNP9, 0)
    # generate_report("YZ9_F6.csv", 1034592395985, cfg.ships_YZ9_F6, cfg.items_YZ9_F6, 0)
    # generate_report("T5ZI_S.csv", 1034877491366, cfg.ships_T5ZI_S, cfg.items_T5ZI_S, 0)
    # generate_report("YAW_7M.csv", 1034857122560, cfg.ships_YAW_7M, cfg.items_YAW_7M, 2)
    generate_report("T0DT_T.csv", 1037022454355, cfg.ships_T0DT_T, cfg.items_T0DT_T, 0)

if __name__ == "__main__":
  main()