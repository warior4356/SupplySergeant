import cfg
from esipy import EsiApp
from esipy import EsiClient
from esipy import EsiSecurity
from time import sleep

item_counts = dict()
ship_counts = dict()


for key in cfg.items:
    item_counts[key] = 0

for key in cfg.ships:
    ship_counts[key] = 0


app = EsiApp().get_latest_swagger

security = EsiSecurity(
    redirect_uri='http://localhost:5000/callback',
    client_id='6979f412d213468481c36acb397f5b1f',
    secret_key= cfg.secret,
    headers={'User-Agent': cfg.agent},
)

esi_client = EsiClient(
    retry_requests=True,
    headers={'User-Agent': cfg.agent},
    security=security
)

# print(security.get_auth_uri(state='SomeRandomGeneratedState', scopes=['esi-markets.structure_markets.v1',
#                                                                       'esi-contracts.read_corporation_contracts.v1']))
security.update_token({
    'access_token': '',  # leave this empty
    'expires_in': -1,  # seconds until expiry, so we force refresh anyway
    'refresh_token': cfg.refresh_token
})

tokens = security.refresh()

# Get orders in citadel with given items
market_results = []

op = app.op['get_markets_structures_structure_id'](
    structure_id=1032766218625,
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
                structure_id=1032766218625,
                page=page,
            )
        )

    market_results = esi_client.multi_request(operations)

# print(market_results[0][1].data)

for pair in market_results:
    for result in pair[1].data:
        if not result.get("is_buy_order") and result.get("type_id") in cfg.items.keys():
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
                and result.get("start_location_id") == 1032766218625:
            sleep(.2)
            op = app.op['get_corporations_corporation_id_contracts_contract_id_items'](
                contract_id=result.get("contract_id"),
                corporation_id=1018389948,
            )
            contents = esi_client.request(op)
            # print(contents.data)

            for item in contents.data:
                # print(item.get("type_id"))
                if item.get("type_id") in cfg.ships.keys():
                    ship_counts[item.get("type_id")] += 1
                    break

print(ship_counts)

out_file = open("outfile.csv", 'w')
out_file.write("Item Name,Number Found,Number Expected,Excess\n")
for key in item_counts:
    out_file.write(cfg.items[key][0] + "," + str(item_counts[key]) + "," + str(cfg.items[key][1])
                   + "," + str(item_counts[key] - cfg.items[key][1]) + "\n")

out_file.write("\n\n\nShip Name,Number Found,Number Expected,Excess\n")
for key in ship_counts:
    out_file.write(cfg.ships[key][0] + "," + str(ship_counts[key]) + "," + str(cfg.ships[key][1])
                   + "," + str(ship_counts[key] - cfg.ships[key][1]) + "\n")

out_file.close()
