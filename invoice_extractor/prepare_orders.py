import csv
import json
from pathlib import Path
# Extract all orders that have Israel as a country from a specified file

order_file_path = Path(r"C:\Users\tmontoya\Python Projects\Small_Python_Projects\invoice_extractor\source-1.csv")
dest_file_path = Path("./invoice_extractor/data/target_orders.json")

with open(order_file_path, mode="r", encoding="utf-8-sig") as source_file:
    reader = csv.DictReader(source_file)

    order_dict = {}
    for entry in reader:
        if entry['ship_to_addr_3'] == "Israel" and entry['order_verify_init'] == "" \
            and entry['ship_to_addr_1'].lower().find("mock") == -1:
            order_dict.update({entry['order_number'] : entry['ship_to_addr_1']})

with open(dest_file_path, mode="w+", encoding="utf-8-sig") as result_file:
    json.dump(order_dict, result_file)
