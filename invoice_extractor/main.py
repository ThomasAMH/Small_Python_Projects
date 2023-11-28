import csv
import json
import os
from pathlib import Path
import xlsxwriter

#FIXME: Add ability to specify currency at the end!
#FIXME: Remove the way the order info dict is prepared. :P

def main():
    """
    Purpose: This script creates a price report by collecting data on orders from various files
    Input:
        - Correctly configured config.json
        - In the data folder
            - .csv files with ORDER_NUMBER, TRACKING_NUMBER and TOTAL fields
            - target_orders.json file

    Output:
        - Creates an xlsx report with the order numbers and price data from across multiple files
    """

    # Paths
    config_path = Path("invoice_extractor/config.json")
    order_file_path = Path("invoice_extractor/data/target_orders.json")
    data_dir = "invoice_extractor/data/"
    completed_order_path = Path("invoice_extractor/data/Completed Report.xlsx")

    # Set config
    with open(config_path, mode = "r", encoding="utf-8-sig") as config_file:
        config_dict = json.load(config_file)

    # Read in target orders into dictionary and set relevant properties
        with open(order_file_path, mode = "r", encoding="utf-8-sig") as order_file_path:
            read_order_info_dict = json.load(order_file_path)
            
            order_info_dict = {}
            for key, value in read_order_info_dict.items():
                temp_obj = ({"street_address" : value})
                temp_obj['tracking_number'] = ""
                order_info_dict.update({clean_order_number(key) : temp_obj})
        
    # Go through each file in each report directory and gather information on the orders
    #Run twice so that any orders that may be only found with a tracking number have a chance to be found
    tracking_number_dict = {}
    loop_through_reports(config_dict, data_dir, order_info_dict, tracking_number_dict)
    loop_through_reports(config_dict, data_dir, order_info_dict, tracking_number_dict)


    # Prepare and write the xlsx report
    workbook = xlsxwriter.Workbook(completed_order_path)
    report_sheet = workbook.add_worksheet("Report")
    report_sheet.write("A1", "Currency")
    report_sheet.write("B1", config_dict['target_currency'])
    report_sheet.write("A2", "Order Number")
    report_sheet.write("B2", "Tracking Number")
    report_sheet.write("C2", "Additional Costs")
    report_sheet.write("D2", "Clearance")
    report_sheet.write("E2", "DG")
    report_sheet.write("F2", "DHL Express")
    report_sheet.write("G2", "Metapack")
    report_sheet.write("H2", "Invoice Detail")
    report_sheet.write("I2", "Total")

    active_row = 3
    grand_total = 0.0
    for order, value in order_info_dict.items():
        running_total = 0.0

        report_sheet.write("A"+str(active_row), order)
        report_sheet.write("B"+str(active_row), value['tracking_number'])

        if "additional_cost" in value:
            report_sheet.write("C"+str(active_row), value['additional_cost'])
            running_total += value['additional_cost']
        
        if "clearance_details" in value:
            report_sheet.write("D"+str(active_row), value['clearance_details'])
            running_total += value['clearance_details']

        if "doterra_dg" in value:
            report_sheet.write("E"+str(active_row), value['doterra_dg'])
            running_total += value['doterra_dg']

        if "doterra_dhl_express" in value:
            report_sheet.write("F"+str(active_row), value['doterra_dhl_express'])
            running_total += value['doterra_dhl_express']

        if "doterra_metapack" in value:
            report_sheet.write("G"+str(active_row), value['doterra_metapack'])
            running_total += value['doterra_metapack']

        if "invoices_detail" in value:
            report_sheet.write("H"+str(active_row), value['invoices_detail'])
            running_total += value['invoices_detail']

        report_sheet.write("I"+str(active_row), running_total)
        grand_total += running_total
    
        active_row += 1
    report_sheet.write("D1", "Grand Total")
    report_sheet.write("E1", grand_total)
    workbook.close()

    print("Report complete!")
    

def loop_through_reports(config_dict, data_dir, order_info_dict, tracking_number_dict):
    for report in config_dict['reports']:
        current_dir = Path(data_dir + report)

        for report_file in current_dir.iterdir():
            with open(report_file, mode = "r", encoding="utf-8-sig", errors="replace") as current_file:
                read_data = csv.DictReader(current_file)

                # Steps for each order
                for line in read_data:

                    if 'ORDER_NUMBER' in line:
                        # Clean order data, as needed
                        if config_dict['reports'][report]['clean_order_number']:
                            order_to_lookup = clean_order_number(line['ORDER_NUMBER'])
                        else:
                            order_to_lookup = line['ORDER_NUMBER']

                        # Added the second clause here because this function is called more than once.
                        # On subseqent calls, if the order already has the file path as a property, no further
                        # Checks are required
                        if order_to_lookup in order_info_dict and report not in order_info_dict[order_to_lookup]:

                        # If there is a order-tracking number pair and it is not assigned, assign it
                            if 'TRACKING_NUMBER' in line:
                                order_info_dict[clean_order_number(line['ORDER_NUMBER'])].update({'tracking_number' : line['TRACKING_NUMBER']})
                                tracking_number_dict.update({clean_order_number(line['ORDER_NUMBER']) : line['TRACKING_NUMBER']})

                        # Add up the value from the TOTAL column, convering currency
                            quantity_to_add = get_quantity_to_add(config_dict, report, line)
                            add_quantity(order_info_dict, quantity_to_add, line, "ORDER_NUMBER", report)

                    # If you can't find the order based on the ORDER_NUMBER, try looking through the tracking codes
                    # This should catch all tracking-only files on the second pass, where the tracking was known somewhere
                    elif 'TRACKING_NUMBER' in line:
                        if line['TRACKING_NUMBER'] in tracking_number_dict:
                            quantity_to_add = get_quantity_to_add(config_dict, report, tracking_number_dict[line['TRACKING_NUMBER']])
                            add_quantity(order_info_dict, quantity_to_add, line, "TRACKING_NUMBER", report, tracking_number_dict)  

    



# Helper functions
def clean_order_number(order_string):
    """
    If there is "DT" or "_DOTERRA" in the order number, remove it
    """
    if order_string.find("DT") != -1:
        order_string = order_string.replace("DT", "")

    if order_string.find("_DOTERRA") != -1:
        order_string = order_string.replace("_DOTERRA", "")

    return order_string

def get_quantity_to_add(config_dict, report, line):
    # Add up the value from the TOTAL column, convering currency conversion as needed
        if config_dict['reports'][report]['currency'] != config_dict['target_currency']:
            return currency_convert(config_dict, line['TOTAL'], config_dict['reports'][report]['currency'], config_dict['target_currency'])
        else:
            return float(line['TOTAL'])

    

def add_quantity(order_info_dict, quantity_to_add, line, ref_type, report, tracking_number_dict = None):
    # Add quanaty and a reference to the file from whence the value comes, if such a reference does not yet exist

    if ref_type == 'ORDER_NUMBER':
        if report not in order_info_dict[clean_order_number(line['ORDER_NUMBER'])]:
            order_info_dict[clean_order_number(line['ORDER_NUMBER'])].update({report : quantity_to_add})
        else:
            order_info_dict[line['ORDER_NUMBER']][report] += quantity_to_add

    elif ref_type == 'TRACKING_NUMBER':
        order_number = tracking_number_dict[line['TRACKING_NUMBER']]
        if report not in order_info_dict[order_number]:
            order_info_dict[order_number].update({report : quantity_to_add})
        else:
            order_info_dict[order_number][report] += quantity_to_add
    

def currency_convert(config_dict, amount, current_currency, target_currency):
    """
    Implement the currency converter
    """
    currency_convert_property_name = current_currency + "-" + target_currency
    return float(amount) * float(config_dict[currency_convert_property_name])
            

if __name__ == "__main__":
    main()