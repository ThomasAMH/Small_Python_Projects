from pathlib import Path
import xlsxwriter
import csv
import json

def main():
    """
    Purpose: This script creates a price report by collecting data on orders from various files
    Input:
        - Correctly configured config.json
        - .csv files with ORDER_NUMBER, TRACKING_NUMBER and TOTAL fields
        - target_orders.csv file

    Output:
        - Creates an xlsx report with the order numbers and price data from across multiple files
    """

    #Paths

    

