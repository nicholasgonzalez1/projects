
import numpy as np
import xlwings as xw
import sys
import csv
from datetime import datetime
import requests
import json
import pandas as pd
# import win32com.client

def main(product, searchType, rtnCount, x_app_id, x_app_key):

    """===STEP 1: clean variables and initiate workbook==="""
    wb = xw.Book.caller()
    product = product.replace("_", " ")
    rtnCount = int(rtnCount)

    """===STEP 2: initiate API search for related products==="""
    url = 'https://trackapi.nutritionix.com/v2/search/instant'
    headers = {
            'x-app-id':x_app_id,
            'x-app-key':x_app_key,
            'x-remote-user-id':'0'}
    data = {
        'query':product}

    response = requests.post(url, headers=headers, data=data)
    dct_search = json.loads(response.text) # dct for dictionary

    """===STEP 3: create pandas dataframe of all products==="""
    products = []
    for pdt in dct_search['branded']:

        #change 'photo' value to http link, not dictionary
        if type(pdt['photo']) is dict:
            pdt['photo'] = pdt['photo']['thumb']

        #add each product to products List
        products.append(pdt)
    df_search = pd.DataFrame(products)
    df_search = df_search.iloc[0:rtnCount]

    """===STEP 4: import nutrients csv==="""
    # dfn = pd.DataFrame(diction1['foods'][0]['full_nutrients'])
    df = pd.read_csv('Full Nutrient USDA Field Mapping.csv')
    df.set_index('attr_id', inplace=True)
    # dfn = pd.merge(dfn, df, on='attr_id', how='left')
    df['name (unit)'] = df['bulk_csv_field'] + " (" + df['unit'] + ")"

    """===STEP 5: locate nutrient values for top searches==="""
    dct_new = {'nix_item_id': [],
              'serving_weight_grams': []}
    item_count = 0

    for i in range(len(df_search)):
        # formulate url
        item_id = df_search.loc[i, 'nix_item_id']
        base_url = 'https://trackapi.nutritionix.com/v2/search/item?nix_item_id='
        url = base_url + item_id

        # execute API search for specific item
        headers = {
                'x-app-id':x_app_id,
                'x-app-key':x_app_key,
                'x-remote-user-id':'0'}
        response = requests.get(url, headers=headers)
        dct_item = json.loads(response.text)
        dct_item = dct_item['foods'][0]

        # input new nix_item_id and serving_weight_grams within each iteration
        dct_new['nix_item_id'].append(dct_item['nix_item_id'])
        dct_new['serving_weight_grams'].append(dct_item['serving_weight_grams'])

        for nutrient in dct_item['full_nutrients']:
            if df.loc[nutrient['attr_id'],'name (unit)'] in dct_new:
                dct_new[df.loc[nutrient['attr_id'],'name (unit)']].append(nutrient['value']) # add new value for new row
            else:
                dct_new[df.loc[nutrient['attr_id'],'name (unit)']] = [] # make new column
                for i in range(item_count):
                    dct_new[df.loc[nutrient['attr_id'],'name (unit)']].append('NaN') # add N/A for all previous rows
                dct_new[df.loc[nutrient['attr_id'],'name (unit)']].append(nutrient['value']) # add new value for new row

        item_count += 1
        for key in dct_new:
            if len(dct_new[key]) != item_count:
                dct_new[key].append('NaN')

    df_new = pd.DataFrame(dct_new)

    """===STEP 6: Connect original search DataFrame with nutrients DataFrame==="""
    df_search.drop(columns=['nf_calories'], inplace=True)
    df_final = pd.merge(df_search, df_new, on='nix_item_id', how='left')

    """===STEP 7: Write df_final to MACROS.xlsx==="""
    # Take the data frame object and convert it to a recordset array, then to list
    rec_array = df_final.to_records()
    rec_array = rec_array.tolist()

    # set the value property equal to the record array.
    wb.sheets['Return'].range("A1:AZ20").value = ""
    wb.sheets['Return'].range("A1:AZ1").value = df_final.columns.to_list()

    #rowNum = {1:"B", 2:"C", 3:"D", 4:"E", 5:"F", 6:"G", 7:"H", 8:"I", 9:"J", 10:"K"}
    for i in range(len(rec_array)):
        rng = "A" + str(i+2) + ":AZ" + str(i+2)
        wb.sheets['Return'].range(rng).value = list(rec_array[i])[1:]
