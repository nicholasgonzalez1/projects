import numpy as np
import xlwings as xw
import pandas as pd

def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def main(cntR_str, cntD_str):

    cntR = int(cntR_str)
    cntD = int(cntD_str)

    """===STEP 1: create pandas dataframe from return table==="""
    wb = xw.Book.caller()
    sht = wb.sheets['Return']
    cols = sht.range('A15:AZ15').value
    cols = list(filter(None, cols))

    dct_add = {}
    continueScript = True
    for i in range(len(cols)):
        # convert column index to column letter
        col = colnum_string(i+1)

        # determine range reference
        if (cntR == 1):
            rng = col + str(16)
            dct_add[cols[i]] = [sht.range(rng).value]
        elif (cntR > 1):
            j = 15 + cntR
            rng = col + str(16) + ':' + col + str(j)
            dct_add[cols[i]] = sht.range(rng).value
        else:
            continueScript = False
            break

    if (continueScript):
        df_add = pd.DataFrame(dct_add)
        df_add = df_add.applymap(str)

        """===STEP 2: create pandas dataframe for current database==="""
        sht = wb.sheets['Foods']
        cols = sht.range('A1:BZ1').value
        cols = list(filter(None, cols))

        dct_cur = {}
        for i in range(len(cols)):
            # convert column index to column letter
            col = colnum_string(i+1)

            # determine range reference
            if (cntD == 1):
                rng = col + str(2)
                dct_cur[cols[i]] = [sht.range(rng).value]
            elif (cntD > 1):
                j = 1 + cntD
                rng = col + str(2) + ':' + col + str(j)
                dct_cur[cols[i]] = sht.range(rng).value
            else:
                break

        df_cur = pd.DataFrame(dct_cur)
        df_cur.drop(columns=['index'], inplace=True)
        df_cur = df_cur.applymap(str)

        """===STEP 3: merge dataframes and update databases==="""
        df_new = df_cur.merge(df_add, how='outer')
        df_new = df_new.drop_duplicates(subset=['food_name'], keep='last')
        sht = wb.sheets['Foods']
        ro = len(df_new) + 1
        co = colnum_string(len(df_new.columns))
        rng = 'A1:' + str(co) + str(ro)
        sht.range(rng).value = df_new
        sht.range('A1').value = 'index'
