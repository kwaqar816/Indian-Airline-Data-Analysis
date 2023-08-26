import os
import pandas as pd
import time
print(os.getcwd())
for each_excel_file in os.listdir('{}\Raw'.format(os.getcwd())):
    if '.xlsx' in each_excel_file:
        print(each_excel_file)
        intial_excel_file_name = each_excel_file.split('_')
        df = pd.read_excel("{}\Raw\{}".format(
            os.getcwd(), each_excel_file), header=None)
        for k in range(0, df.shape[0]):
            for s in range(df.shape[1]):
                print(df.iat[k, s])
                if 'TABLE 1.' in str(df.iat[k, s]):
                    print('We do not need these types of excel file')
                    time.sleep(5)
                    print('Completed in TABLE 1.')

                if 'TABLE 2.' in str(df.iat[k, s]):
                    print('FOUND IT 2')

                    #######################################

                # Finding specific row LOCATION to make it column heading
                    for row in range(0, df.shape[0]):
                        for col in range(df.shape[1]):
                            if (df.iat[row, col] == 'SL. No.') or (df.iat[row, col] == 'NAME OF THE AIRLINE') or (df.iat[row, col] == 'JANUARY'):
                                column_row_location = row

                    # Appending entire uncleaned row to temp_col variable
                    temp_col = []
                    new_column = []
                    for i in df.loc[column_row_location]:
                        i = str(i)
                        if i != 'nan':
                            temp_col.append(i)
                            prev = i
                        else:
                            i = prev + i
                            temp_col.append(i)

                    # Subsituting that row LOCATION with new temp_col
                    df.loc[column_row_location] = temp_col

                    # Merging that row LOCATION and row below it and cleaning it
                    for i, j in zip(df.loc[column_row_location], df.loc[column_row_location+1]):
                        i = str(i).replace('nan', '')
                        j = str(j).replace('nan', '')
                        new_column.append(i + ' ' + j)
                    # print(new_column)

                    row_limits = [['DOMESTIC CARRIERS', 'TOTAL (DOMESTIC CARRIERS)', 'Domestic_Carriers'],
                                  ['FOREIGN CARRIERS', 'TOTAL (FOREIGN CARRIERS)', 'Foreign_Carriers']]
                    for row_limit in row_limits:

                        df_new = pd.DataFrame()
                        for row in range(0, df.shape[0]):
                            for col in range(df.shape[1]):
                                if str(df.iat[row, col]).strip() == row_limit[0]:
                                    row_start = row
                        for row in range(0, df.shape[0]):
                            for col in range(df.shape[1]):
                                if str(df.iat[row, col]).strip() == row_limit[1]:
                                    row_end = row
                        #         break
                        for i in range(row_start+1, row_end):
                            for j in range(0, df.shape[1]):
                                # print(df.iat[i, j])
                                df_new.loc[i, j] = df.iat[i, j]

                        df_new.columns = new_column
                        df_new.drop(df_new.iloc[:, 0:1], axis=1, inplace=True)
                        df_new.dropna(axis=1, how='all', inplace=True)
                        df_new.to_csv('{}\Cleaned\{}_{}.csv'.format(
                            os.getcwd(), intial_excel_file_name[0], row_limit[2]), index=False)

                    #######################################
                    # time.sleep(10000)
                if 'TABLE 3.' in str(df.iat[k, s]):
                    print('FOUND IT 3')
                    #######################################
                    df_new = pd.DataFrame()
                    for row in range(0, df.shape[0]):
                        for col in range(df.shape[1]):
                            # print(row, col, df.iat[row, col])
                            if (df.iat[row, col] == 'Sl. No.') or (df.iat[row, col] == 'NAME OF THE COUNTRY') or (str(df.iat[row, col]).replace('\n', '') == 'PASSENGERS TO INDIA'):
                                column_row_location = row
                                row_start_limit = column_row_location+1
                            if (df.iat[row, col] == 'TOTAL'):
                                row_end_limit = row

                    print(column_row_location)
                    print(row_start_limit)
                    print(row_end_limit)

                    column_header = []
                    for i in df.loc[column_row_location]:
                        column_header.append(str(i).replace('\n', ''))

                    for row in range(row_start_limit, row_end_limit):
                        for col in range(0, df.shape[1]):
                            df_new.loc[row, col] = df.iat[row, col]

                    df_new.columns = column_header
                    df_new.drop(df_new.iloc[:, 0:1], axis=1, inplace=True)
                    df_new.to_csv('{}\Cleaned\{}_CountryWise_Traffic.csv'.format(
                        os.getcwd(), intial_excel_file_name[0]), index=False)

                    #######################################
                    pass
                if 'TABLE 4.' in str(df.iat[k, s]):
                    print('FOUND IT 4')
                    #######################################
                    column_row_location = []
                    column = []
                    row_start_limit = []
                    row_end_limit = []
                    for row in range(0, df.shape[0]):
                        for col in range(df.shape[1]):
                            # print(row, col, df.iat[row, col])
                            if (df.iat[row, col] == 'SL.No.') or (df.iat[row, col] == 'CITY 1') or (str(df.iat[row, col]).replace('\n', '') == 'PASSENGERS FROM CITY 2'):
                                if row not in column_row_location:
                                    column_row_location.append(row)
                                    row_start_limit.append(row+1)
                            if (df.iat[row, col] == 'SUB TOTAL'):
                                if row not in row_end_limit:
                                    row_end_limit.append(row)
                    print(column_row_location)
                    print(row_start_limit)
                    print(row_end_limit)
                    for u in df.loc[column_row_location[0]]:
                        column.append(str(u).replace('\n', ''))
                    for i in range(0, 2):
                        df_new = pd.DataFrame()
                        for row in range(row_start_limit[i], row_end_limit[i]):
                            for col in range(0, df.shape[1]):
                                df_new.loc[row, col] = df.iat[row, col]
                        df_new.columns = column
                        df_new.drop(df_new.iloc[:, 0:1], axis=1, inplace=True)
                        if i == 0:
                            file_name = 'Indian_International'
                            df_new.to_csv('{}\Cleaned\{}_CityWise_{}.csv'.format(
                                os.getcwd(), intial_excel_file_name[0], file_name), index=False)
                        else:
                            file_name = 'Entirely_International'
                            df_new.to_csv('{}\Cleaned\{}_CityWise_{}.csv'.format(
                                os.getcwd(), intial_excel_file_name[0], file_name), index=False)
                    #######################################
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
print('END')
