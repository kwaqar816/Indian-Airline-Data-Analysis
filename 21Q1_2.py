import pandas as pd

df = pd.read_excel(r'D:\Waqar\Power bi\Airline\Raw\21Q1_2.xlsx')

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
print(new_column)

row_limits = [['DOMESTIC CARRIERS', 'TOTAL (DOMESTIC CARRIERS)', 'Domestic data'],
              ['FOREIGN CARRIERS', 'TOTAL (FOREIGN CARRIERS)', 'International data']]
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
    df_new.to_csv('{}.csv'.format(row_limit[2]), index=False)
