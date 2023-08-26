import pandas as pd
df = pd.read_excel(r'D:\Waqar\Power bi\Airline\Raw\21Q1_4.xlsx', header=None)


column_row_location = []
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
for i in range(0, 2):
    df_new = pd.DataFrame()
    for row in range(row_start_limit[i], row_end_limit[i]):
        for col in range(0, df.shape[1]):
            df_new.loc[row, col] = df.iat[row, col]
    df_new.columns = df.loc[column_row_location[i]]
    df_new.drop(df_new.iloc[:, 0:1], axis=1, inplace=True)
    print(df_new)
