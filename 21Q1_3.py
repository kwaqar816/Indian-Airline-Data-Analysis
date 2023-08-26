import pandas as pd
df = pd.read_excel(r'D:\Waqar\Power bi\Airline\Raw\21Q1_3.xlsx', header=None)

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
df_new.to_csv('WaqarKhan.csv', index=False)
