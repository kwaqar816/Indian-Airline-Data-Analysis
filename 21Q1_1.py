import pandas as pd

df = pd.read_excel(r'D:\Waqar\Power bi\Airline\Raw\21Q1_1.xlsx')
df_new = pd.DataFrame()


column_header = []
for i in df.loc[1]:
    column_header.append(str(i).replace('\n', ''))
# print(column_header)
# df.dropna(axis=0, how='any', inplace=True)
# print('Length of column is:- ', len(df.columns))
# print('Length of column is:- ', df.shape[1])
# for row in range(0, df.shape[0]):
#     for col in range(df.shape[1]):
#         print(df.iat[row , col])

print(df.loc[1])
for row in range(0, df.shape[0]):
    for col in range(df.shape[1]):
        if str(df.iat[row, col]).strip() == 'DOMESTIC CARRIERS':
            row_start = row
for row in range(0, df.shape[0]):
    for col in range(df.shape[1]):
        if str(df.iat[row, col]).strip() == 'TOTAL (DOMESTIC CARRIERS)':
            row_end = row
#         break
for i in range(row_start+1, row_end):
    for j in range(0, df.shape[1]):
        # print(df.iat[i, j])
        df_new.loc[i, j] = df.iat[i, j]


# print(df_new)
df_new.columns = column_header
# print(df_new)
df_new.drop(['SL.No.'], axis=1, inplace=True)
# df_new.set_index('NAME OF THE AIRLINE', inplace=True)
print(df_new)
# print(df_new.columns)
df_new.to_csv('new_file.csv', index=False)
# print('Row Start:- ', row_start)
# print('Col End:-', col_end)
# print('Row_End:- ', row_end)
