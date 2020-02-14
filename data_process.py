import pandas as pd
import datetime as dt

indf = pd.read_excel('data1.xlsx', index_col = False, encoding = "gb2312")
#outdf = pd.DataFrame(columns = indf.columns)
output_index_num = 8000
outdf = pd.DataFrame(pd.np.empty((output_index_num, len(indf.columns))) * pd.np.nan, columns = indf.columns)
times1 = dt.datetime.now()
outdf.loc[0] = indf.loc[0]
for outindex in range(0, output_index_num):
    # have some fun here
    for column in indf.columns:
        outdf.at[outindex,column] = indf.at[int(len(indf) * outindex / output_index_num),column]
times2 = dt.datetime.now()
print('Time spent: '+ str(times2-times1))