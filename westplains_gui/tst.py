import pandas as pd

df_dict1 = {'a':[1,2,3], 'b':[4,5,6]}
df1 = pd.DataFrame(df_dict1)
print(df1)
df_dict2 = {'a':[1,2,3,4], 'b':[4,5,6,7]}
df2 = pd.DataFrame(df_dict2)
print(df2)
df = df1.combine_first(df2)
print(df)
