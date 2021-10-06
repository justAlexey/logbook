import pandas as pd

apn1225 = pd.read_excel("1225.xlsx")
apn1328 = pd.read_excel("1328.xlsx")
apn1498 = pd.read_excel("1498.xlsx")

logbook = pd.DataFrame()

apn1328.sort_values(["CM Project No."], inplace=True)
apn1225.dropna(axis=1, inplace=True)
apn1225.dropna(axis=0, inplace=True)

for i in range(len(apn1225.index)):
    logbook["CM Project"] = apn1225.iloc[i]["CM Project"]
# print(logbook)
# print(apn1225.shape)
temp = apn1225[["End Date"]].loc[3]
print(temp)
