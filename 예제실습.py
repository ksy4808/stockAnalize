import win32com.client
instCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
print(instCpCybos.IsConnect)
instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
print(instCpStockCode.GetCount())
print(instCpStockCode.GetData(1, 0))