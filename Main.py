from ExcelStat import ExcelStat

EX = ExcelStat("C:\\Users\\Diego\\Desktop\\Haushaltslisten")
# DATA, _ , _ = EX.get_data_xls("Einkaufsliste ab dem 24.03.18.xlsx")
# print(DATA)
# EX.get_statistics(DATA)
# for FILE in EX.files:
#     DATA = EX.get_data_xls(FILE)
#     EX.get_statistics(DATA)
EX.get_data_all("-s")
