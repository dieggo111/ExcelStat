from ExcelStat import ExcelStat

EX = ExcelStat("C:\\Users\\Diego\\Desktop\\Haushaltslisten")
# DATA = EX.get_data("Einkaufsliste ab dem 20.02.17.xlsx")
# print(DATA)
# EX.get_statistics(DATA)
for FILE in EX.files:
    DATA = EX.get_data(FILE)
    EX.get_statistics(DATA)
