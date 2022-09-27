import numpy as np
from xlsxwriter import Workbook
from datetime import date
import pandas
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import string
from dateutil.relativedelta import relativedelta


#Buffalo template
# Create new Excel file
current_date = date.today()
current_date_modified = current_date.strftime("%Y-%m-%d")

wb05_path = fr"C:\Users\Spareuser\Desktop\Code\template\Daily_templates\Buffalo_05_{current_date_modified}_Valencia.xlsx"
wb05 = Workbook(wb05_path)
wb05.close()

# Load SAP data
SAP_data_path = r"C:\Users\Spareuser\Desktop\Code\template\sample_data.xlsx"
data_loaded = pandas.read_excel(SAP_data_path)

# get 05 data
Buffalo_data = data_loaded[data_loaded["Warehouse"] == "05-US"]
# get data for chairs only
Buffalo_data_chairs = Buffalo_data[~Buffalo_data["Item/Service Description"].isin(
    ["Extended Warranty", "Ipad Holder", "White Glove Service", "Wine Holder", "Carbon Fiber Tray",
     "Accessory Package"])]

Buffalo_data_modified = Buffalo_data_chairs[
    ["Customer/Vendor Name", "Phone Number", "Email", "Street", "City", "State", "Zip Code",
     "Document Number", "Item/Service Description", "Quantity", "CC Notes", "Shipping Method"]]
# rename columns to match the template
Buffalo_data_modified = Buffalo_data_modified.rename(
    {"Customer/Vendor Name": "Name", "Phone Number": "Number", "Street": "Address", "Document Number": "PO #",
     "Item/Service Description": "SKU", "CC Notes": "Note"}, axis=1)
Buffalo_data_modified = Buffalo_data_modified.reset_index(drop=False)
# access "PO #" column
Index_number_Buffalo= Buffalo_data_modified.loc[:, "index"]
empty_rows_05 = []

for i in range(Index_number_Buffalo.index[-1]):
    PO_1 = Buffalo_data_modified.loc[Buffalo_data_modified["index"] == Index_number_Buffalo[i], "PO #"].item()
    PO_1_index = Buffalo_data_modified.loc[Buffalo_data_modified["index"] == Index_number_Buffalo[i], "PO #"].index[-1]
    PO_2 = Buffalo_data_modified.loc[Buffalo_data_modified["index"] == Index_number_Buffalo[i + 1], "PO #"].item()

    if PO_1 is None:
        continue
    else:
        if PO_1 != PO_2:
            Buffalo_data_modified.loc[PO_1_index + 0.5] = [None, None, None, None, None, None,
                                                           None, None, None, None, None, None, None]
            Buffalo_data_modified = Buffalo_data_modified.sort_index().reset_index(drop=True)
            empty_rows_05.append(PO_1_index + 4)
Buffalo_data_modified.drop("index", inplace=True, axis=1)


# replace duplicate with Nan
Buffalo_data_modified.loc[Buffalo_data_modified[
                              ["Name", "Address", "Number", "Email", "City", "State", "Zip Code",
                               "PO #", "Note", "Shipping Method"]].duplicated(), ["Name", "Address", "Number", "Email",
                                                                                  "City", "State",
                                                                                  "Zip Code", "PO #", "Note",
                                                                                  "Shipping Method"]] = np.NAN

# add notes based on shipping method
Shipping_method_05 = Buffalo_data_modified.loc[:, "Shipping Method"]
for i in range(Shipping_method_05.index[-1]):
    if Buffalo_data_modified.loc[i, "Shipping Method"] == 1:
        Buffalo_data_modified.at[i, "Note"] = "white glove service needed"
    elif Buffalo_data_modified.loc[i, "Shipping Method"] == 4:
        Buffalo_data_modified.at[i, "Note"] = "threshold service needed"
    elif Buffalo_data_modified.loc[i, "Shipping Method"] == 3:
        continue

# drop unnecessary rows
Buffalo_data_modified.drop("Shipping Method", inplace=True, axis=1)

with pandas.ExcelWriter(wb05_path, engine='xlsxwriter') as writer:
    # add empty row to the top of the worksheet
    Buffalo_data_modified.to_excel(writer, index=False, startrow=1)
    # define worksheet and workbook
    workbook_05 = writer.book
    worksheet_05 = writer.sheets["Sheet1"]

    # change font style
    cell_format = workbook_05.add_format()
    cell_format.set_font_name("Cambria")
    worksheet_05.set_column("A:L", None, cell_format)
    # merge cells in first row
    merge_format = workbook_05.add_format({
        'bold': 1,
        'border': 0,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#FFCC99'})
    worksheet_05.merge_range("A1:L1", "Elunevision - Buffalo", merge_format)
    for i in empty_rows_05:
        worksheet_05.merge_range(f"A{i}:L{i}", "", merge_format)




