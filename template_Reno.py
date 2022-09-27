import numpy as np
from xlsxwriter import Workbook
from datetime import date
import pandas
import openpyxl
from openpyxl.styles import Font
from IPython.display import display

# Create new Excel file
current_date = date.today()
current_date_modified = current_date.strftime("%Y-%m-%d")

wb13_path = fr"C:\Users\Spareuser\Desktop\Code\template\Daily_templates\Reno_13_{current_date_modified}_Valencia.xlsx"
wb13 = Workbook(wb13_path)
wb13.close()

# Load SAP data
SAP_data_path = r"C:\Users\Spareuser\Desktop\Code\template\sample_data.xlsx"
data_loaded = pandas.read_excel(SAP_data_path)

# get 13 data
Reno_data = data_loaded[data_loaded["Warehouse"] == "13-US NV"]
# get data for chairs only
Reno_data_chairs = Reno_data[~Reno_data["Item/Service Description"].isin([ "Extended Warranty", "Ipad Holder", "White Glove Service" , "Wine Holder", "Carbon Fiber Tray"])]


Reno_data_modified = Reno_data_chairs[
    ["Customer/Vendor Name", "Phone Number", "Email", "Street", "City", "State", "Zip Code",
     "Document Number", "Item/Service Description", "Quantity", "CC Notes", "Shipping Method"]]
# rename columns to match the template
Reno_data_modified = Reno_data_modified.rename(
    {"Customer/Vendor Name": "Name", "Phone Number": "Number", "Street": "Address", "Document Number": "PO #",
     "Item/Service Description": "SKU", "CC Notes": "Note"}, axis=1)
Reno_data_modified = Reno_data_modified.reset_index(drop=False)
# access "PO #" column
Index_number = Reno_data_modified.loc[:, "index"]
empty_rows = []

for i in range(Index_number.index[-1]):
    PO_1 = Reno_data_modified.loc[Reno_data_modified["index"] == Index_number[i], "PO #"].item()
    PO_1_index = Reno_data_modified.loc[Reno_data_modified["index"] == Index_number[i], "PO #"].index[-1]
    PO_2 = Reno_data_modified.loc[Reno_data_modified["index"] == Index_number[i + 1], "PO #"].item()

    if PO_1 is None:
        continue
    else:
        if PO_1 != PO_2:
            Reno_data_modified.loc[PO_1_index + 0.5] = [None, None, None, None, None, None,
                                                        None, None, None, None, None, None, None]
            Reno_data_modified = Reno_data_modified.sort_index().reset_index(drop=True)
            empty_rows.append(PO_1_index + 4)
Reno_data_modified.drop("index", inplace=True, axis=1)

# replace duplicate with Nan
Reno_data_modified.loc[Reno_data_modified[
                           ["Name", "Address", "Number", "Email", "City", "State", "Zip Code",
                            "PO #"]].duplicated(), ["Name", "Address", "Number", "Email", "City", "State",
                                                    "Zip Code", "PO #"]] = np.NAN

# add notes based on shipping method
Shipping_method = Reno_data_modified.loc[:, "Shipping Method"]
for i in range(Shipping_method.index[-1]):
    if Reno_data_modified.loc[i, "Shipping Method"] == 1:
        Reno_data_modified.at[i, "Note"] = "white glove service needed"
    elif Reno_data_modified.loc[i, "Shipping Method"] == 4:
        Reno_data_modified.at[i, "Note"] = "threshold service needed"
    elif Reno_data_modified.loc[i, "Shipping Method"] == 3:
        continue

# drop unnecessary rows
Reno_data_modified.drop("Shipping Method", inplace=True, axis=1)

with pandas.ExcelWriter(wb13_path, engine='xlsxwriter') as writer:
    # add empty row to the top of the worksheet
    Reno_data_modified.to_excel(writer, index=False, startrow=1)
    # define worksheet and workbook
    workbook_13 = writer.book
    worksheet_13 = writer.sheets["Sheet1"]

    # change font style
    cell_format = workbook_13.add_format()
    cell_format.set_font_name("Cambria")
    worksheet_13.set_column("A:L", None, cell_format)
    # merge cells in first row
    merge_format = workbook_13.add_format({
        'bold': 1,
        'border': 0,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#FFFF00'})
    worksheet_13.merge_range("A1:L1", "Elunevision - Reno", merge_format)
    for i in empty_rows:
        worksheet_13.merge_range(f"A{i}:L{i}", "", merge_format)
