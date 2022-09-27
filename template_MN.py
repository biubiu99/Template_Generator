import numpy as np
from xlsxwriter import Workbook
from datetime import date
import pandas
import openpyxl
from openpyxl.styles import Font, PatternFill
from dateutil.relativedelta import relativedelta
from IPython.display import display

# Create new Excel file
current_date = date.today()
current_date_modified = current_date.strftime("%Y-%m-%d")

wb_MN_path = fr"C:\Users\Spareuser\Desktop\Code\template\Daily_templates\MN_11_{current_date_modified}_Valencia.xlsx"

wb_MN = Workbook(wb_MN_path)
wb_MN.close()

# Load SAP data
SAP_data_path = r"C:\Users\Spareuser\Desktop\Code\template\sample_data.xlsx"
data_loaded = pandas.read_excel(SAP_data_path)
# get 05 data
MN_data = data_loaded[data_loaded["Warehouse"] == "11-US-MN"]
# get data for chairs only
MN_data_chairs = MN_data[~MN_data["Item/Service Description"].isin([ "Extended Warranty", "Ipad Holder", "White Glove Service" , "Wine Holder", "Carbon Fiber Tray"])]

# get needed columns from original data
MN_data_modified = MN_data_chairs[
    ["Document Number", "Customer/Vendor Name", "Street", "City", "State", "Zip Code", "Email", "Phone Number",
     "Quantity", "Item/Service Description", "Shipping Method"]]

# rename columns to match the template
MN_data_modified = MN_data_modified.rename(
    {"Document Number": "Shipper Reference", "Customer/Vendor Name": "Consignee Name", "Street": "Consignee Address 1",
     "City": "Consignee City", "State": "Consignee State", "Zip Code": "Consignee Zip", "Email": "Consignee Contact",
     "Phone Number": "Consignee Phone", "Item/Service Description": "Dim1 Description", "Quantity": "Dim1 Piece Count"},
    axis=1)

MN_data_modified.loc[MN_data_modified[
                         ["Shipper Reference", "Consignee Name", "Consignee Address 1", "Consignee City",
                          "Consignee State", "Consignee Zip", "Consignee Contact", "Consignee Phone"]].duplicated(), [
                         "Shipper Reference", "Consignee Name", "Consignee Address 1", "Consignee City",
                         "Consignee State", "Consignee Zip", "Consignee Contact", "Consignee Phone"]] = np.NAN


MN_data_modified = MN_data_modified.reset_index(drop=False)


# add columns as per MN template
MN_data_modified = pandas.concat([MN_data_modified.iloc[:, :1],
                                  pandas.DataFrame('', columns=['Handling Station', 'Housebill', 'Control CustomerId',
                                                                'Payment Type', "Shipper Number", "Shipper Name",
                                                                "Shipper Address 1",
                                                                "Shipper Address 2", "Shipper City", "Shipper State",
                                                                "Shipper Zip", "Shipper Country", "Shipper Contact",
                                                                "Shipper Phone", "Shipper Reference Type"],
                                                   index=MN_data_modified.index),
                                  MN_data_modified.iloc[:, 1:]], axis=1)

MN_data_modified = pandas.concat([MN_data_modified.iloc[:, :17],
                                  pandas.DataFrame('', columns=['Pickup Date', 'Pickup Time', 'Pickup Close Time',
                                                                'Consignee Number'], index=MN_data_modified.index),
                                  MN_data_modified.iloc[:, 17:]], axis=1)




MN_data_modified.insert(loc=23, column="Consignee Address 2", value="")
MN_data_modified.insert(loc=27, column="Consignee Country", value="USA")


MN_data_modified = pandas.concat([MN_data_modified.iloc[:, :30],
                                  pandas.DataFrame('', columns=['Consignee Reference Type', 'Consignee Reference',
                                                                'Scheduled Delivery Date',
                                                                'Scheduled Delivery Time', "Destination Airport",
                                                                "Destination Airport Area", "BillTo Customer Id",
                                                                "Service", "Declared Value", "COD Amount", "FCCOD"],
                                                   index=MN_data_modified.index),
                                  MN_data_modified.iloc[:, 30:]], axis=1)



MN_data_modified.insert(loc=43, column="Dim1 Type", value="Carton")
MN_data_modified = pandas.concat([MN_data_modified.iloc[:, :44],
                                  pandas.DataFrame("",
                                                   columns=["Dim1 Weight", "Dim1 Length", "Dim1 Width", "Dim1 Height",
                                                            "Dim2 Piece Count", "Dim2 Type", "Dim2 Description",
                                                            "Dim2 Weight", "Dim2 Length", "Dim2 Width", "Dim2 Height",
                                                            "Dim3 Piece Count", "Dim3 Type", "Dim3 Description",
                                                            "Dim3 Weight", "Dim3 Length", "Dim3 Width", "Dim3 Height",
                                                            "Dim4 Piece Count", "Dim4 Type", "Dim4 Description",
                                                            "Dim4 Weight", "Dim4 Length", "Dim4 Width", "Dim4 Height",
                                                            "Dim5 Piece Count", "Dim5 Type", "Dim5 Description",
                                                            "Dim5 Weight", "Dim5 Length", "Dim5 Width", "Dim5 Height", "Dim Factor",
                                                            "Special Instructions", "Dangerous Goods", "TopLine Charge",
                                                            "Pickup agent", "IMPORT CODE", "Import Reference Type",
                                                            "Shipment Mode", "Business Unit", "POD Name", "POD Date",
                                                            "POD Time", "Product Code", "Call in", "Call In Name",
                                                            "Call In Phone", "Call In Fax", "Charge code1",
                                                            "Charge Amount1", "Charge Code2", "Charge Amount2",
                                                            "Charge Code3", "Charge Amount3", "Charge code4",
                                                            "Charge Amount4", "Charge Code5", "Charge Amount5",
                                                            "Shipment Status", "Delivered", "Ready for Audit",
                                                            "Ready For Invoicing", "Handling Type", "Handling Units",
                                                            "BilltoRefType", "BilltoRef", "consolNO", "Delivery Agent",
                                                            "Pickup Driver", "Delivery Driver", "Tractor", "Trailer",
                                                            "ETA Date", "ETA Time", "ApptDel", "Consignee Email",
                                                            "Scheduled Delivery Time Range", "INS/DVAL/LL Code",
                                                            "Delivery Instructions", "Pickup Instructions",
                                                            "Shipper Store", "Consignee Store", "Equipment type",
                                                            "Target Pay Min", "Target Pay Max",
                                                            "Notification Email for UPL", "Shipper Show Name",
                                                            "Shipper Venue Name", "Shipper Booth",
                                                            "Shipper Aux (Decorator)", "Consignee Show Name",
                                                            "Consignee Venue Name", "Consignee Booth",
                                                            "Consignee Aux (Decorator", "Vendor Service Code",
                                                            "Vendor No", "Vendor Ref No", "Vendor Amount",
                                                            "Vendor 204 Type", "Linehaul Vendor No", "Linehaul Ref No",
                                                            "Linehaul Amount", "Project Code", "Account Code", "Event Date",
                                                            "Event Time", "Shipper Email", "AccountMgr",
                                                            "Priority Code", "Origin Airport", "Origin Airport Area",
                                                            "Origin Miles", "Destination Miles", "Driver",
                                                            "Controlling Station", "Manifest#", "Manifest Seq",
                                                            "Move Type", "Dim1 Class", "Dim2 Class", "Dim3 Class",
                                                            "Dim4 Class", "Dim5 Class", "AgentPartner",
                                                            "OriginHouseBillNo", "RateCover", "AirlineName", "Flight",
                                                            "MAWB", "VolumnWt", "VolumeWt KILOS", "DepartmentCode",
                                                            "AirlineATA Code", "History"],
                                                   index=MN_data_modified.index),
                                  MN_data_modified.iloc[:, 44:]], axis=1)


# Fill in the blanks
MN_data_modified["BillTo Customer Id"] = MN_data_modified["BillTo Customer Id"].replace("", 1691)
MN_data_modified["Handling Station"] = MN_data_modified["Handling Station"].replace("", "AHD")
MN_data_modified["Control CustomerId"] = MN_data_modified["Control CustomerId"].replace("", 1691)
MN_data_modified["Payment Type"] = MN_data_modified["Payment Type"].replace("", 3)
MN_data_modified["Shipper Country"] = MN_data_modified["Shipper Country"].replace("", "US")
MN_data_modified["Shipper Name"] = MN_data_modified["Shipper Name"].replace("", "ACS LOGISTICS WAREHOUSE 2")
MN_data_modified["Shipper Address 1"] = MN_data_modified["Shipper Address 1"].replace("", "2360 PILOT KNOB ROAD ")
MN_data_modified["Shipper Address 2"] = MN_data_modified["Shipper Address 2"].replace("", "SUITE 600")
MN_data_modified["Shipper City"] = MN_data_modified["Shipper City"].replace("", "MENDOTA HEIGHTS")
MN_data_modified["Shipper State"] = MN_data_modified["Shipper State"].replace("", "MN")
MN_data_modified["Shipper Zip"] = MN_data_modified["Shipper Zip"].replace("", 55120)
MN_data_modified["Shipper Contact"] = MN_data_modified["Shipper Contact"].replace("", "Kwame Djadoo")
MN_data_modified["Shipper Phone"] = MN_data_modified["Shipper Phone"].replace("", "651-209-0037")
MN_data_modified["Shipper Reference Type"] = MN_data_modified["Shipper Reference Type"].replace("", "PO")
pickup_date = date.today() + relativedelta(days=+2)
MN_data_modified["Pickup Date"] = MN_data_modified["Pickup Date"].replace("", f"{pickup_date}")
MN_data_modified["Pickup Time"] = MN_data_modified["Pickup Time"].replace("", "7:00:00 AM")
MN_data_modified["Pickup Close Time"] = MN_data_modified["Pickup Close Time"].replace("", "3:00:00 PM")
MN_data_modified["Dim1 Weight"] = MN_data_modified["Dim1 Weight"].replace("", 125)
MN_data_modified["Dim1 Length"] = MN_data_modified["Dim1 Length"].replace("", 29)
MN_data_modified["Dim1 Width"] = MN_data_modified["Dim1 Width"].replace("", 31)
MN_data_modified["Dim1 Height"] = MN_data_modified["Dim1 Height"].replace("", 33)
MN_data_modified["Shipment Mode"] = MN_data_modified["Shipment Mode"].replace("", "DOMAIR")
MN_data_modified["Shipment Status"] = MN_data_modified["Shipment Status"].replace("", "WEB")

Index_number = MN_data_modified.loc[:, "index"]



# get the value for that cell
# PO_1 = MN_data_modified.loc[MN_data_modified["index"] == Index_number[0], "Shipper Reference"].item()

index = MN_data_modified.loc[MN_data_modified["index"] == Index_number[0], "Shipper Reference"].index.values[0]





for i in range(Index_number.index[-1] + 1):
    if np.isnan(MN_data_modified.at[i, "Shipper Reference"]):
        if MN_data_modified.at[index, "Dim2 Description"] == "":
            MN_data_modified.at[index, "Dim2 Description"] = MN_data_modified.at[i, "Dim1 Description"]
            MN_data_modified.at[index, "Dim2 Piece Count"] = MN_data_modified.at[i, "Dim1 Piece Count"]
            MN_data_modified = MN_data_modified.drop(labels=i, axis=0)
        else:
            MN_data_modified.at[index, "Dim3 Description"] = MN_data_modified.at[i, "Dim1 Description"]
            MN_data_modified.at[index, "Dim3 Piece Count"] = MN_data_modified.at[i, "Dim1 Piece Count"]
            MN_data_modified = MN_data_modified.drop(labels=i, axis=0)
    else:
        index = i


# drop unnecessary rows
MN_data_modified.drop("index", inplace=True, axis=1)
MN_data_modified.reset_index(inplace=True)






PO_number = MN_data_modified.loc[:, "Shipper Reference"]
for i in range(PO_number.index[-1] + 1):
    if MN_data_modified.loc[i, "Shipping Method"] == 1:
        MN_data_modified.at[i, "Service"] = "W4"
        MN_data_modified.at[
            i, "Special Instructions"] = "2 PERSON WHITE GLOVE DELIVERY W/LIGHT NO TOOL ASSEMBLY - Precision scheduling to 4-hour window, 30-minute pre-call, bring inside to client specified location, unpack, follow specified assembly instructions and remove dunnage"
    else:
        MN_data_modified.at[i, "Service"] = "TH"
        MN_data_modified.at[
            i, "Special Instructions"] = "Threshold Delivery - Precision scheduling to 4-hour window, 30-minute pre-call, Delivery of the product through the doorway of the dwelling – In the case of an apartment, delivery would be through the doorway of the apartment, not just the building. 15 minutes"



MN_data_modified.drop("Shipping Method", inplace=True, axis=1)
# display(MN_data_modified.to_string())




for i in range(PO_number.index[-1] + 1):
    if MN_data_modified.at[i, "Dim2 Description"] != "":
        MN_data_modified.at[i, "Dim2 Type"] = "Carton"
        MN_data_modified.at[i, "Dim2 Weight"] = 120
        MN_data_modified.at[i, "Dim2 Length"] = 29
        MN_data_modified.at[i, "Dim2 Width"] = 31
        MN_data_modified.at[i, "Dim2 Height"] = 33
    if MN_data_modified.at[i, "Dim3 Description"] != "":
        MN_data_modified.at[i, "Dim3 Type"] = "Carton"
        MN_data_modified.at[i, "Dim3 Weight"] = 120
        MN_data_modified.at[i, "Dim3 Length"] = 29
        MN_data_modified.at[i, "Dim3 Width"] = 31
        MN_data_modified.at[i, "Dim3 Height"] = 33

#drop index column
MN_data_modified.drop("index", inplace=True, axis=1)



MN_data_modified.drop(MN_data_modified.columns[19], axis=1)
with pandas.ExcelWriter(wb_MN_path, engine='xlsxwriter') as writer:
    # add empty row to the top of the worksheet
    MN_data_modified.to_excel(writer, index=False)
    # define worksheet and workbook
    workbook_11 = writer.book
    worksheet_11 = writer.sheets["Sheet1"]
