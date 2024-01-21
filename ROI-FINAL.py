import pandas as pd
from forex_python.converter import CurrencyRates
from datetime import datetime , timedelta
from dateutil.relativedelta import relativedelta
from forex_python.converter import CurrencyRates

################################################################
# Read both migration jira dump excel and Mapped customer excel

Salesforce = 'C:\\Users\\Incorta BI\\Box\\Incorta Backend\\Salesforce_raw.xlsx'
Mapped_Excel='C:\\Users\\Incorta BI\\Desktop\\Incorta Scripts\\Mapped_Excel_20_01_2024.csv'

Salesforce = pd.read_excel(Salesforce)
Mapped_Excel = pd.read_csv(Mapped_Excel)

################################################################


y=0
z=0
match_count=0
final_roi_excel=[] ## check testscript.py here u will identify all the final excel sheet Column names

################################################################


for iterrate_all_mapped_excel_rows in Mapped_Excel.iterrows():
    i=0
    ROI=0
    currency="USD"
    for iterrate_all_salesforce_rows in Salesforce.iterrows():

        Mapped_Customer_Name=str(Mapped_Excel.loc[z,'Sales Force Name']).lower()
        Salesforce_Customer_Name=str(Salesforce.loc[i,'Account Name']).lower()
        if((Mapped_Customer_Name==Salesforce_Customer_Name) and (Salesforce.loc[i,"Status"] in ("Activated","Fully Signed","Being Activated","Sent for VMware Countersignature","Signed by Customer","VMware Countersignature Requested"))and (Salesforce.loc[i,"PSO Practice"]!="Great Atlantic Migration") and (Salesforce.loc[i,"Total Agreement Value"] not in (0,1))): # if condition mapped customer= salesforce_customer AND status is what we want AND Agreement: created date >= mapped excel First date AND PSO PRACTISE NOT EQUAL "Great Atlantic Migration"

            try:

                Map_Date=Mapped_Excel.loc[z,"First Closed Date"] # getting value of first closed date that we got from previous python and was placed in 3rd column in mapped excel
                Map_Date=datetime.strptime(Map_Date, "%d-%m-%y") # make python understand that this is a time with current format d-m-yy
                New_map_date = Map_Date - relativedelta(days=400) # deducting a year from what is there in the mapped excel
                New_map_date= New_map_date.strftime("%Y/%m/%d") # changing the format to y/m/d so i can compare between two dates
                Salesforce_deal_date=Salesforce.loc[i,"Activated Date"] #getting the deal date from salesforce 25k sheet
                Salesforce_deal_date = Salesforce_deal_date.strftime("%Y/%m/%d") # changing the format to y/m/d so i can compare between two dates
                if(Salesforce_deal_date>=New_map_date):

                    match_count=match_count+1
                    if(Salesforce.loc[i,"Total Agreement Value Currency"]=="USD"):
                        DEAL= Salesforce.loc[i,"Total Agreement Value"]
                        ROI=ROI+DEAL

                    elif( (Salesforce.loc[i,"Total Agreement Value Currency"]!="USD") and (len(Salesforce.loc[i,"Total Agreement Value Currency"])>0) ):

                        c = CurrencyRates()
                        currency=Salesforce.loc[i,"Total Agreement Value Currency"] #can be deleted later
                        exchange_rate = c.get_rate(Salesforce.loc[i,"Total Agreement Value Currency"], "USD")
                        DEAL = (Salesforce.loc[i, "Total Agreement Value"]) * exchange_rate
                        ROI = ROI + DEAL

            except:
                do_nothing=1

        i=i+1
    ROI=round(ROI)
    Mapped_Excel.loc[z,'ROI']=ROI
    print(Mapped_Excel.loc[z, 'Sales Force Name']+ " ROI is " + str(ROI) + " Their local currency is " + currency)
    z=z+1
Mapped_Excel.to_csv("ROI Simple View.csv", index=False)
print(match_count)




