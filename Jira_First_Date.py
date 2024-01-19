import pandas as pd
from datetime import datetime

################################################################
# Read both migration jira dump excel and Mapped customer excel
Migration_Jira_Dump = 'C:\\Users\\Incorta BI\\Box\\Incorta Backend\\Migration Jira Dump.csv'
#Mapped_Customer='C:\\Users\\PowerBI\\OneDrive - VMware, Inc\\Desktop\\Confluence\\Mapped_Customers_final.xlsx'
Mapped_Customer='C:\\Users\\Incorta BI\\Box\\Incorta Backend\\Mapped_Customers_final.xlsx'
Migration_Jira_Dump = pd.read_csv(Migration_Jira_Dump)
Mapped_Customer = pd.read_excel(Mapped_Customer)

#test=Mapped_Customer.loc[0,'Jira Name']
#print(test)


################################################################
Total_Mapped_Jira_Customers=Mapped_Customer['Jira Name'] # This is a list of all Customer names under "Jira Name" in the mapped_customer_final excel
Total_Jira_Dump_Customers=Migration_Jira_Dump['Custom field (Customer)'] # This is a list of all Customer names under "Customer" in Migration Jira Dump
################################################################
z = 0 # This counter is used to update the "First Closed Date " value in Mapped_customer_final (it is syncing with the first for loop)

for Jira_Mapped_Customer in Total_Mapped_Jira_Customers: # This will go through each Jira Customer in the Mapped Excel


        i = 0 # This counter is used to update a list that carries all closed dates of a customer so that we can get the earliest closed date (it is syncing with the second for loop)
        closed_date = []
        for Jira_Dump_Customer in Total_Jira_Dump_Customers: #2nd for loop that goes through all jira dump customers
            if ((Jira_Mapped_Customer==Jira_Dump_Customer) and ((Migration_Jira_Dump.loc[i,'Status']=='Closure') or (Migration_Jira_Dump.loc[i,'Status']=='Completed')))  : # compare between current Jira Mapped customer in the first for loop and current jira dump customer in the second for loop
                date=Migration_Jira_Dump.loc[i,'Custom field (Close Date)']
                #print(Migration_Jira_Dump.loc[i,'Custom field (Product)'])
                date=datetime.strptime(date,"%d/%b/%y %I:%M %p")
                date=date.strftime("%y/%m/%d")

                #print(date)
                closed_date.append(date)
                #print(closed_date)
                #print(Migration_Jira_Dump.loc[i, 'Custom field (Close Date)'])
            i = i + 1
        try:

            minimum_closed_date= min(closed_date)
            minimum_closed_date = datetime.strptime(minimum_closed_date, "%y/%m/%d")

            minimum_closed_date=minimum_closed_date.strftime("%d-%m-%y")
            print(minimum_closed_date)
            #print(minimum_closed_date)
            Mapped_Customer.loc[z,'First Closed Date']= minimum_closed_date
        #print(Mapped_Customer.loc[z,'First Closed Date'])
        except:

            hello=56

        z=z+1

Mapped_Customer.to_csv('khalasna1877.csv', index=False)


