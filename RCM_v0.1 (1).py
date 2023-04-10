
# In[ ]:
    
import os
os.getcwd()
os.chdir("C:/Users/Abhishek.Malan/Desktop/RCM")

# In[ ]:

import pandas as pd

#Loading excel file
file = "RCM_Workings - Copy.xlsx"
data = pd.ExcelFile(file)
print(data.sheet_names)


#Importing sheet names in dataframes
Client_Details = data.parse("Client") 
Vendor_Details = data.parse("Vendor_List")
Transactions = data.parse("Testing")
Directors = data.parse("Directors")


print(Client_Details.columns)
print(Vendor_Details.columns)
print(Transactions.columns)
print(Directors.columns)

# In[ ]:
    
#Is an Insurance agent?
Insurance_Agents = []
for index, row in Vendor_Details.iterrows():
    if "Yes" in str(row["Is an Insurance Agent"]):
        Insurance_Agents.append(row["Name"])
    else:
        pass

#Is an Recovery agent?
Recovery_Agents = []
for index, row in Vendor_Details.iterrows():
    if "Yes" in str(row["Is a Recovery Agent"]):
        Recovery_Agents.append(row["Name"])
    else:
        pass


# can add for lottery/music/.....

#Transactions.isnull().sum()
#Transactions.shape
#Transactions = Transactions.fillna("")
# In[ ]:

#Step1: Service provided by an Insurance Agent to Insurance Company
if "Yes" in str(Client_Details["Is engaged in Insurance Business?"]):
    mask1 = Transactions["Name"].isin(Insurance_Agents)
    if mask1.any():
        Transactions.loc[mask1, "RCM Reason"] = "Service provided by insurance agent"
        Transactions.loc[mask1, "Is RCM Applicable"] = "Yes"
    else:
        pass

# In[ ]:

#Step2: Service provided by a recovery agent to a BANK/FI or NBFC
if "Yes" in str(Client_Details["Is a Bank/NBFC/FI?"]):
    mask2 = Transactions["Name"].isin(Recovery_Agents)
    if mask2.any():
        Transactions.loc[mask2, "RCM Reason"] = "Service provided by recovery agent"
        Transactions.loc[mask2, "Is RCM Applicable"] = "Yes"
    else:
        pass

# In[ ]:

#Step3: Check and update for "Good mentioned u/n/4/2017"
mask3 = Transactions['Sac Codes'].astype(str).str.startswith(('801', '14049010', '2401', '5004', '5005', '5006'))
filtered_df = Transactions[mask3]
Transactions.loc[mask3, "RCM Reason"] = "Good mentioned u/n/4/2017"
Transactions.loc[mask3, "Is RCM Applicable"] = "Yes"

# In[ ]:
    
#Step4: Any service between a client & Director
mask4 = Transactions["Name"].isin(Directors["Name"])
Transactions.loc[mask4, "RCM Reason"] = "Service provided by Director"
Transactions.loc[mask4, "Is RCM Applicable"] = "Yes"

# In[ ]:

#Step5: Any service between Resident Service Recipient (Client) and a Non-Resident Service Provider

#Check for NR Vendors
NR_Vendor = Vendor_Details["Residential Status"] == "NR"
NR_filtered_df = Vendor_Details[NR_Vendor]

#Check for NR Vendors present in the Transactions dataframe
mask5 = Transactions["Name"].isin(NR_filtered_df["Name"])
#Check for Transactions with SAC Codes's first digits as 99
mask5a = Transactions["Sac Codes"].astype(str).str.startswith("99")
#Check for NR Vendor Transactions whose SAC Codes starts with 99
mask5b = mask5 & mask5a


if "R" in str(Client_Details["Residential Status"]):
    mask5b = mask5 & mask5a
    if mask5b.any():
        Transactions.loc[mask5b, "RCM Reason"] = "Service provided by NR"
        Transactions.loc[mask5b, "Is RCM Applicable"] = "Yes"
    else:
        pass

# In[ ]:

#Step6: GTA Services


#Filtering for SAC Code == 996791
GTA_Filter = Transactions["Sac Codes"] == 996791
GTA_Filtered_df = Transactions[GTA_Filter]

#Filtering for Govt Pan Holders
Govt_Filter = Vendor_Details["PAN"] .str[3] == "G"
Govt_Filtered_df = Vendor_Details[Govt_Filter]

#Filtering for Non-Govt Pan Holders
Not_Govt_Filter = Vendor_Details["PAN"] .str[3] != "G"
Not_Govt_Filtered_df = Vendor_Details[Not_Govt_Filter]

#Filtering for Keywords
Transactions["Invoice Description"] = Transactions["Invoice Description"].fillna("")
Transactions["Invoice Description"] = Transactions["Invoice Description"].str.lower()
Keyword_Filter = Transactions["Invoice Description"].str.lower().str.contains("Goods Transport Agency Services", case = False)
Keyword_Filtered_df = Transactions[Keyword_Filter]

#Updating if Sac code = 996791 and Vendor PAN Status not equals to G
mask6 = (Transactions["Name"].isin(Not_Govt_Filtered_df["Name"])) & (GTA_Filter)
Transactions.loc[mask6, "RCM Reason"] = "GTA Services"
Transactions.loc[mask6, "Is RCM Applicable"] = "Yes"

##Updating if Sac code = 996791 and contains keywords
##### NOT CLEAR WHETHER TO UPDATE IF PAN STATUS IS G AND CONTAINS KEYWORDS OR PAN STATUS NOT EQUALS G. 
##### BECAUSE THE SAC CODE 996791 IS FOR GTA SERVICES. SO MAYBE THIS STEP IS NOT NEEDED.

# In[ ]:

#Step7: Legal Representation Services by an advocate or firm

#Filtering for Sac Codes
Sac_Code_Filter = Transactions['Sac Codes'].astype(str).str.startswith(("998211","998212","998213","998214","998215","998216"))
Sac_Code_filtered_df = Transactions[Sac_Code_Filter]

#Filtering for Person or Firm Pan Holders
P_or_F_Filter = Vendor_Details["PAN"].str[3].isin(["P","F"])
P_or_F_Filtered_df = Vendor_Details[P_or_F_Filter]

#Update where Sac codes match with P or F PAN Status
mask7 = (Transactions["Name"].isin(Sac_Code_filtered_df["Name"])) & (P_or_F_Filter)
Transactions.loc[mask7, "RCM Reason"] = "Legal Services by Advocate or a Firm"
Transactions.loc[mask7, "Is RCM Applicable"] = "Yes"

# Update for services by arbitral tribunal
Keyword_Filter2 = Transactions["Invoice Description"].str.lower().str.contains("Services by Arbitral Tribunal", case=False)
Keyword_Filtered2_df = Transactions[Keyword_Filter2]

mask7a = (Transactions["Name"].isin(Keyword_Filtered2_df["Name"])) & (Sac_Code_Filter)
Transactions.loc[mask7a, "RCM Reason"] = "Services by Arbitral Tribunal"
Transactions.loc[mask7a, "Is RCM Applicable"] = "Yes"
    
##### WHAT ABOUT TRANSACTIONS WITH THE GIVEN SAC CODES, PAN STATUS NOT EQUAL TO "P" OR "F" AND NO KEYWORDS IN INVOICE DESCRIPTION??


# In[ ]:

#Step8: Sponsorship Services

# Filter for Sac Code = 998397
Sponsor_Filter = Transactions["Sac Codes"] == 998397
Sponsor_Filtered_df = Transactions[Sponsor_Filter]
Transactions.loc[Sponsor_Filter, "RCM Reason"] = "Sponsorship Services"
Transactions.loc[Sponsor_Filter, "Is RCM Applicable"] = "Yes"

#Filter for keywords
Keyword_Filter3 = Transactions["Invoice Description"].str.lower().str.contains("Sponsorship Services", case=False)
Keyword_Filtered3_df = Transactions[Keyword_Filter3]
mask8 = Transactions["Name"].isin(Keyword_Filtered3_df["Name"])
Transactions.loc[mask8, "RCM Reason"] = "Sponsorship Services"
Transactions.loc[mask8, "Is RCM Applicable"] = "Yes"

# In[ ]:

#Step9: Transportation of Goods by a vessel from a place outside India

#Filter for transactions with SAC code = 996521
Transport_Filter = Transactions["Sac Codes"] == 996521
Transport_Filtered_df = Transactions[Transport_Filter]


## 
## We have already filtered NR Vendors in Step5 :NR_Vendor = Vendor_Details["Residential Status"] == "NR"
##                                               NR_filtered_df = Vendor_Details[NR_Vendor]

#Check for Transactions with SAC Code = 996521 and NR Vendors
NR_Transport_Vendor = Transport_Filtered_df["Name"].isin(NR_filtered_df["Name"])
NR_Transport_Vendor_filtered_df = Transport_Filtered_df[NR_Transport_Vendor]

#Filter for Transactions with SAC code = 996521 and NR Vendor
mask9 = Transactions["Name"].isin(NR_Transport_Vendor_filtered_df["Name"])
Transactions.loc[mask9, "RCM Reason"] = "Transportation of Goods by Vessel from Outside India"
Transactions.loc[mask9, "Is RCM Applicable"] = "Yes"

#Filter for keyword
Keyword_Filter4 = Transactions["Invoice Description"].str.lower().str.contains("Transportation of Goods by Vessel from a place outside India", case=False)
Keyword_Filtered4_df = Transactions[Keyword_Filter4]
mask9a = Transactions["Name"].isin(Keyword_Filtered4_df["Name"])
Transactions.loc[mask9a, "RCM Reason"] = "Transportation of Goods by Vessel from Outside India"
Transactions.loc[mask9a, "Is RCM Applicable"] = "Yes"


#### WHAT ABOUT SAC CODE = 996521, RESIDENT VENDORS??

# In[ ]:

#Step10: Transactions with un-registered supplier

Unregistered_Vendors = Vendor_Details[Vendor_Details["GSTN"].isna()]

# Filter for transactions of un-registered GST vendors
mask10 = Transactions["Name"].isin(Unregistered_Vendors["Name"])
Transactions.loc[mask10, "RCM Reason"] = "Goods/Services received from un-registered service provider"
Transactions.loc[mask10, "Is RCM Applicable"] = "Yes"

# In[ ]:

#Step11: Services provided by Govt./UT or Local Authority

#Filter for Govt. or Local Authority Pan Holders
G_or_L_Filter = Vendor_Details[Vendor_Details["PAN"].str[3].isin(["G","L"])]

#Filter for transactions with SAC codes not in the list
not_in_list = [997211,997212,996811,996812,996411]
mask11 = ~Transactions['Sac Codes'].isin(not_in_list)
not_in_list_df = Transactions[mask11]

#Filter for govt/local authority PAN holders with SAC Codes not in the list
mask11a = not_in_list_df["Name"].isin(G_or_L_Filter["Name"])
Govt_Local_Sac_Code_df = not_in_list_df[mask11a]

#Filter for transactions govt/local authority PAN holders with SAC Codes not in the list
mask12a = Transactions["Name"].isin(Govt_Local_Sac_Code_df["Name"])
Transactions.loc[mask12a, "RCM Reason"] = "Service provided by Govt./UT or Local Authority"
Transactions.loc[mask12a, "Is RCM Applicable"] = "Yes"

# In[ ]:

#Step12: Residual Transactions

#To be used if Transactions.fillna("") is not used before
Transactions["Is RCM Applicable"] = Transactions["Is RCM Applicable"].fillna("No")

#To be used if Transactions.fillna("") is used before
#Transactions["Is RCM Applicable"] = Transactions["Is RCM Applicable"].replace(to_replace=" ", value="No")


# In[ ]:

# Create a new ExcelWriter object
writer = pd.ExcelWriter("all_data.xlsx")

# Write each dataframe to a separate sheet in the same Excel file
Client_Details.to_excel(writer, sheet_name="Client_Details", index=False)
Vendor_Details.to_excel(writer, sheet_name="Vendor_Details", index=False)
Transactions.to_excel(writer, sheet_name="Transactions", index=False)
Directors.to_excel(writer, sheet_name="Directors", index=False)

# Save and close the Excel file
writer.save()
              









