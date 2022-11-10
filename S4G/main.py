import pandas as pd
import numpy as np
import functions as func

products_relation=pd.read_excel("Files/Products_relation.xlsx",sheet_name="Product_relation")
products_relation_family=pd.read_excel("Files/Products_relation.xlsx",sheet_name="Product_family")

opptyLineItem=pd.read_csv("Files/OpportunityLineItem.csv")
products2=pd.read_csv("Files/Product2.csv")
SBQQ=pd.read_csv("Files/SBQQ__QuoteLine__c.csv")

prod_code_oldvalues=products_relation["ProductCode_OLD"]
prod_code_new_values=products_relation["ProductCode"]

prod_name_oldvalues=products_relation["Name_old"]
prod_name_newvalues=products_relation["Name"]
prod_names=dict(zip(prod_name_oldvalues,prod_name_newvalues))

productcodesdict=dict(zip(prod_code_oldvalues,prod_code_new_values))



productcode_df=func.replacecolumvalues(productcodesdict,products2)
productcode_df=func.replacecolumvalues(prod_names,productcode_df)
productcode_df=productcode_df.drop(columns=["ProductFamily__c"])
productcode_df=productcode_df.merge(products_relation_family)
productcode_df["ProductType__c"]="Saas"



opptyLineItem=func.replacecolumvalues(productcodesdict,opptyLineItem)
opptyLineItem=func.replacecolumvalues(prod_names,opptyLineItem)


SBQQ=func.replacecolumvalues(productcodesdict,SBQQ)
SBQQ=func.replacecolumvalues(prod_names,SBQQ)





productcode_df.to_csv("C:\\Users\\Joel Ramirez\\PycharmProjects\\S4G\\Output\\Products2.csv")

opptyLineItem.to_csv("C:\\Users\\Joel Ramirez\\PycharmProjects\\S4G\\Output\\OpportunityLineItem.csv")

SBQQ.to_csv("C:\\Users\\Joel Ramirez\\PycharmProjects\\S4G\\Output\\SBQQ__QuoteLine__c.csv")

