import pandas as pd
import numpy as np
import streamlit as st

def referenceset_transformation(kpi_id,referencesetlist, reference_name_list):

    length=len(referencesetlist)
    df = []
    reference_list_name_new=[]
    x=0
    str_value = "#"
    for i in range(0,length):
        x=x+1
        y=str(x)
        w = str_value + y
        reference_list_name_new.append(w)

    for i in range(0,length):

        reference_set= referencesetlist[i]
        reference_set=reference_set[reference_set["KPI ID"]== kpi_id ]

        reference_set["item"] = reference_name_list[i]
        cols= reference_set.columns.to_list()
        cols.remove("KPI")
        cols.remove("KPI ID")
        cols.remove("item")

        for item in cols:
            reference_set.rename(columns={item: item+" "+reference_list_name_new[i]},inplace=True)

        item=reference_set["item"]

        reference_set.drop(columns=["KPI","KPI ID","item"],inplace=True)
        reference_set = reference_set.replace("-",np.nan)
        reference_set = reference_set.astype(float)

        reference_set["item"]=item
        reference_set.insert(0,"item",reference_set.pop("item"))
        df.append(reference_set)

    df=pd.concat(df)
    df=df.replace(np.nan,"NA")



    return df



def bar_chart_data_frame(kpi_id, company_data, company_name,bu_data, bu_name,reference_set_list,reference_name_list):

    if company_data.empty or company_name=="":
        bu_data = bu_data[bu_data["KPI ID"] == kpi_id]
        bu_data = bu_data.iloc[:, 2:]
        bu_data["item"] = bu_name
        bu_data.insert(0, 'item', bu_data.pop("item"))
        bu_data = bu_data.astype(float, errors="ignore")
        referenceset = referenceset_transformation(kpi_id, reference_set_list, reference_name_list)
        company_bu_reference = bu_data.append(referenceset)
        company_bu_reference = company_bu_reference.replace(np.nan, "NA")
        st.warning("Notice that there is no company data in the file")


    elif bu_data.empty or bu_name=="":
        company_data = company_data[company_data["KPI ID"] == kpi_id]
        company_data = company_data.iloc[:, [2]]
        company_data["item"] = company_name
        company_data.insert(0, 'item', company_data.pop("item"))
        company_data = company_data.astype(float, errors="ignore")
        referenceset = referenceset_transformation(kpi_id, reference_set_list, reference_name_list)
        company_bu_reference = company_data.append(referenceset)
        company_bu_reference = company_bu_reference.replace(np.nan, "NA")
        st.warning("Notice that there is no 'data cuts' (eg: bu,regions, etc) data in the file")

    elif reference_set_list==None or reference_name_list==None:

        company_data = company_data[company_data["KPI ID"] == kpi_id]
        company_data = company_data.iloc[:, [2]]
        company_data["item"] = company_name
        company_data.insert(0, 'item', company_data.pop("item"))
        company_data = company_data.astype(float, errors="ignore")
        bu_data = bu_data[bu_data["KPI ID"] == kpi_id]
        bu_data = bu_data.iloc[:, 2:]
        bu_data["item"] = bu_name
        bu_data.insert(0, 'item', bu_data.pop("item"))
        bu_data = bu_data.astype(float, errors="ignore")

        company_bu_reference = company_data.append(bu_data)
        company_bu_reference = company_bu_reference.replace(np.nan, "NA")
        st.warning("Notice that there is no reference set data data in the file")





    elif company_data.empty and bu_data.empty:
        company_bu_reference = referenceset_transformation(kpi_id, reference_set_list, reference_name_list)
        st.warning("Notice that there is no 'data cuts' (eg: bu,regions, etc) data in the file")




    else:


        company_data = company_data[company_data["KPI ID"] == kpi_id]
        company_data = company_data.iloc[:, [2]]
        company_data["item"] = company_name
        company_data.insert(0, 'item', company_data.pop("item"))
        company_data=company_data.astype(float,errors="ignore")
        bu_data = bu_data[bu_data["KPI ID"] == kpi_id]
        bu_data = bu_data.iloc[:, 2:]
        bu_data["item"] = bu_name
        bu_data.insert(0, 'item', bu_data.pop("item"))
        bu_data=bu_data.astype(float,errors="ignore")
        referenceset = referenceset_transformation(kpi_id, reference_set_list, reference_name_list)
        company_bu = company_data.append(bu_data)
        company_bu_reference = company_bu.append(referenceset)
        company_bu_reference=company_bu_reference.replace(np.nan,"NA")


    return company_bu_reference


def Analysis_sheet_reference_set_transformation(df,df1,rowname1,rowname2):
    df=pd.DataFrame(df)
    df=df.transpose()
    columnslist=df.columns
    newcols=["Bottom Quartile","Median","Top Quartile"]
    colsdict=dict(zip(columnslist,newcols))
    df=df.rename(columns=colsdict)
    df=df.reset_index(drop=True)
    df = pd.concat([df1, df])
    df=df.astype(float)
    df["item"]=[rowname1,rowname2]
    df.insert(0, "item", df.pop("item"))
    df=df.replace(np.nan,"NA")

    return df


@st.cache(suppress_st_warning=True)
def Analysis_sheet_reference_set_transformation_2(df,colsname=[]):
    cols=df.columns
    df=df.rename(columns={cols[0]:colsname[0],cols[1]:colsname[1]})
    df=df[df["Value"]!="nan"]
    df["Value"] = pd.to_numeric(df["Value"])
    return df




def emp_share(company_data,company_name,bu_data,reference_set_data,kpi_id):
    if company_data.empty:
        try:
            bu_data = bu_data[bu_data["KPI ID"].isin(kpi_id)]
            bu_data.set_index("KPI ID", drop=True, inplace=True)
            bu_data.drop(columns=["KPI"], inplace=True)
            bu_data = bu_data.astype("float", errors="ignore")

            reference_set_data = reference_set_data[reference_set_data["KPI ID"].isin(kpi_id)]
            reference_set_data.set_index("KPI ID", drop=True, inplace=True)
            reference_set_data.drop(columns=["KPI"], inplace=True)
            reference_set_data = reference_set_data["Median"]
            reference_set_data = reference_set_data.astype("float", errors="ignore")
            reference_set_data_sum = reference_set_data.sum()
            reference_set_data_values = reference_set_data.div(reference_set_data_sum)

            data = pd.concat([ bu_data, reference_set_data_values], axis=1)
            data.replace(np.nan, 0, inplace=True)
            data.reset_index(inplace=True)

            conditions = [
                (data['KPI ID'] == "KPI25"),  # Account owner FTE/sales FTE
                (data['KPI ID'] == "KPI26"),  # Sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI27"),  # Commercial sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI28"),  # Technical sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI29"),  # Sales operations FTE/sales FTE
                (data['KPI ID'] == "KPI30"),  # Transaction support FTE/sales FTE
                (data['KPI ID'] == "KPI31"),  # Sales administration and other sales support FTE/sales FTE
                (data['KPI ID'] == "KPI32")  # Sales management FTE/sales FTE
            ]

            values = ['Account Owner', 'Sales Enablement', 'Customer Facing', 'Customer Facing', 'Non Customer Facing',
                      'Non Customer Facing', 'Non Customer Facing', 'Sales Management']
            data['Categories'] = np.select(conditions, values)
            data = data.groupby(by=["Categories"], as_index=False, sort=False).sum()
        except:
            st.error("Please fill all the information")
            st.stop()
    elif company_data.empty and bu_data.empty:
        try:
            reference_set_data = reference_set_data[reference_set_data["KPI ID"].isin(kpi_id)]
            reference_set_data.set_index("KPI ID", drop=True, inplace=True)
            reference_set_data.drop(columns=["KPI"], inplace=True)
            reference_set_data = reference_set_data["Median"]
            reference_set_data = reference_set_data.astype("float", errors="ignore")
            reference_set_data_sum = reference_set_data.sum()
            reference_set_data_values = reference_set_data.div(reference_set_data_sum)

            data = reference_set_data_values
            data.replace(np.nan, 0, inplace=True)
            data.reset_index(inplace=True)

            conditions = [
                (data['KPI ID'] == "KPI25"),  # Account owner FTE/sales FTE
                (data['KPI ID'] == "KPI26"),  # Sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI27"),  # Commercial sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI28"),  # Technical sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI29"),  # Sales operations FTE/sales FTE
                (data['KPI ID'] == "KPI30"),  # Transaction support FTE/sales FTE
                (data['KPI ID'] == "KPI31"),  # Sales administration and other sales support FTE/sales FTE
                (data['KPI ID'] == "KPI32")  # Sales management FTE/sales FTE
            ]

            values = ['Account Owner', 'Sales Enablement', 'Customer Facing', 'Customer Facing', 'Non Customer Facing',
                      'Non Customer Facing', 'Non Customer Facing', 'Sales Management']
            data['Categories'] = np.select(conditions, values)
            data = data.groupby(by=["Categories"], as_index=False, sort=False).sum()
        except:
            st.error("Please fill all the information")
            st.stop()


    elif bu_data.empty:

        try:
            company_data = company_data[company_data["KPI ID"].isin(kpi_id)]
            company_data.set_index("KPI ID", drop=True, inplace=True)
            company_data = company_data.drop(columns=["KPI"])
            company_data = company_data.astype("float", errors="ignore")
            company_data.rename(columns={company_data.columns[0]: company_name}, inplace=True)


            reference_set_data = reference_set_data[reference_set_data["KPI ID"].isin(kpi_id)]
            reference_set_data.set_index("KPI ID", drop=True, inplace=True)
            reference_set_data.drop(columns=["KPI"], inplace=True)
            reference_set_data = reference_set_data["Median"]
            reference_set_data = reference_set_data.astype("float", errors="ignore")
            reference_set_data_sum = reference_set_data.sum()
            reference_set_data_values = reference_set_data.div(reference_set_data_sum)

            data = pd.concat([company_data, reference_set_data_values], axis=1)
            data.replace(np.nan, 0, inplace=True)
            data.reset_index(inplace=True)


            conditions = [
                (data['KPI ID'] == "KPI25"),  # Account owner FTE/sales FTE
                (data['KPI ID'] == "KPI26"),  # Sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI27"),  # Commercial sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI28"),  # Technical sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI29"),  # Sales operations FTE/sales FTE
                (data['KPI ID'] == "KPI30"),  # Transaction support FTE/sales FTE
                (data['KPI ID'] == "KPI31"),  # Sales administration and other sales support FTE/sales FTE
                (data['KPI ID'] == "KPI32")  # Sales management FTE/sales FTE
            ]

            values = ['Account Owner', 'Sales Enablement', 'Customer Facing', 'Customer Facing', 'Non Customer Facing',
                      'Non Customer Facing', 'Non Customer Facing', 'Sales Management']
            data['Categories'] = np.select(conditions, values)
            data = data.groupby(by=["Categories"], as_index=False, sort=False).sum()
        except:
            st.error("Please fill all the information")
            st.stop()
    elif reference_set_data.empty:

        try:
            company_data = company_data[company_data["KPI ID"].isin(kpi_id)]
            company_data.set_index("KPI ID", drop=True, inplace=True)
            company_data=company_data.drop(columns=["KPI"])
            company_data=company_data.astype("float",errors="ignore")
            company_data.rename(columns = {company_data.columns[0]:company_name}, inplace = True)


            bu_data = bu_data[bu_data["KPI ID"].isin(kpi_id)]
            bu_data.set_index("KPI ID", drop=True, inplace=True)
            bu_data.drop(columns=["KPI"], inplace=True)
            bu_data=bu_data.astype("float",errors="ignore")


            data = pd.concat([company_data, bu_data], axis=1)
            data.replace(np.nan, 0, inplace=True)
            data.reset_index(inplace=True)

            conditions = [
                (data['KPI ID'] == "KPI25"), #Account owner FTE/sales FTE
                (data['KPI ID'] == "KPI26"), #Sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI27"), #Commercial sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI28"), #Technical sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI29"), #Sales operations FTE/sales FTE
                (data['KPI ID'] == "KPI30"), #Transaction support FTE/sales FTE
                (data['KPI ID'] == "KPI31"), #Sales administration and other sales support FTE/sales FTE
                (data['KPI ID'] == "KPI32") #Sales management FTE/sales FTE
            ]

            values = ['Account Owner', 'Sales Enablement','Customer Facing', 'Customer Facing', 'Non Customer Facing',
                      'Non Customer Facing', 'Non Customer Facing', 'Sales Management']
            data['Categories'] = np.select(conditions, values)
            data = data.groupby(by=["Categories"], as_index=False, sort=False).sum()
        except:
            st.error("Please fill all the information")
            st.stop()

    else:

        try:
            company_data = company_data[company_data["KPI ID"].isin(kpi_id)]
            company_data.set_index("KPI ID", drop=True, inplace=True)
            company_data=company_data.drop(columns=["KPI"])
            company_data=company_data.astype("float",errors="ignore")
            company_data.rename(columns = {company_data.columns[0]:company_name}, inplace = True)


            bu_data = bu_data[bu_data["KPI ID"].isin(kpi_id)]
            bu_data.set_index("KPI ID", drop=True, inplace=True)
            bu_data.drop(columns=["KPI"], inplace=True)
            bu_data=bu_data.astype("float",errors="ignore")


            reference_set_data = reference_set_data[reference_set_data["KPI ID"].isin(kpi_id)]
            reference_set_data.set_index("KPI ID", drop=True, inplace=True)
            reference_set_data.drop(columns=["KPI"],inplace=True)
            reference_set_data=reference_set_data["Median"]
            reference_set_data=reference_set_data.astype("float",errors="ignore")
            reference_set_data_sum = reference_set_data.sum()
            reference_set_data_values = reference_set_data.div(reference_set_data_sum)

            data = pd.concat([company_data, bu_data,reference_set_data_values], axis=1)
            data.replace(np.nan, 0, inplace=True)
            data.reset_index(inplace=True)

            conditions = [
                (data['KPI ID'] == "KPI25"), #Account owner FTE/sales FTE
                (data['KPI ID'] == "KPI26"), #Sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI27"), #Commercial sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI28"), #Technical sales enablement FTE/sales FTE
                (data['KPI ID'] == "KPI29"), #Sales operations FTE/sales FTE
                (data['KPI ID'] == "KPI30"), #Transaction support FTE/sales FTE
                (data['KPI ID'] == "KPI31"), #Sales administration and other sales support FTE/sales FTE
                (data['KPI ID'] == "KPI32") #Sales management FTE/sales FTE
            ]

            values = ['Account Owner', 'Sales Enablement','Customer Facing', 'Customer Facing', 'Non Customer Facing',
                      'Non Customer Facing', 'Non Customer Facing', 'Sales Management']
            data['Categories'] = np.select(conditions, values)
            data = data.groupby(by=["Categories"], as_index=False, sort=False).sum()
        except:
            st.error("Please fill all the information")
            st.stop()

    return data