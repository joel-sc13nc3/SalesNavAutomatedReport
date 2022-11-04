import streamlit as st

@st.cache(suppress_st_warning=True)
def remove_values(df,string_of_values_to_remove):
    try:
        list_of_values=df.columns
        filtered_list = list(filter(lambda x: not x.startswith(string_of_values_to_remove), list_of_values))
        df=df[filtered_list]

    except ValueError:
        st.error("Please enter a valid input")

    return df

@st.cache(suppress_st_warning=True)
def cols_extraction(df,start,end):

    try:
        columns=df.columns
        selection = df[columns[start:end]]
        selection_cols = selection.columns
        selection_cols=map(str,selection_cols)
        selection_cols = list(filter(lambda x: not x.startswith("0"), selection_cols))
        selection=selection[selection_cols]
        old_cols=selection.columns
        new_cols_name=[w.replace('.1','') for w in old_cols]
        rename_cols=dict(zip(old_cols,new_cols_name))
        selection.rename(columns=rename_cols,inplace=True)
    except ValueError:
        st.error("Please check the sales sheet from analysis sheet excel")


    return selection


@st.cache(suppress_st_warning=True)
def remove_end_values(df,value_to_remove='.4'):
    try:
        columns = df.columns
        new_cols_name = [w.replace(value_to_remove, '') for w in columns]
        rename_cols = dict(zip(columns, new_cols_name))
        df.rename(columns=rename_cols, inplace=True)
    except ValueError:
        st.error("Please check dataframe")
    return df


@st.cache(suppress_st_warning=True)
def bu_selection_func(data):
    try:
        df=data.drop(columns=["KPI ID","KPI"])
        cols =df.columns.astype(str)
    except ValueError:
        st.error("Analysis Sheet does not have BU")

    return cols


def remove_if_is_in_list(list_to_keep,remove_list):
    return  [i for i in list_to_keep if i not in remove_list]

