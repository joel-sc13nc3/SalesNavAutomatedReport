class Analysis_sheet():

    def __init__(self, kpi_names=None, company_data=None, company_name=None, bu_data=None, bu_name=None,
                 referenceset1=None, referenceset2=None,
                 referenceset3=None, referenceset4=None, rev_df=None, sales_employee_df=None, gross_margin_df=None,
                 channel_df=None, region_df=None,
                 template=None, rev_df_list=None, sales_employee_df_list=None, gross_margin_df_list=None,
                 channel_df_list=None, region_df_list=None,
                 referenceset_df_list=None, referece_set_list_names=None):
        self.__kpinames = kpi_names
        self.__company_data = company_data
        self.__company_name = company_name
        self.__bu_data = bu_data
        self.__bu_name = bu_name
        self.__referenceset1 = referenceset1
        self.__referenceset2 = referenceset2
        self.__referenceset3 = referenceset3
        self.__referenceset4 = referenceset4
        self.__rev_df = rev_df
        self.__sales_employee_df = sales_employee_df
        self.__gross_margin_df = gross_margin_df
        self.__channel_df = channel_df
        self.__region_df = region_df
        self.__template = template
        self.__sales_employee_df_list = sales_employee_df_list
        self.__gross_margin_df_list = gross_margin_df_list
        self.__channel_df_list = channel_df_list
        self.__region_df_list = region_df_list
        self.__rev_df_list = rev_df_list
        self.__referenceset_df_list = referenceset_df_list
        self.__referenceset_list_names = referece_set_list_names

        @property
        def kpi_names(self):
            return self.__kpi_names

        @kpi_names.setter
        def kpi_names(self, new_val):
            self.__kpi_names = new_val

        @property
        def company_data(self):
            return self.__company_data

        @company_data.setter
        def company_data(self, new_val):
            self.__company_data = new_val

        @property
        def company_name(self):
            return self.__company_name

        @company_name.setter
        def company_name(self, new_val):
            self.__company_name = new_val

        @property
        def bu_data(self):
            return self.__bu_data

        @bu_data.setter
        def bu_data(self, new_val):
            self.__bu_data = new_val

        @property
        def bu_name(self):
            return self.__bu_name

        @bu_name.setter
        def bu_name(self, new_val):
            self.__bu_name = new_val

        ##########################################################
        @property
        def rev_df(self):
            return self.__rev_df

        @rev_df.setter
        def rev_df(self, new_val):
            self.__rev_df = new_val

        @property
        def rev_df_list(self):
            return self.__rev_df

        @rev_df_list.setter
        def rev_df_list(self, new_val):
            self.__rev_df_list = new_val

        @property
        def sales_employee_df(self):
            return self.__sales_employee_df

        @rev_df.setter
        def sales_employee_df(self, new_val):
            self.__sales_employee_df = new_val

        @property
        def gross_margin_df(self):
            return self.__gross_margin_df

        @gross_margin_df.setter
        def gross_margin_df(self, new_val):
            self.__gross_margin_df = new_val

        @property
        def channel_df(self):
            return self.__channel_df

        @channel_df.setter
        def channel_df(self, new_val):
            self.__channel_df = new_val

        @property
        def region_df(self):
            return self.__region_df

        @region_df.setter
        def region_df(self, new_val):
            self.__region_df = new_val

        @property
        def template(self):
            return self.__template

        @template.setter
        def template(self, new_val):
            self.__template = new_val

            ##########################################################

            ##########################################################
        @property
        def rev_df_list(self):
            return self.__rev_df_list

        @rev_df_list.setter
        def rev_df_list(self, new_val):
            self.__rev_df_list = new_val

        @property
        def sales_employee_df_list(self):
            return self.__sales_employee_df_list

        @rev_df_list.setter
        def sales_employee_df_list(self, new_val):
            self.__sales_employee_df_list = new_val

        @property
        def gross_margin_df_list(self):
            return self.__gross_margin_df_list

        @gross_margin_df_list.setter
        def gross_margin_df_list(self, new_val):
            self.__gross_margin_df_list = new_val

        @property
        def channel_df_list(self):
            return self.__channel_df_list

        @channel_df_list.setter
        def channel_df_list(self, new_val):
            self.__channel_df_list = new_val

        @property
        def region_df_list(self):
            return self.__region_df_list

        @region_df_list.setter
        def region_df_list(self, new_val):
            self.__region_df_list = new_val

            ##########################################################

        @property
        def referenceset_df_list(self):
            return self.__referenceset_df_list

        @referenceset_df_list.setter
        def referenceset_df_list(self, new_val):
            self.__referenceset_df_list = new_val

        @property
        def referece_set_list_names(self):
            return self.__referece_set_list_names

        @referece_set_list_names.setter
        def referece_set_list_names(self, new_val):
            self.__referece_set_list_names = new_val