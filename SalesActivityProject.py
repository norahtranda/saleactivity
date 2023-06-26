#---------------------------------------------------------------
#-----------------------IMPORT LIBRARY--------------------------
#---------------------------------------------------------------

import pandas as pd
import numpy as np
import glob
import datetime
import plotly.express as px
import plotly.graph_objs as go
import matplotlib.pyplot as plt
import streamlit as st

#---------------------------------------------------------------
#------------------- SET UP LAYOUT OF WEB-----------------------
#---------------------------------------------------------------

#set layout web
st.set_page_config(page_title="Sales Actitivty Dashboard",
                   layout="wide")
#Set title of dashboard
st.title(":bar_chart: Sales Activity Dashboard")
st.markdown("---")
#set sidebar filter
st.sidebar.header("Please filter hear")

#---------------------------------------------------------------
#------------------- IMPORT DATA SOURCE-------------------------
#---------------------------------------------------------------

#import event data
folder_path = "D:/NGUYET/UC/01.Study/S1/04. COSC480_23S1/Project/data"
@st.cache_data
def get_data_event():
    """Read multiple excel files of event data"""
    files = glob.glob(folder_path + "/Event_*.xlsx")
    list_file = []
    for file in files:
        df = pd.read_excel(file)
        list_file.append(df)
    df_event = pd.concat(list_file, ignore_index=True)
    return df_event

#import retiler data
@st.cache_data
def get_data_retailer():
    """Read retailer data"""
    df_rtl = pd.read_excel(folder_path + "/Account.xlsx")
    return df_rtl


#import salesman data
@st.cache_data
def get_data_sales():
    """Read salesmen data"""
    df_sales = pd.read_excel(folder_path + "/Salesman.xlsx")
    return df_sales

df_event = get_data_event()
df_rtl = get_data_retailer()
df_sales = get_data_sales()


#---------------------------------------------------------------
#------------------- PROCESSING DATA ---------------------------
#---------------------------------------------------------------

@st.cache_data
def processing_data(df_event, df_rtl, df_sales):
    #Merge df_event and df_rtl to take RTL name and it's province
    df_activity = pd.merge(df_event, df_rtl[["Id", "Name", "Market Province"]],
                       left_on="AccountId", right_on="Id")

    #Merge with df_sale to take information of salesman
    df_activity = pd.merge(df_activity, df_sales[["User ID", "ASE Name", "Region"]],
                      left_on="OwnerId", right_on="User ID")

    #Rename column
    rename_columns = {"Id_x":"activity_id",
                  "Name":"retailer_name",
                  "Market Province":"retailer_province",
                  "ASE Name":"sales_name",
                  "Region":"sales_territory",
                  "ActivityDate":"activity_date",
                  "Opportunities_Status__c":"opportunity_status"}
    df_activity = df_activity.rename(columns=rename_columns)

    #Define sales activities: if event include 1.NP, 2.NP, 3.NP, 4.NP will be considered as NP activity and so on

    NP = [(df_activity["1.NP"] == 1) | (df_activity["2.NP"] == 1) | (df_activity["3.NP"] == 1) | (df_activity["4.NP"] == 1)]
    CS = [(df_activity["1.CS"] == 1) | (df_activity["2.CS"] == 1) | (df_activity["3.CS"] == 1) | (df_activity["4.CS"] == 1)]
    US = [(df_activity["1.US"] == 1) | (df_activity["2.US"] == 1) | (df_activity["3.US"] == 1) | (df_activity["4.US"] == 1)]
    SV = [(df_activity["5.SV1"] == 1) | (df_activity["5.SV2"] == 1) | (df_activity["5.SV3"] == 1) | (df_activity["5.SV4"] == 1) | (df_activity["5.SV5"] == 1)]
    SA = [(df_activity["5.SA1"] == 1) | (df_activity["5.SA2"] == 1) | (df_activity["5.SA3"] == 1) | (df_activity["5.SA4"] == 1) | (df_activity["5.SA5"] == 1) | (df_activity["5.SA6"] == 1) | (df_activity["6.SA1"] == 1) | (df_activity["6.SA2"] == 1) | (df_activity["6.SA3"] == 1) | (df_activity["6.SA4"] == 1) | (df_activity["7.SA1"] == 1) | (df_activity["7.SA2"] == 1) | (df_activity["7.SA3"] == 1) | (df_activity["7.SA4"] == 1) | (df_activity["7.SA5"] == 1)]
    OA = [(df_activity["OA1"] == 1) | (df_activity["OA2"] == 1) | (df_activity["OA3"] == 1) | (df_activity["OA4"] == 1)]

    conditions = {"NP": NP,
              "CS": CS,
              "US": US,
              "SV": SV,
              "SA": SA,
              "OA": OA}

    for key, value in conditions.items():
        df_activity[key] = np.select(value, "1")
    
    #take only data needed
    activity_data = df_activity[["activity_id", "activity_date", "sales_territory", "sales_name", "retailer_name", "retailer_province", "opportunity_status", "NP", "CS", "US", "SV", "SA", "OA"]]

    #change data type
    activity_data[["NP", "CS", "US", "SV", "SA", "OA"]] = activity_data[["NP", "CS", "US", "SV", "SA", "OA"]].apply(pd.to_numeric)

    #create column Year
    activity_data["year"] = activity_data["activity_date"].dt.year
    #create column Month
    activity_data["month"]  = activity_data["activity_date"].dt.month_name()
    #create column period
    activity_data["period"] = activity_data["activity_date"].dt.to_period("M")
    #create column sales activity
    activity_data["sales_activity"] = activity_data["NP"] + activity_data["CS"] + activity_data["US"] + activity_data["SV"]
    #create column admin activity
    activity_data["admin_activity"] = activity_data["SA"] + activity_data["OA"]
    #create colum total activity
    activity_data["total_activity"] = activity_data["sales_activity"] + activity_data["admin_activity"]
    #create column new opportunity
    activity_data["new_opportunity"] = activity_data["NP"] + activity_data["CS"] + activity_data["US"]
    return activity_data

#Loading data frame
activity_data = processing_data(df_event, df_rtl, df_sales)
activity_data["period"] = activity_data["period"].astype(str)

#---------------------------------------------------------------
#--------------- SETTING FILTERS AND METRICS--------------------
#---------------------------------------------------------------

#region filter
region_list = ["All"] + activity_data["sales_territory"].unique().tolist()
Region = st.sidebar.selectbox("Select Region", region_list)
#province filter
province_list = ["All"] + activity_data["retailer_province"].unique().tolist()
Province = st.sidebar.selectbox("Select Province", province_list)
#period filter
Period_list = ["All"] + sorted(activity_data["period"].unique().tolist())
Period = st.sidebar.selectbox("Select Period", Period_list)

#filter data frame
if Region != "All":
    activity_data = activity_data.query("sales_territory == @Region")
else:
    activity_data = activity_data
    
if Province != "All":
    activity_data = activity_data.query("retailer_province == @Province")
else:
    activity_data = activity_data

if Period != "All":
    activity_data = activity_data.query("period == @Period")
else:
    activity_data = activity_data

#setting key metrics
total_activity = int(activity_data["total_activity"].sum())
sale_activity = int(activity_data["sales_activity"].sum())
new_opportunity = int(activity_data["new_opportunity"].sum())

left_column, middle_column, right_column = st.columns(3)
with left_column:
    st.metric(":star: Total Activity:", value=f"{total_activity:,}")
with middle_column:
    st.metric(":star: Sale Activity:", value=f"{sale_activity:,}")
with right_column:
    st.metric(":star: New opportunity:", value=f"{new_opportunity:,}")


#---------------------------------------------------------------
#------------------- VISUALIZE DATA ----------------------------
#---------------------------------------------------------------

#------------ PLOT CHARTS CATEGORIZED BY SALESMANS -------------
def plot_stacked_bar(xs, ys, name, col):
    """function plot stacked bar chart"""
    return go.Bar(x=xs, y=ys,
                  name=name,
                  orientation="h",
                  marker=dict(color=col))

def stacked_bar_layout(title, x_name, y_name):
    """funtion setting layout"""
    return go.Layout(
                    title=title,
                    barmode="stack",
                    xaxis={"title": x_name},
                    yaxis={"title": y_name},
                    height=500,
                    width=500,
                    legend=dict(x=0.2, y=-0.2, orientation="h"))

#Plot total activity in number
df_total_act = activity_data.groupby("sales_name")[["sales_activity", "admin_activity", "total_activity"]].sum().reset_index()
total_act1 = plot_stacked_bar(df_total_act["sales_activity"], df_total_act["sales_name"], "Sales Activity", "#3D405B")
total_act2 = plot_stacked_bar(df_total_act["admin_activity"], df_total_act["sales_name"], "Admin Activity", "#E07A5F")
total_act = [total_act1, total_act2]
layout = stacked_bar_layout("Total Sales Activity", "Activity", "Sales Name")

figure_total_act = go.Figure(data=total_act, layout=layout)

#plot total activity in percent
sale_act_per = round(df_total_act["sales_activity"] / df_total_act["total_activity"] *100,0)
admin_act_per = round(df_total_act["admin_activity"] / df_total_act["total_activity"] *100,0)
total_act_per1 = plot_stacked_bar(sale_act_per, df_total_act["sales_name"], "Sales Activity", "#3D405B")
total_act_per2 = plot_stacked_bar(admin_act_per, df_total_act["sales_name"], "Admin Activity", "#E07A5F")
total_act_per = [total_act_per1, total_act_per2]
layout = stacked_bar_layout("Total Sales Activity in Percent", "Activity", "Sales Name")

figure_total_act_per = go.Figure(data=total_act_per, layout=layout)

#plot break down sales activity in number
df_sale_act = activity_data.groupby("sales_name")[["NP", "CS", "US", "SV", "sales_activity", "total_activity"]].sum().sort_values(by="total_activity", ascending=False).reset_index()
sale_act1 = plot_stacked_bar(df_sale_act["NP"], df_sale_act["sales_name"], "NP", "#3D405B")
sale_act2 = plot_stacked_bar(df_sale_act["CS"], df_sale_act["sales_name"], "CS", "#E07A5F")
sale_act3 = plot_stacked_bar(df_sale_act["US"], df_sale_act["sales_name"], "US", "#81B29A")
sale_act4 = plot_stacked_bar(df_sale_act["SV"], df_sale_act["sales_name"], "SV", "#F2CC8F")
sale_act = [sale_act1, sale_act2, sale_act3, sale_act4]
layout = stacked_bar_layout("Sales Activity Categozied by Activity Type", "Activity", "Sales Name")

figure_sale_act = go.Figure(data=sale_act, layout=layout)

#plot break down sales activity in percent
sale_np = round(df_sale_act["NP"]/df_sale_act["sales_activity"]*100,0)
sale_cs = round(df_sale_act["CS"]/df_sale_act["sales_activity"]*100,0)
sale_us = round(df_sale_act["US"]/df_sale_act["sales_activity"]*100,0)
sale_sv = round(df_sale_act["SV"]/df_sale_act["sales_activity"]*100,0)
sale_act_per1 = plot_stacked_bar(sale_np, df_sale_act["sales_name"], "NP", "#3D405B")
sale_act_per2 = plot_stacked_bar(sale_cs, df_sale_act["sales_name"], "CS", "#E07A5F")
sale_act_per3 = plot_stacked_bar(sale_us, df_sale_act["sales_name"], "US", "#81B29A")
sale_act_per4 = plot_stacked_bar(sale_sv, df_sale_act["sales_name"], "SV", "#F2CC8F")
sale_act_per = [sale_act_per1, sale_act_per2, sale_act_per3, sale_act_per4]
layout = stacked_bar_layout("Sales Activity Categozied by Activity Type in Percent", "Activity", "Sales Name")

figure_sale_act_per = go.Figure(data=sale_act_per, layout=layout)

#plot opportunity in number
df_sale_opp = activity_data.groupby("sales_name")[["NP", "CS", "US", "new_opportunity", "total_activity"]].sum().sort_values(by="total_activity", ascending=False).reset_index()
sale_opp1 = plot_stacked_bar(df_sale_opp["NP"], df_sale_opp["sales_name"], "NP", "#3D405B")
sale_opp2 = plot_stacked_bar(df_sale_opp["CS"], df_sale_opp["sales_name"], "CS", "#E07A5F")
sale_opp3 = plot_stacked_bar(df_sale_opp["US"], df_sale_opp["sales_name"], "US", "#81B29A")
sale_opp = [sale_opp1, sale_opp2, sale_opp3]
layout = stacked_bar_layout("Sales Opportunity Categozied by Activity Type", "Activity", "Sales Name")

figure_sale_opp = go.Figure(data=sale_opp, layout=layout)

#plot opportunity in percent
sale_opp_np = round(df_sale_opp["NP"]/df_sale_opp["new_opportunity"]*100,0)
sale_opp_cs = round(df_sale_opp["CS"]/df_sale_opp["new_opportunity"]*100,0)
sale_opp_us = round(df_sale_opp["US"]/df_sale_opp["new_opportunity"]*100,0)
sale_opp_per1 = plot_stacked_bar(sale_opp_np, df_sale_opp["sales_name"], "NP", "#3D405B")
sale_opp_per2 = plot_stacked_bar(sale_opp_cs, df_sale_opp["sales_name"], "CS", "#E07A5F")
sale_opp_per3 = plot_stacked_bar(sale_opp_us, df_sale_opp["sales_name"], "US", "#81B29A")
sale_opp_per = [sale_opp_per1, sale_opp_per2, sale_opp_per3]
layout = stacked_bar_layout("Sales Opportunity Categozied by Activity Type in Percent", "Activity", "Sales Name")

figure_sale_opp_per = go.Figure(data=sale_opp_per, layout=layout)


#------------ PLOT CHARTS COTEGORIZED BY MONTH ----------------

def plot_line(data, xs, ys, col):
    """funtion  plot line chart"""
    return px.line(data, xs, ys,
                   markers=True,
                   color_discrete_sequence=col)

def line_layout(data_plot, title):
    """function layout line chart"""
    data_plot.update_layout(title=title,
                  xaxis_title="Period",
                  yaxis_title="Number of Activity",
                  legend=dict(x=0.3, y=-0.2, orientation="h"),
                  plot_bgcolor="rgba(0,0,0,0)",
                  #xaxis=dict(showgrid=False),
                  #yaxis=dict(showgrid=False),
                  width=500,
                  height=500)

#plot total activity by month
month_act = activity_data.groupby("period")[["sales_activity", "admin_activity"]].sum().reset_index()
figure_month_total_act = plot_line(month_act,"period", ["sales_activity", "admin_activity"], ["#3D405B", "#E07A5F"])
line_layout(figure_month_total_act, "Monthly Activity")

#plot sales activity by month
sale_act_month = activity_data.groupby("period")[["NP", "CS", "US", "SV"]].sum().reset_index()
figure_sale_act_month = plot_line(sale_act_month, "period", ["NP", "CS", "US", "SV"], ["#3D405B", "#E07A5F", "#81B29A", "#F2CC8F"])
line_layout(figure_sale_act_month, "Monthly Sales Activity")

#plot new opportunity by month
sale_opp_month = activity_data.groupby("period")[["NP", "CS", "US"]].sum().reset_index()
figure_sale_opp_month = plot_line(sale_opp_month, "period", ["NP", "CS", "US"], ["#3D405B", "#E07A5F", "#81B29A"])
line_layout(figure_sale_opp_month, "Monthly Opportunities")


#--------------------- PLOT PIE CHART -------------------------

def plot_pie(labels, values, colors):
    """function plot pie chart"""
    data = go.Pie(labels=labels,
                  values=values,
                  hole=0.5,
                  marker=dict(colors=colors))
    return go.Figure(data = data)

def pie_layout(data_plot, title):
    """funtion layout pie chart"""
    data_plot.update_layout(title=title,
                            legend=dict(x=0.3, y=-0.2, orientation="h"),
                            plot_bgcolor="rgba(0,0,0,0)",
                            xaxis=dict(showgrid=False),
                            yaxis=dict(showgrid=False),
                            width=500,
                            height=500)

#pie chart of total activity
labels = ["sales_activity", "admin_activity"]
values = [activity_data["sales_activity"].sum(), activity_data["admin_activity"].sum()]
colors = ["#3D405B", "#E07A5F"]

figure_pie_total_act = plot_pie(labels, values, colors)
pie_layout(figure_pie_total_act, "Percent of Activity")

#pie chart of sale activity
labels = ["NP", "CS", "US", "SV"]
values = [activity_data["NP"].sum(), activity_data["CS"].sum(),activity_data["US"].sum(),activity_data["SV"].sum()]
colors = ["#3D405B", "#E07A5F", "#81B29A", "#F2CC8F"]

figure_pie_sale_act = plot_pie(labels, values, colors)
pie_layout(figure_pie_sale_act, "Percent of Sales Activity")

#pie chart of opportunity
labels = ["NP", "CS", "US"]
values = [activity_data["NP"].sum(), activity_data["CS"].sum(),activity_data["US"].sum()]
colors = ["#3D405B", "#E07A5F", "#81B29A", "#F2CC8F"]

figure_pie_opp_act = plot_pie(labels, values, colors)
pie_layout(figure_pie_opp_act, "Percent of Opportunity")
 
#--------------------- PLOT BOX CHART -------------------------

fig = go.Figure()
def plot_box(fig, data, columns, colors):
    """function plot box plot"""
    for column in columns:
        fig.add_trace(
            go.Box(x=data["sales_territory"],
                   y=data[column],
                   name=column,
                   marker=dict(color=colors[column])))
    return fig

def box_layout(data_plot, name_chart):
    """funtion layout of box plot"""
    data_plot.update_layout(title=f"Box plot of {name_chart}",
                            xaxis_title="Region",
                            yaxis_title="Number of Activity",
                            boxmode="group",# group together boxes of the different traces for each value of x
                            plot_bgcolor="rgba(0,0,0,0)",
                            #xaxis=dict(showgrid=False),
                            #yaxis=dict(showgrid=False),
                            width=1000)
 
#Box plot of sale activity
df_region = activity_data.groupby(["sales_territory", "period"])[["NP", "CS", "US", "SV", "sales_activity", "admin_activity"]].sum().reset_index()
figure_box_sale_act = go.Figure()
columns = ["NP", "CS", "US", "SV"]
colors = {"NP": "#3D405B",
          "CS": "#E07A5F",
          "US": "#81B29A",
          "SV": "#F2CC8F"}
figure_box_sale_act = plot_box(figure_box_sale_act, df_region, columns, colors)
box_layout(figure_box_sale_act, "Sales Activity")

#Box plot of all activity
figure_box_all_act = go.Figure()
columns = ["sales_activity", "admin_activity"]
colors = {"sales_activity": "#3D405B",
          "admin_activity": "#E07A5F"}

figure_box_all_act = plot_box(figure_box_all_act, df_region, columns, colors)
box_layout(figure_box_all_act, "All Activity")

#--------------------- SETTING BUTTONS -------------------------
#set layout button
left_button, right_button = st.columns(2)
with left_button:
    total_act_bt = st.button("Total Activity")
with right_button:
    sale_act_bt = st.button("Sales Activity")

if total_act_bt == True:
    st.plotly_chart(figure_box_all_act)
elif sale_act_bt == True:
    st.plotly_chart(figure_box_sale_act)
else:
    st.plotly_chart(figure_box_all_act)

st.markdown("---")

#--------------------- SHOW CHARTS ON WEB -----------------------

#show "Total Activity" charts
left_column1, right_column1 = st.columns(2)
with left_column1:
    st.plotly_chart(figure_month_total_act)
    st.plotly_chart(figure_total_act)
    st.markdown("---")
with right_column1:
    st.plotly_chart(figure_pie_total_act)
    st.plotly_chart(figure_total_act_per)
    st.markdown("---")


#show "Sales Activity" charts
with left_column1:
    st.plotly_chart(figure_sale_act_month)
    st.plotly_chart(figure_sale_act)
    st.markdown("---")
with right_column1:
    st.plotly_chart(figure_pie_sale_act)
    st.plotly_chart(figure_sale_act_per)
    st.markdown("---")


#show "Opportunity" charts
with left_column1:
    st.plotly_chart(figure_sale_opp_month)
    st.plotly_chart(figure_sale_opp)
with right_column1:
    st.plotly_chart(figure_pie_opp_act)
    st.plotly_chart(figure_sale_opp_per)
    



















