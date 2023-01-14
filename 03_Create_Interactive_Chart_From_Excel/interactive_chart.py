import pandas as pd
import plotly.express as px
import os

try:
    # load the data from the "Data" sheet of the "Financial_Data.xlsx" workbook
    data = pd.read_excel("Financial_Data.xlsx", sheet_name="Data")

    # check if 'Country' and 'Sales' columns exists in data
    if 'Country' not in data.columns or 'Sales' not in data.columns:
        raise ValueError("Columns are missing")
    # group the data by country and calculate the total sales for each country
    sales_by_country = data.groupby("Country")["Sales"].sum().reset_index()

    # create the bar chart
    fig = px.bar(sales_by_country, x="Country", y="Sales", title="Financial Data By Country",
                 labels={"Country": "Country", "Sales": "Total Sales"},
                 color_discrete_sequence=["#00008B"])

    # save the chart to the same directory as the workbook
    fig.write_html("Financial Data By Country.html")

    # display the chart
    fig.show()

except FileNotFoundError as e:
    print(f"{e} not found")
except ValueError as e:
    print(e)
