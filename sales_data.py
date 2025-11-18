import pandas as pd
import matplotlib.pyplot as plt
df = pd.read_excel("sales_data.xlsx")
pivot = pd.pivot_table(df, 
                       values="Total Sales", 
                       index="Product", 
                       columns="Region", 
                       aggfunc="sum", 
                       fill_value=0)
print("\nPivot Table:\n", pivot)
sales_by_product = df.groupby("Product")["Total Sales"].sum()
sales_by_product.plot(kind="bar", title="Total Sales by Product")
plt.ylabel("Sales")
plt.savefig("sales_by_product.png")   
plt.close()
sales_by_date = df.groupby("Date")["Total Sales"].sum()
sales_by_date.plot(kind="line", marker="o", title="Sales Trend Over Time")
plt.ylabel("Sales")
plt.savefig("sales_trend.png")       
plt.close()
pivot.to_excel("pivot_output.xlsx", sheet_name="PivotTable")