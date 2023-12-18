import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
xl = pd.ExcelFile("/content/TableauSalesData.xlsx")

#Combine sales and profit information for each type of product in every region
product_region_sales_profit = SalesData.groupby(["Product Name", "Region"]).agg({"Sales": "sum", "Profit": "sum"}).reset_index()
#Figure out how much profit can be made compared to how much it can be sold for each product and place
product_region_sales_profit["Margin"] = (product_region_sales_profit['Profit'] / product_region_sales_profit["Sales"]) * 100

#Setting a number to look at the top 10 items
n = 10
#Sort the products and regions by total sales from high to low
top_by_sales = product_region_sales_profit.sort_values(by="Sales", ascending=False).head(n)
#Sort the products and regions by profit margin from high to low
top_by_margin = product_region_sales_profit.sort_values(by="Margin", ascending=False).head(n)

#Canon imageCLASS 2200 Advanced Copier has the highest sales volume
plt.figure(figsize=(10, 6))
sns.barplot(x="Sales", y="Product Name", hue="Region", data=top_by_sales)
plt.title("Top Products and Regions by Sales")
plt.xlabel("Total Sales")
plt.ylabel("Product Name")
plt.xticks(rotation=45)
plt.show()

# Calculate the cost of the Canon imageCLASS 2200 Advanced Copier to determine the cost price per unit ($1,819.99)
canonCopier = SalesData[SalesData["Product Name"] == "Canon imageCLASS 2200 Advanced Copier"]
canonCopier["ProfitPerUnit"] = canonCopier["Profit"] / canonCopier["Quantity"]
canonCopier["SalesPricePerUnit"] = canonCopier["Sales"] / canonCopier["Quantity"]
canonCopier["CostPricePerUnit"] = canonCopier["SalesPricePerUnit"] - canonCopier["ProfitPerUnit"]
print(canonCopier[["Product Name", "CostPricePerUnit"]])

#Calculate the sales price per unit of the Canon imageCLASS 2200 Advanced Copier ($3,499.99)
canonCopier = SalesData[SalesData["Product Name"] == "Canon imageCLASS 2200 Advanced Copier"]
canonCopier["ProfitPerUnit"] = canonCopier["Profit"] / canonCopier["Quantity"]
canonCopier["SalesPricePerUnit"] = canonCopier["Sales"] / canonCopier["Quantity"]
canonCopier["CostPricePerUnit"] = canonCopier["SalesPricePerUnit"] - canonCopier["ProfitPerUnit"]
print(canonCopier[["Product Name", "SalesPricePerUnit"]])

#Calculate the profit of the Canon imageCLASS 2200 Advanced Copier ($25199.93)
canonCopier = SalesData[SalesData["Product Name"] == "Canon imageCLASS 2200 Advanced Copier"]
totalProfit = canonCopier["Profit"].sum()
print(f"Profit: ${totalProfit:.2f}")
