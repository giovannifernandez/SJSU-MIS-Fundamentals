"""The provided code analyzes the Same Day Shipping service of Office Solutions.
It examines the usage patterns of different customer groups and regions over time to identify trends.
This helps to gain insights into what factors influence the use of the service, such as customer demographics, location, and seasonal changes.
Based on this information, targeted strategies can be developed to make the service more popular.
Some of these strategies may include targeted marketing campaigns, which can lead to increased sales and customer engagement."""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
xl = pd.ExcelFile("/content/TableauSalesData.xlsx")
SalesData = xl.parse("Orders")

#Calculates the percentage of all orders that chose same day shipping
def SameDayProportion():
    sameDayOrders = SalesData[SalesData["Ship Mode"] == "Same Day"]
    totalOrders = len(SalesData)
    proportion = len(sameDayOrders) / totalOrders
    print(f"Proportion of Same Day Shipping Orders: {proportion:.2%}")

#Shows how popular same day shipping is among different types of customers
def DayBySegment():
    sameDayOrders = SalesData[SalesData["Ship Mode"] == "Same Day"]
    segmentUsage = sameDayOrders["Segment"].value_counts(normalize=True)
    print("Same Day Shipping Usage by Segment(%): ")
    print(segmentUsage)

#Identifies trends in the use of same day shipping across different months
def SeasonalTrends():
    SalesData["Order Date"] = pd.to_datetime(SalesData["Order Date"])
    sameDayOrders = SalesData[SalesData["Ship Mode"] == "Same Day"]
    monthlyTrends = sameDayOrders["Order Date"].dt.to_period("M").value_counts().sort_index()
    print("Monthly Trends in Same Day Shipping:")
    print(monthlyTrends)

#Calculates the average dollar amount spent on orders that use same day shipping
def AverageOrderValue():
    sameDayOrders = SalesData[SalesData["Ship Mode"] == "Same Day"]
    averageValue = sameDayOrders["Sales"].mean()
    print(f"Average Order Value for Same Day Shipping: ${averageValue:.2f}")

#Explores which geographical regions are most likely to choose same day shipping
def DayByRegion():
    sameDayOrders = SalesData[SalesData["Ship Mode"] == "Same Day"]
    regionDistribution = sameDayOrders["Region"].value_counts(normalize=True)
    print("Same Day Shipping Distribution by Region(%):")
    print(regionDistribution)

#Compares the profits from orders with same day shipping to those with other shipping methods
def DayProfitImpact():
    sameDayProfit = SalesData[SalesData["Ship Mode"] == "Same Day"]["Profit"].mean()
    otherModesProfit = SalesData[SalesData["Ship Mode"] != "Same Day"]["Profit"].mean()
    print(f"Average Profit - Same Day Shipping: ${sameDayProfit:.2f}")
    print(f"Average Profit - Other Shipping Modes: ${otherModesProfit:.2f}")

#Displays the  menu and ouputs users' inputs
def menu():
    print("Welcome to the Office Solutions Data Analytics System! To learn more about our sales data, please select a number from the list below:")
    print("\n Enter 1 to see Proportion of Orders with Same Day Shipping" +
          "\n Enter 2 to see Same Day Shipping Usage by Segment" +
          "\n Enter 3 to see Same Day Shipping Trends Over Time" +
          "\n Enter 4 to see Average Order Value with Same Day Shipping" +
          "\n Enter 5 to see Same Day Shipping by Region" +
          "\n Enter 6 to see Impact of Same Day Shipping on Profit" +
          "\n Enter 7 to quit\n")

    choice = input("Please enter your selection here: ")
    print("\n")

    if choice == "1":
        SameDayProportion()
    elif choice == "2":
        DayBySegment()
    elif choice == "3":
        SeasonalTrends()
    elif choice == "4":
        AverageOrderValue()
    elif choice == "5":
        DayByRegion()
    elif choice == "6":
        DayProfitImpact()
    elif choice == "7":
        print("Thank you for using the Office Solutions Data Analytics System.")
    else:
        print("You have entered an invalid option, please select from one of the options provided: ")
        menu()

menu()
