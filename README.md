
# Excel Project 3 - Coffee Orders Dashboard ☕

### Project Overview
Hey! This is my third project using Excel, and I’m excited to share it. I took this on because I wanted to push my skills a bit further, especially in Power Query and with M Language. Also, after watching a [video on YouTube](https://www.youtube.com/watch?v=wP8NWRR0Fdg&list=WL&index=5) about data visualization and design, I felt inspired to make something visually engaging. The core of this project is similar to my [second project](https://github.com/adinramaadin/Excel_Project-2_Bike-Buyers-Dashboard-Project), where I created pivot tables and a dashboard. But this time, I decided to add a bit more spice and challenge myself to handle the entire ETL (Extract, Transform, Load) process efficiently using Power Query. 

The project was a mix of interface-based steps and M code. I challenged myself to handle most of the data transformation through M Language (instead of just using the Power Query interface), writing custom transformations with a "//" prefix to keep track. And no tutorials here! I picked up the logic and syntax patterns by experimenting directly in Power Query—learning through doing.

---

## 🎯 Project Goals and Workflow

### 1. **Data Cleaning and Transformation with Power Query**
  
Using Power Query, I applied several custom M Language transformations to make the data consistent, readable, and insightful. Below is a breakdown of the custom M code sections I created to achieve this.



```
let
    Source = orders,
    #"Merged Queries" = Table.NestedJoin(Source, {"Customer ID"}, customers, {"Customer ID"}, "customers", JoinKind.LeftOuter),
    #"Expanded customers" = Table.ExpandTableColumn(#"Merged Queries", "customers", {"Customer Name", "Email", "Country"}, {"Customer Name.1", "Email.1", "Country.1"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded customers", {"Product ID"}, products, {"Product ID"}, "products", JoinKind.LeftOuter),
    #"Expanded products" = Table.ExpandTableColumn(#"Merged Queries1", "products", {"Coffee Type", "Roast Type", "Size", "Unit Price"}, {"Coffee Type.1", "Roast Type.1", "Size.1", "Unit Price.1"}),
    // Remove original unexpanded columns
    #"Removed Columns" = Table.RemoveColumns(#"Expanded products", {"Customer Name", "Email", "Country", "Coffee Type", "Roast Type", "Size", "Unit Price", "Sales"}),
    // Remove .1 suffixes from column names
    #"Cleaned Column Names" = Table.TransformColumnNames(#"Removed Columns", each Text.Replace(_, ".1", "")),
    // Re-input "Sales" column as Unit Price * Quantity
    #"Added Sales" = Table.AddColumn(#"Cleaned Column Names", "Sales", each [Unit Price] * [Quantity], type number),
    // Replace the Coffee Type Abbreviation
    #"Replaced Coffee Types" = Table.ReplaceValue(#"Added Sales", 
    each [Coffee Type],
    each if [Coffee Type] = "Ara" then "Arabica" 
         else if [Coffee Type] = "Exc" then "Excelsa"
         else if [Coffee Type] = "Lib" then "Liberica"
         else if [Coffee Type] = "Rob" then "Robusta"
         else [Coffee Type], 
    Replacer.ReplaceValue, 
    {"Coffee Type"}
),
    // Replace the Roast Types Abbreviation
    #"Replaced Roast Types" = Table.ReplaceValue(#"Replaced Coffee Types", 
    each [Roast Type],
    each if [Roast Type] = "M" then "Medium"
        else if [Roast Type] = "L" then "Light"
        else if [Roast Type] = "D" then "Dark"
        else [Roast Type],
    Replacer.ReplaceValue,
    {"Roast Type"}
),
    // Change the Date format to "dd-mmm-yyy"
    #"Formatted Order Date" = Table.TransformColumns(#"Replaced Roast Types", {{"Order Date", each Date.ToText(_, "dd-MMM-yyyy"), type text}}),
    // Add " Kg" to the Size column by converting the decimal number to text and appending " Kg"
    #"Added Size with Kg" = Table.TransformColumns(#"Formatted Order Date", {{"Size", each Text.From(_, "en-US") & " Kg", type text}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Size with Kg",{{"Unit Price", type number}, {"Sales", type number}, {"Order ID", type text}, {"Order Date", type date}, {"Customer ID", type text}, {"Product ID", type text}, {"Quantity", Int64.Type}, {"Customer Name", type text}, {"Email", type text}, {"Country", type text}, {"Coffee Type", type text}, {"Roast Type", type text}, {"Size", type text}}),
    #"Removed Duplicates" = Table.Distinct(#"Changed Type", {"Order ID"})
in
    #"Removed Duplicates"
```
then after this (elaborate) i go straight to making pivot tables and chart. this one straightforward i just click some ribbon here and there


![image](https://github.com/user-attachments/assets/f896c24a-7cc8-4a7d-8a7d-c4edd64674b9)

![image](https://github.com/user-attachments/assets/19a3954c-11ef-4fbe-98ed-eb004aadaf4f)


then after this i designed my dashboard
![image](https://github.com/user-attachments/assets/052b1b67-62d8-4bbd-b5db-ebeb3ca957e5)
