
# Excel Project 3 - Coffee Orders Dashboard ☕

## Introduction

Hey! This is my third project using Excel, and I’m excited to share it. I took this on because I wanted to push my skills a bit further, especially in Power Query and with M Language. Inspired by a [YouTube video](https://www.youtube.com/watch?v=wP8NWRR0Fdg&list=WL&index=5) about data visualization and design, I aimed to create a dashboard that’s not only functional but visually engaging.

The core of this project is similar to my [second project](https://github.com/adinramaadin/Excel_Project-2_Bike-Buyers-Dashboard-Project), where I created pivot tables and a dashboard. But this time, I decided to add a bit more spice and challenge myself of managing the full ETL (Extract, Transform, Load) process using Power Query. I went beyond interface-based steps, leveraging M Language for custom data transformations and honing my skills by experimenting directly in Power Query—no tutorials, just hands-on learning!

The project was a mix of interface-based steps and M code. I challenged myself to handle most of the data transformation through M Language (instead of just using the Power Query interface), writing custom transformations with a "//" prefix to keep track. And no tutorials here! I picked up the logic and syntax patterns by experimenting directly in Power Query—learning through doing.

### Dashboard File

You can view the final dashboard file here: [Coffee_Orders_Project.xlsx](https://github.com/adinramaadin/Excel_Project-3_Coffee-Orders-Data/blob/main/Coffee_Orders_Project.xlsx)

### Skills & Techniques Used

In this project, I utilized a variety of Excel skills, including:

- **Power Query** for efficient ETL processing and data transformation
- **M Language** to write custom transformations and calculations
- **Pivot Tables & Charts** for interactive data summaries
- **Dashboard Design** for a visually engaging interface, with careful attention to color schemes and layout for clarity and usability

### Data Source

The dataset used in this project is from [Mo Chen](https://github.com/mochen862/excel-project-coffee-sales) and includes three sheets:
- **Product:** Product ID, Coffee Type, Roast Type, Size, Unit Price, Price per Kg, Profit
- **Customer:** Customer ID, Customer Name, Email, Phone Number, Address, City, Country, Postcode, Loyalty Card status
- **Orders:** Order ID, Order Date, Customer ID, Product ID, Quantity

The data is also available here for reference: [coffeeOrdersData.xlsx](https://github.com/adinramaadin/Excel_Project-3_Coffee-Orders-Data/blob/main/Data/coffeeOrdersData.xlsx).
------

## From Data to Dashboard: Steps in Building the Coffee Orders Dashboard

In this section, I'll walk you through the key steps I took to transform raw data into a dynamic dashboard using Power Query’s interface tools in Excel.

### Step 1: Loading Data into Power Query
To start, I loaded all data from a single workbook containing three sheets: **Orders**, **Products**, and **Customers**. By importing these tables directly into Power Query, I was able to set up a seamless data connection for each sheet, preparing them for transformation and analysis.

### Step 2: Preparing the Orders Table for Merging
Before merging tables, I referenced the **Orders** table as a backup, ensuring I could experiment freely without affecting the original data. This gave me a copy to work with while keeping the initial dataset intact and available for comparison if needed.

### Step 3: Merging Data Across Tables
Using Power Query’s merge functionality, I combined relevant information from the **Customers** and **Products** tables with the **Orders** table. Specifically:
- I joined **Orders** with **Customers** to incorporate details like customer names and locations, enriching each order with customer-specific information.
- I merged **Orders** with **Products** to pull in details like product types, roast levels, and pricing. This allowed each order to be contextualized with the characteristics of the ordered items.

By structuring the data this way, I created a single, comprehensive view of each order, enriched with both customer and product attributes—setting the foundation for insightful analysis in the dashboard.

### Step 4: Cleaning and Transforming Data

With the merged data in place, I applied several transformations in Power Query to clean and prepare it for analysis. Here’s a breakdown of each transformation step by step:

```
// Remove original unexpanded columns
#"Removed Columns" = Table.RemoveColumns(
    #"Expanded products", 
    {"Customer Name", "Email", "Country", "Coffee Type", "Roast Type", "Size", "Unit Price", "Sales"}
),
```

After merging, some columns in the table contained the collumns that has `null` values. To streamline the dataset, I removed columns such as `Customer Name`, `Email`, `Country`, `Coffee Type`, `Roast Type`, `Size`, `Unit Price`, and `Sales`, keeping only the essential data.

```
// Remove .1 suffixes from column names
#"Cleaned Column Names" = Table.TransformColumnNames(
    #"Removed Columns", 
    each Text.Replace(_, ".1", "")
),
```

Some columns had a `.1` suffix due to duplication during merging. To keep the column names consistent and readable, I removed these suffixes using a simple transformation.

```
// Re-input "Sales" column as Unit Price * Quantity
#"Added Sales" = Table.AddColumn(
    #"Cleaned Column Names", 
    "Sales", 
    each [Unit Price] * [Quantity], 
    type number
),
```

Here, I created a new Sales column by calculating it as the product of Unit Price and Quantity. This calculation provided the total sales value for each order, which is essential for financial insights in the dashboard.

```
// Replace the Coffee Type Abbreviation
#"Replaced Coffee Types" = Table.ReplaceValue(
    #"Added Sales", 
    each [Coffee Type], 
    each 
        if [Coffee Type] = "Ara" then "Arabica" 
        else if [Coffee Type] = "Exc" then "Excelsa"
        else if [Coffee Type] = "Lib" then "Liberica"
        else if [Coffee Type] = "Rob" then "Robusta"
        else [Coffee Type], 
    Replacer.ReplaceValue, 
    {"Coffee Type"}
),
```

The `Coffee Type` column contained abbreviations (`Ara`, `Exc`, `Lib`, `Rob`) that weren’t easily understood. I replaced these with their full names—`Arabica`, `Excelsa`, `Liberica`, and `Robusta`—to improve clarity and readability.

```
// Replace the Roast Types Abbreviation
#"Replaced Roast Types" = Table.ReplaceValue(
    #"Replaced Coffee Types", 
    each [Roast Type], 
    each 
        if [Roast Type] = "M" then "Medium"
        else if [Roast Type] = "L" then "Light"
        else if [Roast Type] = "D" then "Dark"
        else [Roast Type],
    Replacer.ReplaceValue,
    {"Roast Type"}
),
```
Similarly, I replaced the abbreviations in the `Roast Typ`e column (`M`, `L`, `D`) with the full roast descriptions—`Medium`, `Light`, and `Dark`. This made the data more intuitive for dashboard users.

```
// Change the Date format to "dd-mmm-yyy"
#"Formatted Order Date" = Table.TransformColumns(
    #"Replaced Roast Types", 
    {{"Order Date", each Date.ToText(_, "dd-MMM-yyyy"), type text}}
),
```
For a more readable format, I converted the `Order Date` column to a string format (`dd`-`MMM`-`yyyy`), making it clearer and more consistent in reports.

```
// Add " Kg" to the Size column by converting the decimal number to text and appending " Kg"
#"Added Size with Kg" = Table.TransformColumns(
    #"Formatted Order Date", 
    {{"Size", each Text.From(_, "en-US") & " Kg", type text}}
),
```

The `Size` column, originally in numerical form, was converted to text and appended with `“ Kg”` to indicate weight. This made the size data more meaningful for the end user.

```
// Change data types to ensure correct representation
#"Changed Type" = Table.TransformColumnTypes(
    #"Added Size with Kg",
    {
        {"Unit Price", type number}, 
        {"Sales", type number}, 
        {"Order ID", type text}, 
        {"Order Date", type date}, 
        {"Customer ID", type text}, 
        {"Product ID", type text}, 
        {"Quantity", Int64.Type}, 
        {"Customer Name", type text}, 
        {"Email", type text}, 
        {"Country", type text}, 
        {"Coffee Type", type text}, 
        {"Roast Type", type text}, 
        {"Size", type text}
    }
),
```

This step sets the appropriate data types for each column, ensuring that numerical and date values are interpreted correctly in Excel and Power Query.

```
// Merge "customers" table to add "Loyalty Card" information based on "Customer ID"
#"Merged Queries2" = Table.NestedJoin(
    #"Changed Type", 
    {"Customer ID"}, 
    customers, 
    {"Customer ID"}, 
    "customers", 
    JoinKind.LeftOuter
),
```

In this step, I joined the `customers` table with the main data table using `Customer ID` as the matching key. This added a column called `customers` that contains related records, specifically to retrieve `Loyalty Card` data.

```
// Expand "customers" column to bring in the "Loyalty Card" data into the main table
#"Expanded customers1" = Table.ExpandTableColumn(
    #"Merged Queries2", 
    "customers", 
    {"Loyalty Card"}, 
    {"Loyalty Card"}
),
```

Here, I expanded the newly joined customers column, selecting only the Loyalty Card field. This step adds the Loyalty Card information as a new column in the main table, making it accessible for further analysis.

```
// Remove duplicate entries based on Order ID
#"Removed Duplicates" = Table.Distinct(
    #"Expanded customers1", 
    {"Order ID"}
```

Finally, this line removes any duplicate rows based on `Order ID`. It ensures that each order appears only once in the dataset, maintaining data accuracy and consistency.

## From Data to Insights: Pivot Tables, Charts, and Analysis

After completing the data cleaning and transformation process in Power Query, I moved on to building pivot tables and visualizations to present the data in a more understandable and actionable way. The goal was to derive meaningful insights and trends that could help guide decision-making.

### 1. Coffee Sales Trend Over Time (Line Chart)

<p align="center">
  
  <img src="https://github.com/user-attachments/assets/5268060e-ecdb-4a62-96a4-75c8f41d76bd" width="600"/>
  
</p>
<p align="center">
  **Coffee Type and Sales Analysis**
</p>

<p align="center">
  <img src="https://github.com/user-attachments/assets/f896c24a-7cc8-4a7d-8a7d-c4edd64674b9" width="600"/>
</p>
<p align="center">
  **Sales Data Pivot Table Overview**
</p>


**Why Line Chart?**
- **Tracking Trends:** The line chart effectively displays sales trends for each coffee type over time, helping to highlight seasonal fluctuations and year-on-year changes.
- **Avoiding Visual Overload:** I considered using a "line area chart" to show the sales trends, but the filled areas made the chart look overly cartoonish. By choosing a simple line chart, the data presentation remains clean and professional, consistent with **The Economist** design style that i try to recreate.

**Key Insights:**
- **Seasonality:** Higher sales in months like January and December suggest seasonal fluctuations.
- **Performance Variations:** Coffee types like Robusta and Arabica show consistent sales, while others like Liberica experience more volatility.

### 2. **Total Sales by Country (Bar Chart)**

<p align="center">
  <img src="https://github.com/user-attachments/assets/18b0e3e0-e1e4-422a-b35a-1123c5bc11e4" width="600"/>
</p>
<p align="center">
  **Sales by Country Chart**
</p>

<p align="center">
  <img src="https://github.com/user-attachments/assets/19a3954c-11ef-4fbe-98ed-eb004aadaf4f" width="600"/>
</p>
<p align="center">
  **Sales by Country Pivot Table Overview**
</p>

**Why Bar Chart?**
- **Comparing Categories:** The bar chart is ideal for comparing total sales across countries, providing a clear view of which regions are performing best.
- **Immediate Impact:** The horizontal bars make it easy to visually compare the contribution of each country to the overall sales, quickly highlighting the leaders (e.g., the United States).

**Key Insights:**
- **Leading Markets:** The United States significantly outperforms other countries in sales.
- **Focus on Growth:** Countries like the United Kingdom and Ireland show significant potential for growth based on current sales trends.

#### 3. **Sales by Sales Representative (Bar Chart)**

<p align="center">
  <img src="https://github.com/user-attachments/assets/0a4e09a8-8a6c-4ac6-b5ea-21ef49339fbf" width="600"/>
</p>
<p align="center">
  **Top 5 Customer Bar Chart**
</p>

<p align="center">
  <img src="https://github.com/user-attachments/assets/19a3954c-11ef-4fbe-98ed-eb004aadaf4f" width="600"/>
</p>
<p align="center">
  **Top 5 Customer Pivot Table Overview**
</p>


**Why Bar Chart?**
- **Individual Performance Comparison:** The bar chart compares sales performance across individual representatives, helping to identify top performers and areas needing improvement.
- **Quick Insights:** The visualization allows managers to quickly identify who is driving the most sales, aiding in performance reviews and incentive planning.

**Key Insights:**
- **Top Performers:** Sales representatives like Alexa Sizey and Allis Wilmore lead in sales.
- **Improvement Opportunities:** Representatives with lower sales should be assessed for training needs or additional support to boost performance.

### Visual Design Process
The visual design was a critical part of this project, focusing on creating a sleek, professional look that aligns with the minimalist aesthetic of **The Economist**.

#### Challenges:
- **Learning by Doing:** The design process took longer than expected as I experimented with different colors, layouts, and formats to achieve a professional look.
- **Color and Formatting Decisions:** I carefully selected colors to ensure contrast and readability, avoiding overly complex color schemes that could distract from the data.

#### Design Goals:
- **Clarity and Simplicity:** The aim was to create charts that were easy to read and interpret, avoiding unnecessary visual elements that might clutter the design.
- **Professional Aesthetic:** The final charts were designed to maintain a clean, minimalist look, focusing on clarity and ease of understanding, in line with professional standards.

## Dashboard Creation and Slicers

After creating the pivot tables and charts, I proceeded to build a dynamic dashboard to enable interactive data exploration. To make the data more accessible and easier to analyze, I added four slicers:

1. **Order Date**: This slicer allows the user to filter the data by the time period (e.g., by year or month). It is an essential component for identifying trends and making time-based comparisons.

2. **Roast Type**: The roast type slicer enables the user to narrow down the analysis to specific coffee types, such as Medium, Light, or Dark roast, which is crucial for understanding which roast categories drive sales.

3. **Size**: The size slicer lets the user filter the data based on the size of the coffee purchased, providing insights into customer preferences in terms of volume.

4. **Loyalty Card**: This slicer differentiates between customers who have a loyalty card and those who don’t, allowing for analysis of the impact of loyalty programs on sales.

The design and configuration of these slicers took a considerable amount of time. The focus on the user experience and ensuring that the slicers functioned smoothly to interact with the visualizations was a critical part of the process. This careful design ensures that users can easily manipulate the data to uncover key insights while maintaining a clean and professional layout.

#### The Design Process:
The attention to detail in the design was inspired by the visual style I aimed for — inspired by **The Economist's** sleek and professional design aesthetic. This involved configuring the slicers and their appearance, adjusting colors, and ensuring a user-friendly interface. The design elements were also carefully chosen to maintain clarity and focus on the data, so that users could quickly grasp the key insights from the dashboard.

<p align="center">
  <img src="https://github.com/user-attachments/assets/a388b4db-17bf-4ea0-9459-e809b0d9034f"/>
</p>

<p align="center">
  **Coffee Sales Dashboard**
</p>

## Conclusion

After performing the data cleaning, transformation, and visualization, the dashboard now provides clear insights into coffee sales trends across different coffee types, countries, and customer behavior. The line chart and bar chart effectively highlight the sales trends over time and across regions. The pivot tables and slicers further allow for deeper insights, enabling users to explore the data by different segments, such as roast type, size, and loyalty card status.

Creating these visualizations and the overall dashboard was a time-consuming process, especially due to the design aspects. I made sure the layout and color schemes adhered closely to the clean, professional design style seen in reputable sources like *The Economist*. This process not only allowed me to present the data in an engaging and interactive way but also helped me sharpen my skills in data visualization and dashboard creation.

Thank you for taking the time to review this project. I hope it provides valuable insights into coffee sales patterns and trends, and serves as a useful tool for decision-making.

---

**Feel free to explore the repository and let me know your thoughts or feedback.**

