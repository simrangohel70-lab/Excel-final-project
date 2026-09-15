# 📊 Final Project – Sales Data Analysis

![Microsoft Excel](https://img.shields.io/badge/Microsoft%20Excel-Data%20Analysis-green?logo=microsoftexcel)
![Sales Analysis](https://img.shields.io/badge/Sales-Analysis-blue)
![Data Cleaning](https://img.shields.io/badge/Data-Cleaning-orange)
![Pivot Table](https://img.shields.io/badge/Pivot%20Table-Analysis-purple)
![Charts](https://img.shields.io/badge/Charts-Visualization-red)
![Dashboard](https://img.shields.io/badge/Dashboard-Completed-success)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen)

---

## 📑 Table of Contents

* [📌 Project Overview](#-project-overview)
* [🎯 Project Objectives](#-project-objectives)
* [📂 Workbook Structure](#-workbook-structure)
* [📋 Dataset Description](#-dataset-description)
* [🧹 Data Cleaning & Preparation](#-data-cleaning--preparation)
* [🧮 Excel Formulas Used](#-excel-formulas-used)
* [🔍 Data Analysis](#-data-analysis)
* [🔄 Pivot Table Analysis](#-pivot-table-analysis)
* [📊 Charts & Visualization](#-charts--visualization)
* [📈 Dashboard](#-dashboard)
* [💡 Key Insights](#-key-insights)
* [🛠️ Tools & Skills](#️-tools--skills)
* [📁 Project File](#-project-file)
* [🎥 Video Link](#-video-link)
* [👩‍💻 Project By](#-project-by)

---

## 🎥 Video Link

🎬 **Project Demonstration Video**

The video explains the complete project workflow, including the formulas, data analysis, Pivot Tables, charts, and dashboard.

🔗 **Video Link:** https://drive.google.com/file/d/1lR4ZCv3Y1kx_cunYEzt0y414KaXW2Hei/view?usp=drive_link


---

## 📌 Project Overview

This project is an **Excel-based Sales Data Analysis project** created to transform raw sales transaction data into meaningful business insights.

The project covers the complete data analysis process, starting from **raw data preparation and cleaning** to **formula-based analysis, Pivot Tables, charts, and an executive dashboard**.

The final workbook provides a clear view of sales performance across different **products, categories, regions, salespeople, payment methods, and order statuses**.

---

## 🎯 Project Objectives

The main objectives of this project are:

* 🧹 Clean and organize sales data
* 🔎 Perform data lookup and analysis using Excel formulas
* 💰 Analyze overall sales performance
* 📦 Analyze product and category performance
* 🌎 Compare sales across different regions
* 👨‍💼 Analyze salesperson performance
* 🔄 Create Pivot Tables for summarized data
* 📊 Create charts for better visualization
* 📈 Build an Executive Sales Analytics Dashboard
* 💡 Extract useful business insights from the dataset

---

## 📂 Workbook Structure

The Excel workbook contains the following sheets:

| Sheet               | Purpose                                                     |
| ------------------- | ----------------------------------------------------------- |
| 📋 **Raw data**     | Contains the original sales transaction data                |
| 🧹 **Cleaned data** | Contains cleaned data and additional formula-based analysis |
| 📊 **Charts**       | Contains visual charts for sales analysis                   |
| 🔄 **Pivot table**  | Contains summarized Pivot Table analysis                    |
| 📈 **Dashboard**    | Contains the final Executive Sales Analytics Dashboard      |

---

## 📋 Dataset Description

The dataset contains **153 sales transactions** with the following columns:

| Column                | Description                        |
| --------------------- | ---------------------------------- |
| 🧾 **Order ID**       | Unique order identification number |
| 📅 **Order Date**     | Date on which the order was placed |
| 🆔 **Product ID**     | Unique product identification      |
| 📦 **Product**        | Product name                       |
| 🏷️ **Category**      | Product category                   |
| 🌎 **Region**         | Sales region                       |
| 👨‍💼 **Salesperson** | Person responsible for the sale    |
| 👤 **Customer**       | Customer identification            |
| 🔢 **Quantity**       | Number of units sold               |
| 💰 **Unit Price**     | Price per unit                     |
| 🏷️ **Discount**      | Discount applied to the order      |
| 💵 **Sales Amount**   | Final sales amount                 |
| 💳 **Payment Method** | Payment method used                |
| 📌 **Order Status**   | Completed, Pending, or Cancelled   |

---

## 🧹 Data Cleaning & Preparation

The **Cleaned data** sheet was created to organize the original dataset and perform additional analysis.

Additional columns were added for:

* 🔍 Product Details
* 💰 High/Low Value classification
* 🔎 Order Search
* ⭐ Result classification
* 🧾 Total Number of Orders

The data was structured in an Excel Table to make formulas and analysis easier to manage.

---

## 🧮 Excel Formulas Used

Several Excel functions were used to perform the analysis.

### 🔹 VLOOKUP

Used to retrieve product details using the Product ID.

```excel
=VLOOKUP("P002",C2:D13,2,FALSE)
```

### 🔹 IF

Used to classify sales into **High Sale** and **Low Sale**.

```excel
=IF(L2>100000,"High Sale","Low Sale")
```

### 🔹 INDEX + MATCH

Used to search for a specific order and return the corresponding product.

```excel
=INDEX(D2:D120,MATCH(1101,A2:A120,0))
```

### 🔹 Nested IF

Used to classify sales performance as **Excellent, Good, or Average**.

```excel
=IF(L2>=100000,"Excellent",IF(L2>=50000,"Good","Average"))
```

### 🔹 ROWS

Used to calculate the total number of orders in the Excel Table.

```excel
=ROWS(Table1[Order ID])
```

---

## 🔍 Data Analysis

The project analyzes sales from multiple perspectives, including:

### 📦 Product Analysis

Sales performance was analyzed for products such as:

* Laptop
* Smartphone
* Monitor
* Tablet
* Office Chair
* Printer
* Desk
* Headphones
* Keyboard
* Mouse

### 🏷️ Category Analysis

The dataset contains four major categories:

* 💻 Electronics
* 🪑 Furniture
* 📎 Office
* 🎧 Accessories

### 🌎 Regional Analysis

Sales were analyzed across:

* East
* West
* North
* South

### 👨‍💼 Salesperson Analysis

The project also compares sales performance among different salespeople.

### 💳 Payment Method Analysis

Orders were analyzed based on:

* Cash
* Card
* UPI
* Bank Transfer

### 📌 Order Status Analysis

Orders were categorized into:

* ✅ Completed
* ⏳ Pending
* ❌ Cancelled

---

## 🔄 Pivot Table Analysis

Pivot Tables were used to summarize large amounts of sales data and make comparisons easier.

The Pivot Table analysis includes summaries such as:

* 📊 Sales by Category
* 🌎 Sales by Region
* 📅 Monthly analysis
* 💰 Sales amount summaries
* 📦 Product-related analysis

Pivot Tables helped convert detailed transaction-level data into easy-to-understand summaries.

---

## 📊 Charts & Visualization

The **Charts** sheet is used to visually represent the analyzed data.

Charts help identify:

* 📈 Sales performance
* 🏆 Top-performing categories
* 🌎 Regional performance
* 📦 Product performance
* 📅 Monthly trends
* 📌 Order-related patterns

These visualizations make the analysis easier to understand and present.

---

## 📈 Dashboard

The final workbook includes an **Executive Sales Analytics Dashboard**.

The dashboard provides a summarized view of the overall sales performance and helps users quickly understand important business metrics.

### Dashboard Highlights

* 💰 Total Sales
* 📦 Total Quantity
* 🧾 Total Orders
* 🏆 Category Performance
* 🌎 Regional Performance
* 📊 Sales Trends
* 📌 Order Status Analysis

The dashboard brings the important results together in one professional view.

---

## 💡 Key Insights

Based on the analysis:

### 💰 Overall Performance

* **Total Sales:** ₹10,588,845
* **Total Quantity Sold:** 699
* **Total Transactions:** 153

### 🏆 Category Performance

**Electronics** is the highest-performing category with approximately **₹8.31 million** in sales.

### 🌎 Regional Performance

The **East Region** generated the highest sales with approximately **₹3.11 million**.

### 📦 Product Performance

**Laptop** is the top-performing product with approximately **₹4.20 million** in sales.

### 📌 Order Status

| Order Status | Orders |
| ------------ | -----: |
| ✅ Completed  |     88 |
| ⏳ Pending    |     35 |
| ❌ Cancelled  |     30 |

This shows that **Completed orders form the largest portion of the transactions**.

---

## 🛠️ Tools & Skills

### 🧰 Tools

* 🟢 Microsoft Excel
* 📊 Excel Charts
* 🔄 Pivot Tables
* 📈 Excel Dashboard

### 💻 Skills Demonstrated

* Data Cleaning
* Data Organization
* Excel Formulas
* Lookup Functions
* Conditional Logic
* Data Analysis
* Pivot Table Analysis
* Data Visualization
* Dashboard Creation
* Business Insights

---

## 📁 Project File

The complete Excel workbook is included in this repository:

**📄 Final Project.xlsx**

It contains the complete workflow from **Raw Data → Cleaned Data → Pivot Analysis → Charts → Dashboard**.

---


---

## 👩‍💻 Project By

### **Simran Gohel**

📊 **Excel | Data Analysis | Visualization | Dashboard**

---

## ⭐ Conclusion

This project demonstrates practical knowledge of **Microsoft Excel and Data Analysis** by converting raw sales data into structured information and meaningful business insights.

The combination of **Excel formulas, data cleaning, Pivot Tables, charts, and dashboard visualization** makes the project useful for understanding real-world sales analysis workflows.

---

⭐ **If you found this project useful, feel free to explore the workbook and analysis.**
