# 📊 Customer Data Analysis Using Excel

## 📌 Project Overview

This project focuses on analyzing customer data using Microsoft Excel to understand customer demographics, membership types, purchasing behavior, spending patterns, customer satisfaction, and discount usage.

The analysis uses Excel functions, formulas, and data summarization techniques to generate meaningful insights that help businesses understand their customers and improve decision-making.

## 🎯 Project Objectives

- Analyze customer demographics based on gender, age, and city.
- Understand customer distribution across membership types.
- Evaluate customer spending and purchasing behavior.
- Calculate total revenue and average spending.
- Analyze customer ratings and satisfaction levels.
- Compare customer behavior with and without discounts.
- Identify purchasing patterns across cities and membership types.
- Understand customer loyalty and purchase frequency.
- Generate meaningful business insights using Excel formulas.

## 📁 Dataset Description

The dataset contains customer information, including demographics, membership details, spending, purchasing activity, ratings, discount usage, and satisfaction levels.

### Dataset Columns

| Column Name | Description |
|---|---|
| Customer ID | Unique identifier assigned to each customer |
| Gender | Gender of the customer |
| Age | Age of the customer |
| City | City where the customer is located |
| Membership Type | Customer membership category: Gold, Silver, or Bronze |
| Total Spend | Total amount spent by the customer |
| Items Purchased | Number of items purchased by the customer |
| Average Rating | Average rating provided by the customer |
| Discount Applied | Indicates whether a discount was applied (TRUE/FALSE) |
| Days Since Last Purchase | Number of days since the customer's previous purchase |
| Satisfaction Level | Customer satisfaction category: Satisfied, Neutral, or Unsatisfied |

## 🛠️ Tools and Technologies

- **Microsoft Excel**
- Excel Formulas and Functions
- Data Analysis and Exploration
- Data Summarization
- Tables

## 📈 Key Analysis and Metrics

### 1. Customer Demographics

- Customer Count by Gender
- Customer Count by City
- Average Age by Gender
- Customer Count by City and Membership Type

### 2. Membership Analysis

- Customer Count by Membership Type
- Average Age by Membership Type
- Average Total Spend by Membership Type
- Total Revenue by Membership Type
- Average Items Purchased by Membership Type
- Average Days Since Last Purchase by Membership Type
- Average Rating by Membership Type

### 3. Spending and Revenue Analysis

- Average Total Spend by Gender
- Average Total Spend by City
- Total Revenue by City
- Total Revenue by Gender and Membership Type
- Average Spend by Gender and Membership Type
- Total Spend by Satisfaction Level and Membership Type

### 4. Purchasing Behavior

- Average Items Purchased by Gender
- Average Items Purchased by Membership Type
- Average Items Purchased by Satisfaction Level
- Average Days Since Last Purchase by City
- Average Days Since Last Purchase by Membership Type

### 5. Customer Satisfaction Analysis

- Customer Count by Satisfaction Level
- Average Spend by Satisfaction Level
- Average Rating by Satisfaction Level
- Average Items Purchased by Satisfaction Level
- Average Rating by City
- Average Rating by City and Membership Type

### 6. Discount Analysis

- Customer Count by Discount Applied (TRUE/FALSE)
- Average Spend With vs. Without Discount
- Average Rating With vs. Without Discount
- Average Items Purchased With vs. Without Discount

### 7. Overall Business Metrics

- Total Customer Count
- Total Revenue

## 🧮 Excel Functions Used

The following Excel functions can be used to analyze and summarize the customer dataset.

| Excel Function | Purpose |
|---|---|
| SUM | Calculates total spending or revenue |
| SUMIF | Calculates total spending based on one condition |
| SUMIFS | Calculates total spending based on multiple conditions |
| COUNT | Counts numeric values |
| COUNTA | Counts non-empty cells |
| COUNTIF | Counts customers matching one condition |
| COUNTIFS | Counts customers matching multiple conditions |
| AVERAGE | Calculates the average of numeric values |
| AVERAGEIF | Calculates averages based on one condition |
| AVERAGEIFS | Calculates averages based on multiple conditions |
| IF | Performs logical tests |
| MIN | Finds the minimum value |
| MAX | Finds the maximum value |
| TEXT | Formats numbers and dates as text |
| CONCAT | Combines text values |
| LEFT | Extracts characters from the beginning of text |
| RIGHT | Extracts characters from the end of text |
| MID | Extracts characters from the middle of text |

## 💻 Sample Excel Formulas

Assume the dataset is formatted as an Excel Table named `Customers`.

### 1. Total Customer Count

```excel
=COUNTA(Customers[Customer ID])
2. Total Revenue
=SUM(Customers[Total Spend])
3. Customer Count by Gender
=COUNTIF(Customers[Gender],A2)
4. Customer Count by Membership Type
=COUNTIF(Customers[Membership Type],A2)
5. Customer Count by Satisfaction Level
=COUNTIF(Customers[Satisfaction Level],A2)
6. Average Age by Gender
=AVERAGEIF(Customers[Gender],A2,Customers[Age])
7. Average Spend by Membership Type
=AVERAGEIF(Customers[Membership Type],A2,Customers[Total Spend])
8. Total Revenue by Membership Type
=SUMIF(Customers[Membership Type],A2,Customers[Total Spend])
9. Total Revenue by City
=SUMIF(Customers[City],A2,Customers[Total Spend])
10. Total Revenue by Gender and Membership Type
=SUMIFS(Customers[Total Spend],Customers[Gender],A2,Customers[Membership Type],B2)
11. Average Spend by Gender and Membership Type
=AVERAGEIFS(Customers[Total Spend],Customers[Gender],A2,Customers[Membership Type],B2)
12. Average Rating by Membership Type
=AVERAGEIF(Customers[Membership Type],A2,Customers[Average Rating])
13. Average Spend by Satisfaction Level
=AVERAGEIF(Customers[Satisfaction Level],A2,Customers[Total Spend])
14. Customer Count by Discount Applied
=COUNTIF(Customers[Discount Applied],TRUE)
To count customers who did not receive a discount:
=COUNTIF(Customers[Discount Applied],FALSE)
15. Average Spend With Discount
=AVERAGEIF(Customers[Discount Applied],TRUE,Customers[Total Spend])
16. Average Spend Without Discount
=AVERAGEIF(Customers[Discount Applied],FALSE,Customers[Total Spend])
Note: In the formulas above, A2 and B2 refer to cells containing the relevant criteria, such as Gender, Membership Type, City, or Satisfaction Level.

🏁 Conclusion
This Customer Data Analysis project demonstrates how Microsoft Excel can be used to explore customer data, calculate business metrics, and identify patterns in spending, purchasing behavior, membership preferences, and satisfaction.
By applying Excel functions, the project transforms raw customer data into useful information that supports data-driven business decisions.
