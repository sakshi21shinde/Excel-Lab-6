# Excel Lab 6 – DMart Retail Business Analysis

## Overview

This repository contains my Excel Lab 6 assignment based on a DMart Retail Business dataset. The objective of this lab is to apply Excel logical and error-handling functions to analyze retail transaction data and generate meaningful results based on different business conditions.

## Dataset

The dataset contains 15,345 DMart retail transaction records. It includes information such as Order ID, Region, City, Category, Customer Type, Order Date, Delivery Date, Sales, Discount, Quantity, Payment Mode, and Profit.

## Excel Functions Used

### IF Function

The IF function is used to evaluate a condition and return different results depending on whether the condition is TRUE or FALSE.

Example:

`=IF(H2>10000,"High","Low")`

This formula categorizes sales greater than 10,000 as "High" and sales of 10,000 or below as "Low".

### Nested IF Function

Nested IF is used when multiple conditions need to be evaluated.

Example:

`=IF(H2>10000,"High",IF(H2>5000,"Medium","Low"))`

This categorizes the sales values into three levels:

- Sales greater than 10,000 → High
- Sales greater than 5,000 → Medium
- Sales of 5,000 or below → Low

### AND Function

The AND function is used when all specified conditions must be TRUE.

Example:

`=IF(AND(H2>5000,Q2>500),"Good","Not Good")`

This checks whether Sales are greater than 5,000 AND Profit is greater than 500. If both conditions are satisfied, the result is "Good"; otherwise, it is "Not Good".

### OR Function

The OR function is used when at least one of the specified conditions must be TRUE.

Example:

`=IF(OR(H2>5000,N2>0.5),"High","Low")`

This checks whether Sales are greater than 5,000 OR Discount is greater than 50%. If either condition is satisfied, the result is "High"; otherwise, it is "Low".

### NOT and AND Functions

The NOT function reverses the logical result of a condition, while AND checks whether all specified conditions are TRUE.

Example:

`=IF(NOT(AND(H2>5000,N2>500)),"Not Good","Good")`

This formula combines NOT and AND to evaluate the given Sales and Discount conditions.

### Profit-to-Revenue Calculation

A profit-to-revenue ratio is calculated to understand the proportion of profit generated from sales.

Example:

`=Table1[[#This Row],[profit]]/Table1[[#This Row],[sales]]`

This calculates the profit-to-revenue ratio for each transaction.

### IFERROR Function

The IFERROR function is used to handle errors that may occur during calculations.

Example:

`=IFERROR(U2,"Error")`

If the calculation produces an error, the formula returns "Error" instead of displaying an Excel error message.

## Objectives

The main objectives of this lab are to understand and apply Excel logical functions to real-world retail data, classify sales based on different conditions, evaluate sales and profit performance, analyze discount-related conditions, calculate profit-to-revenue ratios, and handle calculation errors using IFERROR.

## Tools and Functions

- Microsoft Excel
- IF Function
- Nested IF Function
- AND Function
- OR Function
- NOT Function
- IFERROR Function
- Excel Tables
- Logical Conditions
- Retail Data Analysis

## Skills Practiced

Through this lab, I practiced Excel logical functions, conditional analysis, nested conditions, business data analysis, retail sales analysis, profit analysis, error handling, and formula-based data analysis.

## Repository Contents

The repository contains the Excel workbook `Lab6.xlsx`, which includes the DMart Retail Business dataset and the applied formulas and analysis, along with this `README.md` file containing the documentation for the lab.

## Conclusion

This Excel Lab 6 assignment provided practical experience in applying logical functions such as IF, AND, OR, and NOT to a real-world DMart retail dataset. The use of IFERROR and profit-to-revenue calculations further helped in performing reliable and meaningful business data analysis.

## Author

Sakshi Shinde

Bachelor of Engineering – Computer Engineering
