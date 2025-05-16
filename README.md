# **Window Breakage Rate Analysis Dashboard**

This Shiny app provides an end-to-end analytics solution for understanding and optimizing the breakage rate in window manufacturing processes. It integrates descriptive, predictive, and prescriptive analytics to assist manufacturing technicians and process engineers in making data-driven decisions to reduce product defects.

## **Dataset**
The data is read from an Excel file:
📁 Window_Manufacturing.xlsx

**Columns (after cleaning & transformation):**

- Breakage_Rate: Target variable

- Window_Size, Glass_Thickness, Ambient_Temp, Cut_Speed, Edge_Deletion_Rate, Spacer_Distance, Silicon_Viscosity: Numeric predictors

- Window_Type, Glass_Supplier, Glass_Supplier_Location: Categorical predictors (converted to dummy variables)

- Imputation for missing values is performed using the CART method from the mice package.
