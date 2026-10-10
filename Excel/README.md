# Sales Analysis
## Project Overview
This is a clean, formula-driven solution that classifies e-commerce orders into actionable categories using pure Excel logic with no VBA, no Power Query, no external tools.
The workbook automatically evaluates every order against four business rules and generates the following columns:

- East Completed which identifies fully completed orders from the East region
- Large Completed which flags high-volume (≥ 10 units) completed orders
- Digital Order which detects cash or USSD payments
- Priority Order which elevates orders that are either East and Completed or paid Online

All logic is implemented with nested IF, AND, and OR formulas for maximum transparency, auditability, and portability, which simply drop new order data into the sheet and the classifications update instantly.

This process is ideal for companies, operations teams, and anyone who needs reliable, zero-code order prioritization inside Excel.

## Business Problem
Operation team needs a clear understanding of how sales is achieved by classification, region, products, and method for key performance indication.

This Excel analysis  was carried out to answer the following questions:

- How to quickly identify completed East-region orders?
- How to flag large completed orders (Quantity ≥ 10)?
- How to detect Cash or USSD payments automatically?
- How to mark priority orders based on region, status, or payment method?
- How to eliminate slow and error-prone manual checks?
- How to update classifications instantly with new data?
- How to ensure consistent order labeling across the team?
- How to create transparent and auditable logic?
- How to achieve this without VBA, macros, or external tools?

## Tools & Technologies
- Microsoft Excel
- IF/NESTED IF Function
- OR Function
- AND Function
- GitHub

## Dataset
The dataset contains information about raw sales.

The key fields include:

- Payment
- Category
- Quantity
- Unit Price
- Discount
- Payment Method
