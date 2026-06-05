Project Title
Data Exploration

Step 1: Data Structuring & Baseline Metrics

I started with a transactional dataset containing 50 order records. The dataset included the following baseline columns:

Order ID (formatted as ORD-1001, ORD-1002, etc.)
Date, Region, Product, and Category
Units Sold, Unit Price ($), and Discount (%)

Step 2: Revenue and Category Formulations (Calculated Columns)

I expanded the dataset by creating several calculated columns using mathematical, logical, and text functions across rows 2 to 51.

Price Range Classification

I categorized products based on their unit price using an IF statement. Products with a unit price greater than $500, such as Laptop Pro, were classified as "High Price", while all other products were classified as "Standard Price".

Excel Formula:

=IF(G2 > 500, "High Price", "Standard Price")
Text Manipulation Using Excel Functions

To practice Excel text functions, I used the Order ID column to extract specific portions of the text.

Day Column

I extracted the first two characters from the Order ID, resulting in values such as "OR".

Formula:

=LEFT(A2, 2)
Country Code Column

I extracted the last two characters from the Order ID to generate sequential values ranging from 01 to 50.

Formula:

=RIGHT(A2, 2)

or

=VALUE(RIGHT(A2, 2))
Month Column

I extracted three characters starting from the fourth position of the Order ID, resulting in values such as "-10".

Formula:

=MID(A2, 4, 3)

Step 3: Summary Statistics & KPI Block

I created a KPI summary table on the right side of the worksheet (Columns N and O, Rows 4 to 9) to analyze and summarize the Unit Price data.

Total Price ($10,349.50)

I calculated the total sum of all unit prices using the SUM function.

Formula:

=SUM(G2:G51)
Average Price ($206.99)

I calculated the average unit price across all products using the AVERAGE function.

Formula:

=AVERAGE(G2:G51)
Minimum Price ($9.99)

I identified the lowest-priced product using the MIN function.

Formula:

=MIN(G2:G51)
Maximum Price ($1,299.99)

I identified the highest-priced product using the MAX function.

Formula:

=MAX(G2:G51)
Electronics Price ($7,179.81)

I used the SUMIF function to calculate the total unit price of products belonging to the Electronics category.

Formula:

=SUMIF(E2:E51, "Electronics", G2:G51)
Price Less Than $100 (34 Transactions)

I used the COUNTIF function to determine how many transactions had a unit price below $100.

Formula:

=COUNTIF(G2:G51, "<100")
Conclusion

Through this process, I structured the dataset, created calculated fields, applied logical and text functions, and developed a KPI summary section to analyze pricing trends and product performance effectively.
