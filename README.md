# Excel Analysis Project

![Dashboard](https://github.com/user-attachments/assets/60e84ba6-0e60-4a70-9056-df3a8d7e3979)

 ## Introduction  
 
 As an aspiring data analyst, I completed this project to demonstate my Excel skills.   
 The dataset is from a past challenge hosted by	[Maven Analytics](https://www.mavenanalytics.io) where the goal was to build a dashboard for business executives  
 to understand the companies performances.  
 While the original challenge saw entries completed in Power BI, I felt the dataset showed promise for an Excel project.  

 ### Final Dashboard  
 The final dashboard can be found in this file [Excel Project](https://github.com/toddmatthews86/Excel-Analysis-Project/blob/main/Excel%20Project.xlsx)  
 
 ### The Dataset
 The dataset was presented in 7 csv files -  

 - orders *(the when, where and who of all transactions)*
 - order_details *(the what about each individual transaction)*
 - products
 - categories *(product categories)*
 - customers
 - employees
 - shippers

 ### Skills Used  
 The following skills were needed for the final product to come together.  

 - Pivot Tables
 - Pivot Charts
 - Conditional Formatting
 - DAX (Data Analysis Expressions)
 - Power Query
 - Power Pivot

 ### Building the Dashboard
 While familiarising myself with the dataset, I identified 4 key areas that could serve as the basis for my analysis.
 - Sales
 - Products
 - Customers
 - Shipping

 While it would be possible to build a single page dashboard covering each of these areas, I decided it would be most practical to implement  
 a multi-page design. This design would use a "Home" page as a landing location where a key insight from each section would be displayed.
 A full analysis of each of the 4 sections then occupied its own page. 

 All of the data handling was conducted through an Extract, Transform, Load (ETL) process using Power Query.  
 The files that were required in the transformation process were added to Power Query and their data extracted.  
 The necessary transformations were carried out and then the transformed data was loaded as a connection only and added to the data model.
 
 Once added to the data model, Power Pivot was used to manage relationships between tables, create measures and add pivot tables.
 Data Analysis Expressions (DAX) was invaluable in setting up relationships how I needed them and creating measures when building pivot tables.

 Due to excels inability to create bi-directional relationships between tables, I found myself often using DAX to get around this.

 `
 =CALCULATE([totalRevenue],CROSSFILTER(orders[orderID],order_details[orderID],Both))
 `
 This combination of the CALCULATE and CROSSFILTER functions, allowed me to use measures in a bi-directional manner. This gave me more freedom  
 in how I set up my pivot tables.

 The final table relationships. It is a bit messy and most likely involves some redundancy but it works.  
 Streamlining this would be a future update I would consider.

 ![Table Relationships](https://github.com/user-attachments/assets/597562f8-1987-4e28-8182-fcd05ded010b)

 #### Home Page
 
 ![Home](https://github.com/user-attachments/assets/9a250248-a907-4c3c-959a-898fb7cb8626)

 To make navigating the dashboard easier, each page includes the sidebar in the above image. Each of the 5 icons is linked to the repective  
 pages and clicking on one will take you to that page.

 Unlike the other pages, the Home page is not dynamic. It's purpose is simply to convey quickly and easily an important piece of information  
 from each of the pages. This information was drawn from the Pivot Tables created but for this purpose it was hardcoded in.

 #### Sales

 ![Sales](https://github.com/user-attachments/assets/773ddb3c-0670-4eaa-93ab-172391dc3683)

 Unsurpisingly, revenue is a very important attribute when looking at Sales performance. Therefore my first task when building the Sales  
 pages was to determine the total of each sale so that an overall revenue amount could be calculated.  
 
 This was achieved in Power Query by transforming the order_details table.
 -multiplication of unitPrice and quantity values of each row gave a total price per product but there could be multiple products per order.
 -a combination of pivotting some columns and then unpivotting others resulted in one row per order and a grand total for that order.
 -unnecessary columns were removed leaving just the order number and total, which was then merged with the orders table.

 In Power Pivot I could create a measures for total revenue and median sale from this new column. I also used the quantity column in the order_details  
 table to create a measure for units sold.

 ```
totalRevenue:=SUM(column_name)

medianSale:=MEDIAN(column_name)

unitsSold:=SUM(coumn_name)
```
These measures formed the basis for all of the information displayed in the Sales page.  
From Power Pivot I initialised a pivot table where I used the order date in the rows to group revenue, median sales and units sold values by quarter.  
Originally I had extracted the month from the order date in power query bu this wasnt used as Excel can do it within the pivot table for you.  

I added each measure to the pivot table a second time but this time changed the value field settings to show it as the difference from  
the previous year (either as a % or absolute value). Along with some conditional formatting, this provided insight into how the measures  
compare to the previous year.  

![Show as difference](https://github.com/user-attachments/assets/a95b5d63-2ed5-40ea-b444-6793ae08277b)

I created KPI cards for the three measures so at a glance you can see the performance in the most recent period, and the comparison the the previous period.  

The final dashboard had a second version of this pivot table where units sold was dropped as I only wanted to visualise revenue and median sales.

![Sales Pivot](https://github.com/user-attachments/assets/8387c5fc-5e75-40ad-b8cf-7a71f36e7890)

From the above pivot table I created a pivot chart showing the total revanue and median sale over time.  
This combo chart uses an area chart for revenue and a line chart for median sale.

![Sales Chart](https://github.com/user-attachments/assets/74a0ce47-d3e7-4d00-a3b2-0e220a69cf01)

The final piece of the Sale page is a slicer filtering by employee name. Filtering this way adds another layer of analysis by giving the user control  
over seeing either company wide performance or drilling down to see the performance of individual employees.

**Insights**  
-Revenue has been steadily increasing over time. The last 3 quarters have all shown significant increase in revenue over the same period of the previous year.
-The median sale total has remained steady over the course of the analysis, indicating that the increased revenue is due to an increase in the   
number of sales, rather than their size. Quantity over quality.
-Janet Leverling was the best performed employee both in total revenue and median sale for Q1 2015

####Products

![Products](https://github.com/user-attachments/assets/76fbdb86-8f96-4b3c-8b94-882625e4f203)

There were 5 pivot tables created for this page.  
The first simply took the product category in the rows and the previously created totalRevenue measure as the values.  
Visualising this data in a bar chart easily shows which categories are bringing in the most revenue.

![Categories bar chart](https://github.com/user-attachments/assets/f3653b1a-650d-444f-a43d-7835b9a9862e)

Like the sales page, this page also used a pivot table to show totalRevnue over time and visualised it as an area chart.  
I included a trendline to highlight the revnue trend ove rthe time period.

![Product Revenue](https://github.com/user-attachments/assets/6eefb0c7-3b25-468d-9243-771690c72c70)

The next pivot table was to show how many products were either active or had been discontinued.  
The products data had a column for this where active = 0 and discontinued = 1.  
In Power Query I created a new conditional column where the values were either "ACTIVE" or "DISCONTINUED".

![Products Conditional Column](https://github.com/user-attachments/assets/5319e65b-6261-43e8-b0c4-c241e60302c4)

A count of these values in the pivot table produced the values I needed to be linked to KPI cards.

The final two pivot tables were directly placed into the dashboard and they show the best 5 and worst 5 performing products by revenue.  
The product name was placed in the rows with the quantity and totalRevenue providing the values.  
The top products table was sorted from largest to smallest revenue and filtered by the top 5, while the bottom products was sorted smallest  
largest and filtered by bottom 5.

![Products top bottom 5](https://github.com/user-attachments/assets/7fc3e111-a74f-44fd-8e0e-686af0329ee0)

The slicer on this page allowed the data to be further broken down by category.  
Only the bar chart was not dynamic as that was already filtered by category, all other elements dynamically changed based on the slicer.

**Insights**
-Beverages, Dairy Products, Confections and Seafood were easily the top 4 performing categories. They also the most number of available products.  
-While all categories have show a positive trend in revenue, 4 out of the 6 products in the Meat & Poultry category have been discontinued.  
With only 2 remaining products, this category is unlikely to continue the trend with out an injection of new products.
-The top 5 products represent only 6% of the total products, but their $425k in revenue accounts for 35% of the total revenue.





 

 
