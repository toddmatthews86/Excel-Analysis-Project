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

 ## Home Page
 
 ![Home](https://github.com/user-attachments/assets/9a250248-a907-4c3c-959a-898fb7cb8626)

 To make navigating the dashboard easier, each page includes the sidebar in the above image. Each of the 5 icons is linked to the repective  
 pages and clicking on one will take you to that page.

 Unlike the other pages, the Home page is not dynamic. It's purpose is simply to convey quickly and easily an important piece of information  
 from each of the pages. This information was drawn from the Pivot Tables created but for this purpose it was hardcoded in.

 ## Sales

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
totalRevenue:=SUM(order_totals[orderTotal])

medianSale:=MEDIAN(order_totals[orderTotal])

unitsSold:=SUM(order_details[quantity])
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

- Revenue has been steadily increasing over time. The last 3 quarters have all shown significant increase in revenue over the same period of the previous year.  
- The median sale total has remained steady over the course of the analysis, indicating that the increased revenue is due to an increase in the number of sales, rather than their size. Quantity over quality.  
- Janet Leverling was the best performed employee both in total revenue and median sale for Q1 2015.  

## Products

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

#### Insights  

- Beverages, Dairy Products, Confections and Seafood were easily the top 4 performing categories. They also the most number of available products.  
- While all categories have show a positive trend in revenue, 4 out of the 6 products in the Meat & Poultry category have been discontinued.  
With only 2 remaining products, this category is unlikely to continue the trend with out an injection of new products.  
- The top 5 products represent only 6% of the total products, but their $425k in revenue accounts for 35% of the total revenue.  

## Customers

![Customers](https://github.com/user-attachments/assets/21e02775-b7da-4513-81cc-430277475105)

The main tables this page focused were the customers and products page. As the relationship between these tables needed to pass through multiple other tables,  
I decided to create a new table in Power Query to have most of the necessary information in one place.  
The new table took the orderID, customerID and orderTotals from the orders table and merged with the order_details table.  
The result was a single table that included what products and the quantity each customer ordered.  
There was also simple relationships with the customer and product tables for extracting more information about the products and customers.

To build the pivot table, the company name filled the rows and the previously created measures of totalRevenue,  
medianSale and unitsSold provided the values.  
The table was sorted from highest to lowest in term of revenue and then filtered to only the top 10 results.

![Customers top 10](https://github.com/user-attachments/assets/30dbe4ae-fcc4-4022-aac4-507c4c4129a0)

Next I wanted to find ojut what products the top customers were purchasing, this required another table.  
I loaded a new table into Power Query that contained just the companyID of the top 10 customers.  
This was merged with the orders table to find only the orders for those customers.  
Finally another merge with the order_details table added the products purchased in those orders.

The pivot table used the produce name in the rows and the totalRevenue and unitsSold measures for values, filtered to the top 5.  
The combo chart generated showed the revenue for each product as a column and the unit sold as a line chart but only showing markers and the line removed.  

![Customers popular products](https://github.com/user-attachments/assets/e2db04e5-88fe-4fa6-a740-eee0373d6fb7)

For the other chart I wanted to show the countries where the most revenue is generated from.  
No new tables were required, simply a pivot table with the rows showing the country and totalRevenue again providing the values.    
Rather than using yet another bar or column chart, I mixed things up and used a tree map and I filtered to the top 5 results as this number  
allowed all labels to still be legible.  
I chose this chart as the total amount wasn't as important to me as quickly seeing what countries were included as well as their ratio to each other and the total.   
In addition if you hover the mouse over an area, a tool tip pops up showing the total for that area anyway.  

![Customer countries](https://github.com/user-attachments/assets/e873a394-1930-48e7-b254-b21690876106)

I made some KPI cards showing the number of total, active and new customers by loading the orders table into a sheet from Power Query and filtering by the relevant  
dates for each catergory. I then calculated a count of the unique customerIDs.

`
=SUMPRODUCT(1/COUNTIF(customer_count[customerID],customer_count[customerID]))
`  

![Customer KPI](https://github.com/user-attachments/assets/94616d3e-b21f-4f7f-80d3-96cbd991387f)

A timeline slicer was added to allow control over the time period.

#### Insights

- The top 10 customers generated 45% of the total revenue over the course of the entire analysis and 56% in Q1 2015.  
- QUICK-Stop, Ernst Handel and Save-a-Lot Markets have consistantly featured in the top 10 customers and alone brought in 24% of the total revenue.  
- USA, Germany and to a lesser extent Austria are the countries where most business originates. Further analysis could investigate how
  much that is due to the number of customers or the size of the customers.
- There seems to be a mix of high value, lower volume and low value, high volume products purchased by the top customers. Generally though it is a high value product in the #1 spot.

## Shipping

![Shipping](https://github.com/user-attachments/assets/c9ec5893-ac9a-4209-af43-c17c7980f3b5)

The major element of the Shipping page is the pivot table. I included this table because there is a number of performance indicators I wanted to include to compare  
the shipping companies. To show them all in charts would either make one big messy charts to too many smaller charts.  
I created another new table in Power Query by merging any data relevant to shipping so I could pick and choose what I needed.  
A column showing time to ship was created by using the Subtract Days function. A funtion that was also used in conjuction with a conditional column to assign  
each order a shipping status of "ON TIME", "LATE" or "AWAITING SHIPPING".  
Measures were created for average freight per order, freight-revenue ratio and average days to ship.

The pivot table was constructed the same way as the one in the Sales page. The rows were grouped by quarter, all of my measures provided the values and a second
instance of each measure was displayed as a difference from the previous year. Conditional formatting was used to used to highlight favourable/unfavourable differences.

![Shipping Table](https://github.com/user-attachments/assets/146f79b2-12fa-4d28-bdc8-7a5431cf0507)

The chart shows the average freight cost per order over time and the data is the same as that column in the pivot table.  
I just needed to create a separate pivot table in order to only visualise that data.  

![Shipping ave freight](https://github.com/user-attachments/assets/45858452-4abc-44da-bc0d-0252882e0421)

The last pivot table simply took a count of the shipping status column that was created.  
Visualising this data was ineffective as the values only varied slightly around a 95:5 ratio of ontime to late shipments.  
For this reason I chose to only show the late shipment rate as a KPI card.

![Shipping late](https://github.com/user-attachments/assets/2c7cf950-4e0e-40ff-823a-dc3989a561c3)

To add some dynamic control, 2 slicers were added to this page. One, to change between shipping company which is an obvious  
choice for a Shipping analysis.  

The other altered the country of destination. This was to show both how shipping volume compared between countries but also  
how the freight costs varied.

#### Insights

- Of the 3 companies used, there was no clear cut best option.  
   - United Package shipped the greatest olume but also had the highest average freight cost and freight-revenue ratio.  
   - Speedy Express had the lowest freight-revenue ratio but despite the name, was the slowest ship off orders and most likely to be late in doing so.  
   - Federal Shipping was the most reliable and a cheaper option but shipped low volumes outside the US.  
- Across all companies, the average freight cost has remained steady for the past 12 months. 








 

 
