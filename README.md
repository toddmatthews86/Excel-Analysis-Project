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

 


 


 

 
