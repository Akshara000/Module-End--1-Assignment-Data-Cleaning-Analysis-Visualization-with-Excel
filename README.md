# Module-End--1-Assignment-Data-Cleaning-Analysis-Visualization-with-Excel
Module 1 end evaluation assignment
 have completed Module 1 End evaluation assignment Module-End -1
Module End  Assignment: 
“Data Cleaning, Analysis & Visualization with Excel

Step 1: Data entry and tables
Imported E commerce data set to excel then each sheet make to table by pressing Ctrl+T.
Step 2: Data cleaning
Saved the data the transformed to power query.Four sheets are transformed to power query.
First sheet is customer-Using trim,Capitilise each word cleaned customer name.Country is written 3 ways: USA, usa, United States of America.These are replaced by USA .4 blank Loyalty_Level replaced by Unknown.4 IDs are Started with C- that is replaced by CUST.
Second Sheet is Customer-2 Ids are started with P- it replaced by PROD.Some cost are missed using average value to replace it.And the stock value replaced by median.The column Sub_category deleted because it is irrelevant.
Sales sheet - there are missing values in Quantity,unit price,Discount and total amount.created Quantity updated column the using the custom column option.Formula If [Quantity]=null then Total amount/(unit price*(1-Discount)).Then same like this created Columns-unit price updated ,Total amount Updated,Discount Updated.The basic formula for these calculation is Total amount=Quantity*Unit Price*(1-Discount).After finding the missing values using the base formula there missing value on both unit price and discount. In this case I calculated median of the discount that is 0.14 and added to missing cell in discount then again calculated unit price column by custom column option.
Step 3: Combine tables with XLOOKUP
The dataset loaded to excel.then created a new sheet named-“final lookup” . In order to combine sheets used XLOOKUP here.Pasted Sales_fact sheet table to this new sheet.then created column named Customer name-=XLOOKUP([@[Customer_ID]],’Customer_Dim'!A:A,'Customer_Dim'!B:B,"NA")
Then Created column named State-=XLOOKUP([@[Customer_ID]],’Customer_Dim'!A:A,'Customer_Dim'!F:F,"NA")
Then Created column named Loyality-==XLOOKUP([@[Customer_ID]],’Customer_Dim'!A:A,'Customer_Dim'!H:H,"NA")
Then Created column named Product Name-=XLOOKUP([@[Product_ID]],Product_Dim!A:A,Product_Dim!B:B,"NA")
Then Created column named Category-=XLOOKUP([@[Product_ID]],Product_Dim!A:A,Product_Dim!C:C,"NA")
Then Created column named Brand-=XLOOKUP([@[Product_ID]],Product_Dim!A:A,Product_Dim!D:D,"NA")
Then Created column named Cost-=XLOOKUP([@[Product_ID]],Product_Dim!A:A,Product_Dim!E:E,"NA")
Then Created column named Region-=XLOOKUP([@[Store_ID]],Store!A:A,Store!C:C,"NA")
Then Created column named Store_type-=XLOOKUP([@[Store_ID]],Store!A:A,Store!E:E,"NA")


Step 4: Formulas
Created an new column name ‘Order Size’ on the sheet ‘final lookup’-Filled the column used the function IF()-=IF([@[Total amount]]>1500,”High","Normal")
Then create column named month using TEXT()-=TEXT([@[Order_Date]],”mmm-yyy")

Then created new sheet named- ‘Formula and Statstics’
Total revenue : =SUM(‘final lookup’!J:J)
Total Quantity: =SUM(‘final lookup’!F:F)
Total Revenue by Category Sports(SUMIF):=SUMIF('final lookup'!O:O,"Sports",'final lookup'!J:J)
Orders per payment type Paypal (COUNTIF):=COUNTIF('final lookup'!I:I,"PayPal")
Number of Orders:=COUNTA('final lookup’!A:A)


Step 5: Descriptive statistics
Then Descriptive statistics of Total amount,Quantity,Discount by Data analysis tool and the result showed in the sheet name ‘Formula and Statstics’.

Step 6: Pivot tables
Created sheet named -‘Pivot’
Then create pivot tables named-Revenue VS Category,Region and Category V/S Total amount,Month V/S Revenue, Loyalit&Payment type V/s Rvenue & Order size,Category V/s Discount,Category V/S Quantity and respective chart are also created using the same names. The chart are - Column Chart,Bar Chart,Line chart,Compo Chart,Donut Chart,and line Chart  respectively.

Step 7: Visualization and dashboard

Create Sheet name Dashboard-Header Given “Ecommerce Sales Dashboard”
Then added 4 KPIs-Total Revenue,No of Orders,Total Discount and Top Category.
Then arranged charts from Pivot sheet .Then added slicers that are connected to all the pivot charts created.

Step 8:summary-

Top category: Sports is the best-selling category with 542K in revenue, followed by Electronics at 504K .
Best region: North generated the most revenue at 646K, but South and East are within 4%. West is far behind at 271K, less than half of any other region.
Trend: Monthly revenue is broadly stable at 82K to 108K, peaking in March of both years. The sharp drops in Oct-23 and Oct-25 are partial months, not real declines 
Payments: PayPal is the most-used method (695 orders, 35% of revenue), but Credit Card and COD are almost level, so no method dominates.
