# Data-Modeling-and-Analysis-in-Excel

## Project Description:
## Project: Coffee Sales Analysis
![img alt](https://github.com/nsankareswari-70/Data-Modeling-and-Analysis-in-Excel/blob/32f5ebc8ba6be11308dfac93482b67a258d5f112/cofeespecies.png)
## Objectives

- Identify which species of coffee (e.g., Arabica, Robusta, etc.) are selling the most.

- Find the most frequent customers (by number of orders).

- Determine which customers generated the highest total sales (by revenue).

- Spot overall sales trends to guide marketing and inventory decisions.


## Data sources used for this project are 
Orders - Fact Table   
Customers - Dimension Table   
Products - Dimension Table   

## Data description

- Ara - Arabica
- Rob - Robusta
- Lib - Liberica
- Exc - Excelsa

  Roast Type:
  L - Light Roasted.
  M - Medium Roasted.
  D - Dark Roasted

  ## Creating a calculated column Profit in the orders table as    
  
=VLOOKUP([@[Product ID]],products[#All],7,FALSE)

Finalprofit  =[@Quantity]*[@[Products(Profit)]]

## To find the customers who made most orders

First remove the duplicates from the customerId column

we get 913 Unique customers

![img alt](https://github.com/nsankareswari-70/Data-Modeling-and-Analysis-in-Excel/blob/59bd87b13d439fe72e0f71e0c23eaa69e6c1eeb6/dme3.png)

Now we need to get the count of each customerId in the orders table to find the customer who made the max order.

=COUNTIF(orders[Customer ID],K5)

Now sort the table desending by the column NumberofOrders

![img alt](https://github.com/nsankareswari-70/Data-Modeling-and-Analysis-in-Excel/blob/7d860d95212b9af7f95a8ebc6c42d7e2db374887/dme4.png)

From the above data we can see CustomerID 86579-92122-OC made - 7 orders

  To find that customer from the customer table use vlookup function
  =VLOOKUP(B1006,customers[#All],2,FALSE)

  ![img alt](https://github.com/nsankareswari-70/Data-Modeling-and-Analysis-in-Excel/blob/d5d28018cc976e293c039ba59e2ca7a4696a5268/dme5.png)

  ## Lets create another pivot table to see the profit by coffee type and its Roast type

![img alt](https://github.com/nsankareswari-70/Data-Modeling-and-Analysis-in-Excel/blob/8a7af178a72168eedf5954a51781815bdb3b8558/dme6.png)
  

