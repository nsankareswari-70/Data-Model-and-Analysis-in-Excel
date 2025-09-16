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

  

