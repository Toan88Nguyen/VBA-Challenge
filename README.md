# VBA Homework: The VBA of Wall Street

![stockmarket](https://user-images.githubusercontent.com/120751287/233563584-e1a1f039-0272-43d0-898e-0d233e46609b.jpg)

## Background

I am well on my way to becoming a programmer and Excel master! In this homework assignment, I will use VBA scripting to analyze generated stock market data. Depending on my comfort level with VBA, I may choose to challenge myself with a few of the challenge tasks.

## Instructions

Create a script that loops through all the stocks for one year and outputs the following information:

* The ticker symbol.

* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

* The total stock volume of the stock. The result should match the following image:

**Note:** Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

<p align="center">
<img width="718" alt="moderate_solution" src="https://user-images.githubusercontent.com/120751287/233561023-5d7063de-db10-4fec-9606-573e8802b9e9.png">

* Add functionality to my script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 
 
The solution should match the following image:

<p align="center"> 
<img width="1038" alt="hard_solution" src="https://user-images.githubusercontent.com/120751287/233562240-b041c8c5-bd44-4e0a-a9aa-cb57f9e6309d.png">

* Make the appropriate adjustments to my VBA script to enable it to run on every worksheet (that is, every year) at once.

## Other Considerations

* Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will allow you to test faster. My code should run on this file in less than 3 to 5 minutes.

* Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with one click of a button.
