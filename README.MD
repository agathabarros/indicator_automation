# Indicator Automation

### Objective: Train and create a Complete Project that involves automating a process carried out on the computer

# Description:

Imagine that you work in a large chain of clothing stores with 25 stores spread across Brazil.

Every day, in the morning, the data analysis team calculates the so-called One Pages and sends the store's OnePage to the manager of each store, as well as all the information used to calculate the indicators.

A One Page is a very simple and straight-to-the-point summary, used by the store management team to find out the main indicators of each store and allow on 1 page (hence the name OnePage) both the comparison between different stores, as well as which indicators that store managed to fulfill that day or not.

OnePage example:


![alt text](table.png)


Your role, as a Data Analyst, is to be able to create a process in the most automatic way possible to calculate the OnePage of each store and send an email to the manager of each store with your OnePage in the body of the email and also the file complete with the details of your respective store attached.

Ex: The email to be sent to the Manager of Store A should be as an example



![alt text](<email.png>)



### Important Files and Information

- Emails.xlsx file with the name, store and email of each manager. Note: I suggest replacing each manager's email column with your own email, so you can test the result

- Sales.xlsx file with sales from all stores. Note: Each manager should only receive OnePage and an attached Excel file with their store's sales. Information from another store should not be sent to a manager who is not from that store.

- store.csv file with the name of each store

- At the end, your routine should also send an email to the board of directors (information is also in the Emails.xlsx file) with 2 store rankings attached, 1 daily ranking and another annual ranking. Furthermore, in the body of the email, you must highlight which was the best and worst store of the day and also the best and worst store of the year. A store's ranking is given by the store's revenue.

- Spreadsheets for each store must be saved within the store folder with the date of the spreadsheet in order to create a backup history

### OnePage Indicators

- Revenue -> Target Year: 1,650,000 / Target Day: 1000
- Product Diversity (how many different products were sold in that period) -> Goal Year: 120 / Goal Day: 4
- Average Ticket per Sale -> Target Year: 500 / Target Day: 500

Note: Each indicator must be calculated on the day and year. The day indicator must be that of the last day available in the Sales spreadsheet (the most recent date)

Note2: Tip for the green and red sign character: take the character from this website (https://fsymbols.com/keyboard/windows/alt-codes/list/) and format it with HTML



