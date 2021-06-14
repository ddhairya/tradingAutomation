# Trading Automation

I want to thanks the developer of [yfianace library](https://github.com/ranaroussi/yfinance) , which is the core for this project.

A program to get a details on status of current holding to make a call for the stocks and it give answer on questions such as:

- Am i making a profit or loss!
- How much Qty of share i need to sell to get my actual investsed amount!
- What's my annual return!
- What's my average growth!
- What's current market volume and capital for the company! ...

For this app we will be giving an input of stocklist as an excel file, and it's the only dependency to get the actual output.

> **Note:** When you clone the repo, you have make a change in the path where you are going to store the stocklist.xlsx file. For this repo it's in onedrive you can keep it in any destination.
> 
> `stocks_list = pd.read_excel(OneDrive + "stockslist.xlsx")`

Have user define funtions to get insights
- Annual Return
- Percentage Gain or Loss
- Percentage Growth ( of invested amount)
- Profit Booking ( No of shares required to sell to get invested amount back from profits)
