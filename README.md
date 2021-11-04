# Project Overview

## Background

In order to help my client Steve, I have created a subroutine (macro) that calculates the total daily trade volume and percentage return. While looking into green energy stocks, Steve has decided to research other stock options in order to diversify his parents portfolio.

  

## Purpose of Analysis

Now, since we are looking into other stock options, we will have a much larger data set to gather and calculate data from. Knowing this, we needed to refactor our “AllStockAnalysis” subroutine in order to handle the increased data set.

  
  

# Results

## 2017 Stock Analysis

When looking at the “All Stocks (2017)” table below, you can see that 2017 was actually quite a successful year for green energy stocks. All but one of the stock options are in the green meaning they had a positive return!

  

#### 2017 Stock Option Results

![2017 Stock Option Results](https://github.com/StickySitch/stock-analysis./blob/main/Resources/2017/AllStocks2017Results.png)


  

## 2018 Stock Analysis

Below you will see a much different picture. As you can see, 2018 was a terrible year for the green energy stocks we are looking at. All but two are in the red this time! This means that only two of the twelve green energy stocks had a positive return this year. Not looking good at all!

  

#### 2018 Stock Option Results

![2018 Stock Option Results](https://github.com/StickySitch/stock-analysis./blob/main/Resources/2018/AllStocks2018Results.png)

  

## The Code

Like I mentioned earlier, Steve wants to expand his stock option research by venturing into other areas. To do this, I needed to refactor our old “AllStockAnalysis” subroutine to be more efficient.  
  
  
### The original code  
When working with the original data set (green energy stocks), we used a ```nested for loop``` first to create a variable called ```ticker``` that is used as a reference for the current ticker we are gathering data for.

  
  

```VBA

'Instantiating ticker array and assigning ticker values

Dim tickers(11) As String

tickers(0) = "AY"

tickers(1) = "CSIQ"

tickers(2) = "DQ"

tickers(3) = "ENPH"

tickers(4) = "FSLR"

tickers(5) = "HASI"

tickers(6) = "JKS"

tickers(7) = "RUN"

tickers(8) = "SEDG"

tickers(9) = "SPWR"

tickers(10) = "TERP"

tickers(11) = "VSLR"

'Instantiating starting price and ending price

Dim startingPrice As Single

Dim endingPrice As Single

'Activate data worksheet

Worksheets(yearValue).Activate

'Finding number of populated rows

rowEnd = ActiveSheet.UsedRange.Rows.Count

  

'Looping through tickers

For i = 0 To 11

ticker = tickers(i)

totalVolume = 0

'Activating correct worksheet

Worksheets(yearValue).Activate

  

'Looping through rows in the data

For j = 2 To rowEnd

'Getting total volume for current ticker

If Cells(j, 1).Value = ticker Then

totalVolume = totalVolume + Cells(j, 8).Value

End If

'Getting starting price for current ticker

If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

startingPrice = Cells(j, 6).Value

End If

'Getting ending price for current ticker

If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

endingPrice = Cells(j, 6).Value

End If

Next j

'Outputting data for current ticker

Worksheets("All Stocks Analysis").Activate

Cells(4 + i, 1).Value = ticker

Cells(4 + i, 2).Value = totalVolume

Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

Next i 
```

  

#### The Inner Loop

The inner loop is where the bulk of the work is happening; Going row by row, checking for the current ticker symbol using the ```ticker``` reference mentioned earlier. Once the ticker is found the following data is collected: Trade volumes, the years starting price, and ending price. Once the data is collected for the reference ticker, the inner loop is left and the subroutine continues onto the calculations. The calculations will populate the correct cells and return the following: Total trade volume and the years return percentage. The subroutine will continue going through each ticker in the array collecting and displaying the same information.

  

### The Refactored Code

  

In order to handle the much larger data set, I have made some changes to the code. Instead of using a ```nested loop```, this time I have separate loops. To keep track of the data I have created a ```tickerIndex``` variable, letting our loop know what ticker value we are looking at. Along with this, I have instantiated 3 output arrays: ```tickerVolumes```, ```tickerStartingPrice```, and ```tickerEndingPrice```. With these output arrays, data can be assigned to the correct ticker quickly and efficiently.  
  
#### The First Loop

  

Below you can see a snippet of the first loop.

  

```VBA

'Looping through "tickerVolumes" array and initializing array variables to 0

For i = 0 To 11

'Resizing "tickerVolumes" array to the correct size

ReDim tickerVolumes(i)

'Initializing "tickerVolumes" to 0

tickerVolumes(i) = 0

Next i

```

In our first loop we are simply assigning a value of 0 to each of the tickers corresponding ```tickerVolumes``` variable. This is to make sure our data is accurate when collected.

  
  

#### The Second Loop

  

This is where the data collection begins!

  

```VBA

'Loop over all the rows in the spreadsheet.

For i = 2 To RowCount

'Increase volume traded for current ticker in "tickers" array

tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

'Check if the current row is the first row with the selected tickerIndex.

If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

'Initializes "tickerStartingPrice" to the tickers starting price for the year

tickerStartingPrice(tickerIndex) = Cells(i, 6).Value

End If

'Checks if the current row is the last row with the selected ticker

If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

'Initializes "tickerEndingPrice" to the tickers ending price for the year

tickerEndingPrice(tickerIndex) = Cells(i, 6).Value

'Increases the tickerIndex

tickerIndex = tickerIndex + 1

End If

Next i

```

  

Above you can see the full loop. When going through it, you can see it is very similar to the ```nested loop``` we saw earlier. The big difference between these is the output arrays and ```tickerIndex```. The output arrays and ```tickerIndex``` gives us the ability to gather data much faster and make our variables more dynamic and efficient.

  

```tickerIndex``` is increased by 1, moving onto the next ticker, once the final pieces of data have been collected from the current tickers iteration.

  
  

#### The Third Loop

  

The third loop is where the calculations and the population of cells is done.

  

```VBA

'Loops through your arrays to output the Ticker, Total Daily Trade Volume, and Return.

For i = 0 To 11

'Activating the correct output worksheet

Worksheets("All Stocks Analysis").Activate

tickerIndex = i

'Populating cells with ticker symbols, Ticker Daily Trade Volumes, and yearly return

Cells(4 + i, 1).Value = tickers(tickerIndex)

Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)

Cells(4 + i, 3).Value = tickerEndingPrice(tickerIndex) / tickerStartingPrice(tickerIndex) - 1

Next i
```

  

Above you can see the loop going through each ticker in the ```tickers``` array, using the ```tickerIndex``` as a reference. While going through each ticker, the following calculations are being made: ```Total daily trade volume``` and ```yearly return percentage```.

  

#### Formatting

  

To make everything look nice and readable I’ve gone ahead and added formatting to the subroutine.

  

##### Font and Number Formatting

```VBA

Worksheets("All Stocks Analysis").Activate

'Formatting columns

'Making column headers bold

Range("A3:C3").Font.FontStyle = "Bold"

'Underlining the headers

Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous

'Formatting Total Daily Volume

Range("B4:B15").NumberFormat = "#,##0"

'Formatting return percentage

Range("C4:C15").NumberFormat = "0.0%"

Columns("B").AutoFit

```

Above is the code used to produce some basic formatting. I formatted the headers by making them bold, along with adding an underline. After that, you can see some number formatting. The number formatting changes the ***Return*** columns values to percentages and adds comma separation to the ***Total Daily Volume*** column.

  

#### Cell Formatting

Last but not least, a little cell formatting. The code below is simple; If the yearly return value is **negative**, the cell turns **red**. If the return is **positive**, the cell will turn **green**.

```VBA

dataRowStart = 4

dataRowEnd = 15

'Changing return color format

For i = dataRowStart To dataRowEnd

If Cells(i, 3) > 0 Then

'Makes cell green if the return percentage is positive

Cells(i, 3).Interior.Color = vbGreen

Else

'Makes cell green if the return percentage is positive

Cells(i, 3).Interior.Color = vbRed

End If

Next i

```

  

## Performance

### Original Code

As seen below, the performance of the subroutine is average at best! The analysis alone for 2017 took ```1.31 seconds```, and 2018 took ```1.42 seconds``` to complete. For there only being 12 stock options we are gathering data for, this is quite slow.  
  
![2017 OC Performance](https://github.com/StickySitch/stock-analysis./blob/main/Resources/2017/AllStocks2017Performance.png)

  

![2018 OC Performance](https://github.com/StickySitch/stock-analysis./blob/main/Resources/2018/AllStocks2018Performance.png)

  
  

### Refactored Code

With the new way of handling data by using arrays, our performance has increased greatly! Below you can see that there is about ```1.2 seconds``` shaved off of the completion time for both 2017 and 2018.  
  
![2017 Refact Performance](https://github.com/StickySitch/stock-analysis./blob/main/Resources/2017/Refactored2017Performance.png)

  

![2018 Refact Performance](https://github.com/StickySitch/stock-analysis./blob/main/Resources/2018/Refactored2018Performance.png)

  

Something of note is that even though I added formatting to the refactored code, our performance is still much better than the original subroutine.

  

# Summary

To summarize all of this, I think it is important to note the advantages of refactoring. Clearly there has been a huge improvement in performance. Refactoring in general provides a great opportunity to increase the efficiency of the code which in return decreases the time it takes to collect and calculate data. Time is money! Careful though, refactoring also opens the opportunity for mistakes. These mistakes can cause stability issues making the code worse than when it started.

  

### Original vs Refactored

When looking at the original subroutines performance, we can see a huge disadvantage in speed. The way data is handled causes the subroutine to perform at a much slower rate than the refactored version. There are no real “advantages” for the original code except that it completes the job.  
  
Our refactored code showed much more promise! With our new output arrays and ```tickerIndex```, we are able to handle data much more efficiently. This is where our performance increase comes in. Talk about an advantage!  
  
A big disadvantage for both of these subroutines is the lack of ability for the user to easily add new ticker symbols for data collection.
