This folder contains an example of using TD Ameritrade’s Web API (see developer.tdameritrade.com) to maintain and update a database of 
end-of-day stock prices.  I am writing this in case someone finds it helpful since I have combined information that I found from 
various sources.  I personally use the data from the database for the calculations that I do in some of my Excel workbooks.
Disclaimer: This example includes code that I developed for my own personal use and I have included it as a source of information.  
I am not recommending that anyone else use the code as written and I am not responsible for any consequences that result from doing so.  
This post reflects my current understanding and may contain errors.  Since the code was developed for my own use, I did not attempt 
to make it elegant or efficient.

I think that a TD Ameritrade account is required in order to use TD Ameritrade’s Web API. I am using Visual Studio.Net (Visual Basic.Net) 
for the program described here that accesses the Web API and the database.  The database is a Microsoft SQL Server database and I am 
using a Windows 10 computer.  The versions of Visual Studio.Net and Microsoft SQL Server that I am using are free downloads.  I also 
use the Newtonsoft JSON Nuget package to make it easier to decode the JSON that is returned by the Web API.  Also, I set up a local 
Web site (redirect URL) on my local computer to receive the responses from some of the TD Ameritrade Web APIs and Web pages that allow me to 
manually obtain a new refresh token.

This folder contains an image of the input form of the Visual Basic.Net program.  An access token is required in order to obtain data 
from the Web API.  A refresh token is required in order to obtain an access token. An authorization code is required in order to obtain 
the refresh token.   A consumer key is required in order to obtain the authorization code.  A new refresh token needs to be obtained 
every 3 months whereas a new access token only works for half an hour.

Although there is a button on the program’s input form labelled “Get Refresh Token”, I have not created any code for getting the refresh 
token.  Since a new refresh token is only required every 3 months and using the TD Ameritrade’s Web pages to obtain a new refresh token 
requires logging into a TD Ameritrade account to obtain a new authorization code, I thought it was safer from a security standpoint to 
just use TD Ameritrade’s Web pages.  This is a little tedious since the authorization code that is returned by one Web page has to the 
URL decoded and then entered into another Web page (I use a small Visual Basic.Net program to do the URL decode).  I saved the refresh 
token and consumer key in separate text files for use by the program.

In the program, the button labelled “Get Access Token” calls a subroutine that obtains a new access token for the TD Ameritrade API. 
The subroutine reads the consumer key and refresh token from separate text input files and stores the returned access token in a 
different text file.  The access token that is returned by the Web API is output to a text file for later use by the other subroutines.  
A new access token needs to be obtained first for the download subroutines to work unless it has been less than a half hour since the 
last access token was obtained.

The buttons that are labelled “Download historical data for database” call a subroutine that downloads the end-of-day stock prices for 
a list of ticker symbols.  The end-of-day stock price data is downloaded to a different CSV (comma separated value) file for each ticker 
symbol.  There are 2 buttons because I use 2 different lists of ticker symbols named “ticker_list.txt” and “ticker_list1.txt”.  The 
subroutine checks the “market_price” table in the SQL Server “market_data” database to see whether there is already end-of-day stock 
price data for that ticker symbol in the database.  If end-of-day stock price data already exists, then the subroutine just downloads 
the most recent data up to the current day.  The data is downloaded starting 4 days before the current day in case some recent values 
have changed. If there is no end-of-day stock price data for the ticker symbol in the database, then all of the data up to the current 
day is downloaded to a CSV file.  After downloading the data for 116 ticker symbols, the subroutine waits for one minute 20 seconds 
because the TD Ameritrade API returns an error if more than 120 requests are made in one minute (and I allowed some extra wiggle room).  
The inputs to the subroutine are the filenames of the text files for the consumer key, access token and ticker list, the name of the 
folder where the output CSV files will be stored, and the database data source name.  The database data source name is the name listed 
as the server name for the database engine under Microsoft SQL Server Management Studio.  I’m not sure that the date conversion to the 
number of ticks used by the TD Ameritrade Web APIs is exactly accurate but it is close enough to use for end-of-day stock prices.  
Since the end-of-day data is used, the program is normally executed after the stock market closes.

I use the subroutine UpdateTickerList to construct a new ticker_list.txt file by reading the ticker symbols for the setup sheets of
the Excel workbooks that are listed in the indicator_files.txt. So indicator_files.txt and the subroutine UpdateTickerList are probably 
not needed by someone else.

The buttons that are labelled “Update database” read the CSV file for each ticker symbol and add the end-of-day stock price data to the 
market_price table in the database.  The inputs to the subroutine are the name of the folder where the output CSV files will be stored 
and the database data source name.

There are also buttons labelled “Download fundamental data for database” that download fundamental data for each ticker symbol.  
There are corresponding buttons labelled “Update Database” that store the fundamental data in the “fundamentals” table in the database.  
I don’t currently make use of the fundamental data in my Excel workbooks and the subroutines are similar to subroutines that I described 
above so I’m not going to give a detailed explanation of these subroutines but I have included the code.

The user needs to specify the data source name, the folder locations and the text file names either in the InitializeDefaults subroutine 
or else in the external text file “GetStockData.ini” that is read by the ReadDefaults subroutine if it exists.

Below, I have included information about the structure of the database tables. I have also included an Excel VBA function that shows 
how I read the end-of-day stock prices from the database data into Microsoft Excel.  I normally use the last 120 market days in my 
Excel calculations so that errors in calculations (such as exponential moving averages) will have time to die out.

The Excel VBA function that I use to read the end-of-day stock prices from the database data into Microsoft Excel is below.

Function UpdateWorkSheetFromDatabase%(DataSource$, NumTickers%, tickers$(), NumRowsPerTicker%(), StartRow&, oSheet As Worksheet)
  UpdateWorkSheetFromDatabase = 0
  Dim cn As ADODB.Connection
  Dim rst As ADODB.Recordset
  Dim cmd As ADODB.Command
  Dim i%, j&, RowOffset&
  Dim date1$, open1#, high1#, low1#, close1#, volume1&
  Dim year1$, month1$, day1$, s1$, s2$, msg$
 
  ' Open the connection.
  Dim ConnectionString$
  Set cn = New ADODB.Connection
  ConnectionString = "Provider='SQLOLEDB';Data Source='" & DataSource & "';Initial Catalog='market_data';Integrated Security='SSPI';"
  cn.Open ConnectionString

  RowOffset = 0
  For i% = 1 To NumTickers%
    s1 = tickers(i)
    
    ' Set the command text.
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    With cmd
      ' I want the BOTTOM records of the original table; fortunately I also want them in descending order so I don't need the line that I commented out
      '.CommandText = "SELECT * FROM (SELECT TOP " & Trim$(Str$(num%)) & " * FROM market_price t1 WHERE Ticker='" & s1 & "' ORDER BY t1.Date DESC) t2 ORDER BY t2.Date ASC"
     .CommandText = "Select Top " & Trim$(Str$(NumRowsPerTicker(i))) & " * from market_price where Ticker='" & s1 & "' Order By Date DESC"
     .CommandType = adCmdText
     .Execute
    End With
 
    ' Open the recordset.
    Set rst = New ADODB.Recordset
    Set rst.ActiveConnection = cn
    rst.Source = "market_price"
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.Open cmd
    Dim num_records&
    num_records = rst.RecordCount
    If num_records <> NumRowsPerTicker(i) Then
      rst.Close
      cn.Close
      Set cmd = Nothing
      Set rst = Nothing
      Set cn = Nothing
      msg = "Number of records <> " & Trim$(Str$(NumRowsPerTicker(i))) & " for ticker " & tickers(i)
      MsgBox msg
      Exit Function
    End If

    
    rst.MoveFirst
    For j = 1 To NumRowsPerTicker(i)
      date1 = Trim$(Str$(rst.Fields.Item("Date")))
      open1 = rst.Fields.Item("Open")
      high1 = rst.Fields.Item("High")
      low1 = rst.Fields.Item("Low")
      close1 = rst.Fields.Item("Close")
      volume1 = rst.Fields.Item("Volume")
      
      s2 = ""
      If Len(date1) = 8 Then
        year1 = Mid$(date1, 1, 4)
        month1 = Mid$(date1, 5, 2)
        If Mid$(month1, 1, 1) = "0" Then month1 = Mid$(month1, 2, 1)
        day1 = Mid$(date1, 7, 2)
        If Mid$(day1, 1, 1) = "0" Then day1 = Mid$(day1, 2, 1)
        s2 = month1 & "/" & day1 & "/" & year1
      End If
      
      oSheetList.Cells(StartRow + RowOffset + j - 1&, 1).Value = s1
      oSheetList.Cells(StartRow + RowOffset + j - 1&, 2).Value = s2
      oSheetList.Cells(StartRow + RowOffset + j - 1&, 3).Value = open1
      oSheetList.Cells(StartRow + RowOffset + j - 1&, 4).Value = high1
      oSheetList.Cells(StartRow + RowOffset + j - 1&, 5).Value = low1
      oSheetList.Cells(StartRow + RowOffset + j - 1&, 6).Value = close1
      oSheetList.Cells(StartRow + RowOffset + j - 1&, 7).Value = volume1
      rst.MoveNext
    Next
    RowOffset = RowOffset + NumRowsPerTicker%(i)
  Next

  ' Close the connections and clean up.
  rst.Close
  cn.Close
  Set cmd = Nothing
  Set rst = Nothing
  Set cn = Nothing
  Exit Function

ErrorHandler:
  UpdateWorkSheetFromDatabase = -1
  msg = err.Source & ": Error " & err.Number & ": " & err.Description
  err.Clear
  MsgBox msg
End Function

The structure of the database tables is below.
Table dbo.market_price
COLUMN_NAME,IS_NULLABLE,DATA_TYPE,NUMERIC_PRECISION,CHARACTER_MAXIMUM_LENGTH
Ticker,NO,varchar,NULL,10
Date,NO,int,10,NULL
Open,NO,decimal,18,NULL
High,NO,decimal,18,NULL
Low,NO,decimal,18,NULL
Close,NO,decimal,18,NULL
Volume,NO,bigint,19,NULL

Table dbo.fundamentals
COLUMN_NAME,IS_NULLABLE,DATA_TYPE,NUMERIC_PRECISION,CHARACTER_MAXIMUM_LENGTH
ticker,NO,varchar,NULL,10
high52,YES,decimal,18,NULL
low52,YES,decimal,18,NULL
dividendAmount,YES,decimal,18,NULL
dividendYield,YES,decimal,18,NULL
dividendDate,YES,varchar,NULL,30
peRatio,YES,decimal,18,NULL
pegRatio,YES,decimal,18,NULL
pbRatio,YES,decimal,18,NULL
prRatio,YES,decimal,18,NULL
pcfRatio,YES,decimal,18,NULL
grossMarginTTM,YES,decimal,18,NULL
grossMarginMRQ,YES,decimal,18,NULL
netProfitMarginTTM,YES,decimal,18,NULL
netProfitMarginMRQ,YES,decimal,18,NULL
operatingMarginTTM,YES,decimal,18,NULL
operatingMarginMRQ,YES,decimal,18,NULL
returnOnEquity,YES,decimal,18,NULL
returnOnAssets,YES,decimal,18,NULL
returnOnInvestment,YES,decimal,18,NULL
quickRatio,YES,decimal,18,NULL
currentRatio,YES,decimal,18,NULL
interestCoverage,YES,decimal,18,NULL
totalDebtToCapital,YES,decimal,18,NULL
ltDebtToEquity,YES,decimal,18,NULL
totalDebtToEquity,YES,decimal,18,NULL
epsTTM,YES,decimal,18,NULL
epsChangePercentTTM,YES,decimal,18,NULL
epsChangeYear,YES,decimal,18,NULL
epsChange,YES,decimal,18,NULL
revChangeYear,YES,decimal,18,NULL
revChangeTTM,YES,decimal,18,NULL
revChangeIn,YES,decimal,18,NULL
sharesOutstanding,YES,decimal,18,NULL
marketCapFloat,YES,decimal,18,NULL
marketCap,YES,decimal,18,NULL
bookValuePerShare,YES,decimal,18,NULL
shortIntToFloat,YES,decimal,18,NULL
shortIntDayToCover,YES,decimal,18,NULL
divGrowthRate3Year,YES,decimal,18,NULL
dividendPayAmount,YES,decimal,18,NULL
dividendPayDate,YES,varchar,NULL,30
beta,YES,decimal,18,NULL
vol1DayAvg,YES,decimal,18,NULL
vol10DayAvg,YES,decimal,18,NULL
vol3MonthAvg,YES,decimal,18,NULL
cusip,YES,varchar,NULL,30
description,YES,varchar,NULL,100
exchange,YES,varchar,NULL,30
assetType,YES,varchar,NULL,30

The database views used in the program are below.

View dbo.get_last_date
SELECT        Ticker, MAX(Date) AS Last_Date, COUNT(*) AS Num_of_records
FROM            dbo.market_price AS mp
GROUP BY Ticker

View dbo.get_fundamental_field_names
SELECT        COLUMN_NAME
FROM            INFORMATION_SCHEMA.COLUMNS
WHERE        (TABLE_NAME = N'fundamentals')
