' Updated on 30Jun21 to increase array sizes from 300 to 500 for the maximum number of ticker symbols in the database.
' Modified on 7Aug21 to add downoad of fundamental data.
' Modified on 15Sep21 to use structure INPUTTYPE
' Modified on 3Oct21 to remove some unnecessary code in UpdateTickerList.
' modified on 13Sep22 to allow the user to enter or select the ticker symbol list file names
' in the form.
' Modified on 28Jun23 to handle a problem adding the fundamental data for symbol "$VIX.X" to the database because it did not have a cusip value.
' I modified the code slightly to handle the case where a database table field (such as cusip) is missing from the fundamental input data.
' Modified on 5Jul23 to check that arrays sizes are not exceeded. This was known to be possible but the array sizes were large enough that
' I did not expect them to be exceeded for the number of stock symbols that I expected to use. Note that items that exceed the array sizes
' will simply be ignored; there is no error message.
' Modified on 4Jan24 to use the Polygon.io API as an alternative to the TD Ameritrade API
' Modified on 14Jan24 to add a menu item for importing a Yahoo file exported from a portfolio in order to add the OHLC values for one day to be database.
' Last updated on 15Jan24.


Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Net
Imports System.Web
Imports System.Globalization
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.office.interop
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Runtime.InteropServices.ComTypes

Structure INPUTTYPE
  Dim indicator_file$
  Dim data_source$
  Dim ticker_list_file$
  Dim response_folder$
  Dim csv_folder$
  Dim ticker_list1_file$
  Dim response1_folder$
  Dim csv1_folder$
  Dim fundamental_response_folder$
  Dim fundamental_response1_folder$
  Dim trading_API$
  ' TD Ameritrade API
  Dim file_consumer_key$
  Dim file_refresh_token$
  Dim file_access_token_response$
  Dim file_access_token$
  ' Polygon.io API
  Dim file_polygon_io_api_key$
  ' Yahoo file
  Dim yahoo_file$
  Dim yahoo1_file$
End Structure

Module Module1
  Public client As New Net.Http.HttpClient()
  Public epoch_UTC& = 62135596800000& ' (1/1/1970 12 AM as UTC ticks) / 10000
  Public UserInput As INPUTTYPE

  Sub InitializeDefaults()
    With UserInput
      .indicator_file = "c:\indicator_files.txt"
      .data_source = "your data source name goes here"
      .ticker_list_file = "c:\ticker_list.txt"
      .response_folder = "c:\download_response"
      .csv_folder = "c:\download"
      .ticker_list1_file = "c:\ticker_list1.txt"
      .response1_folder = "c:\download_response1"
      .csv1_folder = "c:\download1"
      .fundamental_response_folder = "c:\download_fundamental_response"
      .fundamental_response1_folder = "c:\download_fundamental_response1"
      .trading_API = "TD Ameritrade"
      ' TD Ameritrade API
      .file_consumer_key = "c:\consumer_key.txt"
      .file_refresh_token = "c:\refresh_token.txt"
      .file_access_token_response = "c:\access_token_response.txt"
      .file_access_token = "c:\access_token.txt"
      ' Polygon.io API
      .file_polygon_io_api_key = "c:\polygon_io_api_key.txt"
      'Yahoo file
      .yahoo_file = "c:\yahoo_file.csv"
      .yahoo1_file = "c:\yahoo1_file.csv"
    End With
  End Sub

  Function ReadDefaults(ByVal sFileName$)
    ReadDefaults = 0
    InitializeDefaults()
    If (Dir(sFileName$) = "") Then Exit Function
    If Not File.Exists(sFileName) Then Exit Function
    Dim line$
    ReadDefaults = -1
    line = ""

    Try
      Dim reader As New StreamReader(sFileName)
      With UserInput
        While (Not reader.EndOfStream)
          line = reader.ReadLine()
          If (line Is Nothing) Then Exit Function
          line = line.Trim
          If line.Length <= 0 Then Exit Function
          Dim items = line.Split(",")
          Select Case (Trim$(items(0)))
            Case "indicator_file"
              .indicator_file = items(1)
            Case "data_source"
              .data_source = items(1)
            Case "ticker_list_file"
              .ticker_list_file = items(1)
            Case "response_folder"
              .response_folder = items(1)
            Case "csv_folder"
              .csv_folder = items(1)
            Case "ticker_list1_file"
              .ticker_list1_file = items(1)
            Case "response1_folder"
              .response1_folder = items(1)
            Case "csv1_folder"
              .csv1_folder = items(1)
            Case "fundamental_response_folder"
              .fundamental_response_folder = items(1)
            Case "fundamental_response1_folder"
              .fundamental_response1_folder = items(1)
            Case "trading_API"
              .trading_API = items(1)
            'TD Ameritrade API
            Case "file_consumer_key"
              .file_consumer_key = items(1)
            Case "file_refresh_token"
              .file_refresh_token = items(1)
            Case "file_access_token_response"
              .file_access_token_response = items(1)
            Case "file_access_token"
              .file_access_token = items(1)
            'Polygon.io API
            Case "file_polygon_io_api_key"
              .file_polygon_io_api_key = items(1)
            'Yahoo file
            Case "yahoo_file"
              .yahoo_file = items(1)
            Case "yahoo1_file"
              .yahoo1_file = items(1)
          End Select
        End While
      End With
      reader.Close()
    Catch e As Exception
      MessageBox.Show("Error in file " & sFileName & ": " & e.Message)
      ReadDefaults = -2
      Exit Function
    End Try
    ReadDefaults = 0
  End Function
  Function ReadEntries(ByVal sFileName$)
    ReadEntries = 0
    If (Dir(sFileName$) = "") Then Exit Function
    If Not File.Exists(sFileName) Then Exit Function
    Dim line$
    ReadEntries = -1
    line = ""

    Try
      Dim reader As New StreamReader(sFileName)
      With UserInput
        While (Not reader.EndOfStream)
          line = reader.ReadLine()
          If (line Is Nothing) Then Exit Function
          line = line.Trim
          If line.Length <= 0 Then Exit Function
          Dim items = line.Split(",")
          Select Case (Trim$(items(0)))
            Case "ticker_list_file"
              .ticker_list_file = items(1)
            Case "ticker_list1_file"
              .ticker_list1_file = items(1)
            Case "trading_API"
              .trading_API = items(1)
            'Yahoo file
            Case "yahoo_file"
              .yahoo_file = items(1)
            Case "yahoo1_file"
              .yahoo1_file = items(1)
          End Select
        End While
      End With
      reader.Close()
    Catch e As Exception
      MessageBox.Show("Error in file " & sFileName & ": " & e.Message)
      ReadEntries = -2
      Exit Function
    End Try
    ReadEntries = 0
  End Function
  Function SaveEntries(ByVal sFileName$)
    SaveEntries = -1
    If File.Exists(sFileName) Then File.Delete(sFileName)
    Try
      Dim writer1 As New StreamWriter(sFileName)
      With UserInput
        writer1.WriteLine("ticker_list_file," & .ticker_list_file.Trim)
        writer1.WriteLine("ticker_list1_file," & .ticker_list1_file.Trim)
        writer1.WriteLine("trading_API," & .trading_API.Trim)
        'Yahoo file
        writer1.WriteLine("yahoo_file," & .yahoo_file.Trim)
        writer1.WriteLine("yahoo1_file," & .yahoo1_file.Trim)
      End With
      writer1.Close()
    Catch e As Exception
      MessageBox.Show("Error writing file " & sFileName & ": " & e.Message)
      SaveEntries = -2
      Exit Function
    End Try
    SaveEntries = 0
  End Function

  Async Sub GetAccessCode(file_consumer_key$, file_refresh_token$, file_access_token_response$, file_access_token$)
    Dim refresh_token$, consumer_key$, err%
    Dim access_token$

    refresh_token = ""
    consumer_key = ""
    err = ReadTextFromFile(file_consumer_key, consumer_key)
    If err < 0 Then Exit Sub
    err = ReadTextFromFile(file_refresh_token, refresh_token)
    If err < 0 Then Exit Sub

    Dim url = "https://api.tdameritrade.com/v1/oauth2/token"

    'Dim contentType$ = "application/json"
    Dim data1 As New Dictionary(Of String, String)
    data1.Add("grant_type", "refresh_token")
    data1.Add("refresh_token", refresh_token)
    data1.Add("client_id", consumer_key & "@AMER.OAUTHAP")
    'data1.Add("redirect_uri", redirect_url)

    Dim content As New Http.FormUrlEncodedContent(data1)
    Dim result$
    'Dim request As New HttpRequestHeader()
    'client.DefaultRequestHeaders.Add("Authorization", "Bearer " & authstring)
    client.DefaultRequestHeaders.Clear()
    Dim response As Http.HttpResponseMessage
    response = Await client.PostAsync(url, content)
    ' will throw an exception if Not successful
    response.EnsureSuccessStatusCode()
    result = response.Content.ReadAsStringAsync().Result

    If File.Exists(file_access_token_response) Then File.Delete(file_access_token_response)
    Dim writer As New StreamWriter(file_access_token_response)
    writer.Write(result)
    writer.Close()

    Dim obj As New JObject()
    obj = JObject.Parse(result)
    access_token = obj("access_token")

    If File.Exists(file_access_token) Then File.Delete(file_access_token)
    Dim writer1 As New StreamWriter(file_access_token)
    writer1.WriteLine(access_token)
    writer1.Close()
  End Sub
  Async Sub GetRefreshCode()

  End Sub

  Async Sub DownloadHistData(file_consumer_key$, file_access_token$, ticker_list_file$, response_folder$, csv_folder$, data_source$)
    Dim consumer_key$, access_token$, content$
    Dim file_price_history_response$, file_price_history$
    Dim err%, i%, n%, count%, limit%, limit1%, s$, s1$, s2$, s3$
    Dim url$, year1$, month1$, day1$

    limit = 299
    limit1 = 500
    access_token = ""
    consumer_key = ""

    err = ReadTextFromFile(file_consumer_key, consumer_key)
    If err < 0 Then Exit Sub
    err = ReadTextFromFile(file_access_token, access_token)
    If err < 0 Then Exit Sub

    Dim tickers$(0 To limit), num_tickers%, line$
    num_tickers = 0
    Try
      Dim reader As New StreamReader(ticker_list_file)
      While Not reader.EndOfStream And num_tickers <= limit
        line = reader.ReadLine()
        If (line Is Nothing) Then Exit While
        line = line.Trim.ToUpper
        If line.Length > 0 Then
          'Dim items = line.Split(",")
          tickers(num_tickers) = line
          num_tickers = num_tickers + 1
        End If
      End While
      reader.Close()
    Catch e As Exception
      MessageBox.Show("Error in file " & ticker_list_file & ": " & e.Message)
      Exit Sub
    End Try

    If num_tickers <= 0 Then
      MessageBox.Show("No ticker symbols found in ticker symbol list file")
    End If

    Dim tickers_db$(0 To limit1), dates1$(0 To limit1), num_tickers_db%
    Dim market_price_db$ = "Data Source=" & data_source & ";Initial Catalog=market_data;Integrated Security=True;"
    Dim cn As New SqlConnection() ' Don't put this statement in a try block; it throws an exception!!!
    cn.ConnectionString = market_price_db
    cn.Open()

    Dim cmd As New SqlCommand, dr As SqlDataReader
    Try
      'Dim cn As New SqlConnection(market_price_db)

      cmd.Connection = cn
      cmd.CommandText = "Select * from dbo.[Get_Last_Date]"
      dr = cmd.ExecuteReader
      num_tickers_db = 0
      While dr.Read() And num_tickers_db <= limit1
        tickers_db(num_tickers_db) = dr("Ticker")
        dates1(num_tickers_db) = dr("Last_Date")
        num_tickers_db = num_tickers_db + 1
      End While
      dr.Close()
      cmd.Dispose()
      cn.Close()
    Catch e As Exception
      cmd.Dispose()
      cn.Close()
      MessageBox.Show(e.Message)
      Exit Sub
    End Try

    ReDim Preserve tickers_db$(0 To num_tickers_db - 1)
    ReDim Preserve dates1$(0 To num_tickers_db - 1)
    Array.Sort(tickers_db, dates1)

    Dim fileEntries As String() = Directory.GetFiles(response_folder)
    Dim fileName As String
    For Each fileName In fileEntries
      File.Delete(fileName)
    Next fileName
    Dim fileEntries1 As String() = Directory.GetFiles(csv_folder)
    For Each fileName In fileEntries1
      File.Delete(fileName)
    Next fileName

    Dim num_ticks&
    Dim estZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")
    Dim CurrentDate = Date.Now()
    Dim CurrentDate1 As Date = CurrentDate.Date
    Dim TickDate As Date = CurrentDate1.AddHours(20) ' add 20 hours to make it beyond closing time
    Dim TickDate1 As Date = Date.SpecifyKind(TickDate, DateTimeKind.Unspecified)
    Dim objUTC = TimeZoneInfo.ConvertTimeToUtc(TickDate1, estZone)
    num_ticks = objUTC.Ticks
    num_ticks = (num_ticks / 10000&) - epoch_UTC
    s3 = Trim(num_ticks.ToString)

    Dim num_requests%, date2$
    client.DefaultRequestHeaders.Clear()
    client.DefaultRequestHeaders.Add("Authorization", "Bearer " & access_token)
    Dim TickDate2 As Date
    num_requests = 0
    For i = 0 To num_tickers - 1
      Application.DoEvents()
      num_requests = num_requests + 1
      If num_requests > 116 Then
        Thread.Sleep(80000) ' pause for 1 minute 20 seconds
        num_requests = 1
      End If

      If num_tickers_db <= 0 Then
        date2 = "0"
      Else
        Dim index1% = Array.BinarySearch(tickers_db, tickers(i))
        If index1 >= 0 Then
          date2 = dates1(index1)
        Else
          date2 = "0"
        End If
      End If

      If date2 = "0" Then
        num_ticks = 0
      Else
        year1 = CInt(Mid$(date2, 1, 4))
        month1 = CInt(Mid$(date2, 5, 2))
        day1 = CInt(Mid$(date2, 7, 2))
        Dim d As New Date(year1, month1, day1)
        TickDate = d.AddDays(-4)
        'TickDate1 = TickDate.AddHours(2)
        TickDate2 = Date.SpecifyKind(TickDate, DateTimeKind.Unspecified)
        objUTC = TimeZoneInfo.ConvertTimeToUtc(TickDate2, estZone)
        num_ticks = objUTC.Ticks
        num_ticks = (num_ticks / 10000&) - epoch_UTC
      End If
      s2 = Trim(num_ticks.ToString)
      url = "https://api.tdameritrade.com/v1/marketdata/" & tickers(i) & "/pricehistory?apikey=" & consumer_key &
                "&periodType=month&frequencyType=daily&startDate=" & s2 & "&endDate=" & s3 & "&needExtendedHoursData=false"
      ' "&periodType=month&frequencyType=daily&endDate=1595221200000"
      ' "&periodType=month&frequencyType=daily&frequency=1"
      '  "&periodType=month&frequencyType=daily&endDate=1595048400000"
      '  "&periodType=month&frequencyType=daily&frequency=1"
      '"&periodType=month&period=7&frequencyType=daily&endDate=" & s1

      Dim response As Http.HttpResponseMessage
      response = Await client.GetAsync(url)
      ' will throw an exception if Not successful
      response.EnsureSuccessStatusCode()
      content = Await response.Content.ReadAsStringAsync()
      file_price_history_response = response_folder & "\" & tickers(i) & "_response.txt"
      file_price_history = csv_folder & "\" & tickers(i) & ".csv"
      If File.Exists(file_price_history_response) Then File.Delete(file_price_history_response)
      Dim writer As New StreamWriter(file_price_history_response)
      writer.Write(content)
      writer.Close()

      Dim date1$, open1$, high1$, low1$, close1$, volume1$
      Dim i32 As Int32
      Dim jss = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Object)(content)
      n = jss("candles").count()
      If n > 0 Then
        If File.Exists(file_price_history) Then File.Delete(file_price_history)
        Dim writer1 As New StreamWriter(file_price_history)
        writer1.WriteLine("rows")
        writer1.WriteLine(Trim$(n.ToString))
        s = "Date,Open,High,Low,Close,Volume"
        writer1.WriteLine(s)
        For i32 = 0 To n - 1
          date1 = jss("candles")(i32)("datetime").ToString
          num_ticks = (CLng(date1) + epoch_UTC) * 10000&
          Dim timeUtc As New DateTime(num_ticks, DateTimeKind.Utc)
          Dim estTime As Date = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, estZone)
          's1 = estTime.ToString("d", CultureInfo.CreateSpecificCulture("en-US"))
          Dim iMonth%, iDay%
          year1 = Trim(estTime.Year.ToString)
          iMonth = estTime.Month
          iDay = estTime.Day
          If iMonth <= 9 Then
            month1 = "0" & Trim(iMonth.ToString)
          Else
            month1 = Trim(iMonth.ToString)
          End If
          If iDay <= 9 Then
            day1 = "0" & Trim(iDay.ToString)
          Else
            day1 = Trim(iDay.ToString)
          End If
          s1 = year1 & month1 & day1
          open1 = jss("candles")(i32)("open").ToString
          high1 = jss("candles")(i32)("high").ToString
          low1 = jss("candles")(i32)("low").ToString
          close1 = jss("candles")(i32)("close").ToString
          volume1 = jss("candles")(i32)("volume").ToString
          s = s1 & "," & open1 & "," & high1 & "," & low1 & "," & close1 & "," & volume1
          writer1.WriteLine(s)
        Next
        writer1.Close()
        count = i + 1
        Form1.lblCount.Text = Trim(count.ToString)
      End If
    Next
    MessageBox.Show("Download finished")
  End Sub

  Async Sub DownloadHistDataPolygonIo(file_polygon_io_API_key$, ticker_list_file$, response_folder$, csv_folder$, data_source$)
    Dim content$, polygon_io_API_key$
    Dim file_price_history_response$, file_price_history$
    Dim err%, i%, n%, count%, s$, s1$, s2$, s3$
    Dim url$, year1$, month1$, day1$

    polygon_io_API_key = ""
    err = ReadTextFromFile(file_polygon_io_API_key, polygon_io_API_key)
    If err < 0 Then Exit Sub

    Dim num_tickers%, line$
    Dim lstTickers As New List(Of String)
    num_tickers = 0
    Try
      Dim reader As New StreamReader(ticker_list_file)
      While Not reader.EndOfStream
        line = reader.ReadLine()
        If (line Is Nothing) Then Exit While
        line = line.Trim.ToUpper
        If line.Length > 0 Then
          'Dim items = line.Split(",")
          lstTickers.Add(line)
          num_tickers = num_tickers + 1
        End If
      End While
      reader.Close()
    Catch e As Exception
      MessageBox.Show("Error in file " & ticker_list_file & ": " & e.Message)
      Exit Sub
    End Try

    If num_tickers <= 0 Then
      MessageBox.Show("No ticker symbols found in ticker symbol list file")
      Exit Sub
    End If
    Dim tickers$() = lstTickers.ToArray

    Dim lstTickersDb, lstDates1 As New List(Of String)
    Dim num_tickers_db%
    Dim market_price_db$ = "Data Source=" & data_source & ";Initial Catalog=market_data;Integrated Security=True;"
    Dim cn As New SqlConnection() ' Don't put this statement in a try block; it throws an exception!!!
    cn.ConnectionString = market_price_db
    cn.Open()

    Dim cmd As New SqlCommand, dr As SqlDataReader
    Try
      'Dim cn As New SqlConnection(market_price_db)

      cmd.Connection = cn
      cmd.CommandText = "Select * from dbo.[Get_Last_Date]"
      dr = cmd.ExecuteReader
      num_tickers_db = 0
      While dr.Read()
        lstTickersDb.Add(dr("Ticker"))
        lstDates1.Add(dr("Last_Date"))
        num_tickers_db = num_tickers_db + 1
      End While
      dr.Close()
      cmd.Dispose()
      cn.Close()
    Catch e As Exception
      cmd.Dispose()
      cn.Close()
      MessageBox.Show(e.Message)
      Exit Sub
    End Try

    If num_tickers_db <= 0 Then
      MessageBox.Show("No ticker symbols found in database")
      Exit Sub
    End If
    Dim tickers_db$() = lstTickersDb.ToArray
    Dim dates1$() = lstDates1.ToArray
    Array.Sort(tickers_db, dates1)

    Dim fileEntries As String() = Directory.GetFiles(response_folder)
    Dim fileName As String
    For Each fileName In fileEntries
      File.Delete(fileName)
    Next fileName
    Dim fileEntries1 As String() = Directory.GetFiles(csv_folder)
    For Each fileName In fileEntries1
      File.Delete(fileName)
    Next fileName

    Dim StartDate As Date
    Dim CurrentDate = Date.Now()
    Dim CurrentDate1 As Date = CurrentDate.Date
    CurrentDate1 = CurrentDate1.AddDays(3.0) ' Add some days
    Dim epoch As New DateTime(1970, 1, 1, 0, 0, 0, 0)

    Dim num_requests%, date2$
    num_requests = 0
    For i = 0 To num_tickers - 1
      Application.DoEvents()
      num_requests = num_requests + 1
      If num_requests > 5 Then
        Thread.Sleep(65000) ' pause for 1 minute 5 seconds
        num_requests = 1
      End If

      If num_tickers_db <= 0 Then
        date2 = "0"
      Else
        Dim index1% = Array.BinarySearch(tickers_db, tickers(i))
        If index1 >= 0 Then
          date2 = dates1(index1)
        Else
          date2 = "0"
        End If
      End If

      If date2 = "0" Then
        s2 = epoch.ToString("yyyy-MM-dd")
      Else
        year1 = CInt(Mid$(date2, 1, 4))
        month1 = CInt(Mid$(date2, 5, 2))
        day1 = CInt(Mid$(date2, 7, 2))
        Dim d As New Date(year1, month1, day1)
        StartDate = d.AddDays(-5.0)  ' Go back 5 days
        s2 = StartDate.ToString("yyyy-MM-dd")
      End If
      s3 = CurrentDate1.ToString("yyyy-MM-dd")
      url = "https://api.polygon.io/v2/aggs/ticker/" & tickers(i) & "/range/1/day/" & s2 & "/" & s3 & "?apiKey=" & polygon_io_API_key

      Dim response As Http.HttpResponseMessage
      response = Await client.GetAsync(url)
      ' will throw an exception if Not successful
      response.EnsureSuccessStatusCode()
      content = Await response.Content.ReadAsStringAsync()
      file_price_history_response = response_folder & "\" & tickers(i) & "_response.txt"
      file_price_history = csv_folder & "\" & tickers(i) & ".csv"
      If File.Exists(file_price_history_response) Then File.Delete(file_price_history_response)
      Dim writer As New StreamWriter(file_price_history_response)
      writer.Write(content)
      writer.Close()

      Dim date1$, open1$, high1$, low1$, close1$, volume1$
      Dim i32 As Int32
      Dim jss = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Object)(content)
      n = jss("results").count()
      If n > 0 Then
        If File.Exists(file_price_history) Then File.Delete(file_price_history)
        Dim writer1 As New StreamWriter(file_price_history)
        writer1.WriteLine("rows")
        writer1.WriteLine(Trim$(n.ToString))
        s = "Date,Open,High,Low,Close,Volume"
        writer1.WriteLine(s)
        Dim t#
        Dim dt As Date
        For i32 = 0 To n - 1
          t = jss("results")(i32)("t").ToString
          dt = epoch
          dt = dt.AddMilliseconds(t)
          date1 = dt.ToString("yyyyMMdd")
          open1 = jss("results")(i32)("o").ToString
          high1 = jss("results")(i32)("h").ToString
          low1 = jss("results")(i32)("l").ToString
          close1 = jss("results")(i32)("c").ToString
          volume1 = jss("results")(i32)("v").ToString
          s = date1 & "," & open1 & "," & high1 & "," & low1 & "," & close1 & "," & volume1
          writer1.WriteLine(s)
        Next
        writer1.Close()
        count = i + 1
        Form1.lblCount.Text = Trim(count.ToString)
      End If
    Next
    MessageBox.Show("Download finished")
  End Sub

  Async Sub DownloadFundamental(file_consumer_key$, file_access_token$, ticker_list_file$, response_folder$, data_source$)
    Dim consumer_key$, access_token$, content$
    Dim file_price_history_response$
    Dim err%, i%, count%, limit%, limit1%
    Dim url$

    limit = 299
    limit1 = 500
    access_token = ""
    consumer_key = ""
    err = ReadTextFromFile(file_consumer_key, consumer_key)
    If err < 0 Then Exit Sub
    err = ReadTextFromFile(file_access_token, access_token)
    If err < 0 Then Exit Sub

    Dim tickers$(0 To limit), num_tickers%, line$
    num_tickers = 0
    Try
      Dim reader As New StreamReader(ticker_list_file)
      While Not reader.EndOfStream And num_tickers <= limit
        line = reader.ReadLine()
        If (line Is Nothing) Then Exit While
        line = line.Trim.ToUpper
        If line.Length > 0 Then
          'Dim items = line.Split(",")
          tickers(num_tickers) = line
          num_tickers = num_tickers + 1
        End If
      End While
      reader.Close()
    Catch e As Exception
      MessageBox.Show("Error in file " & ticker_list_file & ": " & e.Message)
      Exit Sub
    End Try

    If num_tickers <= 0 Then
      MessageBox.Show("No ticker symbols found in ticker symbol list file")
    End If

    Dim tickers_db$(0 To limit1), dates1$(0 To limit1), num_tickers_db%
    Dim market_price_db$ = "Data Source=" & data_source & ";Initial Catalog=market_data;Integrated Security=True;"
    Dim cn As New SqlConnection() ' Don't put this statement in a try block; it throws an exception!!!
    cn.ConnectionString = market_price_db
    cn.Open()

    Dim cmd As New SqlCommand, dr As SqlDataReader
    Try
      'Dim cn As New SqlConnection(market_price_db)

      cmd.Connection = cn
      cmd.CommandText = "SELECT * from dbo.[Get_Last_Date]"
      dr = cmd.ExecuteReader
      num_tickers_db = 0
      While dr.Read() And num_tickers_db <= limit1
        tickers_db(num_tickers_db) = dr("Ticker")
        dates1(num_tickers_db) = dr("Last_Date")
        num_tickers_db = num_tickers_db + 1
      End While
      dr.Close()
      cmd.Dispose()
      cn.Close()
    Catch e As Exception
      cmd.Dispose()
      cn.Close()
      MessageBox.Show(e.Message)
      Exit Sub
    End Try

    ReDim Preserve tickers_db$(0 To num_tickers_db - 1)
    ReDim Preserve dates1$(0 To num_tickers_db - 1)
    Array.Sort(tickers_db, dates1)

    Dim fileEntries As String() = Directory.GetFiles(response_folder)
    Dim fileName As String
    For Each fileName In fileEntries
      File.Delete(fileName)
    Next fileName

    Dim num_requests%
    client.DefaultRequestHeaders.Clear()
    client.DefaultRequestHeaders.Add("Authorization", "Bearer " & access_token)
    num_requests = 0
    For i = 0 To num_tickers - 1
      Application.DoEvents()
      num_requests = num_requests + 1
      If num_requests >= 117 Then
        Thread.Sleep(80000) ' pause for 1 minute 20 seconds
        num_requests = 0
      End If

      'Dim contentType$ = "application/json"
      url = "https://api.tdameritrade.com/v1/instruments?apikey=" & consumer_key &
        "&symbol=" & tickers(i) & "&projection=fundamental"
      'url = "https://api.tdameritrade.com/v1/instruments?symbol=" & tickers(i) & "&projection=symbol-search"

      Dim response As Http.HttpResponseMessage
      response = Await client.GetAsync(url)
      ' will throw an exception if Not successful
      response.EnsureSuccessStatusCode()
      content = Await response.Content.ReadAsStringAsync()

      file_price_history_response = response_folder & "\" & tickers(i) & "_response.txt"
      If File.Exists(file_price_history_response) Then File.Delete(file_price_history_response)
      Dim writer As New StreamWriter(file_price_history_response)
      writer.Write(content)
      writer.Close()

      count = i + 1
      Form1.lblCount.Text = Trim(count.ToString)
    Next
    MessageBox.Show("Download finished")
  End Sub

  Function ReadTextFromFile%(ByVal sFileName$, ByRef sKey$)
    Dim line$
    ReadTextFromFile = -1
    line = ""
    sKey = ""
    If Not File.Exists(sFileName) Then Exit Function
    Try
      Dim reader As New StreamReader(sFileName)
      If reader.EndOfStream Then Exit Function
      line = reader.ReadLine()
      If (line Is Nothing) Then Exit Function
      line = line.Trim
      If line.Length <= 0 Then Exit Function
      sKey = line
      reader.Close()
    Catch e As Exception
      MessageBox.Show("Error in file " & sFileName & ": " & e.Message)
      ReadTextFromFile = -2
      Exit Function
    End Try
    ReadTextFromFile = 0
  End Function

  Sub UpdateDatabase(csv_folder$, data_source$)
    Dim line$
    Dim csvFiles = Directory.EnumerateFiles(csv_folder, "*.csv")

    If csvFiles.Count > 0 Then
      Dim market_price_db$ = "Data Source=" & data_source & ";Initial Catalog=market_data;Integrated Security=True;"
      Dim cn As New SqlConnection() ' Don't put this statement in a try block; it throws an exception!!!
      cn.ConnectionString = market_price_db
      cn.Open()

      Dim cmd As New SqlCommand, dr As SqlDataReader
      Dim ticker$, l%, num_rows%
      cmd.Connection = cn
      For Each currentFile$ In csvFiles
        Dim fileName$ = currentFile.Substring(csv_folder.Length + 1)
        l = InStrRev(fileName, ".") ' Reverse in case the ticker name contains a "." (like BRK.B)
        If l > 1 And l < Len(fileName) Then ' should always be true
          ticker = UCase$(Mid(fileName, 1, l - 1))
          Console.WriteLine("Updating records for ticker " & ticker)
          Dim reader As New StreamReader(currentFile)
          line = reader.ReadLine() ' skip 3 lines
          line = reader.ReadLine()
          line = reader.ReadLine()
          While Not reader.EndOfStream
            line = reader.ReadLine()
            If (line Is Nothing) Then Exit While
            line = line.Trim
            If line.Length > 0 Then
              Dim items = line.Split(",")
              Dim bFound As Boolean
              bFound = False
              Try
                cmd.CommandText = "SELECT [Ticker] from dbo.[market_price] where [Ticker]='" & ticker & "' and [Date]=" & items(0)
                dr = cmd.ExecuteReader
                If dr.HasRows Then bFound = True
                dr.Close()
              Catch e As Exception
                cmd.Dispose()
                cn.Close()
                MessageBox.Show(e.Message)
                Exit Sub
              End Try
              If bFound Then
                cmd.CommandText = "update dbo.[market_price] set [Open]=" & items(1) & ",[High]=" & items(2) &
                  ",[Low]=" & items(3) & ",[Close]=" & items(4) & ",[Volume]=" & items(5) & " where [Ticker]='" & ticker & "' and [Date]=" & items(0)
                num_rows = cmd.ExecuteNonQuery()
                If num_rows <> 1 Then
                  reader.Close()
                  cmd.Dispose()
                  cn.Close()
                  MessageBox.Show("Error updating record Ticker=" & ticker & " Date=" & items(0) & " file=" & fileName)
                  Exit Sub
                End If
              Else
                cmd.CommandText = "insert into dbo.[market_price] values('" & ticker & "'," & items(0) & "," & items(1) & "," _
                & items(2) & "," & items(3) & "," & items(4) & "," & items(5) & ")"
                num_rows = cmd.ExecuteNonQuery()
                If num_rows <> 1 Then
                  reader.Close()
                  cmd.Dispose()
                  cn.Close()
                  MessageBox.Show("Error adding record Ticker=" & ticker & " Date=" & items(0) & " file=" & fileName)
                  Exit Sub
                End If
              End If
            End If
          End While
          reader.Close()
        End If
      Next

      cmd.Dispose()
      cn.Close()
      MessageBox.Show("Database Update finished")
    Else
      MessageBox.Show("Database Update : no files found in folder")
    End If
  End Sub
  Function ReadInputFile(file_name$) As DataTable
    ReadInputFile = Nothing
    Dim symbol_col% = -1
    Dim date_col = -1
    Dim open_col% = -1
    Dim high_col% = -1
    Dim low_col% = -1
    Dim close_col = -1
    Dim volume_col% = -1

    If (Dir(file_name) = "" Or (Not File.Exists(file_name))) Then
      MessageBox.Show("File " & file_name & " was not found")
      Exit Function
    End If
    Dim line$ = "", count% = 0

    Dim i%, j%, n%, num_vars%, s1$
    Dim dt As New DataTable
    Try
      Dim reader As New StreamReader(file_name)
      While (Not reader.EndOfStream)
        line = reader.ReadLine()
        If (line Is Nothing) Then Exit Function
        line = line.Trim
        If line.Length <= 0 Then Exit Function
        Dim items = line.Split(",")
        n = items.Length
        If n < 7 Then Exit Function
        If items(0).Trim.Length <= 0 Then Exit Function

        Dim dr = dt.NewRow()
        If count = 0 Then
          num_vars = 0
          For i = 0 To n - 1
            s1 = items(i).Trim
            If s1.Equals("Symbol") Then
              num_vars += 1
              symbol_col = i
            ElseIf s1.Equals("Date") Then
              num_vars += 1
              date_col = i
            ElseIf s1.Equals("Current Price") Then
              num_vars += 1
              close_col = i
            ElseIf s1.Equals("Open") Then
              num_vars += 1
              open_col = i
            ElseIf s1.Equals("High") Then
              num_vars += 1
              high_col = i
            ElseIf s1.Equals("Low") Then
              num_vars += 1
              low_col = i
            ElseIf s1.Equals("Volume") Then
              num_vars += 1
              volume_col = i
            End If
          Next
          If num_vars <> 7 Then Exit Function

          dt.Columns.Add("Symbol")
          dt.Columns.Add("Date")
          dt.Columns.Add("Open")
          dt.Columns.Add("High")
          dt.Columns.Add("Low")
          dt.Columns.Add("Close")
          dt.Columns.Add("Volume")
        Else
          j = 0
          dr(j) = items(symbol_col).Trim
          j += 1
          dr(j) = items(date_col).Trim
          j += 1
          dr(j) = items(open_col).Trim
          j += 1
          dr(j) = items(high_col).Trim
          j += 1
          dr(j) = items(low_col).Trim
          j += 1
          dr(j) = items(close_col).Trim
          j += 1
          dr(j) = items(volume_col).Trim
          dt.Rows.Add(dr)
        End If
        count += 1
      End While
      reader.Close()
    Catch e As Exception
      MessageBox.Show("Error In file " & file_name & ": " & e.Message)
      Exit Function
    End Try
    ReadInputFile = dt
  End Function
  Function ImportYahooFile(yahoo_file$, ticker_list_file$, response_folder$, csv_folder$, data_source$)
    ImportYahooFile = -1
    Dim file_price_history$
    Dim i%, n%, count%, s$

    If Not File.Exists(yahoo_file) Then Exit Function
    Dim dt_if = ReadInputFile(yahoo_file)
    If IsNothing(dt_if) Then
      Exit Function
    End If
    Dim num_tickers_if% = dt_if.Rows.Count
    If num_tickers_if <= 0 Then
      Exit Function
    End If

    Dim dv = dt_if.DefaultView
    dv.Sort = "Symbol ASC"
    dt_if = dv.ToTable
    Dim tickers_if = dt_if.AsEnumerable.Select(Function(xx) xx(“Symbol”).ToString).ToArray

    ' read ticker symbol list file
    Dim num_tickers%, line$
    Dim lstTickers As New List(Of String)
    num_tickers = 0
    Try
      Dim reader As New StreamReader(ticker_list_file)
      While Not reader.EndOfStream
        line = reader.ReadLine()
        If (line Is Nothing) Then Exit While
        line = line.Trim.ToUpper
        If line.Length > 0 Then
          'Dim items = line.Split(",")
          lstTickers.Add(line)
          num_tickers += 1
        End If
      End While
      reader.Close()
    Catch e As Exception
      MessageBox.Show("Error in file " & ticker_list_file & ": " & e.Message)
      Exit Function
    End Try

    If num_tickers <= 0 Then
      MessageBox.Show("No ticker symbols found in file " & ticker_list_file)
      Exit Function
    End If
    Dim tickers = lstTickers.ToArray

    ' get dates from database
    Dim lstTickersDb, lstDates1 As New List(Of String)
    Dim num_tickers_db%
    Dim market_price_db$ = "Data Source=" & data_source & ";Initial Catalog=market_data;Integrated Security=True;"
    Dim cn As New SqlConnection() ' Don't put this statement in a try block; it throws an exception!!!
    cn.ConnectionString = market_price_db
    cn.Open()

    Dim cmd As New SqlCommand, dr As SqlDataReader
    Try
      'Dim cn As New SqlConnection(market_price_db)

      cmd.Connection = cn
      cmd.CommandText = "Select * from dbo.[Get_Last_Date]"
      dr = cmd.ExecuteReader
      num_tickers_db = 0
      While dr.Read()
        lstTickersDb.Add(dr("Ticker"))
        lstDates1.Add(dr("Last_Date"))
        num_tickers_db += 1
      End While
      dr.Close()
      cmd.Dispose()
      cn.Close()
    Catch e As Exception
      cmd.Dispose()
      cn.Close()
      MessageBox.Show(e.Message)
      Exit Function
    End Try

    If num_tickers_db <= 0 Then
      MessageBox.Show("No ticker symbols found in database")
      Exit Function
    End If
    Dim tickers_db = lstTickersDb.ToArray
    Dim dates1 = lstDates1.ToArray
    Array.Sort(tickers_db, dates1)

    ' delete the old files
    Dim fileEntries As String() = Directory.GetFiles(response_folder)
    Dim fileName As String
    For Each fileName In fileEntries
      File.Delete(fileName)
    Next fileName
    Dim fileEntries1 As String() = Directory.GetFiles(csv_folder)
    For Each fileName In fileEntries1
      File.Delete(fileName)
    Next fileName

    Dim CurrentDate As Date
    CurrentDate = dt_if.Rows(0).Item("Date").ToString

    Dim date2
    For i = 0 To num_tickers - 1
      Application.DoEvents()

      'find date in database
      If num_tickers_db <= 0 Then
        date2 = "0"
      Else
        Dim index1% = Array.BinarySearch(tickers_db, tickers(i))
        If index1 >= 0 Then
          date2 = dates1(index1)
        Else
          date2 = "0"
        End If
      End If
      If date2 = "0" Then
        MessageBox.Show("Error finding ticker " & tickers(i) & " in the database")
        Exit Function
      End If

      ' yahoo ticker symbol is BRK-B instead of BRK.B for example, so we need to replace "." with "-"
      ' Of course this depends on what is in the database which depends on the trading API. I wish this was standardized!
      Dim ticker1$ = tickers(i)
      Dim index3% = ticker1.IndexOf(".")
      If index3 > 0 Then
        ticker1 = ticker1.Replace(".", "-")
      End If
      Dim index2% = Array.BinarySearch(tickers_if, ticker1)
      If index2 < 0 Then
        MessageBox.Show("Error finding ticker " & tickers(i) & " in Yahoo file")
        Exit Function
      End If

      ' save the output files in a format that can be used by the database update
      file_price_history = csv_folder & "\" & tickers(i) & ".csv"

      Dim date1$, open1$, high1$, low1$, close1$, volume1$
      n = 1
      If File.Exists(file_price_history) Then File.Delete(file_price_history)
      Dim writer1 As New StreamWriter(file_price_history)
      writer1.WriteLine("rows")
      writer1.WriteLine(Trim$(n.ToString))
      s = "Date,Open,High,Low,Close,Volume"
      writer1.WriteLine(s)
      date1 = CurrentDate.ToString("yyyyMMdd")
      open1 = dt_if.Rows(index2).Item("Open").ToString
      high1 = dt_if.Rows(index2).Item("High").ToString
      low1 = dt_if.Rows(index2).Item("Low").ToString
      close1 = dt_if.Rows(index2).Item("Close").ToString
      volume1 = dt_if.Rows(index2).Item("Volume").ToString
      s = date1 & "," & open1 & "," & high1 & "," & low1 & "," & close1 & "," & volume1
      writer1.WriteLine(s)
      writer1.Close()
      count = i + 1
      Form2.lblCount.Text = Trim(count.ToString)
    Next
    MessageBox.Show("Yahoo import finished")
    ImportYahooFile = 0
  End Function

  Sub UpdateDatabaseFundamental(response_folder$, data_source$)
    Dim content$, l%, ticker$, num_fields%, num_rows%, i%, s1$, s2$, limit%
    Dim bFound As Boolean
    limit = 100
    Dim field_names$(0 To limit)
    Dim obj1 As Object
    Dim market_price_db$ = "Data Source=" & data_source & ";Initial Catalog=market_data;Integrated Security=True;"

    Dim cn As New SqlConnection() ' Don't put this statement in a try block; it throws an exception!!!
    cn.ConnectionString = market_price_db
    cn.Open()
    Dim cmd As New SqlCommand, dr As SqlDataReader
    cmd.Connection = cn
    Try
      cmd.CommandText = "SELECT * from dbo.[get_fundamental_field_names]"
      dr = cmd.ExecuteReader
      num_fields = 0
      dr.Read() ' skip first field
      While dr.Read() And num_fields <= limit
        field_names(num_fields) = dr("COLUMN_NAME")
        num_fields = num_fields + 1
      End While
      dr.Close()
    Catch e As Exception
      cmd.Dispose()
      cn.Close()
      MessageBox.Show(e.Message)
      Exit Sub
    End Try

    Dim field_values$(0 To num_fields - 1)
    Dim field_types$(0 To num_fields - 1)
    For i = 0 To num_fields - 1
      field_values(i) = ""
      field_types(i) = ""
    Next

    Dim i32 As Int32
    i32 = 0
    Dim txtFiles = Directory.EnumerateFiles(response_folder, "*.txt")
    If txtFiles.Count > 0 Then
      For Each currentFile$ In txtFiles
        Dim fileName$ = currentFile.Substring(response_folder.Length + 1)
        l = InStrRev(fileName, "_") ' Reverse in case the ticker name contains a "_" although it seems unlikely
        If l > 1 And l < Len(fileName) Then ' should always be true
          ticker = UCase$(Mid(fileName, 1, l - 1))
          Dim reader As New StreamReader(currentFile)
          content = reader.ReadLine()
          Dim token1
          token1 = ""
          Dim jss = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Object)(content)
          bFound = False
          bFound = jss.TryGetValue(ticker, token1)
          If bFound = True Then
            bFound = False
            bFound = jss(ticker).TryGetValue("fundamental", token1)
            If bFound = True Then
              For i = 0 To num_fields - 1
                s1 = ""
                s2 = ""
                bFound = False
                bFound = jss(ticker)("fundamental").TryGetValue(field_names(i), token1)
                If bFound Then
                  s1 = token1.ToString
                  obj1 = token1.type
                  s2 = obj1.ToString
                End If
                'If bFound Then s1 = jss(ticker)("fundamental")(field_names(i)).ToString

                If bFound = False Then
                  bFound = jss(ticker).TryGetValue(field_names(i), token1)
                  If bFound Then
                    s1 = token1.ToString
                    obj1 = token1.type
                    s2 = obj1.ToString
                  End If
                  'If bFound Then s1 = jss(ticker)(field_names(i)).ToString
                End If
                field_values(i) = s1.Trim()
                field_types(i) = s2.Trim()
              Next
              bFound = False
              Try
                cmd.CommandText = "SELECT [ticker] from dbo.[fundamentals] where [ticker]='" & ticker & "'"
                dr = cmd.ExecuteReader
                If dr.HasRows Then bFound = True
                dr.Close()
              Catch e As Exception
                cmd.Dispose()
                cn.Close()
                MessageBox.Show(e.Message)
                Exit Sub
              End Try

              If bFound Then
                s1 = "update dbo.[fundamentals] set "
                Dim ii% = 0
                For i = 0 To num_fields - 1
                  If field_types(i) <> "" Then
                    If ii > 0 Then s1 &= ","
                    If field_types(i) = "String" Then
                      s1 = s1 & "[" & field_names(i) & "]='" & field_values(i) & "'"
                    Else
                      s1 = s1 & "[" & field_names(i) & "]=" & field_values(i)
                    End If
                    ii += 1
                  End If
                Next
                s1 = s1 & " where [ticker]='" & ticker & "'"
                cmd.CommandText = s1
                num_rows = cmd.ExecuteNonQuery()
              Else
                s1 = "insert into dbo.[fundamentals] values('" & ticker & "'"
                For i = 0 To num_fields - 1
                  If field_types(i) = "" Then
                    s1 = s1 & "," & "NULL"
                  ElseIf field_types(i) = "String" Then
                    s1 = s1 & ",'" & field_values(i) & "'"
                  Else
                    s1 = s1 & "," & field_values(i)
                  End If
                Next
                s1 &= ")"
                cmd.CommandText = s1
                num_rows = cmd.ExecuteNonQuery()
              End If
            End If
          End If
        End If
      Next
      MessageBox.Show("Database Update finished")
    Else
      MessageBox.Show("Database Update : no files found in folder")

    End If
    cmd.Dispose()
    cn.Close()
  End Sub
  Sub UpdateTickerList(indicator_file$, ticker_list_file$, data_source$)
    Dim line$, n%, i%, j%, k%, limit%, limit1%
    limit = 9
    limit1 = 299
    Dim file_names$(0 To limit)
    Dim tickers$(0 To limit1), num_tickers%

    If File.Exists(ticker_list_file) = True Then
      File.Delete(ticker_list_file)
    End If
    If File.Exists(indicator_file) = False Then
      MessageBox.Show("file " & indicator_file & " does not exist")
      Exit Sub
    End If

    n = 0
    Try
      Dim reader As New StreamReader(indicator_file)
      While Not reader.EndOfStream And n <= limit
        line = reader.ReadLine()
        If (line Is Nothing) Then Exit While
        line = line.Trim
        If line.Length > 0 Then
          file_names(n) = line
          n = n + 1
        End If
      End While
      reader.Close()
    Catch e As Exception
      MessageBox.Show("Error in file " & indicator_file & ": " & e.Message)
      Exit Sub
    End Try

    Dim bFound As Boolean
    num_tickers = 0

    Dim bAppOpen As Boolean, bBookOpen As Boolean
    Dim oApp As New Excel.Application

    For i = 0 To n - 1
      If File.Exists(file_names(i)) = False Then
        MessageBox.Show("file " & file_names(i) & " does not exist")
        Exit Sub
      End If

      Dim oBook As Excel.Workbook
      Dim oSheet As New Excel.Worksheet
      Dim StartRow%, EndRow%, NumRows%, s1$

      bBookOpen = False
      Try
        bAppOpen = True
        oApp.Visible = True
        oBook = oApp.Workbooks.Open(file_names(i))
        bBookOpen = True
        oBook.Activate()
        oApp.WindowState = Excel.XlWindowState.xlMinimized
        Application.DoEvents()
        oApp.ScreenUpdating = False
        oSheet = oBook.Worksheets(1)
        Dim oRange As Excel.Range
        Dim oRange1 As Excel.Range
        StartRow = 2
        oRange = oSheet.Cells(StartRow, 2) ' the only column without an extra row
        oRange1 = oRange.End(Excel.XlDirection.xlDown)
        EndRow = oRange1.Row
        NumRows = EndRow - StartRow + 1
        For j = 1 To NumRows
          s1 = UCase(Trim(oSheet.Cells(StartRow + j - 1, 2).value))
          If s1.Length > 0 Then
            If num_tickers > 0 Then
              bFound = False
              For k = 0 To num_tickers - 1
                If s1 = tickers(k) Then
                  bFound = True
                  Exit For
                End If
              Next
              If Not bFound Then
                tickers(num_tickers) = s1
                num_tickers = num_tickers + 1
                If num_tickers > limit1 Then Exit For
              End If
            Else
              tickers(num_tickers) = s1
              num_tickers = num_tickers + 1
              If num_tickers > limit1 Then Exit For
            End If
          End If
        Next
      Catch e As Exception
        If (bBookOpen) Then
          oBook.Saved = True
          oBook.Close()
        End If
        MessageBox.Show("Error in file " & file_names(i) & ": " & e.Message)
        bBookOpen = False
        oApp.ScreenUpdating = True
        oApp.UserControl = False
        oApp.Quit()
        oBook = Nothing
        oApp = Nothing
        bAppOpen = False
        GC.Collect()
        Exit Sub
      End Try
      oBook.Saved = True
      oBook.Close()
      oBook = Nothing
      bBookOpen = False
      oSheet = Nothing
    Next
    bBookOpen = False
    oApp.ScreenUpdating = True
    oApp.UserControl = False
    oApp.Quit()
    oApp = Nothing
    bAppOpen = False
    GC.Collect()

    If num_tickers > 0 Then
      Dim writer As New StreamWriter(ticker_list_file)
      For i = 0 To num_tickers - 1
        writer.WriteLine(tickers(i))
      Next
      writer.Close()
    End If
    MessageBox.Show("Ticker List Update finished")
  End Sub
End Module

