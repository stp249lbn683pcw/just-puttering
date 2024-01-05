Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.IO

Public Class Form1
  Private Sub butGetRefreshToken_Click(sender As Object, e As EventArgs) Handles butGetRefreshToken.Click
    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    Call GetRefreshCode()
    Me.Cursor = Cursors.Default
  End Sub

  Private Sub butGetAccessToken_Click(sender As Object, e As EventArgs) Handles butGetAccessToken.Click
    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    With UserInput
      Call GetAccessCode(.file_consumer_key, .file_refresh_token, .file_access_token_response, .file_access_token)
    End With
    Me.Cursor = Cursors.Default
  End Sub

  Private Sub butUpdateTickerList_Click(sender As Object, e As EventArgs) Handles butUpdateTickerList.Click
    Dim AppPath$, error1%, sFileName$, ticker_list_file$, ticker_list1_file$, trading_API$
    AppPath = Application.StartupPath
    sFileName = AppPath & "\entries.ini"
    With UserInput
      trading_API = .trading_API
      .trading_API = cmbTradingAPI.Text.Trim
      ticker_list_file = .ticker_list_file
      ticker_list1_file = .ticker_list1_file
      .ticker_list_file = txtFileNameList.Text.Trim
      .ticker_list1_file = txtFileNameList1.Text.Trim
      If ticker_list_file <> txtFileNameList.Text.Trim Or ticker_list1_file <> txtFileNameList1.Text.Trim Or trading_API <> cmbTradingAPI.Text.Trim Then
        error1 = SaveEntries(sFileName)
        If error1 < 0 Then
          MessageBox.Show("Error saving file " & sFileName)
          Exit Sub
        End If
      End If
    End With

    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    With UserInput
      Call UpdateTickerList(.indicator_file, .ticker_list_file, .data_source)
    End With
    Me.Cursor = Cursors.Default
  End Sub

  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Dim AppPath$, error1%, sFileName$
    AppPath = Application.StartupPath
    sFileName = AppPath & "\GetStockData.ini"
    error1 = ReadDefaults(sFileName)
    If error1 < 0 Then MessageBox.Show("Error reading file " & sFileName)
    sFileName = AppPath & "\entries.ini"
    error1 = ReadEntries(sFileName)
    If error1 < 0 Then MessageBox.Show("Error reading file " & sFileName)

    With UserInput
      txtFileNameList.Text = .ticker_list_file.Trim
      txtFileNameList1.Text = .ticker_list1_file.Trim
      cmbTradingAPI.Text = .trading_API

      If .trading_API <> "TD Ameritrade" Then
        butGetRefreshToken.Hide()
        butGetAccessToken.Hide()
        butDownloadFundamental.Hide()
        butDownloadFundamental1.Hide()
        butUpdateFundamental.Hide()
        butUpdateFundamental1.Hide()
      End If
    End With
  End Sub

  Private Sub butDownload_Click(sender As Object, e As EventArgs) Handles butDownload.Click
    Dim AppPath$, error1%, sFileName$, ticker_list_file$, ticker_list1_file$, trading_API$
    AppPath = Application.StartupPath
    sFileName = AppPath & "\entries.ini"
    With UserInput
      trading_API = .trading_API
      .trading_API = cmbTradingAPI.Text.Trim
      ticker_list_file = .ticker_list_file
      ticker_list1_file = .ticker_list1_file
      .ticker_list_file = txtFileNameList.Text.Trim
      .ticker_list1_file = txtFileNameList1.Text.Trim
      If ticker_list_file <> txtFileNameList.Text.Trim Or ticker_list1_file <> txtFileNameList1.Text.Trim Or trading_API <> cmbTradingAPI.Text.Trim Then
        error1 = SaveEntries(sFileName)
        If error1 < 0 Then
          MessageBox.Show("Error saving file " & sFileName)
          Exit Sub
        End If
      End If
    End With

    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    Application.DoEvents()
    With UserInput
      If .trading_API = "TD Ameritrade" Then
        Call DownloadHistData(.file_consumer_key, .file_access_token, .ticker_list_file, .response_folder, .csv_folder, .data_source)
      ElseIf .trading_API = "Polygon.io" Then
        Call DownloadHistDataPolygonIo(.file_polygon_io_api_key, .ticker_list_file, .response_folder, .csv_folder, .data_source)
      End If
    End With
    Me.Cursor = Cursors.Default
  End Sub

  Private Sub butUpdate_Click(sender As Object, e As EventArgs) Handles butUpdate.Click
    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    With UserInput
      Call UpdateDatabase(.csv_folder, .data_source)
    End With
    Me.Cursor = Cursors.Default
  End Sub

  Private Sub butDownload1_Click(sender As Object, e As EventArgs) Handles butDownload1.Click
    Dim AppPath$, error1%, sFileName$, ticker_list_file$, ticker_list1_file$, trading_API$
    AppPath = Application.StartupPath
    sFileName = AppPath & "\entries.ini"
    With UserInput
      trading_API = .trading_API
      .trading_API = cmbTradingAPI.Text.Trim
      ticker_list_file = .ticker_list_file
      ticker_list1_file = .ticker_list1_file
      .ticker_list_file = txtFileNameList.Text.Trim
      .ticker_list1_file = txtFileNameList1.Text.Trim
      If ticker_list_file <> txtFileNameList.Text.Trim Or ticker_list1_file <> txtFileNameList1.Text.Trim Or trading_API <> cmbTradingAPI.Text.Trim Then
        error1 = SaveEntries(sFileName)
        If error1 < 0 Then
          MessageBox.Show("Error saving file " & sFileName)
          Exit Sub
        End If
      End If
    End With

    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    Application.DoEvents()
    With UserInput
      If .trading_API = "TD Ameritrade" Then
        Call DownloadHistData(.file_consumer_key, .file_access_token, .ticker_list1_file, .response1_folder, .csv1_folder, .data_source)
      ElseIf .trading_API = "Polygon.io" Then
        Call DownloadHistDataPolygonIo(.file_polygon_io_api_key, .ticker_list1_file, .response1_folder, .csv1_folder, .data_source)
      End If
    End With
    Me.Cursor = Cursors.Default
  End Sub

  Private Sub butUpdate1_Click(sender As Object, e As EventArgs) Handles butUpdate1.Click
    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    With UserInput
      Call UpdateDatabase(.csv1_folder, .data_source)
    End With
    Me.Cursor = Cursors.Default
  End Sub

  Private Sub butDownloadFundamental_Click(sender As Object, e As EventArgs) Handles butDownloadFundamental.Click
    Dim AppPath$, error1%, sFileName$, ticker_list_file$, ticker_list1_file$, trading_API$
    AppPath = Application.StartupPath
    sFileName = AppPath & "\entries.ini"
    With UserInput
      trading_API = .trading_API
      .trading_API = cmbTradingAPI.Text.Trim
      ticker_list_file = .ticker_list_file
      ticker_list1_file = .ticker_list1_file
      .ticker_list_file = txtFileNameList.Text.Trim
      .ticker_list1_file = txtFileNameList1.Text.Trim
      If ticker_list_file <> txtFileNameList.Text.Trim Or ticker_list1_file <> txtFileNameList1.Text.Trim Or trading_API <> cmbTradingAPI.Text.Trim Then
        error1 = SaveEntries(sFileName)
        If error1 < 0 Then
          MessageBox.Show("Error saving file " & sFileName)
          Exit Sub
        End If
      End If
    End With

    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    Application.DoEvents()
    With UserInput
      Call DownloadFundamental(.file_consumer_key, .file_access_token, .ticker_list_file, .fundamental_response_folder, .data_source)
    End With
    Me.Cursor = Cursors.Default
  End Sub

  Private Sub butUpdateFundamental_Click(sender As Object, e As EventArgs) Handles butUpdateFundamental.Click
    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    With UserInput
      Call UpdateDatabaseFundamental(.fundamental_response_folder, .data_source)
    End With
    Me.Cursor = Cursors.Default
  End Sub

  Private Sub butDownloadFundamental1_Click(sender As Object, e As EventArgs) Handles butDownloadFundamental1.Click
    Dim AppPath$, error1%, sFileName$, ticker_list_file$, ticker_list1_file$, trading_API$
    AppPath = Application.StartupPath
    sFileName = AppPath & "\entries.ini"
    With UserInput
      trading_API = .trading_API
      .trading_API = cmbTradingAPI.Text.Trim
      ticker_list_file = .ticker_list_file
      ticker_list1_file = .ticker_list1_file
      .ticker_list_file = txtFileNameList.Text.Trim
      .ticker_list1_file = txtFileNameList1.Text.Trim
      If ticker_list_file <> txtFileNameList.Text.Trim Or ticker_list1_file <> txtFileNameList1.Text.Trim Or trading_API <> cmbTradingAPI.Text.Trim Then
        error1 = SaveEntries(sFileName)
        If error1 < 0 Then
          MessageBox.Show("Error saving file " & sFileName)
          Exit Sub
        End If
      End If
    End With

    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    Application.DoEvents()
    With UserInput
      Call DownloadFundamental(.file_consumer_key, .file_access_token, .ticker_list1_file, .fundamental_response1_folder, .data_source)
    End With
    Me.Cursor = Cursors.Default
  End Sub

  Private Sub butUpdateFundamental1_Click(sender As Object, e As EventArgs) Handles butUpdateFundamental1.Click
    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    With UserInput
      Call UpdateDatabaseFundamental(.fundamental_response1_folder, .data_source)
    End With
    Me.Cursor = Cursors.Default
  End Sub

  Private Sub butBrowseList_Click(sender As Object, e As EventArgs) Handles butBrowseList.Click
    Dim s1$, s2$

    'Directory.SetCurrentDirectory(CurrentDir)
    s2 = ""
    s1 = txtFileNameList.Text.Trim
    If s1.Length > 0 Then
      If Directory.Exists(Path.GetDirectoryName(s1)) = True Then
        OpenFileDialog1.InitialDirectory = Path.GetDirectoryName(s1)
        s2 = Path.GetFileName(s1)
      End If
    End If

    OpenFileDialog1.FileName = s2
    OpenFileDialog1.DefaultExt = "txt"
    OpenFileDialog1.Filter = "Text files (*.txt)|*.txt"
    OpenFileDialog1.CheckFileExists = True
    OpenFileDialog1.CheckPathExists = True

    If OpenFileDialog1.ShowDialog = DialogResult.OK Then
      s1 = OpenFileDialog1.FileName
      txtFileNameList.Text = s1
      ErrorProvider1.SetError(txtFileNameList, "")
    End If
  End Sub

  Private Sub butBrowseList1_Click(sender As Object, e As EventArgs) Handles butBrowseList1.Click
    Dim s1$, s2$

    'Directory.SetCurrentDirectory(CurrentDir)
    s2 = ""
    s1 = txtFileNameList1.Text.Trim
    If s1.Length > 0 Then
      If Directory.Exists(Path.GetDirectoryName(s1)) = True Then
        OpenFileDialog1.InitialDirectory = Path.GetDirectoryName(s1)
        s2 = Path.GetFileName(s1)
      End If
    End If

    OpenFileDialog1.FileName = s2
    OpenFileDialog1.DefaultExt = "txt"
    OpenFileDialog1.Filter = "Text files (*.txt)|*.txt"
    OpenFileDialog1.CheckFileExists = True
    OpenFileDialog1.CheckPathExists = True

    If OpenFileDialog1.ShowDialog = DialogResult.OK Then
      s1 = OpenFileDialog1.FileName
      txtFileNameList1.Text = s1
      ErrorProvider1.SetError(txtFileNameList1, "")
    End If
  End Sub

  Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
    Dim AppPath$, error1%, sFileName$, ticker_list_file$, ticker_list1_file$, trading_API$
    AppPath = Application.StartupPath
    sFileName = AppPath & "\entries.ini"
    With UserInput
      trading_API = .trading_API
      .trading_API = cmbTradingAPI.Text.Trim
      ticker_list_file = .ticker_list_file
      ticker_list1_file = .ticker_list1_file
      .ticker_list_file = txtFileNameList.Text.Trim
      .ticker_list1_file = txtFileNameList1.Text.Trim
      If ticker_list_file <> txtFileNameList.Text.Trim Or ticker_list1_file <> txtFileNameList1.Text.Trim Or trading_API <> cmbTradingAPI.Text.Trim Then
        error1 = SaveEntries(sFileName)
        If error1 < 0 Then
          MessageBox.Show("Error saving file " & sFileName)
          Exit Sub
        End If
      End If
    End With
  End Sub

  Private Sub cmbTradingAPI_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTradingAPI.SelectedIndexChanged
    If cmbTradingAPI.Text.Trim = "TD Ameritrade" Then
      butGetRefreshToken.Show()
      butGetAccessToken.Show()
      butDownloadFundamental.Show()
      butDownloadFundamental1.Show()
      butUpdateFundamental.Show()
      butUpdateFundamental1.Show()
    Else
      butGetRefreshToken.Hide()
      butGetAccessToken.Hide()
      butDownloadFundamental.Hide()
      butDownloadFundamental1.Hide()
      butUpdateFundamental.Hide()
      butUpdateFundamental1.Hide()
    End If
  End Sub
End Class
