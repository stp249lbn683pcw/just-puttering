Imports System.Data.SqlClient
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
    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    With UserInput
      Call UpdateTickerList(.indicator_file, .ticker_list_file, .data_source)
    End With
    Me.Cursor = Cursors.Default
  End Sub

  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Dim CurrentDir$, error1%, sFileName$
    CurrentDir$ = Application.StartupPath
    sFileName = CurrentDir$ & "\GetStockData.ini"
    error1 = ReadDefaults(sFileName)
    If error1 < 0 Then MessageBox.Show("Error reading file " & sFileName)
  End Sub

  Private Sub butDownload_Click(sender As Object, e As EventArgs) Handles butDownload.Click
    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    Application.DoEvents()
    With UserInput
      Call DownloadHistData(.file_consumer_key, .file_access_token, .ticker_list_file, .response_folder, .csv_folder, .data_source)
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
    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    Application.DoEvents()
    With UserInput
      Call DownloadHistData(.file_consumer_key, .file_access_token, .ticker_list1_file, .response1_folder, .csv1_folder, .data_source)
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
End Class
