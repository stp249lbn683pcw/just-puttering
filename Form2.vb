Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.IO

Public Class Form2
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
      txtYahooFile.Text = .yahoo_file
      txtFileNameList1.Text = .ticker_list1_file.Trim
      txtYahoo1File.Text = .yahoo1_file
    End With
  End Sub

  Private Sub butImportYahooFile_Click(sender As Object, e As EventArgs) Handles butImportYahooFile.Click
    Dim err%
    Dim AppPath$, error1%, sFileName$, ticker_list_file$, ticker_list1_file$
    Dim yahoo_file$, yahoo1_file$
    AppPath = Application.StartupPath
    sFileName = AppPath & "\entries.ini"

    With UserInput
      ticker_list_file = txtFileNameList.Text.Trim
      ticker_list1_file = txtFileNameList1.Text.Trim
      yahoo_file = txtYahooFile.Text.Trim
      yahoo1_file = txtYahoo1File.Text.Trim
    End With

    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    Application.DoEvents()
    With UserInput
      err = ImportYahooFile(yahoo_file, ticker_list_file, .response_folder, .csv_folder, .data_source)
      If err < 0 Then
        MessageBox.Show("Error reading file " & yahoo_file)
        Me.Cursor = Cursors.Default
        Exit Sub
      End If

      If ticker_list_file <> .ticker_list_file Or ticker_list1_file <> .ticker_list1_file Or yahoo_file <> .yahoo_file Or yahoo1_file <> .yahoo1_file Then
        .ticker_list_file = ticker_list_file
        Form1.txtFileNameList.Text = .ticker_list_file
        .ticker_list1_file = ticker_list1_file
        Form1.txtFileNameList1.Text = .ticker_list1_file
        .yahoo_file = yahoo_file
        .yahoo1_file = yahoo1_file
        error1 = SaveEntries(sFileName)
        If error1 < 0 Then
          MessageBox.Show("Error saving file " & sFileName)
        End If
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

  Private Sub butImportYahoo1File_Click(sender As Object, e As EventArgs) Handles butImportYahoo1File.Click
    Dim err%, AppPath$, error1%, sFileName$, ticker_list_file$, ticker_list1_file$
    Dim yahoo_file$, yahoo1_file$
    AppPath = Application.StartupPath
    sFileName = AppPath & "\entries.ini"

    With UserInput
      ticker_list_file = txtFileNameList.Text.Trim
      ticker_list1_file = txtFileNameList1.Text.Trim
      yahoo_file = txtYahooFile.Text.Trim
      yahoo1_file = txtYahoo1File.Text.Trim
    End With

    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    Application.DoEvents()
    With UserInput
      err = ImportYahooFile(yahoo1_file, ticker_list1_file, .response1_folder, .csv1_folder, .data_source)
      If err < 0 Then
        MessageBox.Show("Error reading file " & yahoo1_file)
        Me.Cursor = Cursors.Default
        Exit Sub
      End If

      If ticker_list_file <> .ticker_list_file Or ticker_list1_file <> .ticker_list1_file Or yahoo_file <> .yahoo_file Or yahoo1_file <> .yahoo1_file Then
        .ticker_list_file = ticker_list_file
        Form1.txtFileNameList.Text = .ticker_list_file
        .ticker_list1_file = ticker_list1_file
        Form1.txtFileNameList1.Text = .ticker_list1_file
        .yahoo_file = yahoo_file
        .yahoo1_file = yahoo1_file
        error1 = SaveEntries(sFileName)
        If error1 < 0 Then
          MessageBox.Show("Error saving file " & sFileName)
        End If
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

  Private Sub butBrowseList_Click(sender As Object, e As EventArgs) Handles butBrowseList.Click
    Dim s1$, s2$

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
    Dim AppPath$, error1%, sFileName$, ticker_list_file$, ticker_list1_file$
    Dim yahoo_file$, yahoo1_file$
    AppPath = Application.StartupPath
    sFileName = AppPath & "\entries.ini"

    With UserInput
      ticker_list_file = txtFileNameList.Text.Trim
      ticker_list1_file = txtFileNameList1.Text.Trim
      yahoo_file = txtYahooFile.Text.Trim
      yahoo1_file = txtYahoo1File.Text.Trim

      If ticker_list_file <> .ticker_list_file Or ticker_list1_file <> .ticker_list1_file Or yahoo_file <> .yahoo_file Or yahoo1_file <> .yahoo1_file Then
        .ticker_list_file = ticker_list_file
        Form1.txtFileNameList.Text = .ticker_list_file
        .ticker_list1_file = ticker_list1_file
        Form1.txtFileNameList1.Text = .ticker_list1_file
        .yahoo_file = yahoo_file
        .yahoo1_file = yahoo1_file
        error1 = SaveEntries(sFileName)
        If error1 < 0 Then
          MessageBox.Show("Error saving file " & sFileName)
        End If
      End If
    End With
  End Sub

  Private Sub butYahooFile_Click(sender As Object, e As EventArgs) Handles butYahooFile.Click
    Dim s1$, s2$

    s2 = ""
    s1 = txtYahooFile.Text.Trim
    If s1.Length > 0 Then
      If Directory.Exists(Path.GetDirectoryName(s1)) = True Then
        OpenFileDialog1.InitialDirectory = Path.GetDirectoryName(s1)
        s2 = Path.GetFileName(s1)
      End If
    End If

    OpenFileDialog1.FileName = s2
    OpenFileDialog1.DefaultExt = "csv"
    OpenFileDialog1.Filter = "Text files (*.csv)|*.csv"
    OpenFileDialog1.CheckFileExists = True
    OpenFileDialog1.CheckPathExists = True

    If OpenFileDialog1.ShowDialog = DialogResult.OK Then
      s1 = OpenFileDialog1.FileName
      txtYahooFile.Text = s1
      ErrorProvider1.SetError(txtYahooFile, "")
    End If
  End Sub

  Private Sub butYahoo1File_Click(sender As Object, e As EventArgs) Handles butYahoo1File.Click
    Dim s1$, s2$

    s2 = ""
    s1 = txtYahoo1File.Text.Trim
    If s1.Length > 0 Then
      If Directory.Exists(Path.GetDirectoryName(s1)) = True Then
        OpenFileDialog1.InitialDirectory = Path.GetDirectoryName(s1)
        s2 = Path.GetFileName(s1)
      End If
    End If

    OpenFileDialog1.FileName = s2
    OpenFileDialog1.DefaultExt = "csv"
    OpenFileDialog1.Filter = "Text files (*.csv)|*.csv"
    OpenFileDialog1.CheckFileExists = True
    OpenFileDialog1.CheckPathExists = True

    If OpenFileDialog1.ShowDialog = DialogResult.OK Then
      s1 = OpenFileDialog1.FileName
      txtYahoo1File.Text = s1
      ErrorProvider1.SetError(txtYahoo1File, "")
    End If
  End Sub
End Class
