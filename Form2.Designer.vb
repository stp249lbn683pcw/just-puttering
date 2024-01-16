<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form2))
        Me.butImportYahooFile = New System.Windows.Forms.Button()
        Me.butUpdate = New System.Windows.Forms.Button()
        Me.butImportYahoo1File = New System.Windows.Forms.Button()
        Me.lblCount = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.butUpdate1 = New System.Windows.Forms.Button()
        Me.butBrowseList = New System.Windows.Forms.Button()
        Me.lblInputFileName = New System.Windows.Forms.Label()
        Me.txtFileNameList = New System.Windows.Forms.TextBox()
        Me.butBrowseList1 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtFileNameList1 = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.butYahooFile = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtYahooFile = New System.Windows.Forms.TextBox()
        Me.butYahoo1File = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtYahoo1File = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'butImportYahooFile
        '
        Me.butImportYahooFile.Location = New System.Drawing.Point(66, 275)
        Me.butImportYahooFile.Name = "butImportYahooFile"
        Me.butImportYahooFile.Size = New System.Drawing.Size(265, 23)
        Me.butImportYahooFile.TabIndex = 5
        Me.butImportYahooFile.Text = "Import Yahoo file for database (ticker list)"
        Me.butImportYahooFile.UseVisualStyleBackColor = True
        '
        'butUpdate
        '
        Me.butUpdate.Location = New System.Drawing.Point(66, 314)
        Me.butUpdate.Name = "butUpdate"
        Me.butUpdate.Size = New System.Drawing.Size(216, 23)
        Me.butUpdate.TabIndex = 6
        Me.butUpdate.Text = "Update database"
        Me.butUpdate.UseVisualStyleBackColor = True
        '
        'butImportYahoo1File
        '
        Me.butImportYahoo1File.Location = New System.Drawing.Point(399, 275)
        Me.butImportYahoo1File.Name = "butImportYahoo1File"
        Me.butImportYahoo1File.Size = New System.Drawing.Size(269, 23)
        Me.butImportYahoo1File.TabIndex = 8
        Me.butImportYahoo1File.Text = "Import Yahoo file for database (ticker list 1)"
        Me.butImportYahoo1File.UseVisualStyleBackColor = True
        '
        'lblCount
        '
        Me.lblCount.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.lblCount.Location = New System.Drawing.Point(259, 383)
        Me.lblCount.Name = "lblCount"
        Me.lblCount.Size = New System.Drawing.Size(80, 25)
        Me.lblCount.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(259, 359)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Download Count"
        '
        'butUpdate1
        '
        Me.butUpdate1.Location = New System.Drawing.Point(399, 314)
        Me.butUpdate1.Name = "butUpdate1"
        Me.butUpdate1.Size = New System.Drawing.Size(216, 23)
        Me.butUpdate1.TabIndex = 11
        Me.butUpdate1.Text = "Update database"
        Me.butUpdate1.UseVisualStyleBackColor = True
        '
        'butBrowseList
        '
        Me.butBrowseList.Location = New System.Drawing.Point(508, 40)
        Me.butBrowseList.Name = "butBrowseList"
        Me.butBrowseList.Size = New System.Drawing.Size(75, 23)
        Me.butBrowseList.TabIndex = 18
        Me.butBrowseList.Text = "Browse"
        Me.butBrowseList.UseVisualStyleBackColor = True
        '
        'lblInputFileName
        '
        Me.lblInputFileName.AutoSize = True
        Me.lblInputFileName.Location = New System.Drawing.Point(64, 28)
        Me.lblInputFileName.Name = "lblInputFileName"
        Me.lblInputFileName.Size = New System.Drawing.Size(181, 13)
        Me.lblInputFileName.TabIndex = 17
        Me.lblInputFileName.Text = "Input File for Ticker Symbol List (*.txt)"
        '
        'txtFileNameList
        '
        Me.txtFileNameList.Location = New System.Drawing.Point(67, 44)
        Me.txtFileNameList.Name = "txtFileNameList"
        Me.txtFileNameList.Size = New System.Drawing.Size(404, 20)
        Me.txtFileNameList.TabIndex = 16
        '
        'butBrowseList1
        '
        Me.butBrowseList1.Location = New System.Drawing.Point(508, 173)
        Me.butBrowseList1.Name = "butBrowseList1"
        Me.butBrowseList1.Size = New System.Drawing.Size(75, 23)
        Me.butBrowseList1.TabIndex = 21
        Me.butBrowseList1.Text = "Browse"
        Me.butBrowseList1.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(64, 161)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(190, 13)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "Input File for Ticker Symbol List 1 (*.txt)"
        '
        'txtFileNameList1
        '
        Me.txtFileNameList1.Location = New System.Drawing.Point(67, 177)
        Me.txtFileNameList1.Name = "txtFileNameList1"
        Me.txtFileNameList1.Size = New System.Drawing.Size(404, 20)
        Me.txtFileNameList1.TabIndex = 19
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'butYahooFile
        '
        Me.butYahooFile.Location = New System.Drawing.Point(508, 93)
        Me.butYahooFile.Name = "butYahooFile"
        Me.butYahooFile.Size = New System.Drawing.Size(75, 23)
        Me.butYahooFile.TabIndex = 25
        Me.butYahooFile.Text = "Browse"
        Me.butYahooFile.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(64, 81)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(194, 13)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "Yahoo File for Ticker Symbol List (*.csv)"
        '
        'txtYahooFile
        '
        Me.txtYahooFile.Location = New System.Drawing.Point(67, 97)
        Me.txtYahooFile.Name = "txtYahooFile"
        Me.txtYahooFile.Size = New System.Drawing.Size(404, 20)
        Me.txtYahooFile.TabIndex = 23
        '
        'butYahoo1File
        '
        Me.butYahoo1File.Location = New System.Drawing.Point(508, 225)
        Me.butYahoo1File.Name = "butYahoo1File"
        Me.butYahoo1File.Size = New System.Drawing.Size(75, 23)
        Me.butYahoo1File.TabIndex = 28
        Me.butYahoo1File.Text = "Browse"
        Me.butYahoo1File.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(64, 213)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(203, 13)
        Me.Label5.TabIndex = 27
        Me.Label5.Text = "Yahoo File for Ticker Symbol List 1 (*.csv)"
        '
        'txtYahoo1File
        '
        Me.txtYahoo1File.Location = New System.Drawing.Point(67, 229)
        Me.txtYahoo1File.Name = "txtYahoo1File"
        Me.txtYahoo1File.Size = New System.Drawing.Size(404, 20)
        Me.txtYahoo1File.TabIndex = 26
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Menu
        Me.Label3.Location = New System.Drawing.Point(63, 429)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(681, 129)
        Me.Label3.TabIndex = 33
        Me.Label3.Text = resources.GetString("Label3.Text")
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.ClientSize = New System.Drawing.Size(849, 576)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.butYahoo1File)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtYahoo1File)
        Me.Controls.Add(Me.butYahooFile)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtYahooFile)
        Me.Controls.Add(Me.butBrowseList1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtFileNameList1)
        Me.Controls.Add(Me.butBrowseList)
        Me.Controls.Add(Me.lblInputFileName)
        Me.Controls.Add(Me.txtFileNameList)
        Me.Controls.Add(Me.butUpdate1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblCount)
        Me.Controls.Add(Me.butImportYahoo1File)
        Me.Controls.Add(Me.butUpdate)
        Me.Controls.Add(Me.butImportYahooFile)
        Me.Name = "Form2"
        Me.Text = "Import CSV Fie"
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents butImportYahooFile As Button
    Friend WithEvents butUpdate As Button
    Friend WithEvents butImportYahoo1File As Button
    Friend WithEvents lblCount As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents butUpdate1 As Button
    Friend WithEvents butBrowseList As Button
    Friend WithEvents lblInputFileName As Label
    Friend WithEvents txtFileNameList As TextBox
    Friend WithEvents butBrowseList1 As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents txtFileNameList1 As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents ErrorProvider1 As ErrorProvider
    Friend WithEvents butYahoo1File As Button
    Friend WithEvents Label5 As Label
    Friend WithEvents txtYahoo1File As TextBox
    Friend WithEvents butYahooFile As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents txtYahooFile As TextBox
    Friend WithEvents Label3 As Label
End Class
