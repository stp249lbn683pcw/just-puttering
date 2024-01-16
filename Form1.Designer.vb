<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Me.butGetRefreshToken = New System.Windows.Forms.Button()
        Me.butGetAccessToken = New System.Windows.Forms.Button()
        Me.butUpdateTickerList = New System.Windows.Forms.Button()
        Me.butDownload = New System.Windows.Forms.Button()
        Me.butUpdate = New System.Windows.Forms.Button()
        Me.butDownload1 = New System.Windows.Forms.Button()
        Me.lblCount = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.butUpdate1 = New System.Windows.Forms.Button()
        Me.butUpdateFundamental = New System.Windows.Forms.Button()
        Me.butDownloadFundamental = New System.Windows.Forms.Button()
        Me.butUpdateFundamental1 = New System.Windows.Forms.Button()
        Me.butDownloadFundamental1 = New System.Windows.Forms.Button()
        Me.butBrowseList = New System.Windows.Forms.Button()
        Me.lblInputFileName = New System.Windows.Forms.Label()
        Me.txtFileNameList = New System.Windows.Forms.TextBox()
        Me.butBrowseList1 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtFileNameList1 = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmbTradingAPI = New System.Windows.Forms.ComboBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ImportYahooFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'butGetRefreshToken
        '
        Me.butGetRefreshToken.Location = New System.Drawing.Point(16, 224)
        Me.butGetRefreshToken.Name = "butGetRefreshToken"
        Me.butGetRefreshToken.Size = New System.Drawing.Size(146, 23)
        Me.butGetRefreshToken.TabIndex = 0
        Me.butGetRefreshToken.Text = "Get Refresh Token"
        Me.butGetRefreshToken.UseVisualStyleBackColor = True
        '
        'butGetAccessToken
        '
        Me.butGetAccessToken.Location = New System.Drawing.Point(16, 274)
        Me.butGetAccessToken.Name = "butGetAccessToken"
        Me.butGetAccessToken.Size = New System.Drawing.Size(146, 23)
        Me.butGetAccessToken.TabIndex = 1
        Me.butGetAccessToken.Text = "Get Access Token"
        Me.butGetAccessToken.UseVisualStyleBackColor = True
        '
        'butUpdateTickerList
        '
        Me.butUpdateTickerList.Location = New System.Drawing.Point(199, 224)
        Me.butUpdateTickerList.Name = "butUpdateTickerList"
        Me.butUpdateTickerList.Size = New System.Drawing.Size(362, 23)
        Me.butUpdateTickerList.TabIndex = 4
        Me.butUpdateTickerList.Text = "Update ticker list using setup info from Excel workbooks (ticker list)"
        Me.butUpdateTickerList.UseVisualStyleBackColor = True
        '
        'butDownload
        '
        Me.butDownload.Location = New System.Drawing.Point(199, 274)
        Me.butDownload.Name = "butDownload"
        Me.butDownload.Size = New System.Drawing.Size(265, 23)
        Me.butDownload.TabIndex = 5
        Me.butDownload.Text = "Download historical data for database (ticker list)"
        Me.butDownload.UseVisualStyleBackColor = True
        '
        'butUpdate
        '
        Me.butUpdate.Location = New System.Drawing.Point(199, 313)
        Me.butUpdate.Name = "butUpdate"
        Me.butUpdate.Size = New System.Drawing.Size(216, 23)
        Me.butUpdate.TabIndex = 6
        Me.butUpdate.Text = "Update database"
        Me.butUpdate.UseVisualStyleBackColor = True
        '
        'butDownload1
        '
        Me.butDownload1.Location = New System.Drawing.Point(532, 274)
        Me.butDownload1.Name = "butDownload1"
        Me.butDownload1.Size = New System.Drawing.Size(269, 23)
        Me.butDownload1.TabIndex = 8
        Me.butDownload1.Text = "Download historical data for database (ticker list 1)"
        Me.butDownload1.UseVisualStyleBackColor = True
        '
        'lblCount
        '
        Me.lblCount.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.lblCount.Location = New System.Drawing.Point(356, 495)
        Me.lblCount.Name = "lblCount"
        Me.lblCount.Size = New System.Drawing.Size(80, 25)
        Me.lblCount.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(356, 471)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Download Count"
        '
        'butUpdate1
        '
        Me.butUpdate1.Location = New System.Drawing.Point(532, 313)
        Me.butUpdate1.Name = "butUpdate1"
        Me.butUpdate1.Size = New System.Drawing.Size(216, 23)
        Me.butUpdate1.TabIndex = 11
        Me.butUpdate1.Text = "Update database"
        Me.butUpdate1.UseVisualStyleBackColor = True
        '
        'butUpdateFundamental
        '
        Me.butUpdateFundamental.Location = New System.Drawing.Point(199, 424)
        Me.butUpdateFundamental.Name = "butUpdateFundamental"
        Me.butUpdateFundamental.Size = New System.Drawing.Size(216, 23)
        Me.butUpdateFundamental.TabIndex = 13
        Me.butUpdateFundamental.Text = "Update database"
        Me.butUpdateFundamental.UseVisualStyleBackColor = True
        '
        'butDownloadFundamental
        '
        Me.butDownloadFundamental.Location = New System.Drawing.Point(199, 385)
        Me.butDownloadFundamental.Name = "butDownloadFundamental"
        Me.butDownloadFundamental.Size = New System.Drawing.Size(294, 23)
        Me.butDownloadFundamental.TabIndex = 12
        Me.butDownloadFundamental.Text = "Download fundamental data for database (ticker list)"
        Me.butDownloadFundamental.UseVisualStyleBackColor = True
        '
        'butUpdateFundamental1
        '
        Me.butUpdateFundamental1.Location = New System.Drawing.Point(532, 424)
        Me.butUpdateFundamental1.Name = "butUpdateFundamental1"
        Me.butUpdateFundamental1.Size = New System.Drawing.Size(216, 23)
        Me.butUpdateFundamental1.TabIndex = 15
        Me.butUpdateFundamental1.Text = "Update database"
        Me.butUpdateFundamental1.UseVisualStyleBackColor = True
        '
        'butDownloadFundamental1
        '
        Me.butDownloadFundamental1.Location = New System.Drawing.Point(532, 385)
        Me.butDownloadFundamental1.Name = "butDownloadFundamental1"
        Me.butDownloadFundamental1.Size = New System.Drawing.Size(296, 23)
        Me.butDownloadFundamental1.TabIndex = 14
        Me.butDownloadFundamental1.Text = "Download fundamental data for database (ticker list 1)"
        Me.butDownloadFundamental1.UseVisualStyleBackColor = True
        '
        'butBrowseList
        '
        Me.butBrowseList.Location = New System.Drawing.Point(598, 93)
        Me.butBrowseList.Name = "butBrowseList"
        Me.butBrowseList.Size = New System.Drawing.Size(75, 23)
        Me.butBrowseList.TabIndex = 18
        Me.butBrowseList.Text = "Browse"
        Me.butBrowseList.UseVisualStyleBackColor = True
        '
        'lblInputFileName
        '
        Me.lblInputFileName.AutoSize = True
        Me.lblInputFileName.Location = New System.Drawing.Point(154, 81)
        Me.lblInputFileName.Name = "lblInputFileName"
        Me.lblInputFileName.Size = New System.Drawing.Size(181, 13)
        Me.lblInputFileName.TabIndex = 17
        Me.lblInputFileName.Text = "Input File for Ticker Symbol List (*.txt)"
        '
        'txtFileNameList
        '
        Me.txtFileNameList.Location = New System.Drawing.Point(157, 97)
        Me.txtFileNameList.Name = "txtFileNameList"
        Me.txtFileNameList.Size = New System.Drawing.Size(404, 20)
        Me.txtFileNameList.TabIndex = 16
        '
        'butBrowseList1
        '
        Me.butBrowseList1.Location = New System.Drawing.Point(598, 150)
        Me.butBrowseList1.Name = "butBrowseList1"
        Me.butBrowseList1.Size = New System.Drawing.Size(75, 23)
        Me.butBrowseList1.TabIndex = 21
        Me.butBrowseList1.Text = "Browse"
        Me.butBrowseList1.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(154, 138)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(190, 13)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "Input File for Ticker Symbol List 1 (*.txt)"
        '
        'txtFileNameList1
        '
        Me.txtFileNameList1.Location = New System.Drawing.Point(157, 154)
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
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(154, 46)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(113, 13)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "Select the trading API:"
        '
        'cmbTradingAPI
        '
        Me.cmbTradingAPI.FormattingEnabled = True
        Me.cmbTradingAPI.Items.AddRange(New Object() {"TD Ameritrade", "Polygon.io"})
        Me.cmbTradingAPI.Location = New System.Drawing.Point(273, 43)
        Me.cmbTradingAPI.Name = "cmbTradingAPI"
        Me.cmbTradingAPI.Size = New System.Drawing.Size(121, 21)
        Me.cmbTradingAPI.TabIndex = 23
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportYahooFileToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(849, 24)
        Me.MenuStrip1.TabIndex = 24
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ImportYahooFileToolStripMenuItem
        '
        Me.ImportYahooFileToolStripMenuItem.Name = "ImportYahooFileToolStripMenuItem"
        Me.ImportYahooFileToolStripMenuItem.Size = New System.Drawing.Size(136, 20)
        Me.ImportYahooFileToolStripMenuItem.Text = "Import Yahoo CSV File"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.ClientSize = New System.Drawing.Size(849, 552)
        Me.Controls.Add(Me.cmbTradingAPI)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.butBrowseList1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtFileNameList1)
        Me.Controls.Add(Me.butBrowseList)
        Me.Controls.Add(Me.lblInputFileName)
        Me.Controls.Add(Me.txtFileNameList)
        Me.Controls.Add(Me.butUpdateFundamental1)
        Me.Controls.Add(Me.butDownloadFundamental1)
        Me.Controls.Add(Me.butUpdateFundamental)
        Me.Controls.Add(Me.butDownloadFundamental)
        Me.Controls.Add(Me.butUpdate1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblCount)
        Me.Controls.Add(Me.butDownload1)
        Me.Controls.Add(Me.butUpdate)
        Me.Controls.Add(Me.butDownload)
        Me.Controls.Add(Me.butUpdateTickerList)
        Me.Controls.Add(Me.butGetAccessToken)
        Me.Controls.Add(Me.butGetRefreshToken)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "Form1"
        Me.Text = "Get Stock Data"
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents butGetRefreshToken As Button
    Friend WithEvents butGetAccessToken As Button
    Friend WithEvents butUpdateTickerList As Button
    Friend WithEvents butDownload As Button
    Friend WithEvents butUpdate As Button
    Friend WithEvents butDownload1 As Button
    Friend WithEvents lblCount As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents butUpdate1 As Button
    Friend WithEvents butUpdateFundamental As Button
    Friend WithEvents butDownloadFundamental As Button
    Friend WithEvents butUpdateFundamental1 As Button
    Friend WithEvents butDownloadFundamental1 As Button
    Friend WithEvents butBrowseList As Button
    Friend WithEvents lblInputFileName As Label
    Friend WithEvents txtFileNameList As TextBox
    Friend WithEvents butBrowseList1 As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents txtFileNameList1 As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents ErrorProvider1 As ErrorProvider
    Friend WithEvents cmbTradingAPI As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents ImportYahooFileToolStripMenuItem As ToolStripMenuItem
End Class
