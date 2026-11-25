<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanSalesPerformance
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmLaporanSalesPerformance))
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpSampai = New System.Windows.Forms.DateTimePicker
        Me.dtpDari = New System.Windows.Forms.DateTimePicker
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnShowReport = New System.Windows.Forms.Button
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.PrintForm1 = New Microsoft.VisualBasic.PowerPacks.Printing.PrintForm(Me.components)
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.llbSG_UncheckAll = New System.Windows.Forms.LinkLabel
        Me.llbSG_CheckAll = New System.Windows.Forms.LinkLabel
        Me.chkSalesGroup = New System.Windows.Forms.CheckBox
        Me.clbSalesPersonGroup = New System.Windows.Forms.CheckedListBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.llbSP_UncheckAll = New System.Windows.Forms.LinkLabel
        Me.llbSP_CheckAll = New System.Windows.Forms.LinkLabel
        Me.chkSalesPerson = New System.Windows.Forms.CheckBox
        Me.clbSalesPerson = New System.Windows.Forms.CheckedListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboFormat = New System.Windows.Forms.ComboBox
        Me.LblTo2 = New System.Windows.Forms.Label
        Me.lblPeriod2 = New System.Windows.Forms.Label
        Me.dtpSampai2 = New System.Windows.Forms.DateTimePicker
        Me.dtpDari2 = New System.Windows.Forms.DateTimePicker
        Me.cboReference = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cboFormat1 = New System.Windows.Forms.ComboBox
        Me.cboFormat2 = New System.Windows.Forms.ComboBox
        Me.chkGoods = New System.Windows.Forms.CheckBox
        Me.chkService = New System.Windows.Forms.CheckBox
        Me.chkReturn = New System.Windows.Forms.CheckBox
        Me.cboGroup = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(225, 141)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(25, 17)
        Me.Label3.TabIndex = 44
        Me.Label3.Text = "To"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(14, 142)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 17)
        Me.Label2.TabIndex = 43
        Me.Label2.Text = "Period I"
        '
        'dtpSampai
        '
        Me.dtpSampai.CustomFormat = "MMM yyyy"
        Me.dtpSampai.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpSampai.Location = New System.Drawing.Point(258, 139)
        Me.dtpSampai.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpSampai.Name = "dtpSampai"
        Me.dtpSampai.ShowUpDown = True
        Me.dtpSampai.Size = New System.Drawing.Size(122, 22)
        Me.dtpSampai.TabIndex = 5
        '
        'dtpDari
        '
        Me.dtpDari.CustomFormat = "MMM yyyy"
        Me.dtpDari.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDari.Location = New System.Drawing.Point(95, 139)
        Me.dtpDari.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpDari.Name = "dtpDari"
        Me.dtpDari.ShowUpDown = True
        Me.dtpDari.Size = New System.Drawing.Size(122, 22)
        Me.dtpDari.TabIndex = 4
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(194, 226)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(87, 30)
        Me.btnExit.TabIndex = 12
        Me.btnExit.Text = "&Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnShowReport
        '
        Me.btnShowReport.Location = New System.Drawing.Point(94, 226)
        Me.btnShowReport.Margin = New System.Windows.Forms.Padding(4)
        Me.btnShowReport.Name = "btnShowReport"
        Me.btnShowReport.Size = New System.Drawing.Size(87, 30)
        Me.btnShowReport.TabIndex = 11
        Me.btnShowReport.Text = "&Filter"
        Me.btnShowReport.UseVisualStyleBackColor = True
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.EnableDrillDown = False
        Me.crv1.Location = New System.Drawing.Point(13, 264)
        Me.crv1.Margin = New System.Windows.Forms.Padding(4)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(966, 356)
        Me.crv1.TabIndex = 13
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'PrintForm1
        '
        Me.PrintForm1.DocumentName = "document"
        Me.PrintForm1.Form = Me
        Me.PrintForm1.PrintAction = System.Drawing.Printing.PrintAction.PrintToPrinter
        Me.PrintForm1.PrinterSettings = CType(resources.GetObject("PrintForm1.PrinterSettings"), System.Drawing.Printing.PrinterSettings)
        Me.PrintForm1.PrintFileName = Nothing
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer3.IsSplitterFixed = True
        Me.SplitContainer3.Location = New System.Drawing.Point(410, 8)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.GroupBox5)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer3.Size = New System.Drawing.Size(569, 212)
        Me.SplitContainer3.SplitterDistance = 282
        Me.SplitContainer3.SplitterWidth = 5
        Me.SplitContainer3.TabIndex = 45
        '
        'GroupBox5
        '
        Me.GroupBox5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox5.Controls.Add(Me.llbSG_UncheckAll)
        Me.GroupBox5.Controls.Add(Me.llbSG_CheckAll)
        Me.GroupBox5.Controls.Add(Me.chkSalesGroup)
        Me.GroupBox5.Controls.Add(Me.clbSalesPersonGroup)
        Me.GroupBox5.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox5.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox5.Size = New System.Drawing.Size(282, 212)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
        '
        'llbSG_UncheckAll
        '
        Me.llbSG_UncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbSG_UncheckAll.AutoSize = True
        Me.llbSG_UncheckAll.Location = New System.Drawing.Point(193, 196)
        Me.llbSG_UncheckAll.Name = "llbSG_UncheckAll"
        Me.llbSG_UncheckAll.Size = New System.Drawing.Size(82, 17)
        Me.llbSG_UncheckAll.TabIndex = 3
        Me.llbSG_UncheckAll.TabStop = True
        Me.llbSG_UncheckAll.Text = "Uncheck All"
        '
        'llbSG_CheckAll
        '
        Me.llbSG_CheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbSG_CheckAll.AutoSize = True
        Me.llbSG_CheckAll.Location = New System.Drawing.Point(121, 196)
        Me.llbSG_CheckAll.Name = "llbSG_CheckAll"
        Me.llbSG_CheckAll.Size = New System.Drawing.Size(66, 17)
        Me.llbSG_CheckAll.TabIndex = 2
        Me.llbSG_CheckAll.TabStop = True
        Me.llbSG_CheckAll.Text = "Check All"
        '
        'chkSalesGroup
        '
        Me.chkSalesGroup.AutoSize = True
        Me.chkSalesGroup.Location = New System.Drawing.Point(9, -2)
        Me.chkSalesGroup.Margin = New System.Windows.Forms.Padding(4)
        Me.chkSalesGroup.Name = "chkSalesGroup"
        Me.chkSalesGroup.Size = New System.Drawing.Size(177, 21)
        Me.chkSalesGroup.TabIndex = 0
        Me.chkSalesGroup.Text = "All Sales Person Group"
        Me.chkSalesGroup.UseVisualStyleBackColor = True
        '
        'clbSalesPersonGroup
        '
        Me.clbSalesPersonGroup.CheckOnClick = True
        Me.clbSalesPersonGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbSalesPersonGroup.FormattingEnabled = True
        Me.clbSalesPersonGroup.Location = New System.Drawing.Point(4, 19)
        Me.clbSalesPersonGroup.Margin = New System.Windows.Forms.Padding(4)
        Me.clbSalesPersonGroup.Name = "clbSalesPersonGroup"
        Me.clbSalesPersonGroup.Size = New System.Drawing.Size(274, 174)
        Me.clbSalesPersonGroup.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.llbSP_UncheckAll)
        Me.GroupBox1.Controls.Add(Me.llbSP_CheckAll)
        Me.GroupBox1.Controls.Add(Me.chkSalesPerson)
        Me.GroupBox1.Controls.Add(Me.clbSalesPerson)
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(282, 212)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'llbSP_UncheckAll
        '
        Me.llbSP_UncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbSP_UncheckAll.AutoSize = True
        Me.llbSP_UncheckAll.Location = New System.Drawing.Point(193, 195)
        Me.llbSP_UncheckAll.Name = "llbSP_UncheckAll"
        Me.llbSP_UncheckAll.Size = New System.Drawing.Size(82, 17)
        Me.llbSP_UncheckAll.TabIndex = 3
        Me.llbSP_UncheckAll.TabStop = True
        Me.llbSP_UncheckAll.Text = "Uncheck All"
        '
        'llbSP_CheckAll
        '
        Me.llbSP_CheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbSP_CheckAll.AutoSize = True
        Me.llbSP_CheckAll.Location = New System.Drawing.Point(121, 195)
        Me.llbSP_CheckAll.Name = "llbSP_CheckAll"
        Me.llbSP_CheckAll.Size = New System.Drawing.Size(66, 17)
        Me.llbSP_CheckAll.TabIndex = 2
        Me.llbSP_CheckAll.TabStop = True
        Me.llbSP_CheckAll.Text = "Check All"
        '
        'chkSalesPerson
        '
        Me.chkSalesPerson.AutoSize = True
        Me.chkSalesPerson.Location = New System.Drawing.Point(9, -2)
        Me.chkSalesPerson.Margin = New System.Windows.Forms.Padding(4)
        Me.chkSalesPerson.Name = "chkSalesPerson"
        Me.chkSalesPerson.Size = New System.Drawing.Size(133, 21)
        Me.chkSalesPerson.TabIndex = 0
        Me.chkSalesPerson.Text = "All Sales Person"
        Me.chkSalesPerson.UseVisualStyleBackColor = True
        '
        'clbSalesPerson
        '
        Me.clbSalesPerson.CheckOnClick = True
        Me.clbSalesPerson.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbSalesPerson.FormattingEnabled = True
        Me.clbSalesPerson.Location = New System.Drawing.Point(4, 19)
        Me.clbSalesPerson.Margin = New System.Windows.Forms.Padding(4)
        Me.clbSalesPerson.Name = "clbSalesPerson"
        Me.clbSalesPerson.Size = New System.Drawing.Size(274, 174)
        Me.clbSalesPerson.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 45)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(52, 17)
        Me.Label1.TabIndex = 46
        Me.Label1.Text = "Format"
        '
        'cboFormat
        '
        Me.cboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormat.FormattingEnabled = True
        Me.cboFormat.Items.AddRange(New Object() {"Sales Person Recap", "Sales Person Group Recap", "Sales Person Detail", "Sales Person Group Trend", "Sales Person Group Trend %", "Sales Person Trend", "Sales Person Trend %", "Sales Person Composition"})
        Me.cboFormat.Location = New System.Drawing.Point(94, 42)
        Me.cboFormat.Name = "cboFormat"
        Me.cboFormat.Size = New System.Drawing.Size(309, 24)
        Me.cboFormat.TabIndex = 1
        '
        'LblTo2
        '
        Me.LblTo2.AutoSize = True
        Me.LblTo2.Location = New System.Drawing.Point(225, 171)
        Me.LblTo2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblTo2.Name = "LblTo2"
        Me.LblTo2.Size = New System.Drawing.Size(25, 17)
        Me.LblTo2.TabIndex = 51
        Me.LblTo2.Text = "To"
        '
        'lblPeriod2
        '
        Me.lblPeriod2.AutoSize = True
        Me.lblPeriod2.Location = New System.Drawing.Point(14, 172)
        Me.lblPeriod2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblPeriod2.Name = "lblPeriod2"
        Me.lblPeriod2.Size = New System.Drawing.Size(59, 17)
        Me.lblPeriod2.TabIndex = 50
        Me.lblPeriod2.Text = "Period II"
        '
        'dtpSampai2
        '
        Me.dtpSampai2.CustomFormat = "MMM yyyy"
        Me.dtpSampai2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpSampai2.Location = New System.Drawing.Point(258, 169)
        Me.dtpSampai2.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpSampai2.Name = "dtpSampai2"
        Me.dtpSampai2.ShowUpDown = True
        Me.dtpSampai2.Size = New System.Drawing.Size(122, 22)
        Me.dtpSampai2.TabIndex = 7
        '
        'dtpDari2
        '
        Me.dtpDari2.CustomFormat = "MMM yyyy"
        Me.dtpDari2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDari2.Location = New System.Drawing.Point(95, 169)
        Me.dtpDari2.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpDari2.Name = "dtpDari2"
        Me.dtpDari2.ShowUpDown = True
        Me.dtpDari2.Size = New System.Drawing.Size(122, 22)
        Me.dtpDari2.TabIndex = 6
        '
        'cboReference
        '
        Me.cboReference.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboReference.FormattingEnabled = True
        Me.cboReference.Items.AddRange(New Object() {"Sales Order", "Sales Invoice", "Gross Margin"})
        Me.cboReference.Location = New System.Drawing.Point(94, 12)
        Me.cboReference.Name = "cboReference"
        Me.cboReference.Size = New System.Drawing.Size(123, 24)
        Me.cboReference.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(13, 15)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(74, 17)
        Me.Label6.TabIndex = 52
        Me.Label6.Text = "Reference"
        '
        'cboFormat1
        '
        Me.cboFormat1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormat1.FormattingEnabled = True
        Me.cboFormat1.Items.AddRange(New Object() {"-", "With Budget", "2 Period Comparison"})
        Me.cboFormat1.Location = New System.Drawing.Point(228, 12)
        Me.cboFormat1.Name = "cboFormat1"
        Me.cboFormat1.Size = New System.Drawing.Size(175, 24)
        Me.cboFormat1.TabIndex = 2
        '
        'cboFormat2
        '
        Me.cboFormat2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormat2.FormattingEnabled = True
        Me.cboFormat2.Items.AddRange(New Object() {"-", "Sales Person", "Sales Person Group", "Cust. Segmentation", "Sales Area", "Cust. Type", "Cust. Group", "Product Group", "Product Brand", "Product Sub Brand", "Product Category", "Product"})
        Me.cboFormat2.Location = New System.Drawing.Point(94, 102)
        Me.cboFormat2.Name = "cboFormat2"
        Me.cboFormat2.Size = New System.Drawing.Size(309, 24)
        Me.cboFormat2.TabIndex = 3
        '
        'chkGoods
        '
        Me.chkGoods.AutoSize = True
        Me.chkGoods.Checked = True
        Me.chkGoods.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkGoods.Location = New System.Drawing.Point(180, 198)
        Me.chkGoods.Name = "chkGoods"
        Me.chkGoods.Size = New System.Drawing.Size(72, 21)
        Me.chkGoods.TabIndex = 9
        Me.chkGoods.Text = "Goods"
        Me.chkGoods.UseVisualStyleBackColor = True
        '
        'chkService
        '
        Me.chkService.AutoSize = True
        Me.chkService.Checked = True
        Me.chkService.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkService.Location = New System.Drawing.Point(298, 198)
        Me.chkService.Name = "chkService"
        Me.chkService.Size = New System.Drawing.Size(77, 21)
        Me.chkService.TabIndex = 10
        Me.chkService.Text = "Service"
        Me.chkService.UseVisualStyleBackColor = True
        '
        'chkReturn
        '
        Me.chkReturn.AutoSize = True
        Me.chkReturn.Checked = True
        Me.chkReturn.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkReturn.Location = New System.Drawing.Point(17, 198)
        Me.chkReturn.Name = "chkReturn"
        Me.chkReturn.Size = New System.Drawing.Size(122, 21)
        Me.chkReturn.TabIndex = 8
        Me.chkReturn.Text = "Include Return"
        Me.chkReturn.UseVisualStyleBackColor = True
        '
        'cboGroup
        '
        Me.cboGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGroup.FormattingEnabled = True
        Me.cboGroup.Items.AddRange(New Object() {"Sales Person", "Sales Person Group", "Cust. Segmentation", "Sales Area", "Cust. Type", "Cust. Group", "Product Group", "Product Brand", "Product Sub Brand", "Product Category", "Product"})
        Me.cboGroup.Location = New System.Drawing.Point(94, 72)
        Me.cboGroup.Name = "cboGroup"
        Me.cboGroup.Size = New System.Drawing.Size(309, 24)
        Me.cboGroup.TabIndex = 53
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 75)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(68, 17)
        Me.Label4.TabIndex = 54
        Me.Label4.Text = "Group By"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(13, 105)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(44, 17)
        Me.Label5.TabIndex = 55
        Me.Label5.Text = "Detail"
        '
        'FrmLaporanSalesPerformance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cboGroup)
        Me.Controls.Add(Me.chkReturn)
        Me.Controls.Add(Me.chkService)
        Me.Controls.Add(Me.chkGoods)
        Me.Controls.Add(Me.cboFormat2)
        Me.Controls.Add(Me.cboFormat1)
        Me.Controls.Add(Me.cboReference)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.LblTo2)
        Me.Controls.Add(Me.lblPeriod2)
        Me.Controls.Add(Me.dtpSampai2)
        Me.Controls.Add(Me.dtpDari2)
        Me.Controls.Add(Me.cboFormat)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpSampai)
        Me.Controls.Add(Me.dtpDari)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnShowReport)
        Me.Controls.Add(Me.crv1)
        Me.Controls.Add(Me.SplitContainer3)
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "FrmLaporanSalesPerformance"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales Person Performance Report"
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpSampai As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpDari As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnShowReport As System.Windows.Forms.Button
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents PrintForm1 As Microsoft.VisualBasic.PowerPacks.Printing.PrintForm
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents chkSalesGroup As System.Windows.Forms.CheckBox
    Friend WithEvents clbSalesPersonGroup As System.Windows.Forms.CheckedListBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkSalesPerson As System.Windows.Forms.CheckBox
    Friend WithEvents clbSalesPerson As System.Windows.Forms.CheckedListBox
    Friend WithEvents cboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LblTo2 As System.Windows.Forms.Label
    Friend WithEvents lblPeriod2 As System.Windows.Forms.Label
    Friend WithEvents dtpSampai2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpDari2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents cboReference As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboFormat1 As System.Windows.Forms.ComboBox
    Friend WithEvents cboFormat2 As System.Windows.Forms.ComboBox
    Friend WithEvents chkService As System.Windows.Forms.CheckBox
    Friend WithEvents chkGoods As System.Windows.Forms.CheckBox
    Friend WithEvents llbSG_UncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbSG_CheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbSP_UncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbSP_CheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents chkReturn As System.Windows.Forms.CheckBox
    Friend WithEvents cboGroup As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
End Class
