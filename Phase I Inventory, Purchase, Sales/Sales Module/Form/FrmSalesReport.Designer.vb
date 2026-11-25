<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmSalesReport
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
        Me.BtnPrint = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.CboFormat = New System.Windows.Forms.ComboBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.ChkShowRatio = New System.Windows.Forms.CheckBox
        Me.LblPrintStatus = New System.Windows.Forms.Label
        Me.TxtDescSalesPerson = New System.Windows.Forms.TextBox
        Me.cboPrintStatus = New System.Windows.Forms.ComboBox
        Me.TxtDescCustomer = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtAccCode = New System.Windows.Forms.TextBox
        Me.BtnFindCustomerCode = New System.Windows.Forms.Button
        Me.TxtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.TxtCustomerCode = New System.Windows.Forms.TextBox
        Me.TxtDescAccount = New System.Windows.Forms.TextBox
        Me.BtnFindSales = New System.Windows.Forms.Button
        Me.TxtPrimarySales = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.BtnFindAcc = New System.Windows.Forms.Button
        Me.TxtKodeSales = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAcc = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.txtCodeNo = New System.Windows.Forms.TextBox
        Me.TxtDescProduct = New System.Windows.Forms.TextBox
        Me.BtnFindCodeNo = New System.Windows.Forms.Button
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.BtnClose = New System.Windows.Forms.Button
        Me.ChkAllProduct = New System.Windows.Forms.CheckBox
        Me.Crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.GroupBox1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BtnPrint
        '
        Me.BtnPrint.Location = New System.Drawing.Point(198, 93)
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(53, 20)
        Me.BtnPrint.TabIndex = 9
        Me.BtnPrint.Text = "Print"
        Me.BtnPrint.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.CboFormat)
        Me.GroupBox1.Controls.Add(Me.SplitContainer1)
        Me.GroupBox1.Controls.Add(Me.DtpTo)
        Me.GroupBox1.Controls.Add(Me.BtnPrint)
        Me.GroupBox1.Controls.Add(Me.DtpFrom)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.BtnClose)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(912, 148)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        '
        'CboFormat
        '
        Me.CboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFormat.FormattingEnabled = True
        Me.CboFormat.Items.AddRange(New Object() {"Rekap", "Detail"})
        Me.CboFormat.Location = New System.Drawing.Point(68, 15)
        Me.CboFormat.Name = "CboFormat"
        Me.CboFormat.Size = New System.Drawing.Size(184, 21)
        Me.CboFormat.TabIndex = 1
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(257, 12)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Panel1Collapsed = True
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TabControl1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtCodeNo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtDescProduct)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnFindCodeNo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label19)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtPrimaryCodeNo)
        Me.SplitContainer1.Size = New System.Drawing.Size(652, 130)
        Me.SplitContainer1.SplitterDistance = 357
        Me.SplitContainer1.TabIndex = 22
        Me.SplitContainer1.TabStop = False
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(5, 0)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(641, 124)
        Me.TabControl1.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.ChkShowRatio)
        Me.TabPage1.Controls.Add(Me.LblPrintStatus)
        Me.TabPage1.Controls.Add(Me.TxtDescSalesPerson)
        Me.TabPage1.Controls.Add(Me.cboPrintStatus)
        Me.TabPage1.Controls.Add(Me.TxtDescCustomer)
        Me.TabPage1.Controls.Add(Me.Label6)
        Me.TabPage1.Controls.Add(Me.TxtAccCode)
        Me.TabPage1.Controls.Add(Me.BtnFindCustomerCode)
        Me.TabPage1.Controls.Add(Me.TxtPrimaryCustomer)
        Me.TabPage1.Controls.Add(Me.TxtCustomerCode)
        Me.TabPage1.Controls.Add(Me.TxtDescAccount)
        Me.TabPage1.Controls.Add(Me.BtnFindSales)
        Me.TabPage1.Controls.Add(Me.TxtPrimarySales)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.BtnFindAcc)
        Me.TabPage1.Controls.Add(Me.TxtKodeSales)
        Me.TabPage1.Controls.Add(Me.TxtPrimaryAcc)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Size = New System.Drawing.Size(633, 98)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Additional Filter"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'ChkShowRatio
        '
        Me.ChkShowRatio.AutoSize = True
        Me.ChkShowRatio.Location = New System.Drawing.Point(347, 77)
        Me.ChkShowRatio.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ChkShowRatio.Name = "ChkShowRatio"
        Me.ChkShowRatio.Size = New System.Drawing.Size(142, 17)
        Me.ChkShowRatio.TabIndex = 12
        Me.ChkShowRatio.Text = "Show Ratio (If Available)"
        Me.ChkShowRatio.UseVisualStyleBackColor = True
        '
        'LblPrintStatus
        '
        Me.LblPrintStatus.AutoSize = True
        Me.LblPrintStatus.Location = New System.Drawing.Point(6, 78)
        Me.LblPrintStatus.Name = "LblPrintStatus"
        Me.LblPrintStatus.Size = New System.Drawing.Size(61, 13)
        Me.LblPrintStatus.TabIndex = 28
        Me.LblPrintStatus.Text = "Print Status"
        '
        'TxtDescSalesPerson
        '
        Me.TxtDescSalesPerson.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescSalesPerson.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDescSalesPerson.Location = New System.Drawing.Point(347, 30)
        Me.TxtDescSalesPerson.Name = "TxtDescSalesPerson"
        Me.TxtDescSalesPerson.ReadOnly = True
        Me.TxtDescSalesPerson.Size = New System.Drawing.Size(281, 20)
        Me.TxtDescSalesPerson.TabIndex = 19
        Me.TxtDescSalesPerson.TabStop = False
        '
        'cboPrintStatus
        '
        Me.cboPrintStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrintStatus.FormattingEnabled = True
        Me.cboPrintStatus.Items.AddRange(New Object() {"All Invoice", "Printed Invoice", "Unprinted Invoice"})
        Me.cboPrintStatus.Location = New System.Drawing.Point(115, 76)
        Me.cboPrintStatus.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboPrintStatus.Name = "cboPrintStatus"
        Me.cboPrintStatus.Size = New System.Drawing.Size(179, 21)
        Me.cboPrintStatus.TabIndex = 11
        '
        'TxtDescCustomer
        '
        Me.TxtDescCustomer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescCustomer.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDescCustomer.Location = New System.Drawing.Point(347, 6)
        Me.TxtDescCustomer.Name = "TxtDescCustomer"
        Me.TxtDescCustomer.ReadOnly = True
        Me.TxtDescCustomer.Size = New System.Drawing.Size(281, 20)
        Me.TxtDescCustomer.TabIndex = 22
        Me.TxtDescCustomer.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 7)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(79, 13)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Customer Code"
        '
        'TxtAccCode
        '
        Me.TxtAccCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAccCode.Location = New System.Drawing.Point(115, 53)
        Me.TxtAccCode.Name = "TxtAccCode"
        Me.TxtAccCode.Size = New System.Drawing.Size(179, 20)
        Me.TxtAccCode.TabIndex = 9
        '
        'BtnFindCustomerCode
        '
        Me.BtnFindCustomerCode.Location = New System.Drawing.Point(298, 5)
        Me.BtnFindCustomerCode.Name = "BtnFindCustomerCode"
        Me.BtnFindCustomerCode.Size = New System.Drawing.Size(43, 20)
        Me.BtnFindCustomerCode.TabIndex = 6
        Me.BtnFindCustomerCode.TabStop = False
        Me.BtnFindCustomerCode.Text = "Find"
        Me.BtnFindCustomerCode.UseVisualStyleBackColor = True
        '
        'TxtPrimaryCustomer
        '
        Me.TxtPrimaryCustomer.Location = New System.Drawing.Point(64, 6)
        Me.TxtPrimaryCustomer.Name = "TxtPrimaryCustomer"
        Me.TxtPrimaryCustomer.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryCustomer.TabIndex = 21
        Me.TxtPrimaryCustomer.TabStop = False
        Me.TxtPrimaryCustomer.Text = "-1"
        Me.TxtPrimaryCustomer.Visible = False
        '
        'TxtCustomerCode
        '
        Me.TxtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerCode.Location = New System.Drawing.Point(116, 6)
        Me.TxtCustomerCode.Name = "TxtCustomerCode"
        Me.TxtCustomerCode.Size = New System.Drawing.Size(178, 20)
        Me.TxtCustomerCode.TabIndex = 5
        '
        'TxtDescAccount
        '
        Me.TxtDescAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescAccount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDescAccount.Location = New System.Drawing.Point(347, 54)
        Me.TxtDescAccount.Name = "TxtDescAccount"
        Me.TxtDescAccount.ReadOnly = True
        Me.TxtDescAccount.Size = New System.Drawing.Size(281, 20)
        Me.TxtDescAccount.TabIndex = 24
        Me.TxtDescAccount.TabStop = False
        '
        'BtnFindSales
        '
        Me.BtnFindSales.Location = New System.Drawing.Point(298, 28)
        Me.BtnFindSales.Name = "BtnFindSales"
        Me.BtnFindSales.Size = New System.Drawing.Size(43, 21)
        Me.BtnFindSales.TabIndex = 8
        Me.BtnFindSales.TabStop = False
        Me.BtnFindSales.Text = "Find"
        Me.BtnFindSales.UseVisualStyleBackColor = True
        '
        'TxtPrimarySales
        '
        Me.TxtPrimarySales.Location = New System.Drawing.Point(64, 28)
        Me.TxtPrimarySales.Name = "TxtPrimarySales"
        Me.TxtPrimarySales.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimarySales.TabIndex = 13
        Me.TxtPrimarySales.TabStop = False
        Me.TxtPrimarySales.Text = "-1"
        Me.TxtPrimarySales.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(69, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Sales Person"
        '
        'BtnFindAcc
        '
        Me.BtnFindAcc.Location = New System.Drawing.Point(298, 53)
        Me.BtnFindAcc.Name = "BtnFindAcc"
        Me.BtnFindAcc.Size = New System.Drawing.Size(43, 20)
        Me.BtnFindAcc.TabIndex = 10
        Me.BtnFindAcc.TabStop = False
        Me.BtnFindAcc.Text = "Find"
        Me.BtnFindAcc.UseVisualStyleBackColor = True
        '
        'TxtKodeSales
        '
        Me.TxtKodeSales.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtKodeSales.Location = New System.Drawing.Point(115, 30)
        Me.TxtKodeSales.Name = "TxtKodeSales"
        Me.TxtKodeSales.Size = New System.Drawing.Size(179, 20)
        Me.TxtKodeSales.TabIndex = 7
        '
        'TxtPrimaryAcc
        '
        Me.TxtPrimaryAcc.Location = New System.Drawing.Point(64, 53)
        Me.TxtPrimaryAcc.Name = "TxtPrimaryAcc"
        Me.TxtPrimaryAcc.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryAcc.TabIndex = 17
        Me.TxtPrimaryAcc.TabStop = False
        Me.TxtPrimaryAcc.Text = "-1"
        Me.TxtPrimaryAcc.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 55)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(107, 13)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Receivable  Account"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvItemProductGroup)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Size = New System.Drawing.Size(633, 98)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Product Group"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeight = 25
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_Remarks, Me.ColProductGroup_PrimaryAcuan})
        Me.dgvItemProductGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemProductGroup.Location = New System.Drawing.Point(2, 2)
        Me.dgvItemProductGroup.MultiSelect = False
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(629, 94)
        Me.dgvItemProductGroup.TabIndex = 178
        '
        'ColProductGroup_No
        '
        Me.ColProductGroup_No.Frozen = True
        Me.ColProductGroup_No.HeaderText = "No"
        Me.ColProductGroup_No.Name = "ColProductGroup_No"
        Me.ColProductGroup_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_No.Width = 27
        '
        'ColProductGroup_Level
        '
        Me.ColProductGroup_Level.Frozen = True
        Me.ColProductGroup_Level.HeaderText = "Level"
        Me.ColProductGroup_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColProductGroup_Level.Name = "ColProductGroup_Level"
        Me.ColProductGroup_Level.Width = 39
        '
        'ColProductGroup_CodeNo
        '
        Me.ColProductGroup_CodeNo.Frozen = True
        Me.ColProductGroup_CodeNo.HeaderText = "Group Code"
        Me.ColProductGroup_CodeNo.Name = "ColProductGroup_CodeNo"
        Me.ColProductGroup_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_CodeNo.Width = 70
        '
        'ColProductGroup_Desc
        '
        Me.ColProductGroup_Desc.HeaderText = "Description"
        Me.ColProductGroup_Desc.Name = "ColProductGroup_Desc"
        Me.ColProductGroup_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_Desc.Width = 66
        '
        'ColProductGroup_Remarks
        '
        Me.ColProductGroup_Remarks.HeaderText = "Remarks"
        Me.ColProductGroup_Remarks.Name = "ColProductGroup_Remarks"
        Me.ColProductGroup_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_Remarks.Width = 55
        '
        'ColProductGroup_PrimaryAcuan
        '
        Me.ColProductGroup_PrimaryAcuan.HeaderText = "Primary Acuan"
        Me.ColProductGroup_PrimaryAcuan.Name = "ColProductGroup_PrimaryAcuan"
        Me.ColProductGroup_PrimaryAcuan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_PrimaryAcuan.Visible = False
        Me.ColProductGroup_PrimaryAcuan.Width = 81
        '
        'txtCodeNo
        '
        Me.txtCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCodeNo.Location = New System.Drawing.Point(123, 68)
        Me.txtCodeNo.Name = "txtCodeNo"
        Me.txtCodeNo.Size = New System.Drawing.Size(248, 20)
        Me.txtCodeNo.TabIndex = 7
        '
        'TxtDescProduct
        '
        Me.TxtDescProduct.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescProduct.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDescProduct.Location = New System.Drawing.Point(374, 68)
        Me.TxtDescProduct.Multiline = True
        Me.TxtDescProduct.Name = "TxtDescProduct"
        Me.TxtDescProduct.ReadOnly = True
        Me.TxtDescProduct.Size = New System.Drawing.Size(269, 20)
        Me.TxtDescProduct.TabIndex = 26
        Me.TxtDescProduct.TabStop = False
        '
        'BtnFindCodeNo
        '
        Me.BtnFindCodeNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindCodeNo.Location = New System.Drawing.Point(376, 68)
        Me.BtnFindCodeNo.Name = "BtnFindCodeNo"
        Me.BtnFindCodeNo.Size = New System.Drawing.Size(43, 20)
        Me.BtnFindCodeNo.TabIndex = 11
        Me.BtnFindCodeNo.TabStop = False
        Me.BtnFindCodeNo.Text = "Find"
        Me.BtnFindCodeNo.UseVisualStyleBackColor = True
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(3, 71)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(49, 13)
        Me.Label19.TabIndex = 6
        Me.Label19.Text = "Code No"
        '
        'txtPrimaryCodeNo
        '
        Me.txtPrimaryCodeNo.Location = New System.Drawing.Point(94, 67)
        Me.txtPrimaryCodeNo.Name = "txtPrimaryCodeNo"
        Me.txtPrimaryCodeNo.Size = New System.Drawing.Size(30, 20)
        Me.txtPrimaryCodeNo.TabIndex = 7
        Me.txtPrimaryCodeNo.TabStop = False
        Me.txtPrimaryCodeNo.Text = "-1"
        Me.txtPrimaryCodeNo.Visible = False
        '
        'DtpTo
        '
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(68, 69)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(184, 20)
        Me.DtpTo.TabIndex = 3
        '
        'DtpFrom
        '
        Me.DtpFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFrom.Location = New System.Drawing.Point(68, 43)
        Me.DtpFrom.Name = "DtpFrom"
        Me.DtpFrom.Size = New System.Drawing.Size(184, 20)
        Me.DtpFrom.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 69)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "To"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 45)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "From"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 19)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(39, 13)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "Format"
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(139, 93)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(53, 20)
        Me.BtnClose.TabIndex = 10
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'ChkAllProduct
        '
        Me.ChkAllProduct.AutoSize = True
        Me.ChkAllProduct.Location = New System.Drawing.Point(737, 185)
        Me.ChkAllProduct.Name = "ChkAllProduct"
        Me.ChkAllProduct.Size = New System.Drawing.Size(80, 17)
        Me.ChkAllProduct.TabIndex = 8
        Me.ChkAllProduct.Text = "All Product "
        Me.ChkAllProduct.UseVisualStyleBackColor = True
        '
        'Crv
        '
        Me.Crv.ActiveViewIndex = -1
        Me.Crv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Crv.DisplayGroupTree = False
        Me.Crv.EnableDrillDown = False
        Me.Crv.Location = New System.Drawing.Point(12, 149)
        Me.Crv.Name = "Crv"
        Me.Crv.SelectionFormula = ""
        Me.Crv.ShowCloseButton = False
        Me.Crv.ShowGroupTreeButton = False
        Me.Crv.ShowRefreshButton = False
        Me.Crv.Size = New System.Drawing.Size(912, 335)
        Me.Crv.TabIndex = 2
        Me.Crv.TabStop = False
        Me.Crv.ViewTimeSelectionFormula = ""
        '
        'BackgroundWorker3
        '
        '
        'FrmSalesReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(933, 495)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Crv)
        Me.Controls.Add(Me.ChkAllProduct)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(949, 533)
        Me.Name = "FrmSalesReport"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales Report By Invoice"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnPrint As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindCodeNo As System.Windows.Forms.Button
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryAcc As System.Windows.Forms.TextBox
    Friend WithEvents TxtAccCode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents BtnFindAcc As System.Windows.Forms.Button
    Friend WithEvents TxtPrimarySales As System.Windows.Forms.TextBox
    Friend WithEvents TxtKodeSales As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents BtnFindSales As System.Windows.Forms.Button
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents BtnFindCustomerCode As System.Windows.Forms.Button
    Friend WithEvents CboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtDescSalesPerson As System.Windows.Forms.TextBox
    Friend WithEvents TxtDescAccount As System.Windows.Forms.TextBox
    Friend WithEvents TxtDescCustomer As System.Windows.Forms.TextBox
    Friend WithEvents TxtDescProduct As System.Windows.Forms.TextBox
    Friend WithEvents ChkAllProduct As System.Windows.Forms.CheckBox
    Friend WithEvents LblPrintStatus As System.Windows.Forms.Label
    Friend WithEvents cboPrintStatus As System.Windows.Forms.ComboBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents ChkShowRatio As System.Windows.Forms.CheckBox
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
End Class
