<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanSalesDP
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
        Me.crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.RboInvoiceDate = New System.Windows.Forms.RadioButton
        Me.Label7 = New System.Windows.Forms.Label
        Me.DtpInvoice_DateTo = New System.Windows.Forms.DateTimePicker
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtPrimarySalesPerson = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtSalesPersonCode = New System.Windows.Forms.TextBox
        Me.TxtSalesPersonDesc = New System.Windows.Forms.TextBox
        Me.BtnFindSalesPerson = New System.Windows.Forms.Button
        Me.DtpInvoice_DateFrom = New System.Windows.Forms.DateTimePicker
        Me.RboDP_Date = New System.Windows.Forms.RadioButton
        Me.RboSO_Date = New System.Windows.Forms.RadioButton
        Me.TxtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtCustomerCode = New System.Windows.Forms.TextBox
        Me.TxtCustomerName = New System.Windows.Forms.TextBox
        Me.BtnFindCustomer = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.DtpDP_DateTo = New System.Windows.Forms.DateTimePicker
        Me.DtpDP_DateFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.DtpSO_DateTo = New System.Windows.Forms.DateTimePicker
        Me.DtpSO_DateFrom = New System.Windows.Forms.DateTimePicker
        Me.BtnClose = New System.Windows.Forms.Button
        Me.BtnFilter = New System.Windows.Forms.Button
        Me.CboGroup = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.CboRealizationStatus = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.CboPaymentStatus = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.CboDPStatus = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'crv
        '
        Me.crv.ActiveViewIndex = -1
        Me.crv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv.DisplayGroupTree = False
        Me.crv.EnableDrillDown = False
        Me.crv.Location = New System.Drawing.Point(10, 124)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(725, 380)
        Me.crv.TabIndex = 29
        Me.crv.TabStop = False
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'RboInvoiceDate
        '
        Me.RboInvoiceDate.AutoSize = True
        Me.RboInvoiceDate.Location = New System.Drawing.Point(2, 49)
        Me.RboInvoiceDate.Margin = New System.Windows.Forms.Padding(2)
        Me.RboInvoiceDate.Name = "RboInvoiceDate"
        Me.RboInvoiceDate.Size = New System.Drawing.Size(86, 17)
        Me.RboInvoiceDate.TabIndex = 6
        Me.RboInvoiceDate.Text = "Invoice Date"
        Me.RboInvoiceDate.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(212, 51)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(20, 13)
        Me.Label7.TabIndex = 33
        Me.Label7.Text = "To"
        '
        'DtpInvoice_DateTo
        '
        Me.DtpInvoice_DateTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpInvoice_DateTo.CustomFormat = "dd MMM yyyy"
        Me.DtpInvoice_DateTo.Enabled = False
        Me.DtpInvoice_DateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpInvoice_DateTo.Location = New System.Drawing.Point(235, 47)
        Me.DtpInvoice_DateTo.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpInvoice_DateTo.Name = "DtpInvoice_DateTo"
        Me.DtpInvoice_DateTo.Size = New System.Drawing.Size(126, 20)
        Me.DtpInvoice_DateTo.TabIndex = 8
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(10, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimarySalesPerson)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtSalesPersonCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtSalesPersonDesc)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindSalesPerson)
        Me.SplitContainer1.Panel1.Controls.Add(Me.RboInvoiceDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpInvoice_DateTo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpInvoice_DateFrom)
        Me.SplitContainer1.Panel1.Controls.Add(Me.RboDP_Date)
        Me.SplitContainer1.Panel1.Controls.Add(Me.RboSO_Date)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryCustomer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCustomerCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCustomerName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindCustomer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpDP_DateTo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpDP_DateFrom)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpSO_DateTo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpSO_DateFrom)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnFilter)
        Me.SplitContainer1.Panel2.Controls.Add(Me.CboGroup)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.CboRealizationStatus)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.CboPaymentStatus)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.CboDPStatus)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Size = New System.Drawing.Size(724, 115)
        Me.SplitContainer1.SplitterDistance = 362
        Me.SplitContainer1.TabIndex = 30
        Me.SplitContainer1.TabStop = False
        '
        'TxtPrimarySalesPerson
        '
        Me.TxtPrimarySalesPerson.Location = New System.Drawing.Point(59, 92)
        Me.TxtPrimarySalesPerson.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtPrimarySalesPerson.Name = "TxtPrimarySalesPerson"
        Me.TxtPrimarySalesPerson.Size = New System.Drawing.Size(31, 20)
        Me.TxtPrimarySalesPerson.TabIndex = 39
        Me.TxtPrimarySalesPerson.TabStop = False
        Me.TxtPrimarySalesPerson.Text = "-1"
        Me.TxtPrimarySalesPerson.Visible = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(2, 97)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(69, 13)
        Me.Label9.TabIndex = 38
        Me.Label9.Text = "Sales Person"
        '
        'TxtSalesPersonCode
        '
        Me.TxtSalesPersonCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSalesPersonCode.Location = New System.Drawing.Point(94, 93)
        Me.TxtSalesPersonCode.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtSalesPersonCode.Name = "TxtSalesPersonCode"
        Me.TxtSalesPersonCode.Size = New System.Drawing.Size(67, 20)
        Me.TxtSalesPersonCode.TabIndex = 12
        '
        'TxtSalesPersonDesc
        '
        Me.TxtSalesPersonDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSalesPersonDesc.Location = New System.Drawing.Point(212, 93)
        Me.TxtSalesPersonDesc.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtSalesPersonDesc.Name = "TxtSalesPersonDesc"
        Me.TxtSalesPersonDesc.ReadOnly = True
        Me.TxtSalesPersonDesc.Size = New System.Drawing.Size(150, 20)
        Me.TxtSalesPersonDesc.TabIndex = 14
        Me.TxtSalesPersonDesc.TabStop = False
        '
        'BtnFindSalesPerson
        '
        Me.BtnFindSalesPerson.Location = New System.Drawing.Point(165, 93)
        Me.BtnFindSalesPerson.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnFindSalesPerson.Name = "BtnFindSalesPerson"
        Me.BtnFindSalesPerson.Size = New System.Drawing.Size(43, 20)
        Me.BtnFindSalesPerson.TabIndex = 13
        Me.BtnFindSalesPerson.Text = "Find"
        Me.BtnFindSalesPerson.UseVisualStyleBackColor = True
        '
        'DtpInvoice_DateFrom
        '
        Me.DtpInvoice_DateFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpInvoice_DateFrom.Enabled = False
        Me.DtpInvoice_DateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpInvoice_DateFrom.Location = New System.Drawing.Point(94, 47)
        Me.DtpInvoice_DateFrom.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpInvoice_DateFrom.Name = "DtpInvoice_DateFrom"
        Me.DtpInvoice_DateFrom.Size = New System.Drawing.Size(114, 20)
        Me.DtpInvoice_DateFrom.TabIndex = 7
        '
        'RboDP_Date
        '
        Me.RboDP_Date.AutoSize = True
        Me.RboDP_Date.Location = New System.Drawing.Point(2, 26)
        Me.RboDP_Date.Margin = New System.Windows.Forms.Padding(2)
        Me.RboDP_Date.Name = "RboDP_Date"
        Me.RboDP_Date.Size = New System.Drawing.Size(87, 17)
        Me.RboDP_Date.TabIndex = 3
        Me.RboDP_Date.Text = "DP Inv. Date"
        Me.RboDP_Date.UseVisualStyleBackColor = True
        '
        'RboSO_Date
        '
        Me.RboSO_Date.AutoSize = True
        Me.RboSO_Date.Checked = True
        Me.RboSO_Date.Location = New System.Drawing.Point(2, 3)
        Me.RboSO_Date.Margin = New System.Windows.Forms.Padding(2)
        Me.RboSO_Date.Name = "RboSO_Date"
        Me.RboSO_Date.Size = New System.Drawing.Size(66, 17)
        Me.RboSO_Date.TabIndex = 0
        Me.RboSO_Date.TabStop = True
        Me.RboSO_Date.Text = "SO Date"
        Me.RboSO_Date.UseVisualStyleBackColor = True
        '
        'TxtPrimaryCustomer
        '
        Me.TxtPrimaryCustomer.Location = New System.Drawing.Point(59, 70)
        Me.TxtPrimaryCustomer.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtPrimaryCustomer.Name = "TxtPrimaryCustomer"
        Me.TxtPrimaryCustomer.Size = New System.Drawing.Size(31, 20)
        Me.TxtPrimaryCustomer.TabIndex = 12
        Me.TxtPrimaryCustomer.TabStop = False
        Me.TxtPrimaryCustomer.Text = "-1"
        Me.TxtPrimaryCustomer.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(2, 74)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 13)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Customer"
        '
        'TxtCustomerCode
        '
        Me.TxtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerCode.Location = New System.Drawing.Point(94, 69)
        Me.TxtCustomerCode.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtCustomerCode.Name = "TxtCustomerCode"
        Me.TxtCustomerCode.Size = New System.Drawing.Size(67, 20)
        Me.TxtCustomerCode.TabIndex = 9
        '
        'TxtCustomerName
        '
        Me.TxtCustomerName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerName.Location = New System.Drawing.Point(212, 70)
        Me.TxtCustomerName.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtCustomerName.Name = "TxtCustomerName"
        Me.TxtCustomerName.ReadOnly = True
        Me.TxtCustomerName.Size = New System.Drawing.Size(150, 20)
        Me.TxtCustomerName.TabIndex = 11
        Me.TxtCustomerName.TabStop = False
        '
        'BtnFindCustomer
        '
        Me.BtnFindCustomer.Location = New System.Drawing.Point(165, 69)
        Me.BtnFindCustomer.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnFindCustomer.Name = "BtnFindCustomer"
        Me.BtnFindCustomer.Size = New System.Drawing.Size(43, 20)
        Me.BtnFindCustomer.TabIndex = 10
        Me.BtnFindCustomer.Text = "Find"
        Me.BtnFindCustomer.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(212, 28)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "To"
        '
        'DtpDP_DateTo
        '
        Me.DtpDP_DateTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpDP_DateTo.CustomFormat = "dd MMM yyyy"
        Me.DtpDP_DateTo.Enabled = False
        Me.DtpDP_DateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDP_DateTo.Location = New System.Drawing.Point(235, 24)
        Me.DtpDP_DateTo.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpDP_DateTo.Name = "DtpDP_DateTo"
        Me.DtpDP_DateTo.Size = New System.Drawing.Size(126, 20)
        Me.DtpDP_DateTo.TabIndex = 5
        '
        'DtpDP_DateFrom
        '
        Me.DtpDP_DateFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpDP_DateFrom.Enabled = False
        Me.DtpDP_DateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDP_DateFrom.Location = New System.Drawing.Point(94, 24)
        Me.DtpDP_DateFrom.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpDP_DateFrom.Name = "DtpDP_DateFrom"
        Me.DtpDP_DateFrom.Size = New System.Drawing.Size(114, 20)
        Me.DtpDP_DateFrom.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(212, 5)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(20, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "To"
        '
        'DtpSO_DateTo
        '
        Me.DtpSO_DateTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpSO_DateTo.CustomFormat = "dd MMM yyyy"
        Me.DtpSO_DateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpSO_DateTo.Location = New System.Drawing.Point(235, 1)
        Me.DtpSO_DateTo.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpSO_DateTo.Name = "DtpSO_DateTo"
        Me.DtpSO_DateTo.Size = New System.Drawing.Size(126, 20)
        Me.DtpSO_DateTo.TabIndex = 2
        '
        'DtpSO_DateFrom
        '
        Me.DtpSO_DateFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpSO_DateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpSO_DateFrom.Location = New System.Drawing.Point(94, 1)
        Me.DtpSO_DateFrom.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpSO_DateFrom.Name = "DtpSO_DateFrom"
        Me.DtpSO_DateFrom.Size = New System.Drawing.Size(114, 20)
        Me.DtpSO_DateFrom.TabIndex = 1
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(191, 93)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(52, 21)
        Me.BtnClose.TabIndex = 15
        Me.BtnClose.Text = "&Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'BtnFilter
        '
        Me.BtnFilter.Location = New System.Drawing.Point(135, 93)
        Me.BtnFilter.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnFilter.Name = "BtnFilter"
        Me.BtnFilter.Size = New System.Drawing.Size(52, 21)
        Me.BtnFilter.TabIndex = 14
        Me.BtnFilter.Text = "&Filter"
        Me.BtnFilter.UseVisualStyleBackColor = True
        '
        'CboGroup
        '
        Me.CboGroup.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboGroup.FormattingEnabled = True
        Me.CboGroup.Items.AddRange(New Object() {"By Customer", "By Sales Person"})
        Me.CboGroup.Location = New System.Drawing.Point(135, 70)
        Me.CboGroup.Margin = New System.Windows.Forms.Padding(2)
        Me.CboGroup.Name = "CboGroup"
        Me.CboGroup.Size = New System.Drawing.Size(221, 21)
        Me.CboGroup.TabIndex = 3
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(2, 74)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(50, 13)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "Grouping"
        '
        'CboRealizationStatus
        '
        Me.CboRealizationStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboRealizationStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboRealizationStatus.FormattingEnabled = True
        Me.CboRealizationStatus.Items.AddRange(New Object() {"All", "Realized", "Partially Realized/Not Realized"})
        Me.CboRealizationStatus.Location = New System.Drawing.Point(135, 47)
        Me.CboRealizationStatus.Margin = New System.Windows.Forms.Padding(2)
        Me.CboRealizationStatus.Name = "CboRealizationStatus"
        Me.CboRealizationStatus.Size = New System.Drawing.Size(221, 21)
        Me.CboRealizationStatus.TabIndex = 2
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(2, 51)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(92, 13)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Realization Status"
        '
        'CboPaymentStatus
        '
        Me.CboPaymentStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboPaymentStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPaymentStatus.FormattingEnabled = True
        Me.CboPaymentStatus.Items.AddRange(New Object() {"All", "Paid", "Partially Paid/No Payment"})
        Me.CboPaymentStatus.Location = New System.Drawing.Point(135, 24)
        Me.CboPaymentStatus.Margin = New System.Windows.Forms.Padding(2)
        Me.CboPaymentStatus.Name = "CboPaymentStatus"
        Me.CboPaymentStatus.Size = New System.Drawing.Size(221, 21)
        Me.CboPaymentStatus.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(2, 28)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(81, 13)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "Payment Status"
        '
        'CboDPStatus
        '
        Me.CboDPStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboDPStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboDPStatus.FormattingEnabled = True
        Me.CboDPStatus.Items.AddRange(New Object() {"All", "Downpayment Created", "Downpayment Not Created"})
        Me.CboDPStatus.Location = New System.Drawing.Point(135, 1)
        Me.CboDPStatus.Margin = New System.Windows.Forms.Padding(2)
        Me.CboDPStatus.Name = "CboDPStatus"
        Me.CboDPStatus.Size = New System.Drawing.Size(221, 21)
        Me.CboDPStatus.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(2, 5)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 13)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Down Payment Status"
        '
        'FrmLaporanSalesDP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(744, 514)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.crv)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(760, 553)
        Me.Name = "FrmLaporanSalesDP"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales DownPayment Report"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents RboInvoiceDate As System.Windows.Forms.RadioButton
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents DtpInvoice_DateTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents DtpInvoice_DateFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents RboDP_Date As System.Windows.Forms.RadioButton
    Friend WithEvents RboSO_Date As System.Windows.Forms.RadioButton
    Friend WithEvents TxtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindCustomer As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DtpDP_DateTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpDP_DateFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DtpSO_DateTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpSO_DateFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents CboGroup As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents CboRealizationStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents CboPaymentStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CboDPStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimarySalesPerson As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtSalesPersonCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtSalesPersonDesc As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindSalesPerson As System.Windows.Forms.Button
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents BtnFilter As System.Windows.Forms.Button
End Class
