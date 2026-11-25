<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSalesInvoice_AddData
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.gboCost = New System.Windows.Forms.GroupBox
        Me.dgvCost = New System.Windows.Forms.DataGridView
        Me.ColNoCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrencyCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostNominal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExcRateCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarksCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCurrencyCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccounrCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccountDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.txtPrimaryUnitVolume = New System.Windows.Forms.TextBox
        Me.txtPrimaryUnitWeight = New System.Windows.Forms.TextBox
        Me.txtVar2_Name = New System.Windows.Forms.TextBox
        Me.txtVar1_Name = New System.Windows.Forms.TextBox
        Me.txtVar5_Name = New System.Windows.Forms.TextBox
        Me.txtVar4_Name = New System.Windows.Forms.TextBox
        Me.txtVar3_Name = New System.Windows.Forms.TextBox
        Me.txtVar5_Value = New System.Windows.Forms.TextBox
        Me.txtVar4_Value = New System.Windows.Forms.TextBox
        Me.txtVar3_Value = New System.Windows.Forms.TextBox
        Me.txtVar2_Value = New System.Windows.Forms.TextBox
        Me.txtVar1_Value = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtTransportationMode = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtDestinationPort = New System.Windows.Forms.TextBox
        Me.dtpDestinationDate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtLoadingPort = New System.Windows.Forms.TextBox
        Me.dtpLoadingDate = New System.Windows.Forms.DateTimePicker
        Me.btnFindUnitVolume = New System.Windows.Forms.Button
        Me.txtUnitVolume = New System.Windows.Forms.TextBox
        Me.txtGrossVolume = New System.Windows.Forms.TextBox
        Me.txtNettVolume = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnFindUnitWeight = New System.Windows.Forms.Button
        Me.txtUnitWeight = New System.Windows.Forms.TextBox
        Me.txtGrossWeight = New System.Windows.Forms.TextBox
        Me.txtNettWeight = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.TxtPrimaryCustomerShipTo = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.LlbViewAddressCustomerShipTo = New System.Windows.Forms.LinkLabel
        Me.TxtIdentifierAddressCustomerShipTo = New System.Windows.Forms.TextBox
        Me.CboAddressCustomerShipTo = New System.Windows.Forms.ComboBox
        Me.TxtCustomerCodeShipTo = New System.Windows.Forms.TextBox
        Me.TxtCustomerDescShipTo = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtShippingMarks = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtInvoiceType = New System.Windows.Forms.TextBox
        Me.gboAll = New System.Windows.Forms.GroupBox
        Me.gboCost.SuspendLayout()
        CType(Me.dgvCost, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.gboAll.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(662, 425)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(69, 24)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(588, 425)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(69, 24)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'gboCost
        '
        Me.gboCost.Controls.Add(Me.dgvCost)
        Me.gboCost.Location = New System.Drawing.Point(3, 242)
        Me.gboCost.Name = "gboCost"
        Me.gboCost.Size = New System.Drawing.Size(726, 178)
        Me.gboCost.TabIndex = 1
        Me.gboCost.TabStop = False
        Me.gboCost.Text = "Cost Data"
        '
        'dgvCost
        '
        Me.dgvCost.AllowUserToResizeColumns = False
        Me.dgvCost.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvCost.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvCost.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCost.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNoCost, Me.ColDescCost, Me.ColCurrencyCost, Me.ColCostNominal, Me.ColExcRateCost, Me.ColCostAmount, Me.ColRemarksCost, Me.ColPrimaryCurrencyCost, Me.ColAccounrCode, Me.ColAccountDesc, Me.ColPrimaryAccount})
        Me.dgvCost.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCost.Location = New System.Drawing.Point(3, 16)
        Me.dgvCost.MultiSelect = False
        Me.dgvCost.Name = "dgvCost"
        Me.dgvCost.RowTemplate.Height = 24
        Me.dgvCost.Size = New System.Drawing.Size(720, 159)
        Me.dgvCost.StandardTab = True
        Me.dgvCost.TabIndex = 0
        '
        'ColNoCost
        '
        Me.ColNoCost.Frozen = True
        Me.ColNoCost.HeaderText = "No"
        Me.ColNoCost.Name = "ColNoCost"
        Me.ColNoCost.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNoCost.Width = 27
        '
        'ColDescCost
        '
        Me.ColDescCost.Frozen = True
        Me.ColDescCost.HeaderText = "Description"
        Me.ColDescCost.Name = "ColDescCost"
        Me.ColDescCost.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDescCost.Width = 66
        '
        'ColCurrencyCost
        '
        Me.ColCurrencyCost.HeaderText = "Currency"
        Me.ColCurrencyCost.Name = "ColCurrencyCost"
        Me.ColCurrencyCost.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCurrencyCost.Width = 55
        '
        'ColCostNominal
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColCostNominal.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColCostNominal.HeaderText = "Cost"
        Me.ColCostNominal.Name = "ColCostNominal"
        Me.ColCostNominal.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCostNominal.Width = 34
        '
        'ColExcRateCost
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColExcRateCost.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColExcRateCost.HeaderText = "Exhange Rate to IDR"
        Me.ColExcRateCost.Name = "ColExcRateCost"
        Me.ColExcRateCost.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColExcRateCost.Visible = False
        Me.ColExcRateCost.Width = 115
        '
        'ColCostAmount
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColCostAmount.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColCostAmount.HeaderText = "Cost Amount (IDR)"
        Me.ColCostAmount.Name = "ColCostAmount"
        Me.ColCostAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCostAmount.Visible = False
        Me.ColCostAmount.Width = 101
        '
        'ColRemarksCost
        '
        Me.ColRemarksCost.HeaderText = "Remarks"
        Me.ColRemarksCost.Name = "ColRemarksCost"
        Me.ColRemarksCost.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarksCost.Width = 55
        '
        'ColPrimaryCurrencyCost
        '
        Me.ColPrimaryCurrencyCost.HeaderText = "Primary Currency"
        Me.ColPrimaryCurrencyCost.Name = "ColPrimaryCurrencyCost"
        Me.ColPrimaryCurrencyCost.Visible = False
        Me.ColPrimaryCurrencyCost.Width = 111
        '
        'ColAccounrCode
        '
        Me.ColAccounrCode.HeaderText = "Account No"
        Me.ColAccounrCode.Name = "ColAccounrCode"
        Me.ColAccounrCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccounrCode.Width = 70
        '
        'ColAccountDesc
        '
        Me.ColAccountDesc.HeaderText = "Account Description"
        Me.ColAccountDesc.Name = "ColAccountDesc"
        Me.ColAccountDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccountDesc.Width = 109
        '
        'ColPrimaryAccount
        '
        Me.ColPrimaryAccount.HeaderText = "Primary Account"
        Me.ColPrimaryAccount.Name = "ColPrimaryAccount"
        Me.ColPrimaryAccount.Visible = False
        Me.ColPrimaryAccount.Width = 109
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(5, 9)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryUnitVolume)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryUnitWeight)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtVar2_Name)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtVar1_Name)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtVar5_Name)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtVar4_Name)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtVar3_Name)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtVar5_Value)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtVar4_Value)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtVar3_Value)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtVar2_Value)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtVar1_Value)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtTransportationMode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtDestinationPort)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDestinationDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtLoadingPort)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpLoadingDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindUnitVolume)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtUnitVolume)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtGrossVolume)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtNettVolume)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindUnitWeight)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtUnitWeight)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtGrossWeight)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtNettWeight)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtShippingMarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtInvoiceType)
        Me.SplitContainer1.Size = New System.Drawing.Size(722, 228)
        Me.SplitContainer1.SplitterDistance = 357
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'txtPrimaryUnitVolume
        '
        Me.txtPrimaryUnitVolume.Location = New System.Drawing.Point(85, 23)
        Me.txtPrimaryUnitVolume.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPrimaryUnitVolume.Name = "txtPrimaryUnitVolume"
        Me.txtPrimaryUnitVolume.Size = New System.Drawing.Size(21, 20)
        Me.txtPrimaryUnitVolume.TabIndex = 7
        Me.txtPrimaryUnitVolume.Text = "-1"
        Me.txtPrimaryUnitVolume.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryUnitVolume.Visible = False
        '
        'txtPrimaryUnitWeight
        '
        Me.txtPrimaryUnitWeight.Location = New System.Drawing.Point(85, 2)
        Me.txtPrimaryUnitWeight.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPrimaryUnitWeight.Name = "txtPrimaryUnitWeight"
        Me.txtPrimaryUnitWeight.Size = New System.Drawing.Size(21, 20)
        Me.txtPrimaryUnitWeight.TabIndex = 1
        Me.txtPrimaryUnitWeight.Text = "-1"
        Me.txtPrimaryUnitWeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryUnitWeight.Visible = False
        '
        'txtVar2_Name
        '
        Me.txtVar2_Name.Location = New System.Drawing.Point(4, 139)
        Me.txtVar2_Name.Margin = New System.Windows.Forms.Padding(2)
        Me.txtVar2_Name.MaxLength = 250
        Me.txtVar2_Name.Name = "txtVar2_Name"
        Me.txtVar2_Name.Size = New System.Drawing.Size(102, 20)
        Me.txtVar2_Name.TabIndex = 19
        '
        'txtVar1_Name
        '
        Me.txtVar1_Name.Location = New System.Drawing.Point(4, 116)
        Me.txtVar1_Name.Margin = New System.Windows.Forms.Padding(2)
        Me.txtVar1_Name.MaxLength = 250
        Me.txtVar1_Name.Name = "txtVar1_Name"
        Me.txtVar1_Name.Size = New System.Drawing.Size(102, 20)
        Me.txtVar1_Name.TabIndex = 17
        '
        'txtVar5_Name
        '
        Me.txtVar5_Name.Location = New System.Drawing.Point(4, 207)
        Me.txtVar5_Name.Margin = New System.Windows.Forms.Padding(2)
        Me.txtVar5_Name.MaxLength = 250
        Me.txtVar5_Name.Name = "txtVar5_Name"
        Me.txtVar5_Name.Size = New System.Drawing.Size(102, 20)
        Me.txtVar5_Name.TabIndex = 25
        '
        'txtVar4_Name
        '
        Me.txtVar4_Name.Location = New System.Drawing.Point(4, 184)
        Me.txtVar4_Name.Margin = New System.Windows.Forms.Padding(2)
        Me.txtVar4_Name.MaxLength = 250
        Me.txtVar4_Name.Name = "txtVar4_Name"
        Me.txtVar4_Name.Size = New System.Drawing.Size(102, 20)
        Me.txtVar4_Name.TabIndex = 23
        '
        'txtVar3_Name
        '
        Me.txtVar3_Name.Location = New System.Drawing.Point(4, 162)
        Me.txtVar3_Name.Margin = New System.Windows.Forms.Padding(2)
        Me.txtVar3_Name.MaxLength = 250
        Me.txtVar3_Name.Name = "txtVar3_Name"
        Me.txtVar3_Name.Size = New System.Drawing.Size(102, 20)
        Me.txtVar3_Name.TabIndex = 21
        '
        'txtVar5_Value
        '
        Me.txtVar5_Value.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtVar5_Value.Location = New System.Drawing.Point(110, 207)
        Me.txtVar5_Value.Margin = New System.Windows.Forms.Padding(2)
        Me.txtVar5_Value.MaxLength = 250
        Me.txtVar5_Value.Name = "txtVar5_Value"
        Me.txtVar5_Value.Size = New System.Drawing.Size(246, 20)
        Me.txtVar5_Value.TabIndex = 26
        '
        'txtVar4_Value
        '
        Me.txtVar4_Value.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtVar4_Value.Location = New System.Drawing.Point(110, 184)
        Me.txtVar4_Value.Margin = New System.Windows.Forms.Padding(2)
        Me.txtVar4_Value.MaxLength = 250
        Me.txtVar4_Value.Name = "txtVar4_Value"
        Me.txtVar4_Value.Size = New System.Drawing.Size(246, 20)
        Me.txtVar4_Value.TabIndex = 24
        '
        'txtVar3_Value
        '
        Me.txtVar3_Value.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtVar3_Value.Location = New System.Drawing.Point(110, 162)
        Me.txtVar3_Value.Margin = New System.Windows.Forms.Padding(2)
        Me.txtVar3_Value.MaxLength = 250
        Me.txtVar3_Value.Name = "txtVar3_Value"
        Me.txtVar3_Value.Size = New System.Drawing.Size(246, 20)
        Me.txtVar3_Value.TabIndex = 22
        '
        'txtVar2_Value
        '
        Me.txtVar2_Value.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtVar2_Value.Location = New System.Drawing.Point(110, 139)
        Me.txtVar2_Value.Margin = New System.Windows.Forms.Padding(2)
        Me.txtVar2_Value.MaxLength = 250
        Me.txtVar2_Value.Name = "txtVar2_Value"
        Me.txtVar2_Value.Size = New System.Drawing.Size(246, 20)
        Me.txtVar2_Value.TabIndex = 20
        '
        'txtVar1_Value
        '
        Me.txtVar1_Value.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtVar1_Value.Location = New System.Drawing.Point(110, 116)
        Me.txtVar1_Value.Margin = New System.Windows.Forms.Padding(2)
        Me.txtVar1_Value.MaxLength = 250
        Me.txtVar1_Value.Name = "txtVar1_Value"
        Me.txtVar1_Value.Size = New System.Drawing.Size(246, 20)
        Me.txtVar1_Value.TabIndex = 18
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(2, 96)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(105, 13)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Transportation Mode"
        '
        'txtTransportationMode
        '
        Me.txtTransportationMode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTransportationMode.Location = New System.Drawing.Point(110, 93)
        Me.txtTransportationMode.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTransportationMode.MaxLength = 250
        Me.txtTransportationMode.Name = "txtTransportationMode"
        Me.txtTransportationMode.Size = New System.Drawing.Size(246, 20)
        Me.txtTransportationMode.TabIndex = 16
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(2, 73)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Destination Port"
        '
        'txtDestinationPort
        '
        Me.txtDestinationPort.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDestinationPort.Location = New System.Drawing.Point(110, 71)
        Me.txtDestinationPort.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDestinationPort.MaxLength = 250
        Me.txtDestinationPort.Name = "txtDestinationPort"
        Me.txtDestinationPort.Size = New System.Drawing.Size(134, 20)
        Me.txtDestinationPort.TabIndex = 14
        '
        'dtpDestinationDate
        '
        Me.dtpDestinationDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDestinationDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDestinationDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDestinationDate.Location = New System.Drawing.Point(248, 71)
        Me.dtpDestinationDate.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpDestinationDate.Name = "dtpDestinationDate"
        Me.dtpDestinationDate.Size = New System.Drawing.Size(108, 20)
        Me.dtpDestinationDate.TabIndex = 15
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(2, 50)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Loading Port"
        '
        'txtLoadingPort
        '
        Me.txtLoadingPort.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLoadingPort.Location = New System.Drawing.Point(110, 48)
        Me.txtLoadingPort.Margin = New System.Windows.Forms.Padding(2)
        Me.txtLoadingPort.MaxLength = 250
        Me.txtLoadingPort.Name = "txtLoadingPort"
        Me.txtLoadingPort.Size = New System.Drawing.Size(134, 20)
        Me.txtLoadingPort.TabIndex = 12
        '
        'dtpLoadingDate
        '
        Me.dtpLoadingDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpLoadingDate.CustomFormat = "dd MMM yyyy"
        Me.dtpLoadingDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpLoadingDate.Location = New System.Drawing.Point(248, 48)
        Me.dtpLoadingDate.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpLoadingDate.Name = "dtpLoadingDate"
        Me.dtpLoadingDate.Size = New System.Drawing.Size(108, 20)
        Me.dtpLoadingDate.TabIndex = 13
        '
        'btnFindUnitVolume
        '
        Me.btnFindUnitVolume.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindUnitVolume.Location = New System.Drawing.Point(310, 24)
        Me.btnFindUnitVolume.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFindUnitVolume.Name = "btnFindUnitVolume"
        Me.btnFindUnitVolume.Size = New System.Drawing.Size(45, 21)
        Me.btnFindUnitVolume.TabIndex = 11
        Me.btnFindUnitVolume.TabStop = False
        Me.btnFindUnitVolume.Text = "Find"
        Me.btnFindUnitVolume.UseVisualStyleBackColor = True
        '
        'txtUnitVolume
        '
        Me.txtUnitVolume.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUnitVolume.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtUnitVolume.Location = New System.Drawing.Point(264, 24)
        Me.txtUnitVolume.Margin = New System.Windows.Forms.Padding(2)
        Me.txtUnitVolume.Name = "txtUnitVolume"
        Me.txtUnitVolume.Size = New System.Drawing.Size(42, 20)
        Me.txtUnitVolume.TabIndex = 10
        '
        'txtGrossVolume
        '
        Me.txtGrossVolume.Location = New System.Drawing.Point(187, 24)
        Me.txtGrossVolume.Margin = New System.Windows.Forms.Padding(2)
        Me.txtGrossVolume.Name = "txtGrossVolume"
        Me.txtGrossVolume.Size = New System.Drawing.Size(74, 20)
        Me.txtGrossVolume.TabIndex = 9
        Me.txtGrossVolume.Text = "0"
        Me.txtGrossVolume.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtNettVolume
        '
        Me.txtNettVolume.Location = New System.Drawing.Point(110, 24)
        Me.txtNettVolume.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNettVolume.Name = "txtNettVolume"
        Me.txtNettVolume.Size = New System.Drawing.Size(74, 20)
        Me.txtNettVolume.TabIndex = 8
        Me.txtNettVolume.Text = "0"
        Me.txtNettVolume.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(2, 28)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(97, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Nett/Gross Volume"
        '
        'btnFindUnitWeight
        '
        Me.btnFindUnitWeight.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindUnitWeight.Location = New System.Drawing.Point(310, 1)
        Me.btnFindUnitWeight.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFindUnitWeight.Name = "btnFindUnitWeight"
        Me.btnFindUnitWeight.Size = New System.Drawing.Size(45, 21)
        Me.btnFindUnitWeight.TabIndex = 5
        Me.btnFindUnitWeight.TabStop = False
        Me.btnFindUnitWeight.Text = "Find"
        Me.btnFindUnitWeight.UseVisualStyleBackColor = True
        '
        'txtUnitWeight
        '
        Me.txtUnitWeight.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUnitWeight.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtUnitWeight.Location = New System.Drawing.Point(264, 1)
        Me.txtUnitWeight.Margin = New System.Windows.Forms.Padding(2)
        Me.txtUnitWeight.Name = "txtUnitWeight"
        Me.txtUnitWeight.Size = New System.Drawing.Size(42, 20)
        Me.txtUnitWeight.TabIndex = 4
        '
        'txtGrossWeight
        '
        Me.txtGrossWeight.Location = New System.Drawing.Point(187, 1)
        Me.txtGrossWeight.Margin = New System.Windows.Forms.Padding(2)
        Me.txtGrossWeight.Name = "txtGrossWeight"
        Me.txtGrossWeight.Size = New System.Drawing.Size(74, 20)
        Me.txtGrossWeight.TabIndex = 3
        Me.txtGrossWeight.Text = "0"
        Me.txtGrossWeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtNettWeight
        '
        Me.txtNettWeight.Location = New System.Drawing.Point(110, 1)
        Me.txtNettWeight.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNettWeight.Name = "txtNettWeight"
        Me.txtNettWeight.Size = New System.Drawing.Size(74, 20)
        Me.txtNettWeight.TabIndex = 2
        Me.txtNettWeight.Text = "0"
        Me.txtNettWeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 5)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Nett/Gross Weight"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.TxtPrimaryCustomerShipTo)
        Me.Panel2.Controls.Add(Me.Label21)
        Me.Panel2.Controls.Add(Me.LlbViewAddressCustomerShipTo)
        Me.Panel2.Controls.Add(Me.TxtIdentifierAddressCustomerShipTo)
        Me.Panel2.Controls.Add(Me.CboAddressCustomerShipTo)
        Me.Panel2.Controls.Add(Me.TxtCustomerCodeShipTo)
        Me.Panel2.Controls.Add(Me.TxtCustomerDescShipTo)
        Me.Panel2.Controls.Add(Me.Label13)
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(362, 47)
        Me.Panel2.TabIndex = 313
        '
        'TxtPrimaryCustomerShipTo
        '
        Me.TxtPrimaryCustomerShipTo.Location = New System.Drawing.Point(94, 4)
        Me.TxtPrimaryCustomerShipTo.Name = "TxtPrimaryCustomerShipTo"
        Me.TxtPrimaryCustomerShipTo.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryCustomerShipTo.TabIndex = 312
        Me.TxtPrimaryCustomerShipTo.TabStop = False
        Me.TxtPrimaryCustomerShipTo.Text = "-1"
        Me.TxtPrimaryCustomerShipTo.Visible = False
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(1, 30)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(81, 13)
        Me.Label21.TabIndex = 311
        Me.Label21.Text = "Ship to Address"
        '
        'LlbViewAddressCustomerShipTo
        '
        Me.LlbViewAddressCustomerShipTo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbViewAddressCustomerShipTo.AutoSize = True
        Me.LlbViewAddressCustomerShipTo.Location = New System.Drawing.Point(325, 30)
        Me.LlbViewAddressCustomerShipTo.Name = "LlbViewAddressCustomerShipTo"
        Me.LlbViewAddressCustomerShipTo.Size = New System.Drawing.Size(29, 13)
        Me.LlbViewAddressCustomerShipTo.TabIndex = 198
        Me.LlbViewAddressCustomerShipTo.TabStop = True
        Me.LlbViewAddressCustomerShipTo.Text = "view"
        '
        'TxtIdentifierAddressCustomerShipTo
        '
        Me.TxtIdentifierAddressCustomerShipTo.Location = New System.Drawing.Point(94, 26)
        Me.TxtIdentifierAddressCustomerShipTo.Name = "TxtIdentifierAddressCustomerShipTo"
        Me.TxtIdentifierAddressCustomerShipTo.Size = New System.Drawing.Size(25, 20)
        Me.TxtIdentifierAddressCustomerShipTo.TabIndex = 199
        Me.TxtIdentifierAddressCustomerShipTo.TabStop = False
        Me.TxtIdentifierAddressCustomerShipTo.Text = "-1"
        Me.TxtIdentifierAddressCustomerShipTo.Visible = False
        '
        'CboAddressCustomerShipTo
        '
        Me.CboAddressCustomerShipTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboAddressCustomerShipTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboAddressCustomerShipTo.Enabled = False
        Me.CboAddressCustomerShipTo.FormattingEnabled = True
        Me.CboAddressCustomerShipTo.Location = New System.Drawing.Point(94, 26)
        Me.CboAddressCustomerShipTo.Name = "CboAddressCustomerShipTo"
        Me.CboAddressCustomerShipTo.Size = New System.Drawing.Size(226, 21)
        Me.CboAddressCustomerShipTo.TabIndex = 197
        '
        'TxtCustomerCodeShipTo
        '
        Me.TxtCustomerCodeShipTo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerCodeShipTo.Location = New System.Drawing.Point(94, 4)
        Me.TxtCustomerCodeShipTo.MaxLength = 50
        Me.TxtCustomerCodeShipTo.Name = "TxtCustomerCodeShipTo"
        Me.TxtCustomerCodeShipTo.ReadOnly = True
        Me.TxtCustomerCodeShipTo.Size = New System.Drawing.Size(102, 20)
        Me.TxtCustomerCodeShipTo.TabIndex = 186
        '
        'TxtCustomerDescShipTo
        '
        Me.TxtCustomerDescShipTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerDescShipTo.Location = New System.Drawing.Point(202, 4)
        Me.TxtCustomerDescShipTo.MaxLength = 150
        Me.TxtCustomerDescShipTo.Name = "TxtCustomerDescShipTo"
        Me.TxtCustomerDescShipTo.ReadOnly = True
        Me.TxtCustomerDescShipTo.Size = New System.Drawing.Size(152, 20)
        Me.TxtCustomerDescShipTo.TabIndex = 187
        Me.TxtCustomerDescShipTo.TabStop = False
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(1, 8)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(87, 13)
        Me.Label13.TabIndex = 188
        Me.Label13.Text = "Ship to Customer"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(1, 73)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(80, 13)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "Shipping Marks"
        '
        'txtShippingMarks
        '
        Me.txtShippingMarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtShippingMarks.Location = New System.Drawing.Point(94, 71)
        Me.txtShippingMarks.Margin = New System.Windows.Forms.Padding(2)
        Me.txtShippingMarks.MaxLength = 1500
        Me.txtShippingMarks.Multiline = True
        Me.txtShippingMarks.Name = "txtShippingMarks"
        Me.txtShippingMarks.Size = New System.Drawing.Size(263, 155)
        Me.txtShippingMarks.TabIndex = 10
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(1, 53)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(69, 13)
        Me.Label9.TabIndex = 7
        Me.Label9.Text = "Invoice Type"
        '
        'txtInvoiceType
        '
        Me.txtInvoiceType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtInvoiceType.Location = New System.Drawing.Point(94, 49)
        Me.txtInvoiceType.Margin = New System.Windows.Forms.Padding(2)
        Me.txtInvoiceType.MaxLength = 250
        Me.txtInvoiceType.Name = "txtInvoiceType"
        Me.txtInvoiceType.Size = New System.Drawing.Size(263, 20)
        Me.txtInvoiceType.TabIndex = 8
        '
        'gboAll
        '
        Me.gboAll.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboAll.Controls.Add(Me.gboCost)
        Me.gboAll.Controls.Add(Me.SplitContainer1)
        Me.gboAll.Location = New System.Drawing.Point(4, 2)
        Me.gboAll.Name = "gboAll"
        Me.gboAll.Size = New System.Drawing.Size(732, 427)
        Me.gboAll.TabIndex = 2
        Me.gboAll.TabStop = False
        '
        'FrmTransaksiSalesInvoice_AddData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(740, 456)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.gboAll)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(756, 495)
        Me.Name = "FrmTransaksiSalesInvoice_AddData"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Additional Data"
        Me.gboCost.ResumeLayout(False)
        CType(Me.dgvCost, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.gboAll.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents gboCost As System.Windows.Forms.GroupBox
    Public WithEvents dgvCost As System.Windows.Forms.DataGridView
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnFindUnitVolume As System.Windows.Forms.Button
    Friend WithEvents txtUnitVolume As System.Windows.Forms.TextBox
    Friend WithEvents txtGrossVolume As System.Windows.Forms.TextBox
    Friend WithEvents txtNettVolume As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnFindUnitWeight As System.Windows.Forms.Button
    Friend WithEvents txtUnitWeight As System.Windows.Forms.TextBox
    Friend WithEvents txtGrossWeight As System.Windows.Forms.TextBox
    Friend WithEvents txtNettWeight As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtVar2_Value As System.Windows.Forms.TextBox
    Friend WithEvents txtVar1_Value As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtTransportationMode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtDestinationPort As System.Windows.Forms.TextBox
    Friend WithEvents dtpDestinationDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtLoadingPort As System.Windows.Forms.TextBox
    Friend WithEvents dtpLoadingDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtShippingMarks As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtInvoiceType As System.Windows.Forms.TextBox
    Friend WithEvents txtVar2_Name As System.Windows.Forms.TextBox
    Friend WithEvents txtVar1_Name As System.Windows.Forms.TextBox
    Friend WithEvents txtVar5_Name As System.Windows.Forms.TextBox
    Friend WithEvents txtVar4_Name As System.Windows.Forms.TextBox
    Friend WithEvents txtVar3_Name As System.Windows.Forms.TextBox
    Friend WithEvents txtVar5_Value As System.Windows.Forms.TextBox
    Friend WithEvents txtVar4_Value As System.Windows.Forms.TextBox
    Friend WithEvents txtVar3_Value As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryUnitVolume As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryUnitWeight As System.Windows.Forms.TextBox
    Friend WithEvents gboAll As System.Windows.Forms.GroupBox
    Friend WithEvents ColNoCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrencyCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostNominal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExcRateCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarksCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCurrencyCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccounrCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccountDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TxtPrimaryCustomerShipTo As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents LlbViewAddressCustomerShipTo As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtIdentifierAddressCustomerShipTo As System.Windows.Forms.TextBox
    Friend WithEvents CboAddressCustomerShipTo As System.Windows.Forms.ComboBox
    Friend WithEvents TxtCustomerCodeShipTo As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustomerDescShipTo As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
End Class
