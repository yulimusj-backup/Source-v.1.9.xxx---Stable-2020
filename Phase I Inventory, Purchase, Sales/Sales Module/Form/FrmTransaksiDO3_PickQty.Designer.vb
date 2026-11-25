<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiDO3_PickQty
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.LblRatio = New System.Windows.Forms.Label
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.TxtCodeNo = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtSONumber = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.LblUnitAvailableMulti = New System.Windows.Forms.Label
        Me.TxtQtyAvailableMulti = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.LblUnitAvailableDefault = New System.Windows.Forms.Label
        Me.TxtQtyAvailableDefault = New System.Windows.Forms.TextBox
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.LblUnitDeliveryMulti = New System.Windows.Forms.Label
        Me.TxtQtyDeliveryMulti = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.LblUnitDeliveryDefault = New System.Windows.Forms.Label
        Me.TxtQtyDeliveryDefault = New System.Windows.Forms.TextBox
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.TxtQtyOrderBalanceMulti = New System.Windows.Forms.TextBox
        Me.LblUnitOrderBalanceMulti = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.LblUnitOrderBalanceDefault = New System.Windows.Forms.Label
        Me.TxtQtyOrderBalanceDefault = New System.Windows.Forms.TextBox
        Me.dgvItemDetail = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyAvailable = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyDelivery = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.llbUnselectAll = New System.Windows.Forms.LinkLabel
        Me.llbSelectAll = New System.Windows.Forms.LinkLabel
        Me.llbLocationAccess = New System.Windows.Forms.LinkLabel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.SplitContainer5 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer6 = New System.Windows.Forms.SplitContainer
        Me.TxtWarehouseSource = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.TxtLocationSource = New System.Windows.Forms.TextBox
        Me.btnFindBarcode = New System.Windows.Forms.Button
        Me.TxtTotalRows = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.TxtBarcodeID = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.BtnFindLocationSource = New System.Windows.Forms.Button
        Me.Label14 = New System.Windows.Forms.Label
        Me.TxtScanHere = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtPrimaryWarehouseSource = New System.Windows.Forms.TextBox
        Me.TxtQtyMulti = New System.Windows.Forms.TextBox
        Me.TxtPrimaryLocationSource = New System.Windows.Forms.TextBox
        Me.BtnInputBarcode = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.TxtQtyDefault = New System.Windows.Forms.TextBox
        Me.TxtWarehouseSourceDesc = New System.Windows.Forms.TextBox
        Me.TxtLocationSourceDesc = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtBarcodeLabel = New System.Windows.Forms.TextBox
        Me.lblCounter = New System.Windows.Forms.Label
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.Col1No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1BarcodeType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1Barcode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1LabelID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1Ratio = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1RatioMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1RatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1QtyDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1UnitDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1Batch = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1ExpiredDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1LocationSource = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1WareHouseSource = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1Remark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1PrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1PrimaryUnitDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1LocationSourcePrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1WarehouseSourcePrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1Decimal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1DecimalDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1LocationControl = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1IsService = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.dgvItemDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SplitContainer5.Panel1.SuspendLayout()
        Me.SplitContainer5.Panel2.SuspendLayout()
        Me.SplitContainer5.SuspendLayout()
        Me.SplitContainer6.Panel1.SuspendLayout()
        Me.SplitContainer6.Panel2.SuspendLayout()
        Me.SplitContainer6.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(9, 10)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblRatio)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDesc)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCodeNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtSONumber)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(713, 67)
        Me.SplitContainer1.SplitterDistance = 356
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'LblRatio
        '
        Me.LblRatio.AutoSize = True
        Me.LblRatio.Location = New System.Drawing.Point(74, 48)
        Me.LblRatio.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblRatio.Name = "LblRatio"
        Me.LblRatio.Size = New System.Drawing.Size(32, 13)
        Me.LblRatio.TabIndex = 5
        Me.LblRatio.Text = "Ratio"
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(170, 25)
        Me.TxtDesc.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.ReadOnly = True
        Me.TxtDesc.Size = New System.Drawing.Size(182, 20)
        Me.TxtDesc.TabIndex = 4
        Me.TxtDesc.TabStop = False
        '
        'TxtCodeNo
        '
        Me.TxtCodeNo.Location = New System.Drawing.Point(76, 25)
        Me.TxtCodeNo.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtCodeNo.Name = "TxtCodeNo"
        Me.TxtCodeNo.ReadOnly = True
        Me.TxtCodeNo.Size = New System.Drawing.Size(91, 20)
        Me.TxtCodeNo.TabIndex = 3
        Me.TxtCodeNo.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(2, 28)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Code No"
        '
        'TxtSONumber
        '
        Me.TxtSONumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSONumber.Location = New System.Drawing.Point(76, 2)
        Me.TxtSONumber.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtSONumber.Name = "TxtSONumber"
        Me.TxtSONumber.ReadOnly = True
        Me.TxtSONumber.Size = New System.Drawing.Size(276, 20)
        Me.TxtSONumber.TabIndex = 2
        Me.TxtSONumber.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 5)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "SO Number"
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer4.Location = New System.Drawing.Point(3, 46)
        Me.SplitContainer4.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.LblUnitAvailableMulti)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtQtyAvailableMulti)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label10)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.LblUnitAvailableDefault)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtQtyAvailableDefault)
        Me.SplitContainer4.Size = New System.Drawing.Size(350, 20)
        Me.SplitContainer4.SplitterDistance = 218
        Me.SplitContainer4.SplitterWidth = 3
        Me.SplitContainer4.TabIndex = 11
        Me.SplitContainer4.TabStop = False
        '
        'LblUnitAvailableMulti
        '
        Me.LblUnitAvailableMulti.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnitAvailableMulti.AutoSize = True
        Me.LblUnitAvailableMulti.Location = New System.Drawing.Point(140, 2)
        Me.LblUnitAvailableMulti.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblUnitAvailableMulti.Name = "LblUnitAvailableMulti"
        Me.LblUnitAvailableMulti.Size = New System.Drawing.Size(26, 13)
        Me.LblUnitAvailableMulti.TabIndex = 2
        Me.LblUnitAvailableMulti.Text = "Unit"
        '
        'TxtQtyAvailableMulti
        '
        Me.TxtQtyAvailableMulti.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyAvailableMulti.Location = New System.Drawing.Point(80, 0)
        Me.TxtQtyAvailableMulti.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtQtyAvailableMulti.Name = "TxtQtyAvailableMulti"
        Me.TxtQtyAvailableMulti.ReadOnly = True
        Me.TxtQtyAvailableMulti.Size = New System.Drawing.Size(57, 20)
        Me.TxtQtyAvailableMulti.TabIndex = 1
        Me.TxtQtyAvailableMulti.TabStop = False
        Me.TxtQtyAvailableMulti.Text = "0"
        Me.TxtQtyAvailableMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(1, 2)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(69, 13)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "Available Qty"
        '
        'LblUnitAvailableDefault
        '
        Me.LblUnitAvailableDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnitAvailableDefault.AutoSize = True
        Me.LblUnitAvailableDefault.Location = New System.Drawing.Point(71, 2)
        Me.LblUnitAvailableDefault.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblUnitAvailableDefault.Name = "LblUnitAvailableDefault"
        Me.LblUnitAvailableDefault.Size = New System.Drawing.Size(26, 13)
        Me.LblUnitAvailableDefault.TabIndex = 1
        Me.LblUnitAvailableDefault.Text = "Unit"
        '
        'TxtQtyAvailableDefault
        '
        Me.TxtQtyAvailableDefault.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyAvailableDefault.Location = New System.Drawing.Point(2, 0)
        Me.TxtQtyAvailableDefault.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtQtyAvailableDefault.Name = "TxtQtyAvailableDefault"
        Me.TxtQtyAvailableDefault.ReadOnly = True
        Me.TxtQtyAvailableDefault.Size = New System.Drawing.Size(67, 20)
        Me.TxtQtyAvailableDefault.TabIndex = 0
        Me.TxtQtyAvailableDefault.TabStop = False
        Me.TxtQtyAvailableDefault.Text = "0"
        Me.TxtQtyAvailableDefault.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer3.Location = New System.Drawing.Point(3, 23)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.LblUnitDeliveryMulti)
        Me.SplitContainer3.Panel1.Controls.Add(Me.TxtQtyDeliveryMulti)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label6)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.LblUnitDeliveryDefault)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtQtyDeliveryDefault)
        Me.SplitContainer3.Size = New System.Drawing.Size(350, 20)
        Me.SplitContainer3.SplitterDistance = 218
        Me.SplitContainer3.SplitterWidth = 3
        Me.SplitContainer3.TabIndex = 10
        Me.SplitContainer3.TabStop = False
        '
        'LblUnitDeliveryMulti
        '
        Me.LblUnitDeliveryMulti.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnitDeliveryMulti.AutoSize = True
        Me.LblUnitDeliveryMulti.Location = New System.Drawing.Point(140, 2)
        Me.LblUnitDeliveryMulti.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblUnitDeliveryMulti.Name = "LblUnitDeliveryMulti"
        Me.LblUnitDeliveryMulti.Size = New System.Drawing.Size(26, 13)
        Me.LblUnitDeliveryMulti.TabIndex = 2
        Me.LblUnitDeliveryMulti.Text = "Unit"
        '
        'TxtQtyDeliveryMulti
        '
        Me.TxtQtyDeliveryMulti.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyDeliveryMulti.Location = New System.Drawing.Point(80, 0)
        Me.TxtQtyDeliveryMulti.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtQtyDeliveryMulti.Name = "TxtQtyDeliveryMulti"
        Me.TxtQtyDeliveryMulti.ReadOnly = True
        Me.TxtQtyDeliveryMulti.Size = New System.Drawing.Size(57, 20)
        Me.TxtQtyDeliveryMulti.TabIndex = 1
        Me.TxtQtyDeliveryMulti.TabStop = False
        Me.TxtQtyDeliveryMulti.Text = "0"
        Me.TxtQtyDeliveryMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(1, 2)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Delivery Qty"
        '
        'LblUnitDeliveryDefault
        '
        Me.LblUnitDeliveryDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnitDeliveryDefault.AutoSize = True
        Me.LblUnitDeliveryDefault.Location = New System.Drawing.Point(71, 2)
        Me.LblUnitDeliveryDefault.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblUnitDeliveryDefault.Name = "LblUnitDeliveryDefault"
        Me.LblUnitDeliveryDefault.Size = New System.Drawing.Size(26, 13)
        Me.LblUnitDeliveryDefault.TabIndex = 1
        Me.LblUnitDeliveryDefault.Text = "Unit"
        '
        'TxtQtyDeliveryDefault
        '
        Me.TxtQtyDeliveryDefault.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyDeliveryDefault.Location = New System.Drawing.Point(2, 0)
        Me.TxtQtyDeliveryDefault.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtQtyDeliveryDefault.Name = "TxtQtyDeliveryDefault"
        Me.TxtQtyDeliveryDefault.ReadOnly = True
        Me.TxtQtyDeliveryDefault.Size = New System.Drawing.Size(67, 20)
        Me.TxtQtyDeliveryDefault.TabIndex = 0
        Me.TxtQtyDeliveryDefault.TabStop = False
        Me.TxtQtyDeliveryDefault.Text = "0"
        Me.TxtQtyDeliveryDefault.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(3, 0)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtQtyOrderBalanceMulti)
        Me.SplitContainer2.Panel1.Controls.Add(Me.LblUnitOrderBalanceMulti)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label3)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.LblUnitOrderBalanceDefault)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtQtyOrderBalanceDefault)
        Me.SplitContainer2.Size = New System.Drawing.Size(350, 20)
        Me.SplitContainer2.SplitterDistance = 218
        Me.SplitContainer2.SplitterWidth = 3
        Me.SplitContainer2.TabIndex = 0
        Me.SplitContainer2.TabStop = False
        '
        'TxtQtyOrderBalanceMulti
        '
        Me.TxtQtyOrderBalanceMulti.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyOrderBalanceMulti.Location = New System.Drawing.Point(80, 0)
        Me.TxtQtyOrderBalanceMulti.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtQtyOrderBalanceMulti.Name = "TxtQtyOrderBalanceMulti"
        Me.TxtQtyOrderBalanceMulti.ReadOnly = True
        Me.TxtQtyOrderBalanceMulti.Size = New System.Drawing.Size(57, 20)
        Me.TxtQtyOrderBalanceMulti.TabIndex = 1
        Me.TxtQtyOrderBalanceMulti.TabStop = False
        Me.TxtQtyOrderBalanceMulti.Text = "0"
        Me.TxtQtyOrderBalanceMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblUnitOrderBalanceMulti
        '
        Me.LblUnitOrderBalanceMulti.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnitOrderBalanceMulti.AutoSize = True
        Me.LblUnitOrderBalanceMulti.Location = New System.Drawing.Point(140, 2)
        Me.LblUnitOrderBalanceMulti.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblUnitOrderBalanceMulti.Name = "LblUnitOrderBalanceMulti"
        Me.LblUnitOrderBalanceMulti.Size = New System.Drawing.Size(26, 13)
        Me.LblUnitOrderBalanceMulti.TabIndex = 2
        Me.LblUnitOrderBalanceMulti.Text = "Unit"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(1, 2)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Order Balance"
        '
        'LblUnitOrderBalanceDefault
        '
        Me.LblUnitOrderBalanceDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnitOrderBalanceDefault.AutoSize = True
        Me.LblUnitOrderBalanceDefault.Location = New System.Drawing.Point(71, 2)
        Me.LblUnitOrderBalanceDefault.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblUnitOrderBalanceDefault.Name = "LblUnitOrderBalanceDefault"
        Me.LblUnitOrderBalanceDefault.Size = New System.Drawing.Size(26, 13)
        Me.LblUnitOrderBalanceDefault.TabIndex = 1
        Me.LblUnitOrderBalanceDefault.Text = "Unit"
        '
        'TxtQtyOrderBalanceDefault
        '
        Me.TxtQtyOrderBalanceDefault.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyOrderBalanceDefault.Location = New System.Drawing.Point(2, 0)
        Me.TxtQtyOrderBalanceDefault.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtQtyOrderBalanceDefault.Name = "TxtQtyOrderBalanceDefault"
        Me.TxtQtyOrderBalanceDefault.ReadOnly = True
        Me.TxtQtyOrderBalanceDefault.Size = New System.Drawing.Size(67, 20)
        Me.TxtQtyOrderBalanceDefault.TabIndex = 0
        Me.TxtQtyOrderBalanceDefault.TabStop = False
        Me.TxtQtyOrderBalanceDefault.Text = "0"
        Me.TxtQtyOrderBalanceDefault.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'dgvItemDetail
        '
        Me.dgvItemDetail.AllowUserToAddRows = False
        Me.dgvItemDetail.AllowUserToDeleteRows = False
        Me.dgvItemDetail.AllowUserToResizeColumns = False
        Me.dgvItemDetail.AllowUserToResizeRows = False
        Me.dgvItemDetail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemDetail.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItemDetail.ColumnHeadersHeight = 25
        Me.dgvItemDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColWarehouse, Me.ColLocation, Me.ColBatchNo, Me.ColExpiry, Me.ColQtyAvailable, Me.ColQtyDelivery, Me.ColUnit, Me.ColPrimaryLocation})
        Me.dgvItemDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemDetail.Location = New System.Drawing.Point(3, 3)
        Me.dgvItemDetail.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvItemDetail.Name = "dgvItemDetail"
        Me.dgvItemDetail.RowHeadersVisible = False
        Me.dgvItemDetail.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemDetail.RowTemplate.Height = 24
        Me.dgvItemDetail.Size = New System.Drawing.Size(699, 208)
        Me.dgvItemDetail.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColWarehouse
        '
        Me.ColWarehouse.HeaderText = "Warehouse"
        Me.ColWarehouse.Name = "ColWarehouse"
        Me.ColWarehouse.Width = 87
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.Width = 73
        '
        'ColBatchNo
        '
        Me.ColBatchNo.HeaderText = "Batch No"
        Me.ColBatchNo.Name = "ColBatchNo"
        Me.ColBatchNo.Width = 77
        '
        'ColExpiry
        '
        Me.ColExpiry.HeaderText = "Expiry"
        Me.ColExpiry.Name = "ColExpiry"
        Me.ColExpiry.Width = 60
        '
        'ColQtyAvailable
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyAvailable.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColQtyAvailable.HeaderText = "Qty Available"
        Me.ColQtyAvailable.Name = "ColQtyAvailable"
        Me.ColQtyAvailable.Width = 94
        '
        'ColQtyDelivery
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyDelivery.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyDelivery.HeaderText = "Qty Delivery"
        Me.ColQtyDelivery.Name = "ColQtyDelivery"
        Me.ColQtyDelivery.Width = 89
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Width = 32
        '
        'ColPrimaryLocation
        '
        Me.ColPrimaryLocation.HeaderText = "Primary Location"
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryLocation.Visible = False
        Me.ColPrimaryLocation.Width = 91
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(666, 485)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 20)
        Me.btnCancel.TabIndex = 7
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(605, 485)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 20)
        Me.btnOK.TabIndex = 6
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'llbUnselectAll
        '
        Me.llbUnselectAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUnselectAll.AutoSize = True
        Me.llbUnselectAll.Location = New System.Drawing.Point(653, 241)
        Me.llbUnselectAll.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbUnselectAll.Name = "llbUnselectAll"
        Me.llbUnselectAll.Size = New System.Drawing.Size(63, 13)
        Me.llbUnselectAll.TabIndex = 5
        Me.llbUnselectAll.TabStop = True
        Me.llbUnselectAll.Text = "Unselect All"
        '
        'llbSelectAll
        '
        Me.llbSelectAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbSelectAll.AutoSize = True
        Me.llbSelectAll.Location = New System.Drawing.Point(599, 241)
        Me.llbSelectAll.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbSelectAll.Name = "llbSelectAll"
        Me.llbSelectAll.Size = New System.Drawing.Size(51, 13)
        Me.llbSelectAll.TabIndex = 4
        Me.llbSelectAll.TabStop = True
        Me.llbSelectAll.Text = "Select All"
        '
        'llbLocationAccess
        '
        Me.llbLocationAccess.AutoSize = True
        Me.llbLocationAccess.Location = New System.Drawing.Point(11, 219)
        Me.llbLocationAccess.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbLocationAccess.Name = "llbLocationAccess"
        Me.llbLocationAccess.Size = New System.Drawing.Size(86, 13)
        Me.llbLocationAccess.TabIndex = 2
        Me.llbLocationAccess.TabStop = True
        Me.llbLocationAccess.Text = "Location Access"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.SplitContainer5)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 78)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(710, 137)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'SplitContainer5
        '
        Me.SplitContainer5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer5.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer5.Name = "SplitContainer5"
        '
        'SplitContainer5.Panel1
        '
        Me.SplitContainer5.Panel1.Controls.Add(Me.SplitContainer6)
        '
        'SplitContainer5.Panel2
        '
        Me.SplitContainer5.Panel2.Controls.Add(Me.lblCounter)
        Me.SplitContainer5.Size = New System.Drawing.Size(704, 118)
        Me.SplitContainer5.SplitterDistance = 472
        Me.SplitContainer5.TabIndex = 2
        Me.SplitContainer5.TabStop = False
        '
        'SplitContainer6
        '
        Me.SplitContainer6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer6.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer6.Name = "SplitContainer6"
        '
        'SplitContainer6.Panel1
        '
        Me.SplitContainer6.Panel1.Controls.Add(Me.TxtWarehouseSource)
        Me.SplitContainer6.Panel1.Controls.Add(Me.Label12)
        Me.SplitContainer6.Panel1.Controls.Add(Me.TxtLocationSource)
        Me.SplitContainer6.Panel1.Controls.Add(Me.btnFindBarcode)
        Me.SplitContainer6.Panel1.Controls.Add(Me.TxtTotalRows)
        Me.SplitContainer6.Panel1.Controls.Add(Me.Label15)
        Me.SplitContainer6.Panel1.Controls.Add(Me.TxtBarcodeID)
        Me.SplitContainer6.Panel1.Controls.Add(Me.Label13)
        Me.SplitContainer6.Panel1.Controls.Add(Me.BtnFindLocationSource)
        Me.SplitContainer6.Panel1.Controls.Add(Me.Label14)
        Me.SplitContainer6.Panel1.Controls.Add(Me.TxtScanHere)
        Me.SplitContainer6.Panel1.Controls.Add(Me.Label9)
        '
        'SplitContainer6.Panel2
        '
        Me.SplitContainer6.Panel2.Controls.Add(Me.TxtPrimaryWarehouseSource)
        Me.SplitContainer6.Panel2.Controls.Add(Me.TxtQtyMulti)
        Me.SplitContainer6.Panel2.Controls.Add(Me.TxtPrimaryLocationSource)
        Me.SplitContainer6.Panel2.Controls.Add(Me.BtnInputBarcode)
        Me.SplitContainer6.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer6.Panel2.Controls.Add(Me.Label11)
        Me.SplitContainer6.Panel2.Controls.Add(Me.TxtQtyDefault)
        Me.SplitContainer6.Panel2.Controls.Add(Me.TxtWarehouseSourceDesc)
        Me.SplitContainer6.Panel2.Controls.Add(Me.TxtLocationSourceDesc)
        Me.SplitContainer6.Panel2.Controls.Add(Me.Label8)
        Me.SplitContainer6.Panel2.Controls.Add(Me.TxtBarcodeLabel)
        Me.SplitContainer6.Size = New System.Drawing.Size(472, 118)
        Me.SplitContainer6.SplitterDistance = 269
        Me.SplitContainer6.TabIndex = 0
        Me.SplitContainer6.TabStop = False
        '
        'TxtWarehouseSource
        '
        Me.TxtWarehouseSource.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtWarehouseSource.Location = New System.Drawing.Point(99, 25)
        Me.TxtWarehouseSource.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtWarehouseSource.Name = "TxtWarehouseSource"
        Me.TxtWarehouseSource.ReadOnly = True
        Me.TxtWarehouseSource.Size = New System.Drawing.Size(163, 20)
        Me.TxtWarehouseSource.TabIndex = 7
        Me.TxtWarehouseSource.TabStop = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(4, 51)
        Me.Label12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(61, 13)
        Me.Label12.TabIndex = 2
        Me.Label12.Text = "Total Rows"
        '
        'TxtLocationSource
        '
        Me.TxtLocationSource.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLocationSource.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtLocationSource.Location = New System.Drawing.Point(99, 3)
        Me.TxtLocationSource.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtLocationSource.Name = "TxtLocationSource"
        Me.TxtLocationSource.Size = New System.Drawing.Size(102, 20)
        Me.TxtLocationSource.TabIndex = 5
        '
        'btnFindBarcode
        '
        Me.btnFindBarcode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindBarcode.Location = New System.Drawing.Point(204, 96)
        Me.btnFindBarcode.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFindBarcode.Name = "btnFindBarcode"
        Me.btnFindBarcode.Size = New System.Drawing.Size(58, 20)
        Me.btnFindBarcode.TabIndex = 11
        Me.btnFindBarcode.TabStop = False
        Me.btnFindBarcode.Text = "Find"
        Me.btnFindBarcode.UseVisualStyleBackColor = True
        '
        'TxtTotalRows
        '
        Me.TxtTotalRows.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalRows.Location = New System.Drawing.Point(99, 49)
        Me.TxtTotalRows.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtTotalRows.Name = "TxtTotalRows"
        Me.TxtTotalRows.ReadOnly = True
        Me.TxtTotalRows.Size = New System.Drawing.Size(163, 20)
        Me.TxtTotalRows.TabIndex = 8
        Me.TxtTotalRows.TabStop = False
        Me.TxtTotalRows.Text = "0"
        Me.TxtTotalRows.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(4, 5)
        Me.Label15.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(48, 13)
        Me.Label15.TabIndex = 0
        Me.Label15.Text = "Location"
        '
        'TxtBarcodeID
        '
        Me.TxtBarcodeID.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBarcodeID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBarcodeID.Location = New System.Drawing.Point(99, 96)
        Me.TxtBarcodeID.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtBarcodeID.Name = "TxtBarcodeID"
        Me.TxtBarcodeID.Size = New System.Drawing.Size(102, 20)
        Me.TxtBarcodeID.TabIndex = 10
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(4, 98)
        Me.Label13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(47, 13)
        Me.Label13.TabIndex = 4
        Me.Label13.Text = "Barcode"
        '
        'BtnFindLocationSource
        '
        Me.BtnFindLocationSource.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindLocationSource.Location = New System.Drawing.Point(204, 3)
        Me.BtnFindLocationSource.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnFindLocationSource.Name = "BtnFindLocationSource"
        Me.BtnFindLocationSource.Size = New System.Drawing.Size(56, 20)
        Me.BtnFindLocationSource.TabIndex = 6
        Me.BtnFindLocationSource.TabStop = False
        Me.BtnFindLocationSource.Text = "Find"
        Me.BtnFindLocationSource.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(4, 28)
        Me.Label14.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(62, 13)
        Me.Label14.TabIndex = 1
        Me.Label14.Text = "Warehouse"
        '
        'TxtScanHere
        '
        Me.TxtScanHere.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtScanHere.Location = New System.Drawing.Point(99, 72)
        Me.TxtScanHere.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtScanHere.Name = "TxtScanHere"
        Me.TxtScanHere.Size = New System.Drawing.Size(163, 20)
        Me.TxtScanHere.TabIndex = 9
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(4, 74)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(58, 13)
        Me.Label9.TabIndex = 3
        Me.Label9.Text = "Scan Here"
        '
        'TxtPrimaryWarehouseSource
        '
        Me.TxtPrimaryWarehouseSource.Location = New System.Drawing.Point(8, 25)
        Me.TxtPrimaryWarehouseSource.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtPrimaryWarehouseSource.Name = "TxtPrimaryWarehouseSource"
        Me.TxtPrimaryWarehouseSource.Size = New System.Drawing.Size(32, 20)
        Me.TxtPrimaryWarehouseSource.TabIndex = 2
        Me.TxtPrimaryWarehouseSource.TabStop = False
        Me.TxtPrimaryWarehouseSource.Text = "-1"
        Me.TxtPrimaryWarehouseSource.Visible = False
        '
        'TxtQtyMulti
        '
        Me.TxtQtyMulti.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyMulti.Location = New System.Drawing.Point(93, 48)
        Me.TxtQtyMulti.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtQtyMulti.Name = "TxtQtyMulti"
        Me.TxtQtyMulti.ReadOnly = True
        Me.TxtQtyMulti.Size = New System.Drawing.Size(102, 20)
        Me.TxtQtyMulti.TabIndex = 7
        Me.TxtQtyMulti.TabStop = False
        Me.TxtQtyMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtPrimaryLocationSource
        '
        Me.TxtPrimaryLocationSource.Location = New System.Drawing.Point(8, 3)
        Me.TxtPrimaryLocationSource.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtPrimaryLocationSource.Name = "TxtPrimaryLocationSource"
        Me.TxtPrimaryLocationSource.Size = New System.Drawing.Size(32, 20)
        Me.TxtPrimaryLocationSource.TabIndex = 0
        Me.TxtPrimaryLocationSource.TabStop = False
        Me.TxtPrimaryLocationSource.Text = "-1"
        Me.TxtPrimaryLocationSource.Visible = False
        '
        'BtnInputBarcode
        '
        Me.BtnInputBarcode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnInputBarcode.Location = New System.Drawing.Point(151, 96)
        Me.BtnInputBarcode.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnInputBarcode.Name = "BtnInputBarcode"
        Me.BtnInputBarcode.Size = New System.Drawing.Size(43, 20)
        Me.BtnInputBarcode.TabIndex = 10
        Me.BtnInputBarcode.Text = "Input"
        Me.BtnInputBarcode.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 73)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(87, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Total Qty Default"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 50)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(75, 13)
        Me.Label11.TabIndex = 4
        Me.Label11.Text = "Total Qty Multi"
        '
        'TxtQtyDefault
        '
        Me.TxtQtyDefault.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyDefault.Location = New System.Drawing.Point(93, 71)
        Me.TxtQtyDefault.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtQtyDefault.Name = "TxtQtyDefault"
        Me.TxtQtyDefault.ReadOnly = True
        Me.TxtQtyDefault.Size = New System.Drawing.Size(102, 20)
        Me.TxtQtyDefault.TabIndex = 8
        Me.TxtQtyDefault.TabStop = False
        Me.TxtQtyDefault.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtWarehouseSourceDesc
        '
        Me.TxtWarehouseSourceDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtWarehouseSourceDesc.Location = New System.Drawing.Point(3, 24)
        Me.TxtWarehouseSourceDesc.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtWarehouseSourceDesc.Name = "TxtWarehouseSourceDesc"
        Me.TxtWarehouseSourceDesc.ReadOnly = True
        Me.TxtWarehouseSourceDesc.Size = New System.Drawing.Size(193, 20)
        Me.TxtWarehouseSourceDesc.TabIndex = 3
        Me.TxtWarehouseSourceDesc.TabStop = False
        '
        'TxtLocationSourceDesc
        '
        Me.TxtLocationSourceDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLocationSourceDesc.Location = New System.Drawing.Point(3, 2)
        Me.TxtLocationSourceDesc.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtLocationSourceDesc.Name = "TxtLocationSourceDesc"
        Me.TxtLocationSourceDesc.ReadOnly = True
        Me.TxtLocationSourceDesc.Size = New System.Drawing.Size(193, 20)
        Me.TxtLocationSourceDesc.TabIndex = 1
        Me.TxtLocationSourceDesc.TabStop = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 98)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(47, 13)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = "Label ID"
        '
        'TxtBarcodeLabel
        '
        Me.TxtBarcodeLabel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBarcodeLabel.Location = New System.Drawing.Point(93, 96)
        Me.TxtBarcodeLabel.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtBarcodeLabel.MaxLength = 2
        Me.TxtBarcodeLabel.Name = "TxtBarcodeLabel"
        Me.TxtBarcodeLabel.Size = New System.Drawing.Size(54, 20)
        Me.TxtBarcodeLabel.TabIndex = 9
        Me.TxtBarcodeLabel.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblCounter
        '
        Me.lblCounter.AutoSize = True
        Me.lblCounter.Font = New System.Drawing.Font("Microsoft Sans Serif", 75.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCounter.ForeColor = System.Drawing.Color.Blue
        Me.lblCounter.Location = New System.Drawing.Point(2, 4)
        Me.lblCounter.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblCounter.Name = "lblCounter"
        Me.lblCounter.Size = New System.Drawing.Size(104, 113)
        Me.lblCounter.TabIndex = 0
        Me.lblCounter.Text = "0"
        Me.lblCounter.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(9, 240)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(713, 240)
        Me.TabControl1.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvItemDetail)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(705, 214)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "DO Sales"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvItem)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(705, 214)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Barcode Data"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1No, Me.Col1BarcodeType, Me.Col1Barcode, Me.Col1LabelID, Me.Col1CodeNo, Me.Col1Description, Me.Col1Qty, Me.Col1Unit, Me.Col1Ratio, Me.Col1RatioMulti, Me.Col1RatioDefault, Me.Col1QtyDefault, Me.Col1UnitDefault, Me.Col1Batch, Me.Col1ExpiredDate, Me.Col1LocationSource, Me.Col1WareHouseSource, Me.Col1Remark, Me.Col1PrimaryCodeNo, Me.Col1PrimaryUnits, Me.Col1PrimaryUnitDefault, Me.Col1LocationSourcePrimary, Me.Col1WarehouseSourcePrimary, Me.Col1Decimal, Me.Col1DecimalDefault, Me.Col1LocationControl, Me.Col1IsService})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(3, 3)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(699, 208)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 29
        Me.dgvItem.TabStop = False
        '
        'Col1No
        '
        Me.Col1No.Frozen = True
        Me.Col1No.HeaderText = "No"
        Me.Col1No.Name = "Col1No"
        Me.Col1No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col1No.Width = 27
        '
        'Col1BarcodeType
        '
        Me.Col1BarcodeType.Frozen = True
        Me.Col1BarcodeType.HeaderText = "Barcode Type"
        Me.Col1BarcodeType.Name = "Col1BarcodeType"
        Me.Col1BarcodeType.ReadOnly = True
        Me.Col1BarcodeType.Width = 99
        '
        'Col1Barcode
        '
        Me.Col1Barcode.Frozen = True
        Me.Col1Barcode.HeaderText = "Barcode"
        Me.Col1Barcode.Name = "Col1Barcode"
        Me.Col1Barcode.Width = 72
        '
        'Col1LabelID
        '
        Me.Col1LabelID.Frozen = True
        Me.Col1LabelID.HeaderText = "Label ID"
        Me.Col1LabelID.Name = "Col1LabelID"
        Me.Col1LabelID.Width = 72
        '
        'Col1CodeNo
        '
        Me.Col1CodeNo.Frozen = True
        Me.Col1CodeNo.HeaderText = "Code No"
        Me.Col1CodeNo.Name = "Col1CodeNo"
        Me.Col1CodeNo.Width = 74
        '
        'Col1Description
        '
        Me.Col1Description.HeaderText = "Description"
        Me.Col1Description.Name = "Col1Description"
        Me.Col1Description.Width = 85
        '
        'Col1Qty
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col1Qty.DefaultCellStyle = DataGridViewCellStyle4
        Me.Col1Qty.HeaderText = "Qty"
        Me.Col1Qty.Name = "Col1Qty"
        Me.Col1Qty.Width = 48
        '
        'Col1Unit
        '
        Me.Col1Unit.HeaderText = "Units"
        Me.Col1Unit.Name = "Col1Unit"
        Me.Col1Unit.Width = 56
        '
        'Col1Ratio
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col1Ratio.DefaultCellStyle = DataGridViewCellStyle5
        Me.Col1Ratio.HeaderText = "Ratio"
        Me.Col1Ratio.Name = "Col1Ratio"
        Me.Col1Ratio.Width = 57
        '
        'Col1RatioMulti
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col1RatioMulti.DefaultCellStyle = DataGridViewCellStyle6
        Me.Col1RatioMulti.HeaderText = "Ratio Multi"
        Me.Col1RatioMulti.Name = "Col1RatioMulti"
        Me.Col1RatioMulti.Visible = False
        Me.Col1RatioMulti.Width = 82
        '
        'Col1RatioDefault
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col1RatioDefault.DefaultCellStyle = DataGridViewCellStyle7
        Me.Col1RatioDefault.HeaderText = "Ratio Default"
        Me.Col1RatioDefault.Name = "Col1RatioDefault"
        Me.Col1RatioDefault.Visible = False
        Me.Col1RatioDefault.Width = 94
        '
        'Col1QtyDefault
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col1QtyDefault.DefaultCellStyle = DataGridViewCellStyle8
        Me.Col1QtyDefault.HeaderText = "Qty Default"
        Me.Col1QtyDefault.Name = "Col1QtyDefault"
        Me.Col1QtyDefault.Width = 85
        '
        'Col1UnitDefault
        '
        Me.Col1UnitDefault.HeaderText = "Unit Default"
        Me.Col1UnitDefault.Name = "Col1UnitDefault"
        Me.Col1UnitDefault.Width = 88
        '
        'Col1Batch
        '
        Me.Col1Batch.HeaderText = "Batch No"
        Me.Col1Batch.Name = "Col1Batch"
        Me.Col1Batch.Width = 77
        '
        'Col1ExpiredDate
        '
        Me.Col1ExpiredDate.HeaderText = "Expiry Date"
        Me.Col1ExpiredDate.Name = "Col1ExpiredDate"
        Me.Col1ExpiredDate.Width = 86
        '
        'Col1LocationSource
        '
        Me.Col1LocationSource.HeaderText = "Location"
        Me.Col1LocationSource.Name = "Col1LocationSource"
        Me.Col1LocationSource.Width = 73
        '
        'Col1WareHouseSource
        '
        Me.Col1WareHouseSource.HeaderText = "WareHouse"
        Me.Col1WareHouseSource.Name = "Col1WareHouseSource"
        Me.Col1WareHouseSource.Width = 89
        '
        'Col1Remark
        '
        Me.Col1Remark.HeaderText = "Remarks"
        Me.Col1Remark.Name = "Col1Remark"
        Me.Col1Remark.Visible = False
        Me.Col1Remark.Width = 74
        '
        'Col1PrimaryCodeNo
        '
        Me.Col1PrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.Col1PrimaryCodeNo.Name = "Col1PrimaryCodeNo"
        Me.Col1PrimaryCodeNo.Visible = False
        Me.Col1PrimaryCodeNo.Width = 105
        '
        'Col1PrimaryUnits
        '
        Me.Col1PrimaryUnits.HeaderText = "Primary Units"
        Me.Col1PrimaryUnits.Name = "Col1PrimaryUnits"
        Me.Col1PrimaryUnits.Visible = False
        Me.Col1PrimaryUnits.Width = 93
        '
        'Col1PrimaryUnitDefault
        '
        Me.Col1PrimaryUnitDefault.HeaderText = "Primary Unit Default"
        Me.Col1PrimaryUnitDefault.Name = "Col1PrimaryUnitDefault"
        Me.Col1PrimaryUnitDefault.Visible = False
        Me.Col1PrimaryUnitDefault.Width = 125
        '
        'Col1LocationSourcePrimary
        '
        Me.Col1LocationSourcePrimary.HeaderText = "Primary Location Source"
        Me.Col1LocationSourcePrimary.Name = "Col1LocationSourcePrimary"
        Me.Col1LocationSourcePrimary.ReadOnly = True
        Me.Col1LocationSourcePrimary.Visible = False
        Me.Col1LocationSourcePrimary.Width = 147
        '
        'Col1WarehouseSourcePrimary
        '
        Me.Col1WarehouseSourcePrimary.HeaderText = "Primary Warehouse Source"
        Me.Col1WarehouseSourcePrimary.Name = "Col1WarehouseSourcePrimary"
        Me.Col1WarehouseSourcePrimary.ReadOnly = True
        Me.Col1WarehouseSourcePrimary.Visible = False
        Me.Col1WarehouseSourcePrimary.Width = 161
        '
        'Col1Decimal
        '
        Me.Col1Decimal.HeaderText = "Decimal"
        Me.Col1Decimal.Name = "Col1Decimal"
        Me.Col1Decimal.ReadOnly = True
        Me.Col1Decimal.Visible = False
        Me.Col1Decimal.Width = 70
        '
        'Col1DecimalDefault
        '
        Me.Col1DecimalDefault.HeaderText = "Dec Default"
        Me.Col1DecimalDefault.Name = "Col1DecimalDefault"
        Me.Col1DecimalDefault.ReadOnly = True
        Me.Col1DecimalDefault.Visible = False
        Me.Col1DecimalDefault.Width = 89
        '
        'Col1LocationControl
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col1LocationControl.DefaultCellStyle = DataGridViewCellStyle9
        Me.Col1LocationControl.FillWeight = 30.0!
        Me.Col1LocationControl.HeaderText = "Is Location Controlled"
        Me.Col1LocationControl.MaxInputLength = 3
        Me.Col1LocationControl.MinimumWidth = 30
        Me.Col1LocationControl.Name = "Col1LocationControl"
        Me.Col1LocationControl.ReadOnly = True
        Me.Col1LocationControl.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col1LocationControl.Visible = False
        Me.Col1LocationControl.Width = 115
        '
        'Col1IsService
        '
        Me.Col1IsService.HeaderText = "Is Service"
        Me.Col1IsService.Name = "Col1IsService"
        Me.Col1IsService.Visible = False
        Me.Col1IsService.Width = 79
        '
        'FrmTransaksiDO3_PickQty
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(733, 514)
        Me.Controls.Add(Me.llbUnselectAll)
        Me.Controls.Add(Me.llbSelectAll)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.llbLocationAccess)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(749, 552)
        Me.Name = "FrmTransaksiDO3_PickQty"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Pick By Location"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.Panel2.PerformLayout()
        Me.SplitContainer4.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.PerformLayout()
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.Panel2.PerformLayout()
        Me.SplitContainer3.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.dgvItemDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.SplitContainer5.Panel1.ResumeLayout(False)
        Me.SplitContainer5.Panel2.ResumeLayout(False)
        Me.SplitContainer5.Panel2.PerformLayout()
        Me.SplitContainer5.ResumeLayout(False)
        Me.SplitContainer6.Panel1.ResumeLayout(False)
        Me.SplitContainer6.Panel1.PerformLayout()
        Me.SplitContainer6.Panel2.ResumeLayout(False)
        Me.SplitContainer6.Panel2.PerformLayout()
        Me.SplitContainer6.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtSONumber As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents LblUnitDeliveryMulti As System.Windows.Forms.Label
    Friend WithEvents TxtQtyDeliveryMulti As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents LblUnitDeliveryDefault As System.Windows.Forms.Label
    Friend WithEvents TxtQtyDeliveryDefault As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtQtyOrderBalanceMulti As System.Windows.Forms.TextBox
    Friend WithEvents LblUnitOrderBalanceMulti As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LblUnitOrderBalanceDefault As System.Windows.Forms.Label
    Friend WithEvents TxtQtyOrderBalanceDefault As System.Windows.Forms.TextBox
    Friend WithEvents LblRatio As System.Windows.Forms.Label
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents LblUnitAvailableMulti As System.Windows.Forms.Label
    Friend WithEvents TxtQtyAvailableMulti As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents LblUnitAvailableDefault As System.Windows.Forms.Label
    Friend WithEvents TxtQtyAvailableDefault As System.Windows.Forms.TextBox
    Friend WithEvents dgvItemDetail As System.Windows.Forms.DataGridView
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents llbUnselectAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbSelectAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbLocationAccess As System.Windows.Forms.LinkLabel
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyAvailable As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyDelivery As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents SplitContainer5 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtPrimaryWarehouseSource As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryLocationSource As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtQtyDefault As System.Windows.Forms.TextBox
    Friend WithEvents TxtScanHere As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtBarcodeLabel As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TxtQtyMulti As System.Windows.Forms.TextBox
    Friend WithEvents TxtTotalRows As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents BtnInputBarcode As System.Windows.Forms.Button
    Friend WithEvents btnFindBarcode As System.Windows.Forms.Button
    Friend WithEvents TxtBarcodeID As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TxtWarehouseSourceDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtWarehouseSource As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TxtLocationSourceDesc As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindLocationSource As System.Windows.Forms.Button
    Friend WithEvents TxtLocationSource As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents lblCounter As System.Windows.Forms.Label
    Friend WithEvents SplitContainer6 As System.Windows.Forms.SplitContainer
    Friend WithEvents Col1No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1BarcodeType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1Barcode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1LabelID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1Ratio As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1RatioMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1RatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1QtyDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1UnitDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1Batch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1ExpiredDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1LocationSource As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1WareHouseSource As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1Remark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1PrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1PrimaryUnitDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1LocationSourcePrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1WarehouseSourcePrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1Decimal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1DecimalDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1LocationControl As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1IsService As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
