<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiMonthlyClosing
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryClosing = New System.Windows.Forms.TextBox
        Me.TxtClosingNumber = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.btnCekData = New System.Windows.Forms.Button
        Me.btnFindRow = New System.Windows.Forms.Button
        Me.dtpPeriod = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBeginQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBeginAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOutQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOutAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEndQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEndAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetailCalculation = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvItemAdjust = New System.Windows.Forms.DataGridView
        Me.ColNoAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJenisDocumentAdjustStr = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocumentNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNoAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPriceAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarksAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryStockInAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNoAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJenisDokumen = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvItemAdjust, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(111, 42)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(319, 20)
        Me.dtpDate.TabIndex = 2
        '
        'TxtPrimaryClosing
        '
        Me.TxtPrimaryClosing.Location = New System.Drawing.Point(151, 19)
        Me.TxtPrimaryClosing.Name = "TxtPrimaryClosing"
        Me.TxtPrimaryClosing.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryClosing.TabIndex = 1
        Me.TxtPrimaryClosing.TabStop = False
        Me.TxtPrimaryClosing.Visible = False
        '
        'TxtClosingNumber
        '
        Me.TxtClosingNumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtClosingNumber.Location = New System.Drawing.Point(111, 19)
        Me.TxtClosingNumber.MaxLength = 50
        Me.TxtClosingNumber.Name = "TxtClosingNumber"
        Me.TxtClosingNumber.ReadOnly = True
        Me.TxtClosingNumber.Size = New System.Drawing.Size(319, 20)
        Me.TxtClosingNumber.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Closing Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Transaction No"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(219, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 631)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(894, 22)
        Me.StatusStrip1.TabIndex = 39
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(219, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(219, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(219, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(449, 561)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(333, 561)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 7
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'ToolTip
        '
        Me.ToolTip.AutoPopDelay = 10000
        Me.ToolTip.InitialDelay = 500
        Me.ToolTip.ReshowDelay = 100
        Me.ToolTip.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.ToolTip.ToolTipTitle = "Info"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.Button2)
        Me.gboDetail.Controls.Add(Me.btnCekData)
        Me.gboDetail.Controls.Add(Me.btnFindRow)
        Me.gboDetail.Controls.Add(Me.dtpPeriod)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.TxtRemark)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.dtpDate)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryClosing)
        Me.gboDetail.Controls.Add(Me.TxtClosingNumber)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.Button1)
        Me.gboDetail.Controls.Add(Me.TabControl1)
        Me.gboDetail.Location = New System.Drawing.Point(4, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(887, 593)
        Me.gboDetail.TabIndex = 6
        Me.gboDetail.TabStop = False
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(559, 561)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(79, 26)
        Me.Button2.TabIndex = 151
        Me.Button2.TabStop = False
        Me.Button2.Text = "Add Adjustment"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'btnCekData
        '
        Me.btnCekData.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCekData.Location = New System.Drawing.Point(96, 561)
        Me.btnCekData.Name = "btnCekData"
        Me.btnCekData.Size = New System.Drawing.Size(79, 26)
        Me.btnCekData.TabIndex = 149
        Me.btnCekData.Text = "Check Data"
        Me.btnCekData.UseVisualStyleBackColor = True
        '
        'btnFindRow
        '
        Me.btnFindRow.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnFindRow.Location = New System.Drawing.Point(6, 561)
        Me.btnFindRow.Name = "btnFindRow"
        Me.btnFindRow.Size = New System.Drawing.Size(84, 26)
        Me.btnFindRow.TabIndex = 148
        Me.btnFindRow.Text = "Find Code No"
        Me.btnFindRow.UseVisualStyleBackColor = True
        '
        'dtpPeriod
        '
        Me.dtpPeriod.CustomFormat = "MMM yyyy"
        Me.dtpPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPeriod.Location = New System.Drawing.Point(111, 68)
        Me.dtpPeriod.Name = "dtpPeriod"
        Me.dtpPeriod.ShowUpDown = True
        Me.dtpPeriod.Size = New System.Drawing.Size(319, 20)
        Me.dtpPeriod.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 74)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 146
        Me.Label1.Text = "Period"
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(507, 21)
        Me.TxtRemark.MaxLength = 150
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(368, 66)
        Me.TxtRemark.TabIndex = 5
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(452, 24)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 143
        Me.Label6.Text = "Remarks"
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(181, 561)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(79, 26)
        Me.Button1.TabIndex = 150
        Me.Button1.TabStop = False
        Me.Button1.Text = "Check Data In"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(6, 94)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(875, 461)
        Me.TabControl1.TabIndex = 147
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvItem)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(867, 435)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Closing Item"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDescription, Me.ColUnits, Me.ColBeginQty, Me.ColBeginAmount, Me.ColInQty, Me.ColInAmount, Me.ColOutQty, Me.ColOutAmount, Me.ColEndQty, Me.ColEndAmount, Me.ColRemarks, Me.ColDetailCalculation, Me.ColPrimaryUnits, Me.ColPrimaryCodeNo})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(3, 3)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.Size = New System.Drawing.Size(861, 429)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 6
        Me.dgvItem.TabStop = False
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColCodeNo
        '
        Me.ColCodeNo.Frozen = True
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCodeNo.Width = 50
        '
        'ColDescription
        '
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDescription.Width = 66
        '
        'ColUnits
        '
        Me.ColUnits.HeaderText = "Unit"
        Me.ColUnits.Name = "ColUnits"
        Me.ColUnits.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnits.Width = 32
        '
        'ColBeginQty
        '
        Me.ColBeginQty.HeaderText = "Begining Qty"
        Me.ColBeginQty.Name = "ColBeginQty"
        Me.ColBeginQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColBeginAmount
        '
        Me.ColBeginAmount.HeaderText = "Begining Amount"
        Me.ColBeginAmount.Name = "ColBeginAmount"
        Me.ColBeginAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColInQty
        '
        Me.ColInQty.HeaderText = "In Qty"
        Me.ColInQty.Name = "ColInQty"
        Me.ColInQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColInQty.Width = 50
        '
        'ColInAmount
        '
        Me.ColInAmount.HeaderText = "In Amount"
        Me.ColInAmount.Name = "ColInAmount"
        Me.ColInAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColInAmount.Width = 80
        '
        'ColOutQty
        '
        Me.ColOutQty.HeaderText = "Out Qty"
        Me.ColOutQty.Name = "ColOutQty"
        Me.ColOutQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColOutQty.Width = 50
        '
        'ColOutAmount
        '
        Me.ColOutAmount.HeaderText = "Out Amount"
        Me.ColOutAmount.Name = "ColOutAmount"
        Me.ColOutAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColOutAmount.Width = 80
        '
        'ColEndQty
        '
        Me.ColEndQty.HeaderText = "End Qty"
        Me.ColEndQty.Name = "ColEndQty"
        Me.ColEndQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEndQty.Width = 70
        '
        'ColEndAmount
        '
        Me.ColEndAmount.HeaderText = "End Amount"
        Me.ColEndAmount.Name = "ColEndAmount"
        Me.ColEndAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 55
        '
        'ColDetailCalculation
        '
        Me.ColDetailCalculation.HeaderText = "Detail Calculation"
        Me.ColDetailCalculation.Name = "ColDetailCalculation"
        Me.ColDetailCalculation.Width = 86
        '
        'ColPrimaryUnits
        '
        Me.ColPrimaryUnits.HeaderText = "Primary Units"
        Me.ColPrimaryUnits.Name = "ColPrimaryUnits"
        Me.ColPrimaryUnits.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryUnits.Visible = False
        Me.ColPrimaryUnits.Width = 67
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 70
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvItemAdjust)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(867, 435)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "In Adjustment Item"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvItemAdjust
        '
        Me.dgvItemAdjust.AllowUserToAddRows = False
        Me.dgvItemAdjust.AllowUserToDeleteRows = False
        Me.dgvItemAdjust.AllowUserToOrderColumns = True
        Me.dgvItemAdjust.AllowUserToResizeColumns = False
        Me.dgvItemAdjust.AllowUserToResizeRows = False
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItemAdjust.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvItemAdjust.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemAdjust.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNoAdjust, Me.ColJenisDocumentAdjustStr, Me.ColDocumentNo, Me.ColCodeNoAdjust, Me.ColDescAdjust, Me.ColQty, Me.ColUnitAdjust, Me.ColUnitPriceAdjust, Me.ColRemarksAdjust, Me.ColPrimaryStockInAdjust, Me.ColPrimaryCodeNoAdjust, Me.ColPrimaryUnitAdjust, Me.ColJenisDokumen})
        Me.dgvItemAdjust.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemAdjust.Location = New System.Drawing.Point(3, 3)
        Me.dgvItemAdjust.Name = "dgvItemAdjust"
        Me.dgvItemAdjust.Size = New System.Drawing.Size(861, 429)
        Me.dgvItemAdjust.TabIndex = 0
        '
        'ColNoAdjust
        '
        Me.ColNoAdjust.HeaderText = "No"
        Me.ColNoAdjust.Name = "ColNoAdjust"
        Me.ColNoAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNoAdjust.Width = 27
        '
        'ColJenisDocumentAdjustStr
        '
        Me.ColJenisDocumentAdjustStr.HeaderText = "Document"
        Me.ColJenisDocumentAdjustStr.Name = "ColJenisDocumentAdjustStr"
        Me.ColJenisDocumentAdjustStr.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJenisDocumentAdjustStr.Width = 62
        '
        'ColDocumentNo
        '
        Me.ColDocumentNo.HeaderText = "Document No"
        Me.ColDocumentNo.Name = "ColDocumentNo"
        Me.ColDocumentNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDocumentNo.Width = 79
        '
        'ColCodeNoAdjust
        '
        Me.ColCodeNoAdjust.HeaderText = "Code No"
        Me.ColCodeNoAdjust.Name = "ColCodeNoAdjust"
        Me.ColCodeNoAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCodeNoAdjust.Width = 55
        '
        'ColDescAdjust
        '
        Me.ColDescAdjust.HeaderText = "Description"
        Me.ColDescAdjust.Name = "ColDescAdjust"
        Me.ColDescAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDescAdjust.Width = 66
        '
        'ColQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColQty.HeaderText = "Qty In Adjust"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Width = 73
        '
        'ColUnitAdjust
        '
        Me.ColUnitAdjust.HeaderText = "Units"
        Me.ColUnitAdjust.Name = "ColUnitAdjust"
        Me.ColUnitAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitAdjust.Width = 37
        '
        'ColUnitPriceAdjust
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPriceAdjust.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColUnitPriceAdjust.HeaderText = "Unit Price (IDR)"
        Me.ColUnitPriceAdjust.Name = "ColUnitPriceAdjust"
        Me.ColUnitPriceAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitPriceAdjust.Width = 87
        '
        'ColRemarksAdjust
        '
        Me.ColRemarksAdjust.HeaderText = "Remarks"
        Me.ColRemarksAdjust.Name = "ColRemarksAdjust"
        Me.ColRemarksAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarksAdjust.Width = 55
        '
        'ColPrimaryStockInAdjust
        '
        Me.ColPrimaryStockInAdjust.HeaderText = "Primary Stock In"
        Me.ColPrimaryStockInAdjust.Name = "ColPrimaryStockInAdjust"
        Me.ColPrimaryStockInAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryStockInAdjust.Visible = False
        Me.ColPrimaryStockInAdjust.Width = 73
        '
        'ColPrimaryCodeNoAdjust
        '
        Me.ColPrimaryCodeNoAdjust.HeaderText = "Primary Code No"
        Me.ColPrimaryCodeNoAdjust.Name = "ColPrimaryCodeNoAdjust"
        Me.ColPrimaryCodeNoAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryCodeNoAdjust.Visible = False
        Me.ColPrimaryCodeNoAdjust.Width = 70
        '
        'ColPrimaryUnitAdjust
        '
        Me.ColPrimaryUnitAdjust.HeaderText = "Primary Unit"
        Me.ColPrimaryUnitAdjust.Name = "ColPrimaryUnitAdjust"
        Me.ColPrimaryUnitAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryUnitAdjust.Visible = False
        Me.ColPrimaryUnitAdjust.Width = 62
        '
        'ColJenisDokumen
        '
        Me.ColJenisDokumen.HeaderText = "Jenis Dokumen"
        Me.ColJenisDokumen.Name = "ColJenisDokumen"
        Me.ColJenisDokumen.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJenisDokumen.Visible = False
        Me.ColJenisDokumen.Width = 78
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(225, 599)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(316, 599)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(589, 599)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(407, 599)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(498, 599)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'FrmTransaksiMonthlyClosing
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(894, 653)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnDelete)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MinimumSize = New System.Drawing.Size(900, 678)
        Me.Name = "FrmTransaksiMonthlyClosing"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase - Monthly Closing"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvItemAdjust, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimaryClosing As System.Windows.Forms.TextBox
    Friend WithEvents TxtClosingNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents dtpPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvItemAdjust As System.Windows.Forms.DataGridView
    Friend WithEvents btnFindRow As System.Windows.Forms.Button
    Friend WithEvents btnCekData As System.Windows.Forms.Button
    Friend WithEvents ColNoAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJenisDocumentAdjustStr As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocumentNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNoAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPriceAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarksAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryStockInAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNoAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJenisDokumen As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBeginQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBeginAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOutQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOutAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEndQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEndAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetailCalculation As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
