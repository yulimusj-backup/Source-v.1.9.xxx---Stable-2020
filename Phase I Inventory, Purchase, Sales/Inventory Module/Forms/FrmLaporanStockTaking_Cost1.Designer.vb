<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanStockTaking_Cost1
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
        Me.gboBrand = New System.Windows.Forms.GroupBox
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.clbBrand = New System.Windows.Forms.CheckedListBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.Crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.LblTotalAmount = New System.Windows.Forms.Label
        Me.chk_Diff_Only = New System.Windows.Forms.CheckBox
        Me.chkCounterMeasure1 = New System.Windows.Forms.CheckBox
        Me.chkAnalisis1 = New System.Windows.Forms.CheckBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.Col_no = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyOnHand = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyTaking = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDifference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAnalisis = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCounterMeasure = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEdit = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColPrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.TxtPrimaryStockTaking = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.LblFormat = New System.Windows.Forms.Label
        Me.cboFormat = New System.Windows.Forms.ComboBox
        Me.chkCounterMeasure = New System.Windows.Forms.CheckBox
        Me.chkAnalysis = New System.Windows.Forms.CheckBox
        Me.btnFindStockTaking = New System.Windows.Forms.Button
        Me.TxtStockTaking = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpStockTakingDate = New System.Windows.Forms.DateTimePicker
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.gboBrand.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboBrand
        '
        Me.gboBrand.Controls.Add(Me.LinkLabel2)
        Me.gboBrand.Controls.Add(Me.LinkLabel1)
        Me.gboBrand.Controls.Add(Me.clbBrand)
        Me.gboBrand.Location = New System.Drawing.Point(9, 10)
        Me.gboBrand.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.gboBrand.Name = "gboBrand"
        Me.gboBrand.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.gboBrand.Size = New System.Drawing.Size(361, 127)
        Me.gboBrand.TabIndex = 0
        Me.gboBrand.TabStop = False
        Me.gboBrand.Text = "Brand"
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(272, 6)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel2.TabIndex = 5
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Uncheck All"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(215, 6)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel1.TabIndex = 4
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Check All"
        '
        'clbBrand
        '
        Me.clbBrand.CheckOnClick = True
        Me.clbBrand.ColumnWidth = 300
        Me.clbBrand.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbBrand.FormattingEnabled = True
        Me.clbBrand.Location = New System.Drawing.Point(2, 15)
        Me.clbBrand.MultiColumn = True
        Me.clbBrand.Name = "clbBrand"
        Me.clbBrand.Size = New System.Drawing.Size(357, 109)
        Me.clbBrand.TabIndex = 3
        Me.clbBrand.ThreeDCheckBoxes = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(9, 141)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(835, 445)
        Me.TabControl1.TabIndex = 2
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Crv1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Size = New System.Drawing.Size(827, 419)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Report"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Crv1
        '
        Me.Crv1.ActiveViewIndex = -1
        Me.Crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Crv1.DisplayGroupTree = False
        Me.Crv1.DisplayStatusBar = False
        Me.Crv1.DisplayToolbar = False
        Me.Crv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Crv1.Location = New System.Drawing.Point(2, 2)
        Me.Crv1.Name = "Crv1"
        Me.Crv1.SelectionFormula = ""
        Me.Crv1.ShowCloseButton = False
        Me.Crv1.ShowGroupTreeButton = False
        Me.Crv1.ShowRefreshButton = False
        Me.Crv1.Size = New System.Drawing.Size(823, 415)
        Me.Crv1.TabIndex = 3
        Me.Crv1.ViewTimeSelectionFormula = ""
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.LblTotalAmount)
        Me.TabPage2.Controls.Add(Me.chk_Diff_Only)
        Me.TabPage2.Controls.Add(Me.chkCounterMeasure1)
        Me.TabPage2.Controls.Add(Me.chkAnalisis1)
        Me.TabPage2.Controls.Add(Me.dgvItem)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Size = New System.Drawing.Size(787, 418)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Analysis && Countermeasure"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'LblTotalAmount
        '
        Me.LblTotalAmount.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblTotalAmount.AutoSize = True
        Me.LblTotalAmount.Location = New System.Drawing.Point(708, 300)
        Me.LblTotalAmount.Name = "LblTotalAmount"
        Me.LblTotalAmount.Size = New System.Drawing.Size(82, 13)
        Me.LblTotalAmount.TabIndex = 16
        Me.LblTotalAmount.Text = "Total Amount: 0"
        Me.LblTotalAmount.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'chk_Diff_Only
        '
        Me.chk_Diff_Only.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chk_Diff_Only.AutoSize = True
        Me.chk_Diff_Only.Checked = True
        Me.chk_Diff_Only.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chk_Diff_Only.Location = New System.Drawing.Point(6, 300)
        Me.chk_Diff_Only.Name = "chk_Diff_Only"
        Me.chk_Diff_Only.Size = New System.Drawing.Size(177, 17)
        Me.chk_Diff_Only.TabIndex = 15
        Me.chk_Diff_Only.Text = "Show Item With Difference Only"
        Me.chk_Diff_Only.UseVisualStyleBackColor = True
        '
        'chkCounterMeasure1
        '
        Me.chkCounterMeasure1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkCounterMeasure1.AutoSize = True
        Me.chkCounterMeasure1.Location = New System.Drawing.Point(289, 300)
        Me.chkCounterMeasure1.Name = "chkCounterMeasure1"
        Me.chkCounterMeasure1.Size = New System.Drawing.Size(137, 17)
        Me.chkCounterMeasure1.TabIndex = 14
        Me.chkCounterMeasure1.TabStop = False
        Me.chkCounterMeasure1.Text = "Show Counter Measure"
        Me.chkCounterMeasure1.UseVisualStyleBackColor = True
        '
        'chkAnalisis1
        '
        Me.chkAnalisis1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkAnalisis1.AutoSize = True
        Me.chkAnalisis1.Location = New System.Drawing.Point(189, 300)
        Me.chkAnalisis1.Name = "chkAnalisis1"
        Me.chkAnalisis1.Size = New System.Drawing.Size(94, 17)
        Me.chkAnalisis1.TabIndex = 13
        Me.chkAnalisis1.Text = "Show Analysis"
        Me.chkAnalisis1.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col_no, Me.ColProductCode, Me.ColDesc, Me.ColQtyOnHand, Me.ColQtyTaking, Me.ColDifference, Me.ColAmount, Me.ColAnalisis, Me.ColCounterMeasure, Me.ColEdit, Me.ColPrimaryProduct, Me.ColPrimary})
        Me.dgvItem.Location = New System.Drawing.Point(6, 6)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(783, 288)
        Me.dgvItem.TabIndex = 12
        '
        'Col_no
        '
        Me.Col_no.HeaderText = "No"
        Me.Col_no.Name = "Col_no"
        Me.Col_no.ReadOnly = True
        Me.Col_no.Width = 46
        '
        'ColProductCode
        '
        Me.ColProductCode.HeaderText = "Code No"
        Me.ColProductCode.Name = "ColProductCode"
        Me.ColProductCode.ReadOnly = True
        Me.ColProductCode.Width = 74
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.ReadOnly = True
        Me.ColDesc.Width = 85
        '
        'ColQtyOnHand
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyOnHand.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColQtyOnHand.HeaderText = "Qty On Hand"
        Me.ColQtyOnHand.Name = "ColQtyOnHand"
        Me.ColQtyOnHand.ReadOnly = True
        Me.ColQtyOnHand.Width = 94
        '
        'ColQtyTaking
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyTaking.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyTaking.HeaderText = "Qty Taking"
        Me.ColQtyTaking.Name = "ColQtyTaking"
        Me.ColQtyTaking.ReadOnly = True
        Me.ColQtyTaking.Width = 84
        '
        'ColDifference
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDifference.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColDifference.HeaderText = "Difference"
        Me.ColDifference.Name = "ColDifference"
        Me.ColDifference.ReadOnly = True
        Me.ColDifference.Width = 81
        '
        'ColAmount
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.ReadOnly = True
        Me.ColAmount.Width = 68
        '
        'ColAnalisis
        '
        Me.ColAnalisis.HeaderText = "Analisis"
        Me.ColAnalisis.Name = "ColAnalisis"
        Me.ColAnalisis.ReadOnly = True
        Me.ColAnalisis.Width = 67
        '
        'ColCounterMeasure
        '
        Me.ColCounterMeasure.HeaderText = "Counter Measure"
        Me.ColCounterMeasure.Name = "ColCounterMeasure"
        Me.ColCounterMeasure.ReadOnly = True
        Me.ColCounterMeasure.Width = 113
        '
        'ColEdit
        '
        Me.ColEdit.HeaderText = "Edit"
        Me.ColEdit.Name = "ColEdit"
        Me.ColEdit.ReadOnly = True
        Me.ColEdit.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColEdit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColEdit.Width = 50
        '
        'ColPrimaryProduct
        '
        Me.ColPrimaryProduct.HeaderText = "Primary Product"
        Me.ColPrimaryProduct.Name = "ColPrimaryProduct"
        Me.ColPrimaryProduct.ReadOnly = True
        Me.ColPrimaryProduct.Visible = False
        Me.ColPrimaryProduct.Width = 106
        '
        'ColPrimary
        '
        Me.ColPrimary.HeaderText = "Primary"
        Me.ColPrimary.Name = "ColPrimary"
        Me.ColPrimary.ReadOnly = True
        Me.ColPrimary.Visible = False
        Me.ColPrimary.Width = 66
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.TxtPrimary)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryStockTaking)
        Me.gboDetail.Controls.Add(Me.Button1)
        Me.gboDetail.Controls.Add(Me.LblFormat)
        Me.gboDetail.Controls.Add(Me.cboFormat)
        Me.gboDetail.Controls.Add(Me.chkCounterMeasure)
        Me.gboDetail.Controls.Add(Me.chkAnalysis)
        Me.gboDetail.Controls.Add(Me.btnFindStockTaking)
        Me.gboDetail.Controls.Add(Me.TxtStockTaking)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.btnClose)
        Me.gboDetail.Controls.Add(Me.btnFilter)
        Me.gboDetail.Controls.Add(Me.btnPrint)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.dtpStockTakingDate)
        Me.gboDetail.Controls.Add(Me.TxtReadOnly)
        Me.gboDetail.Location = New System.Drawing.Point(374, 10)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.gboDetail.Size = New System.Drawing.Size(470, 127)
        Me.gboDetail.TabIndex = 1
        Me.gboDetail.TabStop = False
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(10, 93)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimary.TabIndex = 319
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'TxtPrimaryStockTaking
        '
        Me.TxtPrimaryStockTaking.Location = New System.Drawing.Point(206, 54)
        Me.TxtPrimaryStockTaking.Name = "TxtPrimaryStockTaking"
        Me.TxtPrimaryStockTaking.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimaryStockTaking.TabIndex = 317
        Me.TxtPrimaryStockTaking.TabStop = False
        Me.TxtPrimaryStockTaking.Text = "-1"
        Me.TxtPrimaryStockTaking.Visible = False
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(363, 67)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(99, 20)
        Me.Button1.TabIndex = 8
        Me.Button1.TabStop = False
        Me.Button1.Text = "Update Cost"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'LblFormat
        '
        Me.LblFormat.AutoSize = True
        Me.LblFormat.Location = New System.Drawing.Point(8, 71)
        Me.LblFormat.Name = "LblFormat"
        Me.LblFormat.Size = New System.Drawing.Size(39, 13)
        Me.LblFormat.TabIndex = 24
        Me.LblFormat.Text = "Format"
        '
        'cboFormat
        '
        Me.cboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormat.FormattingEnabled = True
        Me.cboFormat.Items.AddRange(New Object() {"Detail", "Recap"})
        Me.cboFormat.Location = New System.Drawing.Point(110, 67)
        Me.cboFormat.Name = "cboFormat"
        Me.cboFormat.Size = New System.Drawing.Size(213, 21)
        Me.cboFormat.TabIndex = 3
        '
        'chkCounterMeasure
        '
        Me.chkCounterMeasure.AutoSize = True
        Me.chkCounterMeasure.Location = New System.Drawing.Point(211, 95)
        Me.chkCounterMeasure.Name = "chkCounterMeasure"
        Me.chkCounterMeasure.Size = New System.Drawing.Size(137, 17)
        Me.chkCounterMeasure.TabIndex = 5
        Me.chkCounterMeasure.TabStop = False
        Me.chkCounterMeasure.Text = "Show Counter Measure"
        Me.chkCounterMeasure.UseVisualStyleBackColor = True
        '
        'chkAnalysis
        '
        Me.chkAnalysis.AutoSize = True
        Me.chkAnalysis.Location = New System.Drawing.Point(110, 95)
        Me.chkAnalysis.Name = "chkAnalysis"
        Me.chkAnalysis.Size = New System.Drawing.Size(94, 17)
        Me.chkAnalysis.TabIndex = 4
        Me.chkAnalysis.Text = "Show Analysis"
        Me.chkAnalysis.UseVisualStyleBackColor = True
        '
        'btnFindStockTaking
        '
        Me.btnFindStockTaking.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindStockTaking.Location = New System.Drawing.Point(320, 17)
        Me.btnFindStockTaking.Name = "btnFindStockTaking"
        Me.btnFindStockTaking.Size = New System.Drawing.Size(37, 23)
        Me.btnFindStockTaking.TabIndex = 1
        Me.btnFindStockTaking.TabStop = False
        Me.btnFindStockTaking.Text = "Find"
        Me.btnFindStockTaking.UseVisualStyleBackColor = True
        '
        'TxtStockTaking
        '
        Me.TxtStockTaking.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtStockTaking.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtStockTaking.Location = New System.Drawing.Point(110, 17)
        Me.TxtStockTaking.Name = "TxtStockTaking"
        Me.TxtStockTaking.Size = New System.Drawing.Size(204, 20)
        Me.TxtStockTaking.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 13)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "Stock Taking No"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(363, 41)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(99, 20)
        Me.btnClose.TabIndex = 7
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(363, 15)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(99, 20)
        Me.btnFilter.TabIndex = 6
        Me.btnFilter.Text = "&Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Location = New System.Drawing.Point(363, 41)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(99, 20)
        Me.btnPrint.TabIndex = 17
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        Me.btnPrint.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 46)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 21
        Me.Label3.Text = "Date"
        '
        'dtpStockTakingDate
        '
        Me.dtpStockTakingDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpStockTakingDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpStockTakingDate.Enabled = False
        Me.dtpStockTakingDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStockTakingDate.Location = New System.Drawing.Point(110, 41)
        Me.dtpStockTakingDate.Name = "dtpStockTakingDate"
        Me.dtpStockTakingDate.ShowUpDown = True
        Me.dtpStockTakingDate.Size = New System.Drawing.Size(247, 20)
        Me.dtpStockTakingDate.TabIndex = 2
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(168, 54)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(100, 20)
        Me.TxtReadOnly.TabIndex = 318
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'FrmLaporanStockTaking_Cost1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(861, 596)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.gboBrand)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimumSize = New System.Drawing.Size(831, 535)
        Me.Name = "FrmLaporanStockTaking_Cost1"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stock Taking Report"
        Me.gboBrand.ResumeLayout(False)
        Me.gboBrand.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gboBrand As System.Windows.Forms.GroupBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents LblFormat As System.Windows.Forms.Label
    Friend WithEvents cboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents chkCounterMeasure As System.Windows.Forms.CheckBox
    Friend WithEvents chkAnalysis As System.Windows.Forms.CheckBox
    Friend WithEvents btnFindStockTaking As System.Windows.Forms.Button
    Friend WithEvents TxtStockTaking As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpStockTakingDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents clbBrand As System.Windows.Forms.CheckedListBox
    Friend WithEvents Crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents chk_Diff_Only As System.Windows.Forms.CheckBox
    Friend WithEvents chkCounterMeasure1 As System.Windows.Forms.CheckBox
    Friend WithEvents chkAnalisis1 As System.Windows.Forms.CheckBox
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents Col_no As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyOnHand As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyTaking As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDifference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAnalisis As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCounterMeasure As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEdit As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColPrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TxtPrimaryStockTaking As System.Windows.Forms.TextBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents LblTotalAmount As System.Windows.Forms.Label
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
End Class
