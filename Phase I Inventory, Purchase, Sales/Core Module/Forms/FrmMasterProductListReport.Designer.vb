<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterProductListReport
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.BtnPrint = New System.Windows.Forms.Button
        Me.RboProduct = New System.Windows.Forms.RadioButton
        Me.RboCategory = New System.Windows.Forms.RadioButton
        Me.RboSubBrand = New System.Windows.Forms.RadioButton
        Me.RboBrand = New System.Windows.Forms.RadioButton
        Me.RboProductGroup = New System.Windows.Forms.RadioButton
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.LinkLabel5 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel6 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel3 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel4 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.DgvSubBrand = New System.Windows.Forms.DataGridView
        Me.ColNo_SB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCode_SB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc_SB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrint_SB = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPrimary_SB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label3 = New System.Windows.Forms.Label
        Me.DgvBrand = New System.Windows.Forms.DataGridView
        Me.ColNo_B = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCode_B = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc_B = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrint_B = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPrimary_B = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label2 = New System.Windows.Forms.Label
        Me.DgvProductGroup = New System.Windows.Forms.DataGridView
        Me.ColNo_PG = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCode_PG = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc_PG = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrint_PG = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPrimary_PG = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.LinkLabel7 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel8 = New System.Windows.Forms.LinkLabel
        Me.DgvProduct = New System.Windows.Forms.DataGridView
        Me.ColNo_P = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCode_P = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc_P = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrint_P = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPrimary_P = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label5 = New System.Windows.Forms.Label
        Me.DgvCategory = New System.Windows.Forms.DataGridView
        Me.ColNo_C = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCode_C = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc_C = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrint_C = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPrimary_C = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label4 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DgvSubBrand, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DgvBrand, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DgvProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DgvProduct, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DgvCategory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.BtnPrint)
        Me.GroupBox1.Controls.Add(Me.RboProduct)
        Me.GroupBox1.Controls.Add(Me.RboCategory)
        Me.GroupBox1.Controls.Add(Me.RboSubBrand)
        Me.GroupBox1.Controls.Add(Me.RboBrand)
        Me.GroupBox1.Controls.Add(Me.RboProductGroup)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 1)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(760, 37)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(5, 15)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(39, 13)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Format"
        '
        'BtnPrint
        '
        Me.BtnPrint.Location = New System.Drawing.Point(416, 10)
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(75, 23)
        Me.BtnPrint.TabIndex = 5
        Me.BtnPrint.Text = "Print"
        Me.BtnPrint.UseVisualStyleBackColor = True
        '
        'RboProduct
        '
        Me.RboProduct.AutoSize = True
        Me.RboProduct.Location = New System.Drawing.Point(352, 13)
        Me.RboProduct.Name = "RboProduct"
        Me.RboProduct.Size = New System.Drawing.Size(62, 17)
        Me.RboProduct.TabIndex = 4
        Me.RboProduct.Text = "Product"
        Me.RboProduct.UseVisualStyleBackColor = True
        '
        'RboCategory
        '
        Me.RboCategory.AutoSize = True
        Me.RboCategory.Location = New System.Drawing.Point(281, 13)
        Me.RboCategory.Name = "RboCategory"
        Me.RboCategory.Size = New System.Drawing.Size(67, 17)
        Me.RboCategory.TabIndex = 3
        Me.RboCategory.Text = "Category"
        Me.RboCategory.UseVisualStyleBackColor = True
        '
        'RboSubBrand
        '
        Me.RboSubBrand.AutoSize = True
        Me.RboSubBrand.Location = New System.Drawing.Point(203, 13)
        Me.RboSubBrand.Name = "RboSubBrand"
        Me.RboSubBrand.Size = New System.Drawing.Size(75, 17)
        Me.RboSubBrand.TabIndex = 2
        Me.RboSubBrand.Text = "Sub Brand"
        Me.RboSubBrand.UseVisualStyleBackColor = True
        '
        'RboBrand
        '
        Me.RboBrand.AutoSize = True
        Me.RboBrand.Location = New System.Drawing.Point(147, 13)
        Me.RboBrand.Name = "RboBrand"
        Me.RboBrand.Size = New System.Drawing.Size(53, 17)
        Me.RboBrand.TabIndex = 1
        Me.RboBrand.Text = "Brand"
        Me.RboBrand.UseVisualStyleBackColor = True
        '
        'RboProductGroup
        '
        Me.RboProductGroup.AutoSize = True
        Me.RboProductGroup.Checked = True
        Me.RboProductGroup.Location = New System.Drawing.Point(50, 13)
        Me.RboProductGroup.Name = "RboProductGroup"
        Me.RboProductGroup.Size = New System.Drawing.Size(94, 17)
        Me.RboProductGroup.TabIndex = 0
        Me.RboProductGroup.TabStop = True
        Me.RboProductGroup.Text = "Product Group"
        Me.RboProductGroup.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(12, 45)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.LinkLabel5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LinkLabel6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LinkLabel3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LinkLabel4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LinkLabel2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LinkLabel1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DgvSubBrand)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DgvBrand)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DgvProductGroup)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.LinkLabel7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LinkLabel8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.DgvProduct)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.DgvCategory)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Size = New System.Drawing.Size(760, 505)
        Me.SplitContainer1.SplitterDistance = 379
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'LinkLabel5
        '
        Me.LinkLabel5.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LinkLabel5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel5.AutoSize = True
        Me.LinkLabel5.Location = New System.Drawing.Point(310, 296)
        Me.LinkLabel5.Name = "LinkLabel5"
        Me.LinkLabel5.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel5.TabIndex = 7
        Me.LinkLabel5.TabStop = True
        Me.LinkLabel5.Text = "Uncheck All"
        Me.LinkLabel5.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'LinkLabel6
        '
        Me.LinkLabel6.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LinkLabel6.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel6.AutoSize = True
        Me.LinkLabel6.Location = New System.Drawing.Point(255, 296)
        Me.LinkLabel6.Name = "LinkLabel6"
        Me.LinkLabel6.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel6.TabIndex = 6
        Me.LinkLabel6.TabStop = True
        Me.LinkLabel6.Text = "Check All"
        Me.LinkLabel6.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'LinkLabel3
        '
        Me.LinkLabel3.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LinkLabel3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel3.AutoSize = True
        Me.LinkLabel3.Location = New System.Drawing.Point(310, 152)
        Me.LinkLabel3.Name = "LinkLabel3"
        Me.LinkLabel3.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel3.TabIndex = 4
        Me.LinkLabel3.TabStop = True
        Me.LinkLabel3.Text = "Uncheck All"
        Me.LinkLabel3.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'LinkLabel4
        '
        Me.LinkLabel4.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LinkLabel4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel4.AutoSize = True
        Me.LinkLabel4.Location = New System.Drawing.Point(255, 152)
        Me.LinkLabel4.Name = "LinkLabel4"
        Me.LinkLabel4.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel4.TabIndex = 3
        Me.LinkLabel4.TabStop = True
        Me.LinkLabel4.Text = "Check All"
        Me.LinkLabel4.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'LinkLabel2
        '
        Me.LinkLabel2.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(313, 7)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel2.TabIndex = 1
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Uncheck All"
        Me.LinkLabel2.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'LinkLabel1
        '
        Me.LinkLabel1.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(257, 7)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel1.TabIndex = 0
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Check All"
        Me.LinkLabel1.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'DgvSubBrand
        '
        Me.DgvSubBrand.AllowUserToAddRows = False
        Me.DgvSubBrand.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.DgvSubBrand.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DgvSubBrand.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvSubBrand.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvSubBrand.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvSubBrand.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvSubBrand.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo_SB, Me.ColCode_SB, Me.ColDesc_SB, Me.ColPrint_SB, Me.ColPrimary_SB})
        Me.DgvSubBrand.Location = New System.Drawing.Point(6, 315)
        Me.DgvSubBrand.Name = "DgvSubBrand"
        Me.DgvSubBrand.RowHeadersVisible = False
        Me.DgvSubBrand.Size = New System.Drawing.Size(369, 187)
        Me.DgvSubBrand.TabIndex = 8
        '
        'ColNo_SB
        '
        Me.ColNo_SB.HeaderText = "No"
        Me.ColNo_SB.Name = "ColNo_SB"
        Me.ColNo_SB.Width = 46
        '
        'ColCode_SB
        '
        Me.ColCode_SB.HeaderText = "Code"
        Me.ColCode_SB.Name = "ColCode_SB"
        Me.ColCode_SB.Width = 57
        '
        'ColDesc_SB
        '
        Me.ColDesc_SB.HeaderText = "Desc"
        Me.ColDesc_SB.Name = "ColDesc_SB"
        Me.ColDesc_SB.Width = 57
        '
        'ColPrint_SB
        '
        Me.ColPrint_SB.HeaderText = "Show"
        Me.ColPrint_SB.Name = "ColPrint_SB"
        Me.ColPrint_SB.Width = 40
        '
        'ColPrimary_SB
        '
        Me.ColPrimary_SB.HeaderText = "Primary"
        Me.ColPrimary_SB.Name = "ColPrimary_SB"
        Me.ColPrimary_SB.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPrimary_SB.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimary_SB.Visible = False
        Me.ColPrimary_SB.Width = 47
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 296)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(57, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Sub Brand"
        '
        'DgvBrand
        '
        Me.DgvBrand.AllowUserToAddRows = False
        Me.DgvBrand.AllowUserToDeleteRows = False
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.DgvBrand.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle2
        Me.DgvBrand.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvBrand.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvBrand.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvBrand.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvBrand.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo_B, Me.ColCode_B, Me.ColDesc_B, Me.ColPrint_B, Me.ColPrimary_B})
        Me.DgvBrand.Location = New System.Drawing.Point(6, 171)
        Me.DgvBrand.Name = "DgvBrand"
        Me.DgvBrand.RowHeadersVisible = False
        Me.DgvBrand.Size = New System.Drawing.Size(369, 117)
        Me.DgvBrand.TabIndex = 5
        '
        'ColNo_B
        '
        Me.ColNo_B.HeaderText = "No"
        Me.ColNo_B.Name = "ColNo_B"
        Me.ColNo_B.Width = 46
        '
        'ColCode_B
        '
        Me.ColCode_B.HeaderText = "Code"
        Me.ColCode_B.Name = "ColCode_B"
        Me.ColCode_B.Width = 57
        '
        'ColDesc_B
        '
        Me.ColDesc_B.HeaderText = "Desc"
        Me.ColDesc_B.Name = "ColDesc_B"
        Me.ColDesc_B.Width = 57
        '
        'ColPrint_B
        '
        Me.ColPrint_B.HeaderText = "Show"
        Me.ColPrint_B.Name = "ColPrint_B"
        Me.ColPrint_B.Width = 40
        '
        'ColPrimary_B
        '
        Me.ColPrimary_B.HeaderText = "Primary"
        Me.ColPrimary_B.Name = "ColPrimary_B"
        Me.ColPrimary_B.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPrimary_B.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimary_B.Visible = False
        Me.ColPrimary_B.Width = 47
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 152)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Brand"
        '
        'DgvProductGroup
        '
        Me.DgvProductGroup.AllowUserToAddRows = False
        Me.DgvProductGroup.AllowUserToDeleteRows = False
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.DgvProductGroup.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle3
        Me.DgvProductGroup.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvProductGroup.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo_PG, Me.ColCode_PG, Me.ColDesc_PG, Me.ColPrint_PG, Me.ColPrimary_PG})
        Me.DgvProductGroup.Location = New System.Drawing.Point(6, 25)
        Me.DgvProductGroup.Name = "DgvProductGroup"
        Me.DgvProductGroup.RowHeadersVisible = False
        Me.DgvProductGroup.Size = New System.Drawing.Size(369, 117)
        Me.DgvProductGroup.TabIndex = 2
        '
        'ColNo_PG
        '
        Me.ColNo_PG.HeaderText = "No"
        Me.ColNo_PG.Name = "ColNo_PG"
        Me.ColNo_PG.Width = 46
        '
        'ColCode_PG
        '
        Me.ColCode_PG.HeaderText = "Code"
        Me.ColCode_PG.Name = "ColCode_PG"
        Me.ColCode_PG.Width = 57
        '
        'ColDesc_PG
        '
        Me.ColDesc_PG.HeaderText = "Desc"
        Me.ColDesc_PG.Name = "ColDesc_PG"
        Me.ColDesc_PG.Width = 57
        '
        'ColPrint_PG
        '
        Me.ColPrint_PG.HeaderText = "Show"
        Me.ColPrint_PG.Name = "ColPrint_PG"
        Me.ColPrint_PG.Width = 40
        '
        'ColPrimary_PG
        '
        Me.ColPrimary_PG.HeaderText = "Primary"
        Me.ColPrimary_PG.Name = "ColPrimary_PG"
        Me.ColPrimary_PG.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPrimary_PG.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimary_PG.Visible = False
        Me.ColPrimary_PG.Width = 47
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Product Group"
        '
        'LinkLabel7
        '
        Me.LinkLabel7.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LinkLabel7.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel7.AutoSize = True
        Me.LinkLabel7.Location = New System.Drawing.Point(310, 9)
        Me.LinkLabel7.Name = "LinkLabel7"
        Me.LinkLabel7.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel7.TabIndex = 1
        Me.LinkLabel7.TabStop = True
        Me.LinkLabel7.Text = "Uncheck All"
        Me.LinkLabel7.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'LinkLabel8
        '
        Me.LinkLabel8.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LinkLabel8.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel8.AutoSize = True
        Me.LinkLabel8.Location = New System.Drawing.Point(255, 9)
        Me.LinkLabel8.Name = "LinkLabel8"
        Me.LinkLabel8.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel8.TabIndex = 0
        Me.LinkLabel8.TabStop = True
        Me.LinkLabel8.Text = "Check All"
        Me.LinkLabel8.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'DgvProduct
        '
        Me.DgvProduct.AllowUserToAddRows = False
        Me.DgvProduct.AllowUserToDeleteRows = False
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.DgvProduct.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle4
        Me.DgvProduct.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvProduct.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvProduct.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvProduct.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvProduct.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo_P, Me.ColCode_P, Me.ColDesc_P, Me.ColPrint_P, Me.ColPrimary_P})
        Me.DgvProduct.Location = New System.Drawing.Point(4, 171)
        Me.DgvProduct.Name = "DgvProduct"
        Me.DgvProduct.RowHeadersVisible = False
        Me.DgvProduct.Size = New System.Drawing.Size(373, 331)
        Me.DgvProduct.TabIndex = 3
        '
        'ColNo_P
        '
        Me.ColNo_P.HeaderText = "No"
        Me.ColNo_P.Name = "ColNo_P"
        Me.ColNo_P.Width = 46
        '
        'ColCode_P
        '
        Me.ColCode_P.HeaderText = "Code"
        Me.ColCode_P.Name = "ColCode_P"
        Me.ColCode_P.Width = 57
        '
        'ColDesc_P
        '
        Me.ColDesc_P.HeaderText = "Desc"
        Me.ColDesc_P.Name = "ColDesc_P"
        Me.ColDesc_P.Width = 57
        '
        'ColPrint_P
        '
        Me.ColPrint_P.HeaderText = "Show"
        Me.ColPrint_P.Name = "ColPrint_P"
        Me.ColPrint_P.Visible = False
        Me.ColPrint_P.Width = 40
        '
        'ColPrimary_P
        '
        Me.ColPrimary_P.HeaderText = "Primary"
        Me.ColPrimary_P.Name = "ColPrimary_P"
        Me.ColPrimary_P.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPrimary_P.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimary_P.Visible = False
        Me.ColPrimary_P.Width = 47
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(1, 152)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(47, 13)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Product "
        '
        'DgvCategory
        '
        Me.DgvCategory.AllowUserToAddRows = False
        Me.DgvCategory.AllowUserToDeleteRows = False
        DataGridViewCellStyle5.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.DgvCategory.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle5
        Me.DgvCategory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvCategory.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvCategory.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvCategory.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvCategory.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo_C, Me.ColCode_C, Me.ColDesc_C, Me.ColPrint_C, Me.ColPrimary_C})
        Me.DgvCategory.Location = New System.Drawing.Point(4, 25)
        Me.DgvCategory.Name = "DgvCategory"
        Me.DgvCategory.RowHeadersVisible = False
        Me.DgvCategory.Size = New System.Drawing.Size(373, 117)
        Me.DgvCategory.TabIndex = 2
        '
        'ColNo_C
        '
        Me.ColNo_C.HeaderText = "No"
        Me.ColNo_C.Name = "ColNo_C"
        Me.ColNo_C.Width = 46
        '
        'ColCode_C
        '
        Me.ColCode_C.HeaderText = "Code"
        Me.ColCode_C.Name = "ColCode_C"
        Me.ColCode_C.Width = 57
        '
        'ColDesc_C
        '
        Me.ColDesc_C.HeaderText = "Desc"
        Me.ColDesc_C.Name = "ColDesc_C"
        Me.ColDesc_C.Width = 57
        '
        'ColPrint_C
        '
        Me.ColPrint_C.HeaderText = "Show"
        Me.ColPrint_C.Name = "ColPrint_C"
        Me.ColPrint_C.Width = 40
        '
        'ColPrimary_C
        '
        Me.ColPrimary_C.HeaderText = "Primary"
        Me.ColPrimary_C.Name = "ColPrimary_C"
        Me.ColPrimary_C.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPrimary_C.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimary_C.Visible = False
        Me.ColPrimary_C.Width = 47
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(1, 6)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Category"
        '
        'FrmMasterProductListReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 561)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MinimumSize = New System.Drawing.Size(800, 599)
        Me.Name = "FrmMasterProductListReport"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Product List Report"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DgvSubBrand, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DgvBrand, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DgvProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DgvProduct, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DgvCategory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents BtnPrint As System.Windows.Forms.Button
    Friend WithEvents RboProduct As System.Windows.Forms.RadioButton
    Friend WithEvents RboCategory As System.Windows.Forms.RadioButton
    Friend WithEvents RboSubBrand As System.Windows.Forms.RadioButton
    Friend WithEvents RboBrand As System.Windows.Forms.RadioButton
    Friend WithEvents RboProductGroup As System.Windows.Forms.RadioButton
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DgvProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents DgvSubBrand As System.Windows.Forms.DataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DgvBrand As System.Windows.Forms.DataGridView
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DgvCategory As System.Windows.Forms.DataGridView
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DgvProduct As System.Windows.Forms.DataGridView
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel5 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel6 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel3 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel4 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel7 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel8 As System.Windows.Forms.LinkLabel
    Friend WithEvents ColNo_B As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCode_B As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc_B As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrint_B As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPrimary_B As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNo_C As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCode_C As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc_C As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrint_C As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPrimary_C As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNo_P As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCode_P As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc_P As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrint_P As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPrimary_P As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNo_SB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCode_SB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc_SB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrint_SB As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPrimary_SB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNo_PG As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCode_PG As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc_PG As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrint_PG As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPrimary_PG As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label6 As System.Windows.Forms.Label

End Class
