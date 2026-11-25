<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ManufacturingOrder_BOM
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
        Dim TreeNode1 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Result:")
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.chkShowProses = New System.Windows.Forms.CheckBox
        Me.trvBOM = New System.Windows.Forms.TreeView
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.TxtProcessDesc = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.txtPrimaryProses = New System.Windows.Forms.TextBox
        Me.btnFindProcess = New System.Windows.Forms.Button
        Me.txtProcessCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtTolerance = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.TxtQtyRef = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.TxtGoodPercentage = New System.Windows.Forms.TextBox
        Me.LblUnitBOM = New System.Windows.Forms.Label
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage7 = New System.Windows.Forms.TabPage
        Me.btnBoMTools = New System.Windows.Forms.Button
        Me.BtnClose = New System.Windows.Forms.Button
        Me.BtnOK = New System.Windows.Forms.Button
        Me.llbAddNewProduct = New System.Windows.Forms.LinkLabel
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.dgvItemBOM = New System.Windows.Forms.DataGridView
        Me.ColBOM_DrillDown = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColBOM_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_GoodPercentage = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_Tolerance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_PrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_PrimaryParentProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_IdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_IdentifierBOM = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_IdentifierBOM_Parent = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBOM_IdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage7.SuspendLayout()
        CType(Me.dgvItemBOM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.chkShowProses)
        Me.SplitContainer3.Panel1.Controls.Add(Me.trvBOM)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TabControl2)
        Me.SplitContainer3.Size = New System.Drawing.Size(1065, 383)
        Me.SplitContainer3.SplitterDistance = 339
        Me.SplitContainer3.SplitterWidth = 5
        Me.SplitContainer3.TabIndex = 1
        Me.SplitContainer3.TabStop = False
        '
        'chkShowProses
        '
        Me.chkShowProses.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkShowProses.AutoSize = True
        Me.chkShowProses.Location = New System.Drawing.Point(4, 357)
        Me.chkShowProses.Margin = New System.Windows.Forms.Padding(4)
        Me.chkShowProses.Name = "chkShowProses"
        Me.chkShowProses.Size = New System.Drawing.Size(112, 21)
        Me.chkShowProses.TabIndex = 1
        Me.chkShowProses.Text = "Show Proses"
        Me.chkShowProses.UseVisualStyleBackColor = True
        '
        'trvBOM
        '
        Me.trvBOM.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.trvBOM.HotTracking = True
        Me.trvBOM.Location = New System.Drawing.Point(0, 0)
        Me.trvBOM.Margin = New System.Windows.Forms.Padding(4)
        Me.trvBOM.Name = "trvBOM"
        TreeNode1.Name = "Node0"
        TreeNode1.Text = "Result:"
        Me.trvBOM.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode1})
        Me.trvBOM.Size = New System.Drawing.Size(338, 343)
        Me.trvBOM.TabIndex = 0
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer4.Location = New System.Drawing.Point(3, 7)
        Me.SplitContainer4.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtProcessDesc)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label23)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label24)
        Me.SplitContainer4.Panel1.Controls.Add(Me.txtPrimaryProses)
        Me.SplitContainer4.Panel1.Controls.Add(Me.btnFindProcess)
        Me.SplitContainer4.Panel1.Controls.Add(Me.txtProcessCode)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer4.Panel2.Controls.Add(Me.txtTolerance)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label25)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtQtyRef)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label26)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label27)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtGoodPercentage)
        Me.SplitContainer4.Panel2.Controls.Add(Me.LblUnitBOM)
        Me.SplitContainer4.Size = New System.Drawing.Size(705, 57)
        Me.SplitContainer4.SplitterDistance = 331
        Me.SplitContainer4.SplitterWidth = 5
        Me.SplitContainer4.TabIndex = 0
        Me.SplitContainer4.TabStop = False
        '
        'TxtProcessDesc
        '
        Me.TxtProcessDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtProcessDesc.Location = New System.Drawing.Point(100, 32)
        Me.TxtProcessDesc.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtProcessDesc.MaxLength = 150
        Me.TxtProcessDesc.Name = "TxtProcessDesc"
        Me.TxtProcessDesc.ReadOnly = True
        Me.TxtProcessDesc.Size = New System.Drawing.Size(224, 22)
        Me.TxtProcessDesc.TabIndex = 2
        Me.TxtProcessDesc.TabStop = False
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(4, 36)
        Me.Label23.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(79, 17)
        Me.Label23.TabIndex = 42
        Me.Label23.Text = "Description"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(3, 5)
        Me.Label24.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(96, 17)
        Me.Label24.TabIndex = 39
        Me.Label24.Text = "Process Code"
        '
        'txtPrimaryProses
        '
        Me.txtPrimaryProses.Location = New System.Drawing.Point(27, 7)
        Me.txtPrimaryProses.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrimaryProses.Name = "txtPrimaryProses"
        Me.txtPrimaryProses.Size = New System.Drawing.Size(48, 22)
        Me.txtPrimaryProses.TabIndex = 40
        Me.txtPrimaryProses.TabStop = False
        Me.txtPrimaryProses.Text = "-1"
        Me.txtPrimaryProses.Visible = False
        '
        'btnFindProcess
        '
        Me.btnFindProcess.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindProcess.Location = New System.Drawing.Point(275, 0)
        Me.btnFindProcess.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindProcess.Name = "btnFindProcess"
        Me.btnFindProcess.Size = New System.Drawing.Size(52, 25)
        Me.btnFindProcess.TabIndex = 1
        Me.btnFindProcess.TabStop = False
        Me.btnFindProcess.Text = "...."
        Me.btnFindProcess.UseVisualStyleBackColor = True
        '
        'txtProcessCode
        '
        Me.txtProcessCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtProcessCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtProcessCode.Location = New System.Drawing.Point(100, 0)
        Me.txtProcessCode.Margin = New System.Windows.Forms.Padding(4)
        Me.txtProcessCode.MaxLength = 150
        Me.txtProcessCode.Name = "txtProcessCode"
        Me.txtProcessCode.Size = New System.Drawing.Size(167, 22)
        Me.txtProcessCode.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(342, 35)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(20, 17)
        Me.Label1.TabIndex = 53
        Me.Label1.Text = "%"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(207, 35)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 17)
        Me.Label2.TabIndex = 52
        Me.Label2.Text = "Tolerance"
        '
        'txtTolerance
        '
        Me.txtTolerance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTolerance.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTolerance.Location = New System.Drawing.Point(284, 32)
        Me.txtTolerance.Margin = New System.Windows.Forms.Padding(4)
        Me.txtTolerance.MaxLength = 7
        Me.txtTolerance.Name = "txtTolerance"
        Me.txtTolerance.Size = New System.Drawing.Size(50, 22)
        Me.txtTolerance.TabIndex = 51
        Me.txtTolerance.Text = "0.00"
        Me.txtTolerance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(184, 36)
        Me.Label25.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(20, 17)
        Me.Label25.TabIndex = 50
        Me.Label25.Text = "%"
        '
        'TxtQtyRef
        '
        Me.TxtQtyRef.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyRef.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtQtyRef.Location = New System.Drawing.Point(133, 0)
        Me.TxtQtyRef.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtQtyRef.MaxLength = 7
        Me.TxtQtyRef.Name = "TxtQtyRef"
        Me.TxtQtyRef.Size = New System.Drawing.Size(186, 22)
        Me.TxtQtyRef.TabIndex = 0
        Me.TxtQtyRef.Text = "1"
        Me.TxtQtyRef.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(4, 36)
        Me.Label26.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(120, 17)
        Me.Label26.TabIndex = 49
        Me.Label26.Text = "Good Percentage"
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(4, 5)
        Me.Label27.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(100, 17)
        Me.Label27.TabIndex = 44
        Me.Label27.Text = "Reference Qty"
        '
        'TxtGoodPercentage
        '
        Me.TxtGoodPercentage.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtGoodPercentage.Location = New System.Drawing.Point(133, 32)
        Me.TxtGoodPercentage.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtGoodPercentage.MaxLength = 7
        Me.TxtGoodPercentage.Name = "TxtGoodPercentage"
        Me.TxtGoodPercentage.Size = New System.Drawing.Size(50, 22)
        Me.TxtGoodPercentage.TabIndex = 1
        Me.TxtGoodPercentage.Text = "100.00"
        Me.TxtGoodPercentage.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblUnitBOM
        '
        Me.LblUnitBOM.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnitBOM.AutoSize = True
        Me.LblUnitBOM.Location = New System.Drawing.Point(329, 5)
        Me.LblUnitBOM.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblUnitBOM.Name = "LblUnitBOM"
        Me.LblUnitBOM.Size = New System.Drawing.Size(33, 17)
        Me.LblUnitBOM.TabIndex = 47
        Me.LblUnitBOM.Text = "Unit"
        '
        'TabControl2
        '
        Me.TabControl2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl2.Controls.Add(Me.TabPage7)
        Me.TabControl2.Location = New System.Drawing.Point(3, 71)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(712, 312)
        Me.TabControl2.TabIndex = 2
        '
        'TabPage7
        '
        Me.TabPage7.Controls.Add(Me.btnBoMTools)
        Me.TabPage7.Controls.Add(Me.BtnClose)
        Me.TabPage7.Controls.Add(Me.BtnOK)
        Me.TabPage7.Controls.Add(Me.llbAddNewProduct)
        Me.TabPage7.Controls.Add(Me.Splitter1)
        Me.TabPage7.Controls.Add(Me.btnUP)
        Me.TabPage7.Controls.Add(Me.btnDown)
        Me.TabPage7.Controls.Add(Me.dgvItemBOM)
        Me.TabPage7.Location = New System.Drawing.Point(4, 25)
        Me.TabPage7.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage7.Size = New System.Drawing.Size(704, 283)
        Me.TabPage7.TabIndex = 0
        Me.TabPage7.Text = "Material"
        Me.TabPage7.UseVisualStyleBackColor = True
        '
        'btnBoMTools
        '
        Me.btnBoMTools.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBoMTools.Location = New System.Drawing.Point(384, 243)
        Me.btnBoMTools.Margin = New System.Windows.Forms.Padding(4)
        Me.btnBoMTools.Name = "btnBoMTools"
        Me.btnBoMTools.Size = New System.Drawing.Size(100, 28)
        Me.btnBoMTools.TabIndex = 21
        Me.btnBoMTools.Text = "BoM Tools"
        Me.btnBoMTools.UseVisualStyleBackColor = True
        Me.btnBoMTools.Visible = False
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(600, 243)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(100, 28)
        Me.BtnClose.TabIndex = 23
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'BtnOK
        '
        Me.BtnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnOK.Location = New System.Drawing.Point(492, 243)
        Me.BtnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(100, 28)
        Me.BtnOK.TabIndex = 22
        Me.BtnOK.Text = "OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'llbAddNewProduct
        '
        Me.llbAddNewProduct.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbAddNewProduct.AutoSize = True
        Me.llbAddNewProduct.Location = New System.Drawing.Point(96, 252)
        Me.llbAddNewProduct.Name = "llbAddNewProduct"
        Me.llbAddNewProduct.Size = New System.Drawing.Size(117, 17)
        Me.llbAddNewProduct.TabIndex = 20
        Me.llbAddNewProduct.TabStop = True
        Me.llbAddNewProduct.Text = "Add New Product"
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(3, 2)
        Me.Splitter1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 279)
        Me.Splitter1.TabIndex = 6
        Me.Splitter1.TabStop = False
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.Inventory_Module.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(7, 243)
        Me.btnUP.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(37, 32)
        Me.btnUP.TabIndex = 4
        Me.btnUP.TabStop = False
        Me.btnUP.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = Global.Inventory_Module.My.Resources.Resources.arrow_down
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(52, 243)
        Me.btnDown.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(37, 32)
        Me.btnDown.TabIndex = 5
        Me.btnDown.TabStop = False
        Me.btnDown.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'dgvItemBOM
        '
        Me.dgvItemBOM.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemBOM.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemBOM.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColBOM_DrillDown, Me.ColBOM_CodeNo, Me.ColBOM_Desc, Me.ColBOM_Qty, Me.ColBOM_Unit, Me.ColBOM_GoodPercentage, Me.ColBOM_Tolerance, Me.ColBOM_Remarks, Me.ColBOM_PrimaryProduct, Me.ColBOM_PrimaryParentProduct, Me.ColBOM_IdentifierWO, Me.ColBOM_IdentifierBOM, Me.ColBOM_IdentifierBOM_Parent, Me.ColBOM_IdentifierJO})
        Me.dgvItemBOM.Location = New System.Drawing.Point(5, 2)
        Me.dgvItemBOM.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItemBOM.MultiSelect = False
        Me.dgvItemBOM.Name = "dgvItemBOM"
        Me.dgvItemBOM.RowTemplate.Height = 24
        Me.dgvItemBOM.Size = New System.Drawing.Size(693, 234)
        Me.dgvItemBOM.StandardTab = True
        Me.dgvItemBOM.TabIndex = 3
        '
        'ColBOM_DrillDown
        '
        Me.ColBOM_DrillDown.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle1.NullValue = "New"
        Me.ColBOM_DrillDown.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColBOM_DrillDown.Frozen = True
        Me.ColBOM_DrillDown.HeaderText = ""
        Me.ColBOM_DrillDown.Name = "ColBOM_DrillDown"
        Me.ColBOM_DrillDown.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColBOM_DrillDown.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColBOM_DrillDown.Width = 39
        '
        'ColBOM_CodeNo
        '
        Me.ColBOM_CodeNo.Frozen = True
        Me.ColBOM_CodeNo.HeaderText = "Code No"
        Me.ColBOM_CodeNo.Name = "ColBOM_CodeNo"
        Me.ColBOM_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBOM_CodeNo.Width = 69
        '
        'ColBOM_Desc
        '
        Me.ColBOM_Desc.HeaderText = "Desc"
        Me.ColBOM_Desc.Name = "ColBOM_Desc"
        Me.ColBOM_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBOM_Desc.Width = 46
        '
        'ColBOM_Qty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle2.NullValue = "0"
        Me.ColBOM_Qty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColBOM_Qty.HeaderText = "Qty"
        Me.ColBOM_Qty.Name = "ColBOM_Qty"
        Me.ColBOM_Qty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBOM_Qty.Width = 36
        '
        'ColBOM_Unit
        '
        Me.ColBOM_Unit.HeaderText = "Unit"
        Me.ColBOM_Unit.Name = "ColBOM_Unit"
        Me.ColBOM_Unit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBOM_Unit.Width = 39
        '
        'ColBOM_GoodPercentage
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.Format = "N2"
        DataGridViewCellStyle3.NullValue = "100.00"
        Me.ColBOM_GoodPercentage.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColBOM_GoodPercentage.HeaderText = "Good Percentage(%)"
        Me.ColBOM_GoodPercentage.Name = "ColBOM_GoodPercentage"
        Me.ColBOM_GoodPercentage.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBOM_GoodPercentage.Width = 148
        '
        'ColBOM_Tolerance
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle4.NullValue = "0.00"
        Me.ColBOM_Tolerance.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColBOM_Tolerance.HeaderText = "Tolerance (%)"
        Me.ColBOM_Tolerance.Name = "ColBOM_Tolerance"
        Me.ColBOM_Tolerance.Width = 123
        '
        'ColBOM_Remarks
        '
        Me.ColBOM_Remarks.HeaderText = "Remarks"
        Me.ColBOM_Remarks.MaxInputLength = 250
        Me.ColBOM_Remarks.Name = "ColBOM_Remarks"
        Me.ColBOM_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBOM_Remarks.Width = 70
        '
        'ColBOM_PrimaryProduct
        '
        Me.ColBOM_PrimaryProduct.HeaderText = "Primary Product"
        Me.ColBOM_PrimaryProduct.Name = "ColBOM_PrimaryProduct"
        Me.ColBOM_PrimaryProduct.Visible = False
        Me.ColBOM_PrimaryProduct.Width = 134
        '
        'ColBOM_PrimaryParentProduct
        '
        Me.ColBOM_PrimaryParentProduct.HeaderText = "Primary Parent Product"
        Me.ColBOM_PrimaryParentProduct.Name = "ColBOM_PrimaryParentProduct"
        Me.ColBOM_PrimaryParentProduct.ReadOnly = True
        Me.ColBOM_PrimaryParentProduct.Visible = False
        Me.ColBOM_PrimaryParentProduct.Width = 180
        '
        'ColBOM_IdentifierWO
        '
        Me.ColBOM_IdentifierWO.HeaderText = "Identifier WO"
        Me.ColBOM_IdentifierWO.Name = "ColBOM_IdentifierWO"
        Me.ColBOM_IdentifierWO.ReadOnly = True
        Me.ColBOM_IdentifierWO.Visible = False
        Me.ColBOM_IdentifierWO.Width = 115
        '
        'ColBOM_IdentifierBOM
        '
        Me.ColBOM_IdentifierBOM.HeaderText = "Identifier BOM"
        Me.ColBOM_IdentifierBOM.Name = "ColBOM_IdentifierBOM"
        Me.ColBOM_IdentifierBOM.ReadOnly = True
        Me.ColBOM_IdentifierBOM.Visible = False
        Me.ColBOM_IdentifierBOM.Width = 122
        '
        'ColBOM_IdentifierBOM_Parent
        '
        Me.ColBOM_IdentifierBOM_Parent.HeaderText = "Identifier BOM Parent"
        Me.ColBOM_IdentifierBOM_Parent.Name = "ColBOM_IdentifierBOM_Parent"
        Me.ColBOM_IdentifierBOM_Parent.ReadOnly = True
        Me.ColBOM_IdentifierBOM_Parent.Visible = False
        Me.ColBOM_IdentifierBOM_Parent.Width = 168
        '
        'ColBOM_IdentifierJO
        '
        Me.ColBOM_IdentifierJO.HeaderText = "Identifier JO"
        Me.ColBOM_IdentifierJO.Name = "ColBOM_IdentifierJO"
        Me.ColBOM_IdentifierJO.ReadOnly = True
        Me.ColBOM_IdentifierJO.Visible = False
        Me.ColBOM_IdentifierJO.Width = 109
        '
        'BackgroundWorker1
        '
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EditToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(105, 28)
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(104, 24)
        Me.EditToolStripMenuItem.Text = "Edit"
        '
        'FrmTrans_ManufacturingOrder_BOM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(1065, 383)
        Me.Controls.Add(Me.SplitContainer3)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1073, 420)
        Me.Name = "FrmTrans_ManufacturingOrder_BOM"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Bill of Material"
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.PerformLayout()
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.Panel2.PerformLayout()
        Me.SplitContainer4.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage7.ResumeLayout(False)
        Me.TabPage7.PerformLayout()
        CType(Me.dgvItemBOM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents chkShowProses As System.Windows.Forms.CheckBox
    Friend WithEvents trvBOM As System.Windows.Forms.TreeView
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtProcessDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryProses As System.Windows.Forms.TextBox
    Friend WithEvents btnFindProcess As System.Windows.Forms.Button
    Friend WithEvents txtProcessCode As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents TxtQtyRef As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents TxtGoodPercentage As System.Windows.Forms.TextBox
    Friend WithEvents LblUnitBOM As System.Windows.Forms.Label
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Public WithEvents dgvItemBOM As System.Windows.Forms.DataGridView
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents llbAddNewProduct As System.Windows.Forms.LinkLabel
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtTolerance As System.Windows.Forms.TextBox
    Friend WithEvents ColBOM_DrillDown As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColBOM_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBOM_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBOM_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBOM_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBOM_GoodPercentage As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBOM_Tolerance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBOM_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBOM_PrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBOM_PrimaryParentProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBOM_IdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBOM_IdentifierBOM As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBOM_IdentifierBOM_Parent As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBOM_IdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnBoMTools As System.Windows.Forms.Button
End Class
