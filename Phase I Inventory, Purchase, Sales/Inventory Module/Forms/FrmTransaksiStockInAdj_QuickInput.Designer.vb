<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockInAdj_QuickInput
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
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtQty = New System.Windows.Forms.TextBox
        Me.TxtRef = New System.Windows.Forms.TextBox
        Me.chkImportFromDB = New System.Windows.Forms.CheckBox
        Me.ChkMirror = New System.Windows.Forms.CheckBox
        Me.chkCoverMinus = New System.Windows.Forms.CheckBox
        Me.dtpCoverMinus_RefDate = New System.Windows.Forms.DateTimePicker
        Me.TxtCoverMinus_LimitData = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.tabAddByGroup = New System.Windows.Forms.TabControl
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dgvImport = New System.Windows.Forms.DataGridView
        Me.Col_Import_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_PONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_SuppDONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_SecondaryQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_SecondaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_Batch = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_ExpiredDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_Location = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_WareHouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col_Import_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_UnitCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_Invalid = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_PrimaryPOno = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_PrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_PrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_PrimarySecondaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Import_PrimaryPdO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.chkShowErrorOnly = New System.Windows.Forms.CheckBox
        Me.txtFilename = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnSelect = New System.Windows.Forms.Button
        Me.btnLoad = New System.Windows.Forms.Button
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn8 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabAddByGroup.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvImport, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeight = 25
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_Remarks, Me.ColProductGroup_PrimaryAcuan})
        Me.dgvItemProductGroup.Location = New System.Drawing.Point(3, 6)
        Me.dgvItemProductGroup.MultiSelect = False
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(1117, 466)
        Me.dgvItemProductGroup.TabIndex = 3
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
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(447, 477)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 19)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(386, 477)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 19)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'BackgroundWorker3
        '
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 479)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(133, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Please specify default qty: "
        '
        'TxtQty
        '
        Me.TxtQty.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtQty.Location = New System.Drawing.Point(139, 477)
        Me.TxtQty.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtQty.Name = "TxtQty"
        Me.TxtQty.Size = New System.Drawing.Size(84, 20)
        Me.TxtQty.TabIndex = 7
        Me.TxtQty.Text = "0"
        Me.TxtQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtRef
        '
        Me.TxtRef.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtRef.Location = New System.Drawing.Point(227, 477)
        Me.TxtRef.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtRef.Name = "TxtRef"
        Me.TxtRef.Size = New System.Drawing.Size(84, 20)
        Me.TxtRef.TabIndex = 8
        Me.TxtRef.TabStop = False
        Me.TxtRef.Text = "0"
        Me.TxtRef.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtRef.Visible = False
        '
        'chkImportFromDB
        '
        Me.chkImportFromDB.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkImportFromDB.AutoSize = True
        Me.chkImportFromDB.Location = New System.Drawing.Point(227, 477)
        Me.chkImportFromDB.Margin = New System.Windows.Forms.Padding(2)
        Me.chkImportFromDB.Name = "chkImportFromDB"
        Me.chkImportFromDB.Size = New System.Drawing.Size(150, 17)
        Me.chkImportFromDB.TabIndex = 9
        Me.chkImportFromDB.Text = "Take data from access db"
        Me.chkImportFromDB.UseVisualStyleBackColor = True
        '
        'ChkMirror
        '
        Me.ChkMirror.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ChkMirror.AutoSize = True
        Me.ChkMirror.Location = New System.Drawing.Point(227, 499)
        Me.ChkMirror.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkMirror.Name = "ChkMirror"
        Me.ChkMirror.Size = New System.Drawing.Size(78, 17)
        Me.ChkMirror.TabIndex = 10
        Me.ChkMirror.Text = "Mirror Data"
        Me.ChkMirror.UseVisualStyleBackColor = True
        '
        'chkCoverMinus
        '
        Me.chkCoverMinus.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkCoverMinus.AutoSize = True
        Me.chkCoverMinus.Location = New System.Drawing.Point(227, 519)
        Me.chkCoverMinus.Margin = New System.Windows.Forms.Padding(2)
        Me.chkCoverMinus.Name = "chkCoverMinus"
        Me.chkCoverMinus.Size = New System.Drawing.Size(111, 17)
        Me.chkCoverMinus.TabIndex = 11
        Me.chkCoverMinus.Text = "Cover Minus Data"
        Me.chkCoverMinus.UseVisualStyleBackColor = True
        '
        'dtpCoverMinus_RefDate
        '
        Me.dtpCoverMinus_RefDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dtpCoverMinus_RefDate.CustomFormat = "dd MMM yyyy"
        Me.dtpCoverMinus_RefDate.Enabled = False
        Me.dtpCoverMinus_RefDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpCoverMinus_RefDate.Location = New System.Drawing.Point(338, 519)
        Me.dtpCoverMinus_RefDate.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpCoverMinus_RefDate.Name = "dtpCoverMinus_RefDate"
        Me.dtpCoverMinus_RefDate.Size = New System.Drawing.Size(148, 20)
        Me.dtpCoverMinus_RefDate.TabIndex = 12
        '
        'TxtCoverMinus_LimitData
        '
        Me.TxtCoverMinus_LimitData.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtCoverMinus_LimitData.Enabled = False
        Me.TxtCoverMinus_LimitData.Location = New System.Drawing.Point(139, 499)
        Me.TxtCoverMinus_LimitData.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtCoverMinus_LimitData.Name = "TxtCoverMinus_LimitData"
        Me.TxtCoverMinus_LimitData.Size = New System.Drawing.Size(84, 20)
        Me.TxtCoverMinus_LimitData.TabIndex = 14
        Me.TxtCoverMinus_LimitData.Text = "0"
        Me.TxtCoverMinus_LimitData.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(2, 501)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(125, 13)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Please specify limit data: "
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.Enabled = False
        Me.TextBox1.Location = New System.Drawing.Point(113, 520)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(84, 20)
        Me.TextBox1.TabIndex = 15
        Me.TextBox1.TabStop = False
        Me.TextBox1.Text = "0"
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox1.Visible = False
        '
        'tabAddByGroup
        '
        Me.tabAddByGroup.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabAddByGroup.Controls.Add(Me.TabPage2)
        Me.tabAddByGroup.Controls.Add(Me.TabPage1)
        Me.tabAddByGroup.Location = New System.Drawing.Point(12, 12)
        Me.tabAddByGroup.Name = "tabAddByGroup"
        Me.tabAddByGroup.SelectedIndex = 0
        Me.tabAddByGroup.Size = New System.Drawing.Size(1131, 571)
        Me.tabAddByGroup.TabIndex = 16
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GroupBox2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1123, 545)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Import"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.dgvImport)
        Me.GroupBox2.Controls.Add(Me.chkShowErrorOnly)
        Me.GroupBox2.Controls.Add(Me.txtFilename)
        Me.GroupBox2.Controls.Add(Me.btnClose)
        Me.GroupBox2.Controls.Add(Me.btnSave)
        Me.GroupBox2.Controls.Add(Me.btnSelect)
        Me.GroupBox2.Controls.Add(Me.btnLoad)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1117, 539)
        Me.GroupBox2.TabIndex = 17
        Me.GroupBox2.TabStop = False
        '
        'dgvImport
        '
        Me.dgvImport.AllowUserToAddRows = False
        Me.dgvImport.AllowUserToDeleteRows = False
        Me.dgvImport.AllowUserToResizeColumns = False
        Me.dgvImport.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvImport.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvImport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvImport.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvImport.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvImport.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvImport.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col_Import_No, Me.Col_Import_PONo, Me.Col_Import_SuppDONo, Me.Col_Import_CodeNo, Me.Col_Import_Description, Me.Col_Import_Qty, Me.Col_Import_Unit, Me.Col_Import_SecondaryQty, Me.Col_Import_SecondaryUnit, Me.Col_Import_Batch, Me.Col_Import_ExpiredDate, Me.Col_Import_Location, Me.Col_Import_WareHouse, Me.col_Import_Remarks, Me.Col_Import_UnitCost, Me.Col_Import_Amount, Me.Col_Import_Invalid, Me.Col_Import_PrimaryPOno, Me.Col_Import_PrimaryCodeNo, Me.Col_Import_PrimaryLocation, Me.Col_Import_PrimaryUnits, Me.Col_Import_PrimarySecondaryUnits, Me.Col_Import_PrimaryPdO})
        Me.dgvImport.Location = New System.Drawing.Point(3, 68)
        Me.dgvImport.MultiSelect = False
        Me.dgvImport.Name = "dgvImport"
        Me.dgvImport.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvImport.RowTemplate.Height = 24
        Me.dgvImport.Size = New System.Drawing.Size(1108, 436)
        Me.dgvImport.StandardTab = True
        Me.dgvImport.TabIndex = 9
        Me.dgvImport.TabStop = False
        '
        'Col_Import_No
        '
        Me.Col_Import_No.Frozen = True
        Me.Col_Import_No.HeaderText = "No"
        Me.Col_Import_No.Name = "Col_Import_No"
        Me.Col_Import_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Import_No.Width = 27
        '
        'Col_Import_PONo
        '
        Me.Col_Import_PONo.Frozen = True
        Me.Col_Import_PONo.HeaderText = "PO. No"
        Me.Col_Import_PONo.Name = "Col_Import_PONo"
        Me.Col_Import_PONo.Visible = False
        Me.Col_Import_PONo.Width = 67
        '
        'Col_Import_SuppDONo
        '
        Me.Col_Import_SuppDONo.Frozen = True
        Me.Col_Import_SuppDONo.HeaderText = "Supp. DO No"
        Me.Col_Import_SuppDONo.Name = "Col_Import_SuppDONo"
        Me.Col_Import_SuppDONo.Visible = False
        Me.Col_Import_SuppDONo.Width = 96
        '
        'Col_Import_CodeNo
        '
        Me.Col_Import_CodeNo.Frozen = True
        Me.Col_Import_CodeNo.HeaderText = "Code No"
        Me.Col_Import_CodeNo.Name = "Col_Import_CodeNo"
        Me.Col_Import_CodeNo.Width = 74
        '
        'Col_Import_Description
        '
        Me.Col_Import_Description.Frozen = True
        Me.Col_Import_Description.HeaderText = "Description"
        Me.Col_Import_Description.Name = "Col_Import_Description"
        Me.Col_Import_Description.Width = 85
        '
        'Col_Import_Qty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col_Import_Qty.DefaultCellStyle = DataGridViewCellStyle2
        Me.Col_Import_Qty.HeaderText = "Qty"
        Me.Col_Import_Qty.Name = "Col_Import_Qty"
        Me.Col_Import_Qty.Width = 48
        '
        'Col_Import_Unit
        '
        Me.Col_Import_Unit.HeaderText = "Units"
        Me.Col_Import_Unit.Name = "Col_Import_Unit"
        Me.Col_Import_Unit.Width = 56
        '
        'Col_Import_SecondaryQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col_Import_SecondaryQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.Col_Import_SecondaryQty.HeaderText = "@"
        Me.Col_Import_SecondaryQty.Name = "Col_Import_SecondaryQty"
        Me.Col_Import_SecondaryQty.Width = 43
        '
        'Col_Import_SecondaryUnit
        '
        Me.Col_Import_SecondaryUnit.HeaderText = "Units"
        Me.Col_Import_SecondaryUnit.Name = "Col_Import_SecondaryUnit"
        Me.Col_Import_SecondaryUnit.Width = 56
        '
        'Col_Import_Batch
        '
        Me.Col_Import_Batch.HeaderText = "Batch No"
        Me.Col_Import_Batch.Name = "Col_Import_Batch"
        Me.Col_Import_Batch.Width = 77
        '
        'Col_Import_ExpiredDate
        '
        Me.Col_Import_ExpiredDate.HeaderText = "Expiry Date"
        Me.Col_Import_ExpiredDate.Name = "Col_Import_ExpiredDate"
        Me.Col_Import_ExpiredDate.Width = 86
        '
        'Col_Import_Location
        '
        Me.Col_Import_Location.HeaderText = "Location"
        Me.Col_Import_Location.Name = "Col_Import_Location"
        Me.Col_Import_Location.Width = 73
        '
        'Col_Import_WareHouse
        '
        Me.Col_Import_WareHouse.HeaderText = "WareHouse"
        Me.Col_Import_WareHouse.Name = "Col_Import_WareHouse"
        Me.Col_Import_WareHouse.Width = 89
        '
        'col_Import_Remarks
        '
        Me.col_Import_Remarks.HeaderText = "Remarks"
        Me.col_Import_Remarks.Name = "col_Import_Remarks"
        Me.col_Import_Remarks.Width = 74
        '
        'Col_Import_UnitCost
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col_Import_UnitCost.DefaultCellStyle = DataGridViewCellStyle4
        Me.Col_Import_UnitCost.HeaderText = "Unit Cost"
        Me.Col_Import_UnitCost.Name = "Col_Import_UnitCost"
        Me.Col_Import_UnitCost.Width = 75
        '
        'Col_Import_Amount
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col_Import_Amount.DefaultCellStyle = DataGridViewCellStyle5
        Me.Col_Import_Amount.HeaderText = "Amount"
        Me.Col_Import_Amount.Name = "Col_Import_Amount"
        Me.Col_Import_Amount.Width = 68
        '
        'Col_Import_Invalid
        '
        Me.Col_Import_Invalid.HeaderText = "Invalid"
        Me.Col_Import_Invalid.Name = "Col_Import_Invalid"
        Me.Col_Import_Invalid.Width = 63
        '
        'Col_Import_PrimaryPOno
        '
        Me.Col_Import_PrimaryPOno.HeaderText = "PrimaryPOno"
        Me.Col_Import_PrimaryPOno.Name = "Col_Import_PrimaryPOno"
        Me.Col_Import_PrimaryPOno.Visible = False
        Me.Col_Import_PrimaryPOno.Width = 93
        '
        'Col_Import_PrimaryCodeNo
        '
        Me.Col_Import_PrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.Col_Import_PrimaryCodeNo.Name = "Col_Import_PrimaryCodeNo"
        Me.Col_Import_PrimaryCodeNo.Visible = False
        Me.Col_Import_PrimaryCodeNo.Width = 105
        '
        'Col_Import_PrimaryLocation
        '
        Me.Col_Import_PrimaryLocation.HeaderText = "Primary Location"
        Me.Col_Import_PrimaryLocation.Name = "Col_Import_PrimaryLocation"
        Me.Col_Import_PrimaryLocation.Visible = False
        Me.Col_Import_PrimaryLocation.Width = 110
        '
        'Col_Import_PrimaryUnits
        '
        Me.Col_Import_PrimaryUnits.HeaderText = "PrimaryUnits"
        Me.Col_Import_PrimaryUnits.Name = "Col_Import_PrimaryUnits"
        Me.Col_Import_PrimaryUnits.Visible = False
        Me.Col_Import_PrimaryUnits.Width = 90
        '
        'Col_Import_PrimarySecondaryUnits
        '
        Me.Col_Import_PrimarySecondaryUnits.HeaderText = "PrimarySecondaryUnits"
        Me.Col_Import_PrimarySecondaryUnits.Name = "Col_Import_PrimarySecondaryUnits"
        Me.Col_Import_PrimarySecondaryUnits.Visible = False
        Me.Col_Import_PrimarySecondaryUnits.Width = 141
        '
        'Col_Import_PrimaryPdO
        '
        Me.Col_Import_PrimaryPdO.HeaderText = "Primary PdO"
        Me.Col_Import_PrimaryPdO.Name = "Col_Import_PrimaryPdO"
        Me.Col_Import_PrimaryPdO.Visible = False
        Me.Col_Import_PrimaryPdO.Width = 90
        '
        'chkShowErrorOnly
        '
        Me.chkShowErrorOnly.AutoSize = True
        Me.chkShowErrorOnly.Location = New System.Drawing.Point(35, 45)
        Me.chkShowErrorOnly.Name = "chkShowErrorOnly"
        Me.chkShowErrorOnly.Size = New System.Drawing.Size(102, 17)
        Me.chkShowErrorOnly.TabIndex = 5
        Me.chkShowErrorOnly.Text = "Show Error Only"
        Me.chkShowErrorOnly.UseVisualStyleBackColor = True
        '
        'txtFilename
        '
        Me.txtFilename.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFilename.Location = New System.Drawing.Point(6, 19)
        Me.txtFilename.Name = "txtFilename"
        Me.txtFilename.ReadOnly = True
        Me.txtFilename.Size = New System.Drawing.Size(862, 20)
        Me.txtFilename.TabIndex = 6
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(1036, 510)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(1036, 19)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 4
        Me.btnSave.Text = "&Import"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnSelect
        '
        Me.btnSelect.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSelect.Location = New System.Drawing.Point(874, 19)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(75, 23)
        Me.btnSelect.TabIndex = 2
        Me.btnSelect.Text = "Selec&t"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'btnLoad
        '
        Me.btnLoad.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLoad.Location = New System.Drawing.Point(955, 19)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 3
        Me.btnLoad.Text = "&Load"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvItemProductGroup)
        Me.TabPage1.Controls.Add(Me.TextBox1)
        Me.TabPage1.Controls.Add(Me.btnCancel)
        Me.TabPage1.Controls.Add(Me.TxtCoverMinus_LimitData)
        Me.TabPage1.Controls.Add(Me.btnOK)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.dtpCoverMinus_RefDate)
        Me.TabPage1.Controls.Add(Me.TxtQty)
        Me.TabPage1.Controls.Add(Me.chkCoverMinus)
        Me.TabPage1.Controls.Add(Me.TxtRef)
        Me.TabPage1.Controls.Add(Me.ChkMirror)
        Me.TabPage1.Controls.Add(Me.chkImportFromDB)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1123, 545)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "AddByGroup"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(35, 45)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(102, 17)
        Me.CheckBox2.TabIndex = 8
        Me.CheckBox2.Text = "Show Error Only"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5, Me.DataGridViewTextBoxColumn6, Me.DataGridViewTextBoxColumn7, Me.DataGridViewTextBoxColumn8})
        Me.DataGridView1.Location = New System.Drawing.Point(6, 68)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(778, 392)
        Me.DataGridView1.TabIndex = 7
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "No"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.Width = 50
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "Status"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "Code"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.HeaderText = "Description"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.HeaderText = "Qty"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.Width = 75
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.HeaderText = "Unit Cost"
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        '
        'DataGridViewTextBoxColumn7
        '
        Me.DataGridViewTextBoxColumn7.HeaderText = "Amount"
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        '
        'DataGridViewTextBoxColumn8
        '
        Me.DataGridViewTextBoxColumn8.HeaderText = "Remarks"
        Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
        '
        'TextBox2
        '
        Me.TextBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox2.Location = New System.Drawing.Point(6, 19)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ReadOnly = True
        Me.TextBox2.Size = New System.Drawing.Size(535, 20)
        Me.TextBox2.TabIndex = 6
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(709, 466)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "&Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(709, 19)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "&Save"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button3.Location = New System.Drawing.Point(547, 19)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Selec&t"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button4.Location = New System.Drawing.Point(628, 19)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 23)
        Me.Button4.TabIndex = 3
        Me.Button4.Text = "&Load"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'FrmTransaksiStockInAdj_QuickInput
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1163, 603)
        Me.Controls.Add(Me.tabAddByGroup)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(576, 400)
        Me.Name = "FrmTransaksiStockInAdj_QuickInput"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Please choose scope"
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabAddByGroup.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.dgvImport, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtQty As System.Windows.Forms.TextBox
    Friend WithEvents TxtRef As System.Windows.Forms.TextBox
    Friend WithEvents chkImportFromDB As System.Windows.Forms.CheckBox
    Friend WithEvents ChkMirror As System.Windows.Forms.CheckBox
    Friend WithEvents chkCoverMinus As System.Windows.Forms.CheckBox
    Friend WithEvents dtpCoverMinus_RefDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtCoverMinus_LimitData As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents tabAddByGroup As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents chkShowErrorOnly As System.Windows.Forms.CheckBox
    Friend WithEvents txtFilename As System.Windows.Forms.TextBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Public WithEvents dgvImport As System.Windows.Forms.DataGridView
    Friend WithEvents Col_Import_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_PONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_SuppDONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_SecondaryQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_SecondaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_Batch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_ExpiredDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_Location As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_WareHouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col_Import_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_UnitCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_Invalid As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_PrimaryPOno As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_PrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_PrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_PrimarySecondaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Import_PrimaryPdO As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
