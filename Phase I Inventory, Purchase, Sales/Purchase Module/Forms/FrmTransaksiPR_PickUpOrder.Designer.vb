<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPR_PickUpOrder
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
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.GboDetail2 = New System.Windows.Forms.GroupBox
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.btnFindPR = New System.Windows.Forms.Button
        Me.TxtPrimaryCode = New System.Windows.Forms.TextBox
        Me.TxtPrimaryPR = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.CboCode = New System.Windows.Forms.ComboBox
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.CboPRNo = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TLP1 = New System.Windows.Forms.TableLayoutPanel
        Me.TxtQty = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtBalance = New System.Windows.Forms.TextBox
        Me.LblUnits_Hidden = New System.Windows.Forms.Label
        Me.BtnInput = New System.Windows.Forms.Button
        Me.LblUnit = New System.Windows.Forms.Label
        Me.TxtRemarksPR = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.ColItem_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_PrimaryPR = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_PRNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_PrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label11 = New System.Windows.Forms.Label
        Me.TxtPrimaryStockInPurchase = New System.Windows.Forms.TextBox
        Me.TxtStockInPurchaseNo = New System.Windows.Forms.TextBox
        Me.TxtRequestByName = New System.Windows.Forms.TextBox
        Me.TLP2 = New System.Windows.Forms.TableLayoutPanel
        Me.TxtNoTransaksi = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.DtpTglTransaksi = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryRefParty = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.BtnFindRefParty = New System.Windows.Forms.Button
        Me.TxtRequestBy = New System.Windows.Forms.TextBox
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtPrimaryStockOutAdjust = New System.Windows.Forms.TextBox
        Me.TxtStockOutAdjustNo = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnGenerateStockOutAdjust = New System.Windows.Forms.Button
        Me.llbShowStockOut = New System.Windows.Forms.LinkLabel
        Me.llbSendEmail = New System.Windows.Forms.LinkLabel
        Me.gboDetail.SuspendLayout()
        Me.GboDetail2.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.TLP1.SuspendLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TLP2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.GboDetail2)
        Me.gboDetail.Controls.Add(Me.DgvItem)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(20, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(793, 474)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'GboDetail2
        '
        Me.GboDetail2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GboDetail2.Controls.Add(Me.SplitContainer2)
        Me.GboDetail2.Location = New System.Drawing.Point(5, 80)
        Me.GboDetail2.Margin = New System.Windows.Forms.Padding(2)
        Me.GboDetail2.Name = "GboDetail2"
        Me.GboDetail2.Padding = New System.Windows.Forms.Padding(2)
        Me.GboDetail2.Size = New System.Drawing.Size(784, 89)
        Me.GboDetail2.TabIndex = 1
        Me.GboDetail2.TabStop = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 10)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnFindPR)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimaryCode)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimaryPR)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer2.Panel1.Controls.Add(Me.CboCode)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtDesc)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer2.Panel1.Controls.Add(Me.CboPRNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label4)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.TLP1)
        Me.SplitContainer2.Panel2.Controls.Add(Me.LblUnits_Hidden)
        Me.SplitContainer2.Panel2.Controls.Add(Me.BtnInput)
        Me.SplitContainer2.Panel2.Controls.Add(Me.LblUnit)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRemarksPR)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer2.Size = New System.Drawing.Size(784, 74)
        Me.SplitContainer2.SplitterDistance = 401
        Me.SplitContainer2.SplitterWidth = 3
        Me.SplitContainer2.TabIndex = 0
        Me.SplitContainer2.TabStop = False
        '
        'btnFindPR
        '
        Me.btnFindPR.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindPR.Location = New System.Drawing.Point(347, 4)
        Me.btnFindPR.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFindPR.Name = "btnFindPR"
        Me.btnFindPR.Size = New System.Drawing.Size(47, 20)
        Me.btnFindPR.TabIndex = 3
        Me.btnFindPR.TabStop = False
        Me.btnFindPR.Text = "Find"
        Me.btnFindPR.UseVisualStyleBackColor = True
        '
        'TxtPrimaryCode
        '
        Me.TxtPrimaryCode.Location = New System.Drawing.Point(34, 32)
        Me.TxtPrimaryCode.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtPrimaryCode.Name = "TxtPrimaryCode"
        Me.TxtPrimaryCode.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryCode.TabIndex = 5
        Me.TxtPrimaryCode.TabStop = False
        Me.TxtPrimaryCode.Text = "-1"
        Me.TxtPrimaryCode.Visible = False
        '
        'TxtPrimaryPR
        '
        Me.TxtPrimaryPR.Location = New System.Drawing.Point(34, 8)
        Me.TxtPrimaryPR.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtPrimaryPR.Name = "TxtPrimaryPR"
        Me.TxtPrimaryPR.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryPR.TabIndex = 1
        Me.TxtPrimaryPR.TabStop = False
        Me.TxtPrimaryPR.Text = "-1"
        Me.TxtPrimaryPR.Visible = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(2, 55)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(60, 13)
        Me.Label9.TabIndex = 7
        Me.Label9.Text = "Description"
        '
        'CboCode
        '
        Me.CboCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboCode.FormattingEnabled = True
        Me.CboCode.Location = New System.Drawing.Point(89, 28)
        Me.CboCode.Margin = New System.Windows.Forms.Padding(2)
        Me.CboCode.Name = "CboCode"
        Me.CboCode.Size = New System.Drawing.Size(305, 21)
        Me.CboCode.TabIndex = 6
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(89, 53)
        Me.TxtDesc.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.ReadOnly = True
        Me.TxtDesc.Size = New System.Drawing.Size(305, 20)
        Me.TxtDesc.TabIndex = 8
        Me.TxtDesc.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(2, 35)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Code No"
        '
        'CboPRNo
        '
        Me.CboPRNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboPRNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPRNo.FormattingEnabled = True
        Me.CboPRNo.Location = New System.Drawing.Point(89, 4)
        Me.CboPRNo.Margin = New System.Windows.Forms.Padding(2)
        Me.CboPRNo.Name = "CboPRNo"
        Me.CboPRNo.Size = New System.Drawing.Size(254, 21)
        Me.CboPRNo.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(2, 11)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(62, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "PR Number"
        '
        'TLP1
        '
        Me.TLP1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TLP1.ColumnCount = 3
        Me.TLP1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TLP1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TLP1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TLP1.Controls.Add(Me.TxtQty, 0, 0)
        Me.TLP1.Controls.Add(Me.Label8, 1, 0)
        Me.TLP1.Controls.Add(Me.TxtBalance, 2, 0)
        Me.TLP1.Location = New System.Drawing.Point(86, 3)
        Me.TLP1.Name = "TLP1"
        Me.TLP1.RowCount = 1
        Me.TLP1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TLP1.Size = New System.Drawing.Size(295, 23)
        Me.TLP1.TabIndex = 1
        '
        'TxtQty
        '
        Me.TxtQty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQty.Location = New System.Drawing.Point(2, 2)
        Me.TxtQty.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtQty.MaxLength = 15
        Me.TxtQty.Name = "TxtQty"
        Me.TxtQty.Size = New System.Drawing.Size(118, 20)
        Me.TxtQty.TabIndex = 0
        Me.TxtQty.Text = "0"
        Me.TxtQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(124, 0)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Padding = New System.Windows.Forms.Padding(0, 5, 0, 0)
        Me.Label8.Size = New System.Drawing.Size(46, 18)
        Me.Label8.TabIndex = 1
        Me.Label8.Text = "Balance"
        '
        'TxtBalance
        '
        Me.TxtBalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBalance.Location = New System.Drawing.Point(174, 2)
        Me.TxtBalance.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtBalance.MaxLength = 250
        Me.TxtBalance.Name = "TxtBalance"
        Me.TxtBalance.ReadOnly = True
        Me.TxtBalance.Size = New System.Drawing.Size(119, 20)
        Me.TxtBalance.TabIndex = 2
        Me.TxtBalance.TabStop = False
        Me.TxtBalance.Text = "0"
        Me.TxtBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblUnits_Hidden
        '
        Me.LblUnits_Hidden.AutoSize = True
        Me.LblUnits_Hidden.Location = New System.Drawing.Point(238, 31)
        Me.LblUnits_Hidden.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblUnits_Hidden.Name = "LblUnits_Hidden"
        Me.LblUnits_Hidden.Size = New System.Drawing.Size(56, 13)
        Me.LblUnits_Hidden.TabIndex = 4
        Me.LblUnits_Hidden.Text = "Qty (Units)"
        Me.LblUnits_Hidden.Visible = False
        '
        'BtnInput
        '
        Me.BtnInput.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnInput.Location = New System.Drawing.Point(322, 28)
        Me.BtnInput.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnInput.Name = "BtnInput"
        Me.BtnInput.Size = New System.Drawing.Size(55, 22)
        Me.BtnInput.TabIndex = 5
        Me.BtnInput.Text = "Input"
        Me.BtnInput.UseVisualStyleBackColor = True
        '
        'LblUnit
        '
        Me.LblUnit.AutoSize = True
        Me.LblUnit.Location = New System.Drawing.Point(2, 6)
        Me.LblUnit.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblUnit.Name = "LblUnit"
        Me.LblUnit.Size = New System.Drawing.Size(56, 13)
        Me.LblUnit.TabIndex = 0
        Me.LblUnit.Text = "Qty (Units)"
        '
        'TxtRemarksPR
        '
        Me.TxtRemarksPR.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarksPR.Location = New System.Drawing.Point(88, 28)
        Me.TxtRemarksPR.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtRemarksPR.MaxLength = 250
        Me.TxtRemarksPR.Multiline = True
        Me.TxtRemarksPR.Name = "TxtRemarksPR"
        Me.TxtRemarksPR.Size = New System.Drawing.Size(231, 41)
        Me.TxtRemarksPR.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(2, 31)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 13)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Remarks"
        '
        'DgvItem
        '
        Me.DgvItem.AllowUserToAddRows = False
        Me.DgvItem.AllowUserToResizeColumns = False
        Me.DgvItem.AllowUserToResizeRows = False
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItem_No, Me.ColItem_PrimaryPR, Me.ColItem_PrimaryCodeNo, Me.ColItem_PRNo, Me.ColItem_CodeNo, Me.ColItem_Desc, Me.ColItem_Qty, Me.ColItem_Unit, Me.ColItem_Remarks, Me.ColItem_PrimaryUnit})
        Me.DgvItem.Location = New System.Drawing.Point(5, 173)
        Me.DgvItem.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvItem.MultiSelect = False
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItem.RowTemplate.Height = 24
        Me.DgvItem.Size = New System.Drawing.Size(780, 264)
        Me.DgvItem.TabIndex = 1
        '
        'ColItem_No
        '
        Me.ColItem_No.HeaderText = "No"
        Me.ColItem_No.Name = "ColItem_No"
        Me.ColItem_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_No.Width = 27
        '
        'ColItem_PrimaryPR
        '
        Me.ColItem_PrimaryPR.HeaderText = "Primary PR"
        Me.ColItem_PrimaryPR.Name = "ColItem_PrimaryPR"
        Me.ColItem_PrimaryPR.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_PrimaryPR.Visible = False
        Me.ColItem_PrimaryPR.Width = 85
        '
        'ColItem_PrimaryCodeNo
        '
        Me.ColItem_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColItem_PrimaryCodeNo.Name = "ColItem_PrimaryCodeNo"
        Me.ColItem_PrimaryCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_PrimaryCodeNo.Visible = False
        Me.ColItem_PrimaryCodeNo.Width = 121
        '
        'ColItem_PRNo
        '
        Me.ColItem_PRNo.HeaderText = "PR. No"
        Me.ColItem_PRNo.MaxInputLength = 250
        Me.ColItem_PRNo.MinimumWidth = 67
        Me.ColItem_PRNo.Name = "ColItem_PRNo"
        Me.ColItem_PRNo.Width = 67
        '
        'ColItem_CodeNo
        '
        Me.ColItem_CodeNo.HeaderText = "Code No"
        Me.ColItem_CodeNo.MaxInputLength = 250
        Me.ColItem_CodeNo.MinimumWidth = 74
        Me.ColItem_CodeNo.Name = "ColItem_CodeNo"
        Me.ColItem_CodeNo.Width = 74
        '
        'ColItem_Desc
        '
        Me.ColItem_Desc.HeaderText = "Description"
        Me.ColItem_Desc.MaxInputLength = 250
        Me.ColItem_Desc.MinimumWidth = 85
        Me.ColItem_Desc.Name = "ColItem_Desc"
        Me.ColItem_Desc.Width = 85
        '
        'ColItem_Qty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_Qty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColItem_Qty.HeaderText = "Qty"
        Me.ColItem_Qty.MaxInputLength = 15
        Me.ColItem_Qty.MinimumWidth = 48
        Me.ColItem_Qty.Name = "ColItem_Qty"
        Me.ColItem_Qty.Width = 48
        '
        'ColItem_Unit
        '
        Me.ColItem_Unit.HeaderText = "Unit"
        Me.ColItem_Unit.MaxInputLength = 250
        Me.ColItem_Unit.MinimumWidth = 51
        Me.ColItem_Unit.Name = "ColItem_Unit"
        Me.ColItem_Unit.Width = 51
        '
        'ColItem_Remarks
        '
        Me.ColItem_Remarks.HeaderText = "Remarks"
        Me.ColItem_Remarks.MaxInputLength = 250
        Me.ColItem_Remarks.MinimumWidth = 74
        Me.ColItem_Remarks.Name = "ColItem_Remarks"
        Me.ColItem_Remarks.Width = 74
        '
        'ColItem_PrimaryUnit
        '
        Me.ColItem_PrimaryUnit.HeaderText = "Primary Unit"
        Me.ColItem_PrimaryUnit.Name = "ColItem_PrimaryUnit"
        Me.ColItem_PrimaryUnit.Visible = False
        Me.ColItem_PrimaryUnit.Width = 110
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryStockInPurchase)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtStockInPurchaseNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRequestByName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TLP2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryRefParty)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label10)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindRefParty)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRequestBy)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryStockOutAdjust)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtStockOutAdjustNo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Size = New System.Drawing.Size(783, 68)
        Me.SplitContainer1.SplitterDistance = 400
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(2, 50)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(112, 13)
        Me.Label11.TabIndex = 6
        Me.Label11.Text = "Stock In Purchase No"
        '
        'TxtPrimaryStockInPurchase
        '
        Me.TxtPrimaryStockInPurchase.Location = New System.Drawing.Point(74, 47)
        Me.TxtPrimaryStockInPurchase.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtPrimaryStockInPurchase.Name = "TxtPrimaryStockInPurchase"
        Me.TxtPrimaryStockInPurchase.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryStockInPurchase.TabIndex = 7
        Me.TxtPrimaryStockInPurchase.TabStop = False
        Me.TxtPrimaryStockInPurchase.Text = "-1"
        Me.TxtPrimaryStockInPurchase.Visible = False
        '
        'TxtStockInPurchaseNo
        '
        Me.TxtStockInPurchaseNo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtStockInPurchaseNo.Location = New System.Drawing.Point(113, 48)
        Me.TxtStockInPurchaseNo.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtStockInPurchaseNo.MaxLength = 250
        Me.TxtStockInPurchaseNo.Name = "TxtStockInPurchaseNo"
        Me.TxtStockInPurchaseNo.ReadOnly = True
        Me.TxtStockInPurchaseNo.Size = New System.Drawing.Size(284, 20)
        Me.TxtStockInPurchaseNo.TabIndex = 8
        Me.TxtStockInPurchaseNo.TabStop = False
        '
        'TxtRequestByName
        '
        Me.TxtRequestByName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRequestByName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRequestByName.Location = New System.Drawing.Point(252, 27)
        Me.TxtRequestByName.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtRequestByName.MaxLength = 250
        Me.TxtRequestByName.Name = "TxtRequestByName"
        Me.TxtRequestByName.ReadOnly = True
        Me.TxtRequestByName.Size = New System.Drawing.Size(146, 20)
        Me.TxtRequestByName.TabIndex = 5
        Me.TxtRequestByName.TabStop = False
        '
        'TLP2
        '
        Me.TLP2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TLP2.ColumnCount = 3
        Me.TLP2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 60.0!))
        Me.TLP2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TLP2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 40.0!))
        Me.TLP2.Controls.Add(Me.TxtNoTransaksi, 0, 0)
        Me.TLP2.Controls.Add(Me.Label2, 1, 0)
        Me.TLP2.Controls.Add(Me.DtpTglTransaksi, 2, 0)
        Me.TLP2.Location = New System.Drawing.Point(111, 0)
        Me.TLP2.Name = "TLP2"
        Me.TLP2.RowCount = 1
        Me.TLP2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TLP2.Size = New System.Drawing.Size(286, 24)
        Me.TLP2.TabIndex = 1
        '
        'TxtNoTransaksi
        '
        Me.TxtNoTransaksi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNoTransaksi.Location = New System.Drawing.Point(2, 2)
        Me.TxtNoTransaksi.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtNoTransaksi.Name = "TxtNoTransaksi"
        Me.TxtNoTransaksi.ReadOnly = True
        Me.TxtNoTransaksi.Size = New System.Drawing.Size(147, 20)
        Me.TxtNoTransaksi.TabIndex = 0
        Me.TxtNoTransaksi.TabStop = False
        '
        'Label2
        '
        Me.Label2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(153, 0)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Padding = New System.Windows.Forms.Padding(0, 5, 0, 0)
        Me.Label2.Size = New System.Drawing.Size(30, 18)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Date"
        '
        'DtpTglTransaksi
        '
        Me.DtpTglTransaksi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTglTransaksi.CustomFormat = "dd MMM yyyy"
        Me.DtpTglTransaksi.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTglTransaksi.Location = New System.Drawing.Point(187, 2)
        Me.DtpTglTransaksi.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpTglTransaksi.Name = "DtpTglTransaksi"
        Me.DtpTglTransaksi.Size = New System.Drawing.Size(97, 20)
        Me.DtpTglTransaksi.TabIndex = 2
        '
        'TxtPrimaryRefParty
        '
        Me.TxtPrimaryRefParty.Location = New System.Drawing.Point(2, 27)
        Me.TxtPrimaryRefParty.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtPrimaryRefParty.Name = "TxtPrimaryRefParty"
        Me.TxtPrimaryRefParty.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryRefParty.TabIndex = 2
        Me.TxtPrimaryRefParty.TabStop = False
        Me.TxtPrimaryRefParty.Text = "-1"
        Me.TxtPrimaryRefParty.Visible = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(2, 29)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(62, 13)
        Me.Label10.TabIndex = 7
        Me.Label10.Text = "Request By"
        '
        'BtnFindRefParty
        '
        Me.BtnFindRefParty.Location = New System.Drawing.Point(190, 26)
        Me.BtnFindRefParty.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnFindRefParty.Name = "BtnFindRefParty"
        Me.BtnFindRefParty.Size = New System.Drawing.Size(57, 20)
        Me.BtnFindRefParty.TabIndex = 4
        Me.BtnFindRefParty.TabStop = False
        Me.BtnFindRefParty.Text = "Find"
        Me.BtnFindRefParty.UseVisualStyleBackColor = True
        '
        'TxtRequestBy
        '
        Me.TxtRequestBy.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRequestBy.Location = New System.Drawing.Point(113, 27)
        Me.TxtRequestBy.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtRequestBy.MaxLength = 250
        Me.TxtRequestBy.Name = "TxtRequestBy"
        Me.TxtRequestBy.Size = New System.Drawing.Size(74, 20)
        Me.TxtRequestBy.TabIndex = 3
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(2, 4)
        Me.TxtPrimary.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimary.TabIndex = 0
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 6)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Document No"
        '
        'TxtPrimaryStockOutAdjust
        '
        Me.TxtPrimaryStockOutAdjust.Location = New System.Drawing.Point(48, 1)
        Me.TxtPrimaryStockOutAdjust.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtPrimaryStockOutAdjust.Name = "TxtPrimaryStockOutAdjust"
        Me.TxtPrimaryStockOutAdjust.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryStockOutAdjust.TabIndex = 1
        Me.TxtPrimaryStockOutAdjust.TabStop = False
        Me.TxtPrimaryStockOutAdjust.Text = "-1"
        Me.TxtPrimaryStockOutAdjust.Visible = False
        '
        'TxtStockOutAdjustNo
        '
        Me.TxtStockOutAdjustNo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtStockOutAdjustNo.Location = New System.Drawing.Point(89, 2)
        Me.TxtStockOutAdjustNo.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtStockOutAdjustNo.MaxLength = 250
        Me.TxtStockOutAdjustNo.Name = "TxtStockOutAdjustNo"
        Me.TxtStockOutAdjustNo.ReadOnly = True
        Me.TxtStockOutAdjustNo.Size = New System.Drawing.Size(288, 20)
        Me.TxtStockOutAdjustNo.TabIndex = 2
        Me.TxtStockOutAdjustNo.TabStop = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(2, 3)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(72, 13)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "Stock Out No"
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(89, 27)
        Me.TxtRemarks.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtRemarks.MaxLength = 250
        Me.TxtRemarks.Multiline = True
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(288, 40)
        Me.TxtRemarks.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(2, 28)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 3
        Me.Label6.Text = "Remarks"
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.Purchase_Module.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 443)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 0
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = Global.Purchase_Module.My.Resources.Resources.arrow_down
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(40, 443)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 5
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(403, 443)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(101, 26)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(288, 443)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(101, 26)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(646, 481)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 6
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(358, 481)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReadOnly.Location = New System.Drawing.Point(739, 481)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(52, 20)
        Me.TxtReadOnly.TabIndex = 8
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(206, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(272, 481)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(184, 481)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(206, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(206, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 527)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(841, 22)
        Me.StatusStrip1.TabIndex = 1
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(206, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(443, 481)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(98, 481)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnGenerateStockOutAdjust
        '
        Me.btnGenerateStockOutAdjust.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnGenerateStockOutAdjust.Location = New System.Drawing.Point(530, 481)
        Me.btnGenerateStockOutAdjust.Name = "btnGenerateStockOutAdjust"
        Me.btnGenerateStockOutAdjust.Size = New System.Drawing.Size(111, 26)
        Me.btnGenerateStockOutAdjust.TabIndex = 5
        Me.btnGenerateStockOutAdjust.Text = "&Create Stock Out"
        Me.btnGenerateStockOutAdjust.UseVisualStyleBackColor = True
        '
        'llbShowStockOut
        '
        Me.llbShowStockOut.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbShowStockOut.AutoSize = True
        Me.llbShowStockOut.Location = New System.Drawing.Point(9, 488)
        Me.llbShowStockOut.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbShowStockOut.Name = "llbShowStockOut"
        Me.llbShowStockOut.Size = New System.Drawing.Size(85, 13)
        Me.llbShowStockOut.TabIndex = 9
        Me.llbShowStockOut.TabStop = True
        Me.llbShowStockOut.Text = "Show Stock Out"
        '
        'llbSendEmail
        '
        Me.llbSendEmail.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbSendEmail.AutoSize = True
        Me.llbSendEmail.Location = New System.Drawing.Point(744, 487)
        Me.llbSendEmail.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbSendEmail.Name = "llbSendEmail"
        Me.llbSendEmail.Size = New System.Drawing.Size(60, 13)
        Me.llbSendEmail.TabIndex = 10
        Me.llbSendEmail.TabStop = True
        Me.llbSendEmail.Text = "Send Email"
        '
        'FrmTransaksiPR_PickUpOrder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(841, 549)
        Me.Controls.Add(Me.llbSendEmail)
        Me.Controls.Add(Me.llbShowStockOut)
        Me.Controls.Add(Me.btnGenerateStockOutAdjust)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnFind)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(849, 576)
        Me.Name = "FrmTransaksiPR_PickUpOrder"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Request - Pick Up Order"
        Me.gboDetail.ResumeLayout(False)
        Me.GboDetail2.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.TLP1.ResumeLayout(False)
        Me.TLP1.PerformLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TLP2.ResumeLayout(False)
        Me.TLP2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents DtpTglTransaksi As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtNoTransaksi As System.Windows.Forms.TextBox
    Private WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents GboDetail2 As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents CboCode As System.Windows.Forms.ComboBox
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CboPRNo As System.Windows.Forms.ComboBox
    Friend WithEvents TxtPrimaryPR As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Private WithEvents TxtQty As System.Windows.Forms.TextBox
    Friend WithEvents LblUnit As System.Windows.Forms.Label
    Private WithEvents TxtRemarksPR As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents BtnInput As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryRefParty As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindRefParty As System.Windows.Forms.Button
    Friend WithEvents TxtRequestBy As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtBalance As System.Windows.Forms.TextBox
    Friend WithEvents LblUnits_Hidden As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TLP1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TLP2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TxtRequestByName As System.Windows.Forms.TextBox
    Private WithEvents TxtStockOutAdjustNo As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Private WithEvents TxtStockInPurchaseNo As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents btnGenerateStockOutAdjust As System.Windows.Forms.Button
    Friend WithEvents llbShowStockOut As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtPrimaryStockInPurchase As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryStockOutAdjust As System.Windows.Forms.TextBox
    Friend WithEvents ColItem_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_PrimaryPR As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_PRNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_PrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents llbSendEmail As System.Windows.Forms.LinkLabel
    Friend WithEvents btnFindPR As System.Windows.Forms.Button
End Class
