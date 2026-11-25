<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPR_Cancel
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
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.ColItem_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_PrimaryCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.TxtNoTransaksi = New System.Windows.Forms.TextBox
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.btnEdit = New System.Windows.Forms.Button
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.btnNew = New System.Windows.Forms.Button
        Me.DtpTglTransaksi = New System.Windows.Forms.DateTimePicker
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TLP2 = New System.Windows.Forms.TableLayoutPanel
        Me.Label2 = New System.Windows.Forms.Label
        Me.TLP1 = New System.Windows.Forms.TableLayoutPanel
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.TxtRequestBy = New System.Windows.Forms.TextBox
        Me.btnFindPR = New System.Windows.Forms.Button
        Me.txtNoPR = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtPrimaryRefParty = New System.Windows.Forms.TextBox
        Me.TxtPrimaryPR = New System.Windows.Forms.TextBox
        Me.txtReason = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BtnFind = New System.Windows.Forms.Button
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TLP2.SuspendLayout()
        Me.TLP1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DgvItem
        '
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItem_No, Me.ColItem_PrimaryCode, Me.ColItem_CodeNo, Me.ColItem_Desc, Me.ColItem_Qty, Me.ColItem_Unit, Me.ColItem_Remarks})
        Me.DgvItem.Location = New System.Drawing.Point(5, 92)
        Me.DgvItem.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DgvItem.MultiSelect = False
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItem.RowTemplate.Height = 24
        Me.DgvItem.Size = New System.Drawing.Size(785, 346)
        Me.DgvItem.TabIndex = 1
        '
        'ColItem_No
        '
        Me.ColItem_No.HeaderText = "No"
        Me.ColItem_No.MinimumWidth = 20
        Me.ColItem_No.Name = "ColItem_No"
        Me.ColItem_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_No.Width = 27
        '
        'ColItem_PrimaryCode
        '
        Me.ColItem_PrimaryCode.HeaderText = "PrimaryProduct"
        Me.ColItem_PrimaryCode.Name = "ColItem_PrimaryCode"
        Me.ColItem_PrimaryCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_PrimaryCode.Visible = False
        Me.ColItem_PrimaryCode.Width = 84
        '
        'ColItem_CodeNo
        '
        Me.ColItem_CodeNo.HeaderText = "Code No"
        Me.ColItem_CodeNo.MaxInputLength = 250
        Me.ColItem_CodeNo.MinimumWidth = 90
        Me.ColItem_CodeNo.Name = "ColItem_CodeNo"
        Me.ColItem_CodeNo.Width = 90
        '
        'ColItem_Desc
        '
        Me.ColItem_Desc.HeaderText = "Description"
        Me.ColItem_Desc.MaxInputLength = 250
        Me.ColItem_Desc.MinimumWidth = 100
        Me.ColItem_Desc.Name = "ColItem_Desc"
        '
        'ColItem_Qty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_Qty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColItem_Qty.HeaderText = "Qty"
        Me.ColItem_Qty.MaxInputLength = 12
        Me.ColItem_Qty.MinimumWidth = 50
        Me.ColItem_Qty.Name = "ColItem_Qty"
        Me.ColItem_Qty.Width = 50
        '
        'ColItem_Unit
        '
        Me.ColItem_Unit.HeaderText = "Unit"
        Me.ColItem_Unit.MaxInputLength = 250
        Me.ColItem_Unit.MinimumWidth = 50
        Me.ColItem_Unit.Name = "ColItem_Unit"
        Me.ColItem_Unit.Width = 51
        '
        'ColItem_Remarks
        '
        Me.ColItem_Remarks.HeaderText = "Remarks"
        Me.ColItem_Remarks.MaxInputLength = 250
        Me.ColItem_Remarks.MinimumWidth = 100
        Me.ColItem_Remarks.Name = "ColItem_Remarks"
        Me.ColItem_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(69, 49)
        Me.TxtRemarks.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtRemarks.MaxLength = 250
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(294, 20)
        Me.TxtRemarks.TabIndex = 1
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.Purchase_Module.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 444)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 2
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
        Me.btnDown.Location = New System.Drawing.Point(40, 444)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 3
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(406, 444)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(101, 26)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(291, 444)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(101, 26)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(496, 482)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 6
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'TxtNoTransaksi
        '
        Me.TxtNoTransaksi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNoTransaksi.Location = New System.Drawing.Point(2, 2)
        Me.TxtNoTransaksi.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtNoTransaksi.Name = "TxtNoTransaksi"
        Me.TxtNoTransaksi.ReadOnly = True
        Me.TxtNoTransaksi.Size = New System.Drawing.Size(175, 20)
        Me.TxtNoTransaksi.TabIndex = 0
        Me.TxtNoTransaksi.TabStop = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 516)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(831, 22)
        Me.StatusStrip1.TabIndex = 8
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(204, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(204, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(204, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(204, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(2, 51)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Request By"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(583, 482)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 7
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(410, 482)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 5
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 51)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Remarks"
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReadOnly.Location = New System.Drawing.Point(745, 482)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(52, 20)
        Me.TxtReadOnly.TabIndex = 6
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(324, 482)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 4
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(2, 4)
        Me.TxtPrimary.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimary.TabIndex = 1
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(237, 482)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 3
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'DtpTglTransaksi
        '
        Me.DtpTglTransaksi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTglTransaksi.CustomFormat = "dd MMM yyyy"
        Me.DtpTglTransaksi.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTglTransaksi.Location = New System.Drawing.Point(215, 2)
        Me.DtpTglTransaksi.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DtpTglTransaksi.Name = "DtpTglTransaksi"
        Me.DtpTglTransaksi.Size = New System.Drawing.Size(116, 20)
        Me.DtpTglTransaksi.TabIndex = 1
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.DgvItem)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(12, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(799, 475)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(5, 9)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TLP2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TLP1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindPR)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtNoPR)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryRefParty)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryPR)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimary)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.AutoScroll = True
        Me.SplitContainer1.Panel2.AutoScrollMinSize = New System.Drawing.Size(0, 71)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtReason)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Size = New System.Drawing.Size(789, 79)
        Me.SplitContainer1.SplitterDistance = 419
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
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
        Me.TLP2.Controls.Add(Me.DtpTglTransaksi, 2, 0)
        Me.TLP2.Controls.Add(Me.Label2, 1, 0)
        Me.TLP2.Location = New System.Drawing.Point(83, 0)
        Me.TLP2.Name = "TLP2"
        Me.TLP2.RowCount = 1
        Me.TLP2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TLP2.Size = New System.Drawing.Size(333, 24)
        Me.TLP2.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(181, 0)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Padding = New System.Windows.Forms.Padding(0, 5, 0, 0)
        Me.Label2.Size = New System.Drawing.Size(30, 18)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Date"
        '
        'TLP1
        '
        Me.TLP1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TLP1.ColumnCount = 3
        Me.TLP1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.50334!))
        Me.TLP1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TLP1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 73.49666!))
        Me.TLP1.Controls.Add(Me.txtDesc, 2, 0)
        Me.TLP1.Controls.Add(Me.TxtRequestBy, 0, 0)
        Me.TLP1.Location = New System.Drawing.Point(83, 49)
        Me.TLP1.Name = "TLP1"
        Me.TLP1.RowCount = 1
        Me.TLP1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TLP1.Size = New System.Drawing.Size(333, 24)
        Me.TLP1.TabIndex = 0
        '
        'txtDesc
        '
        Me.txtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDesc.Location = New System.Drawing.Point(90, 2)
        Me.txtDesc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(241, 20)
        Me.txtDesc.TabIndex = 5
        Me.txtDesc.TabStop = False
        '
        'TxtRequestBy
        '
        Me.TxtRequestBy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRequestBy.Location = New System.Drawing.Point(2, 2)
        Me.TxtRequestBy.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtRequestBy.Name = "TxtRequestBy"
        Me.TxtRequestBy.ReadOnly = True
        Me.TxtRequestBy.Size = New System.Drawing.Size(84, 20)
        Me.TxtRequestBy.TabIndex = 4
        Me.TxtRequestBy.TabStop = False
        '
        'btnFindPR
        '
        Me.btnFindPR.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindPR.Location = New System.Drawing.Point(361, 27)
        Me.btnFindPR.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFindPR.Name = "btnFindPR"
        Me.btnFindPR.Size = New System.Drawing.Size(55, 21)
        Me.btnFindPR.TabIndex = 3
        Me.btnFindPR.TabStop = False
        Me.btnFindPR.Text = "Find"
        Me.btnFindPR.UseVisualStyleBackColor = True
        '
        'txtNoPR
        '
        Me.txtNoPR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNoPR.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtNoPR.Location = New System.Drawing.Point(85, 27)
        Me.txtNoPR.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtNoPR.MaxLength = 250
        Me.txtNoPR.Name = "txtNoPR"
        Me.txtNoPR.Size = New System.Drawing.Size(272, 20)
        Me.txtNoPR.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(2, 28)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(62, 13)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "PR Number"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 6)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Document No"
        '
        'TxtPrimaryRefParty
        '
        Me.TxtPrimaryRefParty.Location = New System.Drawing.Point(2, 49)
        Me.TxtPrimaryRefParty.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimaryRefParty.Name = "TxtPrimaryRefParty"
        Me.TxtPrimaryRefParty.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryRefParty.TabIndex = 4
        Me.TxtPrimaryRefParty.TabStop = False
        Me.TxtPrimaryRefParty.Text = "-1"
        Me.TxtPrimaryRefParty.Visible = False
        '
        'TxtPrimaryPR
        '
        Me.TxtPrimaryPR.Location = New System.Drawing.Point(2, 28)
        Me.TxtPrimaryPR.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimaryPR.Name = "TxtPrimaryPR"
        Me.TxtPrimaryPR.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryPR.TabIndex = 9
        Me.TxtPrimaryPR.Text = "-1"
        Me.TxtPrimaryPR.Visible = False
        '
        'txtReason
        '
        Me.txtReason.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReason.Location = New System.Drawing.Point(69, 4)
        Me.txtReason.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtReason.MaxLength = 250
        Me.txtReason.Multiline = True
        Me.txtReason.Name = "txtReason"
        Me.txtReason.Size = New System.Drawing.Size(294, 42)
        Me.txtReason.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 6)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(44, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Reason"
        '
        'BackgroundWorker1
        '
        '
        'BtnFind
        '
        Me.BtnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnFind.Location = New System.Drawing.Point(151, 482)
        Me.BtnFind.Name = "BtnFind"
        Me.BtnFind.Size = New System.Drawing.Size(80, 26)
        Me.BtnFind.TabIndex = 2
        Me.BtnFind.Text = "&Find"
        Me.BtnFind.UseVisualStyleBackColor = True
        '
        'FrmTransaksiPR_Cancel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(831, 538)
        Me.Controls.Add(Me.BtnFind)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnNew)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimumSize = New System.Drawing.Size(847, 576)
        Me.Name = "FrmTransaksiPR_Cancel"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Request Cancellation"
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TLP2.ResumeLayout(False)
        Me.TLP2.PerformLayout()
        Me.TLP1.ResumeLayout(False)
        Me.TLP1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents TxtNoTransaksi As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents DtpTglTransaksi As System.Windows.Forms.DateTimePicker
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnFindPR As System.Windows.Forms.Button
    Friend WithEvents txtNoPR As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Private WithEvents txtReason As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryPR As System.Windows.Forms.TextBox
    Friend WithEvents TxtRequestBy As System.Windows.Forms.TextBox
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryRefParty As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents BtnFind As System.Windows.Forms.Button
    Friend WithEvents ColItem_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_PrimaryCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TLP1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TLP2 As System.Windows.Forms.TableLayoutPanel
End Class
