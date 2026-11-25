<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPR
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.btnPrint = New System.Windows.Forms.Button
        Me.ChkRFP = New System.Windows.Forms.CheckBox
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.txtNoTransaksi = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtPrimaryOtorisasi = New System.Windows.Forms.TextBox
        Me.DtpTglTransaksi = New System.Windows.Forms.DateTimePicker
        Me.LblOtorisasi = New System.Windows.Forms.Label
        Me.TxtPrimaryPR = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtRequestByName = New System.Windows.Forms.TextBox
        Me.BtnFindRefParty = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtRequestBy = New System.Windows.Forms.TextBox
        Me.TxtPrimaryRefParty = New System.Windows.Forms.TextBox
        Me.llbUpdateDeliveryDate = New System.Windows.Forms.LinkLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.ColItem_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_PrimaryPR = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_PrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_DeliveryDateControl = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColItem_DeliveryDate = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColItem_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_TrackUsing = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColItem_OB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Status = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColItem_Identifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_QtyMPSpare = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.StatusStrip1.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(490, 480)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 5
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'ChkRFP
        '
        Me.ChkRFP.AutoSize = True
        Me.ChkRFP.Location = New System.Drawing.Point(4, 28)
        Me.ChkRFP.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ChkRFP.Name = "ChkRFP"
        Me.ChkRFP.Size = New System.Drawing.Size(123, 17)
        Me.ChkRFP.TabIndex = 2
        Me.ChkRFP.Text = "Ready For Purchase"
        Me.ChkRFP.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 516)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(837, 22)
        Me.StatusStrip1.TabIndex = 24
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(207, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(207, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(207, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(207, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtNoTransaksi)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryOtorisasi)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpTglTransaksi)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblOtorisasi)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkRFP)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryPR)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRequestByName)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnFindRefParty)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRequestBy)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryRefParty)
        Me.SplitContainer1.Size = New System.Drawing.Size(815, 50)
        Me.SplitContainer1.SplitterDistance = 410
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'txtNoTransaksi
        '
        Me.txtNoTransaksi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNoTransaksi.Location = New System.Drawing.Point(68, 4)
        Me.txtNoTransaksi.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtNoTransaksi.Name = "txtNoTransaksi"
        Me.txtNoTransaksi.ReadOnly = True
        Me.txtNoTransaksi.Size = New System.Drawing.Size(163, 20)
        Me.txtNoTransaksi.TabIndex = 0
        Me.txtNoTransaksi.TabStop = False
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(235, 4)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Padding = New System.Windows.Forms.Padding(0, 5, 0, 0)
        Me.Label2.Size = New System.Drawing.Size(30, 18)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Date"
        '
        'TxtPrimaryOtorisasi
        '
        Me.TxtPrimaryOtorisasi.Location = New System.Drawing.Point(68, 4)
        Me.TxtPrimaryOtorisasi.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimaryOtorisasi.Name = "TxtPrimaryOtorisasi"
        Me.TxtPrimaryOtorisasi.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryOtorisasi.TabIndex = 1
        Me.TxtPrimaryOtorisasi.Text = "-1"
        Me.TxtPrimaryOtorisasi.Visible = False
        '
        'DtpTglTransaksi
        '
        Me.DtpTglTransaksi.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTglTransaksi.CustomFormat = "dd MMM yyyy"
        Me.DtpTglTransaksi.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTglTransaksi.Location = New System.Drawing.Point(269, 4)
        Me.DtpTglTransaksi.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DtpTglTransaksi.MinDate = New Date(1950, 1, 1, 0, 0, 0, 0)
        Me.DtpTglTransaksi.Name = "DtpTglTransaksi"
        Me.DtpTglTransaksi.Size = New System.Drawing.Size(137, 20)
        Me.DtpTglTransaksi.TabIndex = 2
        '
        'LblOtorisasi
        '
        Me.LblOtorisasi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblOtorisasi.AutoSize = True
        Me.LblOtorisasi.Location = New System.Drawing.Point(129, 30)
        Me.LblOtorisasi.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblOtorisasi.Name = "LblOtorisasi"
        Me.LblOtorisasi.Size = New System.Drawing.Size(102, 13)
        Me.LblOtorisasi.TabIndex = 3
        Me.LblOtorisasi.Text = "Authorize By: USER"
        '
        'TxtPrimaryPR
        '
        Me.TxtPrimaryPR.Location = New System.Drawing.Point(79, 2)
        Me.TxtPrimaryPR.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimaryPR.Name = "TxtPrimaryPR"
        Me.TxtPrimaryPR.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryPR.TabIndex = 4
        Me.TxtPrimaryPR.Text = "-1"
        Me.TxtPrimaryPR.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 6)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "PR Number"
        '
        'TxtRequestByName
        '
        Me.TxtRequestByName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRequestByName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRequestByName.Location = New System.Drawing.Point(213, 4)
        Me.TxtRequestByName.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtRequestByName.MaxLength = 250
        Me.TxtRequestByName.Name = "TxtRequestByName"
        Me.TxtRequestByName.ReadOnly = True
        Me.TxtRequestByName.Size = New System.Drawing.Size(184, 20)
        Me.TxtRequestByName.TabIndex = 18
        Me.TxtRequestByName.TabStop = False
        '
        'BtnFindRefParty
        '
        Me.BtnFindRefParty.Location = New System.Drawing.Point(152, 2)
        Me.BtnFindRefParty.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.BtnFindRefParty.Name = "BtnFindRefParty"
        Me.BtnFindRefParty.Size = New System.Drawing.Size(56, 22)
        Me.BtnFindRefParty.TabIndex = 16
        Me.BtnFindRefParty.TabStop = False
        Me.BtnFindRefParty.Text = "Find"
        Me.BtnFindRefParty.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 7)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Request By"
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(72, 28)
        Me.TxtRemarks.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtRemarks.MaxLength = 250
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(326, 20)
        Me.TxtRemarks.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 29)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Remarks"
        '
        'TxtRequestBy
        '
        Me.TxtRequestBy.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRequestBy.Location = New System.Drawing.Point(72, 4)
        Me.TxtRequestBy.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtRequestBy.MaxLength = 250
        Me.TxtRequestBy.Name = "TxtRequestBy"
        Me.TxtRequestBy.Size = New System.Drawing.Size(77, 20)
        Me.TxtRequestBy.TabIndex = 0
        '
        'TxtPrimaryRefParty
        '
        Me.TxtPrimaryRefParty.Location = New System.Drawing.Point(91, 4)
        Me.TxtPrimaryRefParty.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimaryRefParty.Name = "TxtPrimaryRefParty"
        Me.TxtPrimaryRefParty.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryRefParty.TabIndex = 17
        Me.TxtPrimaryRefParty.Text = "-1"
        Me.TxtPrimaryRefParty.Visible = False
        '
        'llbUpdateDeliveryDate
        '
        Me.llbUpdateDeliveryDate.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbUpdateDeliveryDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUpdateDeliveryDate.AutoSize = True
        Me.llbUpdateDeliveryDate.LinkColor = System.Drawing.Color.Blue
        Me.llbUpdateDeliveryDate.Location = New System.Drawing.Point(699, 449)
        Me.llbUpdateDeliveryDate.Name = "llbUpdateDeliveryDate"
        Me.llbUpdateDeliveryDate.Size = New System.Drawing.Size(116, 13)
        Me.llbUpdateDeliveryDate.TabIndex = 5
        Me.llbUpdateDeliveryDate.TabStop = True
        Me.llbUpdateDeliveryDate.Text = "update all delivery date"
        Me.llbUpdateDeliveryDate.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(577, 480)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 6
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(318, 480)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 3
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(145, 480)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 1
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(231, 480)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 2
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(419, 442)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(101, 26)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.llbUpdateDeliveryDate)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.DgvItem)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(6, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(825, 473)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.Purchase_Module.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(7, 442)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 30
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
        Me.btnDown.Location = New System.Drawing.Point(41, 442)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 31
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'DgvItem
        '
        Me.DgvItem.AllowUserToAddRows = False
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItem_No, Me.ColItem_PrimaryPR, Me.ColItem_PrimaryProduct, Me.ColItem_CodeNo, Me.ColItem_Desc, Me.ColItem_Qty, Me.ColItem_Unit, Me.ColItem_DeliveryDateControl, Me.ColItem_DeliveryDate, Me.ColItem_Remarks, Me.ColItem_TrackUsing, Me.ColItem_OB, Me.ColItem_Status, Me.ColItem_Identifier, Me.ColItem_QtyMPSpare})
        Me.DgvItem.Location = New System.Drawing.Point(5, 63)
        Me.DgvItem.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DgvItem.MultiSelect = False
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItem.RowTemplate.Height = 24
        Me.DgvItem.Size = New System.Drawing.Size(811, 373)
        Me.DgvItem.TabIndex = 1
        '
        'ColItem_No
        '
        Me.ColItem_No.HeaderText = "No"
        Me.ColItem_No.MinimumWidth = 20
        Me.ColItem_No.Name = "ColItem_No"
        Me.ColItem_No.ReadOnly = True
        Me.ColItem_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_No.Width = 27
        '
        'ColItem_PrimaryPR
        '
        Me.ColItem_PrimaryPR.HeaderText = "ColPrimaryPR_Item"
        Me.ColItem_PrimaryPR.Name = "ColItem_PrimaryPR"
        Me.ColItem_PrimaryPR.ReadOnly = True
        Me.ColItem_PrimaryPR.Visible = False
        Me.ColItem_PrimaryPR.Width = 154
        '
        'ColItem_PrimaryProduct
        '
        Me.ColItem_PrimaryProduct.HeaderText = "PrimaryProduct"
        Me.ColItem_PrimaryProduct.Name = "ColItem_PrimaryProduct"
        Me.ColItem_PrimaryProduct.ReadOnly = True
        Me.ColItem_PrimaryProduct.Visible = False
        Me.ColItem_PrimaryProduct.Width = 130
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
        Me.ColItem_Desc.MinimumWidth = 120
        Me.ColItem_Desc.Name = "ColItem_Desc"
        Me.ColItem_Desc.ReadOnly = True
        Me.ColItem_Desc.Width = 120
        '
        'ColItem_Qty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_Qty.DefaultCellStyle = DataGridViewCellStyle3
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
        Me.ColItem_Unit.ReadOnly = True
        Me.ColItem_Unit.Width = 51
        '
        'ColItem_DeliveryDateControl
        '
        Me.ColItem_DeliveryDateControl.HeaderText = "Delivery Date Control"
        Me.ColItem_DeliveryDateControl.MinimumWidth = 50
        Me.ColItem_DeliveryDateControl.Name = "ColItem_DeliveryDateControl"
        Me.ColItem_DeliveryDateControl.Width = 113
        '
        'ColItem_DeliveryDate
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.ColItem_DeliveryDate.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColItem_DeliveryDate.HeaderText = "Delivery Date"
        Me.ColItem_DeliveryDate.MinimumWidth = 100
        Me.ColItem_DeliveryDate.Name = "ColItem_DeliveryDate"
        Me.ColItem_DeliveryDate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColItem_DeliveryDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'ColItem_Remarks
        '
        Me.ColItem_Remarks.HeaderText = "Remarks"
        Me.ColItem_Remarks.MaxInputLength = 250
        Me.ColItem_Remarks.MinimumWidth = 95
        Me.ColItem_Remarks.Name = "ColItem_Remarks"
        Me.ColItem_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_Remarks.Width = 95
        '
        'ColItem_TrackUsing
        '
        Me.ColItem_TrackUsing.HeaderText = "Track Using"
        Me.ColItem_TrackUsing.Name = "ColItem_TrackUsing"
        Me.ColItem_TrackUsing.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColItem_TrackUsing.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColItem_TrackUsing.Width = 90
        '
        'ColItem_OB
        '
        Me.ColItem_OB.HeaderText = "OB"
        Me.ColItem_OB.Name = "ColItem_OB"
        Me.ColItem_OB.ReadOnly = True
        Me.ColItem_OB.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_OB.Visible = False
        Me.ColItem_OB.Width = 34
        '
        'ColItem_Status
        '
        Me.ColItem_Status.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColItem_Status.HeaderText = "Status"
        Me.ColItem_Status.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline
        Me.ColItem_Status.LinkColor = System.Drawing.Color.Blue
        Me.ColItem_Status.MinimumWidth = 100
        Me.ColItem_Status.Name = "ColItem_Status"
        Me.ColItem_Status.TrackVisitedState = False
        Me.ColItem_Status.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'ColItem_Identifier
        '
        Me.ColItem_Identifier.HeaderText = "Identifier"
        Me.ColItem_Identifier.MaxInputLength = 250
        Me.ColItem_Identifier.Name = "ColItem_Identifier"
        Me.ColItem_Identifier.ReadOnly = True
        Me.ColItem_Identifier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_Identifier.Visible = False
        Me.ColItem_Identifier.Width = 68
        '
        'ColItem_QtyMPSpare
        '
        Me.ColItem_QtyMPSpare.HeaderText = "Qty MP Spare"
        Me.ColItem_QtyMPSpare.Name = "ColItem_QtyMPSpare"
        Me.ColItem_QtyMPSpare.Visible = False
        Me.ColItem_QtyMPSpare.Width = 121
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(304, 442)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(101, 26)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(404, 480)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 4
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReadOnly.Location = New System.Drawing.Point(733, 480)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(52, 20)
        Me.TxtReadOnly.TabIndex = 16
        Me.TxtReadOnly.Visible = False
        '
        'BackgroundWorker1
        '
        '
        'FrmTransaksiPR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(837, 538)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimumSize = New System.Drawing.Size(853, 576)
        Me.Name = "FrmTransaksiPR"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Request"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents ChkRFP As System.Windows.Forms.CheckBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtPrimaryPR As System.Windows.Forms.TextBox
    Friend WithEvents DtpTglTransaksi As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtNoTransaksi As System.Windows.Forms.TextBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryRefParty As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindRefParty As System.Windows.Forms.Button
    Friend WithEvents TxtRequestBy As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LblOtorisasi As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryOtorisasi As System.Windows.Forms.TextBox
    Friend WithEvents TxtRequestByName As System.Windows.Forms.TextBox
    Friend WithEvents llbUpdateDeliveryDate As System.Windows.Forms.LinkLabel
    Friend WithEvents ColItem_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_PrimaryPR As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_PrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_DeliveryDateControl As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColItem_DeliveryDate As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColItem_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_TrackUsing As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColItem_OB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Status As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColItem_Identifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_QtyMPSpare As System.Windows.Forms.DataGridViewTextBoxColumn
End Class