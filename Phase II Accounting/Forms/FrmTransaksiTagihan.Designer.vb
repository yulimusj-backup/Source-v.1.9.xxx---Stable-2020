<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiTagihan
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
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInvoiceNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInvoiceDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDueDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryInvoice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnINput = New System.Windows.Forms.Button
        Me.tspBar1 = New System.Windows.Forms.ToolStripProgressBar
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.btnStopProcess = New System.Windows.Forms.Button
        Me.lblProgress = New System.Windows.Forms.Label
        Me.prgAll = New System.Windows.Forms.ProgressBar
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.TxtPrimaryCust = New System.Windows.Forms.TextBox
        Me.TxtAttn = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtSubject = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtNoTrans = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.btnFindCust = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtCustCode = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtCustName = New System.Windows.Forms.TextBox
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.llbCustAdd = New System.Windows.Forms.LinkLabel
        Me.cboAdd = New System.Windows.Forms.ComboBox
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.TxtSignedBy = New System.Windows.Forms.TextBox
        Me.TxtPosition = New System.Windows.Forms.TextBox
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.TxtCustAddIdentifier = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.gboInput = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtPrimaryInvoice = New System.Windows.Forms.TextBox
        Me.dtpInvoiceDueDate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpInvoiceDate = New System.Windows.Forms.DateTimePicker
        Me.btnFindInvoice = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtInvoiceNumber = New System.Windows.Forms.TextBox
        Me.txtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtCurrItem = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtRemarksItem = New System.Windows.Forms.TextBox
        Me.TxtAmount = New System.Windows.Forms.TextBox
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.gboInput.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(383, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColInvoiceNumber, Me.ColInvoiceDate, Me.ColDueDate, Me.ColCurrency, Me.ColAmount, Me.ColRemarks, Me.ColPrimaryInvoice})
        Me.dgvItem.Location = New System.Drawing.Point(6, 271)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(746, 163)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 3
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
        'ColInvoiceNumber
        '
        Me.ColInvoiceNumber.HeaderText = "Invoice Number"
        Me.ColInvoiceNumber.MaxInputLength = 150
        Me.ColInvoiceNumber.Name = "ColInvoiceNumber"
        Me.ColInvoiceNumber.Width = 107
        '
        'ColInvoiceDate
        '
        Me.ColInvoiceDate.HeaderText = "Invoice Date"
        Me.ColInvoiceDate.Name = "ColInvoiceDate"
        Me.ColInvoiceDate.Width = 93
        '
        'ColDueDate
        '
        Me.ColDueDate.HeaderText = "Due Date"
        Me.ColDueDate.Name = "ColDueDate"
        Me.ColDueDate.Width = 78
        '
        'ColCurrency
        '
        Me.ColCurrency.HeaderText = "Currency"
        Me.ColCurrency.MaxInputLength = 50
        Me.ColCurrency.Name = "ColCurrency"
        Me.ColCurrency.Width = 74
        '
        'ColAmount
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.MaxInputLength = 15
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.Width = 68
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.MaxInputLength = 150
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 55
        '
        'ColPrimaryInvoice
        '
        Me.ColPrimaryInvoice.HeaderText = "Primary_Invoice"
        Me.ColPrimaryInvoice.Name = "ColPrimaryInvoice"
        Me.ColPrimaryInvoice.Visible = False
        Me.ColPrimaryInvoice.Width = 107
        '
        'btnINput
        '
        Me.btnINput.Location = New System.Drawing.Point(113, 55)
        Me.btnINput.Name = "btnINput"
        Me.btnINput.Size = New System.Drawing.Size(77, 20)
        Me.btnINput.TabIndex = 1
        Me.btnINput.Text = "Input"
        Me.btnINput.UseVisualStyleBackColor = True
        '
        'tspBar1
        '
        Me.tspBar1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tspBar1.Name = "tspBar1"
        Me.tspBar1.Size = New System.Drawing.Size(100, 16)
        Me.tspBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.tspBar1.Visible = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate, Me.tspBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 515)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(782, 22)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 31
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(383, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(294, 482)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(88, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(198, 482)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(88, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(102, 482)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(88, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(582, 482)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(88, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(383, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(385, 441)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(88, 26)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(285, 441)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(88, 26)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 537)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(782, 22)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 32
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(383, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.btnStopProcess)
        Me.gboDetail.Controls.Add(Me.lblProgress)
        Me.gboDetail.Controls.Add(Me.prgAll)
        Me.gboDetail.Controls.Add(Me.SplitContainer2)
        Me.gboDetail.Controls.Add(Me.gboInput)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(12, 3)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(758, 473)
        Me.gboDetail.TabIndex = 6
        Me.gboDetail.TabStop = False
        '
        'btnStopProcess
        '
        Me.btnStopProcess.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStopProcess.Location = New System.Drawing.Point(678, 436)
        Me.btnStopProcess.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnStopProcess.Name = "btnStopProcess"
        Me.btnStopProcess.Size = New System.Drawing.Size(75, 26)
        Me.btnStopProcess.TabIndex = 47
        Me.btnStopProcess.Text = "Stop Process"
        Me.btnStopProcess.UseVisualStyleBackColor = True
        Me.btnStopProcess.Visible = False
        '
        'lblProgress
        '
        Me.lblProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(476, 456)
        Me.lblProgress.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(38, 13)
        Me.lblProgress.TabIndex = 46
        Me.lblProgress.Text = "Ready"
        Me.lblProgress.Visible = False
        '
        'prgAll
        '
        Me.prgAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgAll.Location = New System.Drawing.Point(478, 436)
        Me.prgAll.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.prgAll.Name = "prgAll"
        Me.prgAll.Size = New System.Drawing.Size(195, 19)
        Me.prgAll.TabIndex = 45
        Me.prgAll.Visible = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(6, 9)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimaryCust)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtAttn)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtSubject)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtNoTrans)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label13)
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnFindCust)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtCustCode)
        Me.SplitContainer2.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtCustName)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimary)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.llbCustAdd)
        Me.SplitContainer2.Panel2.Controls.Add(Me.cboAdd)
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRemark)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label11)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label15)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtCustAddIdentifier)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label10)
        Me.SplitContainer2.Size = New System.Drawing.Size(746, 158)
        Me.SplitContainer2.SplitterDistance = 368
        Me.SplitContainer2.TabIndex = 1
        Me.SplitContainer2.TabStop = False
        '
        'TxtPrimaryCust
        '
        Me.TxtPrimaryCust.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryCust.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryCust.Location = New System.Drawing.Point(80, 111)
        Me.TxtPrimaryCust.MaxLength = 20
        Me.TxtPrimaryCust.Name = "TxtPrimaryCust"
        Me.TxtPrimaryCust.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimaryCust.TabIndex = 34
        Me.TxtPrimaryCust.TabStop = False
        Me.TxtPrimaryCust.Text = "-1"
        Me.TxtPrimaryCust.Visible = False
        '
        'TxtAttn
        '
        Me.TxtAttn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAttn.Location = New System.Drawing.Point(116, 84)
        Me.TxtAttn.MaxLength = 50
        Me.TxtAttn.Name = "TxtAttn"
        Me.TxtAttn.Size = New System.Drawing.Size(247, 20)
        Me.TxtAttn.TabIndex = 3
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 87)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(26, 13)
        Me.Label9.TabIndex = 33
        Me.Label9.Text = "Attn"
        '
        'TxtSubject
        '
        Me.TxtSubject.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSubject.Location = New System.Drawing.Point(116, 58)
        Me.TxtSubject.MaxLength = 50
        Me.TxtSubject.Name = "TxtSubject"
        Me.TxtSubject.Size = New System.Drawing.Size(247, 20)
        Me.TxtSubject.TabIndex = 2
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 61)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(43, 13)
        Me.Label8.TabIndex = 31
        Me.Label8.Text = "Subject"
        '
        'TxtNoTrans
        '
        Me.TxtNoTrans.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNoTrans.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtNoTrans.Location = New System.Drawing.Point(116, 6)
        Me.TxtNoTrans.MaxLength = 50
        Me.TxtNoTrans.Name = "TxtNoTrans"
        Me.TxtNoTrans.Size = New System.Drawing.Size(247, 20)
        Me.TxtNoTrans.TabIndex = 0
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(6, 36)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(30, 13)
        Me.Label13.TabIndex = 29
        Me.Label13.Text = "Date"
        '
        'btnFindCust
        '
        Me.btnFindCust.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindCust.Location = New System.Drawing.Point(333, 110)
        Me.btnFindCust.Name = "btnFindCust"
        Me.btnFindCust.Size = New System.Drawing.Size(30, 20)
        Me.btnFindCust.TabIndex = 6
        Me.btnFindCust.TabStop = False
        Me.btnFindCust.Text = "..."
        Me.btnFindCust.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 13)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Ref. Number"
        '
        'TxtCustCode
        '
        Me.TxtCustCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustCode.Location = New System.Drawing.Point(116, 110)
        Me.TxtCustCode.MaxLength = 50
        Me.TxtCustCode.Name = "TxtCustCode"
        Me.TxtCustCode.Size = New System.Drawing.Size(211, 20)
        Me.TxtCustCode.TabIndex = 4
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(116, 32)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(247, 20)
        Me.dtpDate.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 113)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(79, 13)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Customer Code"
        '
        'TxtCustName
        '
        Me.TxtCustName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustName.Location = New System.Drawing.Point(116, 136)
        Me.TxtCustName.MaxLength = 75
        Me.TxtCustName.Name = "TxtCustName"
        Me.TxtCustName.ReadOnly = True
        Me.TxtCustName.Size = New System.Drawing.Size(247, 20)
        Me.TxtCustName.TabIndex = 5
        Me.TxtCustName.TabStop = False
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimary.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimary.Location = New System.Drawing.Point(80, 6)
        Me.TxtPrimary.MaxLength = 20
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimary.TabIndex = 0
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 139)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 13)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "Customer Name"
        '
        'llbCustAdd
        '
        Me.llbCustAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCustAdd.AutoSize = True
        Me.llbCustAdd.Location = New System.Drawing.Point(331, 9)
        Me.llbCustAdd.Name = "llbCustAdd"
        Me.llbCustAdd.Size = New System.Drawing.Size(34, 13)
        Me.llbCustAdd.TabIndex = 0
        Me.llbCustAdd.TabStop = True
        Me.llbCustAdd.Text = "Detail"
        '
        'cboAdd
        '
        Me.cboAdd.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboAdd.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAdd.FormattingEnabled = True
        Me.cboAdd.Location = New System.Drawing.Point(113, 5)
        Me.cboAdd.Name = "cboAdd"
        Me.cboAdd.Size = New System.Drawing.Size(212, 21)
        Me.cboAdd.TabIndex = 0
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer3.Location = New System.Drawing.Point(113, 28)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.TxtSignedBy)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtPosition)
        Me.SplitContainer3.Size = New System.Drawing.Size(259, 28)
        Me.SplitContainer3.SplitterDistance = 129
        Me.SplitContainer3.TabIndex = 1
        Me.SplitContainer3.TabStop = False
        '
        'TxtSignedBy
        '
        Me.TxtSignedBy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSignedBy.Location = New System.Drawing.Point(0, 4)
        Me.TxtSignedBy.MaxLength = 50
        Me.TxtSignedBy.Name = "TxtSignedBy"
        Me.TxtSignedBy.Size = New System.Drawing.Size(123, 20)
        Me.TxtSignedBy.TabIndex = 0
        '
        'TxtPosition
        '
        Me.TxtPosition.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPosition.Location = New System.Drawing.Point(3, 4)
        Me.TxtPosition.MaxLength = 50
        Me.TxtPosition.Name = "TxtPosition"
        Me.TxtPosition.Size = New System.Drawing.Size(120, 20)
        Me.TxtPosition.TabIndex = 0
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(113, 59)
        Me.TxtRemark.MaxLength = 250
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(254, 93)
        Me.TxtRemark.TabIndex = 2
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 62)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(49, 13)
        Me.Label11.TabIndex = 35
        Me.Label11.Text = "Remarks"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(3, 36)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(97, 13)
        Me.Label15.TabIndex = 40
        Me.Label15.Text = "Signed By/Position"
        '
        'TxtCustAddIdentifier
        '
        Me.TxtCustAddIdentifier.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustAddIdentifier.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustAddIdentifier.Location = New System.Drawing.Point(77, 7)
        Me.TxtCustAddIdentifier.MaxLength = 20
        Me.TxtCustAddIdentifier.Name = "TxtCustAddIdentifier"
        Me.TxtCustAddIdentifier.Size = New System.Drawing.Size(36, 20)
        Me.TxtCustAddIdentifier.TabIndex = 0
        Me.TxtCustAddIdentifier.TabStop = False
        Me.TxtCustAddIdentifier.Text = "-1"
        Me.TxtCustAddIdentifier.Visible = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(3, 9)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(45, 13)
        Me.Label10.TabIndex = 40
        Me.Label10.Text = "Address"
        '
        'gboInput
        '
        Me.gboInput.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboInput.Controls.Add(Me.SplitContainer1)
        Me.gboInput.Location = New System.Drawing.Point(6, 169)
        Me.gboInput.Name = "gboInput"
        Me.gboInput.Size = New System.Drawing.Size(746, 99)
        Me.gboInput.TabIndex = 2
        Me.gboInput.TabStop = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryInvoice)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpInvoiceDueDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpInvoiceDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindInvoice)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtInvoiceNumber)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtReadOnly)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtCurrItem)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label12)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemarksItem)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtAmount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnINput)
        Me.SplitContainer1.Size = New System.Drawing.Size(740, 80)
        Me.SplitContainer1.SplitterDistance = 364
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 58)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 13)
        Me.Label4.TabIndex = 42
        Me.Label4.Text = "Due Date"
        '
        'TxtPrimaryInvoice
        '
        Me.TxtPrimaryInvoice.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryInvoice.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryInvoice.Location = New System.Drawing.Point(80, 3)
        Me.TxtPrimaryInvoice.MaxLength = 20
        Me.TxtPrimaryInvoice.Name = "TxtPrimaryInvoice"
        Me.TxtPrimaryInvoice.Size = New System.Drawing.Size(29, 20)
        Me.TxtPrimaryInvoice.TabIndex = 0
        Me.TxtPrimaryInvoice.TabStop = False
        Me.TxtPrimaryInvoice.Text = "-1"
        Me.TxtPrimaryInvoice.Visible = False
        '
        'dtpInvoiceDueDate
        '
        Me.dtpInvoiceDueDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpInvoiceDueDate.CustomFormat = "dd MMM yyyy"
        Me.dtpInvoiceDueDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpInvoiceDueDate.Location = New System.Drawing.Point(116, 55)
        Me.dtpInvoiceDueDate.Name = "dtpInvoiceDueDate"
        Me.dtpInvoiceDueDate.Size = New System.Drawing.Size(243, 20)
        Me.dtpInvoiceDueDate.TabIndex = 2
        Me.dtpInvoiceDueDate.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(68, 13)
        Me.Label3.TabIndex = 37
        Me.Label3.Text = "Invoice Date"
        '
        'dtpInvoiceDate
        '
        Me.dtpInvoiceDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpInvoiceDate.CustomFormat = "dd MMM yyyy"
        Me.dtpInvoiceDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpInvoiceDate.Location = New System.Drawing.Point(116, 29)
        Me.dtpInvoiceDate.Name = "dtpInvoiceDate"
        Me.dtpInvoiceDate.Size = New System.Drawing.Size(243, 20)
        Me.dtpInvoiceDate.TabIndex = 1
        Me.dtpInvoiceDate.TabStop = False
        '
        'btnFindInvoice
        '
        Me.btnFindInvoice.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindInvoice.Location = New System.Drawing.Point(329, 3)
        Me.btnFindInvoice.Name = "btnFindInvoice"
        Me.btnFindInvoice.Size = New System.Drawing.Size(30, 20)
        Me.btnFindInvoice.TabIndex = 2
        Me.btnFindInvoice.TabStop = False
        Me.btnFindInvoice.Text = "..."
        Me.btnFindInvoice.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 6)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(82, 13)
        Me.Label6.TabIndex = 35
        Me.Label6.Text = "Invoice Number"
        '
        'TxtInvoiceNumber
        '
        Me.TxtInvoiceNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtInvoiceNumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtInvoiceNumber.Location = New System.Drawing.Point(116, 3)
        Me.TxtInvoiceNumber.MaxLength = 50
        Me.TxtInvoiceNumber.Name = "TxtInvoiceNumber"
        Me.TxtInvoiceNumber.Size = New System.Drawing.Size(207, 20)
        Me.TxtInvoiceNumber.TabIndex = 0
        '
        'txtReadOnly
        '
        Me.txtReadOnly.Location = New System.Drawing.Point(84, 29)
        Me.txtReadOnly.Name = "txtReadOnly"
        Me.txtReadOnly.ReadOnly = True
        Me.txtReadOnly.Size = New System.Drawing.Size(23, 20)
        Me.txtReadOnly.TabIndex = 211
        Me.txtReadOnly.TabStop = False
        Me.txtReadOnly.Visible = False
        '
        'TxtCurrItem
        '
        Me.TxtCurrItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCurrItem.Location = New System.Drawing.Point(113, 3)
        Me.TxtCurrItem.MaxLength = 75
        Me.TxtCurrItem.Name = "TxtCurrItem"
        Me.TxtCurrItem.ReadOnly = True
        Me.TxtCurrItem.Size = New System.Drawing.Size(65, 20)
        Me.TxtCurrItem.TabIndex = 45
        Me.TxtCurrItem.TabStop = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(3, 32)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(49, 13)
        Me.Label12.TabIndex = 44
        Me.Label12.Text = "Remarks"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 6)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(43, 13)
        Me.Label7.TabIndex = 40
        Me.Label7.Text = "Amount"
        '
        'TxtRemarksItem
        '
        Me.TxtRemarksItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarksItem.Location = New System.Drawing.Point(113, 29)
        Me.TxtRemarksItem.MaxLength = 150
        Me.TxtRemarksItem.Name = "TxtRemarksItem"
        Me.TxtRemarksItem.Size = New System.Drawing.Size(251, 20)
        Me.TxtRemarksItem.TabIndex = 0
        '
        'TxtAmount
        '
        Me.TxtAmount.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAmount.Location = New System.Drawing.Point(184, 3)
        Me.TxtAmount.MaxLength = 75
        Me.TxtAmount.Name = "TxtAmount"
        Me.TxtAmount.ReadOnly = True
        Me.TxtAmount.Size = New System.Drawing.Size(182, 20)
        Me.TxtAmount.TabIndex = 0
        Me.TxtAmount.TabStop = False
        Me.TxtAmount.Text = "0"
        Me.TxtAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.AccModule.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 441)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 5
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = Global.AccModule.My.Resources.Resources.arrow_down
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(40, 441)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 6
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(390, 482)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(88, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(486, 482)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(88, 26)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox1.Location = New System.Drawing.Point(0, 4)
        Me.TextBox1.MaxLength = 20
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(121, 20)
        Me.TextBox1.TabIndex = 0
        '
        'TextBox2
        '
        Me.TextBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox2.Location = New System.Drawing.Point(3, 4)
        Me.TextBox2.MaxLength = 20
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(114, 20)
        Me.TextBox2.TabIndex = 0
        '
        'BackgroundWorker1
        '
        '
        'FrmTransaksiTagihan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(782, 559)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnDelete)
        Me.MinimumSize = New System.Drawing.Size(798, 597)
        Me.Name = "FrmTransaksiTagihan"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Collection Letters"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.PerformLayout()
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.Panel2.PerformLayout()
        Me.SplitContainer3.ResumeLayout(False)
        Me.gboInput.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnINput As System.Windows.Forms.Button
    Friend WithEvents tspBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtNoTrans As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustCode As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtCustName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents gboInput As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpInvoiceDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnFindInvoice As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtInvoiceNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpInvoiceDueDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtAmount As System.Windows.Forms.TextBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents TxtSubject As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtAttn As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtCustAddIdentifier As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents btnFindCust As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryInvoice As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TxtRemarksItem As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtSignedBy As System.Windows.Forms.TextBox
    Friend WithEvents TxtPosition As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryCust As System.Windows.Forms.TextBox
    Friend WithEvents TxtCurrItem As System.Windows.Forms.TextBox
    Friend WithEvents txtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents cboAdd As System.Windows.Forms.ComboBox
    Friend WithEvents llbCustAdd As System.Windows.Forms.LinkLabel
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInvoiceNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInvoiceDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDueDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryInvoice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnStopProcess As System.Windows.Forms.Button
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents prgAll As System.Windows.Forms.ProgressBar
End Class
