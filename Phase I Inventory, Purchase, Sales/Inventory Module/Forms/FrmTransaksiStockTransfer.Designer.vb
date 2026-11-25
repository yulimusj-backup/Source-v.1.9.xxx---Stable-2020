<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockTransfer
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTransaksiStockTransfer))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.btnExit = New System.Windows.Forms.Button
        Me.TxtPrimaryStockIn = New System.Windows.Forms.TextBox
        Me.TxtStockInNo = New System.Windows.Forms.TextBox
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.LblProgress = New System.Windows.Forms.Label
        Me.BtnStop = New System.Windows.Forms.Button
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.TxtTransitStatus = New System.Windows.Forms.TextBox
        Me.ChkIsReceived = New System.Windows.Forms.CheckBox
        Me.ChkIsReveivement = New System.Windows.Forms.CheckBox
        Me.chkUseTransit = New System.Windows.Forms.CheckBox
        Me.TxtPrimaryRefDoc = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.llbAdvance = New System.Windows.Forms.LinkLabel
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.gboItem = New System.Windows.Forms.GroupBox
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.cboBatchNo = New System.Windows.Forms.ComboBox
        Me.lblExpiryDate = New System.Windows.Forms.Label
        Me.lblBatchNo = New System.Windows.Forms.Label
        Me.chkLocationControlled = New System.Windows.Forms.CheckBox
        Me.chkExpiryDateControlled = New System.Windows.Forms.CheckBox
        Me.txtPrimaryCategory = New System.Windows.Forms.TextBox
        Me.TxtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.TxtCodeNo = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.DtpExpired = New System.Windows.Forms.DateTimePicker
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.BtnFindCodeNo = New System.Windows.Forms.Button
        Me.BtnFindPONo = New System.Windows.Forms.Button
        Me.lblLocation = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtQtyperBatch = New System.Windows.Forms.TextBox
        Me.txtQtyOnHand = New System.Windows.Forms.TextBox
        Me.lblUnits = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblSecQty = New System.Windows.Forms.Label
        Me.BtnInput = New System.Windows.Forms.Button
        Me.cboSecondaryUnits = New System.Windows.Forms.ComboBox
        Me.btnFindDestination = New System.Windows.Forms.Button
        Me.TxtRemarkItem = New System.Windows.Forms.TextBox
        Me.txtDestinationCode = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtPrimaryDestination = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblWarehouse = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.TxtWarehouse = New System.Windows.Forms.TextBox
        Me.txtPrimaryLocation = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.cboLocation = New System.Windows.Forms.ComboBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.lblSecondaryUnits = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtCustomerName = New System.Windows.Forms.TextBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecondaryQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecondaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatch = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiredDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWareHouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDestinationLoc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouseDestination = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySecondaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDestination = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTransitLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTransitWarehouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyReceived = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitReceived = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryTransitLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.txtPrimarySecondaryUnits = New System.Windows.Forms.TextBox
        Me.txtSecondaryQty = New System.Windows.Forms.TextBox
        Me.txtPrimaryUnits = New System.Windows.Forms.TextBox
        Me.txtPrimaryPDO = New System.Windows.Forms.TextBox
        Me.chkUseSecondaryUnit = New System.Windows.Forms.CheckBox
        Me.chkBatchNoControlled = New System.Windows.Forms.CheckBox
        Me.chkPdoControlled = New System.Windows.Forms.CheckBox
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.cms1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsm_AddByBarcode = New System.Windows.Forms.ToolStripMenuItem
        Me.tsm_AddByRefDoc = New System.Windows.Forms.ToolStripMenuItem
        Me.AddItemByCustom = New System.Windows.Forms.ToolStripMenuItem
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.gboItem.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.cms1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(133, 5)
        Me.TxtRemark.MaxLength = 150
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(294, 44)
        Me.TxtRemark.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(20, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 65
        Me.Label6.Text = "Remarks"
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(277, 5)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(153, 20)
        Me.dtpDate.TabIndex = 1
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(621, 398)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'TxtPrimaryStockIn
        '
        Me.TxtPrimaryStockIn.Location = New System.Drawing.Point(113, -2)
        Me.TxtPrimaryStockIn.Name = "TxtPrimaryStockIn"
        Me.TxtPrimaryStockIn.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryStockIn.TabIndex = 47
        Me.TxtPrimaryStockIn.TabStop = False
        Me.TxtPrimaryStockIn.Visible = False
        '
        'TxtStockInNo
        '
        Me.TxtStockInNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtStockInNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtStockInNo.Location = New System.Drawing.Point(110, 5)
        Me.TxtStockInNo.Name = "TxtStockInNo"
        Me.TxtStockInNo.ReadOnly = True
        Me.TxtStockInNo.Size = New System.Drawing.Size(161, 20)
        Me.TxtStockInNo.TabIndex = 0
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.LblProgress)
        Me.gboDetail.Controls.Add(Me.BtnStop)
        Me.gboDetail.Controls.Add(Me.ProgressBar1)
        Me.gboDetail.Controls.Add(Me.SplitContainer2)
        Me.gboDetail.Controls.Add(Me.llbAdvance)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.gboItem)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtCustomerName)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.txtPrimarySecondaryUnits)
        Me.gboDetail.Controls.Add(Me.txtSecondaryQty)
        Me.gboDetail.Controls.Add(Me.txtPrimaryUnits)
        Me.gboDetail.Controls.Add(Me.txtPrimaryPDO)
        Me.gboDetail.Controls.Add(Me.chkUseSecondaryUnit)
        Me.gboDetail.Controls.Add(Me.chkBatchNoControlled)
        Me.gboDetail.Controls.Add(Me.chkPdoControlled)
        Me.gboDetail.Location = New System.Drawing.Point(9, 1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(872, 391)
        Me.gboDetail.TabIndex = 6
        Me.gboDetail.TabStop = False
        '
        'LblProgress
        '
        Me.LblProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(627, 375)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(30, 13)
        Me.LblProgress.TabIndex = 99
        Me.LblProgress.Text = "Data"
        '
        'BtnStop
        '
        Me.BtnStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnStop.Location = New System.Drawing.Point(785, 355)
        Me.BtnStop.Name = "BtnStop"
        Me.BtnStop.Size = New System.Drawing.Size(75, 26)
        Me.BtnStop.TabIndex = 17
        Me.BtnStop.Text = "Stop"
        Me.BtnStop.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(628, 356)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(151, 18)
        Me.ProgressBar1.TabIndex = 97
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(1, 13)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtStockInNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtTransitStatus)
        Me.SplitContainer2.Panel1.Controls.Add(Me.ChkIsReceived)
        Me.SplitContainer2.Panel1.Controls.Add(Me.ChkIsReveivement)
        Me.SplitContainer2.Panel1.Controls.Add(Me.chkUseTransit)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimaryStockIn)
        Me.SplitContainer2.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimaryRefDoc)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label4)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRemark)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer2.Size = New System.Drawing.Size(869, 55)
        Me.SplitContainer2.SplitterDistance = 434
        Me.SplitContainer2.TabIndex = 70
        Me.SplitContainer2.TabStop = False
        '
        'TxtTransitStatus
        '
        Me.TxtTransitStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTransitStatus.Location = New System.Drawing.Point(110, 30)
        Me.TxtTransitStatus.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtTransitStatus.Name = "TxtTransitStatus"
        Me.TxtTransitStatus.ReadOnly = True
        Me.TxtTransitStatus.Size = New System.Drawing.Size(320, 20)
        Me.TxtTransitStatus.TabIndex = 69
        Me.TxtTransitStatus.TabStop = False
        '
        'ChkIsReceived
        '
        Me.ChkIsReceived.AutoSize = True
        Me.ChkIsReceived.Enabled = False
        Me.ChkIsReceived.Location = New System.Drawing.Point(94, 6)
        Me.ChkIsReceived.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ChkIsReceived.Name = "ChkIsReceived"
        Me.ChkIsReceived.Size = New System.Drawing.Size(83, 17)
        Me.ChkIsReceived.TabIndex = 70
        Me.ChkIsReceived.Text = "Is Received"
        Me.ChkIsReceived.UseVisualStyleBackColor = True
        '
        'ChkIsReveivement
        '
        Me.ChkIsReveivement.AutoSize = True
        Me.ChkIsReveivement.Enabled = False
        Me.ChkIsReveivement.Location = New System.Drawing.Point(94, 32)
        Me.ChkIsReveivement.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ChkIsReveivement.Name = "ChkIsReveivement"
        Me.ChkIsReveivement.Size = New System.Drawing.Size(100, 17)
        Me.ChkIsReveivement.TabIndex = 70
        Me.ChkIsReveivement.TabStop = False
        Me.ChkIsReveivement.Text = "Is Receivement"
        Me.ChkIsReveivement.UseVisualStyleBackColor = True
        '
        'chkUseTransit
        '
        Me.chkUseTransit.AutoSize = True
        Me.chkUseTransit.Enabled = False
        Me.chkUseTransit.Location = New System.Drawing.Point(15, 31)
        Me.chkUseTransit.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkUseTransit.Name = "chkUseTransit"
        Me.chkUseTransit.Size = New System.Drawing.Size(80, 17)
        Me.chkUseTransit.TabIndex = 2
        Me.chkUseTransit.Text = "Use Transit"
        Me.chkUseTransit.UseVisualStyleBackColor = True
        '
        'TxtPrimaryRefDoc
        '
        Me.TxtPrimaryRefDoc.Location = New System.Drawing.Point(82, 30)
        Me.TxtPrimaryRefDoc.Name = "TxtPrimaryRefDoc"
        Me.TxtPrimaryRefDoc.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryRefDoc.TabIndex = 105
        Me.TxtPrimaryRefDoc.TabStop = False
        Me.TxtPrimaryRefDoc.Text = "-1"
        Me.TxtPrimaryRefDoc.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 6)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(94, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Stock Transfer No"
        '
        'llbAdvance
        '
        Me.llbAdvance.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbAdvance.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbAdvance.AutoSize = True
        Me.llbAdvance.Location = New System.Drawing.Point(807, 69)
        Me.llbAdvance.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbAdvance.Name = "llbAdvance"
        Me.llbAdvance.Size = New System.Drawing.Size(50, 13)
        Me.llbAdvance.TabIndex = 69
        Me.llbAdvance.TabStop = True
        Me.llbAdvance.Text = "Advance"
        Me.llbAdvance.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = CType(resources.GetObject("btnUP.BackgroundImage"), System.Drawing.Image)
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 359)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 13
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = CType(resources.GetObject("btnDown.BackgroundImage"), System.Drawing.Image)
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(40, 359)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 14
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'gboItem
        '
        Me.gboItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboItem.Controls.Add(Me.SplitContainer3)
        Me.gboItem.Location = New System.Drawing.Point(6, 69)
        Me.gboItem.Name = "gboItem"
        Me.gboItem.Size = New System.Drawing.Size(860, 153)
        Me.gboItem.TabIndex = 3
        Me.gboItem.TabStop = False
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer3.Location = New System.Drawing.Point(7, 13)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.cboBatchNo)
        Me.SplitContainer3.Panel1.Controls.Add(Me.lblExpiryDate)
        Me.SplitContainer3.Panel1.Controls.Add(Me.lblBatchNo)
        Me.SplitContainer3.Panel1.Controls.Add(Me.chkLocationControlled)
        Me.SplitContainer3.Panel1.Controls.Add(Me.chkExpiryDateControlled)
        Me.SplitContainer3.Panel1.Controls.Add(Me.txtPrimaryCategory)
        Me.SplitContainer3.Panel1.Controls.Add(Me.TxtPrimaryCodeNo)
        Me.SplitContainer3.Panel1.Controls.Add(Me.TxtCodeNo)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label19)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label10)
        Me.SplitContainer3.Panel1.Controls.Add(Me.DtpExpired)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label12)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer3.Panel1.Controls.Add(Me.txtDesc)
        Me.SplitContainer3.Panel1.Controls.Add(Me.BtnFindCodeNo)
        Me.SplitContainer3.Panel1.Controls.Add(Me.BtnFindPONo)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.lblLocation)
        Me.SplitContainer3.Panel2.Controls.Add(Me.SplitContainer1)
        Me.SplitContainer3.Panel2.Controls.Add(Me.lblSecQty)
        Me.SplitContainer3.Panel2.Controls.Add(Me.BtnInput)
        Me.SplitContainer3.Panel2.Controls.Add(Me.cboSecondaryUnits)
        Me.SplitContainer3.Panel2.Controls.Add(Me.btnFindDestination)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtRemarkItem)
        Me.SplitContainer3.Panel2.Controls.Add(Me.txtDestinationCode)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label16)
        Me.SplitContainer3.Panel2.Controls.Add(Me.txtPrimaryDestination)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer3.Panel2.Controls.Add(Me.lblWarehouse)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label15)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtWarehouse)
        Me.SplitContainer3.Panel2.Controls.Add(Me.txtPrimaryLocation)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label14)
        Me.SplitContainer3.Panel2.Controls.Add(Me.cboLocation)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label13)
        Me.SplitContainer3.Panel2.Controls.Add(Me.lblSecondaryUnits)
        Me.SplitContainer3.Size = New System.Drawing.Size(845, 133)
        Me.SplitContainer3.SplitterDistance = 422
        Me.SplitContainer3.TabIndex = 71
        Me.SplitContainer3.TabStop = False
        '
        'cboBatchNo
        '
        Me.cboBatchNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboBatchNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBatchNo.FormattingEnabled = True
        Me.cboBatchNo.Location = New System.Drawing.Point(92, 56)
        Me.cboBatchNo.Name = "cboBatchNo"
        Me.cboBatchNo.Size = New System.Drawing.Size(326, 21)
        Me.cboBatchNo.TabIndex = 3
        '
        'lblExpiryDate
        '
        Me.lblExpiryDate.AutoSize = True
        Me.lblExpiryDate.Location = New System.Drawing.Point(92, 87)
        Me.lblExpiryDate.Name = "lblExpiryDate"
        Me.lblExpiryDate.Size = New System.Drawing.Size(24, 13)
        Me.lblExpiryDate.TabIndex = 97
        Me.lblExpiryDate.Text = "n/a"
        Me.lblExpiryDate.Visible = False
        '
        'lblBatchNo
        '
        Me.lblBatchNo.AutoSize = True
        Me.lblBatchNo.Location = New System.Drawing.Point(92, 61)
        Me.lblBatchNo.Name = "lblBatchNo"
        Me.lblBatchNo.Size = New System.Drawing.Size(24, 13)
        Me.lblBatchNo.TabIndex = 101
        Me.lblBatchNo.Text = "n/a"
        Me.lblBatchNo.Visible = False
        '
        'chkLocationControlled
        '
        Me.chkLocationControlled.AutoSize = True
        Me.chkLocationControlled.Location = New System.Drawing.Point(302, 108)
        Me.chkLocationControlled.Name = "chkLocationControlled"
        Me.chkLocationControlled.Size = New System.Drawing.Size(67, 17)
        Me.chkLocationControlled.TabIndex = 93
        Me.chkLocationControlled.TabStop = False
        Me.chkLocationControlled.Text = "Location"
        Me.chkLocationControlled.UseVisualStyleBackColor = True
        Me.chkLocationControlled.Visible = False
        '
        'chkExpiryDateControlled
        '
        Me.chkExpiryDateControlled.AutoSize = True
        Me.chkExpiryDateControlled.Location = New System.Drawing.Point(92, 108)
        Me.chkExpiryDateControlled.Name = "chkExpiryDateControlled"
        Me.chkExpiryDateControlled.Size = New System.Drawing.Size(80, 17)
        Me.chkExpiryDateControlled.TabIndex = 92
        Me.chkExpiryDateControlled.TabStop = False
        Me.chkExpiryDateControlled.Text = "Expiry Date"
        Me.chkExpiryDateControlled.UseVisualStyleBackColor = True
        Me.chkExpiryDateControlled.Visible = False
        '
        'txtPrimaryCategory
        '
        Me.txtPrimaryCategory.Location = New System.Drawing.Point(132, 4)
        Me.txtPrimaryCategory.Name = "txtPrimaryCategory"
        Me.txtPrimaryCategory.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryCategory.TabIndex = 95
        Me.txtPrimaryCategory.TabStop = False
        Me.txtPrimaryCategory.Visible = False
        '
        'TxtPrimaryCodeNo
        '
        Me.TxtPrimaryCodeNo.Location = New System.Drawing.Point(101, 4)
        Me.TxtPrimaryCodeNo.Name = "TxtPrimaryCodeNo"
        Me.TxtPrimaryCodeNo.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryCodeNo.TabIndex = 53
        Me.TxtPrimaryCodeNo.TabStop = False
        Me.TxtPrimaryCodeNo.Visible = False
        '
        'TxtCodeNo
        '
        Me.TxtCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeNo.Location = New System.Drawing.Point(92, 4)
        Me.TxtCodeNo.MaxLength = 50
        Me.TxtCodeNo.Name = "TxtCodeNo"
        Me.TxtCodeNo.Size = New System.Drawing.Size(288, 20)
        Me.TxtCodeNo.TabIndex = 0
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(3, 7)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(39, 13)
        Me.Label9.TabIndex = 48
        Me.Label9.Text = "PO No"
        Me.Label9.Visible = False
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(3, 61)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(52, 13)
        Me.Label19.TabIndex = 79
        Me.Label19.Text = "Batch No"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(3, 7)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(49, 13)
        Me.Label10.TabIndex = 52
        Me.Label10.Text = "Code No"
        '
        'DtpExpired
        '
        Me.DtpExpired.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpExpired.CustomFormat = "ddd, dd MMM yyyy"
        Me.DtpExpired.Enabled = False
        Me.DtpExpired.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpExpired.Location = New System.Drawing.Point(92, 84)
        Me.DtpExpired.Name = "DtpExpired"
        Me.DtpExpired.Size = New System.Drawing.Size(326, 20)
        Me.DtpExpired.TabIndex = 4
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(3, 87)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(61, 13)
        Me.Label12.TabIndex = 77
        Me.Label12.Text = "Expiry Date"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 33)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(60, 13)
        Me.Label11.TabIndex = 62
        Me.Label11.Text = "Description"
        '
        'txtDesc
        '
        Me.txtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDesc.Location = New System.Drawing.Point(92, 30)
        Me.txtDesc.MaxLength = 6
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(326, 20)
        Me.txtDesc.TabIndex = 2
        Me.txtDesc.TabStop = False
        '
        'BtnFindCodeNo
        '
        Me.BtnFindCodeNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindCodeNo.Location = New System.Drawing.Point(387, 5)
        Me.BtnFindCodeNo.Name = "BtnFindCodeNo"
        Me.BtnFindCodeNo.Size = New System.Drawing.Size(32, 20)
        Me.BtnFindCodeNo.TabIndex = 1
        Me.BtnFindCodeNo.TabStop = False
        Me.BtnFindCodeNo.Text = "...."
        Me.BtnFindCodeNo.UseVisualStyleBackColor = True
        '
        'BtnFindPONo
        '
        Me.BtnFindPONo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindPONo.Location = New System.Drawing.Point(386, 4)
        Me.BtnFindPONo.Name = "BtnFindPONo"
        Me.BtnFindPONo.Size = New System.Drawing.Size(32, 20)
        Me.BtnFindPONo.TabIndex = 1
        Me.BtnFindPONo.Text = "...."
        Me.BtnFindPONo.UseVisualStyleBackColor = True
        Me.BtnFindPONo.Visible = False
        '
        'lblLocation
        '
        Me.lblLocation.AutoSize = True
        Me.lblLocation.Location = New System.Drawing.Point(128, 7)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.Size = New System.Drawing.Size(24, 13)
        Me.lblLocation.TabIndex = 99
        Me.lblLocation.Text = "n/a"
        Me.lblLocation.Visible = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(93, 51)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtQtyperBatch)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtQtyOnHand)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblUnits)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1.Size = New System.Drawing.Size(319, 30)
        Me.SplitContainer1.SplitterDistance = 110
        Me.SplitContainer1.TabIndex = 7
        Me.SplitContainer1.TabStop = False
        '
        'TxtQtyperBatch
        '
        Me.TxtQtyperBatch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatch.Location = New System.Drawing.Point(3, 6)
        Me.TxtQtyperBatch.MaxLength = 18
        Me.TxtQtyperBatch.Name = "TxtQtyperBatch"
        Me.TxtQtyperBatch.Size = New System.Drawing.Size(102, 20)
        Me.TxtQtyperBatch.TabIndex = 0
        Me.TxtQtyperBatch.Text = "0"
        Me.TxtQtyperBatch.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtQtyOnHand
        '
        Me.txtQtyOnHand.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtQtyOnHand.Location = New System.Drawing.Point(63, 5)
        Me.txtQtyOnHand.Name = "txtQtyOnHand"
        Me.txtQtyOnHand.ReadOnly = True
        Me.txtQtyOnHand.Size = New System.Drawing.Size(58, 20)
        Me.txtQtyOnHand.TabIndex = 0
        Me.txtQtyOnHand.TabStop = False
        Me.txtQtyOnHand.Text = "0"
        Me.txtQtyOnHand.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblUnits
        '
        Me.lblUnits.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnits.AutoSize = True
        Me.lblUnits.Location = New System.Drawing.Point(127, 8)
        Me.lblUnits.Name = "lblUnits"
        Me.lblUnits.Size = New System.Drawing.Size(31, 13)
        Me.lblUnits.TabIndex = 1
        Me.lblUnits.Text = "Units"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(50, 13)
        Me.Label2.TabIndex = 105
        Me.Label2.Text = "On Hand"
        '
        'lblSecQty
        '
        Me.lblSecQty.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSecQty.AutoSize = True
        Me.lblSecQty.Location = New System.Drawing.Point(278, 112)
        Me.lblSecQty.Name = "lblSecQty"
        Me.lblSecQty.Size = New System.Drawing.Size(18, 13)
        Me.lblSecQty.TabIndex = 98
        Me.lblSecQty.Text = "@"
        Me.lblSecQty.Visible = False
        '
        'BtnInput
        '
        Me.BtnInput.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnInput.Location = New System.Drawing.Point(360, 105)
        Me.BtnInput.Name = "BtnInput"
        Me.BtnInput.Size = New System.Drawing.Size(49, 25)
        Me.BtnInput.TabIndex = 11
        Me.BtnInput.Text = "&Input"
        Me.BtnInput.UseVisualStyleBackColor = True
        '
        'cboSecondaryUnits
        '
        Me.cboSecondaryUnits.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboSecondaryUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSecondaryUnits.FormattingEnabled = True
        Me.cboSecondaryUnits.Location = New System.Drawing.Point(302, 108)
        Me.cboSecondaryUnits.Name = "cboSecondaryUnits"
        Me.cboSecondaryUnits.Size = New System.Drawing.Size(70, 21)
        Me.cboSecondaryUnits.TabIndex = 8
        Me.cboSecondaryUnits.Visible = False
        '
        'btnFindDestination
        '
        Me.btnFindDestination.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.btnFindDestination.Location = New System.Drawing.Point(377, 82)
        Me.btnFindDestination.Name = "btnFindDestination"
        Me.btnFindDestination.Size = New System.Drawing.Size(32, 21)
        Me.btnFindDestination.TabIndex = 9
        Me.btnFindDestination.TabStop = False
        Me.btnFindDestination.Text = "...."
        Me.btnFindDestination.UseVisualStyleBackColor = True
        '
        'TxtRemarkItem
        '
        Me.TxtRemarkItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarkItem.Location = New System.Drawing.Point(97, 108)
        Me.TxtRemarkItem.MaxLength = 50
        Me.TxtRemarkItem.Name = "TxtRemarkItem"
        Me.TxtRemarkItem.Size = New System.Drawing.Size(257, 20)
        Me.TxtRemarkItem.TabIndex = 10
        '
        'txtDestinationCode
        '
        Me.txtDestinationCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDestinationCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDestinationCode.Location = New System.Drawing.Point(97, 83)
        Me.txtDestinationCode.MaxLength = 50
        Me.txtDestinationCode.Name = "txtDestinationCode"
        Me.txtDestinationCode.Size = New System.Drawing.Size(274, 20)
        Me.txtDestinationCode.TabIndex = 8
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(23, 113)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(49, 13)
        Me.Label16.TabIndex = 63
        Me.Label16.Text = "Remarks"
        '
        'txtPrimaryDestination
        '
        Me.txtPrimaryDestination.Location = New System.Drawing.Point(97, 86)
        Me.txtPrimaryDestination.Name = "txtPrimaryDestination"
        Me.txtPrimaryDestination.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryDestination.TabIndex = 104
        Me.txtPrimaryDestination.TabStop = False
        Me.txtPrimaryDestination.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 88)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(20, 13)
        Me.Label1.TabIndex = 103
        Me.Label1.Text = "To"
        '
        'lblWarehouse
        '
        Me.lblWarehouse.AutoSize = True
        Me.lblWarehouse.Location = New System.Drawing.Point(128, 33)
        Me.lblWarehouse.Name = "lblWarehouse"
        Me.lblWarehouse.Size = New System.Drawing.Size(24, 13)
        Me.lblWarehouse.TabIndex = 100
        Me.lblWarehouse.Text = "n/a"
        Me.lblWarehouse.Visible = False
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(20, 7)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(30, 13)
        Me.Label15.TabIndex = 73
        Me.Label15.Text = "From"
        '
        'TxtWarehouse
        '
        Me.TxtWarehouse.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtWarehouse.Location = New System.Drawing.Point(97, 30)
        Me.TxtWarehouse.MaxLength = 50
        Me.TxtWarehouse.Name = "TxtWarehouse"
        Me.TxtWarehouse.ReadOnly = True
        Me.TxtWarehouse.Size = New System.Drawing.Size(312, 20)
        Me.TxtWarehouse.TabIndex = 6
        Me.TxtWarehouse.TabStop = False
        '
        'txtPrimaryLocation
        '
        Me.txtPrimaryLocation.Location = New System.Drawing.Point(97, 5)
        Me.txtPrimaryLocation.Name = "txtPrimaryLocation"
        Me.txtPrimaryLocation.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryLocation.TabIndex = 94
        Me.txtPrimaryLocation.TabStop = False
        Me.txtPrimaryLocation.Visible = False
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(20, 30)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(62, 13)
        Me.Label14.TabIndex = 65
        Me.Label14.Text = "Warehouse"
        '
        'cboLocation
        '
        Me.cboLocation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboLocation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLocation.FormattingEnabled = True
        Me.cboLocation.Location = New System.Drawing.Point(97, 3)
        Me.cboLocation.Name = "cboLocation"
        Me.cboLocation.Size = New System.Drawing.Size(312, 21)
        Me.cboLocation.TabIndex = 5
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(20, 61)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(23, 13)
        Me.Label13.TabIndex = 63
        Me.Label13.Text = "Qty"
        '
        'lblSecondaryUnits
        '
        Me.lblSecondaryUnits.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSecondaryUnits.AutoSize = True
        Me.lblSecondaryUnits.Location = New System.Drawing.Point(378, 118)
        Me.lblSecondaryUnits.Name = "lblSecondaryUnits"
        Me.lblSecondaryUnits.Size = New System.Drawing.Size(31, 13)
        Me.lblSecondaryUnits.TabIndex = 86
        Me.lblSecondaryUnits.Text = "Units"
        Me.lblSecondaryUnits.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(439, 355)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 16
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(329, 355)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 15
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtCustomerName
        '
        Me.TxtCustomerName.Location = New System.Drawing.Point(116, 95)
        Me.TxtCustomerName.MaxLength = 30
        Me.TxtCustomerName.Name = "TxtCustomerName"
        Me.TxtCustomerName.ReadOnly = True
        Me.TxtCustomerName.Size = New System.Drawing.Size(316, 20)
        Me.TxtCustomerName.TabIndex = 4
        Me.TxtCustomerName.TabStop = False
        Me.TxtCustomerName.Visible = False
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDescription, Me.ColQty, Me.ColUnit, Me.ColSecondaryQty, Me.ColSecondaryUnit, Me.ColBatch, Me.ColExpiredDate, Me.ColLocation, Me.ColWareHouse, Me.ColDestinationLoc, Me.ColWarehouseDestination, Me.ColPrimaryCodeNo, Me.ColPrimaryLocation, Me.ColPrimaryUnits, Me.ColPrimarySecondaryUnits, Me.ColPrimaryDestination, Me.ColTransitLocation, Me.ColTransitWarehouse, Me.ColQtyReceived, Me.ColUnitReceived, Me.ColPrimaryTransitLocation, Me.ColRemark})
        Me.dgvItem.Location = New System.Drawing.Point(6, 228)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.Size = New System.Drawing.Size(854, 123)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 12
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
        Me.ColCodeNo.Width = 74
        '
        'ColDescription
        '
        Me.ColDescription.Frozen = True
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.Width = 85
        '
        'ColQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.Width = 48
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Units"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 56
        '
        'ColSecondaryQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSecondaryQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColSecondaryQty.HeaderText = "@"
        Me.ColSecondaryQty.Name = "ColSecondaryQty"
        Me.ColSecondaryQty.Width = 43
        '
        'ColSecondaryUnit
        '
        Me.ColSecondaryUnit.HeaderText = "Units"
        Me.ColSecondaryUnit.Name = "ColSecondaryUnit"
        Me.ColSecondaryUnit.Width = 56
        '
        'ColBatch
        '
        Me.ColBatch.HeaderText = "Batch No"
        Me.ColBatch.Name = "ColBatch"
        Me.ColBatch.Width = 77
        '
        'ColExpiredDate
        '
        Me.ColExpiredDate.HeaderText = "Expiry Date"
        Me.ColExpiredDate.Name = "ColExpiredDate"
        Me.ColExpiredDate.Width = 86
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.Width = 73
        '
        'ColWareHouse
        '
        Me.ColWareHouse.HeaderText = "WareHouse"
        Me.ColWareHouse.Name = "ColWareHouse"
        Me.ColWareHouse.Width = 89
        '
        'ColDestinationLoc
        '
        Me.ColDestinationLoc.HeaderText = "Destination"
        Me.ColDestinationLoc.Name = "ColDestinationLoc"
        Me.ColDestinationLoc.Width = 85
        '
        'ColWarehouseDestination
        '
        Me.ColWarehouseDestination.HeaderText = "Destination Warehouse"
        Me.ColWarehouseDestination.Name = "ColWarehouseDestination"
        Me.ColWarehouseDestination.Width = 143
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 105
        '
        'ColPrimaryLocation
        '
        Me.ColPrimaryLocation.HeaderText = "Primary Location"
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.Visible = False
        Me.ColPrimaryLocation.Width = 110
        '
        'ColPrimaryUnits
        '
        Me.ColPrimaryUnits.HeaderText = "PrimaryUnits"
        Me.ColPrimaryUnits.Name = "ColPrimaryUnits"
        Me.ColPrimaryUnits.Visible = False
        Me.ColPrimaryUnits.Width = 90
        '
        'ColPrimarySecondaryUnits
        '
        Me.ColPrimarySecondaryUnits.HeaderText = "PrimarySecondaryUnits"
        Me.ColPrimarySecondaryUnits.Name = "ColPrimarySecondaryUnits"
        Me.ColPrimarySecondaryUnits.Visible = False
        Me.ColPrimarySecondaryUnits.Width = 141
        '
        'ColPrimaryDestination
        '
        Me.ColPrimaryDestination.HeaderText = "Primary Destination"
        Me.ColPrimaryDestination.Name = "ColPrimaryDestination"
        Me.ColPrimaryDestination.Visible = False
        Me.ColPrimaryDestination.Width = 122
        '
        'ColTransitLocation
        '
        Me.ColTransitLocation.HeaderText = "Transit Location"
        Me.ColTransitLocation.Name = "ColTransitLocation"
        Me.ColTransitLocation.ReadOnly = True
        Me.ColTransitLocation.Width = 108
        '
        'ColTransitWarehouse
        '
        Me.ColTransitWarehouse.HeaderText = "Transit Warehouse"
        Me.ColTransitWarehouse.Name = "ColTransitWarehouse"
        Me.ColTransitWarehouse.ReadOnly = True
        Me.ColTransitWarehouse.Width = 122
        '
        'ColQtyReceived
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyReceived.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColQtyReceived.HeaderText = "Qty Received"
        Me.ColQtyReceived.Name = "ColQtyReceived"
        Me.ColQtyReceived.ReadOnly = True
        Me.ColQtyReceived.Width = 97
        '
        'ColUnitReceived
        '
        Me.ColUnitReceived.HeaderText = "Unit Received"
        Me.ColUnitReceived.Name = "ColUnitReceived"
        Me.ColUnitReceived.ReadOnly = True
        '
        'ColPrimaryTransitLocation
        '
        Me.ColPrimaryTransitLocation.HeaderText = "Primary Transit Location"
        Me.ColPrimaryTransitLocation.Name = "ColPrimaryTransitLocation"
        Me.ColPrimaryTransitLocation.ReadOnly = True
        Me.ColPrimaryTransitLocation.Width = 145
        '
        'ColRemark
        '
        Me.ColRemark.HeaderText = "Remarks"
        Me.ColRemark.Name = "ColRemark"
        Me.ColRemark.Width = 74
        '
        'txtPrimarySecondaryUnits
        '
        Me.txtPrimarySecondaryUnits.Location = New System.Drawing.Point(446, 228)
        Me.txtPrimarySecondaryUnits.Name = "txtPrimarySecondaryUnits"
        Me.txtPrimarySecondaryUnits.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimarySecondaryUnits.TabIndex = 88
        Me.txtPrimarySecondaryUnits.TabStop = False
        Me.txtPrimarySecondaryUnits.Visible = False
        '
        'txtSecondaryQty
        '
        Me.txtSecondaryQty.Location = New System.Drawing.Point(405, 228)
        Me.txtSecondaryQty.MaxLength = 30
        Me.txtSecondaryQty.Name = "txtSecondaryQty"
        Me.txtSecondaryQty.Size = New System.Drawing.Size(35, 20)
        Me.txtSecondaryQty.TabIndex = 9
        Me.txtSecondaryQty.TabStop = False
        Me.txtSecondaryQty.Text = "0"
        Me.txtSecondaryQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtSecondaryQty.Visible = False
        '
        'txtPrimaryUnits
        '
        Me.txtPrimaryUnits.Location = New System.Drawing.Point(487, 228)
        Me.txtPrimaryUnits.Name = "txtPrimaryUnits"
        Me.txtPrimaryUnits.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimaryUnits.TabIndex = 87
        Me.txtPrimaryUnits.TabStop = False
        Me.txtPrimaryUnits.Visible = False
        '
        'txtPrimaryPDO
        '
        Me.txtPrimaryPDO.Location = New System.Drawing.Point(533, 228)
        Me.txtPrimaryPDO.Name = "txtPrimaryPDO"
        Me.txtPrimaryPDO.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimaryPDO.TabIndex = 96
        Me.txtPrimaryPDO.TabStop = False
        Me.txtPrimaryPDO.Visible = False
        '
        'chkUseSecondaryUnit
        '
        Me.chkUseSecondaryUnit.AutoSize = True
        Me.chkUseSecondaryUnit.Location = New System.Drawing.Point(647, 231)
        Me.chkUseSecondaryUnit.Name = "chkUseSecondaryUnit"
        Me.chkUseSecondaryUnit.Size = New System.Drawing.Size(121, 17)
        Me.chkUseSecondaryUnit.TabIndex = 89
        Me.chkUseSecondaryUnit.TabStop = False
        Me.chkUseSecondaryUnit.Text = "Use Secondary Unit"
        Me.chkUseSecondaryUnit.UseVisualStyleBackColor = True
        Me.chkUseSecondaryUnit.Visible = False
        '
        'chkBatchNoControlled
        '
        Me.chkBatchNoControlled.AutoSize = True
        Me.chkBatchNoControlled.Location = New System.Drawing.Point(574, 231)
        Me.chkBatchNoControlled.Name = "chkBatchNoControlled"
        Me.chkBatchNoControlled.Size = New System.Drawing.Size(71, 17)
        Me.chkBatchNoControlled.TabIndex = 91
        Me.chkBatchNoControlled.TabStop = False
        Me.chkBatchNoControlled.Text = "Batch No"
        Me.chkBatchNoControlled.UseVisualStyleBackColor = True
        Me.chkBatchNoControlled.Visible = False
        '
        'chkPdoControlled
        '
        Me.chkPdoControlled.AutoSize = True
        Me.chkPdoControlled.Location = New System.Drawing.Point(765, 231)
        Me.chkPdoControlled.Name = "chkPdoControlled"
        Me.chkPdoControlled.Size = New System.Drawing.Size(47, 17)
        Me.chkPdoControlled.TabIndex = 90
        Me.chkPdoControlled.TabStop = False
        Me.chkPdoControlled.Text = "PdO"
        Me.chkPdoControlled.UseVisualStyleBackColor = True
        Me.chkPdoControlled.Visible = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 431)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(892, 22)
        Me.StatusStrip1.TabIndex = 30
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
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(219, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(448, 398)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(362, 398)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(190, 398)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(535, 398)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(276, 398)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'cms1
        '
        Me.cms1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsm_AddByBarcode, Me.tsm_AddByRefDoc, Me.AddItemByCustom})
        Me.cms1.Name = "cms1"
        Me.cms1.Size = New System.Drawing.Size(227, 70)
        '
        'tsm_AddByBarcode
        '
        Me.tsm_AddByBarcode.Name = "tsm_AddByBarcode"
        Me.tsm_AddByBarcode.Size = New System.Drawing.Size(226, 22)
        Me.tsm_AddByBarcode.Text = "Add By Barcode"
        '
        'tsm_AddByRefDoc
        '
        Me.tsm_AddByRefDoc.Name = "tsm_AddByRefDoc"
        Me.tsm_AddByRefDoc.Size = New System.Drawing.Size(226, 22)
        Me.tsm_AddByRefDoc.Text = "Add By Reference Document"
        '
        'AddItemByCustom
        '
        Me.AddItemByCustom.Name = "AddItemByCustom"
        Me.AddItemByCustom.Size = New System.Drawing.Size(226, 22)
        Me.AddItemByCustom.Text = "Add Item By Custom"
        '
        'FrmTransaksiStockTransfer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(892, 453)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnNew)
        Me.MinimumSize = New System.Drawing.Size(908, 491)
        Me.Name = "FrmTransaksiStockTransfer"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stock Location Transfer"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.gboItem.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.PerformLayout()
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.Panel2.PerformLayout()
        Me.SplitContainer3.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.cms1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryStockIn As System.Windows.Forms.TextBox
    Friend WithEvents TxtStockInNo As System.Windows.Forms.TextBox
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents gboItem As System.Windows.Forms.GroupBox
    Friend WithEvents txtPrimaryDestination As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboLocation As System.Windows.Forms.ComboBox
    Friend WithEvents lblBatchNo As System.Windows.Forms.Label
    Friend WithEvents lblWarehouse As System.Windows.Forms.Label
    Friend WithEvents lblLocation As System.Windows.Forms.Label
    Friend WithEvents lblSecQty As System.Windows.Forms.Label
    Friend WithEvents lblExpiryDate As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryPDO As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCategory As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryLocation As System.Windows.Forms.TextBox
    Friend WithEvents chkLocationControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkExpiryDateControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkBatchNoControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkPdoControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseSecondaryUnit As System.Windows.Forms.CheckBox
    Friend WithEvents txtPrimarySecondaryUnits As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryUnits As System.Windows.Forms.TextBox
    Friend WithEvents txtSecondaryQty As System.Windows.Forms.TextBox
    Friend WithEvents lblUnits As System.Windows.Forms.Label
    Friend WithEvents TxtQtyperBatch As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents DtpExpired As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemarkItem As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TxtWarehouse As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents BtnFindCodeNo As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents BtnFindPONo As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents BtnInput As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents cboBatchNo As System.Windows.Forms.ComboBox
    Friend WithEvents cboSecondaryUnits As System.Windows.Forms.ComboBox
    Friend WithEvents lblSecondaryUnits As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnFindDestination As System.Windows.Forms.Button
    Friend WithEvents txtDestinationCode As System.Windows.Forms.TextBox
    Friend WithEvents txtQtyOnHand As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents llbAdvance As System.Windows.Forms.LinkLabel
    Friend WithEvents cms1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tsm_AddByBarcode As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsm_AddByRefDoc As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtTransitStatus As System.Windows.Forms.TextBox
    Friend WithEvents chkUseTransit As System.Windows.Forms.CheckBox
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents BtnStop As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecondaryQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecondaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiredDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWareHouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDestinationLoc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouseDestination As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySecondaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDestination As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTransitLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTransitWarehouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyReceived As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitReceived As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryTransitLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChkIsReveivement As System.Windows.Forms.CheckBox
    Friend WithEvents TxtPrimaryRefDoc As System.Windows.Forms.TextBox
    Friend WithEvents ChkIsReceived As System.Windows.Forms.CheckBox
    Friend WithEvents AddItemByCustom As System.Windows.Forms.ToolStripMenuItem
End Class
