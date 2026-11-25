<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ManufacturingOrder
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
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle16 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle17 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle18 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle19 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle20 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle21 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle22 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.btnPrint = New System.Windows.Forms.Button
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.btnStopProcess = New System.Windows.Forms.Button
        Me.lblProgress = New System.Windows.Forms.Label
        Me.prgAll = New System.Windows.Forms.ProgressBar
        Me.llbFilter = New System.Windows.Forms.LinkLabel
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.ColItem_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_JONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Allocation = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColItem_QtyTarget = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_QtyProduce = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_DeliveryDate = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColItem_TargetFinishDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_TargetFinishDateR1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_TargetFinishDateR2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_ActualFinishDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_EditBOM = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColItem_BoM_Tools = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColItem_IsClosed = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColItem_Decimal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Identifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Urutan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_RefDocumentPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_PrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Print = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColItem_SetRequiredDate = New System.Windows.Forms.DataGridViewButtonColumn
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.ClbProcess = New System.Windows.Forms.CheckedListBox
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.ChkGoodPerc = New System.Windows.Forms.CheckBox
        Me.ChkLateWO = New System.Windows.Forms.CheckBox
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.ChkInsufficientRM = New System.Windows.Forms.CheckBox
        Me.llbAddWIP = New System.Windows.Forms.LinkLabel
        Me.rboWO_UseTableView = New System.Windows.Forms.RadioButton
        Me.rboWO_UseTreeView = New System.Windows.Forms.RadioButton
        Me.chkRM_IncludeWIP = New System.Windows.Forms.CheckBox
        Me.chkRM_ShowDetail = New System.Windows.Forms.CheckBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage_RM = New System.Windows.Forms.TabPage
        Me.DGVRM = New System.Windows.Forms.DataGridView
        Me.ColRM_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_WONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_Process = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_QtyTarget = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_QtyRequired = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRM_QtyAllocated = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_QtyAdditional = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRM_QtyWIP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_QtyReplaced = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRM_QtyUsed = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRM_QtyBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_GPTarget = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_GPActual = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_StandByLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_PickUpLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_RequiredDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_PrimaryStandByLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_PrimaryPickUpLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_IdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_IdentifierItem = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_PrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_PrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_IsWIP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DgvRMRekap = New System.Windows.Forms.DataGridView
        Me.ColRMRekap_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMRekap_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMRekap_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMRekap_QtyTarget = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMRekap_QtyRequired = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRMRekap_QtyAllocated = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMRekap_QtyAdditional = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRMRekap_QtyWIP = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRMRekap_QtyReplcaed = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRMRekap_QtyUsed = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRMRekap_QtyBalance = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRMRekap_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMRekap_IdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMRekap_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMRekap_PrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage_WO = New System.Windows.Forms.TabPage
        Me.dgvItemWO = New System.Windows.Forms.DataGridView
        Me.ColWOTbl_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_Status = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColWOTbl_Progress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_WONO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_ProcessCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_ProcessDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_WorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_ShowDetail = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColWOTbl_PrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_IdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_IdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_StartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_StartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_EndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWOTbl_EndTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TGV_WO = New AdvancedDataGridView.TreeGridView
        Me.ColWO2_Nodes = New AdvancedDataGridView.TreeGridColumn
        Me.ColWO2_Status = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColWO2_Progress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWO2_WONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWO2_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWO2_ProductDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWO2_ProcessCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWO2_ProcessDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWO2_WorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWO2_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWO2_ShowDetail = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColWO2_PrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWO2_IdentifierItem = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWO2_IdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage_WIP = New System.Windows.Forms.TabPage
        Me.DGVWIP = New System.Windows.Forms.DataGridView
        Me.ColWIP_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_WONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_Proses = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_QtyWIP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_UnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_IdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_IdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_PrimaryProses = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_PrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtPrimaryUserAuth = New System.Windows.Forms.TextBox
        Me.DtpTargetFinishDate = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.ChkScheduling = New System.Windows.Forms.CheckBox
        Me.DtpTglTransaksi = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtNoTransaksi = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.llbColumnSetting = New System.Windows.Forms.LinkLabel
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage_RM.SuspendLayout()
        CType(Me.DGVRM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DgvRMRekap, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage_WO.SuspendLayout()
        CType(Me.dgvItemWO, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TGV_WO, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage_WIP.SuspendLayout()
        CType(Me.DGVWIP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(509, 456)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 12
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(216, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(435, 417)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(101, 26)
        Me.btnCancel.TabIndex = 13
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(320, 417)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(101, 26)
        Me.btnOK.TabIndex = 12
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.Inventory_Module.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 417)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 14
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = Global.Inventory_Module.My.Resources.Resources.arrow_down
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(40, 417)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 15
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(596, 456)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 13
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(337, 456)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 10
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(164, 456)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 8
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(250, 456)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 9
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(423, 456)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 11
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(216, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(216, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(216, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.btnStopProcess)
        Me.gboDetail.Controls.Add(Me.lblProgress)
        Me.gboDetail.Controls.Add(Me.prgAll)
        Me.gboDetail.Controls.Add(Me.llbFilter)
        Me.gboDetail.Controls.Add(Me.SplitContainer2)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(12, 1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(857, 448)
        Me.gboDetail.TabIndex = 14
        Me.gboDetail.TabStop = False
        '
        'btnStopProcess
        '
        Me.btnStopProcess.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStopProcess.Location = New System.Drawing.Point(777, 417)
        Me.btnStopProcess.Margin = New System.Windows.Forms.Padding(2)
        Me.btnStopProcess.Name = "btnStopProcess"
        Me.btnStopProcess.Size = New System.Drawing.Size(75, 26)
        Me.btnStopProcess.TabIndex = 22
        Me.btnStopProcess.TabStop = False
        Me.btnStopProcess.Text = "Stop Process"
        Me.btnStopProcess.UseVisualStyleBackColor = True
        Me.btnStopProcess.Visible = False
        '
        'lblProgress
        '
        Me.lblProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(575, 431)
        Me.lblProgress.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(38, 13)
        Me.lblProgress.TabIndex = 21
        Me.lblProgress.Text = "Ready"
        Me.lblProgress.Visible = False
        '
        'prgAll
        '
        Me.prgAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgAll.Location = New System.Drawing.Point(579, 410)
        Me.prgAll.Margin = New System.Windows.Forms.Padding(2)
        Me.prgAll.Name = "prgAll"
        Me.prgAll.Size = New System.Drawing.Size(195, 19)
        Me.prgAll.TabIndex = 20
        Me.prgAll.Visible = False
        '
        'llbFilter
        '
        Me.llbFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbFilter.AutoSize = True
        Me.llbFilter.Location = New System.Drawing.Point(73, 423)
        Me.llbFilter.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbFilter.Name = "llbFilter"
        Me.llbFilter.Size = New System.Drawing.Size(74, 13)
        Me.llbFilter.TabIndex = 19
        Me.llbFilter.TabStop = True
        Me.llbFilter.Text = ">> Show Filter"
        Me.llbFilter.Visible = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(5, 68)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.DgvItem)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer2.Size = New System.Drawing.Size(847, 343)
        Me.SplitContainer2.SplitterDistance = 128
        Me.SplitContainer2.SplitterWidth = 3
        Me.SplitContainer2.TabIndex = 18
        '
        'DgvItem
        '
        Me.DgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvItem.ColumnHeadersHeight = 25
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItem_No, Me.ColItem_JONumber, Me.ColItem_CodeNo, Me.ColItem_Desc, Me.ColItem_Allocation, Me.ColItem_QtyTarget, Me.ColItem_QtyProduce, Me.ColItem_Unit, Me.ColItem_DeliveryDate, Me.ColItem_TargetFinishDate, Me.ColItem_TargetFinishDateR1, Me.ColItem_TargetFinishDateR2, Me.ColItem_ActualFinishDate, Me.ColItem_Remarks, Me.ColItem_EditBOM, Me.ColItem_BoM_Tools, Me.ColItem_IsClosed, Me.ColItem_Decimal, Me.ColItem_Identifier, Me.ColItem_Urutan, Me.ColItem_RefDocumentPrimary, Me.ColItem_PrimaryProduct, Me.ColItem_Print, Me.ColItem_SetRequiredDate})
        Me.DgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvItem.Location = New System.Drawing.Point(0, 0)
        Me.DgvItem.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvItem.MultiSelect = False
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItem.RowTemplate.Height = 24
        Me.DgvItem.Size = New System.Drawing.Size(847, 128)
        Me.DgvItem.TabIndex = 17
        '
        'ColItem_No
        '
        Me.ColItem_No.HeaderText = "No"
        Me.ColItem_No.Name = "ColItem_No"
        Me.ColItem_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_No.Width = 27
        '
        'ColItem_JONumber
        '
        Me.ColItem_JONumber.HeaderText = "Job Order No"
        Me.ColItem_JONumber.MaxInputLength = 150
        Me.ColItem_JONumber.Name = "ColItem_JONumber"
        Me.ColItem_JONumber.Width = 95
        '
        'ColItem_CodeNo
        '
        Me.ColItem_CodeNo.HeaderText = "Code No"
        Me.ColItem_CodeNo.MaxInputLength = 150
        Me.ColItem_CodeNo.Name = "ColItem_CodeNo"
        Me.ColItem_CodeNo.Width = 74
        '
        'ColItem_Desc
        '
        Me.ColItem_Desc.HeaderText = "Description"
        Me.ColItem_Desc.MaxInputLength = 150
        Me.ColItem_Desc.Name = "ColItem_Desc"
        Me.ColItem_Desc.Width = 85
        '
        'ColItem_Allocation
        '
        Me.ColItem_Allocation.HeaderText = "Allocation"
        Me.ColItem_Allocation.Name = "ColItem_Allocation"
        Me.ColItem_Allocation.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColItem_Allocation.Width = 59
        '
        'ColItem_QtyTarget
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_QtyTarget.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColItem_QtyTarget.HeaderText = "Qty Target"
        Me.ColItem_QtyTarget.MaxInputLength = 15
        Me.ColItem_QtyTarget.Name = "ColItem_QtyTarget"
        Me.ColItem_QtyTarget.Width = 82
        '
        'ColItem_QtyProduce
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_QtyProduce.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColItem_QtyProduce.HeaderText = "Qty Produce"
        Me.ColItem_QtyProduce.MaxInputLength = 15
        Me.ColItem_QtyProduce.Name = "ColItem_QtyProduce"
        Me.ColItem_QtyProduce.Width = 91
        '
        'ColItem_Unit
        '
        Me.ColItem_Unit.HeaderText = "Units"
        Me.ColItem_Unit.MaxInputLength = 150
        Me.ColItem_Unit.Name = "ColItem_Unit"
        Me.ColItem_Unit.Width = 56
        '
        'ColItem_DeliveryDate
        '
        Me.ColItem_DeliveryDate.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColItem_DeliveryDate.HeaderText = "Delivery Date"
        Me.ColItem_DeliveryDate.Name = "ColItem_DeliveryDate"
        Me.ColItem_DeliveryDate.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColItem_DeliveryDate.Width = 77
        '
        'ColItem_TargetFinishDate
        '
        Me.ColItem_TargetFinishDate.HeaderText = "Target Finish Date"
        Me.ColItem_TargetFinishDate.Name = "ColItem_TargetFinishDate"
        Me.ColItem_TargetFinishDate.Width = 119
        '
        'ColItem_TargetFinishDateR1
        '
        Me.ColItem_TargetFinishDateR1.HeaderText = "Target Finished Date Rev 1"
        Me.ColItem_TargetFinishDateR1.Name = "ColItem_TargetFinishDateR1"
        Me.ColItem_TargetFinishDateR1.Width = 163
        '
        'ColItem_TargetFinishDateR2
        '
        Me.ColItem_TargetFinishDateR2.HeaderText = "Target Finish Date Rev 2"
        Me.ColItem_TargetFinishDateR2.Name = "ColItem_TargetFinishDateR2"
        Me.ColItem_TargetFinishDateR2.Width = 151
        '
        'ColItem_ActualFinishDate
        '
        Me.ColItem_ActualFinishDate.HeaderText = "Actual Finish Date"
        Me.ColItem_ActualFinishDate.Name = "ColItem_ActualFinishDate"
        Me.ColItem_ActualFinishDate.Width = 118
        '
        'ColItem_Remarks
        '
        Me.ColItem_Remarks.HeaderText = "Remarks"
        Me.ColItem_Remarks.MaxInputLength = 250
        Me.ColItem_Remarks.Name = "ColItem_Remarks"
        Me.ColItem_Remarks.Width = 74
        '
        'ColItem_EditBOM
        '
        Me.ColItem_EditBOM.HeaderText = "Edit BOM"
        Me.ColItem_EditBOM.Name = "ColItem_EditBOM"
        Me.ColItem_EditBOM.Width = 58
        '
        'ColItem_BoM_Tools
        '
        Me.ColItem_BoM_Tools.HeaderText = "BoM Tools"
        Me.ColItem_BoM_Tools.Name = "ColItem_BoM_Tools"
        Me.ColItem_BoM_Tools.Width = 64
        '
        'ColItem_IsClosed
        '
        Me.ColItem_IsClosed.HeaderText = "Closed"
        Me.ColItem_IsClosed.Name = "ColItem_IsClosed"
        Me.ColItem_IsClosed.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColItem_IsClosed.Visible = False
        Me.ColItem_IsClosed.Width = 45
        '
        'ColItem_Decimal
        '
        Me.ColItem_Decimal.HeaderText = "Decimal Product"
        Me.ColItem_Decimal.Name = "ColItem_Decimal"
        Me.ColItem_Decimal.Visible = False
        Me.ColItem_Decimal.Width = 110
        '
        'ColItem_Identifier
        '
        Me.ColItem_Identifier.HeaderText = "Identifier"
        Me.ColItem_Identifier.Name = "ColItem_Identifier"
        Me.ColItem_Identifier.Visible = False
        Me.ColItem_Identifier.Width = 72
        '
        'ColItem_Urutan
        '
        Me.ColItem_Urutan.HeaderText = "Urutan"
        Me.ColItem_Urutan.Name = "ColItem_Urutan"
        Me.ColItem_Urutan.Visible = False
        Me.ColItem_Urutan.Width = 64
        '
        'ColItem_RefDocumentPrimary
        '
        Me.ColItem_RefDocumentPrimary.HeaderText = "Primary Ref Document"
        Me.ColItem_RefDocumentPrimary.Name = "ColItem_RefDocumentPrimary"
        Me.ColItem_RefDocumentPrimary.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColItem_RefDocumentPrimary.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_RefDocumentPrimary.Visible = False
        Me.ColItem_RefDocumentPrimary.Width = 119
        '
        'ColItem_PrimaryProduct
        '
        Me.ColItem_PrimaryProduct.HeaderText = "Primary Product"
        Me.ColItem_PrimaryProduct.Name = "ColItem_PrimaryProduct"
        Me.ColItem_PrimaryProduct.Visible = False
        Me.ColItem_PrimaryProduct.Width = 106
        '
        'ColItem_Print
        '
        Me.ColItem_Print.HeaderText = "Print"
        Me.ColItem_Print.Name = "ColItem_Print"
        Me.ColItem_Print.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColItem_Print.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColItem_Print.Width = 53
        '
        'ColItem_SetRequiredDate
        '
        Me.ColItem_SetRequiredDate.HeaderText = "RM Required Date"
        Me.ColItem_SetRequiredDate.Name = "ColItem_SetRequiredDate"
        Me.ColItem_SetRequiredDate.Width = 102
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.GroupBox1)
        Me.SplitContainer3.Panel1Collapsed = True
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.llbAddWIP)
        Me.SplitContainer3.Panel2.Controls.Add(Me.rboWO_UseTableView)
        Me.SplitContainer3.Panel2.Controls.Add(Me.rboWO_UseTreeView)
        Me.SplitContainer3.Panel2.Controls.Add(Me.chkRM_IncludeWIP)
        Me.SplitContainer3.Panel2.Controls.Add(Me.chkRM_ShowDetail)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TabControl1)
        Me.SplitContainer3.Size = New System.Drawing.Size(847, 212)
        Me.SplitContainer3.SplitterDistance = 314
        Me.SplitContainer3.SplitterWidth = 3
        Me.SplitContainer3.TabIndex = 20
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TabControl2)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(314, 100)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Data Filter"
        '
        'TabControl2
        '
        Me.TabControl2.Controls.Add(Me.TabPage3)
        Me.TabControl2.Controls.Add(Me.TabPage4)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl2.Location = New System.Drawing.Point(2, 15)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(310, 83)
        Me.TabControl2.TabIndex = 21
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.ClbProcess)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Size = New System.Drawing.Size(302, 57)
        Me.TabPage3.TabIndex = 0
        Me.TabPage3.Text = "Process"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'ClbProcess
        '
        Me.ClbProcess.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ClbProcess.FormattingEnabled = True
        Me.ClbProcess.Location = New System.Drawing.Point(2, 2)
        Me.ClbProcess.Margin = New System.Windows.Forms.Padding(2)
        Me.ClbProcess.Name = "ClbProcess"
        Me.ClbProcess.Size = New System.Drawing.Size(298, 49)
        Me.ClbProcess.TabIndex = 0
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.ChkGoodPerc)
        Me.TabPage4.Controls.Add(Me.ChkLateWO)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Size = New System.Drawing.Size(302, 57)
        Me.TabPage4.TabIndex = 1
        Me.TabPage4.Text = "WO Filter"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'ChkGoodPerc
        '
        Me.ChkGoodPerc.AutoSize = True
        Me.ChkGoodPerc.Location = New System.Drawing.Point(4, 27)
        Me.ChkGoodPerc.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkGoodPerc.Name = "ChkGoodPerc"
        Me.ChkGoodPerc.Size = New System.Drawing.Size(213, 17)
        Me.ChkGoodPerc.TabIndex = 22
        Me.ChkGoodPerc.Text = "Good Perct Actual < Good Perct Target"
        Me.ChkGoodPerc.UseVisualStyleBackColor = True
        '
        'ChkLateWO
        '
        Me.ChkLateWO.AutoSize = True
        Me.ChkLateWO.Location = New System.Drawing.Point(4, 5)
        Me.ChkLateWO.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkLateWO.Name = "ChkLateWO"
        Me.ChkLateWO.Size = New System.Drawing.Size(69, 17)
        Me.ChkLateWO.TabIndex = 21
        Me.ChkLateWO.Text = "Late WO"
        Me.ChkLateWO.UseVisualStyleBackColor = True
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.ChkInsufficientRM)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Size = New System.Drawing.Size(302, 57)
        Me.TabPage5.TabIndex = 2
        Me.TabPage5.Text = "RM Requirement Filter"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'ChkInsufficientRM
        '
        Me.ChkInsufficientRM.AutoSize = True
        Me.ChkInsufficientRM.Location = New System.Drawing.Point(4, 5)
        Me.ChkInsufficientRM.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkInsufficientRM.Name = "ChkInsufficientRM"
        Me.ChkInsufficientRM.Size = New System.Drawing.Size(97, 17)
        Me.ChkInsufficientRM.TabIndex = 23
        Me.ChkInsufficientRM.Text = "Insufficient RM"
        Me.ChkInsufficientRM.UseVisualStyleBackColor = True
        '
        'llbAddWIP
        '
        Me.llbAddWIP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbAddWIP.AutoSize = True
        Me.llbAddWIP.Location = New System.Drawing.Point(765, 2)
        Me.llbAddWIP.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbAddWIP.Name = "llbAddWIP"
        Me.llbAddWIP.Size = New System.Drawing.Size(50, 13)
        Me.llbAddWIP.TabIndex = 21
        Me.llbAddWIP.TabStop = True
        Me.llbAddWIP.Text = "Add WIP"
        Me.llbAddWIP.Visible = False
        '
        'rboWO_UseTableView
        '
        Me.rboWO_UseTableView.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rboWO_UseTableView.AutoSize = True
        Me.rboWO_UseTableView.Checked = True
        Me.rboWO_UseTableView.Location = New System.Drawing.Point(760, 0)
        Me.rboWO_UseTableView.Margin = New System.Windows.Forms.Padding(2)
        Me.rboWO_UseTableView.Name = "rboWO_UseTableView"
        Me.rboWO_UseTableView.Size = New System.Drawing.Size(78, 17)
        Me.rboWO_UseTableView.TabIndex = 22
        Me.rboWO_UseTableView.TabStop = True
        Me.rboWO_UseTableView.Text = "Table View"
        Me.rboWO_UseTableView.UseVisualStyleBackColor = True
        Me.rboWO_UseTableView.Visible = False
        '
        'rboWO_UseTreeView
        '
        Me.rboWO_UseTreeView.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rboWO_UseTreeView.AutoSize = True
        Me.rboWO_UseTreeView.Location = New System.Drawing.Point(680, 0)
        Me.rboWO_UseTreeView.Margin = New System.Windows.Forms.Padding(2)
        Me.rboWO_UseTreeView.Name = "rboWO_UseTreeView"
        Me.rboWO_UseTreeView.Size = New System.Drawing.Size(73, 17)
        Me.rboWO_UseTreeView.TabIndex = 21
        Me.rboWO_UseTreeView.Text = "Tree View"
        Me.rboWO_UseTreeView.UseVisualStyleBackColor = True
        Me.rboWO_UseTreeView.Visible = False
        '
        'chkRM_IncludeWIP
        '
        Me.chkRM_IncludeWIP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkRM_IncludeWIP.AutoSize = True
        Me.chkRM_IncludeWIP.Location = New System.Drawing.Point(677, 0)
        Me.chkRM_IncludeWIP.Margin = New System.Windows.Forms.Padding(2)
        Me.chkRM_IncludeWIP.Name = "chkRM_IncludeWIP"
        Me.chkRM_IncludeWIP.Size = New System.Drawing.Size(85, 17)
        Me.chkRM_IncludeWIP.TabIndex = 22
        Me.chkRM_IncludeWIP.Text = "Include WIP"
        Me.chkRM_IncludeWIP.UseVisualStyleBackColor = True
        '
        'chkRM_ShowDetail
        '
        Me.chkRM_ShowDetail.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkRM_ShowDetail.AutoSize = True
        Me.chkRM_ShowDetail.Location = New System.Drawing.Point(759, 0)
        Me.chkRM_ShowDetail.Margin = New System.Windows.Forms.Padding(2)
        Me.chkRM_ShowDetail.Name = "chkRM_ShowDetail"
        Me.chkRM_ShowDetail.Size = New System.Drawing.Size(83, 17)
        Me.chkRM_ShowDetail.TabIndex = 21
        Me.chkRM_ShowDetail.Text = "Show Detail"
        Me.chkRM_ShowDetail.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage_RM)
        Me.TabControl1.Controls.Add(Me.TabPage_WO)
        Me.TabControl1.Controls.Add(Me.TabPage_WIP)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(847, 212)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage_RM
        '
        Me.TabPage_RM.Controls.Add(Me.DGVRM)
        Me.TabPage_RM.Controls.Add(Me.DgvRMRekap)
        Me.TabPage_RM.Location = New System.Drawing.Point(4, 22)
        Me.TabPage_RM.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage_RM.Name = "TabPage_RM"
        Me.TabPage_RM.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage_RM.Size = New System.Drawing.Size(839, 186)
        Me.TabPage_RM.TabIndex = 0
        Me.TabPage_RM.Text = "Raw Material Requirement"
        Me.TabPage_RM.UseVisualStyleBackColor = True
        '
        'DGVRM
        '
        Me.DGVRM.AllowUserToResizeColumns = False
        Me.DGVRM.AllowUserToResizeRows = False
        Me.DGVRM.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVRM.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DGVRM.ColumnHeadersHeight = 25
        Me.DGVRM.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DGVRM.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColRM_No, Me.ColRM_WONumber, Me.ColRM_Process, Me.ColRM_CodeNo, Me.ColRM_Desc, Me.ColRM_QtyTarget, Me.ColRM_QtyRequired, Me.ColRM_QtyAllocated, Me.ColRM_QtyAdditional, Me.ColRM_QtyWIP, Me.ColRM_QtyReplaced, Me.ColRM_QtyUsed, Me.ColRM_QtyBalance, Me.ColRM_Unit, Me.ColRM_GPTarget, Me.ColRM_GPActual, Me.ColRM_StandByLocation, Me.ColRM_PickUpLocation, Me.ColRM_RequiredDate, Me.ColRM_Remarks, Me.ColRM_PrimaryStandByLocation, Me.ColRM_PrimaryPickUpLocation, Me.ColRM_IdentifierWO, Me.ColRM_IdentifierItem, Me.ColRM_PrimaryProcess, Me.ColRM_PrimaryCodeNo, Me.ColRM_PrimaryUnit, Me.ColRM_IsWIP})
        Me.DGVRM.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVRM.Location = New System.Drawing.Point(2, 2)
        Me.DGVRM.Margin = New System.Windows.Forms.Padding(2)
        Me.DGVRM.Name = "DGVRM"
        Me.DGVRM.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DGVRM.RowTemplate.Height = 24
        Me.DGVRM.Size = New System.Drawing.Size(835, 182)
        Me.DGVRM.TabIndex = 19
        '
        'ColRM_No
        '
        Me.ColRM_No.HeaderText = "No"
        Me.ColRM_No.Name = "ColRM_No"
        Me.ColRM_No.Width = 46
        '
        'ColRM_WONumber
        '
        Me.ColRM_WONumber.HeaderText = "WO Number"
        Me.ColRM_WONumber.Name = "ColRM_WONumber"
        Me.ColRM_WONumber.Width = 91
        '
        'ColRM_Process
        '
        Me.ColRM_Process.HeaderText = "Process"
        Me.ColRM_Process.Name = "ColRM_Process"
        Me.ColRM_Process.Width = 70
        '
        'ColRM_CodeNo
        '
        Me.ColRM_CodeNo.HeaderText = "Code No"
        Me.ColRM_CodeNo.Name = "ColRM_CodeNo"
        Me.ColRM_CodeNo.Width = 74
        '
        'ColRM_Desc
        '
        Me.ColRM_Desc.HeaderText = "Description"
        Me.ColRM_Desc.Name = "ColRM_Desc"
        Me.ColRM_Desc.Width = 85
        '
        'ColRM_QtyTarget
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_QtyTarget.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColRM_QtyTarget.HeaderText = "Qty Target"
        Me.ColRM_QtyTarget.Name = "ColRM_QtyTarget"
        Me.ColRM_QtyTarget.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRM_QtyTarget.Visible = False
        Me.ColRM_QtyTarget.Width = 82
        '
        'ColRM_QtyRequired
        '
        Me.ColRM_QtyRequired.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_QtyRequired.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColRM_QtyRequired.HeaderText = "Qty Required"
        Me.ColRM_QtyRequired.Name = "ColRM_QtyRequired"
        Me.ColRM_QtyRequired.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRM_QtyRequired.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRM_QtyRequired.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRM_QtyRequired.Width = 94
        '
        'ColRM_QtyAllocated
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_QtyAllocated.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColRM_QtyAllocated.HeaderText = "Qty Allocated"
        Me.ColRM_QtyAllocated.Name = "ColRM_QtyAllocated"
        Me.ColRM_QtyAllocated.Visible = False
        Me.ColRM_QtyAllocated.Width = 95
        '
        'ColRM_QtyAdditional
        '
        Me.ColRM_QtyAdditional.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_QtyAdditional.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColRM_QtyAdditional.HeaderText = "Qty Additional"
        Me.ColRM_QtyAdditional.Name = "ColRM_QtyAdditional"
        Me.ColRM_QtyAdditional.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRM_QtyAdditional.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRM_QtyAdditional.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRM_QtyAdditional.Width = 97
        '
        'ColRM_QtyWIP
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_QtyWIP.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColRM_QtyWIP.HeaderText = "Qty WIP"
        Me.ColRM_QtyWIP.Name = "ColRM_QtyWIP"
        Me.ColRM_QtyWIP.Width = 72
        '
        'ColRM_QtyReplaced
        '
        Me.ColRM_QtyReplaced.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_QtyReplaced.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColRM_QtyReplaced.HeaderText = "Qty Replaced"
        Me.ColRM_QtyReplaced.Name = "ColRM_QtyReplaced"
        Me.ColRM_QtyReplaced.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRM_QtyReplaced.Width = 78
        '
        'ColRM_QtyUsed
        '
        Me.ColRM_QtyUsed.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_QtyUsed.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColRM_QtyUsed.HeaderText = "Qty Used"
        Me.ColRM_QtyUsed.Name = "ColRM_QtyUsed"
        Me.ColRM_QtyUsed.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRM_QtyUsed.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRM_QtyUsed.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRM_QtyUsed.Width = 76
        '
        'ColRM_QtyBalance
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_QtyBalance.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColRM_QtyBalance.HeaderText = "Qty Balance"
        Me.ColRM_QtyBalance.Name = "ColRM_QtyBalance"
        Me.ColRM_QtyBalance.Width = 90
        '
        'ColRM_Unit
        '
        Me.ColRM_Unit.HeaderText = "Unit"
        Me.ColRM_Unit.Name = "ColRM_Unit"
        Me.ColRM_Unit.Width = 51
        '
        'ColRM_GPTarget
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_GPTarget.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColRM_GPTarget.HeaderText = "Good Percentage Target (%)"
        Me.ColRM_GPTarget.Name = "ColRM_GPTarget"
        Me.ColRM_GPTarget.Width = 167
        '
        'ColRM_GPActual
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_GPActual.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColRM_GPActual.HeaderText = "Good Percentage Actual (%)"
        Me.ColRM_GPActual.Name = "ColRM_GPActual"
        Me.ColRM_GPActual.Width = 166
        '
        'ColRM_StandByLocation
        '
        Me.ColRM_StandByLocation.HeaderText = "Stand By Location"
        Me.ColRM_StandByLocation.Name = "ColRM_StandByLocation"
        Me.ColRM_StandByLocation.Width = 119
        '
        'ColRM_PickUpLocation
        '
        Me.ColRM_PickUpLocation.HeaderText = "Pick Up Location"
        Me.ColRM_PickUpLocation.Name = "ColRM_PickUpLocation"
        Me.ColRM_PickUpLocation.ReadOnly = True
        Me.ColRM_PickUpLocation.Width = 114
        '
        'ColRM_RequiredDate
        '
        Me.ColRM_RequiredDate.HeaderText = "Required Date"
        Me.ColRM_RequiredDate.Name = "ColRM_RequiredDate"
        Me.ColRM_RequiredDate.Width = 101
        '
        'ColRM_Remarks
        '
        Me.ColRM_Remarks.HeaderText = "Remarks"
        Me.ColRM_Remarks.Name = "ColRM_Remarks"
        Me.ColRM_Remarks.Width = 74
        '
        'ColRM_PrimaryStandByLocation
        '
        Me.ColRM_PrimaryStandByLocation.HeaderText = "Primary Stand By Location"
        Me.ColRM_PrimaryStandByLocation.Name = "ColRM_PrimaryStandByLocation"
        Me.ColRM_PrimaryStandByLocation.Visible = False
        Me.ColRM_PrimaryStandByLocation.Width = 156
        '
        'ColRM_PrimaryPickUpLocation
        '
        Me.ColRM_PrimaryPickUpLocation.HeaderText = "Primary Pick Up Location"
        Me.ColRM_PrimaryPickUpLocation.Name = "ColRM_PrimaryPickUpLocation"
        Me.ColRM_PrimaryPickUpLocation.ReadOnly = True
        Me.ColRM_PrimaryPickUpLocation.Visible = False
        Me.ColRM_PrimaryPickUpLocation.Width = 151
        '
        'ColRM_IdentifierWO
        '
        Me.ColRM_IdentifierWO.HeaderText = "Identifier JO"
        Me.ColRM_IdentifierWO.Name = "ColRM_IdentifierWO"
        Me.ColRM_IdentifierWO.Visible = False
        Me.ColRM_IdentifierWO.Width = 88
        '
        'ColRM_IdentifierItem
        '
        Me.ColRM_IdentifierItem.HeaderText = "Identifier Item"
        Me.ColRM_IdentifierItem.Name = "ColRM_IdentifierItem"
        Me.ColRM_IdentifierItem.Visible = False
        Me.ColRM_IdentifierItem.Width = 95
        '
        'ColRM_PrimaryProcess
        '
        Me.ColRM_PrimaryProcess.HeaderText = "Primary Process"
        Me.ColRM_PrimaryProcess.Name = "ColRM_PrimaryProcess"
        Me.ColRM_PrimaryProcess.Visible = False
        Me.ColRM_PrimaryProcess.Width = 107
        '
        'ColRM_PrimaryCodeNo
        '
        Me.ColRM_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColRM_PrimaryCodeNo.Name = "ColRM_PrimaryCodeNo"
        Me.ColRM_PrimaryCodeNo.Visible = False
        Me.ColRM_PrimaryCodeNo.Width = 111
        '
        'ColRM_PrimaryUnit
        '
        Me.ColRM_PrimaryUnit.HeaderText = "Primary Unit"
        Me.ColRM_PrimaryUnit.Name = "ColRM_PrimaryUnit"
        Me.ColRM_PrimaryUnit.Visible = False
        Me.ColRM_PrimaryUnit.Width = 88
        '
        'ColRM_IsWIP
        '
        Me.ColRM_IsWIP.HeaderText = "Is WIP"
        Me.ColRM_IsWIP.Name = "ColRM_IsWIP"
        Me.ColRM_IsWIP.Visible = False
        Me.ColRM_IsWIP.Width = 64
        '
        'DgvRMRekap
        '
        Me.DgvRMRekap.AllowUserToAddRows = False
        Me.DgvRMRekap.AllowUserToDeleteRows = False
        Me.DgvRMRekap.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvRMRekap.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvRMRekap.ColumnHeadersHeight = 25
        Me.DgvRMRekap.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvRMRekap.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColRMRekap_No, Me.ColRMRekap_CodeNo, Me.ColRMRekap_Desc, Me.ColRMRekap_QtyTarget, Me.ColRMRekap_QtyRequired, Me.ColRMRekap_QtyAllocated, Me.ColRMRekap_QtyAdditional, Me.ColRMRekap_QtyWIP, Me.ColRMRekap_QtyReplcaed, Me.ColRMRekap_QtyUsed, Me.ColRMRekap_QtyBalance, Me.ColRMRekap_Unit, Me.ColRMRekap_IdentifierJO, Me.ColRMRekap_PrimaryCodeNo, Me.ColRMRekap_PrimaryUnit})
        Me.DgvRMRekap.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvRMRekap.Location = New System.Drawing.Point(2, 2)
        Me.DgvRMRekap.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvRMRekap.Name = "DgvRMRekap"
        Me.DgvRMRekap.ReadOnly = True
        Me.DgvRMRekap.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvRMRekap.RowTemplate.Height = 24
        Me.DgvRMRekap.Size = New System.Drawing.Size(835, 182)
        Me.DgvRMRekap.TabIndex = 20
        '
        'ColRMRekap_No
        '
        Me.ColRMRekap_No.HeaderText = "No"
        Me.ColRMRekap_No.Name = "ColRMRekap_No"
        Me.ColRMRekap_No.ReadOnly = True
        Me.ColRMRekap_No.Width = 46
        '
        'ColRMRekap_CodeNo
        '
        Me.ColRMRekap_CodeNo.HeaderText = "Code No"
        Me.ColRMRekap_CodeNo.Name = "ColRMRekap_CodeNo"
        Me.ColRMRekap_CodeNo.ReadOnly = True
        Me.ColRMRekap_CodeNo.Width = 74
        '
        'ColRMRekap_Desc
        '
        Me.ColRMRekap_Desc.HeaderText = "Description"
        Me.ColRMRekap_Desc.Name = "ColRMRekap_Desc"
        Me.ColRMRekap_Desc.ReadOnly = True
        Me.ColRMRekap_Desc.Width = 85
        '
        'ColRMRekap_QtyTarget
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRMRekap_QtyTarget.DefaultCellStyle = DataGridViewCellStyle13
        Me.ColRMRekap_QtyTarget.HeaderText = "Qty Target"
        Me.ColRMRekap_QtyTarget.Name = "ColRMRekap_QtyTarget"
        Me.ColRMRekap_QtyTarget.ReadOnly = True
        Me.ColRMRekap_QtyTarget.Visible = False
        Me.ColRMRekap_QtyTarget.Width = 82
        '
        'ColRMRekap_QtyRequired
        '
        Me.ColRMRekap_QtyRequired.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRMRekap_QtyRequired.DefaultCellStyle = DataGridViewCellStyle14
        Me.ColRMRekap_QtyRequired.HeaderText = "Qty Required"
        Me.ColRMRekap_QtyRequired.Name = "ColRMRekap_QtyRequired"
        Me.ColRMRekap_QtyRequired.ReadOnly = True
        Me.ColRMRekap_QtyRequired.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRMRekap_QtyRequired.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRMRekap_QtyRequired.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRMRekap_QtyRequired.Width = 94
        '
        'ColRMRekap_QtyAllocated
        '
        DataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRMRekap_QtyAllocated.DefaultCellStyle = DataGridViewCellStyle15
        Me.ColRMRekap_QtyAllocated.HeaderText = "Qty Allocated"
        Me.ColRMRekap_QtyAllocated.Name = "ColRMRekap_QtyAllocated"
        Me.ColRMRekap_QtyAllocated.ReadOnly = True
        Me.ColRMRekap_QtyAllocated.Visible = False
        Me.ColRMRekap_QtyAllocated.Width = 95
        '
        'ColRMRekap_QtyAdditional
        '
        Me.ColRMRekap_QtyAdditional.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRMRekap_QtyAdditional.DefaultCellStyle = DataGridViewCellStyle16
        Me.ColRMRekap_QtyAdditional.HeaderText = "Qty Additional"
        Me.ColRMRekap_QtyAdditional.Name = "ColRMRekap_QtyAdditional"
        Me.ColRMRekap_QtyAdditional.ReadOnly = True
        Me.ColRMRekap_QtyAdditional.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRMRekap_QtyAdditional.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRMRekap_QtyAdditional.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRMRekap_QtyAdditional.Width = 97
        '
        'ColRMRekap_QtyWIP
        '
        DataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRMRekap_QtyWIP.DefaultCellStyle = DataGridViewCellStyle17
        Me.ColRMRekap_QtyWIP.HeaderText = "Qty WIP"
        Me.ColRMRekap_QtyWIP.Name = "ColRMRekap_QtyWIP"
        Me.ColRMRekap_QtyWIP.ReadOnly = True
        Me.ColRMRekap_QtyWIP.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRMRekap_QtyWIP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRMRekap_QtyWIP.Width = 72
        '
        'ColRMRekap_QtyReplcaed
        '
        Me.ColRMRekap_QtyReplcaed.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRMRekap_QtyReplcaed.DefaultCellStyle = DataGridViewCellStyle18
        Me.ColRMRekap_QtyReplcaed.HeaderText = "Qty Replaced"
        Me.ColRMRekap_QtyReplcaed.Name = "ColRMRekap_QtyReplcaed"
        Me.ColRMRekap_QtyReplcaed.ReadOnly = True
        Me.ColRMRekap_QtyReplcaed.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRMRekap_QtyReplcaed.Width = 78
        '
        'ColRMRekap_QtyUsed
        '
        Me.ColRMRekap_QtyUsed.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle19.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRMRekap_QtyUsed.DefaultCellStyle = DataGridViewCellStyle19
        Me.ColRMRekap_QtyUsed.HeaderText = "Qty Used"
        Me.ColRMRekap_QtyUsed.Name = "ColRMRekap_QtyUsed"
        Me.ColRMRekap_QtyUsed.ReadOnly = True
        Me.ColRMRekap_QtyUsed.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRMRekap_QtyUsed.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRMRekap_QtyUsed.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRMRekap_QtyUsed.Width = 76
        '
        'ColRMRekap_QtyBalance
        '
        Me.ColRMRekap_QtyBalance.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRMRekap_QtyBalance.DefaultCellStyle = DataGridViewCellStyle20
        Me.ColRMRekap_QtyBalance.HeaderText = "Qty Balance"
        Me.ColRMRekap_QtyBalance.Name = "ColRMRekap_QtyBalance"
        Me.ColRMRekap_QtyBalance.ReadOnly = True
        Me.ColRMRekap_QtyBalance.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRMRekap_QtyBalance.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRMRekap_QtyBalance.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRMRekap_QtyBalance.Width = 90
        '
        'ColRMRekap_Unit
        '
        Me.ColRMRekap_Unit.HeaderText = "Unit"
        Me.ColRMRekap_Unit.Name = "ColRMRekap_Unit"
        Me.ColRMRekap_Unit.ReadOnly = True
        Me.ColRMRekap_Unit.Width = 51
        '
        'ColRMRekap_IdentifierJO
        '
        Me.ColRMRekap_IdentifierJO.HeaderText = "Identifier JO"
        Me.ColRMRekap_IdentifierJO.Name = "ColRMRekap_IdentifierJO"
        Me.ColRMRekap_IdentifierJO.ReadOnly = True
        Me.ColRMRekap_IdentifierJO.Visible = False
        Me.ColRMRekap_IdentifierJO.Width = 88
        '
        'ColRMRekap_PrimaryCodeNo
        '
        Me.ColRMRekap_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColRMRekap_PrimaryCodeNo.Name = "ColRMRekap_PrimaryCodeNo"
        Me.ColRMRekap_PrimaryCodeNo.ReadOnly = True
        Me.ColRMRekap_PrimaryCodeNo.Visible = False
        Me.ColRMRekap_PrimaryCodeNo.Width = 111
        '
        'ColRMRekap_PrimaryUnit
        '
        Me.ColRMRekap_PrimaryUnit.HeaderText = "Primary Unit"
        Me.ColRMRekap_PrimaryUnit.Name = "ColRMRekap_PrimaryUnit"
        Me.ColRMRekap_PrimaryUnit.ReadOnly = True
        Me.ColRMRekap_PrimaryUnit.Visible = False
        Me.ColRMRekap_PrimaryUnit.Width = 88
        '
        'TabPage_WO
        '
        Me.TabPage_WO.Controls.Add(Me.dgvItemWO)
        Me.TabPage_WO.Controls.Add(Me.TGV_WO)
        Me.TabPage_WO.Location = New System.Drawing.Point(4, 22)
        Me.TabPage_WO.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage_WO.Name = "TabPage_WO"
        Me.TabPage_WO.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage_WO.Size = New System.Drawing.Size(839, 186)
        Me.TabPage_WO.TabIndex = 1
        Me.TabPage_WO.Text = "Work Order"
        Me.TabPage_WO.UseVisualStyleBackColor = True
        '
        'dgvItemWO
        '
        Me.dgvItemWO.AllowUserToAddRows = False
        Me.dgvItemWO.AllowUserToDeleteRows = False
        Me.dgvItemWO.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemWO.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItemWO.ColumnHeadersHeight = 25
        Me.dgvItemWO.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemWO.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColWOTbl_No, Me.ColWOTbl_Status, Me.ColWOTbl_Progress, Me.ColWOTbl_WONO, Me.ColWOTbl_CodeNo, Me.ColWOTbl_Desc, Me.ColWOTbl_ProcessCode, Me.ColWOTbl_ProcessDesc, Me.ColWOTbl_WorkUnit, Me.ColWOTbl_Remarks, Me.ColWOTbl_ShowDetail, Me.ColWOTbl_PrimaryProcess, Me.ColWOTbl_IdentifierWO, Me.ColWOTbl_IdentifierJO, Me.ColWOTbl_StartDate, Me.ColWOTbl_StartTime, Me.ColWOTbl_EndDate, Me.ColWOTbl_EndTime})
        Me.dgvItemWO.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemWO.Location = New System.Drawing.Point(2, 2)
        Me.dgvItemWO.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvItemWO.Name = "dgvItemWO"
        Me.dgvItemWO.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemWO.RowTemplate.Height = 24
        Me.dgvItemWO.Size = New System.Drawing.Size(835, 182)
        Me.dgvItemWO.TabIndex = 21
        '
        'ColWOTbl_No
        '
        Me.ColWOTbl_No.HeaderText = "No"
        Me.ColWOTbl_No.Name = "ColWOTbl_No"
        Me.ColWOTbl_No.Width = 46
        '
        'ColWOTbl_Status
        '
        Me.ColWOTbl_Status.HeaderText = "Status"
        Me.ColWOTbl_Status.Items.AddRange(New Object() {"Free", "Planned", "Scheduled", "On Progress", "Finish"})
        Me.ColWOTbl_Status.Name = "ColWOTbl_Status"
        Me.ColWOTbl_Status.Width = 43
        '
        'ColWOTbl_Progress
        '
        Me.ColWOTbl_Progress.HeaderText = "Progress (%)"
        Me.ColWOTbl_Progress.Name = "ColWOTbl_Progress"
        Me.ColWOTbl_Progress.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_Progress.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_Progress.Width = 71
        '
        'ColWOTbl_WONO
        '
        Me.ColWOTbl_WONO.HeaderText = "WO No"
        Me.ColWOTbl_WONO.Name = "ColWOTbl_WONO"
        Me.ColWOTbl_WONO.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_WONO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_WONO.Width = 49
        '
        'ColWOTbl_CodeNo
        '
        Me.ColWOTbl_CodeNo.HeaderText = "Code No"
        Me.ColWOTbl_CodeNo.Name = "ColWOTbl_CodeNo"
        Me.ColWOTbl_CodeNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_CodeNo.Width = 55
        '
        'ColWOTbl_Desc
        '
        Me.ColWOTbl_Desc.HeaderText = "Description"
        Me.ColWOTbl_Desc.Name = "ColWOTbl_Desc"
        Me.ColWOTbl_Desc.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_Desc.Width = 66
        '
        'ColWOTbl_ProcessCode
        '
        Me.ColWOTbl_ProcessCode.HeaderText = "Process Code"
        Me.ColWOTbl_ProcessCode.Name = "ColWOTbl_ProcessCode"
        Me.ColWOTbl_ProcessCode.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_ProcessCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_ProcessCode.Width = 79
        '
        'ColWOTbl_ProcessDesc
        '
        Me.ColWOTbl_ProcessDesc.HeaderText = "Process Dsecription"
        Me.ColWOTbl_ProcessDesc.Name = "ColWOTbl_ProcessDesc"
        Me.ColWOTbl_ProcessDesc.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_ProcessDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_ProcessDesc.Width = 107
        '
        'ColWOTbl_WorkUnit
        '
        Me.ColWOTbl_WorkUnit.HeaderText = "Work Unit"
        Me.ColWOTbl_WorkUnit.Name = "ColWOTbl_WorkUnit"
        Me.ColWOTbl_WorkUnit.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_WorkUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_WorkUnit.Width = 61
        '
        'ColWOTbl_Remarks
        '
        Me.ColWOTbl_Remarks.HeaderText = "Remarks"
        Me.ColWOTbl_Remarks.Name = "ColWOTbl_Remarks"
        Me.ColWOTbl_Remarks.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWOTbl_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWOTbl_Remarks.Width = 55
        '
        'ColWOTbl_ShowDetail
        '
        Me.ColWOTbl_ShowDetail.HeaderText = "Show Detail"
        Me.ColWOTbl_ShowDetail.Name = "ColWOTbl_ShowDetail"
        Me.ColWOTbl_ShowDetail.Width = 70
        '
        'ColWOTbl_PrimaryProcess
        '
        Me.ColWOTbl_PrimaryProcess.HeaderText = "Primary Process"
        Me.ColWOTbl_PrimaryProcess.Name = "ColWOTbl_PrimaryProcess"
        Me.ColWOTbl_PrimaryProcess.Visible = False
        Me.ColWOTbl_PrimaryProcess.Width = 107
        '
        'ColWOTbl_IdentifierWO
        '
        Me.ColWOTbl_IdentifierWO.HeaderText = "Identifier WO"
        Me.ColWOTbl_IdentifierWO.Name = "ColWOTbl_IdentifierWO"
        Me.ColWOTbl_IdentifierWO.Visible = False
        Me.ColWOTbl_IdentifierWO.Width = 94
        '
        'ColWOTbl_IdentifierJO
        '
        Me.ColWOTbl_IdentifierJO.HeaderText = "Identifier JO"
        Me.ColWOTbl_IdentifierJO.Name = "ColWOTbl_IdentifierJO"
        Me.ColWOTbl_IdentifierJO.Visible = False
        Me.ColWOTbl_IdentifierJO.Width = 88
        '
        'ColWOTbl_StartDate
        '
        Me.ColWOTbl_StartDate.HeaderText = "Start"
        Me.ColWOTbl_StartDate.Name = "ColWOTbl_StartDate"
        Me.ColWOTbl_StartDate.Width = 54
        '
        'ColWOTbl_StartTime
        '
        Me.ColWOTbl_StartTime.HeaderText = ""
        Me.ColWOTbl_StartTime.Name = "ColWOTbl_StartTime"
        Me.ColWOTbl_StartTime.Width = 19
        '
        'ColWOTbl_EndDate
        '
        Me.ColWOTbl_EndDate.HeaderText = "End"
        Me.ColWOTbl_EndDate.Name = "ColWOTbl_EndDate"
        Me.ColWOTbl_EndDate.Width = 51
        '
        'ColWOTbl_EndTime
        '
        Me.ColWOTbl_EndTime.HeaderText = ""
        Me.ColWOTbl_EndTime.Name = "ColWOTbl_EndTime"
        Me.ColWOTbl_EndTime.Width = 19
        '
        'TGV_WO
        '
        Me.TGV_WO.AllowUserToAddRows = False
        Me.TGV_WO.AllowUserToDeleteRows = False
        Me.TGV_WO.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.TGV_WO.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColWO2_Nodes, Me.ColWO2_Status, Me.ColWO2_Progress, Me.ColWO2_WONumber, Me.ColWO2_CodeNo, Me.ColWO2_ProductDesc, Me.ColWO2_ProcessCode, Me.ColWO2_ProcessDesc, Me.ColWO2_WorkUnit, Me.ColWO2_Remarks, Me.ColWO2_ShowDetail, Me.ColWO2_PrimaryProcess, Me.ColWO2_IdentifierItem, Me.ColWO2_IdentifierJO})
        Me.TGV_WO.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TGV_WO.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.TGV_WO.ImageList = Nothing
        Me.TGV_WO.Location = New System.Drawing.Point(2, 2)
        Me.TGV_WO.Margin = New System.Windows.Forms.Padding(2)
        Me.TGV_WO.MultiSelect = False
        Me.TGV_WO.Name = "TGV_WO"
        Me.TGV_WO.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.TGV_WO.Size = New System.Drawing.Size(835, 182)
        Me.TGV_WO.TabIndex = 20
        '
        'ColWO2_Nodes
        '
        Me.ColWO2_Nodes.DefaultNodeImage = Nothing
        Me.ColWO2_Nodes.Frozen = True
        Me.ColWO2_Nodes.HeaderText = ""
        Me.ColWO2_Nodes.Name = "ColWO2_Nodes"
        Me.ColWO2_Nodes.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWO2_Nodes.Width = 5
        '
        'ColWO2_Status
        '
        Me.ColWO2_Status.Frozen = True
        Me.ColWO2_Status.HeaderText = "Status"
        Me.ColWO2_Status.Items.AddRange(New Object() {"Free", "Planned", "Scheduled", "On Progress", "Finish"})
        Me.ColWO2_Status.Name = "ColWO2_Status"
        Me.ColWO2_Status.Width = 43
        '
        'ColWO2_Progress
        '
        DataGridViewCellStyle21.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColWO2_Progress.DefaultCellStyle = DataGridViewCellStyle21
        Me.ColWO2_Progress.Frozen = True
        Me.ColWO2_Progress.HeaderText = "Progress (%)"
        Me.ColWO2_Progress.Name = "ColWO2_Progress"
        Me.ColWO2_Progress.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWO2_Progress.Width = 71
        '
        'ColWO2_WONumber
        '
        Me.ColWO2_WONumber.HeaderText = "WO Number"
        Me.ColWO2_WONumber.Name = "ColWO2_WONumber"
        Me.ColWO2_WONumber.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWO2_WONumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWO2_WONumber.Width = 72
        '
        'ColWO2_CodeNo
        '
        Me.ColWO2_CodeNo.HeaderText = "Code No"
        Me.ColWO2_CodeNo.Name = "ColWO2_CodeNo"
        Me.ColWO2_CodeNo.ReadOnly = True
        Me.ColWO2_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWO2_CodeNo.Width = 55
        '
        'ColWO2_ProductDesc
        '
        Me.ColWO2_ProductDesc.HeaderText = "Description"
        Me.ColWO2_ProductDesc.Name = "ColWO2_ProductDesc"
        Me.ColWO2_ProductDesc.ReadOnly = True
        Me.ColWO2_ProductDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWO2_ProductDesc.Width = 66
        '
        'ColWO2_ProcessCode
        '
        Me.ColWO2_ProcessCode.HeaderText = "Process Code"
        Me.ColWO2_ProcessCode.Name = "ColWO2_ProcessCode"
        Me.ColWO2_ProcessCode.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWO2_ProcessCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWO2_ProcessCode.Width = 79
        '
        'ColWO2_ProcessDesc
        '
        Me.ColWO2_ProcessDesc.HeaderText = "Process Description"
        Me.ColWO2_ProcessDesc.Name = "ColWO2_ProcessDesc"
        Me.ColWO2_ProcessDesc.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWO2_ProcessDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWO2_ProcessDesc.Width = 107
        '
        'ColWO2_WorkUnit
        '
        Me.ColWO2_WorkUnit.HeaderText = "Work Unit"
        Me.ColWO2_WorkUnit.Name = "ColWO2_WorkUnit"
        Me.ColWO2_WorkUnit.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWO2_WorkUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWO2_WorkUnit.Width = 61
        '
        'ColWO2_Remarks
        '
        Me.ColWO2_Remarks.HeaderText = "Remarks"
        Me.ColWO2_Remarks.MaxInputLength = 250
        Me.ColWO2_Remarks.Name = "ColWO2_Remarks"
        Me.ColWO2_Remarks.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWO2_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWO2_Remarks.Width = 55
        '
        'ColWO2_ShowDetail
        '
        Me.ColWO2_ShowDetail.HeaderText = "Show Detail"
        Me.ColWO2_ShowDetail.Name = "ColWO2_ShowDetail"
        Me.ColWO2_ShowDetail.Width = 70
        '
        'ColWO2_PrimaryProcess
        '
        Me.ColWO2_PrimaryProcess.HeaderText = "Primary Process"
        Me.ColWO2_PrimaryProcess.Name = "ColWO2_PrimaryProcess"
        Me.ColWO2_PrimaryProcess.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWO2_PrimaryProcess.Visible = False
        Me.ColWO2_PrimaryProcess.Width = 88
        '
        'ColWO2_IdentifierItem
        '
        Me.ColWO2_IdentifierItem.HeaderText = "Identifier Item"
        Me.ColWO2_IdentifierItem.Name = "ColWO2_IdentifierItem"
        Me.ColWO2_IdentifierItem.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWO2_IdentifierItem.Visible = False
        Me.ColWO2_IdentifierItem.Width = 76
        '
        'ColWO2_IdentifierJO
        '
        Me.ColWO2_IdentifierJO.HeaderText = "Identifier JO"
        Me.ColWO2_IdentifierJO.Name = "ColWO2_IdentifierJO"
        Me.ColWO2_IdentifierJO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWO2_IdentifierJO.Visible = False
        Me.ColWO2_IdentifierJO.Width = 69
        '
        'TabPage_WIP
        '
        Me.TabPage_WIP.Controls.Add(Me.DGVWIP)
        Me.TabPage_WIP.Location = New System.Drawing.Point(4, 22)
        Me.TabPage_WIP.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage_WIP.Name = "TabPage_WIP"
        Me.TabPage_WIP.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage_WIP.Size = New System.Drawing.Size(839, 186)
        Me.TabPage_WIP.TabIndex = 2
        Me.TabPage_WIP.Text = "WIP Allocation"
        Me.TabPage_WIP.UseVisualStyleBackColor = True
        '
        'DGVWIP
        '
        Me.DGVWIP.AllowUserToAddRows = False
        Me.DGVWIP.AllowUserToResizeColumns = False
        Me.DGVWIP.AllowUserToResizeRows = False
        Me.DGVWIP.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVWIP.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DGVWIP.ColumnHeadersHeight = 25
        Me.DGVWIP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DGVWIP.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColWIP_No, Me.ColWIP_WONo, Me.ColWIP_Proses, Me.ColWIP_CodeNo, Me.ColWIP_Desc, Me.ColWIP_QtyWIP, Me.ColWIP_UnitCode, Me.ColWIP_Remarks, Me.ColWIP_IdentifierJO, Me.ColWIP_IdentifierWO, Me.ColWIP_PrimaryProses, Me.ColWIP_PrimaryCodeNo, Me.ColWIP_PrimaryUnit})
        Me.DGVWIP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVWIP.Location = New System.Drawing.Point(2, 2)
        Me.DGVWIP.Margin = New System.Windows.Forms.Padding(2)
        Me.DGVWIP.Name = "DGVWIP"
        Me.DGVWIP.RowHeadersVisible = False
        Me.DGVWIP.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DGVWIP.RowTemplate.Height = 24
        Me.DGVWIP.Size = New System.Drawing.Size(835, 182)
        Me.DGVWIP.TabIndex = 20
        '
        'ColWIP_No
        '
        Me.ColWIP_No.HeaderText = "No"
        Me.ColWIP_No.Name = "ColWIP_No"
        Me.ColWIP_No.Width = 46
        '
        'ColWIP_WONo
        '
        Me.ColWIP_WONo.HeaderText = "WO Number"
        Me.ColWIP_WONo.Name = "ColWIP_WONo"
        Me.ColWIP_WONo.Width = 91
        '
        'ColWIP_Proses
        '
        Me.ColWIP_Proses.HeaderText = "Process"
        Me.ColWIP_Proses.Name = "ColWIP_Proses"
        Me.ColWIP_Proses.Width = 70
        '
        'ColWIP_CodeNo
        '
        Me.ColWIP_CodeNo.HeaderText = "Code No"
        Me.ColWIP_CodeNo.Name = "ColWIP_CodeNo"
        Me.ColWIP_CodeNo.Width = 74
        '
        'ColWIP_Desc
        '
        Me.ColWIP_Desc.HeaderText = "Description"
        Me.ColWIP_Desc.Name = "ColWIP_Desc"
        Me.ColWIP_Desc.Width = 85
        '
        'ColWIP_QtyWIP
        '
        DataGridViewCellStyle22.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColWIP_QtyWIP.DefaultCellStyle = DataGridViewCellStyle22
        Me.ColWIP_QtyWIP.HeaderText = "Qty WIP"
        Me.ColWIP_QtyWIP.Name = "ColWIP_QtyWIP"
        Me.ColWIP_QtyWIP.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWIP_QtyWIP.Width = 72
        '
        'ColWIP_UnitCode
        '
        Me.ColWIP_UnitCode.HeaderText = "Unit"
        Me.ColWIP_UnitCode.Name = "ColWIP_UnitCode"
        Me.ColWIP_UnitCode.Width = 51
        '
        'ColWIP_Remarks
        '
        Me.ColWIP_Remarks.HeaderText = "Remarks"
        Me.ColWIP_Remarks.Name = "ColWIP_Remarks"
        Me.ColWIP_Remarks.Visible = False
        Me.ColWIP_Remarks.Width = 74
        '
        'ColWIP_IdentifierJO
        '
        Me.ColWIP_IdentifierJO.HeaderText = "Identifier JO"
        Me.ColWIP_IdentifierJO.Name = "ColWIP_IdentifierJO"
        Me.ColWIP_IdentifierJO.Visible = False
        Me.ColWIP_IdentifierJO.Width = 88
        '
        'ColWIP_IdentifierWO
        '
        Me.ColWIP_IdentifierWO.HeaderText = "Identifier WO"
        Me.ColWIP_IdentifierWO.Name = "ColWIP_IdentifierWO"
        Me.ColWIP_IdentifierWO.Visible = False
        Me.ColWIP_IdentifierWO.Width = 94
        '
        'ColWIP_PrimaryProses
        '
        Me.ColWIP_PrimaryProses.HeaderText = "Primary Process"
        Me.ColWIP_PrimaryProses.Name = "ColWIP_PrimaryProses"
        Me.ColWIP_PrimaryProses.Visible = False
        Me.ColWIP_PrimaryProses.Width = 107
        '
        'ColWIP_PrimaryCodeNo
        '
        Me.ColWIP_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColWIP_PrimaryCodeNo.Name = "ColWIP_PrimaryCodeNo"
        Me.ColWIP_PrimaryCodeNo.Visible = False
        Me.ColWIP_PrimaryCodeNo.Width = 111
        '
        'ColWIP_PrimaryUnit
        '
        Me.ColWIP_PrimaryUnit.HeaderText = "Primary Unit"
        Me.ColWIP_PrimaryUnit.Name = "ColWIP_PrimaryUnit"
        Me.ColWIP_PrimaryUnit.Visible = False
        Me.ColWIP_PrimaryUnit.Width = 88
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryUserAuth)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpTargetFinishDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkScheduling)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpTglTransaksi)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtNoTransaksi)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TextBox2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimary)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
        Me.SplitContainer1.Size = New System.Drawing.Size(847, 55)
        Me.SplitContainer1.SplitterDistance = 460
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 16
        '
        'TxtPrimaryUserAuth
        '
        Me.TxtPrimaryUserAuth.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryUserAuth.Location = New System.Drawing.Point(43, 27)
        Me.TxtPrimaryUserAuth.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtPrimaryUserAuth.Name = "TxtPrimaryUserAuth"
        Me.TxtPrimaryUserAuth.Size = New System.Drawing.Size(32, 20)
        Me.TxtPrimaryUserAuth.TabIndex = 9
        Me.TxtPrimaryUserAuth.Text = "-1"
        Me.TxtPrimaryUserAuth.Visible = False
        '
        'DtpTargetFinishDate
        '
        Me.DtpTargetFinishDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTargetFinishDate.CustomFormat = "dd MMM yyyy"
        Me.DtpTargetFinishDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTargetFinishDate.Location = New System.Drawing.Point(334, 28)
        Me.DtpTargetFinishDate.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpTargetFinishDate.Name = "DtpTargetFinishDate"
        Me.DtpTargetFinishDate.Size = New System.Drawing.Size(120, 20)
        Me.DtpTargetFinishDate.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(236, 31)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(94, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Target Finish Date"
        '
        'ChkScheduling
        '
        Me.ChkScheduling.AutoSize = True
        Me.ChkScheduling.Location = New System.Drawing.Point(84, 30)
        Me.ChkScheduling.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkScheduling.Name = "ChkScheduling"
        Me.ChkScheduling.Size = New System.Drawing.Size(131, 17)
        Me.ChkScheduling.TabIndex = 5
        Me.ChkScheduling.Text = "Ready For Scheduling"
        Me.ChkScheduling.UseVisualStyleBackColor = True
        '
        'DtpTglTransaksi
        '
        Me.DtpTglTransaksi.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTglTransaksi.CustomFormat = "dd MMM yyyy"
        Me.DtpTglTransaksi.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTglTransaksi.Location = New System.Drawing.Point(334, 4)
        Me.DtpTglTransaksi.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpTglTransaksi.Name = "DtpTglTransaksi"
        Me.DtpTglTransaksi.Size = New System.Drawing.Size(120, 20)
        Me.DtpTglTransaksi.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(236, 8)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Date"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 8)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "MO Number"
        '
        'TxtNoTransaksi
        '
        Me.TxtNoTransaksi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNoTransaksi.Location = New System.Drawing.Point(84, 4)
        Me.TxtNoTransaksi.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtNoTransaksi.Name = "TxtNoTransaksi"
        Me.TxtNoTransaksi.ReadOnly = True
        Me.TxtNoTransaksi.Size = New System.Drawing.Size(130, 20)
        Me.TxtNoTransaksi.TabIndex = 0
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(84, 4)
        Me.TextBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(38, 20)
        Me.TextBox2.TabIndex = 4
        Me.TextBox2.Text = "-1"
        Me.TextBox2.Visible = False
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimary.Location = New System.Drawing.Point(43, 5)
        Me.TxtPrimary.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(11, 20)
        Me.TxtPrimary.TabIndex = 7
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(72, 4)
        Me.TxtRemarks.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtRemarks.MaxLength = 250
        Me.TxtRemarks.Multiline = True
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(310, 44)
        Me.TxtRemarks.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 6)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Remarks"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 492)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(881, 22)
        Me.StatusStrip1.TabIndex = 15
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReadOnly.Location = New System.Drawing.Point(771, 456)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(52, 20)
        Me.TxtReadOnly.TabIndex = 4
        Me.TxtReadOnly.Visible = False
        '
        'llbColumnSetting
        '
        Me.llbColumnSetting.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbColumnSetting.AutoSize = True
        Me.llbColumnSetting.Location = New System.Drawing.Point(10, 462)
        Me.llbColumnSetting.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbColumnSetting.Name = "llbColumnSetting"
        Me.llbColumnSetting.Size = New System.Drawing.Size(83, 13)
        Me.llbColumnSetting.TabIndex = 20
        Me.llbColumnSetting.TabStop = True
        Me.llbColumnSetting.Text = "Columns Setting"
        '
        'BackgroundWorker2
        '
        '
        'FrmTrans_ManufacturingOrder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(881, 514)
        Me.Controls.Add(Me.llbColumnSetting)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(897, 552)
        Me.Name = "FrmTrans_ManufacturingOrder"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Manufacturing Order"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.Panel2.PerformLayout()
        Me.SplitContainer3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.TabPage5.ResumeLayout(False)
        Me.TabPage5.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage_RM.ResumeLayout(False)
        CType(Me.DGVRM, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DgvRMRekap, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage_WO.ResumeLayout(False)
        CType(Me.dgvItemWO, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TGV_WO, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage_WIP.ResumeLayout(False)
        CType(Me.DGVWIP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents DtpTglTransaksi As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtNoTransaksi As System.Windows.Forms.TextBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ChkScheduling As System.Windows.Forms.CheckBox
    Friend WithEvents TGV_WO As AdvancedDataGridView.TreeGridView
    Friend WithEvents llbFilter As System.Windows.Forms.LinkLabel
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents ClbProcess As System.Windows.Forms.CheckedListBox
    Friend WithEvents llbColumnSetting As System.Windows.Forms.LinkLabel
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents ChkGoodPerc As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLateWO As System.Windows.Forms.CheckBox
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents ChkInsufficientRM As System.Windows.Forms.CheckBox
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DtpTargetFinishDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimaryUserAuth As System.Windows.Forms.TextBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage_WO As System.Windows.Forms.TabPage
    Friend WithEvents TabPage_RM As System.Windows.Forms.TabPage
    Friend WithEvents DGVRM As System.Windows.Forms.DataGridView
    Friend WithEvents ColWO2_Nodes As AdvancedDataGridView.TreeGridColumn
    Friend WithEvents ColWO2_Status As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColWO2_Progress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWO2_WONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWO2_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWO2_ProductDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWO2_ProcessCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWO2_ProcessDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWO2_WorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWO2_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWO2_ShowDetail As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColWO2_PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWO2_IdentifierItem As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWO2_IdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnStopProcess As System.Windows.Forms.Button
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents prgAll As System.Windows.Forms.ProgressBar
    Friend WithEvents chkRM_IncludeWIP As System.Windows.Forms.CheckBox
    Friend WithEvents chkRM_ShowDetail As System.Windows.Forms.CheckBox
    Friend WithEvents DgvRMRekap As System.Windows.Forms.DataGridView
    Friend WithEvents rboWO_UseTableView As System.Windows.Forms.RadioButton
    Friend WithEvents rboWO_UseTreeView As System.Windows.Forms.RadioButton
    Friend WithEvents dgvItemWO As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage_WIP As System.Windows.Forms.TabPage
    Friend WithEvents DGVWIP As System.Windows.Forms.DataGridView
    Friend WithEvents llbAddWIP As System.Windows.Forms.LinkLabel
    Friend WithEvents ColWIP_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_WONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_Proses As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_QtyWIP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_UnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_IdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_IdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_PrimaryProses As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_PrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMRekap_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMRekap_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMRekap_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMRekap_QtyTarget As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMRekap_QtyRequired As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRMRekap_QtyAllocated As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMRekap_QtyAdditional As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRMRekap_QtyWIP As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRMRekap_QtyReplcaed As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRMRekap_QtyUsed As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRMRekap_QtyBalance As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRMRekap_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMRekap_IdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMRekap_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMRekap_PrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_WONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_Process As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_QtyTarget As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_QtyRequired As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRM_QtyAllocated As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_QtyAdditional As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRM_QtyWIP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_QtyReplaced As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRM_QtyUsed As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRM_QtyBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_GPTarget As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_GPActual As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_StandByLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_PickUpLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_RequiredDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_PrimaryStandByLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_PrimaryPickUpLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_IdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_IdentifierItem As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_PrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_IsWIP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_JONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Allocation As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColItem_QtyTarget As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_QtyProduce As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_DeliveryDate As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColItem_TargetFinishDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_TargetFinishDateR1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_TargetFinishDateR2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_ActualFinishDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_EditBOM As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColItem_BoM_Tools As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColItem_IsClosed As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColItem_Decimal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Identifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Urutan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_RefDocumentPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_PrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Print As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColItem_SetRequiredDate As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColWOTbl_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_Status As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColWOTbl_Progress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_WONO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_ProcessCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_ProcessDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_WorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_ShowDetail As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColWOTbl_PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_IdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_IdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_StartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_StartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_EndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWOTbl_EndTime As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
