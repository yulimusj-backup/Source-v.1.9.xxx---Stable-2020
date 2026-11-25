<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockOutSales
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTransaksiStockOutSales))
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.btnExit = New System.Windows.Forms.Button
        Me.TxtPrimaryStockIn = New System.Windows.Forms.TextBox
        Me.TxtStockInNo = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtPihakAcuan = New System.Windows.Forms.TextBox
        Me.LblNamaPihakAcuan = New System.Windows.Forms.Label
        Me.LblKodePihakAcuan = New System.Windows.Forms.Label
        Me.TxtKodePihakAcuan = New System.Windows.Forms.TextBox
        Me.BtnCariPihakAcuan = New System.Windows.Forms.Button
        Me.txtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtPrimaryPihakAcuan = New System.Windows.Forms.TextBox
        Me.txtDONo = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnFindDO = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPONoDONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppDONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitsCodeMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecondaryQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecondaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatch = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiredDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWareHouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPOnoDoNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySecondaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPdO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitsMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvItemDO = New System.Windows.Forms.DataGridView
        Me.ColDO2_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_CustCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_CustDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_SONO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_BatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_ExpiryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_QtyMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_QtyPerMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_MasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_TotalQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_UnitsCodeMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_RatioDesc = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColDO2_TotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_SecQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_SecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_PrimarySO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_PrimaryMasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_PrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_PrimarySecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_RatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_RatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_PrimaryUnitsMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_Location = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_Warehouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO2_PrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.lblSecondaryUnits = New System.Windows.Forms.Label
        Me.gboItem = New System.Windows.Forms.GroupBox
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.txtSalesOrderNo = New System.Windows.Forms.TextBox
        Me.btnFindSalesOrder = New System.Windows.Forms.Button
        Me.lblBatchNo = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.lblExpiryDate = New System.Windows.Forms.Label
        Me.BtnFindCodeNo = New System.Windows.Forms.Button
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.DtpExpired = New System.Windows.Forms.DateTimePicker
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.cboBatchNo = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.BtnFindPONo = New System.Windows.Forms.Button
        Me.TxtCodeNo = New System.Windows.Forms.TextBox
        Me.LblQty = New System.Windows.Forms.Label
        Me.txtPrimaryUnitsMulti = New System.Windows.Forms.TextBox
        Me.TxtQtyperBatchMulti = New System.Windows.Forms.TextBox
        Me.lblUnitsMulti = New System.Windows.Forms.Label
        Me.TxtRatioDefault = New System.Windows.Forms.TextBox
        Me.TxtRatioReference = New System.Windows.Forms.TextBox
        Me.lblWarehouse = New System.Windows.Forms.Label
        Me.lblLocation = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.BtnInput = New System.Windows.Forms.Button
        Me.TxtWarehouse = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.TxtRemarkItem = New System.Windows.Forms.TextBox
        Me.txtPrimaryUnitsDefault = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtQtyperBatchDefault = New System.Windows.Forms.TextBox
        Me.lblUnitsDefault = New System.Windows.Forms.Label
        Me.cboLocation = New System.Windows.Forms.ComboBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.chkBatchNoControlled = New System.Windows.Forms.CheckBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.chkPdoControlled = New System.Windows.Forms.CheckBox
        Me.txtPrimaryDO = New System.Windows.Forms.TextBox
        Me.cboSecondaryUnits = New System.Windows.Forms.ComboBox
        Me.lblSecQty = New System.Windows.Forms.Label
        Me.TxtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.txtPrimarySO = New System.Windows.Forms.TextBox
        Me.chkUseSecondaryUnit = New System.Windows.Forms.CheckBox
        Me.chkLocationControlled = New System.Windows.Forms.CheckBox
        Me.chkExpiryDateControlled = New System.Windows.Forms.CheckBox
        Me.txtPrimaryCategory = New System.Windows.Forms.TextBox
        Me.txtPrimarySecondaryUnits = New System.Windows.Forms.TextBox
        Me.txtPrimaryLocation = New System.Windows.Forms.TextBox
        Me.txtPrimaryPDO = New System.Windows.Forms.TextBox
        Me.txtSecondaryQty = New System.Windows.Forms.TextBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DgvItemDO2 = New System.Windows.Forms.DataGridView
        Me.ColDO_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_CustCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_CustDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_SONO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_BatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_ExpiryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_QtyMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_QtyPerMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_MasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_TotalQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_UnitsCodeMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_RatioDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_TotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_SecQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_SecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_PrimarySO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_PrimaryMasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_PrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_PrimarySecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_RatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_RatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_PrimaryUnitsMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_Location = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_Warehouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDO_PrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.AddItemByCustom = New System.Windows.Forms.ToolStripMenuItem
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItemDO, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboItem.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvItemDO2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(85, 29)
        Me.TxtRemark.MaxLength = 150
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(344, 68)
        Me.TxtRemark.TabIndex = 2
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 33)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 65
        Me.Label6.Text = "Remarks"
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(107, 29)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(321, 20)
        Me.dtpDate.TabIndex = 1
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(623, 515)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'TxtPrimaryStockIn
        '
        Me.TxtPrimaryStockIn.Location = New System.Drawing.Point(409, 166)
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
        Me.TxtStockInNo.Location = New System.Drawing.Point(107, 6)
        Me.TxtStockInNo.Name = "TxtStockInNo"
        Me.TxtStockInNo.ReadOnly = True
        Me.TxtStockInNo.Size = New System.Drawing.Size(321, 20)
        Me.TxtStockInNo.TabIndex = 0
        Me.TxtStockInNo.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 33)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Date"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.TabControl1)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(4, 1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(887, 508)
        Me.gboDetail.TabIndex = 6
        Me.gboDetail.TabStop = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 11)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPihakAcuan)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblNamaPihakAcuan)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblKodePihakAcuan)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtStockInNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtKodePihakAcuan)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnCariPihakAcuan)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtReadOnly)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryPihakAcuan)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtDONo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFindDO)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemark)
        Me.SplitContainer1.Size = New System.Drawing.Size(875, 99)
        Me.SplitContainer1.SplitterDistance = 431
        Me.SplitContainer1.TabIndex = 102
        Me.SplitContainer1.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 10)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Stock Out No"
        '
        'TxtPihakAcuan
        '
        Me.TxtPihakAcuan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPihakAcuan.Location = New System.Drawing.Point(107, 76)
        Me.TxtPihakAcuan.MaxLength = 30
        Me.TxtPihakAcuan.Name = "TxtPihakAcuan"
        Me.TxtPihakAcuan.ReadOnly = True
        Me.TxtPihakAcuan.Size = New System.Drawing.Size(321, 20)
        Me.TxtPihakAcuan.TabIndex = 3
        Me.TxtPihakAcuan.TabStop = False
        '
        'LblNamaPihakAcuan
        '
        Me.LblNamaPihakAcuan.AutoSize = True
        Me.LblNamaPihakAcuan.Location = New System.Drawing.Point(9, 80)
        Me.LblNamaPihakAcuan.Name = "LblNamaPihakAcuan"
        Me.LblNamaPihakAcuan.Size = New System.Drawing.Size(82, 13)
        Me.LblNamaPihakAcuan.TabIndex = 71
        Me.LblNamaPihakAcuan.Text = "Customer Name"
        '
        'LblKodePihakAcuan
        '
        Me.LblKodePihakAcuan.AutoSize = True
        Me.LblKodePihakAcuan.Location = New System.Drawing.Point(9, 56)
        Me.LblKodePihakAcuan.Name = "LblKodePihakAcuan"
        Me.LblKodePihakAcuan.Size = New System.Drawing.Size(79, 13)
        Me.LblKodePihakAcuan.TabIndex = 70
        Me.LblKodePihakAcuan.Text = "Customer Code"
        '
        'TxtKodePihakAcuan
        '
        Me.TxtKodePihakAcuan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtKodePihakAcuan.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtKodePihakAcuan.Location = New System.Drawing.Point(107, 52)
        Me.TxtKodePihakAcuan.MaxLength = 6
        Me.TxtKodePihakAcuan.Name = "TxtKodePihakAcuan"
        Me.TxtKodePihakAcuan.ReadOnly = True
        Me.TxtKodePihakAcuan.Size = New System.Drawing.Size(321, 20)
        Me.TxtKodePihakAcuan.TabIndex = 2
        Me.TxtKodePihakAcuan.TabStop = False
        '
        'BtnCariPihakAcuan
        '
        Me.BtnCariPihakAcuan.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnCariPihakAcuan.Location = New System.Drawing.Point(386, 52)
        Me.BtnCariPihakAcuan.Name = "BtnCariPihakAcuan"
        Me.BtnCariPihakAcuan.Size = New System.Drawing.Size(41, 20)
        Me.BtnCariPihakAcuan.TabIndex = 3
        Me.BtnCariPihakAcuan.TabStop = False
        Me.BtnCariPihakAcuan.Text = "...."
        Me.BtnCariPihakAcuan.UseVisualStyleBackColor = True
        Me.BtnCariPihakAcuan.Visible = False
        '
        'txtReadOnly
        '
        Me.txtReadOnly.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtReadOnly.Location = New System.Drawing.Point(43, 76)
        Me.txtReadOnly.MaxLength = 6
        Me.txtReadOnly.Name = "txtReadOnly"
        Me.txtReadOnly.ReadOnly = True
        Me.txtReadOnly.Size = New System.Drawing.Size(25, 20)
        Me.txtReadOnly.TabIndex = 76
        Me.txtReadOnly.TabStop = False
        Me.txtReadOnly.Visible = False
        '
        'TxtPrimaryPihakAcuan
        '
        Me.TxtPrimaryPihakAcuan.Location = New System.Drawing.Point(43, 52)
        Me.TxtPrimaryPihakAcuan.Name = "TxtPrimaryPihakAcuan"
        Me.TxtPrimaryPihakAcuan.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryPihakAcuan.TabIndex = 69
        Me.TxtPrimaryPihakAcuan.TabStop = False
        Me.TxtPrimaryPihakAcuan.Visible = False
        '
        'txtDONo
        '
        Me.txtDONo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDONo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDONo.Location = New System.Drawing.Point(85, 6)
        Me.txtDONo.MaxLength = 50
        Me.txtDONo.Name = "txtDONo"
        Me.txtDONo.Size = New System.Drawing.Size(312, 20)
        Me.txtDONo.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 75
        Me.Label1.Text = "DO. No"
        '
        'btnFindDO
        '
        Me.btnFindDO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindDO.Location = New System.Drawing.Point(400, 6)
        Me.btnFindDO.Name = "btnFindDO"
        Me.btnFindDO.Size = New System.Drawing.Size(28, 20)
        Me.btnFindDO.TabIndex = 1
        Me.btnFindDO.TabStop = False
        Me.btnFindDO.Text = "...."
        Me.btnFindDO.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(8, 112)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(870, 358)
        Me.TabControl1.TabIndex = 7
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.dgvItem)
        Me.TabPage1.Controls.Add(Me.dgvItemDO)
        Me.TabPage1.Controls.Add(Me.TextBox2)
        Me.TabPage1.Controls.Add(Me.lblSecondaryUnits)
        Me.TabPage1.Controls.Add(Me.gboItem)
        Me.TabPage1.Controls.Add(Me.chkBatchNoControlled)
        Me.TabPage1.Controls.Add(Me.Button1)
        Me.TabPage1.Controls.Add(Me.chkPdoControlled)
        Me.TabPage1.Controls.Add(Me.txtPrimaryDO)
        Me.TabPage1.Controls.Add(Me.TxtPrimaryStockIn)
        Me.TabPage1.Controls.Add(Me.cboSecondaryUnits)
        Me.TabPage1.Controls.Add(Me.lblSecQty)
        Me.TabPage1.Controls.Add(Me.TxtPrimaryCodeNo)
        Me.TabPage1.Controls.Add(Me.txtPrimarySO)
        Me.TabPage1.Controls.Add(Me.chkUseSecondaryUnit)
        Me.TabPage1.Controls.Add(Me.chkLocationControlled)
        Me.TabPage1.Controls.Add(Me.chkExpiryDateControlled)
        Me.TabPage1.Controls.Add(Me.txtPrimaryCategory)
        Me.TabPage1.Controls.Add(Me.txtPrimarySecondaryUnits)
        Me.TabPage1.Controls.Add(Me.txtPrimaryLocation)
        Me.TabPage1.Controls.Add(Me.txtPrimaryPDO)
        Me.TabPage1.Controls.Add(Me.txtSecondaryQty)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage1.Size = New System.Drawing.Size(862, 332)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Item Data"
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
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColPONoDONo, Me.ColSuppDONo, Me.ColCodeNo, Me.ColDescription, Me.ColQtyMulti, Me.ColUnitsCodeMulti, Me.ColRatioDesc, Me.ColQty, Me.ColUnit, Me.ColSecondaryQty, Me.ColSecondaryUnit, Me.ColBatch, Me.ColExpiredDate, Me.ColLocation, Me.ColWareHouse, Me.ColRemark, Me.ColPrimaryPOnoDoNo, Me.ColPrimaryCodeNo, Me.ColPrimaryLocation, Me.ColPrimaryUnits, Me.ColPrimarySecondaryUnits, Me.ColPrimaryPdO, Me.ColRatioDefault, Me.ColRatioReference, Me.ColPrimaryUnitsMulti})
        Me.dgvItem.Location = New System.Drawing.Point(5, 147)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(852, 179)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 8
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
        'ColPONoDONo
        '
        Me.ColPONoDONo.Frozen = True
        Me.ColPONoDONo.HeaderText = "SO. No"
        Me.ColPONoDONo.Name = "ColPONoDONo"
        Me.ColPONoDONo.Width = 67
        '
        'ColSuppDONo
        '
        Me.ColSuppDONo.Frozen = True
        Me.ColSuppDONo.HeaderText = "Supp. DO No"
        Me.ColSuppDONo.Name = "ColSuppDONo"
        Me.ColSuppDONo.Visible = False
        Me.ColSuppDONo.Width = 96
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
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.Width = 85
        '
        'ColQtyMulti
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyMulti.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyMulti.HeaderText = "Qty"
        Me.ColQtyMulti.Name = "ColQtyMulti"
        Me.ColQtyMulti.Width = 48
        '
        'ColUnitsCodeMulti
        '
        Me.ColUnitsCodeMulti.HeaderText = "Unit"
        Me.ColUnitsCodeMulti.Name = "ColUnitsCodeMulti"
        Me.ColUnitsCodeMulti.Width = 51
        '
        'ColRatioDesc
        '
        Me.ColRatioDesc.HeaderText = "Ratio"
        Me.ColRatioDesc.Name = "ColRatioDesc"
        Me.ColRatioDesc.Width = 57
        '
        'ColQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColQty.HeaderText = "Qty Default"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.Width = 85
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit Default"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 88
        '
        'ColSecondaryQty
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSecondaryQty.DefaultCellStyle = DataGridViewCellStyle4
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
        'ColRemark
        '
        Me.ColRemark.HeaderText = "Remarks"
        Me.ColRemark.Name = "ColRemark"
        Me.ColRemark.Width = 74
        '
        'ColPrimaryPOnoDoNo
        '
        Me.ColPrimaryPOnoDoNo.HeaderText = "PrimaryPOno"
        Me.ColPrimaryPOnoDoNo.Name = "ColPrimaryPOnoDoNo"
        Me.ColPrimaryPOnoDoNo.Visible = False
        Me.ColPrimaryPOnoDoNo.Width = 93
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
        'ColPrimaryPdO
        '
        Me.ColPrimaryPdO.HeaderText = "Primary PdO"
        Me.ColPrimaryPdO.Name = "ColPrimaryPdO"
        Me.ColPrimaryPdO.Visible = False
        Me.ColPrimaryPdO.Width = 90
        '
        'ColRatioDefault
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRatioDefault.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColRatioDefault.HeaderText = "Ratio Default"
        Me.ColRatioDefault.Name = "ColRatioDefault"
        Me.ColRatioDefault.Visible = False
        Me.ColRatioDefault.Width = 94
        '
        'ColRatioReference
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRatioReference.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColRatioReference.HeaderText = "Ratio Reference"
        Me.ColRatioReference.Name = "ColRatioReference"
        Me.ColRatioReference.Visible = False
        Me.ColRatioReference.Width = 110
        '
        'ColPrimaryUnitsMulti
        '
        Me.ColPrimaryUnitsMulti.HeaderText = "Primary Units Multi"
        Me.ColPrimaryUnitsMulti.Name = "ColPrimaryUnitsMulti"
        Me.ColPrimaryUnitsMulti.Visible = False
        Me.ColPrimaryUnitsMulti.Width = 118
        '
        'dgvItemDO
        '
        Me.dgvItemDO.AllowUserToAddRows = False
        Me.dgvItemDO.AllowUserToDeleteRows = False
        Me.dgvItemDO.AllowUserToResizeColumns = False
        Me.dgvItemDO.AllowUserToResizeRows = False
        DataGridViewCellStyle7.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItemDO.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle7
        Me.dgvItemDO.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemDO.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemDO.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItemDO.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDO2_No, Me.ColDO2_CodeNo, Me.ColDO2_Description, Me.ColDO2_CustCodeNo, Me.ColDO2_CustDesc, Me.ColDO2_SONO, Me.ColDO2_BatchNo, Me.ColDO2_ExpiryDate, Me.ColDO2_QtyMaster, Me.ColDO2_QtyPerMaster, Me.ColDO2_MasterUnits, Me.ColDO2_Qty, Me.ColDO2_TotalQtyMulti, Me.ColDO2_UnitsCodeMulti, Me.ColDO2_RatioDesc, Me.ColDO2_TotalQty, Me.ColDO2_Unit, Me.ColDO2_SecQty, Me.ColDO2_SecUnit, Me.ColDO2_Remarks, Me.ColDO2_PrimaryCodeNo, Me.ColDO2_PrimarySO, Me.ColDO2_PrimaryMasterUnits, Me.ColDO2_PrimaryUnit, Me.ColDO2_PrimarySecUnit, Me.ColDO2_RatioReference, Me.ColDO2_RatioDefault, Me.ColDO2_PrimaryUnitsMulti, Me.ColDO2_Location, Me.ColDO2_Warehouse, Me.ColDO2_PrimaryLocation})
        Me.dgvItemDO.Location = New System.Drawing.Point(281, 216)
        Me.dgvItemDO.MultiSelect = False
        Me.dgvItemDO.Name = "dgvItemDO"
        Me.dgvItemDO.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItemDO.RowTemplate.Height = 24
        Me.dgvItemDO.Size = New System.Drawing.Size(300, 72)
        Me.dgvItemDO.StandardTab = True
        Me.dgvItemDO.TabIndex = 99
        '
        'ColDO2_No
        '
        Me.ColDO2_No.Frozen = True
        Me.ColDO2_No.HeaderText = "No"
        Me.ColDO2_No.Name = "ColDO2_No"
        Me.ColDO2_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDO2_No.Width = 27
        '
        'ColDO2_CodeNo
        '
        Me.ColDO2_CodeNo.Frozen = True
        Me.ColDO2_CodeNo.HeaderText = "Code No"
        Me.ColDO2_CodeNo.Name = "ColDO2_CodeNo"
        Me.ColDO2_CodeNo.Width = 74
        '
        'ColDO2_Description
        '
        Me.ColDO2_Description.Frozen = True
        Me.ColDO2_Description.HeaderText = "Description"
        Me.ColDO2_Description.Name = "ColDO2_Description"
        Me.ColDO2_Description.Width = 85
        '
        'ColDO2_CustCodeNo
        '
        Me.ColDO2_CustCodeNo.HeaderText = "Customer's Code No"
        Me.ColDO2_CustCodeNo.Name = "ColDO2_CustCodeNo"
        Me.ColDO2_CustCodeNo.Visible = False
        Me.ColDO2_CustCodeNo.Width = 162
        '
        'ColDO2_CustDesc
        '
        Me.ColDO2_CustDesc.HeaderText = "Customer's Item Description"
        Me.ColDO2_CustDesc.Name = "ColDO2_CustDesc"
        Me.ColDO2_CustDesc.Visible = False
        Me.ColDO2_CustDesc.Width = 208
        '
        'ColDO2_SONO
        '
        Me.ColDO2_SONO.HeaderText = "SO. No"
        Me.ColDO2_SONO.Name = "ColDO2_SONO"
        Me.ColDO2_SONO.Width = 67
        '
        'ColDO2_BatchNo
        '
        Me.ColDO2_BatchNo.HeaderText = "Batch No"
        Me.ColDO2_BatchNo.Name = "ColDO2_BatchNo"
        Me.ColDO2_BatchNo.Width = 77
        '
        'ColDO2_ExpiryDate
        '
        Me.ColDO2_ExpiryDate.HeaderText = "Expiry Date"
        Me.ColDO2_ExpiryDate.Name = "ColDO2_ExpiryDate"
        Me.ColDO2_ExpiryDate.Width = 86
        '
        'ColDO2_QtyMaster
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDO2_QtyMaster.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColDO2_QtyMaster.HeaderText = "Qty Master"
        Me.ColDO2_QtyMaster.Name = "ColDO2_QtyMaster"
        Me.ColDO2_QtyMaster.Visible = False
        Me.ColDO2_QtyMaster.Width = 102
        '
        'ColDO2_QtyPerMaster
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDO2_QtyPerMaster.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColDO2_QtyPerMaster.HeaderText = "Qty Per Master"
        Me.ColDO2_QtyPerMaster.Name = "ColDO2_QtyPerMaster"
        Me.ColDO2_QtyPerMaster.Visible = False
        Me.ColDO2_QtyPerMaster.Width = 128
        '
        'ColDO2_MasterUnits
        '
        Me.ColDO2_MasterUnits.HeaderText = "Master Unit"
        Me.ColDO2_MasterUnits.Name = "ColDO2_MasterUnits"
        Me.ColDO2_MasterUnits.Visible = False
        Me.ColDO2_MasterUnits.Width = 105
        '
        'ColDO2_Qty
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDO2_Qty.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColDO2_Qty.HeaderText = "Qty"
        Me.ColDO2_Qty.Name = "ColDO2_Qty"
        Me.ColDO2_Qty.Visible = False
        Me.ColDO2_Qty.Width = 55
        '
        'ColDO2_TotalQtyMulti
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDO2_TotalQtyMulti.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColDO2_TotalQtyMulti.HeaderText = "Total Qty"
        Me.ColDO2_TotalQtyMulti.Name = "ColDO2_TotalQtyMulti"
        Me.ColDO2_TotalQtyMulti.Width = 75
        '
        'ColDO2_UnitsCodeMulti
        '
        Me.ColDO2_UnitsCodeMulti.HeaderText = "Units"
        Me.ColDO2_UnitsCodeMulti.Name = "ColDO2_UnitsCodeMulti"
        Me.ColDO2_UnitsCodeMulti.Width = 56
        '
        'ColDO2_RatioDesc
        '
        Me.ColDO2_RatioDesc.HeaderText = "Ratio"
        Me.ColDO2_RatioDesc.Name = "ColDO2_RatioDesc"
        Me.ColDO2_RatioDesc.Width = 38
        '
        'ColDO2_TotalQty
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDO2_TotalQty.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColDO2_TotalQty.HeaderText = "Qty Default"
        Me.ColDO2_TotalQty.Name = "ColDO2_TotalQty"
        Me.ColDO2_TotalQty.Width = 85
        '
        'ColDO2_Unit
        '
        Me.ColDO2_Unit.HeaderText = "Units Default"
        Me.ColDO2_Unit.Name = "ColDO2_Unit"
        Me.ColDO2_Unit.Width = 93
        '
        'ColDO2_SecQty
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDO2_SecQty.DefaultCellStyle = DataGridViewCellStyle13
        Me.ColDO2_SecQty.HeaderText = "@"
        Me.ColDO2_SecQty.Name = "ColDO2_SecQty"
        Me.ColDO2_SecQty.Width = 43
        '
        'ColDO2_SecUnit
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDO2_SecUnit.DefaultCellStyle = DataGridViewCellStyle14
        Me.ColDO2_SecUnit.HeaderText = "@ Unit"
        Me.ColDO2_SecUnit.Name = "ColDO2_SecUnit"
        Me.ColDO2_SecUnit.Width = 65
        '
        'ColDO2_Remarks
        '
        Me.ColDO2_Remarks.HeaderText = "Remarks"
        Me.ColDO2_Remarks.Name = "ColDO2_Remarks"
        Me.ColDO2_Remarks.Width = 74
        '
        'ColDO2_PrimaryCodeNo
        '
        Me.ColDO2_PrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColDO2_PrimaryCodeNo.Name = "ColDO2_PrimaryCodeNo"
        Me.ColDO2_PrimaryCodeNo.Visible = False
        Me.ColDO2_PrimaryCodeNo.Width = 132
        '
        'ColDO2_PrimarySO
        '
        Me.ColDO2_PrimarySO.HeaderText = "PrimarySO"
        Me.ColDO2_PrimarySO.Name = "ColDO2_PrimarySO"
        Me.ColDO2_PrimarySO.Visible = False
        Me.ColDO2_PrimarySO.Width = 101
        '
        'ColDO2_PrimaryMasterUnits
        '
        Me.ColDO2_PrimaryMasterUnits.HeaderText = "Primary Master Unit"
        Me.ColDO2_PrimaryMasterUnits.Name = "ColDO2_PrimaryMasterUnits"
        Me.ColDO2_PrimaryMasterUnits.Visible = False
        Me.ColDO2_PrimaryMasterUnits.Width = 157
        '
        'ColDO2_PrimaryUnit
        '
        Me.ColDO2_PrimaryUnit.HeaderText = "Primary Units"
        Me.ColDO2_PrimaryUnit.Name = "ColDO2_PrimaryUnit"
        Me.ColDO2_PrimaryUnit.Visible = False
        Me.ColDO2_PrimaryUnit.Width = 117
        '
        'ColDO2_PrimarySecUnit
        '
        Me.ColDO2_PrimarySecUnit.HeaderText = "Primary Sec Unit"
        Me.ColDO2_PrimarySecUnit.Name = "ColDO2_PrimarySecUnit"
        Me.ColDO2_PrimarySecUnit.Visible = False
        Me.ColDO2_PrimarySecUnit.Width = 138
        '
        'ColDO2_RatioReference
        '
        Me.ColDO2_RatioReference.HeaderText = "ColRatioReference"
        Me.ColDO2_RatioReference.Name = "ColDO2_RatioReference"
        Me.ColDO2_RatioReference.Visible = False
        Me.ColDO2_RatioReference.Width = 152
        '
        'ColDO2_RatioDefault
        '
        Me.ColDO2_RatioDefault.HeaderText = "Ratio Default"
        Me.ColDO2_RatioDefault.Name = "ColDO2_RatioDefault"
        Me.ColDO2_RatioDefault.Visible = False
        Me.ColDO2_RatioDefault.Width = 115
        '
        'ColDO2_PrimaryUnitsMulti
        '
        Me.ColDO2_PrimaryUnitsMulti.HeaderText = "Primary Units Multi"
        Me.ColDO2_PrimaryUnitsMulti.Name = "ColDO2_PrimaryUnitsMulti"
        Me.ColDO2_PrimaryUnitsMulti.Visible = False
        Me.ColDO2_PrimaryUnitsMulti.Width = 150
        '
        'ColDO2_Location
        '
        Me.ColDO2_Location.HeaderText = "Location"
        Me.ColDO2_Location.Name = "ColDO2_Location"
        Me.ColDO2_Location.Width = 73
        '
        'ColDO2_Warehouse
        '
        Me.ColDO2_Warehouse.HeaderText = "Warehouse"
        Me.ColDO2_Warehouse.Name = "ColDO2_Warehouse"
        Me.ColDO2_Warehouse.Width = 87
        '
        'ColDO2_PrimaryLocation
        '
        Me.ColDO2_PrimaryLocation.HeaderText = "Primary Location"
        Me.ColDO2_PrimaryLocation.Name = "ColDO2_PrimaryLocation"
        Me.ColDO2_PrimaryLocation.Visible = False
        Me.ColDO2_PrimaryLocation.Width = 139
        '
        'TextBox2
        '
        Me.TextBox2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox2.Location = New System.Drawing.Point(468, 166)
        Me.TextBox2.MaxLength = 50
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(25, 20)
        Me.TextBox2.TabIndex = 0
        Me.TextBox2.Visible = False
        '
        'lblSecondaryUnits
        '
        Me.lblSecondaryUnits.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSecondaryUnits.AutoSize = True
        Me.lblSecondaryUnits.Location = New System.Drawing.Point(830, 170)
        Me.lblSecondaryUnits.Name = "lblSecondaryUnits"
        Me.lblSecondaryUnits.Size = New System.Drawing.Size(31, 13)
        Me.lblSecondaryUnits.TabIndex = 86
        Me.lblSecondaryUnits.Text = "Units"
        Me.lblSecondaryUnits.Visible = False
        '
        'gboItem
        '
        Me.gboItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboItem.Controls.Add(Me.LinkLabel1)
        Me.gboItem.Controls.Add(Me.SplitContainer2)
        Me.gboItem.Location = New System.Drawing.Point(6, 4)
        Me.gboItem.Name = "gboItem"
        Me.gboItem.Size = New System.Drawing.Size(850, 145)
        Me.gboItem.TabIndex = 7
        Me.gboItem.TabStop = False
        '
        'LinkLabel1
        '
        Me.LinkLabel1.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(804, -4)
        Me.LinkLabel1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(50, 13)
        Me.LinkLabel1.TabIndex = 197
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Advance"
        Me.LinkLabel1.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 7)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtSalesOrderNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnFindSalesOrder)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblBatchNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label10)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblExpiryDate)
        Me.SplitContainer2.Panel1.Controls.Add(Me.BtnFindCodeNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label12)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label19)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.DtpExpired)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtDesc)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cboBatchNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer2.Panel1.Controls.Add(Me.BtnFindPONo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtCodeNo)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.LblQty)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtPrimaryUnitsMulti)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtQtyperBatchMulti)
        Me.SplitContainer2.Panel2.Controls.Add(Me.lblUnitsMulti)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRatioDefault)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRatioReference)
        Me.SplitContainer2.Panel2.Controls.Add(Me.lblWarehouse)
        Me.SplitContainer2.Panel2.Controls.Add(Me.lblLocation)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TextBox1)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TextBox3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.BtnInput)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtWarehouse)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label16)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRemarkItem)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtPrimaryUnitsDefault)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtQtyperBatchDefault)
        Me.SplitContainer2.Panel2.Controls.Add(Me.lblUnitsDefault)
        Me.SplitContainer2.Panel2.Controls.Add(Me.cboLocation)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label15)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label14)
        Me.SplitContainer2.Size = New System.Drawing.Size(849, 130)
        Me.SplitContainer2.SplitterDistance = 420
        Me.SplitContainer2.TabIndex = 9
        Me.SplitContainer2.TabStop = False
        '
        'txtSalesOrderNo
        '
        Me.txtSalesOrderNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSalesOrderNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSalesOrderNo.Location = New System.Drawing.Point(90, 6)
        Me.txtSalesOrderNo.MaxLength = 50
        Me.txtSalesOrderNo.Name = "txtSalesOrderNo"
        Me.txtSalesOrderNo.Size = New System.Drawing.Size(284, 20)
        Me.txtSalesOrderNo.TabIndex = 0
        '
        'btnFindSalesOrder
        '
        Me.btnFindSalesOrder.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindSalesOrder.Location = New System.Drawing.Point(382, 6)
        Me.btnFindSalesOrder.Name = "btnFindSalesOrder"
        Me.btnFindSalesOrder.Size = New System.Drawing.Size(32, 20)
        Me.btnFindSalesOrder.TabIndex = 1
        Me.btnFindSalesOrder.TabStop = False
        Me.btnFindSalesOrder.Text = "...."
        Me.btnFindSalesOrder.UseVisualStyleBackColor = True
        '
        'lblBatchNo
        '
        Me.lblBatchNo.AutoSize = True
        Me.lblBatchNo.Location = New System.Drawing.Point(90, 83)
        Me.lblBatchNo.Name = "lblBatchNo"
        Me.lblBatchNo.Size = New System.Drawing.Size(24, 13)
        Me.lblBatchNo.TabIndex = 101
        Me.lblBatchNo.Text = "n/a"
        Me.lblBatchNo.Visible = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(5, 34)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(49, 13)
        Me.Label10.TabIndex = 52
        Me.Label10.Text = "Code No"
        '
        'lblExpiryDate
        '
        Me.lblExpiryDate.AutoSize = True
        Me.lblExpiryDate.Location = New System.Drawing.Point(90, 109)
        Me.lblExpiryDate.Name = "lblExpiryDate"
        Me.lblExpiryDate.Size = New System.Drawing.Size(24, 13)
        Me.lblExpiryDate.TabIndex = 97
        Me.lblExpiryDate.Text = "n/a"
        Me.lblExpiryDate.Visible = False
        '
        'BtnFindCodeNo
        '
        Me.BtnFindCodeNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindCodeNo.Location = New System.Drawing.Point(382, 30)
        Me.BtnFindCodeNo.Name = "BtnFindCodeNo"
        Me.BtnFindCodeNo.Size = New System.Drawing.Size(32, 20)
        Me.BtnFindCodeNo.TabIndex = 3
        Me.BtnFindCodeNo.TabStop = False
        Me.BtnFindCodeNo.Text = "...."
        Me.BtnFindCodeNo.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(5, 58)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(60, 13)
        Me.Label11.TabIndex = 62
        Me.Label11.Text = "Description"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(5, 109)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(61, 13)
        Me.Label12.TabIndex = 77
        Me.Label12.Text = "Expiry Date"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(5, 83)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(52, 13)
        Me.Label19.TabIndex = 79
        Me.Label19.Text = "Batch No"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 10)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(79, 13)
        Me.Label2.TabIndex = 79
        Me.Label2.Text = "Sales Order No"
        '
        'DtpExpired
        '
        Me.DtpExpired.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpExpired.CustomFormat = "ddd, dd MMM yyyy"
        Me.DtpExpired.Enabled = False
        Me.DtpExpired.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpExpired.Location = New System.Drawing.Point(90, 105)
        Me.DtpExpired.Name = "DtpExpired"
        Me.DtpExpired.Size = New System.Drawing.Size(324, 20)
        Me.DtpExpired.TabIndex = 6
        '
        'txtDesc
        '
        Me.txtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDesc.Location = New System.Drawing.Point(90, 55)
        Me.txtDesc.MaxLength = 6
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(324, 20)
        Me.txtDesc.TabIndex = 4
        Me.txtDesc.TabStop = False
        '
        'cboBatchNo
        '
        Me.cboBatchNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboBatchNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBatchNo.FormattingEnabled = True
        Me.cboBatchNo.Location = New System.Drawing.Point(90, 79)
        Me.cboBatchNo.Name = "cboBatchNo"
        Me.cboBatchNo.Size = New System.Drawing.Size(324, 21)
        Me.cboBatchNo.TabIndex = 5
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(5, 34)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(39, 13)
        Me.Label9.TabIndex = 48
        Me.Label9.Text = "PO No"
        Me.Label9.Visible = False
        '
        'BtnFindPONo
        '
        Me.BtnFindPONo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindPONo.Location = New System.Drawing.Point(382, 30)
        Me.BtnFindPONo.Name = "BtnFindPONo"
        Me.BtnFindPONo.Size = New System.Drawing.Size(32, 20)
        Me.BtnFindPONo.TabIndex = 1
        Me.BtnFindPONo.Text = "...."
        Me.BtnFindPONo.UseVisualStyleBackColor = True
        Me.BtnFindPONo.Visible = False
        '
        'TxtCodeNo
        '
        Me.TxtCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeNo.Location = New System.Drawing.Point(90, 30)
        Me.TxtCodeNo.MaxLength = 50
        Me.TxtCodeNo.Name = "TxtCodeNo"
        Me.TxtCodeNo.Size = New System.Drawing.Size(284, 20)
        Me.TxtCodeNo.TabIndex = 2
        '
        'LblQty
        '
        Me.LblQty.AutoSize = True
        Me.LblQty.Location = New System.Drawing.Point(5, 58)
        Me.LblQty.Name = "LblQty"
        Me.LblQty.Size = New System.Drawing.Size(23, 13)
        Me.LblQty.TabIndex = 202
        Me.LblQty.Text = "Qty"
        '
        'txtPrimaryUnitsMulti
        '
        Me.txtPrimaryUnitsMulti.Location = New System.Drawing.Point(190, 105)
        Me.txtPrimaryUnitsMulti.Name = "txtPrimaryUnitsMulti"
        Me.txtPrimaryUnitsMulti.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryUnitsMulti.TabIndex = 201
        Me.txtPrimaryUnitsMulti.TabStop = False
        Me.txtPrimaryUnitsMulti.Text = "-1"
        Me.txtPrimaryUnitsMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryUnitsMulti.Visible = False
        '
        'TxtQtyperBatchMulti
        '
        Me.TxtQtyperBatchMulti.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatchMulti.Location = New System.Drawing.Point(87, 54)
        Me.TxtQtyperBatchMulti.MaxLength = 6
        Me.TxtQtyperBatchMulti.Name = "TxtQtyperBatchMulti"
        Me.TxtQtyperBatchMulti.Size = New System.Drawing.Size(102, 20)
        Me.TxtQtyperBatchMulti.TabIndex = 2
        Me.TxtQtyperBatchMulti.Text = "0"
        Me.TxtQtyperBatchMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblUnitsMulti
        '
        Me.lblUnitsMulti.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnitsMulti.AutoSize = True
        Me.lblUnitsMulti.Location = New System.Drawing.Point(194, 58)
        Me.lblUnitsMulti.Name = "lblUnitsMulti"
        Me.lblUnitsMulti.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitsMulti.TabIndex = 200
        Me.lblUnitsMulti.Text = "Units"
        '
        'TxtRatioDefault
        '
        Me.TxtRatioDefault.Location = New System.Drawing.Point(124, 105)
        Me.TxtRatioDefault.Name = "TxtRatioDefault"
        Me.TxtRatioDefault.Size = New System.Drawing.Size(32, 20)
        Me.TxtRatioDefault.TabIndex = 198
        Me.TxtRatioDefault.Text = "1"
        Me.TxtRatioDefault.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtRatioDefault.Visible = False
        '
        'TxtRatioReference
        '
        Me.TxtRatioReference.Location = New System.Drawing.Point(87, 105)
        Me.TxtRatioReference.Name = "TxtRatioReference"
        Me.TxtRatioReference.Size = New System.Drawing.Size(32, 20)
        Me.TxtRatioReference.TabIndex = 197
        Me.TxtRatioReference.Text = "1"
        Me.TxtRatioReference.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtRatioReference.Visible = False
        '
        'lblWarehouse
        '
        Me.lblWarehouse.AutoSize = True
        Me.lblWarehouse.Location = New System.Drawing.Point(84, 34)
        Me.lblWarehouse.Name = "lblWarehouse"
        Me.lblWarehouse.Size = New System.Drawing.Size(24, 13)
        Me.lblWarehouse.TabIndex = 100
        Me.lblWarehouse.Text = "n/a"
        Me.lblWarehouse.Visible = False
        '
        'lblLocation
        '
        Me.lblLocation.AutoSize = True
        Me.lblLocation.Location = New System.Drawing.Point(84, 10)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.Size = New System.Drawing.Size(24, 13)
        Me.lblLocation.TabIndex = 99
        Me.lblLocation.Text = "n/a"
        Me.lblLocation.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(-275, 0)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(25, 20)
        Me.TextBox1.TabIndex = 78
        Me.TextBox1.TabStop = False
        Me.TextBox1.Text = "-1"
        Me.TextBox1.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(-364, 58)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 62
        Me.Label7.Text = "Description"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(-279, 54)
        Me.TextBox3.MaxLength = 6
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ReadOnly = True
        Me.TextBox3.Size = New System.Drawing.Size(52, 20)
        Me.TextBox3.TabIndex = 2
        Me.TextBox3.TabStop = False
        '
        'BtnInput
        '
        Me.BtnInput.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnInput.Location = New System.Drawing.Point(364, 78)
        Me.BtnInput.Name = "BtnInput"
        Me.BtnInput.Size = New System.Drawing.Size(56, 22)
        Me.BtnInput.TabIndex = 5
        Me.BtnInput.Text = "&Input"
        Me.BtnInput.UseVisualStyleBackColor = True
        '
        'TxtWarehouse
        '
        Me.TxtWarehouse.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtWarehouse.Location = New System.Drawing.Point(87, 30)
        Me.TxtWarehouse.MaxLength = 50
        Me.TxtWarehouse.Name = "TxtWarehouse"
        Me.TxtWarehouse.ReadOnly = True
        Me.TxtWarehouse.Size = New System.Drawing.Size(332, 20)
        Me.TxtWarehouse.TabIndex = 1
        Me.TxtWarehouse.TabStop = False
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(5, 83)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(49, 13)
        Me.Label16.TabIndex = 63
        Me.Label16.Text = "Remarks"
        '
        'TxtRemarkItem
        '
        Me.TxtRemarkItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarkItem.Location = New System.Drawing.Point(87, 79)
        Me.TxtRemarkItem.MaxLength = 50
        Me.TxtRemarkItem.Name = "TxtRemarkItem"
        Me.TxtRemarkItem.Size = New System.Drawing.Size(270, 20)
        Me.TxtRemarkItem.TabIndex = 4
        '
        'txtPrimaryUnitsDefault
        '
        Me.txtPrimaryUnitsDefault.Location = New System.Drawing.Point(221, 105)
        Me.txtPrimaryUnitsDefault.Name = "txtPrimaryUnitsDefault"
        Me.txtPrimaryUnitsDefault.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryUnitsDefault.TabIndex = 87
        Me.txtPrimaryUnitsDefault.TabStop = False
        Me.txtPrimaryUnitsDefault.Text = "-1"
        Me.txtPrimaryUnitsDefault.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryUnitsDefault.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(-329, 4)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(79, 13)
        Me.Label5.TabIndex = 79
        Me.Label5.Text = "Sales Order No"
        '
        'TxtQtyperBatchDefault
        '
        Me.TxtQtyperBatchDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatchDefault.Location = New System.Drawing.Point(251, 54)
        Me.TxtQtyperBatchDefault.MaxLength = 6
        Me.TxtQtyperBatchDefault.Name = "TxtQtyperBatchDefault"
        Me.TxtQtyperBatchDefault.Size = New System.Drawing.Size(105, 20)
        Me.TxtQtyperBatchDefault.TabIndex = 3
        Me.TxtQtyperBatchDefault.Text = "0"
        Me.TxtQtyperBatchDefault.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblUnitsDefault
        '
        Me.lblUnitsDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnitsDefault.AutoSize = True
        Me.lblUnitsDefault.Location = New System.Drawing.Point(362, 58)
        Me.lblUnitsDefault.Name = "lblUnitsDefault"
        Me.lblUnitsDefault.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitsDefault.TabIndex = 84
        Me.lblUnitsDefault.Text = "Units"
        '
        'cboLocation
        '
        Me.cboLocation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboLocation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLocation.FormattingEnabled = True
        Me.cboLocation.Location = New System.Drawing.Point(87, 6)
        Me.cboLocation.Name = "cboLocation"
        Me.cboLocation.Size = New System.Drawing.Size(332, 21)
        Me.cboLocation.TabIndex = 0
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(5, 10)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(48, 13)
        Me.Label15.TabIndex = 73
        Me.Label15.Text = "Location"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(5, 34)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(62, 13)
        Me.Label14.TabIndex = 65
        Me.Label14.Text = "Warehouse"
        '
        'chkBatchNoControlled
        '
        Me.chkBatchNoControlled.AutoSize = True
        Me.chkBatchNoControlled.Location = New System.Drawing.Point(161, 168)
        Me.chkBatchNoControlled.Name = "chkBatchNoControlled"
        Me.chkBatchNoControlled.Size = New System.Drawing.Size(71, 17)
        Me.chkBatchNoControlled.TabIndex = 91
        Me.chkBatchNoControlled.TabStop = False
        Me.chkBatchNoControlled.Text = "Batch No"
        Me.chkBatchNoControlled.UseVisualStyleBackColor = True
        Me.chkBatchNoControlled.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(616, 166)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(25, 20)
        Me.Button1.TabIndex = 1
        Me.Button1.TabStop = False
        Me.Button1.Text = "...."
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'chkPdoControlled
        '
        Me.chkPdoControlled.AutoSize = True
        Me.chkPdoControlled.Location = New System.Drawing.Point(233, 168)
        Me.chkPdoControlled.Name = "chkPdoControlled"
        Me.chkPdoControlled.Size = New System.Drawing.Size(47, 17)
        Me.chkPdoControlled.TabIndex = 90
        Me.chkPdoControlled.TabStop = False
        Me.chkPdoControlled.Text = "PdO"
        Me.chkPdoControlled.UseVisualStyleBackColor = True
        Me.chkPdoControlled.Visible = False
        '
        'txtPrimaryDO
        '
        Me.txtPrimaryDO.Location = New System.Drawing.Point(437, 166)
        Me.txtPrimaryDO.Name = "txtPrimaryDO"
        Me.txtPrimaryDO.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryDO.TabIndex = 74
        Me.txtPrimaryDO.TabStop = False
        Me.txtPrimaryDO.Text = "-1"
        Me.txtPrimaryDO.Visible = False
        '
        'cboSecondaryUnits
        '
        Me.cboSecondaryUnits.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboSecondaryUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSecondaryUnits.FormattingEnabled = True
        Me.cboSecondaryUnits.Location = New System.Drawing.Point(760, 166)
        Me.cboSecondaryUnits.Name = "cboSecondaryUnits"
        Me.cboSecondaryUnits.Size = New System.Drawing.Size(70, 21)
        Me.cboSecondaryUnits.TabIndex = 8
        Me.cboSecondaryUnits.Visible = False
        '
        'lblSecQty
        '
        Me.lblSecQty.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSecQty.AutoSize = True
        Me.lblSecQty.Location = New System.Drawing.Point(736, 170)
        Me.lblSecQty.Name = "lblSecQty"
        Me.lblSecQty.Size = New System.Drawing.Size(18, 13)
        Me.lblSecQty.TabIndex = 98
        Me.lblSecQty.Text = "@"
        Me.lblSecQty.Visible = False
        '
        'TxtPrimaryCodeNo
        '
        Me.TxtPrimaryCodeNo.Location = New System.Drawing.Point(644, 166)
        Me.TxtPrimaryCodeNo.Name = "TxtPrimaryCodeNo"
        Me.TxtPrimaryCodeNo.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryCodeNo.TabIndex = 53
        Me.TxtPrimaryCodeNo.TabStop = False
        Me.TxtPrimaryCodeNo.Visible = False
        '
        'txtPrimarySO
        '
        Me.txtPrimarySO.Location = New System.Drawing.Point(529, 166)
        Me.txtPrimarySO.Name = "txtPrimarySO"
        Me.txtPrimarySO.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimarySO.TabIndex = 78
        Me.txtPrimarySO.TabStop = False
        Me.txtPrimarySO.Text = "-1"
        Me.txtPrimarySO.Visible = False
        '
        'chkUseSecondaryUnit
        '
        Me.chkUseSecondaryUnit.AutoSize = True
        Me.chkUseSecondaryUnit.Location = New System.Drawing.Point(282, 168)
        Me.chkUseSecondaryUnit.Name = "chkUseSecondaryUnit"
        Me.chkUseSecondaryUnit.Size = New System.Drawing.Size(121, 17)
        Me.chkUseSecondaryUnit.TabIndex = 89
        Me.chkUseSecondaryUnit.TabStop = False
        Me.chkUseSecondaryUnit.Text = "Use Secondary Unit"
        Me.chkUseSecondaryUnit.UseVisualStyleBackColor = True
        Me.chkUseSecondaryUnit.Visible = False
        '
        'chkLocationControlled
        '
        Me.chkLocationControlled.AutoSize = True
        Me.chkLocationControlled.Location = New System.Drawing.Point(92, 168)
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
        Me.chkExpiryDateControlled.Location = New System.Drawing.Point(12, 168)
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
        Me.txtPrimaryCategory.Location = New System.Drawing.Point(673, 166)
        Me.txtPrimaryCategory.Name = "txtPrimaryCategory"
        Me.txtPrimaryCategory.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryCategory.TabIndex = 95
        Me.txtPrimaryCategory.TabStop = False
        Me.txtPrimaryCategory.Visible = False
        '
        'txtPrimarySecondaryUnits
        '
        Me.txtPrimarySecondaryUnits.Location = New System.Drawing.Point(559, 166)
        Me.txtPrimarySecondaryUnits.Name = "txtPrimarySecondaryUnits"
        Me.txtPrimarySecondaryUnits.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimarySecondaryUnits.TabIndex = 88
        Me.txtPrimarySecondaryUnits.TabStop = False
        Me.txtPrimarySecondaryUnits.Visible = False
        '
        'txtPrimaryLocation
        '
        Me.txtPrimaryLocation.Location = New System.Drawing.Point(588, 166)
        Me.txtPrimaryLocation.Name = "txtPrimaryLocation"
        Me.txtPrimaryLocation.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryLocation.TabIndex = 94
        Me.txtPrimaryLocation.TabStop = False
        Me.txtPrimaryLocation.Visible = False
        '
        'txtPrimaryPDO
        '
        Me.txtPrimaryPDO.Location = New System.Drawing.Point(498, 166)
        Me.txtPrimaryPDO.Name = "txtPrimaryPDO"
        Me.txtPrimaryPDO.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryPDO.TabIndex = 96
        Me.txtPrimaryPDO.TabStop = False
        Me.txtPrimaryPDO.Visible = False
        '
        'txtSecondaryQty
        '
        Me.txtSecondaryQty.Location = New System.Drawing.Point(703, 166)
        Me.txtSecondaryQty.MaxLength = 30
        Me.txtSecondaryQty.Name = "txtSecondaryQty"
        Me.txtSecondaryQty.Size = New System.Drawing.Size(25, 20)
        Me.txtSecondaryQty.TabIndex = 9
        Me.txtSecondaryQty.TabStop = False
        Me.txtSecondaryQty.Text = "0"
        Me.txtSecondaryQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtSecondaryQty.Visible = False
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage2.Controls.Add(Me.DgvItemDO2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage2.Size = New System.Drawing.Size(862, 333)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Do Data"
        '
        'DgvItemDO2
        '
        Me.DgvItemDO2.AllowUserToAddRows = False
        Me.DgvItemDO2.AllowUserToDeleteRows = False
        Me.DgvItemDO2.AllowUserToResizeColumns = False
        Me.DgvItemDO2.AllowUserToResizeRows = False
        DataGridViewCellStyle15.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.DgvItemDO2.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle15
        Me.DgvItemDO2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvItemDO2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvItemDO2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDO_No, Me.ColDO_CodeNo, Me.ColDO_Description, Me.ColDO_CustCodeNo, Me.ColDO_CustDesc, Me.ColDO_SONO, Me.ColDO_BatchNo, Me.ColDO_ExpiryDate, Me.ColDO_QtyMaster, Me.ColDO_QtyPerMaster, Me.ColDO_MasterUnits, Me.ColDO_Qty, Me.ColDO_TotalQtyMulti, Me.ColDO_UnitsCodeMulti, Me.ColDO_RatioDesc, Me.ColDO_TotalQty, Me.ColDO_Unit, Me.ColDO_SecQty, Me.ColDO_SecUnit, Me.ColDO_Remarks, Me.ColDO_PrimaryCodeNo, Me.ColDO_PrimarySO, Me.ColDO_PrimaryMasterUnits, Me.ColDO_PrimaryUnit, Me.ColDO_PrimarySecUnit, Me.ColDO_RatioReference, Me.ColDO_RatioDefault, Me.ColDO_PrimaryUnitsMulti, Me.ColDO_Location, Me.ColDO_Warehouse, Me.ColDO_PrimaryLocation})
        Me.DgvItemDO2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvItemDO2.Location = New System.Drawing.Point(3, 3)
        Me.DgvItemDO2.MultiSelect = False
        Me.DgvItemDO2.Name = "DgvItemDO2"
        Me.DgvItemDO2.ReadOnly = True
        Me.DgvItemDO2.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.DgvItemDO2.RowTemplate.Height = 24
        Me.DgvItemDO2.Size = New System.Drawing.Size(856, 327)
        Me.DgvItemDO2.StandardTab = True
        Me.DgvItemDO2.TabIndex = 7
        '
        'ColDO_No
        '
        Me.ColDO_No.Frozen = True
        Me.ColDO_No.HeaderText = "No"
        Me.ColDO_No.Name = "ColDO_No"
        Me.ColDO_No.ReadOnly = True
        Me.ColDO_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDO_No.Width = 27
        '
        'ColDO_CodeNo
        '
        Me.ColDO_CodeNo.Frozen = True
        Me.ColDO_CodeNo.HeaderText = "Code No"
        Me.ColDO_CodeNo.Name = "ColDO_CodeNo"
        Me.ColDO_CodeNo.ReadOnly = True
        Me.ColDO_CodeNo.Width = 74
        '
        'ColDO_Description
        '
        Me.ColDO_Description.HeaderText = "Description"
        Me.ColDO_Description.Name = "ColDO_Description"
        Me.ColDO_Description.ReadOnly = True
        Me.ColDO_Description.Width = 85
        '
        'ColDO_CustCodeNo
        '
        Me.ColDO_CustCodeNo.HeaderText = "Customer's Code No"
        Me.ColDO_CustCodeNo.Name = "ColDO_CustCodeNo"
        Me.ColDO_CustCodeNo.ReadOnly = True
        Me.ColDO_CustCodeNo.Visible = False
        Me.ColDO_CustCodeNo.Width = 162
        '
        'ColDO_CustDesc
        '
        Me.ColDO_CustDesc.HeaderText = "Customer's Item Description"
        Me.ColDO_CustDesc.Name = "ColDO_CustDesc"
        Me.ColDO_CustDesc.ReadOnly = True
        Me.ColDO_CustDesc.Visible = False
        Me.ColDO_CustDesc.Width = 208
        '
        'ColDO_SONO
        '
        Me.ColDO_SONO.HeaderText = "SO. No"
        Me.ColDO_SONO.Name = "ColDO_SONO"
        Me.ColDO_SONO.ReadOnly = True
        Me.ColDO_SONO.Width = 67
        '
        'ColDO_BatchNo
        '
        Me.ColDO_BatchNo.HeaderText = "Batch No"
        Me.ColDO_BatchNo.Name = "ColDO_BatchNo"
        Me.ColDO_BatchNo.ReadOnly = True
        Me.ColDO_BatchNo.Width = 77
        '
        'ColDO_ExpiryDate
        '
        Me.ColDO_ExpiryDate.HeaderText = "Expiry Date"
        Me.ColDO_ExpiryDate.Name = "ColDO_ExpiryDate"
        Me.ColDO_ExpiryDate.ReadOnly = True
        Me.ColDO_ExpiryDate.Width = 86
        '
        'ColDO_QtyMaster
        '
        DataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDO_QtyMaster.DefaultCellStyle = DataGridViewCellStyle16
        Me.ColDO_QtyMaster.HeaderText = "Qty Master"
        Me.ColDO_QtyMaster.Name = "ColDO_QtyMaster"
        Me.ColDO_QtyMaster.ReadOnly = True
        Me.ColDO_QtyMaster.Visible = False
        Me.ColDO_QtyMaster.Width = 102
        '
        'ColDO_QtyPerMaster
        '
        DataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDO_QtyPerMaster.DefaultCellStyle = DataGridViewCellStyle17
        Me.ColDO_QtyPerMaster.HeaderText = "Qty Per Master"
        Me.ColDO_QtyPerMaster.Name = "ColDO_QtyPerMaster"
        Me.ColDO_QtyPerMaster.ReadOnly = True
        Me.ColDO_QtyPerMaster.Visible = False
        Me.ColDO_QtyPerMaster.Width = 128
        '
        'ColDO_MasterUnits
        '
        Me.ColDO_MasterUnits.HeaderText = "Master Unit"
        Me.ColDO_MasterUnits.Name = "ColDO_MasterUnits"
        Me.ColDO_MasterUnits.ReadOnly = True
        Me.ColDO_MasterUnits.Visible = False
        Me.ColDO_MasterUnits.Width = 105
        '
        'ColDO_Qty
        '
        DataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDO_Qty.DefaultCellStyle = DataGridViewCellStyle18
        Me.ColDO_Qty.HeaderText = "Qty"
        Me.ColDO_Qty.Name = "ColDO_Qty"
        Me.ColDO_Qty.ReadOnly = True
        Me.ColDO_Qty.Visible = False
        Me.ColDO_Qty.Width = 55
        '
        'ColDO_TotalQtyMulti
        '
        DataGridViewCellStyle19.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDO_TotalQtyMulti.DefaultCellStyle = DataGridViewCellStyle19
        Me.ColDO_TotalQtyMulti.HeaderText = "Total Qty"
        Me.ColDO_TotalQtyMulti.Name = "ColDO_TotalQtyMulti"
        Me.ColDO_TotalQtyMulti.ReadOnly = True
        Me.ColDO_TotalQtyMulti.Width = 75
        '
        'ColDO_UnitsCodeMulti
        '
        Me.ColDO_UnitsCodeMulti.HeaderText = "Units"
        Me.ColDO_UnitsCodeMulti.Name = "ColDO_UnitsCodeMulti"
        Me.ColDO_UnitsCodeMulti.ReadOnly = True
        Me.ColDO_UnitsCodeMulti.Width = 56
        '
        'ColDO_RatioDesc
        '
        Me.ColDO_RatioDesc.HeaderText = "Ratio"
        Me.ColDO_RatioDesc.Name = "ColDO_RatioDesc"
        Me.ColDO_RatioDesc.ReadOnly = True
        Me.ColDO_RatioDesc.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColDO_RatioDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDO_RatioDesc.Width = 38
        '
        'ColDO_TotalQty
        '
        DataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDO_TotalQty.DefaultCellStyle = DataGridViewCellStyle20
        Me.ColDO_TotalQty.HeaderText = "Qty Default"
        Me.ColDO_TotalQty.Name = "ColDO_TotalQty"
        Me.ColDO_TotalQty.ReadOnly = True
        Me.ColDO_TotalQty.Width = 85
        '
        'ColDO_Unit
        '
        Me.ColDO_Unit.HeaderText = "Units Default"
        Me.ColDO_Unit.Name = "ColDO_Unit"
        Me.ColDO_Unit.ReadOnly = True
        Me.ColDO_Unit.Width = 93
        '
        'ColDO_SecQty
        '
        DataGridViewCellStyle21.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDO_SecQty.DefaultCellStyle = DataGridViewCellStyle21
        Me.ColDO_SecQty.HeaderText = "@"
        Me.ColDO_SecQty.Name = "ColDO_SecQty"
        Me.ColDO_SecQty.ReadOnly = True
        Me.ColDO_SecQty.Width = 43
        '
        'ColDO_SecUnit
        '
        DataGridViewCellStyle22.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDO_SecUnit.DefaultCellStyle = DataGridViewCellStyle22
        Me.ColDO_SecUnit.HeaderText = "@ Unit"
        Me.ColDO_SecUnit.Name = "ColDO_SecUnit"
        Me.ColDO_SecUnit.ReadOnly = True
        Me.ColDO_SecUnit.Width = 65
        '
        'ColDO_Remarks
        '
        Me.ColDO_Remarks.HeaderText = "Remarks"
        Me.ColDO_Remarks.Name = "ColDO_Remarks"
        Me.ColDO_Remarks.ReadOnly = True
        Me.ColDO_Remarks.Width = 74
        '
        'ColDO_PrimaryCodeNo
        '
        Me.ColDO_PrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColDO_PrimaryCodeNo.Name = "ColDO_PrimaryCodeNo"
        Me.ColDO_PrimaryCodeNo.ReadOnly = True
        Me.ColDO_PrimaryCodeNo.Visible = False
        Me.ColDO_PrimaryCodeNo.Width = 132
        '
        'ColDO_PrimarySO
        '
        Me.ColDO_PrimarySO.HeaderText = "PrimarySO"
        Me.ColDO_PrimarySO.Name = "ColDO_PrimarySO"
        Me.ColDO_PrimarySO.ReadOnly = True
        Me.ColDO_PrimarySO.Visible = False
        Me.ColDO_PrimarySO.Width = 101
        '
        'ColDO_PrimaryMasterUnits
        '
        Me.ColDO_PrimaryMasterUnits.HeaderText = "Primary Master Unit"
        Me.ColDO_PrimaryMasterUnits.Name = "ColDO_PrimaryMasterUnits"
        Me.ColDO_PrimaryMasterUnits.ReadOnly = True
        Me.ColDO_PrimaryMasterUnits.Visible = False
        Me.ColDO_PrimaryMasterUnits.Width = 157
        '
        'ColDO_PrimaryUnit
        '
        Me.ColDO_PrimaryUnit.HeaderText = "Primary Units"
        Me.ColDO_PrimaryUnit.Name = "ColDO_PrimaryUnit"
        Me.ColDO_PrimaryUnit.ReadOnly = True
        Me.ColDO_PrimaryUnit.Visible = False
        Me.ColDO_PrimaryUnit.Width = 117
        '
        'ColDO_PrimarySecUnit
        '
        Me.ColDO_PrimarySecUnit.HeaderText = "Primary Sec Unit"
        Me.ColDO_PrimarySecUnit.Name = "ColDO_PrimarySecUnit"
        Me.ColDO_PrimarySecUnit.ReadOnly = True
        Me.ColDO_PrimarySecUnit.Visible = False
        Me.ColDO_PrimarySecUnit.Width = 138
        '
        'ColDO_RatioReference
        '
        Me.ColDO_RatioReference.HeaderText = "ColRatioReference"
        Me.ColDO_RatioReference.Name = "ColDO_RatioReference"
        Me.ColDO_RatioReference.ReadOnly = True
        Me.ColDO_RatioReference.Visible = False
        Me.ColDO_RatioReference.Width = 152
        '
        'ColDO_RatioDefault
        '
        Me.ColDO_RatioDefault.HeaderText = "Ratio Default"
        Me.ColDO_RatioDefault.Name = "ColDO_RatioDefault"
        Me.ColDO_RatioDefault.ReadOnly = True
        Me.ColDO_RatioDefault.Visible = False
        Me.ColDO_RatioDefault.Width = 115
        '
        'ColDO_PrimaryUnitsMulti
        '
        Me.ColDO_PrimaryUnitsMulti.HeaderText = "Primary Units Multi"
        Me.ColDO_PrimaryUnitsMulti.Name = "ColDO_PrimaryUnitsMulti"
        Me.ColDO_PrimaryUnitsMulti.ReadOnly = True
        Me.ColDO_PrimaryUnitsMulti.Visible = False
        Me.ColDO_PrimaryUnitsMulti.Width = 150
        '
        'ColDO_Location
        '
        Me.ColDO_Location.HeaderText = "Location"
        Me.ColDO_Location.Name = "ColDO_Location"
        Me.ColDO_Location.ReadOnly = True
        Me.ColDO_Location.Width = 73
        '
        'ColDO_Warehouse
        '
        Me.ColDO_Warehouse.HeaderText = "Warehouse"
        Me.ColDO_Warehouse.Name = "ColDO_Warehouse"
        Me.ColDO_Warehouse.ReadOnly = True
        Me.ColDO_Warehouse.Width = 87
        '
        'ColDO_PrimaryLocation
        '
        Me.ColDO_PrimaryLocation.HeaderText = "Primary Location"
        Me.ColDO_PrimaryLocation.Name = "ColDO_PrimaryLocation"
        Me.ColDO_PrimaryLocation.ReadOnly = True
        Me.ColDO_PrimaryLocation.Visible = False
        Me.ColDO_PrimaryLocation.Width = 139
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = CType(resources.GetObject("btnUP.BackgroundImage"), System.Drawing.Image)
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 476)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 10
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
        Me.btnDown.Location = New System.Drawing.Point(40, 476)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 11
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(446, 476)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 9
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(336, 476)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 8
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 550)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(894, 22)
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
        Me.btnDelete.Location = New System.Drawing.Point(451, 515)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(365, 515)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(192, 515)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(279, 515)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(537, 515)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddItemByCustom})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(185, 26)
        '
        'AddItemByCustom
        '
        Me.AddItemByCustom.Name = "AddItemByCustom"
        Me.AddItemByCustom.Size = New System.Drawing.Size(184, 22)
        Me.AddItemByCustom.Text = "Add Item By Custom"
        '
        'FrmTransaksiStockOutSales
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(894, 572)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MinimumSize = New System.Drawing.Size(900, 600)
        Me.Name = "FrmTransaksiStockOutSales"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stock Out - Sales"
        Me.gboDetail.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItemDO, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboItem.ResumeLayout(False)
        Me.gboItem.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DgvItemDO2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryStockIn As System.Windows.Forms.TextBox
    Friend WithEvents TxtStockInNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents gboItem As System.Windows.Forms.GroupBox
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
    Friend WithEvents txtPrimaryUnitsDefault As System.Windows.Forms.TextBox
    Friend WithEvents txtSecondaryQty As System.Windows.Forms.TextBox
    Friend WithEvents lblUnitsDefault As System.Windows.Forms.Label
    Friend WithEvents TxtQtyperBatchDefault As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents DtpExpired As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemarkItem As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TxtWarehouse As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
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
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents LblNamaPihakAcuan As System.Windows.Forms.Label
    Friend WithEvents LblKodePihakAcuan As System.Windows.Forms.Label
    Friend WithEvents BtnCariPihakAcuan As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryPihakAcuan As System.Windows.Forms.TextBox
    Friend WithEvents TxtPihakAcuan As System.Windows.Forms.TextBox
    Friend WithEvents TxtKodePihakAcuan As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnFindDO As System.Windows.Forms.Button
    Friend WithEvents txtPrimaryDO As System.Windows.Forms.TextBox
    Friend WithEvents txtDONo As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnFindSalesOrder As System.Windows.Forms.Button
    Friend WithEvents txtPrimarySO As System.Windows.Forms.TextBox
    Friend WithEvents txtSalesOrderNo As System.Windows.Forms.TextBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents txtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TxtRatioDefault As System.Windows.Forms.TextBox
    Friend WithEvents TxtRatioReference As System.Windows.Forms.TextBox
    Friend WithEvents TxtQtyperBatchMulti As System.Windows.Forms.TextBox
    Friend WithEvents lblUnitsMulti As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryUnitsMulti As System.Windows.Forms.TextBox
    Public WithEvents DgvItemDO2 As System.Windows.Forms.DataGridView
    Friend WithEvents LblQty As System.Windows.Forms.Label
    Public WithEvents dgvItemDO As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPONoDONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppDONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitsCodeMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecondaryQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecondaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiredDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWareHouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPOnoDoNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySecondaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPdO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitsMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_CustCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_CustDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_SONO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_BatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_ExpiryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_QtyMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_QtyPerMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_MasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_TotalQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_UnitsCodeMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_RatioDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_TotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_SecQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_SecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_PrimarySO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_PrimaryMasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_PrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_PrimarySecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_RatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_RatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_PrimaryUnitsMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_Location As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_Warehouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO_PrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_CustCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_CustDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_SONO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_BatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_ExpiryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_QtyMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_QtyPerMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_MasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_TotalQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_UnitsCodeMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_RatioDesc As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColDO2_TotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_SecQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_SecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_PrimarySO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_PrimaryMasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_PrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_PrimarySecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_RatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_RatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_PrimaryUnitsMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_Location As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_Warehouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDO2_PrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents AddItemByCustom As System.Windows.Forms.ToolStripMenuItem
End Class
