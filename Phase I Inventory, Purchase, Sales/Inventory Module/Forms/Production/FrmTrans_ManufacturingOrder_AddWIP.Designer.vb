<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ManufacturingOrder_AddWIP
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
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.DGVWIP = New System.Windows.Forms.DataGridView
        Me.ColWIP_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_WONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_Proses = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_QtyReq = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_QtyProduce = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_QtyWIP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_QtyAvailable = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_UnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_IdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_IdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_PrimaryProses = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_PrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWIP_DefPrecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.chkQtyAvailableOnly = New System.Windows.Forms.CheckBox
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.Label33 = New System.Windows.Forms.Label
        Me.TxtMP_QtyRMReq = New System.Windows.Forms.TextBox
        Me.Label31 = New System.Windows.Forms.Label
        Me.TxtMP_QtyPrOB = New System.Windows.Forms.TextBox
        Me.TxtMP_QtySafetyStock = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.TxtMP_QtySOB = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.TxtMP_QtyPendingDO = New System.Windows.Forms.TextBox
        Me.TxtMP_QtyMOB = New System.Windows.Forms.TextBox
        Me.Label36 = New System.Windows.Forms.Label
        Me.TxtMP_QtyAvailable = New System.Windows.Forms.TextBox
        Me.TxtMP_QtyOnHand = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.TxtMP_QtyPOB = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.TxtMP_QtyGIT = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.Label1 = New System.Windows.Forms.Label
        CType(Me.DGVWIP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(805, 428)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(724, 428)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
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
        Me.DGVWIP.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColWIP_No, Me.ColWIP_WONo, Me.ColWIP_Proses, Me.ColWIP_CodeNo, Me.ColWIP_Desc, Me.ColWIP_QtyReq, Me.ColWIP_QtyProduce, Me.ColWIP_QtyWIP, Me.ColWIP_QtyAvailable, Me.ColWIP_UnitCode, Me.ColWIP_Remarks, Me.ColWIP_IdentifierJO, Me.ColWIP_IdentifierWO, Me.ColWIP_PrimaryProses, Me.ColWIP_PrimaryCodeNo, Me.ColWIP_PrimaryUnit, Me.ColWIP_DefPrecUnit})
        Me.DGVWIP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVWIP.Location = New System.Drawing.Point(3, 18)
        Me.DGVWIP.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DGVWIP.Name = "DGVWIP"
        Me.DGVWIP.RowHeadersVisible = False
        Me.DGVWIP.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DGVWIP.RowTemplate.Height = 24
        Me.DGVWIP.Size = New System.Drawing.Size(862, 232)
        Me.DGVWIP.TabIndex = 21
        '
        'ColWIP_No
        '
        Me.ColWIP_No.HeaderText = "No"
        Me.ColWIP_No.Name = "ColWIP_No"
        Me.ColWIP_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_No.Width = 32
        '
        'ColWIP_WONo
        '
        Me.ColWIP_WONo.HeaderText = "WO Number"
        Me.ColWIP_WONo.Name = "ColWIP_WONo"
        Me.ColWIP_WONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_WONo.Width = 92
        '
        'ColWIP_Proses
        '
        Me.ColWIP_Proses.HeaderText = "Process"
        Me.ColWIP_Proses.Name = "ColWIP_Proses"
        Me.ColWIP_Proses.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_Proses.Width = 65
        '
        'ColWIP_CodeNo
        '
        Me.ColWIP_CodeNo.HeaderText = "Code No"
        Me.ColWIP_CodeNo.Name = "ColWIP_CodeNo"
        Me.ColWIP_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_CodeNo.Width = 69
        '
        'ColWIP_Desc
        '
        Me.ColWIP_Desc.HeaderText = "Description"
        Me.ColWIP_Desc.Name = "ColWIP_Desc"
        Me.ColWIP_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_Desc.Width = 85
        '
        'ColWIP_QtyReq
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColWIP_QtyReq.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColWIP_QtyReq.HeaderText = "Qty Required"
        Me.ColWIP_QtyReq.Name = "ColWIP_QtyReq"
        Me.ColWIP_QtyReq.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_QtyReq.Width = 98
        '
        'ColWIP_QtyProduce
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColWIP_QtyProduce.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColWIP_QtyProduce.HeaderText = "Qty Produce"
        Me.ColWIP_QtyProduce.Name = "ColWIP_QtyProduce"
        Me.ColWIP_QtyProduce.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_QtyProduce.Width = 93
        '
        'ColWIP_QtyWIP
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColWIP_QtyWIP.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColWIP_QtyWIP.HeaderText = "Qty WIP"
        Me.ColWIP_QtyWIP.Name = "ColWIP_QtyWIP"
        Me.ColWIP_QtyWIP.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWIP_QtyWIP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_QtyWIP.Width = 65
        '
        'ColWIP_QtyAvailable
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColWIP_QtyAvailable.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColWIP_QtyAvailable.HeaderText = "Qty Available"
        Me.ColWIP_QtyAvailable.Name = "ColWIP_QtyAvailable"
        Me.ColWIP_QtyAvailable.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_QtyAvailable.Width = 97
        '
        'ColWIP_UnitCode
        '
        Me.ColWIP_UnitCode.HeaderText = "Unit"
        Me.ColWIP_UnitCode.Name = "ColWIP_UnitCode"
        Me.ColWIP_UnitCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_UnitCode.Width = 39
        '
        'ColWIP_Remarks
        '
        Me.ColWIP_Remarks.HeaderText = "Remarks"
        Me.ColWIP_Remarks.Name = "ColWIP_Remarks"
        Me.ColWIP_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_Remarks.Visible = False
        Me.ColWIP_Remarks.Width = 70
        '
        'ColWIP_IdentifierJO
        '
        Me.ColWIP_IdentifierJO.HeaderText = "Identifier JO"
        Me.ColWIP_IdentifierJO.Name = "ColWIP_IdentifierJO"
        Me.ColWIP_IdentifierJO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_IdentifierJO.Visible = False
        Me.ColWIP_IdentifierJO.Width = 90
        '
        'ColWIP_IdentifierWO
        '
        Me.ColWIP_IdentifierWO.HeaderText = "Identifier WO"
        Me.ColWIP_IdentifierWO.Name = "ColWIP_IdentifierWO"
        Me.ColWIP_IdentifierWO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_IdentifierWO.Visible = False
        Me.ColWIP_IdentifierWO.Width = 96
        '
        'ColWIP_PrimaryProses
        '
        Me.ColWIP_PrimaryProses.HeaderText = "Primary Process"
        Me.ColWIP_PrimaryProses.Name = "ColWIP_PrimaryProses"
        Me.ColWIP_PrimaryProses.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_PrimaryProses.Visible = False
        Me.ColWIP_PrimaryProses.Width = 117
        '
        'ColWIP_PrimaryCodeNo
        '
        Me.ColWIP_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColWIP_PrimaryCodeNo.Name = "ColWIP_PrimaryCodeNo"
        Me.ColWIP_PrimaryCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_PrimaryCodeNo.Visible = False
        Me.ColWIP_PrimaryCodeNo.Width = 121
        '
        'ColWIP_PrimaryUnit
        '
        Me.ColWIP_PrimaryUnit.HeaderText = "Primary Unit"
        Me.ColWIP_PrimaryUnit.Name = "ColWIP_PrimaryUnit"
        Me.ColWIP_PrimaryUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_PrimaryUnit.Visible = False
        Me.ColWIP_PrimaryUnit.Width = 91
        '
        'ColWIP_DefPrecUnit
        '
        Me.ColWIP_DefPrecUnit.HeaderText = "DefPrecUnit"
        Me.ColWIP_DefPrecUnit.Name = "ColWIP_DefPrecUnit"
        Me.ColWIP_DefPrecUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWIP_DefPrecUnit.Visible = False
        Me.ColWIP_DefPrecUnit.Width = 90
        '
        'chkQtyAvailableOnly
        '
        Me.chkQtyAvailableOnly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkQtyAvailableOnly.AutoSize = True
        Me.chkQtyAvailableOnly.Checked = True
        Me.chkQtyAvailableOnly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkQtyAvailableOnly.Location = New System.Drawing.Point(12, 429)
        Me.chkQtyAvailableOnly.Name = "chkQtyAvailableOnly"
        Me.chkQtyAvailableOnly.Size = New System.Drawing.Size(246, 21)
        Me.chkQtyAvailableOnly.TabIndex = 22
        Me.chkQtyAvailableOnly.Text = "Show Item With Qty Available Only"
        Me.chkQtyAvailableOnly.UseVisualStyleBackColor = True
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(12, 269)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label33)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtyRMReq)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label31)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtyPrOB)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtySafetyStock)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label16)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtySOB)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label14)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtyPendingDO)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyMOB)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label36)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyAvailable)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyOnHand)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label18)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyPOB)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label10)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label15)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyGIT)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer2.Size = New System.Drawing.Size(868, 154)
        Me.SplitContainer2.SplitterDistance = 432
        Me.SplitContainer2.TabIndex = 72
        Me.SplitContainer2.TabStop = False
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(4, 100)
        Me.Label33.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(123, 17)
        Me.Label33.TabIndex = 71
        Me.Label33.Text = "Raw Material Req."
        '
        'TxtMP_QtyRMReq
        '
        Me.TxtMP_QtyRMReq.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyRMReq.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyRMReq.Location = New System.Drawing.Point(135, 96)
        Me.TxtMP_QtyRMReq.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtMP_QtyRMReq.Name = "TxtMP_QtyRMReq"
        Me.TxtMP_QtyRMReq.ReadOnly = True
        Me.TxtMP_QtyRMReq.Size = New System.Drawing.Size(292, 22)
        Me.TxtMP_QtyRMReq.TabIndex = 70
        Me.TxtMP_QtyRMReq.Text = "0"
        Me.TxtMP_QtyRMReq.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(4, 9)
        Me.Label31.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(106, 17)
        Me.Label31.TabIndex = 69
        Me.Label31.Text = "Purch. Request"
        '
        'TxtMP_QtyPrOB
        '
        Me.TxtMP_QtyPrOB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyPrOB.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyPrOB.Location = New System.Drawing.Point(135, 7)
        Me.TxtMP_QtyPrOB.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtMP_QtyPrOB.Name = "TxtMP_QtyPrOB"
        Me.TxtMP_QtyPrOB.ReadOnly = True
        Me.TxtMP_QtyPrOB.Size = New System.Drawing.Size(292, 22)
        Me.TxtMP_QtyPrOB.TabIndex = 0
        Me.TxtMP_QtyPrOB.Text = "0"
        Me.TxtMP_QtyPrOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtMP_QtySafetyStock
        '
        Me.TxtMP_QtySafetyStock.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtySafetyStock.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtySafetyStock.Location = New System.Drawing.Point(135, 126)
        Me.TxtMP_QtySafetyStock.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtMP_QtySafetyStock.Name = "TxtMP_QtySafetyStock"
        Me.TxtMP_QtySafetyStock.ReadOnly = True
        Me.TxtMP_QtySafetyStock.Size = New System.Drawing.Size(292, 22)
        Me.TxtMP_QtySafetyStock.TabIndex = 4
        Me.TxtMP_QtySafetyStock.Text = "0"
        Me.TxtMP_QtySafetyStock.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(4, 39)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(98, 17)
        Me.Label11.TabIndex = 57
        Me.Label11.Text = "Sales Balance"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(4, 129)
        Me.Label16.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(87, 17)
        Me.Label16.TabIndex = 67
        Me.Label16.Text = "Safety Stock"
        '
        'TxtMP_QtySOB
        '
        Me.TxtMP_QtySOB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtySOB.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtySOB.Location = New System.Drawing.Point(135, 37)
        Me.TxtMP_QtySOB.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtMP_QtySOB.Name = "TxtMP_QtySOB"
        Me.TxtMP_QtySOB.ReadOnly = True
        Me.TxtMP_QtySOB.Size = New System.Drawing.Size(292, 22)
        Me.TxtMP_QtySOB.TabIndex = 1
        Me.TxtMP_QtySOB.Text = "0"
        Me.TxtMP_QtySOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(4, 70)
        Me.Label14.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(85, 17)
        Me.Label14.TabIndex = 59
        Me.Label14.Text = "Pending DO"
        '
        'TxtMP_QtyPendingDO
        '
        Me.TxtMP_QtyPendingDO.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyPendingDO.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyPendingDO.Location = New System.Drawing.Point(135, 66)
        Me.TxtMP_QtyPendingDO.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtMP_QtyPendingDO.Name = "TxtMP_QtyPendingDO"
        Me.TxtMP_QtyPendingDO.ReadOnly = True
        Me.TxtMP_QtyPendingDO.Size = New System.Drawing.Size(292, 22)
        Me.TxtMP_QtyPendingDO.TabIndex = 2
        Me.TxtMP_QtyPendingDO.Text = "0"
        Me.TxtMP_QtyPendingDO.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtMP_QtyMOB
        '
        Me.TxtMP_QtyMOB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyMOB.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyMOB.Location = New System.Drawing.Point(143, 95)
        Me.TxtMP_QtyMOB.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtMP_QtyMOB.Name = "TxtMP_QtyMOB"
        Me.TxtMP_QtyMOB.ReadOnly = True
        Me.TxtMP_QtyMOB.Size = New System.Drawing.Size(274, 22)
        Me.TxtMP_QtyMOB.TabIndex = 74
        Me.TxtMP_QtyMOB.Text = "0"
        Me.TxtMP_QtyMOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.Location = New System.Drawing.Point(4, 99)
        Me.Label36.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(131, 17)
        Me.Label36.TabIndex = 75
        Me.Label36.Text = "Production Balance"
        '
        'TxtMP_QtyAvailable
        '
        Me.TxtMP_QtyAvailable.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyAvailable.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyAvailable.Location = New System.Drawing.Point(143, 126)
        Me.TxtMP_QtyAvailable.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtMP_QtyAvailable.Name = "TxtMP_QtyAvailable"
        Me.TxtMP_QtyAvailable.ReadOnly = True
        Me.TxtMP_QtyAvailable.Size = New System.Drawing.Size(274, 22)
        Me.TxtMP_QtyAvailable.TabIndex = 6
        Me.TxtMP_QtyAvailable.Text = "0"
        Me.TxtMP_QtyAvailable.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtMP_QtyOnHand
        '
        Me.TxtMP_QtyOnHand.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyOnHand.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyOnHand.Location = New System.Drawing.Point(143, 7)
        Me.TxtMP_QtyOnHand.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtMP_QtyOnHand.Name = "TxtMP_QtyOnHand"
        Me.TxtMP_QtyOnHand.ReadOnly = True
        Me.TxtMP_QtyOnHand.Size = New System.Drawing.Size(274, 22)
        Me.TxtMP_QtyOnHand.TabIndex = 0
        Me.TxtMP_QtyOnHand.Text = "0"
        Me.TxtMP_QtyOnHand.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(4, 129)
        Me.Label18.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(65, 17)
        Me.Label18.TabIndex = 69
        Me.Label18.Text = "Available"
        '
        'TxtMP_QtyPOB
        '
        Me.TxtMP_QtyPOB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyPOB.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyPOB.Location = New System.Drawing.Point(143, 66)
        Me.TxtMP_QtyPOB.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtMP_QtyPOB.Name = "TxtMP_QtyPOB"
        Me.TxtMP_QtyPOB.ReadOnly = True
        Me.TxtMP_QtyPOB.Size = New System.Drawing.Size(274, 22)
        Me.TxtMP_QtyPOB.TabIndex = 2
        Me.TxtMP_QtyPOB.Text = "0"
        Me.TxtMP_QtyPOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(4, 70)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(123, 17)
        Me.Label10.TabIndex = 53
        Me.Label10.Text = "Purchase Balance"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(4, 10)
        Me.Label15.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(65, 17)
        Me.Label15.TabIndex = 63
        Me.Label15.Text = "On Hand"
        '
        'TxtMP_QtyGIT
        '
        Me.TxtMP_QtyGIT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyGIT.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyGIT.Location = New System.Drawing.Point(143, 37)
        Me.TxtMP_QtyGIT.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtMP_QtyGIT.Name = "TxtMP_QtyGIT"
        Me.TxtMP_QtyGIT.ReadOnly = True
        Me.TxtMP_QtyGIT.Size = New System.Drawing.Size(274, 22)
        Me.TxtMP_QtyGIT.TabIndex = 1
        Me.TxtMP_QtyGIT.Text = "0"
        Me.TxtMP_QtyGIT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(4, 39)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(113, 17)
        Me.Label9.TabIndex = 51
        Me.Label9.Text = "Goods In Transit"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.Panel1)
        Me.GroupBox1.Controls.Add(Me.DGVWIP)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(868, 253)
        Me.GroupBox1.TabIndex = 73
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Raw Material Requirement (WIP Only)"
        '
        'Timer1
        '
        '
        'BackgroundWorker1
        '
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.ProgressBar1)
        Me.Panel1.Location = New System.Drawing.Point(224, 103)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(420, 54)
        Me.Panel1.TabIndex = 22
        Me.Panel1.Visible = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(3, 3)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(414, 23)
        Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.ProgressBar1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(133, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(155, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Calculating, please wait"
        '
        'FrmTrans_ManufacturingOrder_AddWIP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(892, 463)
        Me.Controls.Add(Me.SplitContainer2)
        Me.Controls.Add(Me.chkQtyAvailableOnly)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximizeBox = False
        Me.MinimumSize = New System.Drawing.Size(900, 500)
        Me.Name = "FrmTrans_ManufacturingOrder_AddWIP"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add WIP"
        CType(Me.DGVWIP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents DGVWIP As System.Windows.Forms.DataGridView
    Friend WithEvents chkQtyAvailableOnly As System.Windows.Forms.CheckBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyRMReq As System.Windows.Forms.TextBox
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyPrOB As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_QtySafetyStock As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtySOB As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyPendingDO As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_QtyMOB As System.Windows.Forms.TextBox
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyAvailable As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_QtyOnHand As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyPOB As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyGIT As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ColWIP_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_WONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_Proses As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_QtyReq As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_QtyProduce As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_QtyWIP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_QtyAvailable As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_UnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_IdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_IdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_PrimaryProses As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_PrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWIP_DefPrecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
End Class
