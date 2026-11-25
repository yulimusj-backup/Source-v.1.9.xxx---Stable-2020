<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockTransfer_Receive
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.dgvFilter = New System.Windows.Forms.DataGridView
        Me.ColFilter_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColFilter_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilter_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilter_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.BtnReceivedCancel = New System.Windows.Forms.Button
        Me.BtnReceived = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtDocNoFilter = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.chkFilter_Received = New System.Windows.Forms.CheckBox
        Me.chkFilter_OnTransit = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.dgvHeader = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocumentNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatus = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReceivedNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReceivedDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReceivedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReceivedOn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColModify = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColReceived = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColEdit = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColCancelReceived = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColPrimaryDoc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryReceiveDoc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.dgvFilter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvHeader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.TxtReadOnly)
        Me.GroupBox1.Controls.Add(Me.BtnReceivedCancel)
        Me.GroupBox1.Controls.Add(Me.BtnReceived)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TxtDocNoFilter)
        Me.GroupBox1.Controls.Add(Me.btnClose)
        Me.GroupBox1.Controls.Add(Me.btnFilter)
        Me.GroupBox1.Controls.Add(Me.chkFilter_Received)
        Me.GroupBox1.Controls.Add(Me.chkFilter_OnTransit)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 10)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(882, 141)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.dgvFilter)
        Me.GroupBox3.Location = New System.Drawing.Point(210, 13)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(459, 123)
        Me.GroupBox3.TabIndex = 14
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Destination"
        '
        'dgvFilter
        '
        Me.dgvFilter.AllowUserToResizeColumns = False
        Me.dgvFilter.AllowUserToResizeRows = False
        Me.dgvFilter.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvFilter.ColumnHeadersHeight = 25
        Me.dgvFilter.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvFilter.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColFilter_Level, Me.ColFilter_Code, Me.ColFilter_Desc, Me.ColFilter_Primary})
        Me.dgvFilter.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvFilter.Location = New System.Drawing.Point(3, 16)
        Me.dgvFilter.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvFilter.MultiSelect = False
        Me.dgvFilter.Name = "dgvFilter"
        Me.dgvFilter.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvFilter.RowTemplate.Height = 24
        Me.dgvFilter.Size = New System.Drawing.Size(453, 104)
        Me.dgvFilter.TabIndex = 6
        '
        'ColFilter_Level
        '
        Me.ColFilter_Level.HeaderText = "Level"
        Me.ColFilter_Level.Items.AddRange(New Object() {"Warehouse", "Location"})
        Me.ColFilter_Level.Name = "ColFilter_Level"
        Me.ColFilter_Level.Width = 39
        '
        'ColFilter_Code
        '
        Me.ColFilter_Code.HeaderText = "Code"
        Me.ColFilter_Code.Name = "ColFilter_Code"
        Me.ColFilter_Code.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColFilter_Code.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColFilter_Code.Width = 38
        '
        'ColFilter_Desc
        '
        Me.ColFilter_Desc.HeaderText = "Description"
        Me.ColFilter_Desc.Name = "ColFilter_Desc"
        Me.ColFilter_Desc.ReadOnly = True
        Me.ColFilter_Desc.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColFilter_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColFilter_Desc.Width = 66
        '
        'ColFilter_Primary
        '
        Me.ColFilter_Primary.HeaderText = "Primary"
        Me.ColFilter_Primary.Name = "ColFilter_Primary"
        Me.ColFilter_Primary.ReadOnly = True
        Me.ColFilter_Primary.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColFilter_Primary.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColFilter_Primary.Visible = False
        Me.ColFilter_Primary.Width = 47
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(81, 13)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(29, 20)
        Me.TxtReadOnly.TabIndex = 11
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'BtnReceivedCancel
        '
        Me.BtnReceivedCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnReceivedCancel.Location = New System.Drawing.Point(674, 64)
        Me.BtnReceivedCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnReceivedCancel.Name = "BtnReceivedCancel"
        Me.BtnReceivedCancel.Size = New System.Drawing.Size(106, 20)
        Me.BtnReceivedCancel.TabIndex = 9
        Me.BtnReceivedCancel.Text = "Cancel Received"
        Me.BtnReceivedCancel.UseVisualStyleBackColor = True
        '
        'BtnReceived
        '
        Me.BtnReceived.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnReceived.Location = New System.Drawing.Point(674, 39)
        Me.BtnReceived.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnReceived.Name = "BtnReceived"
        Me.BtnReceived.Size = New System.Drawing.Size(106, 20)
        Me.BtnReceived.TabIndex = 8
        Me.BtnReceived.Text = "Receive"
        Me.BtnReceived.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 16)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Document No"
        '
        'TxtDocNoFilter
        '
        Me.TxtDocNoFilter.Location = New System.Drawing.Point(80, 14)
        Me.TxtDocNoFilter.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtDocNoFilter.Name = "TxtDocNoFilter"
        Me.TxtDocNoFilter.Size = New System.Drawing.Size(125, 20)
        Me.TxtDocNoFilter.TabIndex = 1
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(785, 14)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(56, 20)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(674, 14)
        Me.btnFilter.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(106, 20)
        Me.btnFilter.TabIndex = 7
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'chkFilter_Received
        '
        Me.chkFilter_Received.AutoSize = True
        Me.chkFilter_Received.Checked = True
        Me.chkFilter_Received.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFilter_Received.Location = New System.Drawing.Point(88, 110)
        Me.chkFilter_Received.Margin = New System.Windows.Forms.Padding(2)
        Me.chkFilter_Received.Name = "chkFilter_Received"
        Me.chkFilter_Received.Size = New System.Drawing.Size(72, 17)
        Me.chkFilter_Received.TabIndex = 5
        Me.chkFilter_Received.Text = "Received"
        Me.chkFilter_Received.UseVisualStyleBackColor = True
        '
        'chkFilter_OnTransit
        '
        Me.chkFilter_OnTransit.AutoSize = True
        Me.chkFilter_OnTransit.Checked = True
        Me.chkFilter_OnTransit.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFilter_OnTransit.Location = New System.Drawing.Point(4, 110)
        Me.chkFilter_OnTransit.Margin = New System.Windows.Forms.Padding(2)
        Me.chkFilter_OnTransit.Name = "chkFilter_OnTransit"
        Me.chkFilter_OnTransit.Size = New System.Drawing.Size(75, 17)
        Me.chkFilter_OnTransit.TabIndex = 4
        Me.chkFilter_OnTransit.Text = "On Transit"
        Me.chkFilter_OnTransit.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.dtpTo)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.dtpFrom)
        Me.GroupBox2.Location = New System.Drawing.Point(4, 39)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Size = New System.Drawing.Size(201, 66)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Transfer Date"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 41)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "To"
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(76, 40)
        Me.dtpTo.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(119, 20)
        Me.dtpTo.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 19)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(76, 17)
        Me.dtpFrom.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(119, 20)
        Me.dtpFrom.TabIndex = 1
        '
        'llbUncheckAll
        '
        Me.llbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUncheckAll.AutoSize = True
        Me.llbUncheckAll.Location = New System.Drawing.Point(823, 139)
        Me.llbUncheckAll.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.Size = New System.Drawing.Size(65, 13)
        Me.llbUncheckAll.TabIndex = 12
        Me.llbUncheckAll.TabStop = True
        Me.llbUncheckAll.Text = "Uncheck All"
        '
        'dgvHeader
        '
        Me.dgvHeader.AllowUserToAddRows = False
        Me.dgvHeader.AllowUserToDeleteRows = False
        Me.dgvHeader.AllowUserToResizeColumns = False
        Me.dgvHeader.AllowUserToResizeRows = False
        Me.dgvHeader.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvHeader.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvHeader.ColumnHeadersHeight = 25
        Me.dgvHeader.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvHeader.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColDocumentNo, Me.ColDate, Me.ColRemarks, Me.ColCreatedBy, Me.ColStatus, Me.ColReceivedNo, Me.ColReceivedDate, Me.ColReceivedBy, Me.ColReceivedOn, Me.ColModify, Me.ColReceived, Me.ColEdit, Me.ColCancelReceived, Me.ColPrimaryDoc, Me.ColPrimaryReceiveDoc})
        Me.dgvHeader.Location = New System.Drawing.Point(9, 155)
        Me.dgvHeader.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvHeader.MultiSelect = False
        Me.dgvHeader.Name = "dgvHeader"
        Me.dgvHeader.ReadOnly = True
        Me.dgvHeader.RowHeadersVisible = False
        Me.dgvHeader.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvHeader.RowTemplate.Height = 24
        Me.dgvHeader.Size = New System.Drawing.Size(882, 378)
        Me.dgvHeader.TabIndex = 13
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 46
        '
        'ColDocumentNo
        '
        Me.ColDocumentNo.Frozen = True
        Me.ColDocumentNo.HeaderText = "Document No"
        Me.ColDocumentNo.Name = "ColDocumentNo"
        Me.ColDocumentNo.ReadOnly = True
        Me.ColDocumentNo.Width = 98
        '
        'ColDate
        '
        DataGridViewCellStyle1.Format = "dd MMM yyyy"
        Me.ColDate.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.Width = 55
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.ReadOnly = True
        Me.ColRemarks.Width = 74
        '
        'ColCreatedBy
        '
        Me.ColCreatedBy.HeaderText = "Created By"
        Me.ColCreatedBy.Name = "ColCreatedBy"
        Me.ColCreatedBy.ReadOnly = True
        Me.ColCreatedBy.Width = 84
        '
        'ColStatus
        '
        Me.ColStatus.HeaderText = "Status"
        Me.ColStatus.Name = "ColStatus"
        Me.ColStatus.ReadOnly = True
        Me.ColStatus.Width = 62
        '
        'ColReceivedNo
        '
        Me.ColReceivedNo.HeaderText = "Received No"
        Me.ColReceivedNo.Name = "ColReceivedNo"
        Me.ColReceivedNo.ReadOnly = True
        Me.ColReceivedNo.Width = 95
        '
        'ColReceivedDate
        '
        DataGridViewCellStyle2.Format = "dd MMM yyyy"
        DataGridViewCellStyle2.NullValue = "1/1/1950"
        Me.ColReceivedDate.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColReceivedDate.HeaderText = "Received Date"
        Me.ColReceivedDate.Name = "ColReceivedDate"
        Me.ColReceivedDate.ReadOnly = True
        Me.ColReceivedDate.Width = 104
        '
        'ColReceivedBy
        '
        Me.ColReceivedBy.HeaderText = "Received By"
        Me.ColReceivedBy.Name = "ColReceivedBy"
        Me.ColReceivedBy.ReadOnly = True
        Me.ColReceivedBy.Width = 93
        '
        'ColReceivedOn
        '
        DataGridViewCellStyle3.Format = "dd MMM yyyy HH:mm"
        DataGridViewCellStyle3.NullValue = "1/1/1950"
        Me.ColReceivedOn.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColReceivedOn.HeaderText = "Received Timestamp"
        Me.ColReceivedOn.Name = "ColReceivedOn"
        Me.ColReceivedOn.ReadOnly = True
        Me.ColReceivedOn.Width = 132
        '
        'ColModify
        '
        Me.ColModify.HeaderText = "Modify"
        Me.ColModify.Name = "ColModify"
        Me.ColModify.ReadOnly = True
        Me.ColModify.Width = 44
        '
        'ColReceived
        '
        Me.ColReceived.HeaderText = "Receive"
        Me.ColReceived.Name = "ColReceived"
        Me.ColReceived.ReadOnly = True
        Me.ColReceived.Width = 53
        '
        'ColEdit
        '
        Me.ColEdit.HeaderText = "Edit"
        Me.ColEdit.Name = "ColEdit"
        Me.ColEdit.ReadOnly = True
        Me.ColEdit.Visible = False
        Me.ColEdit.Width = 31
        '
        'ColCancelReceived
        '
        Me.ColCancelReceived.HeaderText = "Cancel Receive"
        Me.ColCancelReceived.Name = "ColCancelReceived"
        Me.ColCancelReceived.ReadOnly = True
        Me.ColCancelReceived.Width = 89
        '
        'ColPrimaryDoc
        '
        Me.ColPrimaryDoc.HeaderText = "Primary Doc"
        Me.ColPrimaryDoc.Name = "ColPrimaryDoc"
        Me.ColPrimaryDoc.ReadOnly = True
        Me.ColPrimaryDoc.Visible = False
        Me.ColPrimaryDoc.Width = 89
        '
        'ColPrimaryReceiveDoc
        '
        Me.ColPrimaryReceiveDoc.HeaderText = "Primary Receive Doc"
        Me.ColPrimaryReceiveDoc.Name = "ColPrimaryReceiveDoc"
        Me.ColPrimaryReceiveDoc.ReadOnly = True
        Me.ColPrimaryReceiveDoc.Visible = False
        Me.ColPrimaryReceiveDoc.Width = 132
        '
        'llbCheckAll
        '
        Me.llbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.Location = New System.Drawing.Point(762, 139)
        Me.llbCheckAll.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(52, 13)
        Me.llbCheckAll.TabIndex = 11
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Check All"
        '
        'BackgroundWorker1
        '
        '
        'FrmTransaksiStockTransfer_Receive
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(900, 543)
        Me.Controls.Add(Me.llbCheckAll)
        Me.Controls.Add(Me.llbUncheckAll)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.dgvHeader)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(754, 552)
        Me.Name = "FrmTransaksiStockTransfer_Receive"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Location Transfer - Set Transit Status"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.dgvFilter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.dgvHeader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkFilter_Received As System.Windows.Forms.CheckBox
    Friend WithEvents chkFilter_OnTransit As System.Windows.Forms.CheckBox
    Friend WithEvents dgvFilter As System.Windows.Forms.DataGridView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents dgvHeader As System.Windows.Forms.DataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtDocNoFilter As System.Windows.Forms.TextBox
    Friend WithEvents BtnReceivedCancel As System.Windows.Forms.Button
    Friend WithEvents BtnReceived As System.Windows.Forms.Button
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents ColFilter_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColFilter_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilter_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilter_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocumentNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReceivedNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReceivedDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReceivedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReceivedOn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColModify As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColReceived As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColEdit As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColCancelReceived As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColPrimaryDoc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryReceiveDoc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
End Class
