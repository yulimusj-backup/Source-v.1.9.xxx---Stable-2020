<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_ExplainPrOB
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
        Me.dgvData = New System.Windows.Forms.DataGridView
        Me.Button1 = New System.Windows.Forms.Button
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPRNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRefParty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOrderQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSpareQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPR = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeliveryCont = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDeliveryDate = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvData
        '
        Me.dgvData.AllowUserToAddRows = False
        Me.dgvData.AllowUserToDeleteRows = False
        Me.dgvData.AllowUserToResizeColumns = False
        Me.dgvData.AllowUserToResizeRows = False
        Me.dgvData.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvData.ColumnHeadersHeight = 25
        Me.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvData.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColPRNumber, Me.ColDate, Me.ColRefParty, Me.ColRemarks, Me.ColOrderQty, Me.ColSpareQty, Me.ColTotalQty, Me.ColUnit, Me.ColPrimaryPR, Me.ColDeliveryCont, Me.ColDeliveryDate, Me.ColIdentifier})
        Me.dgvData.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvData.Location = New System.Drawing.Point(0, 0)
        Me.dgvData.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvData.Name = "dgvData"
        Me.dgvData.RowHeadersVisible = False
        Me.dgvData.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvData.RowTemplate.Height = 24
        Me.dgvData.Size = New System.Drawing.Size(832, 491)
        Me.dgvData.TabIndex = 6
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(560, 164)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 7
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'ColPRNumber
        '
        Me.ColPRNumber.HeaderText = "Purc. Req. Number"
        Me.ColPRNumber.Name = "ColPRNumber"
        Me.ColPRNumber.ReadOnly = True
        Me.ColPRNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPRNumber.Width = 135
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDate.Width = 44
        '
        'ColRefParty
        '
        Me.ColRefParty.HeaderText = "Ref Party"
        Me.ColRefParty.Name = "ColRefParty"
        Me.ColRefParty.ReadOnly = True
        Me.ColRefParty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRefParty.Width = 73
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.ReadOnly = True
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 70
        '
        'ColOrderQty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColOrderQty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColOrderQty.HeaderText = "Balance Qty"
        Me.ColOrderQty.Name = "ColOrderQty"
        Me.ColOrderQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColOrderQty.Width = 91
        '
        'ColSpareQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSpareQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColSpareQty.HeaderText = "Spare Qty"
        Me.ColSpareQty.Name = "ColSpareQty"
        Me.ColSpareQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSpareQty.Width = 78
        '
        'ColTotalQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColTotalQty.HeaderText = "Total Qty"
        Me.ColTotalQty.Name = "ColTotalQty"
        Me.ColTotalQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTotalQty.Width = 72
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Width = 39
        '
        'ColPrimaryPR
        '
        Me.ColPrimaryPR.HeaderText = "Primary PR"
        Me.ColPrimaryPR.Name = "ColPrimaryPR"
        Me.ColPrimaryPR.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryPR.Visible = False
        Me.ColPrimaryPR.Width = 85
        '
        'ColDeliveryCont
        '
        Me.ColDeliveryCont.HeaderText = "Delivery Date Control"
        Me.ColDeliveryCont.Name = "ColDeliveryCont"
        Me.ColDeliveryCont.Width = 148
        '
        'ColDeliveryDate
        '
        Me.ColDeliveryDate.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColDeliveryDate.HeaderText = "Delivery Date"
        Me.ColDeliveryDate.Name = "ColDeliveryDate"
        Me.ColDeliveryDate.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColDeliveryDate.Width = 99
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColIdentifier.Visible = False
        Me.ColIdentifier.Width = 68
        '
        'FrmTrans_MaterialPlanningDashboard_ExplainPrOB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(832, 491)
        Me.Controls.Add(Me.dgvData)
        Me.Controls.Add(Me.Button1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(653, 528)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_ExplainPrOB"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Request Balance Detail"
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvData As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPRNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRefParty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOrderQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSpareQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPR As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeliveryCont As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDeliveryDate As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
