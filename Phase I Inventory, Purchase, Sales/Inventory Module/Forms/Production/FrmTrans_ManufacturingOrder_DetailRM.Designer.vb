<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ManufacturingOrder_DetailRM
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
        Me.DGVRM = New System.Windows.Forms.DataGridView
        Me.Button1 = New System.Windows.Forms.Button
        Me.ColRM_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_WONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_Process = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_QtyTarget = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_QtyRequired = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRM_QtyAllocated = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_QtyAdditional = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRM_QtyUsed = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRM_QtyBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_GPTarget = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_GPActual = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_StandByLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_PickUpLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_PrimaryStandByLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_PrimaryPickUpLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_IdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_IdentifierItem = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_PrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_PrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRM_IsWIP = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.DGVRM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DGVRM
        '
        Me.DGVRM.AllowUserToAddRows = False
        Me.DGVRM.AllowUserToDeleteRows = False
        Me.DGVRM.AllowUserToResizeColumns = False
        Me.DGVRM.AllowUserToResizeRows = False
        Me.DGVRM.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVRM.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DGVRM.ColumnHeadersHeight = 25
        Me.DGVRM.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DGVRM.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColRM_No, Me.ColRM_WONumber, Me.ColRM_Process, Me.ColRM_CodeNo, Me.ColRM_Desc, Me.ColRM_QtyTarget, Me.ColRM_QtyRequired, Me.ColRM_QtyAllocated, Me.ColRM_QtyAdditional, Me.ColRM_QtyUsed, Me.ColRM_QtyBalance, Me.ColRM_Unit, Me.ColRM_GPTarget, Me.ColRM_GPActual, Me.ColRM_StandByLocation, Me.ColRM_PickUpLocation, Me.ColRM_Remarks, Me.ColRM_PrimaryStandByLocation, Me.ColRM_PrimaryPickUpLocation, Me.ColRM_IdentifierWO, Me.ColRM_IdentifierItem, Me.ColRM_PrimaryProcess, Me.ColRM_PrimaryCodeNo, Me.ColRM_PrimaryUnit, Me.ColRM_IsWIP})
        Me.DGVRM.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVRM.Location = New System.Drawing.Point(0, 0)
        Me.DGVRM.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DGVRM.Name = "DGVRM"
        Me.DGVRM.ReadOnly = True
        Me.DGVRM.RowHeadersVisible = False
        Me.DGVRM.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DGVRM.RowTemplate.Height = 24
        Me.DGVRM.Size = New System.Drawing.Size(892, 413)
        Me.DGVRM.TabIndex = 20
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(571, 153)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 21
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ColRM_No
        '
        Me.ColRM_No.HeaderText = "No"
        Me.ColRM_No.Name = "ColRM_No"
        Me.ColRM_No.ReadOnly = True
        Me.ColRM_No.Width = 51
        '
        'ColRM_WONumber
        '
        Me.ColRM_WONumber.HeaderText = "WO Number"
        Me.ColRM_WONumber.Name = "ColRM_WONumber"
        Me.ColRM_WONumber.ReadOnly = True
        Me.ColRM_WONumber.Width = 111
        '
        'ColRM_Process
        '
        Me.ColRM_Process.HeaderText = "Process"
        Me.ColRM_Process.Name = "ColRM_Process"
        Me.ColRM_Process.ReadOnly = True
        Me.ColRM_Process.Width = 84
        '
        'ColRM_CodeNo
        '
        Me.ColRM_CodeNo.HeaderText = "Code No"
        Me.ColRM_CodeNo.Name = "ColRM_CodeNo"
        Me.ColRM_CodeNo.ReadOnly = True
        Me.ColRM_CodeNo.Width = 88
        '
        'ColRM_Desc
        '
        Me.ColRM_Desc.HeaderText = "Description"
        Me.ColRM_Desc.Name = "ColRM_Desc"
        Me.ColRM_Desc.ReadOnly = True
        Me.ColRM_Desc.Width = 104
        '
        'ColRM_QtyTarget
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_QtyTarget.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColRM_QtyTarget.HeaderText = "Qty Target"
        Me.ColRM_QtyTarget.Name = "ColRM_QtyTarget"
        Me.ColRM_QtyTarget.ReadOnly = True
        Me.ColRM_QtyTarget.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRM_QtyTarget.Visible = False
        Me.ColRM_QtyTarget.Width = 101
        '
        'ColRM_QtyRequired
        '
        Me.ColRM_QtyRequired.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_QtyRequired.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColRM_QtyRequired.HeaderText = "Qty Required"
        Me.ColRM_QtyRequired.Name = "ColRM_QtyRequired"
        Me.ColRM_QtyRequired.ReadOnly = True
        Me.ColRM_QtyRequired.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRM_QtyRequired.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRM_QtyRequired.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRM_QtyRequired.Width = 117
        '
        'ColRM_QtyAllocated
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_QtyAllocated.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColRM_QtyAllocated.HeaderText = "Qty Allocated"
        Me.ColRM_QtyAllocated.Name = "ColRM_QtyAllocated"
        Me.ColRM_QtyAllocated.ReadOnly = True
        Me.ColRM_QtyAllocated.Visible = False
        Me.ColRM_QtyAllocated.Width = 117
        '
        'ColRM_QtyAdditional
        '
        Me.ColRM_QtyAdditional.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_QtyAdditional.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColRM_QtyAdditional.HeaderText = "Qty Additional"
        Me.ColRM_QtyAdditional.Name = "ColRM_QtyAdditional"
        Me.ColRM_QtyAdditional.ReadOnly = True
        Me.ColRM_QtyAdditional.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRM_QtyAdditional.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRM_QtyAdditional.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRM_QtyAdditional.Width = 121
        '
        'ColRM_QtyUsed
        '
        Me.ColRM_QtyUsed.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_QtyUsed.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColRM_QtyUsed.HeaderText = "Qty Used"
        Me.ColRM_QtyUsed.Name = "ColRM_QtyUsed"
        Me.ColRM_QtyUsed.ReadOnly = True
        Me.ColRM_QtyUsed.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRM_QtyUsed.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRM_QtyUsed.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRM_QtyUsed.Width = 92
        '
        'ColRM_QtyBalance
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_QtyBalance.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColRM_QtyBalance.HeaderText = "Qty Balance"
        Me.ColRM_QtyBalance.Name = "ColRM_QtyBalance"
        Me.ColRM_QtyBalance.ReadOnly = True
        Me.ColRM_QtyBalance.Width = 110
        '
        'ColRM_Unit
        '
        Me.ColRM_Unit.HeaderText = "Unit"
        Me.ColRM_Unit.Name = "ColRM_Unit"
        Me.ColRM_Unit.ReadOnly = True
        Me.ColRM_Unit.Width = 58
        '
        'ColRM_GPTarget
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_GPTarget.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColRM_GPTarget.HeaderText = "Good Percentage Target (%)"
        Me.ColRM_GPTarget.Name = "ColRM_GPTarget"
        Me.ColRM_GPTarget.ReadOnly = True
        Me.ColRM_GPTarget.Visible = False
        Me.ColRM_GPTarget.Width = 217
        '
        'ColRM_GPActual
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRM_GPActual.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColRM_GPActual.HeaderText = "Good Percentage Actual (%)"
        Me.ColRM_GPActual.Name = "ColRM_GPActual"
        Me.ColRM_GPActual.ReadOnly = True
        Me.ColRM_GPActual.Visible = False
        Me.ColRM_GPActual.Width = 214
        '
        'ColRM_StandByLocation
        '
        Me.ColRM_StandByLocation.HeaderText = "Stand By Location"
        Me.ColRM_StandByLocation.Name = "ColRM_StandByLocation"
        Me.ColRM_StandByLocation.ReadOnly = True
        Me.ColRM_StandByLocation.Visible = False
        Me.ColRM_StandByLocation.Width = 148
        '
        'ColRM_PickUpLocation
        '
        Me.ColRM_PickUpLocation.HeaderText = "Pick Up Location"
        Me.ColRM_PickUpLocation.Name = "ColRM_PickUpLocation"
        Me.ColRM_PickUpLocation.ReadOnly = True
        Me.ColRM_PickUpLocation.Visible = False
        Me.ColRM_PickUpLocation.Width = 139
        '
        'ColRM_Remarks
        '
        Me.ColRM_Remarks.HeaderText = "Remarks"
        Me.ColRM_Remarks.Name = "ColRM_Remarks"
        Me.ColRM_Remarks.ReadOnly = True
        Me.ColRM_Remarks.Width = 89
        '
        'ColRM_PrimaryStandByLocation
        '
        Me.ColRM_PrimaryStandByLocation.HeaderText = "Primary Stand By Location"
        Me.ColRM_PrimaryStandByLocation.Name = "ColRM_PrimaryStandByLocation"
        Me.ColRM_PrimaryStandByLocation.ReadOnly = True
        Me.ColRM_PrimaryStandByLocation.Visible = False
        Me.ColRM_PrimaryStandByLocation.Width = 200
        '
        'ColRM_PrimaryPickUpLocation
        '
        Me.ColRM_PrimaryPickUpLocation.HeaderText = "Primary Pick Up Location"
        Me.ColRM_PrimaryPickUpLocation.Name = "ColRM_PrimaryPickUpLocation"
        Me.ColRM_PrimaryPickUpLocation.ReadOnly = True
        Me.ColRM_PrimaryPickUpLocation.Visible = False
        Me.ColRM_PrimaryPickUpLocation.Width = 191
        '
        'ColRM_IdentifierWO
        '
        Me.ColRM_IdentifierWO.HeaderText = "Identifier JO"
        Me.ColRM_IdentifierWO.Name = "ColRM_IdentifierWO"
        Me.ColRM_IdentifierWO.ReadOnly = True
        Me.ColRM_IdentifierWO.Visible = False
        Me.ColRM_IdentifierWO.Width = 109
        '
        'ColRM_IdentifierItem
        '
        Me.ColRM_IdentifierItem.HeaderText = "Identifier Item"
        Me.ColRM_IdentifierItem.Name = "ColRM_IdentifierItem"
        Me.ColRM_IdentifierItem.ReadOnly = True
        Me.ColRM_IdentifierItem.Visible = False
        Me.ColRM_IdentifierItem.Width = 117
        '
        'ColRM_PrimaryProcess
        '
        Me.ColRM_PrimaryProcess.HeaderText = "Primary Process"
        Me.ColRM_PrimaryProcess.Name = "ColRM_PrimaryProcess"
        Me.ColRM_PrimaryProcess.ReadOnly = True
        Me.ColRM_PrimaryProcess.Visible = False
        Me.ColRM_PrimaryProcess.Width = 136
        '
        'ColRM_PrimaryCodeNo
        '
        Me.ColRM_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColRM_PrimaryCodeNo.Name = "ColRM_PrimaryCodeNo"
        Me.ColRM_PrimaryCodeNo.ReadOnly = True
        Me.ColRM_PrimaryCodeNo.Visible = False
        Me.ColRM_PrimaryCodeNo.Width = 140
        '
        'ColRM_PrimaryUnit
        '
        Me.ColRM_PrimaryUnit.HeaderText = "Primary Unit"
        Me.ColRM_PrimaryUnit.Name = "ColRM_PrimaryUnit"
        Me.ColRM_PrimaryUnit.ReadOnly = True
        Me.ColRM_PrimaryUnit.Visible = False
        Me.ColRM_PrimaryUnit.Width = 110
        '
        'ColRM_IsWIP
        '
        Me.ColRM_IsWIP.HeaderText = "Is WIP"
        Me.ColRM_IsWIP.Name = "ColRM_IsWIP"
        Me.ColRM_IsWIP.ReadOnly = True
        Me.ColRM_IsWIP.Visible = False
        Me.ColRM_IsWIP.Width = 72
        '
        'FrmTrans_ManufacturingOrder_DetailRM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(892, 413)
        Me.Controls.Add(Me.DGVRM)
        Me.Controls.Add(Me.Button1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(900, 450)
        Me.Name = "FrmTrans_ManufacturingOrder_DetailRM"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Raw Material Detail Qty"
        CType(Me.DGVRM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DGVRM As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ColRM_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_WONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_Process As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_QtyTarget As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_QtyRequired As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRM_QtyAllocated As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_QtyAdditional As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRM_QtyUsed As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRM_QtyBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_GPTarget As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_GPActual As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_StandByLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_PickUpLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_PrimaryStandByLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_PrimaryPickUpLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_IdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_IdentifierItem As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_PrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRM_IsWIP As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
