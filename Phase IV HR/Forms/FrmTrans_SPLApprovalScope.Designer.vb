<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SPLApprovalScope
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
        Me.dgvScope = New System.Windows.Forms.DataGridView
        Me.colNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDepartmentCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDepartmentDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJobPositionCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJobPositionDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPosition = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryEmp = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvScope, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvScope
        '
        Me.dgvScope.AllowUserToAddRows = False
        Me.dgvScope.AllowUserToDeleteRows = False
        Me.dgvScope.AllowUserToResizeColumns = False
        Me.dgvScope.AllowUserToResizeRows = False
        Me.dgvScope.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvScope.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvScope.ColumnHeadersHeight = 25
        Me.dgvScope.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvScope.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colNo, Me.ColDepartmentCode, Me.ColDepartmentDesc, Me.ColJobPositionCode, Me.ColJobPositionDesc, Me.ColEmployeeCode, Me.ColEmployeeName, Me.colRemarks, Me.ColPrimaryDept, Me.ColPrimaryPosition, Me.ColPrimaryEmp})
        Me.dgvScope.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvScope.Location = New System.Drawing.Point(0, 0)
        Me.dgvScope.Name = "dgvScope"
        Me.dgvScope.ReadOnly = True
        Me.dgvScope.RowHeadersVisible = False
        Me.dgvScope.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvScope.RowTemplate.Height = 24
        Me.dgvScope.Size = New System.Drawing.Size(680, 262)
        Me.dgvScope.TabIndex = 1
        '
        'colNo
        '
        Me.colNo.Frozen = True
        Me.colNo.HeaderText = "No"
        Me.colNo.Name = "colNo"
        Me.colNo.ReadOnly = True
        Me.colNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.colNo.Width = 27
        '
        'ColDepartmentCode
        '
        Me.ColDepartmentCode.HeaderText = "Departement Code"
        Me.ColDepartmentCode.Name = "ColDepartmentCode"
        Me.ColDepartmentCode.ReadOnly = True
        Me.ColDepartmentCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDepartmentCode.Width = 102
        '
        'ColDepartmentDesc
        '
        Me.ColDepartmentDesc.HeaderText = "Description"
        Me.ColDepartmentDesc.Name = "ColDepartmentDesc"
        Me.ColDepartmentDesc.ReadOnly = True
        Me.ColDepartmentDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDepartmentDesc.Width = 66
        '
        'ColJobPositionCode
        '
        Me.ColJobPositionCode.HeaderText = "Job Position Code"
        Me.ColJobPositionCode.Name = "ColJobPositionCode"
        Me.ColJobPositionCode.ReadOnly = True
        Me.ColJobPositionCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJobPositionCode.Width = 98
        '
        'ColJobPositionDesc
        '
        Me.ColJobPositionDesc.HeaderText = "Description"
        Me.ColJobPositionDesc.Name = "ColJobPositionDesc"
        Me.ColJobPositionDesc.ReadOnly = True
        Me.ColJobPositionDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJobPositionDesc.Width = 66
        '
        'ColEmployeeCode
        '
        Me.ColEmployeeCode.HeaderText = "Employee Code"
        Me.ColEmployeeCode.Name = "ColEmployeeCode"
        Me.ColEmployeeCode.ReadOnly = True
        Me.ColEmployeeCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEmployeeCode.Width = 87
        '
        'ColEmployeeName
        '
        Me.ColEmployeeName.HeaderText = "Employee Name"
        Me.ColEmployeeName.Name = "ColEmployeeName"
        Me.ColEmployeeName.ReadOnly = True
        Me.ColEmployeeName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEmployeeName.Visible = False
        Me.ColEmployeeName.Width = 90
        '
        'colRemarks
        '
        Me.colRemarks.HeaderText = "Remarks"
        Me.colRemarks.Name = "colRemarks"
        Me.colRemarks.ReadOnly = True
        Me.colRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.colRemarks.Width = 55
        '
        'ColPrimaryDept
        '
        Me.ColPrimaryDept.HeaderText = "Primary Dept"
        Me.ColPrimaryDept.Name = "ColPrimaryDept"
        Me.ColPrimaryDept.ReadOnly = True
        Me.ColPrimaryDept.Visible = False
        Me.ColPrimaryDept.Width = 92
        '
        'ColPrimaryPosition
        '
        Me.ColPrimaryPosition.HeaderText = "Primary Position"
        Me.ColPrimaryPosition.Name = "ColPrimaryPosition"
        Me.ColPrimaryPosition.ReadOnly = True
        Me.ColPrimaryPosition.Visible = False
        Me.ColPrimaryPosition.Width = 106
        '
        'ColPrimaryEmp
        '
        Me.ColPrimaryEmp.HeaderText = "Primary Employee"
        Me.ColPrimaryEmp.Name = "ColPrimaryEmp"
        Me.ColPrimaryEmp.ReadOnly = True
        Me.ColPrimaryEmp.Visible = False
        Me.ColPrimaryEmp.Width = 115
        '
        'FrmTrans_SPLApprovalScope
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(680, 262)
        Me.Controls.Add(Me.dgvScope)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(696, 300)
        Me.Name = "FrmTrans_SPLApprovalScope"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Approval Scope"
        CType(Me.dgvScope, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvScope As System.Windows.Forms.DataGridView
    Friend WithEvents colNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDepartmentCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDepartmentDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJobPositionCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJobPositionDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPosition As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryEmp As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
