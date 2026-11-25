<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterSalaryFormula_OtherDeduction_ItemEmp
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
        Me.DgvItemDetail = New System.Windows.Forms.DataGridView
        Me.DataTable2BindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DseTrans_OtherDeduction = New Core_Module.DseTrans_OtherDeduction
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtEmployeeCode = New System.Windows.Forms.TextBox
        Me.TxtEmployeeName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.DataTable2TableAdapter = New Core_Module.DseTrans_OtherDeductionTableAdapters.DataTable2TableAdapter
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewLinkColumn
        Me.DescriptionDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PrimaryDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PrimaryEmployeeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.DgvItemDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataTable2BindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DseTrans_OtherDeduction, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DgvItemDetail
        '
        Me.DgvItemDetail.AllowUserToAddRows = False
        Me.DgvItemDetail.AllowUserToDeleteRows = False
        Me.DgvItemDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItemDetail.AutoGenerateColumns = False
        Me.DgvItemDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvItemDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.CodeDataGridViewTextBoxColumn, Me.DescriptionDataGridViewTextBoxColumn, Me.PrimaryDataGridViewTextBoxColumn, Me.PrimaryEmployeeDataGridViewTextBoxColumn})
        Me.DgvItemDetail.DataSource = Me.DataTable2BindingSource
        Me.DgvItemDetail.Location = New System.Drawing.Point(9, 62)
        Me.DgvItemDetail.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvItemDetail.Name = "DgvItemDetail"
        Me.DgvItemDetail.ReadOnly = True
        Me.DgvItemDetail.RowTemplate.Height = 24
        Me.DgvItemDetail.Size = New System.Drawing.Size(464, 246)
        Me.DgvItemDetail.TabIndex = 0
        '
        'DataTable2BindingSource
        '
        Me.DataTable2BindingSource.DataMember = "DataTable2"
        Me.DataTable2BindingSource.DataSource = Me.DseTrans_OtherDeduction
        '
        'DseTrans_OtherDeduction
        '
        Me.DseTrans_OtherDeduction.DataSetName = "DseTrans_OtherDeduction"
        Me.DseTrans_OtherDeduction.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 12)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Employee Code"
        '
        'TxtEmployeeCode
        '
        Me.TxtEmployeeCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEmployeeCode.Location = New System.Drawing.Point(94, 10)
        Me.TxtEmployeeCode.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtEmployeeCode.Name = "TxtEmployeeCode"
        Me.TxtEmployeeCode.ReadOnly = True
        Me.TxtEmployeeCode.Size = New System.Drawing.Size(380, 20)
        Me.TxtEmployeeCode.TabIndex = 2
        '
        'TxtEmployeeName
        '
        Me.TxtEmployeeName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEmployeeName.Location = New System.Drawing.Point(94, 32)
        Me.TxtEmployeeName.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtEmployeeName.Name = "TxtEmployeeName"
        Me.TxtEmployeeName.ReadOnly = True
        Me.TxtEmployeeName.Size = New System.Drawing.Size(380, 20)
        Me.TxtEmployeeName.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 35)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Name"
        '
        'DataTable2TableAdapter
        '
        Me.DataTable2TableAdapter.ClearBeforeFill = True
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        '
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "Code"
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "Code"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        Me.CodeDataGridViewTextBoxColumn.ReadOnly = True
        Me.CodeDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.CodeDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'DescriptionDataGridViewTextBoxColumn
        '
        Me.DescriptionDataGridViewTextBoxColumn.DataPropertyName = "Description"
        Me.DescriptionDataGridViewTextBoxColumn.HeaderText = "Description"
        Me.DescriptionDataGridViewTextBoxColumn.Name = "DescriptionDataGridViewTextBoxColumn"
        Me.DescriptionDataGridViewTextBoxColumn.ReadOnly = True
        '
        'PrimaryDataGridViewTextBoxColumn
        '
        Me.PrimaryDataGridViewTextBoxColumn.DataPropertyName = "Primary"
        Me.PrimaryDataGridViewTextBoxColumn.HeaderText = "Primary"
        Me.PrimaryDataGridViewTextBoxColumn.Name = "PrimaryDataGridViewTextBoxColumn"
        Me.PrimaryDataGridViewTextBoxColumn.ReadOnly = True
        Me.PrimaryDataGridViewTextBoxColumn.Visible = False
        '
        'PrimaryEmployeeDataGridViewTextBoxColumn
        '
        Me.PrimaryEmployeeDataGridViewTextBoxColumn.DataPropertyName = "primary employee"
        Me.PrimaryEmployeeDataGridViewTextBoxColumn.HeaderText = "primary employee"
        Me.PrimaryEmployeeDataGridViewTextBoxColumn.Name = "PrimaryEmployeeDataGridViewTextBoxColumn"
        Me.PrimaryEmployeeDataGridViewTextBoxColumn.ReadOnly = True
        Me.PrimaryEmployeeDataGridViewTextBoxColumn.Visible = False
        '
        'FrmMasterSalaryFormula_OtherDeduction_ItemEmp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(480, 318)
        Me.Controls.Add(Me.TxtEmployeeName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtEmployeeCode)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DgvItemDetail)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "FrmMasterSalaryFormula_OtherDeduction_ItemEmp"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Other Deduction"
        CType(Me.DgvItemDetail, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataTable2BindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DseTrans_OtherDeduction, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DgvItemDetail As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtEmployeeCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DseTrans_OtherDeduction As Core_Module.DseTrans_OtherDeduction
    Friend WithEvents DataTable2BindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents DataTable2TableAdapter As Core_Module.DseTrans_OtherDeductionTableAdapters.DataTable2TableAdapter
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents DescriptionDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PrimaryDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PrimaryEmployeeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
