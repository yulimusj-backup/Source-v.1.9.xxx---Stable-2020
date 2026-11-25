<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSupplierAuthorizationDetail
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
        Me.ColAuthorizeBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAuthorizationDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAuthorization = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button1 = New System.Windows.Forms.Button
        Me.dgvDetail = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ColAuthorizeBy
        '
        Me.ColAuthorizeBy.HeaderText = "By"
        Me.ColAuthorizeBy.Name = "ColAuthorizeBy"
        Me.ColAuthorizeBy.ReadOnly = True
        Me.ColAuthorizeBy.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAuthorizeBy.Width = 25
        '
        'ColAuthorizationDate
        '
        Me.ColAuthorizationDate.HeaderText = "Authorization Date"
        Me.ColAuthorizationDate.Name = "ColAuthorizationDate"
        Me.ColAuthorizationDate.ReadOnly = True
        Me.ColAuthorizationDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColAuthorization
        '
        Me.ColAuthorization.HeaderText = "Authorization"
        Me.ColAuthorization.Name = "ColAuthorization"
        Me.ColAuthorization.ReadOnly = True
        Me.ColAuthorization.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAuthorization.Width = 74
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(228, 86)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(115, 29)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'dgvDetail
        '
        Me.dgvDetail.AllowUserToAddRows = False
        Me.dgvDetail.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = Core_Module.GetSMESSBackColorAll
        Me.dgvDetail.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvDetail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDetail.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColAuthorization, Me.ColAuthorizeBy, Me.ColAuthorizationDate})
        Me.dgvDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDetail.Location = New System.Drawing.Point(0, 0)
        Me.dgvDetail.MultiSelect = False
        Me.dgvDetail.Name = "dgvDetail"
        Me.dgvDetail.ReadOnly = True
        Me.dgvDetail.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvDetail.Size = New System.Drawing.Size(460, 176)
        Me.dgvDetail.TabIndex = 2
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'FrmTransaksiSupplierAuthorizationDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(460, 176)
        Me.Controls.Add(Me.dgvDetail)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTransaksiSupplierAuthorizationDetail"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Order Authorization Detail"
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ColAuthorizeBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAuthorizationDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAuthorization As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents dgvDetail As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
