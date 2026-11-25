<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Gaji_Payment
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
        Me.btnExit = New System.Windows.Forms.Button
        Me.DGVPayment = New System.Windows.Forms.DataGridView
        Me.BackgroundWorker_Payment = New System.ComponentModel.BackgroundWorker
        Me.ColDetailPayment_AccountCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetailPayment_AccountDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetailPayment_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetailPayment_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTambahan = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColDetailPayment_PrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetailPayment_Identifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetailPayment_IndexRowSalary = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.DGVPayment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(326, 166)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 7
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'DGVPayment
        '
        Me.DGVPayment.AllowUserToAddRows = False
        Me.DGVPayment.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVPayment.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDetailPayment_AccountCode, Me.ColDetailPayment_AccountDesc, Me.ColDetailPayment_Amount, Me.ColDetailPayment_Remarks, Me.ColTambahan, Me.ColDetailPayment_PrimaryAkun, Me.ColDetailPayment_Identifier, Me.ColDetailPayment_IndexRowSalary})
        Me.DGVPayment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVPayment.Location = New System.Drawing.Point(0, 0)
        Me.DGVPayment.Margin = New System.Windows.Forms.Padding(2)
        Me.DGVPayment.MultiSelect = False
        Me.DGVPayment.Name = "DGVPayment"
        Me.DGVPayment.RowTemplate.Height = 24
        Me.DGVPayment.Size = New System.Drawing.Size(732, 358)
        Me.DGVPayment.TabIndex = 8
        '
        'BackgroundWorker_Payment
        '
        '
        'ColDetailPayment_AccountCode
        '
        Me.ColDetailPayment_AccountCode.HeaderText = "Account Code"
        Me.ColDetailPayment_AccountCode.Name = "ColDetailPayment_AccountCode"
        Me.ColDetailPayment_AccountCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetailPayment_AccountCode.Width = 81
        '
        'ColDetailPayment_AccountDesc
        '
        Me.ColDetailPayment_AccountDesc.HeaderText = "Description"
        Me.ColDetailPayment_AccountDesc.Name = "ColDetailPayment_AccountDesc"
        Me.ColDetailPayment_AccountDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetailPayment_AccountDesc.Width = 66
        '
        'ColDetailPayment_Amount
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDetailPayment_Amount.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColDetailPayment_Amount.HeaderText = "Amount"
        Me.ColDetailPayment_Amount.Name = "ColDetailPayment_Amount"
        Me.ColDetailPayment_Amount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetailPayment_Amount.Width = 49
        '
        'ColDetailPayment_Remarks
        '
        Me.ColDetailPayment_Remarks.HeaderText = "Remarks"
        Me.ColDetailPayment_Remarks.Name = "ColDetailPayment_Remarks"
        Me.ColDetailPayment_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetailPayment_Remarks.Width = 55
        '
        'ColTambahan
        '
        Me.ColTambahan.HeaderText = "++"
        Me.ColTambahan.Name = "ColTambahan"
        Me.ColTambahan.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColTambahan.Text = ""
        Me.ColTambahan.Width = 25
        '
        'ColDetailPayment_PrimaryAkun
        '
        Me.ColDetailPayment_PrimaryAkun.HeaderText = "PrimaryAkun"
        Me.ColDetailPayment_PrimaryAkun.Name = "ColDetailPayment_PrimaryAkun"
        Me.ColDetailPayment_PrimaryAkun.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetailPayment_PrimaryAkun.Visible = False
        Me.ColDetailPayment_PrimaryAkun.Width = 72
        '
        'ColDetailPayment_Identifier
        '
        Me.ColDetailPayment_Identifier.HeaderText = "Identifier"
        Me.ColDetailPayment_Identifier.Name = "ColDetailPayment_Identifier"
        Me.ColDetailPayment_Identifier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetailPayment_Identifier.Visible = False
        Me.ColDetailPayment_Identifier.Width = 53
        '
        'ColDetailPayment_IndexRowSalary
        '
        Me.ColDetailPayment_IndexRowSalary.HeaderText = "Index Row Salary"
        Me.ColDetailPayment_IndexRowSalary.Name = "ColDetailPayment_IndexRowSalary"
        Me.ColDetailPayment_IndexRowSalary.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetailPayment_IndexRowSalary.Visible = False
        Me.ColDetailPayment_IndexRowSalary.Width = 96
        '
        'FrmTrans_Gaji_Payment
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(732, 358)
        Me.Controls.Add(Me.DGVPayment)
        Me.Controls.Add(Me.btnExit)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(748, 396)
        Me.Name = "FrmTrans_Gaji_Payment"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Payment Account Data"
        CType(Me.DGVPayment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents DGVPayment As System.Windows.Forms.DataGridView
    Friend WithEvents BackgroundWorker_Payment As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColDetailPayment_AccountCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetailPayment_AccountDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetailPayment_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetailPayment_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTambahan As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColDetailPayment_PrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetailPayment_Identifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetailPayment_IndexRowSalary As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
