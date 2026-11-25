<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Export_eSPT_Check1
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
        Me.dgvDetail = New System.Windows.Forms.DataGridView
        Me.ColPeriod = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGrossIncome = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrevPeriodIncomeTax = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncomeTax = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvDetail
        '
        Me.dgvDetail.AllowUserToAddRows = False
        Me.dgvDetail.AllowUserToDeleteRows = False
        Me.dgvDetail.AllowUserToResizeColumns = False
        Me.dgvDetail.AllowUserToResizeRows = False
        Me.dgvDetail.ColumnHeadersHeight = 25
        Me.dgvDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColPeriod, Me.ColGrossIncome, Me.ColPrevPeriodIncomeTax, Me.ColIncomeTax})
        Me.dgvDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDetail.Location = New System.Drawing.Point(0, 0)
        Me.dgvDetail.Name = "dgvDetail"
        Me.dgvDetail.ReadOnly = True
        Me.dgvDetail.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvDetail.RowTemplate.Height = 24
        Me.dgvDetail.Size = New System.Drawing.Size(535, 496)
        Me.dgvDetail.TabIndex = 0
        '
        'ColPeriod
        '
        Me.ColPeriod.HeaderText = "Period"
        Me.ColPeriod.Name = "ColPeriod"
        Me.ColPeriod.ReadOnly = True
        Me.ColPeriod.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColGrossIncome
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColGrossIncome.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColGrossIncome.HeaderText = "Gross Income"
        Me.ColGrossIncome.Name = "ColGrossIncome"
        Me.ColGrossIncome.ReadOnly = True
        Me.ColGrossIncome.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColPrevPeriodIncomeTax
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPrevPeriodIncomeTax.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColPrevPeriodIncomeTax.HeaderText = "Prev. Period Income Tax"
        Me.ColPrevPeriodIncomeTax.Name = "ColPrevPeriodIncomeTax"
        Me.ColPrevPeriodIncomeTax.ReadOnly = True
        Me.ColPrevPeriodIncomeTax.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColIncomeTax
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColIncomeTax.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColIncomeTax.HeaderText = "Income Tax"
        Me.ColIncomeTax.Name = "ColIncomeTax"
        Me.ColIncomeTax.ReadOnly = True
        Me.ColIncomeTax.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(396, 117)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 1
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Button1"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'FrmTrans_Export_eSPT_Check1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(535, 496)
        Me.Controls.Add(Me.dgvDetail)
        Me.Controls.Add(Me.btnClose)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "FrmTrans_Export_eSPT_Check1"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Gross Income Data"
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvDetail As System.Windows.Forms.DataGridView
    Friend WithEvents ColPeriod As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGrossIncome As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrevPeriodIncomeTax As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncomeTax As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnClose As System.Windows.Forms.Button
End Class
