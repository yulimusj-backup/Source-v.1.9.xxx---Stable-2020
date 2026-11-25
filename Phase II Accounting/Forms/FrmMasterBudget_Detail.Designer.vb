<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterBudget_Detail
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
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.btnModify = New System.Windows.Forms.Button
        Me.TxtNominal = New System.Windows.Forms.TextBox
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.rbo_Percentage = New System.Windows.Forms.RadioButton
        Me.TxtPencentage = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboMonth = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtBeginingBalance = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.ColMonth = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDifference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDifference_Percentage = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSaldo = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColMonth, Me.ColDifference, Me.ColDifference_Percentage, Me.ColSaldo})
        Me.dgvItem.Location = New System.Drawing.Point(16, 15)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(636, 412)
        Me.dgvItem.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.TxtPrimary)
        Me.GroupBox1.Controls.Add(Me.btnModify)
        Me.GroupBox1.Controls.Add(Me.TxtNominal)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.rbo_Percentage)
        Me.GroupBox1.Controls.Add(Me.TxtPencentage)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.cboMonth)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.TxtBeginingBalance)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(35, 434)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(597, 137)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Quick Add"
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(12, 105)
        Me.TxtPrimary.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(73, 22)
        Me.TxtPrimary.TabIndex = 10
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "0"
        Me.TxtPrimary.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtPrimary.Visible = False
        '
        'btnModify
        '
        Me.btnModify.Location = New System.Drawing.Point(476, 55)
        Me.btnModify.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnModify.Name = "btnModify"
        Me.btnModify.Size = New System.Drawing.Size(108, 28)
        Me.btnModify.TabIndex = 6
        Me.btnModify.Text = "Modify"
        Me.btnModify.UseVisualStyleBackColor = True
        '
        'TxtNominal
        '
        Me.TxtNominal.Enabled = False
        Me.TxtNominal.Location = New System.Drawing.Point(259, 86)
        Me.TxtNominal.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtNominal.Name = "TxtNominal"
        Me.TxtNominal.Size = New System.Drawing.Size(168, 22)
        Me.TxtNominal.TabIndex = 5
        Me.TxtNominal.Text = "0"
        Me.TxtNominal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(144, 87)
        Me.RadioButton2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(80, 21)
        Me.RadioButton2.TabIndex = 3
        Me.RadioButton2.Text = "Nominal"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'rbo_Percentage
        '
        Me.rbo_Percentage.AutoSize = True
        Me.rbo_Percentage.Checked = True
        Me.rbo_Percentage.Location = New System.Drawing.Point(144, 59)
        Me.rbo_Percentage.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.rbo_Percentage.Name = "rbo_Percentage"
        Me.rbo_Percentage.Size = New System.Drawing.Size(102, 21)
        Me.rbo_Percentage.TabIndex = 2
        Me.rbo_Percentage.TabStop = True
        Me.rbo_Percentage.Text = "Percentage"
        Me.rbo_Percentage.UseVisualStyleBackColor = True
        '
        'TxtPencentage
        '
        Me.TxtPencentage.Location = New System.Drawing.Point(259, 58)
        Me.TxtPencentage.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPencentage.Name = "TxtPencentage"
        Me.TxtPencentage.Size = New System.Drawing.Size(168, 22)
        Me.TxtPencentage.TabIndex = 4
        Me.TxtPencentage.Text = "0"
        Me.TxtPencentage.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 59)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 17)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Differrence"
        '
        'cboMonth
        '
        Me.cboMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMonth.FormattingEnabled = True
        Me.cboMonth.Items.AddRange(New Object() {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"})
        Me.cboMonth.Location = New System.Drawing.Point(355, 23)
        Me.cboMonth.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboMonth.Name = "cboMonth"
        Me.cboMonth.Size = New System.Drawing.Size(228, 24)
        Me.cboMonth.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(321, 27)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(24, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "on"
        '
        'TxtBeginingBalance
        '
        Me.TxtBeginingBalance.Location = New System.Drawing.Point(144, 23)
        Me.TxtBeginingBalance.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtBeginingBalance.Name = "TxtBeginingBalance"
        Me.TxtBeginingBalance.Size = New System.Drawing.Size(168, 22)
        Me.TxtBeginingBalance.TabIndex = 0
        Me.TxtBeginingBalance.Text = "0"
        Me.TxtBeginingBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 27)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(126, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Beginning Balance"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(532, 578)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(100, 28)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(424, 578)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(100, 28)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'ColMonth
        '
        Me.ColMonth.HeaderText = "Month"
        Me.ColMonth.Name = "ColMonth"
        Me.ColMonth.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColDifference
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDifference.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColDifference.HeaderText = "Difference"
        Me.ColDifference.Name = "ColDifference"
        Me.ColDifference.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColDifference_Percentage
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDifference_Percentage.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColDifference_Percentage.HeaderText = "% Diference"
        Me.ColDifference_Percentage.Name = "ColDifference_Percentage"
        Me.ColDifference_Percentage.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColSaldo
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSaldo.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColSaldo.HeaderText = "Saldo"
        Me.ColSaldo.Name = "ColSaldo"
        Me.ColSaldo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'FrmMasterBudget_Detail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(679, 620)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.dgvItem)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(687, 657)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(687, 657)
        Me.Name = "FrmMasterBudget_Detail"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Budget"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TxtNominal As System.Windows.Forms.TextBox
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents rbo_Percentage As System.Windows.Forms.RadioButton
    Friend WithEvents TxtPencentage As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboMonth As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtBeginingBalance As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnModify As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents ColMonth As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDifference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDifference_Percentage As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSaldo As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
