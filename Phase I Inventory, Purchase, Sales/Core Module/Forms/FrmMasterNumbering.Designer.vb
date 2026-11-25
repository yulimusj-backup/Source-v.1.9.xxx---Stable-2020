<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterNumbering
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboTransaction = New System.Windows.Forms.ComboBox
        Me.chkAutoNumber = New System.Windows.Forms.CheckBox
        Me.cboReset = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColTipeKomponen = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColTipe_Int = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFormat = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColFormatInt = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtResult = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnSimpan = New System.Windows.Forms.Button
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 18)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(83, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Transaction"
        '
        'cboTransaction
        '
        Me.cboTransaction.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboTransaction.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTransaction.FormattingEnabled = True
        Me.cboTransaction.Location = New System.Drawing.Point(180, 15)
        Me.cboTransaction.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboTransaction.Name = "cboTransaction"
        Me.cboTransaction.Size = New System.Drawing.Size(579, 24)
        Me.cboTransaction.TabIndex = 0
        '
        'chkAutoNumber
        '
        Me.chkAutoNumber.AutoSize = True
        Me.chkAutoNumber.Location = New System.Drawing.Point(180, 50)
        Me.chkAutoNumber.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkAutoNumber.Name = "chkAutoNumber"
        Me.chkAutoNumber.Size = New System.Drawing.Size(109, 21)
        Me.chkAutoNumber.TabIndex = 1
        Me.chkAutoNumber.Text = "AutoNumber"
        Me.chkAutoNumber.UseVisualStyleBackColor = True
        '
        'cboReset
        '
        Me.cboReset.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboReset.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboReset.FormattingEnabled = True
        Me.cboReset.Items.AddRange(New Object() {"Yearly", "Monthly"})
        Me.cboReset.Location = New System.Drawing.Point(487, 48)
        Me.cboReset.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboReset.Name = "cboReset"
        Me.cboReset.Size = New System.Drawing.Size(272, 24)
        Me.cboReset.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(360, 52)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(118, 17)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Reset Numbering"
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(115, 15)
        Me.TxtPrimary.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(56, 22)
        Me.TxtPrimary.TabIndex = 5
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.dgvItem)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TxtResult)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 81)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(744, 458)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Autonumber Components"
        '
        'dgvItem
        '
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColTipeKomponen, Me.ColTipe_Int, Me.ColFormat, Me.ColFormatInt, Me.ColValue})
        Me.dgvItem.Location = New System.Drawing.Point(8, 55)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(728, 395)
        Me.dgvItem.TabIndex = 1
        '
        'ColTipeKomponen
        '
        Me.ColTipeKomponen.HeaderText = "Type"
        Me.ColTipeKomponen.Items.AddRange(New Object() {"Fixed Text", "Month", "Year", "Running Number"})
        Me.ColTipeKomponen.Name = "ColTipeKomponen"
        Me.ColTipeKomponen.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColTipeKomponen.Width = 46
        '
        'ColTipe_Int
        '
        Me.ColTipe_Int.HeaderText = "Tipe"
        Me.ColTipe_Int.Name = "ColTipe_Int"
        Me.ColTipe_Int.Visible = False
        Me.ColTipe_Int.Width = 53
        '
        'ColFormat
        '
        Me.ColFormat.HeaderText = "Format"
        Me.ColFormat.Name = "ColFormat"
        Me.ColFormat.Width = 58
        '
        'ColFormatInt
        '
        Me.ColFormatInt.HeaderText = "Format Int"
        Me.ColFormatInt.Name = "ColFormatInt"
        Me.ColFormatInt.Visible = False
        Me.ColFormatInt.Width = 79
        '
        'ColValue
        '
        Me.ColValue.HeaderText = "Value"
        Me.ColValue.Name = "ColValue"
        Me.ColValue.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColValue.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColValue.Width = 50
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 27)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 17)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Result"
        '
        'TxtResult
        '
        Me.TxtResult.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtResult.Location = New System.Drawing.Point(164, 23)
        Me.TxtResult.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtResult.Name = "TxtResult"
        Me.TxtResult.ReadOnly = True
        Me.TxtResult.Size = New System.Drawing.Size(571, 22)
        Me.TxtResult.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(660, 546)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 28)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnSimpan
        '
        Me.btnSimpan.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSimpan.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnSimpan.Location = New System.Drawing.Point(552, 546)
        Me.btnSimpan.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnSimpan.Name = "btnSimpan"
        Me.btnSimpan.Size = New System.Drawing.Size(100, 28)
        Me.btnSimpan.TabIndex = 4
        Me.btnSimpan.Text = "Save"
        Me.btnSimpan.UseVisualStyleBackColor = True
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.Core_Module.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(16, 546)
        Me.btnUP.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(37, 32)
        Me.btnUP.TabIndex = 215
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = Global.Core_Module.My.Resources.Resources.arrow_down
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(61, 546)
        Me.btnDown.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(37, 32)
        Me.btnDown.TabIndex = 216
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'FrmMasterNumbering
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(776, 590)
        Me.Controls.Add(Me.btnUP)
        Me.Controls.Add(Me.btnDown)
        Me.Controls.Add(Me.btnSimpan)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TxtPrimary)
        Me.Controls.Add(Me.cboReset)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.chkAutoNumber)
        Me.Controls.Add(Me.cboTransaction)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimumSize = New System.Drawing.Size(791, 625)
        Me.Name = "FrmMasterNumbering"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Transaction Numbering"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboTransaction As System.Windows.Forms.ComboBox
    Friend WithEvents chkAutoNumber As System.Windows.Forms.CheckBox
    Friend WithEvents cboReset As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtResult As System.Windows.Forms.TextBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSimpan As System.Windows.Forms.Button
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents ColTipeKomponen As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColTipe_Int As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFormat As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColFormatInt As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColValue As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
