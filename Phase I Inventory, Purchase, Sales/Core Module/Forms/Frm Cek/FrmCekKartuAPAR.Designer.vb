<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCekKartuAPAR
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
        Me.btnStart = New System.Windows.Forms.Button
        Me.btnFindCurrency = New System.Windows.Forms.Button
        Me.dgvCek = New System.Windows.Forms.DataGridView
        Me.rboAP = New System.Windows.Forms.RadioButton
        Me.rboAR = New System.Windows.Forms.RadioButton
        Me.TxtCurrencyCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRefPartyCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRefPartyName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalanceOriginal = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvCek, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(12, 12)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 25)
        Me.btnStart.TabIndex = 0
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'btnFindCurrency
        '
        Me.btnFindCurrency.Location = New System.Drawing.Point(377, 12)
        Me.btnFindCurrency.Name = "btnFindCurrency"
        Me.btnFindCurrency.Size = New System.Drawing.Size(75, 25)
        Me.btnFindCurrency.TabIndex = 1
        Me.btnFindCurrency.Text = "Find"
        Me.btnFindCurrency.UseVisualStyleBackColor = True
        '
        'dgvCek
        '
        Me.dgvCek.AllowUserToAddRows = False
        Me.dgvCek.AllowUserToDeleteRows = False
        Me.dgvCek.AllowUserToResizeColumns = False
        Me.dgvCek.AllowUserToResizeRows = False
        Me.dgvCek.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCek.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCek.ColumnHeadersHeight = 25
        Me.dgvCek.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvCek.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColRefPartyCode, Me.ColRefPartyName, Me.ColBalance, Me.ColBalanceOriginal})
        Me.dgvCek.Location = New System.Drawing.Point(12, 43)
        Me.dgvCek.Name = "dgvCek"
        Me.dgvCek.ReadOnly = True
        Me.dgvCek.RowTemplate.Height = 24
        Me.dgvCek.Size = New System.Drawing.Size(754, 208)
        Me.dgvCek.TabIndex = 2
        '
        'rboAP
        '
        Me.rboAP.AutoSize = True
        Me.rboAP.Checked = True
        Me.rboAP.Location = New System.Drawing.Point(93, 14)
        Me.rboAP.Name = "rboAP"
        Me.rboAP.Size = New System.Drawing.Size(47, 21)
        Me.rboAP.TabIndex = 3
        Me.rboAP.Text = "AP"
        Me.rboAP.UseVisualStyleBackColor = True
        '
        'rboAR
        '
        Me.rboAR.AutoSize = True
        Me.rboAR.Location = New System.Drawing.Point(146, 14)
        Me.rboAR.Name = "rboAR"
        Me.rboAR.Size = New System.Drawing.Size(48, 21)
        Me.rboAR.TabIndex = 4
        Me.rboAR.Text = "AR"
        Me.rboAR.UseVisualStyleBackColor = True
        '
        'TxtCurrencyCode
        '
        Me.TxtCurrencyCode.Location = New System.Drawing.Point(271, 13)
        Me.TxtCurrencyCode.Name = "TxtCurrencyCode"
        Me.TxtCurrencyCode.Size = New System.Drawing.Size(100, 22)
        Me.TxtCurrencyCode.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(200, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 17)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Currency"
        '
        'txtPrimaryCurrency
        '
        Me.txtPrimaryCurrency.Location = New System.Drawing.Point(458, 13)
        Me.txtPrimaryCurrency.Name = "txtPrimaryCurrency"
        Me.txtPrimaryCurrency.Size = New System.Drawing.Size(41, 22)
        Me.txtPrimaryCurrency.TabIndex = 7
        Me.txtPrimaryCurrency.TabStop = False
        Me.txtPrimaryCurrency.Text = "-1"
        Me.txtPrimaryCurrency.Visible = False
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 51
        '
        'ColRefPartyCode
        '
        Me.ColRefPartyCode.HeaderText = "Ref. Party Code"
        Me.ColRefPartyCode.Name = "ColRefPartyCode"
        Me.ColRefPartyCode.ReadOnly = True
        Me.ColRefPartyCode.Width = 133
        '
        'ColRefPartyName
        '
        Me.ColRefPartyName.HeaderText = "Ref. Party Name"
        Me.ColRefPartyName.Name = "ColRefPartyName"
        Me.ColRefPartyName.ReadOnly = True
        Me.ColRefPartyName.Width = 137
        '
        'ColBalance
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalance.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColBalance.HeaderText = "Balance"
        Me.ColBalance.Name = "ColBalance"
        Me.ColBalance.ReadOnly = True
        Me.ColBalance.Width = 84
        '
        'ColBalanceOriginal
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalanceOriginal.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColBalanceOriginal.HeaderText = "Balance Original"
        Me.ColBalanceOriginal.Name = "ColBalanceOriginal"
        Me.ColBalanceOriginal.ReadOnly = True
        Me.ColBalanceOriginal.Width = 137
        '
        'FrmCekKartuAPAR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(778, 263)
        Me.Controls.Add(Me.txtPrimaryCurrency)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtCurrencyCode)
        Me.Controls.Add(Me.rboAR)
        Me.Controls.Add(Me.rboAP)
        Me.Controls.Add(Me.dgvCek)
        Me.Controls.Add(Me.btnFindCurrency)
        Me.Controls.Add(Me.btnStart)
        Me.Name = "FrmCekKartuAPAR"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cek Kartu Balance vs Balance Original"
        CType(Me.dgvCek, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents btnFindCurrency As System.Windows.Forms.Button
    Friend WithEvents dgvCek As System.Windows.Forms.DataGridView
    Friend WithEvents rboAP As System.Windows.Forms.RadioButton
    Friend WithEvents rboAR As System.Windows.Forms.RadioButton
    Friend WithEvents TxtCurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRefPartyCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRefPartyName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalanceOriginal As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
