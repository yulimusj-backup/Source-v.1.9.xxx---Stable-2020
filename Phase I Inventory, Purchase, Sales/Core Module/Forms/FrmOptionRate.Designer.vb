<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmOptionRate
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
        Me.BtnOK = New System.Windows.Forms.Button
        Me.DgvRate = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCurr = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.DgvRate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BtnOK
        '
        Me.BtnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnOK.Location = New System.Drawing.Point(397, 289)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(110, 31)
        Me.BtnOK.TabIndex = 0
        Me.BtnOK.Text = "&OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'DgvRate
        '
        Me.DgvRate.AllowUserToAddRows = False
        Me.DgvRate.AllowUserToDeleteRows = False
        Me.DgvRate.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvRate.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvRate.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCurrency, Me.ColRate, Me.ColPrimaryCurr})
        Me.DgvRate.Location = New System.Drawing.Point(12, 12)
        Me.DgvRate.Name = "DgvRate"
        Me.DgvRate.RowTemplate.Height = 24
        Me.DgvRate.Size = New System.Drawing.Size(494, 268)
        Me.DgvRate.TabIndex = 1
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 50
        '
        'ColCurrency
        '
        Me.ColCurrency.HeaderText = "Currency"
        Me.ColCurrency.Name = "ColCurrency"
        '
        'ColRate
        '
        Me.ColRate.HeaderText = "Rate"
        Me.ColRate.Name = "ColRate"
        '
        'ColPrimaryCurr
        '
        Me.ColPrimaryCurr.HeaderText = "Primary Curr"
        Me.ColPrimaryCurr.Name = "ColPrimaryCurr"
        Me.ColPrimaryCurr.Visible = False
        '
        'FrmOptionRate
        '
        Me.AcceptButton = Me.BtnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(518, 330)
        Me.Controls.Add(Me.DgvRate)
        Me.Controls.Add(Me.BtnOK)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmOptionRate"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FrmOptionRate"
        CType(Me.DgvRate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents BtnOK As System.Windows.Forms.Button
    Public WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents ColCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents ColRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents ColPrimaryCurr As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents DgvRate As System.Windows.Forms.DataGridView
End Class
