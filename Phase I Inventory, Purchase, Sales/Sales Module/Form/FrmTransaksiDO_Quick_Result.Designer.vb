<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiDO_Quick_Result
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
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colSONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrint = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColPrimaryDO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button1 = New System.Windows.Forms.Button
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DgvItem
        '
        Me.DgvItem.AllowUserToAddRows = False
        Me.DgvItem.AllowUserToDeleteRows = False
        Me.DgvItem.AllowUserToResizeColumns = False
        Me.DgvItem.AllowUserToResizeRows = False
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.colSONo, Me.ColDONo, Me.ColPrint, Me.ColPrimaryDO})
        Me.DgvItem.Location = New System.Drawing.Point(11, 11)
        Me.DgvItem.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.RowTemplate.Height = 24
        Me.DgvItem.Size = New System.Drawing.Size(501, 317)
        Me.DgvItem.TabIndex = 1
        '
        'ColNo
        '
        Me.ColNo.FillWeight = 50.0!
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.ColNo.Width = 50
        '
        'colSONo
        '
        Me.colSONo.HeaderText = "SO No"
        Me.colSONo.Name = "colSONo"
        Me.colSONo.ReadOnly = True
        '
        'ColDONo
        '
        Me.ColDONo.HeaderText = "DO No"
        Me.ColDONo.Name = "ColDONo"
        Me.ColDONo.ReadOnly = True
        '
        'ColPrint
        '
        Me.ColPrint.HeaderText = "Print"
        Me.ColPrint.Name = "ColPrint"
        Me.ColPrint.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPrint.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'ColPrimaryDO
        '
        Me.ColPrimaryDO.HeaderText = "Primary DO"
        Me.ColPrimaryDO.Name = "ColPrimaryDO"
        Me.ColPrimaryDO.Visible = False
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(446, 332)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(66, 24)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "&OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiDO_Quick_Result
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(523, 367)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DgvItem)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(531, 398)
        Me.Name = "FrmTransaksiDO_Quick_Result"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FrmTransaksiDO_Quick_Result"
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colSONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrint As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColPrimaryDO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
