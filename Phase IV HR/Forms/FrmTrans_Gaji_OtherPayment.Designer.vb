<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Gaji_OtherPayment
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
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColComponentName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeKomponen = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColTaxElement = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifierKomponen = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryRefDoc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnOK = New System.Windows.Forms.Button
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 30
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColComponentName, Me.ColDocNo, Me.ColDate, Me.ColTipeKomponen, Me.ColTaxElement, Me.ColAmount, Me.ColIdentifierKomponen, Me.ColIsRutin, Me.ColPrimaryRefDoc})
        Me.dgvItem.Location = New System.Drawing.Point(9, 9)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowHeadersVisible = False
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(494, 226)
        Me.dgvItem.TabIndex = 2
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColComponentName
        '
        Me.ColComponentName.HeaderText = "Component Name"
        Me.ColComponentName.Name = "ColComponentName"
        Me.ColComponentName.ReadOnly = True
        Me.ColComponentName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColComponentName.Width = 98
        '
        'ColDocNo
        '
        Me.ColDocNo.HeaderText = "Document No"
        Me.ColDocNo.Name = "ColDocNo"
        Me.ColDocNo.ReadOnly = True
        Me.ColDocNo.Width = 98
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.Width = 55
        '
        'ColTipeKomponen
        '
        Me.ColTipeKomponen.HeaderText = "Type"
        Me.ColTipeKomponen.Items.AddRange(New Object() {"Income", "Deduction", "Company Expense", "Employee Expense"})
        Me.ColTipeKomponen.Name = "ColTipeKomponen"
        Me.ColTipeKomponen.ReadOnly = True
        Me.ColTipeKomponen.Width = 37
        '
        'ColTaxElement
        '
        Me.ColTaxElement.HeaderText = "Tax Element"
        Me.ColTaxElement.Items.AddRange(New Object() {"GAJI/PENSIUN ATAU THT/JHT", "TUNJANGAN LAINNYA, UANG LEMBUR DAN SEBAGAINYA", "HONORARIUM DAN IMBALAN LAIN SEJENISNYA", "PREMI ASURANSI YANG DIBAYARKAN PEMBERI KERJA", "PENERIMAAN DALAM BENTUK NATURA DAN KENIKMATAN LAINNYA YANG DIKENAKAN PEMOTONGAN P" & _
                        "Ph PASAL 21", "TANTIEM, BONUS, GRATIFIKASI, JASA PRODUKSI DAN THR", "IURAN PENSIUN ATAU IURAN THT/JHT", "BUKAN OBJEK PAJAK", "PESANGON"})
        Me.ColTaxElement.Name = "ColTaxElement"
        Me.ColTaxElement.ReadOnly = True
        Me.ColTaxElement.Visible = False
        Me.ColTaxElement.Width = 72
        '
        'ColAmount
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.ReadOnly = True
        Me.ColAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAmount.Width = 49
        '
        'ColIdentifierKomponen
        '
        Me.ColIdentifierKomponen.HeaderText = "Identifier Salary Component"
        Me.ColIdentifierKomponen.Name = "ColIdentifierKomponen"
        Me.ColIdentifierKomponen.ReadOnly = True
        Me.ColIdentifierKomponen.Visible = False
        Me.ColIdentifierKomponen.Width = 161
        '
        'ColIsRutin
        '
        Me.ColIsRutin.HeaderText = "IsRutin"
        Me.ColIsRutin.Name = "ColIsRutin"
        Me.ColIsRutin.ReadOnly = True
        Me.ColIsRutin.Visible = False
        Me.ColIsRutin.Width = 65
        '
        'ColPrimaryRefDoc
        '
        Me.ColPrimaryRefDoc.HeaderText = "Primary Ref Doc"
        Me.ColPrimaryRefDoc.Name = "ColPrimaryRefDoc"
        Me.ColPrimaryRefDoc.ReadOnly = True
        Me.ColPrimaryRefDoc.Visible = False
        Me.ColPrimaryRefDoc.Width = 109
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnOK.Location = New System.Drawing.Point(447, 243)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 19)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmTrans_Gaji_OtherPayment
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnOK
        Me.ClientSize = New System.Drawing.Size(512, 271)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.dgvItem)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTrans_Gaji_OtherPayment"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Other Payment"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColComponentName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeKomponen As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColTaxElement As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifierKomponen As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryRefDoc As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
