<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterCommission_Cari
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
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslQuery = New System.Windows.Forms.ToolStripStatusLabel
        Me.dgvHasil = New System.Windows.Forms.DataGridView
        Me.gboKriteria = New System.Windows.Forms.GroupBox
        Me.txtprimary = New System.Windows.Forms.TextBox
        Me.cboCriteria = New System.Windows.Forms.ComboBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.txtCari = New System.Windows.Forms.TextBox
        Me.StatusStrip1.SuspendLayout()
        CType(Me.dgvHasil, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboKriteria.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslQuery})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 322)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(469, 22)
        Me.StatusStrip1.TabIndex = 31
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslQuery
        '
        Me.tslQuery.Name = "tslQuery"
        Me.tslQuery.Size = New System.Drawing.Size(0, 17)
        '
        'dgvHasil
        '
        Me.dgvHasil.AllowUserToAddRows = False
        Me.dgvHasil.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvHasil.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvHasil.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvHasil.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvHasil.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvHasil.Location = New System.Drawing.Point(12, 67)
        Me.dgvHasil.MultiSelect = False
        Me.dgvHasil.Name = "dgvHasil"
        Me.dgvHasil.ReadOnly = True
        Me.dgvHasil.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvHasil.RowTemplate.Height = 24
        Me.dgvHasil.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvHasil.Size = New System.Drawing.Size(445, 245)
        Me.dgvHasil.TabIndex = 30
        Me.dgvHasil.TabStop = False
        '
        'gboKriteria
        '
        Me.gboKriteria.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboKriteria.Controls.Add(Me.txtprimary)
        Me.gboKriteria.Controls.Add(Me.cboCriteria)
        Me.gboKriteria.Controls.Add(Me.btnOK)
        Me.gboKriteria.Controls.Add(Me.btnClose)
        Me.gboKriteria.Controls.Add(Me.btnFind)
        Me.gboKriteria.Controls.Add(Me.txtCari)
        Me.gboKriteria.Location = New System.Drawing.Point(12, 0)
        Me.gboKriteria.Name = "gboKriteria"
        Me.gboKriteria.Size = New System.Drawing.Size(445, 65)
        Me.gboKriteria.TabIndex = 29
        Me.gboKriteria.TabStop = False
        '
        'txtprimary
        '
        Me.txtprimary.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtprimary.Location = New System.Drawing.Point(133, 38)
        Me.txtprimary.Name = "txtprimary"
        Me.txtprimary.Size = New System.Drawing.Size(29, 20)
        Me.txtprimary.TabIndex = 6
        Me.txtprimary.Text = "-1"
        Me.txtprimary.Visible = False
        '
        'cboCriteria
        '
        Me.cboCriteria.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCriteria.FormattingEnabled = True
        Me.cboCriteria.Items.AddRange(New Object() {"Code", "Description", "Remarks"})
        Me.cboCriteria.Location = New System.Drawing.Point(6, 12)
        Me.cboCriteria.Name = "cboCriteria"
        Me.cboCriteria.Size = New System.Drawing.Size(121, 21)
        Me.cboCriteria.TabIndex = 5
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnOK.Location = New System.Drawing.Point(313, 38)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(60, 23)
        Me.btnOK.TabIndex = 4
        Me.btnOK.TabStop = False
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnClose.Location = New System.Drawing.Point(379, 38)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(60, 23)
        Me.btnClose.TabIndex = 2
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFind.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnFind.Location = New System.Drawing.Point(379, 10)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(60, 23)
        Me.btnFind.TabIndex = 1
        Me.btnFind.TabStop = False
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'txtCari
        '
        Me.txtCari.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCari.Location = New System.Drawing.Point(133, 12)
        Me.txtCari.MaxLength = 150
        Me.txtCari.Name = "txtCari"
        Me.txtCari.Size = New System.Drawing.Size(240, 20)
        Me.txtCari.TabIndex = 0
        '
        'FrmMasterCommission_Cari
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(469, 344)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.dgvHasil)
        Me.Controls.Add(Me.gboKriteria)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximumSize = New System.Drawing.Size(485, 382)
        Me.MinimizeBox = False
        Me.Name = "FrmMasterCommission_Cari"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Find Sales Commission Data"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.dgvHasil, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboKriteria.ResumeLayout(False)
        Me.gboKriteria.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslQuery As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents dgvHasil As System.Windows.Forms.DataGridView
    Friend WithEvents gboKriteria As System.Windows.Forms.GroupBox
    Friend WithEvents txtprimary As System.Windows.Forms.TextBox
    Public WithEvents cboCriteria As System.Windows.Forms.ComboBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Public WithEvents txtCari As System.Windows.Forms.TextBox
End Class
