<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterWorkUnit_Cari
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
        Me.gboKriteria = New System.Windows.Forms.GroupBox
        Me.txtPrimaryGroup = New System.Windows.Forms.TextBox
        Me.cboCriteria = New System.Windows.Forms.ComboBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.txtCari = New System.Windows.Forms.TextBox
        Me.dgvHasil = New System.Windows.Forms.DataGridView
        Me.StatusStrip1.SuspendLayout()
        Me.gboKriteria.SuspendLayout()
        CType(Me.dgvHasil, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslQuery})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 393)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(629, 22)
        Me.StatusStrip1.TabIndex = 18
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslQuery
        '
        Me.tslQuery.Name = "tslQuery"
        Me.tslQuery.Size = New System.Drawing.Size(0, 17)
        '
        'gboKriteria
        '
        Me.gboKriteria.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboKriteria.Controls.Add(Me.txtPrimaryGroup)
        Me.gboKriteria.Controls.Add(Me.cboCriteria)
        Me.gboKriteria.Controls.Add(Me.btnOK)
        Me.gboKriteria.Controls.Add(Me.btnClose)
        Me.gboKriteria.Controls.Add(Me.btnFind)
        Me.gboKriteria.Controls.Add(Me.txtCari)
        Me.gboKriteria.Location = New System.Drawing.Point(16, 1)
        Me.gboKriteria.Margin = New System.Windows.Forms.Padding(4)
        Me.gboKriteria.Name = "gboKriteria"
        Me.gboKriteria.Padding = New System.Windows.Forms.Padding(4)
        Me.gboKriteria.Size = New System.Drawing.Size(587, 80)
        Me.gboKriteria.TabIndex = 17
        Me.gboKriteria.TabStop = False
        '
        'txtPrimaryGroup
        '
        Me.txtPrimaryGroup.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryGroup.Location = New System.Drawing.Point(273, 47)
        Me.txtPrimaryGroup.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrimaryGroup.Name = "txtPrimaryGroup"
        Me.txtPrimaryGroup.Size = New System.Drawing.Size(20, 22)
        Me.txtPrimaryGroup.TabIndex = 25
        Me.txtPrimaryGroup.TabStop = False
        Me.txtPrimaryGroup.Visible = False
        '
        'cboCriteria
        '
        Me.cboCriteria.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCriteria.FormattingEnabled = True
        Me.cboCriteria.Items.AddRange(New Object() {"Code", "Description", "Remarks", "Employee", "Workshift", "Machine", "Location In", "Location Out"})
        Me.cboCriteria.Location = New System.Drawing.Point(8, 15)
        Me.cboCriteria.Margin = New System.Windows.Forms.Padding(4)
        Me.cboCriteria.Name = "cboCriteria"
        Me.cboCriteria.Size = New System.Drawing.Size(160, 24)
        Me.cboCriteria.TabIndex = 5
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnOK.Location = New System.Drawing.Point(411, 47)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 28)
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
        Me.btnClose.Location = New System.Drawing.Point(499, 47)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(80, 28)
        Me.btnClose.TabIndex = 2
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFind.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnFind.Location = New System.Drawing.Point(499, 12)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 28)
        Me.btnFind.TabIndex = 1
        Me.btnFind.TabStop = False
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'txtCari
        '
        Me.txtCari.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCari.Location = New System.Drawing.Point(177, 15)
        Me.txtCari.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCari.Name = "txtCari"
        Me.txtCari.Size = New System.Drawing.Size(312, 22)
        Me.txtCari.TabIndex = 0
        '
        'dgvHasil
        '
        Me.dgvHasil.AllowUserToAddRows = False
        Me.dgvHasil.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvHasil.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvHasil.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvHasil.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvHasil.Location = New System.Drawing.Point(16, 84)
        Me.dgvHasil.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvHasil.MultiSelect = False
        Me.dgvHasil.Name = "dgvHasil"
        Me.dgvHasil.ReadOnly = True
        Me.dgvHasil.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvHasil.Size = New System.Drawing.Size(587, 292)
        Me.dgvHasil.TabIndex = 19
        Me.dgvHasil.TabStop = False
        '
        'FrmMasterWorkUnit_Cari
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(629, 415)
        Me.Controls.Add(Me.dgvHasil)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.gboKriteria)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(637, 452)
        Me.Name = "FrmMasterWorkUnit_Cari"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Find Work Unit Data"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboKriteria.ResumeLayout(False)
        Me.gboKriteria.PerformLayout()
        CType(Me.dgvHasil, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslQuery As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents gboKriteria As System.Windows.Forms.GroupBox
    Friend WithEvents txtPrimaryGroup As System.Windows.Forms.TextBox
    Public WithEvents cboCriteria As System.Windows.Forms.ComboBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Public WithEvents txtCari As System.Windows.Forms.TextBox
    Friend WithEvents dgvHasil As System.Windows.Forms.DataGridView
End Class
