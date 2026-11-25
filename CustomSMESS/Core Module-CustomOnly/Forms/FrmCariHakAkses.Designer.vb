<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCariHakAkses
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
        Me.cmdOK = New System.Windows.Forms.Button
        Me.cmdKeluar = New System.Windows.Forms.Button
        Me.DGHasil = New System.Windows.Forms.DataGridView
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.cmdSearch = New System.Windows.Forms.Button
        Me.LblTxtCari = New System.Windows.Forms.Label
        Me.TxtCari = New System.Windows.Forms.TextBox
        Me.RBDenganNama = New System.Windows.Forms.RadioButton
        Me.RBDenganID = New System.Windows.Forms.RadioButton
        CType(Me.DGHasil, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOK.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdOK.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOK.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmdOK.Location = New System.Drawing.Point(312, 363)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(108, 30)
        Me.cmdOK.TabIndex = 2
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'cmdKeluar
        '
        Me.cmdKeluar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdKeluar.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdKeluar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdKeluar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdKeluar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.cmdKeluar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdKeluar.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmdKeluar.Location = New System.Drawing.Point(426, 363)
        Me.cmdKeluar.Name = "cmdKeluar"
        Me.cmdKeluar.Size = New System.Drawing.Size(108, 30)
        Me.cmdKeluar.TabIndex = 3
        Me.cmdKeluar.Text = "&Keluar"
        Me.cmdKeluar.UseVisualStyleBackColor = False
        '
        'DGHasil
        '
        Me.DGHasil.AllowUserToAddRows = False
        Me.DGHasil.AllowUserToDeleteRows = False
        Me.DGHasil.AllowUserToOrderColumns = True
        Me.DGHasil.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DGHasil.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGHasil.Location = New System.Drawing.Point(6, 113)
        Me.DGHasil.MultiSelect = False
        Me.DGHasil.Name = "DGHasil"
        Me.DGHasil.ReadOnly = True
        Me.DGHasil.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DGHasil.Size = New System.Drawing.Size(527, 239)
        Me.DGHasil.TabIndex = 1
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox3.Controls.Add(Me.cmdSearch)
        Me.GroupBox3.Controls.Add(Me.LblTxtCari)
        Me.GroupBox3.Controls.Add(Me.TxtCari)
        Me.GroupBox3.Controls.Add(Me.RBDenganNama)
        Me.GroupBox3.Controls.Add(Me.RBDenganID)
        Me.GroupBox3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.GroupBox3.Location = New System.Drawing.Point(6, 12)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(522, 95)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        '
        'cmdSearch
        '
        Me.cmdSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSearch.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdSearch.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdSearch.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.cmdSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSearch.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmdSearch.Location = New System.Drawing.Point(403, 56)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.Size = New System.Drawing.Size(88, 30)
        Me.cmdSearch.TabIndex = 1
        Me.cmdSearch.Text = "Ca&ri"
        Me.cmdSearch.UseVisualStyleBackColor = False
        '
        'LblTxtCari
        '
        Me.LblTxtCari.AutoSize = True
        Me.LblTxtCari.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblTxtCari.ForeColor = System.Drawing.SystemColors.WindowText
        Me.LblTxtCari.Location = New System.Drawing.Point(22, 62)
        Me.LblTxtCari.Name = "LblTxtCari"
        Me.LblTxtCari.Size = New System.Drawing.Size(87, 13)
        Me.LblTxtCari.TabIndex = 18
        Me.LblTxtCari.Text = "Kode Hak Akses"
        '
        'TxtCari
        '
        Me.TxtCari.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtCari.Location = New System.Drawing.Point(179, 60)
        Me.TxtCari.MaxLength = 50
        Me.TxtCari.Name = "TxtCari"
        Me.TxtCari.Size = New System.Drawing.Size(181, 22)
        Me.TxtCari.TabIndex = 0
        '
        'RBDenganNama
        '
        Me.RBDenganNama.AutoSize = True
        Me.RBDenganNama.Checked = True
        Me.RBDenganNama.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBDenganNama.ForeColor = System.Drawing.SystemColors.WindowText
        Me.RBDenganNama.Location = New System.Drawing.Point(26, 25)
        Me.RBDenganNama.Name = "RBDenganNama"
        Me.RBDenganNama.Size = New System.Drawing.Size(94, 17)
        Me.RBDenganNama.TabIndex = 1
        Me.RBDenganNama.TabStop = True
        Me.RBDenganNama.Text = "Dengan Nama"
        Me.RBDenganNama.UseVisualStyleBackColor = True
        '
        'RBDenganID
        '
        Me.RBDenganID.AutoSize = True
        Me.RBDenganID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBDenganID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.RBDenganID.Location = New System.Drawing.Point(179, 25)
        Me.RBDenganID.Name = "RBDenganID"
        Me.RBDenganID.Size = New System.Drawing.Size(91, 17)
        Me.RBDenganID.TabIndex = 0
        Me.RBDenganID.Text = "Dengan Kode"
        Me.RBDenganID.UseVisualStyleBackColor = True
        '
        'FrmCariHakAkses
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdKeluar
        Me.ClientSize = New System.Drawing.Size(540, 407)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdKeluar)
        Me.Controls.Add(Me.DGHasil)
        Me.Controls.Add(Me.GroupBox3)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(556, 443)
        Me.Name = "FrmCariHakAkses"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " Cari Hak Akses"
        CType(Me.DGHasil, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdKeluar As System.Windows.Forms.Button
    Friend WithEvents DGHasil As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdSearch As System.Windows.Forms.Button
    Friend WithEvents LblTxtCari As System.Windows.Forms.Label
    Friend WithEvents TxtCari As System.Windows.Forms.TextBox
    Friend WithEvents RBDenganNama As System.Windows.Forms.RadioButton
    Friend WithEvents RBDenganID As System.Windows.Forms.RadioButton
End Class
