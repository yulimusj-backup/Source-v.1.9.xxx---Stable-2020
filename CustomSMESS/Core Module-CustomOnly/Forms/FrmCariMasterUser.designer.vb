<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCariMasterUser
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
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.txtPrimaryGroup = New System.Windows.Forms.TextBox
        Me.cmdSearch = New System.Windows.Forms.Button
        Me.LblTxtCari = New System.Windows.Forms.Label
        Me.TxtCari = New System.Windows.Forms.TextBox
        Me.RBDenganNama = New System.Windows.Forms.RadioButton
        Me.RBDenganID = New System.Windows.Forms.RadioButton
        Me.DGHasil = New System.Windows.Forms.DataGridView
        Me.GroupBox3.SuspendLayout()
        CType(Me.DGHasil, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdOK.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Lime
        Me.cmdOK.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.cmdOK.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.Color.Black
        Me.cmdOK.Location = New System.Drawing.Point(317, 364)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(108, 30)
        Me.cmdOK.TabIndex = 38
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'cmdKeluar
        '
        Me.cmdKeluar.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdKeluar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdKeluar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Red
        Me.cmdKeluar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.cmdKeluar.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdKeluar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdKeluar.ForeColor = System.Drawing.Color.Black
        Me.cmdKeluar.Location = New System.Drawing.Point(431, 364)
        Me.cmdKeluar.Name = "cmdKeluar"
        Me.cmdKeluar.Size = New System.Drawing.Size(108, 30)
        Me.cmdKeluar.TabIndex = 37
        Me.cmdKeluar.Text = "&Keluar"
        Me.cmdKeluar.UseVisualStyleBackColor = False
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox3.Controls.Add(Me.txtPrimaryGroup)
        Me.GroupBox3.Controls.Add(Me.cmdSearch)
        Me.GroupBox3.Controls.Add(Me.LblTxtCari)
        Me.GroupBox3.Controls.Add(Me.TxtCari)
        Me.GroupBox3.Controls.Add(Me.RBDenganNama)
        Me.GroupBox3.Controls.Add(Me.RBDenganID)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.GroupBox3.Location = New System.Drawing.Point(12, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(528, 108)
        Me.GroupBox3.TabIndex = 36
        Me.GroupBox3.TabStop = False
        '
        'txtPrimaryGroup
        '
        Me.txtPrimaryGroup.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryGroup.Location = New System.Drawing.Point(362, 30)
        Me.txtPrimaryGroup.Name = "txtPrimaryGroup"
        Me.txtPrimaryGroup.Size = New System.Drawing.Size(14, 21)
        Me.txtPrimaryGroup.TabIndex = 39
        Me.txtPrimaryGroup.TabStop = False
        Me.txtPrimaryGroup.Visible = False
        '
        'cmdSearch
        '
        Me.cmdSearch.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdSearch.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Yellow
        Me.cmdSearch.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.cmdSearch.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSearch.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmdSearch.Location = New System.Drawing.Point(419, 56)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.Size = New System.Drawing.Size(97, 27)
        Me.cmdSearch.TabIndex = 23
        Me.cmdSearch.Text = "Ca&ri"
        Me.cmdSearch.UseVisualStyleBackColor = False
        '
        'LblTxtCari
        '
        Me.LblTxtCari.AutoSize = True
        Me.LblTxtCari.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblTxtCari.Location = New System.Drawing.Point(25, 61)
        Me.LblTxtCari.Name = "LblTxtCari"
        Me.LblTxtCari.Size = New System.Drawing.Size(67, 13)
        Me.LblTxtCari.TabIndex = 18
        Me.LblTxtCari.Text = "Kode Kriteria"
        '
        'TxtCari
        '
        Me.TxtCari.Location = New System.Drawing.Point(182, 62)
        Me.TxtCari.Name = "TxtCari"
        Me.TxtCari.Size = New System.Drawing.Size(179, 21)
        Me.TxtCari.TabIndex = 17
        '
        'RBDenganNama
        '
        Me.RBDenganNama.AutoSize = True
        Me.RBDenganNama.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBDenganNama.Location = New System.Drawing.Point(182, 29)
        Me.RBDenganNama.Name = "RBDenganNama"
        Me.RBDenganNama.Size = New System.Drawing.Size(94, 17)
        Me.RBDenganNama.TabIndex = 1
        Me.RBDenganNama.Text = "Dengan Nama"
        Me.RBDenganNama.UseVisualStyleBackColor = True
        '
        'RBDenganID
        '
        Me.RBDenganID.AutoSize = True
        Me.RBDenganID.Checked = True
        Me.RBDenganID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBDenganID.Location = New System.Drawing.Point(24, 29)
        Me.RBDenganID.Name = "RBDenganID"
        Me.RBDenganID.Size = New System.Drawing.Size(91, 17)
        Me.RBDenganID.TabIndex = 0
        Me.RBDenganID.TabStop = True
        Me.RBDenganID.Text = "Dengan Kode"
        Me.RBDenganID.UseVisualStyleBackColor = True
        '
        'DGHasil
        '
        Me.DGHasil.AllowUserToAddRows = False
        Me.DGHasil.AllowUserToDeleteRows = False
        Me.DGHasil.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGHasil.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DGHasil.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGHasil.Location = New System.Drawing.Point(13, 114)
        Me.DGHasil.MultiSelect = False
        Me.DGHasil.Name = "DGHasil"
        Me.DGHasil.ReadOnly = True
        Me.DGHasil.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DGHasil.Size = New System.Drawing.Size(527, 239)
        Me.DGHasil.TabIndex = 35
        '
        'FrmCariMasterUser
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdKeluar
        Me.ClientSize = New System.Drawing.Size(552, 407)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdKeluar)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.DGHasil)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmCariMasterUser"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cari Data User"
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.DGHasil, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdKeluar As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdSearch As System.Windows.Forms.Button
    Friend WithEvents LblTxtCari As System.Windows.Forms.Label
    Friend WithEvents RBDenganNama As System.Windows.Forms.RadioButton
    Friend WithEvents RBDenganID As System.Windows.Forms.RadioButton
    Friend WithEvents DGHasil As System.Windows.Forms.DataGridView
    Friend WithEvents txtPrimaryGroup As System.Windows.Forms.TextBox
    Public WithEvents TxtCari As System.Windows.Forms.TextBox
End Class
