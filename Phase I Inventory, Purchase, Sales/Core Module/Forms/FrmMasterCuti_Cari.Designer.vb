<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterCuti_Cari
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
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.txtprimarygroup = New System.Windows.Forms.TextBox
        Me.cboCriteria = New System.Windows.Forms.ComboBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.tslQuery = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.gboKriteria = New System.Windows.Forms.GroupBox
        Me.txtCari = New System.Windows.Forms.TextBox
        Me.dgvHasil = New System.Windows.Forms.DataGridView
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.TblmastercutiBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DseMasterCuti = New Core_Module.DseMasterCuti
        Me.Tblmaster_cutiTableAdapter = New Core_Module.DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Me.ColPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPeriod = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.gboKriteria.SuspendLayout()
        CType(Me.dgvHasil, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.TblmastercutiBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DseMasterCuti, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtprimarygroup
        '
        Me.txtprimarygroup.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtprimarygroup.Location = New System.Drawing.Point(481, 38)
        Me.txtprimarygroup.Name = "txtprimarygroup"
        Me.txtprimarygroup.Size = New System.Drawing.Size(35, 20)
        Me.txtprimarygroup.TabIndex = 6
        Me.txtprimarygroup.TabStop = False
        Me.txtprimarygroup.Text = "-1"
        Me.txtprimarygroup.Visible = False
        '
        'cboCriteria
        '
        Me.cboCriteria.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCriteria.FormattingEnabled = True
        Me.cboCriteria.Items.AddRange(New Object() {"Code", "Description", "Remarks", "Kode Karyawan", "Nama Karyawan"})
        Me.cboCriteria.Location = New System.Drawing.Point(6, 12)
        Me.cboCriteria.Name = "cboCriteria"
        Me.cboCriteria.Size = New System.Drawing.Size(157, 21)
        Me.cboCriteria.TabIndex = 5
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnOK.Location = New System.Drawing.Point(522, 38)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(60, 23)
        Me.btnOK.TabIndex = 4
        Me.btnOK.TabStop = False
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'tslQuery
        '
        Me.tslQuery.Name = "tslQuery"
        Me.tslQuery.Size = New System.Drawing.Size(0, 17)
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnClose.Location = New System.Drawing.Point(588, 38)
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
        Me.btnFind.Location = New System.Drawing.Point(588, 10)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(60, 23)
        Me.btnFind.TabIndex = 1
        Me.btnFind.TabStop = False
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'gboKriteria
        '
        Me.gboKriteria.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboKriteria.Controls.Add(Me.txtprimarygroup)
        Me.gboKriteria.Controls.Add(Me.cboCriteria)
        Me.gboKriteria.Controls.Add(Me.btnOK)
        Me.gboKriteria.Controls.Add(Me.btnClose)
        Me.gboKriteria.Controls.Add(Me.btnFind)
        Me.gboKriteria.Controls.Add(Me.txtCari)
        Me.gboKriteria.Location = New System.Drawing.Point(12, 1)
        Me.gboKriteria.Name = "gboKriteria"
        Me.gboKriteria.Size = New System.Drawing.Size(654, 65)
        Me.gboKriteria.TabIndex = 33
        Me.gboKriteria.TabStop = False
        '
        'txtCari
        '
        Me.txtCari.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCari.Location = New System.Drawing.Point(169, 12)
        Me.txtCari.Name = "txtCari"
        Me.txtCari.Size = New System.Drawing.Size(413, 20)
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
        Me.dgvHasil.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColPrimary, Me.ColCode, Me.ColDescription, Me.ColPeriod, Me.ColRemarks})
        Me.dgvHasil.Location = New System.Drawing.Point(12, 72)
        Me.dgvHasil.MultiSelect = False
        Me.dgvHasil.Name = "dgvHasil"
        Me.dgvHasil.ReadOnly = True
        Me.dgvHasil.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvHasil.Size = New System.Drawing.Size(654, 309)
        Me.dgvHasil.TabIndex = 32
        Me.dgvHasil.TabStop = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslQuery})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 386)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(678, 22)
        Me.StatusStrip1.TabIndex = 34
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'TblmastercutiBindingSource
        '
        Me.TblmastercutiBindingSource.DataMember = "tblmaster_cuti"
        Me.TblmastercutiBindingSource.DataSource = Me.DseMasterCuti
        '
        'DseMasterCuti
        '
        Me.DseMasterCuti.DataSetName = "DseMasterCuti"
        Me.DseMasterCuti.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'Tblmaster_cutiTableAdapter
        '
        Me.Tblmaster_cutiTableAdapter.ClearBeforeFill = True
        '
        'ColPrimary
        '
        Me.ColPrimary.HeaderText = "Primary"
        Me.ColPrimary.Name = "ColPrimary"
        Me.ColPrimary.ReadOnly = True
        Me.ColPrimary.Visible = False
        Me.ColPrimary.Width = 66
        '
        'ColCode
        '
        Me.ColCode.HeaderText = "Code"
        Me.ColCode.Name = "ColCode"
        Me.ColCode.ReadOnly = True
        Me.ColCode.Width = 57
        '
        'ColDescription
        '
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.ReadOnly = True
        Me.ColDescription.Width = 85
        '
        'ColPeriod
        '
        Me.ColPeriod.HeaderText = "Period"
        Me.ColPeriod.Name = "ColPeriod"
        Me.ColPeriod.ReadOnly = True
        Me.ColPeriod.Width = 62
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.ReadOnly = True
        Me.ColRemarks.Width = 74
        '
        'FrmMasterCuti_Cari
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(678, 408)
        Me.Controls.Add(Me.gboKriteria)
        Me.Controls.Add(Me.dgvHasil)
        Me.Controls.Add(Me.StatusStrip1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(694, 446)
        Me.Name = "FrmMasterCuti_Cari"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Find Leave Data"
        Me.gboKriteria.ResumeLayout(False)
        Me.gboKriteria.PerformLayout()
        CType(Me.dgvHasil, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.TblmastercutiBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DseMasterCuti, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtprimarygroup As System.Windows.Forms.TextBox
    Public WithEvents cboCriteria As System.Windows.Forms.ComboBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents tslQuery As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents gboKriteria As System.Windows.Forms.GroupBox
    Public WithEvents txtCari As System.Windows.Forms.TextBox
    Friend WithEvents dgvHasil As System.Windows.Forms.DataGridView
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents DseMasterCuti As Core_Module.DseMasterCuti
    Friend WithEvents TblmastercutiBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Tblmaster_cutiTableAdapter As Core_Module.DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
    Friend WithEvents ColPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPeriod As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
