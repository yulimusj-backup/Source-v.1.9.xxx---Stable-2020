<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterSalaryFormula_OtherDeduction_Cari
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
        Me.tslQuery = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboKriteria = New System.Windows.Forms.GroupBox
        Me.cboCriteria = New System.Windows.Forms.ComboBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.txtCari = New System.Windows.Forms.TextBox
        Me.dgvHasil = New System.Windows.Forms.DataGridView
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.DseTrans_OtherDeduction = New Core_Module.DseTrans_OtherDeduction
        Me.DataTable1BindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DataTable1TableAdapter = New Core_Module.DseTrans_OtherDeductionTableAdapters.DataTable1TableAdapter
        Me.PrimaryDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DescriptionDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CodeSlipDescDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.AccountCodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.gboKriteria.SuspendLayout()
        CType(Me.dgvHasil, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.DseTrans_OtherDeduction, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataTable1BindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        Me.gboKriteria.Controls.Add(Me.cboCriteria)
        Me.gboKriteria.Controls.Add(Me.btnOK)
        Me.gboKriteria.Controls.Add(Me.btnClose)
        Me.gboKriteria.Controls.Add(Me.btnFind)
        Me.gboKriteria.Controls.Add(Me.txtCari)
        Me.gboKriteria.Location = New System.Drawing.Point(12, 0)
        Me.gboKriteria.Name = "gboKriteria"
        Me.gboKriteria.Size = New System.Drawing.Size(656, 65)
        Me.gboKriteria.TabIndex = 47
        Me.gboKriteria.TabStop = False
        '
        'cboCriteria
        '
        Me.cboCriteria.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCriteria.FormattingEnabled = True
        Me.cboCriteria.Items.AddRange(New Object() {"Code", "Description", "Slip Code", "Reffered account"})
        Me.cboCriteria.Location = New System.Drawing.Point(6, 12)
        Me.cboCriteria.Name = "cboCriteria"
        Me.cboCriteria.Size = New System.Drawing.Size(121, 21)
        Me.cboCriteria.TabIndex = 5
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnOK.Location = New System.Drawing.Point(524, 38)
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
        Me.btnClose.Location = New System.Drawing.Point(590, 38)
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
        Me.btnFind.Location = New System.Drawing.Point(590, 10)
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
        Me.txtCari.Size = New System.Drawing.Size(451, 20)
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
        Me.dgvHasil.AutoGenerateColumns = False
        Me.dgvHasil.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvHasil.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.PrimaryDataGridViewTextBoxColumn, Me.CodeDataGridViewTextBoxColumn, Me.DescriptionDataGridViewTextBoxColumn, Me.CodeSlipDescDataGridViewTextBoxColumn, Me.AccountCodeDataGridViewTextBoxColumn})
        Me.dgvHasil.DataSource = Me.DataTable1BindingSource
        Me.dgvHasil.Location = New System.Drawing.Point(12, 72)
        Me.dgvHasil.MultiSelect = False
        Me.dgvHasil.Name = "dgvHasil"
        Me.dgvHasil.ReadOnly = True
        Me.dgvHasil.RowTemplate.Height = 24
        Me.dgvHasil.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvHasil.Size = New System.Drawing.Size(656, 308)
        Me.dgvHasil.TabIndex = 45
        Me.dgvHasil.TabStop = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslQuery})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 387)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(688, 22)
        Me.StatusStrip1.TabIndex = 46
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'DseTrans_OtherDeduction
        '
        Me.DseTrans_OtherDeduction.DataSetName = "DseTrans_OtherDeduction"
        Me.DseTrans_OtherDeduction.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'DataTable1BindingSource
        '
        Me.DataTable1BindingSource.DataMember = "DataTable1"
        Me.DataTable1BindingSource.DataSource = Me.DseTrans_OtherDeduction
        '
        'DataTable1TableAdapter
        '
        Me.DataTable1TableAdapter.ClearBeforeFill = True
        '
        'PrimaryDataGridViewTextBoxColumn
        '
        Me.PrimaryDataGridViewTextBoxColumn.DataPropertyName = "Primary"
        Me.PrimaryDataGridViewTextBoxColumn.HeaderText = "Primary"
        Me.PrimaryDataGridViewTextBoxColumn.Name = "PrimaryDataGridViewTextBoxColumn"
        Me.PrimaryDataGridViewTextBoxColumn.ReadOnly = True
        Me.PrimaryDataGridViewTextBoxColumn.Visible = False
        Me.PrimaryDataGridViewTextBoxColumn.Width = 66
        '
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "Code"
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "Code"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        Me.CodeDataGridViewTextBoxColumn.ReadOnly = True
        Me.CodeDataGridViewTextBoxColumn.Width = 57
        '
        'DescriptionDataGridViewTextBoxColumn
        '
        Me.DescriptionDataGridViewTextBoxColumn.DataPropertyName = "Description"
        Me.DescriptionDataGridViewTextBoxColumn.HeaderText = "Description"
        Me.DescriptionDataGridViewTextBoxColumn.Name = "DescriptionDataGridViewTextBoxColumn"
        Me.DescriptionDataGridViewTextBoxColumn.ReadOnly = True
        Me.DescriptionDataGridViewTextBoxColumn.Width = 85
        '
        'CodeSlipDescDataGridViewTextBoxColumn
        '
        Me.CodeSlipDescDataGridViewTextBoxColumn.DataPropertyName = "Code Slip desc"
        Me.CodeSlipDescDataGridViewTextBoxColumn.HeaderText = "Code Slip desc"
        Me.CodeSlipDescDataGridViewTextBoxColumn.Name = "CodeSlipDescDataGridViewTextBoxColumn"
        Me.CodeSlipDescDataGridViewTextBoxColumn.ReadOnly = True
        Me.CodeSlipDescDataGridViewTextBoxColumn.Visible = False
        Me.CodeSlipDescDataGridViewTextBoxColumn.Width = 103
        '
        'AccountCodeDataGridViewTextBoxColumn
        '
        Me.AccountCodeDataGridViewTextBoxColumn.DataPropertyName = "Account Code"
        Me.AccountCodeDataGridViewTextBoxColumn.HeaderText = "Account Code"
        Me.AccountCodeDataGridViewTextBoxColumn.Name = "AccountCodeDataGridViewTextBoxColumn"
        Me.AccountCodeDataGridViewTextBoxColumn.ReadOnly = True
        Me.AccountCodeDataGridViewTextBoxColumn.Visible = False
        '
        'FrmMasterSalaryFormula_OtherDeduction_Cari
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(688, 409)
        Me.Controls.Add(Me.gboKriteria)
        Me.Controls.Add(Me.dgvHasil)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(698, 446)
        Me.Name = "FrmMasterSalaryFormula_OtherDeduction_Cari"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Find Other Deduction"
        Me.gboKriteria.ResumeLayout(False)
        Me.gboKriteria.PerformLayout()
        CType(Me.dgvHasil, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.DseTrans_OtherDeduction, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataTable1BindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tslQuery As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents gboKriteria As System.Windows.Forms.GroupBox
    Public WithEvents cboCriteria As System.Windows.Forms.ComboBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Public WithEvents txtCari As System.Windows.Forms.TextBox
    Friend WithEvents dgvHasil As System.Windows.Forms.DataGridView
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents DseTrans_OtherDeduction As Core_Module.DseTrans_OtherDeduction
    Friend WithEvents DataTable1BindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents DataTable1TableAdapter As Core_Module.DseTrans_OtherDeductionTableAdapters.DataTable1TableAdapter
    Friend WithEvents PrimaryDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DescriptionDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CodeSlipDescDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AccountCodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
