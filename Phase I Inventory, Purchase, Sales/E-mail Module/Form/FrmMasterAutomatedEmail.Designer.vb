<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterAutomatedEmail
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMasterAutomatedEmail))
        Me.DsTransaksiEmail = New E_mail_Module.DsTransaksiEmail
        Me.Tbltrans_email_automatedBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Tbltrans_email_automatedTableAdapter = New E_mail_Module.DsTransaksiEmailTableAdapters.tbltrans_email_automatedTableAdapter
        Me.Tbltrans_email_automatedBindingNavigator = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorSeparator2 = New System.Windows.Forms.ToolStripSeparator
        Me.BindingNavigatorAddNewItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorDeleteItem = New System.Windows.Forms.ToolStripButton
        Me.Tbltrans_email_automatedBindingNavigatorSaveItem = New System.Windows.Forms.ToolStripButton
        Me.Tbltrans_email_automatedDataGridView = New System.Windows.Forms.DataGridView
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.DsTransaksiEmail, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Tbltrans_email_automatedBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Tbltrans_email_automatedBindingNavigator, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Tbltrans_email_automatedBindingNavigator.SuspendLayout()
        CType(Me.Tbltrans_email_automatedDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DsTransaksiEmail
        '
        Me.DsTransaksiEmail.DataSetName = "DsTransaksiEmail"
        Me.DsTransaksiEmail.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'Tbltrans_email_automatedBindingSource
        '
        Me.Tbltrans_email_automatedBindingSource.DataMember = "tbltrans_email_automated"
        Me.Tbltrans_email_automatedBindingSource.DataSource = Me.DsTransaksiEmail
        '
        'Tbltrans_email_automatedTableAdapter
        '
        Me.Tbltrans_email_automatedTableAdapter.ClearBeforeFill = True
        '
        'Tbltrans_email_automatedBindingNavigator
        '
        Me.Tbltrans_email_automatedBindingNavigator.AddNewItem = Me.BindingNavigatorAddNewItem
        Me.Tbltrans_email_automatedBindingNavigator.BindingSource = Me.Tbltrans_email_automatedBindingSource
        Me.Tbltrans_email_automatedBindingNavigator.CountItem = Me.BindingNavigatorCountItem
        Me.Tbltrans_email_automatedBindingNavigator.DeleteItem = Me.BindingNavigatorDeleteItem
        Me.Tbltrans_email_automatedBindingNavigator.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorSeparator2, Me.BindingNavigatorAddNewItem, Me.BindingNavigatorDeleteItem, Me.Tbltrans_email_automatedBindingNavigatorSaveItem})
        Me.Tbltrans_email_automatedBindingNavigator.Location = New System.Drawing.Point(0, 0)
        Me.Tbltrans_email_automatedBindingNavigator.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.Tbltrans_email_automatedBindingNavigator.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.Tbltrans_email_automatedBindingNavigator.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.Tbltrans_email_automatedBindingNavigator.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.Tbltrans_email_automatedBindingNavigator.Name = "Tbltrans_email_automatedBindingNavigator"
        Me.Tbltrans_email_automatedBindingNavigator.PositionItem = Me.BindingNavigatorPositionItem
        Me.Tbltrans_email_automatedBindingNavigator.Size = New System.Drawing.Size(667, 25)
        Me.Tbltrans_email_automatedBindingNavigator.TabIndex = 0
        Me.Tbltrans_email_automatedBindingNavigator.Text = "BindingNavigator1"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveFirstItem.Text = "Move first"
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMovePreviousItem.Text = "Move previous"
        '
        'BindingNavigatorSeparator
        '
        Me.BindingNavigatorSeparator.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorPositionItem
        '
        Me.BindingNavigatorPositionItem.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem.AutoSize = False
        Me.BindingNavigatorPositionItem.Name = "BindingNavigatorPositionItem"
        Me.BindingNavigatorPositionItem.Size = New System.Drawing.Size(50, 23)
        Me.BindingNavigatorPositionItem.Text = "0"
        Me.BindingNavigatorPositionItem.ToolTipText = "Current position"
        '
        'BindingNavigatorCountItem
        '
        Me.BindingNavigatorCountItem.Name = "BindingNavigatorCountItem"
        Me.BindingNavigatorCountItem.Size = New System.Drawing.Size(35, 15)
        Me.BindingNavigatorCountItem.Text = "of {0}"
        Me.BindingNavigatorCountItem.ToolTipText = "Total number of items"
        '
        'BindingNavigatorSeparator1
        '
        Me.BindingNavigatorSeparator1.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator1.Size = New System.Drawing.Size(6, 6)
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(23, 20)
        Me.BindingNavigatorMoveNextItem.Text = "Move next"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(23, 20)
        Me.BindingNavigatorMoveLastItem.Text = "Move last"
        '
        'BindingNavigatorSeparator2
        '
        Me.BindingNavigatorSeparator2.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator2.Size = New System.Drawing.Size(6, 6)
        '
        'BindingNavigatorAddNewItem
        '
        Me.BindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorAddNewItem.Image = CType(resources.GetObject("BindingNavigatorAddNewItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorAddNewItem.Name = "BindingNavigatorAddNewItem"
        Me.BindingNavigatorAddNewItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorAddNewItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorAddNewItem.Text = "Add new"
        '
        'BindingNavigatorDeleteItem
        '
        Me.BindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorDeleteItem.Image = CType(resources.GetObject("BindingNavigatorDeleteItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorDeleteItem.Name = "BindingNavigatorDeleteItem"
        Me.BindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorDeleteItem.Size = New System.Drawing.Size(23, 20)
        Me.BindingNavigatorDeleteItem.Text = "Delete"
        '
        'Tbltrans_email_automatedBindingNavigatorSaveItem
        '
        Me.Tbltrans_email_automatedBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.Tbltrans_email_automatedBindingNavigatorSaveItem.Image = CType(resources.GetObject("Tbltrans_email_automatedBindingNavigatorSaveItem.Image"), System.Drawing.Image)
        Me.Tbltrans_email_automatedBindingNavigatorSaveItem.Name = "Tbltrans_email_automatedBindingNavigatorSaveItem"
        Me.Tbltrans_email_automatedBindingNavigatorSaveItem.Size = New System.Drawing.Size(23, 23)
        Me.Tbltrans_email_automatedBindingNavigatorSaveItem.Text = "Save Data"
        '
        'Tbltrans_email_automatedDataGridView
        '
        Me.Tbltrans_email_automatedDataGridView.AutoGenerateColumns = False
        Me.Tbltrans_email_automatedDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.Tbltrans_email_automatedDataGridView.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.Tbltrans_email_automatedDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Tbltrans_email_automatedDataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5})
        Me.Tbltrans_email_automatedDataGridView.DataSource = Me.Tbltrans_email_automatedBindingSource
        Me.Tbltrans_email_automatedDataGridView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Tbltrans_email_automatedDataGridView.Location = New System.Drawing.Point(0, 25)
        Me.Tbltrans_email_automatedDataGridView.Name = "Tbltrans_email_automatedDataGridView"
        Me.Tbltrans_email_automatedDataGridView.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.Tbltrans_email_automatedDataGridView.Size = New System.Drawing.Size(667, 380)
        Me.Tbltrans_email_automatedDataGridView.TabIndex = 1
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.DataPropertyName = "Primary"
        Me.DataGridViewTextBoxColumn1.HeaderText = "Primary"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.Width = 66
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.DataPropertyName = "Subject"
        Me.DataGridViewTextBoxColumn2.HeaderText = "Subject"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.Width = 68
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.DataPropertyName = "Description"
        Me.DataGridViewTextBoxColumn3.HeaderText = "Description"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Width = 85
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.DataPropertyName = "Subject English"
        Me.DataGridViewTextBoxColumn4.HeaderText = "Subject English"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.Width = 105
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.DataPropertyName = "Description English"
        Me.DataGridViewTextBoxColumn5.HeaderText = "Description English"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.Width = 122
        '
        'FrmMasterAutomatedEmail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(667, 405)
        Me.Controls.Add(Me.Tbltrans_email_automatedDataGridView)
        Me.Controls.Add(Me.Tbltrans_email_automatedBindingNavigator)
        Me.Name = "FrmMasterAutomatedEmail"
        Me.Text = "FrmMasterAutomatedEmail-Programmer's Only"
        CType(Me.DsTransaksiEmail, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Tbltrans_email_automatedBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Tbltrans_email_automatedBindingNavigator, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Tbltrans_email_automatedBindingNavigator.ResumeLayout(False)
        Me.Tbltrans_email_automatedBindingNavigator.PerformLayout()
        CType(Me.Tbltrans_email_automatedDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DsTransaksiEmail As E_mail_Module.DsTransaksiEmail
    Friend WithEvents Tbltrans_email_automatedBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Tbltrans_email_automatedTableAdapter As E_mail_Module.DsTransaksiEmailTableAdapters.tbltrans_email_automatedTableAdapter
    Friend WithEvents Tbltrans_email_automatedBindingNavigator As System.Windows.Forms.BindingNavigator
    Friend WithEvents BindingNavigatorAddNewItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorCountItem As System.Windows.Forms.ToolStripLabel
    Friend WithEvents BindingNavigatorDeleteItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveFirstItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorPositionItem As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents Tbltrans_email_automatedBindingNavigatorSaveItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents Tbltrans_email_automatedDataGridView As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
