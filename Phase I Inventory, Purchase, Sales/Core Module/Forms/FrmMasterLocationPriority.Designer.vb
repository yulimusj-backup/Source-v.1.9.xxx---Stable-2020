<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterLocationPriority
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMasterLocationPriority))
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColLocationCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocationDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouseCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouseDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.chkIncludeOtherLocation = New System.Windows.Forms.CheckBox
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(199, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Please specify location priority"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColLocationCode, Me.ColLocationDesc, Me.ColWarehouseCode, Me.ColWarehouseDesc, Me.ColPrimaryLocation})
        Me.dgvItem.Location = New System.Drawing.Point(12, 30)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(661, 285)
        Me.dgvItem.TabIndex = 1
        '
        'ColLocationCode
        '
        Me.ColLocationCode.HeaderText = "Location Code"
        Me.ColLocationCode.Name = "ColLocationCode"
        Me.ColLocationCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLocationCode.Width = 105
        '
        'ColLocationDesc
        '
        Me.ColLocationDesc.HeaderText = "Description"
        Me.ColLocationDesc.Name = "ColLocationDesc"
        Me.ColLocationDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLocationDesc.Width = 85
        '
        'ColWarehouseCode
        '
        Me.ColWarehouseCode.HeaderText = "Warehouse Code"
        Me.ColWarehouseCode.Name = "ColWarehouseCode"
        Me.ColWarehouseCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWarehouseCode.Width = 124
        '
        'ColWarehouseDesc
        '
        Me.ColWarehouseDesc.HeaderText = "Description"
        Me.ColWarehouseDesc.Name = "ColWarehouseDesc"
        Me.ColWarehouseDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWarehouseDesc.Width = 85
        '
        'ColPrimaryLocation
        '
        Me.ColPrimaryLocation.HeaderText = "Primary Location"
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.Visible = False
        Me.ColPrimaryLocation.Width = 139
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(596, 329)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(77, 25)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(512, 329)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(77, 25)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = CType(resources.GetObject("btnUP.BackgroundImage"), System.Drawing.Image)
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(16, 329)
        Me.btnUP.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(37, 32)
        Me.btnUP.TabIndex = 11
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = CType(resources.GetObject("btnDown.BackgroundImage"), System.Drawing.Image)
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(61, 329)
        Me.btnDown.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(37, 32)
        Me.btnDown.TabIndex = 12
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'chkIncludeOtherLocation
        '
        Me.chkIncludeOtherLocation.AutoSize = True
        Me.chkIncludeOtherLocation.Location = New System.Drawing.Point(117, 329)
        Me.chkIncludeOtherLocation.Name = "chkIncludeOtherLocation"
        Me.chkIncludeOtherLocation.Size = New System.Drawing.Size(165, 21)
        Me.chkIncludeOtherLocation.TabIndex = 13
        Me.chkIncludeOtherLocation.Text = "Include other location"
        Me.chkIncludeOtherLocation.UseVisualStyleBackColor = True
        Me.chkIncludeOtherLocation.Visible = False
        '
        'FrmMasterLocationPriority
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(688, 386)
        Me.Controls.Add(Me.chkIncludeOtherLocation)
        Me.Controls.Add(Me.btnUP)
        Me.Controls.Add(Me.btnDown)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(696, 395)
        Me.Name = "FrmMasterLocationPriority"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Location Priority"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents ColLocationCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocationDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouseCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouseDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Public WithEvents chkIncludeOtherLocation As System.Windows.Forms.CheckBox
End Class
