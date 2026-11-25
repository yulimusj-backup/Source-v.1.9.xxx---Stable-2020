<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ManufacturingOrder_SetFinishJO
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColItem_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_MONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_JONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_QtyTarget = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_QtyProduce = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_IsClosed = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColItem_IdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.chkOngoingJOOnly = New System.Windows.Forms.CheckBox
        Me.llbUnCheckAll = New System.Windows.Forms.LinkLabel
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.Button1 = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.ChkShow100PectOnly = New System.Windows.Forms.CheckBox
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(76, 9)
        Me.dtpFrom.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(120, 20)
        Me.dtpFrom.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 13)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(50, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "MO Date"
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(220, 9)
        Me.dtpTo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(120, 20)
        Me.dtpTo.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(200, 13)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(16, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "to"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItem_No, Me.ColItem_MONumber, Me.ColItem_JONumber, Me.ColItem_CodeNo, Me.ColItem_Desc, Me.ColItem_QtyTarget, Me.ColItem_QtyProduce, Me.ColItem_Unit, Me.ColItem_Remarks, Me.ColItem_IsClosed, Me.ColItem_IdentifierJO})
        Me.dgvItem.Location = New System.Drawing.Point(9, 51)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(720, 456)
        Me.dgvItem.TabIndex = 18
        '
        'ColItem_No
        '
        Me.ColItem_No.Frozen = True
        Me.ColItem_No.HeaderText = "No"
        Me.ColItem_No.Name = "ColItem_No"
        Me.ColItem_No.Width = 46
        '
        'ColItem_MONumber
        '
        Me.ColItem_MONumber.Frozen = True
        Me.ColItem_MONumber.HeaderText = "Manufacturing Order Number"
        Me.ColItem_MONumber.Name = "ColItem_MONumber"
        Me.ColItem_MONumber.Width = 169
        '
        'ColItem_JONumber
        '
        Me.ColItem_JONumber.HeaderText = "Job Order No"
        Me.ColItem_JONumber.Name = "ColItem_JONumber"
        Me.ColItem_JONumber.Width = 95
        '
        'ColItem_CodeNo
        '
        Me.ColItem_CodeNo.HeaderText = "Code No"
        Me.ColItem_CodeNo.Name = "ColItem_CodeNo"
        Me.ColItem_CodeNo.Width = 74
        '
        'ColItem_Desc
        '
        Me.ColItem_Desc.HeaderText = "Description"
        Me.ColItem_Desc.Name = "ColItem_Desc"
        Me.ColItem_Desc.Width = 85
        '
        'ColItem_QtyTarget
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_QtyTarget.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColItem_QtyTarget.HeaderText = "Qty Target"
        Me.ColItem_QtyTarget.Name = "ColItem_QtyTarget"
        Me.ColItem_QtyTarget.Width = 82
        '
        'ColItem_QtyProduce
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_QtyProduce.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColItem_QtyProduce.HeaderText = "Qty Produce"
        Me.ColItem_QtyProduce.Name = "ColItem_QtyProduce"
        Me.ColItem_QtyProduce.Width = 91
        '
        'ColItem_Unit
        '
        Me.ColItem_Unit.HeaderText = "Unit"
        Me.ColItem_Unit.Name = "ColItem_Unit"
        Me.ColItem_Unit.Width = 51
        '
        'ColItem_Remarks
        '
        Me.ColItem_Remarks.HeaderText = "Remarks"
        Me.ColItem_Remarks.Name = "ColItem_Remarks"
        Me.ColItem_Remarks.Width = 74
        '
        'ColItem_IsClosed
        '
        Me.ColItem_IsClosed.HeaderText = "Closed"
        Me.ColItem_IsClosed.Name = "ColItem_IsClosed"
        Me.ColItem_IsClosed.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColItem_IsClosed.Width = 45
        '
        'ColItem_IdentifierJO
        '
        Me.ColItem_IdentifierJO.HeaderText = "Identifier JO"
        Me.ColItem_IdentifierJO.Name = "ColItem_IdentifierJO"
        Me.ColItem_IdentifierJO.Visible = False
        Me.ColItem_IdentifierJO.Width = 88
        '
        'chkOngoingJOOnly
        '
        Me.chkOngoingJOOnly.AutoSize = True
        Me.chkOngoingJOOnly.Location = New System.Drawing.Point(76, 32)
        Me.chkOngoingJOOnly.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkOngoingJOOnly.Name = "chkOngoingJOOnly"
        Me.chkOngoingJOOnly.Size = New System.Drawing.Size(169, 17)
        Me.chkOngoingJOOnly.TabIndex = 19
        Me.chkOngoingJOOnly.Text = "Show Ongoing Job Order Only"
        Me.chkOngoingJOOnly.UseVisualStyleBackColor = True
        '
        'llbUnCheckAll
        '
        Me.llbUnCheckAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbUnCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUnCheckAll.AutoSize = True
        Me.llbUnCheckAll.LinkColor = System.Drawing.Color.Blue
        Me.llbUnCheckAll.Location = New System.Drawing.Point(650, 36)
        Me.llbUnCheckAll.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbUnCheckAll.Name = "llbUnCheckAll"
        Me.llbUnCheckAll.Size = New System.Drawing.Size(77, 13)
        Me.llbUnCheckAll.TabIndex = 20
        Me.llbUnCheckAll.TabStop = True
        Me.llbUnCheckAll.Text = "Unchecked All"
        Me.llbUnCheckAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'llbCheckAll
        '
        Me.llbCheckAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.LinkColor = System.Drawing.Color.Blue
        Me.llbCheckAll.Location = New System.Drawing.Point(587, 36)
        Me.llbCheckAll.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(64, 13)
        Me.llbCheckAll.TabIndex = 21
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Checked All"
        Me.llbCheckAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(373, 218)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(56, 19)
        Me.Button1.TabIndex = 22
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(345, 9)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(75, 21)
        Me.btnShow.TabIndex = 23
        Me.btnShow.Text = "&Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'ChkShow100PectOnly
        '
        Me.ChkShow100PectOnly.AutoSize = True
        Me.ChkShow100PectOnly.Location = New System.Drawing.Point(253, 32)
        Me.ChkShow100PectOnly.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ChkShow100PectOnly.Name = "ChkShow100PectOnly"
        Me.ChkShow100PectOnly.Size = New System.Drawing.Size(158, 17)
        Me.ChkShow100PectOnly.TabIndex = 24
        Me.ChkShow100PectOnly.Text = "Show 100% completion only"
        Me.ChkShow100PectOnly.UseVisualStyleBackColor = True
        '
        'FrmTrans_ManufacturingOrder_SetFinishJO
        '
        Me.AcceptButton = Me.btnShow
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(740, 514)
        Me.Controls.Add(Me.ChkShow100PectOnly)
        Me.Controls.Add(Me.btnShow)
        Me.Controls.Add(Me.llbCheckAll)
        Me.Controls.Add(Me.llbUnCheckAll)
        Me.Controls.Add(Me.chkOngoingJOOnly)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimumSize = New System.Drawing.Size(756, 552)
        Me.Name = "FrmTrans_ManufacturingOrder_SetFinishJO"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Set Job Order Status"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents chkOngoingJOOnly As System.Windows.Forms.CheckBox
    Friend WithEvents llbUnCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents ColItem_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_MONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_JONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_QtyTarget As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_QtyProduce As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_IsClosed As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColItem_IdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChkShow100PectOnly As System.Windows.Forms.CheckBox
End Class
