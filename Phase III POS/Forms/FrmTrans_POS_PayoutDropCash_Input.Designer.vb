<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_POS_PayoutDropCash_Input
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
        Me.ColTipe = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAutoNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPayout = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySession = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtPrimarySession = New System.Windows.Forms.TextBox
        Me.TxtPrimaryPayOutDropCash = New System.Windows.Forms.TextBox
        Me.CboProdCode = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboPayout_DropCash = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnInput = New System.Windows.Forms.Button
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtAmount = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColTipe, Me.ColCode, Me.ColDesc, Me.ColAmount, Me.ColRemarks, Me.ColPrimaryAutoNumber, Me.ColPrimaryPayout, Me.ColPrimarySession})
        Me.dgvItem.Location = New System.Drawing.Point(12, 97)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.Size = New System.Drawing.Size(560, 324)
        Me.dgvItem.TabIndex = 1
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColTipe
        '
        Me.ColTipe.HeaderText = "PayOut/Drop Cash"
        Me.ColTipe.Items.AddRange(New Object() {"Payout", "Drop Cash"})
        Me.ColTipe.Name = "ColTipe"
        Me.ColTipe.Width = 103
        '
        'ColCode
        '
        Me.ColCode.HeaderText = "Code"
        Me.ColCode.Name = "ColCode"
        Me.ColCode.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCode.Width = 38
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDesc.Width = 66
        '
        'ColAmount
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.MaxInputLength = 20
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.Width = 68
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.MaxInputLength = 250
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 55
        '
        'ColPrimaryAutoNumber
        '
        Me.ColPrimaryAutoNumber.HeaderText = "Primary Auto Number"
        Me.ColPrimaryAutoNumber.Name = "ColPrimaryAutoNumber"
        Me.ColPrimaryAutoNumber.Visible = False
        Me.ColPrimaryAutoNumber.Width = 131
        '
        'ColPrimaryPayout
        '
        Me.ColPrimaryPayout.HeaderText = "Primary Payout"
        Me.ColPrimaryPayout.Name = "ColPrimaryPayout"
        Me.ColPrimaryPayout.Visible = False
        Me.ColPrimaryPayout.Width = 102
        '
        'ColPrimarySession
        '
        Me.ColPrimarySession.HeaderText = "Primary Session"
        Me.ColPrimarySession.Name = "ColPrimarySession"
        Me.ColPrimarySession.Visible = False
        Me.ColPrimarySession.Width = 106
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(12, 12)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimarySession)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryPayOutDropCash)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboProdCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDesc)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboPayout_DropCash)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnInput)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtAmount)
        Me.SplitContainer1.Size = New System.Drawing.Size(560, 79)
        Me.SplitContainer1.SplitterDistance = 279
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'TxtPrimarySession
        '
        Me.TxtPrimarySession.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimarySession.Location = New System.Drawing.Point(61, 56)
        Me.TxtPrimarySession.Name = "TxtPrimarySession"
        Me.TxtPrimarySession.Size = New System.Drawing.Size(37, 20)
        Me.TxtPrimarySession.TabIndex = 9
        Me.TxtPrimarySession.TabStop = False
        Me.TxtPrimarySession.Text = "-1"
        Me.TxtPrimarySession.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtPrimarySession.Visible = False
        '
        'TxtPrimaryPayOutDropCash
        '
        Me.TxtPrimaryPayOutDropCash.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryPayOutDropCash.Location = New System.Drawing.Point(61, 32)
        Me.TxtPrimaryPayOutDropCash.Name = "TxtPrimaryPayOutDropCash"
        Me.TxtPrimaryPayOutDropCash.Size = New System.Drawing.Size(37, 20)
        Me.TxtPrimaryPayOutDropCash.TabIndex = 8
        Me.TxtPrimaryPayOutDropCash.TabStop = False
        Me.TxtPrimaryPayOutDropCash.Text = "-1"
        Me.TxtPrimaryPayOutDropCash.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtPrimaryPayOutDropCash.Visible = False
        '
        'CboProdCode
        '
        Me.CboProdCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboProdCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboProdCode.FormattingEnabled = True
        Me.CboProdCode.Location = New System.Drawing.Point(104, 30)
        Me.CboProdCode.Name = "CboProdCode"
        Me.CboProdCode.Size = New System.Drawing.Size(172, 21)
        Me.CboProdCode.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 61)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Description"
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(104, 57)
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.ReadOnly = True
        Me.TxtDesc.Size = New System.Drawing.Size(172, 20)
        Me.TxtDesc.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Code"
        '
        'cboPayout_DropCash
        '
        Me.cboPayout_DropCash.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboPayout_DropCash.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPayout_DropCash.FormattingEnabled = True
        Me.cboPayout_DropCash.Items.AddRange(New Object() {"Payout", "Drop Cash"})
        Me.cboPayout_DropCash.Location = New System.Drawing.Point(104, 3)
        Me.cboPayout_DropCash.Name = "cboPayout_DropCash"
        Me.cboPayout_DropCash.Size = New System.Drawing.Size(172, 21)
        Me.cboPayout_DropCash.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(95, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Payout/Drop Cash"
        '
        'btnInput
        '
        Me.btnInput.Location = New System.Drawing.Point(102, 54)
        Me.btnInput.Name = "btnInput"
        Me.btnInput.Size = New System.Drawing.Size(50, 23)
        Me.btnInput.TabIndex = 2
        Me.btnInput.Text = "Input"
        Me.btnInput.UseVisualStyleBackColor = True
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(102, 28)
        Me.TxtRemarks.MaxLength = 250
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(172, 20)
        Me.TxtRemarks.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 35)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 13)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Remarks"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 6)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(43, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Amount"
        '
        'TxtAmount
        '
        Me.TxtAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAmount.Location = New System.Drawing.Point(102, 3)
        Me.TxtAmount.MaxLength = 12
        Me.TxtAmount.Name = "TxtAmount"
        Me.TxtAmount.Size = New System.Drawing.Size(172, 20)
        Me.TxtAmount.TabIndex = 0
        Me.TxtAmount.Text = "0"
        Me.TxtAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(497, 427)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(416, 427)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmTrans_POS_PayoutDropCash_Input
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(584, 462)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.dgvItem)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(600, 500)
        Me.Name = "FrmTrans_POS_PayoutDropCash_Input"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Payout Drop Cash"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents cboPayout_DropCash As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnInput As System.Windows.Forms.Button
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtAmount As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryPayOutDropCash As System.Windows.Forms.TextBox
    Friend WithEvents CboProdCode As System.Windows.Forms.ComboBox
    Friend WithEvents TxtPrimarySession As System.Windows.Forms.TextBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipe As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAutoNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPayout As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySession As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
