<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ManufacturingOrder_Exp_QtyReq
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.DgvItemBundling = New System.Windows.Forms.DataGridView
        Me.Button1 = New System.Windows.Forms.Button
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtJONumber = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtCodeNo = New System.Windows.Forms.TextBox
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.TxtExplain = New System.Windows.Forms.TextBox
        Me.TxtExplain1 = New System.Windows.Forms.TextBox
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMQTY = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWipQTY = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyRef = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyRefRM = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGPProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGPRM = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDecimal_Product = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDecimalRM = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.DgvItemBundling, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SuspendLayout()
        '
        'DgvItemBundling
        '
        Me.DgvItemBundling.AllowUserToAddRows = False
        Me.DgvItemBundling.AllowUserToDeleteRows = False
        Me.DgvItemBundling.AllowUserToResizeColumns = False
        Me.DgvItemBundling.AllowUserToResizeRows = False
        Me.DgvItemBundling.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItemBundling.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DgvItemBundling.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DgvItemBundling.ColumnHeadersHeight = 25
        Me.DgvItemBundling.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItemBundling.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColProcess, Me.ColDocNo, Me.ColCodeNo, Me.ColDesc, Me.ColQty, Me.ColUnit, Me.ColRMCode, Me.ColRMDesc, Me.ColRMQTY, Me.ColWipQTY, Me.ColRMUnit, Me.ColQtyRef, Me.ColQtyRefRM, Me.ColGPProcess, Me.ColGPRM, Me.ColDecimal_Product, Me.ColDecimalRM})
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DgvItemBundling.DefaultCellStyle = DataGridViewCellStyle5
        Me.DgvItemBundling.Location = New System.Drawing.Point(12, 80)
        Me.DgvItemBundling.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DgvItemBundling.Name = "DgvItemBundling"
        Me.DgvItemBundling.ReadOnly = True
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DgvItemBundling.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.DgvItemBundling.RowHeadersVisible = False
        Me.DgvItemBundling.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItemBundling.RowTemplate.Height = 24
        Me.DgvItemBundling.Size = New System.Drawing.Size(968, 237)
        Me.DgvItemBundling.TabIndex = 2
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(726, 93)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 4
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtJONumber)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtCodeNo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtDesc)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Size = New System.Drawing.Size(968, 63)
        Me.SplitContainer1.SplitterDistance = 483
        Me.SplitContainer1.TabIndex = 5
        '
        'TxtJONumber
        '
        Me.TxtJONumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtJONumber.Location = New System.Drawing.Point(97, 3)
        Me.TxtJONumber.Name = "TxtJONumber"
        Me.TxtJONumber.ReadOnly = True
        Me.TxtJONumber.Size = New System.Drawing.Size(383, 22)
        Me.TxtJONumber.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 17)
        Me.Label1.TabIndex = 66
        Me.Label1.Text = "JO Number"
        '
        'TxtCodeNo
        '
        Me.TxtCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeNo.Location = New System.Drawing.Point(96, 3)
        Me.TxtCodeNo.Name = "TxtCodeNo"
        Me.TxtCodeNo.ReadOnly = True
        Me.TxtCodeNo.Size = New System.Drawing.Size(382, 22)
        Me.TxtCodeNo.TabIndex = 0
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(96, 31)
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.ReadOnly = True
        Me.TxtDesc.Size = New System.Drawing.Size(382, 22)
        Me.TxtDesc.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(1, 34)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(79, 17)
        Me.Label3.TabIndex = 72
        Me.Label3.Text = "Description"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(1, 6)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(63, 17)
        Me.Label4.TabIndex = 70
        Me.Label4.Text = "Code No"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(12, 322)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtExplain)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtExplain1)
        Me.SplitContainer2.Size = New System.Drawing.Size(968, 231)
        Me.SplitContainer2.SplitterDistance = 483
        Me.SplitContainer2.TabIndex = 6
        '
        'TxtExplain
        '
        Me.TxtExplain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TxtExplain.Location = New System.Drawing.Point(0, 0)
        Me.TxtExplain.Multiline = True
        Me.TxtExplain.Name = "TxtExplain"
        Me.TxtExplain.ReadOnly = True
        Me.TxtExplain.Size = New System.Drawing.Size(483, 231)
        Me.TxtExplain.TabIndex = 4
        Me.TxtExplain.TabStop = False
        '
        'TxtExplain1
        '
        Me.TxtExplain1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TxtExplain1.Location = New System.Drawing.Point(0, 0)
        Me.TxtExplain1.Multiline = True
        Me.TxtExplain1.Name = "TxtExplain1"
        Me.TxtExplain1.ReadOnly = True
        Me.TxtExplain1.Size = New System.Drawing.Size(481, 231)
        Me.TxtExplain1.TabIndex = 5
        Me.TxtExplain1.TabStop = False
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'ColProcess
        '
        Me.ColProcess.HeaderText = "Process"
        Me.ColProcess.Name = "ColProcess"
        Me.ColProcess.ReadOnly = True
        Me.ColProcess.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProcess.Width = 65
        '
        'ColDocNo
        '
        Me.ColDocNo.HeaderText = "Document No"
        Me.ColDocNo.Name = "ColDocNo"
        Me.ColDocNo.ReadOnly = True
        Me.ColDocNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColDocNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDocNo.Visible = False
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.ReadOnly = True
        Me.ColCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCodeNo.Visible = False
        Me.ColCodeNo.Width = 69
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.ReadOnly = True
        Me.ColDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDesc.Visible = False
        Me.ColDesc.Width = 85
        '
        'ColQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Visible = False
        Me.ColQty.Width = 36
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.ReadOnly = True
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Visible = False
        Me.ColUnit.Width = 39
        '
        'ColRMCode
        '
        Me.ColRMCode.HeaderText = "RM Code"
        Me.ColRMCode.Name = "ColRMCode"
        Me.ColRMCode.ReadOnly = True
        Me.ColRMCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRMCode.Width = 72
        '
        'ColRMDesc
        '
        Me.ColRMDesc.HeaderText = "RM Descrition"
        Me.ColRMDesc.Name = "ColRMDesc"
        Me.ColRMDesc.ReadOnly = True
        Me.ColRMDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRMDesc.Width = 102
        '
        'ColRMQTY
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRMQTY.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColRMQTY.HeaderText = "RM Qty"
        Me.ColRMQTY.Name = "ColRMQTY"
        Me.ColRMQTY.ReadOnly = True
        Me.ColRMQTY.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRMQTY.Width = 61
        '
        'ColWipQTY
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColWipQTY.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColWipQTY.HeaderText = "WIP Qty"
        Me.ColWipQTY.Name = "ColWipQTY"
        Me.ColWipQTY.ReadOnly = True
        Me.ColWipQTY.Width = 84
        '
        'ColRMUnit
        '
        Me.ColRMUnit.HeaderText = "Unit"
        Me.ColRMUnit.Name = "ColRMUnit"
        Me.ColRMUnit.ReadOnly = True
        Me.ColRMUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRMUnit.Width = 39
        '
        'ColQtyRef
        '
        Me.ColQtyRef.HeaderText = "Qty Ref"
        Me.ColQtyRef.Name = "ColQtyRef"
        Me.ColQtyRef.ReadOnly = True
        Me.ColQtyRef.Visible = False
        Me.ColQtyRef.Width = 81
        '
        'ColQtyRefRM
        '
        Me.ColQtyRefRM.HeaderText = "Qty Ref RM"
        Me.ColQtyRefRM.Name = "ColQtyRefRM"
        Me.ColQtyRefRM.ReadOnly = True
        Me.ColQtyRefRM.Visible = False
        Me.ColQtyRefRM.Width = 106
        '
        'ColGPProcess
        '
        Me.ColGPProcess.HeaderText = "GP Process"
        Me.ColGPProcess.Name = "ColGPProcess"
        Me.ColGPProcess.ReadOnly = True
        Me.ColGPProcess.Visible = False
        Me.ColGPProcess.Width = 108
        '
        'ColGPRM
        '
        Me.ColGPRM.HeaderText = "GP RM"
        Me.ColGPRM.Name = "ColGPRM"
        Me.ColGPRM.ReadOnly = True
        Me.ColGPRM.Visible = False
        Me.ColGPRM.Width = 78
        '
        'ColDecimal_Product
        '
        Me.ColDecimal_Product.HeaderText = "Decimal_Product"
        Me.ColDecimal_Product.Name = "ColDecimal_Product"
        Me.ColDecimal_Product.ReadOnly = True
        Me.ColDecimal_Product.Visible = False
        Me.ColDecimal_Product.Width = 140
        '
        'ColDecimalRM
        '
        Me.ColDecimalRM.HeaderText = "Decimal RM"
        Me.ColDecimalRM.Name = "ColDecimalRM"
        Me.ColDecimalRM.ReadOnly = True
        Me.ColDecimalRM.Visible = False
        Me.ColDecimalRM.Width = 108
        '
        'FrmTrans_ManufacturingOrder_Exp_QtyReq
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(992, 565)
        Me.Controls.Add(Me.SplitContainer2)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.DgvItemBundling)
        Me.Controls.Add(Me.Button1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1000, 484)
        Me.Name = "FrmTrans_ManufacturingOrder_Exp_QtyReq"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Qty Required"
        CType(Me.DgvItemBundling, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DgvItemBundling As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtJONumber As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtExplain As System.Windows.Forms.TextBox
    Friend WithEvents TxtExplain1 As System.Windows.Forms.TextBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMQTY As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWipQTY As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyRef As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyRefRM As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGPProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGPRM As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDecimal_Product As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDecimalRM As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
