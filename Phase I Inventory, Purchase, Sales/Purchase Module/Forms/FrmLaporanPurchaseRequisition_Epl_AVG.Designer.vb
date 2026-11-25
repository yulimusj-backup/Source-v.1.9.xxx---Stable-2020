<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanPurchaseRequisition_Epl_AVG
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.LblUnit4 = New System.Windows.Forms.Label
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvDO = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPeriod = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.LblUnit1 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.LblUnit3 = New System.Windows.Forms.Label
        Me.LblUnit2 = New System.Windows.Forms.Label
        Me.TxtAvgPerMonth = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtTotalPeriode = New System.Windows.Forms.TextBox
        Me.TxtTotal = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtTotalHS = New System.Windows.Forms.TextBox
        Me.TxtTotalDO = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.dgvHS = New System.Windows.Forms.DataGridView
        Me.ColNoHS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHSNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPeriodHS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyTotalHS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFormula = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtySales = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvDO, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvHS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LblUnit4
        '
        Me.LblUnit4.AutoSize = True
        Me.LblUnit4.Location = New System.Drawing.Point(265, 121)
        Me.LblUnit4.Name = "LblUnit4"
        Me.LblUnit4.Size = New System.Drawing.Size(50, 13)
        Me.LblUnit4.TabIndex = 34
        Me.LblUnit4.Text = "Total DO"
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvDO)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(689, 317)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Delivery Order"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvDO
        '
        Me.dgvDO.AllowUserToAddRows = False
        Me.dgvDO.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvDO.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvDO.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDO.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvDO.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColPeriod, Me.ColQty, Me.ColUnits})
        Me.dgvDO.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDO.Location = New System.Drawing.Point(3, 3)
        Me.dgvDO.Name = "dgvDO"
        Me.dgvDO.ReadOnly = True
        Me.dgvDO.Size = New System.Drawing.Size(683, 311)
        Me.dgvDO.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 46
        '
        'ColPeriod
        '
        Me.ColPeriod.HeaderText = "Period"
        Me.ColPeriod.Name = "ColPeriod"
        Me.ColPeriod.ReadOnly = True
        Me.ColPeriod.Width = 62
        '
        'ColQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle2.NullValue = Nothing
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.Width = 48
        '
        'ColUnits
        '
        Me.ColUnits.HeaderText = "Units"
        Me.ColUnits.Name = "ColUnits"
        Me.ColUnits.ReadOnly = True
        Me.ColUnits.Width = 56
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 144)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(697, 343)
        Me.TabControl1.TabIndex = 30
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvHS)
        Me.TabPage2.Controls.Add(Me.Button1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(689, 317)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Historical Sales"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'LblUnit1
        '
        Me.LblUnit1.AutoSize = True
        Me.LblUnit1.Location = New System.Drawing.Point(265, 17)
        Me.LblUnit1.Name = "LblUnit1"
        Me.LblUnit1.Size = New System.Drawing.Size(50, 13)
        Me.LblUnit1.TabIndex = 29
        Me.LblUnit1.Text = "Total DO"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(265, 95)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(48, 13)
        Me.Label8.TabIndex = 33
        Me.Label8.Text = "Month(s)"
        '
        'LblUnit3
        '
        Me.LblUnit3.AutoSize = True
        Me.LblUnit3.Location = New System.Drawing.Point(265, 69)
        Me.LblUnit3.Name = "LblUnit3"
        Me.LblUnit3.Size = New System.Drawing.Size(50, 13)
        Me.LblUnit3.TabIndex = 32
        Me.LblUnit3.Text = "Total DO"
        '
        'LblUnit2
        '
        Me.LblUnit2.AutoSize = True
        Me.LblUnit2.Location = New System.Drawing.Point(265, 43)
        Me.LblUnit2.Name = "LblUnit2"
        Me.LblUnit2.Size = New System.Drawing.Size(50, 13)
        Me.LblUnit2.TabIndex = 31
        Me.LblUnit2.Text = "Total DO"
        '
        'TxtAvgPerMonth
        '
        Me.TxtAvgPerMonth.Location = New System.Drawing.Point(159, 118)
        Me.TxtAvgPerMonth.Name = "TxtAvgPerMonth"
        Me.TxtAvgPerMonth.ReadOnly = True
        Me.TxtAvgPerMonth.Size = New System.Drawing.Size(100, 20)
        Me.TxtAvgPerMonth.TabIndex = 28
        Me.TxtAvgPerMonth.TabStop = False
        Me.TxtAvgPerMonth.Text = "0"
        Me.TxtAvgPerMonth.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 121)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(99, 13)
        Me.Label6.TabIndex = 27
        Me.Label6.Text = "Average Per Month"
        '
        'TxtTotalPeriode
        '
        Me.TxtTotalPeriode.Location = New System.Drawing.Point(159, 92)
        Me.TxtTotalPeriode.Name = "TxtTotalPeriode"
        Me.TxtTotalPeriode.ReadOnly = True
        Me.TxtTotalPeriode.Size = New System.Drawing.Size(100, 20)
        Me.TxtTotalPeriode.TabIndex = 26
        Me.TxtTotalPeriode.TabStop = False
        Me.TxtTotalPeriode.Text = "0"
        Me.TxtTotalPeriode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtTotal
        '
        Me.TxtTotal.Location = New System.Drawing.Point(159, 66)
        Me.TxtTotal.Name = "TxtTotal"
        Me.TxtTotal.ReadOnly = True
        Me.TxtTotal.Size = New System.Drawing.Size(100, 20)
        Me.TxtTotal.TabIndex = 25
        Me.TxtTotal.TabStop = False
        Me.TxtTotal.Text = "0"
        Me.TxtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 95)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(37, 13)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Period"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 69)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(31, 13)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "Total"
        '
        'TxtTotalHS
        '
        Me.TxtTotalHS.Location = New System.Drawing.Point(159, 40)
        Me.TxtTotalHS.Name = "TxtTotalHS"
        Me.TxtTotalHS.ReadOnly = True
        Me.TxtTotalHS.Size = New System.Drawing.Size(100, 20)
        Me.TxtTotalHS.TabIndex = 22
        Me.TxtTotalHS.TabStop = False
        Me.TxtTotalHS.Text = "0"
        Me.TxtTotalHS.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtTotalDO
        '
        Me.TxtTotalDO.Location = New System.Drawing.Point(159, 14)
        Me.TxtTotalDO.Name = "TxtTotalDO"
        Me.TxtTotalDO.ReadOnly = True
        Me.TxtTotalDO.Size = New System.Drawing.Size(100, 20)
        Me.TxtTotalDO.TabIndex = 21
        Me.TxtTotalDO.TabStop = False
        Me.TxtTotalDO.Text = "0"
        Me.TxtTotalDO.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 43)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(106, 13)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "Total Historical Sales"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 13)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Total DO"
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(401, 58)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 18
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'dgvHS
        '
        Me.dgvHS.AllowUserToAddRows = False
        Me.dgvHS.AllowUserToDeleteRows = False
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvHS.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvHS.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvHS.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvHS.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNoHS, Me.ColHSNumber, Me.ColPeriodHS, Me.ColQtyTotalHS, Me.ColFormula, Me.ColQtySales})
        Me.dgvHS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvHS.Location = New System.Drawing.Point(3, 3)
        Me.dgvHS.Name = "dgvHS"
        Me.dgvHS.ReadOnly = True
        Me.dgvHS.Size = New System.Drawing.Size(683, 311)
        Me.dgvHS.TabIndex = 0
        '
        'ColNoHS
        '
        Me.ColNoHS.HeaderText = "No"
        Me.ColNoHS.Name = "ColNoHS"
        Me.ColNoHS.ReadOnly = True
        Me.ColNoHS.Width = 46
        '
        'ColHSNumber
        '
        Me.ColHSNumber.HeaderText = "Transaction Number"
        Me.ColHSNumber.Name = "ColHSNumber"
        Me.ColHSNumber.ReadOnly = True
        Me.ColHSNumber.Width = 128
        '
        'ColPeriodHS
        '
        Me.ColPeriodHS.HeaderText = "Historical Sales Period"
        Me.ColPeriodHS.Name = "ColPeriodHS"
        Me.ColPeriodHS.ReadOnly = True
        Me.ColPeriodHS.Width = 137
        '
        'ColQtyTotalHS
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyTotalHS.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColQtyTotalHS.HeaderText = "Total Qty"
        Me.ColQtyTotalHS.Name = "ColQtyTotalHS"
        Me.ColQtyTotalHS.ReadOnly = True
        Me.ColQtyTotalHS.Width = 75
        '
        'ColFormula
        '
        Me.ColFormula.HeaderText = "Formula"
        Me.ColFormula.Name = "ColFormula"
        Me.ColFormula.ReadOnly = True
        Me.ColFormula.Width = 69
        '
        'ColQtySales
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtySales.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColQtySales.HeaderText = "Sales Qty"
        Me.ColQtySales.Name = "ColQtySales"
        Me.ColQtySales.ReadOnly = True
        Me.ColQtySales.Width = 77
        '
        'FrmLaporanPurchaseRequisition_Epl_AVG
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(721, 500)
        Me.Controls.Add(Me.LblUnit4)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.LblUnit1)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.LblUnit3)
        Me.Controls.Add(Me.LblUnit2)
        Me.Controls.Add(Me.TxtAvgPerMonth)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TxtTotalPeriode)
        Me.Controls.Add(Me.TxtTotal)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TxtTotalHS)
        Me.Controls.Add(Me.TxtTotalDO)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(737, 536)
        Me.Name = "FrmLaporanPurchaseRequisition_Epl_AVG"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FrmLaporanPurchaseRequisition_Epl_2"
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvDO, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvHS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LblUnit4 As System.Windows.Forms.Label
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents LblUnit1 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents LblUnit3 As System.Windows.Forms.Label
    Friend WithEvents LblUnit2 As System.Windows.Forms.Label
    Friend WithEvents TxtAvgPerMonth As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalPeriode As System.Windows.Forms.TextBox
    Friend WithEvents TxtTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalHS As System.Windows.Forms.TextBox
    Friend WithEvents TxtTotalDO As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents dgvDO As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPeriod As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvHS As System.Windows.Forms.DataGridView
    Friend WithEvents ColNoHS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHSNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPeriodHS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyTotalHS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFormula As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtySales As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
