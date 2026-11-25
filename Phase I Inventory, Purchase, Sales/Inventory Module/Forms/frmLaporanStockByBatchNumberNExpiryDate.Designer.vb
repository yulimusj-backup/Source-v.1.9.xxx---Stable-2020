<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLaporanStockByBatchNumberNExpiryDate
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
        Me.lbBrandCodeOrCodeNumber = New System.Windows.Forms.Label
        Me.txtBrandCodeOrCodeNumber = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.cboCode = New System.Windows.Forms.ComboBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.cboTgl = New System.Windows.Forms.ComboBox
        Me.dtpTgl1 = New System.Windows.Forms.DateTimePicker
        Me.lbTgl1 = New System.Windows.Forms.Label
        Me.lbTgl2 = New System.Windows.Forms.Label
        Me.dtpTgl2 = New System.Windows.Forms.DateTimePicker
        Me.btnSearch = New System.Windows.Forms.Button
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'lbBrandCodeOrCodeNumber
        '
        Me.lbBrandCodeOrCodeNumber.AutoSize = True
        Me.lbBrandCodeOrCodeNumber.Location = New System.Drawing.Point(150, 13)
        Me.lbBrandCodeOrCodeNumber.Name = "lbBrandCodeOrCodeNumber"
        Me.lbBrandCodeOrCodeNumber.Size = New System.Drawing.Size(63, 13)
        Me.lbBrandCodeOrCodeNumber.TabIndex = 8
        Me.lbBrandCodeOrCodeNumber.Text = "Brand Code"
        '
        'txtBrandCodeOrCodeNumber
        '
        Me.txtBrandCodeOrCodeNumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtBrandCodeOrCodeNumber.Location = New System.Drawing.Point(238, 10)
        Me.txtBrandCodeOrCodeNumber.Name = "txtBrandCodeOrCodeNumber"
        Me.txtBrandCodeOrCodeNumber.Size = New System.Drawing.Size(137, 20)
        Me.txtBrandCodeOrCodeNumber.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.crv)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 90)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(739, 416)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        '
        'crv
        '
        Me.crv.ActiveViewIndex = -1
        Me.crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv.DisplayGroupTree = False
        Me.crv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crv.Location = New System.Drawing.Point(3, 16)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(733, 397)
        Me.crv.TabIndex = 0
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'cboCode
        '
        Me.cboCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCode.FormattingEnabled = True
        Me.cboCode.Items.AddRange(New Object() {"Product Code", "Brand Code"})
        Me.cboCode.Location = New System.Drawing.Point(6, 19)
        Me.cboCode.Name = "cboCode"
        Me.cboCode.Size = New System.Drawing.Size(129, 21)
        Me.cboCode.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(412, 33)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(99, 24)
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(412, 7)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(99, 24)
        Me.btnFilter.TabIndex = 5
        Me.btnFilter.Text = "&Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'cboTgl
        '
        Me.cboTgl.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTgl.FormattingEnabled = True
        Me.cboTgl.Items.AddRange(New Object() {"After", "Before", "Between"})
        Me.cboTgl.Location = New System.Drawing.Point(6, 48)
        Me.cboTgl.Name = "cboTgl"
        Me.cboTgl.Size = New System.Drawing.Size(129, 21)
        Me.cboTgl.TabIndex = 1
        '
        'dtpTgl1
        '
        Me.dtpTgl1.CustomFormat = "dd MMM yyyy"
        Me.dtpTgl1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTgl1.Location = New System.Drawing.Point(238, 38)
        Me.dtpTgl1.Name = "dtpTgl1"
        Me.dtpTgl1.Size = New System.Drawing.Size(168, 20)
        Me.dtpTgl1.TabIndex = 3
        '
        'lbTgl1
        '
        Me.lbTgl1.AutoSize = True
        Me.lbTgl1.Location = New System.Drawing.Point(150, 41)
        Me.lbTgl1.Name = "lbTgl1"
        Me.lbTgl1.Size = New System.Drawing.Size(29, 13)
        Me.lbTgl1.TabIndex = 9
        Me.lbTgl1.Text = "After"
        '
        'lbTgl2
        '
        Me.lbTgl2.AutoSize = True
        Me.lbTgl2.Location = New System.Drawing.Point(150, 70)
        Me.lbTgl2.Name = "lbTgl2"
        Me.lbTgl2.Size = New System.Drawing.Size(25, 13)
        Me.lbTgl2.TabIndex = 10
        Me.lbTgl2.Text = "thru"
        Me.lbTgl2.Visible = False
        '
        'dtpTgl2
        '
        Me.dtpTgl2.CustomFormat = "dd MMM yyyy"
        Me.dtpTgl2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTgl2.Location = New System.Drawing.Point(238, 64)
        Me.dtpTgl2.Name = "dtpTgl2"
        Me.dtpTgl2.Size = New System.Drawing.Size(168, 20)
        Me.dtpTgl2.TabIndex = 4
        Me.dtpTgl2.Visible = False
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(379, 9)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(27, 21)
        Me.btnSearch.TabIndex = 2
        Me.btnSearch.TabStop = False
        Me.btnSearch.Text = "...."
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'txtPrimary
        '
        Me.txtPrimary.Location = New System.Drawing.Point(412, 9)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(58, 20)
        Me.txtPrimary.TabIndex = 5
        Me.txtPrimary.Text = "-1"
        Me.txtPrimary.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cboCode)
        Me.GroupBox2.Controls.Add(Me.cboTgl)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 7)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(141, 77)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Filter"
        '
        'frmLaporanStockByBatchNumberNExpiryDate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(746, 514)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.txtPrimary)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.dtpTgl2)
        Me.Controls.Add(Me.lbTgl2)
        Me.Controls.Add(Me.lbTgl1)
        Me.Controls.Add(Me.dtpTgl1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtBrandCodeOrCodeNumber)
        Me.Controls.Add(Me.lbBrandCodeOrCodeNumber)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MinimumSize = New System.Drawing.Size(752, 542)
        Me.Name = "frmLaporanStockByBatchNumberNExpiryDate"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stock By Batch No"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbBrandCodeOrCodeNumber As System.Windows.Forms.Label
    Friend WithEvents txtBrandCodeOrCodeNumber As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cboCode As System.Windows.Forms.ComboBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents cboTgl As System.Windows.Forms.ComboBox
    Friend WithEvents dtpTgl1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents lbTgl1 As System.Windows.Forms.Label
    Friend WithEvents lbTgl2 As System.Windows.Forms.Label
    Friend WithEvents dtpTgl2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
End Class
