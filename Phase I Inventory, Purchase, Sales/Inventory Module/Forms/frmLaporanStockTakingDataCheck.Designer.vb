<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLaporanStockTakingDataCheck
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.CrvLaporan = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.cboPeriod = New System.Windows.Forms.ComboBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtTakingNo = New System.Windows.Forms.TextBox
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.btnSearch = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.CrvLaporan)
        Me.GroupBox1.Location = New System.Drawing.Point(5, 54)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(702, 391)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'CrvLaporan
        '
        Me.CrvLaporan.ActiveViewIndex = -1
        Me.CrvLaporan.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrvLaporan.DisplayGroupTree = False
        Me.CrvLaporan.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrvLaporan.Location = New System.Drawing.Point(3, 16)
        Me.CrvLaporan.Name = "CrvLaporan"
        Me.CrvLaporan.SelectionFormula = ""
        Me.CrvLaporan.ShowCloseButton = False
        Me.CrvLaporan.ShowGroupTreeButton = False
        Me.CrvLaporan.ShowRefreshButton = False
        Me.CrvLaporan.Size = New System.Drawing.Size(696, 372)
        Me.CrvLaporan.TabIndex = 0
        Me.CrvLaporan.ViewTimeSelectionFormula = ""
        '
        'cboPeriod
        '
        Me.cboPeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPeriod.FormattingEnabled = True
        Me.cboPeriod.Items.AddRange(New Object() {"Expiry Date Data Check", "Qty Data Check"})
        Me.cboPeriod.Location = New System.Drawing.Point(5, 6)
        Me.cboPeriod.Name = "cboPeriod"
        Me.cboPeriod.Size = New System.Drawing.Size(169, 21)
        Me.cboPeriod.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(458, 28)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(99, 20)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(458, 6)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(99, 20)
        Me.btnFilter.TabIndex = 3
        Me.btnFilter.Text = "&Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(198, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Taking No."
        '
        'txtTakingNo
        '
        Me.txtTakingNo.Location = New System.Drawing.Point(264, 7)
        Me.txtTakingNo.Name = "txtTakingNo"
        Me.txtTakingNo.Size = New System.Drawing.Size(148, 20)
        Me.txtTakingNo.TabIndex = 1
        '
        'txtPrimary
        '
        Me.txtPrimary.Location = New System.Drawing.Point(299, 7)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimary.TabIndex = 22
        Me.txtPrimary.Text = "-1"
        Me.txtPrimary.Visible = False
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(418, 6)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(27, 21)
        Me.btnSearch.TabIndex = 2
        Me.btnSearch.Text = "...."
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'frmLaporanStockTakingDataCheck
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(712, 448)
        Me.Controls.Add(Me.txtPrimary)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.txtTakingNo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.cboPeriod)
        Me.Controls.Add(Me.GroupBox1)
        Me.MinimizeBox = False
        Me.Name = "frmLaporanStockTakingDataCheck"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Laporan Stock Taking Data Check"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CrvLaporan As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents cboPeriod As System.Windows.Forms.ComboBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtTakingNo As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents btnSearch As System.Windows.Forms.Button
End Class
