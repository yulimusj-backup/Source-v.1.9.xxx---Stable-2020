<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCetakQuotation
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
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtnPreview = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboVersion = New System.Windows.Forms.ComboBox
        Me.cboLanguage = New System.Windows.Forms.ComboBox
        Me.crv2 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.CrvDetail = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.Location = New System.Drawing.Point(5, 72)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(850, 488)
        Me.crv1.TabIndex = 2
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(324, 41)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.BtnPreview)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.cboVersion)
        Me.GroupBox1.Controls.Add(Me.cboLanguage)
        Me.GroupBox1.Location = New System.Drawing.Point(5, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(846, 68)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'BtnPreview
        '
        Me.BtnPreview.Location = New System.Drawing.Point(324, 14)
        Me.BtnPreview.Name = "BtnPreview"
        Me.BtnPreview.Size = New System.Drawing.Size(75, 23)
        Me.BtnPreview.TabIndex = 2
        Me.BtnPreview.Text = "Preview"
        Me.BtnPreview.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(18, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(42, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Version"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Language"
        '
        'cboVersion
        '
        Me.cboVersion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboVersion.FormattingEnabled = True
        Me.cboVersion.Location = New System.Drawing.Point(100, 41)
        Me.cboVersion.Name = "cboVersion"
        Me.cboVersion.Size = New System.Drawing.Size(207, 21)
        Me.cboVersion.TabIndex = 1
        '
        'cboLanguage
        '
        Me.cboLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLanguage.FormattingEnabled = True
        Me.cboLanguage.Location = New System.Drawing.Point(100, 14)
        Me.cboLanguage.Name = "cboLanguage"
        Me.cboLanguage.Size = New System.Drawing.Size(207, 21)
        Me.cboLanguage.TabIndex = 0
        '
        'crv2
        '
        Me.crv2.ActiveViewIndex = -1
        Me.crv2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv2.DisplayGroupTree = False
        Me.crv2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crv2.Location = New System.Drawing.Point(3, 3)
        Me.crv2.Name = "crv2"
        Me.crv2.SelectionFormula = ""
        Me.crv2.ShowCloseButton = False
        Me.crv2.ShowGroupTreeButton = False
        Me.crv2.ShowRefreshButton = False
        Me.crv2.Size = New System.Drawing.Size(836, 456)
        Me.crv2.TabIndex = 5
        Me.crv2.ViewTimeSelectionFormula = ""
        '
        'CrvDetail
        '
        Me.CrvDetail.ActiveViewIndex = -1
        Me.CrvDetail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrvDetail.DisplayGroupTree = False
        Me.CrvDetail.DisplayStatusBar = False
        Me.CrvDetail.DisplayToolbar = False
        Me.CrvDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrvDetail.Location = New System.Drawing.Point(3, 3)
        Me.CrvDetail.Name = "CrvDetail"
        Me.CrvDetail.SelectionFormula = ""
        Me.CrvDetail.ShowCloseButton = False
        Me.CrvDetail.ShowGroupTreeButton = False
        Me.CrvDetail.ShowRefreshButton = False
        Me.CrvDetail.Size = New System.Drawing.Size(836, 456)
        Me.CrvDetail.TabIndex = 6
        Me.CrvDetail.ViewTimeSelectionFormula = ""
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(5, 72)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(850, 488)
        Me.TabControl1.TabIndex = 7
        Me.TabControl1.TabStop = False
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.crv2)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(842, 462)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Page 1"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.CrvDetail)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(842, 462)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Page 2"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'FrmCetakQuotation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(858, 564)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.crv1)
        Me.Controls.Add(Me.GroupBox1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(874, 541)
        Me.Name = "FrmCetakQuotation"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Quotation"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboVersion As System.Windows.Forms.ComboBox
    Friend WithEvents cboLanguage As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BtnPreview As System.Windows.Forms.Button
    Friend WithEvents crv2 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents CrvDetail As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
End Class
