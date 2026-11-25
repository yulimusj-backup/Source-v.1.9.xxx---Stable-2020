<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanSalaryStaff_Cetak_IGI
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
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.crvReport1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.crvReport2 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.crvReport3 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.crvReport4 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.BtnClose = New System.Windows.Forms.Button
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.crvReport5 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage5)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(984, 632)
        Me.TabControl1.TabIndex = 54
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.crvReport1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(976, 606)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Page 1"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'crvReport1
        '
        Me.crvReport1.ActiveViewIndex = -1
        Me.crvReport1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvReport1.DisplayGroupTree = False
        Me.crvReport1.DisplayStatusBar = False
        Me.crvReport1.DisplayToolbar = False
        Me.crvReport1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvReport1.Location = New System.Drawing.Point(3, 3)
        Me.crvReport1.Name = "crvReport1"
        Me.crvReport1.SelectionFormula = ""
        Me.crvReport1.ShowCloseButton = False
        Me.crvReport1.ShowGroupTreeButton = False
        Me.crvReport1.ShowRefreshButton = False
        Me.crvReport1.Size = New System.Drawing.Size(970, 600)
        Me.crvReport1.TabIndex = 282
        Me.crvReport1.ViewTimeSelectionFormula = ""
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.crvReport2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(976, 606)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Page 2"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'crvReport2
        '
        Me.crvReport2.ActiveViewIndex = -1
        Me.crvReport2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvReport2.DisplayGroupTree = False
        Me.crvReport2.DisplayStatusBar = False
        Me.crvReport2.DisplayToolbar = False
        Me.crvReport2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvReport2.Location = New System.Drawing.Point(3, 3)
        Me.crvReport2.Name = "crvReport2"
        Me.crvReport2.SelectionFormula = ""
        Me.crvReport2.ShowCloseButton = False
        Me.crvReport2.ShowGroupTreeButton = False
        Me.crvReport2.ShowRefreshButton = False
        Me.crvReport2.Size = New System.Drawing.Size(970, 600)
        Me.crvReport2.TabIndex = 282
        Me.crvReport2.ViewTimeSelectionFormula = ""
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.crvReport3)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(976, 606)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Page 3"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'crvReport3
        '
        Me.crvReport3.ActiveViewIndex = -1
        Me.crvReport3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvReport3.DisplayGroupTree = False
        Me.crvReport3.DisplayStatusBar = False
        Me.crvReport3.DisplayToolbar = False
        Me.crvReport3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvReport3.Location = New System.Drawing.Point(0, 0)
        Me.crvReport3.Name = "crvReport3"
        Me.crvReport3.SelectionFormula = ""
        Me.crvReport3.ShowCloseButton = False
        Me.crvReport3.ShowGroupTreeButton = False
        Me.crvReport3.ShowRefreshButton = False
        Me.crvReport3.Size = New System.Drawing.Size(976, 606)
        Me.crvReport3.TabIndex = 282
        Me.crvReport3.ViewTimeSelectionFormula = ""
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.crvReport4)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(976, 606)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Page 4"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'crvReport4
        '
        Me.crvReport4.ActiveViewIndex = -1
        Me.crvReport4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvReport4.DisplayGroupTree = False
        Me.crvReport4.DisplayStatusBar = False
        Me.crvReport4.DisplayToolbar = False
        Me.crvReport4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvReport4.Location = New System.Drawing.Point(0, 0)
        Me.crvReport4.Name = "crvReport4"
        Me.crvReport4.SelectionFormula = ""
        Me.crvReport4.ShowCloseButton = False
        Me.crvReport4.ShowGroupTreeButton = False
        Me.crvReport4.ShowRefreshButton = False
        Me.crvReport4.Size = New System.Drawing.Size(976, 606)
        Me.crvReport4.TabIndex = 282
        Me.crvReport4.ViewTimeSelectionFormula = ""
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(508, 103)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(60, 62)
        Me.BtnClose.TabIndex = 55
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.crvReport5)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(976, 606)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Page 5"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'crvReport5
        '
        Me.crvReport5.ActiveViewIndex = -1
        Me.crvReport5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvReport5.DisplayGroupTree = False
        Me.crvReport5.DisplayStatusBar = False
        Me.crvReport5.DisplayToolbar = False
        Me.crvReport5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvReport5.Location = New System.Drawing.Point(3, 3)
        Me.crvReport5.Name = "crvReport5"
        Me.crvReport5.SelectionFormula = ""
        Me.crvReport5.ShowCloseButton = False
        Me.crvReport5.ShowGroupTreeButton = False
        Me.crvReport5.ShowRefreshButton = False
        Me.crvReport5.Size = New System.Drawing.Size(970, 600)
        Me.crvReport5.TabIndex = 283
        Me.crvReport5.ViewTimeSelectionFormula = ""
        '
        'FrmLaporanSalaryStaff_Cetak_IGI
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(984, 632)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.BtnClose)
        Me.MinimizeBox = False
        Me.Name = "FrmLaporanSalaryStaff_Cetak_IGI"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents crvReport1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents crvReport2 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents crvReport3 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents crvReport4 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents crvReport5 As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
