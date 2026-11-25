<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanStandarCompetence
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.LblUnchecked = New System.Windows.Forms.LinkLabel
        Me.LblChecked = New System.Windows.Forms.LinkLabel
        Me.GboLaporan = New System.Windows.Forms.GroupBox
        Me.clbJobPosition = New System.Windows.Forms.CheckedListBox
        Me.GboLaporan.SuspendLayout()
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
        Me.crv1.Location = New System.Drawing.Point(12, 112)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(860, 438)
        Me.crv1.TabIndex = 79
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(765, 180)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 23)
        Me.btnClose.TabIndex = 84
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnShow.Location = New System.Drawing.Point(796, 12)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(76, 23)
        Me.btnShow.TabIndex = 83
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'LblUnchecked
        '
        Me.LblUnchecked.AutoSize = True
        Me.LblUnchecked.BackColor = System.Drawing.SystemColors.Control
        Me.LblUnchecked.Location = New System.Drawing.Point(80, 98)
        Me.LblUnchecked.Name = "LblUnchecked"
        Me.LblUnchecked.Size = New System.Drawing.Size(77, 13)
        Me.LblUnchecked.TabIndex = 86
        Me.LblUnchecked.TabStop = True
        Me.LblUnchecked.Text = "Unchecked All"
        '
        'LblChecked
        '
        Me.LblChecked.AutoSize = True
        Me.LblChecked.BackColor = System.Drawing.SystemColors.Control
        Me.LblChecked.Location = New System.Drawing.Point(18, 98)
        Me.LblChecked.Name = "LblChecked"
        Me.LblChecked.Size = New System.Drawing.Size(64, 13)
        Me.LblChecked.TabIndex = 87
        Me.LblChecked.TabStop = True
        Me.LblChecked.Text = "Checked All"
        '
        'GboLaporan
        '
        Me.GboLaporan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GboLaporan.Controls.Add(Me.clbJobPosition)
        Me.GboLaporan.Location = New System.Drawing.Point(12, 0)
        Me.GboLaporan.Name = "GboLaporan"
        Me.GboLaporan.Size = New System.Drawing.Size(767, 106)
        Me.GboLaporan.TabIndex = 88
        Me.GboLaporan.TabStop = False
        Me.GboLaporan.Text = "Job Position"
        '
        'clbJobPosition
        '
        Me.clbJobPosition.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.clbJobPosition.CheckOnClick = True
        Me.clbJobPosition.FormattingEnabled = True
        Me.clbJobPosition.Location = New System.Drawing.Point(3, 16)
        Me.clbJobPosition.Name = "clbJobPosition"
        Me.clbJobPosition.Size = New System.Drawing.Size(761, 79)
        Me.clbJobPosition.TabIndex = 86
        '
        'FrmLaporanStandarCompetence
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(884, 562)
        Me.Controls.Add(Me.LblUnchecked)
        Me.Controls.Add(Me.LblChecked)
        Me.Controls.Add(Me.crv1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnShow)
        Me.Controls.Add(Me.GboLaporan)
        Me.Name = "FrmLaporanStandarCompetence"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Standar Competence Report"
        Me.GboLaporan.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents LblUnchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents LblChecked As System.Windows.Forms.LinkLabel
    Friend WithEvents GboLaporan As System.Windows.Forms.GroupBox
    Friend WithEvents clbJobPosition As System.Windows.Forms.CheckedListBox
End Class
