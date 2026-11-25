<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SPL_Cetak
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
        Me.crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.crvDetail = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.BtnClose = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(789, 558)
        Me.TabControl1.TabIndex = 16
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.crv)
        Me.TabPage1.Controls.Add(Me.BtnClose)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(781, 532)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Main Format"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'crv
        '
        Me.crv.ActiveViewIndex = -1
        Me.crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv.DisplayGroupTree = False
        Me.crv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crv.EnableDrillDown = False
        Me.crv.Location = New System.Drawing.Point(3, 3)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(775, 526)
        Me.crv.TabIndex = 17
        Me.crv.TabStop = False
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.crvDetail)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(781, 532)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Detail Format"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'crvDetail
        '
        Me.crvDetail.ActiveViewIndex = -1
        Me.crvDetail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvDetail.DisplayGroupTree = False
        Me.crvDetail.DisplayStatusBar = False
        Me.crvDetail.DisplayToolbar = False
        Me.crvDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvDetail.EnableDrillDown = False
        Me.crvDetail.Location = New System.Drawing.Point(3, 3)
        Me.crvDetail.Name = "crvDetail"
        Me.crvDetail.SelectionFormula = ""
        Me.crvDetail.ShowCloseButton = False
        Me.crvDetail.ShowGroupTreeButton = False
        Me.crvDetail.ShowRefreshButton = False
        Me.crvDetail.Size = New System.Drawing.Size(775, 526)
        Me.crvDetail.TabIndex = 16
        Me.crvDetail.TabStop = False
        Me.crvDetail.ViewTimeSelectionFormula = ""
        '
        'BtnClose
        '
        Me.BtnClose.Location = New System.Drawing.Point(703, 3)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 23)
        Me.BtnClose.TabIndex = 18
        Me.BtnClose.TabStop = False
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'FrmTrans_SPL_Cetak
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 558)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "FrmTrans_SPL_Cetak"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Overtime Order"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents crvDetail As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents BtnClose As System.Windows.Forms.Button
End Class
