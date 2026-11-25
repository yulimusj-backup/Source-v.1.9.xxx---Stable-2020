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
        Me.crvSPL = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'crvSPL
        '
        Me.crvSPL.ActiveViewIndex = -1
        Me.crvSPL.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvSPL.DisplayGroupTree = False
        Me.crvSPL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvSPL.EnableDrillDown = False
        Me.crvSPL.Location = New System.Drawing.Point(0, 0)
        Me.crvSPL.Name = "crvSPL"
        Me.crvSPL.SelectionFormula = ""
        Me.crvSPL.ShowCloseButton = False
        Me.crvSPL.ShowGroupTreeButton = False
        Me.crvSPL.ShowRefreshButton = False
        Me.crvSPL.Size = New System.Drawing.Size(784, 462)
        Me.crvSPL.TabIndex = 0
        Me.crvSPL.ViewTimeSelectionFormula = ""
        '
        'FrmTrans_SPL_Cetak
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 462)
        Me.Controls.Add(Me.crvSPL)
        Me.Name = "FrmTrans_SPL_Cetak"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Overtime Order Print"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents crvSPL As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
