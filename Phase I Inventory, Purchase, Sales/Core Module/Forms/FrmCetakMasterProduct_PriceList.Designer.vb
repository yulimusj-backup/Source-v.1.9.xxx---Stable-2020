<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCetakMasterProduct_PriceList
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
        Me.CrvLaporanMasterPriceList = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'CrvLaporanMasterPriceList
        '
        Me.CrvLaporanMasterPriceList.ActiveViewIndex = -1
        Me.CrvLaporanMasterPriceList.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CrvLaporanMasterPriceList.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrvLaporanMasterPriceList.DisplayGroupTree = False
        Me.CrvLaporanMasterPriceList.Location = New System.Drawing.Point(16, 15)
        Me.CrvLaporanMasterPriceList.Margin = New System.Windows.Forms.Padding(4)
        Me.CrvLaporanMasterPriceList.Name = "CrvLaporanMasterPriceList"
        Me.CrvLaporanMasterPriceList.SelectionFormula = ""
        Me.CrvLaporanMasterPriceList.ShowCloseButton = False
        Me.CrvLaporanMasterPriceList.ShowGroupTreeButton = False
        Me.CrvLaporanMasterPriceList.ShowRefreshButton = False
        Me.CrvLaporanMasterPriceList.Size = New System.Drawing.Size(959, 603)
        Me.CrvLaporanMasterPriceList.TabIndex = 16
        Me.CrvLaporanMasterPriceList.ViewTimeSelectionFormula = ""
        '
        'FrmCetakMasterProduct_PriceList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.CrvLaporanMasterPriceList)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "FrmCetakMasterProduct_PriceList"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Price List"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CrvLaporanMasterPriceList As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
