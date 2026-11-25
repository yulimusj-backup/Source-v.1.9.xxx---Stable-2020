<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanAtt
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
        Me.GboData = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.CboFomat = New System.Windows.Forms.ComboBox
        Me.BtnClose = New System.Windows.Forms.Button
        Me.LblSetting = New System.Windows.Forms.LinkLabel
        Me.BtnShow = New System.Windows.Forms.Button
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.CrvReport = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.GboData.SuspendLayout()
        Me.SuspendLayout()
        '
        'GboData
        '
        Me.GboData.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GboData.Controls.Add(Me.Label3)
        Me.GboData.Controls.Add(Me.CboFomat)
        Me.GboData.Controls.Add(Me.BtnClose)
        Me.GboData.Controls.Add(Me.LblSetting)
        Me.GboData.Controls.Add(Me.BtnShow)
        Me.GboData.Controls.Add(Me.dtpFrom)
        Me.GboData.Controls.Add(Me.Label1)
        Me.GboData.Controls.Add(Me.dtpTo)
        Me.GboData.Controls.Add(Me.Label2)
        Me.GboData.Location = New System.Drawing.Point(12, 1)
        Me.GboData.Name = "GboData"
        Me.GboData.Size = New System.Drawing.Size(960, 126)
        Me.GboData.TabIndex = 0
        Me.GboData.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 18)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 323
        Me.Label3.Text = "Format"
        '
        'CboFomat
        '
        Me.CboFomat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFomat.FormattingEnabled = True
        Me.CboFomat.Items.AddRange(New Object() {"Recap", "Detail"})
        Me.CboFomat.Location = New System.Drawing.Point(67, 15)
        Me.CboFomat.Name = "CboFomat"
        Me.CboFomat.Size = New System.Drawing.Size(153, 21)
        Me.CboFomat.TabIndex = 322
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(226, 42)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(54, 19)
        Me.BtnClose.TabIndex = 4
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'LblSetting
        '
        Me.LblSetting.AutoSize = True
        Me.LblSetting.Location = New System.Drawing.Point(11, 98)
        Me.LblSetting.Name = "LblSetting"
        Me.LblSetting.Size = New System.Drawing.Size(71, 13)
        Me.LblSetting.TabIndex = 2
        Me.LblSetting.TabStop = True
        Me.LblSetting.Text = "Leave setting"
        '
        'BtnShow
        '
        Me.BtnShow.Location = New System.Drawing.Point(226, 15)
        Me.BtnShow.Name = "BtnShow"
        Me.BtnShow.Size = New System.Drawing.Size(54, 19)
        Me.BtnShow.TabIndex = 3
        Me.BtnShow.Text = "Show"
        Me.BtnShow.UseVisualStyleBackColor = True
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(67, 42)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(153, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 45)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 320
        Me.Label1.Text = "From"
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(67, 68)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(153, 20)
        Me.dtpTo.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 74)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 321
        Me.Label2.Text = "To"
        '
        'CrvReport
        '
        Me.CrvReport.ActiveViewIndex = -1
        Me.CrvReport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CrvReport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrvReport.DisplayGroupTree = False
        Me.CrvReport.EnableDrillDown = False
        Me.CrvReport.Location = New System.Drawing.Point(12, 133)
        Me.CrvReport.Name = "CrvReport"
        Me.CrvReport.SelectionFormula = ""
        Me.CrvReport.ShowCloseButton = False
        Me.CrvReport.ShowGotoPageButton = False
        Me.CrvReport.ShowGroupTreeButton = False
        Me.CrvReport.ShowRefreshButton = False
        Me.CrvReport.Size = New System.Drawing.Size(960, 496)
        Me.CrvReport.TabIndex = 1
        Me.CrvReport.ViewTimeSelectionFormula = ""
        '
        'FrmLaporanAtt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(984, 642)
        Me.Controls.Add(Me.CrvReport)
        Me.Controls.Add(Me.GboData)
        Me.MinimumSize = New System.Drawing.Size(1000, 680)
        Me.Name = "FrmLaporanAtt"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Attendance Summary"
        Me.GboData.ResumeLayout(False)
        Me.GboData.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GboData As System.Windows.Forms.GroupBox
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LblSetting As System.Windows.Forms.LinkLabel
    Friend WithEvents BtnShow As System.Windows.Forms.Button
    Friend WithEvents CrvReport As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CboFomat As System.Windows.Forms.ComboBox
End Class
