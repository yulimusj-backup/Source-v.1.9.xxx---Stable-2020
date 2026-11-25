<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanOvertimeAnalysis
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
        Me.LlbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.LlbCheckAll = New System.Windows.Forms.LinkLabel
        Me.clbDept = New System.Windows.Forms.CheckedListBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.crvLaporanAnalysisOvertime = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LlbUncheckAll)
        Me.GroupBox1.Controls.Add(Me.LlbCheckAll)
        Me.GroupBox1.Controls.Add(Me.clbDept)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 7)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(363, 87)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Department"
        '
        'LlbUncheckAll
        '
        Me.LlbUncheckAll.AutoSize = True
        Me.LlbUncheckAll.Location = New System.Drawing.Point(292, 0)
        Me.LlbUncheckAll.Name = "LlbUncheckAll"
        Me.LlbUncheckAll.Size = New System.Drawing.Size(65, 13)
        Me.LlbUncheckAll.TabIndex = 66
        Me.LlbUncheckAll.TabStop = True
        Me.LlbUncheckAll.Text = "Uncheck All"
        '
        'LlbCheckAll
        '
        Me.LlbCheckAll.AutoSize = True
        Me.LlbCheckAll.Location = New System.Drawing.Point(234, 0)
        Me.LlbCheckAll.Name = "LlbCheckAll"
        Me.LlbCheckAll.Size = New System.Drawing.Size(52, 13)
        Me.LlbCheckAll.TabIndex = 65
        Me.LlbCheckAll.TabStop = True
        Me.LlbCheckAll.Text = "Check All"
        '
        'clbDept
        '
        Me.clbDept.FormattingEnabled = True
        Me.clbDept.Location = New System.Drawing.Point(6, 17)
        Me.clbDept.Name = "clbDept"
        Me.clbDept.Size = New System.Drawing.Size(348, 64)
        Me.clbDept.TabIndex = 2
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(548, 39)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(548, 11)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(75, 23)
        Me.btnShow.TabIndex = 3
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "MMMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(422, 39)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.ShowUpDown = True
        Me.dtpTo.Size = New System.Drawing.Size(120, 20)
        Me.dtpTo.TabIndex = 1
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "MMMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(422, 12)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.ShowUpDown = True
        Me.dtpFrom.Size = New System.Drawing.Size(120, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(383, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "To"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(383, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From"
        '
        'crvLaporanAnalysisOvertime
        '
        Me.crvLaporanAnalysisOvertime.ActiveViewIndex = -1
        Me.crvLaporanAnalysisOvertime.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvLaporanAnalysisOvertime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvLaporanAnalysisOvertime.DisplayGroupTree = False
        Me.crvLaporanAnalysisOvertime.Location = New System.Drawing.Point(12, 105)
        Me.crvLaporanAnalysisOvertime.Name = "crvLaporanAnalysisOvertime"
        Me.crvLaporanAnalysisOvertime.SelectionFormula = ""
        Me.crvLaporanAnalysisOvertime.ShowCloseButton = False
        Me.crvLaporanAnalysisOvertime.ShowGroupTreeButton = False
        Me.crvLaporanAnalysisOvertime.ShowRefreshButton = False
        Me.crvLaporanAnalysisOvertime.Size = New System.Drawing.Size(768, 449)
        Me.crvLaporanAnalysisOvertime.TabIndex = 1
        Me.crvLaporanAnalysisOvertime.TabStop = False
        Me.crvLaporanAnalysisOvertime.ViewTimeSelectionFormula = ""
        '
        'FrmLaporanOvertimeAnalysis
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(792, 566)
        Me.Controls.Add(Me.crvLaporanAnalysisOvertime)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnShow)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpFrom)
        Me.Name = "FrmLaporanOvertimeAnalysis"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Overtime Analysis Report"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents crvLaporanAnalysisOvertime As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents clbDept As System.Windows.Forms.CheckedListBox
    Friend WithEvents LlbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbCheckAll As System.Windows.Forms.LinkLabel
End Class
