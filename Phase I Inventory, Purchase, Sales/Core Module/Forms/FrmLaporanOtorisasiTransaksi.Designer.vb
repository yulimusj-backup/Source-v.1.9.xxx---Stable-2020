<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanOtorisasiTransaksi
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
        Me.TxtDocumentNo = New System.Windows.Forms.TextBox
        Me.clbTipeOtorisasi = New System.Windows.Forms.CheckedListBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtRequestBy = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtApprovedBy = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.llbCekAll = New System.Windows.Forms.LinkLabel
        Me.llbUncekAll = New System.Windows.Forms.LinkLabel
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.TxtDocumentNo)
        Me.GroupBox1.Controls.Add(Me.clbTipeOtorisasi)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.TxtRequestBy)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.TxtApprovedBy)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.llbCekAll)
        Me.GroupBox1.Controls.Add(Me.llbUncekAll)
        Me.GroupBox1.Controls.Add(Me.btnClose)
        Me.GroupBox1.Controls.Add(Me.btnShow)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.dtpTo)
        Me.GroupBox1.Controls.Add(Me.dtpFrom)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(968, 146)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filter"
        '
        'TxtDocumentNo
        '
        Me.TxtDocumentNo.Location = New System.Drawing.Point(358, 72)
        Me.TxtDocumentNo.Name = "TxtDocumentNo"
        Me.TxtDocumentNo.Size = New System.Drawing.Size(221, 22)
        Me.TxtDocumentNo.TabIndex = 4
        '
        'clbTipeOtorisasi
        '
        Me.clbTipeOtorisasi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.clbTipeOtorisasi.CheckOnClick = True
        Me.clbTipeOtorisasi.ColumnWidth = 175
        Me.clbTipeOtorisasi.FormattingEnabled = True
        Me.clbTipeOtorisasi.Location = New System.Drawing.Point(585, 16)
        Me.clbTipeOtorisasi.MultiColumn = True
        Me.clbTipeOtorisasi.Name = "clbTipeOtorisasi"
        Me.clbTipeOtorisasi.Size = New System.Drawing.Size(377, 106)
        Me.clbTipeOtorisasi.TabIndex = 5
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(238, 77)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(94, 17)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Document No"
        '
        'TxtRequestBy
        '
        Me.TxtRequestBy.Location = New System.Drawing.Point(358, 44)
        Me.TxtRequestBy.Name = "TxtRequestBy"
        Me.TxtRequestBy.Size = New System.Drawing.Size(221, 22)
        Me.TxtRequestBy.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(238, 49)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(81, 17)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Request By"
        '
        'TxtApprovedBy
        '
        Me.TxtApprovedBy.Location = New System.Drawing.Point(358, 16)
        Me.TxtApprovedBy.Name = "TxtApprovedBy"
        Me.TxtApprovedBy.Size = New System.Drawing.Size(221, 22)
        Me.TxtApprovedBy.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(238, 21)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(89, 17)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Approved By"
        '
        'llbCekAll
        '
        Me.llbCekAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbCekAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCekAll.AutoSize = True
        Me.llbCekAll.Location = New System.Drawing.Point(808, 132)
        Me.llbCekAll.Name = "llbCekAll"
        Me.llbCekAll.Size = New System.Drawing.Size(66, 17)
        Me.llbCekAll.TabIndex = 6
        Me.llbCekAll.TabStop = True
        Me.llbCekAll.Text = "Check All"
        Me.llbCekAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'llbUncekAll
        '
        Me.llbUncekAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbUncekAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUncekAll.AutoSize = True
        Me.llbUncekAll.Location = New System.Drawing.Point(880, 132)
        Me.llbUncekAll.Name = "llbUncekAll"
        Me.llbUncekAll.Size = New System.Drawing.Size(82, 17)
        Me.llbUncekAll.TabIndex = 7
        Me.llbUncekAll.TabStop = True
        Me.llbUncekAll.Text = "Uncheck All"
        Me.llbUncekAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(137, 81)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(79, 23)
        Me.btnClose.TabIndex = 9
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(52, 81)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(79, 23)
        Me.btnShow.TabIndex = 8
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 54)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(25, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "To"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "From"
        '
        'dtpTo
        '
        Me.dtpTo.Location = New System.Drawing.Point(52, 49)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(180, 22)
        Me.dtpTo.TabIndex = 1
        '
        'dtpFrom
        '
        Me.dtpFrom.Location = New System.Drawing.Point(52, 21)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(180, 22)
        Me.dtpFrom.TabIndex = 0
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.Location = New System.Drawing.Point(12, 165)
        Me.crv1.Margin = New System.Windows.Forms.Padding(4)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(967, 463)
        Me.crv1.TabIndex = 1
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(326, 269)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(67, 22)
        Me.TextBox1.TabIndex = 13
        Me.TextBox1.TabStop = False
        Me.TextBox1.Visible = False
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(463, 305)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(67, 22)
        Me.TxtReadOnly.TabIndex = 14
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'FrmLaporanOtorisasiTransaksi
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.crv1)
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "FrmLaporanOtorisasiTransaksi"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Approval Report"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents clbTipeOtorisasi As System.Windows.Forms.CheckedListBox
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents llbUncekAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCekAll As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtDocumentNo As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtRequestBy As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtApprovedBy As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
End Class
