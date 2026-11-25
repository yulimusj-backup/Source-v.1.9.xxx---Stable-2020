<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporan_PR_Balance
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.btnFindProduct = New System.Windows.Forms.Button
        Me.BtnFindRefParty = New System.Windows.Forms.Button
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GboDetail1 = New System.Windows.Forms.GroupBox
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.ChkPeriod = New System.Windows.Forms.CheckBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtRequestBy = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtProductCode = New System.Windows.Forms.TextBox
        Me.TxtPrimaryRefParty = New System.Windows.Forms.TextBox
        Me.TxtPrimaryProduct = New System.Windows.Forms.TextBox
        Me.cboCodeNo = New System.Windows.Forms.ComboBox
        Me.RboRequestBy = New System.Windows.Forms.RadioButton
        Me.RboProduct = New System.Windows.Forms.RadioButton
        Me.TxtDescRefParty = New System.Windows.Forms.TextBox
        Me.TxtDescProduct = New System.Windows.Forms.TextBox
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GboDetail1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(456, 81)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 20)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(456, 59)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(75, 20)
        Me.btnFilter.TabIndex = 2
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'crv
        '
        Me.crv.ActiveViewIndex = -1
        Me.crv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv.DisplayGroupTree = False
        Me.crv.EnableDrillDown = False
        Me.crv.Location = New System.Drawing.Point(16, 128)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(921, 401)
        Me.crv.TabIndex = 9
        Me.crv.TabStop = False
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'btnFindProduct
        '
        Me.btnFindProduct.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindProduct.Location = New System.Drawing.Point(348, 58)
        Me.btnFindProduct.Name = "btnFindProduct"
        Me.btnFindProduct.Size = New System.Drawing.Size(31, 23)
        Me.btnFindProduct.TabIndex = 7
        Me.btnFindProduct.TabStop = False
        Me.btnFindProduct.Text = "..."
        Me.btnFindProduct.UseVisualStyleBackColor = True
        '
        'BtnFindRefParty
        '
        Me.BtnFindRefParty.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindRefParty.Location = New System.Drawing.Point(348, 83)
        Me.BtnFindRefParty.Name = "BtnFindRefParty"
        Me.BtnFindRefParty.Size = New System.Drawing.Size(31, 23)
        Me.BtnFindRefParty.TabIndex = 12
        Me.BtnFindRefParty.TabStop = False
        Me.BtnFindRefParty.Text = "..."
        Me.BtnFindRefParty.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(16, 12)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GboDetail1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRequestBy)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindRefParty)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtProductCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindProduct)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryRefParty)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryProduct)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboCodeNo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.RboRequestBy)
        Me.SplitContainer1.Panel2.Controls.Add(Me.RboProduct)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtDescRefParty)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtDescProduct)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFilter)
        Me.SplitContainer1.Size = New System.Drawing.Size(921, 110)
        Me.SplitContainer1.SplitterDistance = 383
        Me.SplitContainer1.TabIndex = 26
        Me.SplitContainer1.TabStop = False
        '
        'GboDetail1
        '
        Me.GboDetail1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GboDetail1.Controls.Add(Me.TableLayoutPanel1)
        Me.GboDetail1.Controls.Add(Me.ChkPeriod)
        Me.GboDetail1.Controls.Add(Me.Label6)
        Me.GboDetail1.Location = New System.Drawing.Point(7, 3)
        Me.GboDetail1.Name = "GboDetail1"
        Me.GboDetail1.Size = New System.Drawing.Size(372, 49)
        Me.GboDetail1.TabIndex = 0
        Me.GboDetail1.TabStop = False
        Me.GboDetail1.Text = "| |"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.AutoSize = True
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.DtpTo, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.DtpFrom, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(80, 17)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(287, 28)
        Me.TableLayoutPanel1.TabIndex = 32
        '
        'DtpTo
        '
        Me.DtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(159, 3)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(125, 20)
        Me.DtpTo.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(133, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Padding = New System.Windows.Forms.Padding(0, 5, 0, 0)
        Me.Label3.Size = New System.Drawing.Size(20, 18)
        Me.Label3.TabIndex = 30
        Me.Label3.Text = "To"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DtpFrom
        '
        Me.DtpFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFrom.Location = New System.Drawing.Point(3, 3)
        Me.DtpFrom.Name = "DtpFrom"
        Me.DtpFrom.Size = New System.Drawing.Size(124, 20)
        Me.DtpFrom.TabIndex = 0
        '
        'ChkPeriod
        '
        Me.ChkPeriod.AutoSize = True
        Me.ChkPeriod.Location = New System.Drawing.Point(6, 0)
        Me.ChkPeriod.Name = "ChkPeriod"
        Me.ChkPeriod.Size = New System.Drawing.Size(96, 17)
        Me.ChkPeriod.TabIndex = 0
        Me.ChkPeriod.Text = "Filter By Period"
        Me.ChkPeriod.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 26)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(30, 13)
        Me.Label6.TabIndex = 27
        Me.Label6.Text = "From"
        '
        'TxtRequestBy
        '
        Me.TxtRequestBy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRequestBy.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRequestBy.Location = New System.Drawing.Point(90, 82)
        Me.TxtRequestBy.Name = "TxtRequestBy"
        Me.TxtRequestBy.Size = New System.Drawing.Size(252, 20)
        Me.TxtRequestBy.TabIndex = 1
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(10, 85)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(62, 13)
        Me.Label7.TabIndex = 29
        Me.Label7.Text = "Request By"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(10, 63)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(72, 13)
        Me.Label8.TabIndex = 27
        Me.Label8.Text = "Product Code"
        '
        'TxtProductCode
        '
        Me.TxtProductCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtProductCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtProductCode.Location = New System.Drawing.Point(90, 58)
        Me.TxtProductCode.Name = "TxtProductCode"
        Me.TxtProductCode.Size = New System.Drawing.Size(252, 20)
        Me.TxtProductCode.TabIndex = 0
        '
        'TxtPrimaryRefParty
        '
        Me.TxtPrimaryRefParty.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryRefParty.Location = New System.Drawing.Point(54, 78)
        Me.TxtPrimaryRefParty.Name = "TxtPrimaryRefParty"
        Me.TxtPrimaryRefParty.Size = New System.Drawing.Size(33, 20)
        Me.TxtPrimaryRefParty.TabIndex = 32
        Me.TxtPrimaryRefParty.TabStop = False
        Me.TxtPrimaryRefParty.Text = "-1"
        Me.TxtPrimaryRefParty.Visible = False
        '
        'TxtPrimaryProduct
        '
        Me.TxtPrimaryProduct.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryProduct.Location = New System.Drawing.Point(54, 56)
        Me.TxtPrimaryProduct.Name = "TxtPrimaryProduct"
        Me.TxtPrimaryProduct.Size = New System.Drawing.Size(33, 20)
        Me.TxtPrimaryProduct.TabIndex = 31
        Me.TxtPrimaryProduct.TabStop = False
        Me.TxtPrimaryProduct.Text = "-1"
        Me.TxtPrimaryProduct.Visible = False
        '
        'cboCodeNo
        '
        Me.cboCodeNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCodeNo.FormattingEnabled = True
        Me.cboCodeNo.Items.AddRange(New Object() {"Show Code No", "Show Description", "Show Code No & Description"})
        Me.cboCodeNo.Location = New System.Drawing.Point(118, 14)
        Me.cboCodeNo.Margin = New System.Windows.Forms.Padding(2)
        Me.cboCodeNo.Name = "cboCodeNo"
        Me.cboCodeNo.Size = New System.Drawing.Size(138, 21)
        Me.cboCodeNo.TabIndex = 29
        '
        'RboRequestBy
        '
        Me.RboRequestBy.AutoSize = True
        Me.RboRequestBy.Location = New System.Drawing.Point(3, 33)
        Me.RboRequestBy.Name = "RboRequestBy"
        Me.RboRequestBy.Size = New System.Drawing.Size(112, 17)
        Me.RboRequestBy.TabIndex = 1
        Me.RboRequestBy.TabStop = True
        Me.RboRequestBy.Text = "Group By Request"
        Me.RboRequestBy.UseVisualStyleBackColor = True
        '
        'RboProduct
        '
        Me.RboProduct.AutoSize = True
        Me.RboProduct.Location = New System.Drawing.Point(3, 15)
        Me.RboProduct.Name = "RboProduct"
        Me.RboProduct.Size = New System.Drawing.Size(109, 17)
        Me.RboProduct.TabIndex = 0
        Me.RboProduct.TabStop = True
        Me.RboProduct.Text = "Group By Product"
        Me.RboProduct.UseVisualStyleBackColor = True
        '
        'TxtDescRefParty
        '
        Me.TxtDescRefParty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescRefParty.Location = New System.Drawing.Point(3, 83)
        Me.TxtDescRefParty.Name = "TxtDescRefParty"
        Me.TxtDescRefParty.ReadOnly = True
        Me.TxtDescRefParty.Size = New System.Drawing.Size(447, 20)
        Me.TxtDescRefParty.TabIndex = 28
        Me.TxtDescRefParty.TabStop = False
        '
        'TxtDescProduct
        '
        Me.TxtDescProduct.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescProduct.Location = New System.Drawing.Point(3, 59)
        Me.TxtDescProduct.Name = "TxtDescProduct"
        Me.TxtDescProduct.ReadOnly = True
        Me.TxtDescProduct.Size = New System.Drawing.Size(447, 20)
        Me.TxtDescProduct.TabIndex = 27
        Me.TxtDescProduct.TabStop = False
        '
        'FrmLaporan_PR_Balance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(953, 538)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.crv)
        Me.MinimumSize = New System.Drawing.Size(969, 576)
        Me.Name = "FrmLaporan_PR_Balance"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "PR Balance"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.GboDetail1.ResumeLayout(False)
        Me.GboDetail1.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnFindProduct As System.Windows.Forms.Button
    Friend WithEvents BtnFindRefParty As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtRequestBy As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtProductCode As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtDescRefParty As System.Windows.Forms.TextBox
    Friend WithEvents TxtDescProduct As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryRefParty As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryProduct As System.Windows.Forms.TextBox
    Friend WithEvents GboDetail1 As System.Windows.Forms.GroupBox
    Friend WithEvents RboRequestBy As System.Windows.Forms.RadioButton
    Friend WithEvents RboProduct As System.Windows.Forms.RadioButton
    Friend WithEvents ChkPeriod As System.Windows.Forms.CheckBox
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents cboCodeNo As System.Windows.Forms.ComboBox
End Class
