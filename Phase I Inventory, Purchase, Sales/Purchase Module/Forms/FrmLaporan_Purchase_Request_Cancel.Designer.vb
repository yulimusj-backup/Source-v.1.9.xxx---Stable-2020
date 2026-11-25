<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporan_Purchase_Request_Cancel
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryRefParty = New System.Windows.Forms.TextBox
        Me.TxtPrimaryProduct = New System.Windows.Forms.TextBox
        Me.TxtRequestBy = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtProductCode = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtPrimaryPR = New System.Windows.Forms.TextBox
        Me.TxtDescRefParty = New System.Windows.Forms.TextBox
        Me.txtNoPR = New System.Windows.Forms.TextBox
        Me.TxtDescProduct = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.BtnFindPR = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(320, 59)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 20)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(320, 35)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(75, 20)
        Me.btnFilter.TabIndex = 1
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
        Me.crv.Location = New System.Drawing.Point(12, 101)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(803, 429)
        Me.crv.TabIndex = 24
        Me.crv.TabStop = False
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'btnFindProduct
        '
        Me.btnFindProduct.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindProduct.Location = New System.Drawing.Point(359, 29)
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
        Me.BtnFindRefParty.Location = New System.Drawing.Point(359, 53)
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
        Me.SplitContainer1.Location = New System.Drawing.Point(12, 12)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TableLayoutPanel1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryRefParty)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryProduct)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRequestBy)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindRefParty)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtProductCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindProduct)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryPR)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtDescRefParty)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtNoPR)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtDescProduct)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnFindPR)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFilter)
        Me.SplitContainer1.Size = New System.Drawing.Size(803, 83)
        Me.SplitContainer1.SplitterDistance = 401
        Me.SplitContainer1.TabIndex = 26
        Me.SplitContainer1.TabStop = False
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(85, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(308, 26)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'DtpTo
        '
        Me.DtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(170, 3)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(135, 20)
        Me.DtpTo.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(144, 0)
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
        Me.DtpFrom.Size = New System.Drawing.Size(135, 20)
        Me.DtpFrom.TabIndex = 0
        '
        'TxtPrimaryRefParty
        '
        Me.TxtPrimaryRefParty.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryRefParty.Location = New System.Drawing.Point(54, 49)
        Me.TxtPrimaryRefParty.Name = "TxtPrimaryRefParty"
        Me.TxtPrimaryRefParty.Size = New System.Drawing.Size(33, 20)
        Me.TxtPrimaryRefParty.TabIndex = 32
        Me.TxtPrimaryRefParty.Text = "-1"
        Me.TxtPrimaryRefParty.Visible = False
        '
        'TxtPrimaryProduct
        '
        Me.TxtPrimaryProduct.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryProduct.Location = New System.Drawing.Point(54, 25)
        Me.TxtPrimaryProduct.Name = "TxtPrimaryProduct"
        Me.TxtPrimaryProduct.Size = New System.Drawing.Size(33, 20)
        Me.TxtPrimaryProduct.TabIndex = 31
        Me.TxtPrimaryProduct.Text = "-1"
        Me.TxtPrimaryProduct.Visible = False
        '
        'TxtRequestBy
        '
        Me.TxtRequestBy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRequestBy.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRequestBy.Location = New System.Drawing.Point(88, 53)
        Me.TxtRequestBy.Name = "TxtRequestBy"
        Me.TxtRequestBy.Size = New System.Drawing.Size(265, 20)
        Me.TxtRequestBy.TabIndex = 1
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(4, 56)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(62, 13)
        Me.Label7.TabIndex = 29
        Me.Label7.Text = "Request By"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(4, 32)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(44, 13)
        Me.Label8.TabIndex = 27
        Me.Label8.Text = "Product"
        '
        'TxtProductCode
        '
        Me.TxtProductCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtProductCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtProductCode.Location = New System.Drawing.Point(88, 29)
        Me.TxtProductCode.Name = "TxtProductCode"
        Me.TxtProductCode.Size = New System.Drawing.Size(265, 20)
        Me.TxtProductCode.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(4, 7)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(30, 13)
        Me.Label6.TabIndex = 27
        Me.Label6.Text = "From"
        '
        'TxtPrimaryPR
        '
        Me.TxtPrimaryPR.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryPR.Location = New System.Drawing.Point(54, 2)
        Me.TxtPrimaryPR.Name = "TxtPrimaryPR"
        Me.TxtPrimaryPR.Size = New System.Drawing.Size(33, 20)
        Me.TxtPrimaryPR.TabIndex = 36
        Me.TxtPrimaryPR.Text = "-1"
        Me.TxtPrimaryPR.Visible = False
        '
        'TxtDescRefParty
        '
        Me.TxtDescRefParty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescRefParty.Location = New System.Drawing.Point(7, 61)
        Me.TxtDescRefParty.Name = "TxtDescRefParty"
        Me.TxtDescRefParty.ReadOnly = True
        Me.TxtDescRefParty.Size = New System.Drawing.Size(307, 20)
        Me.TxtDescRefParty.TabIndex = 28
        Me.TxtDescRefParty.TabStop = False
        '
        'txtNoPR
        '
        Me.txtNoPR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNoPR.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtNoPR.Location = New System.Drawing.Point(92, 6)
        Me.txtNoPR.Name = "txtNoPR"
        Me.txtNoPR.Size = New System.Drawing.Size(222, 20)
        Me.txtNoPR.TabIndex = 0
        '
        'TxtDescProduct
        '
        Me.TxtDescProduct.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescProduct.Location = New System.Drawing.Point(7, 34)
        Me.TxtDescProduct.Name = "TxtDescProduct"
        Me.TxtDescProduct.ReadOnly = True
        Me.TxtDescProduct.Size = New System.Drawing.Size(307, 20)
        Me.TxtDescProduct.TabIndex = 27
        Me.TxtDescProduct.TabStop = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(4, 9)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(62, 13)
        Me.Label9.TabIndex = 34
        Me.Label9.Text = "PR Number"
        '
        'BtnFindPR
        '
        Me.BtnFindPR.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindPR.Location = New System.Drawing.Point(320, 4)
        Me.BtnFindPR.Name = "BtnFindPR"
        Me.BtnFindPR.Size = New System.Drawing.Size(31, 23)
        Me.BtnFindPR.TabIndex = 33
        Me.BtnFindPR.TabStop = False
        Me.BtnFindPR.Text = "..."
        Me.BtnFindPR.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox1.Location = New System.Drawing.Point(54, 49)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(33, 20)
        Me.TextBox1.TabIndex = 32
        Me.TextBox1.Text = "-1"
        '
        'TextBox2
        '
        Me.TextBox2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox2.Location = New System.Drawing.Point(54, 25)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(33, 20)
        Me.TextBox2.TabIndex = 31
        Me.TextBox2.Text = "-1"
        '
        'TextBox3
        '
        Me.TextBox3.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox3.Location = New System.Drawing.Point(92, 53)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(329, 20)
        Me.TextBox3.TabIndex = 30
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "dd MMM yyyy"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(288, 3)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(170, 20)
        Me.DateTimePicker1.TabIndex = 30
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 29
        Me.Label1.Text = "Request By"
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.CustomFormat = "dd MMM yyyy"
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker2.Location = New System.Drawing.Point(92, 3)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(164, 20)
        Me.DateTimePicker2.TabIndex = 28
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 13)
        Me.Label2.TabIndex = 27
        Me.Label2.Text = "Product"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(427, 53)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(31, 23)
        Me.Button1.TabIndex = 12
        Me.Button1.TabStop = False
        Me.Button1.Text = "..."
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(262, 7)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 29
        Me.Label4.Text = "To"
        '
        'TextBox4
        '
        Me.TextBox4.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox4.Location = New System.Drawing.Point(92, 29)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(329, 20)
        Me.TextBox4.TabIndex = 28
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(4, 7)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 13)
        Me.Label5.TabIndex = 27
        Me.Label5.Text = "From"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(427, 29)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(31, 23)
        Me.Button2.TabIndex = 7
        Me.Button2.TabStop = False
        Me.Button2.Text = "..."
        Me.Button2.UseVisualStyleBackColor = True
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(7, 56)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.ReadOnly = True
        Me.TextBox5.Size = New System.Drawing.Size(402, 20)
        Me.TextBox5.TabIndex = 28
        Me.TextBox5.TabStop = False
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(92, 6)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(41, 17)
        Me.CheckBox1.TabIndex = 25
        Me.CheckBox1.Text = "OB"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(9, 6)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(77, 17)
        Me.CheckBox2.TabIndex = 18
        Me.CheckBox2.Text = "All Product"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(7, 32)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.ReadOnly = True
        Me.TextBox6.Size = New System.Drawing.Size(402, 20)
        Me.TextBox6.TabIndex = 27
        Me.TextBox6.TabStop = False
        '
        'Button3
        '
        Me.Button3.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button3.Location = New System.Drawing.Point(412, 54)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 20)
        Me.Button3.TabIndex = 23
        Me.Button3.Text = "Close"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(412, 30)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 20)
        Me.Button4.TabIndex = 22
        Me.Button4.Text = "Filter"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'FrmLaporan_Purchase_Request_Cancel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(827, 538)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.crv)
        Me.MinimumSize = New System.Drawing.Size(843, 576)
        Me.Name = "FrmLaporan_Purchase_Request_Cancel"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Request Cancel Report"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
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
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryPR As System.Windows.Forms.TextBox
    Friend WithEvents txtNoPR As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents BtnFindPR As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
End Class
