<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterKompetensi
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtDefPrec = New System.Windows.Forms.TextBox
        Me.LblDecimal = New System.Windows.Forms.Label
        Me.cboDataType = New System.Windows.Forms.ComboBox
        Me.txtPrimaryStockIn = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtDataCode = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtDefValueDigit = New System.Windows.Forms.TextBox
        Me.cboDefValue = New System.Windows.Forms.ComboBox
        Me.DtpDefValue = New System.Windows.Forms.DateTimePicker
        Me.TxtDefValueText = New System.Windows.Forms.TextBox
        Me.cboAlfabetic = New System.Windows.Forms.ComboBox
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tspBar1 = New System.Windows.Forms.ToolStripProgressBar
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.gboDetail.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 103)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 39
        Me.Label1.Text = "Remarks"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(237, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.txtDefPrec)
        Me.gboDetail.Controls.Add(Me.LblDecimal)
        Me.gboDetail.Controls.Add(Me.cboDataType)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.txtPrimaryStockIn)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.txtDescription)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.txtDataCode)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtDefValueDigit)
        Me.gboDetail.Controls.Add(Me.cboDefValue)
        Me.gboDetail.Controls.Add(Me.DtpDefValue)
        Me.gboDetail.Controls.Add(Me.TxtDefValueText)
        Me.gboDetail.Controls.Add(Me.cboAlfabetic)
        Me.gboDetail.Controls.Add(Me.txtRemarks)
        Me.gboDetail.Location = New System.Drawing.Point(12, 6)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(466, 172)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 76)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(71, 13)
        Me.Label6.TabIndex = 43
        Me.Label6.Text = "Default Value"
        '
        'txtDefPrec
        '
        Me.txtDefPrec.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDefPrec.Location = New System.Drawing.Point(337, 72)
        Me.txtDefPrec.MaxLength = 15
        Me.txtDefPrec.Name = "txtDefPrec"
        Me.txtDefPrec.Size = New System.Drawing.Size(123, 20)
        Me.txtDefPrec.TabIndex = 4
        Me.txtDefPrec.Text = "0"
        Me.txtDefPrec.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblDecimal
        '
        Me.LblDecimal.AutoSize = True
        Me.LblDecimal.Location = New System.Drawing.Point(286, 76)
        Me.LblDecimal.Name = "LblDecimal"
        Me.LblDecimal.Size = New System.Drawing.Size(45, 13)
        Me.LblDecimal.TabIndex = 41
        Me.LblDecimal.Text = "Decimal"
        '
        'cboDataType
        '
        Me.cboDataType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDataType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDataType.FormattingEnabled = True
        Me.cboDataType.Items.AddRange(New Object() {"Number", "Date", "Text", "True/False", "Alphabetic"})
        Me.cboDataType.Location = New System.Drawing.Point(349, 19)
        Me.cboDataType.Name = "cboDataType"
        Me.cboDataType.Size = New System.Drawing.Size(111, 21)
        Me.cboDataType.TabIndex = 1
        '
        'txtPrimaryStockIn
        '
        Me.txtPrimaryStockIn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryStockIn.Location = New System.Drawing.Point(79, 19)
        Me.txtPrimaryStockIn.Name = "txtPrimaryStockIn"
        Me.txtPrimaryStockIn.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryStockIn.TabIndex = 36
        Me.txtPrimaryStockIn.TabStop = False
        Me.txtPrimaryStockIn.Text = "-1"
        Me.txtPrimaryStockIn.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 13)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Description"
        '
        'txtDescription
        '
        Me.txtDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescription.Location = New System.Drawing.Point(122, 45)
        Me.txtDescription.MaxLength = 250
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(338, 20)
        Me.txtDescription.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(57, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Data Type"
        Me.Label4.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 13)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Code"
        '
        'txtDataCode
        '
        Me.txtDataCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDataCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDataCode.Location = New System.Drawing.Point(122, 19)
        Me.txtDataCode.MaxLength = 45
        Me.txtDataCode.Name = "txtDataCode"
        Me.txtDataCode.Size = New System.Drawing.Size(215, 20)
        Me.txtDataCode.TabIndex = 0
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(239, 140)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(88, 26)
        Me.btnCancel.TabIndex = 11
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(139, 140)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(88, 26)
        Me.btnOK.TabIndex = 10
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtDefValueDigit
        '
        Me.TxtDefValueDigit.Location = New System.Drawing.Point(122, 72)
        Me.TxtDefValueDigit.MaxLength = 15
        Me.TxtDefValueDigit.Name = "TxtDefValueDigit"
        Me.TxtDefValueDigit.Size = New System.Drawing.Size(152, 20)
        Me.TxtDefValueDigit.TabIndex = 3
        Me.TxtDefValueDigit.Text = "0"
        Me.TxtDefValueDigit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboDefValue
        '
        Me.cboDefValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDefValue.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDefValue.FormattingEnabled = True
        Me.cboDefValue.Items.AddRange(New Object() {"No", "Yes"})
        Me.cboDefValue.Location = New System.Drawing.Point(122, 72)
        Me.cboDefValue.Name = "cboDefValue"
        Me.cboDefValue.Size = New System.Drawing.Size(338, 21)
        Me.cboDefValue.TabIndex = 5
        '
        'DtpDefValue
        '
        Me.DtpDefValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpDefValue.Location = New System.Drawing.Point(122, 72)
        Me.DtpDefValue.Name = "DtpDefValue"
        Me.DtpDefValue.Size = New System.Drawing.Size(338, 20)
        Me.DtpDefValue.TabIndex = 7
        '
        'TxtDefValueText
        '
        Me.TxtDefValueText.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDefValueText.Location = New System.Drawing.Point(122, 72)
        Me.TxtDefValueText.MaxLength = 250
        Me.TxtDefValueText.Name = "TxtDefValueText"
        Me.TxtDefValueText.Size = New System.Drawing.Size(338, 20)
        Me.TxtDefValueText.TabIndex = 8
        '
        'cboAlfabetic
        '
        Me.cboAlfabetic.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboAlfabetic.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAlfabetic.FormattingEnabled = True
        Me.cboAlfabetic.Items.AddRange(New Object() {"A", "B", "C", "D", "E", "F"})
        Me.cboAlfabetic.Location = New System.Drawing.Point(122, 72)
        Me.cboAlfabetic.Name = "cboAlfabetic"
        Me.cboAlfabetic.Size = New System.Drawing.Size(338, 21)
        Me.cboAlfabetic.TabIndex = 6
        '
        'txtRemarks
        '
        Me.txtRemarks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRemarks.Location = New System.Drawing.Point(122, 99)
        Me.txtRemarks.MaxLength = 250
        Me.txtRemarks.Multiline = True
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(338, 35)
        Me.txtRemarks.TabIndex = 9
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 215)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(490, 22)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 24
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(237, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(237, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(237, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate, Me.tspBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 237)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(490, 22)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 23
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tspBar1
        '
        Me.tspBar1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tspBar1.Name = "tspBar1"
        Me.tspBar1.Size = New System.Drawing.Size(100, 16)
        Me.tspBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.tspBar1.Visible = False
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(200, 184)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(88, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(104, 184)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(88, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(8, 184)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(88, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(392, 184)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(88, 26)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(296, 184)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(88, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'FrmMasterKompetensi
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(490, 259)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(506, 297)
        Me.Name = "FrmMasterKompetensi"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Competence Data"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryStockIn As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDataCode As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tspBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents cboDataType As System.Windows.Forms.ComboBox
    Friend WithEvents txtDefPrec As System.Windows.Forms.TextBox
    Friend WithEvents LblDecimal As System.Windows.Forms.Label
    Friend WithEvents DtpDefValue As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtDefValueText As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboDefValue As System.Windows.Forms.ComboBox
    Friend WithEvents cboAlfabetic As System.Windows.Forms.ComboBox
    Friend WithEvents TxtDefValueDigit As System.Windows.Forms.TextBox
End Class
