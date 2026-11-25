<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterCompanyData
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtCompanyName = New System.Windows.Forms.TextBox
        Me.TxtAddLine1 = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtAddLine2 = New System.Windows.Forms.TextBox
        Me.TxtAddLine3 = New System.Windows.Forms.TextBox
        Me.TxtAddLine4 = New System.Windows.Forms.TextBox
        Me.TxtAddLine5 = New System.Windows.Forms.TextBox
        Me.TxtTaxAddLine2 = New System.Windows.Forms.TextBox
        Me.TxtTaxAddLine1 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtNPWPNumber = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtCity = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dgvAddress = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAddressCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Line1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Line2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Line3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Line4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Line5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarksAdd = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAddressIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.TxtPrimaryCompany = New System.Windows.Forms.TextBox
        Me.txtReadOnly = New System.Windows.Forms.TextBox
        Me.dtp_PKP_Reg_Date = New System.Windows.Forms.DateTimePicker
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvAddress, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Company Name"
        '
        'TxtCompanyName
        '
        Me.TxtCompanyName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCompanyName.Location = New System.Drawing.Point(131, 12)
        Me.TxtCompanyName.MaxLength = 50
        Me.TxtCompanyName.Name = "TxtCompanyName"
        Me.TxtCompanyName.Size = New System.Drawing.Size(336, 20)
        Me.TxtCompanyName.TabIndex = 0
        '
        'TxtAddLine1
        '
        Me.TxtAddLine1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAddLine1.Location = New System.Drawing.Point(131, 38)
        Me.TxtAddLine1.MaxLength = 50
        Me.TxtAddLine1.Name = "TxtAddLine1"
        Me.TxtAddLine1.Size = New System.Drawing.Size(336, 20)
        Me.TxtAddLine1.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Address"
        '
        'TxtAddLine2
        '
        Me.TxtAddLine2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAddLine2.Location = New System.Drawing.Point(131, 64)
        Me.TxtAddLine2.MaxLength = 50
        Me.TxtAddLine2.Name = "TxtAddLine2"
        Me.TxtAddLine2.Size = New System.Drawing.Size(336, 20)
        Me.TxtAddLine2.TabIndex = 2
        '
        'TxtAddLine3
        '
        Me.TxtAddLine3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAddLine3.Location = New System.Drawing.Point(131, 90)
        Me.TxtAddLine3.MaxLength = 50
        Me.TxtAddLine3.Name = "TxtAddLine3"
        Me.TxtAddLine3.Size = New System.Drawing.Size(336, 20)
        Me.TxtAddLine3.TabIndex = 3
        '
        'TxtAddLine4
        '
        Me.TxtAddLine4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAddLine4.Location = New System.Drawing.Point(131, 116)
        Me.TxtAddLine4.MaxLength = 50
        Me.TxtAddLine4.Name = "TxtAddLine4"
        Me.TxtAddLine4.Size = New System.Drawing.Size(336, 20)
        Me.TxtAddLine4.TabIndex = 4
        '
        'TxtAddLine5
        '
        Me.TxtAddLine5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAddLine5.Location = New System.Drawing.Point(131, 142)
        Me.TxtAddLine5.MaxLength = 50
        Me.TxtAddLine5.Name = "TxtAddLine5"
        Me.TxtAddLine5.Size = New System.Drawing.Size(336, 20)
        Me.TxtAddLine5.TabIndex = 5
        '
        'TxtTaxAddLine2
        '
        Me.TxtTaxAddLine2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTaxAddLine2.Location = New System.Drawing.Point(131, 272)
        Me.TxtTaxAddLine2.MaxLength = 100
        Me.TxtTaxAddLine2.Name = "TxtTaxAddLine2"
        Me.TxtTaxAddLine2.Size = New System.Drawing.Size(336, 20)
        Me.TxtTaxAddLine2.TabIndex = 10
        '
        'TxtTaxAddLine1
        '
        Me.TxtTaxAddLine1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTaxAddLine1.Location = New System.Drawing.Point(131, 246)
        Me.TxtTaxAddLine1.MaxLength = 100
        Me.TxtTaxAddLine1.Name = "TxtTaxAddLine1"
        Me.TxtTaxAddLine1.Size = New System.Drawing.Size(336, 20)
        Me.TxtTaxAddLine1.TabIndex = 9
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 249)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(66, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Tax Address"
        '
        'TxtNPWPNumber
        '
        Me.TxtNPWPNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNPWPNumber.Location = New System.Drawing.Point(131, 194)
        Me.TxtNPWPNumber.MaxLength = 50
        Me.TxtNPWPNumber.Name = "TxtNPWPNumber"
        Me.TxtNPWPNumber.Size = New System.Drawing.Size(336, 20)
        Me.TxtNPWPNumber.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 197)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "NPWP Number"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 226)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(113, 13)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "PKP Registration Date"
        '
        'TxtCity
        '
        Me.TxtCity.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCity.Location = New System.Drawing.Point(131, 168)
        Me.TxtCity.MaxLength = 50
        Me.TxtCity.Name = "TxtCity"
        Me.TxtCity.Size = New System.Drawing.Size(336, 20)
        Me.TxtCity.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 171)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(24, 13)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "City"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.dgvAddress)
        Me.GroupBox1.Location = New System.Drawing.Point(15, 298)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(452, 175)
        Me.GroupBox1.TabIndex = 11
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Shipment Data"
        '
        'dgvAddress
        '
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvAddress.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvAddress.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAddress.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvAddress.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAddress.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColAddressCode, Me.Col_Line1, Me.Col_Line2, Me.Col_Line3, Me.Col_Line4, Me.Col_Line5, Me.ColRemarksAdd, Me.ColAddressIdentifier})
        Me.dgvAddress.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAddress.Location = New System.Drawing.Point(3, 16)
        Me.dgvAddress.MultiSelect = False
        Me.dgvAddress.Name = "dgvAddress"
        Me.dgvAddress.Size = New System.Drawing.Size(446, 156)
        Me.dgvAddress.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "#"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 21
        '
        'ColAddressCode
        '
        Me.ColAddressCode.HeaderText = "Address Code"
        Me.ColAddressCode.Name = "ColAddressCode"
        Me.ColAddressCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAddressCode.Width = 79
        '
        'Col_Line1
        '
        Me.Col_Line1.HeaderText = "Line 1"
        Me.Col_Line1.Name = "Col_Line1"
        Me.Col_Line1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line1.Width = 42
        '
        'Col_Line2
        '
        Me.Col_Line2.HeaderText = "Line 2"
        Me.Col_Line2.Name = "Col_Line2"
        Me.Col_Line2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line2.Width = 42
        '
        'Col_Line3
        '
        Me.Col_Line3.HeaderText = "Line 3"
        Me.Col_Line3.Name = "Col_Line3"
        Me.Col_Line3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line3.Width = 42
        '
        'Col_Line4
        '
        Me.Col_Line4.HeaderText = "Line 4"
        Me.Col_Line4.Name = "Col_Line4"
        Me.Col_Line4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line4.Width = 42
        '
        'Col_Line5
        '
        Me.Col_Line5.HeaderText = "Line 5"
        Me.Col_Line5.Name = "Col_Line5"
        Me.Col_Line5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line5.Width = 42
        '
        'ColRemarksAdd
        '
        Me.ColRemarksAdd.HeaderText = "Remarks"
        Me.ColRemarksAdd.Name = "ColRemarksAdd"
        Me.ColRemarksAdd.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarksAdd.Width = 55
        '
        'ColAddressIdentifier
        '
        Me.ColAddressIdentifier.HeaderText = "Identifier"
        Me.ColAddressIdentifier.Name = "ColAddressIdentifier"
        Me.ColAddressIdentifier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAddressIdentifier.Visible = False
        Me.ColAddressIdentifier.Width = 53
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(392, 479)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 13
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(311, 479)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 12
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 510)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(479, 22)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 21
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(232, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(232, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 532)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(479, 22)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 20
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(232, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(232, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'TxtPrimaryCompany
        '
        Me.TxtPrimaryCompany.Location = New System.Drawing.Point(95, 12)
        Me.TxtPrimaryCompany.Name = "TxtPrimaryCompany"
        Me.TxtPrimaryCompany.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryCompany.TabIndex = 24
        Me.TxtPrimaryCompany.TabStop = False
        Me.TxtPrimaryCompany.Visible = False
        '
        'txtReadOnly
        '
        Me.txtReadOnly.Location = New System.Drawing.Point(102, 246)
        Me.txtReadOnly.Name = "txtReadOnly"
        Me.txtReadOnly.ReadOnly = True
        Me.txtReadOnly.Size = New System.Drawing.Size(23, 20)
        Me.txtReadOnly.TabIndex = 212
        Me.txtReadOnly.TabStop = False
        Me.txtReadOnly.Visible = False
        '
        'dtp_PKP_Reg_Date
        '
        Me.dtp_PKP_Reg_Date.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtp_PKP_Reg_Date.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtp_PKP_Reg_Date.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtp_PKP_Reg_Date.Location = New System.Drawing.Point(131, 220)
        Me.dtp_PKP_Reg_Date.Name = "dtp_PKP_Reg_Date"
        Me.dtp_PKP_Reg_Date.Size = New System.Drawing.Size(336, 20)
        Me.dtp_PKP_Reg_Date.TabIndex = 8
        '
        'FrmMasterCompanyData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(479, 554)
        Me.Controls.Add(Me.dtp_PKP_Reg_Date)
        Me.Controls.Add(Me.txtReadOnly)
        Me.Controls.Add(Me.TxtPrimaryCompany)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TxtCity)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TxtNPWPNumber)
        Me.Controls.Add(Me.TxtTaxAddLine2)
        Me.Controls.Add(Me.TxtTaxAddLine1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TxtAddLine5)
        Me.Controls.Add(Me.TxtAddLine4)
        Me.Controls.Add(Me.TxtAddLine3)
        Me.Controls.Add(Me.TxtAddLine2)
        Me.Controls.Add(Me.TxtAddLine1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtCompanyName)
        Me.Controls.Add(Me.Label1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(495, 550)
        Me.Name = "FrmMasterCompanyData"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Company Data"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.dgvAddress, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtCompanyName As System.Windows.Forms.TextBox
    Friend WithEvents TxtAddLine1 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtAddLine2 As System.Windows.Forms.TextBox
    Friend WithEvents TxtAddLine3 As System.Windows.Forms.TextBox
    Friend WithEvents TxtAddLine4 As System.Windows.Forms.TextBox
    Friend WithEvents TxtAddLine5 As System.Windows.Forms.TextBox
    Friend WithEvents TxtTaxAddLine2 As System.Windows.Forms.TextBox
    Friend WithEvents TxtTaxAddLine1 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtNPWPNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtCity As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents dgvAddress As System.Windows.Forms.DataGridView
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TxtPrimaryCompany As System.Windows.Forms.TextBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAddressCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Line1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Line2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Line3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Line4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Line5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarksAdd As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAddressIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents txtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents dtp_PKP_Reg_Date As System.Windows.Forms.DateTimePicker
End Class
