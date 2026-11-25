<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterCustomer_CekEmailCP
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
        Me.dgvCP = New System.Windows.Forms.DataGridView
        Me.ColPersonCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPerson_Customer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPersonName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPersonPosition = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPersonPhoneNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPersonFax = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPersonEmail = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPersonSendCL = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPerson_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPersonIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPerson_PrimaryCust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPersonPrimaryItem = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdd = New System.Windows.Forms.DataGridViewButtonColumn
        Me.Button1 = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        CType(Me.dgvCP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvCP
        '
        Me.dgvCP.AllowUserToAddRows = False
        Me.dgvCP.AllowUserToDeleteRows = False
        Me.dgvCP.AllowUserToResizeColumns = False
        Me.dgvCP.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvCP.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvCP.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvCP.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColPersonCode, Me.ColPerson_Customer, Me.ColPersonName, Me.ColPersonPosition, Me.ColPersonPhoneNumber, Me.ColPersonFax, Me.ColPersonEmail, Me.ColPersonSendCL, Me.ColPerson_Remarks, Me.ColPersonIdentifier, Me.ColPerson_PrimaryCust, Me.ColPersonPrimaryItem, Me.ColAdd})
        Me.dgvCP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCP.Location = New System.Drawing.Point(0, 0)
        Me.dgvCP.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvCP.MultiSelect = False
        Me.dgvCP.Name = "dgvCP"
        Me.dgvCP.RowTemplate.Height = 24
        Me.dgvCP.Size = New System.Drawing.Size(892, 513)
        Me.dgvCP.TabIndex = 60
        '
        'ColPersonCode
        '
        Me.ColPersonCode.HeaderText = "#"
        Me.ColPersonCode.Name = "ColPersonCode"
        Me.ColPersonCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPersonCode.Width = 22
        '
        'ColPerson_Customer
        '
        Me.ColPerson_Customer.HeaderText = "Customer"
        Me.ColPerson_Customer.Name = "ColPerson_Customer"
        Me.ColPerson_Customer.Width = 93
        '
        'ColPersonName
        '
        Me.ColPersonName.HeaderText = "Contact Person Name"
        Me.ColPersonName.Name = "ColPersonName"
        Me.ColPersonName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPersonName.Width = 152
        '
        'ColPersonPosition
        '
        Me.ColPersonPosition.HeaderText = "Position"
        Me.ColPersonPosition.Name = "ColPersonPosition"
        Me.ColPersonPosition.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPersonPosition.Width = 64
        '
        'ColPersonPhoneNumber
        '
        Me.ColPersonPhoneNumber.HeaderText = "Phone Number"
        Me.ColPersonPhoneNumber.Name = "ColPersonPhoneNumber"
        Me.ColPersonPhoneNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPersonPhoneNumber.Width = 109
        '
        'ColPersonFax
        '
        Me.ColPersonFax.HeaderText = "Fax Number"
        Me.ColPersonFax.Name = "ColPersonFax"
        Me.ColPersonFax.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPersonFax.Width = 90
        '
        'ColPersonEmail
        '
        Me.ColPersonEmail.HeaderText = "Email Address"
        Me.ColPersonEmail.Name = "ColPersonEmail"
        Me.ColPersonEmail.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPersonEmail.Width = 104
        '
        'ColPersonSendCL
        '
        Me.ColPersonSendCL.HeaderText = "Send Collection Letter"
        Me.ColPersonSendCL.Name = "ColPersonSendCL"
        Me.ColPersonSendCL.Width = 153
        '
        'ColPerson_Remarks
        '
        Me.ColPerson_Remarks.HeaderText = "Remarks"
        Me.ColPerson_Remarks.Name = "ColPerson_Remarks"
        Me.ColPerson_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPerson_Remarks.Width = 70
        '
        'ColPersonIdentifier
        '
        Me.ColPersonIdentifier.HeaderText = "Identifier"
        Me.ColPersonIdentifier.Name = "ColPersonIdentifier"
        Me.ColPersonIdentifier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPersonIdentifier.Visible = False
        Me.ColPersonIdentifier.Width = 68
        '
        'ColPerson_PrimaryCust
        '
        Me.ColPerson_PrimaryCust.HeaderText = "Primary Cust"
        Me.ColPerson_PrimaryCust.Name = "ColPerson_PrimaryCust"
        Me.ColPerson_PrimaryCust.Visible = False
        Me.ColPerson_PrimaryCust.Width = 113
        '
        'ColPersonPrimaryItem
        '
        Me.ColPersonPrimaryItem.HeaderText = "Primary"
        Me.ColPersonPrimaryItem.Name = "ColPersonPrimaryItem"
        Me.ColPersonPrimaryItem.Visible = False
        Me.ColPersonPrimaryItem.Width = 81
        '
        'ColAdd
        '
        Me.ColAdd.HeaderText = "Add More"
        Me.ColAdd.Name = "ColAdd"
        Me.ColAdd.Width = 75
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(679, 185)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 61
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'FrmMasterCustomer_CekEmailCP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(892, 513)
        Me.Controls.Add(Me.dgvCP)
        Me.Controls.Add(Me.Button1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(900, 550)
        Me.Name = "FrmMasterCustomer_CekEmailCP"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Check Contact Person Email"
        CType(Me.dgvCP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvCP As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ColPersonCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPerson_Customer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPersonName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPersonPosition As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPersonPhoneNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPersonFax As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPersonEmail As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPersonSendCL As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPerson_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPersonIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPerson_PrimaryCust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPersonPrimaryItem As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdd As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
