<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiRepairInvoice
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
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage_SO = New System.Windows.Forms.TabPage
        Me.btnInvRefresh = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.chkInvShowAll = New System.Windows.Forms.CheckBox
        Me.dtpInvPeriod = New System.Windows.Forms.DateTimePicker
        Me.dgvSO = New System.Windows.Forms.DataGridView
        Me.ColSO_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSO_SONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSO_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSO_CustCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSO_CustName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSO_SalesPersonCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSO_SalesPersonName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSO_PrimaryCust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSO_PrimarySalesPerson = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSO_PrimarySO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage_Customer = New System.Windows.Forms.TabPage
        Me.brnReloadCust = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.ChkCust_ShowAll = New System.Windows.Forms.CheckBox
        Me.dtpCust_Period = New System.Windows.Forms.DateTimePicker
        Me.dgvCustomer = New System.Windows.Forms.DataGridView
        Me.ColCust_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCust_CustCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCust_Customer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCust_SalesPersonCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCust_SalesPersonName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCust_PrimaryCust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCust_PrimarySalesPerson = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabControl1.SuspendLayout()
        Me.TabPage_SO.SuspendLayout()
        CType(Me.dgvSO, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage_Customer.SuspendLayout()
        CType(Me.dgvCustomer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage_SO)
        Me.TabControl1.Controls.Add(Me.TabPage_Customer)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(992, 633)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage_SO
        '
        Me.TabPage_SO.Controls.Add(Me.btnInvRefresh)
        Me.TabPage_SO.Controls.Add(Me.Label2)
        Me.TabPage_SO.Controls.Add(Me.chkInvShowAll)
        Me.TabPage_SO.Controls.Add(Me.dtpInvPeriod)
        Me.TabPage_SO.Controls.Add(Me.dgvSO)
        Me.TabPage_SO.Location = New System.Drawing.Point(4, 25)
        Me.TabPage_SO.Name = "TabPage_SO"
        Me.TabPage_SO.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_SO.Size = New System.Drawing.Size(984, 604)
        Me.TabPage_SO.TabIndex = 0
        Me.TabPage_SO.Text = "Sales Order"
        Me.TabPage_SO.UseVisualStyleBackColor = True
        '
        'btnInvRefresh
        '
        Me.btnInvRefresh.Location = New System.Drawing.Point(473, 7)
        Me.btnInvRefresh.Name = "btnInvRefresh"
        Me.btnInvRefresh.Size = New System.Drawing.Size(75, 23)
        Me.btnInvRefresh.TabIndex = 8
        Me.btnInvRefresh.Text = "Refresh"
        Me.btnInvRefresh.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 10)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 17)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Period"
        '
        'chkInvShowAll
        '
        Me.chkInvShowAll.AutoSize = True
        Me.chkInvShowAll.Location = New System.Drawing.Point(259, 8)
        Me.chkInvShowAll.Name = "chkInvShowAll"
        Me.chkInvShowAll.Size = New System.Drawing.Size(208, 21)
        Me.chkInvShowAll.TabIndex = 6
        Me.chkInvShowAll.Text = "Show data from all customer"
        Me.chkInvShowAll.UseVisualStyleBackColor = True
        '
        'dtpInvPeriod
        '
        Me.dtpInvPeriod.CustomFormat = "MMM yyyy"
        Me.dtpInvPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpInvPeriod.Location = New System.Drawing.Point(108, 7)
        Me.dtpInvPeriod.Name = "dtpInvPeriod"
        Me.dtpInvPeriod.ShowUpDown = True
        Me.dtpInvPeriod.Size = New System.Drawing.Size(145, 22)
        Me.dtpInvPeriod.TabIndex = 5
        '
        'dgvSO
        '
        Me.dgvSO.AllowUserToAddRows = False
        Me.dgvSO.AllowUserToDeleteRows = False
        Me.dgvSO.AllowUserToResizeColumns = False
        Me.dgvSO.AllowUserToResizeRows = False
        Me.dgvSO.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvSO.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSO.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSO.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvSO.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSO_No, Me.ColSO_SONo, Me.ColSO_Date, Me.ColSO_CustCode, Me.ColSO_CustName, Me.ColSO_SalesPersonCode, Me.ColSO_SalesPersonName, Me.ColSO_PrimaryCust, Me.ColSO_PrimarySalesPerson, Me.ColSO_PrimarySO})
        Me.dgvSO.Location = New System.Drawing.Point(8, 35)
        Me.dgvSO.Name = "dgvSO"
        Me.dgvSO.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvSO.RowTemplate.Height = 24
        Me.dgvSO.Size = New System.Drawing.Size(968, 562)
        Me.dgvSO.TabIndex = 4
        '
        'ColSO_No
        '
        Me.ColSO_No.HeaderText = "No"
        Me.ColSO_No.Name = "ColSO_No"
        Me.ColSO_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColSO_No.Width = 51
        '
        'ColSO_SONo
        '
        Me.ColSO_SONo.HeaderText = "Sales Order No"
        Me.ColSO_SONo.Name = "ColSO_SONo"
        Me.ColSO_SONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColSO_SONo.Width = 131
        '
        'ColSO_Date
        '
        Me.ColSO_Date.HeaderText = "Date"
        Me.ColSO_Date.Name = "ColSO_Date"
        Me.ColSO_Date.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColSO_Date.Width = 63
        '
        'ColSO_CustCode
        '
        Me.ColSO_CustCode.HeaderText = "Customer Code"
        Me.ColSO_CustCode.Name = "ColSO_CustCode"
        Me.ColSO_CustCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColSO_CustCode.Width = 130
        '
        'ColSO_CustName
        '
        Me.ColSO_CustName.HeaderText = "Customer Name"
        Me.ColSO_CustName.Name = "ColSO_CustName"
        Me.ColSO_CustName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColSO_CustName.Width = 134
        '
        'ColSO_SalesPersonCode
        '
        Me.ColSO_SalesPersonCode.HeaderText = "Sales Person Code"
        Me.ColSO_SalesPersonCode.Name = "ColSO_SalesPersonCode"
        Me.ColSO_SalesPersonCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColSO_SalesPersonCode.Width = 154
        '
        'ColSO_SalesPersonName
        '
        Me.ColSO_SalesPersonName.HeaderText = "Sales Person Name"
        Me.ColSO_SalesPersonName.Name = "ColSO_SalesPersonName"
        Me.ColSO_SalesPersonName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColSO_SalesPersonName.Width = 158
        '
        'ColSO_PrimaryCust
        '
        Me.ColSO_PrimaryCust.HeaderText = "Primary Cust"
        Me.ColSO_PrimaryCust.Name = "ColSO_PrimaryCust"
        Me.ColSO_PrimaryCust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColSO_PrimaryCust.Visible = False
        Me.ColSO_PrimaryCust.Width = 113
        '
        'ColSO_PrimarySalesPerson
        '
        Me.ColSO_PrimarySalesPerson.HeaderText = "Primary Sales Person"
        Me.ColSO_PrimarySalesPerson.Name = "ColSO_PrimarySalesPerson"
        Me.ColSO_PrimarySalesPerson.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColSO_PrimarySalesPerson.Visible = False
        Me.ColSO_PrimarySalesPerson.Width = 169
        '
        'ColSO_PrimarySO
        '
        Me.ColSO_PrimarySO.HeaderText = "Primary SO"
        Me.ColSO_PrimarySO.Name = "ColSO_PrimarySO"
        Me.ColSO_PrimarySO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColSO_PrimarySO.Visible = False
        Me.ColSO_PrimarySO.Width = 105
        '
        'TabPage_Customer
        '
        Me.TabPage_Customer.Controls.Add(Me.brnReloadCust)
        Me.TabPage_Customer.Controls.Add(Me.Label1)
        Me.TabPage_Customer.Controls.Add(Me.ChkCust_ShowAll)
        Me.TabPage_Customer.Controls.Add(Me.dtpCust_Period)
        Me.TabPage_Customer.Controls.Add(Me.dgvCustomer)
        Me.TabPage_Customer.Location = New System.Drawing.Point(4, 25)
        Me.TabPage_Customer.Name = "TabPage_Customer"
        Me.TabPage_Customer.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_Customer.Size = New System.Drawing.Size(984, 604)
        Me.TabPage_Customer.TabIndex = 1
        Me.TabPage_Customer.Text = "Customer"
        Me.TabPage_Customer.UseVisualStyleBackColor = True
        '
        'brnReloadCust
        '
        Me.brnReloadCust.Location = New System.Drawing.Point(259, 6)
        Me.brnReloadCust.Name = "brnReloadCust"
        Me.brnReloadCust.Size = New System.Drawing.Size(75, 23)
        Me.brnReloadCust.TabIndex = 9
        Me.brnReloadCust.Text = "Refresh"
        Me.brnReloadCust.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Period"
        '
        'ChkCust_ShowAll
        '
        Me.ChkCust_ShowAll.AutoSize = True
        Me.ChkCust_ShowAll.Location = New System.Drawing.Point(259, 6)
        Me.ChkCust_ShowAll.Name = "ChkCust_ShowAll"
        Me.ChkCust_ShowAll.Size = New System.Drawing.Size(208, 21)
        Me.ChkCust_ShowAll.TabIndex = 2
        Me.ChkCust_ShowAll.TabStop = False
        Me.ChkCust_ShowAll.Text = "Show data from all customer"
        Me.ChkCust_ShowAll.UseVisualStyleBackColor = True
        Me.ChkCust_ShowAll.Visible = False
        '
        'dtpCust_Period
        '
        Me.dtpCust_Period.CustomFormat = "MMM yyyy"
        Me.dtpCust_Period.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpCust_Period.Location = New System.Drawing.Point(108, 6)
        Me.dtpCust_Period.Name = "dtpCust_Period"
        Me.dtpCust_Period.ShowUpDown = True
        Me.dtpCust_Period.Size = New System.Drawing.Size(145, 22)
        Me.dtpCust_Period.TabIndex = 1
        '
        'dgvCustomer
        '
        Me.dgvCustomer.AllowUserToAddRows = False
        Me.dgvCustomer.AllowUserToDeleteRows = False
        Me.dgvCustomer.AllowUserToResizeColumns = False
        Me.dgvCustomer.AllowUserToResizeRows = False
        Me.dgvCustomer.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCustomer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCustomer.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvCustomer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvCustomer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCust_No, Me.ColCust_CustCode, Me.ColCust_Customer, Me.ColCust_SalesPersonCode, Me.ColCust_SalesPersonName, Me.ColCust_PrimaryCust, Me.ColCust_PrimarySalesPerson})
        Me.dgvCustomer.Location = New System.Drawing.Point(8, 34)
        Me.dgvCustomer.Name = "dgvCustomer"
        Me.dgvCustomer.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvCustomer.RowTemplate.Height = 24
        Me.dgvCustomer.Size = New System.Drawing.Size(968, 562)
        Me.dgvCustomer.TabIndex = 0
        '
        'ColCust_No
        '
        Me.ColCust_No.HeaderText = "No"
        Me.ColCust_No.Name = "ColCust_No"
        Me.ColCust_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColCust_No.Width = 51
        '
        'ColCust_CustCode
        '
        Me.ColCust_CustCode.HeaderText = "Customer Code"
        Me.ColCust_CustCode.Name = "ColCust_CustCode"
        Me.ColCust_CustCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColCust_CustCode.Width = 130
        '
        'ColCust_Customer
        '
        Me.ColCust_Customer.HeaderText = "Customer Name"
        Me.ColCust_Customer.Name = "ColCust_Customer"
        Me.ColCust_Customer.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColCust_Customer.Width = 134
        '
        'ColCust_SalesPersonCode
        '
        Me.ColCust_SalesPersonCode.HeaderText = "Sales Person Code"
        Me.ColCust_SalesPersonCode.Name = "ColCust_SalesPersonCode"
        Me.ColCust_SalesPersonCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColCust_SalesPersonCode.Width = 154
        '
        'ColCust_SalesPersonName
        '
        Me.ColCust_SalesPersonName.HeaderText = "Sales Person Name"
        Me.ColCust_SalesPersonName.Name = "ColCust_SalesPersonName"
        Me.ColCust_SalesPersonName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColCust_SalesPersonName.Width = 158
        '
        'ColCust_PrimaryCust
        '
        Me.ColCust_PrimaryCust.HeaderText = "Primary Cust"
        Me.ColCust_PrimaryCust.Name = "ColCust_PrimaryCust"
        Me.ColCust_PrimaryCust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColCust_PrimaryCust.Visible = False
        Me.ColCust_PrimaryCust.Width = 113
        '
        'ColCust_PrimarySalesPerson
        '
        Me.ColCust_PrimarySalesPerson.HeaderText = "Primary Sales Person"
        Me.ColCust_PrimarySalesPerson.Name = "ColCust_PrimarySalesPerson"
        Me.ColCust_PrimarySalesPerson.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColCust_PrimarySalesPerson.Visible = False
        Me.ColCust_PrimarySalesPerson.Width = 169
        '
        'FrmTransaksiRepairInvoice
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.TabControl1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "FrmTransaksiRepairInvoice"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Repair Sales Person Data"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage_SO.ResumeLayout(False)
        Me.TabPage_SO.PerformLayout()
        CType(Me.dgvSO, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage_Customer.ResumeLayout(False)
        Me.TabPage_Customer.PerformLayout()
        CType(Me.dgvCustomer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage_SO As System.Windows.Forms.TabPage
    Friend WithEvents btnInvRefresh As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkInvShowAll As System.Windows.Forms.CheckBox
    Friend WithEvents dtpInvPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents dgvSO As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage_Customer As System.Windows.Forms.TabPage
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ChkCust_ShowAll As System.Windows.Forms.CheckBox
    Friend WithEvents dtpCust_Period As System.Windows.Forms.DateTimePicker
    Friend WithEvents dgvCustomer As System.Windows.Forms.DataGridView
    Friend WithEvents brnReloadCust As System.Windows.Forms.Button
    Friend WithEvents ColSO_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSO_SONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSO_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSO_CustCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSO_CustName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSO_SalesPersonCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSO_SalesPersonName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSO_PrimaryCust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSO_PrimarySalesPerson As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSO_PrimarySO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCust_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCust_CustCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCust_Customer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCust_SalesPersonCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCust_SalesPersonName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCust_PrimaryCust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCust_PrimarySalesPerson As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
