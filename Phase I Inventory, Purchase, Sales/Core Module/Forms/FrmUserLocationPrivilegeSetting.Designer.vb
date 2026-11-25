<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmUserLocationPrivilegeSetting
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgvUser = New System.Windows.Forms.DataGridView
        Me.colUser_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colUser_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colUser_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colUser_DefaultSetting = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColUser_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label2 = New System.Windows.Forms.Label
        Me.dgvLocation = New System.Windows.Forms.DataGridView
        Me.BtnClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.cboDefaultUserNotInList = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.colLocation_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colLocation_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.colLocation_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colLocation_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colLocation_StockInPurchase = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colLocation_StockInSalesReturn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockInAdjust = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockOutSales = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockOutPurcahseReturn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockOutAdjust = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockTransferOut = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockTransferIn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockBundlingBefore = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockBundlingAfter = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_ProductionlogBefore = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_ProductionLogAfter = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colLocationPrimaryRefParty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvLocation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(4, 42)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dgvUser)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvLocation)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnClose)
        Me.SplitContainer1.Size = New System.Drawing.Size(1048, 481)
        Me.SplitContainer1.SplitterDistance = 239
        Me.SplitContainer1.SplitterWidth = 5
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 0)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(101, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "User Code List"
        '
        'dgvUser
        '
        Me.dgvUser.AllowUserToResizeColumns = False
        Me.dgvUser.AllowUserToResizeRows = False
        Me.dgvUser.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvUser.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvUser.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colUser_No, Me.colUser_Code, Me.colUser_Name, Me.colUser_DefaultSetting, Me.ColUser_Primary})
        Me.dgvUser.Location = New System.Drawing.Point(4, 20)
        Me.dgvUser.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvUser.MultiSelect = False
        Me.dgvUser.Name = "dgvUser"
        Me.dgvUser.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvUser.RowTemplate.Height = 24
        Me.dgvUser.Size = New System.Drawing.Size(1040, 218)
        Me.dgvUser.TabIndex = 0
        '
        'colUser_No
        '
        Me.colUser_No.HeaderText = "No"
        Me.colUser_No.Name = "colUser_No"
        Me.colUser_No.ReadOnly = True
        Me.colUser_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.colUser_No.Width = 32
        '
        'colUser_Code
        '
        Me.colUser_Code.HeaderText = "User Code"
        Me.colUser_Code.Name = "colUser_Code"
        '
        'colUser_Name
        '
        Me.colUser_Name.HeaderText = "User Name"
        Me.colUser_Name.Name = "colUser_Name"
        Me.colUser_Name.Width = 104
        '
        'colUser_DefaultSetting
        '
        Me.colUser_DefaultSetting.HeaderText = "Location not Defined"
        Me.colUser_DefaultSetting.Items.AddRange(New Object() {"Can Access Location", "Cannot Access Location"})
        Me.colUser_DefaultSetting.Name = "colUser_DefaultSetting"
        Me.colUser_DefaultSetting.Width = 145
        '
        'ColUser_Primary
        '
        Me.ColUser_Primary.HeaderText = "Primary"
        Me.ColUser_Primary.Name = "ColUser_Primary"
        Me.ColUser_Primary.ReadOnly = True
        Me.ColUser_Primary.Visible = False
        Me.ColUser_Primary.Width = 66
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 5)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(111, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Location Access"
        '
        'dgvLocation
        '
        Me.dgvLocation.AllowUserToResizeColumns = False
        Me.dgvLocation.AllowUserToResizeRows = False
        Me.dgvLocation.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvLocation.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvLocation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvLocation.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colLocation_No, Me.colLocation_Level, Me.colLocation_Code, Me.colLocation_Desc, Me.colLocation_StockInPurchase, Me.colLocation_StockInSalesReturn, Me.ColLocation_StockInAdjust, Me.ColLocation_StockOutSales, Me.ColLocation_StockOutPurcahseReturn, Me.ColLocation_StockOutAdjust, Me.ColLocation_StockTransferOut, Me.ColLocation_StockTransferIn, Me.ColLocation_StockBundlingBefore, Me.ColLocation_StockBundlingAfter, Me.ColLocation_ProductionlogBefore, Me.ColLocation_ProductionLogAfter, Me.colLocationPrimaryRefParty})
        Me.dgvLocation.Location = New System.Drawing.Point(4, 25)
        Me.dgvLocation.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvLocation.MultiSelect = False
        Me.dgvLocation.Name = "dgvLocation"
        Me.dgvLocation.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvLocation.RowTemplate.Height = 24
        Me.dgvLocation.Size = New System.Drawing.Size(1040, 209)
        Me.dgvLocation.TabIndex = 1
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(833, 146)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(100, 28)
        Me.BtnClose.TabIndex = 5
        Me.BtnClose.TabStop = False
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BtnCancel)
        Me.GroupBox1.Controls.Add(Me.btnOK)
        Me.GroupBox1.Controls.Add(Me.cboDefaultUserNotInList)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.SplitContainer1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(1056, 570)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'BtnCancel
        '
        Me.BtnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnCancel.Location = New System.Drawing.Point(945, 532)
        Me.BtnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(100, 28)
        Me.BtnCancel.TabIndex = 4
        Me.BtnCancel.Text = "&Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(837, 532)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(100, 28)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'cboDefaultUserNotInList
        '
        Me.cboDefaultUserNotInList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDefaultUserNotInList.FormattingEnabled = True
        Me.cboDefaultUserNotInList.Items.AddRange(New Object() {"Can access location", "Cannot access location"})
        Me.cboDefaultUserNotInList.Location = New System.Drawing.Point(327, 14)
        Me.cboDefaultUserNotInList.Margin = New System.Windows.Forms.Padding(4)
        Me.cboDefaultUserNotInList.Name = "cboDefaultUserNotInList"
        Me.cboDefaultUserNotInList.Size = New System.Drawing.Size(266, 24)
        Me.cboDefaultUserNotInList.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 18)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(308, 17)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Default access if user not include in this setting:"
        '
        'BackgroundWorker2
        '
        '
        'colLocation_No
        '
        Me.colLocation_No.Frozen = True
        Me.colLocation_No.HeaderText = "No"
        Me.colLocation_No.Name = "colLocation_No"
        Me.colLocation_No.ReadOnly = True
        Me.colLocation_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.colLocation_No.Width = 32
        '
        'colLocation_Level
        '
        Me.colLocation_Level.Frozen = True
        Me.colLocation_Level.HeaderText = "Location Level"
        Me.colLocation_Level.Items.AddRange(New Object() {"Warehouse", "Location"})
        Me.colLocation_Level.Name = "colLocation_Level"
        Me.colLocation_Level.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colLocation_Level.Width = 106
        '
        'colLocation_Code
        '
        Me.colLocation_Code.Frozen = True
        Me.colLocation_Code.HeaderText = "Code"
        Me.colLocation_Code.Name = "colLocation_Code"
        Me.colLocation_Code.Width = 66
        '
        'colLocation_Desc
        '
        Me.colLocation_Desc.HeaderText = "Description"
        Me.colLocation_Desc.Name = "colLocation_Desc"
        Me.colLocation_Desc.ReadOnly = True
        Me.colLocation_Desc.Width = 104
        '
        'colLocation_StockInPurchase
        '
        Me.colLocation_StockInPurchase.HeaderText = "Stock in Purchase"
        Me.colLocation_StockInPurchase.Name = "colLocation_StockInPurchase"
        Me.colLocation_StockInPurchase.Width = 128
        '
        'colLocation_StockInSalesReturn
        '
        Me.colLocation_StockInSalesReturn.HeaderText = "Stock in Sales Return"
        Me.colLocation_StockInSalesReturn.Name = "colLocation_StockInSalesReturn"
        Me.colLocation_StockInSalesReturn.Width = 150
        '
        'ColLocation_StockInAdjust
        '
        Me.ColLocation_StockInAdjust.HeaderText = "Stock in Adjust"
        Me.ColLocation_StockInAdjust.Name = "ColLocation_StockInAdjust"
        Me.ColLocation_StockInAdjust.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColLocation_StockInAdjust.Width = 107
        '
        'ColLocation_StockOutSales
        '
        Me.ColLocation_StockOutSales.HeaderText = "Stock out Sales"
        Me.ColLocation_StockOutSales.Name = "ColLocation_StockOutSales"
        Me.ColLocation_StockOutSales.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColLocation_StockOutSales.Width = 112
        '
        'ColLocation_StockOutPurcahseReturn
        '
        Me.ColLocation_StockOutPurcahseReturn.HeaderText = "Stock out Purchase Return"
        Me.ColLocation_StockOutPurcahseReturn.Name = "ColLocation_StockOutPurcahseReturn"
        Me.ColLocation_StockOutPurcahseReturn.Width = 184
        '
        'ColLocation_StockOutAdjust
        '
        Me.ColLocation_StockOutAdjust.HeaderText = "Stock out Adjust"
        Me.ColLocation_StockOutAdjust.Name = "ColLocation_StockOutAdjust"
        Me.ColLocation_StockOutAdjust.Width = 116
        '
        'ColLocation_StockTransferOut
        '
        Me.ColLocation_StockTransferOut.HeaderText = "Stock Transfer Out"
        Me.ColLocation_StockTransferOut.Name = "ColLocation_StockTransferOut"
        Me.ColLocation_StockTransferOut.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColLocation_StockTransferOut.Width = 134
        '
        'ColLocation_StockTransferIn
        '
        Me.ColLocation_StockTransferIn.HeaderText = "Stock Transfer In"
        Me.ColLocation_StockTransferIn.Name = "ColLocation_StockTransferIn"
        Me.ColLocation_StockTransferIn.Width = 122
        '
        'ColLocation_StockBundlingBefore
        '
        Me.ColLocation_StockBundlingBefore.HeaderText = "Stock Bundling Before"
        Me.ColLocation_StockBundlingBefore.Name = "ColLocation_StockBundlingBefore"
        Me.ColLocation_StockBundlingBefore.Width = 154
        '
        'ColLocation_StockBundlingAfter
        '
        Me.ColLocation_StockBundlingAfter.HeaderText = "Stock Bundling After"
        Me.ColLocation_StockBundlingAfter.Name = "ColLocation_StockBundlingAfter"
        Me.ColLocation_StockBundlingAfter.Width = 142
        '
        'ColLocation_ProductionlogBefore
        '
        Me.ColLocation_ProductionlogBefore.HeaderText = "Production Log RM"
        Me.ColLocation_ProductionlogBefore.Name = "ColLocation_ProductionlogBefore"
        Me.ColLocation_ProductionlogBefore.Width = 135
        '
        'ColLocation_ProductionLogAfter
        '
        Me.ColLocation_ProductionLogAfter.HeaderText = "Production Log FG"
        Me.ColLocation_ProductionLogAfter.Name = "ColLocation_ProductionLogAfter"
        Me.ColLocation_ProductionLogAfter.Width = 133
        '
        'colLocationPrimaryRefParty
        '
        Me.colLocationPrimaryRefParty.HeaderText = "Primary Ref Party"
        Me.colLocationPrimaryRefParty.Name = "colLocationPrimaryRefParty"
        Me.colLocationPrimaryRefParty.ReadOnly = True
        Me.colLocationPrimaryRefParty.Visible = False
        Me.colLocationPrimaryRefParty.Width = 144
        '
        'FrmUserLocationPrivilegeSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(1056, 570)
        Me.Controls.Add(Me.GroupBox1)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimumSize = New System.Drawing.Size(1064, 607)
        Me.Name = "FrmUserLocationPrivilegeSetting"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "User Location Privilege Setting"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvLocation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents dgvLocation As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboDefaultUserNotInList As System.Windows.Forms.ComboBox
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents colUser_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colUser_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colUser_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colUser_DefaultSetting As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColUser_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents colLocation_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colLocation_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents colLocation_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colLocation_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colLocation_StockInPurchase As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colLocation_StockInSalesReturn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockInAdjust As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockOutSales As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockOutPurcahseReturn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockOutAdjust As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockTransferOut As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockTransferIn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockBundlingBefore As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockBundlingAfter As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_ProductionlogBefore As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_ProductionLogAfter As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colLocationPrimaryRefParty As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
