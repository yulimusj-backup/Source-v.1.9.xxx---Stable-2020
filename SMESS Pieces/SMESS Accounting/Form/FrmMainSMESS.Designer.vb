<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMainSMESS
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
        Me.btnAccounting = New System.Windows.Forms.Button
        Me.btnMaster = New System.Windows.Forms.Button
        Me.btnInternalMessage = New System.Windows.Forms.Button
        Me.PnlMasterData = New System.Windows.Forms.Panel
        Me.FrmMasterSalesArea1 = New System.Windows.Forms.Button
        Me.FrmMasterStorageArea1 = New System.Windows.Forms.Button
        Me.FrmMasterGudang1 = New System.Windows.Forms.Button
        Me.FrmMasterSupplier1 = New System.Windows.Forms.Button
        Me.FrmMasterCustomerGroup1 = New System.Windows.Forms.Button
        Me.FrmMasterCustomer1 = New System.Windows.Forms.Button
        Me.FrmMasterCustomerSegmentation1 = New System.Windows.Forms.Button
        Me.FrmMasterCustomerType1 = New System.Windows.Forms.Button
        Me.FrmMasterProductGroup1 = New System.Windows.Forms.Button
        Me.FrmMasterCategory1 = New System.Windows.Forms.Button
        Me.FrmMasterSubBrand1 = New System.Windows.Forms.Button
        Me.FrmMasterBrand1 = New System.Windows.Forms.Button
        Me.FrmTransaksiBudget1 = New System.Windows.Forms.Button
        Me.FrmMasterCurrency1 = New System.Windows.Forms.Button
        Me.FrmMasterAkun1 = New System.Windows.Forms.Button
        Me.FrmMasterUnit1 = New System.Windows.Forms.Button
        Me.FrmMasterProduct1 = New System.Windows.Forms.Button
        Me.Label12 = New System.Windows.Forms.Label
        Me.PnlAccounting = New System.Windows.Forms.Panel
        Me.FrmTransaksiJournal_View1 = New System.Windows.Forms.Button
        Me.FrmTransaksiReminder1 = New System.Windows.Forms.Button
        Me.FrmTransaksiTagihan1 = New System.Windows.Forms.Button
        Me.FrmTransaksiClosing1 = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.FrmTransaksiJournal_Closing1 = New System.Windows.Forms.Button
        Me.FrmTransaksiJournal_Adjust1 = New System.Windows.Forms.Button
        Me.FrmTransaksiJournal_Purchase21 = New System.Windows.Forms.Button
        Me.FrmTransaksiJournal_Purchase11 = New System.Windows.Forms.Button
        Me.FrmTransaksiSalesCycle22 = New System.Windows.Forms.Button
        Me.FrmTransaksiSalesCycle12 = New System.Windows.Forms.Button
        Me.FrmTransaksiGeneralJournal1 = New System.Windows.Forms.Button
        Me.Label13 = New System.Windows.Forms.Label
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.PnlMasterData.SuspendLayout()
        Me.PnlAccounting.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.BackColor = System.Drawing.Color.White
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnAccounting)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnMaster)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnInternalMessage)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.BackColor = System.Drawing.Color.White
        Me.SplitContainer1.Panel2.Controls.Add(Me.PnlMasterData)
        Me.SplitContainer1.Panel2.Controls.Add(Me.PnlAccounting)
        Me.SplitContainer1.Size = New System.Drawing.Size(896, 606)
        Me.SplitContainer1.SplitterDistance = 121
        Me.SplitContainer1.TabIndex = 0
        '
        'btnAccounting
        '
        Me.btnAccounting.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.btnAccounting.Location = New System.Drawing.Point(15, 95)
        Me.btnAccounting.Name = "btnAccounting"
        Me.btnAccounting.Size = New System.Drawing.Size(89, 77)
        Me.btnAccounting.TabIndex = 4
        Me.btnAccounting.Text = "Accounting"
        Me.btnAccounting.UseVisualStyleBackColor = True
        '
        'btnMaster
        '
        Me.btnMaster.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.btnMaster.BackColor = System.Drawing.SystemColors.Control
        Me.btnMaster.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnMaster.Location = New System.Drawing.Point(15, 12)
        Me.btnMaster.Name = "btnMaster"
        Me.btnMaster.Size = New System.Drawing.Size(89, 77)
        Me.btnMaster.TabIndex = 0
        Me.btnMaster.Text = "Master Data"
        Me.btnMaster.UseVisualStyleBackColor = False
        '
        'btnInternalMessage
        '
        Me.btnInternalMessage.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.btnInternalMessage.Location = New System.Drawing.Point(15, 178)
        Me.btnInternalMessage.Name = "btnInternalMessage"
        Me.btnInternalMessage.Size = New System.Drawing.Size(89, 77)
        Me.btnInternalMessage.TabIndex = 5
        Me.btnInternalMessage.Text = "Internal Message"
        Me.btnInternalMessage.UseVisualStyleBackColor = True
        '
        'PnlMasterData
        '
        Me.PnlMasterData.BackColor = System.Drawing.Color.White
        Me.PnlMasterData.Controls.Add(Me.FrmMasterSalesArea1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterStorageArea1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterGudang1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterSupplier1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterCustomerGroup1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterCustomer1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterCustomerSegmentation1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterCustomerType1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterProductGroup1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterCategory1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterSubBrand1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterBrand1)
        Me.PnlMasterData.Controls.Add(Me.FrmTransaksiBudget1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterCurrency1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterAkun1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterUnit1)
        Me.PnlMasterData.Controls.Add(Me.FrmMasterProduct1)
        Me.PnlMasterData.Controls.Add(Me.Label12)
        Me.PnlMasterData.Location = New System.Drawing.Point(13, 10)
        Me.PnlMasterData.Name = "PnlMasterData"
        Me.PnlMasterData.Size = New System.Drawing.Size(746, 577)
        Me.PnlMasterData.TabIndex = 10
        '
        'FrmMasterSalesArea1
        '
        Me.FrmMasterSalesArea1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterSalesArea1.Location = New System.Drawing.Point(413, 126)
        Me.FrmMasterSalesArea1.Name = "FrmMasterSalesArea1"
        Me.FrmMasterSalesArea1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterSalesArea1.TabIndex = 24
        Me.FrmMasterSalesArea1.Text = "Customer Sales Area"
        Me.FrmMasterSalesArea1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterSalesArea1.UseVisualStyleBackColor = True
        '
        'FrmMasterStorageArea1
        '
        Me.FrmMasterStorageArea1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterStorageArea1.Location = New System.Drawing.Point(536, 221)
        Me.FrmMasterStorageArea1.Name = "FrmMasterStorageArea1"
        Me.FrmMasterStorageArea1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterStorageArea1.TabIndex = 23
        Me.FrmMasterStorageArea1.Text = "Storage Area"
        Me.FrmMasterStorageArea1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterStorageArea1.UseVisualStyleBackColor = True
        '
        'FrmMasterGudang1
        '
        Me.FrmMasterGudang1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterGudang1.Location = New System.Drawing.Point(413, 221)
        Me.FrmMasterGudang1.Name = "FrmMasterGudang1"
        Me.FrmMasterGudang1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterGudang1.TabIndex = 22
        Me.FrmMasterGudang1.Text = "Warehouse"
        Me.FrmMasterGudang1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterGudang1.UseVisualStyleBackColor = True
        '
        'FrmMasterSupplier1
        '
        Me.FrmMasterSupplier1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterSupplier1.Location = New System.Drawing.Point(48, 221)
        Me.FrmMasterSupplier1.Name = "FrmMasterSupplier1"
        Me.FrmMasterSupplier1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterSupplier1.TabIndex = 21
        Me.FrmMasterSupplier1.Text = "Supplier"
        Me.FrmMasterSupplier1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterSupplier1.UseVisualStyleBackColor = True
        '
        'FrmMasterCustomerGroup1
        '
        Me.FrmMasterCustomerGroup1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterCustomerGroup1.Location = New System.Drawing.Point(48, 126)
        Me.FrmMasterCustomerGroup1.Name = "FrmMasterCustomerGroup1"
        Me.FrmMasterCustomerGroup1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterCustomerGroup1.TabIndex = 20
        Me.FrmMasterCustomerGroup1.Text = "Customer Group"
        Me.FrmMasterCustomerGroup1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterCustomerGroup1.UseVisualStyleBackColor = True
        '
        'FrmMasterCustomer1
        '
        Me.FrmMasterCustomer1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterCustomer1.Location = New System.Drawing.Point(536, 126)
        Me.FrmMasterCustomer1.Name = "FrmMasterCustomer1"
        Me.FrmMasterCustomer1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterCustomer1.TabIndex = 19
        Me.FrmMasterCustomer1.Text = "Customer"
        Me.FrmMasterCustomer1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterCustomer1.UseVisualStyleBackColor = True
        '
        'FrmMasterCustomerSegmentation1
        '
        Me.FrmMasterCustomerSegmentation1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterCustomerSegmentation1.Location = New System.Drawing.Point(291, 126)
        Me.FrmMasterCustomerSegmentation1.Name = "FrmMasterCustomerSegmentation1"
        Me.FrmMasterCustomerSegmentation1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterCustomerSegmentation1.TabIndex = 18
        Me.FrmMasterCustomerSegmentation1.Text = "Customer Segmentation"
        Me.FrmMasterCustomerSegmentation1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterCustomerSegmentation1.UseVisualStyleBackColor = True
        '
        'FrmMasterCustomerType1
        '
        Me.FrmMasterCustomerType1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterCustomerType1.Location = New System.Drawing.Point(165, 126)
        Me.FrmMasterCustomerType1.Name = "FrmMasterCustomerType1"
        Me.FrmMasterCustomerType1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterCustomerType1.TabIndex = 17
        Me.FrmMasterCustomerType1.Text = "Customer Type"
        Me.FrmMasterCustomerType1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterCustomerType1.UseVisualStyleBackColor = True
        '
        'FrmMasterProductGroup1
        '
        Me.FrmMasterProductGroup1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterProductGroup1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.FrmMasterProductGroup1.Location = New System.Drawing.Point(48, 34)
        Me.FrmMasterProductGroup1.Name = "FrmMasterProductGroup1"
        Me.FrmMasterProductGroup1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterProductGroup1.TabIndex = 16
        Me.FrmMasterProductGroup1.Text = "Product Group"
        Me.FrmMasterProductGroup1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterProductGroup1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.FrmMasterProductGroup1.UseVisualStyleBackColor = True
        '
        'FrmMasterCategory1
        '
        Me.FrmMasterCategory1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterCategory1.Location = New System.Drawing.Point(413, 34)
        Me.FrmMasterCategory1.Name = "FrmMasterCategory1"
        Me.FrmMasterCategory1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterCategory1.TabIndex = 15
        Me.FrmMasterCategory1.Text = "Category"
        Me.FrmMasterCategory1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterCategory1.UseVisualStyleBackColor = True
        '
        'FrmMasterSubBrand1
        '
        Me.FrmMasterSubBrand1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterSubBrand1.Location = New System.Drawing.Point(291, 34)
        Me.FrmMasterSubBrand1.Name = "FrmMasterSubBrand1"
        Me.FrmMasterSubBrand1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterSubBrand1.TabIndex = 14
        Me.FrmMasterSubBrand1.Text = "Sub Brand"
        Me.FrmMasterSubBrand1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterSubBrand1.UseVisualStyleBackColor = True
        '
        'FrmMasterBrand1
        '
        Me.FrmMasterBrand1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterBrand1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.FrmMasterBrand1.Location = New System.Drawing.Point(165, 34)
        Me.FrmMasterBrand1.Name = "FrmMasterBrand1"
        Me.FrmMasterBrand1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterBrand1.TabIndex = 13
        Me.FrmMasterBrand1.Text = "Brand"
        Me.FrmMasterBrand1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterBrand1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiBudget1
        '
        Me.FrmTransaksiBudget1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmTransaksiBudget1.Location = New System.Drawing.Point(165, 318)
        Me.FrmTransaksiBudget1.Name = "FrmTransaksiBudget1"
        Me.FrmTransaksiBudget1.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiBudget1.TabIndex = 12
        Me.FrmTransaksiBudget1.Text = "Account Budget"
        Me.FrmTransaksiBudget1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmTransaksiBudget1.UseVisualStyleBackColor = True
        '
        'FrmMasterCurrency1
        '
        Me.FrmMasterCurrency1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterCurrency1.Location = New System.Drawing.Point(291, 221)
        Me.FrmMasterCurrency1.Name = "FrmMasterCurrency1"
        Me.FrmMasterCurrency1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterCurrency1.TabIndex = 11
        Me.FrmMasterCurrency1.Text = "Currency"
        Me.FrmMasterCurrency1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterCurrency1.UseVisualStyleBackColor = True
        '
        'FrmMasterAkun1
        '
        Me.FrmMasterAkun1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterAkun1.Location = New System.Drawing.Point(48, 318)
        Me.FrmMasterAkun1.Name = "FrmMasterAkun1"
        Me.FrmMasterAkun1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterAkun1.TabIndex = 10
        Me.FrmMasterAkun1.Text = "Chart Of Account"
        Me.FrmMasterAkun1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterAkun1.UseVisualStyleBackColor = True
        '
        'FrmMasterUnit1
        '
        Me.FrmMasterUnit1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterUnit1.Location = New System.Drawing.Point(165, 221)
        Me.FrmMasterUnit1.Name = "FrmMasterUnit1"
        Me.FrmMasterUnit1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterUnit1.TabIndex = 9
        Me.FrmMasterUnit1.Text = "Units"
        Me.FrmMasterUnit1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterUnit1.UseVisualStyleBackColor = True
        '
        'FrmMasterProduct1
        '
        Me.FrmMasterProduct1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmMasterProduct1.Location = New System.Drawing.Point(536, 34)
        Me.FrmMasterProduct1.Name = "FrmMasterProduct1"
        Me.FrmMasterProduct1.Size = New System.Drawing.Size(89, 77)
        Me.FrmMasterProduct1.TabIndex = 8
        Me.FrmMasterProduct1.Text = "Product"
        Me.FrmMasterProduct1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FrmMasterProduct1.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(22, 14)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(85, 17)
        Me.Label12.TabIndex = 8
        Me.Label12.Text = "Master Data"
        '
        'PnlAccounting
        '
        Me.PnlAccounting.BackColor = System.Drawing.Color.White
        Me.PnlAccounting.Controls.Add(Me.FrmTransaksiJournal_View1)
        Me.PnlAccounting.Controls.Add(Me.FrmTransaksiReminder1)
        Me.PnlAccounting.Controls.Add(Me.FrmTransaksiTagihan1)
        Me.PnlAccounting.Controls.Add(Me.FrmTransaksiClosing1)
        Me.PnlAccounting.Controls.Add(Me.Label9)
        Me.PnlAccounting.Controls.Add(Me.FrmTransaksiJournal_Closing1)
        Me.PnlAccounting.Controls.Add(Me.FrmTransaksiJournal_Adjust1)
        Me.PnlAccounting.Controls.Add(Me.FrmTransaksiJournal_Purchase21)
        Me.PnlAccounting.Controls.Add(Me.FrmTransaksiJournal_Purchase11)
        Me.PnlAccounting.Controls.Add(Me.FrmTransaksiSalesCycle22)
        Me.PnlAccounting.Controls.Add(Me.FrmTransaksiSalesCycle12)
        Me.PnlAccounting.Controls.Add(Me.FrmTransaksiGeneralJournal1)
        Me.PnlAccounting.Controls.Add(Me.Label13)
        Me.PnlAccounting.Location = New System.Drawing.Point(13, 10)
        Me.PnlAccounting.Name = "PnlAccounting"
        Me.PnlAccounting.Size = New System.Drawing.Size(746, 577)
        Me.PnlAccounting.TabIndex = 18
        '
        'FrmTransaksiJournal_View1
        '
        Me.FrmTransaksiJournal_View1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmTransaksiJournal_View1.Location = New System.Drawing.Point(291, 151)
        Me.FrmTransaksiJournal_View1.Name = "FrmTransaksiJournal_View1"
        Me.FrmTransaksiJournal_View1.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiJournal_View1.TabIndex = 22
        Me.FrmTransaksiJournal_View1.Text = "Journal View"
        Me.FrmTransaksiJournal_View1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiReminder1
        '
        Me.FrmTransaksiReminder1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmTransaksiReminder1.Location = New System.Drawing.Point(291, 274)
        Me.FrmTransaksiReminder1.Name = "FrmTransaksiReminder1"
        Me.FrmTransaksiReminder1.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiReminder1.TabIndex = 21
        Me.FrmTransaksiReminder1.Text = "Reminder"
        Me.FrmTransaksiReminder1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiTagihan1
        '
        Me.FrmTransaksiTagihan1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmTransaksiTagihan1.Location = New System.Drawing.Point(154, 274)
        Me.FrmTransaksiTagihan1.Name = "FrmTransaksiTagihan1"
        Me.FrmTransaksiTagihan1.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiTagihan1.TabIndex = 20
        Me.FrmTransaksiTagihan1.Text = "Collection Letter"
        Me.FrmTransaksiTagihan1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiClosing1
        '
        Me.FrmTransaksiClosing1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmTransaksiClosing1.Location = New System.Drawing.Point(25, 274)
        Me.FrmTransaksiClosing1.Name = "FrmTransaksiClosing1"
        Me.FrmTransaksiClosing1.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiClosing1.TabIndex = 19
        Me.FrmTransaksiClosing1.Text = "Journal Locking"
        Me.FrmTransaksiClosing1.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(22, 251)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(44, 17)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "Other"
        '
        'FrmTransaksiJournal_Closing1
        '
        Me.FrmTransaksiJournal_Closing1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmTransaksiJournal_Closing1.Location = New System.Drawing.Point(154, 151)
        Me.FrmTransaksiJournal_Closing1.Name = "FrmTransaksiJournal_Closing1"
        Me.FrmTransaksiJournal_Closing1.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiJournal_Closing1.TabIndex = 14
        Me.FrmTransaksiJournal_Closing1.Text = "Closing Journal"
        Me.FrmTransaksiJournal_Closing1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiJournal_Adjust1
        '
        Me.FrmTransaksiJournal_Adjust1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmTransaksiJournal_Adjust1.Location = New System.Drawing.Point(25, 151)
        Me.FrmTransaksiJournal_Adjust1.Name = "FrmTransaksiJournal_Adjust1"
        Me.FrmTransaksiJournal_Adjust1.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiJournal_Adjust1.TabIndex = 13
        Me.FrmTransaksiJournal_Adjust1.Text = "Adjustment Journal"
        Me.FrmTransaksiJournal_Adjust1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiJournal_Purchase21
        '
        Me.FrmTransaksiJournal_Purchase21.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmTransaksiJournal_Purchase21.Location = New System.Drawing.Point(556, 46)
        Me.FrmTransaksiJournal_Purchase21.Name = "FrmTransaksiJournal_Purchase21"
        Me.FrmTransaksiJournal_Purchase21.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiJournal_Purchase21.TabIndex = 12
        Me.FrmTransaksiJournal_Purchase21.Text = "Purchase Return Receivement Journal"
        Me.FrmTransaksiJournal_Purchase21.UseVisualStyleBackColor = True
        '
        'FrmTransaksiJournal_Purchase11
        '
        Me.FrmTransaksiJournal_Purchase11.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmTransaksiJournal_Purchase11.Location = New System.Drawing.Point(425, 46)
        Me.FrmTransaksiJournal_Purchase11.Name = "FrmTransaksiJournal_Purchase11"
        Me.FrmTransaksiJournal_Purchase11.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiJournal_Purchase11.TabIndex = 11
        Me.FrmTransaksiJournal_Purchase11.Text = "Purchase Payment Journal"
        Me.FrmTransaksiJournal_Purchase11.UseVisualStyleBackColor = True
        '
        'FrmTransaksiSalesCycle22
        '
        Me.FrmTransaksiSalesCycle22.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmTransaksiSalesCycle22.Location = New System.Drawing.Point(291, 46)
        Me.FrmTransaksiSalesCycle22.Name = "FrmTransaksiSalesCycle22"
        Me.FrmTransaksiSalesCycle22.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiSalesCycle22.TabIndex = 10
        Me.FrmTransaksiSalesCycle22.Text = "Sales Return Payment Journal"
        Me.FrmTransaksiSalesCycle22.UseVisualStyleBackColor = True
        '
        'FrmTransaksiSalesCycle12
        '
        Me.FrmTransaksiSalesCycle12.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmTransaksiSalesCycle12.Location = New System.Drawing.Point(154, 46)
        Me.FrmTransaksiSalesCycle12.Name = "FrmTransaksiSalesCycle12"
        Me.FrmTransaksiSalesCycle12.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiSalesCycle12.TabIndex = 9
        Me.FrmTransaksiSalesCycle12.Text = "Sales Receivement Journal"
        Me.FrmTransaksiSalesCycle12.UseVisualStyleBackColor = True
        '
        'FrmTransaksiGeneralJournal1
        '
        Me.FrmTransaksiGeneralJournal1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FrmTransaksiGeneralJournal1.Location = New System.Drawing.Point(25, 46)
        Me.FrmTransaksiGeneralJournal1.Name = "FrmTransaksiGeneralJournal1"
        Me.FrmTransaksiGeneralJournal1.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiGeneralJournal1.TabIndex = 8
        Me.FrmTransaksiGeneralJournal1.Text = "General Journal"
        Me.FrmTransaksiGeneralJournal1.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(22, 14)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(55, 17)
        Me.Label13.TabIndex = 8
        Me.Label13.Text = "Journal"
        '
        'FrmMainSMESS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(896, 606)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MinimumSize = New System.Drawing.Size(912, 642)
        Me.Name = "FrmMainSMESS"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Main Panel"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.PnlMasterData.ResumeLayout(False)
        Me.PnlMasterData.PerformLayout()
        Me.PnlAccounting.ResumeLayout(False)
        Me.PnlAccounting.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnInternalMessage As System.Windows.Forms.Button
    Friend WithEvents btnMaster As System.Windows.Forms.Button
    Friend WithEvents btnAccounting As System.Windows.Forms.Button
    Friend WithEvents PnlMasterData As System.Windows.Forms.Panel
    Friend WithEvents FrmTransaksiBudget1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterCurrency1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterAkun1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterUnit1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterProduct1 As System.Windows.Forms.Button
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents FrmMasterProductGroup1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterCategory1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterSubBrand1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterBrand1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterCustomerGroup1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterCustomer1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterCustomerSegmentation1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterCustomerType1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterStorageArea1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterGudang1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterSupplier1 As System.Windows.Forms.Button
    Friend WithEvents FrmMasterSalesArea1 As System.Windows.Forms.Button
    Friend WithEvents PnlAccounting As System.Windows.Forms.Panel
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents FrmTransaksiJournal_Closing1 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiJournal_Adjust1 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiJournal_Purchase21 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiJournal_Purchase11 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiSalesCycle22 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiSalesCycle12 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiGeneralJournal1 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiReminder1 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiTagihan1 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiClosing1 As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents FrmTransaksiJournal_View1 As System.Windows.Forms.Button
End Class
