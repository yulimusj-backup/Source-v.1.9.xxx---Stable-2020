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
        Me.btnMaster = New System.Windows.Forms.Button
        Me.btnInternalMessage = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.PnlInventory = New System.Windows.Forms.Panel
        Me.FrmTransaksiStockTakingProcess1 = New System.Windows.Forms.Button
        Me.frmTransaksiStockTag1 = New System.Windows.Forms.Button
        Me.FrmTransaksiStockCard1 = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.frmTransaksiStockTaking1 = New System.Windows.Forms.Button
        Me.FrmTransaksiStockPindah1 = New System.Windows.Forms.Button
        Me.FrmTransaksiStockBundling1 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.FrmTransaksiStockOutAdj1 = New System.Windows.Forms.Button
        Me.FrmTransaksiStockOutPurchaseReturn2 = New System.Windows.Forms.Button
        Me.FrmTransaksiStockOutSales2 = New System.Windows.Forms.Button
        Me.FrmTransaksiStockInAdj1 = New System.Windows.Forms.Button
        Me.FrmTransaksiStockInSalesReturn2 = New System.Windows.Forms.Button
        Me.FrmTransaksiStockInPurchase2 = New System.Windows.Forms.Button
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
        Me.FrmMasterCurrency1 = New System.Windows.Forms.Button
        Me.FrmMasterUnit1 = New System.Windows.Forms.Button
        Me.FrmMasterProduct1 = New System.Windows.Forms.Button
        Me.Label12 = New System.Windows.Forms.Label
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.PnlInventory.SuspendLayout()
        Me.PnlMasterData.SuspendLayout()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnMaster)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnInternalMessage)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Button2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.BackColor = System.Drawing.Color.White
        Me.SplitContainer1.Panel2.Controls.Add(Me.PnlMasterData)
        Me.SplitContainer1.Panel2.Controls.Add(Me.PnlInventory)
        Me.SplitContainer1.Size = New System.Drawing.Size(896, 606)
        Me.SplitContainer1.SplitterDistance = 121
        Me.SplitContainer1.TabIndex = 0
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
        Me.btnInternalMessage.Location = New System.Drawing.Point(16, 178)
        Me.btnInternalMessage.Name = "btnInternalMessage"
        Me.btnInternalMessage.Size = New System.Drawing.Size(89, 77)
        Me.btnInternalMessage.TabIndex = 5
        Me.btnInternalMessage.Text = "Internal Message"
        Me.btnInternalMessage.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button2.Location = New System.Drawing.Point(16, 95)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(89, 77)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Inventory"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'PnlInventory
        '
        Me.PnlInventory.BackColor = System.Drawing.Color.White
        Me.PnlInventory.Controls.Add(Me.FrmTransaksiStockTakingProcess1)
        Me.PnlInventory.Controls.Add(Me.frmTransaksiStockTag1)
        Me.PnlInventory.Controls.Add(Me.FrmTransaksiStockCard1)
        Me.PnlInventory.Controls.Add(Me.Label8)
        Me.PnlInventory.Controls.Add(Me.frmTransaksiStockTaking1)
        Me.PnlInventory.Controls.Add(Me.FrmTransaksiStockPindah1)
        Me.PnlInventory.Controls.Add(Me.FrmTransaksiStockBundling1)
        Me.PnlInventory.Controls.Add(Me.Label3)
        Me.PnlInventory.Controls.Add(Me.Label2)
        Me.PnlInventory.Controls.Add(Me.Label1)
        Me.PnlInventory.Controls.Add(Me.FrmTransaksiStockOutAdj1)
        Me.PnlInventory.Controls.Add(Me.FrmTransaksiStockOutPurchaseReturn2)
        Me.PnlInventory.Controls.Add(Me.FrmTransaksiStockOutSales2)
        Me.PnlInventory.Controls.Add(Me.FrmTransaksiStockInAdj1)
        Me.PnlInventory.Controls.Add(Me.FrmTransaksiStockInSalesReturn2)
        Me.PnlInventory.Controls.Add(Me.FrmTransaksiStockInPurchase2)
        Me.PnlInventory.Location = New System.Drawing.Point(13, 10)
        Me.PnlInventory.Name = "PnlInventory"
        Me.PnlInventory.Size = New System.Drawing.Size(746, 577)
        Me.PnlInventory.TabIndex = 9
        '
        'FrmTransaksiStockTakingProcess1
        '
        Me.FrmTransaksiStockTakingProcess1.Location = New System.Drawing.Point(258, 271)
        Me.FrmTransaksiStockTakingProcess1.Name = "FrmTransaksiStockTakingProcess1"
        Me.FrmTransaksiStockTakingProcess1.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiStockTakingProcess1.TabIndex = 17
        Me.FrmTransaksiStockTakingProcess1.Text = "Report && Data Processing"
        Me.FrmTransaksiStockTakingProcess1.UseVisualStyleBackColor = True
        '
        'frmTransaksiStockTag1
        '
        Me.frmTransaksiStockTag1.Location = New System.Drawing.Point(142, 271)
        Me.frmTransaksiStockTag1.Name = "frmTransaksiStockTag1"
        Me.frmTransaksiStockTag1.Size = New System.Drawing.Size(89, 77)
        Me.frmTransaksiStockTag1.TabIndex = 16
        Me.frmTransaksiStockTag1.Text = "Stock Tag"
        Me.frmTransaksiStockTag1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiStockCard1
        '
        Me.FrmTransaksiStockCard1.Location = New System.Drawing.Point(25, 402)
        Me.FrmTransaksiStockCard1.Name = "FrmTransaksiStockCard1"
        Me.FrmTransaksiStockCard1.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiStockCard1.TabIndex = 15
        Me.FrmTransaksiStockCard1.Text = "Stock Card"
        Me.FrmTransaksiStockCard1.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(22, 251)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(90, 17)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Stock Taking"
        '
        'frmTransaksiStockTaking1
        '
        Me.frmTransaksiStockTaking1.Location = New System.Drawing.Point(25, 271)
        Me.frmTransaksiStockTaking1.Name = "frmTransaksiStockTaking1"
        Me.frmTransaksiStockTaking1.Size = New System.Drawing.Size(89, 77)
        Me.frmTransaksiStockTaking1.TabIndex = 13
        Me.frmTransaksiStockTaking1.Text = "Stock Taking"
        Me.frmTransaksiStockTaking1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiStockPindah1
        '
        Me.FrmTransaksiStockPindah1.Location = New System.Drawing.Point(259, 402)
        Me.FrmTransaksiStockPindah1.Name = "FrmTransaksiStockPindah1"
        Me.FrmTransaksiStockPindah1.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiStockPindah1.TabIndex = 12
        Me.FrmTransaksiStockPindah1.Text = "Location Transfer"
        Me.FrmTransaksiStockPindah1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiStockBundling1
        '
        Me.FrmTransaksiStockBundling1.Location = New System.Drawing.Point(142, 402)
        Me.FrmTransaksiStockBundling1.Name = "FrmTransaksiStockBundling1"
        Me.FrmTransaksiStockBundling1.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiStockBundling1.TabIndex = 11
        Me.FrmTransaksiStockBundling1.Text = "Stock Bundling"
        Me.FrmTransaksiStockBundling1.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(22, 382)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 17)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Others"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(22, 131)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 17)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Stock Out"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(22, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 17)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Stock In"
        '
        'FrmTransaksiStockOutAdj1
        '
        Me.FrmTransaksiStockOutAdj1.Location = New System.Drawing.Point(258, 151)
        Me.FrmTransaksiStockOutAdj1.Name = "FrmTransaksiStockOutAdj1"
        Me.FrmTransaksiStockOutAdj1.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiStockOutAdj1.TabIndex = 7
        Me.FrmTransaksiStockOutAdj1.Text = "Stock Out Adjust"
        Me.FrmTransaksiStockOutAdj1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiStockOutPurchaseReturn2
        '
        Me.FrmTransaksiStockOutPurchaseReturn2.Location = New System.Drawing.Point(142, 151)
        Me.FrmTransaksiStockOutPurchaseReturn2.Name = "FrmTransaksiStockOutPurchaseReturn2"
        Me.FrmTransaksiStockOutPurchaseReturn2.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiStockOutPurchaseReturn2.TabIndex = 6
        Me.FrmTransaksiStockOutPurchaseReturn2.Text = "Stock Out Purchase Return"
        Me.FrmTransaksiStockOutPurchaseReturn2.UseVisualStyleBackColor = True
        '
        'FrmTransaksiStockOutSales2
        '
        Me.FrmTransaksiStockOutSales2.Location = New System.Drawing.Point(25, 151)
        Me.FrmTransaksiStockOutSales2.Name = "FrmTransaksiStockOutSales2"
        Me.FrmTransaksiStockOutSales2.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiStockOutSales2.TabIndex = 5
        Me.FrmTransaksiStockOutSales2.Text = "Stock Out Sales"
        Me.FrmTransaksiStockOutSales2.UseVisualStyleBackColor = True
        '
        'FrmTransaksiStockInAdj1
        '
        Me.FrmTransaksiStockInAdj1.Location = New System.Drawing.Point(258, 34)
        Me.FrmTransaksiStockInAdj1.Name = "FrmTransaksiStockInAdj1"
        Me.FrmTransaksiStockInAdj1.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiStockInAdj1.TabIndex = 4
        Me.FrmTransaksiStockInAdj1.Text = "Stock In Adjust"
        Me.FrmTransaksiStockInAdj1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiStockInSalesReturn2
        '
        Me.FrmTransaksiStockInSalesReturn2.Location = New System.Drawing.Point(142, 34)
        Me.FrmTransaksiStockInSalesReturn2.Name = "FrmTransaksiStockInSalesReturn2"
        Me.FrmTransaksiStockInSalesReturn2.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiStockInSalesReturn2.TabIndex = 2
        Me.FrmTransaksiStockInSalesReturn2.Text = "Stock In Sales Return"
        Me.FrmTransaksiStockInSalesReturn2.UseVisualStyleBackColor = True
        '
        'FrmTransaksiStockInPurchase2
        '
        Me.FrmTransaksiStockInPurchase2.Location = New System.Drawing.Point(25, 34)
        Me.FrmTransaksiStockInPurchase2.Name = "FrmTransaksiStockInPurchase2"
        Me.FrmTransaksiStockInPurchase2.Size = New System.Drawing.Size(89, 77)
        Me.FrmTransaksiStockInPurchase2.TabIndex = 1
        Me.FrmTransaksiStockInPurchase2.Text = "Stock In Purchase"
        Me.FrmTransaksiStockInPurchase2.UseVisualStyleBackColor = True
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
        Me.PnlMasterData.Controls.Add(Me.FrmMasterCurrency1)
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
        Me.PnlInventory.ResumeLayout(False)
        Me.PnlInventory.PerformLayout()
        Me.PnlMasterData.ResumeLayout(False)
        Me.PnlMasterData.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnInternalMessage As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents PnlInventory As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents FrmTransaksiStockOutAdj1 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiStockOutPurchaseReturn2 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiStockOutSales2 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiStockInAdj1 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiStockInSalesReturn2 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiStockInPurchase2 As System.Windows.Forms.Button
    Friend WithEvents btnMaster As System.Windows.Forms.Button
    Friend WithEvents frmTransaksiStockTaking1 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiStockPindah1 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiStockBundling1 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents FrmTransaksiStockTakingProcess1 As System.Windows.Forms.Button
    Friend WithEvents frmTransaksiStockTag1 As System.Windows.Forms.Button
    Friend WithEvents FrmTransaksiStockCard1 As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents PnlMasterData As System.Windows.Forms.Panel
    Friend WithEvents FrmMasterCurrency1 As System.Windows.Forms.Button
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
End Class
