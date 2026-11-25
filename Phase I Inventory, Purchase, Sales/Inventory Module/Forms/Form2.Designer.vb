<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.btnStep1 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.prgStep1 = New System.Windows.Forms.ProgressBar
        Me.LblPrg_Step1 = New System.Windows.Forms.Label
        Me.LblPrgStep2 = New System.Windows.Forms.Label
        Me.prgStep2 = New System.Windows.Forms.ProgressBar
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnStep2 = New System.Windows.Forms.Button
        Me.LblPrgStep3 = New System.Windows.Forms.Label
        Me.PrgStep3 = New System.Windows.Forms.ProgressBar
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnStep3 = New System.Windows.Forms.Button
        Me.LblPrgStep4 = New System.Windows.Forms.Label
        Me.prgStep4 = New System.Windows.Forms.ProgressBar
        Me.Label7 = New System.Windows.Forms.Label
        Me.btnStep4 = New System.Windows.Forms.Button
        Me.bgwStep1 = New System.ComponentModel.BackgroundWorker
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.lblPrgStep1_2a = New System.Windows.Forms.Label
        Me.prb_Step1_2a = New System.Windows.Forms.ProgressBar
        Me.dtpStep13 = New System.Windows.Forms.DateTimePicker
        Me.btnStep13 = New System.Windows.Forms.Button
        Me.Label13 = New System.Windows.Forms.Label
        Me.btnStep12 = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnDelAdj = New System.Windows.Forms.Button
        Me.btnStep9 = New System.Windows.Forms.Button
        Me.lblPrgStep9 = New System.Windows.Forms.Label
        Me.prgStep9 = New System.Windows.Forms.ProgressBar
        Me.Label16 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label11 = New System.Windows.Forms.Label
        Me.btnStep11 = New System.Windows.Forms.Button
        Me.LblPrgStep8 = New System.Windows.Forms.Label
        Me.prgStep8 = New System.Windows.Forms.ProgressBar
        Me.Label12 = New System.Windows.Forms.Label
        Me.btnStep7 = New System.Windows.Forms.Button
        Me.LblPrgStep7 = New System.Windows.Forms.Label
        Me.prgStep7 = New System.Windows.Forms.ProgressBar
        Me.Label9 = New System.Windows.Forms.Label
        Me.dtpStep1 = New System.Windows.Forms.DateTimePicker
        Me.LblPrgStep6 = New System.Windows.Forms.Label
        Me.prgStep6 = New System.Windows.Forms.ProgressBar
        Me.Label10 = New System.Windows.Forms.Label
        Me.LblPrgStep5 = New System.Windows.Forms.Label
        Me.prgStep5 = New System.Windows.Forms.ProgressBar
        Me.Label8 = New System.Windows.Forms.Label
        Me.dtpResetStock = New System.Windows.Forms.DateTimePicker
        Me.btnStep1_2 = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.prb_Step1_2 = New System.Windows.Forms.ProgressBar
        Me.lblPrgStep1_2 = New System.Windows.Forms.Label
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.btnShowOverhead = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpPeriodOverhead = New System.Windows.Forms.DateTimePicker
        Me.dgvOverhead = New System.Windows.Forms.DataGridView
        Me.ColOv_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOv_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOv_RefUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOv_AccCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOv_AccDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOv_Balance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOv_AllocatedPerc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.dgvStock = New System.Windows.Forms.DataGridView
        Me.ColStock_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStock_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStock_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStock_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStock_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStock_Batch = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStock_Expiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.dgvStockMinus = New System.Windows.Forms.DataGridView
        Me.ColDetailMinus_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetailMinus_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetailMinus_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetailMinus_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetailMinus_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetailMinus_Batch = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetailMinus_Expiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage6 = New System.Windows.Forms.TabPage
        Me.dgvStockTidakCocok = New System.Windows.Forms.DataGridView
        Me.ColTdkCocok_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTdkCocok_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTdkCocok_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage7 = New System.Windows.Forms.TabPage
        Me.dgvCekSC = New System.Windows.Forms.DataGridView
        Me.ColCekSC_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCekSC_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCekSC_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCekSC_BatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCekSC_Location = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCekSC_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCekSC_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage8 = New System.Windows.Forms.TabPage
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.dtp_Step8_4 = New System.Windows.Forms.DateTimePicker
        Me.dgvAdjustment = New System.Windows.Forms.DataGridView
        Me.ColDGVAdjust_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDGVAdjust_DocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDGVAdjust_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDGVAdjust_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDGVAdjust_Type = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDGVAdjust_CreatedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDGVAdjust_CreatedDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDGVAdjust_ModifiedDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDGVAdjust_Include = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDGVAdjust_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnStep8_4 = New System.Windows.Forms.Button
        Me.Label17 = New System.Windows.Forms.Label
        Me.btnStep8_3 = New System.Windows.Forms.Button
        Me.Label22 = New System.Windows.Forms.Label
        Me.btnStep8_2 = New System.Windows.Forms.Button
        Me.Label20 = New System.Windows.Forms.Label
        Me.btnStep8_1 = New System.Windows.Forms.Button
        Me.Label18 = New System.Windows.Forms.Label
        Me.TabControl3 = New System.Windows.Forms.TabControl
        Me.TabPage9 = New System.Windows.Forms.TabPage
        Me.dgvAdjustmentItem1 = New System.Windows.Forms.DataGridView
        Me.ColAdjItem1_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdjItem1_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdjItem1_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdjItem1_Location = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdjItem1_Batch = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdjItem1_Expiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdjItem1_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdjItem1_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdjItem1_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage10 = New System.Windows.Forms.TabPage
        Me.dgvAdjustmentItem2 = New System.Windows.Forms.DataGridView
        Me.ColItemAdj2_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAdj2_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAdj2_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAdj2_Location = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAdj2_BatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAdj2_Expiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAdj2_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAdj2_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage11 = New System.Windows.Forms.TabPage
        Me.dgvFinal = New System.Windows.Forms.DataGridView
        Me.ColFinal_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_GroupCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_GroupDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_BrandCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_BrandDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_SubBrandCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_SubBrandDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_CategoryCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_CategoryDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_Qtym3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_Warehouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_WarehouseDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_UnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_Amouunt = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_AccCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFinal_AccDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage12 = New System.Windows.Forms.TabPage
        Me.btnStep20 = New System.Windows.Forms.Button
        Me.Label19 = New System.Windows.Forms.Label
        Me.btnStep1_RemoveMaintenance = New System.Windows.Forms.Button
        Me.btnStep19_Set = New System.Windows.Forms.Button
        Me.Label25 = New System.Windows.Forms.Label
        Me.prgStep18 = New System.Windows.Forms.ProgressBar
        Me.LblPrgStep18 = New System.Windows.Forms.Label
        Me.btnStep18 = New System.Windows.Forms.Button
        Me.Label24 = New System.Windows.Forms.Label
        Me.prgStep17 = New System.Windows.Forms.ProgressBar
        Me.LblPrgStep17 = New System.Windows.Forms.Label
        Me.btnStep17 = New System.Windows.Forms.Button
        Me.Label23 = New System.Windows.Forms.Label
        Me.prgStep16 = New System.Windows.Forms.ProgressBar
        Me.LblPrgStep16 = New System.Windows.Forms.Label
        Me.btnStep16 = New System.Windows.Forms.Button
        Me.Label21 = New System.Windows.Forms.Label
        Me.prgStep15 = New System.Windows.Forms.ProgressBar
        Me.LblPrgStep15 = New System.Windows.Forms.Label
        Me.btnStep15 = New System.Windows.Forms.Button
        Me.Label15 = New System.Windows.Forms.Label
        Me.dtpStep14 = New System.Windows.Forms.DateTimePicker
        Me.btnStep14 = New System.Windows.Forms.Button
        Me.Label14 = New System.Windows.Forms.Label
        Me.bgwStep1_2 = New System.ComponentModel.BackgroundWorker
        Me.bgwStep2 = New System.ComponentModel.BackgroundWorker
        Me.bgwStep3 = New System.ComponentModel.BackgroundWorker
        Me.bgwStep4 = New System.ComponentModel.BackgroundWorker
        Me.bgwStep5 = New System.ComponentModel.BackgroundWorker
        Me.bgwStep6 = New System.ComponentModel.BackgroundWorker
        Me.bgwStep15 = New System.ComponentModel.BackgroundWorker
        Me.bgwStep16 = New System.ComponentModel.BackgroundWorker
        Me.bgwStep17 = New System.ComponentModel.BackgroundWorker
        Me.TabPage13 = New System.Windows.Forms.TabPage
        Me.Label26 = New System.Windows.Forms.Label
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.btnCek_Bundling = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvOverhead, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        CType(Me.dgvStock, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage5.SuspendLayout()
        CType(Me.dgvStockMinus, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage6.SuspendLayout()
        CType(Me.dgvStockTidakCocok, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage7.SuspendLayout()
        CType(Me.dgvCekSC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage8.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvAdjustment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl3.SuspendLayout()
        Me.TabPage9.SuspendLayout()
        CType(Me.dgvAdjustmentItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage10.SuspendLayout()
        CType(Me.dgvAdjustmentItem2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage11.SuspendLayout()
        CType(Me.dgvFinal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage12.SuspendLayout()
        Me.TabPage13.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnStep1
        '
        Me.btnStep1.Location = New System.Drawing.Point(696, 6)
        Me.btnStep1.Name = "btnStep1"
        Me.btnStep1.Size = New System.Drawing.Size(75, 23)
        Me.btnStep1.TabIndex = 0
        Me.btnStep1.Text = "Step 1"
        Me.btnStep1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(235, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Step 1: pindahkan tanggal transaksi"
        '
        'prgStep1
        '
        Me.prgStep1.Location = New System.Drawing.Point(406, 6)
        Me.prgStep1.Name = "prgStep1"
        Me.prgStep1.Size = New System.Drawing.Size(284, 23)
        Me.prgStep1.TabIndex = 2
        '
        'LblPrg_Step1
        '
        Me.LblPrg_Step1.AutoSize = True
        Me.LblPrg_Step1.Location = New System.Drawing.Point(403, 32)
        Me.LblPrg_Step1.Name = "LblPrg_Step1"
        Me.LblPrg_Step1.Size = New System.Drawing.Size(59, 17)
        Me.LblPrg_Step1.TabIndex = 3
        Me.LblPrg_Step1.Text = "pending"
        '
        'LblPrgStep2
        '
        Me.LblPrgStep2.AutoSize = True
        Me.LblPrgStep2.Location = New System.Drawing.Point(403, 124)
        Me.LblPrgStep2.Name = "LblPrgStep2"
        Me.LblPrgStep2.Size = New System.Drawing.Size(59, 17)
        Me.LblPrgStep2.TabIndex = 7
        Me.LblPrgStep2.Text = "pending"
        '
        'prgStep2
        '
        Me.prgStep2.Location = New System.Drawing.Point(406, 98)
        Me.prgStep2.Name = "prgStep2"
        Me.prgStep2.Size = New System.Drawing.Size(284, 23)
        Me.prgStep2.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 101)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(293, 17)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Step 2: buat stock out adjust u/ stock positive"
        '
        'btnStep2
        '
        Me.btnStep2.Location = New System.Drawing.Point(696, 98)
        Me.btnStep2.Name = "btnStep2"
        Me.btnStep2.Size = New System.Drawing.Size(75, 23)
        Me.btnStep2.TabIndex = 4
        Me.btnStep2.Text = "Step 2"
        Me.btnStep2.UseVisualStyleBackColor = True
        '
        'LblPrgStep3
        '
        Me.LblPrgStep3.AutoSize = True
        Me.LblPrgStep3.Location = New System.Drawing.Point(403, 170)
        Me.LblPrgStep3.Name = "LblPrgStep3"
        Me.LblPrgStep3.Size = New System.Drawing.Size(59, 17)
        Me.LblPrgStep3.TabIndex = 11
        Me.LblPrgStep3.Text = "pending"
        '
        'PrgStep3
        '
        Me.PrgStep3.Location = New System.Drawing.Point(406, 144)
        Me.PrgStep3.Name = "PrgStep3"
        Me.PrgStep3.Size = New System.Drawing.Size(284, 23)
        Me.PrgStep3.TabIndex = 10
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 147)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(290, 17)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Step 3: buat stock in adjust u/ stock negative"
        '
        'btnStep3
        '
        Me.btnStep3.Location = New System.Drawing.Point(696, 144)
        Me.btnStep3.Name = "btnStep3"
        Me.btnStep3.Size = New System.Drawing.Size(75, 23)
        Me.btnStep3.TabIndex = 8
        Me.btnStep3.Text = "Step 3"
        Me.btnStep3.UseVisualStyleBackColor = True
        '
        'LblPrgStep4
        '
        Me.LblPrgStep4.AutoSize = True
        Me.LblPrgStep4.Location = New System.Drawing.Point(403, 216)
        Me.LblPrgStep4.Name = "LblPrgStep4"
        Me.LblPrgStep4.Size = New System.Drawing.Size(59, 17)
        Me.LblPrgStep4.TabIndex = 15
        Me.LblPrgStep4.Text = "pending"
        '
        'prgStep4
        '
        Me.prgStep4.Location = New System.Drawing.Point(406, 190)
        Me.prgStep4.Name = "prgStep4"
        Me.prgStep4.Size = New System.Drawing.Size(284, 23)
        Me.prgStep4.TabIndex = 14
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 193)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(370, 17)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "Step 4: buat stock in adjust dr positive, sejumlah negative"
        '
        'btnStep4
        '
        Me.btnStep4.Location = New System.Drawing.Point(696, 190)
        Me.btnStep4.Name = "btnStep4"
        Me.btnStep4.Size = New System.Drawing.Size(75, 23)
        Me.btnStep4.TabIndex = 12
        Me.btnStep4.Text = "Step 4"
        Me.btnStep4.UseVisualStyleBackColor = True
        '
        'bgwStep1
        '
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage7)
        Me.TabControl1.Controls.Add(Me.TabPage8)
        Me.TabControl1.Controls.Add(Me.TabPage11)
        Me.TabControl1.Controls.Add(Me.TabPage12)
        Me.TabControl1.Controls.Add(Me.TabPage13)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1085, 609)
        Me.TabControl1.TabIndex = 16
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.lblPrgStep1_2a)
        Me.TabPage1.Controls.Add(Me.prb_Step1_2a)
        Me.TabPage1.Controls.Add(Me.dtpStep13)
        Me.TabPage1.Controls.Add(Me.btnStep13)
        Me.TabPage1.Controls.Add(Me.Label13)
        Me.TabPage1.Controls.Add(Me.btnStep12)
        Me.TabPage1.Controls.Add(Me.Label6)
        Me.TabPage1.Controls.Add(Me.btnDelAdj)
        Me.TabPage1.Controls.Add(Me.btnStep9)
        Me.TabPage1.Controls.Add(Me.lblPrgStep9)
        Me.TabPage1.Controls.Add(Me.prgStep9)
        Me.TabPage1.Controls.Add(Me.Label16)
        Me.TabPage1.Controls.Add(Me.Button1)
        Me.TabPage1.Controls.Add(Me.Label11)
        Me.TabPage1.Controls.Add(Me.btnStep11)
        Me.TabPage1.Controls.Add(Me.LblPrgStep8)
        Me.TabPage1.Controls.Add(Me.prgStep8)
        Me.TabPage1.Controls.Add(Me.Label12)
        Me.TabPage1.Controls.Add(Me.btnStep7)
        Me.TabPage1.Controls.Add(Me.LblPrgStep7)
        Me.TabPage1.Controls.Add(Me.prgStep7)
        Me.TabPage1.Controls.Add(Me.Label9)
        Me.TabPage1.Controls.Add(Me.dtpStep1)
        Me.TabPage1.Controls.Add(Me.LblPrgStep6)
        Me.TabPage1.Controls.Add(Me.prgStep6)
        Me.TabPage1.Controls.Add(Me.Label10)
        Me.TabPage1.Controls.Add(Me.LblPrgStep5)
        Me.TabPage1.Controls.Add(Me.prgStep5)
        Me.TabPage1.Controls.Add(Me.Label8)
        Me.TabPage1.Controls.Add(Me.dtpResetStock)
        Me.TabPage1.Controls.Add(Me.btnStep1_2)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.prb_Step1_2)
        Me.TabPage1.Controls.Add(Me.lblPrgStep1_2)
        Me.TabPage1.Controls.Add(Me.prgStep1)
        Me.TabPage1.Controls.Add(Me.LblPrgStep4)
        Me.TabPage1.Controls.Add(Me.btnStep1)
        Me.TabPage1.Controls.Add(Me.prgStep4)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.Label7)
        Me.TabPage1.Controls.Add(Me.LblPrg_Step1)
        Me.TabPage1.Controls.Add(Me.btnStep4)
        Me.TabPage1.Controls.Add(Me.btnStep2)
        Me.TabPage1.Controls.Add(Me.LblPrgStep3)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.PrgStep3)
        Me.TabPage1.Controls.Add(Me.prgStep2)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.LblPrgStep2)
        Me.TabPage1.Controls.Add(Me.btnStep3)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1077, 580)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Update Data"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'lblPrgStep1_2a
        '
        Me.lblPrgStep1_2a.AutoSize = True
        Me.lblPrgStep1_2a.Location = New System.Drawing.Point(403, 78)
        Me.lblPrgStep1_2a.Name = "lblPrgStep1_2a"
        Me.lblPrgStep1_2a.Size = New System.Drawing.Size(59, 17)
        Me.lblPrgStep1_2a.TabIndex = 59
        Me.lblPrgStep1_2a.Text = "pending"
        '
        'prb_Step1_2a
        '
        Me.prb_Step1_2a.Location = New System.Drawing.Point(406, 52)
        Me.prb_Step1_2a.Name = "prb_Step1_2a"
        Me.prb_Step1_2a.Size = New System.Drawing.Size(284, 23)
        Me.prb_Step1_2a.TabIndex = 58
        '
        'dtpStep13
        '
        Me.dtpStep13.CustomFormat = "MMM yyyy"
        Me.dtpStep13.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStep13.Location = New System.Drawing.Point(406, 522)
        Me.dtpStep13.Name = "dtpStep13"
        Me.dtpStep13.ShowUpDown = True
        Me.dtpStep13.Size = New System.Drawing.Size(170, 22)
        Me.dtpStep13.TabIndex = 55
        '
        'btnStep13
        '
        Me.btnStep13.Location = New System.Drawing.Point(696, 524)
        Me.btnStep13.Name = "btnStep13"
        Me.btnStep13.Size = New System.Drawing.Size(75, 23)
        Me.btnStep13.TabIndex = 54
        Me.btnStep13.Text = "Step 13"
        Me.btnStep13.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(8, 527)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(247, 17)
        Me.Label13.TabIndex = 53
        Me.Label13.Text = "Step 13: cek end amount hasil closing"
        '
        'btnStep12
        '
        Me.btnStep12.Location = New System.Drawing.Point(696, 495)
        Me.btnStep12.Name = "btnStep12"
        Me.btnStep12.Size = New System.Drawing.Size(75, 23)
        Me.btnStep12.TabIndex = 52
        Me.btnStep12.Text = "Step 12"
        Me.btnStep12.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(8, 498)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(198, 17)
        Me.Label6.TabIndex = 51
        Me.Label6.Text = "Step 12: bundling per account"
        '
        'btnDelAdj
        '
        Me.btnDelAdj.Location = New System.Drawing.Point(300, 98)
        Me.btnDelAdj.Name = "btnDelAdj"
        Me.btnDelAdj.Size = New System.Drawing.Size(100, 23)
        Me.btnDelAdj.TabIndex = 50
        Me.btnDelAdj.Text = "Delete"
        Me.btnDelAdj.UseVisualStyleBackColor = True
        '
        'btnStep9
        '
        Me.btnStep9.Location = New System.Drawing.Point(696, 403)
        Me.btnStep9.Name = "btnStep9"
        Me.btnStep9.Size = New System.Drawing.Size(75, 23)
        Me.btnStep9.TabIndex = 49
        Me.btnStep9.Text = "Step 9"
        Me.btnStep9.UseVisualStyleBackColor = True
        '
        'lblPrgStep9
        '
        Me.lblPrgStep9.AutoSize = True
        Me.lblPrgStep9.Location = New System.Drawing.Point(403, 429)
        Me.lblPrgStep9.Name = "lblPrgStep9"
        Me.lblPrgStep9.Size = New System.Drawing.Size(59, 17)
        Me.lblPrgStep9.TabIndex = 48
        Me.lblPrgStep9.Text = "pending"
        '
        'prgStep9
        '
        Me.prgStep9.Location = New System.Drawing.Point(406, 403)
        Me.prgStep9.Name = "prgStep9"
        Me.prgStep9.Size = New System.Drawing.Size(284, 23)
        Me.prgStep9.TabIndex = 47
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(8, 406)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(359, 17)
        Me.Label16.TabIndex = 46
        Me.Label16.Text = "Step 9: generate laporan stock end 2017 - kirim pak leo"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(696, 374)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 41
        Me.Button1.Text = "Step 8"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(8, 377)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(269, 17)
        Me.Label11.TabIndex = 38
        Me.Label11.Text = "Step 8: cari data adjustment closing 2017"
        '
        'btnStep11
        '
        Me.btnStep11.Location = New System.Drawing.Point(696, 449)
        Me.btnStep11.Name = "btnStep11"
        Me.btnStep11.Size = New System.Drawing.Size(75, 23)
        Me.btnStep11.TabIndex = 37
        Me.btnStep11.Text = "Step 11"
        Me.btnStep11.UseVisualStyleBackColor = True
        '
        'LblPrgStep8
        '
        Me.LblPrgStep8.AutoSize = True
        Me.LblPrgStep8.Location = New System.Drawing.Point(403, 475)
        Me.LblPrgStep8.Name = "LblPrgStep8"
        Me.LblPrgStep8.Size = New System.Drawing.Size(59, 17)
        Me.LblPrgStep8.TabIndex = 36
        Me.LblPrgStep8.Text = "pending"
        '
        'prgStep8
        '
        Me.prgStep8.Location = New System.Drawing.Point(406, 449)
        Me.prgStep8.Name = "prgStep8"
        Me.prgStep8.Size = New System.Drawing.Size(284, 23)
        Me.prgStep8.TabIndex = 35
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(8, 452)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(155, 17)
        Me.Label12.TabIndex = 34
        Me.Label12.Text = "Step 11: hitung amount"
        '
        'btnStep7
        '
        Me.btnStep7.Location = New System.Drawing.Point(696, 328)
        Me.btnStep7.Name = "btnStep7"
        Me.btnStep7.Size = New System.Drawing.Size(75, 23)
        Me.btnStep7.TabIndex = 33
        Me.btnStep7.Text = "Step 7"
        Me.btnStep7.UseVisualStyleBackColor = True
        '
        'LblPrgStep7
        '
        Me.LblPrgStep7.AutoSize = True
        Me.LblPrgStep7.Location = New System.Drawing.Point(403, 354)
        Me.LblPrgStep7.Name = "LblPrgStep7"
        Me.LblPrgStep7.Size = New System.Drawing.Size(59, 17)
        Me.LblPrgStep7.TabIndex = 32
        Me.LblPrgStep7.Text = "pending"
        '
        'prgStep7
        '
        Me.prgStep7.Location = New System.Drawing.Point(406, 328)
        Me.prgStep7.Name = "prgStep7"
        Me.prgStep7.Size = New System.Drawing.Size(284, 23)
        Me.prgStep7.TabIndex = 31
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(8, 331)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(126, 17)
        Me.Label9.TabIndex = 30
        Me.Label9.Text = "Step 7: data check"
        '
        'dtpStep1
        '
        Me.dtpStep1.CustomFormat = "dd MMM yyyy"
        Me.dtpStep1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStep1.Location = New System.Drawing.Point(249, 7)
        Me.dtpStep1.Name = "dtpStep1"
        Me.dtpStep1.Size = New System.Drawing.Size(151, 22)
        Me.dtpStep1.TabIndex = 29
        '
        'LblPrgStep6
        '
        Me.LblPrgStep6.AutoSize = True
        Me.LblPrgStep6.Location = New System.Drawing.Point(403, 308)
        Me.LblPrgStep6.Name = "LblPrgStep6"
        Me.LblPrgStep6.Size = New System.Drawing.Size(59, 17)
        Me.LblPrgStep6.TabIndex = 28
        Me.LblPrgStep6.Text = "pending"
        '
        'prgStep6
        '
        Me.prgStep6.Location = New System.Drawing.Point(406, 282)
        Me.prgStep6.Name = "prgStep6"
        Me.prgStep6.Size = New System.Drawing.Size(284, 23)
        Me.prgStep6.TabIndex = 27
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(8, 285)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(284, 17)
        Me.Label10.TabIndex = 26
        Me.Label10.Text = "Step 6: stock in adjust dari sisa positive-nya"
        '
        'LblPrgStep5
        '
        Me.LblPrgStep5.AutoSize = True
        Me.LblPrgStep5.Location = New System.Drawing.Point(403, 262)
        Me.LblPrgStep5.Name = "LblPrgStep5"
        Me.LblPrgStep5.Size = New System.Drawing.Size(59, 17)
        Me.LblPrgStep5.TabIndex = 24
        Me.LblPrgStep5.Text = "pending"
        '
        'prgStep5
        '
        Me.prgStep5.Location = New System.Drawing.Point(406, 236)
        Me.prgStep5.Name = "prgStep5"
        Me.prgStep5.Size = New System.Drawing.Size(284, 23)
        Me.prgStep5.TabIndex = 23
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 239)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(262, 17)
        Me.Label8.TabIndex = 22
        Me.Label8.Text = "Step 5: buat stock out adjust u/ negative"
        '
        'dtpResetStock
        '
        Me.dtpResetStock.CustomFormat = "dd MMM yyyy"
        Me.dtpResetStock.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpResetStock.Location = New System.Drawing.Point(249, 52)
        Me.dtpResetStock.Name = "dtpResetStock"
        Me.dtpResetStock.Size = New System.Drawing.Size(151, 22)
        Me.dtpResetStock.TabIndex = 20
        '
        'btnStep1_2
        '
        Me.btnStep1_2.Location = New System.Drawing.Point(696, 52)
        Me.btnStep1_2.Name = "btnStep1_2"
        Me.btnStep1_2.Size = New System.Drawing.Size(75, 23)
        Me.btnStep1_2.TabIndex = 16
        Me.btnStep1_2.Text = "Step 1.2"
        Me.btnStep1_2.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 55)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(189, 17)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "Step 1.2: tentukan end stock"
        '
        'prb_Step1_2
        '
        Me.prb_Step1_2.Location = New System.Drawing.Point(777, 52)
        Me.prb_Step1_2.Name = "prb_Step1_2"
        Me.prb_Step1_2.Size = New System.Drawing.Size(284, 23)
        Me.prb_Step1_2.TabIndex = 18
        '
        'lblPrgStep1_2
        '
        Me.lblPrgStep1_2.AutoSize = True
        Me.lblPrgStep1_2.Location = New System.Drawing.Point(774, 78)
        Me.lblPrgStep1_2.Name = "lblPrgStep1_2"
        Me.lblPrgStep1_2.Size = New System.Drawing.Size(59, 17)
        Me.lblPrgStep1_2.TabIndex = 19
        Me.lblPrgStep1_2.Text = "pending"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.btnShowOverhead)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.dtpPeriodOverhead)
        Me.TabPage2.Controls.Add(Me.dgvOverhead)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1077, 580)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Detail Overhead"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'btnShowOverhead
        '
        Me.btnShowOverhead.Location = New System.Drawing.Point(196, 6)
        Me.btnShowOverhead.Name = "btnShowOverhead"
        Me.btnShowOverhead.Size = New System.Drawing.Size(75, 23)
        Me.btnShowOverhead.TabIndex = 19
        Me.btnShowOverhead.Text = "Show"
        Me.btnShowOverhead.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 17)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Period"
        '
        'dtpPeriodOverhead
        '
        Me.dtpPeriodOverhead.CustomFormat = "MMM yyyy"
        Me.dtpPeriodOverhead.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPeriodOverhead.Location = New System.Drawing.Point(61, 6)
        Me.dtpPeriodOverhead.Name = "dtpPeriodOverhead"
        Me.dtpPeriodOverhead.ShowUpDown = True
        Me.dtpPeriodOverhead.Size = New System.Drawing.Size(129, 22)
        Me.dtpPeriodOverhead.TabIndex = 17
        '
        'dgvOverhead
        '
        Me.dgvOverhead.AllowUserToAddRows = False
        Me.dgvOverhead.AllowUserToDeleteRows = False
        Me.dgvOverhead.AllowUserToResizeColumns = False
        Me.dgvOverhead.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvOverhead.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvOverhead.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        Me.dgvOverhead.ColumnHeadersHeight = 25
        Me.dgvOverhead.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvOverhead.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColOv_No, Me.ColOv_Desc, Me.ColOv_RefUnit, Me.ColOv_AccCode, Me.ColOv_AccDesc, Me.ColOv_Balance, Me.ColOv_AllocatedPerc})
        Me.dgvOverhead.Location = New System.Drawing.Point(6, 35)
        Me.dgvOverhead.Name = "dgvOverhead"
        Me.dgvOverhead.RowHeadersVisible = False
        Me.dgvOverhead.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvOverhead.RowTemplate.Height = 24
        Me.dgvOverhead.Size = New System.Drawing.Size(913, 539)
        Me.dgvOverhead.TabIndex = 0
        '
        'ColOv_No
        '
        Me.ColOv_No.HeaderText = "No"
        Me.ColOv_No.Name = "ColOv_No"
        Me.ColOv_No.Width = 51
        '
        'ColOv_Desc
        '
        Me.ColOv_Desc.HeaderText = "Description"
        Me.ColOv_Desc.Name = "ColOv_Desc"
        Me.ColOv_Desc.Width = 104
        '
        'ColOv_RefUnit
        '
        Me.ColOv_RefUnit.HeaderText = "RefUnit"
        Me.ColOv_RefUnit.Name = "ColOv_RefUnit"
        Me.ColOv_RefUnit.Width = 80
        '
        'ColOv_AccCode
        '
        Me.ColOv_AccCode.HeaderText = "Acc. Code"
        Me.ColOv_AccCode.Name = "ColOv_AccCode"
        Me.ColOv_AccCode.Width = 97
        '
        'ColOv_AccDesc
        '
        Me.ColOv_AccDesc.HeaderText = "Acc. Desc"
        Me.ColOv_AccDesc.Name = "ColOv_AccDesc"
        Me.ColOv_AccDesc.Width = 96
        '
        'ColOv_Balance
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColOv_Balance.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColOv_Balance.HeaderText = "Balance"
        Me.ColOv_Balance.Name = "ColOv_Balance"
        Me.ColOv_Balance.Width = 84
        '
        'ColOv_AllocatedPerc
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColOv_AllocatedPerc.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColOv_AllocatedPerc.HeaderText = "Allocated %"
        Me.ColOv_AllocatedPerc.Name = "ColOv_AllocatedPerc"
        Me.ColOv_AllocatedPerc.Width = 107
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.TabControl2)
        Me.TabPage3.Location = New System.Drawing.Point(4, 25)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(1077, 580)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Detail Stock"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'TabControl2
        '
        Me.TabControl2.Controls.Add(Me.TabPage4)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Controls.Add(Me.TabPage6)
        Me.TabControl2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl2.Location = New System.Drawing.Point(3, 3)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(1071, 574)
        Me.TabControl2.TabIndex = 2
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.dgvStock)
        Me.TabPage4.Location = New System.Drawing.Point(4, 25)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(1063, 545)
        Me.TabPage4.TabIndex = 0
        Me.TabPage4.Text = "Stock"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'dgvStock
        '
        Me.dgvStock.AllowUserToAddRows = False
        Me.dgvStock.AllowUserToDeleteRows = False
        Me.dgvStock.AllowUserToResizeColumns = False
        Me.dgvStock.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvStock.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        Me.dgvStock.ColumnHeadersHeight = 25
        Me.dgvStock.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvStock.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColStock_No, Me.ColStock_CodeNo, Me.ColStock_Desc, Me.ColStock_Qty, Me.ColStock_Unit, Me.ColStock_Batch, Me.ColStock_Expiry})
        Me.dgvStock.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvStock.Location = New System.Drawing.Point(3, 3)
        Me.dgvStock.Name = "dgvStock"
        Me.dgvStock.RowHeadersVisible = False
        Me.dgvStock.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvStock.RowTemplate.Height = 24
        Me.dgvStock.Size = New System.Drawing.Size(1057, 539)
        Me.dgvStock.TabIndex = 2
        '
        'ColStock_No
        '
        Me.ColStock_No.HeaderText = "No"
        Me.ColStock_No.Name = "ColStock_No"
        Me.ColStock_No.Width = 51
        '
        'ColStock_CodeNo
        '
        Me.ColStock_CodeNo.HeaderText = "Code No"
        Me.ColStock_CodeNo.Name = "ColStock_CodeNo"
        Me.ColStock_CodeNo.Width = 88
        '
        'ColStock_Desc
        '
        Me.ColStock_Desc.HeaderText = "Description"
        Me.ColStock_Desc.Name = "ColStock_Desc"
        Me.ColStock_Desc.Width = 104
        '
        'ColStock_Qty
        '
        Me.ColStock_Qty.HeaderText = "Qty"
        Me.ColStock_Qty.Name = "ColStock_Qty"
        Me.ColStock_Qty.Width = 55
        '
        'ColStock_Unit
        '
        Me.ColStock_Unit.HeaderText = "Unit"
        Me.ColStock_Unit.Name = "ColStock_Unit"
        Me.ColStock_Unit.Width = 58
        '
        'ColStock_Batch
        '
        Me.ColStock_Batch.HeaderText = "Batch"
        Me.ColStock_Batch.Name = "ColStock_Batch"
        Me.ColStock_Batch.Width = 69
        '
        'ColStock_Expiry
        '
        Me.ColStock_Expiry.HeaderText = "Expiry"
        Me.ColStock_Expiry.Name = "ColStock_Expiry"
        Me.ColStock_Expiry.Width = 71
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.dgvStockMinus)
        Me.TabPage5.Location = New System.Drawing.Point(4, 25)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(1063, 545)
        Me.TabPage5.TabIndex = 1
        Me.TabPage5.Text = "Detail Minus"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'dgvStockMinus
        '
        Me.dgvStockMinus.AllowUserToAddRows = False
        Me.dgvStockMinus.AllowUserToDeleteRows = False
        Me.dgvStockMinus.AllowUserToResizeColumns = False
        Me.dgvStockMinus.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvStockMinus.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        Me.dgvStockMinus.ColumnHeadersHeight = 25
        Me.dgvStockMinus.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvStockMinus.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDetailMinus_No, Me.ColDetailMinus_CodeNo, Me.ColDetailMinus_Desc, Me.ColDetailMinus_Qty, Me.ColDetailMinus_Unit, Me.ColDetailMinus_Batch, Me.ColDetailMinus_Expiry})
        Me.dgvStockMinus.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvStockMinus.Location = New System.Drawing.Point(3, 3)
        Me.dgvStockMinus.Name = "dgvStockMinus"
        Me.dgvStockMinus.RowHeadersVisible = False
        Me.dgvStockMinus.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvStockMinus.RowTemplate.Height = 24
        Me.dgvStockMinus.Size = New System.Drawing.Size(1057, 539)
        Me.dgvStockMinus.TabIndex = 3
        '
        'ColDetailMinus_No
        '
        Me.ColDetailMinus_No.HeaderText = "No"
        Me.ColDetailMinus_No.Name = "ColDetailMinus_No"
        Me.ColDetailMinus_No.Width = 51
        '
        'ColDetailMinus_CodeNo
        '
        Me.ColDetailMinus_CodeNo.HeaderText = "Code No"
        Me.ColDetailMinus_CodeNo.Name = "ColDetailMinus_CodeNo"
        Me.ColDetailMinus_CodeNo.Width = 88
        '
        'ColDetailMinus_Desc
        '
        Me.ColDetailMinus_Desc.HeaderText = "Description"
        Me.ColDetailMinus_Desc.Name = "ColDetailMinus_Desc"
        Me.ColDetailMinus_Desc.Width = 104
        '
        'ColDetailMinus_Qty
        '
        Me.ColDetailMinus_Qty.HeaderText = "Qty"
        Me.ColDetailMinus_Qty.Name = "ColDetailMinus_Qty"
        Me.ColDetailMinus_Qty.Width = 55
        '
        'ColDetailMinus_Unit
        '
        Me.ColDetailMinus_Unit.HeaderText = "Unit"
        Me.ColDetailMinus_Unit.Name = "ColDetailMinus_Unit"
        Me.ColDetailMinus_Unit.Width = 58
        '
        'ColDetailMinus_Batch
        '
        Me.ColDetailMinus_Batch.HeaderText = "Batch"
        Me.ColDetailMinus_Batch.Name = "ColDetailMinus_Batch"
        Me.ColDetailMinus_Batch.Width = 69
        '
        'ColDetailMinus_Expiry
        '
        Me.ColDetailMinus_Expiry.HeaderText = "Expiry"
        Me.ColDetailMinus_Expiry.Name = "ColDetailMinus_Expiry"
        Me.ColDetailMinus_Expiry.Width = 71
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.dgvStockTidakCocok)
        Me.TabPage6.Location = New System.Drawing.Point(4, 25)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage6.Size = New System.Drawing.Size(1063, 545)
        Me.TabPage6.TabIndex = 2
        Me.TabPage6.Text = "Detail Tidak Cocok"
        Me.TabPage6.UseVisualStyleBackColor = True
        '
        'dgvStockTidakCocok
        '
        Me.dgvStockTidakCocok.AllowUserToAddRows = False
        Me.dgvStockTidakCocok.AllowUserToDeleteRows = False
        Me.dgvStockTidakCocok.AllowUserToResizeColumns = False
        Me.dgvStockTidakCocok.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvStockTidakCocok.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        Me.dgvStockTidakCocok.ColumnHeadersHeight = 25
        Me.dgvStockTidakCocok.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvStockTidakCocok.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColTdkCocok_No, Me.ColTdkCocok_CodeNo, Me.ColTdkCocok_Desc})
        Me.dgvStockTidakCocok.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvStockTidakCocok.Location = New System.Drawing.Point(3, 3)
        Me.dgvStockTidakCocok.Name = "dgvStockTidakCocok"
        Me.dgvStockTidakCocok.RowHeadersVisible = False
        Me.dgvStockTidakCocok.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvStockTidakCocok.RowTemplate.Height = 24
        Me.dgvStockTidakCocok.Size = New System.Drawing.Size(1057, 539)
        Me.dgvStockTidakCocok.TabIndex = 4
        '
        'ColTdkCocok_No
        '
        Me.ColTdkCocok_No.HeaderText = "No"
        Me.ColTdkCocok_No.Name = "ColTdkCocok_No"
        Me.ColTdkCocok_No.Width = 51
        '
        'ColTdkCocok_CodeNo
        '
        Me.ColTdkCocok_CodeNo.HeaderText = "Code No"
        Me.ColTdkCocok_CodeNo.Name = "ColTdkCocok_CodeNo"
        Me.ColTdkCocok_CodeNo.Width = 88
        '
        'ColTdkCocok_Desc
        '
        Me.ColTdkCocok_Desc.HeaderText = "Description"
        Me.ColTdkCocok_Desc.Name = "ColTdkCocok_Desc"
        Me.ColTdkCocok_Desc.Width = 104
        '
        'TabPage7
        '
        Me.TabPage7.Controls.Add(Me.dgvCekSC)
        Me.TabPage7.Location = New System.Drawing.Point(4, 25)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage7.Size = New System.Drawing.Size(1077, 580)
        Me.TabPage7.TabIndex = 3
        Me.TabPage7.Text = "Checking Stock Card Adjust"
        Me.TabPage7.UseVisualStyleBackColor = True
        '
        'dgvCekSC
        '
        Me.dgvCekSC.AllowUserToAddRows = False
        Me.dgvCekSC.AllowUserToDeleteRows = False
        Me.dgvCekSC.AllowUserToResizeColumns = False
        Me.dgvCekSC.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCekSC.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        Me.dgvCekSC.ColumnHeadersHeight = 25
        Me.dgvCekSC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvCekSC.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCekSC_No, Me.ColCekSC_CodeNo, Me.ColCekSC_Desc, Me.ColCekSC_BatchNo, Me.ColCekSC_Location, Me.ColCekSC_Qty, Me.ColCekSC_Unit})
        Me.dgvCekSC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCekSC.Location = New System.Drawing.Point(3, 3)
        Me.dgvCekSC.Name = "dgvCekSC"
        Me.dgvCekSC.RowHeadersVisible = False
        Me.dgvCekSC.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvCekSC.RowTemplate.Height = 24
        Me.dgvCekSC.Size = New System.Drawing.Size(1071, 574)
        Me.dgvCekSC.TabIndex = 5
        '
        'ColCekSC_No
        '
        Me.ColCekSC_No.HeaderText = "No"
        Me.ColCekSC_No.Name = "ColCekSC_No"
        Me.ColCekSC_No.Width = 51
        '
        'ColCekSC_CodeNo
        '
        Me.ColCekSC_CodeNo.HeaderText = "Code No"
        Me.ColCekSC_CodeNo.Name = "ColCekSC_CodeNo"
        Me.ColCekSC_CodeNo.Width = 88
        '
        'ColCekSC_Desc
        '
        Me.ColCekSC_Desc.HeaderText = "Description"
        Me.ColCekSC_Desc.Name = "ColCekSC_Desc"
        Me.ColCekSC_Desc.Width = 104
        '
        'ColCekSC_BatchNo
        '
        Me.ColCekSC_BatchNo.HeaderText = "Batch No"
        Me.ColCekSC_BatchNo.Name = "ColCekSC_BatchNo"
        Me.ColCekSC_BatchNo.Width = 91
        '
        'ColCekSC_Location
        '
        Me.ColCekSC_Location.HeaderText = "Location"
        Me.ColCekSC_Location.Name = "ColCekSC_Location"
        Me.ColCekSC_Location.Width = 87
        '
        'ColCekSC_Qty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColCekSC_Qty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColCekSC_Qty.HeaderText = "Qty"
        Me.ColCekSC_Qty.Name = "ColCekSC_Qty"
        Me.ColCekSC_Qty.Width = 55
        '
        'ColCekSC_Unit
        '
        Me.ColCekSC_Unit.HeaderText = "Unit"
        Me.ColCekSC_Unit.Name = "ColCekSC_Unit"
        Me.ColCekSC_Unit.Width = 58
        '
        'TabPage8
        '
        Me.TabPage8.Controls.Add(Me.SplitContainer1)
        Me.TabPage8.Location = New System.Drawing.Point(4, 25)
        Me.TabPage8.Name = "TabPage8"
        Me.TabPage8.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage8.Size = New System.Drawing.Size(1077, 580)
        Me.TabPage8.TabIndex = 4
        Me.TabPage8.Text = "Reverse Adjustment"
        Me.TabPage8.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 3)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.LinkLabel1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtp_Step8_4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dgvAdjustment)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnStep8_4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label17)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnStep8_3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label22)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnStep8_2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label20)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnStep8_1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label18)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TabControl3)
        Me.SplitContainer1.Size = New System.Drawing.Size(1071, 574)
        Me.SplitContainer1.SplitterDistance = 323
        Me.SplitContainer1.TabIndex = 0
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(998, 6)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(66, 17)
        Me.LinkLabel1.TabIndex = 59
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Check All"
        '
        'dtp_Step8_4
        '
        Me.dtp_Step8_4.CustomFormat = "MMM yyyy"
        Me.dtp_Step8_4.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtp_Step8_4.Location = New System.Drawing.Point(775, 35)
        Me.dtp_Step8_4.Name = "dtp_Step8_4"
        Me.dtp_Step8_4.Size = New System.Drawing.Size(126, 22)
        Me.dtp_Step8_4.TabIndex = 58
        '
        'dgvAdjustment
        '
        Me.dgvAdjustment.AllowUserToAddRows = False
        Me.dgvAdjustment.AllowUserToDeleteRows = False
        Me.dgvAdjustment.AllowUserToResizeColumns = False
        Me.dgvAdjustment.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAdjustment.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAdjustment.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        Me.dgvAdjustment.ColumnHeadersHeight = 25
        Me.dgvAdjustment.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvAdjustment.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDGVAdjust_No, Me.ColDGVAdjust_DocNo, Me.ColDGVAdjust_Date, Me.ColDGVAdjust_Remarks, Me.ColDGVAdjust_Type, Me.ColDGVAdjust_CreatedBy, Me.ColDGVAdjust_CreatedDate, Me.ColDGVAdjust_ModifiedDate, Me.ColDGVAdjust_Include, Me.ColDGVAdjust_Primary})
        Me.dgvAdjustment.Location = New System.Drawing.Point(6, 64)
        Me.dgvAdjustment.Name = "dgvAdjustment"
        Me.dgvAdjustment.RowHeadersVisible = False
        Me.dgvAdjustment.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvAdjustment.RowTemplate.Height = 24
        Me.dgvAdjustment.Size = New System.Drawing.Size(1062, 256)
        Me.dgvAdjustment.TabIndex = 57
        '
        'ColDGVAdjust_No
        '
        Me.ColDGVAdjust_No.HeaderText = "No"
        Me.ColDGVAdjust_No.Name = "ColDGVAdjust_No"
        Me.ColDGVAdjust_No.Width = 51
        '
        'ColDGVAdjust_DocNo
        '
        Me.ColDGVAdjust_DocNo.HeaderText = "Document No"
        Me.ColDGVAdjust_DocNo.Name = "ColDGVAdjust_DocNo"
        Me.ColDGVAdjust_DocNo.Width = 119
        '
        'ColDGVAdjust_Date
        '
        Me.ColDGVAdjust_Date.HeaderText = "Date"
        Me.ColDGVAdjust_Date.Name = "ColDGVAdjust_Date"
        Me.ColDGVAdjust_Date.Width = 63
        '
        'ColDGVAdjust_Remarks
        '
        Me.ColDGVAdjust_Remarks.HeaderText = "Remarks"
        Me.ColDGVAdjust_Remarks.Name = "ColDGVAdjust_Remarks"
        Me.ColDGVAdjust_Remarks.Width = 89
        '
        'ColDGVAdjust_Type
        '
        Me.ColDGVAdjust_Type.HeaderText = "Type"
        Me.ColDGVAdjust_Type.Name = "ColDGVAdjust_Type"
        Me.ColDGVAdjust_Type.Width = 65
        '
        'ColDGVAdjust_CreatedBy
        '
        Me.ColDGVAdjust_CreatedBy.HeaderText = "Created By"
        Me.ColDGVAdjust_CreatedBy.Name = "ColDGVAdjust_CreatedBy"
        Me.ColDGVAdjust_CreatedBy.Width = 103
        '
        'ColDGVAdjust_CreatedDate
        '
        Me.ColDGVAdjust_CreatedDate.HeaderText = "Created Date"
        Me.ColDGVAdjust_CreatedDate.Name = "ColDGVAdjust_CreatedDate"
        Me.ColDGVAdjust_CreatedDate.Width = 117
        '
        'ColDGVAdjust_ModifiedDate
        '
        Me.ColDGVAdjust_ModifiedDate.HeaderText = "Modified Date"
        Me.ColDGVAdjust_ModifiedDate.Name = "ColDGVAdjust_ModifiedDate"
        Me.ColDGVAdjust_ModifiedDate.Width = 120
        '
        'ColDGVAdjust_Include
        '
        Me.ColDGVAdjust_Include.HeaderText = "Include"
        Me.ColDGVAdjust_Include.Name = "ColDGVAdjust_Include"
        Me.ColDGVAdjust_Include.Width = 59
        '
        'ColDGVAdjust_Primary
        '
        Me.ColDGVAdjust_Primary.HeaderText = "Primary"
        Me.ColDGVAdjust_Primary.Name = "ColDGVAdjust_Primary"
        Me.ColDGVAdjust_Primary.Visible = False
        Me.ColDGVAdjust_Primary.Width = 81
        '
        'btnStep8_4
        '
        Me.btnStep8_4.Location = New System.Drawing.Point(694, 35)
        Me.btnStep8_4.Name = "btnStep8_4"
        Me.btnStep8_4.Size = New System.Drawing.Size(75, 23)
        Me.btnStep8_4.TabIndex = 56
        Me.btnStep8_4.Text = "Step 8.4"
        Me.btnStep8_4.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(411, 38)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(277, 17)
        Me.Label17.TabIndex = 54
        Me.Label17.Text = "Step 8.4: Pindahkan data adjustment 2017"
        '
        'btnStep8_3
        '
        Me.btnStep8_3.Location = New System.Drawing.Point(694, 6)
        Me.btnStep8_3.Name = "btnStep8_3"
        Me.btnStep8_3.Size = New System.Drawing.Size(75, 23)
        Me.btnStep8_3.TabIndex = 53
        Me.btnStep8_3.Text = "Step 8.3"
        Me.btnStep8_3.UseVisualStyleBackColor = True
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(411, 9)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(234, 17)
        Me.Label22.TabIndex = 50
        Me.Label22.Text = "Step 8.3: Cek data adjustment 2017"
        '
        'btnStep8_2
        '
        Me.btnStep8_2.Location = New System.Drawing.Point(314, 32)
        Me.btnStep8_2.Name = "btnStep8_2"
        Me.btnStep8_2.Size = New System.Drawing.Size(75, 23)
        Me.btnStep8_2.TabIndex = 49
        Me.btnStep8_2.Text = "Step 8.2"
        Me.btnStep8_2.UseVisualStyleBackColor = True
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(3, 35)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(305, 17)
        Me.Label20.TabIndex = 46
        Me.Label20.Text = "Step 8.2: Simpan data adjustment closing 2017"
        '
        'btnStep8_1
        '
        Me.btnStep8_1.Location = New System.Drawing.Point(314, 3)
        Me.btnStep8_1.Name = "btnStep8_1"
        Me.btnStep8_1.Size = New System.Drawing.Size(75, 23)
        Me.btnStep8_1.TabIndex = 45
        Me.btnStep8_1.Text = "Step 8.1"
        Me.btnStep8_1.UseVisualStyleBackColor = True
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(3, 6)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(283, 17)
        Me.Label18.TabIndex = 42
        Me.Label18.Text = "Step 8.1: Cari data adjustment closing 2017"
        '
        'TabControl3
        '
        Me.TabControl3.Controls.Add(Me.TabPage9)
        Me.TabControl3.Controls.Add(Me.TabPage10)
        Me.TabControl3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl3.Location = New System.Drawing.Point(0, 0)
        Me.TabControl3.Name = "TabControl3"
        Me.TabControl3.SelectedIndex = 0
        Me.TabControl3.Size = New System.Drawing.Size(1071, 247)
        Me.TabControl3.TabIndex = 59
        '
        'TabPage9
        '
        Me.TabPage9.Controls.Add(Me.dgvAdjustmentItem1)
        Me.TabPage9.Location = New System.Drawing.Point(4, 25)
        Me.TabPage9.Name = "TabPage9"
        Me.TabPage9.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage9.Size = New System.Drawing.Size(1063, 218)
        Me.TabPage9.TabIndex = 0
        Me.TabPage9.Text = "Detail Item"
        Me.TabPage9.UseVisualStyleBackColor = True
        '
        'dgvAdjustmentItem1
        '
        Me.dgvAdjustmentItem1.AllowUserToAddRows = False
        Me.dgvAdjustmentItem1.AllowUserToDeleteRows = False
        Me.dgvAdjustmentItem1.AllowUserToResizeColumns = False
        Me.dgvAdjustmentItem1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAdjustmentItem1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        Me.dgvAdjustmentItem1.ColumnHeadersHeight = 25
        Me.dgvAdjustmentItem1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvAdjustmentItem1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAdjItem1_No, Me.ColAdjItem1_CodeNo, Me.ColAdjItem1_Desc, Me.ColAdjItem1_Location, Me.ColAdjItem1_Batch, Me.ColAdjItem1_Expiry, Me.ColAdjItem1_Qty, Me.ColAdjItem1_Unit, Me.ColAdjItem1_Remarks})
        Me.dgvAdjustmentItem1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAdjustmentItem1.Location = New System.Drawing.Point(3, 3)
        Me.dgvAdjustmentItem1.Name = "dgvAdjustmentItem1"
        Me.dgvAdjustmentItem1.ReadOnly = True
        Me.dgvAdjustmentItem1.RowHeadersVisible = False
        Me.dgvAdjustmentItem1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvAdjustmentItem1.RowTemplate.Height = 24
        Me.dgvAdjustmentItem1.Size = New System.Drawing.Size(1057, 212)
        Me.dgvAdjustmentItem1.TabIndex = 59
        '
        'ColAdjItem1_No
        '
        Me.ColAdjItem1_No.HeaderText = "No"
        Me.ColAdjItem1_No.Name = "ColAdjItem1_No"
        Me.ColAdjItem1_No.ReadOnly = True
        Me.ColAdjItem1_No.Width = 51
        '
        'ColAdjItem1_CodeNo
        '
        Me.ColAdjItem1_CodeNo.HeaderText = "Code No"
        Me.ColAdjItem1_CodeNo.Name = "ColAdjItem1_CodeNo"
        Me.ColAdjItem1_CodeNo.ReadOnly = True
        Me.ColAdjItem1_CodeNo.Width = 88
        '
        'ColAdjItem1_Desc
        '
        Me.ColAdjItem1_Desc.HeaderText = "Desc"
        Me.ColAdjItem1_Desc.Name = "ColAdjItem1_Desc"
        Me.ColAdjItem1_Desc.ReadOnly = True
        Me.ColAdjItem1_Desc.Width = 65
        '
        'ColAdjItem1_Location
        '
        Me.ColAdjItem1_Location.HeaderText = "Location"
        Me.ColAdjItem1_Location.Name = "ColAdjItem1_Location"
        Me.ColAdjItem1_Location.ReadOnly = True
        Me.ColAdjItem1_Location.Width = 87
        '
        'ColAdjItem1_Batch
        '
        Me.ColAdjItem1_Batch.HeaderText = "Batch No"
        Me.ColAdjItem1_Batch.Name = "ColAdjItem1_Batch"
        Me.ColAdjItem1_Batch.ReadOnly = True
        Me.ColAdjItem1_Batch.Width = 91
        '
        'ColAdjItem1_Expiry
        '
        Me.ColAdjItem1_Expiry.HeaderText = "Expiry"
        Me.ColAdjItem1_Expiry.Name = "ColAdjItem1_Expiry"
        Me.ColAdjItem1_Expiry.ReadOnly = True
        Me.ColAdjItem1_Expiry.Width = 71
        '
        'ColAdjItem1_Qty
        '
        Me.ColAdjItem1_Qty.HeaderText = "Qty"
        Me.ColAdjItem1_Qty.Name = "ColAdjItem1_Qty"
        Me.ColAdjItem1_Qty.ReadOnly = True
        Me.ColAdjItem1_Qty.Width = 55
        '
        'ColAdjItem1_Unit
        '
        Me.ColAdjItem1_Unit.HeaderText = "Unit"
        Me.ColAdjItem1_Unit.Name = "ColAdjItem1_Unit"
        Me.ColAdjItem1_Unit.ReadOnly = True
        Me.ColAdjItem1_Unit.Width = 58
        '
        'ColAdjItem1_Remarks
        '
        Me.ColAdjItem1_Remarks.HeaderText = "Remarks"
        Me.ColAdjItem1_Remarks.Name = "ColAdjItem1_Remarks"
        Me.ColAdjItem1_Remarks.ReadOnly = True
        Me.ColAdjItem1_Remarks.Width = 89
        '
        'TabPage10
        '
        Me.TabPage10.Controls.Add(Me.dgvAdjustmentItem2)
        Me.TabPage10.Location = New System.Drawing.Point(4, 25)
        Me.TabPage10.Name = "TabPage10"
        Me.TabPage10.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage10.Size = New System.Drawing.Size(1063, 218)
        Me.TabPage10.TabIndex = 1
        Me.TabPage10.Text = "Rekap Item"
        Me.TabPage10.UseVisualStyleBackColor = True
        '
        'dgvAdjustmentItem2
        '
        Me.dgvAdjustmentItem2.AllowUserToAddRows = False
        Me.dgvAdjustmentItem2.AllowUserToDeleteRows = False
        Me.dgvAdjustmentItem2.AllowUserToResizeColumns = False
        Me.dgvAdjustmentItem2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAdjustmentItem2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        Me.dgvAdjustmentItem2.ColumnHeadersHeight = 25
        Me.dgvAdjustmentItem2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvAdjustmentItem2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItemAdj2_No, Me.ColItemAdj2_CodeNo, Me.ColItemAdj2_Desc, Me.ColItemAdj2_Location, Me.ColItemAdj2_BatchNo, Me.ColItemAdj2_Expiry, Me.ColItemAdj2_Qty, Me.ColItemAdj2_Unit})
        Me.dgvAdjustmentItem2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAdjustmentItem2.Location = New System.Drawing.Point(3, 3)
        Me.dgvAdjustmentItem2.Name = "dgvAdjustmentItem2"
        Me.dgvAdjustmentItem2.ReadOnly = True
        Me.dgvAdjustmentItem2.RowHeadersVisible = False
        Me.dgvAdjustmentItem2.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvAdjustmentItem2.RowTemplate.Height = 24
        Me.dgvAdjustmentItem2.Size = New System.Drawing.Size(1057, 212)
        Me.dgvAdjustmentItem2.TabIndex = 60
        '
        'ColItemAdj2_No
        '
        Me.ColItemAdj2_No.HeaderText = "No"
        Me.ColItemAdj2_No.Name = "ColItemAdj2_No"
        Me.ColItemAdj2_No.ReadOnly = True
        Me.ColItemAdj2_No.Width = 51
        '
        'ColItemAdj2_CodeNo
        '
        Me.ColItemAdj2_CodeNo.HeaderText = "Code No"
        Me.ColItemAdj2_CodeNo.Name = "ColItemAdj2_CodeNo"
        Me.ColItemAdj2_CodeNo.ReadOnly = True
        Me.ColItemAdj2_CodeNo.Width = 88
        '
        'ColItemAdj2_Desc
        '
        Me.ColItemAdj2_Desc.HeaderText = "Desc"
        Me.ColItemAdj2_Desc.Name = "ColItemAdj2_Desc"
        Me.ColItemAdj2_Desc.ReadOnly = True
        Me.ColItemAdj2_Desc.Width = 65
        '
        'ColItemAdj2_Location
        '
        Me.ColItemAdj2_Location.HeaderText = "Location"
        Me.ColItemAdj2_Location.Name = "ColItemAdj2_Location"
        Me.ColItemAdj2_Location.ReadOnly = True
        Me.ColItemAdj2_Location.Width = 87
        '
        'ColItemAdj2_BatchNo
        '
        Me.ColItemAdj2_BatchNo.HeaderText = "Batch No"
        Me.ColItemAdj2_BatchNo.Name = "ColItemAdj2_BatchNo"
        Me.ColItemAdj2_BatchNo.ReadOnly = True
        Me.ColItemAdj2_BatchNo.Width = 91
        '
        'ColItemAdj2_Expiry
        '
        Me.ColItemAdj2_Expiry.HeaderText = "Expiry"
        Me.ColItemAdj2_Expiry.Name = "ColItemAdj2_Expiry"
        Me.ColItemAdj2_Expiry.ReadOnly = True
        Me.ColItemAdj2_Expiry.Width = 71
        '
        'ColItemAdj2_Qty
        '
        Me.ColItemAdj2_Qty.HeaderText = "Qty"
        Me.ColItemAdj2_Qty.Name = "ColItemAdj2_Qty"
        Me.ColItemAdj2_Qty.ReadOnly = True
        Me.ColItemAdj2_Qty.Width = 55
        '
        'ColItemAdj2_Unit
        '
        Me.ColItemAdj2_Unit.HeaderText = "Unit"
        Me.ColItemAdj2_Unit.Name = "ColItemAdj2_Unit"
        Me.ColItemAdj2_Unit.ReadOnly = True
        Me.ColItemAdj2_Unit.Width = 58
        '
        'TabPage11
        '
        Me.TabPage11.Controls.Add(Me.dgvFinal)
        Me.TabPage11.Location = New System.Drawing.Point(4, 25)
        Me.TabPage11.Name = "TabPage11"
        Me.TabPage11.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage11.Size = New System.Drawing.Size(1077, 580)
        Me.TabPage11.TabIndex = 5
        Me.TabPage11.Text = "Final Result"
        Me.TabPage11.UseVisualStyleBackColor = True
        '
        'dgvFinal
        '
        Me.dgvFinal.AllowUserToAddRows = False
        Me.dgvFinal.AllowUserToDeleteRows = False
        Me.dgvFinal.AllowUserToResizeColumns = False
        Me.dgvFinal.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvFinal.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        Me.dgvFinal.ColumnHeadersHeight = 25
        Me.dgvFinal.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvFinal.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColFinal_No, Me.ColFinal_GroupCode, Me.ColFinal_GroupDesc, Me.ColFinal_BrandCode, Me.ColFinal_BrandDesc, Me.ColFinal_SubBrandCode, Me.ColFinal_SubBrandDesc, Me.ColFinal_CategoryCode, Me.ColFinal_CategoryDesc, Me.ColFinal_CodeNo, Me.ColFinal_Desc, Me.ColFinal_Qty, Me.ColFinal_Unit, Me.ColFinal_Qtym3, Me.ColFinal_Warehouse, Me.ColFinal_WarehouseDesc, Me.ColFinal_UnitPrice, Me.ColFinal_Amouunt, Me.ColFinal_AccCode, Me.ColFinal_AccDesc})
        Me.dgvFinal.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvFinal.Location = New System.Drawing.Point(3, 3)
        Me.dgvFinal.Name = "dgvFinal"
        Me.dgvFinal.RowHeadersVisible = False
        Me.dgvFinal.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvFinal.RowTemplate.Height = 24
        Me.dgvFinal.Size = New System.Drawing.Size(1071, 574)
        Me.dgvFinal.TabIndex = 5
        '
        'ColFinal_No
        '
        Me.ColFinal_No.HeaderText = "No"
        Me.ColFinal_No.Name = "ColFinal_No"
        Me.ColFinal_No.Width = 51
        '
        'ColFinal_GroupCode
        '
        Me.ColFinal_GroupCode.HeaderText = "Group Code"
        Me.ColFinal_GroupCode.Name = "ColFinal_GroupCode"
        Me.ColFinal_GroupCode.Width = 110
        '
        'ColFinal_GroupDesc
        '
        Me.ColFinal_GroupDesc.HeaderText = "Group Desc"
        Me.ColFinal_GroupDesc.Name = "ColFinal_GroupDesc"
        Me.ColFinal_GroupDesc.Width = 109
        '
        'ColFinal_BrandCode
        '
        Me.ColFinal_BrandCode.HeaderText = "Brand Code"
        Me.ColFinal_BrandCode.Name = "ColFinal_BrandCode"
        Me.ColFinal_BrandCode.Width = 108
        '
        'ColFinal_BrandDesc
        '
        Me.ColFinal_BrandDesc.HeaderText = "Brand Desc"
        Me.ColFinal_BrandDesc.Name = "ColFinal_BrandDesc"
        Me.ColFinal_BrandDesc.Width = 107
        '
        'ColFinal_SubBrandCode
        '
        Me.ColFinal_SubBrandCode.HeaderText = "Sub Brand Code"
        Me.ColFinal_SubBrandCode.Name = "ColFinal_SubBrandCode"
        Me.ColFinal_SubBrandCode.Width = 137
        '
        'ColFinal_SubBrandDesc
        '
        Me.ColFinal_SubBrandDesc.HeaderText = "Sub Brand Desc"
        Me.ColFinal_SubBrandDesc.Name = "ColFinal_SubBrandDesc"
        Me.ColFinal_SubBrandDesc.Width = 136
        '
        'ColFinal_CategoryCode
        '
        Me.ColFinal_CategoryCode.HeaderText = "Category Code"
        Me.ColFinal_CategoryCode.Name = "ColFinal_CategoryCode"
        Me.ColFinal_CategoryCode.Width = 127
        '
        'ColFinal_CategoryDesc
        '
        Me.ColFinal_CategoryDesc.HeaderText = "Category Description"
        Me.ColFinal_CategoryDesc.Name = "ColFinal_CategoryDesc"
        Me.ColFinal_CategoryDesc.Width = 165
        '
        'ColFinal_CodeNo
        '
        Me.ColFinal_CodeNo.HeaderText = "Code No"
        Me.ColFinal_CodeNo.Name = "ColFinal_CodeNo"
        Me.ColFinal_CodeNo.Width = 88
        '
        'ColFinal_Desc
        '
        Me.ColFinal_Desc.HeaderText = "Description"
        Me.ColFinal_Desc.Name = "ColFinal_Desc"
        Me.ColFinal_Desc.Width = 104
        '
        'ColFinal_Qty
        '
        Me.ColFinal_Qty.HeaderText = "Qty"
        Me.ColFinal_Qty.Name = "ColFinal_Qty"
        Me.ColFinal_Qty.Width = 55
        '
        'ColFinal_Unit
        '
        Me.ColFinal_Unit.HeaderText = "Unit"
        Me.ColFinal_Unit.Name = "ColFinal_Unit"
        Me.ColFinal_Unit.Width = 58
        '
        'ColFinal_Qtym3
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColFinal_Qtym3.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColFinal_Qtym3.HeaderText = "Qty M3"
        Me.ColFinal_Qtym3.Name = "ColFinal_Qtym3"
        Me.ColFinal_Qtym3.Width = 78
        '
        'ColFinal_Warehouse
        '
        Me.ColFinal_Warehouse.HeaderText = "Warehouse"
        Me.ColFinal_Warehouse.Name = "ColFinal_Warehouse"
        Me.ColFinal_Warehouse.Width = 106
        '
        'ColFinal_WarehouseDesc
        '
        Me.ColFinal_WarehouseDesc.HeaderText = "Warehouse Desc"
        Me.ColFinal_WarehouseDesc.Name = "ColFinal_WarehouseDesc"
        Me.ColFinal_WarehouseDesc.Width = 142
        '
        'ColFinal_UnitPrice
        '
        Me.ColFinal_UnitPrice.HeaderText = "Unit Price"
        Me.ColFinal_UnitPrice.Name = "ColFinal_UnitPrice"
        Me.ColFinal_UnitPrice.Width = 94
        '
        'ColFinal_Amouunt
        '
        Me.ColFinal_Amouunt.HeaderText = "Amount"
        Me.ColFinal_Amouunt.Name = "ColFinal_Amouunt"
        Me.ColFinal_Amouunt.Width = 81
        '
        'ColFinal_AccCode
        '
        Me.ColFinal_AccCode.HeaderText = "Acc. Code"
        Me.ColFinal_AccCode.Name = "ColFinal_AccCode"
        Me.ColFinal_AccCode.Width = 97
        '
        'ColFinal_AccDesc
        '
        Me.ColFinal_AccDesc.HeaderText = "Acc. Desc"
        Me.ColFinal_AccDesc.Name = "ColFinal_AccDesc"
        Me.ColFinal_AccDesc.Width = 96
        '
        'TabPage12
        '
        Me.TabPage12.Controls.Add(Me.btnStep20)
        Me.TabPage12.Controls.Add(Me.Label19)
        Me.TabPage12.Controls.Add(Me.btnStep1_RemoveMaintenance)
        Me.TabPage12.Controls.Add(Me.btnStep19_Set)
        Me.TabPage12.Controls.Add(Me.Label25)
        Me.TabPage12.Controls.Add(Me.prgStep18)
        Me.TabPage12.Controls.Add(Me.LblPrgStep18)
        Me.TabPage12.Controls.Add(Me.btnStep18)
        Me.TabPage12.Controls.Add(Me.Label24)
        Me.TabPage12.Controls.Add(Me.prgStep17)
        Me.TabPage12.Controls.Add(Me.LblPrgStep17)
        Me.TabPage12.Controls.Add(Me.btnStep17)
        Me.TabPage12.Controls.Add(Me.Label23)
        Me.TabPage12.Controls.Add(Me.prgStep16)
        Me.TabPage12.Controls.Add(Me.LblPrgStep16)
        Me.TabPage12.Controls.Add(Me.btnStep16)
        Me.TabPage12.Controls.Add(Me.Label21)
        Me.TabPage12.Controls.Add(Me.prgStep15)
        Me.TabPage12.Controls.Add(Me.LblPrgStep15)
        Me.TabPage12.Controls.Add(Me.btnStep15)
        Me.TabPage12.Controls.Add(Me.Label15)
        Me.TabPage12.Controls.Add(Me.dtpStep14)
        Me.TabPage12.Controls.Add(Me.btnStep14)
        Me.TabPage12.Controls.Add(Me.Label14)
        Me.TabPage12.Location = New System.Drawing.Point(4, 25)
        Me.TabPage12.Name = "TabPage12"
        Me.TabPage12.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage12.Size = New System.Drawing.Size(1077, 580)
        Me.TabPage12.TabIndex = 6
        Me.TabPage12.Text = "Update Data - cont'"
        Me.TabPage12.UseVisualStyleBackColor = True
        '
        'btnStep20
        '
        Me.btnStep20.Location = New System.Drawing.Point(379, 250)
        Me.btnStep20.Name = "btnStep20"
        Me.btnStep20.Size = New System.Drawing.Size(131, 25)
        Me.btnStep20.TabIndex = 79
        Me.btnStep20.Text = "Cek Bundling"
        Me.btnStep20.UseVisualStyleBackColor = True
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(6, 253)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(148, 17)
        Me.Label19.TabIndex = 78
        Me.Label19.Text = "Step 20: Cek Bundling"
        '
        'btnStep1_RemoveMaintenance
        '
        Me.btnStep1_RemoveMaintenance.Location = New System.Drawing.Point(516, 219)
        Me.btnStep1_RemoveMaintenance.Name = "btnStep1_RemoveMaintenance"
        Me.btnStep1_RemoveMaintenance.Size = New System.Drawing.Size(178, 25)
        Me.btnStep1_RemoveMaintenance.TabIndex = 77
        Me.btnStep1_RemoveMaintenance.Text = "Remove Maintenance"
        Me.btnStep1_RemoveMaintenance.UseVisualStyleBackColor = True
        '
        'btnStep19_Set
        '
        Me.btnStep19_Set.Location = New System.Drawing.Point(379, 219)
        Me.btnStep19_Set.Name = "btnStep19_Set"
        Me.btnStep19_Set.Size = New System.Drawing.Size(131, 25)
        Me.btnStep19_Set.TabIndex = 76
        Me.btnStep19_Set.Text = "Set Maintenance"
        Me.btnStep19_Set.UseVisualStyleBackColor = True
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(6, 222)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(171, 17)
        Me.Label25.TabIndex = 75
        Me.Label25.Text = "Step 19: Set Maintenance"
        '
        'prgStep18
        '
        Me.prgStep18.Location = New System.Drawing.Point(379, 173)
        Me.prgStep18.Name = "prgStep18"
        Me.prgStep18.Size = New System.Drawing.Size(284, 23)
        Me.prgStep18.TabIndex = 73
        '
        'LblPrgStep18
        '
        Me.LblPrgStep18.AutoSize = True
        Me.LblPrgStep18.Location = New System.Drawing.Point(376, 199)
        Me.LblPrgStep18.Name = "LblPrgStep18"
        Me.LblPrgStep18.Size = New System.Drawing.Size(59, 17)
        Me.LblPrgStep18.TabIndex = 74
        Me.LblPrgStep18.Text = "pending"
        '
        'btnStep18
        '
        Me.btnStep18.Location = New System.Drawing.Point(669, 173)
        Me.btnStep18.Name = "btnStep18"
        Me.btnStep18.Size = New System.Drawing.Size(75, 23)
        Me.btnStep18.TabIndex = 72
        Me.btnStep18.Text = "Step 18"
        Me.btnStep18.UseVisualStyleBackColor = True
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(6, 176)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(148, 17)
        Me.Label24.TabIndex = 71
        Me.Label24.Text = "Step 18: Edit Bundling"
        '
        'prgStep17
        '
        Me.prgStep17.Location = New System.Drawing.Point(379, 127)
        Me.prgStep17.Name = "prgStep17"
        Me.prgStep17.Size = New System.Drawing.Size(284, 23)
        Me.prgStep17.TabIndex = 69
        '
        'LblPrgStep17
        '
        Me.LblPrgStep17.AutoSize = True
        Me.LblPrgStep17.Location = New System.Drawing.Point(376, 153)
        Me.LblPrgStep17.Name = "LblPrgStep17"
        Me.LblPrgStep17.Size = New System.Drawing.Size(59, 17)
        Me.LblPrgStep17.TabIndex = 70
        Me.LblPrgStep17.Text = "pending"
        '
        'btnStep17
        '
        Me.btnStep17.Location = New System.Drawing.Point(669, 127)
        Me.btnStep17.Name = "btnStep17"
        Me.btnStep17.Size = New System.Drawing.Size(75, 23)
        Me.btnStep17.TabIndex = 68
        Me.btnStep17.Text = "Step 17"
        Me.btnStep17.UseVisualStyleBackColor = True
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(6, 130)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(225, 17)
        Me.Label23.TabIndex = 67
        Me.Label23.Text = "Step 17: Hapus closing sementara"
        '
        'prgStep16
        '
        Me.prgStep16.Location = New System.Drawing.Point(379, 81)
        Me.prgStep16.Name = "prgStep16"
        Me.prgStep16.Size = New System.Drawing.Size(284, 23)
        Me.prgStep16.TabIndex = 65
        '
        'LblPrgStep16
        '
        Me.LblPrgStep16.AutoSize = True
        Me.LblPrgStep16.Location = New System.Drawing.Point(376, 107)
        Me.LblPrgStep16.Name = "LblPrgStep16"
        Me.LblPrgStep16.Size = New System.Drawing.Size(59, 17)
        Me.LblPrgStep16.TabIndex = 66
        Me.LblPrgStep16.Text = "pending"
        '
        'btnStep16
        '
        Me.btnStep16.Location = New System.Drawing.Point(669, 81)
        Me.btnStep16.Name = "btnStep16"
        Me.btnStep16.Size = New System.Drawing.Size(75, 23)
        Me.btnStep16.TabIndex = 64
        Me.btnStep16.Text = "Step 16"
        Me.btnStep16.UseVisualStyleBackColor = True
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(6, 84)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(213, 17)
        Me.Label21.TabIndex = 63
        Me.Label21.Text = "Step 16: Buat closing sementara"
        '
        'prgStep15
        '
        Me.prgStep15.Location = New System.Drawing.Point(379, 35)
        Me.prgStep15.Name = "prgStep15"
        Me.prgStep15.Size = New System.Drawing.Size(284, 23)
        Me.prgStep15.TabIndex = 61
        '
        'LblPrgStep15
        '
        Me.LblPrgStep15.AutoSize = True
        Me.LblPrgStep15.Location = New System.Drawing.Point(376, 61)
        Me.LblPrgStep15.Name = "LblPrgStep15"
        Me.LblPrgStep15.Size = New System.Drawing.Size(59, 17)
        Me.LblPrgStep15.TabIndex = 62
        Me.LblPrgStep15.Text = "pending"
        '
        'btnStep15
        '
        Me.btnStep15.Location = New System.Drawing.Point(669, 35)
        Me.btnStep15.Name = "btnStep15"
        Me.btnStep15.Size = New System.Drawing.Size(75, 23)
        Me.btnStep15.TabIndex = 60
        Me.btnStep15.Text = "Step 15"
        Me.btnStep15.UseVisualStyleBackColor = True
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(6, 38)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(367, 17)
        Me.Label15.TabIndex = 59
        Me.Label15.Text = "Step 15: Hapus semua item aop yang tidak memiliki GRN"
        '
        'dtpStep14
        '
        Me.dtpStep14.CustomFormat = "MMM yyyy"
        Me.dtpStep14.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStep14.Location = New System.Drawing.Point(379, 4)
        Me.dtpStep14.Name = "dtpStep14"
        Me.dtpStep14.ShowUpDown = True
        Me.dtpStep14.Size = New System.Drawing.Size(126, 22)
        Me.dtpStep14.TabIndex = 58
        '
        'btnStep14
        '
        Me.btnStep14.Location = New System.Drawing.Point(669, 6)
        Me.btnStep14.Name = "btnStep14"
        Me.btnStep14.Size = New System.Drawing.Size(75, 23)
        Me.btnStep14.TabIndex = 57
        Me.btnStep14.Text = "Step 14"
        Me.btnStep14.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(6, 9)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(273, 17)
        Me.Label14.TabIndex = 56
        Me.Label14.Text = "Step 14: cek end qty closing vs stock card"
        '
        'bgwStep1_2
        '
        Me.bgwStep1_2.WorkerReportsProgress = True
        '
        'bgwStep2
        '
        Me.bgwStep2.WorkerReportsProgress = True
        '
        'bgwStep3
        '
        Me.bgwStep3.WorkerReportsProgress = True
        '
        'bgwStep4
        '
        Me.bgwStep4.WorkerReportsProgress = True
        '
        'bgwStep5
        '
        Me.bgwStep5.WorkerReportsProgress = True
        '
        'bgwStep6
        '
        Me.bgwStep6.WorkerReportsProgress = True
        '
        'bgwStep15
        '
        Me.bgwStep15.WorkerReportsProgress = True
        '
        'bgwStep16
        '
        Me.bgwStep16.WorkerReportsProgress = True
        '
        'bgwStep17
        '
        Me.bgwStep17.WorkerReportsProgress = True
        '
        'TabPage13
        '
        Me.TabPage13.Controls.Add(Me.btnCek_Bundling)
        Me.TabPage13.Controls.Add(Me.ProgressBar1)
        Me.TabPage13.Controls.Add(Me.Label26)
        Me.TabPage13.Location = New System.Drawing.Point(4, 25)
        Me.TabPage13.Name = "TabPage13"
        Me.TabPage13.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage13.Size = New System.Drawing.Size(1077, 580)
        Me.TabPage13.TabIndex = 7
        Me.TabPage13.Text = "Data Check"
        Me.TabPage13.UseVisualStyleBackColor = True
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(6, 14)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(170, 17)
        Me.Label26.TabIndex = 0
        Me.Label26.Text = "Cek Bundling Jml Item = 0"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(257, 6)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(733, 25)
        Me.ProgressBar1.TabIndex = 1
        '
        'btnCek_Bundling
        '
        Me.btnCek_Bundling.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCek_Bundling.Location = New System.Drawing.Point(996, 6)
        Me.btnCek_Bundling.Name = "btnCek_Bundling"
        Me.btnCek_Bundling.Size = New System.Drawing.Size(75, 25)
        Me.btnCek_Bundling.TabIndex = 2
        Me.btnCek_Bundling.Text = "Start"
        Me.btnCek_Bundling.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1109, 633)
        Me.Controls.Add(Me.TabControl1)
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "Form2"
        Me.Text = "Form2"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.dgvOverhead, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        CType(Me.dgvStock, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage5.ResumeLayout(False)
        CType(Me.dgvStockMinus, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage6.ResumeLayout(False)
        CType(Me.dgvStockTidakCocok, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage7.ResumeLayout(False)
        CType(Me.dgvCekSC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage8.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvAdjustment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl3.ResumeLayout(False)
        Me.TabPage9.ResumeLayout(False)
        CType(Me.dgvAdjustmentItem1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage10.ResumeLayout(False)
        CType(Me.dgvAdjustmentItem2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage11.ResumeLayout(False)
        CType(Me.dgvFinal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage12.ResumeLayout(False)
        Me.TabPage12.PerformLayout()
        Me.TabPage13.ResumeLayout(False)
        Me.TabPage13.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnStep1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents prgStep1 As System.Windows.Forms.ProgressBar
    Friend WithEvents LblPrg_Step1 As System.Windows.Forms.Label
    Friend WithEvents LblPrgStep2 As System.Windows.Forms.Label
    Friend WithEvents prgStep2 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnStep2 As System.Windows.Forms.Button
    Friend WithEvents LblPrgStep3 As System.Windows.Forms.Label
    Friend WithEvents PrgStep3 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnStep3 As System.Windows.Forms.Button
    Friend WithEvents LblPrgStep4 As System.Windows.Forms.Label
    Friend WithEvents prgStep4 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnStep4 As System.Windows.Forms.Button
    Friend WithEvents bgwStep1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents btnShowOverhead As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpPeriodOverhead As System.Windows.Forms.DateTimePicker
    Friend WithEvents dgvOverhead As System.Windows.Forms.DataGridView
    Friend WithEvents ColOv_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOv_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOv_RefUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOv_AccCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOv_AccDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOv_Balance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOv_AllocatedPerc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnStep1_2 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents prb_Step1_2 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblPrgStep1_2 As System.Windows.Forms.Label
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents bgwStep1_2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents dgvStock As System.Windows.Forms.DataGridView
    Friend WithEvents ColStock_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStock_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStock_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStock_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStock_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStock_Batch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStock_Expiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvStockMinus As System.Windows.Forms.DataGridView
    Friend WithEvents ColDetailMinus_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetailMinus_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetailMinus_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetailMinus_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetailMinus_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetailMinus_Batch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetailMinus_Expiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents dgvStockTidakCocok As System.Windows.Forms.DataGridView
    Friend WithEvents ColTdkCocok_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTdkCocok_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTdkCocok_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dtpResetStock As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblPrgStep6 As System.Windows.Forms.Label
    Friend WithEvents prgStep6 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents LblPrgStep5 As System.Windows.Forms.Label
    Friend WithEvents prgStep5 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents bgwStep2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents dtpStep1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents bgwStep3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnStep7 As System.Windows.Forms.Button
    Friend WithEvents LblPrgStep7 As System.Windows.Forms.Label
    Friend WithEvents prgStep7 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnStep11 As System.Windows.Forms.Button
    Friend WithEvents LblPrgStep8 As System.Windows.Forms.Label
    Friend WithEvents prgStep8 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents bgwStep4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwStep5 As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwStep6 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
    Friend WithEvents dgvCekSC As System.Windows.Forms.DataGridView
    Friend WithEvents btnStep9 As System.Windows.Forms.Button
    Friend WithEvents lblPrgStep9 As System.Windows.Forms.Label
    Friend WithEvents prgStep9 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents ColCekSC_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCekSC_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCekSC_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCekSC_BatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCekSC_Location As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCekSC_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCekSC_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnDelAdj As System.Windows.Forms.Button
    Friend WithEvents TabPage8 As System.Windows.Forms.TabPage
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnStep8_4 As System.Windows.Forms.Button
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents btnStep8_3 As System.Windows.Forms.Button
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents btnStep8_2 As System.Windows.Forms.Button
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents btnStep8_1 As System.Windows.Forms.Button
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents dgvAdjustment As System.Windows.Forms.DataGridView
    Friend WithEvents TabControl3 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage9 As System.Windows.Forms.TabPage
    Friend WithEvents dgvAdjustmentItem1 As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage10 As System.Windows.Forms.TabPage
    Friend WithEvents dtp_Step8_4 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dgvAdjustmentItem2 As System.Windows.Forms.DataGridView
    Friend WithEvents ColDGVAdjust_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDGVAdjust_DocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDGVAdjust_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDGVAdjust_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDGVAdjust_Type As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDGVAdjust_CreatedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDGVAdjust_CreatedDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDGVAdjust_ModifiedDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDGVAdjust_Include As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDGVAdjust_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjItem1_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjItem1_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjItem1_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjItem1_Location As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjItem1_Batch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjItem1_Expiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjItem1_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjItem1_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjItem1_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents ColItemAdj2_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAdj2_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAdj2_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAdj2_Location As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAdj2_BatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAdj2_Expiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAdj2_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAdj2_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage11 As System.Windows.Forms.TabPage
    Friend WithEvents dgvFinal As System.Windows.Forms.DataGridView
    Friend WithEvents ColFinal_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_GroupCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_GroupDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_BrandCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_BrandDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_SubBrandCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_SubBrandDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_CategoryCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_CategoryDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_Qtym3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_Warehouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_WarehouseDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_UnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_Amouunt As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_AccCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFinal_AccDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnStep12 As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnStep13 As System.Windows.Forms.Button
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents dtpStep13 As System.Windows.Forms.DateTimePicker
    Friend WithEvents TabPage12 As System.Windows.Forms.TabPage
    Friend WithEvents dtpStep14 As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnStep14 As System.Windows.Forms.Button
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents prgStep15 As System.Windows.Forms.ProgressBar
    Friend WithEvents LblPrgStep15 As System.Windows.Forms.Label
    Friend WithEvents btnStep15 As System.Windows.Forms.Button
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents bgwStep15 As System.ComponentModel.BackgroundWorker
    Friend WithEvents prgStep16 As System.Windows.Forms.ProgressBar
    Friend WithEvents LblPrgStep16 As System.Windows.Forms.Label
    Friend WithEvents btnStep16 As System.Windows.Forms.Button
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents bgwStep16 As System.ComponentModel.BackgroundWorker
    Friend WithEvents prgStep17 As System.Windows.Forms.ProgressBar
    Friend WithEvents LblPrgStep17 As System.Windows.Forms.Label
    Friend WithEvents btnStep17 As System.Windows.Forms.Button
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents bgwStep17 As System.ComponentModel.BackgroundWorker
    Friend WithEvents prgStep18 As System.Windows.Forms.ProgressBar
    Friend WithEvents LblPrgStep18 As System.Windows.Forms.Label
    Friend WithEvents btnStep18 As System.Windows.Forms.Button
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents lblPrgStep1_2a As System.Windows.Forms.Label
    Friend WithEvents prb_Step1_2a As System.Windows.Forms.ProgressBar
    Friend WithEvents btnStep1_RemoveMaintenance As System.Windows.Forms.Button
    Friend WithEvents btnStep19_Set As System.Windows.Forms.Button
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents btnStep20 As System.Windows.Forms.Button
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents TabPage13 As System.Windows.Forms.TabPage
    Friend WithEvents btnCek_Bundling As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label26 As System.Windows.Forms.Label
End Class
