Option Explicit On
Option Strict Off

Public Class ClsTransaksiMaterialPlanningGroup
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mDocumentNo As String = ""
    Dim mDesc As String = ""
    Dim mTanggalTransaksi As Date = Now
    Dim mUrutan As Double = -1
    Dim mRemarks As String = ""
    Dim mStartPeriod As Date = #1/1/1950#
    Dim mEndPeriod As Date = #1/1/1950#
    Dim mIncludeSales As Integer = 0
    Dim mIncludeAdjustment As Integer = 0
    Dim mIncludeBundling As Integer = 0
    Dim mTipe As EnumTipe = EnumTipe.Gross_Margin_Cumulative_Percentage
    Dim mCodeNoFilter As String = ""
    Dim mProductDescFilter As String = ""
    Dim mStopBeforeRefValue As Integer = 0

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Enum ProductLevel As Integer
        ProductGroup
        Brand
        SubBrand
        Category
        Product
    End Enum

    Public Enum EnumLeadTimeStrategy As Integer
        Based_on_usage
        User_Defined
        Based_on_prefered_supplier '-> untuk tipe ini sudah obselete, karena sudah bisa memasukkan item berdasarkan prefered supplier
    End Enum

    Public Property StopBeforeRefValue() As Boolean
        Get
            If Me.mStopBeforeRefValue = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                Me.mStopBeforeRefValue = 1
            Else
                Me.mStopBeforeRefValue = 0
            End If
        End Set
    End Property

    Public Property ProductDescFilter() As String
        Get
            Return Me.mProductDescFilter
        End Get
        Set(ByVal value As String)
            mProductDescFilter = value
        End Set
    End Property

    Public Property CodeNoFilter() As String
        Get
            Return Me.mCodeNoFilter
        End Get
        Set(ByVal value As String)
            mCodeNoFilter = value
        End Set
    End Property

    Public Property Tipe() As EnumTipe
        Get
            Return Me.mTipe
        End Get
        Set(ByVal value As EnumTipe)
            Me.mTipe = value
        End Set
    End Property

    Public Property IncludeBundling() As Boolean
        Get
            If Me.mIncludeBundling = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                Me.mIncludeBundling = 1
            Else
                Me.mIncludeBundling = 0
            End If
        End Set
    End Property

    Public Property IncludeAdjustment() As Boolean
        Get
            If Me.mIncludeAdjustment = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                Me.mIncludeAdjustment = 1
            Else
                Me.mIncludeAdjustment = 0
            End If
        End Set
    End Property

    Public Property IncludeSales() As Boolean
        Get
            If Me.mIncludeSales = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                Me.mIncludeSales = 1
            Else
                Me.mIncludeSales = 0
            End If
        End Set
    End Property

    Public Property StartPeriod() As Date
        Get
            Return Me.mStartPeriod
        End Get
        Set(ByVal value As Date)
            mStartPeriod = value
        End Set
    End Property

    Public Property EndPeriod() As Date
        Get
            Return Me.mEndPeriod
        End Get
        Set(ByVal value As Date)
            mEndPeriod = value
        End Set
    End Property

    Public Property Urutan() As Double
        Get
            Return Me.mUrutan
        End Get
        Set(ByVal value As Double)
            mUrutan = value
        End Set
    End Property

    Public Property Desc() As String
        Get
            Return Me.mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
        End Set
    End Property

    Public Enum EnumTipe As Integer
        Gross_Margin_Cumulative_Percentage
        Nett_Sales_Cumulative_Percentage
        Usage_Qty_Cumulative_Percentage
        Gross_Margin_Percentage
        Nett_Sales_Percentage
        Usage_Qty_Percentage
        Gross_Margin_Minimum_Value
        Nett_Sales_Minimum_Value
        Usage_Qty_Minimum_Value
    End Enum

    Public Property Remarks() As String
        Get
            Return Me.mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Public Property TanggalTransaksi() As Date
        Get
            Return Me.mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
        End Set
    End Property

    Public Property DocumentNo() As String
        Get
            Return Me.mDocumentNo
        End Get
        Set(ByVal value As String)
            Me.mDocumentNo = value
        End Set
    End Property

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property CreatedBy() As String
        Get
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter
        Dim Jml As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Jml = DTA.ScalarQuery_GetJml(vNoTransaksi)
            If Jml.HasValue = False Then
                Return False
            Else
                If Jml.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vDgvItemGroup As DataGridView, ByVal vDGVItemSupplier As DataGridView, ByVal vDGVItemProduct As DataGridView, ByVal vDGVItemProductGroup As DataGridView, _
    ByVal vFrm As FrmTrans_MaterialPlanningDashboard_Data1, _
    ByVal TBLItem_Invoice As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoiceDataTable, _
    ByVal TBLItem_SR As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_srDataTable, _
    ByVal TBLItem_Bundling As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundlingDataTable, _
    ByVal TBLItem_Adjustment As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustmentDataTable) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter
        Dim TBL As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_groupDataTable
        Dim row As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_groupRow
        Dim ObjItem As New ClsTransaksiMaterialPlanningGroupItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mPrimary = .primary
                    Me.mProductDescFilter = .product_desc_filter
                    Me.mRemarks = .remarks
                    Me.mStartPeriod = .start_period
                    Me.mStopBeforeRefValue = .stop_right_before_ref_value
                    Me.mTanggalTransaksi = .tanggal
                    Me.mTipe = .tipe
                    Me.mUrutan = .urutan
                    Me.mCodeNoFilter = .code_no_filter
                    Me.mDesc = .description
                    Me.mDocumentNo = .document_no
                    Me.mEndPeriod = .end_period
                    Me.mIncludeAdjustment = .include_adjustment
                    Me.mIncludeBundling = .include_budling
                    Me.mIncludeSales = .include_sales

                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With
                'salin data item 
                ObjItem.IsiDGV(Me.mPrimary, vDGVItemSupplier, vDgvItemGroup, vDGVItemProduct, vDGVItemProductGroup, vFrm, Me.mTipe, _
                TBLItem_Invoice, TBLItem_SR, TBLItem_Bundling, TBLItem_Adjustment)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDgvItemGroup, vDGVItemSupplier, vDGVItemProductGroup, vDGVItemProduct)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItemGroup As DataGridView, ByVal vDgvItemSupplier As DataGridView, _
    ByVal vDGVItemProductGroup As DataGridView, ByVal vDGVItemProduct As DataGridView)
        Try
            Me.mCreatedBy = ""
            Me.mCreatedDate = Now
            Me.mModifiedBy = ""
            Me.mModifiedDate = Now

            Me.mPrimary = -1
            Me.mRemarks = ""
            Me.mTanggalTransaksi = Now
            Me.mTipe = EnumTipe.Gross_Margin_Cumulative_Percentage
            Me.mProductDescFilter = ""
            Me.mStartPeriod = #1/1/1950#
            Me.mEndPeriod = #1/1/1950#
            Me.mStopBeforeRefValue = 0
            Me.mUrutan = 0
            Me.mCodeNoFilter = ""
            Me.mDesc = ""
            Me.mDocumentNo = ""
            Me.mIncludeAdjustment = 0
            Me.mIncludeBundling = 0
            Me.mIncludeSales = 0

            Me.mTimestamp = Now
            vDgvItemGroup.Rows.Clear()
            vDGVItemProduct.Rows.Clear()
            vDGVItemProductGroup.Rows.Clear()
            vDgvItemSupplier.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function ApakahBolehEdit(ByVal vStrictEdit As Boolean, ByVal vUserCode As String, ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter
        Dim CreatedBy As String = ""
        Dim ObjUser As New Core_Module.ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cek strict edit
            If vStrictEdit = True Then
                CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                If CreatedBy = vUserCode Then
                Else
                    If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_MaterialPlanningDashboard_Data1_OverrideEdit") = False Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                            MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjUser = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vStrictDelete As Boolean, ByVal vUserCode As String, ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter
        Dim CreatedBy As String = ""
        Dim ObjUser As New Core_Module.ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cek strict edit
            If vStrictDelete = True Then
                CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                If CreatedBy = vUserCode Then
                Else
                    If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_MaterialPlanningDashboard_Data1_OverrideDelete") = False Then
                        MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                            MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjUser = Nothing
        End Try
    End Function

    'pada fungsi ini diasumsikan semua dgv yang masuk sudah allow user to add row = false
    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vDGVGroup As DataGridView, ByVal vDGVProduct As DataGridView, _
    ByVal vFrm As FrmTrans_MaterialPlanningDashboard_Data1) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter
        Dim Jml As Nullable(Of Long)
        Dim vPrimaryAcuan As Nullable(Of Double)
        Dim vDocNoAcuan As String = ""
        Dim rowDGV, rowDgvBefore As DataGridViewRow
        Dim TotalPerc As Decimal = 0
        Dim ObjClosing As New ClsTransaksiMonthlyClosingItem
        Dim TipeTrans As EnumTipe = EnumTipe.Gross_Margin_Cumulative_Percentage

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                Jml = DTA.ScalarQuery_GetJml(Me.mDocumentNo)
                If Jml.HasValue = True Then
                    If Jml.Value > 0 Then
                        MsgBox("Document number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Else
                vDocNoAcuan = DTA.ScalarQuery_GetDocumentNo(Me.mPrimary)
                If vDocNoAcuan <> Me.mDocumentNo Then
                    vPrimaryAcuan = DTA.ScalarQuery_GetPrimary(Me.mDocumentNo)
                    If vPrimaryAcuan.HasValue = True Then
                        If vPrimaryAcuan.Value <> Me.mPrimary Then
                            MsgBox("Document number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If

            '1.b Field berikut ini kosong: no transaksi, desc, 
            Me.mDocumentNo = Me.mDocumentNo.Trim
            If Me.mDocumentNo = "" Then
                MsgBox("Please specify document number.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Me.mDesc = Me.mDesc.Trim
            If Me.mDesc = "" Then
                MsgBox("Please specify document number.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '2. jika salah satu dari dgv berikut kosong: group, dgv product
            If vDGVGroup.Rows.Count = 0 Then
                MsgBox("Please specify at least 1 group.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If vDGVProduct.Rows.Count = 0 Then
                MsgBox("Please specify at least 1 product.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '3. period from harus lebih kecil dari period to
            If Me.mStartPeriod > Me.mEndPeriod Then
                MsgBox("Start period must be smaller than end period.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '5. jika include sales, include adjustment, dan include bundling semuanya bernilai 0
            If Me.mIncludeSales = 0 And Me.mIncludeBundling = 0 And Me.mIncludeAdjustment = 0 Then
                MsgBox("Please make sure at least 1 of these data is checked: 'Include Sales', 'Include Bundling', or 'Include Adjustment'.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '6. untuk tipe transaksi cumulative percetage, total cumulative percetage harus 100 %
            If Me.mTipe = EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            Me.mTipe = EnumTipe.Nett_Sales_Cumulative_Percentage Or _
            Me.mTipe = EnumTipe.Usage_Qty_Cumulative_Percentage Then
                For Each rowDGV In vDGVGroup.Rows
                    TotalPerc += CDec(rowDGV.Cells(vFrm.ColGroup_RefValue.Index).Value)
                Next
                If TotalPerc <> 100 Then
                    MsgBox("Please make sure the sum of ref. value data is 100%.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '7. untuk tipe transaksi percetage, nilai tidak boleh melebih 100%, baris yang atas harus selalu lebih besar nilainya dari baris yang dibawahnya
            If Me.mTipe = EnumTipe.Gross_Margin_Percentage Or _
            Me.mTipe = EnumTipe.Nett_Sales_Percentage Or _
            Me.mTipe = EnumTipe.Usage_Qty_Percentage Then
                For Each rowDGV In vDGVGroup.Rows
                    If CDec(rowDGV.Cells(vFrm.ColGroup_RefValue.Index).Value) > 100 Then
                        MsgBox("Please make sure maximum ref. value data is always < 100%.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If rowDGV.Index > 0 Then
                        rowDgvBefore = vDGVGroup.Rows(rowDGV.Index - 1)
                        If CDec(rowDGV.Cells(vFrm.ColGroup_RefValue.Index).Value) >= CDec(rowDgvBefore.Cells(vFrm.ColGroup_RefValue.Index).Value) Then
                            MsgBox("Please make sure ref. value data is smaller than the previous row.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                Next
            End If

            '8. untuk setiap item ref value pada dgv group tidak boleh nilai 0
            TipeTrans = Me.GetTipeTransaksi(vFrm.cboGroupBy.Text)
            If TipeTrans = EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            TipeTrans = EnumTipe.Nett_Sales_Cumulative_Percentage Or _
            TipeTrans = EnumTipe.Usage_Qty_Cumulative_Percentage Then
                For Each rowDGV In vDGVGroup.Rows
                    If CDec(rowDGV.Cells(vFrm.ColGroup_RefValue.Index).Value) = 0 Then
                        MsgBox("Please make sure ref. value data is always > 0", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                Next
            End If

            '9. jika tipe transaksi melibatkan gross margin, maka harus sudah closing mulai dari start period - end period
            If Me.mTipe = EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            Me.mTipe = EnumTipe.Gross_Margin_Minimum_Value Or _
            Me.mTipe = EnumTipe.Gross_Margin_Percentage Then
                If ObjClosing.GetPrimaryClosing(Me.mEndPeriod) = -1 Then
                    MsgBox("Please make sure there is inventory closing on " + Format(Me.mEndPeriod, "MMM yyyy"), MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '10. 

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjClosing = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItemGroup As DataGridView, ByVal vDGVItemSupplier As DataGridView, _
    ByVal vdgvItemProductGroup As DataGridView, ByVal vDGVItemProduct As DataGridView, ByVal vFrm As FrmTrans_MaterialPlanningDashboard_Data1, _
    ByVal vUserCode As String, _
    ByVal TBLItem_Invoice As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoiceDataTable, _
    ByVal TBLItem_SR As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_srDataTable, _
    ByVal TBLItem_Bundling As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundlingDataTable, _
    ByVal TBLItem_Adjustment As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustmentDataTable) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter
        Dim ObjItem As New ClsTransaksiMaterialPlanningGroupItem
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True, vdgvItemGroup, vDGVItemProduct, vFrm) = True Then
                'insert query
                If DTA.Insert(Me.mDocumentNo, Me.mTanggalTransaksi, 0, Me.mDesc, Me.mStartPeriod, Me.mEndPeriod, _
                Me.mRemarks, Me.mIncludeSales, Me.mIncludeAdjustment, Me.mIncludeBundling, Me.mTipe, Me.mCodeNoFilter, _
                Me.mProductDescFilter, vUserCode, Now, vUserCode, Now, Now, Me.mStopBeforeRefValue) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mDocumentNo).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'ganti data identifier data pada item group dan item product
                    vdgvItemGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    vDGVItemProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    For Each rowDGV In vdgvItemGroup.Rows
                        rowDGV.Cells(vFrm.ColGroup_Identifier.Index).Value = Me.mPrimary.ToString + _
                        rowDGV.Cells(vFrm.ColGroup_Identifier.Index).Value.ToString.Substring(2)
                    Next
                    For Each rowDGV In vDGVItemProduct.Rows
                        rowDGV.Cells(vFrm.ColCodeNo_IdentifierGroup.Index).Value = Me.mPrimary.ToString + _
                        rowDGV.Cells(vFrm.ColCodeNo_IdentifierGroup.Index).Value.ToString.Substring(2)
                    Next
                    vdgvItemGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    vDGVItemProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

                    'simpan data item baru
                    ObjItem.Simpan(Me.mPrimary, vDGVItemSupplier, vdgvItemGroup, vDGVItemProduct, vdgvItemProductGroup, vFrm, _
                    TBLItem_Invoice, TBLItem_SR, TBLItem_Bundling, TBLItem_Adjustment)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message.StartsWith("Duplicate entry '" + Me.mDocumentNo + "' for key") = True Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Return Me.SimpanBaru(vdgvItemGroup, vDGVItemSupplier, vdgvItemProductGroup, vDGVItemProduct, vFrm, vUserCode, _
                TBLItem_Invoice, TBLItem_SR, TBLItem_Bundling, TBLItem_Adjustment)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vdgvItemGroup As DataGridView, ByVal vDGVItemSupplier As DataGridView, _
    ByVal vdgvItemProductGroup As DataGridView, ByVal vDGVItemProduct As DataGridView, ByVal vFrm As FrmTrans_MaterialPlanningDashboard_Data1, _
    ByVal vUserCode As String, _
    ByVal TBLItem_Invoice As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoiceDataTable, _
    ByVal TBLItem_SR As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_srDataTable, _
    ByVal TBLItem_Bundling As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundlingDataTable, _
    ByVal TBLItem_Adjustment As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustmentDataTable) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter
        Dim ObjItem As New ClsTransaksiMaterialPlanningGroupItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(False, vdgvItemGroup, vDGVItemProduct, vFrm) = True Then
                If DTA.Update(Me.mDocumentNo, Me.mTanggalTransaksi, 0, Me.mDesc, Me.mStartPeriod, Me.mEndPeriod, Me.mRemarks, _
                Me.mIncludeSales, Me.mIncludeAdjustment, Me.mIncludeBundling, Me.mTipe, Me.mCodeNoFilter, Me.mProductDescFilter, _
                vUserCode, Now, Now, Me.mStopBeforeRefValue, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan(Me.mPrimary, vDGVItemSupplier, vdgvItemGroup, vDGVItemProduct, vdgvItemProductGroup, vFrm, _
                    TBLItem_Invoice, TBLItem_SR, TBLItem_Bundling, TBLItem_Adjustment)
                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message.StartsWith("Duplicate entry '" + Me.mDocumentNo + "' for key") = True Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Return Me.SimpanBaru(vdgvItemGroup, vDGVItemSupplier, vdgvItemProductGroup, vDGVItemProduct, vFrm, vUserCode, _
                TBLItem_Invoice, TBLItem_SR, TBLItem_Bundling, TBLItem_Adjustment)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vStrictDelete As Boolean, ByVal vUserCode As String) As Boolean
        Dim Obj As New ClsTransaksiMaterialPlanningGroupItem
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vStrictDelete, vUserCode, vPrimaryTransaksi) = True Then
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj = Nothing
        End Try
    End Function

    Public Function GetDocumentNo(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter
        Dim RetVal As String

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            RetVal = DTA.ScalarQuery_GetDocumentNo(vPrimary)
            If String.IsNullOrEmpty(RetVal) = True Then
                Return ""
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vDocumenNo As String) As Double
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            RetVal = DTA.ScalarQuery_GetPrimary(vDocumenNo)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vDocumentNo As String) As Double
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            RetVal = DTA.ScalarQuery_GetJml(vDocumentNo)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group1TableAdapter
        Dim Jml As Nullable(Of Long)
        Dim ProductGroupLevel As Integer = 1
        Dim Level As ClsTransaksiMaterialPlanningGroup.ProductLevel = ProductLevel.Brand

        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                'by no trnsksi
                Jml = DTA.ScalarQuery_GetJmlByDocumentNoLIKE(Criteria + "%")
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                'by remarks
                Jml = DTA.ScalarQuery_GetJmlByRemarksLIKE(Criteria + "%")
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                'by desc
                Jml = DTA.ScalarQuery_GetJmlByDescLIKE(Criteria + "%")
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                'by date
                Jml = DTA.ScalarQuery_GetJmlByTgl(vTanggalSampai, vTanggalDari)
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByCategoryCode Then
                'by product category code
                Level = ProductLevel.Category
                Jml = DTA.ScalarQuery_GetJmlByCategoryCodeLIKE(Level, Criteria + "%")
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByCategoryDesc Then
                'by product category desc
                Level = ProductLevel.Category
                Jml = DTA.ScalarQuery_GetJmlByCategoryDescLIKE(Level, Criteria + "%")
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.BySubBrandCode Then
                'by product sub brand code
                ProductGroupLevel = 1
                Level = ProductLevel.SubBrand
                Jml = DTA.ScalarQuery_GetJmlByProductGroupCodeLIKE(Level, Criteria + "%", ProductGroupLevel)
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.BySubBrandDesc Then
                'by product sub brand desc
                ProductGroupLevel = 1
                Level = ProductLevel.SubBrand
                Jml = DTA.ScalarQuery_GetJmlByProductGroupDescLIKE(Level, Criteria + "%", ProductGroupLevel)
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByBrandCode Then
                'by product brand code
                ProductGroupLevel = 2
                Level = ProductLevel.Brand
                Jml = DTA.ScalarQuery_GetJmlByProductGroupCodeLIKE(Level, Criteria + "%", ProductGroupLevel)
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByBrandDesc Then
                'by product brand desc
                ProductGroupLevel = 2
                Level = ProductLevel.Brand
                Jml = DTA.ScalarQuery_GetJmlByProductGroupDescLIKE(Level, ProductGroupLevel, Criteria + "%")
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductGroupCode Then
                'by product group code
                ProductGroupLevel = 3
                Level = ProductLevel.ProductGroup
                Jml = DTA.ScalarQuery_GetJmlByProductGroupCodeLIKE(Level, Criteria + "%", ProductGroupLevel)
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductGroupDesc Then
                'by product group desc
                ProductGroupLevel = 3
                Level = ProductLevel.ProductGroup
                Jml = DTA.ScalarQuery_GetJmlByProductGroupDescLIKE(Level, ProductGroupLevel, Criteria + "%")
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                'by supplier code
                Jml = DTA.ScalarQuery_GetJmlBySupplierCodeLIKE(Criteria + "%")
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaSupplier Then
                'by supplier code
                Jml = DTA.ScalarQuery_GetJmlBySupplierDescLIKE(Criteria + "%")
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                'by product code desc
                Jml = DTA.ScalarQuery_GetJmlProductDescLIKE(Criteria + "%")
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                'by product code no
                Jml = DTA.ScalarQuery_GetJmlByCodeNoLIKE(Criteria + "%")
                If Jml.HasValue = True Then
                    Return Jml.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Dim ProductGroupLevel As Integer = 1
        Dim Level As ClsTransaksiMaterialPlanningGroup.ProductLevel = ProductLevel.Brand

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_DocumentNoLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                'by remarks
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                'by desc
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                'by date
                QueryResult = DTA.FillBy_Date(TBL, vTanggalSampai, vTanggalDari)
            ElseIf SearchCriteria = EnumSearchCriteria.ByCategoryCode Then
                'by product category code
                Level = ProductLevel.Category
                QueryResult = DTA.FillBy_CategoryCode(TBL, Level, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCategoryDesc Then
                'by product category desc
                Level = ProductLevel.Category
                QueryResult = DTA.FillBy_CategoryDesc(TBL, Level, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySubBrandCode Then
                'by product sub brand code
                ProductGroupLevel = 1
                Level = ProductLevel.SubBrand
                QueryResult = DTA.FillBy_ProductGroupCodeLIKE(TBL, Level, Criteria + "%", ProductGroupLevel)
            ElseIf SearchCriteria = EnumSearchCriteria.BySubBrandDesc Then
                'by product sub brand desc
                ProductGroupLevel = 1
                Level = ProductLevel.SubBrand
                QueryResult = DTA.FillBy_ProductGroupDescLIKE(TBL, Level, ProductGroupLevel, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByBrandCode Then
                'by product group code
                ProductGroupLevel = 2
                Level = ProductLevel.Brand
                QueryResult = DTA.FillBy_ProductGroupCodeLIKE(TBL, Level, Criteria + "%", ProductGroupLevel)
            ElseIf SearchCriteria = EnumSearchCriteria.ByBrandDesc Then
                'by product brand desc
                ProductGroupLevel = 2
                Level = ProductLevel.Brand
                QueryResult = DTA.FillBy_ProductGroupDescLIKE(TBL, Level, ProductGroupLevel, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductGroupCode Then
                'by product group code
                ProductGroupLevel = 3
                Level = ProductLevel.ProductGroup
                QueryResult = DTA.FillBy_ProductGroupCodeLIKE(TBL, Level, Criteria + "%", ProductGroupLevel)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductGroupDesc Then
                'by product group desc
                ProductGroupLevel = 3
                Level = ProductLevel.ProductGroup
                QueryResult = DTA.FillBy_ProductGroupDescLIKE(TBL, Level, ProductGroupLevel, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                'by supplier code
                QueryResult = DTA.FillBy_SupplierCode(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaSupplier Then
                'by supplier name
                QueryResult = DTA.FillBy_SupplierNameLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                'by product code desc
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                'by product code no
                QueryResult = DTA.FillBy_CodeNoLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(6).HeaderText = "Remarks"
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(4).DefaultCellStyle.Format = "MMM yyyy"
                DGV.Columns(5).DefaultCellStyle.Format = "MMM yyyy"
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetTimestamp(vPrimary)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipeTransaksi_ByPrimary(ByVal vPrimary As Double) As EnumTipe
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter
        Dim RetVal As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetTipe(vPrimary)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return EnumTipe.Gross_Margin_Cumulative_Percentage
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumTipe.Gross_Margin_Cumulative_Percentage
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipeTransaksi(ByVal vTipeTransStr As String) As EnumTipe
        Try
            If vTipeTransStr = "Gross Margin Cumulative Percentage" Then
                Return EnumTipe.Gross_Margin_Cumulative_Percentage
            ElseIf vTipeTransStr = "Nett Sales Cumulative Percentage" Then
                Return EnumTipe.Nett_Sales_Cumulative_Percentage
            ElseIf vTipeTransStr = "Usage Qty Cumulative Percentage" Then
                Return EnumTipe.Usage_Qty_Cumulative_Percentage
            ElseIf vTipeTransStr = "Gross Margin Percentage" Then
                Return EnumTipe.Gross_Margin_Percentage
            ElseIf vTipeTransStr = "Nett Sales Percentage" Then
                Return EnumTipe.Nett_Sales_Percentage
            ElseIf vTipeTransStr = "Usage Qty Percentage" Then
                Return EnumTipe.Usage_Qty_Percentage
            ElseIf vTipeTransStr = "Gross Margin Minimum Value" Then
                Return EnumTipe.Gross_Margin_Minimum_Value
            ElseIf vTipeTransStr = "Nett Sales Minimum Value" Then
                Return EnumTipe.Nett_Sales_Minimum_Value
            ElseIf vTipeTransStr = "Usage Qty Minimum Value" Then
                Return EnumTipe.Usage_Qty_Minimum_Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetTipeTransaksiStr(ByVal vTipeTrans As EnumTipe) As String
        Try
            If vTipeTrans = EnumTipe.Gross_Margin_Cumulative_Percentage Then
                Return "Gross Margin Cumulative Percentage"
            ElseIf vTipeTrans = EnumTipe.Nett_Sales_Cumulative_Percentage Then
                Return "Nett Sales Cumulative Percentage"
            ElseIf vTipeTrans = EnumTipe.Usage_Qty_Cumulative_Percentage Then
                Return "Usage Qty Cumulative Percentage"
            ElseIf vTipeTrans = EnumTipe.Gross_Margin_Percentage Then
                Return "Gross Margin Percentage"
            ElseIf vTipeTrans = EnumTipe.Nett_Sales_Percentage Then
                Return "Nett Sales Percentage"
            ElseIf vTipeTrans = EnumTipe.Usage_Qty_Percentage Then
                Return "Usage Qty Percentage"
            ElseIf vTipeTrans = EnumTipe.Gross_Margin_Minimum_Value Then
                Return "Gross Margin Minimum Value"
            ElseIf vTipeTrans = EnumTipe.Nett_Sales_Minimum_Value Then
                Return "Nett Sales Minimum Value"
            ElseIf vTipeTrans = EnumTipe.Usage_Qty_Minimum_Value Then
                Return "Usage Qty Minimum Value"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub Cetak(ByVal vPrimaryTrans As Double, ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vReport As CrpTransaksiMaterialPlanningGroup_Cetak, ByVal DS As DseTransaksiMaterialPlanningGroup)
        Dim DTAItemGroup As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_groupTableAdapter
        Dim TBLItemGroup As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_groupDataTable = DS.tbltrans_material_planning_group_item_group
        Dim rowItemGroup As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_groupRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim Strategy1 As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy = EnumLeadTimeStrategy.Based_on_prefered_supplier
        Dim ObjItem As New ClsTransaksiMaterialPlanningGroupItem
        Dim DTAItemProduct As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_product1TableAdapter
        Dim TBLItemProduct As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_product1DataTable = DS.tbltrans_material_planning_group_item_product1
        Dim rowItemProduct As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_product1Row
        Dim TipeTrans As EnumTipe = EnumTipe.Gross_Margin_Cumulative_Percentage
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim SubTotalAmountGroup As Decimal = 0
        Dim SubTotalQtyGroup As Decimal = 0
        Dim ListGroupCode As New List(Of String)
        Dim TotalQty As Decimal = 0
        Dim TotalAmount As Decimal = 0
        Dim DV As Data.DataRowView
        Dim KodeGroup As String = ""
        Dim Total1 As Decimal = 0
        Dim Total2 As Decimal = 0
        Dim Total3 As Decimal = 0
        Dim Total4 As Decimal = 0
        Dim TotalDummy1 As Decimal = 0
        Dim TotalDummy2 As Decimal = 0
        Dim DefPrecBase As Integer = 0

        Try
            DTAItemGroup.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemProduct.Connection.ConnectionString = DTAItemGroup.Connection.ConnectionString
            TipeTrans = Me.GetTipeTransaksi_ByPrimary(vPrimaryTrans)
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency

            'item group
            TBLItemGroup.Rows.Clear()
            DTAItemGroup.FillBy_PrimaryTransaksi(TBLItemGroup, vPrimaryTrans)
            For Each rowItemGroup In TBLItemGroup.Rows
                With rowItemGroup
                    .optimum_stock_strategy_str = ObjItem.GetLeadTimeStrategyStr(.optimum_stock_strategy)
                    Strategy1 = .optimum_stock_strategy
                    If Strategy1 = EnumLeadTimeStrategy.Based_on_usage Then
                        .optimum_stock_qy_str = ObjCD.BuangNolDibelakangDecimal(ObjCD.DigitGrouping(.optimum_stock_days))
                        .optimum_stock_strategy_unit = "days"
                    ElseIf Strategy1 = EnumLeadTimeStrategy.User_Defined Then
                        .optimum_stock_qy_str = ObjCD.BuangNolDibelakangDecimal(ObjCD.DigitGrouping(.optimum_stock_qty))
                        .optimum_stock_strategy_unit = "unit"
                    End If

                    .safety_stock_strategy_str = ObjItem.GetLeadTimeStrategyStr(.safety_stock_strategy)
                    Strategy1 = .safety_stock_strategy
                    If Strategy1 = EnumLeadTimeStrategy.Based_on_usage Then
                        .safety_stock_qty_str = ObjCD.BuangNolDibelakangDecimal(ObjCD.DigitGrouping(.safety_stock_days))
                        .safety_stock_stretegy_unit = "days"
                    ElseIf Strategy1 = EnumLeadTimeStrategy.User_Defined Then
                        .safety_stock_qty_str = ObjCD.BuangNolDibelakangDecimal(ObjCD.DigitGrouping(.safety_stock_qty))
                        .safety_stock_stretegy_unit = "unit"
                    End If

                    .lead_time_stock_strategy = ObjItem.GetLeadTimeStrategyStr(.lead_time_strategy)
                    Strategy1 = .lead_time_strategy
                    If Strategy1 = EnumLeadTimeStrategy.Based_on_usage Then
                        .lead_time_stock_str = ObjCD.BuangNolDibelakangDecimal(ObjCD.DigitGrouping(.lead_time_days))
                        .lead_time_stock_strategy_unit = "days"
                    ElseIf Strategy1 = EnumLeadTimeStrategy.User_Defined Then
                        .lead_time_stock_str = ObjCD.BuangNolDibelakangDecimal(ObjCD.DigitGrouping(.lead_time_qty))
                        .lead_time_stock_strategy_unit = "unit"
                    End If

                    .procurement_strategy_str = ObjItem.GetProcurementStrategyStr(.procurement_strategy)
                End With
            Next
            If TipeTrans = EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            TipeTrans = EnumTipe.Gross_Margin_Percentage Or _
            TipeTrans = EnumTipe.Nett_Sales_Cumulative_Percentage Or _
            TipeTrans = EnumTipe.Nett_Sales_Percentage Or _
            TipeTrans = EnumTipe.Usage_Qty_Cumulative_Percentage Or _
            TipeTrans = EnumTipe.Usage_Qty_Percentage Then
                'decimal ref value = 2
                For Each rowItemGroup In TBLItemGroup.Rows
                    rowItemGroup.ref_value_str = ObjCD.DigitGrouping(rowItemGroup.ref_value, 2) + " %"
                Next
            ElseIf TipeTrans = EnumTipe.Gross_Margin_Minimum_Value Or _
            TipeTrans = EnumTipe.Nett_Sales_Minimum_Value Then
                'decimal IDR
                For Each rowItemGroup In TBLItemGroup.Rows
                    rowItemGroup.ref_value_str = ObjCD.DigitGrouping(rowItemGroup.ref_value, DefPrecBase)
                Next
            ElseIf Tipe = EnumTipe.Usage_Qty_Minimum_Value Then
                'buang nol di belakang koma, max 6 decimal
                For Each rowItemGroup In TBLItemGroup.Rows
                    rowItemGroup.ref_value_str = ObjCD.BuangNolDibelakangDecimal(ObjCD.DigitGrouping(rowItemGroup.ref_value, 6))
                Next
            End If

            'item product
            DTAItemProduct.Fill(TBLItemProduct, vPrimaryTrans)
            If MsgBox("Exclude zero amount item?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                For ij As Double = TBLItemProduct.Rows.Count - 1 To 0 Step -1
                    rowItemProduct = TBLItemProduct.Rows(ij)
                    With rowItemProduct
                        If .nett_sales = 0 And .total_qty = 0 And .gross_margin = 0 Then
                            TBLItemProduct.Rows.RemoveAt(ij)
                        End If
                    End With
                Next
            End If
            TotalAmount = 0
            TotalQty = 0
            For Each rowItemProduct In TBLItemProduct.Rows
                With rowItemProduct
                    If TipeTrans = EnumTipe.Gross_Margin_Cumulative_Percentage Or _
                    TipeTrans = EnumTipe.Gross_Margin_Minimum_Value Or _
                    TipeTrans = EnumTipe.Gross_Margin_Percentage Then
                        .amount1 = .gross_margin
                        .amount1_avg = .gross_margin_average
                        .amount1_perc = .gross_margin_perc
                        TotalAmount += .gross_margin
                    Else
                        .amount1 = .nett_sales
                        .amount1_avg = .nett_sales_average
                        .amount1_perc = .nett_sales_perc
                        TotalAmount += .nett_sales
                    End If
                    If ListGroupCode.Contains(.group_code) = False Then
                        ListGroupCode.Add(.group_code)
                    End If
                    TotalQty += .total_qty
                End With
            Next

            'hitung sub total
            For Each KodeGroup In ListGroupCode
                TBLItemProduct.DefaultView.RowFilter = "[group code] = '" + KodeGroup + "'"
                SubTotalAmountGroup = 0
                SubTotalQtyGroup = 0
                For Each DV In TBLItemProduct.DefaultView
                    rowItemProduct = DV.Row
                    With rowItemProduct
                        If TipeTrans = EnumTipe.Gross_Margin_Cumulative_Percentage Or _
                        TipeTrans = EnumTipe.Gross_Margin_Minimum_Value Or _
                        TipeTrans = EnumTipe.Gross_Margin_Percentage Then
                            SubTotalAmountGroup += .gross_margin
                        Else
                            SubTotalAmountGroup += .nett_sales
                        End If
                        SubTotalQtyGroup += .total_qty
                    End With
                Next
                Total1 = Math.Round(((SubTotalAmountGroup + TotalDummy1) / TotalAmount) * 100, 2, MidpointRounding.AwayFromZero)
                Total2 = Math.Round(((SubTotalQtyGroup + TotalDummy2) / TotalQty) * 100, 2, MidpointRounding.AwayFromZero)
                Total3 = Math.Round(((SubTotalAmountGroup) / TotalAmount) * 100, 2, MidpointRounding.AwayFromZero)
                Total4 = Math.Round(((SubTotalQtyGroup) / TotalQty) * 100, 2, MidpointRounding.AwayFromZero)
                For Each DV In TBLItemProduct.DefaultView
                    rowItemProduct = DV.Row
                    With rowItemProduct
                        .group_amount_perc = Total1
                        .group_qty_perc = Total2
                        .group_amount1_perc = Total3
                        .group_qty1_perc = Total4
                    End With
                Next

                TotalDummy1 += SubTotalAmountGroup
                TotalDummy2 += SubTotalQtyGroup
            Next

            'tampilkan cetakan
            vReport.SetDataSource(DS)
            vReport.SetParameterValue(vReport.Parameter_DocumentNo.ParameterFieldName, Me.GetDocumentNo(vPrimaryTrans))
            If TipeTrans = EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            TipeTrans = EnumTipe.Gross_Margin_Minimum_Value Or _
            TipeTrans = EnumTipe.Gross_Margin_Percentage Then
                vReport.SetParameterValue(vReport.Parameter_Amount1_Avg.ParameterFieldName, "Avg. GM")
                vReport.SetParameterValue(vReport.Parameter_Amount1_Str.ParameterFieldName, "Gross Margin")
            Else
                vReport.SetParameterValue(vReport.Parameter_Amount1_Avg.ParameterFieldName, "Avg. NS")
                vReport.SetParameterValue(vReport.Parameter_Amount1_Str.ParameterFieldName, "Nett Sales")
            End If
            vReport.SetParameterValue(vReport.Parameter_DefPrecIDR.ParameterFieldName, DefPrecBase)
            vReport.SetParameterValue(vReport.Parameter_EndDate.ParameterFieldName, Me.GetEndPeriod(vPrimaryTrans))
            vReport.SetParameterValue(vReport.Parameter_StartDate.ParameterFieldName, Me.GetStartPeriod(vPrimaryTrans))
            vReport.SetParameterValue(vReport.Parameter_Group_By.ParameterFieldName, Me.GetTipeTransaksiStr(TipeTrans))

            vCrv.ReportSource = vReport
            vCrv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemGroup.Dispose()
            ObjCD = Nothing
            ObjItem = Nothing
            DTAItemProduct.Dispose()
            ObjCur = Nothing
            ListGroupCode = Nothing
        End Try
    End Sub

    Public Function GetStartPeriod(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetStartPeriod(vPrimary)
            If RetVal.HasValue = True Then
                Return RetVal
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetEndPeriod(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_groupTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetEndPeriod(vPrimary)
            If RetVal.HasValue = True Then
                Return RetVal
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
