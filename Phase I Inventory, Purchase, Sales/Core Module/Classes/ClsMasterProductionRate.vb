Option Explicit On
Option Strict Off

Public Class ClsMasterProductionRate
    Dim objCd As New Core_Module.ClsCekData
    Dim objItem As New ClsMasterProductionRate_Item
    Dim objProduct As New clsMasterProduct
    Dim objCategory As New ClsMasterCategory
    Dim objGroup As New ClsMasterProductGroup
    Dim mTipeDependency As clsMasterProduct.EnumTipeDependency = clsMasterProduct.EnumTipeDependency.End_Start1
    Dim mDelayDependency As Double = 0
    Enum EnumSearchBy As Integer
        DOCUMENT_NO
        _DATE
        PROCESS_CODE
        PROCESS_DESC
        REMARKS
    End Enum
    Public Enum ProductLevel As Integer
        ProductGroup
        Brand
        SubBrand
        Category
        Product
    End Enum

    Private mPrimary As Double
    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Private mUrutan As Double
    Public Property Urutan() As Double
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Double)
            mUrutan = value
        End Set
    End Property

    Private mDocumentNo As String
    Public Property DocumentNo() As String
        Get
            Return mDocumentNo
        End Get
        Set(ByVal value As String)
            mDocumentNo = value
        End Set
    End Property

    Private mDate As Date
    Public Property _Date() As Date
        Get
            Return mDate
        End Get
        Set(ByVal value As Date)
            mDate = value
        End Set
    End Property

    Private mPrimaryProcess As Double
    Public Property PrimaryProcess() As Double
        Get
            Return mPrimaryProcess
        End Get
        Set(ByVal value As Double)
            mPrimaryProcess = value
        End Set
    End Property

    Private mRemarks As String
    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Private mCreatedBy As String
    Public Property CreatedBy() As String
        Get
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Private mCreatedDate As Date
    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Private mModifiedBy As String
    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Private mModifiedDate As Date
    Public Property ModifiedDate() As Date
        Get
            Return mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Private mTimeStamp As Date
    Public Property TimeStamp() As Date
        Get
            Return mTimeStamp
        End Get
        Set(ByVal value As Date)
            mTimeStamp = value
        End Set
    End Property

    Public Property TipeDependency() As clsMasterProduct.EnumTipeDependency
        Get
            Return Me.mTipeDependency
        End Get
        Set(ByVal value As clsMasterProduct.EnumTipeDependency)
            Me.mTipeDependency = value
        End Set
    End Property

    Public Property DelayDependency() As Double
        Get
            Return Me.mDelayDependency
        End Get
        Set(ByVal value As Double)
            Me.mDelayDependency = value
        End Set
    End Property

    Public Function SimpanBaru(ByVal vUserCode As String, ByVal vFrm As FrmMasterProductionRate) As Boolean
        Dim dta As New DseMasterProductionRateTableAdapters.tblmsater_production_rateTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Me.mUrutan = CInt(Me.mDocumentNo.Substring(0, 6))
            If dta.InsertQuery(Me.mUrutan, Me.mDocumentNo, Me.mDate, Me.mPrimaryProcess, Me.mRemarks, vUserCode, Now, vUserCode, Now, Me.mTipeDependency, Me.mDelayDependency) <= 0 Then Return False
            Me.mPrimary = Me.GetPrimary(Me.mDocumentNo)
            Me.mTimeStamp = Me.GetTimestamp(Me.mPrimary)
            Me.mCreatedDate = Me.mTimeStamp
            Me.mModifiedDate = Me.mTimeStamp
            Me.mCreatedBy = vUserCode
            Me.mModifiedBy = vUserCode
            If Me.objItem.SimpanItem(vFrm, Me.mPrimary) = False Then Return False
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vUserCode As String, ByVal vFrm As FrmMasterProductionRate) As Boolean
        Dim dta As New DseMasterProductionRateTableAdapters.tblmsater_production_rateTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Me.mUrutan = CInt(Me.mDocumentNo.Substring(0, 6))
            If dta.UpdateQuery(Me.mUrutan, Me.mDocumentNo, Me.mDate, Me.mPrimaryProcess, Me.mRemarks, vUserCode, Now, Me.mTipeDependency, Me.mDelayDependency, Me.mPrimary) <= 0 Then Return False
            Me.mTimeStamp = Me.GetTimestamp(Me.mPrimary)
            Me.mModifiedDate = Me.mTimeStamp
            Me.mModifiedBy = vUserCode
            Me.objItem.HapusItem(Me.mPrimary)
            If Me.objItem.SimpanItem(vFrm, Me.mPrimary) = False Then Return False
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim dta As New DseMasterProductionRateTableAdapters.tblmsater_production_rateTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.DeleteQuery(vPrimary) <= 0 Then Return False
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vDocumentNo As String, ByVal vShowMessage As Boolean) As Boolean
        Try
            If vBaru = True Then
                If Me.ApakahSudahAda(vBaru, Me.mDocumentNo) = True Then
                    Me.mDocumentNo = Me.GetNoBaru(Me.mDate)
                End If
            Else
                Dim documentNo_ As String = Me.GetDocumentNo(Me.mPrimary)
                If documentNo_ <> Me.mDocumentNo Then
                    If Me.ApakahSudahAda(vBaru, Me.mDocumentNo) = True Then
                        Me.mDocumentNo = Me.GetNoBaru(Me.mDate)
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vStrictEdit As Boolean, ByVal vUserCode As String, ByVal vPrimary As Double, ByVal vShowMessage As Boolean) As Boolean
        Try
            If vStrictEdit Then
                If Me.mCreatedBy <> vUserCode Then
                    If vShowMessage Then MsgBox("Only user with user code: " & Me.mCreatedBy & " can edit this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vStrictDelete As Boolean, ByVal vUserCode As String, ByVal vPrimary As Double, ByVal vShowMessage As Boolean) As Boolean
        Try
            If vStrictDelete Then
                If Me.mCreatedBy <> vUserCode Then
                    If vShowMessage Then MsgBox("Only user with usercode: " & Me.mCreatedBy & " can delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function GetPrimary(ByVal vDocumentNo As String) As Double
        Dim dta As New DseMasterProductionRateTableAdapters.tblmsater_production_rateTableAdapter
        Dim v As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            v = dta.ScalarQuery_GetPrimary(vDocumentNo)
            If v.HasValue Then Return v.Value
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Public Function GetDocumentNo(ByVal vPrimary As String) As String
        Dim dta As New DseMasterProductionRateTableAdapters.tblmsater_production_rateTableAdapter
        Dim v As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            v = dta.ScalarQuery_GetDocumentNo(vPrimary)
            If String.IsNullOrEmpty(v) = False Then Return v
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetDate(ByVal vPrimary As Boolean) As Date
        Dim dta As New DseMasterProductionRateTableAdapters.tblmsater_production_rateTableAdapter
        Dim v As New Nullable(Of Date)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            v = dta.ScalarQuery_GetDate(vPrimary)
            If v.HasValue = False Then Return v.Value
            Return #1/1/1950#
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim dta As New DseMasterProductionRateTableAdapters.tblmsater_production_rateTableAdapter
        Dim v As New Nullable(Of Date)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            v = dta.ScalarQuery_GetTimestamp(vPrimary)
            If v.HasValue Then Return v.Value
            Return #1/1/1950#
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vFrm As FrmMasterProductionRate) As Boolean
        Dim dta As New DseMasterProductionRateTableAdapters.tblmsater_production_rateTableAdapter
        Dim tbl As New DseMasterProductionRate.tblmsater_production_rateDataTable
        Dim row As DseMasterProductionRate.tblmsater_production_rateRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Fill(tbl, vPrimary) > 0 Then
                row = tbl.Rows(0)
                With Me
                    .mPrimary = row.primary
                    .mUrutan = row.urutan
                    .mDocumentNo = row.document_no
                    .mDate = row._date
                    .mPrimaryProcess = row.primary_process
                    .mRemarks = row.remarks
                    .mCreatedBy = row.created_by
                    .mCreatedDate = row.created_date
                    .mModifiedBy = row.modified_by
                    .mModifiedDate = row.modified_date
                    .mTimeStamp = row.timestamp
                    .mTipeDependency = row.tipe_dependency
                    .mDelayDependency = row.delay_dependency
                End With
                Me.objItem.IsiDgvItem(vFrm, vPrimary)
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function GetJml(ByVal vDocumentNo As String) As Integer
        Dim dta As New DseMasterProductionRateTableAdapters.tblmsater_production_rateTableAdapter
        Dim v As New Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            v = dta.ScalarQuery_GetJml(vDocumentNo)
            If v.HasValue Then Return v.Value
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function GetJml(ByVal vCriteria As String, ByVal vSearchBy As ClsMasterProductionRate.EnumSearchBy, _
            ByVal vDari As Date, ByVal vSampai As Date) As Double
        Dim DTA As New DseMasterProductionRateTableAdapters.DataTable1TableAdapter
        Dim v As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchBy.DOCUMENT_NO Then
                v = DTA.ScalarQuery_GetJmlByDocNoLIKE(vCriteria + "%")
                If v.HasValue Then
                    Return v.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchBy.PROCESS_CODE Then
                v = DTA.ScalarQuery_GetJmlByProcessCodeLIKE(vCriteria + "%")
                If v.HasValue Then
                    Return v.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchBy.PROCESS_DESC Then
                v = DTA.ScalarQuery_GetJmlByProcessDescLIKE(vCriteria + "%")
                If v.HasValue Then
                    Return v.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchBy.REMARKS Then
                v = DTA.ScalarQuery_GetJmlByRemarksLIKE(vCriteria + "%")
                If v.HasValue Then
                    Return v.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchBy._DATE Then
                v = DTA.ScalarQuery_GetJmlByDate(vDari, vSampai)
                If v.HasValue Then
                    Return v.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDgv(ByVal vCriteria As String, ByVal vDgv As DataGridView, ByVal vTbl As DseMasterProductionRate.DataTable1DataTable, _
        ByVal vSearchCriteria As ClsMasterProductionRate.EnumSearchBy, ByVal vTsl As ToolStripLabel, ByVal vDari As Date, ByVal vSampai As Date) As Boolean
        Dim dta As New DseMasterProductionRateTableAdapters.DataTable1TableAdapter
        Dim queryResult As Double
        Dim startT As Date

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            startT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If vSearchCriteria = EnumSearchBy.DOCUMENT_NO Then
                queryResult = dta.FillBy_DocumentNoLIKE(vTbl, vCriteria + "%")
            ElseIf vSearchCriteria = EnumSearchBy.PROCESS_CODE Then
                queryResult = dta.FillBy_ProcessCodeLIKE(vTbl, vCriteria + "%")
            ElseIf vSearchCriteria = EnumSearchBy.PROCESS_DESC Then
                queryResult = dta.FillBy_ProcessDescLIKE(vTbl, vCriteria + "%")
            ElseIf vSearchCriteria = EnumSearchBy.REMARKS Then
                queryResult = dta.FillBy_RemarksLIKE(vTbl, vCriteria + "%")
            ElseIf vSearchCriteria = EnumSearchBy._DATE Then
                queryResult = dta.FillBy_Date(vTbl, vDari, vSampai)
            End If

            vDgv.DataSource = vTbl.DefaultView
            vDgv.Refresh()
            If vDgv.ColumnCount > 0 Then
                vDgv.Columns(0).Visible = False
                vDgv.Columns(0).Frozen = True
                vDgv.Columns(1).Frozen = True
                vDgv.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
            End If

            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Dim TS As TimeSpan = Now - startT

            If queryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    vTsl.Text = CStr(queryResult) + " result(s) in " + objCd.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    vTsl.Text = CStr(queryResult) + " result(s) in " + objCd.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                vTsl.Text = "No result found"
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vBaru As Boolean, ByVal vDocumentNo As String) As Boolean
        Try
            If Me.GetJml(vDocumentNo) <= 0 Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function getNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseMasterProductionRateTableAdapters.tblmsater_production_rateTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim v As New Nullable(Of Double)
        Try
            TanggalStart = objCd.GetFisrtDateOfTheYear(vTanggalTransaksi)
            TanggalEnd = objCd.GetLastDateOfTheYear(vTanggalTransaksi)
            v = DTA.ScalarQuery_GetMaxUrutan(TanggalStart, TanggalEnd)
            If v.HasValue = True Then
                NoUrut = CStr(v.Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            While NoUrut.Length < 6
                NoUrut = "0" + NoUrut
            End While

            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While
            Return NoUrut + "/RATE/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsExistProcessAndProduct(ByVal vPrimaryProcess As Double, ByVal vGroupLevel As Double, ByVal vPrimaryRefParty As Double, _
    ByVal vPrimaryTrans As Double) As Boolean
        Dim dta As New DseMasterProductionRateTableAdapters.tblmsater_production_rate_item_workunitTableAdapter
        Dim v As New Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            v = dta.ScalarQuery_GetJmlByProcessAndProduct(vPrimaryProcess, vGroupLevel, vPrimaryRefParty, vPrimaryTrans)
            If v.HasValue Then
                If v.Value > 0 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function getMinimalGroupLevel(ByVal vPrimary As Double) As Integer
        Dim dta As New DseMasterProductionRateTableAdapters.tbltrans_production_rate_item_productTableAdapter
        Dim v As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            v = dta.ScalarQuery_GetMinimalGroupLevel(vPrimary)
            If v.HasValue Then
                Return v.Value
            End If
            Return 4
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 4
        End Try
    End Function

    Public Function GetProductionRate(ByVal vPrimaryProduct As Double, ByVal vPrimaryProcess As Double, ByVal vPrimaryCategory As Double) As Double
        Dim primaryProdRate As Double = -1
        Try
            primaryProdRate = Me.getPrimaryTranksiWithParams(4, vPrimaryProduct, vPrimaryProcess)
            If primaryProdRate = -1 Then
                primaryProdRate = Me.getPrimaryTranksiWithParams(3, vPrimaryCategory, vPrimaryProcess)
            End If
            If primaryProdRate = -1 Then
                primaryProdRate = Me.getPrimaryTranksiWithParams(2, getPrimarySubBrandByPrimaryCategory(vPrimaryCategory), vPrimaryProcess)
            End If
            If primaryProdRate = -1 Then
                primaryProdRate = Me.getPrimaryTranksiWithParams(1, objCategory.GetPrimaryBrand(vPrimaryCategory), vPrimaryProcess)
            End If
            If primaryProdRate = -1 Then
                primaryProdRate = Me.getPrimaryTranksiWithParams(0, Me.getParentPrimary(objCategory.GetPrimaryBrand(vPrimaryCategory)), vPrimaryProcess)
            End If

            Return primaryProdRate
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Function getPrimaryTranksiWithParams(ByVal vGroupLevel As Double, ByVal vPrimaryRefParty As Double, ByVal vPrimaryProcess As Double) As Double
        Dim dta As New DseMasterProductionRateTableAdapters.tbltrans_production_rate_item_productTableAdapter
        Dim v As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            v = dta.ScalarQuery_GetJmlByGroupAndPrimary(vGroupLevel, vPrimaryRefParty, vPrimaryProcess)
            If v.HasValue Then Return v.Value
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Function getPrimarySubBrandByPrimaryCategory(ByVal vPrimaryCategory As Double) As Double
        Dim dta As New DseMasterProductionRateTableAdapters.tbltrans_production_rate_item_productTableAdapter
        Dim n As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            n = dta.ScalarQuery_GetPrimarySubBrand(vPrimaryCategory)
            If n.HasValue Then
                Return n.Value
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Function getParentPrimary(ByVal vPrimaryChild As Double) As Double
        Dim dta As New DseMasterProductionRateTableAdapters.tbltrans_production_rate_item_productTableAdapter
        Dim n As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            n = dta.ScalarQuery_GetParentPrimary(vPrimaryChild)
            If n.HasValue Then
                Return n.Value
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Public Function UpdateDgvWorkUnit(ByVal vDgv As DataGridView, ByVal vFrm As frmMasterProduct, ByVal vPrimaryProdRate As Double) As Boolean
        Dim dtaWU As New DseMasterProductionRateTableAdapters.DataTable2TableAdapter
        Dim tblWU As New DseMasterProductionRate.DataTable2DataTable
        Dim rowWU As DseMasterProductionRate.DataTable2Row
        Try
            dtaWU.Connection.ConnectionString = Core_Module.GetConstring
            dtaWU.Fill(tblWU, vPrimaryProdRate)
            vDgv.Rows.Clear()
            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgv.AllowUserToAddRows = False
            For Each rowWU In tblWU
                Dim dr As DataGridViewRow = vDgv.Rows(vDgv.Rows.Add())
                With dr
                    .Cells(vFrm.ColWorkUnit_Code.Index).Value = rowWU.Code
                    .Cells(vFrm.ColWorkUnit_Desc.Index).Value = rowWU.Desc
                    .Cells(vFrm.ColWorkUnit_Efficiency.Index).Value = Me.objCd.DigitGrouping(rowWU.eficiency, 2)
                    .Cells(vFrm.ColWorkUnit_No.Index).Value = vDgv.Rows.Count
                    .Cells(vFrm.ColWorkUnit_PrimaryWorkUnit.Index).Value = rowWU.primary_work_unit
                    .Cells(vFrm.ColWorkUnit_Rate.Index).Value = Me.objCd.DigitGrouping(rowWU.rate, vFrm.TxtDefaultPrecision.Text)
                    .Cells(vFrm.ColWorkUnit_Remarks.Index).Value = rowWU.remarks
                End With
            Next 
            vDgv.AllowUserToAddRows = True
            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function GetProductLevel(ByVal vProductGroupLevelStr As String) As ClsMasterProductionRate.ProductLevel
        Try
            If vProductGroupLevelStr = "Product Group" Then
                Return ClsMasterProductionRate.ProductLevel.ProductGroup
            ElseIf vProductGroupLevelStr = "Brand" Then
                Return ClsMasterProductionRate.ProductLevel.Brand
            ElseIf vProductGroupLevelStr = "Sub Brand" Then
                Return ClsMasterProductionRate.ProductLevel.SubBrand
            ElseIf vProductGroupLevelStr = "Category" Then
                Return ClsMasterProductionRate.ProductLevel.Category
            Else
                Return ClsMasterProductionRate.ProductLevel.Product
            End If

            Return ClsMasterProductionRate.ProductLevel.Product
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ClsMasterProductionRate.ProductLevel.Product
        End Try
    End Function

    Public Function GetProductLevelStr(ByVal vProductGroupLevel As ClsMasterProductionRate.ProductLevel) As String
        Try
            If vProductGroupLevel = ClsMasterProductionRate.ProductLevel.ProductGroup Then
                Return "Product Group"
            ElseIf vProductGroupLevel = ClsMasterProductionRate.ProductLevel.Brand Then
                Return "Brand"
            ElseIf vProductGroupLevel = ClsMasterProductionRate.ProductLevel.SubBrand Then
                Return "Sub Brand"
            ElseIf vProductGroupLevel = ClsMasterProductionRate.ProductLevel.Category Then
                Return "Category"
            Else
                Return "Product"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
End Class
