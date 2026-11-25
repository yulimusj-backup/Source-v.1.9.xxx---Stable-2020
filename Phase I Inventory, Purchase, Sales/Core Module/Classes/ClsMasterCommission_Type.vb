Option Explicit On
Option Strict Off
Public Class ClsMasterCommission_Type
    Dim ObjItem As New ClsMasterCommission_Type_Item

#Region "Property"
    Dim mPrimary As Double = -1
    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Dim mCode As String = ""
    Public Property Code() As String
        Get
            Return mCode
        End Get
        Set(ByVal value As String)
            mCode = value
        End Set
    End Property

    Dim mDesc As String = ""
    Public Property Desc() As String
        Get
            Return mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
        End Set
    End Property

    Dim mBase_Calculation As ClsMasterCommission_Type_Item.Enum_BaseCalculation
    Public Property Base_Calculation() As ClsMasterCommission_Type_Item.Enum_BaseCalculation
        Get
            Return mBase_Calculation
        End Get
        Set(ByVal value As ClsMasterCommission_Type_Item.Enum_BaseCalculation)
            mBase_Calculation = value
        End Set
    End Property

    Dim mLimitType As ClsMasterCommission_Type_Item.Enum_LimitType
    Public Property LimitType() As ClsMasterCommission_Type_Item.Enum_LimitType
        Get
            Return mLimitType
        End Get
        Set(ByVal value As ClsMasterCommission_Type_Item.Enum_LimitType)
            mLimitType = value
        End Set
    End Property

    Dim mProgresiveCalculation As ClsMasterCommission_Type_Item.Enum_ProgresiveCalculation
    Public Property ProgresiveCalculation() As ClsMasterCommission_Type_Item.Enum_ProgresiveCalculation
        Get
            Return mProgresiveCalculation
        End Get
        Set(ByVal value As ClsMasterCommission_Type_Item.Enum_ProgresiveCalculation)
            mProgresiveCalculation = value
        End Set
    End Property

    Dim mRemarks As String = ""
    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Dim mCreatedBy As String = ""
    Public Property CreatedBy() As String
        Get
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Dim mCreatedDate As Date = Now
    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Dim mModifiedBy As String = ""
    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Dim mModifiedDate As Date = Now
    Public Property ModifiedDate() As Date
        Get
            Return mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Dim mTimestamp As Date = Now
    Public Property TimeStamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property
#End Region

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim dta As New DseMasterCommissionTableAdapters.tblmaster_sales_commissionTableAdapter
        Dim retval As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetPrimary(vCode)
            If retval.HasValue = True Then
                If retval.Value <> -1 Then
                    Return retval
                Else
                    Return -1
                End If
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim dta As New DseMasterCommissionTableAdapters.tblmaster_sales_commissionTableAdapter
        Dim retval As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetCode(vPrimary)
            If String.IsNullOrEmpty(retval) = False Then
                Return retval
            Else
                Return ""
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim dta As New DseMasterCommissionTableAdapters.tblmaster_sales_commissionTableAdapter
        Dim retval As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetDesc(vPrimary)
            If String.IsNullOrEmpty(retval) = False Then
                Return retval
            Else
                Return ""
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterCommissionTableAdapters.tblmaster_sales_commissionTableAdapter
        Dim Retval As New Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            If Retval.HasValue = True Then
                Return Retval
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vCode As String) As Double
        Dim DTA As New DseMasterCommissionTableAdapters.tblmaster_sales_commissionTableAdapter
        Dim Retval As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJml(vCode)
            If Retval.HasValue = True Then
                If String.IsNullOrEmpty(Retval) = False Then
                    Return Retval
                Else
                    Return 0
                End If
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahDiAllocate(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim Dta As New DseMasterCommissionTableAdapters.tblmaster_sales_commissionTableAdapter
        Dim Retval As New Nullable(Of Long)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Retval = Dta.ScalarQuery_GetJml_InAllocateComm(vPrimaryTransaksi)
            If Retval.HasValue = True Then
                If Retval.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
            Return False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function SimpanBaru(ByVal vDGVItem As DataGridView, _
         ByVal vUserCode As String, ByVal Frm As FrmMasterCommission) As Boolean
        Dim DTA As New DseMasterCommissionTableAdapters.tblmaster_sales_commissionTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(vDGVItem, True, Frm) = True Then
                If DTA.Insert(mCode, mDesc, mBase_Calculation, mLimitType, mProgresiveCalculation, mRemarks, vUserCode, Now, vUserCode, Now, Now) > 0 Then

                    'update data yang berubah setelah insert query
                    Me.mPrimary = Me.GetPrimary(Me.mCode)
                    Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItem(mPrimary, vDGVItem, Frm)
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

    Public Overridable Function SimpanPerubahan(ByVal vDGVItem As DataGridView, _
   ByVal vUserCode As String, ByVal Frm As FrmMasterCommission) As Boolean
        Dim DTA As New DseMasterCommissionTableAdapters.tblmaster_sales_commissionTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(vDGVItem, False, Frm) = True Then
                If DTA.Update(mCode, mDesc, mBase_Calculation, mLimitType, mProgresiveCalculation, mRemarks, vUserCode, Now, Now, mPrimary, mTimestamp) > 0 Then
                    Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    ObjItem.HapusItem(Me.mPrimary)
                    ObjItem.SimpanItem(Me.mPrimary, Frm.dgvItem, Frm)
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

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, _
   ByVal vUserCode As String, ByVal Frm As FrmMasterCommission, ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterCommissionTableAdapters.tblmaster_sales_commissionTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimaryTransaksi, vUserCode, vCode) = True Then
                If DTA.Delete(vPrimaryTransaksi, Me.mTimestamp) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahDataBenar(ByVal vDGV As DataGridView, ByVal vBaru As Boolean, ByVal vFrm As FrmMasterCommission) As Boolean
        Try
            vDGV.AllowUserToAddRows = False
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.Code) = True Then
                    MsgBox("Code no duplication, please change Stock In No.", _
                    MsgBoxStyle.Information, "Information")
                    vFrm.txtCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.GetCode(Me.Primary)
                If ProductGroupLama <> Me.Code Then
                    If Me.ApakahSudahAda(Me.Code) = True Then
                        MsgBox("Code no duplication, please change Stock In No.", _
                        MsgBoxStyle.Information, "Information")
                        vFrm.txtCode.Focus()
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Try
            If Me.GetJml(vCode) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterCommissionTableAdapters.tblmaster_sales_commissionTableAdapter
        Dim Created_By As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If ApakahSudahDiAllocate(vPrimary) = True Then
                MsgBox("Cannot delete because this data have been exist in commission allocate form.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterCommissionTableAdapters.tblmaster_sales_commissionTableAdapter
        Dim Created_By As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, ByVal Frm As FrmMasterCommission, ByVal ObjCD As Core_Module.ClsCekData) As Boolean
        Dim DTA As New DseMasterCommissionTableAdapters.tblmaster_sales_commissionTableAdapter
        Dim TBL As New DseMasterCommission.tblmaster_sales_commissionDataTable
        Dim row As DseMasterCommission.tblmaster_sales_commissionRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(vPrimary))
            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseMasterCommission.tblmaster_sales_commissionRow)
                With row
                    Me.mPrimary = row.Primary
                    Me.mCode = row.Code
                    Me.mDesc = row.Description
                    If row.Base_Calculation = 0 Then
                        Me.mBase_Calculation = ClsMasterCommission_Type_Item.Enum_BaseCalculation.Fixed_Amount_For_Each_Sold_Item
                    ElseIf row.Base_Calculation = 1 Then
                        Me.mBase_Calculation = ClsMasterCommission_Type_Item.Enum_BaseCalculation.Percentage_of_Sales_Amount
                    End If
                    If row.Limit_Type = 0 Then
                        Me.mLimitType = ClsMasterCommission_Type_Item.Enum_LimitType.By_Qty
                    ElseIf row.Limit_Type = 1 Then
                        Me.mLimitType = ClsMasterCommission_Type_Item.Enum_LimitType.By_Amount
                    End If
                    If row.Progresive_Calculation = 0 Then
                        Me.mProgresiveCalculation = ClsMasterCommission_Type_Item.Enum_ProgresiveCalculation.NoProgresive
                    ElseIf row.Progresive_Calculation = 1 Then
                        Me.mProgresiveCalculation = ClsMasterCommission_Type_Item.Enum_ProgresiveCalculation.Progresive
                    End If
                    Me.mRemarks = row.Remarks
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                Me.ObjItem.IsiDGVItem(vPrimary, vDGVItem, Frm, ObjCD, Frm.vDefCurr, Me.mLimitType, Me.mBase_Calculation)

            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
            ByVal DGV As DataGridView, ByVal TBL As DseMasterCommission.tblmaster_sales_commission1DataTable, _
            ByVal SearchCriteria As Core_Module.EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
            Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
            Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseMasterCommissionTableAdapters.tblmaster_sales_commission1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If SearchCriteria = Core_Module.EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            End If

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Code No."
                DGV.Columns(2).HeaderText = "Description"
                DGV.Columns(3).HeaderText = "Remarks"
                DGV.Columns(1).MinimumWidth = 90
                DGV.Columns(2).MinimumWidth = 70
                DGV.Columns(3).MinimumWidth = 90
            End If
            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > 10000000 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Ticks / TimeSpan.TicksPerSecond) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, 0, MidpointRounding.AwayFromZero)) + _
                    " miliseconds"
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
            TBL.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       Core_Module.EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
                                     Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Double
        Dim DTA As New DseMasterCommissionTableAdapters.tblmaster_sales_commission1TableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = Core_Module.EnumSearchCriteria.ByCode Then
                RetVal = DTA.ScalarQuery_GetJmlByCodeLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByDescription Then
                RetVal = DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByRemarks Then
                RetVal = DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
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
End Class
