Option Explicit On
Option Strict Off
Public Class ClsMasterWorkUnitGroup
    Dim mPrimary As Double = -1
    Dim mCode As String = ""
    Dim mDesc As String = ""
    Dim mRemarks As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Dim ObjCD As New Core_Module.ClsCekData

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property Code() As String
        Get
            Return mCode
        End Get
        Set(ByVal value As String)
            mCode = value
        End Set
    End Property

    Public Property Desc() As String
        Get
            Return mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
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

    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
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

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseMasterWorkUnitGroupTableAdapters.tblmaster_workunitgroupTableAdapter
        Dim Retval As New Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimary(vCode)
            If Retval.HasValue = True Then
                If Retval.Value > -1 Then
                    Return Retval.Value
                Else
                    Return -1
                End If
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterWorkUnitGroupTableAdapters.tblmaster_workunitgroupTableAdapter
        Dim Retval As New Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            If Retval.HasValue = True Then
                Return Retval.Value
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

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterWorkUnitGroupTableAdapters.tblmaster_workunitgroupTableAdapter
        Dim Retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetCode(vPrimary)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Retval
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterWorkUnitGroupTableAdapters.tblmaster_workunitgroupTableAdapter
        Dim Retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetDesc(vPrimary)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Retval
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanBaru(ByVal Frm As FrmMasterWorkUnitGroup, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterWorkUnitGroupTableAdapters.tblmaster_workunitgroupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True, Frm) = True Then

                If DTA.Insert(Me.mCode, Me.mDesc, Me.mRemarks, _
                 vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    Me.mPrimary = Me.GetPrimary(Me.mCode)
                    Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal Frm As FrmMasterWorkUnitGroup) As Boolean
        Dim DTA As New DseMasterWorkUnitGroupTableAdapters.tblmaster_workunitgroupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(False, Frm) = True Then

                If DTA.Update(Me.mCode, Me.mDesc, Me.mRemarks, _
                 vUserCode, Now, Now, Me.mPrimary) > 0 Then
                    Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As String, ByVal Frm As FrmMasterWorkUnitGroup, ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterWorkUnitGroupTableAdapters.tblmaster_workunitgroupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Delete(vPrimaryTransaksi) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vFrm As FrmMasterWorkUnitGroup) As Boolean
        Try
            'tidak boleh double
            If vBaru = True Then
                If Me.Apakahsudahada(mCode) = True Then
                    MsgBox("Duplication code, please change your code", _
                   MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                Dim productGroupLama As String = Me.GetCode(mPrimary)
                If productGroupLama <> mCode Then
                    If Me.Apakahsudahada(mCode) = True Then
                        MsgBox("Duplication code, please change your code", _
                         MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If

                'data salah jika nilai timestamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim oldTimeStamp As Date = Me.GetTimestamp(mPrimary)
                If mTimestamp <> oldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.CariDetail(Me.mPrimary, vFrm, Me.ObjCD)
                    End If
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal Frm As FrmMasterWorkUnitGroup, _
                                ByVal ObjCD As Core_Module.ClsCekData) As Boolean
        Dim DTA As New DseMasterWorkUnitGroupTableAdapters.tblmaster_workunitgroupTableAdapter
        Dim TBL As New DseMasterWorkUnitGroup.tblmaster_workunitgroupDataTable
        Dim row As DseMasterWorkUnitGroup.tblmaster_workunitgroupRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(vPrimary))
            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseMasterWorkUnitGroup.tblmaster_workunitgroupRow)
                With row
                    Me.mPrimary = row.primary
                    Me.mCode = row.code
                    Me.mDesc = row.desc
                    Me.mRemarks = row.remarks
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
         ByVal DGV As DataGridView, ByVal TBL As DseMasterWorkUnitGroup.tblmaster_workunitgroup1DataTable, _
         ByVal SearchCriteria As Core_Module.EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean

        Dim DTA As New DseMasterWorkUnitGroupTableAdapters.tblmaster_workunitgroup1TableAdapter
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
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Code"
                DGV.Columns(2).HeaderText = "Description"
                DGV.Columns(3).HeaderText = "Remarks"

                DGV.Columns(1).MinimumWidth = 50
                DGV.Columns(2).MinimumWidth = 80
                DGV.Columns(3).MinimumWidth = 50

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
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
            TBL.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal vCode As String) As Double
        Dim DTA As New DseMasterWorkUnitGroupTableAdapters.tblmaster_workunitgroupTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByCode(vCode)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       Core_Module.EnumSearchCriteria) As Double
        Dim DTA As New DseMasterWorkUnitGroupTableAdapters.tblmaster_workunitgroup1TableAdapter
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
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Apakahsudahada(ByVal vCode As String) As Boolean
        Dim Retval As Nullable(Of Long)
        Dim DTA As New DseMasterWorkUnitGroupTableAdapters.tblmaster_workunitgroupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJml(vCode)
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
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vCode As String, Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Try
            If ApakahSudahAdaDI_WU(vCode) = True Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function ApakahSudahAdaDI_WU(ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterWorkUnitGroupTableAdapters.tblmaster_workunitgroupTableAdapter
        Dim Retval As Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJmlIn_WU(vCode)
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
            Return True
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
