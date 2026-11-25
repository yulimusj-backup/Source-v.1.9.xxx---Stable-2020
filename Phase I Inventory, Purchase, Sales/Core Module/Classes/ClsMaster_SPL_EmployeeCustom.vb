Option Explicit On
Option Strict Off

Public Class ClsMaster_SPL_EmployeeCustom
    Dim mPrimary As Double = -1
    Dim mCode As String = ""
    Dim mDesc As String = ""
    Dim mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

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
            Return Me.mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return Me.mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return Me.mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return Me.mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return Me.mTimestamp
        End Get
        Set(ByVal value As Date)
            Me.mTimestamp = value
        End Set
    End Property

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseMaster_SPL_EmployeeCustomTableAdapters.tbltrans_spl_overtime_order_employee_custom_scalarTableAdapter
        Dim Dummy As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimary(vCode)
            If Dummy.HasValue Then
                Return Dummy.Value
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

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMaster_SPL_EmployeeCustomTableAdapters.tbltrans_spl_overtime_order_employee_custom_scalarTableAdapter
        Dim Dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetCode(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim DTA As New DseMaster_SPL_EmployeeCustomTableAdapters.tbltrans_spl_overtime_order_employee_custom_scalarTableAdapter
        Dim Dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetDesc(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMaster_SPL_EmployeeCustomTableAdapters.tbltrans_spl_overtime_order_employee_custom_scalarTableAdapter
        Dim Dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTimestamp(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return "1/1/1950"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DseMaster_SPL_EmployeeCustomTableAdapters.tbltrans_spl_overtime_order_employee_custom_scalarTableAdapter
        Dim dummy As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            dummy = DTA.ScalarQuery_GetJml(vCode)
            If dummy.HasValue Then
                If dummy.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmMaster_SPL_EmployeeCustom) As Boolean
        Dim Code As String = ""

        Try
            'tidak boleh ada kode yang sama
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mCode) = True Then
                    If vShowMsg = True Then
                        MsgBox("Duplication in code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.GetCode(Me.mPrimary)
                If CodeLama <> Me.mCode Then
                    If Me.ApakahSudahAda(Me.mCode) = True Then
                        If vShowMsg = True Then
                            MsgBox("Duplication in code, please change your code", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double) As Boolean

        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, ByVal DGV As DataGridView, _
    ByVal TBL As DseMaster_SPL_EmployeeCustom.tbltrans_spl_overtime_order_employee_custom_fillDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMaster_SPL_EmployeeCustomTableAdapters.tbltrans_spl_overtime_order_employee_custom_fillTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.Fill(TBL, Criteria + "%", "%", "%", "%", "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.Fill(TBL, "%", Criteria + "%", "%", "%", "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.Fill(TBL, "%", "%", Criteria + "%", "%", "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByEmployeeCode Then
                QueryResult = DTA.Fill(TBL, "%", "%", "%", Criteria + "%", "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByEmployeeName Then
                QueryResult = DTA.Fill(TBL, "%", "%", "%", "%", Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Code"
                DGV.Columns(2).HeaderText = "Description"
                DGV.Columns(3).HeaderText = "Remarks"
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > TimeSpan.TicksPerSecond Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Seconds) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, _
                    0, MidpointRounding.AwayFromZero)) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function CariDetail(ByVal Primary As Double, ByVal vDGVItem As DataGridView, _
    ByVal vObjCD As ClsCekData, ByVal i As FrmMaster_SPL_EmployeeCustom) As Boolean
        Dim DTA As New DseMaster_SPL_EmployeeCustomTableAdapters.tbltrans_spl_overtime_order_employee_customTableAdapter
        Dim TBL As New DseMaster_SPL_EmployeeCustom.tbltrans_spl_overtime_order_employee_customDataTable
        Dim row As DseMaster_SPL_EmployeeCustom.tbltrans_spl_overtime_order_employee_customRow
        Dim ObjItem As New ClsMaster_SPL_EmployeeCustom_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(Primary))

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .primary
                    Me.mCode = .Code
                    Me.mDesc = .Description
                    Me.mRemarks = .Remarks
                   
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .Timestamp
                End With

                'cari data item employee
                ObjItem.IsiDgv(Primary, vDGVItem, i, vObjCD)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As EnumSearchCriteria) As Double
        Dim DTA As New DseMaster_SPL_EmployeeCustomTableAdapters.tbltrans_spl_overtime_order_employee_custom_fillTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If vSearchBy = EnumSearchCriteria.ByCode Then
                RetVal = DTA.ScalarQuery_GetJmlLike("%", "%", "%", "%", vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                RetVal = DTA.ScalarQuery_GetJmlLike("%", "%", "%", vKriteria + "%", "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                RetVal = DTA.ScalarQuery_GetJmlLike("%", "%", vKriteria + "%", "%", "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByEmployeeCode Then
                RetVal = DTA.ScalarQuery_GetJmlLike(vKriteria + "%", "%", "%", "%", "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByEmployeeName Then
                RetVal = DTA.ScalarQuery_GetJmlLike("%", vKriteria + "%", "%", "%", "%")
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

    Public Function SimpanBaru(ByVal vUserCode As String, ByVal vPrimary As Double, ByVal vObjCD As ClsCekData, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmMaster_SPL_EmployeeCustom) As Boolean
        Dim DTA As New DseMaster_SPL_EmployeeCustomTableAdapters.tbltrans_spl_overtime_order_employee_customTableAdapter
        Dim ObjItem As New ClsMaster_SPL_EmployeeCustom_Item

        Try

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True, True, vDGVItem, i) = True Then
                If DTA.Insert(Me.mCode, Me.mDesc, Me.mRemarks, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    Me.mPrimary = Me.GetPrimary(Me.mCode)
                    Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item Employee
                    ObjItem.SimpanItem(Me.mPrimary, vDGVItem, i)

                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vUserCode As String, ByVal vPrimary As Double, ByVal vObjCD As ClsCekData, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmMaster_SPL_EmployeeCustom) As Boolean
        Dim DTA As New DseMaster_SPL_EmployeeCustomTableAdapters.tbltrans_spl_overtime_order_employee_customTableAdapter
        Dim ObjItem As New ClsMaster_SPL_EmployeeCustom_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(False, True, vDGVItem, i) = True Then
                If DTA.Update(Me.mCode, Me.mDesc, Me.mRemarks, vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama Employee
                    ObjItem.HapusItem(Me.mPrimary)

                    'simpan data item Employee
                    ObjItem.SimpanItem(Me.mPrimary, vDGVItem, i)

                    Return True
                Else
                    Return False
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal i As FrmMaster_SPL_EmployeeCustom) As Boolean
        Dim DTA As New DseMaster_SPL_EmployeeCustomTableAdapters.tbltrans_spl_overtime_order_employee_customTableAdapter
        Dim ObjItem As New ClsMaster_SPL_EmployeeCustom_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary) = True Then
                DTA.Delete(vPrimary)

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = -1
            mCode = ""
            mDesc = ""
            mRemarks = ""
          
            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today

            'Item
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
