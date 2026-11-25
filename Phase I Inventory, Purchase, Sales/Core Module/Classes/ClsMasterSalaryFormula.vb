Option Explicit On
Option Strict Off

Public Class ClsMasterSalaryFormula
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mCode As String = ""
    Dim mDesc As String = ""
    Dim mRemark As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As DateTime = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As DateTime = Now
    Dim mTimestamp As Date = Now

    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Enum EnumSearch
        Code
        Desc
        Remarks
    End Enum

    Public Enum EnumTipe_OT As Integer
        Weekday
        Weekend
        Holiday
    End Enum

    Public Enum EnumTipe_OT_2 As Integer
        OT_1    '1.5x
        OT_2    '2x
        OT_3    '3x
        OT_4    '4x
    End Enum

#Region "Property"
    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            Me.mPrimary = value
        End Set
    End Property

    Public Property Code() As String
        Get
            Return Me.mCode
        End Get
        Set(ByVal value As String)
            Me.mCode = value
        End Set
    End Property

    Public Property Desc() As String
        Get
            Return Me.mDesc
        End Get
        Set(ByVal value As String)
            Me.mDesc = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return Me.mRemark
        End Get
        Set(ByVal value As String)
            Me.mRemark = value
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
#End Region

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim dta As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formulaTableAdapter
        Dim dummy As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetJmlByCode(vCode)
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
            dta.Dispose()
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimarySalaryFormula As Double, ByVal i As FrmMasterSalaryFormula) As Boolean
        Dim ObjPayrollAkses As New ClsMasterPayrollAccess
        Try
            '2. Cek ke payroll akses jika sama sekali tidak memiliki akses, maka tidak boleh edit, dan hapus 
            If Me.GetJmlItemComponent(vPrimarySalaryFormula) <> 0 Or Me.GetJmlItemEmployee(vPrimarySalaryFormula) <> 0 Then
                If ObjPayrollAkses.ApakahMemilikiPayrollAccess(vPrimarySalaryFormula, i.UserCode) = False Then
                    MsgBox("User code : " & i.UserCode.ToString & " doesn't have access.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Erorr")
            Return False
        Finally
            ObjPayrollAkses = Nothing
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String) As Boolean
        Dim DTA1 As New DseMasterSalaryFormulaTableAdapters.tblmaster_payroll_access_item_salaryformulaTableAdapter
        Dim DTA2 As New DseMasterSalaryFormulaTableAdapters.tbltrans_salaryTableAdapter
        Dim DummyLong As Nullable(Of Long)
        Dim ObjPayrollAkses As New ClsMasterPayrollAccess
        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = Core_Module.GetConstring

            DummyLong = DTA2.ScalarQuery_GetJmlByFormula(vPrimary)
            If DummyLong.HasValue = True Then
                If DummyLong.Value > 0 Then
                    MsgBox("You cannot delete this data because master salary formula is referenced by salary calculation.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                Return False
            End If
 

            '2. Cek ke payroll akses jika sama sekali tidak memiliki akses, maka tidak boleh edit, dan hapus  
            If Me.GetJmlItemComponent(vPrimary) <> 0 Or Me.GetJmlItemEmployee(vPrimary) <> 0 Then
                If ObjPayrollAkses.ApakahMemilikiPayrollAccess(vPrimary, vUserCode) = False Then
                    MsgBox("User code : " & vUserCode.ToString & " doesn't have access.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Erorr")
            Return False
        Finally
            DTA1.Dispose()
            DTA2.Dispose()
            ObjPayrollAkses = Nothing
        End Try
    End Function

    Public Function GetCodeByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formulaTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetCodeByPrimary(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetTimeStampByPrimary(ByVal vPrimary As Double) As Date
        Dim dta As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formulaTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetTimestampByPrimary(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return "1/1/1950"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double) As Boolean
        Dim dta As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formulaTableAdapter
        Dim TBL As New DseMasterSalaryFormula.tblmaster_salary_formulaDataTable
        Dim row As DseMasterSalaryFormula.tblmaster_salary_formulaRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .primary
                    Me.mCode = .code
                    Me.mDesc = .desc
                    Me.Remarks = .Remarks
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dta.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Sub SimpanBaru_ParameterGaji(ByVal vPrimary As Double, ByVal vCode As String, ByVal vDesc As String, ByVal vRemarks As String, _
    ByVal vUnits As String, ByVal vTipe As String, ByVal vIsVariable As Integer, ByVal vIsConstraint As Integer)
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_parameterTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Insert(vCode, vDesc, vRemarks, vUnits, vTipe, vIsVariable, vIsConstraint, vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Overridable Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterSalaryFormula_SimpanBaru, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Public Function SimpanBaru1(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formulaTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mCode, Me.mDesc, Me.mRemark, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimaryByCode(Me.mCode)
                Me.mTimestamp = DTA.ScalarQuery_GetTimestampByPrimary(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterSalaryFormula_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vUser As String) As Boolean
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formulaTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Update(Me.mCode, Me.mDesc, Me.mRemark, vUser, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestampByPrimary(Me.Primary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUser

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUser As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterSalaryDescription_Hapus, vUser)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, vUser)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUser As String) As Boolean
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formulaTableAdapter
        Dim DTALog As New DseMasterSalaryFormulaTableAdapters.tbl_log_userTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUser) = True Then
                DTALog.Insert(Core_Module.ClsLogUser.EnumActivityType.Delete, vUser, Now, "Deleting Salary Formula " & _
                Me.GetCodeByPrimary(vPrimary), vPrimary, Me.GetCodeByPrimary(vPrimary))
                DTA.Delete(vPrimary)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTALog.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, ByVal DGV As DataGridView, _
    ByVal SearchCriteria As EnumSearch, ByVal tsl As ToolStripLabel, _
    ByVal vDta As Core_Module.DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formulaTableAdapter, _
    ByVal vtbl As DseMasterSalaryFormula.tblmaster_salary_formulaDataTable) As Boolean
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            vDta.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearch.Code Then
                QueryResult = vDta.FillBy_LikeCode(vtbl, Criteria + "%")
            ElseIf SearchCriteria = EnumSearch.Desc Then
                QueryResult = vDta.FillBy_LikeDesc(vtbl, Criteria + "%")
            ElseIf SearchCriteria = EnumSearch.Remarks Then
                QueryResult = vDta.FillBy_LikeRemarks(vtbl, Criteria + "%")
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
        End Try
    End Function

    Public Function GetDescription(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formulaTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetDescription(vPrimary)
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

    Public Function GetJmlParameter_ByPrimary(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_parameterTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlPrimary(vPrimary)
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
    EnumSearch, ByVal vDta As Core_Module.DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formulaTableAdapter _
    ) As Double
        Dim Dummy As Nullable(Of Long)
        Try
            vDta.Connection.ConnectionString = Core_Module.GetConstring

            If vSearchBy = EnumSearch.Code Then
                Dummy = vDta.ScalarQuery_GetJmlLikeCode(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearch.Desc Then
                Dummy = vDta.ScalarQuery_GetJmlLikeDesc(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearch.Remarks Then
                Dummy = vDta.ScalarQuery_GetJmlLikeRemarks(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetCodePayrollAcces(ByVal vPrimarySalary As Double) As String
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.DataTable1TableAdapter
        Dim Tbl As New DseMasterSalaryFormula.DataTable1DataTable
        Dim row As DseMasterSalaryFormula.DataTable1Row
        Dim str As String = ""
        Dim i As Integer = 1
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(Tbl, vPrimarySalary)
            For Each row In Tbl.Rows
                str = str + i.ToString + ". " + row.Code.ToString + vbNewLine
            Next

            Return str
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            Tbl.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAdaEmployee(ByVal vPrimaryEmp As Double, ByVal vPrimarySalaryFormulaAcuan As Double) As Boolean
        Dim Dta As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_component_detailemployeeTableAdapter
        Dim Jml As Nullable(Of Long)

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Jml = Dta.ScalarQuery_GetJmlByPrimaryKaryawan(vPrimaryEmp, vPrimarySalaryFormulaAcuan)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
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
            Dta.Connection.ConnectionString = Core_Module.GetConstring
        End Try
    End Function

    Public Function GetPrimaryDefaultAccount(ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_component_detailemployeeTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryDefaultAccount(vPrimaryEmployee)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Public Function GetPrimaryByCode(ByVal vSalaryCode As String) As Double
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formulaTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryByCode(vSalaryCode)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_ItemSalaryCalc_ByIdentifierSalaryFormula(ByVal vIdentifier As String) As Double
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formulaTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlItemSalary(vIdentifier)
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

    Public Sub New()
        Me.DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
    End Sub

    Public Function GetJmlConstraint(ByVal vPrimaryParameter As Double, ByVal vTipe As Integer, _
    ByVal vTipe2 As ClsMasterSalaryFormula_item.EnumTipeConstraitVariable, ByVal vRefVal As String) As Double
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponent_constraintTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJml(vTipe, vTipe2, vPrimaryParameter, vRefVal)
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

    Public Function GetJmlItemEmployee(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_component_detailemployeeTableAdapter
        Dim Jml As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Jml = DTA.ScalarQuery_GetJml(vPrimary)

            If Jml.HasValue = True Then
                Return Jml.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlItemComponent(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
        Dim Jml As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Jml = DTA.ScalarQueryGetJmlByPrimaryTransaksi(vPrimary)

            If Jml.HasValue = True Then
                Return Jml.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
