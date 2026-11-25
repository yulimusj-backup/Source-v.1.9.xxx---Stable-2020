Option Explicit On
Option Strict Off

Public Class ClsMasterSalaryFormula_SlipDesc
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mCode As String = ""
    Dim mDesc As String = ""
    Dim mRemark As String = ""
    Dim mType As EnumType = EnumType.Deduction
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

    Enum EnumType
        Deduction
        Income
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

    Public Property Type() As EnumType
        Get
            Return Me.mType
        End Get
        Set(ByVal value As EnumType)
            Me.mType = value
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
        Dim dta As New DseMasterSalaryFormula_SlipDescTableAdapters.tblmaster_salary_slip_descriptionTableAdapter
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

    Public Function ApakahBolehEdit(ByVal vPrimary As Double) As Boolean
        'Dim DTA As New DseMasterSalaryFormula_SlipDescTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
        'Dim DummyLong As Nullable(Of Long)
        Try
            'DTA.Connection.ConnectionString = Core_Module.GetConstring

            'DummyLong = DTA.ScalarQuery_GetJmlByPrimarySlipDesc(vPrimary)
            'If DummyLong.HasValue = True Then
            '    If DummyLong.Value > 0 Then
            '        MsgBox("You cannot edit this data because this data is referring to master salary formula component.", _
            '        MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'Else
            '    Return False
            'End If

            'benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Erorr")
        Finally
            'DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterSalaryFormula_SlipDescTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
        Dim DummyLong As Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DummyLong = DTA.ScalarQuery_GetJmlByPrimarySlipDesc(vPrimary)
            If DummyLong.HasValue = True Then
                If DummyLong.Value > 0 Then
                    MsgBox("You cannot delete this data because this data is referring to master salary formula component.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                Return False
            End If

            'benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Erorr")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCodeByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseMasterSalaryFormula_SlipDescTableAdapters.tblmaster_salary_slip_descriptionTableAdapter
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

    Public Function GetDescByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseMasterSalaryFormula_SlipDescTableAdapters.tblmaster_salary_slip_descriptionTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetDescByPrimary(vPrimary)
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

    Public Function GetPrimaryByCode(ByVal vCode As String) As Double
        Dim dta As New DseMasterSalaryFormula_SlipDescTableAdapters.tblmaster_salary_slip_descriptionTableAdapter
        Dim Dummy As Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetPrimaryByCode(vCode)
            If Dummy.HasValue Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetTypeByPrimary(ByVal vPrimary As Double) As Integer
        Dim dta As New DseMasterSalaryFormula_SlipDescTableAdapters.tblmaster_salary_slip_descriptionTableAdapter
        Dim Dummy As Nullable(Of Integer)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetTypeByPrimary(vPrimary)
            If Dummy.HasValue Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetTimeStampByPrimary(ByVal vPrimary As Double) As Date
        Dim dta As New DseMasterSalaryFormula_SlipDescTableAdapters.tblmaster_salary_slip_descriptionTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetTimeStampByPrimary(vPrimary)
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
        Dim dta As New DseMasterSalaryFormula_SlipDescTableAdapters.tblmaster_salary_slip_descriptionTableAdapter
        Dim TBL As New DseMasterSalaryFormula_SlipDesc.tblmaster_salary_slip_descriptionDataTable
        Dim row As DseMasterSalaryFormula_SlipDesc.tblmaster_salary_slip_descriptionRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mCode = .Code
                    Me.mDesc = .Description
                    Me.mType = .Type
                    Me.Remarks = .Remarks
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
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
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterSalaryDescription_SimpanBaru, vUserCode)

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
        Dim DTA As New DseMasterSalaryFormula_SlipDescTableAdapters.tblmaster_salary_slip_descriptionTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mCode, Me.mDesc, Me.mType, Me.mRemark, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimaryByCode(Me.mCode)
                Me.mTimestamp = DTA.ScalarQuery_GetTimeStampByPrimary(Me.mPrimary)
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
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterSalaryDescription_SimpanPerubahan, vUserCode)

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
        Dim DTA As New DseMasterSalaryFormula_SlipDescTableAdapters.tblmaster_salary_slip_descriptionTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehEdit(Me.mPrimary) = True Then
                If DTA.Update(Me.mCode, Me.mDesc, Me.mType, Me.mRemark, vUser, Now, Now, _
                Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStampByPrimary(Me.Primary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUser

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
                            retval = Me.Hapus1(vPrimaryTransaksi)
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

    Public Function Hapus1(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterSalaryFormula_SlipDescTableAdapters.tblmaster_salary_slip_descriptionTableAdapter
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
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, ByVal DGV As DataGridView, _
    ByVal SearchCriteria As EnumSearch, ByVal tsl As ToolStripLabel, _
    ByVal vDta As Core_Module.DseMasterSalaryFormula_SlipDescTableAdapters.tblmaster_salary_slip_descriptionTableAdapter, _
    ByVal vtbl As DseMasterSalaryFormula_SlipDesc.tblmaster_salary_slip_descriptionDataTable) As Boolean
        Dim QueryResult As Double
        Dim StartT As Date = Now
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
    EnumSearch, ByVal vDta As Core_Module.DseMasterSalaryFormula_SlipDescTableAdapters.tblmaster_salary_slip_descriptionTableAdapter _
    ) As Double
        Dim Dummy As Nullable(Of Long)
        Try
            vDta.Connection.ConnectionString = Core_Module.GetConstring

            If vSearchBy = EnumSearch.Code Then
                Dummy = vDta.ScalarQuery_GetJmlByLikeCode(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearch.Desc Then
                Dummy = vDta.ScalarQuery_GetJmlByLikeDesc(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearch.Remarks Then
                Dummy = vDta.ScalarQuery_GetJmlByLikeRemarks(vKriteria + "%")
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

    Public Sub New()
        Me.DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
    End Sub
End Class
