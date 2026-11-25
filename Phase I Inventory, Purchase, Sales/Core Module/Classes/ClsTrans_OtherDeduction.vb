Option Explicit On
Option Strict Off

Public Class ClsTrans_OtherDeduction
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mCode As String = ""
    Dim mDesc As String = ""
    Dim mPrimarySlip As Double = -1
    Dim mPrimaryRefAcc As Double = -1
    Dim mFormula As String = ""
    Dim mCalPeriod As EnumCalPeriod = EnumCalPeriod.Daily
    Dim mCalOn As ClsMasterSalaryFormula_item.EnumCalculateOn = ClsMasterSalaryFormula_item.EnumCalculateOn.Jan
    Dim mYearlyMaxVal As String = ""
    Dim mMonthlyMaxValue As String = ""
    Dim mRoundingUnit As String = ""
    Dim mRoundingType As EnumRoundingType = EnumRoundingType.No_Rounding
    Dim mYearlyMinVal As String = ""
    Dim mMonthlyMinVal As String = ""
    Dim mYearlyDate As Date = #1/1/1950#
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As DateTime = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As DateTime = Now
    Dim mTimestamp As Date = Now

    Dim ObjItem As New ClsTrans_OtherDeduction_item

    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Enum EnumCalPeriod
        Daily
        Monthly
        Yearly
    End Enum

    Enum EnumRoundingType
        No_Rounding
        Round_Down
        Round_Middlepoint
        Round_Up
    End Enum

    Enum EnumSearch
        code
        Desc
        Slip
        Acc
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

    Public Property Formula() As String
        Get
            Return Me.mFormula
        End Get
        Set(ByVal value As String)
            Me.mFormula = value
        End Set
    End Property

    Public Property PrimarySlip() As Double
        Get
            Return Me.mPrimarySlip
        End Get
        Set(ByVal value As Double)
            Me.mPrimarySlip = value
        End Set
    End Property

    Public Property PrimaryRefAcc() As Double
        Get
            Return Me.mPrimaryRefAcc
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryRefAcc = value
        End Set
    End Property

    Public Property CalPeriod() As EnumCalPeriod
        Get
            Return Me.mCalPeriod
        End Get
        Set(ByVal value As EnumCalPeriod)
            Me.mCalPeriod = value
        End Set
    End Property

    Public Property CalOn() As ClsMasterSalaryFormula_item.EnumCalculateOn
        Get
            Return Me.mCalOn
        End Get
        Set(ByVal value As ClsMasterSalaryFormula_item.EnumCalculateOn)
            Me.mCalOn = value
        End Set
    End Property

    Public Property YearlyMaxVal() As String
        Get
            Return Me.mYearlyMaxVal
        End Get
        Set(ByVal value As String)
            Me.mYearlyMaxVal = value
        End Set
    End Property

    Public Property MonthlyMaxValue() As String
        Get
            Return Me.mMonthlyMaxValue
        End Get
        Set(ByVal value As String)
            Me.mMonthlyMaxValue = value
        End Set
    End Property

    Public Property RoundingUnit() As String
        Get
            Return Me.mRoundingUnit
        End Get
        Set(ByVal value As String)
            Me.mRoundingUnit = value
        End Set
    End Property

    Public Property RoundingType() As EnumRoundingType
        Get
            Return Me.mRoundingType
        End Get
        Set(ByVal value As EnumRoundingType)
            Me.mRoundingType = value
        End Set
    End Property

    Public Property YearlyMinVal() As String
        Get
            Return Me.mYearlyMinVal
        End Get
        Set(ByVal value As String)
            Me.mYearlyMinVal = value
        End Set
    End Property

    Public Property MonthlyMinVal() As String
        Get
            Return Me.mMonthlyMinVal
        End Get
        Set(ByVal value As String)
            Me.mMonthlyMinVal = value
        End Set
    End Property

    Public Property YearlyDate() As Date
        Get
            Return Me.mYearlyDate
        End Get
        Set(ByVal value As Date)
            Me.mYearlyDate = value
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
        Dim dta As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeductionTableAdapter
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

    Public Overridable Function ApakahBolehEdit(ByVal vUser As String, ByVal vPrimary As Double, _
    Optional ByVal vStrictUserEdit As Boolean = True, Optional ByVal vShowMsg As Boolean = True) As Boolean
        Dim dta As New DseTrans_OtherDeductionTableAdapters.tbltrans_salary_item_otherdeductionTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If vStrictUserEdit = True Then
                If vUser <> Me.CreatedBy Then
                    If vShowMsg = True Then
                        MsgBox("Only " + Me.CreatedBy + " user is allowed to edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            Dim jml As Nullable(Of Long)
            jml = dta.ScalarQuery_GetJmlSalaryByOD(vPrimary)
            If jml.HasValue = True Then
                If jml.Value > 0 Then
                    If vShowMsg = True Then
                        MsgBox("These data cannot be changed because this data is refering to calculate salary.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vUser As String, ByVal vPrimary As Double) As Boolean
        Dim dta As New DseTrans_OtherDeductionTableAdapters.tbltrans_salary_item_otherdeductionTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If vUser <> Me.CreatedBy Then
                MsgBox("Only " + Me.CreatedBy + " user is allowed to delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek disalary
            Dim jml As Nullable(Of Long)
            jml = dta.ScalarQuery_GetJmlSalaryByOD(vPrimary)
            If jml.HasValue = True Then
                If jml.Value > 0 Then
                    MsgBox("These data cannot be deleted because this data is refering to calculate salary.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetCodeByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeductionTableAdapter
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

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeductionTableAdapter
        Dim Dummy As System.Nullable(Of Double)
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
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimeStampByPrimary(ByVal vPrimary As Double) As Date
        Dim dta As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeductionTableAdapter
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

    Public Sub CariDetail(ByVal vPrimary As Double, ByVal vTblVar As DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_variableDataTable, _
    ByVal vTblCons As DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_constraintDataTable, ByVal vFrm As FrmMasterSalaryFormula_OtherDeduction)
        Dim Dta As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeductionTableAdapter
        Dim Tbl As New DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeductionDataTable
        Dim row As DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeductionRow
        Dim DtaVar As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeduction_variableTableAdapter
        Dim DtaCon As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeduction_constraintTableAdapter
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            DtaVar.Connection.ConnectionString = Core_Module.GetConstring
            DtaCon.Connection.ConnectionString = Core_Module.GetConstring
            'cari header berdasarkan primary
            Dta.FillBy_Primary(Tbl, vPrimary)
            'cari variabel dan constrain berdasarkan primary transaksi
            DtaVar.FillBy_PrimaryTransaksi(vTblVar, vPrimary)
            DtaCon.FillBy_PrimaryTransaksi(vTblCons, vPrimary)
            If Tbl.Rows.Count > 0 Then
                row = Tbl.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mCode = .Code
                    Me.mDesc = .Description
                    Me.mPrimarySlip = .Primary_Slip_Desc
                    Me.mPrimaryRefAcc = .Primary_Referred_Account
                    Me.mFormula = .Formula
                    Me.mCalPeriod = .Calculation_Period
                    Me.mCalOn = .Calculate_On
                    Me.mYearlyMaxVal = .Yearly_Max_Value
                    Me.mMonthlyMaxValue = .Monthly_Max_Value
                    Me.mRoundingUnit = .Rounding_Unit
                    Me.mRoundingType = .Rounding_Type
                    Me.mYearlyMinVal = .Yearly_Min_Value
                    Me.mMonthlyMinVal = .Monthly_Min_Value
                    Me.mYearlyDate = .Yearly_Date
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp

                    Me.ObjItem.IsiDGVItem(vFrm, vPrimary)
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
            Tbl.Dispose()
            DtaVar.Dispose()
            DtaCon.Dispose()
        End Try
    End Sub

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vFrm As FrmMasterSalaryFormula_OtherDeduction) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_OtherDeduction_SimpanBaru, vUserCode)

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
                            retval = Me.SimpanBaru1(vUserCode, vFrm)
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

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal frm As FrmMasterSalaryFormula_OtherDeduction) As Boolean
        Dim DTA As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeductionTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mCode, Me.mDesc, Me.mPrimarySlip, Me.mPrimaryRefAcc, Me.mFormula, Me.CalPeriod, Me.mCalOn, _
                    Me.mYearlyMaxVal, Me.mMonthlyMaxValue, Me.mRoundingUnit, Me.mRoundingType, Me.mYearlyMinVal, Me.mMonthlyMinVal, Me.mYearlyDate, _
                    vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimaryByCode(Me.mCode)
                Me.mTimestamp = DTA.ScalarQuery_GetTimeStampByPrimary(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan item, simpan table var, simpan table cons
                Me.ObjItem.simpan(frm, Me.mPrimary)
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

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vFrm As FrmMasterSalaryFormula_OtherDeduction) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_OtherDeduction_SimpanPerubahan, vUserCode)

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
                            retval = Me.SimpanPerubahan1(vUserCode, vFrm)
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

    Public Function SimpanPerubahan1(ByVal vUserCode As String, ByVal frm As FrmMasterSalaryFormula_OtherDeduction) As Boolean
        Dim DTA As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeductionTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Update(Me.mCode, Me.mDesc, Me.mPrimarySlip, Me.mPrimaryRefAcc, Me.mFormula, Me.CalPeriod, Me.mCalOn, _
                    Me.mYearlyMaxVal, Me.mMonthlyMaxValue, Me.mRoundingUnit, Me.mRoundingType, Me.mYearlyMinVal, Me.mMonthlyMinVal, Me.mYearlyDate, _
                     vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then

                'hapus item
                Me.ObjItem.hapus(Me.mPrimary)

                'simpan item, simpan table var, simpan table cons
                Me.ObjItem.simpan(frm, Me.mPrimary)
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
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_OtherDeduction_Hapus, vUser)

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
        Dim DTA As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeductionTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vUser, vPrimary) = True Then
                DTA.Delete(vPrimary)
                'pastiin item ke hapus
                Me.ObjItem.hapus(vPrimary)
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

    Public Function IsiDGV(ByVal Criteria As String, ByVal DGV As Windows.Forms.DataGridView, _
   ByVal TBL As DseTrans_OtherDeduction.DataTable1DataTable, ByVal SearchCriteria As ClsTrans_OtherDeduction.EnumSearch, _
   ByVal tsl As Windows.Forms.ToolStripLabel, ByVal dta As DseTrans_OtherDeductionTableAdapters.DataTable1TableAdapter) As Boolean
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearch.code Then
                QueryResult = dta.FillBy_LikeCode(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearch.Desc Then
                QueryResult = dta.FillBy_LikeDesc(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearch.Slip Then
                QueryResult = dta.FillBy_LikeSlip(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearch.Acc Then
                QueryResult = dta.FillBy_LikeAcc(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()

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
            DGV.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.Arrow
        End Try
    End Function

    Public Function GetDescription(ByVal vPrimary As Double) As String
        Dim DTA As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeductionTableAdapter
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ClsTrans_OtherDeduction.EnumSearch) As Double
        Dim Dummy As Nullable(Of Long)
        Dim dta As New DseTrans_OtherDeductionTableAdapters.DataTable1TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearch.code Then
                Dummy = dta.ScalarQuery_GetJmlLikeCode(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearch.Desc Then
                Dummy = dta.ScalarQuery_GetJmlLikeDesc(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearch.Slip Then
                Dummy = dta.ScalarQuery_GetJmlLikeSlip(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearch.Acc Then
                Dummy = dta.ScalarQuery_GetJmlLikeAcc(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function CekCalculate(ByVal vPrimaryEmp As Double, ByVal vPrimaryOther As Double) As Boolean
        Dim Dummy As Nullable(Of Long)
        Dim dta As New DseTrans_OtherDeductionTableAdapters.DataTable3TableAdapter
        Try
            Dummy = dta.ScalarQuery(vPrimaryEmp, vPrimaryOther)
            If Dummy.HasValue Then
                If Dummy.Value > 0 Then
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

    Public Sub New()
        Me.DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
    End Sub
End Class
