Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class ClsTrans_SPL_Budget
    Dim mPrimary As Double = -1
    Dim mNoTransaksi As String = ""
    Dim mUrutan As Integer = -1
    Dim mDate As Date = #1/1/1950#
    Dim mPeriod As Date = #1/1/1950#
    Dim mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As DateTime = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As DateTime = Now
    Dim mTimestamp As Date = Now

    Dim ObjItem As New ClsTrans_SPL_Budget_Item
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjUser As New Core_Module.ClsUser

    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Enum EnumSearch As Integer
        DocNo
        Tgl
        Remarks
    End Enum

    Enum EnumCariData As Integer
        AllDept
        Other
    End Enum

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property NoTransaksi() As String
        Get
            Return mNoTransaksi
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property

    Public Property Urutan() As Integer
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Integer)
            mUrutan = value
        End Set
    End Property

    Public Property Date_() As Date
        Get
            Return mDate
        End Get
        Set(ByVal value As Date)
            mDate = value
        End Set
    End Property

    Public Property Period() As Date
        Get
            Return mPeriod
        End Get
        Set(ByVal value As Date)
            mPeriod = value
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

    Public Function ApakahBolehedit(ByVal vUserCode As String, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim Dta As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_scalarTableAdapter

        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_SPL_Budget_Edit") = False Then
                MsgBox("You don't have permission to perform this action", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '1.strict edit
            Dim UserCode As String = Dta.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If UserCode <> vUserCode Then
                If Me.ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_SPL_Budget_EditUserLain") = False Then
                    MsgBox("Only user with ID:" + UserCode + " can edit this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            ''2.status aktif tidak boleh di edit
            'Dim status As Integer = Dta.ScalarQuery_GetStatusByPrimary(vPrimaryTransaksi)
            'If status = ClsTransaksiLoan.EnumStatusLoan.Finished Then
            '    MsgBox("This document no : " + Me.NoTransaksi.ToString + " cannot edit because loan status is finished.", _
            '    MsgBoxStyle.Information, "Information")
            'End If

            '4.tidak boleh jika sudah ada di transaksi loan payment

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vUserCode As String, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim Dta As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_scalarTableAdapter

        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_SPL_Budget_Hapus") = False Then
                MsgBox("You don't have permission to perform this action", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '1.strict edit
            Dim UserCode As String = Dta.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If UserCode <> vUserCode Then
                If Me.ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_SPL_Budget_HapusUserLain") = False Then
                    MsgBox("Only user with ID:" + UserCode + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            ''2.status aktif tidak boleh di edit
            'Dim status As Integer = Dta.ScalarQuery_GetStatusByPrimary(vPrimaryTransaksi)
            'If status = ClsTransaksiLoan.EnumStatusLoan.Finished Then
            '    MsgBox("This document no : " + Me.NoTransaksi.ToString + " cannot delete because loan status is finished.", _
            '    MsgBoxStyle.Information, "Information")
            'End If

            ''3.Jika sudah terjadi closing tidak boleh di edit
            'If vBeginingBalance = False Then
            '    Dim closing As Date = ObjClosing.GetClosingDate
            '    If Me.Date_ < closing Then
            '        MsgBox("You cannot delete this data because this data is closed in period : " + Format(closing, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If

            ''4. tidak boleh hapus jika sudah ada payment yang merujuk
            'If Me.GetJmlLoanPayment(vPrimaryTransaksi) > 0 Then
            '    MsgBox("You cannot delete this data, there are loan payment refering to this data", MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Private Function ApakahDataBenar(ByVal Frm As FrmTrans_SPL_Budget, ByVal vBaru As Boolean, _
    ByVal vDGV As DataGridView) As Boolean
        Dim i As New FrmTrans_SPL_Budget

        Try
            '1.tidak boleh double
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.NoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mDate, Me.mPrimary)
                End If
            Else
                Dim NoTransaksiLama As String = Me.GetNoTransaksiByPrimary(Me.mPrimary)
                If NoTransaksiLama <> Me.mNoTransaksi Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mDate, Me.mPrimary)
                    End If
                End If
            End If

            'dalam satu periode hanya bisa input satu department saja.
            For Each row As DataGridViewRow In vDGV.Rows
                If row.Cells(i.ColDeptPrimary.Index).Value <> -1 And String.IsNullOrEmpty(row.Cells(i.ColDeptPrimary.Index).Value) = False Then
                    If Me.ApakahDepartmentSudahDigunakan(Me.mPrimary, row.Cells(i.ColDeptPrimary.Index).Value, Me.mPeriod) = True Then
                        MsgBox("Department (" & row.Cells(i.ColDeptCode.Index).Value & " - " & row.Cells(i.ColDeptDesc.Index).Value & ") has been used in period", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Function

    Public Function ApakahDepartmentSudahDigunakan(ByVal vPrimaryTransaksi As Double, ByVal vPrimaryDept As Double, _
    ByVal vPeriod As Date) As Boolean
        Dim Dta As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_scalarTableAdapter
        Dim JmlDept As Nullable(Of Long)

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            JmlDept = Dta.ScalarQuery_GetJmlDeptInPeriod(vPrimaryTransaksi, vPeriod, vPrimaryDept)
            If JmlDept.HasValue = True Then
                If JmlDept.Value > 0 Then
                    Return True
                End If
            Else
                Return False
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetTanggalTransaksi(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_scalarTableAdapter
        Dim Dummy As System.Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTransDate(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vPrimaryTrans As Double) As String
        Dim DTA As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_scalarTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim JmlDb As Nullable(Of Double)
        Try
            'cari tanggal sampai dan dari
            TanggalEnd = Me.ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = Me.ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            JmlDb = DTA.ScalarQuery_GetMaxUrutan(TanggalStart, TanggalEnd, vPrimaryTrans)
            If JmlDb.HasValue = True Then
                NoUrut = CStr(CInt(JmlDb.Value)) + 1
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            'pastikan no urut terdiri atas 6 huruf
            While NoUrut.Length < 6
                NoUrut = "0" + NoUrut
            End While

            'bulan 
            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While
            Return NoUrut + "/" + "OTB" + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vFrm As FrmTrans_SPL_Budget, _
    ByVal vDgvItem As Windows.Forms.DataGridView) As Boolean
        Dim DTA As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budgetTableAdapter
        Dim TBL As New DseTrans_SPL_Budget.tbltrans_spl_budgetDataTable
        Dim row As DseTrans_SPL_Budget.tbltrans_spl_budgetRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mUrutan = .Urutan
                    Me.mDate = ._Date
                    Me.mPeriod = .Period
                    Me.mRemarks = .Remarks

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp

                    Me.ObjItem.IsiDGVItem(vDgvItem, .Primary, vFrm)
                End With
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim dta As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_scalarTableAdapter
        Dim dummy As Nullable(Of Long)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetJml(vNoTransaksi)
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

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim d As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_scalarTableAdapter
        Dim n As New Nullable(Of Double)

        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            n = d.ScalarQuery_GetPrimary(vNoTransaksi)
            If n.HasValue Then
                Return n.Value
            End If

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            d.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksiByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_scalarTableAdapter
        Dim Dummy As String

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetNoTransaksi(vPrimary)
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
        Dim dta As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_scalarTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetTimestamp(vPrimary)
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

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vDGV As DataGridView, _
   ByVal frm As FrmTrans_SPL_Budget) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_SimpanBaru, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > Me.ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If CInt(Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value) = CInt(Me.vPrimary_Antrian) Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode, vDGV, frm)
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

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal vDGV As DataGridView, ByVal frm As FrmTrans_SPL_Budget) As Boolean
        Dim DTA As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budgetTableAdapter
        Dim BufferDoc As String = Me.mNoTransaksi

        Try
            If Me.ApakahDataBenar(frm, True, vDGV) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring

                Me.mUrutan = CDbl(Me.mNoTransaksi.Substring(0, 6))
                If DTA.Insert(Me.mNoTransaksi, Me.mUrutan, Me.mDate, Me.mPeriod, Me.mRemarks, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    Me.mPrimary = Me.GetPrimary(Me.mNoTransaksi)
                    Me.mTimestamp = Me.GetTimeStampByPrimary(Me.mPrimary)
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan item
                    Me.ObjItem.SimpanItem(Me.Primary, vDGV, frm)

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

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vDGV As DataGridView, _
   ByVal frm As FrmTrans_SPL_Budget) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > Me.ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If CInt(Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value) = CInt(Me.vPrimary_Antrian) Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode, vDGV, frm)
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

    Public Function SimpanPerubahan1(ByVal vUserCode As String, ByVal vDGV As DataGridView, ByVal frm As FrmTrans_SPL_Budget) As Boolean
        Dim DTA As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budgetTableAdapter
        Dim BufferDoc As String = Me.mNoTransaksi
        Try
            If Me.ApakahDataBenar(frm, False, vDGV) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring

                Me.mUrutan = CDbl(Me.mNoTransaksi.Substring(0, 6))
                If DTA.Update(Me.mNoTransaksi, Me.mUrutan, Me.mDate, Me.mPeriod, Me.mRemarks, vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = Me.GetTimeStampByPrimary(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus item
                    Me.ObjItem.HapusItem(Me.mPrimary)

                    'simpan item
                    Me.ObjItem.SimpanItem(Me.Primary, vDGV, frm)

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

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, ByVal frm As FrmTrans_SPL_Budget) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_Hapus, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > Me.ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If CInt(Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value) = CInt(Me.vPrimary_Antrian) Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, vUserCode, frm)
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

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal frm As FrmTrans_SPL_Budget) As Boolean
        Dim DTA As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budgetTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahBolehHapus(vUserCode, vPrimary) = True Then
                DTA.Delete(vPrimary)
                Me.ObjItem.HapusItem(vPrimary)

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

    Public Function IsiDGV(ByVal Criteria As String, ByVal vDGV As DataGridView, _
    ByVal vTBL As DseTrans_SPL_Budget.tbltrans_spl_budget_fillDataTable, ByVal SearchCriteria As EnumSearch, _
    ByVal tsl As ToolStripLabel, ByVal vTanggalDari As Date, _
    ByVal vTanggalSampai As Date, ByVal vFrm As FrmTrans_SPL_Budget, Optional ByVal vJml As Integer = 0) As Boolean
        Dim DTA As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_fillTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearch.DocNo Then
                QueryResult = DTA.Fill_like(vTBL, Criteria + "%", "%", #1/1/1950#, #1/1/2200#)
            ElseIf SearchCriteria = EnumSearch.Tgl Then
                QueryResult = DTA.Fill_like(vTBL, "%", "%", vTanggalDari, vTanggalSampai)
            ElseIf SearchCriteria = EnumSearch.Remarks Then
                QueryResult = DTA.Fill_like(vTBL, "%", Criteria + "%", #1/1/1950#, #1/1/2200#)
            End If

            'update dgv
            vDGV.DataSource = vTBL.DefaultView
            vDGV.Refresh()
            If vDGV.ColumnCount > 0 Then
                vDGV.Columns(0).Visible = False
                vDGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                vDGV.Columns(3).DefaultCellStyle.Format = "yyyy"
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
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vJml = QueryResult

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal vCriteria As String, ByVal vDGV As DataGridView, ByVal vSearchCriteria As EnumSearch, _
    ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date) As Double
        Dim Dta As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_fillTableAdapter
        Dim Dummy As Nullable(Of Long)
        Try

            Dta.Connection.ConnectionString = Core_Module.GetConstring

            If vSearchCriteria = EnumSearch.DocNo Then
                Dummy = Dta.ScalarQuery_GetJmlLike(vCriteria + "%", #1/1/1950#, #1/1/2200#, "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = EnumSearch.Tgl Then
                Dummy = Dta.ScalarQuery_GetJmlLike("%", vTanggalDari, vTanggalSampai, "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = EnumSearch.Remarks Then
                Dummy = Dta.ScalarQuery_GetJmlLike("%", #1/1/1950#, #1/1/2200#, vCriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function Cetak(ByVal vPrimary As Double, ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer) As Boolean
        Dim report As New CrpTrans_SPL_Budget
        Dim ds As New DseTrans_SPL_BudgetCetak
        Dim dta As New DseTrans_SPL_BudgetCetakTableAdapters.DataTable1TableAdapter
        Dim tbl, tblDum As New DseTrans_SPL_BudgetCetak.DataTable1DataTable
        Dim row As DseTrans_SPL_BudgetCetak.DataTable1Row
        Dim dta2 As New DseTrans_SPL_BudgetCetakTableAdapters.tbl_total_ot_attendanceTableAdapter
        Dim tbl2 As New DseTrans_SPL_BudgetCetak.tbl_total_ot_attendanceDataTable
        Dim row2 As DseTrans_SPL_BudgetCetak.tbl_total_ot_attendanceRow
        Dim dta5 As New DseTrans_SPL_BudgetTableAdapters.tblmaster_departemenTableAdapter
        Dim tbl5 As New DseTrans_SPL_Budget.tblmaster_departemenDataTable
        Dim row5 As DseTrans_SPL_Budget.tblmaster_departemenRow
        Dim tblC As DseTrans_SPL_BudgetCetak.tbl_cetakDataTable = ds.tbl_cetak
        Dim rowC As DseTrans_SPL_BudgetCetak.tbl_cetakRow

        Dim ParNoTransaksi As String = ""
        Dim ParTglTransaksi As Date = Now
        Dim ParPeriod As String = ""
        Dim ParRemarks As String = ""
        Dim DateFirst As Date = Now
        Dim DateLast As Date = Now
        Dim Pemberitahuan As String = ""
        Dim hasil As MsgBoxResult

        Dim ActualOT_Jan, ActualOT_Feb, ActualOT_Mar, ActualOT_Apr As Double
        Dim ActualOT_May, ActualOT_Jun, ActualOT_Jul, ActualOT_Aug As Double
        Dim ActualOT_Sep, ActualOT_Oct, ActualOT_Nov, ActualOT_Dec As Double
        Dim GT_OT_Jan, GT_OT_Feb, GT_OT_Mar, GT_OT_Apr As Double
        Dim GT_OT_May, GT_OT_Jun, GT_OT_Jul, GT_OT_Aug As Double
        Dim GT_OT_Sep, GT_OT_Oct, GT_OT_Nov, GT_OT_Dec, GT_OT_Total As Double
        Dim GT_Act_OT_Jan, GT_Act_OT_Feb, GT_Act_OT_Mar, GT_Act_OT_Apr As Double
        Dim GT_Act_OT_May, GT_Act_OT_Jun, GT_Act_OT_Jul, GT_Act_OT_Aug As Double
        Dim GT_Act_OT_Sep, GT_Act_OT_Oct, GT_Act_OT_Nov, GT_Act_OT_Dec, GT_Act_OT_Total As Double
        Dim TotalBudgetOT As Double = 0
        Dim TotalActualOT As Double = 0
        Dim MaxUrutan As Double = 0

        Dim RedColorOverBudget As Boolean = False

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta2.Connection.ConnectionString = Core_Module.GetConstring
            dta5.Connection.ConnectionString = Core_Module.GetConstring

            If dta.Fill(tbl, vPrimary) > 0 Then
                'Header Page Open
                For i As Integer = 0 To tbl.Rows.Count - 1
                    row = tbl.Rows(i)
                    If i = 0 Then
                        ParNoTransaksi = row.No_Transaksi
                        ParTglTransaksi = row._Date
                        ParPeriod = row.Period.Year
                        ParRemarks = row.Remarks

                        DateFirst = Me.ObjCD.GetFisrtDateOfTheYear(row.Period)
                        DateLast = Me.ObjCD.GetLastDateOfTheYear(row.Period)
                        dta.Fill_2(tblDum, vPrimary, row.Period.Date)

                        dta5.Fill(tbl5)
                    End If

                    tbl5.DefaultView.RowFilter = "[Primary] = " & row.Primary_Department
                    While tbl5.DefaultView.Count > 0
                        row5 = tbl5.DefaultView(0).Row
                        tbl5.Rows.Remove(row5)
                    End While
                Next
                'Header Page Close

                'Message Open
                For i As Integer = 0 To tblDum.Rows.Count - 1
                    row = tblDum.Rows(i)
                    tbl5.DefaultView.RowFilter = "[Primary] = " & row.Primary_Department
                    While tbl5.DefaultView.Count > 0
                        row5 = tbl5.DefaultView(0).Row
                        tbl5.Rows.Remove(row5)
                    End While
                Next

                If tbl5.Rows.Count > 0 Then
                    For Each row5 In tbl5
                        If Pemberitahuan = "" Then
                            Pemberitahuan = row5.Description
                        Else
                            Pemberitahuan &= ", " & row5.Description
                        End If
                    Next

                    hasil = MsgBox("The following departments that have not been registered in the " & ParPeriod & " period" & vbCrLf & _
                    Pemberitahuan & vbCrLf & "Continue printing ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.No Then
                        Return False
                    End If
                End If
                'Message Close

                Dim ij As New FrmTrans_SPL_Budget_Opt
                ij.ShowDialog()
                If ij.ChkORedColor.Checked = True Then
                    RedColorOverBudget = True
                End If

                'Get OT Attendance Open
                dta2.ClearBeforeFill = False
                For i As Integer = 1 To 12
                    DateFirst = CDate(i & "/1/" & ParPeriod)
                    DateLast = Me.ObjCD.GetLasDateOfTheMonth(DateFirst)
                    dta2.Fill(tbl2, DateFirst, DateLast)
                    For Each row2 In tbl2
                        If row2.Period = CDate("1/1/1950") Then
                            row2.Period = DateFirst
                        End If
                    Next
                Next
                'Get OT Attendance Close

                'Isi Tabel Cetak Open
                For i As Integer = 0 To tbl.Rows.Count - 1
                    row = tbl.Rows(i)

                    rowC = tblC.Newtbl_cetakRow
                    With rowC
                        .Primary_Department = row.Primary_Department
                        .Dept_Code = row.Dept_Code
                        .Dept_Desc = row.Dept_Desc

                        'Isi Budget
                        If row.Jan > 0 Then
                            .Jan = Me.ObjCD.Convert_ToString(row.Jan)
                        End If
                        If row.Feb > 0 Then
                            .Feb = Me.ObjCD.Convert_ToString(row.Feb)
                        End If
                        If row.Mar > 0 Then
                            .Mar = Me.ObjCD.Convert_ToString(row.Mar)
                        End If
                        If row.Apr > 0 Then
                            .Apr = Me.ObjCD.Convert_ToString(row.Apr)
                        End If
                        If row.May > 0 Then
                            .May = Me.ObjCD.Convert_ToString(row.May)
                        End If
                        If row.Jun > 0 Then
                            .Jun = Me.ObjCD.Convert_ToString(row.Jun)
                        End If
                        If row.Jul > 0 Then
                            .Jul = Me.ObjCD.Convert_ToString(row.Jul)
                        End If
                        If row.Aug > 0 Then
                            .Aug = Me.ObjCD.Convert_ToString(row.Aug)
                        End If
                        If row.Sep > 0 Then
                            .Sep = Me.ObjCD.Convert_ToString(row.Sep)
                        End If
                        If row.Oct > 0 Then
                            .Oct = Me.ObjCD.Convert_ToString(row.Oct)
                        End If
                        If row.Nov > 0 Then
                            .Nov = Me.ObjCD.Convert_ToString(row.Nov)
                        End If
                        If row.Dec > 0 Then
                            .Dec = Me.ObjCD.Convert_ToString(row.Dec)
                        End If

                        'Isi Actual Attendance
                        ActualOT_Jan = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("1/1/" & ParPeriod), tbl2)
                        If ActualOT_Jan > 0 Then
                            .Jan_Act = Me.ObjCD.Convert_ToString(ActualOT_Jan)
                        End If
                        ActualOT_Feb = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("2/1/" & ParPeriod), tbl2)
                        If ActualOT_Feb > 0 Then
                            .Feb_Act = Me.ObjCD.Convert_ToString(ActualOT_Feb)
                        End If
                        ActualOT_Mar = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("3/1/" & ParPeriod), tbl2)
                        If ActualOT_Mar > 0 Then
                            .Mar_Act = Me.ObjCD.Convert_ToString(ActualOT_Mar)
                        End If
                        ActualOT_Apr = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("4/1/" & ParPeriod), tbl2)
                        If ActualOT_Apr > 0 Then
                            .Apr_Act = Me.ObjCD.Convert_ToString(ActualOT_Apr)
                        End If
                        ActualOT_May = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("5/1/" & ParPeriod), tbl2)
                        If ActualOT_May > 0 Then
                            .May_Act = Me.ObjCD.Convert_ToString(ActualOT_May)
                        End If
                        ActualOT_Jun = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("6/1/" & ParPeriod), tbl2)
                        If ActualOT_Jun > 0 Then
                            .Jun_Act = Me.ObjCD.Convert_ToString(ActualOT_Jun)
                        End If
                        ActualOT_Jul = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("7/1/" & ParPeriod), tbl2)
                        If ActualOT_Jul > 0 Then
                            .Jul_Act = Me.ObjCD.Convert_ToString(ActualOT_Jul)
                        End If
                        ActualOT_Aug = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("8/1/" & ParPeriod), tbl2)
                        If ActualOT_Aug > 0 Then
                            .Aug_Act = Me.ObjCD.Convert_ToString(ActualOT_Aug)
                        End If
                        ActualOT_Sep = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("9/1/" & ParPeriod), tbl2)
                        If ActualOT_Sep > 0 Then
                            .Sep_Act = Me.ObjCD.Convert_ToString(ActualOT_Sep)
                        End If
                        ActualOT_Oct = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("10/1/" & ParPeriod), tbl2)
                        If ActualOT_Oct > 0 Then
                            .Oct_Act = Me.ObjCD.Convert_ToString(ActualOT_Oct)
                        End If
                        ActualOT_Nov = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("11/1/" & ParPeriod), tbl2)
                        If ActualOT_Nov > 0 Then
                            .Nov_Act = Me.ObjCD.Convert_ToString(ActualOT_Nov)
                        End If
                        ActualOT_Dec = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("12/1/" & ParPeriod), tbl2)
                        If ActualOT_Dec > 0 Then
                            .Dec_Act = Me.ObjCD.Convert_ToString(ActualOT_Dec)
                        End If

                        'HITUNG TOTAL BUDGET OT
                        TotalBudgetOT = row.Jan + row.Feb + row.Mar + row.Apr + row.May + row.Jun + _
                                        row.Jul + row.Aug + row.Sep + row.Oct + row.Nov + row.Dec
                        If TotalBudgetOT > 0 Then
                            .Total = Me.ObjCD.Convert_ToString(TotalBudgetOT)
                        End If
                        'HITUNG TOTAL ACTUAL OT
                        TotalActualOT = ActualOT_Jan + ActualOT_Feb + ActualOT_Mar + ActualOT_Apr + _
                                        ActualOT_May + ActualOT_Jun + ActualOT_Jul + ActualOT_Aug + _
                                        ActualOT_Sep + ActualOT_Oct + ActualOT_Nov + ActualOT_Dec
                        If TotalActualOT > 0 Then
                            .Total_Act = Me.ObjCD.Convert_ToString(TotalActualOT)
                        End If

                        'Tentukan Apakah Text Color Red Jika Over Budget
                        If row.Jan < ActualOT_Jan Then
                            .Is_Red_Jan = 1
                        End If
                        If row.Feb < ActualOT_Feb Then
                            .Is_Red_Feb = 1
                        End If
                        If row.Mar < ActualOT_Mar Then
                            .Is_Red_Mar = 1
                        End If
                        If row.Apr < ActualOT_Apr Then
                            .Is_Red_Apr = 1
                        End If
                        If row.May < ActualOT_May Then
                            .Is_Red_May = 1
                        End If
                        If row.Jun < ActualOT_Jun Then
                            .Is_Red_Jun = 1
                        End If
                        If row.Jul < ActualOT_Jul Then
                            .Is_Red_Jul = 1
                        End If
                        If row.Aug < ActualOT_Aug Then
                            .Is_Red_Aug = 1
                        End If
                        If row.Sep < ActualOT_Sep Then
                            .Is_Red_Sep = 1
                        End If
                        If row.Oct < ActualOT_Oct Then
                            .Is_Red_Oct = 1
                        End If
                        If row.Nov < ActualOT_Nov Then
                            .Is_Red_Nov = 1
                        End If
                        If row.Dec < ActualOT_Dec Then
                            .Is_Red_Dec = 1
                        End If
                        If TotalBudgetOT < TotalActualOT Then
                            .Is_Red_Total = 1
                        End If
                    End With

                    MaxUrutan += 1
                    rowC.Urutan = MaxUrutan
                    tblC.Rows.Add(rowC)

                    With row
                        GT_OT_Jan += .Jan
                        GT_OT_Feb += .Feb
                        GT_OT_Mar += .Mar
                        GT_OT_Apr += .Apr
                        GT_OT_May += .May
                        GT_OT_Jun += .Jun
                        GT_OT_Jul += .Jul
                        GT_OT_Aug += .Aug
                        GT_OT_Sep += .Sep
                        GT_OT_Oct += .Oct
                        GT_OT_Nov += .Nov
                        GT_OT_Dec += .Dec
                        GT_OT_Total += TotalBudgetOT

                        GT_Act_OT_Jan += ActualOT_Jan
                        GT_Act_OT_Feb += ActualOT_Feb
                        GT_Act_OT_Mar += ActualOT_Mar
                        GT_Act_OT_Apr += ActualOT_Apr
                        GT_Act_OT_May += ActualOT_May
                        GT_Act_OT_Jun += ActualOT_Jun
                        GT_Act_OT_Jul += ActualOT_Jul
                        GT_Act_OT_Aug += ActualOT_Aug
                        GT_Act_OT_Sep += ActualOT_Sep
                        GT_Act_OT_Oct += ActualOT_Oct
                        GT_Act_OT_Nov += ActualOT_Nov
                        GT_Act_OT_Dec += ActualOT_Dec
                        GT_Act_OT_Total += TotalActualOT
                    End With
                Next
                'Isi Tabel Cetak Close

                'Isi Tabel Grand Total Open
                rowC = tblC.Newtbl_cetakRow
                With rowC
                    .Primary_Department = -1
                    .Dept_Code = "Grand Total"
                    .Dept_Desc = "Grand Total"

                    If GT_OT_Jan > 0 Then
                        .Jan = Me.ObjCD.Convert_ToString(GT_OT_Jan)
                    End If
                    If GT_OT_Feb > 0 Then
                        .Feb = Me.ObjCD.Convert_ToString(GT_OT_Feb)
                    End If
                    If GT_OT_Mar > 0 Then
                        .Mar = Me.ObjCD.Convert_ToString(GT_OT_Mar)
                    End If
                    If GT_OT_Apr > 0 Then
                        .Apr = Me.ObjCD.Convert_ToString(GT_OT_Apr)
                    End If
                    If GT_OT_May > 0 Then
                        .May = Me.ObjCD.Convert_ToString(GT_OT_May)
                    End If
                    If GT_OT_Jun > 0 Then
                        .Jun = Me.ObjCD.Convert_ToString(GT_OT_Jun)
                    End If
                    If GT_OT_Jul > 0 Then
                        .Jul = Me.ObjCD.Convert_ToString(GT_OT_Jul)
                    End If
                    If GT_OT_Aug > 0 Then
                        .Aug = Me.ObjCD.Convert_ToString(GT_OT_Aug)
                    End If
                    If GT_OT_Sep > 0 Then
                        .Sep = Me.ObjCD.Convert_ToString(GT_OT_Sep)
                    End If
                    If GT_OT_Oct > 0 Then
                        .Oct = Me.ObjCD.Convert_ToString(GT_OT_Oct)
                    End If
                    If GT_OT_Nov > 0 Then
                        .Nov = Me.ObjCD.Convert_ToString(GT_OT_Nov)
                    End If
                    If GT_OT_Dec > 0 Then
                        .Dec = Me.ObjCD.Convert_ToString(GT_OT_Dec)
                    End If
                    If GT_OT_Total > 0 Then
                        .Total = Me.ObjCD.Convert_ToString(GT_OT_Total)
                    End If

                    If GT_Act_OT_Jan > 0 Then
                        .Jan_Act = Me.ObjCD.Convert_ToString(GT_Act_OT_Jan)
                    End If
                    If GT_Act_OT_Feb > 0 Then
                        .Feb_Act = Me.ObjCD.Convert_ToString(GT_Act_OT_Feb)
                    End If
                    If GT_Act_OT_Mar > 0 Then
                        .Mar_Act = Me.ObjCD.Convert_ToString(GT_Act_OT_Mar)
                    End If
                    If GT_Act_OT_Apr > 0 Then
                        .Apr_Act = Me.ObjCD.Convert_ToString(GT_Act_OT_Apr)
                    End If
                    If GT_Act_OT_May > 0 Then
                        .May_Act = Me.ObjCD.Convert_ToString(GT_Act_OT_May)
                    End If
                    If GT_Act_OT_Jun > 0 Then
                        .Jun_Act = Me.ObjCD.Convert_ToString(GT_Act_OT_Jun)
                    End If
                    If GT_Act_OT_Jul > 0 Then
                        .Jul_Act = Me.ObjCD.Convert_ToString(GT_Act_OT_Jul)
                    End If
                    If GT_Act_OT_Aug > 0 Then
                        .Aug_Act = Me.ObjCD.Convert_ToString(GT_Act_OT_Aug)
                    End If
                    If GT_Act_OT_Sep > 0 Then
                        .Sep_Act = Me.ObjCD.Convert_ToString(GT_Act_OT_Sep)
                    End If
                    If GT_Act_OT_Oct > 0 Then
                        .Oct_Act = Me.ObjCD.Convert_ToString(GT_Act_OT_Oct)
                    End If
                    If GT_Act_OT_Nov > 0 Then
                        .Nov_Act = Me.ObjCD.Convert_ToString(GT_Act_OT_Nov)
                    End If
                    If GT_Act_OT_Dec > 0 Then
                        .Dec_Act = Me.ObjCD.Convert_ToString(GT_Act_OT_Dec)
                    End If
                    If GT_Act_OT_Total > 0 Then
                        .Total_Act = Me.ObjCD.Convert_ToString(GT_Act_OT_Total)
                    End If

                    'Tentukan Apakah Text Color Red Jika Over Budget
                    If GT_OT_Jan < GT_Act_OT_Jan Then
                        .Is_Red_Jan = 1
                    End If
                    If GT_OT_Feb < GT_Act_OT_Feb Then
                        .Is_Red_Feb = 1
                    End If
                    If GT_OT_Mar < GT_Act_OT_Mar Then
                        .Is_Red_Mar = 1
                    End If
                    If GT_OT_Apr < GT_Act_OT_Apr Then
                        .Is_Red_Apr = 1
                    End If
                    If GT_OT_May < GT_Act_OT_May Then
                        .Is_Red_May = 1
                    End If
                    If GT_OT_Jun < GT_Act_OT_Jun Then
                        .Is_Red_Jun = 1
                    End If
                    If GT_OT_Jul < GT_Act_OT_Jul Then
                        .Is_Red_Jul = 1
                    End If
                    If GT_OT_Aug < GT_Act_OT_Aug Then
                        .Is_Red_Aug = 1
                    End If
                    If GT_OT_Sep < GT_Act_OT_Sep Then
                        .Is_Red_Sep = 1
                    End If
                    If GT_OT_Oct < GT_Act_OT_Oct Then
                        .Is_Red_Oct = 1
                    End If
                    If GT_OT_Nov < GT_Act_OT_Nov Then
                        .Is_Red_Nov = 1
                    End If
                    If GT_OT_Dec < GT_Act_OT_Dec Then
                        .Is_Red_Dec = 1
                    End If
                    If GT_OT_Total < GT_Act_OT_Total Then
                        .Is_Red_Total = 1
                    End If

                    MaxUrutan += 1
                    rowC.Urutan = MaxUrutan
                End With
               
                tblC.Rows.Add(rowC)
                'Isi Tabel Grand Total Close
            End If

            report.SetDataSource(ds)
            report.SetParameterValue(report.Parameter_ParamNoTransaksi.ParameterFieldName, ParNoTransaksi)
            report.SetParameterValue(report.Parameter_ParamTglTransaksi.ParameterFieldName, ParTglTransaksi)
            report.SetParameterValue(report.Parameter_ParamPeriod.ParameterFieldName, ParPeriod)
            report.SetParameterValue(report.Parameter_ParamRemarks.ParameterFieldName, ParRemarks)
            report.SetParameterValue(report.Parameter_ParamSetOverBudgetColor.ParameterFieldName, RedColorOverBudget)
            report.SetParameterValue(report.Parameter_MaxUrutan.ParameterFieldName, MaxUrutan)
            vCrv.ReportSource = report
            vCrv.Show()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function Cetak_CalculateActualOT(ByVal vPrimaryDept As Double, ByVal vDate As Date, _
    ByVal vTbl2 As DseTrans_SPL_BudgetCetak.tbl_total_ot_attendanceDataTable) As Double
        Dim row2 As DseTrans_SPL_BudgetCetak.tbl_total_ot_attendanceRow
        Dim TotalOT As Double = 0

        Try
            vTbl2.DefaultView.RowFilter = "[Period] = #" & vDate.Date & "# AND [primary departemen] = " & vPrimaryDept
            If vTbl2.DefaultView.Count > 0 Then
                row2 = vTbl2.DefaultView(0).Row
                TotalOT += row2.Total_OT_Attendance
            End If

            Return TotalOT
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Public Function Cetak(ByVal vPrimary As Double, ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer) As Boolean
    '    Dim report As New CrpTrans_SPL_Budget
    '    Dim ds As New DseTrans_SPL_BudgetCetak
    '    Dim dta As New DseTrans_SPL_BudgetCetakTableAdapters.DataTable1TableAdapter
    '    Dim tbl, tblDum As New DseTrans_SPL_BudgetCetak.DataTable1DataTable
    '    Dim row As DseTrans_SPL_BudgetCetak.DataTable1Row
    '    Dim dta2 As New DseTrans_SPL_BudgetCetakTableAdapters.DataTable2TableAdapter
    '    Dim tbl2 As New DseTrans_SPL_BudgetCetak.DataTable2DataTable
    '    Dim dta3 As New DseTrans_SPL_BudgetCetakTableAdapters.DataTable3TableAdapter
    '    Dim tbl3 As New DseTrans_SPL_BudgetCetak.DataTable3DataTable
    '    Dim dta4 As New DseTrans_SPL_BudgetCetakTableAdapters.tblmaster_karyawanTableAdapter
    '    Dim tbl4 As New DseTrans_SPL_BudgetCetak.tblmaster_karyawanDataTable
    '    Dim dta5 As New DseTrans_SPL_BudgetTableAdapters.tblmaster_departemenTableAdapter
    '    Dim tbl5 As New DseTrans_SPL_Budget.tblmaster_departemenDataTable
    '    Dim row5 As DseTrans_SPL_Budget.tblmaster_departemenRow
    '    Dim tblC As DseTrans_SPL_BudgetCetak.tbl_cetakDataTable = ds.tbl_cetak
    '    Dim rowC As DseTrans_SPL_BudgetCetak.tbl_cetakRow

    '    Dim ParNoTransaksi As String = ""
    '    Dim ParTglTransaksi As Date = Now
    '    Dim ParPeriod As String = ""
    '    Dim ParRemarks As String = ""
    '    Dim DateFirst As Date = Now
    '    Dim DateLast As Date = Now
    '    Dim Pemberitahuan As String = ""
    '    Dim hasil As MsgBoxResult

    '    Dim ActualOT_Jan, ActualOT_Feb, ActualOT_Mar, ActualOT_Apr As Double
    '    Dim ActualOT_May, ActualOT_Jun, ActualOT_Jul, ActualOT_Aug As Double
    '    Dim ActualOT_Sep, ActualOT_Oct, ActualOT_Nov, ActualOT_Dec As Double
    '    Dim TotalBudgetOT As Double = 0
    '    Dim TotalActualOT As Double = 0

    '    Try
    '        dta.Connection.ConnectionString = Core_Module.GetConstring
    '        dta2.Connection.ConnectionString = Core_Module.GetConstring
    '        dta3.Connection.ConnectionString = Core_Module.GetConstring
    '        dta4.Connection.ConnectionString = Core_Module.GetConstring
    '        dta5.Connection.ConnectionString = Core_Module.GetConstring

    '        If dta.Fill(tbl, vPrimary) > 0 Then

    '            For i As Integer = 0 To tbl.Rows.Count - 1
    '                row = tbl.Rows(i)
    '                If i = 0 Then
    '                    ParNoTransaksi = row.No_Transaksi
    '                    ParTglTransaksi = row._Date
    '                    ParPeriod = row.Period.Year
    '                    ParRemarks = row.Remarks

    '                    DateFirst = Me.ObjCD.GetFisrtDateOfTheYear(row.Period)
    '                    DateLast = Me.ObjCD.GetLastDateOfTheYear(row.Period)
    '                    dta.Fill_2(tblDum, vPrimary, row.Period.Date)
    '                    dta2.Fill(tbl2, DateLast, DateFirst, ClsTrans_SPL.EnumStatus.Approved)
    '                    dta3.Fill(tbl3, DateFirst, DateLast)
    '                    dta4.Fill(tbl4)
    '                    dta5.Fill(tbl5)
    '                End If

    '                tbl5.DefaultView.RowFilter = "[Primary] = " & row.Primary_Department
    '                While tbl5.DefaultView.Count > 0
    '                    row5 = tbl5.DefaultView(0).Row
    '                    tbl5.Rows.Remove(row5)
    '                End While
    '            Next

    '            For i As Integer = 0 To tblDum.Rows.Count - 1
    '                row = tblDum.Rows(i)
    '                tbl5.DefaultView.RowFilter = "[Primary] = " & row.Primary_Department
    '                While tbl5.DefaultView.Count > 0
    '                    row5 = tbl5.DefaultView(0).Row
    '                    tbl5.Rows.Remove(row5)
    '                End While
    '            Next

    '            If tbl5.Rows.Count > 0 Then
    '                For Each row5 In tbl5
    '                    If Pemberitahuan = "" Then
    '                        Pemberitahuan = row5.Description
    '                    Else
    '                        Pemberitahuan &= ", " & row5.Description
    '                    End If
    '                Next

    '                hasil = MsgBox("The following departments that have not been registered in the " & ParPeriod & " period" & vbCrLf & _
    '                Pemberitahuan & vbCrLf & "Continue printing ?", MsgBoxStyle.YesNo, "Confirm")
    '                If hasil = MsgBoxResult.No Then
    '                    Return False
    '                End If
    '            End If

    '            For i As Integer = 0 To tbl.Rows.Count - 1
    '                row = tbl.Rows(i)

    '                rowC = tblC.Newtbl_cetakRow
    '                With rowC
    '                    .Primary_Department = row.Primary_Department
    '                    .Dept_Code = row.Dept_Code
    '                    .Dept_Desc = row.Dept_Desc
    '                    If row.Jan > 0 Then
    '                        .Jan = Me.ObjCD.Convert_ToString(row.Jan)
    '                    End If
    '                    If row.Feb > 0 Then
    '                        .Feb = Me.ObjCD.Convert_ToString(row.Feb)
    '                    End If
    '                    If row.Mar > 0 Then
    '                        .Mar = Me.ObjCD.Convert_ToString(row.Mar)
    '                    End If
    '                    If row.Apr > 0 Then
    '                        .Apr = Me.ObjCD.Convert_ToString(row.Apr)
    '                    End If
    '                    If row.May > 0 Then
    '                        .May = Me.ObjCD.Convert_ToString(row.May)
    '                    End If
    '                    If row.Jun > 0 Then
    '                        .Jun = Me.ObjCD.Convert_ToString(row.Jun)
    '                    End If
    '                    If row.Jul > 0 Then
    '                        .Jul = Me.ObjCD.Convert_ToString(row.Jul)
    '                    End If
    '                    If row.Aug > 0 Then
    '                        .Aug = Me.ObjCD.Convert_ToString(row.Aug)
    '                    End If
    '                    If row.Sep > 0 Then
    '                        .Sep = Me.ObjCD.Convert_ToString(row.Sep)
    '                    End If
    '                    If row.Oct > 0 Then
    '                        .Oct = Me.ObjCD.Convert_ToString(row.Oct)
    '                    End If
    '                    If row.Nov > 0 Then
    '                        .Nov = Me.ObjCD.Convert_ToString(row.Nov)
    '                    End If
    '                    If row.Dec > 0 Then
    '                        .Dec = Me.ObjCD.Convert_ToString(row.Dec)
    '                    End If

    '                    'JAN
    '                    ActualOT_Jan = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("1/1/" & ParPeriod), tbl2, tbl3, tbl4)
    '                    If ActualOT_Jan > 0 Then
    '                        .Jan_Act = Me.ObjCD.Convert_ToString(ActualOT_Jan)
    '                    End If
    '                    'FEB
    '                    ActualOT_Feb = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("2/1/" & ParPeriod), tbl2, tbl3, tbl4)
    '                    If ActualOT_Feb > 0 Then
    '                        .Feb_Act = Me.ObjCD.Convert_ToString(ActualOT_Feb)
    '                    End If
    '                    'MAR
    '                    ActualOT_Mar = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("3/1/" & ParPeriod), tbl2, tbl3, tbl4)
    '                    If ActualOT_Mar > 0 Then
    '                        .Mar_Act = Me.ObjCD.Convert_ToString(ActualOT_Mar)
    '                    End If
    '                    'APR
    '                    ActualOT_Apr = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("4/1/" & ParPeriod), tbl2, tbl3, tbl4)
    '                    If ActualOT_Apr > 0 Then
    '                        .Apr_Act = Me.ObjCD.Convert_ToString(ActualOT_Apr)
    '                    End If
    '                    'MAY
    '                    ActualOT_May = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("5/1/" & ParPeriod), tbl2, tbl3, tbl4)
    '                    If ActualOT_May > 0 Then
    '                        .May_Act = Me.ObjCD.Convert_ToString(ActualOT_May)
    '                    End If
    '                    'JUN
    '                    ActualOT_Jun = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("6/1/" & ParPeriod), tbl2, tbl3, tbl4)
    '                    If ActualOT_Jun > 0 Then
    '                        .Jun_Act = Me.ObjCD.Convert_ToString(ActualOT_Jun)
    '                    End If
    '                    'JUL
    '                    ActualOT_Jul = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("7/1/" & ParPeriod), tbl2, tbl3, tbl4)
    '                    If ActualOT_Jul > 0 Then
    '                        .Jul_Act = Me.ObjCD.Convert_ToString(ActualOT_Jul)
    '                    End If
    '                    'AUG
    '                    ActualOT_Aug = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("8/1/" & ParPeriod), tbl2, tbl3, tbl4)
    '                    If ActualOT_Aug > 0 Then
    '                        .Aug_Act = Me.ObjCD.Convert_ToString(ActualOT_Aug)
    '                    End If
    '                    'SEP
    '                    ActualOT_Sep = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("9/1/" & ParPeriod), tbl2, tbl3, tbl4)
    '                    If ActualOT_Sep > 0 Then
    '                        .Sep_Act = Me.ObjCD.Convert_ToString(ActualOT_Sep)
    '                    End If
    '                    'OCT
    '                    ActualOT_Oct = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("10/1/" & ParPeriod), tbl2, tbl3, tbl4)
    '                    If ActualOT_Oct > 0 Then
    '                        .Oct_Act = Me.ObjCD.Convert_ToString(ActualOT_Oct)
    '                    End If
    '                    'NOV
    '                    ActualOT_Nov = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("11/1/" & ParPeriod), tbl2, tbl3, tbl4)
    '                    If ActualOT_Nov > 0 Then
    '                        .Nov_Act = Me.ObjCD.Convert_ToString(ActualOT_Nov)
    '                    End If
    '                    'DEC
    '                    ActualOT_Dec = Me.Cetak_CalculateActualOT(row.Primary_Department, CDate("12/1/" & ParPeriod), tbl2, tbl3, tbl4)
    '                    If ActualOT_Dec > 0 Then
    '                        .Dec_Act = Me.ObjCD.Convert_ToString(ActualOT_Dec)
    '                    End If

    '                    'HITUNG TOTAL BUDGET OT
    '                    TotalBudgetOT = row.Jan + row.Feb + row.Mar + row.Apr + row.May + row.Jun + _
    '                                    row.Jul + row.Aug + row.Sep + row.Oct + row.Nov + row.Dec
    '                    If TotalBudgetOT > 0 Then
    '                        .Total = Me.ObjCD.Convert_ToString(TotalBudgetOT)
    '                    End If
    '                    'HITUNG TOTAL ACTUAL OT
    '                    TotalActualOT = ActualOT_Jan + ActualOT_Feb + ActualOT_Mar + ActualOT_Apr + _
    '                                    ActualOT_May + ActualOT_Jun + ActualOT_Jul + ActualOT_Aug + _
    '                                    ActualOT_Sep + ActualOT_Oct + ActualOT_Nov + ActualOT_Dec
    '                    If TotalActualOT > 0 Then
    '                        .Total_Act = Me.ObjCD.Convert_ToString(TotalActualOT)
    '                    End If
    '                End With
    '                tblC.Rows.Add(rowC)
    '            Next
    '        End If

    '        report.SetDataSource(ds)
    '        report.SetParameterValue(report.Parameter_ParamNoTransaksi.ParameterFieldName, ParNoTransaksi)
    '        report.SetParameterValue(report.Parameter_ParamTglTransaksi.ParameterFieldName, ParTglTransaksi)
    '        report.SetParameterValue(report.Parameter_ParamPeriod.ParameterFieldName, ParPeriod)
    '        report.SetParameterValue(report.Parameter_ParamRemarks.ParameterFieldName, ParRemarks)
    '        vCrv.ReportSource = report
    '        vCrv.Show()

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return False
    '    End Try
    'End Function

    'Private Function Cetak_CalculateActualOT(ByVal vPrimaryDept As Double, ByVal vDate As Date, _
    'ByVal vTbl2 As DseTrans_SPL_BudgetCetak.DataTable2DataTable, _
    'ByVal vTbl3 As DseTrans_SPL_BudgetCetak.DataTable3DataTable, _
    'ByVal vTbl4 As DseTrans_SPL_BudgetCetak.tblmaster_karyawanDataTable) As Double
    '    Dim tbl2 As New DseTrans_SPL_BudgetCetak.DataTable2DataTable
    '    Dim tbl3 As New DseTrans_SPL_BudgetCetak.DataTable3DataTable
    '    Dim row2 As DseTrans_SPL_BudgetCetak.DataTable2Row
    '    Dim row3 As DseTrans_SPL_BudgetCetak.DataTable3Row
    '    Dim row4 As DseTrans_SPL_BudgetCetak.tblmaster_karyawanRow
    '    Dim DateFirst As Date = Me.ObjCD.GetFirstDateofTheMonth(vDate)
    '    Dim DateLast As Date = Me.ObjCD.GetLasDateOfTheMonth(vDate)
    '    Dim TotalOT As Double = 0

    '    Try
    '        vTbl2.DefaultView.RowFilter = "[Date] >= #" & DateFirst.Date & "# AND [Date] <= #" & _
    '        DateLast.Date & "# AND [primary departemen] = " & vPrimaryDept
    '        If vTbl2.DefaultView.Count > 0 Then
    '            tbl2.Merge(vTbl2.DefaultView.ToTable)
    '        End If
    '        vTbl3.DefaultView.RowFilter = "[Date] >= #" & DateFirst.Date & "# AND [Date] <= #" & _
    '        DateLast.Date & "# AND [primary departemen] = " & vPrimaryDept
    '        If vTbl3.DefaultView.Count > 0 Then
    '            tbl3.Merge(vTbl3.DefaultView.ToTable)
    '        End If

    '        While DateFirst.Date <= DateLast.Date
    '            vTbl4.DefaultView.RowFilter = "[primary departemen] = " & vPrimaryDept
    '            For Each dv As DataRowView In vTbl4.DefaultView
    '                row4 = dv.Row
    '                tbl3.DefaultView.RowFilter = "[Date] = #" & DateFirst.Date & "# AND [primary karyawan] = " & row4.Primary
    '                If tbl3.DefaultView.Count > 0 Then
    '                    row3 = tbl3.DefaultView(0).Row
    '                    TotalOT += row3.Total_OT_Attendance
    '                Else
    '                    tbl2.DefaultView.RowFilter = "[Date] = #" & DateFirst.Date & "# AND [Primary Employee] = " & row4.Primary
    '                    If tbl2.DefaultView.Count > 0 Then
    '                        row2 = tbl2.DefaultView(0).Row
    '                        TotalOT += row2.Total_OT_Approved
    '                    End If
    '                End If
    '            Next

    '            DateFirst = DateFirst.Date.AddDays(1)
    '        End While

    '        Return TotalOT
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function
End Class
