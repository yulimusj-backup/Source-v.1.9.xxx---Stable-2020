Option Explicit On
Option Strict Off

Imports System.Windows.Forms

Public Class ClsTransaksiLoan
    Dim mPrimary As Double = -1
    Dim mDocNo As String = ""
    Dim mJournalNo As String = ""
    Dim mJournalDesc As String = ""
    Dim mUrutan As Integer = -1
    Dim mDate As Date = #1/1/1950#
    Dim mFirstPaymentDate As Date = #1/1/1950#
    Dim mPrimaryKaryawan As Double = -1
    Dim mLoanAmount As Decimal = 0
    Dim mLoanType As EnumLoanType = EnumLoanType.NoRate
    Dim mPersenRate As Decimal = 0
    Dim mInstallmentTo As Double = -1
    Dim mPrimaryAccCashOut As Double = -1
    Dim mPrimaryAccRev As Double = -1
    Dim mRemarks As String = ""
    Dim mStatus As EnumStatusLoan = EnumStatusLoan.OnGoing
    Dim mPrimaryJurnal As Double = -1
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As DateTime = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As DateTime = Now
    Dim mTimestamp As Date = Now

    Dim ObjItem As New ClsTransaksiLoan_Item
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjJurnal As New AccModule.ClsTransaksiJournal

    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Enum EnumFrmTransLoanDari As Integer
        FrmTransLoan
        FrmTransLoan_Beginning_Balance
    End Enum

    Enum EnumLoanType
        NoRate
        Annuity
        Flat
        Effective
    End Enum

    Enum EnumStatusLoan
        OnGoing
        Finished
    End Enum

    Enum EnumSreach
        Document_no
        Date_
        Employee_code
        Employee_name
        Remarks
    End Enum

    Enum EnumJurnalOtomatis
        Simpan
        Edit
        Hapus
    End Enum

#Region "Property"
    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property DocNo() As String
        Get
            Return mDocNo
        End Get
        Set(ByVal value As String)
            mDocNo = value
        End Set
    End Property

    Public Property JournalNo() As String
        Get
            Return mJournalNo
        End Get
        Set(ByVal value As String)
            mJournalNo = value
        End Set
    End Property

    Public Property JournalDesc() As String
        Get
            Return mJournalDesc
        End Get
        Set(ByVal value As String)
            mJournalDesc = value
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

    Public Property FirstPaymentDate() As Date
        Get
            Return mFirstPaymentDate
        End Get
        Set(ByVal value As Date)
            mFirstPaymentDate = value
        End Set
    End Property

    Public Property PrimaryKaryawan() As Double
        Get
            Return mPrimaryKaryawan
        End Get
        Set(ByVal value As Double)
            mPrimaryKaryawan = value
        End Set
    End Property

    Public Property LoanAmount() As Decimal
        Get
            Return mLoanAmount
        End Get
        Set(ByVal value As Decimal)
            mLoanAmount = value
        End Set
    End Property

    Public Property LoanType() As EnumLoanType
        Get
            Return mLoanType
        End Get
        Set(ByVal value As EnumLoanType)
            mLoanType = value
        End Set
    End Property

    Public Property PersenRate() As Decimal
        Get
            Return mPersenRate
        End Get
        Set(ByVal value As Decimal)
            mPersenRate = value
        End Set
    End Property

    Public Property InstallmentTo() As Double
        Get
            Return mInstallmentTo
        End Get
        Set(ByVal value As Double)
            mInstallmentTo = value
        End Set
    End Property

    Public Property PrimaryAccCashOut() As Double
        Get
            Return mPrimaryAccCashOut
        End Get
        Set(ByVal value As Double)
            mPrimaryAccCashOut = value
        End Set
    End Property

    Public Property PrimaryAccRev() As Double
        Get
            Return mPrimaryAccRev
        End Get
        Set(ByVal value As Double)
            mPrimaryAccRev = value
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

    Public Property Status() As EnumStatusLoan
        Get
            Return mStatus
        End Get
        Set(ByVal value As EnumStatusLoan)
            mStatus = value
        End Set
    End Property

    Public Property PrimaryJurnal() As Double
        Get
            Return mPrimaryJurnal
        End Get
        Set(ByVal value As Double)
            mPrimaryJurnal = value
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

    Public Function ApakahBolehedit(ByVal vUserCode As String, ByVal vPrimaryTransaksi As Double, ByVal vBeginingBalance As Boolean) As Boolean
        Dim Dta As New DseTransaksiLoanTableAdapters.tbltrans_loanTableAdapter
        Dim ObjClosing As New AccModule.ClsTransaksiClosing
        Try
            '1.strict edit
            Dim UserCode As String = Dta.ScalarQuery_GetUserCodeByPrimary(vPrimaryTransaksi)
            If UserCode <> vUserCode Then
                MsgBox("Only user with ID:" + UserCode + " can edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '2.status aktif tidak boleh di edit
            Dim status As Integer = Dta.ScalarQuery_GetStatusByPrimary(vPrimaryTransaksi)
            If status = ClsTransaksiLoan.EnumStatusLoan.Finished Then
                MsgBox("This document no : " + Me.DocNo.ToString + " cannot edit because loan status is finished.", _
                MsgBoxStyle.Information, "Information")
            End If

            '3.Jika sudah terjadi closing tidak boleh di edit, hanya berlaku untuk yang bukan begining balance
            If vBeginingBalance = False Then
                Dim closing As Date = ObjClosing.GetClosingDate
                If Me.Date_ < closing Then
                    MsgBox("You cannot edit this data because this data is closed in period : " + Format(closing, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '4.tidak boleh jika sudah ada di transaksi loan payment
            If Me.GetJmlLoanPayment(vPrimaryTransaksi) > 0 Then
                MsgBox("You cannot edit this data, there are loan payment refering to this data", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vUserCode As String, ByVal vPrimaryTransaksi As Double, ByVal vBeginingBalance As Boolean) As Boolean
        Dim Dta As New DseTransaksiLoanTableAdapters.tbltrans_loanTableAdapter
        Dim ObjClosing As New AccModule.ClsTransaksiClosing
        Try
            '1.strict edit
            Dim UserCode As String = Dta.ScalarQuery_GetUserCodeByPrimary(vPrimaryTransaksi)
            If UserCode <> vUserCode Then
                MsgBox("Only user with ID:" + UserCode + " can delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '2.status aktif tidak boleh di edit
            Dim status As Integer = Dta.ScalarQuery_GetStatusByPrimary(vPrimaryTransaksi)
            If status = ClsTransaksiLoan.EnumStatusLoan.Finished Then
                MsgBox("This document no : " + Me.DocNo.ToString + " cannot delete because loan status is finished.", _
                MsgBoxStyle.Information, "Information")
            End If

            '3.Jika sudah terjadi closing tidak boleh di edit
            If vBeginingBalance = False Then
                Dim closing As Date = ObjClosing.GetClosingDate
                If Me.Date_ < closing Then
                    MsgBox("You cannot delete this data because this data is closed in period : " + Format(closing, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '4. tidak boleh hapus jika sudah ada payment yang merujuk
            If Me.GetJmlLoanPayment(vPrimaryTransaksi) > 0 Then
                MsgBox("You cannot delete this data, there are loan payment refering to this data", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
            ObjClosing = Nothing
        End Try
    End Function

    Private Function ApakahDataBenar(ByVal FrmLoan As FrmTransaksiLoan, ByVal vBaru As Boolean) As Boolean
        Dim ObjClosing As New AccModule.ClsTransaksiClosing
        Dim payrollConventionDate As Date = ObjClosing.GetPayrollConvertionDate()
        Try
            '1.tidak boleh double
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.DocNo) = True Then
                    Me.mDocNo = Me.GetNoBaru(Me.mDate)
                End If
            Else
                Dim NoTransaksiLama As String = Me.GetDocNoByPrimary(Me.mPrimary)
                If NoTransaksiLama <> Me.mDocNo Then
                    If Me.ApakahSudahAda(Me.mDocNo) = True Then
                        Me.mDocNo = Me.GetNoBaru(Me.mDate)
                    End If
                End If
            End If
            'If Me.ApakahSudahAda(Me.DocNo) = True Then
            '    Return False
            'End If
            '2.Harus input Convention date HR
            If payrollConventionDate = #1/1/1950# Then
                MsgBox("Please input payroll convention date in global setting data.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If FrmLoan.EnumDipanggilDari = ClsTransaksiLoan.EnumFrmTransLoanDari.FrmTransLoan Then
                '3.Harus dibuat setelah konvention date HR
                If Me.mDate < payrollConventionDate Then
                    MsgBox("Please make sure date in transaction loan, loan's date must be bigger than payroll convention date.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf FrmLoan.EnumDipanggilDari = EnumFrmTransLoanDari.FrmTransLoan_Beginning_Balance Then
                '3.Harus dibuat setelah konvention date HR
                If Me.mDate >= payrollConventionDate Then
                    MsgBox("Please make sure date in transaction beginning balance loan, loan's date must be smaller than payroll convention date.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjClosing = Nothing
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiLoanTableAdapters.tbltrans_loanTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim JmlDb As Nullable(Of Integer)
        Try
            'cari tanggal sampai dan dari
            TanggalEnd = Me.ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = Me.ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)


            'cari no urut
            JmlDb = DTA.ScalarQuery_GetMaxUrutan(TanggalStart, TanggalEnd)
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
            Return NoUrut + "/" + "LOAN" + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vFrm As FrmTransaksiLoan, _
   ByVal vDgvItem As Windows.Forms.DataGridView, ByVal vDefPrec As Integer) As Boolean
        Dim DTA As New DseTransaksiLoanTableAdapters.tbltrans_loanTableAdapter
        Dim TBL As New DseTransaksiLoan.tbltrans_loanDataTable
        Dim row As DseTransaksiLoan.tbltrans_loanRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mDocNo = .Document_No
                    Me.mUrutan = .Urutan
                    Me.mDate = ._Date
                    Me.mFirstPaymentDate = .First_Payment_Date
                    Me.mPrimaryKaryawan = .Primary_Karyawan
                    Me.mLoanAmount = .Loan_Amount
                    Me.mLoanType = .Loan_Type
                    Me.mPersenRate = .Persen_Rate
                    Me.mInstallmentTo = .Installment_To
                    Me.mPrimaryAccCashOut = .Primary_Cash_Out_Account
                    Me.mPrimaryAccRev = .Primary_Receivable_Account
                    Me.mRemarks = .Remarks
                    Me.mStatus = .Status
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mPrimaryJurnal = .Primary_Jurnal
                    Me.mJournalNo = Me.ObjJurnal.GetRefCode(.Primary_Jurnal)
                    Me.mJournalDesc = Me.ObjJurnal.GetDescription(.Primary_Jurnal)

                    Me.ObjItem.IsiDGVItem(vDgvItem, .Primary, vFrm, vDefPrec)
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

    Public Function ApakahSudahAda(ByVal vDocNo As String) As Boolean
        Dim dta As New DseTransaksiLoanTableAdapters.tbltrans_loanTableAdapter
        Dim dummy As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetJmlByDocNo(vDocNo)
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

    Public Function GetDocNoByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseTransaksiLoanTableAdapters.tbltrans_loanTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetDocNoByPrimary(vPrimary)
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
        Dim dta As New DseTransaksiLoanTableAdapters.tbltrans_loanTableAdapter
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

    Public Sub IsiCboCashAcc(ByVal vCbo As ComboBox)
        Dim DTA As New DseTransaksiLoanTableAdapters.tblmaster_akunTableAdapter
        Dim TBL As New DseTransaksiLoan.tblmaster_akunDataTable

        Try
            vCbo.DataSource = Nothing
            DTA.FillBy_AccCashOut(TBL, Core_Module.ClsMasterAkun1.EnumTipeAkun.Asset, _
            Core_Module.ClsMasterAkun1.EnumTipeAkun2.Cash_And_Bank)

            vCbo.DataSource = TBL
            vCbo.DisplayMember = "Account Code"

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetJmlLoanPayment(ByVal vPrimaryLoan As Double) As Double
        Dim dta As New DseTransaksiLoanTableAdapters.tbltrans_loan_payment_temTableAdapter
        Dim Jml As Nullable(Of Long)
        Try
            Jml = dta.ScalarQuery_GetJmlByPrimaryLoan(vPrimaryLoan)
            If Jml.HasValue Then
                Return Jml.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vDGV As DataGridView, _
   ByVal frm As FrmTransaksiLoan) As Boolean
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

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal vDGV As DataGridView, ByVal frm As FrmTransaksiLoan) As Boolean
        Dim DTA As New DseTransaksiLoanTableAdapters.tbltrans_loanTableAdapter
        Dim BufferDoc As String = Me.mDocNo
        Try
            If Me.ApakahDataBenar(frm, True) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                Me.mUrutan = CDbl(Me.mDocNo.Substring(0, 6))
                If DTA.Insert(Me.mDocNo, Me.mUrutan, Me.mDate, Me.mFirstPaymentDate, Me.mPrimaryKaryawan, Me.mLoanAmount, Me.mLoanType, _
                 Me.mPersenRate, Me.mInstallmentTo, Me.mPrimaryAccCashOut, Me.mPrimaryAccRev, Me.mRemarks, Me.mStatus, vUserCode, Now, vUserCode, _
                 Now, Now, -1) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimaryByDocNo(Me.mDocNo)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStampByPrimary(Me.mPrimary)
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan item
                    Me.ObjItem.SimpanItem(Me.Primary, vDGV, frm)

                    'jurnal otomatis
                    If frm.EnumDipanggilDari = EnumFrmTransLoanDari.FrmTransLoan Then
                        Me.UpdateJournalNumber(BufferDoc)
                        Me.JurnalOtomatis(vUserCode, EnumJurnalOtomatis.Simpan, frm)
                    End If

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
   ByVal frm As FrmTransaksiLoan) As Boolean
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

    Public Function SimpanPerubahan1(ByVal vUserCode As String, ByVal vDGV As DataGridView, ByVal frm As FrmTransaksiLoan) As Boolean
        Dim DTA As New DseTransaksiLoanTableAdapters.tbltrans_loanTableAdapter
        Dim BufferDoc As String = Me.mDocNo
        Try
            If Me.ApakahDataBenar(frm, False) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                Me.mUrutan = CDbl(Me.mDocNo.Substring(0, 6))
                If DTA.Update(Me.mDocNo, Me.mUrutan, Me.mDate, Me.mFirstPaymentDate, Me.mPrimaryKaryawan, Me.mLoanAmount, Me.mLoanType, _
                Me.mPersenRate, Me.mInstallmentTo, Me.mPrimaryAccCashOut, Me.mPrimaryAccRev, Me.mRemarks, Me.mStatus, vUserCode, Now, Now, _
                Me.mPrimaryJurnal, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStampByPrimary(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus item
                    Me.ObjItem.HapusItem(Me.mPrimary)

                    'simpan item
                    Me.ObjItem.SimpanItem(Me.Primary, vDGV, frm)

                    'jurnal otomatis
                    If frm.EnumDipanggilDari = EnumFrmTransLoanDari.FrmTransLoan Then
                        Me.UpdateJournalNumber(BufferDoc)
                        Me.JurnalOtomatis(vUserCode, EnumJurnalOtomatis.Edit, frm, Me.mPrimaryJurnal)
                    End If
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

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, ByVal frm As FrmTransaksiLoan) As Boolean
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

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal frm As FrmTransaksiLoan) As Boolean
        Dim DTA As New DseTransaksiLoanTableAdapters.tbltrans_loanTableAdapter
        Dim IsBeginingBalance As Boolean = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If frm.EnumDipanggilDari = EnumFrmTransLoanDari.FrmTransLoan_Beginning_Balance Then
                IsBeginingBalance = True
            Else
                IsBeginingBalance = False
            End If
            If Me.ApakahBolehHapus(vUserCode, vPrimary, IsBeginingBalance) = True Then
                DTA.Delete(vPrimary)
                Me.ObjItem.HapusItem(vPrimary)
                'jurnal otomatis
                If frm.EnumDipanggilDari = EnumFrmTransLoanDari.FrmTransLoan Then
                    Me.JurnalOtomatis(vUserCode, EnumJurnalOtomatis.Hapus, frm, ObjJurnal.GetPrimary(frm.txtJournalNo.Text))
                End If

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

    Public Function IsiDGV(ByVal Criteria As String, ByVal vDGV As DataGridView, ByVal vDTA As DseTransaksiLoanTableAdapters.DataTable1TableAdapter, _
    ByVal vTBL As DseTransaksiLoan.DataTable1DataTable, ByVal SearchCriteria As EnumSreach, ByVal tsl As ToolStripLabel, ByVal vTanggalDari As Date, _
    ByVal vTanggalSampai As Date, ByVal vFrm As FrmTransaksiLoan, Optional ByVal vJml As Integer = 0, Optional ByVal AllFilterData As Boolean = False) As Boolean
        Dim QueryResult As Double
        Dim StartT As Date
        Dim ObjClosing As New AccModule.ClsTransaksiClosing
        Dim PayrollConventionDate As Date = ObjClosing.GetPayrollConvertionDate()
        Try
            vDTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If AllFilterData = True Then
                'Analogi Beginning loan dan loan, jika beginning sebelum convention date, loan setelah
                If SearchCriteria = EnumSreach.Document_no Then
                    QueryResult = vDTA.FillBy_LikeDocNo(vTBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSreach.Date_ Then
                    QueryResult = vDTA.FillBy_Date(vTBL, vTanggalDari, vTanggalSampai)
                ElseIf SearchCriteria = EnumSreach.Employee_code Then
                    QueryResult = vDTA.FillBy_LikeEmpCode(vTBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSreach.Employee_name Then
                    QueryResult = vDTA.FillBy_LikeEmpName(vTBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSreach.Remarks Then
                    QueryResult = vDTA.FillBy_LikeRemark(vTBL, Criteria + "%")
                End If
            Else
                If vFrm.EnumDipanggilDari = EnumFrmTransLoanDari.FrmTransLoan Then
                    If SearchCriteria = EnumSreach.Document_no Then
                        QueryResult = vDTA.FillBy_LikeDocNoLoan(vTBL, Criteria + "%", PayrollConventionDate)
                    ElseIf SearchCriteria = EnumSreach.Date_ Then
                        QueryResult = vDTA.FillBy_DateLoan(vTBL, PayrollConventionDate, vTanggalDari, vTanggalSampai)
                    ElseIf SearchCriteria = EnumSreach.Employee_code Then
                        QueryResult = vDTA.FillBy_LikeEmpCodeLoan(vTBL, Criteria + "%", PayrollConventionDate)
                    ElseIf SearchCriteria = EnumSreach.Employee_name Then
                        QueryResult = vDTA.FillBy_LikeEmpNameLoan(vTBL, Criteria + "%", PayrollConventionDate)
                    ElseIf SearchCriteria = EnumSreach.Remarks Then
                        QueryResult = vDTA.FillBy_LikeRemarkLoan(vTBL, Criteria + "%", PayrollConventionDate)
                    End If
                ElseIf vFrm.EnumDipanggilDari = EnumFrmTransLoanDari.FrmTransLoan_Beginning_Balance Then
                    If SearchCriteria = EnumSreach.Document_no Then
                        QueryResult = vDTA.FillBy_LikeDocNoLoanBeginning(vTBL, Criteria + "%", PayrollConventionDate)
                    ElseIf SearchCriteria = EnumSreach.Date_ Then
                        QueryResult = vDTA.FillBy_DateLoanBeginning(vTBL, PayrollConventionDate, vTanggalDari, vTanggalSampai)
                    ElseIf SearchCriteria = EnumSreach.Employee_code Then
                        QueryResult = vDTA.FillBy_LikeEmpCodeLoanBeginning(vTBL, Criteria + "%", PayrollConventionDate)
                    ElseIf SearchCriteria = EnumSreach.Employee_name Then
                        QueryResult = vDTA.FillBy_LikeEmpNameLoanBeginning(vTBL, Criteria + "%", PayrollConventionDate)
                    ElseIf SearchCriteria = EnumSreach.Remarks Then
                        QueryResult = vDTA.FillBy_LikeRemarkLoanBeginning(vTBL, Criteria + "%", PayrollConventionDate)
                    End If
                End If
            End If

            'update dgv
            'vDGV.Refresh()

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
            ObjClosing = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vCriteria As String, ByVal vDGV As DataGridView, ByVal vSearchCriteria As EnumSreach, _
    ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date) As Double
        Dim Dta As New DseTransaksiLoanTableAdapters.tbl_transaksi_loan_viewTableAdapter
        Dim Dummy As Nullable(Of Long)
        Try

            Dta.Connection.ConnectionString = Core_Module.GetConstring

            If vSearchCriteria = EnumSreach.Document_no Then
                Dummy = Dta.ScalarQuery_GetJmlByDocNo(vCriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = EnumSreach.Date_ Then
                Dummy = Dta.ScalarQuery_GetJmlByDate(vTanggalDari, vTanggalSampai)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = EnumSreach.Employee_code Then
                Dummy = Dta.ScalarQuery_GetJmlByEmpCode(vCriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = EnumSreach.Employee_name Then
                Dummy = Dta.ScalarQuery_GetJmlByEmpName(vCriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = EnumSreach.Remarks Then
                Dummy = Dta.ScalarQuery_GetJmlByRemaks(vCriteria + "%")
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

    Public Sub JurnalOtomatis(ByVal vUserCode As String, ByVal vJournalOtomatis As EnumJurnalOtomatis, ByVal frm As FrmTransaksiLoan, _
    Optional ByVal vPrimaryJournal As Double = -1)
        Dim PrimaryJournal As Double = -1
        Dim ObjAcc As New Core_Module.ClsMasterAkun1
        Dim DtaJournalHeader As New DseTransaksiLoanTableAdapters.tbltrans_jurnalTableAdapter
        Dim DtaJournalItem As New DseTransaksiLoanTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim Dta As New DseTransaksiLoanTableAdapters.tbltrans_loanTableAdapter
        Dim desc As String = ""
        Try

            desc = Me.mJournalDesc
            'desc = frm.TxtDocumentNo.Text + "-" + frm.TxtEmployeeCode.Text + "-" + frm.TxtEmployeeName.Text
            If vJournalOtomatis = EnumJurnalOtomatis.Simpan Then
                'baru
                'simpan header
                DtaJournalHeader.Insert(Me.mJournalNo, Me.Date_, 0, AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, Me.Primary, _
                  Me.PrimaryKaryawan, desc, vUserCode, Now, vUserCode, Now, Now, -1, -1, 0, 1, 0, 0, 0, 0, AccModule.ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Loan, _
                ObjAcc.GetPrimaryCurrency(Me.PrimaryAccCashOut), AccModule.ClsTransaksiJournal.EnumTipeInOut._In, -1, -1)
                PrimaryJournal = DtaJournalHeader.ScalarQuery_GetPrimaryByDocNo(Me.mJournalNo)
                'simpan item
                'debet
                DtaJournalItem.Insert(PrimaryJournal, Me.PrimaryAccRev, Me.LoanAmount, 0, "", Me.LoanAmount, 0, 1, -1, -1)
                'kredit
                DtaJournalItem.Insert(PrimaryJournal, Me.PrimaryAccCashOut, 0, Me.LoanAmount, "", 0, Me.LoanAmount, 1, -1, -1)
                'update primary jurnal di tabel loan
                Dta.UpdateQuery_UpdatePrimaryJournal(PrimaryJournal, Me.Primary)
                Me.mPrimaryJurnal = PrimaryJournal
            ElseIf vJournalOtomatis = EnumJurnalOtomatis.Edit Then
                'edit
                'simpan header
                DtaJournalHeader.Update(Me.mJournalNo, Me.Date_, 0, AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, Me.Primary, _
                 Me.PrimaryKaryawan, desc, vUserCode, Now, Now, -1, -1, 0, 1, 0, 0, 0, 0, AccModule.ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Loan, _
                ObjAcc.GetPrimaryCurrency(Me.PrimaryAccCashOut), AccModule.ClsTransaksiJournal.EnumTipeInOut._In, -1, -1, vPrimaryJournal)
                vPrimaryJournal = Me.mPrimaryJurnal
                'hapus item
                DtaJournalItem.Delete(vPrimaryJournal)
                'simpan item
                'debet
                DtaJournalItem.Insert(vPrimaryJournal, Me.PrimaryAccRev, Me.LoanAmount, 0, "", Me.LoanAmount, 0, 1, -1, -1)
                'kredit
                DtaJournalItem.Insert(vPrimaryJournal, Me.PrimaryAccCashOut, 0, Me.LoanAmount, "", 0, Me.LoanAmount, 1, -1, -1)
            ElseIf vJournalOtomatis = EnumJurnalOtomatis.Hapus Then
                'hapus header
                DtaJournalHeader.Delete(vPrimaryJournal)
                'hapus item
                DtaJournalItem.Delete(vPrimaryJournal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaJournalHeader.Dispose()
            DtaJournalItem.Dispose()
            Dta.Dispose()
            ObjAcc = Nothing
        End Try
    End Sub

    Private Function ApakaJurnalNonIDR(ByVal vPrimaryAkun As Double) As Boolean
        Dim DTA As New AccModule.DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlNonIDR(Core_Module.PrimaryBaseCurrency, vPrimaryAkun).HasValue = True Then
                If DTA.ScalarQuery_GetJmlNonIDR(Core_Module.PrimaryBaseCurrency, vPrimaryAkun).Value <= 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub CetakDetail(ByVal vPrimaryLoan As Double, ByVal Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vTypeLoan As String)
        Dim ds As New DseTransaksiLoan
        Dim Dta As New DseTransaksiLoanTableAdapters.DataTable2TableAdapter
        Dim Tbl As DseTransaksiLoan.DataTable2DataTable = ds.DataTable2
        Dim Report As New CrpCetkDetail

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill(Tbl, vPrimaryLoan)

            Report.SetDataSource(ds)
            Report.SetParameterValue(Report.Parameter_LoanType.ParameterFieldName, vTypeLoan)
            Crv.ReportSource = Report
            Crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ds.Dispose()
            Dta.Dispose()
        End Try
    End Sub

    Public Function CetakJournal(ByVal vPrimary As Double, _
        ByVal Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer) As Boolean
        Dim Ds As New DseTransaksiLoan
        Dim DTA As New DseTransaksiLoanTableAdapters.DataTable3TableAdapter
        Dim TBL As DseTransaksiLoan.DataTable3DataTable = Ds.DataTable3
        Dim FormatDecimal As Integer = 0
        Dim Negatif As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Report As New CrpCetakJurnal
        Dim Report2 As New CrpCetakJurnal2
        Dim Report3 As New CrpCetakJurnal3
        Dim CetakJurnalPortrait As String = ""
        Dim KodeCur As String = ""
        Dim row As DseTransaksiLoan.DataTable3Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimary)
            KodeCur = ObjCur.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)

            CetakJurnalPortrait = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CetakJurnalPortrait)

            If CetakJurnalPortrait = "" Or CetakJurnalPortrait = "0" Then
                If Me.ApakaJurnalNonIDR(vPrimary) = False Then
                    'cetak lanscape tanda currency no idr
                    Report2.SetDataSource(Ds)
                    Report2.SetParameterValue(Report2.Parameter_FormatDecimal.ParameterFieldName, ObjCur.GetDecimal_DefaultCurrency)
                    If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                        Negatif = 1
                    Else
                        Negatif = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                        Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                    End If
                    Report2.SetParameterValue(Report2.Parameter_Negatif.ParameterFieldName, Negatif)
                    Report2.SetParameterValue(Report2.Parameter_KodeBaseCurrency.ParameterFieldName, KodeCur)
                    Report2.SetParameterValue(Report2.Parameter_Judul.ParameterFieldName, "Loan")

                    Crv.ReportSource = Report2
                    'Crv.Zoom(1)
                    Crv.Show()
                Else
                    Report3.SetDataSource(Ds)
                    For Each row In Ds.DataTable3.Rows
                        If row.Credit_Original > 0 Then
                            row.Debet_Original = row.Credit_Original
                        End If
                    Next
                    Report3.SetParameterValue(Report3.Parameter_FormatDecimal.ParameterFieldName, ObjCur.GetDecimal_DefaultCurrency)
                    If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                        Negatif = 1
                    Else
                        Negatif = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                        Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                    End If
                    Report3.SetParameterValue(Report3.Parameter_Negatif.ParameterFieldName, Negatif)
                    Report3.SetParameterValue(Report3.Parameter_KodeBaseCurrency.ParameterFieldName, KodeCur)
                    Report3.SetParameterValue(Report3.Parameter_Judul.ParameterFieldName, "Loan")

                    Crv.ReportSource = Report3
                    'Crv.Zoom(1)
                    Crv.Show()
                End If
            Else
                'cetak portrait
                Report.SetDataSource(Ds)
                Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, ObjCur.GetDecimal_DefaultCurrency)
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                    Negatif = 1
                Else
                    Negatif = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                End If
                Report.SetParameterValue(Report.Parameter_Negatif.ParameterFieldName, Negatif)
                Report.SetParameterValue(Report.Parameter_KodeBaseCurrency.ParameterFieldName, KodeCur)
                Report.SetParameterValue(Report.Parameter_Judul.ParameterFieldName, "Loan")

                Crv.ReportSource = Report
                Crv.Show()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCur = Nothing
            ObjSetting = Nothing
        End Try
    End Function

    Public Function GetInstallmentTo(ByVal vLoanNumber As String) As Double
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_itemTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetInstallmentTo(vLoanNumber)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub UpdateStatusLoan(ByVal vStatus As EnumStatusLoan, ByVal vLoanNumber As String)
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetStatusLoan(vStatus, vLoanNumber)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetPrimaryJournal(ByVal vPrimaryLoan As Double) As Double
        Dim d As New DseTransaksiLoanTableAdapters.tbltrans_loanTableAdapter
        Dim n As New Nullable(Of Double)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            n = d.ScalarQuery_GetPrimaryJournal(vPrimaryLoan)
            If n.HasValue Then Return n.Value
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            d.Dispose()
        End Try
    End Function

    Public Function GetJmlByJournalNo(ByVal vJournalNo As String) As Integer
        Dim d As New DseTransaksiLoanTableAdapters.tbltrans_loanTableAdapter
        Dim n As New Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            n = d.ScalarQuery_GetJmlByJournalNo(vJournalNo)
            If n.HasValue Then Return n.Value
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
        End Try
    End Function

    'fungsi ini untuk update no jurnal saat no document ke ubah karena antrian simpan berbarengan
    'selain update no document update juga no jurnal jika no jurnal sama dengan nilai buffer no document
    Private Sub UpdateJournalNumber(ByVal vBufferDoc As String)

        Try
            If Me.JournalNo = vBufferDoc Then
                Me.JournalNo = Me.DocNo
            Else
                If Me.ObjJurnal.ApakahSudahAda(Me.JournalNo) = True Then
                    Me.JournalNo = Me.DocNo
                End If
            End If
            If Me.JournalDesc.Contains(vBufferDoc) = True Then
                Me.JournalDesc = Me.JournalDesc.Replace(vBufferDoc, Me.DocNo)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
