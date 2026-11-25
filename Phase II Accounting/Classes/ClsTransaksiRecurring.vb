Option Explicit On
Option Strict Off

Public Class ClsTransaksiRecurring
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mRefNumber As String = ""
    Private mDate As Date = Now
    Private mUrutan As Double = -1
    Private mPrimaryDokumenAcuan As Double = -1
    Private mPrimaryPihakAcuan As Double = -1
    Private mDesc As String = ""
    Private mPrimaryAkunPayTo As Double = -1
    Private mPrimaryAkunDifference As Double = -1
    Private mPaid As Decimal = 0
    Private mExchangeRate As Decimal = 1
    Private mPaid_IDR As Decimal = 0
    Private mARPaid_IDR As Decimal = 0
    Private mDifference As Decimal = 0
    Private mExchageRateGainLoss As Decimal = 0
    Private mCreatedBy As String = ""
    Private mCreatedDate As Date = Now
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date = Now
    Private mTimestamp As Date = Now
    Private mPrimaryCurr As Double = -1
    Private mReferenceParty As String = ""

    Private mIgnoreIncrementOnSkip As Integer = 0
    Private mFrequency As Integer = 0
    Private mEvery As Integer = 0
    Private mLimit As Integer = 0
    Private mLimitValue As Double = 0
    Private mAktive As Integer = 0
    Private mStartDate As Date = Now
    Private mEndDate As Date = Now


    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian

    Public Enum EnumFrequency
        Daily
        Weekly
        Bi_Weekly
        Monthly
        Quarter_Yearly
        Semester
        Yearly
    End Enum

    Public Enum EnumStatus
        Planed
        Created
        Skiped
        Locked
    End Enum

    Public Enum EnumAction
        Create
        Update
        Edit
        Delete
        Skip
        Locks
        Delete_All
        Update_All
    End Enum

    Public Enum EnumEvery_Month
        First_Day_Of_The_Month
        Last_Day_Of_The_Month
        First_Monday
        First_Tuesday
        First_Wednesday
        First_Thursday
        First_Friday
        First_Saturday
        First_Sunday
        Last_Monday
        Last_Tuesday
        Last_Wednesday
        Last_Thusday
        Last_Friday
        Last_Saturday
        Last_Sunday
        _01
        _02
        _03
        _04
        _05
        _06
        _07
        _08
        _09
        _10
        _11
        _12
        _13
        _14
        _15
        _16
        _17
        _18
        _19
        _20
        _21
        _22
        _23
        _24
        _25
        _26
        _27
        _28
        _29
        _30
        _31
    End Enum
    Public Enum EnumEvery_Week
        Monday
        Tuesday
        Wednesday
        Thursday
        Friday
        Saturday
        Sunday
    End Enum

    Public Enum EnumIncrementStyle
        No_Increment
        Fixed_Increment
        Percentage_Increment
    End Enum

    Public Property ReferenceParty() As String
        Get
            Return Me.mReferenceParty
        End Get
        Set(ByVal value As String)
            mReferenceParty = value
        End Set
    End Property

    Public Property PrimaryCurrency() As Double
        Get
            Return Me.mPrimaryCurr
        End Get
        Set(ByVal value As Double)
            mPrimaryCurr = value
        End Set
    End Property

    Public Property IgnoreIncrementOnSkip() As Integer
        Get
            Return Me.mIgnoreIncrementOnSkip
        End Get
        Set(ByVal value As Integer)
            mIgnoreIncrementOnSkip = value
        End Set
    End Property

    Public Property Frequency() As Integer
        Get
            Return Me.mFrequency
        End Get
        Set(ByVal value As Integer)
            mFrequency = value
        End Set
    End Property

    Public Property Every() As Integer
        Get
            Return Me.mEvery
        End Get
        Set(ByVal value As Integer)
            mEvery = value
        End Set
    End Property


    Public Property Limit() As Integer
        Get
            Return Me.mLimit
        End Get
        Set(ByVal value As Integer)
            mLimit = value
        End Set
    End Property

    Public Property LimitValue() As Double
        Get
            Return Me.mLimitValue
        End Get
        Set(ByVal value As Double)
            mLimitValue = value
        End Set
    End Property

    Public Property Active() As Integer
        Get
            Return Me.mAktive
        End Get
        Set(ByVal value As Integer)
            mAktive = value
        End Set
    End Property

    Public Property StartDate() As Date
        Get
            Return Me.mStartDate
        End Get
        Set(ByVal value As Date)
            mStartDate = value
        End Set
    End Property

    Public Property EndDate() As Date
        Get
            Return Me.mEndDate
        End Get
        Set(ByVal value As Date)
            mEndDate = value
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

    Public Property Timestamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property

    'Public Property TipeJurnal2() As EnumTipeJurnal2
    '    Get
    '        Return Me.mTipeJurnal2
    '    End Get
    '    Set(ByVal value As EnumTipeJurnal2)
    '        mTipeJurnal2 = value
    '    End Set
    'End Property

    Public Property ExchageRateGainLoss() As Decimal
        Get
            Return Me.mExchageRateGainLoss
        End Get
        Set(ByVal value As Decimal)
            mExchageRateGainLoss = value
        End Set
    End Property

    Public Property Difference() As Decimal
        Get
            Return Me.mDifference
        End Get
        Set(ByVal value As Decimal)
            mDifference = value
        End Set
    End Property

    Public Property ARPaid_IDR() As Decimal
        Get
            Return Me.mARPaid_IDR
        End Get
        Set(ByVal value As Decimal)
            mARPaid_IDR = value
        End Set
    End Property

    Public Property Paid_IDR() As Decimal
        Get
            Return Me.mPaid_IDR
        End Get
        Set(ByVal value As Decimal)
            mPaid_IDR = value
        End Set
    End Property

    Public Property ExchangeRate() As Decimal
        Get
            Return Me.mExchangeRate
        End Get
        Set(ByVal value As Decimal)
            mExchangeRate = value
        End Set
    End Property

    Public Property Paid() As Decimal
        Get
            Return Me.mPaid
        End Get
        Set(ByVal value As Decimal)
            mPaid = value
        End Set
    End Property

    Public Property PrimaryAkunPayTo() As Double
        Get
            Return Me.mPrimaryAkunPayTo
        End Get
        Set(ByVal value As Double)
            mPrimaryAkunPayTo = value
        End Set
    End Property

    Public Property PrimaryAkunDifference() As Double
        Get
            Return Me.mPrimaryAkunDifference
        End Get
        Set(ByVal value As Double)
            mPrimaryAkunDifference = value
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

    Public Property RefNumber() As String
        Get
            Return mRefNumber
        End Get
        Set(ByVal value As String)
            mRefNumber = value
        End Set
    End Property

    Public Property Tanggal() As Date
        Get
            Return mDate
        End Get
        Set(ByVal value As Date)
            mDate = value
        End Set
    End Property

    Public Property Urutan() As Double
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Double)
            mUrutan = value
        End Set
    End Property

    'Public Property TipeJournal() As EnumTipeJournal
    '    Get
    '        Return mTipeJournal
    '    End Get
    '    Set(ByVal value As EnumTipeJournal)
    '        mTipeJournal = value
    '    End Set
    'End Property

    Public Property PrimaryDokumenAcuan() As Double
        Get
            Return mPrimaryDokumenAcuan
        End Get
        Set(ByVal value As Double)
            mPrimaryDokumenAcuan = value
        End Set
    End Property

    Public Property PrimaryPihakAcuan() As Double
        Get
            Return mPrimaryPihakAcuan
        End Get
        Set(ByVal value As Double)
            mPrimaryPihakAcuan = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
        End Set
    End Property

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journalTableAdapter
        Dim TanggalAwal, TanggalAkhir As Date
        Dim NoUrut As String
        Dim ObjCD As New Core_Module.ClsCekData
        Dim RetVal As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            TanggalAwal = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)
            TanggalAkhir = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            If DTA.ScalarQuery_GetMaxUrutan(TanggalAwal, TanggalAkhir).HasValue Then
                NoUrut = CStr(DTA.ScalarQuery_GetMaxUrutan(TanggalAwal, TanggalAkhir).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            If IsNumeric(NoUrut) Then
                Me.mUrutan = CDbl(NoUrut)
            End If
            While NoUrut.Length < 6
                NoUrut = "0" + NoUrut
            End While

BuatNoBaru:
            RetVal = NoUrut + "/" + "RCR" + "/" + CStr(Format(vTanggalTransaksi, "MM")) + "/" + CStr(Format(vTanggalTransaksi, "yyyy"))

            If ApakahSudahAda(RetVal) = True Then
                NoUrut += 1
                GoTo BuatNoBaru
            End If

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function CariDetail(ByVal vPrimaryMaster As Double, _
       ByVal vDGVItem As DataGridView, ByVal frm As FrmTransaksiRecuring) As Boolean
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journalTableAdapter
        Dim TBL As New DseTransaksiRecurring.tbltrans_recurring_journalDataTable
        Dim row As DseTransaksiRecurring.tbltrans_recurring_journalRow
        Dim ObjItem As New ClsTransaksiRecurring_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fil basis data
            If DTA.Fill(TBL, vPrimaryMaster) > 0 Then
                'salin data header
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = vPrimaryMaster
                    Me.mRefNumber = .Recurring_No
                    Me.mDate = .Recurring_Date
                    Me.mDesc = .Description
                    Me.mIgnoreIncrementOnSkip = .Ignore_Increment_on_skip
                    Me.mFrequency = .Frequency
                    Me.mEvery = .Every
                    Me.mLimit = .Limit
                    Me.mLimitValue = .Limit_Times
                    Me.mAktive = .Active
                    Me.mStartDate = .Start_Date
                    Me.mEndDate = .End_Date
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'salin data item ada 2 item 
                'CodeNo
                ObjItem.IsiDGVItem(vDGVItem, vPrimaryMaster, frm)
                ObjItem.IsiDGVEntries(frm.dgvEntries, vPrimaryMaster, frm)
                ObjItem.IsiDGVReminder(frm.dgvReminder, vPrimaryMaster, frm)
                'nilai balikan
                Return True
            Else
                'Me.ResetObject(vDGVItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
        ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
        Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journalTableAdapter
        Dim DTAItem As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_transaction_entriesTableAdapter
        Dim TBLItem As New DseTransaksiRecurring.tbltrans_recurring_journal_transaction_entriesDataTable
        Dim RowItem As DseTransaksiRecurring.tbltrans_recurring_journal_transaction_entriesRow
        Dim DTAJ As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim CreatedBy As String = ""
        Dim ObjClosing As New ClsTransaksiClosing
        Dim PrimaryLastJurnal As Double = -1
        Dim LastRevaluationDate, DummyDate As Date
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA_Closing As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim NoClosing As String = ""

        Dim ObjJournal As New ClsTransaksiJournal

        Dim TipeJurnal1 As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
        Dim TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Closing.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            '0. pengecekan jika ada transaksi status = Locked
            If DTAItem.ScalarQuery_GetJmlByStatus(vPrimaryTransaksi, ClsTransaksiRecurring.EnumStatus.Created, ClsTransaksiRecurring.EnumStatus.Locked).HasValue Then
                If DTAItem.ScalarQuery_GetJmlByStatus(vPrimaryTransaksi, ClsTransaksiRecurring.EnumStatus.Created, ClsTransaksiRecurring.EnumStatus.Locked).Value > 0 Then
                    If TampilkanPesan = True Then
                        MsgBox("You can't delete this data because this data already exist in journal transaction." _
                               , MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            '1. jika tanggal sebelum tanggal closing maka tidak bisa dihapus
            Dim dtClosing As Date = Date.Now
            dtClosing = ObjClosing.GetClosingDate
            TBLItem.Rows.Clear()
            DTAItem.FillBy_TanggalClosing_N_Status(TBLItem, vPrimaryTransaksi, dtClosing, ClsTransaksiRecurring.EnumStatus.Created)
            If TBLItem.Rows.Count > 0 Then
                Dim strMsg As String = ""
                For Each RowItem In TBLItem.Rows
                    With RowItem
                        If String.IsNullOrEmpty(strMsg) Then
                            strMsg &= .Journal_No
                        Else
                            strMsg &= ", " & .Journal_No
                        End If
                    End With
                Next
                If TampilkanPesan = True Then
                    MsgBox("You can't delete this data because journal no: " & strMsg & _
                        " have transaction date <= closing date", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '2. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If CreatedBy <> vSalesStaff Then
                If TampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            If TBLItem.Rows.Count > 0 Then
                Dim strMsg As String = ""
                For Each RowItem In TBLItem.Rows
                    With RowItem
                        If .Primary_Journal > -1 Then
                            '4. jika tipenya adalah end of month journal, maka hanya bisa diedit untuk transaksi yang terakhir saja
                            If DTAJ.ScalarQuery_GetTipeJurnal2(.Primary_Journal).HasValue = True Then
                                TipeJurnal2 = DTAJ.ScalarQuery_GetTipeJurnal2(.Primary_Journal).Value

                                If TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation Then
                                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal

                                    If DTAJ.ScalarQuery_GetPrimaryMaxByTipeJurnal2(TipeJurnal1, TipeJurnal2).HasValue = True Then
                                        PrimaryLastJurnal = DTAJ.ScalarQuery_GetPrimaryMaxByTipeJurnal2(TipeJurnal1, TipeJurnal2).Value
                                        If PrimaryLastJurnal <> vPrimaryTransaksi Then
                                            If strMsg = "" Then
                                                strMsg &= .Journal_No
                                            Else
                                                strMsg &= ", " & .Journal_No
                                            End If
                                            'MsgBox("You cannot delete this data, because it's not the last period of end of month journal.", _
                                            'MsgBoxStyle.Information, "Information")
                                            ' Return False
                                            'Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End With
                Next

                If strMsg <> "" Then
                    MsgBox("You can't delete this data, because journal no : " & strMsg & " not the last period of end of month journal.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '5. jika membuat jurnal umum, selain jurnal revaluasi, tidak boleh membuat jurnal yg memiliki mata uang non IDR
            'pada periode yang sudah direvaluasi
            If TBLItem.Rows.Count > 0 Then
                Dim strMsg As String = ""
                Dim strMsg2 As String = ""
                For Each RowItem In TBLItem.Rows
                    With RowItem
                        TipeJurnal1 = ObjJournal.GetTipeJournal(.Primary_Journal)
                        LastRevaluationDate = ObjCD.GetLasDateOfTheMonth(ObjJournal.GetTglRevaluasiSelanjutnya - TimeSpan.FromDays(33))
                        DummyDate = ObjCD.GetFirstDateofTheMonth(ObjJournal.GetTanggalJournal(.Primary_Journal))
                        If (TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal And TipeJurnal2 <> ClsTransaksiJournal.EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation) Or _
                        TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Or TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                            If LastRevaluationDate >= DummyDate Then
                                If ObjJournal.ApakaJurnalNonIDR(.Primary_Journal) = True Then
                                    If strMsg = "" Then
                                        strMsg &= "Journal No: " & .Journal_No & "Period of " & DummyDate
                                    Else
                                        strMsg &= vbNewLine & ", Journal No: " & .Journal_No & "Period of " & DummyDate
                                    End If
                                End If
                            End If
                        End If
                        If TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Or TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                            If LastRevaluationDate >= DummyDate Then
                                If ObjJournal.ApakahAdaPaymentNonIDR(vPrimaryTransaksi) = True Then
                                    If strMsg2 = "" Then
                                        strMsg2 &= "Journal No: " & .Journal_No & "Period of " & DummyDate
                                    Else
                                        strMsg2 &= vbNewLine & ", Journal No: " & .Journal_No & "Period of " & DummyDate
                                    End If
                                End If
                            End If
                        End If
                    End With
                Next
                If strMsg <> "" Then
                    MsgBox(strMsg & " has been revaluate, you cannot edit this data", _
                           MsgBoxStyle.Information, "Information")
                    Return False
                End If

                If strMsg2 <> "" Then
                    MsgBox(strMsg2 & " has been revaluate, you cannot edit this data", _
                           MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If







            '3. cek ke transaksi custom
            'Dim ObjCust As New CustomSMESS.ClsTransaksiJournal
            'If ObjCust.ApakahBolehHapusCustom(vPrimaryTransaksi) = False Then
            '    Return False
            'End If
            'ObjCust = Nothing


            'jika jurnal tersebut dirujuk oleh purchase closing
            TBLItem.Rows.Clear()
            If DTAItem.FillBy(TBLItem, vPrimaryTransaksi) > 0 Then
                For Each RowItem In TBLItem.Rows
                    With RowItem
                        If .Primary_Journal > -1 Then
                            If String.IsNullOrEmpty(DTA_Closing.ScalarQuery_GetNoTransaksi(.Primary_Journal)) = False Then
                                If NoClosing = "" Then
                                    NoClosing &= "Journal No " & .Journal_No & " Closing " & DTA_Closing.ScalarQuery_GetNoTransaksi(.Primary_Journal)
                                Else
                                    NoClosing &= ", Journal No " & .Journal_No & " Closing " & DTA_Closing.ScalarQuery_GetNoTransaksi(.Primary_Journal)
                                End If
                            End If
                        End If
                    End With
                Next
                If NoClosing <> "" Then
                    MsgBox("This journal referred by inventory closing: " & vbNewLine & NoClosing & vbNewLine & _
                           ". You can't delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If


            'nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjClosing = Nothing
            ObjCD = Nothing
            DTA_Closing = Nothing
            DTAItem = Nothing
            ObjJournal = Nothing
            DTAJ.Dispose()
        End Try
    End Function


    Public Function Hapus(ByVal vPrimaryTransaksi As Double, _
       ByVal StrictSalesStaffDelete As Boolean, ByVal usercode As String, ByVal frm As FrmTransaksiRecuring) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_General_Hapus, usercode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, StrictSalesStaffDelete, usercode, frm)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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
            ObjAntri = Nothing
            Me.DTAInOUT.Dispose()
        End Try
    End Function

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As Boolean, ByVal usercode As String, ByVal frm As FrmTransaksiRecuring) As Boolean
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, usercode) = True Then
                'hapus data header, data item otomatis dihapuskan oleh server
                'frm.tms_DeleteAll_Click(Nothing, Nothing)
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByReffNo(vNoTransaksi).HasValue Then
                If DTA.ScalarQuery_GetJmlByReffNo(vNoTransaksi).Value > 0 Then
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

    Public Overridable Function ApakahDataBenar1(ByVal vBaru As Boolean) As Boolean
        Dim LastRevaluationDate As Date = Today
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjClosing As New ClsTransaksiClosing

        Try

            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mRefNumber) = True Then
                    Me.mRefNumber = Me.GetNoBaru(Me.mDate)
                    MsgBox("Duplication in recurring no:'" + Me.mRefNumber + "', please change your number ", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                'no po tidak boleh duplikat
                Dim NoLama As String = Me.GetNoTransasi(Me.mPrimary)
                If Me.RefNumber <> NoLama Then
                    If Me.ApakahSudahAda(Me.mRefNumber) = True Then
                        Me.mRefNumber = Me.GetNoBaru(Me.mDate)
                        MsgBox("Duplication in Recurring No '" + Me.mRefNumber + "', please change your number", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            ''Cek DGVItem
            'If Me.dgvItem.Rows.Count < 2 Then
            '    MsgBox("Minimum item = 2, please input item account entries again", MsgBoxStyle.Information, "Information")
            '    Me.dgvItem.Focus()
            '    Return False
            'End If

            ''Cek DgvEtries
            'If Me.dgvEntries.Rows.Count < 1 Then
            '    MsgBox("Minimum item = 1, please input item transaction entries again", MsgBoxStyle.Information, "Information")
            '    Me.dgvEntries.Focus()
            '    Return False
            'End If

            ''Cek DgvReminder
            'If Me.dgvReminder.Rows.Count < 1 Then
            '    MsgBox("Minimum item = 1, please input item reminder entries again", MsgBoxStyle.Information, "Information")
            '    Me.dgvReminder.Focus()
            '    Return False
            'End If

            '' Cek debet kredit harus balance
            'Me.HitungUnbalance()
            'If CDec(Me.TxtUnbalance.Text) <> 0 Then
            '    MsgBox("This is an unbalance account. Please add numbers in debet or credit column so that a balanced calculation", _
            '    MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            'Dim Unbalance As Decimal = 0
            'Unbalance = Me.HitungUnbalanceIncrement
            'If Unbalance <> 0 Then
            '    MsgBox("This is an unbalance increment value. Please add numbers in increment value of debet or credit column so that a balanced calculation", _
            '             MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            ''tidak boleh bikin transaksi di belakang convertion date
            'If Me.ObjClosing.GetConvertionDate.Date > Me.dtpStartDate.Value.Date Then
            '    MsgBox("Minimum transaction start date is " + Format(Me.ObjClosing.GetConvertionDate.Date, "dd MMM yyyy") + ".", _
            '    MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            ''cek apakah ada account number yang debet dan kreditnya bernilai 0
            'For Each row In Me.dgvItem.Rows
            '    If Me.dgvItem.Item(Me.ColDebet.Index, row.Index).Value = 0 And _
            '    Me.dgvItem.Item(Me.ColKredit.Index, row.Index).Value = 0 Then
            '        MsgBox("Please add positive number in account no. " + Me.dgvItem.Item(Me.ColAccount.Index, row.Index).Value, _
            '        MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If

            '    If Me.dgvItem.Item(Me.ColAcc_IncrementStyle.Index, row.Index).Value <> _
            '       Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.No_Increment) AndAlso _
            '       Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, row.Index).Value = 0 Then
            '        MsgBox("Please add positive number in increment value if you choose increment style as " & _
            '        Me.dgvItem.Item(Me.ColAcc_IncrementStyle.Index, row.Index).Value, _
            '        MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'Next

            'If Me.dgvReminder.Rows.Count > 0 Then
            '    For Each row In Me.dgvReminder.Rows
            '        If row.Index <> Me.dgvReminder.Rows.Count Then
            '            If String.IsNullOrEmpty(Me.dgvReminder.Item(Me.ColReminder_UserCode.Index, row.Index).Value) Then
            '                MsgBox("User code is null please input item reminder entries again", MsgBoxStyle.Information, "Information")
            '                Me.dgvReminder.Focus()
            '                Return False
            '                Exit For
            '            End If

            '            'If String.IsNullOrEmpty(Me.dgvReminder.Item(Me.ColReminder_Remarks.Index, row.Index).Value) Then
            '            '    MsgBox("remarks of reminder is null please input item reminder entries again", MsgBoxStyle.Information, "Information")
            '            '    Me.dgvReminder.Focus()
            '            '    Return False
            '            '    Exit For
            '            'End If
            '        End If
            '    Next
            'End If

            '''cek data yang salah, double dalam kode
            ''' If Me.Baru = True Then
            ''Dim strJournal As String = ""
            ''For Each row In Me.dgvEntries.Rows
            ''    With row
            ''        If .Cells(Me.ColTrans_Status.Index).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
            ''            If .Cells(Me.ColTrans_PrimaryJournal.Index).Value = -1 Then
            ''                If Me.ObjJournal.ApakahSudahAda(.Cells(Me.ColTrans_JournalNumber.Index).Value) = True Then
            ''                    If strJournal = "" Then
            ''                        strJournal &= .Cells(Me.ColTrans_JournalNumber.Index).Value
            ''                    Else
            ''                        strJournal &= ", " & .Cells(Me.ColTrans_JournalNumber.Index).Value
            ''                    End If
            ''                End If
            ''            Else
            ''                Dim CodeLama As String = Me.ObjJournal.GetRefCode(.Cells(Me.ColTrans_PrimaryJournal.Index).Value)
            ''                If CodeLama <> .Cells(Me.ColTrans_JournalNumber.Index).Value Then
            ''                    If Me.ObjJournal.ApakahSudahAda(.Cells(Me.ColTrans_JournalNumber.Index).Value) = True Then
            ''                        If strJournal = "" Then
            ''                            strJournal &= .Cells(Me.ColTrans_JournalNumber.Index).Value
            ''                        Else
            ''                            strJournal &= ", " & .Cells(Me.ColTrans_JournalNumber.Index).Value
            ''                        End If
            ''                    End If
            ''                End If
            ''            End If
            ''        End If
            ''    End With
            ''Next
            ''If strJournal = "" Then
            ''    MsgBox("Duplication in no journal = " & strJournal & ", please change your number ", _
            ''    MsgBoxStyle.Information, "Information")
            ''    Me.dgvEntries.Focus()
            ''    Return False
            ''End If

            ''cek data yang salah, double dalam kode
            'If Me.Baru = True Then
            '    If Me.ObjRecurring.ApakahSudahAda(Me.TxtRefNo.Text) = True Then
            '        MsgBox("Duplication in Recurring No =" + Me.TxtRefNo.Text + ", please change your number", _
            '        MsgBoxStyle.Information, "Information")
            '        Me.TxtRefNo.Focus()
            '        Return False
            '    End If
            'Else
            '    Dim TipeCodeLama As String = Me.ObjRecurring.GetNoTransasi(Me.TxtPrimaryJournal.Text)
            '    If TipeCodeLama <> Me.TxtRefNo.Text Then
            '        If Me.ObjRecurring.ApakahSudahAda(Me.TxtRefNo.Text) = True Then
            '            MsgBox("Duplication in Recurring No =" + Me.TxtRefNo.Text + ", please change your number", _
            '            MsgBoxStyle.Information, "Information")
            '            Me.TxtRefNo.Focus()
            '            Return False
            '        End If
            '    End If

            '    'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
            '    'proses edit
            '    Dim OldTimeStamp As Date = Me.ObjRecurring.GetTimestamp(Me.TxtPrimaryJournal.Text)
            '    If Me.ObjRecurring.Timestamp <> OldTimeStamp Then
            '        If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
            '        MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
            '            Me.ObjRecurring.CariDetail(Me.ObjRecurring.Primary, Me.dgvItem, Me)
            '            Me.SalinkeField()
            '        End If
            '        Return False
            '    End If
            'End If

            ''pastikan semua kode akun ada di dalam master akun
            'For Each row In Me.dgvItem.Rows
            '    If IsNothing(row.Cells(Me.ColPrimaryAkun.Index).Value) = True Then
            '        row.Selected = True
            '        MsgBox("Invalid account in row " + row.Index.ToString, MsgBoxStyle.Information, "Information")
            '        Me.dgvItem.ClearSelection()
            '        Return False
            '    Else
            '        If ObjAkun.ApakahSudahAda_Aktif_NonHeader2(row.Cells(Me.ColPrimaryAkun.Index).Value) = False Then
            '            MsgBox("Invalid account in row " + row.Index.ToString, MsgBoxStyle.Information, "Information")
            '            Me.dgvItem.ClearSelection()
            '            row.Selected = True
            '            Return False
            '        End If
            '    End If
            'Next

            '''jika sudah ada closing jurnal untuk tahun ybs maka tidak bisa input data baru
            ''If Me.ObjJurnal.GetJML(ClsTransaksiJournal.EnumTipeJournal.ClosingJournal, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpDate.Value), _
            ''Me.ObjCD.GetLastDateOfTheYear(Me.dtpDate.Value.Date)) > 0 Then
            ''    MsgBox("There are closing jurnal for this year you cannot save this data", MsgBoxStyle.Information, "Information")
            ''    Return False
            ''End If

            ''kalo sampe sini return true
            'Return True


            ''1. jika tanggal Journal sebelum tanggal closing maka tidak bisa diedit
            'If vDGVEntries.RowCount > 0 Then
            '    Dim strMsg As String = ""
            '    For Each rowDGVEntries In vDGVEntries.Rows
            '        With rowDGVEntries
            '            If CDate(.Cells(vIndexJournalDate).Value).Date <= ObjClosing.GetClosingDate.Date Then
            '                If strMsg = "" Then
            '                    strMsg &= .Cells(vIndexJournalNo).Value.ToString
            '                Else
            '                    strMsg &= ", " & .Cells(vIndexJournalNo).Value.ToString
            '                End If
            '            End If
            '        End With
            '    Next
            '    If strMsg <> "" Then
            '        MsgBox("You cannot save this data because journal no : " & vbNewLine & strMsg & vbNewLine & _
            '               " have transaction date <= closing date" , MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If


            ''2. jika membuat jurnal umum, selain jurnal revaluasi, tidak boleh membuat jurnal yg memiliki mata uang non IDR
            ''pada periode yang sudah direvaluasi
            'LastRevaluationDate = ObjCD.GetLasDateOfTheMonth(Me.GetTglRevaluasiSelanjutnya - TimeSpan.FromDays(33))
            'DummyDate = ObjCD.GetFirstDateofTheMonth(Me.Tanggal)
            ''If (Me.TipeJournal = EnumTipeJournal.GeneralJournal And Me.TipeJurnal2 <> EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation) Or _
            ''Me.TipeJournal = EnumTipeJournal.PurchasePaymentJournal Or Me.TipeJournal = EnumTipeJournal.SalesReceivementJournal Then
            'If LastRevaluationDate >= DummyDate Then
            '    For Each rowDGV In vDGVItem.Rows
            '        If rowDGV.Cells(vIndexPrimaryCurr).Value <> Core_Module.PrimaryBaseCurrency Then
            '            MsgBox("The period of " + Format(DummyDate, "MMMMM yyyy") + " has been revaluate, you cannot save this data", _
            '            MsgBoxStyle.Information, "Information")
            '            Return False
            '        End If
            '    Next
            'End If
            ''End If
            'If Me.TipeJournal = EnumTipeJournal.SalesReceivementJournal Or Me.TipeJournal = EnumTipeJournal.PurchasePaymentJournal Then
            '    If LastRevaluationDate >= DummyDate Then
            '        For Each rowDGV In vDGVBalance.Rows
            '            If rowDGV.Cells(vIndexPrimaryCurrencyBalance).Value <> Core_Module.PrimaryBaseCurrency And _
            '            CDec(rowDGV.Cells(vIndexPayment).Value) <> 0 Then
            '                MsgBox("The period of " + Format(DummyDate, "MMMMM yyyy") + " has been revaluate, you cannot save this data", _
            '                MsgBoxStyle.Information, "Information")
            '                Return False
            '            End If
            '        Next
            '    End If
            'End If

            'return value
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            ObjCD = Nothing
            ObjClosing = Nothing
        End Try
    End Function

    Public Function GetNoTransasi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoTransaksi(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetNoTransaksi(vPrimary)
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


    Public Function SimpanBaru(ByVal vDGVItem As DataGridView, _
          ByVal vUserCode As String, ByVal Frm As FrmTransaksiRecuring) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier.ToString, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_General_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vDGVItem, vUserCode, Frm)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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
            ' Me.DTAInOUT.Dispose()
        End Try
    End Function

    Public Function SimpanBaru1(ByVal vDGVItem As DataGridView, _
           ByVal vUserCode As String, ByVal Frm As FrmTransaksiRecuring) As Boolean
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journalTableAdapter
        Dim ObjItem As New ClsTransaksiRecurring_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.mUrutan = CDbl(Me.mRefNumber.Substring(0, 6))
            'Me.mUrutan = -1
            'Me.mPrimaryDokumenAcuan = -1
            'Me.mPrimaryPihakAcuan = -1
            If Me.ApakahDataBenar1(True) = True Then

                If DTA.Insert(Me.mRefNumber, _
                              Me.mDate, Me.mDesc, _
                              Me.mIgnoreIncrementOnSkip, Me.mFrequency, _
                              Me.mEvery, Me.mLimit, Me.mLimitValue, _
                              Me.mAktive, Me.mStartDate, Me.mEndDate, _
                              vUserCode, Now, vUserCode, Now, Now, Me.mUrutan) > 0 Then

                    'update data yang berubah setelah insert query
                    Me.mPrimary = Me.GetPrimary(Me.RefNumber)
                    Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanDGVItem(Frm.dgvItem, Me.mPrimary, Frm)
                    ObjItem.SimpanDGVEntries(Frm.dgvEntries, Me.mPrimary, Frm)
                    ObjItem.SimpanDGVReminders(Frm.dgvReminder, Me.mPrimary, Frm)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mRefNumber + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru1(vDGVItem, vUserCode, Frm)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vDGVItem As DataGridView, ByVal vUserCode As String, ByVal frm As FrmTransaksiRecuring) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_General_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vDGVItem, vUserCode, frm)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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
            ObjAntri = Nothing
            SW = Nothing
            Me.DTAInOUT.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1(ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, ByVal frm As FrmTransaksiRecuring) As Boolean
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journalTableAdapter
        Dim ObjItem As New ClsTransaksiRecurring_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'Me.mTipeJournal = TipeJournal
            Me.mUrutan = CDbl(Me.mRefNumber.Substring(0, 6))
            'Me.mPrimaryDokumenAcuan = -1
            'Me.mPrimaryPihakAcuan = -1
            If Me.ApakahDataBenar1(False) = True Then
                If DTA.Update(Me.mRefNumber, Me.mDate, Me.mDesc, _
                              Me.mIgnoreIncrementOnSkip, Me.mFrequency, _
                              Me.mEvery, Me.mLimit, Me.mLimitValue, _
                              Me.mAktive, Me.mStartDate, Me.mEndDate, _
                              vUserCode, Date.Now, Date.Now, Me.mUrutan, Me.mPrimary) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)

                    'simpan data item baru
                    ObjItem.SimpanDGVItem(frm.dgvItem, Me.mPrimary, frm)
                    ObjItem.SimpanDGVEntries(frm.dgvEntries, Me.mPrimary, frm)
                    ObjItem.SimpanDGVReminders(frm.dgvReminder, Me.mPrimary, frm)

                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mRefNumber + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanPerubahan1(vDGVItem, vUserCode, frm)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journalTableAdapter
        Dim Primary As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vNoTransaksi).HasValue AndAlso DTA.ScalarQuery_GetPrimary(vNoTransaksi).Value > -1 Then
                Primary = DTA.ScalarQuery_GetPrimary(vNoTransaksi).Value
            End If
            Return Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Public Function GetPrimaryTransaksiByPrimaryJournal(ByVal vPrimaryJournal As Double) As Double
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_transaction_entriesTableAdapter
        Dim Primary As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryTransaksiBy(vPrimaryJournal).HasValue AndAlso DTA.ScalarQuery_GetPrimaryTransaksiBy(vPrimaryJournal).Value > -1 Then
                Primary = DTA.ScalarQuery_GetPrimaryTransaksiBy(vPrimaryJournal).Value
            End If
            Return Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
    ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journalTableAdapter
        Dim ObjClosing As New ClsTransaksiClosing
        'Dim TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation
        'Dim TipeJurnal1 As ClsTransaksiJournal.EnumTipeJournal = EnumTipeJournal.GeneralJournal
        Dim PrimaryLastJurnal As Double = -1
        Dim LastRevaluationDate, DummyDate As Date
        Dim ObjCD As New Core_Module.ClsCekData
        Dim NoClosing As String = ""
        Dim DTA_Closing As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter

        Try
            ''1. jika tanggal sebelum tanggal closing maka tidak bisa diedit 
            'DTA_Closing.Connection.ConnectionString = Core_Module.GetConstring
            'If Me.GetTanggalJournal(vPrimaryTransaksi).Date <= ObjClosing.GetClosingDate.Date Then
            '    If TampilkanPesan = True Then
            '        MsgBox("You cannot edit this data because transaction date <= closing date" _
            '        , MsgBoxStyle.Information, "Information")
            '    End If
            '    Return False
            'End If

            '2. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If CreatedBy <> vSalesStaff Then
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            '3. cek ke transaksi custom
            'Dim ObjCust As New CustomSMESS.ClsTransaksiJournal
            'If ObjCust.ApakahBolehEditCustom(vPrimaryTransaksi) = False Then
            '    Return False
            'End If
            'ObjCust = Nothing

            ''4. jika tipenya adalah end of month journal, maka hanya bisa diedit untuk transaksi yang terakhir saja
            'If DTA.ScalarQuery_GetTipeJurnal2(vPrimaryTransaksi).HasValue = True Then
            '    TipeJurnal2 = DTA.ScalarQuery_GetTipeJurnal2(vPrimaryTransaksi).Value
            '    If TipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation Then
            '        TipeJurnal1 = EnumTipeJournal.GeneralJournal

            '        If DTA.ScalarQuery_GetPrimaryMaxByTipeJurnal2(TipeJurnal1, TipeJurnal2).HasValue = True Then
            '            PrimaryLastJurnal = DTA.ScalarQuery_GetPrimaryMaxByTipeJurnal2(TipeJurnal1, TipeJurnal2).Value
            '            If PrimaryLastJurnal <> vPrimaryTransaksi Then
            '                MsgBox("You cannot edit this data, because it's not the last period of end of month journal.", _
            '                MsgBoxStyle.Information, "Information")
            '                Return False
            '            End If
            '        End If
            '    End If

            '    '5. jika membuat jurnal umum, selain jurnal revaluasi, tidak boleh membuat jurnal yg memiliki mata uang non IDR
            '    'pada periode yang sudah direvaluasi
            '    TipeJurnal1 = Me.GetTipeJournal(vPrimaryTransaksi)
            '    LastRevaluationDate = ObjCD.GetLasDateOfTheMonth(Me.GetTglRevaluasiSelanjutnya - TimeSpan.FromDays(33))
            '    DummyDate = ObjCD.GetFirstDateofTheMonth(Me.GetTanggalJournal(vPrimaryTransaksi))
            '    If (TipeJurnal1 = EnumTipeJournal.GeneralJournal And TipeJurnal2 <> EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation) Or _
            '    TipeJurnal1 = EnumTipeJournal.PurchasePaymentJournal Or TipeJurnal1 = EnumTipeJournal.SalesReceivementJournal Then
            '        If LastRevaluationDate >= DummyDate Then
            '            If Me.ApakaJurnalNonIDR(vPrimaryTransaksi) = True Then
            '                MsgBox("The period of " + Format(DummyDate, "MMMMM yyyy") + " has been revaluate, you cannot edit this data", _
            '                    MsgBoxStyle.Information, "Information")
            '                Return False
            '            End If
            '        End If
            '    End If
            '    If TipeJurnal1 = EnumTipeJournal.SalesReceivementJournal Or TipeJurnal1 = EnumTipeJournal.PurchasePaymentJournal Then
            '        If LastRevaluationDate >= DummyDate Then
            '            If Me.ApakahAdaPaymentNonIDR(vPrimaryTransaksi) = True Then
            '                MsgBox("The period of " + Format(DummyDate, "MMMMM yyyy") + " has been revaluate, you cannot edit this data", _
            '                    MsgBoxStyle.Information, "Information")
            '                Return False
            '            End If
            '        End If
            '    End If
            'End If

            'If String.IsNullOrEmpty(DTA_Closing.ScalarQuery_GetNoTransaksi(vPrimaryTransaksi)) = False Then
            '    NoClosing = DTA_Closing.ScalarQuery_GetNoTransaksi(vPrimaryTransaksi)
            'Else
            '    NoClosing = ""
            'End If
            'If NoClosing <> "" Then
            '    MsgBox("This journal referred by inventory closing: " + NoClosing + ". You cannot delete this data.", MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjClosing = Nothing
            ObjCD = Nothing
            DTA_Closing.Dispose()
        End Try
    End Function


    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiRecurring.tblCariDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date) As Boolean
        Dim DTA As New DseTransaksiRecurringTableAdapters.tblCariTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            vTanggalDari = Format(CDate(vTanggalDari), "yyyy-MM-dd")
            vTanggalSampai = Format(CDate(vTanggalSampai), "yyyy-MM-dd")
            If SearchCriteria = EnumSearchCriteria.ByReferenceCode Then
                QueryResult = DTA.FillBy_RecurringNoLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescriptionLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDate Then
                QueryResult = DTA.FillBy_PeriodeRecurring(TBL, vTanggalDari, vTanggalSampai)
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Recurring No"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Recurring date"
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(3).HeaderText = "Description"
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJmlReminderAktif(ByVal vTglAcuan As Date, ByVal vPrimaryUser As Double) As Double
        Dim DTA As New DseTransaksiRecurringTableAdapters.RecurringActiveTableAdapter
        Dim TBL As New DseTransaksiRecurring.RecurringActiveDataTable
        Dim row As DseTransaksiRecurring.RecurringActiveRow
        Dim StatusAcuan As EnumStatus = EnumStatus.Planed
        Dim RetVal As Double
        Dim IntStart As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryUser, CInt(StatusAcuan))
            'buang yang tida sesuai tanggal
            For ij As Integer = TBL.Rows.Count - 1 To 0 Step -1
                row = TBL.Rows(ij)
                With row
                    IntStart = 0
                    If IsNumeric(.Start_Alarm_days_before) Then
                        IntStart = CInt(.Start_Alarm_days_before)
                    End If
                    If CDate(vTglAcuan.AddDays(IntStart)).Date < .Journal_Date.Date Then
                        TBL.Rows.RemoveAt(ij)
                    End If
                End With
            Next

            RetVal = TBL.Rows.Count
            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vStartDate As Date, ByVal vEndDate As Date) As Double
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journalTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlBy_StartEndDate(vEndDate, vStartDate).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlBy_StartEndDate(vEndDate, vStartDate).Value
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ModGlobalVariable.EnumSearchCriteria, _
    ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date) As Double
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journalTableAdapter

        Try
            If vSearchBy = EnumSearchCriteria.ByReferenceCode Then
                Return DTA.ScalarQuery_GetJmlByReffNoLike(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                Return DTA.ScalarQuery_GetJmlByDescLike(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByDate Then
                Return DTA.ScalarQuery_GetJmlBy_StartEndDate(vTanggalSampai, vTanggalDari).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
