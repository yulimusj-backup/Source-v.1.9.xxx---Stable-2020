Option Explicit On
Option Strict Off

Public Class ClsTransaksiJournal
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mRefNumber As String = ""
    Private mDate As Date = Now
    Private mUrutan As Double = -1
    Private mTipeJournal As EnumTipeJournal = EnumTipeJournal.AdjustmenJournal
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
    Private mTipeJurnal2 As EnumTipeJurnal2 = EnumTipeJurnal2.Jurnal_NoGrouping
    Private mCreatedBy As String = ""
    Private mCreatedDate As Date = Now
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date = Now
    Private mTimestamp As Date = Now
    Private mTipeInOut As EnumTipeInOut = EnumTipeInOut._In
    Private mPrimaryCurr As Double = -1
    Private mReferenceParty As String = ""
    Private mPaymentMethod As String = ""
    Private mPrimaryExChangerateGainLoss As Double = -1

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    'Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian

    Public Enum EnumTipeInOut As Integer
        _In
        out
    End Enum

    Public Enum EnumTipeJurnal2 As Integer
        Jurnal_NoGrouping
        GeneralJurnal_MonthlyRevaluation
        Cash_n_Bank
        AssetAcquistion_Other
        AssetRevaluation_Other
        AssetDepreciation
        AssetDisposal
        AssetSetFinish
        HR_Payroll_Payment
        HR_Payroll_Loan
        HR_Payroll_Loan_Payment
        GeneralJurnal_Recurring
        Credit_Note
        Debit_Note
    End Enum

    'ini juga ada di clstransaksi user
    Public Enum EnumJournalAccesType As Integer
        FullAccess
        Selected_Users_Only
        Respected_User_Only
    End Enum

    Public Property TipeInOut() As EnumTipeInOut
        Get
            Return Me.mTipeInOut
        End Get
        Set(ByVal value As EnumTipeInOut)
            mTipeInOut = value
        End Set
    End Property

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

    Public Property PaymentMethod() As String
        Get
            Return Me.mPaymentMethod
        End Get
        Set(ByVal value As String)
            mPaymentMethod = value
        End Set
    End Property
    Public Property PrimaryExChangeRateGainLoss() As Double
        Get
            Return Me.mPrimaryExChangerateGainLoss
        End Get
        Set(ByVal value As Double)
            mPrimaryExChangerateGainLoss = value
        End Set
    End Property
    Public Property TipeJurnal2() As EnumTipeJurnal2
        Get
            Return Me.mTipeJurnal2
        End Get
        Set(ByVal value As EnumTipeJurnal2)
            mTipeJurnal2 = value
        End Set
    End Property

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

    Public Property TipeJournal() As EnumTipeJournal
        Get
            Return mTipeJournal
        End Get
        Set(ByVal value As EnumTipeJournal)
            mTipeJournal = value
        End Set
    End Property

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

    Public Enum EnumTipeJournal As Integer
        GeneralJournal
        AdjustmenJournal
        ClosingJournal
        SalesJournal
        SalesReturnJournal
        SalesDownPayment
        PurchaseJournal
        PurchaseReturnJournal
        PurchaseDownPaymentJournal
        SalesReceivementJournal
        SalesReturnPaymentJournal
        PurchasePaymentJournal
        PurchaseReturnReceivement
    End Enum

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

    Public Function GetTanggalJournal(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetTanggalJournal(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTanggalJournal(vPrimary).Value
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
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim ObjClosing As New ClsTransaksiClosing
        Dim TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation
        Dim TipeJurnal1 As ClsTransaksiJournal.EnumTipeJournal = EnumTipeJournal.GeneralJournal
        Dim PrimaryLastJurnal As Double = -1
        Dim LastRevaluationDate, DummyDate As Date
        Dim ObjCD As New Core_Module.ClsCekData
        Dim NoClosing As String = ""
        Dim DTA_Closing As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim ObjUser As New Core_Module.ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Closing.Connection.ConnectionString = DTA.Connection.ConnectionString

            '1. jika tanggal sebelum tanggal closing maka tidak bisa diedit 
            DTA_Closing.Connection.ConnectionString = Core_Module.GetConstring
            If Me.GetTanggalJournal(vPrimaryTransaksi).Date <= ObjClosing.GetClosingDate.Date Then
                If TampilkanPesan = True Then
                    MsgBox("You cannot edit this data because transaction date <= closing date" _
                    , MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '2. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            TipeJurnal1 = Me.GetTipeJournal(vPrimaryTransaksi)
            TipeJurnal2 = DTA.ScalarQuery_GetTipeJurnal2(vPrimaryTransaksi).Value

            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If TipeJurnal1 = EnumTipeJournal.GeneralJournal And TipeJurnal2 = EnumTipeJurnal2.Jurnal_NoGrouping Then
                    If CreatedBy <> vSalesStaff Then
                        If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiGeneralJournal_EditUserLain") = False Then
                            If TampilkanPesan = True Then
                                MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                ElseIf TipeJurnal1 = EnumTipeJournal.GeneralJournal And TipeJurnal2 = EnumTipeJurnal2.Cash_n_Bank Then
                    If CreatedBy <> vSalesStaff Then
                        If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiJournal_CashBank_EditUserLain") = False Then
                            If TampilkanPesan = True Then
                                MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                ElseIf TipeJurnal1 = EnumTipeJournal.GeneralJournal And TipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation Then
                    If CreatedBy <> vSalesStaff Then
                        If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiJournal_EndOfMonth_EditUserLain") = False Then
                            If TampilkanPesan = True Then
                                MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                ElseIf TipeJurnal1 = EnumTipeJournal.ClosingJournal Then
                    If CreatedBy <> vSalesStaff Then
                        If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiJournal_Closing_EditUserLain") = False Then
                            If TampilkanPesan = True Then
                                MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                ElseIf TipeJurnal1 = EnumTipeJournal.SalesReceivementJournal Then
                    If CreatedBy <> vSalesStaff Then
                        If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiSalesCycle1_EditUserLain") = False Then
                            If TampilkanPesan = True Then
                                MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                ElseIf TipeJurnal1 = EnumTipeJournal.PurchasePaymentJournal Then
                    If CreatedBy <> vSalesStaff Then
                        If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiJournal_Purchase1_EditUserLain") = False Then
                            If TampilkanPesan = True Then
                                MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                Else
                    If CreatedBy <> vSalesStaff Then
                        If TampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            '3. cek ke transaksi custom
            Dim ObjCust As New CustomSMESS.ClsTransaksiJournal
            If ObjCust.ApakahBolehEditCustom(vPrimaryTransaksi) = False Then
                Return False
            End If
            ObjCust = Nothing

            '4. jika tipenya adalah end of month journal, maka hanya bisa diedit untuk transaksi yang terakhir saja
            If DTA.ScalarQuery_GetTipeJurnal2(vPrimaryTransaksi).HasValue = True Then
                'TipeJurnal2 = DTA.ScalarQuery_GetTipeJurnal2(vPrimaryTransaksi).Value
                If TipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal

                    If DTA.ScalarQuery_GetPrimaryMaxByTipeJurnal2(TipeJurnal1, TipeJurnal2).HasValue = True Then
                        PrimaryLastJurnal = DTA.ScalarQuery_GetPrimaryMaxByTipeJurnal2(TipeJurnal1, TipeJurnal2).Value
                        If PrimaryLastJurnal <> vPrimaryTransaksi Then
                            MsgBox("You cannot edit this data, because it's not the last period of end of month journal.", _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If

                '5. jika membuat jurnal umum, selain jurnal revaluasi, tidak boleh membuat jurnal yg memiliki mata uang non IDR
                'pada periode yang sudah direvaluasi
                'TipeJurnal1 = Me.GetTipeJournal(vPrimaryTransaksi)
                LastRevaluationDate = ObjCD.GetLasDateOfTheMonth(Me.GetTglRevaluasiSelanjutnya - TimeSpan.FromDays(33))
                DummyDate = ObjCD.GetFirstDateofTheMonth(Me.GetTanggalJournal(vPrimaryTransaksi))
                If (TipeJurnal1 = EnumTipeJournal.GeneralJournal And TipeJurnal2 <> EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation) Or _
                TipeJurnal1 = EnumTipeJournal.PurchasePaymentJournal Or TipeJurnal1 = EnumTipeJournal.SalesReceivementJournal Then
                    If LastRevaluationDate >= DummyDate Then
                        If Me.ApakaJurnalNonIDR(vPrimaryTransaksi) = True Then
                            MsgBox("The period of " + Format(DummyDate, "MMMMM yyyy") + " has been revaluate, you cannot edit this data", _
                                MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
                If TipeJurnal1 = EnumTipeJournal.SalesReceivementJournal Or TipeJurnal1 = EnumTipeJournal.PurchasePaymentJournal Then
                    If LastRevaluationDate >= DummyDate Then
                        If Me.ApakahAdaPaymentNonIDR(vPrimaryTransaksi) = True Then
                            MsgBox("The period of " + Format(DummyDate, "MMMMM yyyy") + " has been revaluate, you cannot edit this data", _
                                MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If

            If String.IsNullOrEmpty(DTA_Closing.ScalarQuery_GetNoTransaksi(vPrimaryTransaksi)) = False Then
                NoClosing = DTA_Closing.ScalarQuery_GetNoTransaksi(vPrimaryTransaksi)
            Else
                NoClosing = ""
            End If
            If NoClosing <> "" Then
                MsgBox("This journal referred by inventory closing: " + NoClosing + ". You cannot delete this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

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

    Public Overridable Function ApakahBolehHapus_EndOFMonthJournal(ByVal vPrimaryTransaksi As Double, _
       ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
       Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim vTipeJurnal1 As EnumTipeJournal = EnumTipeJournal.GeneralJournal
        Dim vTipeJurnal2 As EnumTipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation
        Dim PrimaryJurnal As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            '1. jika bukan merupakan jurnal end of month yang terakhir, maka tidak boleh dihapuskan
            If DTA.ScalarQuery_GetPrimaryMaxByTipeJurnal2(vTipeJurnal1, vTipeJurnal2).HasValue = True Then
                PrimaryJurnal = DTA.ScalarQuery_GetPrimaryMaxByTipeJurnal2(vTipeJurnal1, vTipeJurnal2).Value

                If PrimaryJurnal <> vPrimaryTransaksi Then
                    MsgBox("You cannot delete end of month journal on period " + Format(Me.GetTanggalJournal(PrimaryJurnal), _
                    "MMM yyyy"), MsgBoxStyle.Information, "Information")

                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
        ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, ByVal BukaDariRecurring As Boolean, _
        Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim CreatedBy As String = ""
        Dim ObjClosing As New ClsTransaksiClosing
        Dim TipeJurnal1 As ClsTransaksiJournal.EnumTipeJournal = EnumTipeJournal.GeneralJournal
        Dim TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation
        Dim PrimaryLastJurnal As Double = -1
        Dim LastRevaluationDate, DummyDate As Date
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA_Closing As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim NoClosing As String = ""
        Dim objRecurring As New ClsTransaksiRecurring
        Dim ObjUser As New Core_Module.ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Closing.Connection.ConnectionString = DTA.Connection.ConnectionString

            '--------------------------------------------------------------------------------------------------------------
            'erna code ----------------------------------------------------------------------------------------------------
            If objRecurring.GetPrimaryTransaksiByPrimaryJournal(vPrimaryTransaksi) > -1 AndAlso BukaDariRecurring = False Then
                If TampilkanPesan = True Then
                    Dim strNo As String = ""
                    strNo = objRecurring.GetNoTransasi(objRecurring.GetPrimaryTransaksiByPrimaryJournal(vPrimaryTransaksi))
                    MsgBox("This data is recurring journal, you cannot delete this data on this form, " & _
                    "please use form recurring journal to delete this data. Recurring journal No : " & strNo _
                    , MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            '--------------------------------------------------------------------------------------------------------------


            '1. jika tanggal sebelum tanggal closing maka tidak bisa dihapus
            If Me.GetTanggalJournal(vPrimaryTransaksi).Date <= ObjClosing.GetClosingDate.Date Then
                If TampilkanPesan = True Then
                    MsgBox("You cannot delete this data because transaction date <= closing date" _
                    , MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '2. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            TipeJurnal1 = Me.GetTipeJournal(vPrimaryTransaksi)
            TipeJurnal2 = DTA.ScalarQuery_GetTipeJurnal2(vPrimaryTransaksi).Value

            If StrictSalesStaffDelete = True Then
                If TipeJurnal1 = EnumTipeJournal.GeneralJournal And TipeJurnal2 = EnumTipeJurnal2.Jurnal_NoGrouping Then
                    If CreatedBy <> vSalesStaff Then
                        If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiGeneralJournal_HapusUserLain") = False Then
                            If TampilkanPesan = True Then
                                MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                ElseIf TipeJurnal1 = EnumTipeJournal.GeneralJournal And TipeJurnal2 = EnumTipeJurnal2.Cash_n_Bank Then
                    If CreatedBy <> vSalesStaff Then
                        If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiJournal_CashBank_HapusUserLain") = False Then
                            If TampilkanPesan = True Then
                                MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                ElseIf TipeJurnal1 = EnumTipeJournal.GeneralJournal And TipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation Then
                    If CreatedBy <> vSalesStaff Then
                        If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiJournal_EndOfMonth_HapusUserLain") = False Then
                            If TampilkanPesan = True Then
                                MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                ElseIf TipeJurnal1 = EnumTipeJournal.ClosingJournal Then
                    If CreatedBy <> vSalesStaff Then
                        If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiJournal_Closing_HapusUserLain") = False Then
                            If TampilkanPesan = True Then
                                MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                ElseIf TipeJurnal1 = EnumTipeJournal.SalesReceivementJournal Then
                    If CreatedBy <> vSalesStaff Then
                        If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiSalesCycle1_HapusUserLain") = False Then
                            If TampilkanPesan = True Then
                                MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                ElseIf TipeJurnal1 = EnumTipeJournal.PurchasePaymentJournal Then
                    If CreatedBy <> vSalesStaff Then
                        If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiJournal_Purchase1_HapusUserLain") = False Then
                            If TampilkanPesan = True Then
                                MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                Else
                    If CreatedBy <> vSalesStaff Then
                        If TampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            '3. cek ke transaksi custom
            Dim ObjCust As New CustomSMESS.ClsTransaksiJournal
            If ObjCust.ApakahBolehHapusCustom(vPrimaryTransaksi) = False Then
                Return False
            End If
            ObjCust = Nothing

            '4. jika tipenya adalah end of month journal, maka hanya bisa diedit untuk transaksi yang terakhir saja
            If DTA.ScalarQuery_GetTipeJurnal2(vPrimaryTransaksi).HasValue = True Then
                TipeJurnal2 = DTA.ScalarQuery_GetTipeJurnal2(vPrimaryTransaksi).Value
                If TipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal

                    If DTA.ScalarQuery_GetPrimaryMaxByTipeJurnal2(TipeJurnal1, TipeJurnal2).HasValue = True Then
                        PrimaryLastJurnal = DTA.ScalarQuery_GetPrimaryMaxByTipeJurnal2(TipeJurnal1, TipeJurnal2).Value
                        If PrimaryLastJurnal <> vPrimaryTransaksi Then
                            MsgBox("You cannot delete this data, because it's not the last period of end of month journal.", _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If

                '5. jika membuat jurnal umum, selain jurnal revaluasi, tidak boleh membuat jurnal yg memiliki mata uang non IDR
                'pada periode yang sudah direvaluasi
                TipeJurnal1 = Me.GetTipeJournal(vPrimaryTransaksi)
                LastRevaluationDate = ObjCD.GetLasDateOfTheMonth(Me.GetTglRevaluasiSelanjutnya - TimeSpan.FromDays(33))
                DummyDate = ObjCD.GetFirstDateofTheMonth(Me.GetTanggalJournal(vPrimaryTransaksi))
                If (TipeJurnal1 = EnumTipeJournal.GeneralJournal And TipeJurnal2 <> EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation) Or _
                TipeJurnal1 = EnumTipeJournal.PurchasePaymentJournal Or TipeJurnal1 = EnumTipeJournal.SalesReceivementJournal Then
                    If LastRevaluationDate >= DummyDate Then
                        If Me.ApakaJurnalNonIDR(vPrimaryTransaksi) = True Then
                            MsgBox("The period of " + Format(DummyDate, "MMMMM yyyy") + " has been revaluate, you cannot edit this data", _
                                MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
                If TipeJurnal1 = EnumTipeJournal.SalesReceivementJournal Or TipeJurnal1 = EnumTipeJournal.PurchasePaymentJournal Then
                    If LastRevaluationDate >= DummyDate Then
                        If Me.ApakahAdaPaymentNonIDR(vPrimaryTransaksi) = True Then
                            MsgBox("The period of " + Format(DummyDate, "MMMMM yyyy") + " has been revaluate, you cannot edit this data", _
                                MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If

            'jika jurnal tersebut dirujuk oleh purchase closing
            If String.IsNullOrEmpty(DTA_Closing.ScalarQuery_GetNoTransaksi(vPrimaryTransaksi)) = False Then
                NoClosing = DTA_Closing.ScalarQuery_GetNoTransaksi(vPrimaryTransaksi)
            Else
                NoClosing = ""
            End If
            If NoClosing <> "" Then
                MsgBox("This journal referred by inventory closing: " + NoClosing + ". You cannot delete this data.", MsgBoxStyle.Information, "Information")
                Return False
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
        End Try
    End Function

    Public Function GetPrimaryMax(ByVal vTipeJurnal As EnumTipeJournal) As Double
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryMaxByTipeJurnal(vTipeJurnal).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryMaxByTipeJurnal(vTipeJurnal).Value
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

    Public Function GetPrimary(ByVal vRefCode As String) As Double
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vRefCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vRefCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetRefCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetRefNumber(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetRefNumber(vPrimary)
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

    Public Function GetTipeJournal2(ByVal VPrimary As Double) As Integer
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTipeJurnal2(VPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTipeJurnal2(VPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipeJournal(ByVal VPrimary As Double) As Integer
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTipeJournal(VPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTipeJournal(VPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetCreatedBy(vPrimary)
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
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
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

    Public Function ApakahSudahRevaluasi(ByVal vDate As Date) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlRevaluasi(TipeJurnal2, vDate).HasValue Then
                If DTA.ScalarQuery_GetJmlRevaluasi(TipeJurnal2, vDate).Value > 0 Then
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

    Public Function ApakahSudahAda(ByVal vRefCode As String) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByRefCode(vRefCode).HasValue Then
                If DTA.ScalarQuery_GetJmlByRefCode(vRefCode).Value > 0 Then
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

    Public Overridable Function ApakahDataBenar1(ByVal vBaru As Boolean, ByVal vDGVItem As DataGridView, ByVal vIndexPrimaryCurr As Integer, ByVal vDate As Date, _
    Optional ByVal vDGVBalance As DataGridView = Nothing, Optional ByVal vIndexPrimaryCurrencyBalance As Integer = -1, _
    Optional ByVal vIndexPayment As Integer = -1, Optional ByVal vIndexTipeDocAcuan As Integer = -1, _
    Optional ByVal vIndexPrimaryDocAcuan As Integer = -1, Optional ByVal vIndexDocNo As Integer = -1, Optional ByVal vPrimaryPihakAcuan As Double = -1, _
    Optional ByVal vIndexExchangeRate As Integer = -1, Optional ByVal vIndexTanggalTransaksiAsal As Integer = -1) As Boolean
        Dim LastRevaluationDate As Date = Today
        Dim DummyDate As Date
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim ObjClosing As New ClsTransaksiClosing
        Dim TipeKartu As ClsTransaksi_Kartu_APAR.EnumJenisKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP

        Try
            '1. jika tanggal sebelum tanggal closing maka tidak bisa diedit 
            If vDate.Date <= ObjClosing.GetClosingDate.Date Then
                MsgBox("You cannot save this data because transaction date <= closing date" _
                    , MsgBoxStyle.Information, "Information")
                Return False
            End If

            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mRefNumber) = True Then
                    MsgBox("Duplication in no journal '" + Me.mRefNumber + "', please change your number", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                'no po tidak boleh duplikat
                Dim NoLama As String = Me.GetRefCode(Me.mPrimary)
                If Me.RefNumber <> NoLama Then
                    If Me.ApakahSudahAda(Me.mRefNumber) = True Then
                        MsgBox("Duplication in no journal '" + Me.mRefNumber + "', please change your number", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            '2. jika membuat jurnal umum, selain jurnal revaluasi, tidak boleh membuat jurnal yg memiliki mata uang non IDR
            'pada periode yang sudah direvaluasi
            LastRevaluationDate = ObjCD.GetLasDateOfTheMonth(Me.GetTglRevaluasiSelanjutnya - TimeSpan.FromDays(33))
            DummyDate = ObjCD.GetFirstDateofTheMonth(Me.Tanggal)
            If (Me.TipeJournal = EnumTipeJournal.GeneralJournal And Me.TipeJurnal2 <> EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation) Or _
            Me.TipeJournal = EnumTipeJournal.PurchasePaymentJournal Or Me.TipeJournal = EnumTipeJournal.SalesReceivementJournal Then
                If LastRevaluationDate >= DummyDate Then
                    For Each rowDGV In vDGVItem.Rows
                        If rowDGV.Cells(vIndexPrimaryCurr).Value <> Core_Module.PrimaryBaseCurrency Then
                            MsgBox("The period of " + Format(DummyDate, "MMMMM yyyy") + " has been revaluate, you cannot save this data", _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    Next
                End If
            End If
            If Me.TipeJournal = EnumTipeJournal.SalesReceivementJournal Or Me.TipeJournal = EnumTipeJournal.PurchasePaymentJournal Then
                If LastRevaluationDate >= DummyDate Then
                    For Each rowDGV In vDGVBalance.Rows
                        If rowDGV.Cells(vIndexPrimaryCurrencyBalance).Value <> Core_Module.PrimaryBaseCurrency And _
                        CDec(rowDGV.Cells(vIndexPayment).Value) <> 0 Then
                            MsgBox("The period of " + Format(DummyDate, "MMMMM yyyy") + " has been revaluate, you cannot save this data", _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    Next
                End If

                If Me.TipeJournal = EnumTipeJournal.SalesReceivementJournal Then
                    TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR
                ElseIf Me.TipeJournal = EnumTipeJournal.PurchasePaymentJournal Then
                    TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP
                End If
                If Me.ApakahRateKonsisten(vDGVBalance, vIndexTipeDocAcuan, vIndexPrimaryDocAcuan, TipeKartu, _
                vIndexPayment, vIndexDocNo, vPrimaryPihakAcuan, vIndexPrimaryCurrencyBalance, vIndexExchangeRate) = False Then
                    Return False
                End If

                'cek tanggal pembayaran minimal sama dengan tanggal transaksi acuan
                For Each rowDGV In vDGVBalance.Rows
                    If CDec(rowDGV.Cells(vIndexPayment).Value) <> 0 Then
                        If vDate < CDate(rowDGV.Cells(vIndexTanggalTransaksiAsal).value) Then
                            vDGVBalance.ClearSelection()
                            rowDGV.Selected = True

                            MsgBox("Please do payment on later date than the reference transaction", _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                Next
            End If

            'return value
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            ObjCD = Nothing
            ObjClosing = Nothing
        End Try
    End Function

    'fungsi ini untuk mnegecek:
    '1. rate pada saat payment = rate pada kartu APAR, bisa tidak sama jika user melakukan edit transaksi invoice, ketika saat yg bersamaan orang akunting menyimpan data payment
    '2. rate transaksi IDR = 1
    '3. cek transaksi yg dibayar masih ada ketika melakukan pembayaran
    Private Function ApakahRateKonsisten(ByVal vDGVBalance As DataGridView, ByVal vIndexTipeDocAcuan As Integer, ByVal vIndexPrimaryDocAcuan As Integer, _
    ByVal vTipeKartu As ClsTransaksi_Kartu_APAR.EnumJenisKartu, ByVal vIndexPayment As Integer, ByVal vIndexDocNo As Integer, _
    ByVal vPrimaryPihakAcuan As Double, ByVal vIndexPrimaryCurrency As Integer, ByVal vIndexExchangeRate As Integer) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim PrimaryMax, PrimaryMin As Double
        Dim DTA As New DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_ap2TableAdapter
        Dim TBL As New DseTrans_Kartu_APAR.tbltransaksi_kartu_ar_ap2DataTable
        Dim row As DseTrans_Kartu_APAR.tbltransaksi_kartu_ar_ap2Row
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim Ketemu As Boolean = False
        Dim TipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_DP_Manual
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic 
        Dim PrimaryBaseCurrency As Double = 1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'mencari primary max min
            PrimaryMax = Double.MinValue
            PrimaryMin = Double.MaxValue
            For Each rowDGV In vDGVBalance.Rows
                If CDec(rowDGV.Cells(vIndexPayment).Value) <> 0 Then
                    If PrimaryMax < CDbl(rowDGV.Cells(vIndexPrimaryDocAcuan).Value) Then
                        PrimaryMax = CDbl(rowDGV.Cells(vIndexPrimaryDocAcuan).Value)
                    End If
                    If PrimaryMin > CDbl(rowDGV.Cells(vIndexPrimaryDocAcuan).Value) Then
                        PrimaryMin = CDbl(rowDGV.Cells(vIndexPrimaryDocAcuan).Value)
                    End If
                End If
            Next

            If PrimaryMax = Double.MinValue Then
                'tidak ada payment currency non idr, tidak perlu cek -> return true
                Return True
            Else
                'cek rate atas setiap doc yg dimaksud
                DTA.Fill(TBL, vPrimaryPihakAcuan, CInt(vTipeKartu), PrimaryMax, PrimaryMin)
                For Each rowDGV In vDGVBalance.Rows
                    If CDec(rowDGV.Cells(vIndexPayment).Value) <> 0 Then
                        Ketemu = False

                        For Each row In TBL.Rows
                            If row.Primary_Transaksi_Asal = rowDGV.Cells(vIndexPrimaryDocAcuan).Value Then
                                TipeBalance = CInt(rowDGV.Cells(vIndexTipeDocAcuan).Value)
                                TipeTransaksiAsal = CInt(row.Type_Transaksi_Asal)

                                If Me.ApakahCocok(TipeTransaksiAsal, TipeBalance) = True Then
                                    If row.exchange_rate <> CDec(rowDGV.Cells(vIndexExchangeRate).Value) Then
                                        If rowDGV.Cells(vIndexPrimaryCurrency).Value <> PrimaryBaseCurrency Then
                                            'cek rate hanya untuk transaksi non IDR
                                            vDGVBalance.ClearSelection()
                                            rowDGV.Selected = True

                                            MsgBox("Document: " + rowDGV.Cells(vIndexDocNo).Value.ToString + " uses exchange rate: " + _
                                            ObjCD.DigitGrouping(row.exchange_rate, 4) + ", please reload data.", _
                                            MsgBoxStyle.Information, "Information")

                                            Return False
                                        Else
                                            If CDec(rowDGV.Cells(vIndexExchangeRate).Value) <> 1 Then
                                                'rate harus 1, untuk transaksi IDR
                                                vDGVBalance.ClearSelection()
                                                rowDGV.Selected = True

                                                MsgBox("Document: " + rowDGV.Cells(vIndexDocNo).Value.ToString + " (IDR) should use use exchange rate: 1, please contact your vendor.", _
                                                MsgBoxStyle.Information, "Information")

                                                Return False
                                            Else
                                                Ketemu = True
                                                Exit For
                                            End If
                                        End If
                                    Else
                                        Ketemu = True
                                        Exit For
                                    End If
                                End If
                            End If
                        Next

                        If Ketemu = False Then
                            vDGVBalance.ClearSelection()
                            rowDGV.Selected = True

                            MsgBox("Reference document not found for: " + rowDGV.Cells(vIndexDocNo).Value.ToString + ", please reload transaction.", _
                            MsgBoxStyle.Information, "Information")

                            Return False
                        End If
                    End If 
                Next

                'kalo sampe sini, maka rate cocok semua, return true
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCur = Nothing
        End Try
    End Function

    Private Function ApakahCocok(ByVal TipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal, _
    ByVal TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance) As Boolean
        Try
            If TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_DP_Manual And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_DP Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual_Transfer And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_DP_Manual And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_DP Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual_Transfer And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual_Transfer And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual_Transfer And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice Then
                Return True
            ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis And TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function GetJml() As Double
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml4().HasValue = True Then
                Return DTA.ScalarQuery_GetJml4().Value
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

    Public Function GetJml(ByVal vTipeJurnal As EnumTipeJournal, ByVal vStartDate As Date, ByVal vEndDate As Date) As Double
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml2(vEndDate, vStartDate, vTipeJurnal).HasValue = True Then
                Return DTA.ScalarQuery_GetJml2(vEndDate, vStartDate, vTipeJurnal).Value
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

    Public Function GetTglRevaluasiSelanjutnya() As Date
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim TipeTransaksi1 As EnumTipeJournal = EnumTipeJournal.GeneralJournal
        Dim TipeTransaksi2 As EnumTipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_MaxDate_ByTipeTransaksi1_2(TipeTransaksi1, TipeTransaksi2).HasValue = True Then
                Return ObjCD.GetLasDateOfTheMonth(DTA.ScalarQuery_MaxDate_ByTipeTransaksi1_2(TipeTransaksi1, _
                TipeTransaksi2).Value + TimeSpan.FromDays(2))
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vTipeJurnal As EnumTipeJournal, ByVal vTipeJurnal2 As EnumTipeJurnal2) As Double
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJml3(vTipeJurnal, vTipeJurnal2).HasValue = True Then
                Return DTA.ScalarQuery_GetJml3(vTipeJurnal, vTipeJurnal2).Value
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

    Public Function GetJml(ByVal vTipeJurnal As EnumTipeJournal) As Double
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vTipeJurnal).HasValue = True Then
                Return DTA.ScalarQuery_GetJml(vTipeJurnal).Value
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
   ByVal TipeJournal As Integer, ByVal vTanggalDari As Date, _
   ByVal vTanggalSampai As Date, ByVal TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2) As Double
        Dim DTA As New DseTransaksiJournalTableAdapters.tblCariTrans_jurnalTableAdapter
        Dim TipeJurnalPurchasePayment As ClsTransaksiJournal.EnumTipeJournal = EnumTipeJournal.PurchasePaymentJournal
        Dim TipeJurnalSalesPayment As ClsTransaksiJournal.EnumTipeJournal = EnumTipeJournal.SalesReceivementJournal
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
        Dim TipeDocAcuan1 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
        Dim TipeDocAcuan2 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByReferenceCode Then
                Return DTA.ScalarQuery_GetJmlByRefNoLIKE(vKriteria + "%", TipeJournal, TipeJurnal2).Value
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                Return DTA.ScalarQuery_GetJmlByDesc(vKriteria + "%", TipeJournal, TipeJurnal2).Value
            ElseIf vSearchBy = EnumSearchCriteria.ByDate Then
                Return DTA.ScalarQuery_GetJmlByTanggalDariSampai(vTanggalSampai, vTanggalDari, TipeJournal, TipeJurnal2).Value
            ElseIf vSearchBy = EnumSearchCriteria.BySupplierCode Then
                Return DTA.ScalarQuery_GetJmlBySupplierCode(CInt(TipeJurnalPurchasePayment), vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.BySupplierName Then
                Return DTA.ScalarQuery_GetJmlBySupplierName(CInt(TipeJurnalPurchasePayment), vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByAOPNumber Then
                TipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
                TipeDocAcuan1 = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
                TipeDocAcuan2 = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
                Return DTA.ScalarQuery_GetJmlByAOPNumber(vKriteria + "%", TipeJurnalPurchasePayment, TipePayment, TipeDocAcuan1, TipeDocAcuan2)
            ElseIf vSearchBy = EnumSearchCriteria.ByPurchaseReturnNo Then
                TipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
                TipeDocAcuan1 = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
                TipeDocAcuan2 = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
                Return DTA.ScalarQuery_GetJmlByAOPNumber(vKriteria + "%", TipeJurnalPurchasePayment, TipePayment, TipeDocAcuan1, TipeDocAcuan2)
            ElseIf vSearchBy = EnumSearchCriteria.BySupplierInvoiceNo Then
                TipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
                TipeDocAcuan1 = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
                TipeDocAcuan2 = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
                Return DTA.ScalarQuery_GetJmlBySupplierInvoiceNumber(TipeJurnalPurchasePayment, TipePayment, TipeDocAcuan1, TipeDocAcuan2, vKriteria + "%")
            ElseIf vSearchBy = EnumSearchCriteria.ByCustomerCode Then
                Return DTA.ScalarQuery_GetJmlByCustCode(vKriteria + "%", TipeJurnalSalesPayment)
            ElseIf vSearchBy = EnumSearchCriteria.ByCustomerName Then
                Return DTA.ScalarQuery_GetJmlByCustName(vKriteria + "%", TipeJurnalSalesPayment)
            ElseIf vSearchBy = EnumSearchCriteria.BySalesReturnNo Then
                TipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesInvoice_SalesInvoiceDP
                TipeDocAcuan1 = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
                TipeDocAcuan2 = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
                Return DTA.ScalarQuery_GetJmlByAOPNumber(vKriteria + "%", TipeJurnalSalesPayment, TipePayment, TipeDocAcuan1, TipeDocAcuan2)
            ElseIf vSearchBy = EnumSearchCriteria.ByInvoiceNumber Then
                TipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesInvoice_SalesInvoiceDP
                TipeDocAcuan1 = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
                TipeDocAcuan2 = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
                Return DTA.ScalarQuery_GetJmlByAOPNumber(vKriteria + "%", TipeJurnalSalesPayment, TipePayment, TipeDocAcuan1, TipeDocAcuan2)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
  ByVal DGV As DataGridView, ByVal TBL As DseTransaksiJournal.tblCariTrans_jurnalDataTable, _
  ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal TipeJournal As Integer, _
   ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date, ByVal vTipeJurnal2 As EnumTipeJurnal2, _
   ByVal FilterCreatedBy As String) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tblCariTrans_jurnalTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            StartT = Now
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByReferenceCode Then
                QueryResult = DTA.FillBy_RefCodeLIKE(TBL, Criteria + "%", TipeJournal, vTipeJurnal2)
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DecriptionLIKE(TBL, Criteria + "%", TipeJournal, vTipeJurnal2)
            ElseIf SearchCriteria = EnumSearchCriteria.ByDate Then
                QueryResult = DTA.FillBy_TanggalDariSampai(TBL, vTanggalSampai, vTanggalDari, TipeJournal, vTipeJurnal2)
            End If

            'update dgv
            If FilterCreatedBy <> "" Then
                TBL.DefaultView.RowFilter = FilterCreatedBy
                QueryResult = TBL.DefaultView.Count
            Else
                TBL.DefaultView.RowFilter = ""
            End If
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "No Journal"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Transaction date"
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(3).HeaderText = "Description"
                DGV.Columns(4).Visible = False
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

    Public Function IsiDGV_SalesPayment(ByVal Criteria As String, _
  ByVal DGV As DataGridView, ByVal TBL As DseTransaksiJournal.tbltrans_jurnal2DataTable, _
  ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal TipeJournal As Integer, _
   ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date, ByVal vTipeJurnal2 As EnumTipeJurnal2, _
    ByVal FilterCreatedBy As String) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal2TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesInvoice_SalesInvoiceDP
        Dim TipeDocAcuan1 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
        Dim TipeDocAcuan2 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual

        Try
            StartT = Now
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByReferenceCode Then
                QueryResult = DTA.FillBy_RefNoLike(TBL, TipeJournal, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, TipeJournal, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDate Then
                QueryResult = DTA.FillBy_TanggalDariSampai(TBL, TipeJournal, vTanggalSampai, vTanggalDari)
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerCode Then
                QueryResult = DTA.FillBy_CustCodeLIKE(TBL, TipeJournal, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerName Then
                QueryResult = DTA.FillBy_CustNameLIKE(TBL, TipeJournal, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByInvoiceNumber Then
                QueryResult = DTA.FillBy_InvoiceNumber(TBL, TipeJournal, TipePayment, TipeDocAcuan1, TipeDocAcuan2, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesReturnNo Then
                TipeDocAcuan1 = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
                TipeDocAcuan2 = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
                QueryResult = DTA.FillBy_InvoiceNumber(TBL, TipeJournal, TipePayment, TipeDocAcuan1, TipeDocAcuan2, Criteria + "%")
            End If

            'update dgv
            If FilterCreatedBy <> "" Then
                TBL.DefaultView.RowFilter = FilterCreatedBy
                QueryResult = TBL.DefaultView.Count
            Else
                TBL.DefaultView.RowFilter = ""
            End If
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "No Journal"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Transaction date"
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(5).HeaderText = "Customer Code"
                DGV.Columns(5).HeaderText = "Description"
                DGV.Columns(6).Visible = False
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

    Public Function IsiDGV_PurchasePayment(ByVal Criteria As String, _
  ByVal DGV As DataGridView, ByVal TBL As DseTransaksiJournal.tbltrans_jurnal1DataTable, _
  ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal TipeJournal As Integer, _
   ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date, ByVal vTipeJurnal2 As EnumTipeJurnal2, _
    ByVal FilterCreatedBy As String) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
        Dim TipeDocAcuan1 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
        Dim TipeDocAcuan2 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual

        Try
            StartT = Now
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByReferenceCode Then
                QueryResult = DTA.FillBy_RefNoLIKE(TBL, Criteria + "%", TipeJournal)
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%", TipeJournal)
            ElseIf SearchCriteria = EnumSearchCriteria.ByDate Then
                QueryResult = DTA.FillBy_TglSampaiDari(TBL, vTanggalSampai, vTanggalDari, TipeJournal)
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierCode Then
                QueryResult = DTA.FillBy_SupplierCodeLIKE(TBL, Criteria + "%", TipeJournal)
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                QueryResult = DTA.FillBy_SupplierName(TBL, Criteria + "%", TipeJournal)
            ElseIf SearchCriteria = EnumSearchCriteria.ByAOPNumber Then
                QueryResult = DTA.FillBy_AOPNumberLIKE(TBL, Criteria + "%", TipeDocAcuan1, TipeDocAcuan2, TipeJournal, TipePayment)
            ElseIf SearchCriteria = EnumSearchCriteria.ByPurchaseReturnNo Then
                TipeDocAcuan1 = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
                TipeDocAcuan2 = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
                QueryResult = DTA.FillBy_AOPNumberLIKE(TBL, Criteria + "%", TipeDocAcuan1, TipeDocAcuan2, TipeJournal, TipePayment)
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierInvoiceNo Then
                QueryResult = DTA.FillBy_SupplierInvoiceNo(TBL, TipeDocAcuan1, TipeDocAcuan2, TipeJournal, TipePayment, Criteria + "%")
            End If

            'update dgv
            If FilterCreatedBy <> "" Then
                TBL.DefaultView.RowFilter = FilterCreatedBy
                QueryResult = TBL.DefaultView.Count
            Else
                TBL.DefaultView.RowFilter = ""
            End If
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "No Journal"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Transaction date"
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(5).HeaderText = "Description"
                DGV.Columns(3).HeaderText = "Supplier Code"
                DGV.Columns(6).Visible = False
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

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = -1
            mRefNumber = ""
            mDate = Now
            mTipeJournal = EnumTipeJournal.AdjustmenJournal
            mUrutan = 0
            mDesc = ""
            mPrimaryDokumenAcuan = -1
            mPrimaryPihakAcuan = -1
            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function CariDetail_EndOfMonth(ByVal vPrimaryMaster As Double, _
   ByVal vDGVItem As DataGridView, ByVal frm As FrmTransaksiJournal_EndOfMonth, _
   ByVal tipe_journal As Integer, ByVal vDGVCurrency As DataGridView, ByVal vDGVItemAR As DataGridView, _
    ByVal vDGVItemAP As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnalDataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnalRow
        Dim ObjItem As New ClsTransaksiJournal_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryMaster, tipe_journal) > 0 Then
                'salin data header
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = vPrimaryMaster
                    Me.mRefNumber = .Ref_Number
                    Me.mDate = ._Date
                    Me.mUrutan = .Urutan
                    Me.mTipeJournal = .Tipe_Journal
                    Me.mPrimaryDokumenAcuan = .Primary_Document_Acuan
                    Me.mPrimaryPihakAcuan = .Primary_Pihak_Acuan
                    Me.mDesc = .Desc
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mTipeJurnal2 = .Tipe_Jurnal_2

                    Me.mPrimaryAkunDifference = .primary_akun_difference
                    Me.mPrimaryAkunPayTo = .primary_akun_pay_to
                    Me.mPaid = .paid
                    Me.mPaid_IDR = .paid_IDR
                    Me.mARPaid_IDR = .AR_paid_IDR
                    Me.mDifference = .difference
                    Me.mExchangeRate = .exchange_rate
                    Me.mExchageRateGainLoss = .exchange_rate_gain_loss
                    Me.mReferenceParty = .Reference_Party
                End With
                'salin data item ada 2 item 
                'CodeNo
                ObjItem.IsiDGVItem_EndOfMonth(vDGVItem, vPrimaryMaster, frm, vDGVCurrency, vDGVItemAR, vDGVItemAP)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDGVItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function CariDetail(ByVal vPrimaryMaster As Double, _
    ByVal vDGVItem As DataGridView, ByVal frm As FrmTransaksiJournal, _
    ByVal tipe_journal As Integer) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnalDataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnalRow
        Dim ObjItem As New ClsTransaksiJournal_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryMaster, tipe_journal) > 0 Then
                'salin data header
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = vPrimaryMaster
                    Me.mRefNumber = .Ref_Number
                    Me.mDate = ._Date
                    Me.mUrutan = .Urutan
                    Me.mTipeJournal = .Tipe_Journal
                    Me.mPrimaryDokumenAcuan = .Primary_Document_Acuan
                    Me.mPrimaryPihakAcuan = .Primary_Pihak_Acuan
                    Me.mDesc = .Desc
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mTipeJurnal2 = .Tipe_Jurnal_2

                    Me.mPrimaryAkunDifference = .primary_akun_difference
                    Me.mPrimaryAkunPayTo = .primary_akun_pay_to
                    Me.mPaid = .paid
                    Me.mPaid_IDR = .paid_IDR
                    Me.mARPaid_IDR = .AR_paid_IDR
                    Me.mDifference = .difference
                    Me.mExchangeRate = .exchange_rate
                    Me.mExchageRateGainLoss = .exchange_rate_gain_loss
                    Me.mReferenceParty = .Reference_Party
                End With
                'salin data item ada 2 item 
                'CodeNo
                ObjItem.IsiDGVItem(vDGVItem, vPrimaryMaster, frm)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDGVItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function CariDetail_DepresiasiAsset(ByVal vPrimaryMaster As Double, _
   ByVal vDGVItem As DataGridView, ByVal frm As FrmTransaksiJournal_AssetDepreciation, _
   ByVal tipe_journal As Integer, ByVal vDGVAssetDetail As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnalDataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnalRow
        Dim ObjItem As New ClsTransaksiJournal_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryMaster, tipe_journal) > 0 Then
                'salin data header
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = vPrimaryMaster
                    Me.mRefNumber = .Ref_Number
                    Me.mDate = ._Date
                    Me.mUrutan = .Urutan
                    Me.mTipeJournal = .Tipe_Journal
                    Me.mPrimaryDokumenAcuan = .Primary_Document_Acuan
                    Me.mPrimaryPihakAcuan = .Primary_Pihak_Acuan
                    Me.mDesc = .Desc
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mTipeJurnal2 = .Tipe_Jurnal_2

                    Me.mPrimaryAkunDifference = .primary_akun_difference
                    Me.mPrimaryAkunPayTo = .primary_akun_pay_to
                    Me.mPaid = .paid
                    Me.mPaid_IDR = .paid_IDR
                    Me.mARPaid_IDR = .AR_paid_IDR
                    Me.mDifference = .difference
                    Me.mExchangeRate = .exchange_rate
                    Me.mExchageRateGainLoss = .exchange_rate_gain_loss
                    Me.mReferenceParty = .Reference_Party
                End With

                'salin data item ada 2 item 
                'CodeNo
                ObjItem.IsiDGVItem_Asset(vDGVItem, vPrimaryMaster, frm, vDGVAssetDetail)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDGVItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function CariDetail(ByVal vPrimaryMaster As Double, _
 ByVal vDGVItem As DataGridView, ByVal frm As FrmTransaksiJournal_CashBank, _
 ByVal tipe_journal As Integer, ByVal vDGVItem_payment As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnalDataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnalRow
        Dim ObjItem As New ClsTransaksiJournal_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryMaster, tipe_journal) > 0 Then
                'salin data header
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = vPrimaryMaster
                    Me.mRefNumber = .Ref_Number
                    Me.mDate = ._Date
                    Me.mUrutan = .Urutan
                    Me.mTipeJournal = .Tipe_Journal
                    Me.mPrimaryDokumenAcuan = .Primary_Document_Acuan
                    Me.mPrimaryPihakAcuan = .Primary_Pihak_Acuan
                    Me.mDesc = .Desc
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mTipeJurnal2 = .Tipe_Jurnal_2

                    Me.mPrimaryAkunDifference = .primary_akun_difference
                    Me.mPrimaryAkunPayTo = .primary_akun_pay_to
                    Me.mPaid = .paid
                    Me.mPaid_IDR = .paid_IDR
                    Me.mARPaid_IDR = .AR_paid_IDR
                    Me.mDifference = .difference
                    Me.mExchangeRate = .exchange_rate
                    Me.mExchageRateGainLoss = .exchange_rate_gain_loss
                    Me.mPrimaryCurr = .Primary_Currency
                    Me.mTipeInOut = .Type_InOut
                    Me.mReferenceParty = .Reference_Party
                End With
                'salin data item ada 2 item 
                'CodeNo
                ObjItem.IsiDGVItem(vDGVItem, vPrimaryMaster, frm, vDGVItem_payment)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDGVItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vDGVItem As DataGridView, _
           ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal, ByVal TipeJournal As Integer) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_General_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    DTAInOUT.Delete(Me.vPrimary_Antrian)
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vDGVItem, vUserCode, Frm, TipeJournal)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Overridable Function SimpanBaru1(ByVal vDGVItem As DataGridView, _
           ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal, ByVal TipeJournal As Integer) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.mTipeJournal = TipeJournal
            Me.mUrutan = -1
            Me.mPrimaryDokumenAcuan = -1
            Me.mPrimaryPihakAcuan = -1
            If Me.ApakahDataBenar1(True, vDGVItem, Frm.ColPrimaryCurrency.Index, Me.Tanggal) = True Then

                If DTA.Insert(Me.RefNumber, Me.Tanggal, Me.mUrutan, Me.mTipeJournal, Me.mPrimaryDokumenAcuan, _
                Me.mPrimaryPihakAcuan, Me.Description, vUserCode, Now, _
                 vUserCode, Now, Now, Me.mPrimaryAkunPayTo, Me.mPrimaryAkunDifference, Me.mPaid, Me.mExchangeRate, Me.mPaid_IDR, _
                 Me.mARPaid_IDR, Me.mDifference, Me.mExchageRateGainLoss, Me.mTipeJurnal2, Me.mTipeInOut, Me.mPrimaryCurr, Me.mReferenceParty, Me.mPaymentMethod) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mRefNumber).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItem(vDGVItem, Me.mPrimary, Frm)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mRefNumber + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru1(vDGVItem, vUserCode, Frm, TipeJournal)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal, ByVal TipeJournal As Integer) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_General_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vDGVItem, vUserCode, frm, TipeJournal)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1(ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal, ByVal TipeJournal As Integer) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.mTipeJournal = TipeJournal
            Me.mUrutan = -1
            Me.mPrimaryDokumenAcuan = -1
            Me.mPrimaryPihakAcuan = -1
            If Me.ApakahDataBenar1(False, vDGVItem, frm.ColPrimaryCurrency.Index, Me.Tanggal) = True Then
                If DTA.Update(Me.RefNumber, Me.Tanggal, Me.mUrutan, Me.mTipeJournal, Me.mPrimaryDokumenAcuan, _
                Me.mPrimaryPihakAcuan, Me.Description, vUserCode, Now, Now, _
                Me.mPrimaryAkunPayTo, Me.mPrimaryAkunDifference, Me.mPaid, Me.mExchangeRate, Me.mPaid_IDR, Me.mARPaid_IDR, Me.mDifference, _
                Me.mExchageRateGainLoss, Me.mTipeJurnal2, Me.mTipeInOut, Me.mPrimaryCurr, Me.mReferenceParty, Me.mPaymentMethod, Me.Primary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)

                    'simpan data item baru
                    ObjItem.SimpanItem(vDGVItem, Me.mPrimary, frm)

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
                Return Me.SimpanPerubahan1(vDGVItem, vUserCode, frm, TipeJournal)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1(ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal_AssetDepreciation, ByVal TipeJournal As Integer, ByVal vDGVAsset As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Me.mTipeJournal = TipeJournal
            Me.mUrutan = -1
            Me.mPrimaryDokumenAcuan = -1
            Me.mPrimaryPihakAcuan = -1
            If Me.ApakahDataBenar1(False, vDGVItem, frm.ColAsset_PrimaryCurrency.Index, Me.Tanggal) = True Then
                If DTA.Update(Me.RefNumber, Me.Tanggal, Me.mUrutan, Me.mTipeJournal, Me.mPrimaryDokumenAcuan, _
                Me.mPrimaryPihakAcuan, Me.Description, vUserCode, Now, Now, _
                Me.mPrimaryAkunPayTo, Me.mPrimaryAkunDifference, Me.mPaid, Me.mExchangeRate, Me.mPaid_IDR, Me.mARPaid_IDR, Me.mDifference, _
                Me.mExchageRateGainLoss, Me.mTipeJurnal2, Me.mTipeInOut, Me.mPrimaryCurr, Me.mReferenceParty, Me.mPaymentMethod, Me.Primary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)

                    'simpan data item baru
                    ObjItem.SimpanItem(vDGVItem, Me.mPrimary, frm, vDGVAsset)

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
                Return Me.SimpanPerubahan1(vDGVItem, vUserCode, frm, TipeJournal, vDGVAsset)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru_EndOfMonth(ByVal vDGVItem As DataGridView, _
           ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal_EndOfMonth, ByVal TipeJournal As Integer, _
           ByVal vDGVCurrency As DataGridView, ByVal vDGVItemAR As DataGridView, ByVal vDGVItemAP As DataGridView) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_EOM_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    DTAInOUT.Delete(Me.vPrimary_Antrian)
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1_EndOfMonth(vDGVItem, vUserCode, Frm, TipeJournal, vDGVCurrency, vDGVItemAR, vDGVItemAP)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Overridable Function SimpanBaru1_EndOfMonth(ByVal vDGVItem As DataGridView, _
           ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal_EndOfMonth, ByVal TipeJournal As Integer, _
           ByVal vDGVCurrency As DataGridView, ByVal vDGVItemAR As DataGridView, ByVal vDGVItemAP As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Me.mTipeJournal = TipeJournal
            Me.mUrutan = -1
            Me.mPrimaryDokumenAcuan = -1
            Me.mPrimaryPihakAcuan = -1

            If Me.ApakahDataBenar1(True, vDGVItem, Frm.ColPrimaryCurrency.Index, Me.Tanggal) = True Then
                If DTA.Insert(Me.RefNumber, Me.Tanggal, Me.mUrutan, Me.mTipeJournal, Me.mPrimaryDokumenAcuan, _
                Me.mPrimaryPihakAcuan, Me.Description, vUserCode, Now, _
                 vUserCode, Now, Now, Me.mPrimaryAkunPayTo, Me.mPrimaryAkunDifference, Me.mPaid, Me.mExchangeRate, Me.mPaid_IDR, _
                 Me.mARPaid_IDR, Me.mDifference, Me.mExchageRateGainLoss, Me.mTipeJurnal2, Me.mTipeInOut, Me.mPrimaryCurr, Me.mReferenceParty, Me.mPaymentMethod) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mRefNumber).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItem_EndOfMonth(vDGVItem, Me.mPrimary, Frm, vDGVCurrency, vDGVItemAR, vDGVItemAP)

                    'update data journal ar non-idr, setelah bulan ini
                    Me.UpdateJournalAR_AP_NonIDR_Simpan(Me.mPrimary, Me.mDate, Frm, vDGVCurrency, vDGVItemAR, Me.mDate, vDGVItemAP)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mRefNumber + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru1_EndOfMonth(vDGVItem, vUserCode, Frm, TipeJournal, vDGVCurrency, vDGVItemAR, vDGVItemAP)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Private Sub UpdateJournalAR_AP_NonIDR_Simpan(ByVal vPrimaryJurnal As Double, ByVal Periode As Date, ByVal Frm As FrmTransaksiJournal_EndOfMonth, _
    ByVal vDGVCurrency As DataGridView, ByVal vDGVItemAR As DataGridView, ByVal vTanggalJurnal As Date, ByVal vDGVItemAP As DataGridView)
        Try
            'update balance + kartu ar
            Me.UpdateBalance_Kartu_AR(vPrimaryJurnal, Frm, vDGVItemAR, vTanggalJurnal)
            'update jurnal untuk semua transaksi sales payment dengan tanggal 1 hari setelah revaluatioin
            Me.UpdateJurnalAR(vTanggalJurnal + TimeSpan.FromDays(1))
            'update balance + kartu ap
            Me.UpdateBalance_Kartu_AP(vPrimaryJurnal, Frm, vDGVItemAP, vTanggalJurnal)
            'update jurnal untuk semua transaksi purchase payment dengan tanggal 1 hari setelah revaluatioin
            Me.UpdateJurnalAP(vTanggalJurnal + TimeSpan.FromDays(1))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateBalance_Kartu_AR(ByVal vPrimaryJurnal As Double, ByVal Frm As FrmTransaksiJournal_EndOfMonth, _
    ByVal vDGVItemAR As DataGridView, ByVal vTanggalJurnal As Date)
        Dim rowDGV As DataGridViewRow
        Dim DTABalance As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DefPrec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DTAInv As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TBLInv As New DseTransaksiJournal_EndOfMonth.tbltrans_salesinvoiceDataTable
        Dim DTASR As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_salesreturnTableAdapter
        Dim TBLSR As New DseTransaksiJournal_EndOfMonth.tbltrans_salesreturnDataTable
        Dim DTAInvManual As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_historicalsalesTableAdapter
        Dim TBLInvManual As New DseTransaksiJournal_EndOfMonth.tbltrans_historicalsalesDataTable
        Dim DTASRMAnual As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Dim TBLSRManual As New DseTransaksiJournal_EndOfMonth.tbltrans_historicalsalesreturn_manualDataTable
        Dim DokumenAcuan As String = ""
        Dim MinDate, MaxDate As Date
        Dim DTA_CNDN As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_credit_noteTableAdapter
        Dim TBL_CNDN As New DseTransaksiJournal_EndOfMonth.tbltrans_credit_noteDataTable

        Try
            MinDate = Date.MaxValue
            MaxDate = Date.MinValue
            For Each rowDGV In vDGVItemAR.Rows
                If CDate(rowDGV.Cells(Frm.ColItemAR_Tanggal.Index).Value).Date < MinDate.Date Then
                    MinDate = CDate(rowDGV.Cells(Frm.ColItemAR_Tanggal.Index).Value).Date
                End If
                If CDate(rowDGV.Cells(Frm.ColItemAR_Tanggal.Index).Value).Date > MaxDate.Date Then
                    MaxDate = CDate(rowDGV.Cells(Frm.ColItemAR_Tanggal.Index).Value).Date
                End If
            Next

            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            DTASR.Connection.ConnectionString = DTAInv.Connection.ConnectionString
            DTAInvManual.Connection.ConnectionString = DTAInv.Connection.ConnectionString
            DTASRMAnual.Connection.ConnectionString = DTAInv.Connection.ConnectionString
            DTABalance.Connection.ConnectionString = Core_Module.GetConstring
            DTA_CNDN.Connection.ConnectionString = Core_Module.GetConstring

            DefPrec = ObjCur.GetDecimal_DefaultCurrency

            'fill tabel untuk no transaksi acuan
            DTAInv.Fill(TBLInv, Core_Module.PrimaryBaseCurrency, MaxDate, MinDate)
            DTASR.Fill(TBLSR, Core_Module.PrimaryBaseCurrency, MaxDate, MinDate)
            DTAInvManual.Fill(TBLInvManual)
            DTASRMAnual.Fill(TBLSRManual)
            DTA_CNDN.FillBy_TglSampaiDari(TBL_CNDN, MaxDate, MinDate, Core_Module.PrimaryBaseCurrency)

            For Each rowDGV In vDGVItemAR.Rows
                With rowDGV
                    If CDec(.Cells(Frm.ColItemAR_Selisih.Index).Value) <> 0 Then
                        'get no dokumen acuan
                        DokumenAcuan = Me.GetNumberTransaksiAR(TBLInv, TBLSR, TBLInvManual, TBLSRManual, _
                        .Cells(Frm.ColItemAR_TipeTransaksiAsal.Index).Value, .Cells(Frm.ColItemAR_PrimaryTransaksiAsal.Index).Value, TBL_CNDN)

                        'update balance ar untuk transaksi ybs, ganti exchange rate dan hitung ulang saldo idr
                        DTABalance.UpdateQuery_Revaluate(Frm.GetExchangeRate(CDbl(rowDGV.Cells(Frm.ColItemAR_PrimaryCurrency.Index).Value)), _
                        DefPrec, ObjKartu.GetTipeSalesBalance(.Cells(Frm.ColItemAR_TipeTransaksiAsal.Index).Value), _
                        .Cells(Frm.ColItemAR_PrimaryTransaksiAsal.Index).Value, .Cells(Frm.ColItemAR_PrimaryCurrency.Index).Value)

                        'update kartu ar, insert masuk kartu sejumlah selisih saldo, dan nol untuk saldo original
                        If CDec(.Cells(Frm.ColItemAR_Selisih.Index).Value) > 0 Then
                            'masukkan pada kredit
                            ObjKartu.Simpan(vPrimaryJurnal, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Revaluation, DokumenAcuan, .Cells(Frm.ColItemAR_PrimaryPihakAcuan.Index).Value, _
                            0, CDec(.Cells(Frm.ColItemAR_Selisih.Index).Value), "Revaluate AR " + Format(vTanggalJurnal, "MMM yyyy"), ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, vTanggalJurnal, _
                            .Cells(Frm.ColItemAR_PrimaryTransaksiAsal.Index).Value, .Cells(Frm.ColItemAR_Tanggal.Index).Value, .Cells(Frm.ColItemAR_PaymentTermAcuan.Index).Value, _
                            .Cells(Frm.ColItemAR_TipeTransaksiAsal.Index).Value, _
                            .Cells(Frm.ColItemAR_PrimaryCurrency.Index).Value, 0, 0, 1, .Cells(Frm.ColItemAR_PrimaryAkunAR.Index).Value)
                        Else
                            'masukkan pada debet, nilai absoultenya
                            ObjKartu.Simpan(vPrimaryJurnal, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Revaluation, DokumenAcuan, .Cells(Frm.ColItemAR_PrimaryPihakAcuan.Index).Value, _
                            Math.Abs(CDec(.Cells(Frm.ColItemAR_Selisih.Index).Value)), 0, "Revaluate AR " + Format(vTanggalJurnal, "MMM yyyy"), ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, vTanggalJurnal, _
                            .Cells(Frm.ColItemAR_PrimaryTransaksiAsal.Index).Value, .Cells(Frm.ColItemAR_Tanggal.Index).Value, .Cells(Frm.ColItemAR_PaymentTermAcuan.Index).Value, .Cells(Frm.ColItemAR_TipeTransaksiAsal.Index).Value, _
                            .Cells(Frm.ColItemAR_PrimaryCurrency.Index).Value, 0, 0, 1, .Cells(Frm.ColItemAR_PrimaryAkunAR.Index).Value)
                        End If
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTABalance.Dispose()
            ObjKartu = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
            DTAInv.Dispose()
            TBLInv.Dispose()
            DTASR.Dispose()
            TBLSR.Dispose()
            DTAInvManual.Dispose()
            TBLInvManual.Dispose()
            DTASRMAnual.Dispose()
            TBLSRManual.Dispose()
            DTA_CNDN.Dispose()
            TBL_CNDN.Dispose()
        End Try
    End Sub

    Private Sub UpdateBalance_Kartu_AP(ByVal vPrimaryJurnal As Double, ByVal Frm As FrmTransaksiJournal_EndOfMonth, _
  ByVal dgvItemAP As DataGridView, ByVal vTanggalJurnal As Date)
        Dim rowDGV As DataGridViewRow
        Dim DTABalance As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DefPrec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DTAAOP As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_aopTableAdapter
        Dim TBLAOP As New DseTransaksiJournal_EndOfMonth.tbltrans_aopDataTable
        Dim DTAPR As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_purchasereturnTableAdapter
        Dim TBLPR As New DseTransaksiJournal_EndOfMonth.tbltrans_purchasereturnDataTable
        Dim DTAAOPManual As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Dim TBLAOPManual As New DseTransaksiJournal_EndOfMonth.tbltrans_historicalpurchasebalancemanualDataTable
        Dim DTAPRMAnual As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Dim TBLPRManual As New DseTransaksiJournal_EndOfMonth.tbltrans_historicalpurchasereturnbalancemanualDataTable
        Dim DokumenAcuan As String = ""
        Dim TipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual
        Dim DTA_CNDN As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_credit_noteTableAdapter
        Dim TBL_CNDN As New DseTransaksiJournal_EndOfMonth.tbltrans_credit_noteDataTable
        Dim MinDate, MaxDate As Date

        Try
            MinDate = Date.MaxValue
            MaxDate = Date.MinValue
            For Each rowDGV In dgvItemAP.Rows
                If CDate(rowDGV.Cells(Frm.ColItemAP_Tanggal.Index).Value).Date < MinDate.Date Then
                    MinDate = CDate(rowDGV.Cells(Frm.ColItemAP_Tanggal.Index).Value).Date
                End If
                If CDate(rowDGV.Cells(Frm.ColItemAP_Tanggal.Index).Value).Date > MaxDate.Date Then
                    MaxDate = CDate(rowDGV.Cells(Frm.ColItemAP_Tanggal.Index).Value).Date
                End If
            Next

            DTAAOP.Connection.ConnectionString = Core_Module.GetConstring
            DTAPR.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
            DTAAOPManual.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
            DTAPRMAnual.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
            DTABalance.Connection.ConnectionString = Core_Module.GetConstring
            DTA_CNDN.Connection.ConnectionString = Core_Module.GetConstring
            DefPrec = ObjCur.GetDecimal_DefaultCurrency

            'fill tabel untuk no transaksi acuan
            DTAAOP.Fill(TBLAOP, Core_Module.PrimaryBaseCurrency, MaxDate, MinDate)
            DTAPR.Fill(TBLPR, Core_Module.PrimaryBaseCurrency, MaxDate, MinDate)
            DTAAOPManual.Fill(TBLAOPManual)
            DTAPRMAnual.Fill(TBLPRManual)
            DTA_CNDN.FillBy_TglSampaiDari(TBL_CNDN, MaxDate, MinDate, Core_Module.PrimaryBaseCurrency)

            For Each rowDGV In dgvItemAP.Rows
                With rowDGV
                    'update kartu ar, insert masuk kartu sejumlah selisih saldo, dan nol untuk saldo original
                    If CDec(.Cells(Frm.ColItemAP_Selisih.Index).Value) <> 0 Then
                        'get no dokumen acuan
                        DokumenAcuan = Me.GetNumberTransaksiAP(TBLAOP, TBLPR, TBLAOPManual, TBLPRManual, _
                        .Cells(Frm.ColItemAP_TipeTransaksiAsal.Index).Value, .Cells(Frm.ColItemAP_PrimaryTransaksiAsal.Index).Value, TBL_CNDN)

                        'update balance ar untuk transaksi ybs, ganti exchange rate dan hitung ulang saldo idr
                        DTABalance.UpdateQuery_Revaluate(Frm.GetExchangeRate(CDbl(rowDGV.Cells(Frm.ColItemAP_PrimaryCurrency.Index).Value)), _
                         DefPrec, ObjKartu.GetTipeSalesBalance(.Cells(Frm.ColItemAP_TipeTransaksiAsal.Index).Value), _
                         .Cells(Frm.ColItemAP_PrimaryTransaksiAsal.Index).Value, .Cells(Frm.ColItemAP_PrimaryCurrency.Index).Value)

                        If CDec(.Cells(Frm.ColItemAP_Selisih.Index).Value) > 0 Then
                            'masukkan pada kredit
                            ObjKartu.Simpan(vPrimaryJurnal, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Revaluation, DokumenAcuan, .Cells(Frm.ColItemAP_PrimaryPihakAcuan.Index).Value, _
                            0, CDec(.Cells(Frm.ColItemAP_Selisih.Index).Value), "Revaluate AP " + Format(vTanggalJurnal, "MMM yyyy"), ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, vTanggalJurnal, _
                            .Cells(Frm.ColItemAP_PrimaryTransaksiAsal.Index).Value, .Cells(Frm.ColItemAP_Tanggal.Index).Value, .Cells(Frm.ColItemAP_PaymentTermAcuan.Index).Value, _
                            .Cells(Frm.ColItemAP_TipeTransaksiAsal.Index).Value, _
                            .Cells(Frm.ColItemAP_PrimaryCurrency.Index).Value, 0, 0, 1, .Cells(Frm.ColItemAP_PrimaryAkunAP.Index).Value)
                        Else
                            'masukkan pada debet, nilai absoultenya
                            ObjKartu.Simpan(vPrimaryJurnal, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Revaluation, DokumenAcuan, .Cells(Frm.ColItemAP_PrimaryPihakAcuan.Index).Value, _
                            Math.Abs(CDec(.Cells(Frm.ColItemAP_Selisih.Index).Value)), 0, "Revaluate AP " + Format(vTanggalJurnal, "MMM yyyy"), ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, vTanggalJurnal, _
                            .Cells(Frm.ColItemAP_PrimaryTransaksiAsal.Index).Value, .Cells(Frm.ColItemAP_Tanggal.Index).Value, .Cells(Frm.ColItemAP_PaymentTermAcuan.Index).Value, .Cells(Frm.ColItemAP_TipeTransaksiAsal.Index).Value, _
                            .Cells(Frm.ColItemAP_PrimaryCurrency.Index).Value, 0, 0, 1, .Cells(Frm.ColItemAP_PrimaryAkunAP.Index).Value)
                        End If
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTABalance.Dispose()
            ObjKartu = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
            DTAAOP.Dispose()
            TBLAOP.Dispose()
            DTAPR.Dispose()
            TBLPR.Dispose()
            DTAAOPManual.Dispose()
            TBLAOPManual.Dispose()
            DTAPRMAnual.Dispose()
            TBLPRManual.Dispose()
            DTA_CNDN.Dispose()
            TBL_CNDN.Dispose()
        End Try
    End Sub

    Private Sub UpdateJurnalAR(ByVal vStartDate As Date)
        Dim i As New FrmTransaksiJournal_Sales
        Dim DTA As New DseTransaksiJournal_EndOfMonthTableAdapters.tblsales_payment_jurnal1TableAdapter
        Dim TBL As New DseTransaksiJournal_EndOfMonth.tblsales_payment_jurnal1DataTable
        Dim row As DseTransaksiJournal_EndOfMonth.tblsales_payment_jurnal1Row
        Dim TipeJurnal As ClsTransaksiJournal.EnumTipeJournal = EnumTipeJournal.SalesReceivementJournal
        Dim ij_i As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vStartDate, TipeJurnal, Core_Module.PrimaryBaseCurrency)

            i.Show()
            i.MinimumSize = New Drawing.Size(100, 100)
            i.Size = New Drawing.Size(100, 100)
            i.Location = New Point(-100, -100)

            For Each row In TBL.Rows
                i.AktifkanDetail()
                i.ObjJurnal.CariDetailSalesJournal(row.Primary, i.dgvItem, i, i.TipeJournal, i.DgvDifference)
                i.SalinkeField()
                ij_i = i.TxtPrimaryAccountPayTo.Text
                i.UpdateJurnal()
                If i.TxtPrimaryAccountPayTo.Text <> ij_i Then
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                End If
                i.ObjJurnal.SimpanPerubahanSalesJournal1(i.dgvItem, i.ObjJurnal.CreatedBy, i, i.TipeJournal, i.dgvSalesBalance, i.DgvDifference)
            Next

            i.Hide()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub UpdateJurnalAP(ByVal vStartDate As Date)
        Dim i As New FrmTransaksiJournal_Purchase
        Dim DTA As New DseTransaksiJournal_EndOfMonthTableAdapters.tblsales_payment_jurnal1TableAdapter
        Dim TBL As New DseTransaksiJournal_EndOfMonth.tblsales_payment_jurnal1DataTable
        Dim row As DseTransaksiJournal_EndOfMonth.tblsales_payment_jurnal1Row
        Dim TipeJurnal As ClsTransaksiJournal.EnumTipeJournal = EnumTipeJournal.PurchasePaymentJournal

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vStartDate, TipeJurnal, Core_Module.PrimaryBaseCurrency)

            i.Show()
            i.MinimumSize = New Drawing.Size(100, 100)
            i.Size = New Drawing.Size(100, 100)
            i.Location = New Point(-100, -100)

            For Each row In TBL.Rows
                i.AktifkanDetail()
                i.ObjJurnal.CariDetailPurchaseJournal(row.Primary, i.dgvItem, i, i.TipeJournal, i.DgvDifference)
                i.SalinkeField()
                i.UpdateJurnal(True)
                i.ObjJurnal.SimpanPerubahanPurchaseJournal1(i.dgvItem, i.ObjJurnal.CreatedBy, i, i.TipeJournal, i.dgvPurchaseBalance, i.DgvDifference)
            Next

            i.Hide()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Function GetNumberTransaksiAR(ByVal vTBLInv As DseTransaksiJournal_EndOfMonth.tbltrans_salesinvoiceDataTable, ByVal vTBLSR As DseTransaksiJournal_EndOfMonth.tbltrans_salesreturnDataTable, _
  ByVal vTBLInvManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalsalesDataTable, ByVal vTBLSRManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalsalesreturn_manualDataTable, _
  ByVal vTipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal, ByVal vPrimaryTransaksiAsal As Double, _
    ByVal vTBL_CNDN As DseTransaksiJournal_EndOfMonth.tbltrans_credit_noteDataTable) As String
        Dim Filter As String = ""
        Dim rowInv As DseTransaksiJournal_EndOfMonth.tbltrans_salesinvoiceRow
        Dim rowInvManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalsalesRow
        Dim rowSRManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalsalesreturn_manualRow
        Dim rowSR As DseTransaksiJournal_EndOfMonth.tbltrans_salesreturnRow
        Dim row_CNDN As DseTransaksiJournal_EndOfMonth.tbltrans_credit_noteRow

        Try
            Filter = "Primary = " + vPrimaryTransaksiAsal.ToString
            If vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual Then
                vTBLInvManual.DefaultView.RowFilter = Filter
                If vTBLInvManual.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowInvManual = vTBLInvManual.DefaultView(0).Row
                    Return rowInvManual.Invoice_Number
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis Then
                vTBLInv.DefaultView.RowFilter = Filter
                If vTBLInv.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowInv = vTBLInv.DefaultView(0).Row
                    Return rowInv.No_Transaksi
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual Then
                vTBLSRManual.DefaultView.RowFilter = Filter
                If vTBLSRManual.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowSRManual = vTBLSRManual.DefaultView(0).Row
                    Return rowSRManual.AOSR_Number
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis Then
                vTBLSR.DefaultView.RowFilter = Filter
                If vTBLSR.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowSR = vTBLSR.DefaultView(0).Row
                    Return rowSR.No_Transaksi
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Or _
            vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Then
                vTBL_CNDN.DefaultView.RowFilter = Filter
                If vTBL_CNDN.DefaultView.Count = 0 Then
                    Return ""
                Else
                    row_CNDN = vTBL_CNDN.DefaultView(0).Row
                    Return row_CNDN.no_transksi
                End If
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetNoTransaksi_ByDocumentAcuan(ByVal vTipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance, ByVal vPrimaryDocumentAcuan As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal3TableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnal3DataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnal3Row
        Dim RetVal As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vTipeBalance, vPrimaryDocumentAcuan)
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    If RetVal = "" Then
                        RetVal = row.Ref_Number
                    Else
                        RetVal = RetVal + ", " + row.Ref_Number
                    End If
                Next

                Return RetVal
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Private Function GetNumberTransaksiAP(ByVal vTBLAOP As DseTransaksiJournal_EndOfMonth.tbltrans_aopDataTable, ByVal vTBLPR As DseTransaksiJournal_EndOfMonth.tbltrans_purchasereturnDataTable, _
 ByVal vTBLAOPManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalpurchasebalancemanualDataTable, ByVal vTBLPRManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalpurchasereturnbalancemanualDataTable, _
 ByVal vTipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal, ByVal vPrimaryTransaksiAsal As Double, _
    ByVal vTBL_CNDN As DseTransaksiJournal_EndOfMonth.tbltrans_credit_noteDataTable) As String
        Dim Filter As String = ""
        Dim rowAOP As DseTransaksiJournal_EndOfMonth.tbltrans_aopRow
        Dim rowAOPManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalpurchasebalancemanualRow
        Dim rowPRManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalpurchasereturnbalancemanualRow
        Dim rowPR As DseTransaksiJournal_EndOfMonth.tbltrans_purchasereturnRow
        Dim row_CNDN As DseTransaksiJournal_EndOfMonth.tbltrans_credit_noteRow

        Try
            Filter = "Primary = " + vPrimaryTransaksiAsal.ToString
            If vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual Then
                vTBLAOPManual.DefaultView.RowFilter = Filter
                If vTBLAOPManual.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowAOPManual = vTBLAOPManual.DefaultView(0).Row
                    Return rowAOPManual.AOP_Number
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis Then
                vTBLAOP.DefaultView.RowFilter = Filter
                If vTBLAOP.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowAOP = vTBLAOP.DefaultView(0).Row
                    Return rowAOP.No_Transaksi
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual Then
                vTBLPRManual.DefaultView.RowFilter = Filter
                If vTBLPRManual.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowPRManual = vTBLPRManual.DefaultView(0).Row
                    Return rowPRManual.AOPR_Number
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis Then
                vTBLPR.DefaultView.RowFilter = Filter
                If vTBLPR.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowPR = vTBLPR.DefaultView(0).Row
                    Return rowPR.No_Transaksi
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Or _
            vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Then
                vTBL_CNDN.DefaultView.RowFilter = Filter
                If vTBL_CNDN.DefaultView.Count = 0 Then
                    Return ""
                Else
                    row_CNDN = vTBL_CNDN.DefaultView(0).Row
                    Return row_CNDN.no_transksi
                End If
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub UpdateJournalAR_AP_NonIDR_Hapus(ByVal vPrimaryJurnal As Double, ByVal Periode As Date, ByVal Frm As FrmTransaksiJournal_EndOfMonth, _
    ByVal vTanggalJurnal As Date)
        Dim DTABalance As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
        Dim DTABalance1 As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter

        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DefPrec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Dim DTAItemAR As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_revaluasi_arTableAdapter
        Dim TBLItemAR As New DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_arDataTable
        Dim rowItemAR As DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_arRow

        Dim DTAItemAP As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_revaluasi_apTableAdapter
        Dim TBLItemAP As New DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_apDataTable
        Dim rowItemAP As DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_apRow

        Try
            DTABalance.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemAR.Connection.ConnectionString = DTABalance.Connection.ConnectionString
            DTAItemAP.Connection.ConnectionString = DTABalance.Connection.ConnectionString
            DTABalance1.Connection.ConnectionString = DTABalance.Connection.ConnectionString
            DefPrec = ObjCur.GetDecimal_DefaultCurrency

            'balance
            DTAItemAR.FillBy_PrimaryJurnal(TBLItemAR, vPrimaryJurnal)
            For Each rowItemAR In TBLItemAR.Rows
                With rowItemAR
                    'update balance ar untuk transaksi ybs, ganti exchange rate dan hitung ulang saldo idr
                    DTABalance.UpdateQuery_Revaluate(.exchange_rate_lama, _
                    DefPrec, ObjKartu.GetTipeSalesBalance(.tipe_transaksi_asal), _
                    .primary_transaksi_asal, .primary_currency)
                End With
            Next
            DTAItemAP.FillBy_PrimaryTransaksi(TBLItemAP, vPrimaryJurnal)
            For Each rowItemAP In TBLItemAP.Rows
                With rowItemAP
                    'update balance ar untuk transaksi ybs, ganti exchange rate dan hitung ulang saldo idr
                    DTABalance1.UpdateQuery_Revaluate(.exchange_rate_lama, _
                    DefPrec, ObjKartu.GetTipeSalesBalance(.tipe_transaksi_asal), _
                    .primary_transaksi_asal, .primary_currency)
                End With
            Next

            'update kartu ar, hapus item kartu ap ar dari transaksi revaluasi
            ObjKartu.Hapus(vPrimaryJurnal, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Revaluation)

            'update jurnal untuk semua transaksi sales payment dengan tanggal 1 hari setelah revaluatioin
            Me.UpdateJurnalAR(vTanggalJurnal + TimeSpan.FromDays(1))
            Me.UpdateJurnalAP(vTanggalJurnal + TimeSpan.FromDays(1))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTABalance.Dispose()
            ObjKartu = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
            DTAItemAR.Dispose()
            TBLItemAR.Dispose()
            DTAItemAP.Dispose()
            TBLItemAP.Dispose()
            DTABalance1.Dispose()
        End Try
    End Sub

    Public Overridable Function SimpanPerubahan_EndOfMonth(ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal_EndOfMonth, ByVal TipeJournal As Integer, _
    ByVal vDGVCurrency As DataGridView, ByVal vDGVItemAR As DataGridView, ByVal vDGVItemAP As DataGridView) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()

            'ambil no antrian
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_EOM_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1_EndOfMonth(vDGVItem, vUserCode, frm, TipeJournal, vDGVCurrency, vDGVItemAR, vDGVItemAP)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1_EndOfMonth(ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal_EndOfMonth, ByVal TipeJournal As Integer, _
    ByVal vDGVCurrency As DataGridView, ByVal vDGVItemAR As DataGridView, ByVal vDGVItemAP As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Me.mTipeJournal = TipeJournal
            Me.mUrutan = -1
            Me.mPrimaryDokumenAcuan = -1
            Me.mPrimaryPihakAcuan = -1
            If Me.ApakahDataBenar1(False, vDGVItem, frm.ColPrimaryCurrency.Index, Me.Tanggal) = True Then
                If DTA.Update(Me.RefNumber, Me.Tanggal, Me.mUrutan, Me.mTipeJournal, Me.mPrimaryDokumenAcuan, _
                Me.mPrimaryPihakAcuan, Me.Description, vUserCode, Now, Now, _
                Me.mPrimaryAkunPayTo, Me.mPrimaryAkunDifference, Me.mPaid, Me.mExchangeRate, Me.mPaid_IDR, Me.mARPaid_IDR, Me.mDifference, _
                Me.mExchageRateGainLoss, Me.mTipeJurnal2, Me.mTipeInOut, Me.mPrimaryCurr, Me.mReferenceParty, Me.mPaymentMethod, Me.Primary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)

                    'simpan data item baru
                    ObjItem.SimpanItem_EndOfMonth(vDGVItem, Me.mPrimary, frm, vDGVCurrency, vDGVItemAR, vDGVItemAP)

                    'update data journal ar non-idr, setelah bulan ini
                    Me.UpdateJournalAR_AP_NonIDR_Hapus(Me.mPrimary, Me.mDate, frm, Me.mDate)
                    Me.UpdateJournalAR_AP_NonIDR_Simpan(Me.mPrimary, Me.mDate, frm, vDGVCurrency, vDGVItemAR, Me.mDate, vDGVItemAP)

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
                Return Me.SimpanPerubahan1_EndOfMonth(vDGVItem, vUserCode, frm, TipeJournal, vDGVCurrency, vDGVItemAR, vDGVItemAP)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahanSalesJournal(ByVal vDGVItem As DataGridView, _
       ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal_Sales, ByVal vTipeJournal As Integer, _
       ByVal vDGVSalesBalance As DataGridView, ByVal vDGVItemOther As DataGridView) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()

            'ambil no antrian
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_Sales_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahanSalesJournal1(vDGVItem, vUserCode, frm, vTipeJournal, vDGVSalesBalance, vDGVItemOther)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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
        End Try
    End Function

    ''Untuk Sales Journal
    Public Overridable Function SimpanPerubahanSalesJournal1(ByVal vDGVItem As DataGridView, _
       ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal_Sales, ByVal vTipeJournal As Integer, _
       ByVal vDGVSalesBalance As DataGridView, ByVal vDGVAkunOther As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item
        Dim ObjCustom As New CustomSMESS.ClsTransaksiJournal

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Me.mTipeJournal = TipeJournal
            Me.mUrutan = -1
            Me.mPrimaryDokumenAcuan = -1
            If Me.ApakahDataBenar1(False, vDGVItem, frm.ColPrimaryCurrency.Index, Me.Tanggal, vDGVSalesBalance, frm.ColPrimaryCurrencyBalance.Index, _
            frm.ColReceivementOriginal.Index, frm.ColTipeDocumentAcuan_Sales.Index, frm.ColPrimaryDocumenAcuan_Sales.Index, frm.ColInvoiceNo.Index, _
            Me.mPrimaryPihakAcuan, frm.ColExhangeRateBalance.Index, frm.ColInvoiceDate.Index) = True Then
                'pastikan semua item payment, original balancenya- sama dengan original balance yang ada didalam database,
                'untuk memastikan data sales balance selalu konsisten
                If Me.CekKonsistensiSalesBalance(vDGVSalesBalance, frm) = False Then
                    Return False
                End If

                If DTA.Update(Me.RefNumber, Me.Tanggal, Me.mUrutan, Me.mTipeJournal, Me.mPrimaryDokumenAcuan, _
                Me.mPrimaryPihakAcuan, Me.Description, vUserCode, Now, Now, _
                Me.mPrimaryAkunPayTo, Me.mPrimaryAkunDifference, Me.mPaid, Me.mExchangeRate, Me.mPaid_IDR, Me.mARPaid_IDR, Me.mDifference, _
                Me.mExchageRateGainLoss, Me.mTipeJurnal2, Me.mTipeInOut, Me.mPrimaryCurr, Me.mReferenceParty, Me.mPaymentMethod, Me.Primary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama
                    ObjCustom.FollowUpHapus_SalesJournal(Me.Primary)
                    ObjItem.HapusItem_SalesJurnal(Me.mPrimary, vTipeJournal)

                    'simpan data item baru
                    ObjItem.SimpanItemSalesJournal(vDGVItem, Me.mPrimary, frm, vDGVSalesBalance, _
                    vTipeJournal, Me.mRefNumber, Me.mPrimaryPihakAcuan, Me.mDate, vDGVAkunOther)

                    ObjCustom.FollowUpSimpan_SalesJournal(Me.mPrimary)

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
                Return Me.SimpanPerubahanSalesJournal1(vDGVItem, vUserCode, frm, vTipeJournal, vDGVSalesBalance, vDGVAkunOther)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjCustom = Nothing
        End Try
    End Function

    ''Untuk Purchase Journal
    ''
    Public Overridable Function SimpanPerubahanPurchaseJournal(ByVal vDGVItem As DataGridView, _
       ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal_Purchase, ByVal vTipeJournal As Integer, _
       ByVal vDGVPurchaseBalance As DataGridView, ByVal vDGVAkunOther As DataGridView) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()

            'ambil no antrian
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_Purchase_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahanPurchaseJournal1(vDGVItem, vUserCode, frm, vTipeJournal, vDGVPurchaseBalance, vDGVAkunOther)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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
        End Try
    End Function

    Public Overridable Function SimpanPerubahanPurchaseJournal1(ByVal vDGVItem As DataGridView, _
       ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal_Purchase, ByVal vTipeJournal As Integer, _
       ByVal vDGVPurchaseBalance As DataGridView, ByVal vDGVAkunOther As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim DTA2 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_NewTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Me.mTipeJournal = vTipeJournal
            Me.mUrutan = -1
            Me.mPrimaryDokumenAcuan = -1

            If Me.ApakahDataBenar1(False, vDGVItem, frm.ColPrimaryCurrency.Index, Me.Tanggal, vDGVPurchaseBalance, frm.ColPrimaryCurrencyBalance.Index, _
            frm.ColPaymentOriginal.Index, frm.ColTipeDocumentAcuan_Purchase.Index, frm.ColPrimaryDocumentAcuan_Purchase.Index, frm.ColAOPNumber.Index, _
            Me.PrimaryPihakAcuan, frm.ColExchangeRateBalance.Index, frm.ColAOPDate.Index) = True Then
                'cek konsistensi purchase balance
                If Me.CekKonsistensiPurchaseBalance(vDGVPurchaseBalance, frm) = False Then
                    Return False
                End If

                If DTA2.UpdateQuery(Me.RefNumber, Me.Tanggal, Me.mUrutan, Me.mTipeJournal, Me.mPrimaryDokumenAcuan, _
                Me.mPrimaryPihakAcuan, Me.Description, vUserCode, Now, Now, _
                Me.mPrimaryAkunPayTo, Me.mPrimaryAkunDifference, Me.mPaid, Me.mExchangeRate, Me.mPaid_IDR, Me.mARPaid_IDR, Me.mDifference, _
                Me.mExchageRateGainLoss, Me.mTipeJurnal2, Me.mTipeInOut, Me.mPrimaryCurr, Me.mReferenceParty, Me.mPaymentMethod, _
                Me.mPrimaryExChangerateGainLoss, Me.Primary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama
                    ObjItem.HapusItem_PurchaseJurnal(Me.mPrimary, vTipeJournal)

                    'simpan data item baru
                    ObjItem.SimpanItemPurchaseJournal(vDGVItem, Me.mPrimary, frm, vDGVPurchaseBalance, _
                    vTipeJournal, Me.mRefNumber, Me.mPrimaryPihakAcuan, Me.mDate, vDGVAkunOther)

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
                Return Me.SimpanPerubahanPurchaseJournal1(vDGVItem, vUserCode, frm, vTipeJournal, vDGVPurchaseBalance, vDGVAkunOther)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, _
      ByVal StrictSalesStaffDelete As Boolean, ByVal usercode As String, ByVal VFromRecurring As Boolean) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()

            'ambil no antrian
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_General_Hapus, usercode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, StrictSalesStaffDelete, usercode, VFromRecurring)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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
        End Try
    End Function

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As Boolean, ByVal usercode As String, ByVal vFromRecurring As Boolean) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, usercode, vFromRecurring) = True Then
                'hapus data header, data item otomatis dihapuskan oleh server
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

    Public Function Hapus_EndOFMonthJournal(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As Boolean, ByVal usercode As String, ByVal Frm As FrmTransaksiJournal_EndOfMonth, _
    ByVal vDGVCurrency As DataGridView, ByVal vDGVItemAR As DataGridView) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()

            'ambil no antrian
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_EOM_Hapus, usercode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus_EndOFMonthJournal1(vPrimaryTransaksi, StrictSalesStaffDelete, usercode, Frm, vDGVCurrency, vDGVItemAR)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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
        End Try
    End Function

    Public Overridable Function Hapus_EndOFMonthJournal1(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As Boolean, ByVal usercode As String, ByVal Frm As FrmTransaksiJournal_EndOfMonth, _
    ByVal vDGVCurrency As DataGridView, ByVal vDGVItemAR As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter

        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, usercode, False) = True Then
                If Me.ApakahBolehHapus_EndOFMonthJournal(vPrimaryTransaksi, StrictSalesStaffDelete, usercode) = True Then
                    'update data journal ar non-idr, setelah bulan ini
                    Me.UpdateJournalAR_AP_NonIDR_Hapus(Me.mPrimary, Me.mDate, Frm, Me.GetTanggalJournal(vPrimaryTransaksi))

                    'hapus data header, data item otomatis dihapuskan oleh server
                    If DTA.Delete(vPrimaryTransaksi) > 0 Then
                        Return True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus_SalesJurnal(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As Boolean, ByVal usercode As String, ByVal vTipeJournal As ClsTransaksiJournal.EnumTipeJournal) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()

            'ambil no antrian
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_Sales_Hapus, usercode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus_SalesJurnal1(vPrimaryTransaksi, StrictSalesStaffDelete, usercode, vTipeJournal)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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
        End Try
    End Function

    Public Overridable Function Hapus_SalesJurnal1(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As Boolean, ByVal usercode As String, ByVal vTipeJournal As ClsTransaksiJournal.EnumTipeJournal) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item
        Dim ObjCustom As New CustomSMESS.ClsTransaksiJournal

        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, usercode, False) = True Then
                ObjCustom.FollowUpHapus_SalesJournal(vPrimaryTransaksi)

                'hapuskan data item
                ObjItem.HapusItem_SalesJurnal(vPrimaryTransaksi, vTipeJournal)
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjCustom = Nothing
        End Try
    End Function

    Public Function Hapus_PurchaseJurnal(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As Boolean, ByVal usercode As String, _
    ByVal vTipeJournal As ClsTransaksiJournal.EnumTipeJournal) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()

            'ambil no antrian
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_Purchase_Hapus, usercode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus_PurchaseJurnal1(vPrimaryTransaksi, StrictSalesStaffDelete, usercode, vTipeJournal)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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
        End Try
    End Function

    Public Overridable Function Hapus_PurchaseJurnal1(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As Boolean, ByVal usercode As String, _
    ByVal vTipeJournal As ClsTransaksiJournal.EnumTipeJournal) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item

        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, usercode, False) = True Then
                'hapuskan data item
                ObjItem.HapusItem_PurchaseJurnal(vPrimaryTransaksi, vTipeJournal)
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function ApakaJurnalNonIDR(ByVal vPrimaryAkun As Double) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter

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

    Public Function CetakJournal(ByVal vStartDate As Date, ByVal vEndDate As Date, _
      ByVal TBL As DseTransaksiJournal_Cetak.TblCetak_JurnalDataTable, ByVal CRV As _
      CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vclb As Windows.Forms.CheckedListBox, ByVal vOriginal As Boolean, _
      ByVal vShowItemRemarks As Boolean, _
      ByVal vJournalDesc As String, ByVal vRemarksItemDesc As String, ByVal vDGVFilter As DataGridView, ByVal vFrm As FrmTransaksiJournal_View, ByVal vUserCode As String) As Boolean
        Dim row As DseTransaksiJournal_Cetak.TblCetak_JurnalRow
        Dim RowCetak As DseTransaksiJournal_Cetak.DataTable1Row
        Dim Filter As String = ""
        Dim DataView As DataRowView
        Dim Report As New CrpCetakTransaksiJournal
        Dim DS As New DseTransaksiJournal_Cetak
        Dim DTA As New DseTransaksiJournal_CetakTableAdapters.TblCetak_JurnalTableAdapter
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DefaultPrecision As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim FormatNegatif As Integer = 0
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim DefCur As String = ""
        Dim ListPrimaryJurnal As New List(Of Double)
        Dim Lanjut As Boolean = True
        Dim TipeJurnal1 As EnumTipeJournal = EnumTipeJournal.GeneralJournal
        Dim TipeJurnal2 As EnumTipeJurnal2 = EnumTipeJurnal2.Cash_n_Bank
        Dim ListCreatedBy As New List(Of String)
        Dim TipeAccess As EnumJournalAccesType = EnumJournalAccesType.FullAccess

        Try
            vDGVFilter.AllowUserToAddRows = False
            FormatDecimal = ObjCur.GetDecimal_DefaultCurrency
            DefCur = ObjCur.GetCurrencyCode(1)

            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) <> "" Then
                FormatNegatif = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            Else
                FormatNegatif = 1
            End If

            DTA.FillBy_TanggalDariSampai(TBL, vEndDate, vStartDate)

            If vDGVFilter.RowCount > 0 Then
                Filter = ""
                For Each rowdgv As DataGridViewRow In vDGVFilter.Rows
                    If Filter = "" Then
                        Filter = "[Primary Akun] = " + rowdgv.Cells(vFrm.ColPrimaryAkun.Index).Value.ToString
                    Else
                        Filter = Filter + " or [Primary Akun] = " + rowdgv.Cells(vFrm.ColPrimaryAkun.Index).Value.ToString
                    End If
                Next
                TBL.DefaultView.RowFilter = Filter
                For Each DataView In TBL.DefaultView
                    row = DataView.Row
                    If ListPrimaryJurnal.Contains(row.Primary) = False Then
                        ListPrimaryJurnal.Add(row.Primary)
                    End If
                Next
            End If

            Dim CheckedItem As String = ""
            Filter = ""
            For i As Integer = 0 To vclb.CheckedItems.Count - 1
                CheckedItem = vclb.CheckedItems.Item(i)
                If CheckedItem = "General Journal" Then
                    'general journal
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.Jurnal_NoGrouping
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Cash And Bank" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.Cash_n_Bank
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "End of Month Journal" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Sales Journal" Then
                    TipeJurnal1 = EnumTipeJournal.SalesJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Sales Return Journal" Then
                    TipeJurnal1 = EnumTipeJournal.SalesReturnJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Sales Down Payment" Then
                    TipeJurnal1 = EnumTipeJournal.SalesDownPayment
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase Journal" Then
                    TipeJurnal1 = EnumTipeJournal.PurchaseJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase Return Journal" Then
                    TipeJurnal1 = EnumTipeJournal.PurchaseReturnJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase DownPayment Journal" Then
                    TipeJurnal1 = EnumTipeJournal.PurchaseDownPaymentJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Sales Receivement Journal" Then
                    TipeJurnal1 = EnumTipeJournal.SalesReceivementJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase Payment Journal" Then
                    TipeJurnal1 = EnumTipeJournal.PurchasePaymentJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Credit Note" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.Credit_Note
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Debit Note" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.Debit_Note
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                End If
            Next
            If vJournalDesc.EndsWith("%") = False Then
                vJournalDesc = vJournalDesc + "%"
            End If
            If vRemarksItemDesc.EndsWith("%") = False Then
                vRemarksItemDesc = vRemarksItemDesc + "%"
            End If
            vJournalDesc = vJournalDesc.Replace("%", "*")
            vRemarksItemDesc = vRemarksItemDesc.Replace("%", "*")
            If vJournalDesc <> "" Then
                Filter = "Desc like '" + vJournalDesc + "' AND " + "(" + Filter + ")"
                'filter = "Desc like '" + vJournalDesc + "'"
            End If
            If vRemarksItemDesc <> "" Then
                Filter = "Remarks like '" + vRemarksItemDesc + "' AND " + Filter
            End If

            TBL.DefaultView.RowFilter = Filter
            TBL.DefaultView.Sort = "Date, [Ref Number], [Primary Item]"

            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Jurnal_IsLedgerAccessActive) = "1" Then
                TipeAccess = Me.GetLedgerAccess(vUserCode)
            Else
                TipeAccess = EnumJournalAccesType.FullAccess
            End If
            If TipeAccess = EnumJournalAccesType.Respected_User_Only Or _
            TipeAccess = EnumJournalAccesType.Selected_Users_Only Then
                Me.GetLedgerAccess(vUserCode, ListCreatedBy)
            End If

            If TBL.DefaultView.Count > 0 Then

                For Each DataView In TBL.DefaultView
                    row = DataView.Row
                    Lanjut = False
                    If vDGVFilter.RowCount = 0 Then
                        Lanjut = True
                    Else
                        If ListPrimaryJurnal.Contains(row.Primary) = True Then
                            Lanjut = True
                        End If
                    End If
                    If Lanjut = True Then
                        If TipeAccess = EnumJournalAccesType.Respected_User_Only Or _
                        TipeAccess = EnumJournalAccesType.Selected_Users_Only Then
                            If ListCreatedBy.Contains(row.Created_By) = False Then
                                Lanjut = False
                            End If
                        End If
                    End If

                    If Lanjut = True Then
                        RowCetak = DS.DataTable1.NewDataTable1Row
                        RowCetak.Primary = row.Primary
                        RowCetak.No_Jurnal = row.Ref_Number
                        If vOriginal = True Then
                            RowCetak.Debet = row.Debet_Original
                            RowCetak.Kredit = row.Credit_Original
                            RowCetak.Currency_Code = row.Currency_Code
                            RowCetak.Def_Prec_Cur = row.Default_Precision
                        Else
                            RowCetak.Debet = row.Debet
                            RowCetak.Kredit = row.Kredit
                            RowCetak.Currency_Code = DefCur
                            RowCetak.Def_Prec_Cur = FormatDecimal
                        End If
                        If vShowItemRemarks = True And row.Remarks <> "" Then
                            RowCetak.Account = row.Account + vbCrLf + row.Remarks
                        Else
                            RowCetak.Account = row.Account
                        End If
                        RowCetak.Tanggal = row._Date
                        RowCetak.Desc = row.Desc
                        DS.DataTable1.Rows.Add(RowCetak)
                    End If
                Next
            End If
            TBL.DefaultView.RowFilter = ""

            ''set grouping by Tanggal
            Dim Tanggal As Date = #1/1/1950#
            Dim JurnalNumber As String = ""
            For ij As Integer = 0 To DS.DataTable1.Rows.Count - 1
                RowCetak = DS.DataTable1.Rows(ij)
                If RowCetak.Tanggal <> Tanggal Then
                    Tanggal = RowCetak.Tanggal
                    RowCetak.Suppress_Tanggal = 0
                Else
                    RowCetak.Suppress_Tanggal = 1
                End If
                If RowCetak.No_Jurnal <> JurnalNumber Then
                    JurnalNumber = RowCetak.No_Jurnal
                    RowCetak.Suppress_Jurnal = 0
                Else
                    RowCetak.Suppress_Jurnal = 1
                End If
            Next

            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_Format_Day.ParameterFieldName, FormatDay)
            Report.SetParameterValue(Report.Parameter_Format_Month.ParameterFieldName, FormatMonth)
            Report.SetParameterValue(Report.Parameter_Format_Year.ParameterFieldName, FormatYears)
            Report.SetParameterValue(Report.Parameter_TanggalDari.ParameterFieldName, vStartDate.Date)
            Report.SetParameterValue(Report.Parameter_TanggalSampai.ParameterFieldName, vEndDate.Date)
            Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            Report.SetParameterValue(Report.Parameter_FormatNegatif.ParameterFieldName, FormatNegatif)

            CRV.ReportSource = Report
            CRV.Show()
            vDGVFilter.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            ObjCur = Nothing
            ObjSetting = Nothing
            ListPrimaryJurnal = Nothing
            vDGVFilter.AllowUserToAddRows = True
            ListCreatedBy = Nothing
        End Try
    End Function

    Public Function CetakJournal_DetailInvoice_Purchase(ByVal vPrimary As Double, _
   ByVal Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal DS As DseTransaksiJournal, ByVal vDGV As DataGridView, _
    ByVal vFrm As FrmTransaksiJournal_Purchase) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tblCetakJurnalTableAdapter
        Dim TBL As DseTransaksiJournal.tblCetakJurnalDataTable = DS.tblCetakJurnal
        Dim FormatDecimal As Integer = 0
        Dim Negatif As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Report As New CrpCetakJurnal_DetailInv
        Dim Report2 As New CrpCetakJurnal2_DetailInv
        Dim Report3 As New CrpCetakJurnal3_DetailInv
        Dim CetakJurnalPortrait As String = ""
        Dim KodeCur As String = ""
        Dim row As DseTransaksiJournal.tblCetakJurnalRow
        Dim rowDGV As DataGridViewRow
        Dim JmlBaris As Integer = 0

        Dim rowDetail As DseTransaksiJournal.TBLDetailInvoiceRow

        Try
            'isikan data detail invoice
            DS.TBLDetailInvoice.Rows.Clear()
            For Each rowDGV In vDGV.Rows
                If CDec(rowDGV.Cells(vFrm.ColPaymentOriginal.Index).Value) <> 0 Then
                    rowDetail = DS.TBLDetailInvoice.NewTBLDetailInvoiceRow()

                    rowDetail.Currency = rowDGV.Cells(vFrm.ColCurrencyBalance.Index).Value
                    rowDetail.No_Invoice = rowDGV.Cells(vFrm.ColSupplierInvoiceNumber.Index).Value
                    If rowDetail.No_Invoice = "" Then
                        rowDetail.No_Invoice = rowDGV.Cells(vFrm.ColAOPNumber.Index).Value
                    End If
                    rowDetail.Payment = rowDGV.Cells(vFrm.ColPaymentOriginal.Index).Value
                    rowDetail.Tanggal_Invoice = rowDGV.Cells(vFrm.ColAOPDate.Index).Value

                    DS.TBLDetailInvoice.Rows.Add(rowDetail)
                End If
            Next


            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimary)
            KodeCur = ObjCur.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)

            JmlBaris = DS.TBLDetailInvoice.Rows.Count

            CetakJurnalPortrait = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CetakJurnalPortrait)

            If CetakJurnalPortrait = "" Or CetakJurnalPortrait = "0" Then
                If Me.ApakaJurnalNonIDR(vPrimary) = False Then
                    'cetak lanscape tanda currency no idr
                    Report2.SetDataSource(DS)
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
                    Report2.SetParameterValue(Report2.Parameter_Judul.ParameterFieldName, "PURCHASE PAYMENT JOURNAL")
                    Report2.SetParameterValue(Report2.Parameter_JmlBaris.ParameterFieldName, JmlBaris)

                    Crv.ReportSource = Report2
                    Crv.Zoom(1)
                    Crv.Show()
                Else
                    Report3.SetDataSource(DS)
                    For Each row In DS.tblCetakJurnal.Rows
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
                    Report3.SetParameterValue(Report3.Parameter_Judul.ParameterFieldName, "PURCHASE PAYMENT JOURNAL")
                    Report3.SetParameterValue(Report3.Parameter_JmlBaris.ParameterFieldName, JmlBaris)

                    Crv.ReportSource = Report3
                    Crv.Zoom(1)
                    Crv.Show()
                End If
            Else
                'cetak portrait
                Report.SetDataSource(DS)
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
                Report.SetParameterValue(Report.Parameter_Judul.ParameterFieldName, "PURCHASE PAYMENT JOURNAL")
                Report.SetParameterValue(Report.Parameter_JmlBaris.ParameterFieldName, JmlBaris)

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

    Public Function CetakJournal_DetailInvoice_Sales(ByVal vPrimary As Double, _
  ByVal Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal DS As DseTransaksiJournal, ByVal vDGV As DataGridView, _
   ByVal vFrm As FrmTransaksiJournal_Sales) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tblCetakJurnalTableAdapter
        Dim TBL As DseTransaksiJournal.tblCetakJurnalDataTable = DS.tblCetakJurnal
        Dim FormatDecimal As Integer = 0
        Dim Negatif As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Report As New CrpCetakJurnal_DetailInv
        Dim Report2 As New CrpCetakJurnal2_DetailInv
        Dim Report3 As New CrpCetakJurnal3_DetailInv
        Dim CetakJurnalPortrait As String = ""
        Dim KodeCur As String = ""
        Dim row As DseTransaksiJournal.tblCetakJurnalRow
        Dim rowDGV As DataGridViewRow

        Dim rowDetail As DseTransaksiJournal.TBLDetailInvoiceRow

        Try
            'isikan data detail invoice
            DS.TBLDetailInvoice.Rows.Clear()
            For Each rowDGV In vDGV.Rows
                If CDec(rowDGV.Cells(vFrm.ColReceivementOriginal.Index).Value) <> 0 Then
                    rowDetail = DS.TBLDetailInvoice.NewTBLDetailInvoiceRow()

                    rowDetail.Currency = rowDGV.Cells(vFrm.ColCurrencyBalance.Index).Value
                    rowDetail.No_Invoice = rowDGV.Cells(vFrm.ColInvoiceNo.Index).Value
                    rowDetail.Payment = rowDGV.Cells(vFrm.ColReceivementOriginal.Index).Value
                    rowDetail.Tanggal_Invoice = rowDGV.Cells(vFrm.ColInvoiceDate.Index).Value

                    DS.TBLDetailInvoice.Rows.Add(rowDetail)
                End If
            Next

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimary)
            KodeCur = ObjCur.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)

            CetakJurnalPortrait = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CetakJurnalPortrait)

            If CetakJurnalPortrait = "" Or CetakJurnalPortrait = "0" Then
                If Me.ApakaJurnalNonIDR(vPrimary) = False Then
                    'cetak lanscape tanda currency no idr
                    Report2.SetDataSource(DS)
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
                    Report2.SetParameterValue(Report2.Parameter_Judul.ParameterFieldName, "SALES RECEIVEMENT JOURNAL")

                    Crv.ReportSource = Report2
                    'Crv.Zoom(1)
                    Crv.Show()
                Else
                    Report3.SetDataSource(DS)
                    For Each row In DS.tblCetakJurnal.Rows
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
                    Report3.SetParameterValue(Report3.Parameter_Judul.ParameterFieldName, "SALES RECEIVEMENT JOURNAL")

                    Crv.ReportSource = Report3
                    'Crv.Zoom(1)
                    Crv.Show()
                End If
            Else
                'cetak portrait
                Report.SetDataSource(DS)
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
                Report.SetParameterValue(Report.Parameter_Judul.ParameterFieldName, "SALES RECEIVEMENT JOURNAL")

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

    Public Function CetakJournal_OriginalCurrency(ByVal vStartDate As Date, ByVal vEndDate As Date, _
    ByVal TBL As DseTransaksiJournal_Cetak.TblCetak_JurnalDataTable, ByVal CRV As _
    CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vclb As Windows.Forms.CheckedListBox, ByVal vShowItemRemarks As Boolean, _
    ByVal vJournalDesc As String, ByVal vRemarksItemDesc As String, ByVal vDGVFilter As DataGridView, ByVal vFrm As FrmTransaksiJournal_View, _
    ByVal vUserCode As String) As Boolean
        Dim row As DseTransaksiJournal_Cetak.TblCetak_JurnalRow
        Dim RowCetak As DseTransaksiJournal_Cetak.DataTable1Row
        Dim Filter As String = ""
        Dim DataView As DataRowView
        Dim Report As New CrpCetakTransaksiJournal2
        Dim DS As New DseTransaksiJournal_Cetak
        Dim DTA As New DseTransaksiJournal_CetakTableAdapters.TblCetak_JurnalTableAdapter
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DefaultPrecision As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim FormatNegatif As Integer = 0
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim ListPrimaryJurnal As New List(Of Double)
        Dim Lanjut As Boolean = True
        Dim TipeJurnal1 As EnumTipeJournal = EnumTipeJournal.GeneralJournal
        Dim TipeJurnal2 As EnumTipeJurnal2 = EnumTipeJurnal2.Cash_n_Bank
        Dim ListCreatedBy As New List(Of String)
        Dim TipeAccess As EnumJournalAccesType = EnumJournalAccesType.FullAccess

        Try
            vDGVFilter.AllowUserToAddRows = False

            FormatDecimal = ObjCur.GetDecimal_DefaultCurrency
            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) <> "" Then
                FormatNegatif = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            Else
                FormatNegatif = 1
            End If

            DTA.FillBy_TanggalDariSampai(TBL, vEndDate, vStartDate)

            If vDGVFilter.RowCount > 0 Then
                Filter = ""
                For Each rowdgv As DataGridViewRow In vDGVFilter.Rows
                    If Filter = "" Then
                        Filter = "[Primary Akun] = " + rowdgv.Cells(vFrm.ColPrimaryAkun.Index).Value.ToString
                    Else
                        Filter = Filter + " or [Primary Akun] = " + rowdgv.Cells(vFrm.ColPrimaryAkun.Index).Value.ToString
                    End If
                Next
                TBL.DefaultView.RowFilter = Filter
                For Each DataView In TBL.DefaultView
                    row = DataView.Row
                    If ListPrimaryJurnal.Contains(row.Primary) = False Then
                        ListPrimaryJurnal.Add(row.Primary)
                    End If
                Next
            End If

            Dim CheckedItem As String = ""
            Filter = ""
            For i As Integer = 0 To vclb.CheckedItems.Count - 1
                CheckedItem = vclb.CheckedItems.Item(i)
                If CheckedItem = "General Journal" Then
                    'general journal
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.Jurnal_NoGrouping
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Cash And Bank" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.Cash_n_Bank
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "End of Month Journal" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Loan Journal" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.HR_Payroll_Loan
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Loan Payment Journal" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.HR_Payroll_Loan_Payment
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Salary Payment Journal" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.HR_Payroll_Payment
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Sales Journal" Then
                    TipeJurnal1 = EnumTipeJournal.SalesJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Sales Return Journal" Then
                    TipeJurnal1 = EnumTipeJournal.SalesReturnJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Sales Down Payment" Then
                    TipeJurnal1 = EnumTipeJournal.SalesDownPayment
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase Journal" Then
                    TipeJurnal1 = EnumTipeJournal.PurchaseJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase Return Journal" Then
                    TipeJurnal1 = EnumTipeJournal.PurchaseReturnJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase DownPayment Journal" Then
                    TipeJurnal1 = EnumTipeJournal.PurchaseDownPaymentJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Sales Receivement Journal" Then
                    TipeJurnal1 = EnumTipeJournal.SalesReceivementJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase Payment Journal" Then
                    TipeJurnal1 = EnumTipeJournal.PurchasePaymentJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Credit Note" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.Credit_Note
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Debit Note" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.Debit_Note
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                End If
            Next

            'pastikan tidak ada % atau * ditengah2
            If vJournalDesc.Length > 0 Then
                If vJournalDesc.IndexOf("%", 1) - 1 Then
                    If vJournalDesc.StartsWith("%") Then
                        vJournalDesc = "%" + vJournalDesc.Replace("%", "")
                    Else
                        vJournalDesc = vJournalDesc.Replace("%", "")
                    End If
                End If
                If vJournalDesc.IndexOf("*", 1) - 1 Then
                    If vJournalDesc.StartsWith("*") = True Then
                        vJournalDesc = "*" + vJournalDesc.Replace("*", "")
                    Else
                        vJournalDesc = vJournalDesc.Replace("*", "")
                    End If
                End If
            End If
            If vRemarksItemDesc.Length > 0 Then
                If vRemarksItemDesc.IndexOf("%", 1) <> -1 Then
                    If vRemarksItemDesc.StartsWith("%") = True Then
                        vRemarksItemDesc = "%" + vRemarksItemDesc.Replace("%", "")
                    Else
                        vRemarksItemDesc = vRemarksItemDesc.Replace("%", "")
                    End If
                End If
                If vRemarksItemDesc.IndexOf("*", 1) - 1 Then
                    If vRemarksItemDesc.StartsWith("*") = True Then
                        vRemarksItemDesc = "*" + vRemarksItemDesc.Replace("*", "")
                    Else
                        vRemarksItemDesc = vRemarksItemDesc.Replace("*", "")
                    End If
                End If
            End If

            If vJournalDesc.EndsWith("%") = False And vJournalDesc.Length > 0 Then
                vJournalDesc = vJournalDesc + "%"
            End If
            If vRemarksItemDesc.EndsWith("%") = False And vRemarksItemDesc.Length > 0 Then
                vRemarksItemDesc = vRemarksItemDesc + "%"
            End If
            If vJournalDesc <> "" Then
                Filter = "Desc like  '" + vJournalDesc + "' AND " + "(" + Filter + ")"
                'filter = "Desc like '" + vJournalDesc + "'"
            End If
            If vRemarksItemDesc <> "" Then
                Filter = "Remarks like '" + vRemarksItemDesc + "' And " + Filter
            End If
            TBL.DefaultView.RowFilter = Filter
            TBL.DefaultView.Sort = "Date, [Ref Number], [Primary Item]"

            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
             Core_Module.ClsTransaksiSetting.EnumNamaParameter.Jurnal_IsLedgerAccessActive) = "1" Then
                TipeAccess = Me.GetLedgerAccess(vUserCode)
            Else
                TipeAccess = EnumJournalAccesType.FullAccess
            End If
            If TipeAccess = EnumJournalAccesType.Respected_User_Only Or _
            TipeAccess = EnumJournalAccesType.Selected_Users_Only Then
                Me.GetLedgerAccess(vUserCode, ListCreatedBy)
            End If

            If TBL.DefaultView.Count > 0 Then
                For Each DataView In TBL.DefaultView
                    row = DataView.Row

                    Lanjut = False
                    If vDGVFilter.RowCount = 0 Then
                        Lanjut = True
                    Else
                        If ListPrimaryJurnal.Contains(row.Primary) = True Then
                            Lanjut = True
                        End If
                    End If
                    If Lanjut = True Then
                        If TipeAccess = EnumJournalAccesType.Respected_User_Only Or _
                        TipeAccess = EnumJournalAccesType.Selected_Users_Only Then
                            If ListCreatedBy.Contains(row.Created_By) = False Then
                                Lanjut = False
                            End If
                        End If
                    End If

                    If Lanjut = True Then
                        RowCetak = DS.DataTable1.NewDataTable1Row
                        RowCetak.Primary = row.Primary
                        RowCetak.No_Jurnal = row.Ref_Number
                        RowCetak.Debet = row.Debet
                        RowCetak.Kredit = row.Kredit
                        If vShowItemRemarks = True And row.Remarks <> "" Then
                            RowCetak.Account = row.Account + vbCrLf + row.Remarks
                        Else
                            RowCetak.Account = row.Account
                        End If
                        RowCetak.Tanggal = row._Date
                        RowCetak.Desc = row.Desc
                        RowCetak.Debet_Original = row.Debet_Original
                        RowCetak.Kredit_Original = row.Credit_Original
                        If row.Ratio = 1 Then
                            RowCetak.Exchange_Rate = 0
                        Else
                            RowCetak.Exchange_Rate = row.Ratio
                        End If
                        RowCetak.Currency_Code = row.Currency_Code
                        RowCetak.Def_Prec_Cur = row.Default_Precision

                        DS.DataTable1.Rows.Add(RowCetak)
                    End If
                Next
            End If
            TBL.DefaultView.RowFilter = ""

            ''set grouping by Tanggal
            Dim Tanggal As Date = #1/1/1950#
            Dim JurnalNumber As String = ""
            For ij As Integer = 0 To DS.DataTable1.Rows.Count - 1
                RowCetak = DS.DataTable1.Rows(ij)
                If RowCetak.Tanggal <> Tanggal Then
                    Tanggal = RowCetak.Tanggal
                    RowCetak.Suppress_Tanggal = 0
                Else
                    RowCetak.Suppress_Tanggal = 1
                End If
                If RowCetak.No_Jurnal <> JurnalNumber Then
                    JurnalNumber = RowCetak.No_Jurnal
                    RowCetak.Suppress_Jurnal = 0
                Else
                    RowCetak.Suppress_Jurnal = 1
                End If
            Next

            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_Format_Day.ParameterFieldName, FormatDay)
            Report.SetParameterValue(Report.Parameter_Format_Month.ParameterFieldName, FormatMonth)
            Report.SetParameterValue(Report.Parameter_Format_Year.ParameterFieldName, FormatYears)
            Report.SetParameterValue(Report.Parameter_TanggalDari.ParameterFieldName, vStartDate.Date)
            Report.SetParameterValue(Report.Parameter_TanggalSampai.ParameterFieldName, vEndDate.Date)
            Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            Report.SetParameterValue(Report.Parameter_FormatNegatif.ParameterFieldName, FormatNegatif)

            CRV.ReportSource = Report
            CRV.Show()

            vDGVFilter.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            ObjCur = Nothing
            ObjSetting = Nothing
            ListPrimaryJurnal = Nothing
            vDGVFilter.AllowUserToAddRows = True
            ListCreatedBy = Nothing
        End Try
    End Function

    Public Function CetakJournal(ByVal vPrimary As Double, _
    ByVal Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal DS As DseTransaksiJournal) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tblCetakJurnalTableAdapter
        Dim TBL As DseTransaksiJournal.tblCetakJurnalDataTable = DS.tblCetakJurnal
        Dim FormatDecimal As Integer = 0
        Dim Negatif As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Report As New CrpCetakJurnal
        Dim Report2 As New CrpCetakJurnal2
        Dim Report3 As New CrpCetakJurnal3
        Dim CetakJurnalPortrait As String = ""
        Dim KodeCur As String = ""
        Dim row As DseTransaksiJournal.tblCetakJurnalRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimary)
            KodeCur = ObjCur.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)

            CetakJurnalPortrait = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CetakJurnalPortrait)

            If CetakJurnalPortrait = "" Or CetakJurnalPortrait = "0" Then
                If Me.ApakaJurnalNonIDR(vPrimary) = False Then
                    'cetak lanscape tanda currency no idr
                    Report2.SetDataSource(DS)
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

                    Crv.ReportSource = Report2
                    Crv.Zoom(1)
                    Crv.Show()
                Else
                    Report3.SetDataSource(DS)
                    For Each row In DS.tblCetakJurnal.Rows
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

                    Crv.ReportSource = Report3
                    Crv.Zoom(1)
                    Crv.Show()
                End If
            Else
                'cetak portrait
                Report.SetDataSource(DS)
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

    Public Function CetakJournal(ByVal vStartDate As Date, ByVal vEndDate As Date, _
     ByVal TBL As DseTransaksiJournal_Cetak.TblCetak_JurnalDataTable, ByVal CRV As _
     CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vclb As Windows.Forms.CheckedListBox, ByVal vOriginal As Boolean, _
     ByVal vShowItemRemarks As Boolean, _
     ByVal vJournalDesc As String, ByVal vRemarksItemDesc As String, ByVal vDGVFilter As DataGridView, ByVal vFrm As FrmTransaksiJournal_View) As Boolean
        Dim row As DseTransaksiJournal_Cetak.TblCetak_JurnalRow
        Dim RowCetak As DseTransaksiJournal_Cetak.DataTable1Row
        Dim Filter As String = ""
        Dim DataView As DataRowView
        Dim Report As New CrpCetakTransaksiJournal
        Dim DS As New DseTransaksiJournal_Cetak
        Dim DTA As New DseTransaksiJournal_CetakTableAdapters.TblCetak_JurnalTableAdapter
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DefaultPrecision As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim FormatNegatif As Integer = 0
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim DefCur As String = ""
        Dim ListPrimaryJurnal As New List(Of Double)
        Dim Lanjut As Boolean = True
        Dim TipeJurnal1 As EnumTipeJournal = EnumTipeJournal.GeneralJournal
        Dim TipeJurnal2 As EnumTipeJurnal2 = EnumTipeJurnal2.Cash_n_Bank

        Try
            vDGVFilter.AllowUserToAddRows = False
            FormatDecimal = ObjCur.GetDecimal_DefaultCurrency
            DefCur = ObjCur.GetCurrencyCode(1)

            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) <> "" Then
                FormatNegatif = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            Else
                FormatNegatif = 1
            End If

            DTA.FillBy_TanggalDariSampai(TBL, vEndDate, vStartDate)

            If vDGVFilter.RowCount > 0 Then
                Filter = ""
                For Each rowdgv As DataGridViewRow In vDGVFilter.Rows
                    If Filter = "" Then
                        Filter = "[Primary Akun] = " + rowdgv.Cells(vFrm.ColPrimaryAkun.Index).Value.ToString
                    Else
                        Filter = Filter + " or [Primary Akun] = " + rowdgv.Cells(vFrm.ColPrimaryAkun.Index).Value.ToString
                    End If
                Next
                TBL.DefaultView.RowFilter = Filter
                For Each DataView In TBL.DefaultView
                    row = DataView.Row
                    If ListPrimaryJurnal.Contains(row.Primary) = False Then
                        ListPrimaryJurnal.Add(row.Primary)
                    End If
                Next
            End If

            Dim CheckedItem As String = ""
            Filter = ""
            For i As Integer = 0 To vclb.CheckedItems.Count - 1
                CheckedItem = vclb.CheckedItems.Item(i)
                If CheckedItem = "General Journal" Then
                    'general journal
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.Jurnal_NoGrouping
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Cash And Bank" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.Cash_n_Bank
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "End of Month Journal" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Sales Journal" Then
                    TipeJurnal1 = EnumTipeJournal.SalesJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Sales Return Journal" Then
                    TipeJurnal1 = EnumTipeJournal.SalesReturnJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Sales Down Payment" Then
                    TipeJurnal1 = EnumTipeJournal.SalesDownPayment
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase Journal" Then
                    TipeJurnal1 = EnumTipeJournal.PurchaseJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase Return Journal" Then
                    TipeJurnal1 = EnumTipeJournal.PurchaseReturnJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase DownPayment Journal" Then
                    TipeJurnal1 = EnumTipeJournal.PurchaseDownPaymentJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Sales Receivement Journal" Then
                    TipeJurnal1 = EnumTipeJournal.SalesReceivementJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase Payment Journal" Then
                    TipeJurnal1 = EnumTipeJournal.PurchasePaymentJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Credit Note" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.Credit_Note
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Debit Note" Then
                    TipeJurnal1 = EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = EnumTipeJurnal2.Debit_Note
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                End If
            Next
            If vJournalDesc.EndsWith("%") = False Then
                vJournalDesc = vJournalDesc + "%"
            End If
            If vRemarksItemDesc.EndsWith("%") = False Then
                vRemarksItemDesc = vRemarksItemDesc + "%"
            End If
            vJournalDesc = vJournalDesc.Replace("%", "*")
            vRemarksItemDesc = vRemarksItemDesc.Replace("%", "*")
            If vJournalDesc <> "" Then
                Filter = "Desc like '" + vJournalDesc + "' AND " + "(" + Filter + ")"
                'filter = "Desc like '" + vJournalDesc + "'"
            End If
            If vRemarksItemDesc <> "" Then
                Filter = "Remarks like '" + vRemarksItemDesc + "' AND " + Filter
            End If

            TBL.DefaultView.RowFilter = Filter
            TBL.DefaultView.Sort = "Date, Tipe Journal, Primary"
            If TBL.DefaultView.Count > 0 Then

                For Each DataView In TBL.DefaultView
                    row = DataView.Row
                    Lanjut = False
                    If vDGVFilter.RowCount = 0 Then
                        Lanjut = True
                    Else
                        If ListPrimaryJurnal.Contains(row.Primary) = True Then
                            Lanjut = True
                        End If
                    End If

                    If Lanjut = True Then
                        RowCetak = DS.DataTable1.NewDataTable1Row
                        RowCetak.Primary = row.Primary
                        RowCetak.No_Jurnal = row.Ref_Number
                        If vOriginal = True Then
                            RowCetak.Debet = row.Debet_Original
                            RowCetak.Kredit = row.Credit_Original
                            RowCetak.Currency_Code = row.Currency_Code
                            RowCetak.Def_Prec_Cur = row.Default_Precision
                        Else
                            RowCetak.Debet = row.Debet
                            RowCetak.Kredit = row.Kredit
                            RowCetak.Currency_Code = DefCur
                            RowCetak.Def_Prec_Cur = FormatDecimal
                        End If
                        If vShowItemRemarks = True And row.Remarks <> "" Then
                            RowCetak.Account = row.Account + vbCrLf + row.Remarks
                        Else
                            RowCetak.Account = row.Account
                        End If
                        RowCetak.Tanggal = row._Date
                        RowCetak.Desc = row.Desc
                        DS.DataTable1.Rows.Add(RowCetak)
                    End If
                Next
            End If
            TBL.DefaultView.RowFilter = ""

            ''set grouping by Tanggal
            Dim Tanggal As Date = #1/1/1950#
            Dim JurnalNumber As String = ""
            For ij As Integer = 0 To DS.DataTable1.Rows.Count - 1
                RowCetak = DS.DataTable1.Rows(ij)
                If RowCetak.Tanggal <> Tanggal Then
                    Tanggal = RowCetak.Tanggal
                    RowCetak.Suppress_Tanggal = 0
                Else
                    RowCetak.Suppress_Tanggal = 1
                End If
                If RowCetak.No_Jurnal <> JurnalNumber Then
                    JurnalNumber = RowCetak.No_Jurnal
                    RowCetak.Suppress_Jurnal = 0
                Else
                    RowCetak.Suppress_Jurnal = 1
                End If
            Next

            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_Format_Day.ParameterFieldName, FormatDay)
            Report.SetParameterValue(Report.Parameter_Format_Month.ParameterFieldName, FormatMonth)
            Report.SetParameterValue(Report.Parameter_Format_Year.ParameterFieldName, FormatYears)
            Report.SetParameterValue(Report.Parameter_TanggalDari.ParameterFieldName, vStartDate.Date)
            Report.SetParameterValue(Report.Parameter_TanggalSampai.ParameterFieldName, vEndDate.Date)
            Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            Report.SetParameterValue(Report.Parameter_FormatNegatif.ParameterFieldName, FormatNegatif)

            CRV.ReportSource = Report
            CRV.Show()
            vDGVFilter.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            ObjCur = Nothing
            ObjSetting = Nothing
            ListPrimaryJurnal = Nothing
            vDGVFilter.AllowUserToAddRows = True
        End Try
    End Function

    Public Function GetTipeTransaksi(ByVal vTipeJournal As String) As Integer
        Try
            If vTipeJournal = "General Journal" Then
                Return EnumTipeJournal.GeneralJournal
            ElseIf vTipeJournal = "Adjustment Journal" Then
                Return EnumTipeJournal.AdjustmenJournal
            ElseIf vTipeJournal = "Closing Journal" Then
                Return EnumTipeJournal.ClosingJournal
            ElseIf vTipeJournal = "Sales Journal" Then
                Return EnumTipeJournal.SalesJournal
            ElseIf vTipeJournal = "Sales Return Journal" Then
                Return EnumTipeJournal.SalesReturnJournal
            ElseIf vTipeJournal = "Sales Down Payment" Then
                Return EnumTipeJournal.SalesDownPayment
            ElseIf vTipeJournal = "Purchase Journal" Then
                Return EnumTipeJournal.PurchaseJournal
            ElseIf vTipeJournal = "Purchase Return Journal" Then
                Return EnumTipeJournal.PurchaseReturnJournal
            ElseIf vTipeJournal = "Purchase DownPayment Journal" Then
                Return EnumTipeJournal.PurchaseDownPaymentJournal
            ElseIf vTipeJournal = "Sales Receivement Journal" Then
                Return EnumTipeJournal.SalesReceivementJournal
            ElseIf vTipeJournal = "Sales ReturnPayment Journal" Then
                Return EnumTipeJournal.SalesReturnPaymentJournal
            ElseIf vTipeJournal = "Purchase Payment Journal" Then
                Return EnumTipeJournal.PurchasePaymentJournal
            ElseIf vTipeJournal = "Purchase Return Receivement" Then
                Return EnumTipeJournal.PurchaseReturnReceivement
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Get_SalesBalance_Jurnal(ByVal vPrimarySalesBalance As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiJournal.DataTable2DataTable
        Dim row As DseTransaksiJournal.DataTable2Row
        Dim retval As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesInvoice_SalesInvoiceDP
        Dim TipeSalesBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual

        Try
            DTA.Fill(TBL, TipePayment, vPrimarySalesBalance, TipeSalesBalance)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Ref_Number
                Else
                    retval = retval + ", " + row.Ref_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_SalesBalance_SR_Jurnal(ByVal vPrimarySalesReturnBalance As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiJournal.DataTable2DataTable
        Dim row As DseTransaksiJournal.DataTable2Row
        Dim retval As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesInvoice_SalesInvoiceDP
        Dim TipeSalesBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR

        Try
            DTA.Fill(TBL, TipePayment, vPrimarySalesReturnBalance, TipeSalesBalance)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Ref_Number
                Else
                    retval = retval + ", " + row.Ref_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_SalesBalanceTransfer_Jurnal_Item(ByVal vPrimarySalesInvoice As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiJournal.DataTable2DataTable
        Dim row As DseTransaksiJournal.DataTable2Row
        Dim retval As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesInvoice_SalesInvoiceDP
        Dim TipeSalesBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice

        Try
            DTA.Fill(TBL, TipePayment, vPrimarySalesInvoice, TipeSalesBalance)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Ref_Number
                Else
                    retval = retval + ", " + row.Ref_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_SalesBalanceTransfer_Jurnal(ByVal vPrimaryTransaksi As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiJournal.DataTable2DataTable
        Dim row As DseTransaksiJournal.DataTable2Row
        Dim retval As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesInvoice_SalesInvoiceDP
        Dim TipeSalesBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice

        Try
            DTA.FillBy(TBL, TipePayment, TipeSalesBalance, vPrimaryTransaksi)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Ref_Number
                Else
                    retval = retval + ", " + row.Ref_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_SalesReturnBalance_Jurnal(ByVal vPrimarySalesReturnBalance As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiJournal.DataTable2DataTable
        Dim row As DseTransaksiJournal.DataTable2Row
        Dim retval As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesReturn
        Dim TipeSalesBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual

        Try
            DTA.Fill(TBL, TipePayment, vPrimarySalesReturnBalance, TipeSalesBalance)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Ref_Number
                Else
                    retval = retval + ", " + row.Ref_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_SalesReturnBalanceTransfer_Jurnal(ByVal vPrimaryTransaksi As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.DataTable3TableAdapter
        Dim TBL As New DseTransaksiJournal.DataTable3DataTable
        Dim row As DseTransaksiJournal.DataTable3Row
        Dim retval As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesReturn
        Dim TipeSalesBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice

        Try
            DTA.Fill(TBL, TipePayment, TipeSalesBalance, vPrimaryTransaksi)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Ref_Number
                Else
                    retval = retval + ", " + row.Ref_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_PurchaseBalance_Jurnal(ByVal vPrimaryPurchaseBalance As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiJournal.DataTable2DataTable
        Dim row As DseTransaksiJournal.DataTable2Row
        Dim retval As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
        Dim TipeSalesBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual

        Try
            DTA.Fill(TBL, TipePayment, vPrimaryPurchaseBalance, TipeSalesBalance)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Ref_Number
                Else
                    retval = retval + ", " + row.Ref_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_PurchaseBalanceTransfer_Jurnal(ByVal vPrimaryTransaksi As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.DataTable4TableAdapter
        Dim TBL As New DseTransaksiJournal.DataTable4DataTable
        Dim row As DseTransaksiJournal.DataTable4Row
        Dim retval As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
        Dim TipePurchaseBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice

        Try
            DTA.Fill(TBL, TipePayment, TipePurchaseBalance, vPrimaryTransaksi)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Ref_Number
                Else
                    retval = retval + ", " + row.Ref_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_PurchaseReturnBalance_Jurnal(ByVal vPrimaryPurchaseBalance As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiJournal.DataTable2DataTable
        Dim row As DseTransaksiJournal.DataTable2Row
        Dim retval As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
        Dim TipeSalesBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR

        Try
            DTA.Fill(TBL, TipePayment, vPrimaryPurchaseBalance, TipeSalesBalance)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Ref_Number
                Else
                    retval = retval + ", " + row.Ref_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_PurchaseReturnBalanceTransfer_Jurnal(ByVal vPrimaryTransaksi As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.DataTable5TableAdapter
        Dim TBL As New DseTransaksiJournal.DataTable5DataTable
        Dim row As DseTransaksiJournal.DataTable5Row
        Dim retval As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.Purchase_Return
        Dim TipePurchaseBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = _
        ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice

        Try
            DTA.Fill(TBL, TipePayment, TipePurchaseBalance, vPrimaryTransaksi)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Ref_Number
                Else
                    retval = retval + ", " + row.Ref_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_ByPrimaryAkun(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlJournalByPrimaryAkun(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlJournalByPrimaryAkun(vPrimaryAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    '' penambahan fungsi cari detail untuk tipe Sales Jurnal
    ''
    Public Overridable Function CariDetailSalesJournal(ByVal vPrimaryMaster As Double, _
        ByVal vDGVItem As DataGridView, ByVal frm As FrmTransaksiJournal_Sales, _
        ByVal tipe_journal As Integer, ByVal vDGVItemOther As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_NewTableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnal_NewDataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnal_NewRow
        Dim ObjItem As New ClsTransaksiJournal_Item

        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryMaster, tipe_journal) > 0 Then
                'salin data header
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = vPrimaryMaster
                    Me.mRefNumber = .Ref_Number
                    Me.mDate = ._Date
                    Me.mUrutan = .Urutan
                    Me.mTipeJournal = .Tipe_Journal
                    Me.mPrimaryDokumenAcuan = .Primary_Document_Acuan
                    Me.mPrimaryPihakAcuan = .Primary_Pihak_Acuan
                    Me.mDesc = .Desc
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mPrimaryAkunDifference = .primary_akun_difference
                    Me.mPrimaryAkunPayTo = .primary_akun_pay_to
                    Me.mPaid = .paid
                    Me.mPaid_IDR = .paid_IDR
                    Me.mARPaid_IDR = .AR_paid_IDR
                    Me.mDifference = .difference
                    Me.mExchangeRate = .exchange_rate
                    Me.mExchageRateGainLoss = .exchange_rate_gain_loss
                    Me.mTipeJurnal2 = .Tipe_Jurnal_2
                    Me.mReferenceParty = .Reference_Party
                    Me.mPaymentMethod = .payment_method
                    Me.mPrimaryExChangerateGainLoss = .primary_exhange_gain_loss
                End With

                'salin data item ada 2 item 
                'CodeNo
                ObjItem.IsiDGVItemSalesJournal(vDGVItem, vPrimaryMaster, frm, vDGVItemOther)

                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDGVItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    '' penambahan fungsi cari detail untuk tipe Purchase Jurnal
    ''
    Public Overridable Function CariDetailPurchaseJournal(ByVal vPrimaryMaster As Double, _
          ByVal vDGVItem As DataGridView, ByVal frm As FrmTransaksiJournal_Purchase, _
          ByVal tipe_journal As Integer, ByVal vDGVAkunOther As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_NewTableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnal_NewDataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnal_NewRow
        Dim ObjItem As New ClsTransaksiJournal_Item
        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryMaster, tipe_journal) > 0 Then
                'salin data header
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = vPrimaryMaster
                    Me.mRefNumber = .Ref_Number
                    Me.mDate = ._Date
                    Me.mUrutan = .Urutan
                    Me.mTipeJournal = .Tipe_Journal
                    Me.mPrimaryDokumenAcuan = .Primary_Document_Acuan
                    Me.mPrimaryPihakAcuan = .Primary_Pihak_Acuan
                    Me.mDesc = .Desc
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mPrimaryAkunDifference = .primary_akun_difference
                    Me.mPrimaryAkunPayTo = .primary_akun_pay_to
                    Me.mPaid = .paid
                    Me.mPaid_IDR = .paid_IDR
                    Me.mARPaid_IDR = .AR_paid_IDR
                    Me.mDifference = .difference
                    Me.mExchangeRate = .exchange_rate
                    Me.mExchageRateGainLoss = .exchange_rate_gain_loss
                    Me.mTipeJurnal2 = .Tipe_Jurnal_2
                    Me.mReferenceParty = .Reference_Party
                    If String.IsNullOrEmpty(.payment_method) OrElse IsDBNull(.payment_method) Then
                        Me.mPaymentMethod = ""
                    Else
                        Me.mPaymentMethod = .payment_method
                    End If
                    Me.mPrimaryExChangerateGainLoss = .primary_exhange_gain_loss
                End With
                'salin data item ada 2 item 
                'CodeNo
                ObjItem.IsiDGVItemPurchaseJournal(vDGVItem, vPrimaryMaster, frm, vDGVAkunOther)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDGVItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    ''Untuk Sales Journal
    ''
    Public Overridable Function SimpanBaruSalesJournal(ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal_Sales, ByVal vTipeJournal As ClsTransaksiJournal.EnumTipeJournal, _
    ByVal vDGVSalesBalance As DataGridView, ByVal vDGVAkunOther As DataGridView) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()
            'ambil no antrian
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_Sales_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    DTAInOUT.Delete(Me.vPrimary_Antrian)
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaruSalesJournal1(vDGVItem, vUserCode, Frm, vTipeJournal, vDGVSalesBalance, vDGVAkunOther)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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

    'pastikan semua item payment, original balancenya- sama dengan original balance yang ada didalam database,
    'untuk memastikan data sales balance selalu konsisten
    'jika tidak konsisten mengembalikan nilai false
    Private Function CekKonsistensiSalesBalance(ByVal vDGVSalesBalance As DataGridView, ByVal Frm As FrmTransaksiJournal_Sales) As Boolean
        Dim ObjBalanceSales As New ClsTransaksiSalesBalance
        Dim DTA_BS As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In vDGVSalesBalance.Rows
                If CDec(rowDGV.Cells(Frm.ColBalanceOriginal.Index).Value) > 0 Then
                    Dim BalanceOriginal As Decimal = 0
                    BalanceOriginal = ObjBalanceSales.GetBalanceOriginal(rowDGV.Cells(Frm.ColPrimarySalesBalance.Index).Value, DTA_BS)
                    If CDec(rowDGV.Cells(Frm.ColBalanceOriginal02.Index).Value) <> BalanceOriginal Then
                        MsgBox("Other user have enter sales balance data as you did, process stop.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            ObjBalanceSales = Nothing
            DTA_BS = Nothing
        End Try
    End Function

    'pastikan semua item payment, original balancenya- sama dengan original balance yang ada didalam database,
    'untuk memastikan data purchase balance selalu konsisten
    'jika tidak konsisten mengembalikan nilai false
    Private Function CekKonsistensiPurchaseBalance(ByVal vDGVPurchaseBalance As DataGridView, ByVal Frm As FrmTransaksiJournal_Purchase) As Boolean
        Dim ObjBalancePurchase As New ClsTransaksiPurchaseBalance
        Dim DTA_BS As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In vDGVPurchaseBalance.Rows
                If CDec(rowDGV.Cells(Frm.ColBalanceOriginal.Index).Value) > 0 Then
                    If rowDGV.Cells(Frm.ColBalanceOriginal02.Index).Value <> _
                    ObjBalancePurchase.GetBalanceOriginal(rowDGV.Cells(Frm.ColPrimaryPurchaseBalance.Index).Value, DTA_BS) Then
                        MsgBox("Other user have enter purchase balance data as you did, process stop.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            ObjBalancePurchase = Nothing
            DTA_BS = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaruSalesJournal1(ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal_Sales, ByVal vTipeJournal As ClsTransaksiJournal.EnumTipeJournal, _
    ByVal vDGVSalesBalance As DataGridView, ByVal vDGVAkunOther As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item
        Dim ObjCustom As New CustomSMESS.ClsTransaksiJournal

        Try
            Me.mTipeJournal = vTipeJournal
            Me.mUrutan = -1
            Me.mPrimaryDokumenAcuan = -1

            If Me.ApakahDataBenar1(True, vDGVItem, Frm.ColPrimaryCurrency.Index, Me.Tanggal, vDGVSalesBalance, Frm.ColPrimaryCurrencyBalance.Index, _
            Frm.ColReceivementOriginal.Index, Frm.ColTipeDocumentAcuan_Sales.Index, Frm.ColPrimaryDocumenAcuan_Sales.Index, Frm.ColInvoiceNo.Index, _
            Me.mPrimaryPihakAcuan, Frm.ColExhangeRateBalance.Index, Frm.ColInvoiceDate.Index) = True Then
                'pastikan semua item payment, original balancenya- sama dengan original balance yang ada didalam database,
                'untuk memastikan data sales balance selalu konsisten
                If Me.CekKonsistensiSalesBalance(vDGVSalesBalance, Frm) = False Then
                    Return False
                End If

                If DTA.Insert(Me.RefNumber, Me.Tanggal, Me.mUrutan, Me.mTipeJournal, Me.mPrimaryDokumenAcuan, _
                Me.mPrimaryPihakAcuan, Me.Description, vUserCode, Now, _
                 vUserCode, Now, Now, Me.mPrimaryAkunPayTo, Me.mPrimaryAkunDifference, Me.mPaid, Me.mExchangeRate, Me.mPaid_IDR, _
                 Me.mARPaid_IDR, Me.mDifference, Me.mExchageRateGainLoss, Me.mTipeJurnal2, Me.mTipeInOut, Me.mPrimaryCurr, Me.mReferenceParty, Me.mPaymentMethod) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mRefNumber).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItemSalesJournal(vDGVItem, Me.mPrimary, Frm, vDGVSalesBalance, _
                    vTipeJournal, Me.mRefNumber, Me.mPrimaryPihakAcuan, Me.mDate, vDGVAkunOther)

                    ObjCustom.FollowUpSimpan_SalesJournal(Me.mPrimary)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mRefNumber + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaruSalesJournal1(vDGVItem, vUserCode, Frm, vTipeJournal, vDGVSalesBalance, vDGVAkunOther)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjCustom = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vDGVItem As DataGridView, ByVal vDGVItem_Payment As DataGridView, _
    ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal_CashBank, ByVal vDgvCek As DataGridView, _
     ByVal vFrmCek As FrmTransaksiJournal_CashBank_Cheque) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()
            'ambil no antrian
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_Cash_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    DTAInOUT.Delete(Me.vPrimary_Antrian)
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vDGVItem, vDGVItem_Payment, vUserCode, Frm, vDgvCek, vFrmCek)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Overridable Function SimpanBaru(ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal_AssetDepreciation, _
    ByVal dgvItemAsset As DataGridView) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()
            'ambil no antrian
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_Asset_Dep_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    DTAInOUT.Delete(Me.vPrimary_Antrian)
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vDGVItem, vUserCode, Frm, dgvItemAsset)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Overridable Function SimpanBaru1(ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal_AssetDepreciation, _
    ByVal vDGVItemAsset As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item
        Dim ObjCek As New ClsTransaksiJournal_Item_Cheque
        Try
            Me.mUrutan = -1
            Me.mPrimaryDokumenAcuan = -1
            Me.mPrimaryPihakAcuan = -1
            If Me.ApakahDataBenar1(True, vDGVItem, Frm.ColAsset_PrimaryCurrency.Index, Me.Tanggal) = True Then
                If DTA.Insert(Me.RefNumber, Me.Tanggal, Me.Urutan, Me.TipeJournal, Me.PrimaryDokumenAcuan, _
                Me.PrimaryPihakAcuan, Me.Description, vUserCode, Now, _
                 vUserCode, Now, Now, Me.PrimaryAkunPayTo, Me.PrimaryAkunDifference, Me.Paid, Me.ExchangeRate, Me.Paid_IDR, _
                 Me.ARPaid_IDR, Me.Difference, Me.ExchageRateGainLoss, Me.TipeJurnal2, Me.TipeInOut, Me.PrimaryCurrency, Me.mReferenceParty, Me.mPaymentMethod) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mRefNumber).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItem(vDGVItem, Me.mPrimary, Frm, vDGVItemAsset)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mRefNumber + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru1(vDGVItem, vUserCode, Frm, vDGVItemAsset)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjCek = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru1(ByVal vDGVItem As DataGridView, ByVal vDGVPayment As DataGridView, _
    ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal_CashBank, ByVal vDgvCek As DataGridView, _
    ByVal vFrmCek As FrmTransaksiJournal_CashBank_Cheque) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item
        Dim ObjCek As New ClsTransaksiJournal_Item_Cheque
        Try
            Me.mUrutan = -1
            Me.mPrimaryDokumenAcuan = -1
            Me.mPrimaryPihakAcuan = -1
            If Me.ApakahDataBenar1(True, vDGVItem, Frm.ColPrimaryCurrency.Index, Me.Tanggal) = True Then
                If DTA.Insert(Me.RefNumber, Me.Tanggal, Me.Urutan, Me.TipeJournal, Me.PrimaryDokumenAcuan, _
                Me.PrimaryPihakAcuan, Me.Description, vUserCode, Now, _
                 vUserCode, Now, Now, Me.PrimaryAkunPayTo, Me.PrimaryAkunDifference, Me.Paid, Me.ExchangeRate, Me.Paid_IDR, _
                 Me.ARPaid_IDR, Me.Difference, Me.ExchageRateGainLoss, Me.TipeJurnal2, Me.TipeInOut, Me.PrimaryCurrency, Me.mReferenceParty, Me.mPaymentMethod) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mRefNumber).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItem(vDGVItem, Me.mPrimary, Frm, vDGVPayment)
                    ObjCek.SimpanCheque(Me.mPrimary, vDgvCek, vFrmCek)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mRefNumber + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru1(vDGVItem, vDGVPayment, vUserCode, Frm, vDgvCek, vFrmCek)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjCek = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vDGVItem As DataGridView, ByVal vDGVItem_pay As DataGridView, _
   ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal_CashBank, ByVal frmcek As FrmTransaksiJournal_CashBank_Cheque, _
   ByVal DgvCek As DataGridView) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()

            'ambil no antrian
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_Cash_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vDGVItem, vDGVItem_pay, vUserCode, frm, frmcek, DgvCek)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vDGVItem As DataGridView, _
   ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal_AssetDepreciation, ByVal vDGVItemAsset As DataGridView, ByVal vTipeJurnal As ClsTransaksiJournal.EnumTipeJournal) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()

            'ambil no antrian
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_Asset_Dep_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vDGVItem, vUserCode, frm, vTipeJurnal, vDGVItemAsset)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1(ByVal vDGVItem As DataGridView, ByVal vDGVItem_Pay As DataGridView, _
    ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal_CashBank, ByVal vFrmCek As FrmTransaksiJournal_CashBank_Cheque, _
    ByVal vDgvCek As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item
        Dim ObjCek As New ClsTransaksiJournal_Item_Cheque

        Try
            Me.mUrutan = -1
            Me.mPrimaryDokumenAcuan = -1
            Me.mPrimaryPihakAcuan = -1
            If Me.ApakahDataBenar1(False, vDGVItem, frm.ColPrimaryCurrency.Index, Me.Tanggal) = True Then
                If DTA.Update(Me.RefNumber, Me.Tanggal, Me.Urutan, Me.TipeJournal, Me.PrimaryDokumenAcuan, _
                Me.PrimaryPihakAcuan, Me.Description, vUserCode, Now, Now, _
                Me.PrimaryAkunPayTo, Me.PrimaryAkunDifference, Me.Paid, Me.ExchangeRate, Me.Paid_IDR, Me.ARPaid_IDR, Me.Difference, _
                Me.ExchageRateGainLoss, Me.TipeJurnal2, Me.TipeInOut, Me.PrimaryCurrency, Me.mReferenceParty, Me.mPaymentMethod, Me.Primary, Me.Timestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)
                    ObjCek.HapusItem(Me.mPrimary)

                    'simpan data item baru
                    ObjItem.SimpanItem(vDGVItem, Me.mPrimary, frm, vDGVItem_Pay)
                    ObjCek.SimpanCheque(Me.mPrimary, vDgvCek, vFrmCek)

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
                Return Me.SimpanPerubahan1(vDGVItem, vDGVItem_Pay, vUserCode, frm, vFrmCek, vDgvCek)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    ''Untuk Purchase Journal
    ''
    Public Overridable Function SimpanBaruPurchaseJournal(ByVal vDGVItem As DataGridView, _
          ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal_Purchase, ByVal vTipeJournal As ClsTransaksiJournal.EnumTipeJournal, _
    ByVal vDGVPurchaseBalance As DataGridView, ByVal vDGVAkunOther As DataGridView) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()
            'ambil no antrian
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_Purchase_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    DTAInOUT.Delete(Me.vPrimary_Antrian)
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaruPurchaseJournal1(vDGVItem, vUserCode, Frm, vTipeJournal, vDGVPurchaseBalance _
                            , vDGVAkunOther)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Overridable Function SimpanBaruPurchaseJournal1(ByVal vDGVItem As DataGridView, _
          ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal_Purchase, ByVal vTipeJournal As ClsTransaksiJournal.EnumTipeJournal, _
    ByVal vDGVPurchaseBalance As DataGridView, ByVal vDGVAkunOther As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim DTA2 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_NewTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item
        Try
            Me.mTipeJournal = vTipeJournal
            Me.mUrutan = -1
            Me.mPrimaryDokumenAcuan = -1
            If Me.ApakahDataBenar1(True, vDGVItem, Frm.ColPrimaryCurrency.Index, Me.Tanggal, vDGVPurchaseBalance, Frm.ColPrimaryCurrencyBalance.Index, _
            Frm.ColPaymentOriginal.Index, Frm.ColTipeDocumentAcuan_Purchase.Index, Frm.ColPrimaryDocumentAcuan_Purchase.Index, Frm.ColAOPNumber.Index, _
            Me.PrimaryPihakAcuan, Frm.ColExchangeRateBalance.Index, Frm.ColAOPDate.Index) = True Then
                'cek konsistensi purchase balance
                If Me.CekKonsistensiPurchaseBalance(vDGVPurchaseBalance, Frm) = False Then
                    Return False
                End If

                If DTA2.InsertQuery(Me.RefNumber, Me.Tanggal, Me.mUrutan, Me.mTipeJournal, Me.mPrimaryDokumenAcuan, _
                Me.mPrimaryPihakAcuan, Me.Description, vUserCode, Now, _
                 vUserCode, Now, Now, Me.mPrimaryAkunPayTo, Me.mPrimaryAkunDifference, Me.mPaid, Me.mExchangeRate, Me.mPaid_IDR, _
                 Me.mARPaid_IDR, Me.mDifference, Me.mExchageRateGainLoss, Me.mTipeJurnal2, Me.mTipeInOut, Me.mPrimaryCurr, _
                 Me.mReferenceParty, Me.mPaymentMethod, Me.mPrimaryExChangerateGainLoss) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mRefNumber).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItemPurchaseJournal(vDGVItem, Me.mPrimary, Frm, vDGVPurchaseBalance, vTipeJournal, _
                     Me.mRefNumber, Me.mPrimaryPihakAcuan, Me.mDate, vDGVAkunOther)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mRefNumber + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaruPurchaseJournal1(vDGVItem, vUserCode, Frm, vTipeJournal, vDGVPurchaseBalance, vDGVAkunOther)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function GetCurrentYearEarning1(ByVal vEndDate As Date, ByVal vPrimaryDept As Double, ByVal vPrimaryJob As Double) As Decimal
        Dim DTA As New DseTransaksiJournalTableAdapters.tblmaster_akunTableAdapter
        Dim TBL As New DseTransaksiJournal.tblmaster_akunDataTable
        Dim row As DseTransaksiJournal.tblmaster_akunRow
        Dim ObjSetting As New ClsTransaksiClosing
        Dim ConvertionDate As Date
        Dim TotalDebet As Decimal = 0
        Dim TotalKredit As Decimal = 0
        Dim StartPeriod, EndPeriod As Date
        Dim PrimaryAkunRetainedEarning As String = -1
        Dim ObjCD As New Core_Module.ClsCekData
        Dim StartPeriod2 As Date = #1/1/1950#
        Dim JmlClosingJurnal As Double = 0
        Dim Loop1 As Boolean = True

        Try
            ConvertionDate = ObjSetting.GetConvertionDate

            StartPeriod = ObjCD.GetFisrtDateOfTheYear(vEndDate)
            EndPeriod = ObjCD.GetLasDateOfTheMonth(vEndDate)

            'cek apakah dperiode tahun sebelumnya sudah ada closing jurnal, jika belum, maka start period harus diundur
            Loop1 = True
            While Loop1 = True
                StartPeriod2 = ObjCD.GetFisrtDateOfTheYear(StartPeriod - TimeSpan.FromDays(10))
                JmlClosingJurnal = Me.GetJml(EnumTipeJournal.ClosingJournal, StartPeriod2, StartPeriod)
                If JmlClosingJurnal = 0 Then
                    If StartPeriod2 < ConvertionDate Then
                        Loop1 = False
                    ElseIf StartPeriod2 = ConvertionDate Then
                        StartPeriod = StartPeriod2
                        Loop1 = False
                    Else
                        Loop1 = True
                    End If
                Else
                    Loop1 = False
                End If

                If Loop1 = True Then
                    StartPeriod = StartPeriod2
                End If
            End While

            'masukkan jurnal closing
            ConvertionDate = ObjCD.GetFisrtDateOfTheYear(ConvertionDate)

            If vPrimaryJob = -1 And vPrimaryDept = -1 Then
                DTA.Fill(TBL, EndPeriod, StartPeriod)
            ElseIf vPrimaryDept <> -1 And vPrimaryJob = -1 Then
                DTA.FillBy_PrimaryDept(TBL, EndPeriod, StartPeriod, vPrimaryDept)
            ElseIf vPrimaryDept = -1 And vPrimaryJob <> -1 Then
                DTA.FillBy_PrimaryJob(TBL, EndPeriod, StartPeriod, vPrimaryJob)
            Else
                DTA.FillBy_PrimaryDept_PrimaryJob(TBL, EndPeriod, StartPeriod, vPrimaryDept, vPrimaryJob)
            End If

            'jika yang diclosing adalah period pertama, maka harus masukkan data beginning balance
            If StartPeriod = ObjCD.GetFirstDateofTheMonth(ConvertionDate) Then
                For Each row In TBL.Rows
                    If row.Saldo_Normal = 0 Then
                        row.Debet += row.Beginning_Balance
                    Else
                        row.Kredit += row.Beginning_Balance
                    End If
                Next
            End If

            TotalDebet = 0
            TotalKredit = 0
            For Each row In TBL.Rows
                With row
                    If row.Debet <> 0 Or row.Kredit <> 0 Then
                        TotalDebet += row.Debet
                        TotalKredit += row.Kredit
                    End If
                End With
            Next

            Return TotalKredit - TotalDebet
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function Get_AOP_Jurnal(ByVal vPrimaryAOP As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.DataTable6TableAdapter
        Dim TBL As New DseTransaksiJournal.DataTable6DataTable
        Dim row As DseTransaksiJournal.DataTable6Row
        Dim retval As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
        Dim TipeSalesBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
        Dim listJurnal As New List(Of String)

        Try
            DTA.Fill(TBL, TipePayment, vPrimaryAOP, TipeSalesBalance)
            For Each row In TBL.Rows
                If listJurnal.Contains(row.Ref_Number) = False Then
                    listJurnal.Add(row.Ref_Number)
                    If retval = "" Then
                        retval = row.Ref_Number
                    Else
                        retval = retval + ", " + row.Ref_Number
                    End If
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            listJurnal = Nothing
        End Try
    End Function
    '-----------------------------------------------------------------------------------------------
    'Erna Code 
    Public Function ApakahAdaPaymentNonIDR(ByVal vPrimaryJurnal As Double) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_paymentTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlNonIDR(Core_Module.PrimaryBaseCurrency, vPrimaryJurnal).HasValue = True Then
                If DTA.ScalarQuery_GetJmlNonIDR(Core_Module.PrimaryBaseCurrency, vPrimaryJurnal).Value > 0 Then
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
    '-----------------------------------------------------------------------------------------------

    Public Function Get_PurchaseReturn_Jurnal(ByVal vPrimaryPurchaseReturn As Double) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.DataTable6TableAdapter
        Dim TBL As New DseTransaksiJournal.DataTable6DataTable
        Dim row As DseTransaksiJournal.DataTable6Row
        Dim retval As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
        Dim TipeSalesBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
        Dim ListNoJurnal As New List(Of String)

        Try
            DTA.Fill(TBL, TipePayment, vPrimaryPurchaseReturn, TipeSalesBalance)
            For Each row In TBL.Rows
                If ListNoJurnal.Contains(row.Ref_Number) = False Then
                    If retval = "" Then
                        retval = row.Ref_Number
                    Else
                        retval = retval + ", " + row.Ref_Number
                    End If
                    ListNoJurnal.Add(row.Ref_Number)
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ListNoJurnal = Nothing
        End Try
    End Function

    Public Sub GetLedgerAccess(ByVal vKodeUser As String, ByVal vListKodeUser As List(Of String))
        Dim DTA As New DseTransaksiJournalTableAdapters.LedgerAccess_DetailUserTableAdapter
        Dim TBL As New DseTransaksiJournal.LedgerAccess_DetailUserDataTable
        Dim row As DseTransaksiJournal.LedgerAccess_DetailUserRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_kode_user_ref(TBL, vKodeUser)
            vListKodeUser.Clear()
            vListKodeUser.Add(vKodeUser)
            For Each row In TBL.Rows
                vListKodeUser.Add(row.Kode_User)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetLedgerAccess(ByVal vKodeUser As String) As EnumJournalAccesType
        Dim DTA As New DseTransaksiJournalTableAdapters.LedgerAccessTableAdapter
        Dim RetVal As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetTipeAccess(vKodeUser)
            If RetVal.HasValue = False Then
                Return EnumJournalAccesType.Respected_User_Only
            Else
                Return RetVal.Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahMemilikiLedgerAccess(ByVal vKodeUserViewer As String, ByVal vKodeUserViewed As String) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.LedgerAccessTableAdapter
        Dim TipeAccess As Nullable(Of Integer)
        Dim TipeAccess1 As EnumJournalAccesType = EnumJournalAccesType.FullAccess
        Dim DTA1 As New DseTransaksiJournalTableAdapters.LedgerAccess_DetailUserTableAdapter
        Dim TBL1 As New DseTransaksiJournal.LedgerAccess_DetailUserDataTable

        Try
            If vKodeUserViewed = vKodeUserViewer Then
                Return True
            Else
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTA1.Connection.ConnectionString = Core_Module.GetConstring

                TipeAccess = DTA.ScalarQuery_GetTipeAccess(vKodeUserViewer)
                If TipeAccess.HasValue = False Then
                    'hanya bisa lihat jurnal2 yang dibuat oleh dirinya sendiri
                    Return False
                Else
                    TipeAccess1 = TipeAccess.Value
                    If TipeAccess1 = EnumJournalAccesType.FullAccess Then
                        'full access, bisa melihat semua jurnal
                        Return True
                    Else
                        'selected user only, cek apakah user ybs masuk ke dalam daftar yg boleh di view
                        DTA1.FillBy_kode_user(TBL1, vKodeUserViewer, vKodeUserViewed)
                        If TBL1.Rows.Count = 0 Then
                            Return False
                        Else
                            Return True
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Function

    Public Function GetDescription(ByVal vPrimaryJurnal As Double) As String
        Dim DTAJurnal As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim RetVal As String = ""

        Try
            DTAJurnal.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTAJurnal.ScalarQuery_GetDesc(vPrimaryJurnal)

            If String.IsNullOrEmpty(RetVal) = True Then
                Return ""
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTAJurnal.Dispose()
        End Try
    End Function
End Class
