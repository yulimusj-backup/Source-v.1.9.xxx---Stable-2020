Option Explicit On
Option Strict Off
Imports Core_Module

Public Class ClsTrans_Session

    Private mPrimary As Double = -1
    Private mSessionID As String = ""
    Private mPrimaryTransServer As Double = -1
    Private mStatus As EnumStatusSession = EnumStatusSession.Open
    Private mStartDate As DateTime = "1950/01/01 00:00:00"
    Private mEndDate As DateTime = "1950/01/01 00:00:00"
    Private mCashierCode As String = ""
    Private mPrimaryBranch As Double = -1
    Private mRemarks As String = ""
    Private mBeginningCash As Decimal = 0
    Private mCashSales As Decimal = 0
    Private mDropCash As Decimal = 0
    Private mPayout As Decimal = 0
    Private mBalance As Decimal = 0
    Private mPrimaryBeginningCA As Double = -1
    Private mPrimaryCashierCA As Double = -1
    Private mPrimaryBalanceCA As Double = -1
    Private mPrimaryJurnalOpenSession As Double = -1
    Private mPrimaryJurnalCloseSession As Double = -1
    Private mStatusTransfer As EnumStatusTransfer = EnumStatusTransfer.BelumTransfer
    Private mStartTransferTime As DateTime = "1950/01/01 00:00:00"
    Private mEndTransferTime As DateTime = "1950/01/01 00:00:00"
    Private mUrut As Double = 0

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiPOS
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New DseTransaksiAntrianTableAdapters.tbltrans_antrianTableAdapter

    Public Property PrimaryTransServer() As Double
        Get
            Return mPrimaryTransServer
        End Get
        Set(ByVal value As Double)
            mPrimaryTransServer = value
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

    Public Property Urut() As Double
        Get
            Return mUrut
        End Get
        Set(ByVal value As Double)
            mUrut = value
        End Set
    End Property

    Public Enum EnumStatusTransfer As Integer
        BelumTransfer
        SudahTransfer
    End Enum

    Public Property StatusTransfer() As EnumStatusTransfer
        Get
            Return mStatusTransfer
        End Get
        Set(ByVal value As EnumStatusTransfer)
            mStatusTransfer = value
        End Set
    End Property

    Public Property BeginningCash() As Decimal
        Get
            Return mBeginningCash
        End Get
        Set(ByVal value As Decimal)
            mBeginningCash = value
        End Set
    End Property

    Public Property CashSales() As Decimal
        Get
            Return mCashSales
        End Get
        Set(ByVal value As Decimal)
            mCashSales = value
        End Set
    End Property

    Public Property DropCash() As Decimal
        Get
            Return mDropCash
        End Get
        Set(ByVal value As Decimal)
            mDropCash = value
        End Set
    End Property

    Public Property Payout() As Decimal
        Get
            Return mPayout
        End Get
        Set(ByVal value As Decimal)
            mPayout = value
        End Set
    End Property

    Public Property Balance() As Decimal
        Get
            Return mBalance
        End Get
        Set(ByVal value As Decimal)
            mBalance = value
        End Set
    End Property

    Public Property PrimaryBeginningCA() As Double
        Get
            Return mPrimaryBeginningCA
        End Get
        Set(ByVal value As Double)
            mPrimaryBeginningCA = value
        End Set
    End Property

    Public Property PrimaryCashierCA() As Double
        Get
            Return mPrimaryCashierCA
        End Get
        Set(ByVal value As Double)
            mPrimaryCashierCA = value
        End Set
    End Property

    Public Property PrimaryBalanceCA() As Double
        Get
            Return mPrimaryBalanceCA
        End Get
        Set(ByVal value As Double)
            mPrimaryBalanceCA = value
        End Set
    End Property

    Public Property PrimaryJurnalOpenSession() As Double
        Get
            Return mPrimaryJurnalOpenSession
        End Get
        Set(ByVal value As Double)
            mPrimaryJurnalOpenSession = value
        End Set
    End Property

    Public Property PrimaryJurnalCloseSession() As Double
        Get
            Return mPrimaryJurnalCloseSession
        End Get
        Set(ByVal value As Double)
            mPrimaryJurnalCloseSession = value
        End Set
    End Property

    Public Property StartTransferTime() As DateTime
        Get
            Return mStartTransferTime
        End Get
        Set(ByVal value As DateTime)
            mStartTransferTime = value
        End Set
    End Property

    Public Property EndTransferTime() As DateTime
        Get
            Return mEndTransferTime
        End Get
        Set(ByVal value As DateTime)
            mEndTransferTime = value
        End Set
    End Property

    Public Enum EnumStatusSession As Integer
        Open
        Closed
    End Enum

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property SessionID() As String
        Get
            Return mSessionID
        End Get
        Set(ByVal value As String)
            mSessionID = value
        End Set
    End Property

    Public Property Status() As EnumStatusSession
        Get
            Return mStatus
        End Get
        Set(ByVal value As EnumStatusSession)
            mStatus = value
        End Set
    End Property

    Public Property StartDate() As DateTime
        Get
            Return mStartDate
        End Get
        Set(ByVal value As DateTime)
            mStartDate = value
        End Set
    End Property

    Public Property EndDate() As DateTime
        Get
            Return mEndDate
        End Get
        Set(ByVal value As DateTime)
            mEndDate = value
        End Set
    End Property

    Public Property CashierCode() As String
        Get
            Return mCashierCode
        End Get
        Set(ByVal value As String)
            mCashierCode = value
        End Set
    End Property

    Public Property PrimaryBranch() As Double
        Get
            Return mPrimaryBranch
        End Get
        Set(ByVal value As Double)
            mPrimaryBranch = value
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

    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307"
        End Try
    End Function

    Public Function GetPrimaryBranch(ByVal vPrimarySession As Double) As Double
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimaryBranch(vPrimarySession).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryBranch(vPrimarySession).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlBySessionID(ByVal vSesionID As String) As Double
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlBySessionID(vSesionID).HasValue Then
                Return DTA.ScalarQuery_GetJmlBySessionID(vSesionID).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vSessionID As String) As Double
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimaryBranch(vSessionID).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryBranch(vSessionID).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSessionID(ByVal vPrimary As Double) As String
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSessionID(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetSessionID(vPrimary)
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
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
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

    Public Function ApakahSudahAda(ByVal vRefName As String) As Boolean
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlBySessionID(vRefName).Value > 0 Then
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

    Private Sub GetNoAntrian()
        Try
            DTAAntrian.Connection.ConnectionString = Me.GetConstring
            Me.vIdentifier = Me.ObjRnd.Next(1000000)
            'jangan sampai sama dengan identifier yang sudah ada
            While Me.DTAAntrian.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 0
                Me.vIdentifier = Me.ObjRnd.Next(1000000)
            End While
            'masukan data baru ke antrian
            Me.DTAAntrian.Insert(Me.vTipeAntrian, Me.vIdentifier)

            'cek jangan sampai ada 2 data yang sama
            While Me.DTAAntrian.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 1
                'bila ada yang sama maka harus masukkan data baru
                DTAAntrian.DeleteQuery(Me.vTipeAntrian, Me.vIdentifier)
                'minta identifier baru
                Me.vIdentifier = Me.ObjRnd.Next(1000000)
                'jangan sampai sama dengan identifier yang sudah ada
                While Me.DTAAntrian.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 0
                    Me.vIdentifier = Me.ObjRnd.Next(1000000)
                    'masukan data baru ke antrian
                    Me.DTAAntrian.Insert(Me.vTipeAntrian, Me.vIdentifier)
                End While
            End While

            'ambil no antrian
            Me.vPrimary_Antrian = DTAAntrian.ScalarQuery_GetPrimary(Me.vIdentifier, Me.vTipeAntrian).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Ketentuan Get No Baru
    'Format : xxxxx/ses/kodecabang/mm/yyyy
    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vPrimaryBranch As Double) As String
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Dim ObjSession As New ClsTrans_Session
        Dim ObjBranch As New ClsMast_Cabang
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData
        Dim BranchStr As String = "-"
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetMaxUrutan(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetMaxUrutan(TanggalEnd, TanggalStart).Value + 1)
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

            'Branch 
            BranchStr = ObjBranch.GetCode(vPrimaryBranch)

            Return NoUrut + "/" + "SES" + "/" + BranchStr + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            ObjSession = Nothing
            ObjBranch = Nothing
        End Try
    End Function

    ''' <summary>
    ''' data ini digunakan untuk mengecek kebenaran data yang diberikan
    ''' </summary>
    ''' <param name="vBaru"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vPrimary As Double, _
    ByVal vShowMsg As Boolean, ByVal vObjCD As ClsCekData, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmTrans_Session) As Boolean
        'Dim ObjClosing As New ClsTransaksiClosing
        Dim ListPrimaryAkun As New List(Of Double)
        Try
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mSessionID) = True Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Duplication in session id, please change your code", _
                            MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Duplikasi session id, silahkan ubah code", _
                            MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            Else
                Dim NoTransaksiLama As String = Me.GetSessionID(Me.mPrimary)
                If NoTransaksiLama <> Me.mSessionID Then
                    If Me.ApakahSudahAda(Me.mSessionID) = True Then
                        If vShowMsg = True Then
                            If vEnumBhs = EnumBahasa.Inggris Then
                                MsgBox("Duplication in code, please change your code", _
                                MsgBoxStyle.Information, "Information")
                            ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                                MsgBox("Duplikasi code, silahkan ubah code", _
                                MsgBoxStyle.Information, "Informasi")
                            End If
                        End If
                        Return False
                    End If
                End If
            End If

            'yang harus diisikan:
            '1. kode customer
            If String.IsNullOrEmpty(Me.mSessionID) = True Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input session id.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan session id.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If

            '3. Branch Details
            If Me.mPrimaryBranch = "-1" Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input branch code.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan branch code.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If
            If Me.mPrimaryBalanceCA = "-1" Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input balance cash account.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan server balance cash account.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If
            If Me.mPrimaryBeginningCA = "-1" Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input beginning cash account.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan server beginning cash account.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If
            If Me.mPrimaryCashierCA = "-1" Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input cashier cash account.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan cashier cash account.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If

            '4. Remarks
            If String.IsNullOrEmpty(Me.mRemarks) = False Then
                If vObjCD.CekDataKosong(Me.mRemarks) = False Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Data tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            End If

            'Cek kode akun
            If ListPrimaryAkun.Contains(Me.mPrimaryCashierCA) = False Then
                ListPrimaryAkun.Add(Me.mPrimaryCashierCA)
            Else
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Cashier cash account code has been used in another account code.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Cashier cash account code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If

            If ListPrimaryAkun.Contains(Me.mPrimaryBeginningCA) = False Then
                'ListPrimaryAkun.Add(Me.mPrimaryBeginningCA)
            Else
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Beginning cash account code has been used in another account code.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Beginning cash account code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If

            If ListPrimaryAkun.Contains(Me.mPrimaryBalanceCA) = False Then
                'ListPrimaryAkun.Add(Me.mPrimaryBalanceCA)
            Else
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Balance cash account code has been used in another account code.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Balance cash account code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If

            'Tanggal Transaksi berada dalam Periode yang sudah closing 
            'If ObjClosing.GetClosingDate.Date >= Me.TanggalTransaksi Then
            'MsgBox("Periode <= " + Format(ObjClosing.GetClosingDate.Date, "dd MMM yyyy") _
            '+ " sudah di kunci. Hubungi admin Anda.", MsgBoxStyle.Information, "Informasi")
            'Return False
            'End If

            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'ObjClosing = Nothing
            ListPrimaryAkun = Nothing
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictEdit As Boolean, ByVal vShowMsg As Boolean, ByVal vEnumBhs As EnumBahasa) As Boolean
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Dim pCreatedBy As String = ""
        'Dim ObjClosing As New ClsTransaksiClosing
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            Return True
            Exit Try

            '1. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictEdit = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If

                If pCreatedBy <> vUserCode Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("User can edit this data: " + pCreatedBy, _
                            MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Data ini hanya dapat diedit oleh: " + pCreatedBy, _
                            MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            'ObjClosing = Nothing
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, ByVal vEnumBhs As EnumBahasa, _
    ByVal i As FrmTrans_Session) As Boolean
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Dim pCreatedBy As String = ""
        Dim ObjPOS As New ClsTrans_POS
        'Dim ObjClosing As New ClsTransaksiClosing
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            '1. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictDelete = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If
                If pCreatedBy <> vUserCode Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("User can delete this data: " + pCreatedBy, _
                            MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Data ini hanya dapat didelete oleh: " + pCreatedBy, _
                            MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            End If

            '2. Trans POS
            If ObjPOS.GetJmlByPrimarySession(vPrimary) > 0 Then
                If vEnumBhs = EnumBahasa.Inggris Then
                    MsgBox("Data cannot be deleted because there are other transaction that refering this data", MsgBoxStyle.Information, "Information")
                ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                    MsgBox("Data tidak dapat dihapus karena ada transaksi lain yang merujuk", MsgBoxStyle.Information, "Informasi")
                End If
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            'ObjClosing = Nothing
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
        ByVal DGV As DataGridView, ByVal TBL As DseTrans_Session.tbltrans_session_cabang1DataTable, _
        ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabang1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = EnumSearchCriteria.BySessionID Then
                QueryResult = DTA.FillBy_SessionIDLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(1).HeaderText = "Session ID"
                DGV.Columns(2).Visible = False
                DGV.Columns(3).Visible = False
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy - HH:mm"
                DGV.Columns(3).DefaultCellStyle.Format = "dd MMM yyyy - HH:mm"
                DGV.Columns(4).HeaderText = "Remarks"
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

    Public Function CariDetail(ByVal vPrimary As Double, _
    ByVal vObjCD As ClsCekData, ByVal i As FrmTrans_Session, _
    ByVal TBLItem As DseTrans_Session.DGVPaymentDetailsDataTable) As Boolean
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabang2TableAdapter
        Dim TBL As New DseTrans_Session.tbltrans_session_cabang2DataTable
        Dim row As DseTrans_Session.tbltrans_session_cabang2Row
        Dim ObjItem As New ClsTrans_Session_Item
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    mPrimary = .primary
                    mSessionID = .session_id
                    mPrimaryTransServer = .primary_transaksi_server
                    mStatus = .status
                    mStartDate = .start_date
                    mEndDate = .end_date
                    mCashierCode = .Cashier_Code
                    mPrimaryBranch = .primary_branch
                    mRemarks = .remarks
                    mBeginningCash = .beginng_cash
                    mCashSales = .cash_sales
                    mDropCash = .drop_cash
                    mPayout = .pay_out
                    mBalance = .balance
                    mPrimaryBeginningCA = .primary_begining_cash_account
                    mPrimaryCashierCA = .primary_cashier_cash_account
                    mPrimaryBalanceCA = .primary_balance_cash_account
                    mPrimaryJurnalOpenSession = .primary_jurnal_open_session
                    mPrimaryJurnalCloseSession = .primary_jurnal_close_session
                    mStatusTransfer = EnumStatusTransfer.BelumTransfer
                    mStartTransferTime = .start_transfer_time
                    mEndTransferTime = .end_transfer_time

                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With
            End If
            ObjItem.IsiDGVSucceeded(vPrimary, i.dgv_DetailTrans, i, vObjCD)
            ObjItem.IsiDGVCancelled(vPrimary, i.dgv_cancelled_Trans, i, vObjCD)
            ObjItem.IsiDGVPaymentHeader(vPrimary, i, vObjCD, TBLItem)
            ObjItem.IsiDGVCancelledDetails(vPrimary, i.dgv_CancelledItem, i, vObjCD)
            ObjItem.IsiDGVPayout(vPrimary, i.dgv_PayOutS, i, vObjCD)
            ObjItem.IsiDGVPendingPOS(vPrimary, i.DGVPendingTrans, i, vObjCD)
            'Hitung Balance 
            Me.PerhitunganBalance(vPrimary, Me.mBeginningCash, i, vObjCD)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function CariDetailDataTransfer(ByVal vPrimary As Double, _
    ByVal vObjCD As ClsCekData, ByVal i As FrmTrans_Data_Transfer, _
    ByVal TBLItem As DseTrans_Session.DGVPaymentDetailsDataTable) As Boolean
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabang2TableAdapter
        Dim TBL As New DseTrans_Session.tbltrans_session_cabang2DataTable
        Dim row As DseTrans_Session.tbltrans_session_cabang2Row
        Dim ObjItem As New ClsTrans_Session_Item
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    mPrimary = .primary
                    mSessionID = .session_id
                    mPrimaryTransServer = .primary_transaksi_server
                    mStatus = .status
                    mStartDate = .start_date
                    mEndDate = .end_date
                    mCashierCode = .Cashier_Code
                    mPrimaryBranch = .primary_branch
                    mRemarks = .remarks
                    mBeginningCash = .beginng_cash
                    mCashSales = .cash_sales
                    mDropCash = .drop_cash
                    mPayout = .pay_out
                    mBalance = .balance
                    mPrimaryBeginningCA = .primary_begining_cash_account
                    mPrimaryCashierCA = .primary_cashier_cash_account
                    mPrimaryBalanceCA = .primary_balance_cash_account
                    mPrimaryJurnalOpenSession = .primary_jurnal_open_session
                    mPrimaryJurnalCloseSession = .primary_jurnal_close_session
                    mStatusTransfer = EnumStatusTransfer.BelumTransfer
                    mStartTransferTime = .start_transfer_time
                    mEndTransferTime = .end_transfer_time

                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With
            End If
            ObjItem.IsiDGVSucceededDataTransfer(vPrimary, i.dgv_DetailTrans, i, vObjCD)
            ObjItem.IsiDGVCancelledDataTransfer(vPrimary, i.dgv_cancelled_Trans, i, vObjCD)
            ObjItem.IsiDGVPaymentHeaderDataTransfer(vPrimary, i, vObjCD, TBLItem)
            ObjItem.IsiDGVCancelledDetailsDataTransfer(vPrimary, i.dgv_CancelledItem, i, vObjCD)
            ObjItem.IsiDGVPayoutDataTransfer(vPrimary, i.dgv_PayOutS, i, vObjCD)
            ObjItem.IsiDGVPendingPOSDataTransfer(vPrimary, i.DGVPendingTrans, i, vObjCD)
            'Hitung Balance 
            Me.PerhitunganBalanceDataTransfer(Me.mBeginningCash, i, vObjCD)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function GetBeginningCash(ByVal vPrimary As Double) As Decimal
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            If DTA.ScalarQuery_GetJmlBeginningCash(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetJmlBeginningCash(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNilaiBalance(ByVal vPrimarySession As Double, ByVal Objcd As ClsCekData) As Decimal
        Dim ObjPOS As New ClsTrans_POS
        Dim BeginningCashSales As Decimal
        Dim CashSales As Decimal
        Dim DropCash As Decimal
        Dim Payouts As Decimal
        Dim Withdrawal As Decimal
        Dim PayoutsWithdrawal As Decimal
        Dim Retval As Decimal
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_pos_cabang_item_paymentTableAdapter
        Dim TBL As New DseTrans_Session.tbltrans_pos_cabang_item_paymentDataTable
        Dim Row As DseTrans_Session.tbltrans_pos_cabang_item_paymentRow
        Dim TBL11 As New DseTrans_Session.tbltrans_pos_cabang_item_paymentDataTable
        Dim DTAPayout As New DseTrans_SessionTableAdapters.DataTable1TableAdapter
        Dim TBLPayout As New DseTrans_Session.DataTable1DataTable
        Dim rowPayout As DseTrans_Session.DataTable1Row
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTAPayout.Connection.ConnectionString = Me.GetConstring

            DTA.FillBy_PrimarySessionYgCash(TBL, vPrimarySession)
            DTA.FillBy_PrimarySessionYgWithdrawal(TBL11, vPrimarySession)
            DTAPayout.Fill(TBLPayout, vPrimarySession)

            BeginningCashSales = Me.GetBeginningCash(vPrimarySession)
            'Cash Sales (ambil nilai yang cash saja)
            For Each Row In TBL.Rows
                CashSales += CDec(Row.amount) - CDec(ObjPOS.GetChange(Row.primary_transaksi))
            Next
            'DropCash & Payout
            For Each rowPayout In TBLPayout.Rows
                If rowPayout.tipe_pay_out_drop_cash = ClsMast_PayoutDropCash.EnumTipe.Drop_Cash Then
                    DropCash += CDec(rowPayout.amount)
                Else
                    Payouts += CDec(rowPayout.amount)
                End If
            Next
            For Each Row In TBL11.Rows
                Withdrawal += Row.amount
            Next
            PayoutsWithdrawal = Payouts + Withdrawal
            'Balance
            Retval = CDec(BeginningCashSales) + CDec(CashSales) _
            + CDec(DropCash) - CDec(PayoutsWithdrawal)

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            TBL11.Dispose()
            DTAPayout.Dispose()
            TBLPayout.Dispose()
            ObjPOS = Nothing
        End Try
    End Function

    Public Function PerhitunganBalance(ByVal vPrimarySession As Double, ByVal BeginningCashSales As Decimal, _
    ByVal i As FrmTrans_Session, ByVal Objcd As ClsCekData) As Boolean
        Dim ObjPOS As New ClsTrans_POS
        Dim RowDGV As DataGridViewRow
        Dim CashSales As Decimal
        Dim DropCash As Decimal
        Dim Payouts As Decimal
        Dim Withdrawal As Decimal
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_pos_cabang_item_paymentTableAdapter
        Dim TBL As New DseTrans_Session.tbltrans_pos_cabang_item_paymentDataTable
        Dim Row As DseTrans_Session.tbltrans_pos_cabang_item_paymentRow
        Dim TBL11 As New DseTrans_Session.tbltrans_pos_cabang_item_paymentDataTable
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.FillBy_PrimarySessionYgCash(TBL, vPrimarySession)
            DTA.FillBy_PrimarySessionYgWithdrawal(TBL11, vPrimarySession)

            i.TxtCash_Sales.Clear()
            i.TxtDropCash.Clear()
            i.TxtPayouts.Clear()
            i.TxtBalance.Clear()
            'Cash Sales (ambil nilai yang cash saja)
            For Each Row In TBL.Rows
                CashSales += CDec(Row.amount) - CDec(ObjPOS.GetChange(Row.primary_transaksi))
            Next
            i.TxtCash_Sales.Text = CashSales
            'DropCash & Payout
            For Each RowDGV In i.dgv_PayOutS.Rows
                If RowDGV.Cells(i.Col_PayOutD_Tipe.Index).Value = "Drop cash" Then
                    DropCash += CDec(RowDGV.Cells(i.Col_PayOutD_Amount.Index).Value)
                Else
                    Payouts += CDec(RowDGV.Cells(i.Col_PayOutD_Amount.Index).Value)
                End If
            Next
            i.TxtDropCash.Text = DropCash
            For Each Row In TBL11.Rows
                Withdrawal += Row.amount
            Next
            i.TxtPayouts.Text = Payouts + Withdrawal
            'Balance
            i.TxtBalance.Text = Objcd.DigitGrouping((CDec(BeginningCashSales) + CDec(i.TxtCash_Sales.Text) _
            + CDec(i.TxtDropCash.Text)) - CDec(i.TxtPayouts.Text), i.DefPrec, i.RoundingTipe)
            i.TxtCash_Sales.Text = Objcd.DigitGrouping(i.TxtCash_Sales.Text, i.DefPrec, i.RoundingTipe)
            i.TxtDropCash.Text = Objcd.DigitGrouping(i.TxtDropCash.Text, i.DefPrec, i.RoundingTipe)
            i.TxtPayouts.Text = Objcd.DigitGrouping(i.TxtPayouts.Text, i.DefPrec, i.RoundingTipe)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            TBL11.Dispose()
            ObjPOS = Nothing
        End Try
    End Function

    Public Function PerhitunganBalanceDataTransfer(ByVal BeginningCashSales As Decimal, _
    ByVal i As FrmTrans_Data_Transfer, ByVal Objcd As ClsCekData) As Boolean
        Dim ObjPOS As New ClsTrans_POS
        Dim RowDGV As DataGridViewRow
        Dim CashSales As Decimal
        Dim DropCash As Decimal
        Dim Payouts As Decimal
        Dim Withdrawal As Decimal
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_pos_cabang_item_paymentTableAdapter
        Dim TBL As New DseTrans_Session.tbltrans_pos_cabang_item_paymentDataTable
        Dim Row As DseTrans_Session.tbltrans_pos_cabang_item_paymentRow
        Dim TBL11 As New DseTrans_Session.tbltrans_pos_cabang_item_paymentDataTable
        Try
            i.TxtCash_Sales.Clear()
            i.TxtDropCash.Clear()
            i.TxtPayouts.Clear()
            i.TxtBalance.Clear()

            DTA.FillBy_PrimarySessionYgCash(TBL, i.TxtPrimary.Text)
            DTA.FillBy_PrimarySessionYgWithdrawal(TBL11, i.TxtPrimary.Text)
            'Cash Sales (ambil nilai yang cash saja)
            For Each Row In TBL.Rows
                CashSales += CDec(Row.amount) - CDec(ObjPOS.GetChange(Row.primary_transaksi))
            Next
            'For Each RowDGV In i.dgv_PaymentD.Rows
            'CashSales += CDec(RowDGV.Cells(i.Col_TransD_Payment.Index).Value) - CDec(ObjPOS.GetChange(RowDGV.Cells(i.Col_TransD_PrimaryPOS.Index).Value))
            'Next
            i.TxtCash_Sales.Text = CashSales
            'DropCash & Payout
            For Each RowDGV In i.dgv_PayOutS.Rows
                If RowDGV.Cells(i.Col_PayOutD_Tipe.Index).Value = "Drop cash" Then
                    DropCash += CDec(RowDGV.Cells(i.Col_PayOutD_Amount.Index).Value)
                Else
                    Payouts += CDec(RowDGV.Cells(i.Col_PayOutD_Amount.Index).Value)
                End If
            Next
            i.TxtDropCash.Text = DropCash
            For Each Row In TBL11.Rows
                Withdrawal += Row.amount
            Next
            i.TxtPayouts.Text = Payouts + Withdrawal
            'Balance
            i.TxtBalance.Text = Objcd.DigitGrouping((CDec(BeginningCashSales) + CDec(i.TxtCash_Sales.Text) _
            + CDec(i.TxtDropCash.Text)) - CDec(i.TxtPayouts.Text), i.DefPrec, i.RoundingTipe)
            i.TxtCash_Sales.Text = Objcd.DigitGrouping(i.TxtCash_Sales.Text, i.DefPrec, i.RoundingTipe)
            i.TxtDropCash.Text = Objcd.DigitGrouping(i.TxtDropCash.Text, i.DefPrec, i.RoundingTipe)
            i.TxtPayouts.Text = Objcd.DigitGrouping(i.TxtPayouts.Text, i.DefPrec, i.RoundingTipe)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            TBL11.Dispose()
            ObjPOS = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As EnumSearchCriteria) As Double
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabang1TableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If vSearchBy = EnumSearchCriteria.ByRefName Then
                Return DTA.ScalarQuery_GetJmlBySessionIDLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub New()
        Try
            Me.DTAAntrian.Connection.ConnectionString = Me.GetConstring
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, _
    ByVal vObjCD As ClsCekData, ByVal vPrimary As Double, ByVal vEnumBhs As EnumBahasa, _
    ByVal i As FrmTrans_Session) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Try
            'ambil no antrian
            Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.SimpanBaru1(vUserCode, vPrimary, vObjCD, _
                        vEnumBhs, i)

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
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanBaru1(ByVal vUserCode As String, _
    ByVal vPrimary As Double, ByVal vObjCD As ClsCekData, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmTrans_Session) As Boolean
        Dim DTA1 As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabang2TableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA1.Connection.ConnectionString = Me.GetConstring

            If Me.ApakahDataBenar(True, vPrimary, True, vObjCD, vEnumBhs, i) = True Then
                Me.mUrut = CInt(Me.mSessionID.Substring(0, 6))
                If DTA.Insert(Me.mSessionID, Me.mPrimaryTransServer, Me.mStatus, Me.mStartDate, Me.mEndDate, _
                Me.mCashierCode, Me.mPrimaryBranch, Me.mRemarks, Me.mBeginningCash, Me.mCashSales, _
                Me.mDropCash, Me.mPayout, Me.mBalance, Me.mPrimaryBeginningCA, Me.mPrimaryCashierCA, _
                 Me.mPrimaryBalanceCA, Me.mPrimaryJurnalOpenSession, Me.mPrimaryJurnalCloseSession, _
                 vUserCode, Now, vUserCode, Now, Now, Me.mStatusTransfer, Me.mStartTransferTime, Me.mEndTransferTime, Me.mUrut) > 0 Then
                    Me.mPrimary = DTA1.ScalarQuery_GetPrimary(Me.mSessionID)
                    Me.mTimestamp = DTA1.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

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
            DTA1.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, _
    ByVal vObjCD As ClsCekData, ByVal vPrimary As Double, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmTrans_Session) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Try
            'ambil no antrian
            Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.SimpanPerubahan1(vUserCode, vPrimary, vObjCD, vEnumBhs, i)

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
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vUserCode As String, _
    ByVal vPrimary As Double, ByVal vObjCD As ClsCekData, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmTrans_Session) As Boolean
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Dim DTA1 As New DseTrans_SessionTableAdapters.tbltrans_session_cabang2TableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA1.Connection.ConnectionString = Me.GetConstring

            If Me.ApakahDataBenar(False, vPrimary, True, vObjCD, vEnumBhs, i) = True Then
                Me.mUrut = Me.mSessionID.Substring(0, 6)
                If DTA1.Update(Me.mSessionID, Me.mPrimaryTransServer, Me.mStatus, Me.mStartDate, Me.mEndDate, _
                Me.mCashierCode, Me.mPrimaryBranch, Me.mRemarks, Me.mBeginningCash, Me.mCashSales, _
                Me.mDropCash, Me.mPayout, Me.mBalance, Me.mPrimaryBeginningCA, Me.mPrimaryCashierCA, _
                Me.mPrimaryBalanceCA, Me.mPrimaryJurnalOpenSession, Me.mPrimaryJurnalCloseSession, _
                vUserCode, Now, Now, Me.mStatusTransfer, Me.mStartTransferTime, Me.mEndTransferTime, _
                Me.mUrut, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmTrans_Session) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Try
            'ambil no antrian
            Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.Hapus1(vPrimary, vUserCode, _
                        vStrictDelete, vShowMsg, vEnumBhs, i)
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
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmTrans_Session) As Boolean
        Dim DTA1 As New DseTrans_SessionTableAdapters.tbltrans_session_cabang2TableAdapter
        Try
            DTA1.Connection.ConnectionString = Me.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUserCode, _
                vStrictDelete, vShowMsg, vEnumBhs, i) = True Then
                'Hapus
                DTA1.Delete(vPrimary)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
        End Try
    End Function

    Public Sub ResetObject(ByVal i As FrmTrans_Session)
        Try
            mPrimary = -1
            mSessionID = ""
            mPrimaryTransServer = -1
            mStatus = EnumStatusSession.Open
            mStartDate = "1950/01/01 00:00:00"
            mEndDate = "1950/01/01 00:00:00"
            mCashierCode = ""
            mPrimaryBranch = -1
            mRemarks = ""
            mBeginningCash = 0
            mCashSales = 0
            mDropCash = 0
            mPayout = 0
            mBalance = 0
            mPrimaryBeginningCA = -1
            mPrimaryCashierCA = -1
            mPrimaryBalanceCA = -1
            mPrimaryJurnalOpenSession = -1
            mPrimaryJurnalCloseSession = -1
            mStatusTransfer = EnumStatusTransfer.BelumTransfer
            mStartTransferTime = "1950/01/01 00:00:00"
            mEndTransferTime = "1950/01/01 00:00:00"


            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today

            i.dgv_DetailTrans.Rows.Clear()
            i.dgv_cancelled_Trans.Rows.Clear()
            i.dgv_PaymentD.Rows.Clear()
            i.dgv_Payment_Sum.Rows.Clear()
            i.dgv_CancelledItem.Rows.Clear()
            i.dgv_PayOutS.Rows.Clear()
            i.DGVPendingTrans.Rows.Clear()

            i.Txt_TransDetail_TotalTransQty.Clear()
            i.TxtSalesItemQty.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetStartDate(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetStartDate(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetStartDate(vPrimary)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByKasirNBranch(ByVal vPrimaryKasir As Double, ByVal vPrimaryBranch As Double) As Double
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabang2TableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByKodeKasirNBranch(vPrimaryBranch, vPrimaryKasir).HasValue Then
                Return DTA.ScalarQuery_GetJmlByKodeKasirNBranch(vPrimaryBranch, vPrimaryKasir).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiCBOSessionID(ByVal vCashierCode As String, _
    ByVal vCboSession As ComboBox, ByVal vTBL As DseTrans_Session.tbltrans_session_cabang1_3DataTable, _
    ByVal DTA As DseTrans_SessionTableAdapters.tbltrans_session_cabang1_3TableAdapter)
        Try
            'get connection
            DTA.Connection.ConnectionString = Me.GetConstring
            'kosongkan cbo
            vCboSession.DataSource = Nothing
            vCboSession.Items.Clear()

            'isi tabel
            DTA.Fill(vTBL, vCashierCode)

            'hubungkan dengan cbo
            vCboSession.DataSource = vTBL.DefaultView
            vTBL.DefaultView.Sort = "[session id] ASC"
            vCboSession.DisplayMember = "session id"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetJmlByCashierCode(ByVal vUserCode As String) As Double
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabang2TableAdapter
        Try
            'get connection
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByCashierCode(vUserCode).HasValue Then
                Return DTA.ScalarQuery_GetJmlByCashierCode(vUserCode).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryByCashierCode(ByVal vUserCode As String) As Double
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabang2TableAdapter
        Try
            'get connection
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimaryByCashierCode(vUserCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryByCashierCode(vUserCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetStatus(ByVal vPrimary As Double) As EnumStatusSession
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            '0. Open
            '1. Close
            If DTA.ScalarQuery_GetStatus(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetStatus(vPrimary).Value
            Else
                Return EnumStatusSession.Open
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTotalTransaksiPOS(ByVal vPrimarySession As Double) As Double
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetTotalTransaksiPOS(vPrimarySession).HasValue Then
                Return DTA.ScalarQuery_GetTotalTransaksiPOS(vPrimarySession).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTotalQtyPOS(ByVal vPrimarySession As Double) As Decimal
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetTotalQtyPOSByPrimarySession(vPrimarySession).HasValue Then
                Return DTA.ScalarQuery_GetTotalQtyPOSByPrimarySession(vPrimarySession).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlPOSBelumPaid(ByVal vPrimarySession As Double) As Decimal
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlPOSYangBelumPaidByPrimarySession(vPrimarySession).HasValue Then
                Return DTA.ScalarQuery_GetJmlPOSYangBelumPaidByPrimarySession(vPrimarySession).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SetStatusSession(ByVal vStatus As ClsTrans_Session.EnumStatusSession, ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabang2TableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            DTA.UpdateQuery_SetStatusSession(vStatus, vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function LoadDataPOSBuatSession(ByVal lbo As ListBox, _
    ByVal TBL As DseTrans_Session.tbltrans_pos_cabang1DataTable, _
    ByVal vPrimarySession As Double) As Boolean
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_pos_cabang1TableAdapter
        Dim Row As DseTrans_Session.tbltrans_pos_cabang1Row
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.Fill(TBL, vPrimarySession)
            lbo.DataSource = TBL.DefaultView
            For Each Row In TBL.Rows
                Row.Combination = Row.no_trans.ToString
            Next
            lbo.DisplayMember = "Combination"
            lbo.ColumnWidth = 180
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJumlahUntukValidasiSession(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_pos_cabang1TableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimarySession(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimarySession(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPrimary(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimary(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimary(vPrimary).Value
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
