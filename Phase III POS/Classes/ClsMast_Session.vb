Option Explicit On
Option Strict Off
Imports Core_Module

Public Class ClsMast_Session
    Private mPrimary As Double = -1
    Private mRefName As String = ""
    Private mSessionPeriod As DateTime = "1950/01/01 00:00:00"
    Private mSessionThrough As DateTime = "1950/01/01 00:00:00"
    Private mCashierCode As String = ""
    Private mPrimaryBranch As Double = -1
    Private mRemarks As String = ""
    Private mBeginningCash As Decimal = 0
    Private mPrimaryBeginningCA As Double = -1
    Private mPrimaryCashierCA As Double = -1
    Private mPrimaryBalanceCA As Double = -1

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New DseTransaksiAntrianTableAdapters.tbltrans_antrianTableAdapter

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property RefName() As String
        Get
            Return mRefName
        End Get
        Set(ByVal value As String)
            mRefName = value
        End Set
    End Property

    Public Property SessionPeriod() As DateTime
        Get
            Return mSessionPeriod
        End Get
        Set(ByVal value As DateTime)
            mSessionPeriod = value
        End Set
    End Property

    Public Property SessionThrough() As DateTime
        Get
            Return mSessionThrough
        End Get
        Set(ByVal value As DateTime)
            mSessionThrough = value
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

    Public Property BeginningCash() As Decimal
        Get
            Return mBeginningCash
        End Get
        Set(ByVal value As Decimal)
            mBeginningCash = value
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

    Public Property PrimaryBeginningCashAccount() As Double
        Get
            Return mPrimaryBeginningCA
        End Get
        Set(ByVal value As Double)
            mPrimaryBeginningCA = value
        End Set
    End Property

    Public Property PrimaryCashierCashAccount() As Double
        Get
            Return mPrimaryCashierCA
        End Get
        Set(ByVal value As Double)
            mPrimaryCashierCA = value
        End Set
    End Property

    Public Property PrimaryBalanceCashAccount() As Double
        Get
            Return mPrimaryBalanceCA
        End Get
        Set(ByVal value As Double)
            mPrimaryBalanceCA = value
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

    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307"
        End Try
    End Function

    Public Function GetEndDateSession(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMast_SessionTableAdapters.tblmaster_sessionTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetSessionThroughDate(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetSessionThroughDate(vPrimary).Value
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

    ''Coding standarnya 
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

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMast_SessionTableAdapters.tblmaster_sessionTableAdapter
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

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseMast_SessionTableAdapters.tblmaster_sessionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimary(vNoTransaksi).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vNoTransaksi).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCashier(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMast_SessionTableAdapters.tblmaster_sessionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimaryCashier(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryCashier(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryBranch(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMast_SessionTableAdapters.tblmaster_sessionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimaryBranch(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryBranch(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByRefName(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseMast_SessionTableAdapters.tblmaster_sessionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByRefName(vNoTransaksi).HasValue Then
                Return DTA.ScalarQuery_GetJmlByRefName(vNoTransaksi).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetRefName(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_SessionTableAdapters.tblmaster_sessionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetRefName(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetRefName(vPrimary)
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

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_SessionTableAdapters.tblmaster_sessionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
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

    Public Function ApakahSudahAda(ByVal vRefName As String) As Boolean
        Dim DTA As New DseMast_SessionTableAdapters.tblmaster_sessionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByRefName(vRefName).Value > 0 Then
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

    ''' <summary>
    ''' data ini digunakan untuk mengecek kebenaran data yang diberikan
    ''' </summary>
    ''' <param name="vBaru"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vPrimary As Double, _
    ByVal vShowMsg As Boolean, ByVal vObjCD As ClsCekData, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_Session) As Boolean
        'Dim ObjClosing As New ClsTransaksiClosing
        Dim ListPrimaryAkun As New List(Of Double)
        Try
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mRefName) = True Then
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
            Else
                Dim NoTransaksiLama As String = Me.GetRefName(Me.mPrimary)
                If NoTransaksiLama <> Me.mRefName Then
                    If Me.ApakahSudahAda(Me.mRefName) = True Then
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
            If String.IsNullOrEmpty(Me.mRefName) = True Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input code.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan code.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If

            '2. Description
            If String.IsNullOrEmpty(Me.mCashierCode) = True Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input cashier code.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan cashier code.", MsgBoxStyle.Information, "Informasi")
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
                ''ListPrimaryAkun.Add(Me.mPrimaryBeginningCA)
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
                ''ListPrimaryAkun.Add(Me.mPrimaryBalanceCA)
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
        Dim DTA As New DseMast_SessionTableAdapters.tblmaster_sessionTableAdapter
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
    ByVal i As FrmMast_Session) As Boolean
        Dim DTA As New DseMast_SessionTableAdapters.tblmaster_sessionTableAdapter
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

            '2. transaksi POS
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
            ObjPOS = Nothing
            'ObjClosing = Nothing
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
        ByVal DGV As DataGridView, ByVal TBL As DseMast_Session.DataTable1DataTable, _
        ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMast_SessionTableAdapters.DataTable1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = EnumSearchCriteria.ByRefName Then
                QueryResult = DTA.FillBy_RefNumberLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCashierCode Then
                QueryResult = DTA.FillBy_CashierCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByBranchCode Then
                QueryResult = DTA.FillBy_BranchCodeLIKE(TBL, Criteria + "%")
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
                DGV.Columns(1).HeaderText = "Reference Name"
                DGV.Columns(2).DefaultCellStyle.Format = "HH:mm"
                DGV.Columns(3).DefaultCellStyle.Format = "HH:mm"
                DGV.Columns(4).HeaderText = "Cashier Code"
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
    ByVal vObjCD As ClsCekData, ByVal i As FrmMast_Session) As Boolean
        Dim DTA As New DseMast_SessionTableAdapters.tblmaster_session1TableAdapter
        Dim TBL As New DseMast_Session.tblmaster_session1DataTable
        Dim row As DseMast_Session.tblmaster_session1Row
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mRefName = .Reference_Name
                    Me.mSessionPeriod = .Session_Period
                    Me.mSessionThrough = .Session_Through
                    Me.mRemarks = .Remarks
                    Me.mBeginningCash = .Beginning_Cash
                    Me.mCashierCode = .Cashier_Code
                    Me.mPrimaryBranch = .Primary_Branch
                    Me.mPrimaryBeginningCA = .Primary_Begining_Cash_Account
                    Me.mPrimaryCashierCA = .Primary_Cashier_Cash_Account
                    Me.mPrimaryBalanceCA = .Primary_Balance_Cash_Account

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As EnumSearchCriteria) As Double
        Dim DTA As New DseMast_SessionTableAdapters.DataTable1TableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If vSearchBy = EnumSearchCriteria.ByRefName Then
                Return DTA.ScalarQuery_GetJmlByRefNameLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByCashierCode Then
                Return DTA.ScalarQuery_GetJmlByCashierCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByBranchCode Then
                Return DTA.ScalarQuery_GetJmlByBranchCodeLIKE(vKriteria + "%")
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
    ByVal vObjCD As ClsCekData, _
    ByVal vPrimary As Double, ByVal vEnumBhs As EnumBahasa, _
    ByVal i As FrmMast_Session) As Boolean
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
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_Session) As Boolean
        Dim DTA As New DseMast_SessionTableAdapters.tblmaster_session1TableAdapter
        Dim DTA1 As New DseMast_SessionTableAdapters.tblmaster_sessionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA1.Connection.ConnectionString = Me.GetConstring

            If Me.ApakahDataBenar(True, vPrimary, True, vObjCD, vEnumBhs, i) = True Then
                If DTA.Insert(Me.mRefName, Me.mSessionPeriod, Me.mSessionThrough, Me.mCashierCode, Me.mPrimaryBranch, _
                Me.mRemarks, Me.mBeginningCash, Me.mPrimaryBeginningCA, Me.mPrimaryCashierCA, Me.mPrimaryBalanceCA, _
                vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    Me.mPrimary = DTA1.ScalarQuery_GetPrimary(Me.mRefName)
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
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_Session) As Boolean
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
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_Session) As Boolean
        Dim DTA As New DseMast_SessionTableAdapters.tblmaster_sessionTableAdapter
        Dim DTA1 As New DseMast_SessionTableAdapters.tblmaster_session1TableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA1.Connection.ConnectionString = Me.GetConstring

            If Me.ApakahDataBenar(False, vPrimary, True, vObjCD, vEnumBhs, i) = True Then
                If DTA1.Update(Me.mRefName, Me.mSessionPeriod, Me.mSessionThrough, Me.mCashierCode, Me.mPrimaryBranch, _
                Me.mRemarks, Me.mBeginningCash, Me.mPrimaryBeginningCA, Me.mPrimaryCashierCA, Me.mPrimaryBalanceCA, _
                              vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
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
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_Session) As Boolean
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
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_Session) As Boolean
        Dim DTA1 As New DseMast_SessionTableAdapters.tblmaster_session1TableAdapter
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

    Public Sub ResetObject()
        Try
            mPrimary = -1
            mRefName = ""
            mSessionPeriod = "1950/01/01 00:00:00"
            mSessionThrough = "1950/01/01 00:00:00"
            mCashierCode = ""
            mPrimaryBranch = -1
            mBeginningCash = 0
            mPrimaryBeginningCA = -1
            mPrimaryCashierCA = -1
            mPrimaryBalanceCA = -1
            mRemarks = ""

            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class
