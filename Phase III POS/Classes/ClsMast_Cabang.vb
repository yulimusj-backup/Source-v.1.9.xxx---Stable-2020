Option Explicit On
Option Strict Off
Imports Core_Module

Public Class ClsMast_Cabang
    Private mPrimary As Double = -1
    Private mCode As String = ""
    Private mDesc As String = ""
    Private mRemarks As String = ""
    Private mMainBranch As Integer = 0
    Private mServerName As String = ""
    Private mUserId As String = ""
    Private mPassword As String = ""
    Private mDatabaseName As String = ""
    Private mPort As String = ""

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

    Public Property MainBranch() As Integer
        Get
            Return mMainBranch
        End Get
        Set(ByVal value As Integer)
            mMainBranch = value
        End Set
    End Property

    Public Property ServerName() As String
        Get
            Return mServerName
        End Get
        Set(ByVal value As String)
            mServerName = value
        End Set
    End Property

    Public Property UserID() As String
        Get
            Return mUserId
        End Get
        Set(ByVal value As String)
            mUserId = value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return mPassword
        End Get
        Set(ByVal value As String)
            mPassword = value
        End Set
    End Property

    Public Property DatabaseName() As String
        Get
            Return mDatabaseName
        End Get
        Set(ByVal value As String)
            mDatabaseName = value
        End Set
    End Property

    Public Property Port() As String
        Get
            Return mPort
        End Get
        Set(ByVal value As String)
            mPort = value
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

    Public Function GetPrimary(ByVal vKode As String) As Double
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimary(vKode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vKode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCode(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetCode(vPrimary)
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
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
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

    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDesc(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetDesc(vPrimary)
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
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
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

    Public Function GetJmlByCode(ByVal vCode As String) As Double
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByCode(vCode).HasValue Then
                Return DTA.ScalarQuery_GetJmlByCode(vCode).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetServer(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetServerName(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetServerName(vPrimary)
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

    Public Function GetUserID(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetUserID(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetUserID(vPrimary)
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

    Public Function GetPassword(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetPassword(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetPassword(vPrimary)
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

    Public Function GetDatabaseName(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDatabaseName(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetDatabaseName(vPrimary)
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

    Public Function GetPort(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetPort(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetPort(vPrimary)
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

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByCode(vCode).Value > 0 Then
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
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_Cabang) As Boolean
        'Dim ObjClosing As New ClsTransaksiClosing
        Try
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mCode) = True Then
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
                Dim NoTransaksiLama As String = Me.GetCode(Me.mPrimary)
                If NoTransaksiLama <> Me.mCode Then
                    If Me.ApakahSudahAda(Me.mCode) = True Then
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
            If String.IsNullOrEmpty(Me.mCode) = True Then
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
            If String.IsNullOrEmpty(Me.mDesc) = True Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input description.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan decription.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If

            '2b. Remarks
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

            '3. Branch Details
            If String.IsNullOrEmpty(Me.mServerName) = True Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input server name.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan server name.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If
            If String.IsNullOrEmpty(Me.mUserId) = True Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input user id.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan server user id.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If
            If String.IsNullOrEmpty(Me.mPassword) = True Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input password.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan server password.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If
            If String.IsNullOrEmpty(Me.mDatabaseName) = True Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input database name.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan server database name.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If
            If String.IsNullOrEmpty(Me.mPort) = True Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input port.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan server port.", MsgBoxStyle.Information, "Informasi")
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
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictEdit As Boolean, ByVal vShowMsg As Boolean, ByVal vEnumBhs As EnumBahasa) As Boolean
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
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

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            'ObjClosing = Nothing
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, ByVal vEnumBhs As EnumBahasa, _
    ByVal i As FrmMast_Cabang) As Boolean
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Dim pCreatedBy As String = ""
        'Dim ObjClosing As New ClsTransaksiClosing
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            Return True
            Exit Try

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

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            'ObjClosing = Nothing
        End Try
    End Function

    Public Function GetMainBranch(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetMainBranch(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetMainBranch(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
     ByVal DGV As DataGridView, ByVal TBL As DseMast_Cabang.tblmaster_cabang1DataTable, _
     ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabang1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescriptionLIKE(TBL, Criteria + "%")
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

    Public Sub IsiCheckListBox(ByVal vClb As CheckedListBox, ByVal vTBL As DseMast_Cabang.tblmaster_cabang2DataTable)
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabang2TableAdapter

        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.Fill(vTBL)
            vClb.DataSource = vTBL.DefaultView
            vClb.DisplayMember = "Cabang"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function CariDetail(ByVal vPrimary As Double, _
    ByVal vObjCD As ClsCekData, ByVal i As FrmMast_Cabang) As Boolean
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Dim TBL As New DseMast_Cabang.tblmaster_cabangDataTable
        Dim row As DseMast_Cabang.tblmaster_cabangRow
        Dim ObjItem As New ClsMast_Cabang_Item
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mCode = .Code
                    Me.mDesc = .Description
                    Me.mMainBranch = .Main_Branch
                    Me.mRemarks = .Remarks
                    Me.mServerName = .Server_Name
                    Me.mUserId = .User_Id
                    Me.mPassword = .Password
                    Me.mDatabaseName = .Database_Name
                    Me.mPort = .Port

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
            End If

            'cari data item
            ObjItem.IsiDGV(Primary, i.dgvLocation, i, vObjCD)

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
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabang1TableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                Return DTA.ScalarQuery_GetJmlByCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                Return DTA.ScalarQuery_GetJmlByDescriptionLIKE(vKriteria + "%").Value
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
    ByVal i As FrmMast_Cabang) As Boolean
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
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_Cabang) As Boolean
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Dim ObjItem As New ClsMast_Cabang_Item
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If Me.ApakahDataBenar(True, vPrimary, True, vObjCD, vEnumBhs, i) = True Then
                If DTA.Insert(Me.mCode, Me.mDesc, Me.mRemarks, Me.mMainBranch, _
                Me.mServerName, Me.mUserId, Me.mPassword, Me.mDatabaseName, Me.mPort, _
                vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mCode)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, i.dgvLocation, i)
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

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, _
    ByVal vObjCD As ClsCekData, ByVal vPrimary As Double, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_Cabang) As Boolean
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
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_Cabang) As Boolean
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Dim ObjItem As New ClsMast_Cabang_Item
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If Me.ApakahDataBenar(False, vPrimary, True, vObjCD, vEnumBhs, i) = True Then
                If DTA.Update(Me.mCode, Me.mDesc, Me.mRemarks, Me.mMainBranch, _
                              Me.mServerName, Me.mUserId, Me.mPassword, Me.mDatabaseName, Me.mPort, _
                              vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)
                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, i.dgvLocation, i)
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

    Public Overridable Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_Cabang) As Boolean
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
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_Cabang) As Boolean
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Dim ObjItem As New ClsMast_Cabang_Item
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUserCode, _
                vStrictDelete, vShowMsg, vEnumBhs, i) = True Then
                'Hapus item
                ObjItem.HapusItem(vPrimary)
                'Hapus
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

    Public Sub ResetObject(ByVal i As FrmMast_Cabang)
        Try
            mPrimary = -1
            mCode = ""
            mDesc = ""
            mMainBranch = 0
            mRemarks = ""

            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today

            i.dgvLocation.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetJmlByServerDBName(ByVal vServer As String, ByVal vDatabaseName As String) As Double
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByServerDBName(vServer, vDatabaseName).HasValue Then
                Return DTA.ScalarQuery_GetJmlByServerDBName(vServer, vDatabaseName).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByMainBranch() As Double
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByMainBranch().HasValue Then
                Return DTA.ScalarQuery_GetJmlByMainBranch().Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByMainBranchPosisiEdit(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByMainBranchPosisiEdit(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetJmlByMainBranchPosisiEdit(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByServerDBNamePosisiEdit(ByVal vServer As String, ByVal vDatabaseName As String, _
    ByVal vPrimary As Double) As Double
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByServerDBNamePosisiEdit(vServer, vDatabaseName, vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetJmlByServerDBNamePosisiEdit(vServer, vDatabaseName, vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlUntukCabangSaja() As Double
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlBuatYgCabangAja().HasValue Then
                Return DTA.ScalarQuery_GetJmlBuatYgCabangAja().Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetServerNameUtama() As String
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetServerNameUtama()) = False Then
                Return DTA.ScalarQuery_GetServerNameUtama()
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetUserIDUtama() As String
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetUserIDUtama()) = False Then
                Return DTA.ScalarQuery_GetUserIDUtama()
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetPasswordUtama() As String
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetPasswordUtama()) = False Then
                Return DTA.ScalarQuery_GetPasswordUtama()
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetPortUtama() As String
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetPortUtama()) = False Then
                Return DTA.ScalarQuery_GetPortUtama()
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetDatabaseNameUtama() As String
        Dim DTA As New DseMast_CabangTableAdapters.tblmaster_cabangTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDatabaseNameUtama()) = False Then
                Return DTA.ScalarQuery_GetDatabaseNameUtama()
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
End Class
