Option Strict Off
Option Explicit On

Public Class ClsMasterUserGroup

    'setelah simpan group tertentu, kelas ini juga harus melakukan update terhadap hak akses 
    'yang dimiliki oleh user yang merujuk kepadanya

    Private mPrimary As Double = -1
    Private mGroupCode As String = ""
    Private mDescription As String = ""

    Private mCreatedBy As String
    Private mCreatedDate As Date
    Private mModifiedBy As String
    Private mModifiedDate As Date
    Private mTimeStamp As Date

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property GroupCode() As String
        Get
            Return mGroupCode
        End Get
        Set(ByVal value As String)
            mGroupCode = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return mDescription
        End Get
        Set(ByVal value As String)
            mDescription = value
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

    Public Property TimeStamp() As Date
        Get
            Return mTimeStamp
        End Get
        Set(ByVal value As Date)
            mTimeStamp = value
        End Set
    End Property

    ''' <summary>
    ''' Untuk Cek No PO
    ''' </summary>
    ''' <param name="vNoTransaksi"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmaster_usergroupTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi).Value > 0 Then
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

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
       ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
       Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmaster_usergroupTableAdapter

        Try
            'batasan edit Requirement FG:
            '1. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(CDec(vPrimaryTransaksi))
                If CreatedBy <> vSalesStaff Then
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If


            'sampe sini balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Try
            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mGroupCode) = True Then
                    MsgBox("Sudah ada group dengan code " + Me.mGroupCode + ", ganti group code.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                'no po tidak boleh duplikat
                Dim NoLama As String = Me.GetGroupCode(Me.mPrimary)
                If Me.GroupCode <> NoLama Then
                    If Me.ApakahSudahAda(Me.mGroupCode) = True Then
                        MsgBox("Sudah ada group  dengan code  " + Me.mGroupCode + ", ganti group code.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            'return value
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
      ByVal vTrv As TreeView) As Boolean
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmaster_usergroupTableAdapter
        Dim TBL As New DseMasterUserGroup.tblmaster_usergroupDataTable
        Dim row As DseMasterUserGroup.tblmaster_usergroupRow
        Dim ObjItem As New ClsMasterUserGroupItem

        Try
            'fil basis data
            If DTA.FillBy_PrimaryTransaksi(TBL, CDec(vPrimaryTransaksi)) > 0 Then
                'salin data header
                row = TryCast(TBL.Rows(0), DseMasterUserGroup.tblmaster_usergroupRow)
                With row
                    mPrimary = .Primary
                    mGroupCode = .Group_Code
                    mDescription = .Desc

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimeStamp = .Timestamp
                End With
                'salin data item 
                ObjItem.IsiDGV(vTrv, vPrimaryTransaksi)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vTrv)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Sub ResetObject(ByVal vTrv As Windows.Forms.TreeView)
        Dim Nd As TreeNode
        Dim ObjItem As New ClsMasterUserGroupItem

        Try
            mPrimary = -1
            mGroupCode = ""
            Me.mDescription = ""

            mCreatedBy = ""
            mCreatedDate = Now
            mModifiedBy = ""
            mModifiedDate = Now
            mTimeStamp = Now

            For Each Nd In vTrv.Nodes
                ObjItem.KosongkanTreeView(Nd)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function SimpanBaru(ByVal vTrv As Windows.Forms.TreeView, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmaster_usergroupTableAdapter
        Dim ObjItem As New ClsMasterUserGroupItem

        Try
            If Me.ApakahDataBenar(True) = True Then
                If DTA.Insert(Me.GroupCode, Me.Description, _
                   vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mGroupCode).Value
                    Me.mTimeStamp = DTA.ScalarQuery_GetTimestamp(CDec(Me.mPrimary)).Value
                    Me.mCreatedDate = Me.mTimeStamp
                    Me.mModifiedDate = Me.mTimeStamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item baru
                    ObjItem.Simpan(vTrv, Me.mPrimary)

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

    Public Overridable Function SimpanPerubahan(ByVal vTrv As Windows.Forms.TreeView, ByVal vUserCode As String, ByVal VPrimaryGroup As Double) As Boolean
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmaster_usergroupTableAdapter
        Dim ObjItem As New ClsMasterUserGroupItem

        Dim DTA1 As New DseMasterUserGroupTableAdapters.tbluserTableAdapter
        Dim row As DseMasterUserGroup.tbluserRow
        Dim DS As New DseMasterUserGroup
        Dim row1 As DseMasterUserGroup.Tbl_JmlUserRow

        Try
            If Me.ApakahDataBenar(False) = True Then
                'hitung level otorisasi yang dibutuhkan
                If DTA.Update(Me.GroupCode, Me.Description, _
                  vUserCode, Now, Now, CDec(Me.Primary), Me.TimeStamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimeStamp = DTA.ScalarQuery_GetTimestamp(CDec(Me.mPrimary)).Value
                    Me.mModifiedDate = Me.mTimeStamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan(vTrv, Me.mPrimary)
                    'nilai balikan


                    ''buat Fungsi untuk mengUpdate Tbl Antara
                    ''Bersihkan
                    DS.Tbl_JmlUser.Clear()
                    '1 Cek dulu Primary Group ini di pake sama Berapa Kode User]
                    For Each row In DTA1.GetData(CDec(VPrimaryGroup))
                        row1 = DS.Tbl_JmlUser.NewTbl_JmlUserRow
                        row1.Primary = row.Primary_Group
                        row1.KodeUser = row.Kode_User
                        DS.Tbl_JmlUser.Rows.Add(row1)
                    Next
                    '2 Baru UpDate
                    For Each row1 In DS.Tbl_JmlUser
                        Me.SimpanPerubahanTBLAntara(row1.KodeUser, VPrimaryGroup)
                    Next

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
            DTA1.Dispose()
            DS.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahanTBLAntara(ByVal VKodeUser As String, ByVal VPrimaryGroup As Double) As Boolean
        Dim DTA As New DSUserTableAdapters.TblPerbandinganTableAdapter
        Dim DTA1 As New DSUserTableAdapters.tblantarauserhakakses1TableAdapter
        Dim row As DSUser.TblPerbandinganRow
        Dim DS As New DSUser
        Dim row1 As DSUser.Tbl_DataGroupRow
        Dim ObjUser As New ClsUser
        Try
            ''Bersihkan
            DS.Tbl_DataGroup.Clear()

            '' Kumpulkan Data
            For Each row In DTA.GetDataBy_PrimaryGroup(CDec(VPrimaryGroup))
                row1 = DS.Tbl_DataGroup.NewTbl_DataGroupRow
                row1.PrimaryGroup = row.Primary
                row1.KodeHakAkses = row.Kode_Hak_Akses.ToString
                row1.Check = row.Checked
                DS.Tbl_DataGroup.Rows.Add(row1)
            Next

            ''Hapus Tabel Antara
            ObjUser.HapusTBLAntara(VKodeUser)

            ''Simpan Table Antara
            For Each row1 In DS.Tbl_DataGroup
                DTA1.Insert(VKodeUser, row1.KodeHakAkses)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DS.Dispose()
            DTA.Dispose()
            DTA1.Dispose()
            ObjUser = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim dta As New DseMasterUserGroupTableAdapters.tblmaster_usergroupTableAdapter
        Try
            '1. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            Dim CreatedBy As String = dta.ScalarQuery_GetCreatedBy(CDec(vPrimaryTransaksi))
            If CreatedBy = vSalesStaff Then
                '2.Kalau Merujuk Ke Master User
                ''Cek dulu Primary Group ini di pake sama Berapa Kode User
                If Me.GetJmlPrimaryGroupByKodeUser(vPrimaryTransaksi) > 0 Then
                    Return False
                End If
                Return True
            Else
                If TampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function GetJmlPrimaryGroupByKodeUser(ByVal VPrimaryGroup As Double) As Double
        Dim DTA As New DseMasterUserGroupTableAdapters.tbluserTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlKodeUserByPrimaryGroup(CDec(VPrimaryGroup)).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlKodeUserByPrimaryGroup(CDec(VPrimaryGroup)).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Public Function GetCreatedBy(ByVal vPrimaryGroup As Double) As String
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmaster_usergroupTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(CDec(vPrimaryGroup))) = False Then
                Return DTA.ScalarQuery_GetCreatedBy(CDec(vPrimaryGroup))
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

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, _
    ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean) As Boolean
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmaster_usergroupTableAdapter

        Dim DTA1 As New DseMasterUserGroupTableAdapters.tbluserTableAdapter
        Dim row As DseMasterUserGroup.tbluserRow
        Dim DS As New DseMasterUserGroup
        Dim row1 As DseMasterUserGroup.Tbl_JmlUserRow
        Dim objUser As New ClsUser
        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, _
            UserCode) = True Then

                ''Hapus tabel antara terlebi dahulu
                ''Bersihkan
                DS.Tbl_JmlUser.Clear()
                '1 Cek dulu Primary Group ini di pake sama Berapa Kode User]
                For Each row In DTA1.GetData(CDec(vPrimaryTransaksi))
                    row1 = DS.Tbl_JmlUser.NewTbl_JmlUserRow
                    row1.Primary = row.Primary_Group
                    row1.KodeUser = row.Kode_User
                    DS.Tbl_JmlUser.Rows.Add(row1)
                Next

                '2 Baru Hapus
                For Each row1 In DS.Tbl_JmlUser
                    objUser.HapusTBLAntara(row1.KodeUser)
                Next

                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(CDec(vPrimaryTransaksi)) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DS.Dispose()
            objUser = Nothing
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmaster_usergroupTableAdapter

        Try
            If DTA.ScalarQuery_GetTimestamp(CDec(vPrimary)).HasValue = True Then
                Return DTA.ScalarQuery_GetTimestamp(CDec(vPrimary)).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoTrans As String) As Double
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmaster_usergroupTableAdapter
        Try
            Return DTA.ScalarQuery_GetPrimary(vNoTrans).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmaster_usergroup1TableAdapter
        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlLIKEGroupCode(Criteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterUserGroup.tblmaster_usergroup1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmaster_usergroup1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_GroupCodeLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL
            DGV.Refresh()

            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                ''DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                ''DGV.Columns(3).DefaultCellStyle.Format = "MMM yyyy"
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > 10000000 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Ticks / 10000000) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / 10000, 0, MidpointRounding.AwayFromZero)) + " miliseconds"
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

    Public Function GetGroupCode(ByVal vprimary As Double) As String
        Dim dta As New DseMasterUserGroupTableAdapters.tblmaster_usergroupTableAdapter
        Try
            If String.IsNullOrEmpty(dta.ScalarQuery_GetGroupCode(CDec(vprimary))) = False Then
                Return dta.ScalarQuery_GetGroupCode(CDec(vprimary))
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
End Class
