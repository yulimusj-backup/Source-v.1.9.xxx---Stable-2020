
Option Explicit On
Option Strict Off
Public Class ClsMasterRefParty
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mCode As String = ""
    Private mDesc As String = ""
    Private mEmail As String = ""
    Private mAktif As EnumAktif
    Private mRemarks As String = ""
    Private mCreatedBy As String = ""
    Private mCreatedDate As Date = Now
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date = Now
    Private mTimestamp As Date = Now


    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

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

    Public Property Email() As String
        Get
            Return mEmail
        End Get
        Set(ByVal value As String)
            mEmail = value
        End Set
    End Property

    Public Property Aktif() As EnumAktif
        Get
            Return mAktif
        End Get
        Set(ByVal value As EnumAktif)
            mAktif = value
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
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
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

    Public Enum EnumAktif
        No
        Yes
    End Enum

    '
    'Scalar Get 
    '
    Public Function GetPrimaryByAktif(ByVal vNoTransaksi As String, ByVal vAktif As Integer) As Double
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter
        Dim Primary As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryByAktifasi(vAktif, vNoTransaksi).HasValue AndAlso DTA.ScalarQuery_GetPrimary(vNoTransaksi).Value > -1 Then
                Primary = DTA.ScalarQuery_GetPrimaryByAktifasi(vAktif, vNoTransaksi).Value
            End If
            Return Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vRefPartyCode As String) As Double
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter
        Dim Primary As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimary(vRefPartyCode).HasValue Then
                If DTA.ScalarQuery_GetPrimary(vRefPartyCode).Value > -1 Then
                    Primary = DTA.ScalarQuery_GetPrimary(vRefPartyCode).Value
                End If
            End If
            Return Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetEmail(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter
        Dim RetVal As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetEmail(vPrimary)
            If String.IsNullOrEmpty(RetVal) = True Then
                Return ""
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter
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

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetCode(vPrimary)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDesc(ByVal vPrimaryRefParty As Double) As String
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetDesc(vPrimaryRefParty)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByCode2(ByVal vCode As String) As Double
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetJmlByCode2(vCode)) = False Then
                Return DTA.ScalarQuery_GetJmlByCode2(vCode)
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

    Public Function GetJmlByAktif(ByVal vAktif As Integer, ByVal vCode As String) As Double
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetJmlByAktif(vAktif, vCode)) = False Then
                Return DTA.ScalarQuery_GetJmlByAktif(vAktif, vCode)
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


    Public Function GetJmlPrimaryRefPartyInPR(ByVal vPrimaryRefParty As Double) As Double
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter
        Dim Retval As Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJmlPrimaryRefPartyInPR(vPrimaryRefParty)
            If Retval.HasValue = True Then
                If Retval.Value > 0 Then
                    Return Retval
                Else
                    Return 0
                End If
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    '
    'Function Simpan Baru & Simpan Perubahan & Hapus
    '
    Public Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Insert(Me.mCode, Me.mDesc, Me.mRemarks, Me.mAktif, Me.mEmail, vUserCode, Now, _
                          vUserCode, Now, Now) > 0 Then

                Me.mPrimary = Me.GetPrimary(Me.Code)
                Me.mTimestamp = Me.GetTimestamp(CDec(Me.mPrimary))
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

    Public Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Update(Me.mCode, Me.mDesc, Me.mRemarks, Me.mAktif, Me.mEmail, _
                          vUserCode, Now, Now, CDec(Me.mPrimary)) > 0 Then
                Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
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

    Public Function Hapus(ByVal vPrimary As Double, ByVal vUser As String) As Boolean
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(CDec(vPrimary))
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal Primary As Double) As Boolean
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter
        Dim TBL As New DseMasterRefParty.tblmaster_refpartyDataTable
        Dim row As DseMasterRefParty.tblmaster_refpartyRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_primary(TBL, CDec(Primary))
            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseMasterRefParty.tblmaster_refpartyRow)
                With row
                    Me.mPrimary = row.primary
                    Me.mCode = row.code
                    Me.mDesc = row.desc
                    Me.mEmail = row.email
                    Me.mAktif = row.aktif
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
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

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
           ByVal DGV As DataGridView, ByVal TBL As DseMasterRefParty.tblmaster_refparty1DataTable, _
           ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, Optional ByVal vAktif As Boolean = False) As Boolean
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refparty1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If vAktif = False Then
                If SearchCriteria = EnumSearchCriteria.ByCode Then
                    QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                    QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByEmail Then
                    QueryResult = DTA.FillBy_EmailLIKE(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                    QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
                End If
            Else
                If SearchCriteria = EnumSearchCriteria.ByCode Then
                    QueryResult = DTA.FillBy_CodeLIKEAktif(TBL, Criteria + "%", 1)
                ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                    QueryResult = DTA.FillBy_DescLIKEAktif(TBL, Criteria + "%", 1)
                ElseIf SearchCriteria = EnumSearchCriteria.ByEmail Then
                    QueryResult = DTA.FillBy_EmailLIKEAktif(TBL, Criteria + "%", 1)
                ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                    QueryResult = DTA.FillBy_RemarksLIKEAktif(TBL, Criteria + "%", 1)
                End If
            End If

            'update DGV
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Code"
                DGV.Columns(2).HeaderText = "Description"
                DGV.Columns(3).HeaderText = "Email"
                DGV.Columns(4).HeaderText = "Remarks"

                DGV.Columns(1).MinimumWidth = 80
                DGV.Columns(2).MinimumWidth = 100
                DGV.Columns(3).MinimumWidth = 90
                DGV.Columns(4).MinimumWidth = 80
            End If
            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + _
                    " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria) As Double
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                RetVal = DTA.ScalarQuery_GetJmlByCodeLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                RetVal = DTA.ScalarQuery_GetJmlByDesc(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByEmail Then
                RetVal = DTA.ScalarQuery_GetJmlByEmail(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                RetVal = DTA.ScalarQuery_GetJmlByRemarks(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByCode(vCode).HasValue Then
                If DTA.ScalarQuery_GetJmlByCode(vCode).Value > 0 Then
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

    Public Function ApakahSudahAda(ByVal vCode As String, ByVal vAktif As ClsMasterRefParty.EnumAktif) As Boolean
        Dim DTA As New DseMasterRefPartyTableAdapters.tblmaster_refpartyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByAktif(vCode, vAktif).HasValue Then
                If DTA.ScalarQuery_GetJmlByAktif(vCode, vAktif).Value > 0 Then
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

    Public Function ApakahSudahAda_ByPrimary(ByVal vPrimaryRefParty As Double) As Boolean
        Try
            If Me.GetJmlPrimaryRefPartyInPR(vPrimaryRefParty) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vCode As String, ByVal vSalesStaff As String, Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Try
            If Me.GetJmlByCode2(vCode) > 0 Then
                If vTampilkanPesan = True Then
                    MsgBox("The following ref. party data is reference by one or several purchase request, you can’t delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    '''master gg pake antrian


    ' fungsi2 untuk Antrian

    'Public Overridable Function SimpanBaru1(ByVal vUserCode As String) As Boolean
    '    Dim CheckInterval As Double = 0
    '    Dim TryCounter As Integer = 0
    '    Dim Start As Double = 0
    '    Dim retval As Boolean = True

    '    Try
    '        DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
    '        'ambil no antrian
    '        Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
    '        Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterRefParty_SimpanBaru, vUserCode)

    '        'coba simpan data baru
    '        Me.Clock_Antri.Reset()
    '        Me.Clock_Antri.Start()
    '        While Start = 0
    '            'hanya mencoba ke database setelah interval waktu tertentu saja
    '            If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
    '                If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
    '                    'clock antri matikan
    '                    Me.Clock_Antri.Stop()
    '                    'sudah saatnya melakukan proses simpan baru
    '                    retval = Me.SimpanBaru1(vUserCode)
    '                    'hapus data antri untuk transaksi ini
    '                    If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
    '                        MsgBox("Logic error(1), contact your provider when you see this message", _
    '                        MsgBoxStyle.Information, "Information")
    '                    End If
    '                    'nilai balikan
    '                    Return retval
    '                Else
    '                    'kontrol interval pengecekan
    '                    TryCounter += 1
    '                    CheckInterval = ((TryCounter \ 5) + 1) * 500
    '                    Me.Clock_Antri.Reset()
    '                    Me.Clock_Antri.Start()
    '                End If
    '            End If
    '        End While
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    'Public Overridable Function SimpanPerubahan1(ByVal vUserCode As String) As Boolean
    '    Dim CheckInterval As Double = 0
    '    Dim TryCounter As Integer = 0
    '    Dim Start As Double = 0
    '    Dim retval As Boolean = True

    '    Try
    '        DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
    '        'ambil no antrian
    '        Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
    '        Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterRefParty_SimpanPerubahan, vUserCode)

    '        'coba simpan data baru
    '        Me.Clock_Antri.Reset()
    '        Me.Clock_Antri.Start()
    '        While Start = 0
    '            'hanya mencoba ke database setelah interval waktu tertentu saja
    '            If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
    '                If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
    '                    'clock antri matikan
    '                    Me.Clock_Antri.Stop()
    '                    'sudah saatnya melakukan proses simpan baru
    '                    retval = Me.SimpanPerubahan1(vUserCode)
    '                    'hapus data antri untuk transaksi ini
    '                    If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
    '                        MsgBox("Logic error(1), contact your provider when you see this message", _
    '                        MsgBoxStyle.Information, "Information")
    '                    End If
    '                    'nilai balikan
    '                    Return retval
    '                Else
    '                    'kontrol interval pengecekan
    '                    TryCounter += 1
    '                    CheckInterval = ((TryCounter \ 5) + 1) * 500
    '                    Me.Clock_Antri.Reset()
    '                    Me.Clock_Antri.Start()
    '                End If
    '            End If
    '        End While
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    'Public Function Hapus1(ByVal vPrimaryTransaksi As Double, ByVal vUser As String) As Boolean
    '    Dim CheckInterval As Double = 0
    '    Dim TryCounter As Integer = 0
    '    Dim Start As Double = 0
    '    Dim retval As Boolean = True

    '    Try
    '        DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
    '        'ambil no antrian
    '        Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
    '        Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterRefParty_Hapus, vUser)
    '        'coba simpan data baru
    '        Me.Clock_Antri.Reset()
    '        Me.Clock_Antri.Start()
    '        While Start = 0
    '            'hanya mencoba ke database setelah interval waktu tertentu saja
    '            If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
    '                If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
    '                    'clock antri matikan
    '                    Me.Clock_Antri.Stop()
    '                    'sudah saatnya melakukan proses simpan baru
    '                    retval = Me.Hapus1(vPrimaryTransaksi, vUser)
    '                    'hapus data antri untuk transaksi ini
    '                    If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
    '                        MsgBox("Logic error(1), contact your provider when you see this message", _
    '                        MsgBoxStyle.Information, "Information")
    '                    End If
    '                    'nilai balikan
    '                    Return retval
    '                Else
    '                    'kontrol interval pengecekan
    '                    TryCounter += 1
    '                    CheckInterval = ((TryCounter \ 5) + 1) * 500
    '                    Me.Clock_Antri.Reset()
    '                    Me.Clock_Antri.Start()
    '                End If
    '            End If
    '        End While
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    'Public Sub New()
    '    Me.DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
    'End Sub
End Class