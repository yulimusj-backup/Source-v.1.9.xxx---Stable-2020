Option Explicit On
Option Strict Off

Public Class ClsMasterProses
    Private mPrimary As Double = -1
    Private mCode As String = ""
    Private mDesc As String = ""
    Private mRemarks As String = ""
    Private mAktif As Integer = Enum_Aktif.Aktif

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
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian

    Public Enum Enum_Aktif As Integer
        NonAktif
        Aktif
    End Enum

    Public Property Aktif() As Enum_Aktif
        Get
            Return mAktif
        End Get
        Set(ByVal value As Enum_Aktif)
            mAktif = value
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

    Public Function GetConstring() As String
        Try
            Return Core_Module.GetConstring()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=fitri-pc;user id=root;Password=123;persist security info=True;database=payroll_db;port=3307;default command timeout=19000000"
        End Try
    End Function

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_prosesTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimary(vCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vCode).Value
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
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_prosesTableAdapter
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

    Public Function GetAktif(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_prosesTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetAktif(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetAktif(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_prosesTableAdapter
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

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_prosesTableAdapter
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

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_prosesTableAdapter
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

    ''' <summary>
    ''' data ini digunakan untuk mengecek kebenaran data yang diberikan
    ''' </summary>
    ''' <param name="vBaru"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vShowMsg As Boolean) As Boolean
        Try
            'tidak boleh ada kode yang sama
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mCode) = True Then
                    If vShowMsg = True Then
                        MsgBox("Code duplication, please change code.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.GetCode(Me.mPrimary)
                If CodeLama <> Me.mCode Then
                    If Me.ApakahSudahAda(Me.mCode) = True Then
                        If vShowMsg = True Then
                            MsgBox("Code duplication, please change code.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            'yang harus diisikan:
            '1. kode
            If Me.mCode = "" Then
                If vShowMsg = True Then
                    MsgBox("Please specify code.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '2. nama
            If Me.mDesc = "" Then
                If vShowMsg = True Then
                    MsgBox("Please specify description.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vStrictEdit As Boolean, _
    ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_prosesTableAdapter
        Dim pCreatedBy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            '1. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictEdit = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If

                'If pCreatedBy <> vUserCode Then
                'If vShowMsg = True Then
                'MsgBox("User yang bisa mengedit transaksi ini adalah: " + pCreatedBy, MsgBoxStyle.Information, "Information")
                'End If
                'Return False
                'End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vStrictDelete As Boolean, _
    ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_prosesTableAdapter
        Dim pCreatedBy As String = ""
        Dim ObjKaryawan As New Core_Module.ClsMasterKaryawan
        Dim Dummy As Nullable(Of Long)
        Dim ObjCustom As New ClsIdentifier

        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            '1.tidak boleh dihapus bila sudah ada yang merefer:

            '2. tidak boleh hapus jika melanggar aturan strict edit
            'If vStrictDelete = True Then
            '    If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
            '        pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
            '    Else
            '        pCreatedBy = ""
            '    End If

            '    If pCreatedBy <> vUserCode Then
            '        If vShowMsg = True Then
            '            MsgBox("User yang bisa menghapus transaksi ini adalah: " + pCreatedBy, MsgBoxStyle.Information, "Information")
            '        End If
            '        Return False
            '    End If
            'End If

            'tidak boleh hapus jika ada product yang merujuk pada proses ini
            Dummy = DTA.ScalarQuery_GetJmlByProduct(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
                    MsgBox("There is product referring to this data, you can't delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'tidak boleh hapus jika bundling yang merujuk pada proses ini
            Dummy = DTA.ScalarQuery_GetJmlByStockBundling(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
                    MsgBox("There is stock bundling transaction referring to this data, you can't delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If ObjCustom.GetIdentifier = "smess_lmb" Then
                If Me.ApakahBolehHapus_CustomLMB(vPrimary) = False Then
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjKaryawan = Nothing
            ObjCustom = Nothing
        End Try
    End Function

    Public Function ApakahBolehHapus_CustomLMB(ByVal vPrimaryProses As Double) As Boolean
        Dim DTA As New DseMasterProsesTableAdapters.tbltrans_custom_stockbundling_setting_warehouse_itemTableAdapter
        Dim TBL As New DseMasterProses.tbltrans_custom_stockbundling_setting_warehouse_itemDataTable
        Dim row As DseMasterProses.tbltrans_custom_stockbundling_setting_warehouse_itemRow 

        Try
            'tidak boleh hapus jika sudah di rujuk pada master setting bundling
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryProses)
            If TBL.Rows.Count > 0 Then
                MsgBox("This data is referenced by bundling warehouse setting, " + _
                vbCrLf + "you can't delete this data.", MsgBoxStyle.Information, "Information")

                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    ''
    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterProses.tblmaster_proses_cariDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal vStatusAktif As Boolean, _
    ByVal vSearchAll As Boolean) As Boolean
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_proses_cariTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Dim StatusAktif_Int As Integer
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If vStatusAktif = True Then
                StatusAktif_Int = 1
            Else
                StatusAktif_Int = 0
            End If

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If vSearchAll = False Then
                If SearchCriteria = EnumSearchCriteria.ByCode Then
                    QueryResult = DTA.FillBy_Code(TBL, Criteria + "%", StatusAktif_Int)
                ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                    QueryResult = DTA.FillBy_Desc(TBL, Criteria + "%", StatusAktif_Int)
                End If
            Else
                If SearchCriteria = EnumSearchCriteria.ByCode Then
                    QueryResult = DTA.FillBy_CodeLIKEALL(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                    QueryResult = DTA.FillBy_DescLIKEALL(TBL, Criteria + "%")
                End If
            End If


            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                'DGV.Columns(1).HeaderText = "Kode Proses"
                'DGV.Columns(2).HeaderText = "Deskripsi"
                'DGV.Columns(3).HeaderText = "Keterangan"
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

    Public Function CariDetail(ByVal Primary As Double) As Boolean
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_prosesTableAdapter
        Dim TBL As New DseMasterProses.tblmaster_prosesDataTable
        Dim row As DseMasterProses.tblmaster_prosesRow
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.FillBy_Primary(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mCode = .Code
                    Me.mDesc = .Desc
                    Me.mRemarks = .Remarks
                    Me.mAktif = .Aktif

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

    Public Function GetJml(ByVal vKode As String) As Double
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_prosesTableAdapter

        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByCode(vKode).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByCode(vKode).Value
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As EnumSearchCriteria, ByVal vStatusAktif As Boolean, _
    ByVal vSearchAll As Boolean) As Double
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_proses_cariTableAdapter
        Dim StatusAktif_Int As Integer

        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If vStatusAktif = True Then
                StatusAktif_Int = 1
            Else
                StatusAktif_Int = 0
            End If

            If vSearchAll = False Then
                If vSearchBy = EnumSearchCriteria.ByCode Then
                    Return DTA.ScalarQuery_GetJmlCodeLIKE(vKriteria + "%", StatusAktif_Int).Value
                ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                    Return DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%", StatusAktif_Int).Value
                End If
            Else
                If vSearchBy = EnumSearchCriteria.ByCode Then
                    Return DTA.ScalarQuery_GetJmlByCodeLIKEALL(vKriteria + "%").Value
                ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                    Return DTA.ScalarQuery_GetJmlByDescLIKEALL(vKriteria + "%").Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub New()
        Try
            Me.DTAInOUT.Connection.ConnectionString = Me.GetConstring
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterProses, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode)
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
        End Try
    End Function

    Public Function SimpanBaru1(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_prosesTableAdapter
        Dim Status_Int As Integer

        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If Me.ApakahDataBenar(True, True) = True Then
                If Me.mAktif = Enum_Aktif.Aktif Then
                    Status_Int = 1
                Else
                    Status_Int = 0
                End If
                If DTA.Insert(Me.mCode, Me.mDesc, Me.mRemarks, Me.mAktif, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mCode)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
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
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterProses, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode)
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
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_prosesTableAdapter
        Dim StatusAktif_Int As Integer

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(False, True) = True Then
                If Me.mAktif = Enum_Aktif.Aktif Then
                    StatusAktif_Int = 1
                Else
                    StatusAktif_Int = 0
                End If

                DTA.Connection.ConnectionString = Me.GetConstring
                If DTA.Update(Me.mCode, Me.mDesc, Me.mRemarks, Me.mAktif, vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
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
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, vPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterProses, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimary, vUserCode, vStrictDelete, vShowMsg)
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
        End Try
    End Function

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_prosesTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            If Me.ApakahBolehHapus(vPrimary, vUserCode, vStrictDelete, vShowMsg) = True Then
                DTA.Delete(vPrimary)
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

    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_prosesTableAdapter
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

    Public Sub FillTabelProses(ByVal vTBL As DseMasterProses.tblmaster_proses1DataTable)
        Dim DTA As New DseMasterProsesTableAdapters.tblmaster_proses1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vTBL.Rows.Clear()

            DTA.Fill(vTBL)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
