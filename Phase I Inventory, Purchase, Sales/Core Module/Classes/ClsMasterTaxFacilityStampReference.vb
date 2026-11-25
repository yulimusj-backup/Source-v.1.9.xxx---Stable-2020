Option Explicit On
Option Strict Off

Public Class ClsMasterTaxFacilityStampReference
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mType As New ClsMasterTaxProductReference.EnumType
    Dim mCode As String = ""
    Dim mDescription As String = ""
    Dim mRemarks As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Enum EnumType
        Goods
        Services
    End Enum

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property Type() As ClsMasterTaxProductReference.EnumType
        Get
            Return Me.mType
        End Get
        Set(ByVal value As ClsMasterTaxProductReference.EnumType)
            mType = value
        End Set
    End Property

    Public Property Code() As String
        Get
            Return Me.mCode
        End Get
        Set(ByVal value As String)
            mCode = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return Me.mDescription
        End Get
        Set(ByVal value As String)
            mDescription = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return Me.mRemarks
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

    Public Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterTaxFacilityStampReferenceTableAdapters.tblmaster_taxcapfasilitas_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mCode, Me.mDescription, Me.mType, Me.mRemarks, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mCode).Value
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(CDec(Me.mPrimary)).Value
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

    'Public Function GetJml(ByVal vCode As String, ByVal vTipe As ClsMasterTaxProductReference.EnumType) As Double
    '    Dim DTA As New DseMasterTaxFacilityStampReferenceTableAdapters.tblmaster_taxcapfasilitas_referenceTableAdapter

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If DTA.ScalarQuery_GetJml(vCode, vTipe).HasValue = True Then
    '            Return DTA.ScalarQuery_GetJml(vCode, vTipe).Value
    '        Else
    '            Return 0
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return 0
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria) As Double
        Dim DTA As New DseMasterTaxFacilityStampReferenceTableAdapters.tblcariTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                If DTA.ScalarQuery_GetJmlByCodeLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByCodeLIKE(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                If DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                If DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").Value
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

    Public Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterTaxFacilityStampReferenceTableAdapters.tblmaster_taxcapfasilitas_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Update(Me.mCode, Me.mDescription, Me.mType, Me.mRemarks, vUserCode, Now, Now, CDec(Me.mPrimary), Me.mTimestamp) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(CDec(Me.mPrimary)).Value
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

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterTaxFacilityStampReferenceTableAdapters.tblmaster_taxcapfasilitas_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary) = True Then
                DTA.Delete(CDec(vPrimary))
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

    Public Function ApakahSudahAda(ByVal vmCode As String) As Boolean
        Dim DTA As New DseMasterTaxFacilityStampReferenceTableAdapters.tblmaster_taxcapfasilitas_referenceTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlAllType(vmCode).HasValue Then
                If DTA.ScalarQuery_GetJmlAllType(vmCode).Value > 0 Then
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

    Public Function ApakahSudahAda(ByVal vmCode As String, ByVal vTipe As ClsMasterTaxProductReference.EnumType) As Boolean
        Dim DTA As New DseMasterTaxFacilityStampReferenceTableAdapters.tblmaster_taxcapfasilitas_referenceTableAdapter

        Try
            If DTA.ScalarQuery_GetJml(vmCode, vTipe).HasValue Then
                If DTA.ScalarQuery_GetJml(vmCode, vTipe).Value > 0 Then
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

    Public Function ApakahBolehHapus(ByVal vPrimary As Double) As Boolean
        Try
            'tidak boleh dihapus bila sudah ada yang merefer:
            '1. Master Product
            If Me.GetJmlTransactionByTaxFacilityStampReference(vPrimary) > 0 Then
                MsgBox("You cannot delete this data, because there are transaction using this data.", _
                                 MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterTaxFacilityStampReference.tblcariDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterTaxFacilityStampReferenceTableAdapters.tblcariTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(3).Visible = False
                DGV.Columns(1).HeaderText = "Code"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
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
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function CariDetail(ByVal Primary As Double) As Boolean
        Dim DTA As New DseMasterTaxFacilityStampReferenceTableAdapters.tblmaster_taxcapfasilitas_referenceTableAdapter
        Dim TBL As New DseMasterTaxFacilityStampReference.tblmaster_taxcapfasilitas_referenceDataTable
        Dim row As DseMasterTaxFacilityStampReference.tblmaster_taxcapfasilitas_referenceRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(Primary))

            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseMasterTaxFacilityStampReference.tblmaster_taxcapfasilitas_referenceRow)
                With row
                    Me.mPrimary = row.Primary
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mDescription = .Description
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mCode = .Code
                    Me.mType = .Type
                    Me.mRemarks = .Remarks
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

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterTaxFacilityStampReferenceTableAdapters.tblmaster_taxcapfasilitas_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCode(CDec(vPrimary))) = False Then
                Return DTA.ScalarQuery_GetCode(CDec(vPrimary))
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

    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterTaxFacilityStampReferenceTableAdapters.tblmaster_taxcapfasilitas_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDesc(CDec(vPrimary))) = False Then
                Return DTA.ScalarQuery_GetDesc(CDec(vPrimary))
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

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterTaxFacilityStampReferenceTableAdapters.tblmaster_taxcapfasilitas_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTimestamp(CDec(vPrimary)).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vmCode As String) As Double
        Dim DTA As New DseMasterTaxFacilityStampReferenceTableAdapters.tblmaster_taxcapfasilitas_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vmCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vmCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlTransactionByTaxFacilityStampReference(ByVal vPrimaryTaxFacilityStamp As Double) As Double
        Dim DTA As New DseMasterTaxFacilityStampReferenceTableAdapters.tblmaster_taxcapfasilitas_referenceTableAdapter
        Dim r As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            r = DTA.ScalarQuery_GetJmlCustomInvoiceByPrimaryFacilityStamp(vPrimaryTaxFacilityStamp)

            If r.HasValue = True Then
                If r.Value > 0 Then
                    Return r.Value
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

    Public Function SimpanItem_UsingScript(ByVal i As FrmMasterTaxFacilityStampReference) As Boolean
        Dim conStr As String = My.Settings.Core_DbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Try
            'open connection
            Conn.Open()
            If Me.ApakahSudahAdaData = False Then
                sql = Me.GetInsertString()
                script.Query = sql
                'execute script
                script.Execute()
                MsgBox("Insert data completed.", MsgBoxStyle.Information, "Information")
            End If

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString() As String
        Dim retval As String = ""
        Try
            retval = "Insert Into `tblmaster_taxcapfasilitas_reference` (`type`,`Code`,`Description`) Values" & _
            "('0','TD.01101','1 - untuk Kawasan Bebas')," & _
            "('0','TD.01102','2 - untuk Tempat Penimbunan Berikat')," & _
            "('0','TD.01103','3 - untuk Hibah dan Bantuan Luar Negeri')," & _
            "('0','TD.01104','4 - untuk Avtur')," & _
            "('0','TD.01105','5 - untuk Lainnya')," & _
            "('0','TD.01106','6 - untuk Kontraktor Perjanjian Karya Pengusahaan Pertambangan Batubara Generasi I')," & _
            "('0','TD.01107','7 - untuk Penyerahan bahan bakar minyak untuk Kapal Angkutan Laut Luar Negeri')," & _
            "('0','TD.01108','8 - untuk Penyerahan jasa kena pajak terkait alat angkutan tertentu')," & _
            "('0','TD.01109','9 - untuk Penyerahan BKP Tertentu di KEK')," & _
            "('0','TD.01110','10 - untuk BKP tertentu yang bersifat strategis berupa anode slime')," & _
            "('0','TD.01111','11 - untuk Penyerahan alat angkutan tertentu dan/atau Jasa Kena Pajak terkait alat angkutan tertentu')," & _
            "('0','TD.01112','12 - untuk Penyerahan kepada Kontraktor Kerja Sama Migas yang mengikuti ketentuan Peraturan Pemerintah Nomor 27 Tahun 2017')," & _
            "('0','TD.01113','13 - Penyerahan Rumah Tapak dan Satuan Rumah Susun Rumah Susun Ditanggung Pemerintah Tahun Anggaran 2021')," & _
            "('0','TD.01114','14 - Penyerahan Jasa Sewa Ruangan atau Bangunan Kepada Pedagang Eceran yang Ditanggung Pemerintah Tahun Anggaran 2021')," & _
            "('0','TD.01115','15 - Penyerahan Barang dan Jasa Dalam Rangka Penanganan Pandemi COVID-19 (PMK 239/PMK. 03/2020)')," & _
            "('0','TD.01116','16 - Insentif PMK-103/PMK.010/2021 berupa PPN atas Penyerahan Rumah Tapak dan Unit Hunian Rumah Susun yang Ditanggung Pemerintah Tahun Anggaran 2021')," & _
            "('0','TD.01117','17 - Kawasan Ekonomi Khusus PP nomor 40 Tahun 2021')," & _
            "('0','TD.01118','18 - Kawasan Bebas PP nomor 41 Tahun 2021')," & _
            "('0','TD.01119','19 - Penyerahan Rumah Tapak dan Unit Hunian Rumah Susun yang Ditanggung Pemerintah Tahun Anggaran 2022')," & _
            "('0','TD.01120','20 - PPN Ditanggung Pemerintah dalam rangka Penanganan Pandemi Corona Virus')," & _
            "('0','TD.01121','21 - Penyerahan kepada Kontraktor Kerja Sama Migas yang mengikuti ketentuan Peraturan Pemerintah Nomor 53 Tahun 2017')," & _
            "('0','TD.01122','22 - BKP strategis tertentu dalam bentuk anode slime dan emas butiran')," & _
            "('0','TD.01123','23 - untuk penyerahan kertas koran dan/atau majalah')," & _
            "('0','TD.01124','24 - PPN tidak dipungut oleh Pemerintah lainnya')," & _
            "('0','TD.01125','25 - BKP dan JKP tertentu')," & _
            "('0','TD.01126','26 - Penyerahan BKP dan JKP di Ibu Kota Negara baru')," & _
            "('0','TD.01127','27 - Penyerahan kendaraan listrik berbasis baterai')," & _
            "('1','TD.01101','1 - untuk BKP dan JKP Tertentu')," & _
            "('1','TD.01102','2 - untuk BKP Tertentu yang Bersifat Strategis')," & _
            "('1','TD.01103','3 - untuk Jasa Kebandarudaraan')," & _
            "('1','TD.01104','4 - untuk Lainnya')," & _
            "('1','TD.01105','5 - untuk BKP Tertentu yang Bersifat Strategis sesuai PP Nomor 81 Tahun 2015')," & _
            "('1','TD.01106','6 - untuk Penyerahan Jasa Kepelabuhan Tertentu untuk kegiatan angkutan laut Luar Negeri')," & _
            "('1','TD.01107','7 - untuk Penyerahan Air Bersih')," & _
            "('1','TD.01108','8 - Penyerahan BKP tertentu yang bersifat strategis berdasarkan PP 48 Tahun 2020')," & _
            "('1','TD.01109','9 - Penyerahan kepada Perwakilan Negara Asing dan Badan Internasional serta Pejabatnya')," & _
            "('1','TD.01110','10 - BKP dan JKP tertentu');"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function ApakahSudahAdaData() As Boolean
        Dim d As New DseMasterTaxFacilityStampReferenceTableAdapters.tblmaster_taxcapfasilitas_referenceTableAdapter
        Dim l As New Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            l = d.ScalarQuery_GetJmlTotal()
            If l.HasValue = True Then
                If l.Value > 0 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
            d.Dispose()
        End Try
    End Function


End Class
