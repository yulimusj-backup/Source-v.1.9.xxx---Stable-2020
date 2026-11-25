Option Explicit On
Option Strict Off

Public Class ClsMasterTaxAddInfoReference
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
        Dim DTA As New DseMasterTaxAddInfoReferenceTableAdapters.tblmaster_taxaddinfo_referenceTableAdapter

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
    '    Dim DTA As New DseMasterTaxAddInfoReferenceTableAdapters.tblmaster_taxaddinfo_referenceTableAdapter

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
        Dim DTA As New DseMasterTaxAddInfoReferenceTableAdapters.tblcariTableAdapter

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
        Dim DTA As New DseMasterTaxAddInfoReferenceTableAdapters.tblmaster_taxaddinfo_referenceTableAdapter

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
        Dim DTA As New DseMasterTaxAddInfoReferenceTableAdapters.tblmaster_taxaddinfo_referenceTableAdapter

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
        Dim DTA As New DseMasterTaxAddInfoReferenceTableAdapters.tblmaster_taxaddinfo_referenceTableAdapter

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
        Dim DTA As New DseMasterTaxAddInfoReferenceTableAdapters.tblmaster_taxaddinfo_referenceTableAdapter

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
            If Me.GetJmlTransactionByTaxAddInfoReference(vPrimary) > 0 Then
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
    ByVal DGV As DataGridView, ByVal TBL As DseMasterTaxAddInfoReference.tblcariDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterTaxAddInfoReferenceTableAdapters.tblcariTableAdapter
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
        Dim DTA As New DseMasterTaxAddInfoReferenceTableAdapters.tblmaster_taxaddinfo_referenceTableAdapter
        Dim TBL As New DseMasterTaxAddInfoReference.tblmaster_taxaddinfo_referenceDataTable
        Dim row As DseMasterTaxAddInfoReference.tblmaster_taxaddinfo_referenceRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(Primary))

            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseMasterTaxAddInfoReference.tblmaster_taxaddinfo_referenceRow)
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
        Dim DTA As New DseMasterTaxAddInfoReferenceTableAdapters.tblmaster_taxaddinfo_referenceTableAdapter

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
        Dim DTA As New DseMasterTaxAddInfoReferenceTableAdapters.tblmaster_taxaddinfo_referenceTableAdapter

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
        Dim DTA As New DseMasterTaxAddInfoReferenceTableAdapters.tblmaster_taxaddinfo_referenceTableAdapter

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
        Dim DTA As New DseMasterTaxAddInfoReferenceTableAdapters.tblmaster_taxaddinfo_referenceTableAdapter

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

    Public Function GetJmlTransactionByTaxAddInfoReference(ByVal vPrimaryTaxProduct As Double) As Double
        Dim DTA As New DseMasterTaxAddInfoReferenceTableAdapters.tblmaster_taxaddinfo_referenceTableAdapter
        Dim r As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            r = DTA.ScalarQuery_GetJmlCustomInvoiceByPrimaryAddInfo(vPrimaryTaxProduct)

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

    Public Function SimpanItem_UsingScript(ByVal i As FrmMasterTaxAddInfoReference) As Boolean
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
            retval = "Insert Into `tblmaster_taxaddinfo_reference` (`type`,`Code`,`Description`) Values" & _
             "('0','TD.00501','1 - Pajak Pertambahan Nilai Tidak Dipungut berdasarkan PP Nomor 10 Tahun 2012')," & _
             "('0','TD.00502','2 - Pajak Pertambahan Nilai atau Pajak Pertambahan Nilai dan Pajak Penjualan atas Barang Mewah tidak dipungut')," & _
             "('0','TD.00503','3 - Pajak Pertambahan Nilai dan Pajak Penjualan atas Barang Mewah Tidak Dipungut')," & _
             "('0','TD.00504','4 - Pajak Pertambahan Nilai Tidak Dipungut Sesuai PP Nomor 71 Tahun 2012')," & _
             "('0','TD.00505','5 - (Tidak ada Cap)')," & _
             "('0','TD.00506','6 - PPN dan/atau PPnBM tidak dipungut berdasarkan PMK No. 194/PMK.03/2012')," & _
             "('0','TD.00507','7 - PPN Tidak Dipungut Berdasarkan PP Nomor 15 Tahun 2015')," & _
             "('0','TD.00508','8 - PPN Tidak Dipungut Berdasarkan PP Nomor 69 Tahun 2015')," & _
             "('0','TD.00509','9 - PPN Tidak Dipungut Berdasarkan PP Nomor 96 Tahun 2015')," & _
             "('0','TD.00510','10 - PPN Tidak Dipungut Berdasarkan PP Nomor 106 Tahun 2015')," & _
             "('0','TD.00511','11 - PPN Tidak Dipungut Sesuai PP Nomor 50 Tahun 2019')," & _
             "('0','TD.00512','12 - PPN atau PPN dan PPnBM Tidak Dipungut Sesuai Dengan PP Nomor 27 Tahun 2017')," & _
             "('0','TD.00513','13 - PPN ditanggung PEMERINTAH EX PMK 21/PMK.010/21')," & _
             "('0','TD.00514','14 - PPN DITANGGUNG PEMERINTAH EKS PMK 102/PMK.010/2021')," & _
             "('0','TD.00515','15 - PPN DITANGGUNG PEMERINTAH EKS PMK 239/PMK.03/2020')," & _
             "('0','TD.00516','16 - Insentif PPN DITANGGUNG PEMERINTAH EKSEKUSI PMK NOMOR 103/PMK.010/2021')," & _
             "('0','TD.00517','17 - PAJAK PERTAMBAHAN NILAI TIDAK DIPUNGUT BERDASARKAN PP NOMOR 40 TAHUN 2021')," & _
             "('0','TD.00518','18 - PAJAK PERTAMBAHAN NILAI TIDAK DIPUNGUT BERDASARKAN PP NOMOR 41 TAHUN 2021')," & _
             "('0','TD.00519','19 - PPN DITANGGUNG PEMERINTAH EKS PMK 6/PMK.010/2022')," & _
             "('0','TD.00520','20 - PPN DITANGGUNG PEMERINTAH EKSEKUSI PMK NOMOR 226/PMK.03/2021')," & _
             "('0','TD.00521','21 - PPN ATAU PPN DAN PPnBM TIDAK DIPUNGUT SESUAI DENGAN PP NOMOR 53 TAHUN 2017')," & _
             "('0','TD.00522','22 - PPN tidak dipungut berdasarkan PP Nomor 70 Tahun 2021')," & _
             "('0','TD.00523','23 - PPN ditanggung Pemerintah Ex PMK-125/PMK.01/2020')," & _
             "('0','TD.00524','24 - (Tidak ada Cap)')," & _
             "('0','TD.00525','25 - PPN tidak dipungut berdasarkan PP Nomor 49 Tahun 2022')," & _
             "('0','TD.00526','26 - PPN tidak dipungut berdasarkan PP Nomor 12 Tahun 2023')," & _
             "('0','TD.00527','27 - PPN ditanggung Pemerintah berdasarkan PMK Nomor 38 Tahun 2023')," & _
             "('1','TD.00501','1 - PPN Dibebaskan Sesuai PP Nomor 146 Tahun 2000 Sebagaimana Telah Diubah Dengan PP Nomor 38 Tahun 2003')," & _
             "('1','TD.00502','2 - PPN Dibebaskan Sesuai PP Nomor 12 Tahun 2001 Sebagaimana Telah Beberapa Kali Diubah Terakhir Dengan PP Nomor 31 Tahun 2007')," & _
             "('1','TD.00503','3 - PPN dibebaskan berdasarkan Peraturan Pemerintah Nomor 28 Tahun 2009')," & _
             "('1','TD.00504','4 - (Tidak ada cap)')," & _
             "('1','TD.00505','5 - PPN Dibebaskan Sesuai Dengan PP Nomor 81 Tahun 2015')," & _
             "('1','TD.00506','6 - PPN Dibebaskan Berdasarkan PP Nomor 74 Tahun 2015')," & _
             "('1','TD.00507','7 - (tanpa cap)')," & _
             "('1','TD.00508','8 - PPN DIBEBASKAN SESUAI PP NOMOR 81 TAHUN 2015 SEBAGAIMANA TELAH DIUBAH DENGAN PP 48 TAHUN 2020')," & _
             "('1','TD.00509','9 - PPN DIBEBASKAN BERDASARKAN PP NOMOR 47 TAHUN 2020')," & _
             "('1','TD.00510','10  -PPN Dibebaskan berdasarkan PP Nomor 49 Tahun 2022');"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function ApakahSudahAdaData() As Boolean
        Dim d As New DseMasterTaxAddInfoReferenceTableAdapters.tblmaster_taxaddinfo_referenceTableAdapter
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
