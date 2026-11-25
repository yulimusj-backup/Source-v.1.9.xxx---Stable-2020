Option Explicit On
Option Strict Off
Imports Microsoft.Win32
Imports System.Net
Public Class ClsTransaksiTagihan
    Dim ObjCD As New Core_Module.ClsCekData
    Dim objEmailSettings As New Core_Module.ClsEmailSettings

    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering

    Private mPrimary As Double = -1
    Private mNoTransaksi As String = ""
    Private mTanggalTransaksi As Date = Now
    Private mPrimaryCust As Double = -1
    Private mIndentifier As String = ""
    Private mSubject As String = ""
    Private mAttn As String = ""
    Private mSignedBy As String = ""
    Private mSignedPosition As String = ""
    Private mRemarks As String = ""

    Private mCreatedBy As String = ""
    Private mCreatedDate As Date = Now
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date = Now
    Private mTimestamp As Date = Now

    Public Property SignedPosition() As String
        Get
            Return mSignedPosition
        End Get
        Set(ByVal value As String)
            mSignedPosition = value
        End Set
    End Property

    Public Property SignedBy() As String
        Get
            Return mSignedBy
        End Get
        Set(ByVal value As String)
            mSignedBy = value
        End Set
    End Property

    Public Property Subject() As String
        Get
            Return mSubject
        End Get
        Set(ByVal value As String)
            mSubject = value
        End Set
    End Property

    Public Property Attn() As String
        Get
            Return mAttn
        End Get
        Set(ByVal value As String)
            mAttn = value
        End Set
    End Property

    Public Property PrimaryCust() As Double
        Get
            Return mPrimaryCust
        End Get
        Set(ByVal value As Double)
            mPrimaryCust = value
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

    Public Property Notransaksi() As String
        Get
            Return mNoTransaksi
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property

    Public Property TanggalTransaksi() As Date
        Get
            Return mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
        End Set
    End Property

    Public Property Identifier() As String
        Get
            Return mIndentifier
        End Get
        Set(ByVal value As String)
            mIndentifier = value
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

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihanTableAdapter
        Try
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
        Dim DTA As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihanTableAdapter
        Try
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary)
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
        Dim DTa As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihanTableAdapter
        Try
            If DTa.ScalarQuery_GetPrimary(vNoTrans).HasValue Then
                Return DTa.ScalarQuery_GetPrimary(vNoTrans)
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTa.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihanTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoTransaksiByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetNoTransaksiByPrimary(vPrimary)
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

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
        ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
        Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihanTableAdapter
        Dim CreatedBy As String = ""
        Try
            '1. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            If StrictEditting = True Then
                CreatedBy = Me.GetCreatedBy(vPrimaryTransaksi)
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

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
          ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
          Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihanTableAdapter
        Dim ObjItem As New ClsTransaksiTagihanItem
        Try

            '1. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            Dim CreatedBy As String = Me.GetCreatedBy(vPrimaryTransaksi)
            If CreatedBy = vSalesStaff Then
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
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function


    Public Function ApakahSudahAda(ByVal vNotrans As String) As Boolean
        Dim DTA As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihanTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByNoTransaksi(vNotrans).HasValue Then
                If DTA.ScalarQuery_GetJmlByNoTransaksi(vNotrans).Value > 0 Then
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

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
        ByVal vDGVItem As DataGridView, ByVal frm As FrmTransaksiTagihan) As Boolean
        Dim DTA As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihanTableAdapter
        Dim TBL As New DseTransaksiTagihan.tbltrans_surattagihanDataTable
        Dim row As DseTransaksiTagihan.tbltrans_surattagihanRow
        Dim ObjItem As New ClsTransaksiTagihanItem

        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = vPrimaryTransaksi
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mIndentifier = .Identifier_Alamat_Cust
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    Me.mSubject = .Subject
                    Me.mSignedBy = .Signed_By
                    Me.mSignedPosition = .Signed_By_Position
                    Me.mAttn = .Attn
                    Me.mRemarks = .Remarks
                    Me.mPrimaryCust = .Primary_Customer

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'salin data item ada 2 item 
                'CodeNo
                ObjItem.IsiDGVItem(vDGVItem, vPrimaryTransaksi, frm)

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

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = -1
            mNoTransaksi = ""
            mTanggalTransaksi = Now
            mPrimaryCust = -1
            mIndentifier = ""
            mSubject = ""
            mAttn = ""
            mSignedBy = ""
            mSignedPosition = ""
            mRemarks = ""

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

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Try
            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.CollectionLetter) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        MsgBox("Document number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Else
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.CollectionLetter) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimary(Me.mNoTransaksi) <> Me.mPrimary Then
                            Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                        End If
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimary(Me.mNoTransaksi) <> Me.mPrimary Then
                            MsgBox("Document number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vDGVItem As DataGridView, _
        ByVal vUserCode As String, ByVal Frm As FrmTransaksiTagihan) As Boolean
        Dim DTA As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihanTableAdapter
        Dim ObjItem As New ClsTransaksiTagihanItem
        Dim Urutan As Double

        Try
            If Me.ApakahDataBenar(True) = True Then
                Urutan = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, _
                Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.CollectionLetter, Me.mTanggalTransaksi)

                If DTA.Insert(Me.mNoTransaksi, Me.TanggalTransaksi, Me.PrimaryCust, Me.Identifier, _
                 Me.Subject, Me.Attn, Me.SignedBy, Me.SignedPosition, Me.mRemarks, vUserCode, _
                  Now, vUserCode, Now, Now, Urutan) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
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
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mRefNumber = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru(vDGVItem, vUserCode, Frm)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, ByVal Frm As FrmTransaksiTagihan) As Boolean
        Dim DTA As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihanTableAdapter
        Dim ObjItem As New ClsTransaksiTagihanItem
        Dim Urutan As Double = 0

        Try
            If Me.ApakahDataBenar(False) = True Then
                Urutan = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, _
                Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.CollectionLetter, Me.mTanggalTransaksi)

                If DTA.Update(Me.mNoTransaksi, Me.TanggalTransaksi, Me.PrimaryCust, Me.Identifier, _
                 Me.Subject, Me.Attn, Me.SignedBy, Me.SignedPosition, Me.mRemarks, vUserCode, _
                 Now, Now, Urutan, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.hapusitem(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.SimpanItem(vDGVItem, Me.mPrimary, Frm)
                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mRefNumber = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanPerubahan(vDGVItem, vUserCode, Frm)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, _
    ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean) As Boolean
        Dim DTA As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihanTableAdapter
        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, UserCode) = True Then
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ModGlobalVariable.EnumSearchCriteria, _
             Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
             Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiTagihanTableAdapters.tblcaritrans_surat_tagihan_viewTableAdapter
        Try
            If vSearchBy = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(vKriteria + "%")
            ElseIf vSearchBy = EnumSearchCriteria.ByDate Then
                Return DTA.ScalarQuery_GetJmlByTglDariSampai(vTanggalSampai, vTanggalDari)
            ElseIf vSearchBy = EnumSearchCriteria.ByCustomerCode Then
                Return DTA.ScalarQueryGetJmlByCustCodeLIKE(vKriteria + "%")
            ElseIf vSearchBy = EnumSearchCriteria.ByCustomerName Then
                Return DTA.ScalarQuery_GetJmlByCustNameLIKE(vKriteria + "%")
            ElseIf vSearchBy = EnumSearchCriteria.ByInvoiceNumber Then
                Return DTA.ScalarQuery_GetJmlByINVNoLIKE(vKriteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiTagihan.tblcaritrans_surat_tagihan_viewDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiTagihanTableAdapters.tblcaritrans_surat_tagihan_viewTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDate Then
                QueryResult = DTA.FillBy_TanggalDariSampai(TBL, vTanggalSampai, vTanggalDari)
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerCode Then
                QueryResult = DTA.FillBy_CustCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerName Then
                QueryResult = DTA.FillBy_CustNameLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByInvoiceNumber Then
                QueryResult = DTA.FillBy_INVnoLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Reference Number"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(3).HeaderText = "Customer Code"
                DGV.Columns(2).HeaderText = "Date"
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(4).HeaderText = "Customer Name"
                DGV.Columns(5).HeaderText = "Remarks"
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) _
                    + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + _
                    ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) _
                    + " miliseconds"
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

    Public Function GetJml_CariInvoice(ByVal vKriteria As String, ByVal vSearchBy As ModGlobalVariable.EnumSearchCriteria, _
             Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
             Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiTagihanTableAdapters.DataTable4TableAdapter
        Try
            If vSearchBy = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(vKriteria + "%")
            ElseIf vSearchBy = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlByTanggalSampaiDari(vTanggalSampai, vTanggalDari)
            ElseIf vSearchBy = EnumSearchCriteria.ByCustomerCode Then
                Return DTA.ScalarQuery_GetJmlByCustCodeLIKE(vKriteria + "%")
            ElseIf vSearchBy = EnumSearchCriteria.ByCustomerName Then
                Return DTA.ScalarQuery_GetJmlByCustNameLIKE(vKriteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV_CariInvoice(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiTagihan.DataTable4DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiTagihanTableAdapters.DataTable4TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TanggalSampaiDari(TBL, vTanggalSampai, vTanggalDari)
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerCode Then
                QueryResult = DTA.FillBy_CustCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerName Then
                QueryResult = DTA.FillBy_CustNameLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) _
                    + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + _
                    ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) _
                    + " miliseconds"
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

    Public Function CekApakahNoInvoicePunyaCustomerNya(ByVal vCustCode As String, ByVal vInvoice As String) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_viewTableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_viewDataTable
        Dim Row As DseTransaksiSalesInvoice.tbltrans_salesinvoice_viewRow
        Try
            If DTA.FillBy_CustCode(TBL, vCustCode) > 0 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                For Each Row In TBL.Rows
                    If vInvoice = Row.Invoice_No Then
                        Return True
                        Exit Function
                    End If
                Next
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function CekInvoiceNumber(ByVal vDGV As DataGridView, ByVal vPrimaryTagihan As Double, _
      ByVal Frm As FrmTransaksiTagihan) As String
        Dim DTA As New DseTransaksiTagihanTableAdapters.DataTable3TableAdapter
        Dim TBL As New DseTransaksiTagihan.DataTable3DataTable
        Dim Row As DseTransaksiTagihan.DataTable3Row
        Dim RowDGV As DataGridViewRow
        Dim RowMsg As DseTransaksiTagihan.TBL_MSGRow
        Dim DS As New DseTransaksiTagihan
        Dim vRetval As String = ""
        Dim vPembanding As String = ""
        Dim vNo As Integer = 1

        Try
            If vDGV.Rows.Count > 0 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                DTA.Fill(TBL)
                For Each RowDGV In vDGV.Rows
                    For Each Row In TBL.Rows
                        If RowDGV.Cells(Frm.ColInvoiceNumber.Index).Value = Row.Invoice_Number And _
                        vPrimaryTagihan <> Row.Primary_Transaksi Then
                            DS.TBL_MSG.DefaultView.RowFilter = "[Invoice_No] = '" + Row.Invoice_Number + "' AND [No_Tagihan] = '" + Row.No_Surat_Tagihan + "'"
                            If DS.TBL_MSG.DefaultView.Count = 0 Then
                                RowMsg = DS.TBL_MSG.NewRow
                                RowMsg.Invoice_No = Row.Invoice_Number
                                RowMsg.No_Tagihan = Row.No_Surat_Tagihan
                                DS.TBL_MSG.Rows.Add(RowMsg)
                            End If
                            DS.TBL_MSG.DefaultView.RowFilter = ""
                        End If
                    Next
                Next
                ''Buat msgbox('')
                For i As Integer = 0 To DS.TBL_MSG.Rows.Count - 1
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    RowMsg = DS.TBL_MSG.Rows(i)
                    If i = 0 Then
                        vPembanding = RowMsg.Invoice_No
                        vRetval = vNo.ToString + "." + "   " + RowMsg.Invoice_No + " " + "on" + " " + RowMsg.No_Tagihan
                        vNo += 1
                    Else
                        If RowMsg.Invoice_No = vPembanding Then
                            vRetval += "," + "  " + RowMsg.No_Tagihan
                        Else
                            vPembanding = RowMsg.Invoice_No
                            vRetval += vbCrLf + vNo.ToString + "." + "   " + RowMsg.Invoice_No + " " + "on" + " " + RowMsg.No_Tagihan
                            vNo += 1
                        End If
                    End If
                Next
            End If

            Return vRetval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DS.Dispose()
        End Try
    End Function

    Public Function CetakTransaksiTagihan(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimary As Double, _
            ByVal vPrimaryCust As Double, Optional ByVal vRealPrint As Boolean = False, Optional ByVal vPrinterName As String = "", _
            Optional ByVal vExport As Boolean = False, Optional ByVal vPrimaryEmailAccount As Double = -1) As Boolean
        Dim Report As New CrpCetakTransaksiSuratTagihan
        Dim DS As New DseTransaksiTagihan
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim DTAHeader As New DseTransaksiTagihanTableAdapters.tbl_Cetak_HeaderTableAdapter
        Dim DTAItem As New DseTransaksiTagihanTableAdapters.Tbl_Cetak_ItemTableAdapter
        Dim DTACust As New DseTransaksiTagihanTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim TBLHeader As DseTransaksiTagihan.tbl_Cetak_HeaderDataTable = DS.tbl_Cetak_Header
        Dim TBLCust As DseTransaksiTagihan.tblmaster_customer_item_addressDataTable = DS.tblmaster_customer_item_address
        Dim TBlItem As DseTransaksiTagihan.Tbl_Cetak_ItemDataTable = DS.Tbl_Cetak_Item
        Dim rowHeader As DseTransaksiTagihan.tbl_Cetak_HeaderRow
        Dim rowCust As DseTransaksiTagihan.tblmaster_customer_item_addressRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DefaultPrecision As Integer = 0
        Dim vRetval As String = ""
        Dim vNo As String = ""
        Dim vSubject As String = ""
        Dim vCustName As String = ""
        Dim vSignUp As String = ""
        Dim vSignUpPosition As String = ""
        Dim vTanggal As Date = Now
        Dim vAttn As String = ""
        Dim ObjCompany As New Core_Module.ClsMasterCompanyData
        Dim City As String = ""
        Dim RemarksBwh As String = ""

        Try
            Dim ObjSetting As New Core_Module.ClsTransaksiSetting
            Dim ObjCustom As New CustomSMESS.ClsTransaksiTagihan
            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SuratTagihan, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomPrint) = "1" Then
                Return ObjCustom.cetakSuratTagihan(CRV, vPrimary, vPrimaryCust, vRealPrint, vPrinterName, _
                vExport, vPrimaryEmailAccount)
                Exit Function
            End If

            If DTAHeader.FillBy_Primary(TBLHeader, vPrimary) > 0 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                DTAHeader.FillBy_Primary(TBLHeader, vPrimary)
                For Each rowHeader In TBLHeader.Rows
                    vTanggal = rowHeader.Tanggal_Transaksi
                    vNo = rowHeader.No_Transaksi
                    vSubject = rowHeader.Subject
                    vCustName = rowHeader.Customer_Name
                    vSignUp = rowHeader.Signed_By
                    vSignUpPosition = rowHeader.Signed_By_Position
                    vAttn = rowHeader.Attn
                    RemarksBwh = rowHeader.Remarks
                    Exit For
                Next

                If DTAItem.FillBy_Primary(TBlItem, vPrimary) Then
                    rowHeader = TBLHeader.Rows(0)
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    DTAItem.FillBy_Primary(TBlItem, vPrimary)
                    DTACust.FillBy_Identifier(TBLCust, rowHeader.Identifier_Alamat_Cust)

                    For Each rowHeader In TBLHeader.Rows
                        For Each rowCust In TBLCust.Rows
                            If rowHeader.Primary_Customer = rowCust.Primary_Customer Then
                                If String.IsNullOrEmpty(rowCust.Line_1) = False Then
                                    vRetval = rowCust.Line_1.ToString + vbCrLf
                                End If
                                If String.IsNullOrEmpty(rowCust.Line_2) = False Then
                                    vRetval += rowCust.Line_2.ToString + vbCrLf
                                End If
                                If String.IsNullOrEmpty(rowCust.Line_3) = False Then
                                    vRetval += rowCust.Line_3.ToString + vbCrLf
                                End If
                                If String.IsNullOrEmpty(rowCust.Line_4) = False Then
                                    vRetval += rowCust.Line_4.ToString + vbCrLf
                                End If
                                If String.IsNullOrEmpty(rowCust.Line_5) = False Then
                                    vRetval += rowCust.Line_5.ToString
                                End If
                                Exit For
                            End If
                        Next
                    Next
                End If
            End If

            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_Format_Day.ParameterFieldName, FormatDay)
            Report.SetParameterValue(Report.Parameter_Format_Month.ParameterFieldName, FormatMonth)
            Report.SetParameterValue(Report.Parameter_Format_Year.ParameterFieldName, FormatYears)
            Report.SetParameterValue(Report.Parameter_Customer_Address.ParameterFieldName, vRetval)
            Report.SetParameterValue(Report.Parameter_Tanggal.ParameterFieldName, vTanggal)
            Report.SetParameterValue(Report.Parameter_No_Trans.ParameterFieldName, vNo)
            Report.SetParameterValue(Report.Parameter_Subject.ParameterFieldName, vSubject)
            Report.SetParameterValue(Report.Parameter_Customer_Name.ParameterFieldName, vCustName)
            Report.SetParameterValue(Report.Parameter_Sign_Up.ParameterFieldName, vSignUp)
            Report.SetParameterValue(Report.Parameter_Sign_Up_Position.ParameterFieldName, vSignUpPosition)
            Report.SetParameterValue(Report.Parameter_Attn.ParameterFieldName, vAttn)
            Report.SetParameterValue(Report.Parameter_Remarks_bwh.ParameterFieldName, RemarksBwh)

            City = ObjCompany.GetCompany_City
            If City <> "" Then
                City = City + ","
            End If
            Report.SetParameterValue(Report.Parameter_Company_City.ParameterFieldName, City)
            CRV.ReportSource = Report
            CRV.Show()

            'rnew
            If vExport = True Then
                If My.Computer.FileSystem.DirectoryExists(Application.StartupPath & "\CL") = False Then
                    My.Computer.FileSystem.CreateDirectory(Application.StartupPath & "\CL")
                End If
                If My.Computer.FileSystem.FileExists(Application.StartupPath & "\CL\" & Me.objEmailSettings.GetPrefxFilename(vPrimaryEmailAccount, Core_Module.ClsEmailSettings.EnumTipeTransaksiEmail.CollectionLetter) & vNo.Replace("/", "-") & ".pdf") = True Then
                    My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\CL\" & Me.objEmailSettings.GetPrefxFilename(vPrimaryEmailAccount, Core_Module.ClsEmailSettings.EnumTipeTransaksiEmail.CollectionLetter) & vNo.Replace("/", "-") & ".pdf", FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                Report.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, Application.StartupPath & "\CL\" & Me.objEmailSettings.GetPrefxFilename(vPrimaryEmailAccount, Core_Module.ClsEmailSettings.EnumTipeTransaksiEmail.CollectionLetter) & vNo.Replace("/", "-") & ".pdf")
            End If

            ''cetak ke fisik jika diminta :(
            If vRealPrint = True Then
                Report.PrintOptions.PrinterName = vPrinterName
                Report.PrintToPrinter(1, True, 1, 99)
                Exit Function
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTAHeader.Dispose()
            DTAItem.Dispose()
            DTACust.Dispose()
            ObjCompany = Nothing
        End Try
    End Function

    Public Sub IsiDGVBelumLunas(ByVal vTanggalSuratTagihan As Date, ByVal vDGV As DataGridView, _
            ByVal vPrimaryCust As Double, ByVal i As FrmTransaksiTagihan, ByVal ObjCD As Core_Module.ClsCekData)
        Dim DTA As New DseTransaksiTagihanTableAdapters.balance_invoice_viewTableAdapter
        Dim TBL As New DseTransaksiTagihan.balance_invoice_viewDataTable
        Dim row As DseTransaksiTagihan.balance_invoice_viewRow
        Dim rowdgv As DataGridViewRow
        Dim ObjInv As New ClsTransaksiSalesInvoice
        Dim PrimaryInv As Double = -1

        Try
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDGV.Rows.Clear()
            DTA.Fill(TBL, vPrimaryCust, vTanggalSuratTagihan)
            For Each row In TBL.Rows
                PrimaryInv = ObjInv.GetPrimary(row.Invoice_Number)
                If PrimaryInv <> -1 Then
                    vDGV.Rows.Add()
                    rowdgv = vDGV.Rows(vDGV.RowCount - 1)

                    With rowdgv
                        .Cells(i.ColAmount.Index).Value = ObjCD.DigitGrouping(row.balance_original, row.Default_Precision)
                        .Cells(i.ColCurrency.Index).Value = row.Currency_Code
                        .Cells(i.ColDueDate.Index).Value = Format(row.Due_Date, "dd MMM yyyy")
                        .Cells(i.ColInvoiceDate.Index).Value = Format(row.Invoice_Date, "dd MMM yyyy")
                        .Cells(i.ColInvoiceNumber.Index).Value = row.Invoice_Number
                        .Cells(i.ColNo.Index).Value = rowdgv.Index + 1
                        .Cells(i.ColPrimaryInvoice.Index).Value = PrimaryInv
                        .Cells(i.ColRemarks.Index).Value = ""
                    End With
                End If
            Next

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjInv = Nothing
        End Try
    End Sub

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihanTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As Double = 0
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            Me.ObjNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.CollectionLetter, vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetUrutanMax(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetUrutanMax(TanggalEnd, TanggalStart).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            Return Me.ObjNumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.CollectionLetter)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function LoadEmailAccount_CollectionLetter(ByVal vUserCode As String, ByVal vTbl As DseTransaksiTagihan.DataTable5DataTable) As Boolean
        Dim dta As New DseTransaksiTagihanTableAdapters.DataTable5TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.FillBy_UserCode(vTbl, vUserCode)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vPrimaryCustomer As Double, ByVal vDatetimeMoreThan As Date) As Boolean
        Dim dta As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihanTableAdapter
        Dim count As Nullable(Of Long) = 0
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            count = dta.ScalarQuery_GetJmlByPrimaryCustomerAndCreatedDateMoreThan(vPrimaryCustomer, vDatetimeMoreThan)
            If count.HasValue Then
                If count.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function
End Class
