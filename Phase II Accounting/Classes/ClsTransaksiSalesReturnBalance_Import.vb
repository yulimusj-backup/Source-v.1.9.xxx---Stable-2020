Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesReturnBalance_Import
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mTotalSalesReturn As Double = -1
    Private mTotalVAT As Double = -1
    Private mTotalDutystamp As Double = -1
    Private mGrandTotal As Double = -1
    Private mTotalPaid As Double = -1
    Private mRemarks As String = ""
    Private mNoTransaksi As String = ""
    Private mUrutan As Double = 0

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Property NoTransaksi() As String
        Get
            Return mNoTransaksi
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property

    Public Property Urutan() As String
        Get
            Return mUrutan
        End Get
        Set(ByVal value As String)
            mUrutan = value
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

    Public Property TotalSalesReturn() As Double
        Get
            Return mTotalSalesReturn
        End Get
        Set(ByVal value As Double)
            mTotalSalesReturn = value
        End Set
    End Property

    Public Property TotalVAT() As Double
        Get
            Return mTotalVAT
        End Get
        Set(ByVal value As Double)
            mTotalVAT = value
        End Set
    End Property

    Public Property TotalDutystamp() As Double
        Get
            Return mTotalDutystamp
        End Get
        Set(ByVal value As Double)
            mTotalDutystamp = value
        End Set
    End Property

    Public Property GrandTotal() As Double
        Get
            Return mGrandTotal
        End Get
        Set(ByVal value As Double)
            mGrandTotal = value
        End Set
    End Property

    Public Property TotalPaid() As Double
        Get
            Return mTotalPaid
        End Get
        Set(ByVal value As Double)
            mTotalPaid = value
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

    Public Function GetPaymentTerm(ByVal vPrimarySR As Double) As String
        Dim DTA As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetPaymentTerm(vPrimarySR)) = False Then
                Return DTA.ScalarQuery_GetPaymentTerm(vPrimarySR)
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

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_importTableAdapter
        Try
            If dta.ScalarQuery_GetPrimary(vNoTransaksi).HasValue Then
                Return dta.ScalarQuery_GetPrimary(vNoTransaksi).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetNotransaksi(ByVal vPrimary As Double) As String
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_importTableAdapter
        Try
            If String.IsNullOrEmpty(dta.ScalarQuery_GetNoTransaksi(vPrimary)) = False Then
                Return dta.ScalarQuery_GetNoTransaksi(vPrimary)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_importTableAdapter
        Try
            If dta.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return dta.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_importTableAdapter
        Try
            If String.IsNullOrEmpty(dta.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                Return dta.ScalarQuery_GetCreatedBy(vPrimary)
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

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
          ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
          Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_importTableAdapter
        Dim ObjItem As New ClsTransaksiSalesReturnBalance_Import_Item
        Dim ObjJurnal As New ClsTransaksiJournal
        Dim Jurnal As String = ""

        Try

            '1. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            Dim CreatedBy As String = dta.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If CreatedBy = vSalesStaff Then
            Else
                If TampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '2. jika sudah ada sales return payment
            Jurnal = ObjJurnal.Get_SalesReturnBalanceTransfer_Jurnal(vPrimaryTransaksi)
            If Jurnal <> "" Then
                If TampilkanPesan = True Then
                    MsgBox("This transaction referred by sales return payment journal: " + Jurnal + _
                    ". You cannot delete this data.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If

            'nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            ObjItem = Nothing
            ObjJurnal = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
     ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
     Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_importTableAdapter
        Dim CreatedBy As String = ""
        Dim ObjJurnal As New ClsTransaksiJournal
        Dim Jurnal As String = ""

        Try

            '1. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            If StrictEditting = True Then
                CreatedBy = dta.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If CreatedBy <> vSalesStaff Then
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            '2. jika sudah ada sales receivement
            Jurnal = ObjJurnal.Get_SalesReturnBalanceTransfer_Jurnal(vPrimaryTransaksi)
            If Jurnal <> "" Then
                If TampilkanPesan = True Then
                    MsgBox("This transaction referred by sales return payment journal: " + Jurnal + _
                    ". You cannot edit this data.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If

            'sampe sini balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            ObjJurnal = Nothing
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_importTableAdapter
        Try
            If dta.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi).HasValue Then
                If dta.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi).Value > 0 Then
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
            dta.Dispose()
        End Try
    End Function

    ''Auto Number
    ''Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
    Public Function GetNoBaru() As String
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_importTableAdapter
        'Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As String
        'Dim Bulan As String
        'Dim ObjSupplier As New Core_Module.clsMasterSupplier
        Dim TipeStr As String = ""
        Dim objcd As New Core_Module.ClsCekData
        Try
            'cari tanggal sampai dan dari
            ''TanggalEnd = objcd.GetLasDateOfTheMonth(vTanggalTransaksi)
            ''TanggalStart = objcd.GetFirstDateofTheMonth(vTanggalTransaksi)

            'cari no urut
            If dta.ScalarQuery_GetMAXUrutan().HasValue = True Then
                NoUrut = CStr(dta.ScalarQuery_GetMAXUrutan().Value + 1)
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
            ''Return "SBI" + "/" + CStr(Format(vTanggalTransaksi, "MMM")) + "/" + _
            ''CStr(Format(vTanggalTransaksi, "yy")) + "/" + NoUrut

            Return "SRBI" + "/" + NoUrut
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    '' Merujuk pada Auto Number 
    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Try
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru()
                End If
            End If
            ''nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vDGVItem As DataGridView, _
     ByVal vUserCode As String, ByVal Frm As FrmTransaksiSalesReturnBalance_Import) As Boolean
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_importTableAdapter
        Dim ObjItem As New ClsTransaksiSalesReturnBalance_Import_Item
        Try
            If Me.ApakahDataBenar(True) = True Then
                Me.mUrutan = CDbl(Me.mNoTransaksi.Substring(5, 6))
                If dta.Insert(Me.mNoTransaksi, Me.mUrutan, Me.mTotalSalesReturn, _
                  Me.TotalVAT, Me.mTotalDutystamp, Me.mGrandTotal, Me.TotalPaid, Me.Remarks, _
                  vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = dta.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
                    Me.mTimestamp = dta.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItem(vDGVItem, Me.mPrimary, Frm)
                    'simpan data balance
                    Me.TransferBalance(vDGVItem, vUserCode, Frm)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 3" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru(vDGVItem, vUserCode, Frm)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            dta.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, ByVal Frm As FrmTransaksiSalesReturnBalance_Import) As Boolean
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_importTableAdapter
        Dim ObjItem As New ClsTransaksiSalesReturnBalance_Import_Item
        Try
            If Me.ApakahDataBenar(False) = True Then
                Me.Urutan = CDbl(Me.mNoTransaksi.Substring(5, 6))
                If dta.Update(Me.mNoTransaksi, Me.mUrutan, Me.mTotalSalesReturn, _
                  Me.TotalVAT, Me.mTotalDutystamp, Me.mGrandTotal, Me.TotalPaid, Me.Remarks, vUserCode, _
                 Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = dta.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data balance
                    Me.HapusBalance(Me.mPrimary)
                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.SimpanItem(vDGVItem, Me.mPrimary, Frm)
                    'simpan data balance
                    Me.TransferBalance(vDGVItem, vUserCode, Frm)

                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, _
    ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean) As Boolean
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_importTableAdapter
        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, UserCode) = True Then
                'hapus data balance
                Me.HapusBalance(vPrimaryTransaksi)
                'hapus data header, data item otomatis dihapuskan oleh server
                If dta.Delete(vPrimaryTransaksi) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = -1
            mNoTransaksi = ""
            mUrutan = 0
            mTotalSalesReturn = 0
            mTotalVAT = 0
            mTotalDutystamp = 0
            mGrandTotal = 0
            mTotalPaid = 0
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

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vDGVItem As DataGridView, ByVal frm As FrmTransaksiSalesReturnBalance_Import) As Boolean
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_importTableAdapter
        Dim TBL As New DseTransaksiSalesReturnBalance_Import.tbltrans_salesreturnbalance_importDataTable
        Dim row As DseTransaksiSalesReturnBalance_Import.tbltrans_salesreturnbalance_importRow
        Dim ObjItem As New ClsTransaksiSalesReturnBalance_Import_Item
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            'fil basis data
            If dta.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = vPrimaryTransaksi
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mUrutan = .Urutan
                    Me.mTotalSalesReturn = ObjCD.DigitGrouping(.Total_Sales_Return)
                    Me.mTotalVAT = ObjCD.DigitGrouping(.Total_VAT)
                    Me.mTotalDutystamp = ObjCD.DigitGrouping(.Total_Dutystamp)
                    Me.mGrandTotal = ObjCD.DigitGrouping(.Grand_Total)
                    Me.mTotalPaid = ObjCD.DigitGrouping(.Total_Paid)
                    Me.mRemarks = .Remarks
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'salin data item ada 2 item 
                'CodeNo
                ObjItem.IsiDGVItem(vDGVItem, frm, vPrimaryTransaksi)

                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDGVItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ModGlobalVariable.EnumSearchCriteria, _
      Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
      Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiSalesReturnBalance_ImportTableAdapters.Tbl_CariTableAdapter
        Try
            If vSearchBy = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(vKriteria + "%")
            ElseIf vSearchBy = EnumSearchCriteria.ByInvoiceNumber Then
                Return DTA.ScalarQuery_GetJmlBySRNumberLIKE(vKriteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiSalesReturnBalance_Import.Tbl_CariDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiSalesReturnBalance_ImportTableAdapters.Tbl_CariTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByInvoiceNumber Then
                QueryResult = DTA.FillBy_SRNumberLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "No Transaksi"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Remarks"
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

    ''Validasi Key valuetion (double data)
    Public Function CekInvoiceNumber_Item(ByVal vPrimarySBI As Double, ByVal vPrimaryInvoice As Double, ByVal vInvoiceNumber As String, _
    ByVal Frm As FrmTransaksiSalesReturnBalance_Import) As String
        Dim vRetval As String = ""
        Dim vNoTransaksi As String = ""
        Dim DTA As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_import_itemTableAdapter
        Dim TBL As New DseTransaksiSalesReturnBalance_Import.tbltrans_salesreturnbalance_import_itemDataTable
        Dim Row As DseTransaksiSalesReturnBalance_Import.tbltrans_salesreturnbalance_import_itemRow
        Try
            DTA.Fill(TBL)
            For Each Row In TBL.Rows
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                ''If vPrimarySBI <> Row.Primary_Transaksi Then
                If vPrimaryInvoice = Row.Primary_Invoice And vInvoiceNumber = Row.Inv_Number Then
                    vNoTransaksi = Me.GetNotransaksi(Row.Primary_Transaksi)
                    vRetval = "Data sales return" + " " + vInvoiceNumber + " " + "has been recorded on sales return balance import no" + " " + vNoTransaksi + _
                    ", please input another sales return."
                    Exit For
                End If
                ''End If
            Next

            Return vRetval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function CekJmlInvoiceNumber(ByVal vPrimaryInvoice As Double) As Double
        Dim DTA As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_import_itemTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByPrimaryInvoice(vPrimaryInvoice).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryInvoice(vPrimaryInvoice).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CekJmlInvoiceNumber_EDIT(ByVal vPrimaryInvoice As Double, ByVal vPrimarySBI As Double) As Double
        Dim DTA As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_import_itemTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByPrimaryInvoice__Transaksi(vPrimaryInvoice, vPrimarySBI).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryInvoice__Transaksi(vPrimaryInvoice, vPrimarySBI).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''Moving Data Sales,VAT, dan Dutystamp
    Public Function GetDutystamp(ByVal vPrimaryInv As Double) As Double
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnTableAdapter
        Try
            If dta.ScalarQuery_GetDutystamp(vPrimaryInv).HasValue Then
                Return dta.ScalarQuery_GetDutystamp(vPrimaryInv)
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetVAT(ByVal vPrimaryInv As Double) As Double
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnTableAdapter
        Try
            If dta.ScalarQuery_GetVAT(vPrimaryInv).HasValue Then
                Return dta.ScalarQuery_GetVAT(vPrimaryInv)
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetSales(ByVal vPrimaryInv As Double) As Double
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnTableAdapter
        Try
            If dta.ScalarQuery_GetJmlSalesReturn(vPrimaryInv).HasValue Then
                Return dta.ScalarQuery_GetJmlSalesReturn(vPrimaryInv)
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCurrency(ByVal vPrimaryInv As Double) As Double
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnTableAdapter
        Try
            Return dta.ScalarQuery_GetPrimaryCurr(vPrimaryInv).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetTaxExchangeRate(ByVal vPrimaryInv As Double) As Double
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnTableAdapter
        Try
            Return dta.ScalarQuery_GetTaxExchangeRate(vPrimaryInv).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetNilaiSalesInvoice(ByVal vGetsales As Double, ByVal vGetPrimaryCurrency As Double, ByVal vPrimaryInv As Double) As Double
        Dim Total As Double = 0
        Dim Tax As Double = 1
        Try
            If vGetPrimaryCurrency <> 1 Then
                Tax = Me.GetTaxExchangeRate(vPrimaryInv)
                Total = Math.Round(vGetsales * Tax, 0, MidpointRounding.AwayFromZero)
            Else
                Total = vGetsales
            End If
            Return Total
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''Hitung Invoice Due date 
    Public Function Get_DUE_DATE(ByVal vInvoicedate As Date, ByVal vPrimaryInvoice As Double) As Date
        Dim vDueDate As Date = Now
        Dim Term As String = ""
        Dim ij As Double = 0
        Dim dta As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnTableAdapter
        Try
            'Dim i As String = "30"
            'Dim ij As Double = 0
            ''If IsNumeric(i) = True Then
            ''ij = CDbl(i)
            ''Else
            ''ij = 0
            ''End If
            ''DueDate = #1/1/1950# + TimeSpan.FromDays(ij)
            Term = dta.ScalarQuery_GetPaymentTerm(vPrimaryInvoice)
            If String.IsNullOrEmpty(Term) = False Then
                If IsNumeric(Term) = True Then
                    ij = CDbl(Term)
                Else
                    ij = 0
                End If
                ''vDueDate = #1/1/1950# + TimeSpan.FromDays(ij)
                vDueDate = vInvoicedate + TimeSpan.FromDays(ij)
            End If

            Return vDueDate
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    ''Hitung Total pada Header
    Public Sub GetTotalHeader(ByVal vDGVItem As DataGridView, ByVal frm As FrmTransaksiSalesReturnBalance_Import, _
    ByRef TotalSales As Double, ByRef TotalVAT As Double, ByRef TotalBalance As Double, _
    ByRef TotalDutystamp As Double)
        Dim RowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            For Each RowDGV In vDGVItem.Rows
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                TotalSales += Math.Round(CDbl(RowDGV.Cells(frm.ColSales.Index).Value), 0, MidpointRounding.AwayFromZero)
                TotalVAT += Math.Round(CDbl(RowDGV.Cells(frm.ColVAT.Index).Value), 0, MidpointRounding.AwayFromZero)
                TotalDutystamp += Math.Round(CDbl(RowDGV.Cells(frm.ColDutystamp.Index).Value), 0, MidpointRounding.AwayFromZero)
                TotalBalance += Math.Round(CDbl(RowDGV.Cells(frm.ColBalance.Index).Value), 0, MidpointRounding.AwayFromZero)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Public Function GetGrandTotal(ByVal frm As FrmTransaksiSalesReturnBalance_Import, ByVal vDGVItem As DataGridView) As Double
        Dim RowDGV As DataGridViewRow
        Dim GrandTotal As Double = 0
        Try
            If vDGVItem.Rows.Count > 0 Then
                For Each RowDGV In vDGVItem.Rows
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    GrandTotal += Math.Round(CDbl(RowDGV.Cells(frm.ColSales.Index).Value) + CDbl(RowDGV.Cells(frm.ColVAT.Index).Value) + _
                    CDbl(RowDGV.Cells(frm.ColDutystamp.Index).Value), 0, MidpointRounding.AwayFromZero)
                Next
            End If
            Return GrandTotal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function HitungBalance(ByVal vSales As Double, ByVal vVAT As Double, ByVal vDutystamp As Double, _
    ByVal vRecieved As Double) As String
        Dim vBalance As Double = 0
        Try
            vBalance = Math.Round(((vSales + vVAT + vDutystamp) - vRecieved), 0, MidpointRounding.AwayFromZero)
            Return vBalance
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub TransferBalance(ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, ByVal Frm As FrmTransaksiSalesReturnBalance_Import)
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_balance_salesreturnTableAdapter
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = _
        ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice
        Dim rowdgv As DataGridViewRow

        Try
            For Each rowdgv In vDGVItem.Rows
                With rowdgv
                    DTA.Insert(.Cells(Frm.ColPrimaryCust.Index).Value, .Cells(Frm.ColPrimaryInv.Index).Value, TipeBalance, _
                    .Cells(Frm.ColSR_Number.Index).Value, .Cells(Frm.ColSR_Date.Index).Value, .Cells(Frm.ColRemarks.Index).Value, .Cells(Frm.Col_SR_DueDate.Index).Value, _
                    .Cells(Frm.ColSales.Index).Value, .Cells(Frm.ColVAT.Index).Value, .Cells(Frm.ColDutystamp.Index).Value, .Cells(Frm.ColBalance.Index).Value, .Cells(Frm.ColPaid.Index).Value, _
                    Me.mCreatedBy, Me.mCreatedDate, Me.mModifiedBy, Me.mModifiedDate, Me.mTimestamp)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub HapusBalance(ByVal vPrimaryTransaksi As Double)
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_balance_salesreturnTableAdapter
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = _
        ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice
        Dim DTAItem As New DseTransaksiSalesReturnBalance_ImportTableAdapters.tbltrans_salesreturnbalance_import_itemTableAdapter
        Dim TBLItem As New DseTransaksiSalesReturnBalance_Import.tbltrans_salesreturnbalance_import_itemDataTable
        Dim rowItem As DseTransaksiSalesReturnBalance_Import.tbltrans_salesreturnbalance_import_itemRow

        Try
            DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimaryTransaksi)
            For Each rowItem In TBLItem.Rows
                DTA.Delete(rowItem.Primary_Invoice, TipeBalance)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAItem.Dispose()
        End Try
    End Sub
End Class
