Option Explicit On
Option Strict Off

Public Class ClsTransaksiMonthlyClosing
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mNoTransaksi As String = ""
    Private mUrutan As Integer = 0
    'tanggal dilakukannya closing
    Private mClosingDate As Date = Today
    ''' <summary>
    ''' periode closing yang diwakili oleh tanggal pertama dari bulan yang bersangkutan
    ''' </summary>
    ''' <remarks></remarks>
    Private mClosingPeriode As Date = Today
    Private mRemarks As String

    Private mCreatedBy As String = ""
    Private mCreatedDate As Date = Today
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date = Today
    Private mTimestamp As Date

    'elemen yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian
    Dim ObjProduct As New Core_Module.clsMasterProduct


    Public Property Remarks() As String
        Get
            Return Me.mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Public Property ClosingPeriod() As Date
        Get
            Return Me.mClosingPeriode
        End Get
        Set(ByVal value As Date)
            mClosingPeriode = value
        End Set
    End Property

    Public Property ClosingDate() As Date
        Get
            Return Me.mClosingDate
        End Get
        Set(ByVal value As Date)
            mClosingDate = value
        End Set
    End Property

    Public Property Urutan() As Integer
        Get
            Return Me.mUrutan
        End Get
        Set(ByVal value As Integer)
            mUrutan = value
        End Set
    End Property

    Public Property NoTransaksi() As String
        Get
            Return Me.mNoTransaksi
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property

    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            Me.mPrimary = value
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

    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
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


    Public Function Cetak() As Boolean
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vNoTransaksi).Value > 0 Then
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

    Public Function GetNoBaru(ByVal vPeriodeClosing As Date) As String
        Dim Bulan As String = ""

        Try
            'bulan 
            Bulan = CStr(vPeriodeClosing.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While
            Return Bulan + "/" + CStr(vPeriodeClosing.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, _
    ByVal TampilkanPesan As Boolean) As Boolean
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim Bulan As String = ""
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'msgbox("belum")
            'data salah bila:
            '0. bila sudah ada transaksi closing untuk periode yang bersangkutan
            If vBaru = True Then
                If DTA.ScalarQuery_GetJml2(Me.mClosingPeriode).Value > 0 Then
                    If TampilkanPesan = True Then
                        Bulan = ObjCD.GetMonthOfTheYear(Me.mClosingPeriode.Month)
                        MsgBox("The purchase transaction for " + Bulan + " " + _
                        Me.mClosingPeriode.Year.ToString + "is already closed.", MsgBoxStyle.Information, _
                        "Information")
                    End If
                    Return False
                End If
            End If

            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                'set ulang no buat jaga- jaga
                Me.mNoTransaksi = Me.GetNoBaru(Me.mClosingPeriode)
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mClosingPeriode)
                End If
            End If

            '2. Jika semua transaksi purchase sedang dikunci

            '3. Jika sudah melewati batas penguncian transaksi

            '4. jika tidak ada item yang akan disimpan

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
      ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
      Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim ClosingDate, LastRevDate As Date
        Dim ObjRev As New AccModule.ClsTransaksiJournal
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjClosingJur As New AccModule.ClsTransaksiClosing

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'batasan purchase closing:
            '1. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If CreatedBy = vSalesStaff Then
                Else
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            '2. Jika sudah ada closing untuk bulan berikutnya
            If DTA.ScalarQuery_GetJml3(vPrimaryTransaksi).Value > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("Only the latest purchase closing can be editted.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '3. jika sudah ada transaksi revaluasi, maka tidak bisa edit: jika closing date inv <= tanggal revaluasi terakhir, maka tidak bisa edit
            ClosingDate = ObjCD.GetFirstDateofTheMonth(Me.GetClosingPeriod(vPrimaryTransaksi))
            LastRevDate = ObjCD.GetFirstDateofTheMonth(ObjRev.GetTglRevaluasiSelanjutnya - TimeSpan.FromDays(33))
            If ClosingDate <= LastRevDate Then
                MsgBox("To edit this data, please delete revaluation journal for the respected month.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '4 berkaitan dengan jurnal locking
            If ObjCD.GetLasDateOfTheMonth(ClosingPeriod) <= ObjClosingJur.GetClosingDate.Date Then
                MsgBox("You cannot edit this data because " + Format(ObjCD.GetLasDateOfTheMonth(ClosingPeriod), "dd MMM yyyy") + " <= journal lock date" _
                    , MsgBoxStyle.Information, "Information")
                Return False
            End If

            '3. Jika sudah pernah dicetak maka, ketika akan melakukan edit harus dengan persetujuan 
            'orang yang memiliki hak khusus
            'msgbox("belum")

            'sampe sini balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjRev = Nothing
            ObjCD = Nothing
            ObjClosingJur = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
        ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
        Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim ClosingDate, LastRevDate As Date
        Dim ObjRev As New AccModule.ClsTransaksiJournal
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjClosingJur As New AccModule.ClsTransaksiClosing

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            '1. Jika strict sales Staff delete maka yang boleh hapus hanya sales staff yang bersangkutan
            Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If CreatedBy = vSalesStaff Then
            Else
                If TampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '2. Jika sudah ada closing untuk bulan berikutnya
            If DTA.ScalarQuery_GetJml3(vPrimaryTransaksi).Value > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("Only the latest purchase closing can be deleted.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '3. jika sudah ada revaluasi untuk bulan ybs, maka tidak bisa delete
            ClosingDate = ObjCD.GetFirstDateofTheMonth(Me.GetClosingPeriod(vPrimaryTransaksi))
            LastRevDate = ObjCD.GetFirstDateofTheMonth(ObjRev.GetTglRevaluasiSelanjutnya - TimeSpan.FromDays(33))
            If ClosingDate <= LastRevDate Then
                MsgBox("To delete this data, please delete revaluation journal for the respected month.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '4 berkaitan dengan jurnal locking
            If ObjCD.GetLasDateOfTheMonth(ClosingPeriod) <= ObjClosingJur.GetClosingDate.Date Then
                MsgBox("You cannot delete this data because " + Format(ObjCD.GetLasDateOfTheMonth(ClosingPeriod), "dd MMM yyyy") + " <= journal lock date" _
                    , MsgBoxStyle.Information, "Information")
                Return False
            End If

            '3. Jika sudah pernah dicetak, maka minta minta otorisasi sebelum menghapus invoice
            'msgbox("belum")

            'nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjRev = Nothing
            ObjCD = Nothing
            ObjClosingJur = Nothing
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim dta As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(dta.ScalarQuery_GetNoTransaksi(vPrimary)) = False Then
                Return dta.ScalarQuery_GetNoTransaksi(vPrimary)
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

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vDgvItem As DataGridView, ByVal vDGVItemAdjust As DataGridView, _
    ByVal vTBLDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, _
    ByVal vTBLBundling As DseTransaksiMonthlyClosing2.TblBundling1DataTable, _
    ByVal i As FrmTransaksiMonthlyClosing2, _
    ByVal vDGVAkun As DataGridView, ByVal vTBLSettingItem As DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_tipe_adjustmentDataTable, _
    ByVal vDGVCostGroup As DataGridView, ByVal vTBLCostGroup_Account As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountDataTable, _
    ByVal vTBLCostGroup_Scope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeDataTable, _
    ByVal vTBLCostGroup_CodeNo As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noDataTable, _
    ByVal vFrm As FrmTransaksiMonthlyClosing2, _
    ByVal vTBLCostGroup_CodeNo2 As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_no2DataTable, _
    ByVal vTblProductQty As Inventory_Module.DseLaporanStockDetail2.TblCetakDataTable) As Boolean
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing.tbltrans_purchase_closingDataTable

        Dim row As DseTransaksiMonthlyClosing.tbltrans_purchase_closingRow
        Dim ObjItem As New ClsTransaksiMonthlyClosingItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mPrimary = .Primary
                    Me.mRemarks = .Remark
                    Me.mClosingPeriode = .Closing_Period
                    Me.mClosingDate = .Closing_Date

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With

                'salin data item 
                ObjItem.IsiDGV(vDgvItem, vDGVItemAdjust, Me.mPrimary, vTBLDetail, _
                vTBLBundling, i, vDGVAkun, vTBLSettingItem, _
                vDGVCostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, vTBLCostGroup_CodeNo, vFrm, vTBLCostGroup_CodeNo2, vTblProductQty)

                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDgvItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            Me.mNoTransaksi = ""
            Me.mPrimary = -1
            Me.mRemarks = ""
            Me.mClosingPeriode = Today
            Me.mClosingDate = Today

            Me.mCreatedBy = ""
            Me.mCreatedDate = Today
            Me.mModifiedBy = ""
            Me.mModifiedDate = Today
            Me.mTimestamp = Today


            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiMonthlyClosing.tbltrans_purchase_closing1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closing1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(3).DefaultCellStyle.Format = "MMM yyyy"
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

    ''' <summary>
    ''' mengembalikan nilai true jika primary transaksi yang diberikan adalah primary transaksi closing yang
    ''' paling awal.
    ''' </summary>
    ''' <param name="vPrimaryTransaksi"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ApakahClosingPalingAwal(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter

        Try
            If vPrimaryTransaksi = -1 Then
                If Me.GetJml = 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                If DTA.ScalarQuery_GetJml5(vPrimaryTransaksi).Value > 0 Then
                    Return False
                Else
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan no do untuk periode tertentu yang belum memiliki invoice
    ''' </summary>
    ''' <param name="vStartDODate"></param>
    ''' <param name="vEndDODate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDOPending(ByVal vStartDODate As Date, ByVal vEndDODate As Date) As String
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_doTableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing.tbltrans_doDataTable
        Dim row As DseTransaksiMonthlyClosing.tbltrans_doRow
        Dim Dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy(TBL, vEndDODate, vStartDODate)
            For Each row In TBL.Rows
                If Dummy = "" Then
                    Dummy = row.DO_Number
                Else
                    Dummy = Dummy + ", " + row.DO_Number
                End If
            Next
            Return Dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetDO_StockOutBedaPeriode(ByVal vStartDate As Date, ByVal vEndDate As Date) As String
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.DataTable28TableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing.DataTable28DataTable
        Dim Dummy As String = ""
        Dim row As DseTransaksiMonthlyClosing.DataTable28Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.FillBy_TanggalSampaiDari(TBL, vEndDate, vStartDate) > 0 Then
                For Each row In TBL.Rows
                    If Dummy = "" Then
                        Dummy = row.DO_Number
                    Else
                        Dummy = Dummy + ", " + row.DO_Number
                    End If
                Next

                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetDO_BelumInvoice(ByVal vStartDate As Date, ByVal vEndDate As Date) As String
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_do1TableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing.tbltrans_do1DataTable
        Dim Dummy As String = ""
        Dim row As DseTransaksiMonthlyClosing.tbltrans_do1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Fill(TBL, vEndDate, vStartDate) > 0 Then
                For Each row In TBL.Rows
                    If Dummy = "" Then
                        Dummy = row.No_Transaksi
                    Else
                        Dummy = Dummy + ", " + row.No_Transaksi
                    End If
                Next

                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetSR_StockInSalesReturnBedaPeriode(ByVal vStartDate As Date, ByVal vEndDate As Date) As String
        Dim TBL As New DseTransaksiMonthlyClosing.DataTable24DataTable
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.DataTable24TableAdapter
        Dim row As DseTransaksiMonthlyClosing.DataTable24Row
        Dim dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Fill(TBL, vEndDate, vStartDate) > 0 Then
                For Each row In TBL.Rows
                    If dummy = "" Then
                        dummy = row.No_Transaksi
                    Else
                        dummy = dummy + ", " + row.No_Transaksi
                    End If
                Next

                Return dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetStockInSalesReturn_SRBedaPeriode(ByVal vStartDate As Date, ByVal vEndDate As Date) As String
        Dim TBL As New DseTransaksiMonthlyClosing.tbltransaksi_stockinout1DataTable
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltransaksi_stockinout1TableAdapter
        Dim row As DseTransaksiMonthlyClosing.tbltransaksi_stockinout1Row
        Dim dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Fill(TBL, vEndDate, vStartDate) > 0 Then
                For Each row In TBL.Rows
                    If dummy = "" Then
                        dummy = row.Stock_Sales_Return_No
                    Else
                        dummy = dummy + ", " + row.Stock_Sales_Return_No
                    End If
                Next

                Return dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetGRN_AOP_BedaPeriode(ByVal vStartDate As Date, ByVal vEndDate As Date) As String
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.DataTable22TableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing.DataTable22DataTable
        Dim row As DseTransaksiMonthlyClosing.DataTable22Row
        Dim dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Fill(TBL, vEndDate, vStartDate) > 0 Then
                For Each row In TBL.Rows
                    If dummy = "" Then
                        dummy = row.GRN_Number + "(AOP: " + row.AOP_Number + ")"
                    Else
                        dummy = dummy + ", " + row.GRN_Number + "(AOP: " + row.AOP_Number + ")"
                    End If
                Next

                Return dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function Get_PR_Rev_BelumMasukAOP(ByVal vStartDate As Date, ByVal vEndDate As Date) As String
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.DataTable30TableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing.DataTable30DataTable
        Dim row As DseTransaksiMonthlyClosing.DataTable30Row
        Dim dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Fill(TBL, vEndDate, vStartDate) > 0 Then
                For Each row In TBL.Rows
                    If dummy = "" Then
                        dummy = row.No_Transaksi
                    Else
                        dummy = dummy + ", " + row.No_Transaksi
                    End If
                Next

                Return dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vClosingPeriod As Date) As Double
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimary2(vClosingPeriod).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary2(vClosingPeriod).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_DO_BelumInvoice(ByVal vStartDate As Date, ByVal vEndDate As Date) As String
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.DataTable17TableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing.DataTable17DataTable
        Dim row As DseTransaksiMonthlyClosing.DataTable17Row
        Dim dummy As String = ""
        Dim TipeStockOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = _
        Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(TBL, TipeStockOut, vEndDate, vStartDate) > 0 Then
                For Each row In TBL.Rows
                    If dummy = "" Then
                        dummy = row.No_Transaksi
                    Else
                        dummy = dummy + ", " + row.No_Transaksi
                    End If
                Next

                Return dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function Get_PR_BelumStockOut(ByVal vStartDate As Date, ByVal vEndDate As Date) As String
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchasereturnTableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing.tbltrans_purchasereturnDataTable
        Dim row As DseTransaksiMonthlyClosing.tbltrans_purchasereturnRow
        Dim dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Fill(TBL, vEndDate, vStartDate) > 0 Then
                For Each row In TBL.Rows
                    If dummy = "" Then
                        dummy = row.No_Transaksi
                    Else
                        dummy = dummy + ", " + row.No_Transaksi
                    End If
                Next

                Return dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function Get_SR_RevInvoice_BelumRevisi(ByVal vStartDate As Date, ByVal vEndDate As Date) As String
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_salesreturnTableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing.tbltrans_salesreturnDataTable
        Dim row As DseTransaksiMonthlyClosing.tbltrans_salesreturnRow
        Dim dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.FillBy_TanggalSampaiDari(TBL, vEndDate, vStartDate) > 0 Then
                For Each row In TBL.Rows
                    If dummy = "" Then
                        dummy = row.No_Transaksi
                    Else
                        dummy = dummy + ", " + row.No_Transaksi
                    End If
                Next

                Return dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetAOP_GRNBEdaPeriode(ByVal vStartDate As Date, ByVal vEndDate As Date) As String
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.DataTable23TableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing.DataTable23DataTable
        Dim row As DseTransaksiMonthlyClosing.DataTable23Row
        Dim dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Fill(TBL, vEndDate, vStartDate) > 0 Then
                For Each row In TBL.Rows
                    If dummy = "" Then
                        dummy = row.Grn_Number + "(AOP: " + row.AOP_Number + ")"
                    Else
                        dummy = dummy + ", " + row.Grn_Number + "(AOP: " + row.AOP_Number + ")"
                    End If
                Next

                Return dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetInvoice_StockOutSalesPreviousPeriode(ByVal vStartDODate As Date, ByVal vEndDODate As Date) As String
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.DataTable18TableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing.DataTable18DataTable

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Fill(TBL, vEndDODate, vStartDODate) > 0 Then
                Return "-"
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetJml() As Double
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Return DTA.ScalarQuery_GetJml4.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
       ByVal SearchCriteria As EnumSearchCriteria, Optional _
       ByVal vTanggalDari As Date = (#1/1/1950#), Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, _
        ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean) As Boolean
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim ObjST As New Inventory_Module.clsTransaksiStockTaking
        Dim DTA_Jurnal As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim vClosingPeriod As Date = #1/1/1950#

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Jurnal.Connection.ConnectionString = DTA.Connection.ConnectionString

            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, _
            UserCode) = True Then
                vClosingPeriod = Me.GetClosingPeriod(vPrimaryTransaksi)

                'hapus jurnal
                DTA_Jurnal.DeleteQuery(DTA.ScalarQuery_GetPrimaryJurnal(vPrimaryTransaksi).Value)

                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    'hapus data laporan stock taking --> fungsi sudah ada dalam simpan data penyesuaian
                    'ObjST.HapusDataPenyesuaianStockTaking(vPrimaryTransaksi)

                    'kosongkan nilai amount pada stock out material usage
                    Me.UpdateDataStockOutMaterialUsing_HapusClosing(vPrimaryTransaksi, vClosingPeriod)

                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjST = Nothing
            DTA_Jurnal.Dispose()
        End Try
    End Function

    Private Sub UpdateDataStockOutMaterialUsing_SimpanClosing(ByVal vDGVItem As DataGridView, _
    ByVal vTBLItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, _
    ByVal vFrm As FrmTransaksiMonthlyClosing2, ByVal vTBLSetting As DseTransaksiMonthlyClosing2.tbl_setting_adjustmentDataTable)
        Dim Amount As Decimal = 0
        Dim DV As Data.DataRowView
        Dim TBLDummy As New DseTransaksiMonthlyClosing2.TblOutAdjustDataTable
        Dim rowDummy, rowDummy2 As DseTransaksiMonthlyClosing2.TblOutAdjustRow
        Dim TipeOutAdjust As Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.NoTipe
        Dim ObjAsset As New Core_Module.ClsMasterAsset
        Dim ListPrimary As New List(Of Double)
        Dim PrimaryAkunInv As Double = -1
        Dim DTA_OutAdjust As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_out_adjustTableAdapter

        Try
            '1. constring
            DTA_OutAdjust.Connection.ConnectionString = Core_Module.GetConstring

            '2. buat tbl dummy
            Me.GetTBLDummy_OutAdjust_MaterialUsingAsset(vDGVItem, TBLDummy, vTBLItemDetail, vFrm, vTBLSetting, vFrm.LblProgress, vFrm.ProgressBar1)

            '3. update nilai amount pada stock out revaluasi dan acquisition
            ListPrimary.Clear()
            For Each rowDummy In TBLDummy.Rows
                Amount = 0
                If ListPrimary.Contains(rowDummy.Primary_Stock_Out) = False Then
                    ListPrimary.Add(rowDummy.Primary_Stock_Out)
                    TBLDummy.DefaultView.RowFilter = "[Primary Stock Out] = " + rowDummy.Primary_Stock_Out.ToString
                    For Each DV In TBLDummy.DefaultView
                        rowDummy2 = DV.Row
                        Amount += rowDummy2.Amount
                    Next
                    'update nilai amount ke stock out adjust
                    DTA_OutAdjust.UpdateQuery_UpdateAmount(Amount, rowDummy.Primary_Stock_Out)
                End If
            Next

            '4. update acquired value + book value untuk aquisition
            ListPrimary.Clear()
            TipeOutAdjust = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetAcquisition
            For Each rowDummy In TBLDummy.Rows
                If rowDummy.Tipe_Stock_Out = CInt(TipeOutAdjust) Then
                    If ListPrimary.Contains(rowDummy.Primary_Asset) = False Then
                        ListPrimary.Add(rowDummy.Primary_Asset)
                        ObjAsset.UpdateAcquiredValue(rowDummy.Primary_Asset)
                        ObjAsset.UpdateBookValue(rowDummy.Primary_Asset)
                    End If
                End If
            Next

            '5. update book value
            ListPrimary.Clear()
            TipeOutAdjust = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetRevaluation
            For Each rowDummy In TBLDummy.Rows
                If rowDummy.Tipe_Stock_Out = CInt(TipeOutAdjust) Then
                    If ListPrimary.Contains(rowDummy.Primary_Asset) = False Then
                        ListPrimary.Add(rowDummy.Primary_Asset)
                        ObjAsset.UpdateBookValue(rowDummy.Primary_Asset)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLDummy = Nothing
            ObjAsset = Nothing
            ListPrimary = Nothing
            DTA_OutAdjust.Dispose()
        End Try
    End Sub

    Public Sub GetTBLDummy_OutAdjust_MaterialUsingAsset(ByVal vDGVItem As DataGridView, _
    ByVal vTBLDummy As DseTransaksiMonthlyClosing2.TblOutAdjustDataTable, _
    ByVal vTBLItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, _
    ByVal vFrm As FrmTransaksiMonthlyClosing2, ByVal vTBLSetting As DseTransaksiMonthlyClosing2.tbl_setting_adjustmentDataTable, ByVal vLblProgressBar As Label, ByVal vPrg As ProgressBar)
        Dim rowItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim TipeOutAdjust As Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.NoTipe
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
        Dim TipeTransaksi As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
        Dim Amount As Decimal = 0
        Dim DV As Data.DataRowView
        Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut
        Dim PrimaryAsset As Double = -1
        Dim rowDummy As DseTransaksiMonthlyClosing2.TblOutAdjustRow
        Dim ObjAsset As New Core_Module.ClsMasterAsset
        Dim row_setting As DseTransaksiMonthlyClosing2.tbl_setting_adjustmentRow

        Dim DTAOutItem As New DseTransaksiMonthlyClosing3TableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TBLOutItem As New DseTransaksiMonthlyClosing3.tbltransaksi_stockinout_itemDataTable
        Dim rowOutItem As DseTransaksiMonthlyClosing3.tbltransaksi_stockinout_itemRow
        Dim TipeTrans1 As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust

        Try
            DTAOutItem.Fill(TBLOutItem, CInt(TipeTrans1), Me.ObjCD.GetLasDateOfTheMonth(vFrm.dtpPeriod.Value), Me.ObjCD.GetFirstDateofTheMonth(vFrm.dtpPeriod.Value))

            vTBLDummy.Rows.Clear()
            vPrg.Value = 0
            vPrg.Maximum = TBLOutItem.Rows.Count
            For Each rowOutItem In TBLOutItem.Rows
                vPrg.Value += 1
                If vPrg.Value > 25 Then
                    If Math.IEEERemainder(vPrg.Value, 100) = 0 Then
                        vLblProgressBar.Text = "Preparing - Data " + vPrg.Value.ToString + " of " + vPrg.Maximum.ToString
                    ElseIf vPrg.Value = vPrg.Maximum Then
                        vLblProgressBar.Text = "Preparing - Data " + vPrg.Value.ToString + " of " + vPrg.Maximum.ToString
                    End If
                Else
                    vLblProgressBar.Text = "Preparing - Data " + vPrg.Value.ToString + " of " + vPrg.Maximum.ToString
                End If


                If IsNumeric(rowOutItem.Primary_Product) AndAlso CDbl(rowOutItem.Primary_Product) > -1 Then
                    TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust
                    TipeTransaksi = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
                    vTBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + rowOutItem.Primary_Product.ToString + _
                     " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString + " AND [tipe item detail] = " + CInt(TipeTransaksi).ToString

                    Amount = 0

                    For Each DV In vTBLItemDetail.DefaultView
                        rowItemDetail = DV.Row
                        vTBLDummy.DefaultView.RowFilter = "[Primary Stock Out] = " + rowItemDetail.primary_transaksi_asal.ToString + _
                        " AND [Primary Product] = " + rowItemDetail.primary_code_no.ToString

                        If vTBLDummy.DefaultView.Count = 0 Then
                            'buat baris baru
                            rowDummy = vTBLDummy.NewTblOutAdjustRow

                            PrimaryAsset = ObjStockInOut.GetPrimaryAsset(rowItemDetail.primary_transaksi_asal)
                            rowDummy.Amount = rowItemDetail.amount
                            rowDummy.Primary_Adjustment_Type = -1
                            TipeOutAdjust = ObjStockInOut.GetTipeAdjustment(rowItemDetail.primary_transaksi_asal)
                            If TipeOutAdjust = -1 Then
                                TipeOutAdjust = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.NoTipe
                            End If
                            If TipeOutAdjust = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.NoTipe Then
                                rowDummy.Primary_Adjustment_Type = ObjStockInOut.GetPrimaryTipeAdjustment(rowItemDetail.primary_transaksi_asal)
                                If rowDummy.Primary_Adjustment_Type = -1 Then
                                    rowDummy.Primary_Account_Expense = rowOutItem.Primary_Adjustment_Account
                                    If rowItemDetail.primary_adjustment_type <> -1 Then
                                        'coding disini untuk menentukan primary akun adjustment sesuai dengan tipe stock out adjust
                                        rowDummy.Primary_Account_Expense = rowItemDetail.primary_loss_account
                                    End If
                                Else
                                    vTBLSetting.DefaultView.RowFilter = "Primary = " + rowDummy.Primary_Adjustment_Type.ToString
                                    row_setting = vTBLSetting.DefaultView(0).Row
                                    rowDummy.Primary_Account_Expense = row_setting.PrimaryLoss
                                End If
                            ElseIf TipeOutAdjust = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetAcquisition Then
                                rowDummy.Primary_Account_Expense = ObjAsset.GetPrimaryAkunOnProgress(PrimaryAsset)
                            ElseIf TipeOutAdjust = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetRevaluation Then
                                rowDummy.Primary_Account_Expense = ObjAsset.GetPrimaryAkunAsset(PrimaryAsset)
                            Else
                                'ini disimpan klo coding stock out custom
                                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                                rowDummy.Primary_Account_Expense = rowOutItem.Primary_Adjustment_Account
                            End If
                            rowDummy.Primary_Product = rowItemDetail.primary_code_no
                            rowDummy.Primary_Stock_Out = rowItemDetail.primary_transaksi_asal
                            rowDummy.Primary_Account_Inventory = rowOutItem.Primary_Inventory_Account
                            rowDummy.Tipe_Stock_Out = CInt(TipeOutAdjust)
                            rowDummy.Primary_Asset = PrimaryAsset

                            vTBLDummy.Rows.Add(rowDummy)
                        Else
                            'tambahkan amount
                            rowDummy = vTBLDummy.DefaultView(0).Row
                            rowDummy.Amount += rowItemDetail.amount
                        End If
                    Next
                End If
            Next 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjStockInOut = Nothing
            ObjAsset = Nothing
        End Try
    End Sub

    Private Sub UpdateDataStockOutMaterialUsing_HapusClosing(ByVal vPrimaryTransaksi As Double, ByVal vClosingPeriod As Date)
        Dim DTA_OutAdjust As New DseTransaksiMonthlyClosing2TableAdapters.TblOutAdjust_AssetTableAdapter
        Dim TBL_OutAdjust As New DseTransaksiMonthlyClosing2.TblOutAdjust_AssetDataTable
        Dim row_OutAdjust As DseTransaksiMonthlyClosing2.TblOutAdjust_AssetRow
        Dim TipeTransaksi As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
        Dim TglSampai, TglDari As Date
        Dim ObjCD As New Core_Module.ClsCekData
        Dim TipeStockOutAdjust2 As Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.NoTipe
        Dim ListPrimaryUpdateAcq As New List(Of Double)
        Dim ListPrimaryUpdateRev As New List(Of Double)
        Dim ObjAsset As New Core_Module.ClsMasterAsset

        Try
            DTA_OutAdjust.Connection.ConnectionString = Core_Module.GetConstring

            TglSampai = ObjCD.GetLasDateOfTheMonth(vClosingPeriod)
            TglDari = ObjCD.GetFirstDateofTheMonth(vClosingPeriod)
            DTA_OutAdjust.UpdateQuery_SetAmountNol(TipeTransaksi, TglSampai, TglDari)

            'update nilai book value dan acquired value
            DTA_OutAdjust.Fill(TBL_OutAdjust, TipeTransaksi, TglSampai, TglDari, TipeStockOutAdjust2)
            For Each row_OutAdjust In TBL_OutAdjust.Rows
                TipeStockOutAdjust2 = row_OutAdjust.In_Out_Adjust
                If TipeStockOutAdjust2 = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.InOutCustom Then
                    MsgBox("logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                ElseIf TipeStockOutAdjust2 = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetAcquisition Then
                    If ListPrimaryUpdateAcq.Contains(row_OutAdjust.Primary_Asset) = False Then
                        ListPrimaryUpdateAcq.Add(row_OutAdjust.Primary_Asset)
                        ObjAsset.UpdateAcquiredValue(row_OutAdjust.Primary_Asset)
                    End If
                ElseIf TipeStockOutAdjust2 = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetRevaluation Then
                    If ListPrimaryUpdateRev.Contains(row_OutAdjust.Primary_Asset) = False Then
                        ListPrimaryUpdateRev.Add(row_OutAdjust.Primary_Asset)
                        ObjAsset.UpdateBookValue(row_OutAdjust.Primary_Asset)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_OutAdjust.Dispose()
            TBL_OutAdjust.Dispose()
            ObjCD = Nothing
            ListPrimaryUpdateAcq = Nothing
            ListPrimaryUpdateRev = Nothing
            ObjAsset = Nothing
        End Try
    End Sub

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vDGVItemAdjust As DataGridView, ByVal vUserCode As String, _
    ByVal vFRM As FrmTransaksiMonthlyClosing2, _
    ByVal vTBLDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, _
    ByVal vTBLBundling As DseTransaksiMonthlyClosing2.TblBundling1DataTable, _
    ByVal vDGVAkun As DataGridView, ByVal vTBLSetting As DseTransaksiMonthlyClosing2.tbl_setting_adjustmentDataTable, _
    ByVal vTBLSettingItem As DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_tipe_adjustmentDataTable, _
    ByVal vTBLPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    ByVal vDGVCostGroup As DataGridView, ByVal vTBLCostGroup_Account As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountDataTable, _
    ByVal vTBLCostGroup_Scope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeDataTable, _
    ByVal vTBLCostGroup_CodeNo As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noDataTable, _
    ByVal vTBLCostGroup_CodeNo2 As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_no2DataTable, _
    ByVal vTblProductQty As Inventory_Module.DseLaporanStockDetail2.TblCetakDataTable, _
    Optional ByVal IsOnlySaveHeader As Boolean = False) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiAOP_SimpanBaru, vUserCode)
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
                            If IsOnlySaveHeader = False Then
                                'sudah saatnya melakukan proses simpan baru
                                retval = Me.SimpanBaru1(vdgvItem, vDGVItemAdjust, vUserCode, _
                                vFRM, vTBLDetail, vTBLBundling, _
                                vDGVAkun, vTBLSetting, vTBLSettingItem, vTBLPool, _
                                vDGVCostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, vTBLCostGroup_CodeNo, vTBLCostGroup_CodeNo2, vTblProductQty)
                            Else
                                'sudah saatnya melakukan proses simpan baru
                                retval = Me.SimpanHeader(vFRM, vUserCode)
                            End If
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

    Public Function SimpanHeader(ByVal vfrm As FrmTransaksiMonthlyClosing2, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mNoTransaksi, Me.mClosingDate, Me.mClosingPeriode, _
                            Me.mRemarks, vUserCode, Now, vUserCode, Now, Now, 0) > 0 Then
                vfrm.PrimaryTemporary = Me.GetPrimary(Me.mNoTransaksi)
                vfrm.saveTemporaryAktif = True
                Return True
            Else
                vfrm.PrimaryTemporary = -1
                vfrm.saveTemporaryAktif = False
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru1(ByVal vdgvItem As DataGridView, ByVal vDGVItemAdjust As DataGridView, ByVal vUserCode As String, _
    ByVal vFRM As FrmTransaksiMonthlyClosing2, _
    ByVal vTBLDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, _
    ByVal vTBLBundling As DseTransaksiMonthlyClosing2.TblBundling1DataTable, _
    ByVal vDGVAkun As DataGridView, ByVal vTBLSetting As DseTransaksiMonthlyClosing2.tbl_setting_adjustmentDataTable, _
    ByVal vTBLSettingItem As DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_tipe_adjustmentDataTable, _
    ByVal vTBLpool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    ByVal vDGVCostGroup As DataGridView, ByVal vTBLCostGroup_Account As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountDataTable, _
    ByVal vTBLCostGroup_Scope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeDataTable, _
    ByVal vTBLCostGroup_CodeNo As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noDataTable, _
    ByVal vTBLCostGroup_CodeNo2 As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_no2DataTable, _
    ByVal vTblProductQty As Inventory_Module.DseLaporanStockDetail2.TblCetakDataTable) As Boolean
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim ObjItem As New ClsTransaksiMonthlyClosingItem
        Dim ObjStockTaking As New Inventory_Module.clsTransaksiStockTaking
        Dim vPrimaryStockTaking As Double = -1
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowCostGroup_Acc As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim rowCostGroup_Scope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow
        Dim rowCostGroup_CodeNo As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noRow
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(True, True) = True Then
                'insert query
                If DTA.Insert(Me.mNoTransaksi, Me.mClosingDate, Me.mClosingPeriode, _
                 Me.mRemarks, vUserCode, Now, vUserCode, Now, Now, 1) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = Me.GetPrimary(Me.NoTransaksi)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'update identifier
                    For Each rowDGV In vDGVCostGroup.Rows
                        If String.IsNullOrEmpty(rowDGV.Cells(vFRM.ColCostGroup_GroupDesc.Index).Value) = False Then
                            rowDGV.Cells(vFRM.ColCostGroup_Identifier.Index).Value = Me.mPrimary.ToString + "-" + rowDGV.Cells(vFRM.ColCostGroup_Identifier.Index).Value.ToString.Substring(3)
                        End If
                    Next
                    For Each rowCostGroup_Acc In vTBLCostGroup_Account.Rows
                        rowCostGroup_Acc.identifier = Me.mPrimary.ToString + "-" + rowCostGroup_Acc.identifier.Substring(3)
                    Next
                    For Each rowCostGroup_CodeNo In vTBLCostGroup_CodeNo.Rows
                        rowCostGroup_CodeNo.identifier = Me.mPrimary.ToString + "-" + rowCostGroup_CodeNo.identifier.Substring(3)
                    Next
                    For Each rowCostGroup_Scope In vTBLCostGroup_Scope.Rows
                        rowCostGroup_Scope.identifier = Me.mPrimary.ToString + "-" + rowCostGroup_Scope.identifier.Substring(3)
                    Next

                    'simpan data item baru
                    ObjItem.Simpan(vdgvItem, vDGVItemAdjust, Me.mPrimary, vFRM, vTBLDetail, _
                    vTBLBundling, vDGVAkun, vTBLSettingItem, vTBLpool, _
                    vDGVCostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, vTBLCostGroup_CodeNo, vTBLCostGroup_CodeNo2, vTblProductQty)

                    'jika ada stock taking pada periode ini maka datanya harus di masukkan ke dalam laporan penyesuaian
                    vPrimaryStockTaking = ObjStockTaking.GetPrimary_ByPeriode(ObjCD.GetFirstDateofTheMonth(Me.mClosingPeriode), _
                    ObjCD.GetLasDateOfTheMonth(Me.mClosingPeriode))
                    If vPrimaryStockTaking <> -1 Then
                        ObjStockTaking.SimpanDataPenyesuaian_StockTaking(Me.mPrimary, vPrimaryStockTaking)
                    End If

                    'simpan data cost baru
                    If Me.GetJml > 1 Then
                        Me.SimpanDataCost(Me.mPrimary, Me.mClosingPeriode)
                    End If

                    'simpan data jurnal
                    If vDGVAkun.RowCount <> 0 Then
                        Me.TransferJurnal(Me.mPrimary, Me.mClosingPeriode, vUserCode, vDGVAkun, vFRM, -1)
                    End If

                    'simpan data amount stock out adjust material usage + update nilai asset
                    'Me.UpdateDataStockOutMaterialUsing_SimpanClosing(vdgvItem, vTBLDetail, vFRM, vTBLSetting)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mClosingPeriode)
                Return Me.SimpanBaru1(vdgvItem, vDGVItemAdjust, vUserCode, vFRM, vTBLDetail, _
                    vTBLBundling, vDGVAkun, vTBLSetting, vTBLSettingItem, vTBLpool, _
                    vDGVCostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, vTBLCostGroup_CodeNo, vTBLCostGroup_CodeNo2, vTblProductQty)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjStockTaking = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Private Sub KurangkanItemReturn1(ByVal vTBLReturn As DseTransaksiMonthlyClosing.DataTable31DataTable, _
   ByVal vTBLInvoice As DseTransaksiMonthlyClosing.TblUpdateCostDataTable)
        Dim Filter As String = ""
        Dim row1 As DseTransaksiMonthlyClosing.TblUpdateCostRow
        Dim rowReturn As DseTransaksiMonthlyClosing.DataTable31Row
        Dim i As Integer = 0

        Try
            For Each rowReturn In vTBLReturn.Rows
                Filter = "[Primary Transaksi] =" + rowReturn.Primary_Sales_Invoice.ToString + " AND [Primary Product] =" + _
                rowReturn.Primary_Product.ToString + " AND [Primary SO] =" + rowReturn.Primary_SO.ToString + _
                " AND [Primary DO] =" + rowReturn.Primary_DO.ToString
                vTBLInvoice.DefaultView.RowFilter = Filter
                i = 0
                While vTBLInvoice.DefaultView.Count > i And rowReturn.Total_Qty > 0
                    row1 = vTBLInvoice.DefaultView(i).Row
                    If row1.Total_Qty >= rowReturn.Total_Qty Then
                        row1.Total_Qty -= rowReturn.Total_Qty
                        rowReturn.Total_Qty = 0
                    Else
                        rowReturn.Total_Qty -= row1.Total_Qty
                        row1.Total_Qty = 0
                    End If

                    If row1.Total_Qty = 0 Then
                        vTBLInvoice.Rows.Remove(row1)
                    End If
                End While
            Next

            vTBLInvoice.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SimpanDataCost(ByVal vPrimaryClosing As Double, ByVal vPeriodeClosing As Date)
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales
        Dim TipeItem As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
        Dim DTAClosing As New Sales_Module.DseTransaksiSalesInvoiceTableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim TBLClosing As New Sales_Module.DseTransaksiSalesInvoice.tbltrans_purchase_closing_item_detailDataTable
        Dim rowClosing As Sales_Module.DseTransaksiSalesInvoice.tbltrans_purchase_closing_item_detailRow
        Dim DTAItem As New Sales_Module.DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Dim TBLItem As New Sales_Module.DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemDataTable
        Dim rowItem As Sales_Module.DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemRow
        Dim StartDate, EndDate As Date
        Dim ObjCD As New Core_Module.ClsCekData
        Dim TotalCost As Decimal = 0
        Dim DummyScript As String = ""
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, "")
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DefPrec As Integer = 0
        Dim DummyCost As Decimal = 0
        Dim i As Double = 0
        Dim j As Double = 0

        Try
            DefPrec = ObjCur.GetDecimal_DefaultCurrency
            DTAClosing.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DTAClosing.Connection.ConnectionString

            'load data invoice + kurangkan item return
            StartDate = ObjCD.GetFirstDateofTheMonth(vPeriodeClosing)
            EndDate = ObjCD.GetLasDateOfTheMonth(vPeriodeClosing)
            DTAItem.FillBy_TanggalSampaiDari_GoodsOnly(TBLItem, EndDate, StartDate)
            Me.KurangkanItemReturn_Cost_GoodsOnly(TBLItem, StartDate, EndDate)

            'load data closing item detail
            DTAClosing.Fill(TBLClosing, vPrimaryClosing, TipeTransaksiAsal, TipeItem)

            'untuk setiap item invoice cari cost outnya
            For Each rowItem In TBLItem.Rows
                If rowItem.Total_Qty > 0 Then
                    i = 0
                    TotalCost = 0

                    TBLClosing.DefaultView.RowFilter = "[primary transaksi asal] = " + rowItem.Primary_Transaksi.ToString + " AND " + _
                    "[primary code no] = " + rowItem.Primary_Product.ToString

                    If TBLClosing.DefaultView.Count = 0 Then
                        MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    Else
                        While rowItem.Total_Qty > 0
                            i += 1
                            j = 0

                            While j < TBLClosing.DefaultView.Count
                                rowClosing = TBLClosing.DefaultView(j).Row
                                If rowItem.Total_Qty >= rowClosing.qty Then
                                    TotalCost += rowClosing.amount
                                    rowClosing.amount = 0
                                    rowItem.Total_Qty -= rowClosing.qty
                                Else
                                    DummyCost = Math.Round((rowItem.Total_Qty * rowClosing.amount) / rowClosing.qty, DefPrec, MidpointRounding.AwayFromZero)
                                    rowClosing.qty -= rowItem.Total_Qty
                                    rowClosing.amount -= DummyCost
                                    TotalCost += DummyCost
                                    rowItem.Total_Qty = 0
                                End If

                                If rowClosing.amount = 0 Then
                                    TBLClosing.Rows.Remove(rowClosing)
                                End If

                                If rowItem.Total_Qty = 0 Then
                                    Exit While
                                End If

                                i += 1
                                If i > 10000 Then
                                    MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                                    Exit Sub
                                End If
                            End While

                            If i > 10000 Then
                                MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        End While
                    End If

                    'set script
                    If DummyScript = "" Then
                        DummyScript = "UPDATE       tbltrans_salesinvoice_item SET Cost = " + TotalCost.ToString + " WHERE        (`Primary` =" + rowItem.Primary.ToString + ");"
                    Else
                        DummyScript = DummyScript + vbCrLf + "UPDATE       tbltrans_salesinvoice_item SET Cost = " + TotalCost.ToString + " WHERE        (`Primary` =" + rowItem.Primary.ToString + ");"
                    End If

                    If DummyScript.Length > 50000 Then
                        'update menggunakan script
                        'open connection
                        Conn.Open()
                        script.Query = DummyScript
                        'execute script
                        script.Execute()
                        Conn.Close()

                        DummyScript = ""
                    End If
                End If
            Next

            'update menggunakan script - untuk batch terakhir
            'open connection
            Conn.Open()
            script.Query = DummyScript
            'execute script
            script.Execute()
            Conn.Close()

            'cek tabel item detailnya harus tidak ada
            If TBLClosing.Rows.Count > 0 Then
                MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAClosing.Dispose()
            TBLClosing.Dispose()
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
            Conn.Dispose()
            script = Nothing
            ObjCur = Nothing
        End Try
    End Sub

    Private Sub KurangkanItemReturn_Cost_GoodsOnly(ByVal vTBLItem As Sales_Module.DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemDataTable, _
    ByVal vStartPeriod As Date, ByVal vEndPeriod As Date)
        Dim Filter As String = ""
        Dim row1 As Sales_Module.DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemRow
        Dim i As Integer = 0
        Dim TBLReturn As New Sales_Module.DseLaporanSalesAnalysis.TblReturnInvoiceDataTable
        Dim DTAReturn As New Sales_Module.DseLaporanSalesAnalysisTableAdapters.TblReturnInvoiceTableAdapter
        Dim rowReturn As Sales_Module.DseLaporanSalesAnalysis.TblReturnInvoiceRow

        Try
            Sales_Module.SetConnectionString(DTAReturn)
            DTAReturn.FillBy_GoodsOnly(TBLReturn, vEndPeriod, vStartPeriod)

            For Each rowReturn In TBLReturn.Rows
                Filter = "[Primary Transaksi] =" + rowReturn.Primary_Sales_Invoice.ToString + " AND [Primary Product] =" + _
                rowReturn.Primary_Product.ToString + " AND [Primary SO] =" + rowReturn.Primary_SO.ToString + _
                " AND [Primary DO] =" + rowReturn.Primary_DO.ToString
                vTBLItem.DefaultView.RowFilter = Filter

                i = 0
                While vTBLItem.DefaultView.Count > i And rowReturn.Total_Qty > 0
                    row1 = vTBLItem.DefaultView(i).Row
                    If row1.Total_Qty >= rowReturn.Total_Qty Then
                        row1.Total_Qty -= rowReturn.Total_Qty
                        rowReturn.Total_Qty = 0
                    Else
                        rowReturn.Total_Qty -= row1.Total_Qty
                        row1.Total_Qty = 0
                    End If

                    'If row1.Total_Qty = 0 Then
                    '    vTBLItem.Rows.Remove(row1)
                    'End If
                    i += 1
                End While
            Next

            vTBLItem.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAReturn.Dispose()
            TBLReturn.Dispose()
        End Try
    End Sub

    Private Function Get_UpdateScript(ByVal vPrimaryItem As Double, ByVal vTotalCost As Decimal) As String
        Try
            Return "UPDATE       tbltrans_salesinvoice_item SET Cost = " + vTotalCost.ToString + "WHERE        (`Primary` =" + vPrimaryItem.ToString + ");"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetPrimary(ByVal VnoTrans As String) As Double
        Dim dta As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetPrimary(VnoTrans).HasValue Then
                Return dta.ScalarQuery_GetPrimary(VnoTrans).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vDGVItemAdjust As DataGridView, _
    ByVal vUserCode As String, ByVal vFRM As FrmTransaksiMonthlyClosing2, _
    ByVal vTBLDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, _
    ByVal vTBLBundling As DseTransaksiMonthlyClosing2.TblBundling1DataTable, _
    ByVal vDGVAkun As DataGridView, ByVal vTBLSetting As DseTransaksiMonthlyClosing2.tbl_setting_adjustmentDataTable, _
    ByVal vTBLSettingItem As DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_tipe_adjustmentDataTable, _
    ByVal vtblPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    ByVal vDGVCostGroup As DataGridView, ByVal vTBLCostGroup_Account As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountDataTable, _
    ByVal vTBLCostGroup_Scope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeDataTable, _
    ByVal vTBLCostGroup_CodeNo As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noDataTable, _
    ByVal vTBLCostGroup_CodeNo2 As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_no2DataTable, _
    ByVal vTblProductQty As Inventory_Module.DseLaporanStockDetail2.TblCetakDataTable) As Boolean
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim ObjItem As New ClsTransaksiMonthlyClosingItem
        Dim ObjST As New Inventory_Module.clsTransaksiStockTaking
        Dim vPrimaryStockTaking As Double = -1
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA_Jurnal As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim vPrimaryJurnalLama As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Jurnal.Connection.ConnectionString = DTA.Connection.ConnectionString

            If Me.ApakahDataBenar(False, True) = True Then
                'hitung level otorisasi yang dibutuhkan
                If DTA.Update(Me.mNoTransaksi, Me.mClosingDate, Me.mClosingPeriode, _
                 Me.mRemarks, vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, vDGVItemAdjust, Me.mPrimary, vFRM, vTBLDetail, _
                    vTBLBundling, vDGVAkun, vTBLSettingItem, vtblPool, _
                    vDGVCostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, vTBLCostGroup_CodeNo, vTBLCostGroup_CodeNo2, vTblProductQty)

                    'hapus data laporan stock taking
                    'ObjST.HapusDataPenyesuaianStockTaking(Me.mPrimary) -> fungsinya sudah ada di dalam simpan data penyesuaian
                    'jika ada stock taking pada periode ini maka datanya harus di masukkan ke dalam laporan penyesuaian
                    vPrimaryStockTaking = ObjST.GetPrimary_ByPeriode(ObjCD.GetFirstDateofTheMonth(Me.mClosingPeriode), _
                    ObjCD.GetLasDateOfTheMonth(Me.mClosingPeriode))
                    If vPrimaryStockTaking <> -1 Then
                        ObjST.SimpanDataPenyesuaian_StockTaking(Me.mPrimary, vPrimaryStockTaking)
                    End If

                    'simpan data cost baru
                    If Me.GetJml > 1 Then
                        Me.SimpanDataCost(Me.mPrimary, Me.mClosingPeriode)
                    End If

                    'simpan data jurnal
                    vPrimaryJurnalLama = DTA.ScalarQuery_GetPrimaryJurnal(Me.mPrimary).Value
                    DTA_Jurnal.DeleteQuery(vPrimaryJurnalLama)
                    If vDGVAkun.RowCount <> 0 Then
                        Me.TransferJurnal(Me.mPrimary, Me.mClosingPeriode, vUserCode, _
                        vDGVAkun, vFRM, vPrimaryJurnalLama)
                    End If

                    'hapus data amount stock out adjust material usage + update nilai asset
                    'Me.UpdateDataStockOutMaterialUsing_HapusClosing(Me.mPrimary, Me.mClosingDate)

                    'simpan data amount stock out adjust material usage + update nilai asset
                    'Me.UpdateDataStockOutMaterialUsing_SimpanClosing(dgvItem, vTBLDetail, vFRM, vTBLSetting)

                    'nilai balikan
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
            ObjST = Nothing
            ObjCD = Nothing
            DTA_Jurnal = Nothing
        End Try
    End Function

    Public Function GetNextClosingPeriode() As Date
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetMaxClosingPeriod.HasValue = True Then
                Return ObjCD.GetFirstDateofTheMonth(DTA.ScalarQuery_GetMaxClosingPeriod.Value + TimeSpan.FromDays(32))
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiTBL4Closing(ByVal TBL As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable, _
     ByVal AkhirPeriodeClosing As Date, ByVal AwalPeriodeClosing As Date, ByVal vSortByBundling As Boolean)
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item2TableAdapter
        Dim row As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2Row
        Dim TBLDummy As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable
        Dim TblDummy2 As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable
        Dim i As Double = 0

        Dim DTA1 As New DseTransaksiMonthlyClosingTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TBL1 As New DseTransaksiMonthlyClosing.tbltransaksi_stockinout_itemDataTable
        Dim row1a As DseTransaksiMonthlyClosing.tbltransaksi_stockinout_itemRow
        Dim Filter As String = ""
        Dim Dummy As String = ""
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCurr As New Core_Module.clsMasterCurrency
        Dim DefPrec As Integer = 0

        Dim DTA2 As New DseTransaksiMonthlyClosingTableAdapters.tbltransaksi_stockinout_item1TableAdapter
        Dim TBL2 As New DseTransaksiMonthlyClosing.tbltransaksi_stockinout_item1DataTable
        Dim row2 As DseTransaksiMonthlyClosing.tbltransaksi_stockinout_item1Row

        Dim row_a As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2Row

        Dim DTA_Prod As New DseTransaksiMonthlyClosingTableAdapters.tblmaster_productTableAdapter
        Dim TBL_Prod As New DseTransaksiMonthlyClosing.tblmaster_productDataTable

        Try
            DefPrec = ObjCurr.GetDecimal_DefaultCurrency
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Prod.Connection.ConnectionString = Core_Module.GetConstring

            DTA.Fill(TBL, AkhirPeriodeClosing, AwalPeriodeClosing)
            DTA.Fill(TBLDummy, AkhirPeriodeClosing, AwalPeriodeClosing)
            DTA.Fill(TblDummy2, AkhirPeriodeClosing, AwalPeriodeClosing)

            'qty in dan qty out dikurangkan dengan qty in sales return rev invoice
            DTA1.Fill(TBL1, AkhirPeriodeClosing, AwalPeriodeClosing)
            For Each row1a In TBL1.Rows
                Filter = "[Primary Product] = " + row1a.Primary_Product.ToString
                TBL.DefaultView.RowFilter = Filter
                If TBL.DefaultView.Count <> 0 Then
                    row_a = TBL.DefaultView(0).Row
                    row_a.In_Qty -= row1a.JmlIn
                    row_a.Out_Qty -= row1a.JmlIn
                End If
            Next

            'qty out dan qty in dikurangkan dengan qty out purchase return rev aop
            DTA2.Fill(TBL2, AkhirPeriodeClosing, AwalPeriodeClosing)
            For Each row2 In TBL2.Rows
                Filter = "[Primary Product] = " + row2.Primary_Product.ToString
                TBL.DefaultView.RowFilter = Filter
                If TBL.DefaultView.Count <> 0 Then
                    row_a = TBL.DefaultView(0).Row
                    row_a.In_Qty -= row2.JmlOut
                    row_a.Out_Qty -= row2.JmlOut
                    row_a.In_Amount -= CDec(ObjCD.DigitGrouping((row2.JmlOut * row2.At_Price), DefPrec))
                End If
            Next

            'kembalikan filter
            TBL.DefaultView.RowFilter = Dummy

            If vSortByBundling = True Then
                Dim TBLItemAfter As New Inventory_Module.dseTransaksiStockBundling.tbltransaksi_stockbundling_item_after1DataTable
                Dim DTAItemAfter As New Inventory_Module.dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_after1TableAdapter
                Dim TBLItemBefore As New Inventory_Module.dseTransaksiStockBundling.tbltransaksi_stockbundling_item_before4DataTable
                Dim DTAItemBefore As New Inventory_Module.dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_before4TableAdapter

                Inventory_Module.SetConnectionString(DTAItemAfter)
                Inventory_Module.SetConnectionString(DTAItemBefore)

                DTAItemBefore.Fill(TBLItemBefore, AkhirPeriodeClosing, AwalPeriodeClosing)
                DTAItemAfter.Fill(TBLItemAfter, AkhirPeriodeClosing, AwalPeriodeClosing)
                Dim row1 As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2Row
                Filter = ""

                For Each row In TBL.Rows
                    Filter = "[Primary Product] = " + row.Primary_Product.ToString
                    TBLItemBefore.DefaultView.RowFilter = Filter
                    TBLItemAfter.DefaultView.RowFilter = Filter

                    If TBLItemAfter.DefaultView.Count > 0 Then
                        row.Ada_Item_After_Bundling = 1
                    Else
                        row.Ada_Item_After_Bundling = 0
                    End If
                    If TBLItemBefore.DefaultView.Count > 0 Then
                        row.Ada_Item_Before_Bundling = 1
                    Else
                        row.Ada_Item_Before_Bundling = 0
                    End If

                    row1 = TBLDummy.Rows(i)
                    row1.Ada_Item_After_Bundling = row.Ada_Item_After_Bundling
                    row1.Ada_Item_Before_Bundling = row.Ada_Item_Before_Bundling
                    row1 = TblDummy2.Rows(i)
                    row1.Ada_Item_After_Bundling = row.Ada_Item_After_Bundling
                    row1.Ada_Item_Before_Bundling = row.Ada_Item_Before_Bundling
                    i += 1
                Next

                'buang yang ada di item after dan yang ada (diitem after dan item before)
                i = 0
                While i < TBL.Rows.Count
                    row = TBL.Rows(i)
                    If row.Ada_Item_After_Bundling = 1 And row.Ada_Item_Before_Bundling = 1 Then
                        TBL.Rows.Remove(row)
                    ElseIf row.Ada_Item_After_Bundling = 1 Then
                        TBL.Rows.Remove(row)
                    Else
                        i += 1
                    End If
                End While
                'pd tbldummy sisakan, yang ada di item buffer dan item after
                i = 0
                While i < TBLDummy.Rows.Count
                    row = TBLDummy.Rows(i)
                    If row.Ada_Item_After_Bundling = 1 And row.Ada_Item_Before_Bundling = 1 Then
                        i += 1
                    Else
                        TBLDummy.Rows.Remove(row)
                    End If
                End While
                'pd tbldummy2 sisakan, yang ada di item after saja.
                i = 0
                While i < TblDummy2.Rows.Count
                    row = TblDummy2.Rows(i)
                    If row.Ada_Item_After_Bundling = 1 And row.Ada_Item_Before_Bundling = 0 Then
                        i += 1
                    Else
                        TblDummy2.Rows.Remove(row)
                    End If
                End While
                'urutkan item2 yang ada di item after dan item before, masuk duluan
                For Each row In TBLDummy.Rows
                    row1 = TBL.Newtbltrans_purchaseclosing_item2Row

                    row1.Ada_Item_After_Bundling = row.Ada_Item_After_Bundling
                    row1.Ada_Item_Before_Bundling = row.Ada_Item_Before_Bundling
                    row1.Begin_Amount = row.Begin_Amount
                    row1.Begin_Qty = row.Begin_Qty
                    row1.Description = row.Description
                    row1.End_Qty = row.End_Qty
                    row1.In_Amount = row.In_Amount
                    row1.In_Qty = row.In_Qty
                    row1.Out_Qty = row.Out_Qty
                    row1.Primary_Product = row.Primary_Product
                    row1.Primary_Unit = row.Primary_Unit
                    row1.Product_Code = row.Product_Code
                    row1.Unit_Code = row.Unit_Code
                    row1._Decimal = row._Decimal

                    TBL.Rows.Add(row1)
                Next

                'urutan kan, pastikan yang ada di item after(saja) transaksi bundling, menjadi yang paling bawah.
                For Each row In TblDummy2.Rows
                    row1 = TBL.Newtbltrans_purchaseclosing_item2Row

                    row1.Ada_Item_After_Bundling = row.Ada_Item_After_Bundling
                    row1.Ada_Item_Before_Bundling = row.Ada_Item_Before_Bundling
                    row1.Begin_Amount = row.Begin_Amount
                    row1.Begin_Qty = row.Begin_Qty
                    row1.Description = row.Description
                    row1.End_Qty = row.End_Qty
                    row1.In_Amount = row.In_Amount
                    row1.In_Qty = row.In_Qty
                    row1.Out_Qty = row.Out_Qty
                    row1.Primary_Product = row.Primary_Product
                    row1.Primary_Unit = row.Primary_Unit
                    row1.Product_Code = row.Product_Code
                    row1.Unit_Code = row.Unit_Code
                    row1._Decimal = row._Decimal

                    TBL.Rows.Add(row1)
                Next

                TBLItemAfter.Dispose()
                DTAItemAfter.Dispose()
                TBLItemBefore.Dispose()
                DTAItemBefore.Dispose()
            End If

            For Each row In TBL.Rows
                row.In_Qty = Math.Round(row.In_Qty, row._Decimal)
                row.Begin_Qty = Math.Round(row.Begin_Qty, row._Decimal)
                row.Out_Qty = Math.Round(row.Out_Qty, row._Decimal)
                row.End_Qty = Math.Round(row.End_Qty, row._Decimal)
            Next

            'buang baris yang jasa
            DTA_Prod.Fill(TBL_Prod)
            i = 0
            While i <= TBL.Rows.Count - 1
                row = TBL.Rows(i)
                Filter = "Primary = " + row.Primary_Product.ToString
                TBL_Prod.DefaultView.RowFilter = Filter
                If TBL_Prod.DefaultView.Count = 0 Then
                    TBL.Rows.Remove(row)
                Else
                    i += 1
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBLDummy.Dispose()
            TblDummy2.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
            DTA2.Dispose()
            TBL2.Dispose()
            ObjCurr = Nothing
            ObjCD = Nothing
            DTA_Prod = Nothing
            TBL_Prod = Nothing
        End Try
    End Sub

    Public Function GetLastClosingPeriod() As Date
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.DataTable26TableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing.DataTable26DataTable
        Dim row As DseTransaksiMonthlyClosing.DataTable26Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(TBL) > 0 Then
                row = TBL.Rows(0)
                If row.IsMaxClosingPeriodNull = True Then
                    Return #1/1/1950#
                End If
                Return row.MaxClosingPeriod
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetClosingPeriod(ByVal vPrimaryClosing As Double) As Date
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetClosingPeriod(vPrimaryClosing).HasValue = True Then
                Return DTA.ScalarQuery_GetClosingPeriod(vPrimaryClosing).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub TransferJurnal(ByVal vPrimaryClosing As Double, ByVal vClosingPeriod As Date, ByVal vUserCode As String, _
    ByVal vDGVAkun As DataGridView, ByVal vFrm As FrmTransaksiMonthlyClosing2, ByVal vPrimaryJurnalLama As Double)
        Dim DTA_Jurnal As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim DTA_JurnalItem As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim TipeJurnal As AccModule.ClsTransaksiJournal.EnumTipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
        Dim rowDGV As DataGridViewRow
        Dim MonthStr As String = ""
        Dim PrimaryAkun As Double = -1
        Dim ObjJurnal As New AccModule.ClsTransaksiJournal
        Dim Dummy As Integer = 1
        Dim MonthStrLama As String = ""
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrimaryJurnal As Double = -1
        Dim DTA_Closing As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter

        Try
            DTA_Jurnal.Connection.ConnectionString = Core_Module.GetConstring
            DTA_JurnalItem.Connection.ConnectionString = DTA_Jurnal.Connection.ConnectionString
            DTA_Closing.Connection.ConnectionString = DTA_Jurnal.Connection.ConnectionString

            'hapuskan data jurnal lama
            DTA_Jurnal.DeleteQuery(vPrimaryJurnalLama)
            'buat jurnal baru
            MonthStr = Format(vClosingPeriod, "MMM yyyy")
            MonthStrLama = MonthStr
            While ObjJurnal.ApakahSudahAda(MonthStr) = True
                MonthStr = MonthStrLama + Dummy.ToString
                Dummy += 1
            End While
            'data header

            DTA_Jurnal.Insert(("COGS-" + MonthStr).ToUpper, ObjCD.GetLasDateOfTheMonth(vClosingPeriod), _
            0, TipeJurnal, vPrimaryClosing, -1, "COGS: " + Format(vClosingPeriod, "MMM yyyy") _
            , vUserCode, Now, vUserCode, Now, Now)

            'data item
            PrimaryJurnal = DTA_Jurnal.ScalarQuery_GetPrimary2(("COGS-" + MonthStr).ToUpper).Value

            'simpan data jurnal
            For Each rowDGV In vDGVAkun.Rows
                With rowDGV
                    DTA_JurnalItem.InsertQuery(PrimaryJurnal, .Cells(vFrm.ColJournal_PrimaryAkun.Index).Value, _
                    .Cells(vFrm.ColJournal_Debet.Index).Value, .Cells(vFrm.ColJournal_Kredit.Index).Value, _
                    .Cells(vFrm.ColJournal_Remarks.Index).Value, .Cells(vFrm.ColJournal_Debet_OriginalCurr.Index).Value, _
                    .Cells(vFrm.ColJournal_Credit_OriginalCurrency.Index).Value, _
                    .Cells(vFrm.ColJournal_ExchangeRate.Index).Value)
                End With
            Next

            'update data jurnal ke dalam tabel header closing
            DTA_Closing.UpdateQuery_SetPrimaryJurnalCOGS(PrimaryJurnal, vPrimaryClosing)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Jurnal.Dispose()
            DTA_JurnalItem.Dispose()
            ObjCD = Nothing
            ObjJurnal = Nothing
            DTA_Closing.Dispose()
        End Try
    End Sub

    Public Function ApakahAdaCyclicBundling(ByVal vPeriod As Date) As Boolean
        Dim DTAAft As New DseTransaksiMonthlyClosing2TableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim TBLAft As New DseTransaksiMonthlyClosing2.tbltransaksi_stockbundling_item_afterDataTable
        Dim rowAft, rowAft1 As DseTransaksiMonthlyClosing2.tbltransaksi_stockbundling_item_afterRow
        Dim DTABefore As New DseTransaksiMonthlyClosing2TableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim TBLBefore As New DseTransaksiMonthlyClosing2.tbltransaksi_stockbundling_item_beforeDataTable
        Dim rowBefore As DseTransaksiMonthlyClosing2.tbltransaksi_stockbundling_item_beforeRow
        Dim ObjBundling As New Inventory_Module.ClsTransaksiStockBundling
        Dim DTASC As New DseTransaksiMonthlyClosing2TableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim TBLSC As New DseTransaksiMonthlyClosing2.tbltransaksi_stock_cardDataTable
        Dim rowSC As DseTransaksiMonthlyClosing2.tbltransaksi_stock_cardRow
        Dim TipeTransaksiSC As Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal = Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Bundling
        Dim IndAft As Integer
        Dim TBLCek As New DseTransaksiMonthlyClosing2.tbltransaksi_stockbundling_item_afterDataTable
        Dim QtyAwal As Decimal = 0
        Dim Dummy As Nullable(Of Decimal)
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTAAft.Connection.ConnectionString = Core_Module.GetConstring
            DTABefore.Connection.ConnectionString = Core_Module.GetConstring
            DTASC.Connection.ConnectionString = DTAAft.Connection.ConnectionString

            DTAAft.Fill(TBLAft, ObjCD.GetLasDateOfTheMonth(vPeriod), ObjCD.GetFirstDateofTheMonth(vPeriod))
            DTABefore.Fill(TBLBefore, ObjCD.GetLasDateOfTheMonth(vPeriod), ObjCD.GetFirstDateofTheMonth(vPeriod))

            'cari pasangan primary bundling dan primary code no yang ada di item after dan item before
            IndAft = 0
            While IndAft <= TBLAft.Rows.Count - 1
                rowAft = TBLAft.Rows(IndAft)
                For Each rowBefore In TBLBefore.Rows
                    If rowAft.Primary_Transaksi = rowBefore.Primary_Transaksi Then
                        If rowAft.Primary_Product = rowBefore.Primary_Product Then
                            rowAft1 = TBLCek.Newtbltransaksi_stockbundling_item_afterRow
                            rowAft1.Primary_Product = rowAft.Primary_Product
                            rowAft1.Qty_After = rowAft.Qty_After
                            rowAft1.Primary_Transaksi = rowAft.Primary_Transaksi
                            rowAft1.Qty_Before = rowBefore.Qty
                            rowAft1.Siklik = 0
                            rowAft1.No_Transaksi = rowAft.No_Transaksi
                            rowAft1.Product_Code = rowAft.Product_Code
                            rowAft1.Description = rowAft.Description
                            rowAft1.Tanggal_Transaksi = rowAft.Tanggal_Transaksi
                            TBLCek.Rows.Add(rowAft1)
                        End If
                    ElseIf rowAft.Primary_Transaksi < rowBefore.Primary_Transaksi Then
                        Exit For
                    End If
                Next
                If IndAft + 1 <= TBLAft.Rows.Count - 1 Then
                    rowAft1 = TBLAft.Rows(IndAft + 1)
                    If rowAft1.Primary_Transaksi <> rowAft.Primary_Transaksi Then
                        'baris terakhir untuk primary transaksi tsb pada tabel aft
                        'buang baris pada tabel before, supaya proses pencariannya lebih cepat
                        rowBefore = TBLBefore.Rows(0)
                        While rowBefore.Primary_Transaksi = rowAft.Primary_Transaksi
                            TBLBefore.Rows.RemoveAt(0)
                            If TBLBefore.Rows.Count > 0 Then
                                rowBefore = TBLBefore.Rows(0)
                            Else
                                Exit While
                            End If
                        End While
                    End If
                End If
                IndAft += 1
            End While

            'untuk semua pasangan yang bersangkutan, cari yang ketika mau dihitung, 
            'ketika qty dikurangi qty out, qtynya tidak cukup
            For Each rowAft In TBLCek.Rows
                Dummy = DTASC.ScalarQuery_GetJml(rowAft.Primary_Product, (rowAft.Tanggal_Transaksi - TimeSpan.FromDays(1)).Date)
                If Dummy.HasValue = True Then
                    QtyAwal = Dummy.Value
                Else
                    QtyAwal = 0
                End If
                DTASC.Fill(TBLSC, rowAft.Primary_Product, rowAft.Tanggal_Transaksi)
                For Each rowSC In TBLSC.Rows
                    If rowSC.Type_Transaksi_Asal <> Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Bundling Then
                        QtyAwal += (rowSC._In - rowSC.Out)
                    ElseIf rowSC.Primary_Transaksi_Asal < rowAft.Primary_Transaksi Then
                        QtyAwal += (rowSC._In - rowSC.Out)
                    End If
                Next

                QtyAwal -= rowAft.Qty_Before
                If QtyAwal <= 0 Then
                    rowAft.Siklik = 1
                End If
            Next

            'jika ada yang siklik maka tampilkan
            For Each rowAft In TBLCek.Rows
                If rowAft.Siklik = 1 Then
                    Dim ij As New FrmTransaksiMonthlyClosing2_CyclicBundling
                    For Each rowAft1 In TBLCek.Rows
                        If rowAft1.Siklik = 1 Then
                            ij.dgvItem.Rows.Add()
                            rowDGV = ij.dgvItem.Rows(ij.dgvItem.Rows.Count - 1)
                            rowDGV.Cells(ij.ColBundlingNo.Index).Value = rowAft1.No_Transaksi
                            rowDGV.Cells(ij.ColCodeNo.Index).Value = rowAft1.Product_Code
                            rowDGV.Cells(ij.ColDescription.Index).Value = rowAft1.Description
                            rowDGV.Cells(ij.ColNo.Index).Value = rowDGV.Index + 1
                        End If
                    Next
                    ij.ShowDialog()
                    ij.Dispose()
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAAft.Dispose()
            TBLAft.Dispose()
            DTABefore.Dispose()
            TBLBefore.Dispose()
            ObjBundling = Nothing
            DTASC.Dispose()
            TBLSC.Dispose()
            TBLCek.Dispose()
            ObjCD = Nothing
        End Try
    End Function
End Class
