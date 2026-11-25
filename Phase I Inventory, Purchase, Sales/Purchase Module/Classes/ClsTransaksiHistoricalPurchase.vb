Option Explicit On
Option Strict Off

Public Class ClsTransaksiHistoricalPurchase
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double
    Private mNoTransaksi As String
    Private mUrutan As Int64
    Private mTglTransaksi As Date
    Private mStartPeriod As Date
    Private mEndPeriod As Date
    Private mPrimaryCurrency As Double
    Private mAmount As Double
    Private mRemarks As String
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
    Public Property NoTransaksi() As String
        Get
            Return mNoTransaksi
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property
    Public Property Urutan() As Int16
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Int16)
            mUrutan = value
        End Set
    End Property
    Public Property TglTransaksi() As Date
        Get
            Return mTglTransaksi
        End Get
        Set(ByVal value As Date)
            mTglTransaksi = value
        End Set
    End Property
    Public Property StartPeriod() As Date
        Get
            Return mStartPeriod
        End Get
        Set(ByVal value As Date)
            mStartPeriod = value
        End Set
    End Property
    Public Property EndPeriod() As Date
        Get
            Return mEndPeriod
        End Get
        Set(ByVal value As Date)
            mEndPeriod = value
        End Set
    End Property
    Public Property PrimaryCurrency() As Double
        Get
            Return mPrimaryCurrency
        End Get
        Set(ByVal value As Double)
            mPrimaryCurrency = value
        End Set
    End Property
    Public Property Amount() As Double
        Get
            Return mAmount
        End Get
        Set(ByVal value As Double)
            mAmount = value
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
    Public Property TimeStamp() As Date
        Get
            Return mTimeStamp
        End Get
        Set(ByVal value As Date)
            mTimeStamp = value
        End Set
    End Property

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = -1
            mNoTransaksi = ""
            mUrutan = 0
            mTglTransaksi = Today
            mStartPeriod = Today
            mEndPeriod = Today
            mPrimaryCurrency = -1
            mAmount = 0
            mRemarks = ""
            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimeStamp = Today
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchaseTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If String.IsNullOrEmpty(DTA.ScalarQuery_getNoMax(TanggalStart, TanggalEnd)) = False Then
                NoUrut = CStr(CInt(CStr(DTA.ScalarQuery_getNoMax(TanggalStart, TanggalEnd)).Substring(0, 5)) + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If
            'pastikan no urut terdiri atas 5 huruf
            While NoUrut.Length < 5
                NoUrut = "0" + NoUrut
            End While

            'bulan 
            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While
            Return NoUrut + "/HST/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vNoTransaksi As String, _
   ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
   Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchaseTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                If DTA.ScalarQuery_getJmlByNoTransaksiLike(vNoTransaksi + "%").HasValue = True Then
                    Return DTA.ScalarQuery_getJmlByNoTransaksiLike(vNoTransaksi + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                If DTA.ScalarQuery_getJmlByTglDariSampai(vTanggalDari, vTanggalSampai).HasValue = True Then
                    Return DTA.ScalarQuery_getJmlByTglDariSampai(vTanggalDari, vTanggalSampai).Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                If DTA.ScalarQuery_GetJmlByProductCode(vNoTransaksi + "%").HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByProductCode(vNoTransaksi + "%").Value
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

    Public Function IsiDGV(ByVal vNoTransaksi As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiHistoricalPurchase.tbltrans_historicalpurchase1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchase1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLike(TBL, vNoTransaksi + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_tgldarisampai(TBL, vTanggalDari, vTanggalSampai)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLike(TBL, vNoTransaksi + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                Dim ij As New DataGridViewCellStyle
                Dim ik As New DataGridViewCellStyle

                ij.Format = "MMM yyyy"
                ik.Format = "dd MMM yyyy"

                DGV.Columns(TBL.PrimaryColumn.ColumnName).Visible = False
                DGV.Columns(TBL.No_TransaksiColumn.ColumnName).HeaderText = "Transaction No"
                DGV.Columns(TBL.No_TransaksiColumn.ColumnName).Frozen = True
                DGV.Columns(TBL.Tanggal_TransaksiColumn.ColumnName).HeaderText = "Transaction Date"
                DGV.Columns(TBL.Tanggal_TransaksiColumn.ColumnName).DefaultCellStyle = ik
                DGV.Columns(TBL.Start_PeriodColumn.ColumnName).DefaultCellStyle = ij
                DGV.Columns(TBL.End_PeriodeColumn.ColumnName).DefaultCellStyle = ij
                DGV.Columns(TBL.Currency_CodeColumn.ColumnName).HeaderText = "Currency"
                DGV.Columns(TBL.AmountColumn.ColumnName).Visible = False
                DGV.Columns(TBL.Currency_CodeColumn.ColumnName).Visible = False
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

    Public Function SimpanBaru(ByVal dgv As DataGridView, ByVal objItem As ClsTransaksiHostoricalPurchaseItem, ByVal vUser As String) As Boolean
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchaseTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.NoTransaksi, Me.Urutan, Me.TglTransaksi, Me.StartPeriod, Me.EndPeriod, Me.PrimaryCurrency, Me.Amount, Me.Remarks, vUser, Now, vUser, Now, Now) > 0 Then
                Me.Primary = DTA.ScalarQuery_getPrimary(Me.NoTransaksi)
                Me.TimeStamp = DTA.ScalarQuery_getTimeStamp(Me.Primary)
                Me.mCreatedDate = Me.TimeStamp
                Me.ModifiedDate = Me.TimeStamp

                Return objItem.Simpan(dgv, Me.Primary)
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahAdaTglOverSlap(ByVal ListString As List(Of String), ByVal StartDate As Date, ByVal EndDate As Date, ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchaseTableAdapter
        Dim TBL As DseTransaksiHistoricalPurchase.tbltrans_historicalpurchaseDataTable
        Dim row As DseTransaksiHistoricalPurchase.tbltrans_historicalpurchaseRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL = DTA.GetData_StartDateEndDate_By_TglDariSampai(StartDate, EndDate, vPrimary)
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    ListString.Add(row.No_Transaksi + " with start date: " + Format(row.Start_Period, "dd MMM yyyy") + _
                    " and end date: " + Format(row.End_Periode, "dd MMM yyyy"))
                Next
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal dgv As DataGridView, ByVal objItem As ClsTransaksiHostoricalPurchaseItem, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchaseTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If DTA.Update(Me.NoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.TglTransaksi, Me.StartPeriod, Me.EndPeriod, Me.PrimaryCurrency, Me.Amount, Me.Remarks, vUserCode, Now, Now, Me.Primary, Me.TimeStamp) > 0 Then
                Me.mUrutan = Me.mNoTransaksi.Substring(0, 5)
                Me.mTimeStamp = DTA.ScalarQuery_getTimeStamp(CStr(Me.mPrimary))
                Me.mModifiedDate = Me.mTimeStamp
                Me.mModifiedBy = vUserCode

                If objItem.Hapus(Me.Primary) = True Then
                    Return objItem.Simpan(dgv, Me.Primary)
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
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal objItem As ClsTransaksiHostoricalPurchaseItem) As Boolean
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchaseTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Delete(Me.Primary) > 0 Then
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

    Public Function CariDetail(ByVal dgv As DataGridView, ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchaseTableAdapter
        Dim DTAItem As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchase_item1TableAdapter
        Dim TBL As DseTransaksiHistoricalPurchase.tbltrans_historicalpurchaseDataTable
        Dim TBLItem As DseTransaksiHistoricalPurchase.tbltrans_historicalpurchase_item1DataTable
        Dim Row As DseTransaksiHistoricalPurchase.tbltrans_historicalpurchaseRow
        Dim RowItem As DseTransaksiHistoricalPurchase.tbltrans_historicalpurchase_item1Row
        Dim CD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            dgv.Rows.Clear()
            TBL = DTA.GetDataBy_Primary(vPrimary)
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If TBL.Rows.Count > 0 Then
                For Each Row In TBL.Rows
                    Me.Amount = Row.Amount
                    Me.CreatedBy = Row.Created_By
                    Me.CreatedDate = Row.Created_Date
                    Me.EndPeriod = Row.End_Periode
                    Me.ModifiedBy = Row.Modified_By
                    Me.ModifiedDate = Row.Modified_Date
                    Me.NoTransaksi = Row.No_Transaksi
                    Me.Primary = Row.Primary
                    Me.PrimaryCurrency = Row.Primary_Currency
                    Me.Remarks = Row.Remarks
                    Me.StartPeriod = Row.Start_Period
                    Me.TglTransaksi = Row.Tanggal_Transaksi
                    Me.TimeStamp = Row.Timestamp
                    Me.Urutan = Row.Urutan
                    TBLItem = DTAItem.GetDataBy_PrimaryTransaksi(vPrimary)
                    If TBLItem.Rows.Count > 0 Then
                        For Each RowItem In TBLItem.Rows
                            dgv.Rows.Add("", RowItem.Product_Code, RowItem.Description, RowItem.Qty, RowItem.Unit_Code, CD.DigitGrouping(CStr(RowItem.Amount)), RowItem.Remarks, RowItem.Primary_Product, RowItem.Primary_Unit)
                        Next
                    End If
                Next
            End If

            CD.SetNumber(dgv)
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTAItem.Dispose()
            DTA.Dispose()
            CD = Nothing
        End Try
    End Function

    Public Function GetPrimary(ByVal RefDate As Date) As Double
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchaseTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary2(RefDate).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary2(RefDate).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchaseTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_getPrimary(vNoTransaksi).HasValue Then
                Return DTA.ScalarQuery_getPrimary(vNoTransaksi).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTglTransaksi(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchaseTableAdapter
        Try
            If DTA.ScalarQuery_getTglTransaksiByPrimary(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_getTglTransaksiByPrimary(vPrimary).Value
            Else
                'Ini sengaja di setting tgl #1/1/1888#,jangan di ganti now or today
                Return #1/1/1888#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub GetStartNEndDate(ByVal vPrimaryHP As Double, ByRef vStartDate As Date, ByRef vEndDate As Date)
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchaseTableAdapter
        Dim TBL As New DseTransaksiHistoricalPurchase.tbltrans_historicalpurchaseDataTable

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_Primary(TBL, vPrimaryHP) > 0 Then
                Dim row As DseTransaksiHistoricalPurchase.tbltrans_historicalpurchaseRow = TBL.Rows(0)
                vStartDate = row.Start_Period
                vEndDate = row.End_Periode
            Else
                vStartDate = CDate("1/1/1950")
                vEndDate = CDate("1/1/1950")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchaseTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_getTimeStamp(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_getTimeStamp(vPrimary).Value
            Else
                Return Now
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchaseTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_getNoTransaksi(vPrimary)) = False Then
                Return DTA.ScalarQuery_getNoTransaksi(vPrimary)
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
End Class
