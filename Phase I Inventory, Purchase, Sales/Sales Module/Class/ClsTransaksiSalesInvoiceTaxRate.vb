Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesInvoiceTaxRate
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mKMKNumber As String = ""
    Dim mKMK_Date As Date = Today
    Dim mStartDate As Date = Today
    Dim mEndDate As Date = Today
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public DefaultPrec As Integer = 2

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property StartDate() As Date
        Get
            Return Me.mStartDate
        End Get
        Set(ByVal value As Date)
            mStartDate = value
        End Set
    End Property

    Public Property EndDate() As Date
        Get
            Return Me.mEndDate
        End Get
        Set(ByVal value As Date)
            mEndDate = value
        End Set
    End Property

    Public Property KMKDate() As Date
        Get
            Return Me.mKMK_Date
        End Get
        Set(ByVal value As Date)
            mKMK_Date = value
        End Set
    End Property

    Public Property KMKNumber() As String
        Get
            Return Me.mKMKNumber
        End Get
        Set(ByVal value As String)
            mKMKNumber = value
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

    Public Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Dim PrimaryKMK As Double = -1

        Try
            'periodenya konflik
            If Me.ApakahAdaYgKonflik(Me.mPrimary, Me.mStartDate, Me.mEndDate, True) = True Then
                Return False
            End If

            'jika tanggal sampai lebih kecil atau sama dengan tanggal dari
            If Me.mEndDate <= Me.mStartDate Then
                MsgBox("End date must be large than start date", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'kmk number harus diisikan
            If Me.mKMKNumber = "" Then
                MsgBox("Please specify KMK number.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'jika kmk number konflik
            If vBaru = True Then
                If Me.GetPrimary(Me.mKMKNumber) <> -1 Then
                    MsgBox("Data duplication in KMK number, please change KMK number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                If Me.GetPrimary(Me.mKMKNumber) <> Me.mPrimary Then
                    MsgBox("Data duplication in KMK number, please change KMK number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahDataTerakhir(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrateTableAdapter
        Dim MaxEndDate As Date

        Try
            If DTA.ScalarQuery_GetMaxEndDate().HasValue = True Then
                MaxEndDate = DTA.ScalarQuery_GetMaxEndDate().Value

                If DTA.ScalarQuery_GetPrimaryByEndDate(MaxEndDate).HasValue = True Then
                    If DTA.ScalarQuery_GetPrimaryByEndDate(MaxEndDate).Value = vPrimary Then
                        Return True
                    Else
                        Return False
                    End If
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

    Public Function ApakahAdaYgKonflik(ByVal vPrimaryTrans As Double, ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrateTableAdapter
        Dim TBL As New DseTransaksiSalesInvoice_TaxRate.tbltrans_salesinvoice_taxrateDataTable
        Dim row As DseTransaksiSalesInvoice_TaxRate.tbltrans_salesinvoice_taxrateRow

        Try
            DTA.Fill(TBL)
            For Each row In TBL.Rows
                If row.primary = vPrimaryTrans Then
                    TBL.Rows.Remove(row)

                    Exit For
                End If
            Next

            'yang dimaksud adalah konflik periodenya.
            For Each row In TBL.Rows
                '1. tanggal sampai berada diantara tgl dari sampai yg lain
                If vEndDate <= row.end_date And vEndDate >= row.start_date Then
                    If vShowMsg = True Then
                        MsgBox("This transaction period conficts with: " + row.kmk_number + ". Please change period", _
                        MsgBoxStyle.Information, "Information")
                    End If

                    Return True
                End If
                '2. tbl dari berada diantara tgl dari sampai yg lain
                If vStartDate <= row.end_date And vStartDate >= row.start_date Then
                    If vShowMsg = True Then
                        MsgBox("This transaction period conficts with: " + row.kmk_number + ". Please change period", _
                        MsgBoxStyle.Information, "Information")
                    End If

                    Return True
                End If
                '3. tgl dari kurang dari tanggal dari yg lain dan pada saat yg bersamaa tanggal sampai lebih dari tgl sampai yg lain 
                If vStartDate <= row.start_date And vEndDate >= row.end_date Then
                    If vShowMsg = True Then
                        MsgBox("This transaction period conficts with: " + row.kmk_number + ". Please change period", _
                        MsgBoxStyle.Information, "Information")
                    End If

                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function SimpanBaru(ByVal vUserCode As String, ByVal vDGV As DataGridView, ByVal i As FrmTransaksiSalesInvoiceSettingPajak) As Boolean
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrateTableAdapter
        Try
            If Me.ApakahDataBenar(True) = True Then
                If DTA.Insert(Me.mKMKNumber, Me.mKMK_Date, Me.mStartDate, Me.mEndDate, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mKMKNumber).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(CDec(Me.mPrimary)).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan item
                    Me.SimpanItem(vDGV, i, Me.mPrimary)

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

    Private Sub SimpanItem(ByVal vDGV As DataGridView, ByVal i As FrmTransaksiSalesInvoiceSettingPajak, ByVal vPrimaryTrans As Double)
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrate_itemTableAdapter
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In vDGV.Rows
                DTA.Insert(vPrimaryTrans, rowDGV.Cells(i.ColRateKurs.Index).Value, rowDGV.Cells(i.ColRatePrimaryCurrency.Index).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria, ByVal vDate As Date) As Double
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrate1TableAdapter

        Try
            If vSearchBy = EnumSearchCriteria.ByNoTransaksi Then
                If DTA.ScalarQuery_GetJmlByKMKNumberLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByKMKNumberLIKE(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByTanggalTransaksi Then
                If DTA.ScalarQuery_GetJml_ByTanggal(vDate).HasValue Then
                    Return DTA.ScalarQuery_GetJml_ByTanggal(vDate).Value
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

    Public Function SimpanPerubahan(ByVal vUserCode As String, ByVal vDGV As DataGridView, ByVal i As FrmTransaksiSalesInvoiceSettingPajak) As Boolean
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrateTableAdapter
        Dim DTAItem As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrate_itemTableAdapter

        Try
            If Me.ApakahDataBenar(False) = True Then
                If DTA.Update(Me.mKMKNumber, Me.mKMK_Date, Me.mStartDate, Me.mEndDate, vUserCode, Now, CDec(Me.mPrimary), Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(CDec(Me.mPrimary)).Value
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus item lama
                    DTAItem.Delete(Me.mPrimary)
                    'simpan item
                    Me.SimpanItem(vDGV, i, Me.mPrimary)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAItem.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrateTableAdapter

        Try
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

    Public Function GetKMKNumber(ByVal vTgl As Date) As String
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrateTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetKMKNumber2(vTgl)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetKMKNumber2(vTgl)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetKMKDate(ByVal vTgl As Date) As Date
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrateTableAdapter

        Try
            If DTA.ScalarQuery_GetKMKDate(vTgl).HasValue = False Then
                Return #1/1/1950#
            Else
                Return DTA.ScalarQuery_GetKMKDate(vTgl)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTaxRate(ByVal vTanggal As Date, ByVal vPrimaryCurrency As Double) As Decimal
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrate_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetRate(vPrimaryCurrency, vTanggal).HasValue = True Then
                Return DTA.ScalarQuery_GetRate(vPrimaryCurrency, vTanggal).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vTgl As Date) As Boolean
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrateTableAdapter

        Try
            If DTA.ScalarQuery_GetJml2(vTgl).HasValue = True Then
                If DTA.ScalarQuery_GetJml2(vTgl).Value > 0 Then
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

    Public Function ApakahSudahAda(ByVal vKMKNumber As String) As Boolean
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrateTableAdapter

        Try
            If DTA.ScalarQuery_GetJml(vKMKNumber).HasValue Then
                If DTA.ScalarQuery_GetJml(vKMKNumber).Value > 0 Then
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
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrateTableAdapter
        Dim TBL As New DseTransaksiSalesInvoice_TaxRate.tbltrans_salesinvoice_taxrateDataTable
        Dim row As DseTransaksiSalesInvoice_TaxRate.tbltrans_salesinvoice_taxrateRow
        Dim TglDr, TglSampai As Date
        Dim DTA_Inv As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim TBL_Inv As New DseTransaksiSalesInvoice.tbltrans_salesinvoice1DataTable
        Dim row_Inv As DseTransaksiSalesInvoice.tbltrans_salesinvoice1Row
        Dim ListNo As New List(Of String)
        Dim Dummy, NoTrans As String

        Try
            '1. jika ada invoice non IDR pada periode tgl dr sampai maka data tidak boleh dihapuskan
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                TglDr = row.start_date
                TglSampai = row.end_date

                DTA_Inv.FillBy_NonIDRInvoice(TBL_Inv, TglSampai, TglDr)
                If TBL_Inv.Rows.Count <> 0 Then
                    For Each row_Inv In TBL_Inv.Rows
                        If ListNo.Count < 5 Then
                            ListNo.Add(row_Inv.No_Transaksi)
                        Else
                            Exit For
                        End If
                    Next

                    Dummy = ""
                    For Each NoTrans In ListNo
                        If Dummy = "" Then
                            Dummy = NoTrans
                        Else
                            Dummy = Dummy + ", " + NoTrans
                        End If
                    Next

                    MsgBox("The following invoice number use this data: " + Dummy + ", you cannot delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA_Inv.Dispose()
            TBL_Inv.Dispose()
            ListNo = Nothing
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiSalesInvoice_TaxRate.tbltrans_salesinvoice_taxrate1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal vTanggal As Date) As Boolean
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrate1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_KmkNumberLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_Date(TBL, vTanggal)
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
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

    Public Sub IsikanDGVKosong(ByVal vDGV As DataGridView, ByVal i As FrmTransaksiSalesInvoiceSettingPajak)
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tblmaster_currencyTableAdapter
        Dim TBL As New DseTransaksiSalesInvoice_TaxRate.tblmaster_currencyDataTable
        Dim row As DseTransaksiSalesInvoice_TaxRate.tblmaster_currencyRow
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Fill(TBL)

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGV.Rows.Clear()

            For Each row In TBL.Rows
                If row.Primary <> 1 Then
                    vDGV.Rows.Add()
                    rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)

                    With rowDGV
                        .Cells(i.ColRatePrimaryCurrency.Index).Value = row.Primary
                        .Cells(i.ColRateNo.Index).Value = vDGV.Rows.Count
                        .Cells(i.ColRateKurs.Index).Value = ObjCD.DigitGrouping(0, Me.DefaultPrec)
                        .Cells(i.ColRateCurrency.Index).Value = row.Currency_Code
                    End With
                End If
            Next

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
        End Try
    End Sub

    Public Function CariDetail(ByVal Primary As Double, ByVal vDGV As DataGridView, ByVal i As FrmTransaksiSalesInvoiceSettingPajak) As Boolean
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrateTableAdapter
        Dim TBL As New DseTransaksiSalesInvoice_TaxRate.tbltrans_salesinvoice_taxrateDataTable
        Dim row As DseTransaksiSalesInvoice_TaxRate.tbltrans_salesinvoice_taxrateRow

        Dim DTAItem As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrate_itemTableAdapter
        Dim TBLItem As New DseTransaksiSalesInvoice_TaxRate.tbltrans_salesinvoice_taxrate_itemDataTable
        Dim rowItem As DseTransaksiSalesInvoice_TaxRate.tbltrans_salesinvoice_taxrate_itemRow
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.FillBy_Primary(TBL, CDec(Primary))

            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseTransaksiSalesInvoice_TaxRate.tbltrans_salesinvoice_taxrateRow)
                With row
                    Me.mPrimary = row.primary
                    Me.mKMKNumber = row.kmk_number
                    Me.mKMK_Date = row.kmk_date
                    Me.mStartDate = row.start_date
                    Me.mEndDate = row.end_date

                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp

                    'isikan data item
                    Me.IsikanDGVKosong(vDGV, i)
                    DTAItem.FillBy_PrimaryTrans(TBLItem, Me.mPrimary)
                    For Each rowItem In TBLItem.Rows
                        For Each rowDGV As DataGridViewRow In vDGV.Rows
                            If rowDGV.Cells(i.ColRatePrimaryCurrency.Index).Value = rowItem.primary_currency Then
                                rowDGV.Cells(i.ColRateKurs.Index).Value = ObjCD.DigitGrouping(rowItem.rate, Me.DefaultPrec)
                            End If
                        Next
                    Next
                End With
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetKMKNumber(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrateTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetKmkNumber(CDec(vPrimary))) = False Then
                Return DTA.ScalarQuery_GetKmkNumber(CDec(vPrimary))
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
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrateTableAdapter

        Try
            Return DTA.ScalarQuery_GetTimestamp(CDec(vPrimary)).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vTaxRateCode As String) As Double
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrateTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimary(vTaxRateCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vTaxRateCode).Value
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
End Class
