Option Explicit On
Option Strict Off

Public Class ClsTransaksiStockCard

    Dim obj As New Core_Module.clsMasterProduct

    Public Enum EnumJenisTransaksiAsal As Integer
        StockIn_Purchase
        StockIn_SalesReturn
        StockIn_Adjust
        StockOut_Sales
        StockOut_PurchaseReturn
        StockOut_Adjust
        Stock_Bundling
        Stock_Taking
        StockIn_Process
        StockIn_BackProcess
        StockOut_Process
        StockIn_Transfer
    End Enum

    Public Enum EnumJenisStockCard As Integer
        Stock_Taking
        Stock_In
        Stock_Out
    End Enum

    Public Sub UpdateDocNo(ByVal vTipeDocAsal As EnumJenisTransaksiAsal, ByVal vPrimaryTransaksiAsal As Double, ByVal vDocNo As String)
        Dim DTA As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_card1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_UpdateDocNo(vDocNo, vPrimaryTransaksiAsal, CInt(vTipeDocAsal))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function Simpan(ByVal vPrimaryTransaksiAsal As Double, ByVal vJenisTransaksi As EnumJenisTransaksiAsal, ByVal vDocumentNo As String, _
   ByVal vPrimaryProduct As Double, ByVal vQtyIn As Double, ByVal vQtyOut As Double, ByVal vDescription As String, _
   ByVal vInOut As EnumJenisStockCard, ByVal vDate As Date, ByVal DTA As dseTransaksiStockCardTableAdapters.tbltransaksi_stock_cardTableAdapter, _
   ByVal vBatchNo As String, ByVal vExpiryDate As Date, ByVal vPrimaryLocation As Double, ByVal vRemarksItem As String) As Boolean
        Try
            If DTA.InsertQuery1(CDec(vPrimaryTransaksiAsal), vJenisTransaksi, CDec(vPrimaryProduct), _
            vDocumentNo, CDec(vQtyIn), CDec(vQtyOut), vDescription, vInOut, vDate, vBatchNo, vExpiryDate, vPrimaryLocation, vRemarksItem) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Simpan(ByVal vPrimaryTransaksiAsal As Double, ByVal vJenisTransaksi As EnumJenisTransaksiAsal, ByVal vDocumentNo As String, _
    ByVal vPrimaryProduct As Double, ByVal vQtyIn As Double, ByVal vQtyOut As Double, ByVal vDescription As String, _
    ByVal vInOut As EnumJenisStockCard, ByVal vDate As Date, _
    ByVal vBatchNo As String, ByVal vExpiryDate As Date, ByVal vPrimaryLocation As Double, ByVal vRemarksItem As String) As Boolean
        Dim DTA As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_cardTableAdapter
        Try
            If DTA.InsertQuery1(vPrimaryTransaksiAsal, vJenisTransaksi, vPrimaryProduct, _
            vDocumentNo, vQtyIn, vQtyOut, vDescription, vInOut, vDate, vBatchNo, vExpiryDate, vPrimaryLocation, vRemarksItem) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksiAsal As Double, ByVal vJenisTransaksiAsal As EnumJenisTransaksiAsal) As Boolean
        Dim DTA As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_cardTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.DeleteQuery(vPrimaryTransaksiAsal, vJenisTransaksiAsal) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini mengembalikan stock pada tanggal tertentu untuk product yang dberikan primarynya
    ''' </summary>
    ''' <param name="vPrimaryProduct"></param>
    ''' <param name="Tanggal"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetStock(ByVal vPrimaryProduct As Double, ByVal Tanggal As Date) As Decimal
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim TotalIn, TotalOut As Decimal
        Dim DTA As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim RetvalIn As Nullable(Of Decimal)
        Dim RetvalOut As Nullable(Of Double)

        Try
            RetvalIn = DTA.ScalarQuery_GetTotalIn(vPrimaryProduct, Tanggal, CDate("01/01/1950"))
            If RetvalIn.HasValue Then
                TotalIn = RetvalIn.Value
            Else
                TotalIn = 0
            End If
            RetvalOut = DTA.ScalarQuery_GetTotalOut(vPrimaryProduct, Tanggal, CDate("01/01/1950"))
            If RetvalOut.HasValue Then
                TotalOut = RetvalOut.Value
            Else
                TotalOut = 0
            End If

            Return TotalIn - TotalOut 'ObjProduct.GetStockAwal(vPrimaryProduct) + TotalIn - TotalOut
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjProduct = Nothing
            DTA.Dispose()
        End Try
    End Function

    Public Function PrintDGV(ByVal vPrimaryProduct As Double, ByVal vStartDate As Date, ByVal vEndDate As Date, _
    ByVal vDGV As DataGridView, ByVal TBL As dseTransaksiStockCard.tbltransaksi_stock_card2DataTable) As Boolean
        Dim DTA1 As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_card2TableAdapter
        Dim StartingBalance As Double = Me.GetStock(vPrimaryProduct, vStartDate - TimeSpan.FromDays(1))
        Dim row1, row2 As dseTransaksiStockCard.tbltransaksi_stock_card2Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim AngkaDecimal As Integer = 0

        Try
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            AngkaDecimal = ObjProduct.GetAngkaDecimal(vPrimaryProduct)

            DTA1.Fill(TBL, vPrimaryProduct, vEndDate, vStartDate)
            'masukkan starting balance
            If TBL.Rows.Count = 0 Then
                TBL.Rows.Add(vStartDate, Format(vStartDate, "dd MMM yyyy"), "-", "Starting Balance", 0, 0, StartingBalance)
            Else
                Dim row As dseTransaksiStockCard.tbltransaksi_stock_card2Row = TBL.NewRow
                row._Date = vStartDate
                row.Balance = StartingBalance
                row.Document_No = "-"
                row.Moving_In = 0
                row.Remarks = "Starting Balance"
                row.Moving_Out = 0

                TBL.Rows.InsertAt(row, 0)
            End If

            'hitung balance untuk setiap itemnya
            For i As Integer = 0 To TBL.Rows.Count - 1
                row1 = TBL.Rows(i)
                If i <> 0 Then
                    row2 = TBL.Rows(i - 1)

                    row1.Balance = row2.Balance + row1.Moving_In - row1.Moving_Out
                End If
                row1.Transaction_Date = Format(row1._Date, "dd MMM yyyy")

                'format data
                If row1.Balance = 0 Then
                    row1.End_Str = "-"
                Else
                    row1.End_Str = ObjCD.DigitGrouping(row1.Balance, AngkaDecimal)
                End If
                If row1.Moving_In = 0 Then
                    row1.Moving_In_Str = "-"
                Else
                    row1.Moving_In_Str = ObjCD.DigitGrouping(row1.Moving_In, AngkaDecimal)
                End If
                If row1.Moving_Out = 0 Then
                    row1.Moving_Out_Str = "-"
                Else
                    row1.Moving_Out_Str = ObjCD.DigitGrouping(row1.Moving_Out, AngkaDecimal)
                End If
            Next

            vDGV.DataSource = TBL.DefaultView
            vDGV.Columns(0).Visible = False
            vDGV.Columns(4).Visible = False
            vDGV.Columns(5).Visible = False
            vDGV.Columns(6).Visible = False
            vDGV.Columns(7).HeaderText = "Moving In"
            vDGV.Columns(8).HeaderText = "Moving Out"
            vDGV.Columns(9).HeaderText = "Balance"
            vDGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            vDGV.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            vDGV.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            Dim col As DataGridViewColumn
            For Each col In vDGV.Columns
                col.SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            vDGV.Refresh()

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            ObjCD = Nothing
            ObjProduct = Nothing
        End Try
    End Function

    Public Function Print(ByVal vPrimaryProduct As Double, ByVal vStartDate As Date, ByVal vEndDate As Date, _
    ByVal vReport As crpCetakStockCard, ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vProductCode As String, ByVal vProductDesc As String, ByVal vUnit As String, ByVal vDecimal As Integer, ByVal vShowDetail As Boolean, ByVal vShowRemarks As Boolean) As Boolean
        Dim DTA1 As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_card1TableAdapter
        Dim DS As New dseTransaksiStockCard
        Dim TBL1 As dseTransaksiStockCard.tbltransaksi_stock_card1DataTable = DS.tbltransaksi_stock_card1
        Dim row1, row2 As dseTransaksiStockCard.tbltransaksi_stock_card1Row
        Dim StartingBalance As Double = Me.GetStock(vPrimaryProduct, vStartDate - TimeSpan.FromDays(1))

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA1.Fill(TBL1, vPrimaryProduct, vEndDate, vStartDate)

            'masukkan starting balance

            Dim row As dseTransaksiStockCard.tbltransaksi_stock_card1Row = TBL1.NewRow
            row.Primary = 999999999999
            row.Primary_Product = vPrimaryProduct
            row.Primary_Transaksi_Asal = 1
            row.Type_Transaksi_Asal = 0
            row.InOut = 0
            row._Date = vStartDate
            If Me.GetProductType(vPrimaryProduct) = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                row.Balance = 0
            Else
                row.Balance = StartingBalance
            End If
            row.Document_No = "-"
            row._In = 0
            row.Description = "Starting Balance"
            row.Remarks = ""
            row.Location = ""
            row.Batch_Number = ""
            row.primary_location = -1
            row.Expiry = #1/1/1950#
            row.Product_Code = vProductCode
            row.Product_Desc = vProductDesc
            row.Out = 0
            row.Type = 0
            TBL1.Rows.InsertAt(row, 0)

            'hitung balance untuk setiap itemnya
            For i As Integer = 0 To TBL1.Rows.Count - 1
                If i <> 0 Then
                    row1 = TBL1.Rows(i)
                    row2 = TBL1.Rows(i - 1)

                    If Me.GetProductType(row1.Primary_Product) = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                        row1.Balance = 0
                    Else
                        row1.Balance = row2.Balance + row1._In - row1.Out
                    End If
                End If
                If vShowDetail = True Then
                    row1 = TBL1.Rows(i)
                    If row1.Batch_Number <> "" Then
                        row1.Description = row1.Description + vbCrLf + _
                        "Batch No: " + row1.Batch_Number
                    End If
                    If row1.Expiry <> #1/1/1950# Then
                        row1.Description = row1.Description + " - Expiry: " + Format(row1.Expiry, "dd MMM yyyy")
                    End If
                    If row1.Location <> "" Then
                        row1.Description = row1.Description + vbCrLf + _
                        "Location: " + row1.Location
                    End If
                End If
                If vShowRemarks = True Then
                    row1 = TBL1.Rows(i)
                    If row1.Remarks.Trim <> "" Then
                        row1.Description = row1.Description + vbCrLf + _
                        "Remarks: " + row1.Remarks.Trim
                    End If
                End If
            Next

            'tampilkan report
            vReport.SetDataSource(DS)
            vReport.SetParameterValue(0, vStartDate)
            vReport.SetParameterValue(1, vEndDate)
            vReport.SetParameterValue(2, vProductCode)
            vReport.SetParameterValue(3, vUnit)
            vReport.SetParameterValue(vReport.Parameter_Decimal.ParameterFieldName, vDecimal)

            vCrv.ReportSource = vReport
            vCrv.ShowRefreshButton = False
            vCrv.Zoom(90)
            vCrv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            DTA1.Dispose()
        End Try
    End Function

    Public Function Print_PerCategory(ByVal vTipeGroup As FrmTransaksiStockCard2.EnumProductGroupLevel, ByVal vPrimaryCategory As Double, ByVal vStartDate As Date, ByVal vEndDate As Date, _
            ByVal vReport As crpCetakStockCard2, ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vShowDetail As Boolean, ByVal vJudul As String, _
            ByVal vShowAll As Boolean, ByVal vShowRemarks As Boolean) As Boolean
        Dim DTA1 As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_card1TableAdapter
        Dim DS As New dseTransaksiStockCard
        Dim TBL1 As dseTransaksiStockCard.tbltransaksi_stock_card1DataTable = DS.tbltransaksi_stock_card1
        Dim row1, row2 As dseTransaksiStockCard.tbltransaksi_stock_card1Row
        Dim primaryProductBefore As Double = -1

        Dim dtaProductCode As New dseTransaksiStockCardTableAdapters.tblmaster_productTableAdapter
        Dim tblProductCode As New dseTransaksiStockCard.tblmaster_productDataTable
        Dim rowProductCode As dseTransaksiStockCard.tblmaster_productRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Select Case vTipeGroup
                Case FrmTransaksiStockCard2.EnumProductGroupLevel.Category
                    DTA1.FillBy_PrimaryCategory(TBL1, vEndDate, vStartDate, vPrimaryCategory)
                    dtaProductCode.FillBy_PrimaryCategory(tblProductCode, vPrimaryCategory)
                    Exit Select
                Case FrmTransaksiStockCard2.EnumProductGroupLevel.SubBrand
                    DTA1.FillBy_PrimarySubBrand(TBL1, vEndDate, vStartDate, vPrimaryCategory)
                    dtaProductCode.FillBy_PrimarySubBrand(tblProductCode, vPrimaryCategory)
                    Exit Select
                Case FrmTransaksiStockCard2.EnumProductGroupLevel.Brand
                    DTA1.FillBy_PrimaryBrand(TBL1, vEndDate, vStartDate, vPrimaryCategory)
                    dtaProductCode.FillBy_PrimaryBrand(tblProductCode, vPrimaryCategory)
                    Exit Select
                Case FrmTransaksiStockCard2.EnumProductGroupLevel.ProductGroup
                    DTA1.FillBy_PrimaryProductGroup(TBL1, vEndDate, vStartDate, vPrimaryCategory)
                    dtaProductCode.FillBy_ProductGroup(tblProductCode, vPrimaryCategory)
                    Exit Select
            End Select

            Dim rowCount As Integer = TBL1.Rows.Count - 1
            Dim rowAktif As dseTransaksiStockCard.tbltransaksi_stock_card1Row
            Dim rowDiperiksaTerakhir As Integer = -1

            While rowDiperiksaTerakhir <> TBL1.Rows.Count - 1
                For i As Integer = rowDiperiksaTerakhir + 1 To rowCount
                    rowAktif = TBL1.Rows(i)
                    If primaryProductBefore <> rowAktif.Primary_Product Then
                        primaryProductBefore = rowAktif.Primary_Product
                        Dim row As dseTransaksiStockCard.tbltransaksi_stock_card1Row = TBL1.NewRow
                        row.Primary = 999999999999
                        row.Primary_Product = rowAktif.Primary_Product
                        row.Primary_Transaksi_Asal = 1
                        row.Type_Transaksi_Asal = 0
                        row.InOut = 0
                        row._Date = vStartDate
                        If rowAktif.Type = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                            row.Balance = 0
                        Else
                            row.Balance = Me.GetStock(rowAktif.Primary_Product, vStartDate - TimeSpan.FromDays(1))
                        End If
                        row.Document_No = "-"
                        row._In = 0
                        row.Description = "Starting Balance"
                        row.Remarks = ""
                        row.Location = ""
                        row.Batch_Number = ""
                        row.primary_location = -1
                        row.Expiry = #1/1/1950#
                        row.Product_Code = rowAktif.Product_Code
                        row.Product_Desc = rowAktif.Product_Desc
                        row.Out = 0
                        row.Unit_Code = Me.obj.GetUnitCode(rowAktif.Primary_Product)
                        row.Type = 0
                        TBL1.Rows.InsertAt(row, i)
                        rowCount = TBL1.Rows.Count - 1
                    Else
                        If i <> 0 Then
                            row1 = TBL1.Rows(i)
                            row2 = TBL1.Rows(i - 1)

                            If row1.Type = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                                row1.Balance = 0
                            Else
                                row1.Balance = row2.Balance + row1._In - row1.Out
                            End If
                        End If
                        If vShowDetail = True Then
                            row1 = TBL1.Rows(i)
                            If row1.Batch_Number <> "" Then
                                row1.Description = row1.Description + vbCrLf + _
                                "Batch No: " + row1.Batch_Number
                            End If
                            If row1.Expiry <> #1/1/1950# Then
                                row1.Description = row1.Description + " - Expiry: " + Format(row1.Expiry, "dd MMM yyyy")
                            End If
                            If row1.Location <> "" Then
                                row1.Description = row1.Description + vbCrLf + _
                                "Location: " + row1.Location
                            End If
                        End If
                        If vShowRemarks = True Then
                            row1 = TBL1.Rows(i)
                            If row1.Remarks.Trim <> "" Then
                                row1.Description = row1.Description + vbCrLf + _
                                "Remarks: " + row1.Remarks.Trim
                            End If
                        End If
                    End If
                    rowDiperiksaTerakhir = i
                Next
            End While

            For Each rowProductCode In tblProductCode
                TBL1.DefaultView.RowFilter = "[Primary Product] = " & rowProductCode.Primary
                If TBL1.DefaultView.Count = 0 Then
                    If vShowAll = False Then
                        If Me.GetStock(rowProductCode.Primary, vStartDate - TimeSpan.FromDays(1)) > 0 Then
                            Dim row As dseTransaksiStockCard.tbltransaksi_stock_card1Row = TBL1.NewRow
                            row.Primary = 999999999999
                            row.Primary_Product = rowProductCode.Primary
                            row.Primary_Transaksi_Asal = 1
                            row.Type_Transaksi_Asal = 0
                            row.InOut = 0
                            row._Date = vStartDate
                            If Me.GetProductType(rowProductCode.Primary) = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                                row.Balance = 0
                            Else
                                row.Balance = Me.GetStock(rowProductCode.Primary, vStartDate - TimeSpan.FromDays(1))
                            End If
                            row.Document_No = "-"
                            row._In = 0
                            row.Description = "Starting Balance"
                            row.Remarks = ""
                            row.Location = ""
                            row.Batch_Number = ""
                            row.primary_location = -1
                            row.Expiry = #1/1/1950#
                            row.Product_Code = rowProductCode.Product_Code
                            row.Product_Desc = rowProductCode.Description
                            row.Out = 0
                            row.Unit_Code = Me.obj.GetUnitCode(rowProductCode.Primary)
                            row.Type = 0
                            TBL1.Rows.InsertAt(row, TBL1.Rows.Count + 1)
                        End If
                    Else
                        Dim row As dseTransaksiStockCard.tbltransaksi_stock_card1Row = TBL1.NewRow
                        row.Primary = 999999999999
                        row.Primary_Product = rowProductCode.Primary
                        row.Primary_Transaksi_Asal = 1
                        row.Type_Transaksi_Asal = 0
                        row.InOut = 0
                        row._Date = vStartDate
                        If Me.GetProductType(rowProductCode.Primary) = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                            row.Balance = 0
                        Else
                            row.Balance = Me.GetStock(rowProductCode.Primary, vStartDate - TimeSpan.FromDays(1))
                        End If
                        row.Document_No = "-"
                        row._In = 0
                        row.Description = "Starting Balance"
                        row.Remarks = ""
                        row.Location = ""
                        row.Batch_Number = ""
                        row.primary_location = -1
                        row.Expiry = #1/1/1950#
                        row.Product_Code = rowProductCode.Product_Code
                        row.Product_Desc = rowProductCode.Description
                        row.Out = 0
                        row.Unit_Code = Me.obj.GetUnitCode(rowProductCode.Primary)
                        row.Type = 0
                        TBL1.Rows.InsertAt(row, TBL1.Rows.Count + 1)
                    End If
                End If
            Next

            Me.SortTable(TBL1, "Product Code", "asc")

            vReport.SetDataSource(DS)
            vReport.SetParameterValue(0, vStartDate)
            vReport.SetParameterValue(1, vEndDate)
            vReport.SetParameterValue("Judul", vJudul)
            vCrv.ReportSource = vReport
            vCrv.ShowRefreshButton = False
            vCrv.Zoom(90)
            vCrv.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            DTA1.Dispose()
            dtaProductCode.Dispose()
            tblProductCode.Dispose()
        End Try
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Sub SortTable(ByRef mytable As Data.DataTable, ByVal column As String, ByVal vSort As String)
        ''Addi parameter    :[vSort] by Rosadin, AKA Rosadin :p
        ''source site       :http://www.vbforums.com/showthread.php?498623-Sorting-a-datatable-for-REALS!

        'Sorting a table for reals! By SiLentThReaD, AKA Numbchucks.

        'create a blank table. tablecopy is the name
        Dim tablecopy As New Data.DataTable
        'make an exact copy of the table you passed in by reference, and
        'store it in your tablecopy table
        tablecopy = mytable.Copy()
        'clear the contents of the table you just passed in, because
        'you will soon fill it with newly sorted rows
        mytable.Clear()
        'create a rows array
        Dim foundRows() As Data.DataRow
        'assign the foundrows array the values returned
        'from the select. In this case, I left the
        'filterexpression blank, because I want to get
        'everything. I all I need to change is the sort 
        'order.
        foundRows = tablecopy.Select("", column + " " + vSort)

        'iterate through each row and import them into
        'the table that you passed in by reference
        Dim dbrow As Data.DataRow
        For Each dbrow In foundRows
            mytable.ImportRow(dbrow)
        Next

        'finally clear the contents of the tablecopy you created.
        tablecopy.Clear()

    End Sub

    Private Function GetProductType(ByVal vPrimary As Double) As Integer
        Dim dta As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_card1TableAdapter
        Dim temp As Nullable(Of Integer)
        Try
            temp = dta.ScalarQuery_GetType(vPrimary).Value
            If temp.HasValue = True Then
                Return temp
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
