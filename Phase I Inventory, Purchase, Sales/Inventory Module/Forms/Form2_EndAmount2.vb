Option Explicit On
Option Strict Off

Public Class Form2_EndAmount2
    Public Period As Date = #1/1/1950#
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim objStockInOut As New clsTransaksiStockInOut

    Dim DefPrec As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim PrimaryClosing As Double = -1

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Form2_EndAmount2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Dummy As String = ""
        Dim DTA As New DataSet2_EndAmountTableAdapters.tbltrans_custom_sch_cost1TableAdapter
        Dim TBL As New DataSet2_EndAmount.tbltrans_custom_sch_cost1DataTable
        Dim row As DataSet2_EndAmount.tbltrans_custom_sch_cost1Row
        Dim RefDate As Date = Me.ObjCD.GetFirstDateofTheMonth(Period.Date)
        Dim ObjClosing As New ClsTransaksiMonthlyClosingItem
        Dim JmlClosing As Decimal = 0
        Dim DTAClosing As New DataSet2_EndAmountTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim TBLClosing As New DataSet2_EndAmount.tbltrans_purchase_closingDataTable 
        Dim rowClosing As DataSet2_EndAmount.tbltrans_purchase_closingRow
        Dim rowDGV As DataGridViewRow

        Try
            Me.Period = RefDate
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.dgvDetailItem.AlternatingRowsDefaultCellStyle.BackColor = Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor
            Me.dgvItem.ClearSelection()
            Me.DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency

            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItem.Rows.Clear()

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAClosing.Connection.ConnectionString = DTA.Connection.ConnectionString

            PrimaryClosing = ObjClosing.GetPrimaryClosing(RefDate)
            If PrimaryClosing <> -1 Then
                'If MsgBox("Check closing amount for: " + Format(RefDate, "MMM yyyy"), MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                'load data reference amount
                DTA.FillBy_Period(TBL, RefDate)
                If TBL.Rows.Count = 0 Then
                    MsgBox("No reference account data.", MsgBoxStyle.Information, "Information")
                End If
                DTAClosing.Fill(TBLClosing, PrimaryClosing) 

                'load data closing per account
                JmlClosing = 0
                For Each row In TBL.Rows
                    For Each rowClosing In TBLClosing.Rows
                        If rowClosing.Primary_Inventory_Account = row.primary_akun Then
                            If rowClosing.Ending_Amount <> row.amount Then
                                Me.dgvItem.Rows.Add()
                                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                                rowDGV.Cells(Me.ColAcc_AccDesc.Index).Value = row.Account_Code
                                rowDGV.Cells(Me.ColAcc_AccountCode.Index).Value = row.Description
                                rowDGV.Cells(Me.ColAcc_Amount.Index).Value = Me.ObjCD.DigitGrouping(row.amount, Me.DefPrec)
                                rowDGV.Cells(Me.ColAcc_AmountClosing.Index).Value = Me.ObjCD.DigitGrouping(rowClosing.Ending_Amount, Me.DefPrec)
                                rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value = row.primary_akun

                                If Dummy = "" Then
                                    Dummy = row.Account_Code
                                Else
                                    Dummy = Dummy + ", " + row.Account_Code
                                End If
                            End If
                            Exit For
                        End If
                    Next
                Next

                If Dummy <> "" Then
                    MsgBox("Data discrepancy found", MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Data checked", MsgBoxStyle.Information, "Information")
                End If
                'End If
            Else
                MsgBox("Please do monthly closing first: " + Format(RefDate, "MMM yyyy"), MsgBoxStyle.Information, "Information")
            End If
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Dim DTARef As New DataSet2_EndAmountTableAdapters.DataTable2TableAdapter
        Dim TBLRef As New DataSet2_EndAmount.DataTable2DataTable
        Dim rowRef As DataSet2_EndAmount.DataTable2Row
        Dim DTAClosing As New DataSet2_EndAmountTableAdapters.DataTable1TableAdapter
        Dim TBLClosing As New DataSet2_EndAmount.DataTable1DataTable '-> item purchase closing
        Dim TBLClosing1 As New DataSet2_EndAmount.DataTable1DataTable '-> item detail purchase closing
        Dim rowClosing As DataSet2_EndAmount.DataTable1Row
        Dim rowDGV, rowDGV1 As DataGridViewRow

        Try
            Me.dgvDetailItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvDetailItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvDetailItem.Rows.Clear()

            If Me.dgvItem.SelectedCells.Count <> 0 Then
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.SelectedCells(0).RowIndex)

                DTARef.Fill(TBLRef, Me.Period, rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value)
                DTAClosing.Fill(TBLClosing, Me.PrimaryClosing, rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value)
                DTAClosing.FillBy1(TBLClosing1, Me.PrimaryClosing, rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value)

                'masukkan data row closing ke tbl ref
                For ij As Double = TBLClosing.Rows.Count - 1 To 0 Step -1
                    rowClosing = TBLClosing.Rows(ij)
                    TBLRef.DefaultView.RowFilter = "[primary product] = " + rowClosing.Primary_Product.ToString
                    If TBLRef.DefaultView.Count = 0 Then
                        'tambahkan baris baru
                        If rowClosing.Ending_Amount <> 0 Then
                            rowRef = TBLRef.NewDataTable2Row
                            With rowRef
                                .amount = 0
                                .Amount_Closing = rowClosing.Ending_Amount
                                .Description = rowClosing.Description
                                .primary_product = rowClosing.Primary_Product
                                .Product_Code = rowClosing.Product_Code
                                .Qty = rowClosing.Ending_Qty
                                .Unit_Code = rowClosing.Unit_Code
                                ._Decimal = rowClosing._Decimal
                            End With
                            TBLRef.Rows.Add(rowRef)
                        End If
                    Else
                        'masukkan data amount closing
                        If TBLRef.DefaultView.Count = 1 Then
                            rowRef = TBLRef.DefaultView(0).Row
                            rowRef.Amount_Closing = rowClosing.Ending_Amount
                        Else
                            MsgBox("Logic error", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                Next

                For ij As Double = TBLClosing1.Rows.Count - 1 To 0 Step -1
                    rowClosing = TBLClosing1.Rows(ij)
                    TBLRef.DefaultView.RowFilter = "[primary product] = " + rowClosing.Primary_Product.ToString
                    If TBLRef.DefaultView.Count = 0 Then
                        'tambahkan baris baru
                        If rowClosing.Ending_Amount <> 0 Then
                            rowRef = TBLRef.NewDataTable2Row
                            With rowRef
                                .amount = 0
                                .Amount_Closing1 = rowClosing.Ending_Amount
                                .Description = rowClosing.Description
                                .primary_product = rowClosing.Primary_Product
                                .Product_Code = rowClosing.Product_Code
                                .Qty = rowClosing.Ending_Qty
                                .Unit_Code = rowClosing.Unit_Code
                                ._Decimal = rowClosing._Decimal
                            End With
                            TBLRef.Rows.Add(rowRef)
                        End If
                    Else
                        'masukkan data amount closing
                        If TBLRef.DefaultView.Count = 1 Then
                            rowRef = TBLRef.DefaultView(0).Row
                            rowRef.Amount_Closing1 = rowClosing.Ending_Amount
                        Else
                            MsgBox("Logic error", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                Next

                'masukkan data tbl ref ke dgv item product
                For Each rowRef In TBLRef.Rows
                    If rowRef.Amount_Closing <> rowRef.amount Or rowRef.Amount_Closing <> rowRef.Amount_Closing1 Then
                        Me.dgvDetailItem.Rows.Add()
                        rowDGV1 = Me.dgvDetailItem.Rows(Me.dgvDetailItem.Rows.Count - 1)
                        rowDGV1.Cells(Me.ColItem_ClosingAmount.Index).Value = Me.ObjCD.DigitGrouping(rowRef.Amount_Closing, Me.DefPrec)
                        rowDGV1.Cells(Me.ColItem_ClosingAmountDetail.Index).Value = Me.ObjCD.DigitGrouping(rowRef.Amount_Closing1, Me.DefPrec)
                        rowDGV1.Cells(Me.ColItem_CodeNo.Index).Value = rowRef.Product_Code
                        rowDGV1.Cells(Me.ColItem_Desc.Index).Value = rowRef.Description
                        rowDGV1.Cells(Me.ColItem_Qty.Index).Value = Me.ObjCD.DigitGrouping(rowRef.Qty, rowRef._Decimal)
                        rowDGV1.Cells(Me.ColItem_RefAmount.Index).Value = Me.ObjCD.DigitGrouping(rowRef.amount, Me.DefPrec)
                        rowDGV1.Cells(Me.ColItem_Unit.Index).Value = rowRef.Unit_Code
                    Else
                        'MsgBox("")
                    End If
                Next
            End If

            Me.dgvDetailItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvDetailItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnRepair_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRepair.Click
        Dim DTARef As New DataSet2_EndAmountTableAdapters.DataTable2TableAdapter
        Dim TBLRef As New DataSet2_EndAmount.DataTable2DataTable
        Dim rowRef As DataSet2_EndAmount.DataTable2Row
        Dim DTAClosing As New DataSet2_EndAmountTableAdapters.DataTable1TableAdapter
        Dim TBLClosing As New DataSet2_EndAmount.DataTable1DataTable
        Dim rowClosing As DataSet2_EndAmount.DataTable1Row
        Dim rowDGV1 As DataGridViewRow

        Dim DTAAkun As New DataSet2_EndAmountTableAdapters.tbltrans_custom_sch_costTableAdapter
        Dim TBLAkun As New DataSet2_EndAmount.tbltrans_custom_sch_costDataTable
        Dim rowAkun As DataSet2_EndAmount.tbltrans_custom_sch_costRow

        Dim PrimaryStockIn As Double = -1
        Dim JmlTdkRepair As Double = 0

        Dim DTAItem As New DataSet2_EndAmountTableAdapters.tbltrans_purchase_closing_itemTableAdapter
        Dim TBLItem As New DataSet2_EndAmount.tbltrans_purchase_closing_itemDataTable
        Dim rowItem As DataSet2_EndAmount.tbltrans_purchase_closing_itemRow
        Dim DTAItemD As New DataSet2_EndAmountTableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim TBLItemD As New DataSet2_EndAmount.tbltrans_purchase_closing_item_detailDataTable
        Dim TBLItemD_End As New DataSet2_EndAmount.tbltrans_purchase_closing_item_detailDataTable
        Dim rowItemD, rowItemD_End As DataSet2_EndAmount.tbltrans_purchase_closing_item_detailRow
        Dim RefRemarks As String
        Dim ListRemarks As String()
        Dim TipeInOut As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
        Dim MulaiRepair As Boolean = False
        Dim ListDetail As String()
        'tipe transaksi asal, primary transaksi asal, qty, amount
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
        Dim PrimaryTransaksiAsal As Double = -1
        Dim Qty As Decimal = 0
        Dim Amount As Decimal = 0
        Dim QtyTukerTotal, AmountTukerTotal, QtyTuker, AmountTuker As Decimal
        Dim RemarksBaru As String = ""
        Dim DefPrecIdr As Integer = Me.ObjCur.GetDecimal_DefaultCurrency
        Dim Freq_Pindah As Integer = 0

        Try
            If MsgBox("Start data repair?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor

                PrimaryStockIn = Me.objStockInOut.GetPrimary("RESET2017-5")

                Me.dgvDetailItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvDetailItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.dgvDetailItem.Rows.Clear()

                TBLAkun.Rows.Clear()
                DTAAkun.FillBy_Period(TBLAkun, Me.Period)
                For Each rowAkun In TBLAkun.Rows
                    TBLRef.Rows.Clear()
                    TBLClosing.Rows.Clear()

                    DTARef.Fill(TBLRef, Me.Period, rowAkun.primary_akun)
                    DTAClosing.Fill(TBLClosing, Me.PrimaryClosing, rowAkun.primary_akun)

                    For ij As Double = TBLClosing.Rows.Count - 1 To 0 Step -1
                        rowClosing = TBLClosing.Rows(ij)
                        TBLRef.DefaultView.RowFilter = "[primary product] = " + rowClosing.Primary_Product.ToString
                        If TBLRef.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            If rowClosing.Ending_Amount <> 0 Then
                                rowRef = TBLRef.NewDataTable2Row
                                With rowRef
                                    .amount = 0
                                    .Amount_Closing = rowClosing.Ending_Amount
                                    .Description = rowClosing.Description
                                    .primary_product = rowClosing.Primary_Product
                                    .Product_Code = rowClosing.Product_Code
                                    .Qty = rowClosing.Ending_Qty
                                    .Unit_Code = rowClosing.Unit_Code
                                    ._Decimal = rowClosing._Decimal
                                End With
                                TBLRef.Rows.Add(rowRef)
                            End If
                        Else
                            'masukkan data amount closing
                            If TBLRef.DefaultView.Count = 1 Then
                                rowRef = TBLRef.DefaultView(0).Row
                                rowRef.Amount_Closing = rowClosing.Ending_Amount
                            Else
                                MsgBox("Logic error", MsgBoxStyle.Information, "Information")
                            End If
                        End If
                    Next

                    'masukkan data tbl ref ke dgv item product
                    For Each rowRef In TBLRef.Rows
                        If rowRef.Amount_Closing <> rowRef.amount Then
                            Me.dgvDetailItem.Rows.Add()
                            rowDGV1 = Me.dgvDetailItem.Rows(Me.dgvDetailItem.Rows.Count - 1)

                            rowDGV1.Cells(Me.ColItem_ClosingAmount.Index).Value = Me.ObjCD.DigitGrouping(rowRef.Amount_Closing, Me.DefPrec)
                            rowDGV1.Cells(Me.ColItem_CodeNo.Index).Value = rowRef.Product_Code
                            rowDGV1.Cells(Me.ColItem_Desc.Index).Value = rowRef.Description
                            rowDGV1.Cells(Me.ColItem_Qty.Index).Value = Me.ObjCD.DigitGrouping(rowRef.Qty, rowRef._Decimal)
                            rowDGV1.Cells(Me.ColItem_RefAmount.Index).Value = Me.ObjCD.DigitGrouping(rowRef.amount, Me.DefPrec)
                            rowDGV1.Cells(Me.ColItem_Unit.Index).Value = rowRef.Unit_Code

                            'syarat: ada item detail yang menggunakan stock in yang primary-nya diberikan
                            MulaiRepair = False
                            TBLItemD.Rows.Clear()
                            DTAItemD.Fill(TBLItemD, PrimaryClosing, rowRef.primary_product)
                            DTAItem.Fill(TBLItem, PrimaryClosing, rowRef.primary_product)
                            rowItem = TBLItem.Rows(0)

                            For Each rowItemD In TBLItemD.Rows
                                If rowItemD.tipe_item_detail = CInt(TipeInOut) Then
                                    ListRemarks = rowItemD.remarks.Split(";")
                                    For Each RefRemarks In ListRemarks
                                        ListDetail = RefRemarks.Split(",")
                                        'tipe transaksi asal, primary transaksi asal, qty, amount
                                        TipeTransaksiAsal = ListDetail(0)
                                        PrimaryTransaksiAsal = ListDetail(1)
                                        Qty = ListDetail(2)
                                        Amount = ListDetail(3)

                                        If TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust And PrimaryTransaksiAsal = PrimaryStockIn Then
                                            MulaiRepair = True
                                            Exit For
                                        End If
                                    Next
                                End If
                                If MulaiRepair = True Then
                                    Exit For
                                End If
                            Next

                            If MulaiRepair = True Then
                                'pindahkan item detail end balance ke tabel yang berbeda
                                TBLItemD_End.Rows.Clear()
                                For ij As Double = TBLItemD.Rows.Count - 1 To 0 Step -1
                                    rowItemD = TBLItemD.Rows(ij)
                                    If rowItemD.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem Or _
                                    rowItemD.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem Then
                                        TBLItemD.Rows.Remove(rowItemD)
                                    ElseIf rowItemD.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem Then
                                        rowItemD_End = TBLItemD_End.Newtbltrans_purchase_closing_item_detailRow
                                        For sd As Integer = 0 To TBLItemD_End.Columns.Count - 1
                                            rowItemD_End.Item(sd) = rowItemD.Item(sd)
                                        Next
                                        TBLItemD_End.Rows.Add(rowItemD_End)
                                        TBLItemD.Rows.Remove(rowItemD)
                                    End If
                                Next

                                'mulai proses perbaikan, yang harus diperbaiki: amount out pada qty, amount pada item detail, jumlah baris pada end item detail
                                'jurnal yang terjadi diabaikan, seharusnya tidak boleh di abaikan
                                'tapi ga akan kena cek
                                'kurangkan out amount item detail sejumlah amount yang menggunakan stock in adjust tsb: 
                                '1. kurangkan amountnya dan detail item remarksnya
                                '2. tambahkan qty dan amount stock in adjust pada item end detail 
                                '3. ganti dengan item end balance yang bukan dari primary closing yang diberikan
                                '   a. hapus item ybs
                                '   b. pindahkan ke stock out yang tadi: tambahkan amount out-nya dan detail remarksnya
                                For Each rowItemD In TBLItemD.Rows
                                    If rowItemD.tipe_item_detail = CInt(TipeInOut) Then
                                        MulaiRepair = False
                                        Freq_Pindah = 0

                                        If rowItemD.tipe_item_detail = CInt(TipeInOut) Then
                                            ListRemarks = rowItemD.remarks.Split(";")
                                            For Each RefRemarks In ListRemarks
                                                ListDetail = RefRemarks.Split(",")
                                                'tipe transaksi asal, primary transaksi asal, qty, amount
                                                TipeTransaksiAsal = ListDetail(0)
                                                PrimaryTransaksiAsal = ListDetail(1)
                                                Qty = ListDetail(2)
                                                Amount = ListDetail(3)

                                                If TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust _
                                                And PrimaryTransaksiAsal = PrimaryStockIn Then
                                                    Freq_Pindah += 1
                                                    If Freq_Pindah > 1 Then
                                                        'coding belum mengakomodasi kalo pindah 2x
                                                        MsgBox("")
                                                    End If

                                                    'disini mulai proses perbaikannya
                                                    '1. kurangkan amountnya pada item detail
                                                    rowItemD.amount -= Amount
                                                    rowItem.Out_Amount -= Amount
                                                    rowItem.Ending_Amount += Amount

                                                    QtyTukerTotal = 0
                                                    AmountTukerTotal = 0
                                                    RemarksBaru = ""
                                                    For Each rowItemD_End In TBLItemD_End.Rows
                                                        If rowItemD_End.primary_transaksi_asal = PrimaryStockIn And _
                                                        rowItemD_End.tipe_transaksi_asal = CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust) Then
                                                            '2. tambahkan qty dan amount stock in adjust pada item end detail 
                                                            rowItemD_End.qty += Qty
                                                            rowItemD_End.amount += Amount
                                                        Else
                                                            If rowItemD_End.qty > 0 And Qty - QtyTukerTotal > 0 Then
                                                                '3. ganti dengan item end balance yang bukan dari primary closing yang diberikan
                                                                '   a. kurangkan qty dan amount pada end item
                                                                If Qty - QtyTukerTotal < rowItemD_End.qty Then
                                                                    'ambil sebagian dari item end
                                                                    QtyTuker = Qty - QtyTukerTotal
                                                                    AmountTuker = Math.Round((QtyTuker * rowItemD_End.amount) / rowItemD_End.qty, _
                                                                    DefPrecIdr, MidpointRounding.AwayFromZero)
                                                                    rowItemD_End.qty -= QtyTuker
                                                                    rowItemD_End.amount -= AmountTuker
                                                                    QtyTukerTotal += QtyTuker
                                                                    AmountTukerTotal += AmountTuker
                                                                Else
                                                                    'ambil semua dari item end
                                                                    QtyTuker = rowItemD_End.qty
                                                                    AmountTuker = rowItemD_End.amount
                                                                    QtyTukerTotal += QtyTuker
                                                                    AmountTukerTotal += AmountTuker
                                                                    rowItemD_End.qty = 0
                                                                    rowItemD_End.amount = 0
                                                                End If
                                                                If RemarksBaru = "" Then
                                                                    RemarksBaru = rowItemD_End.tipe_transaksi_asal.ToString + "," + rowItemD_End.primary_transaksi_asal.ToString + _
                                                                    QtyTuker.ToString + "," + AmountTuker.ToString
                                                                Else
                                                                    RemarksBaru = RemarksBaru + ";" + rowItemD_End.tipe_transaksi_asal.ToString + "," + rowItemD_End.primary_transaksi_asal.ToString + _
                                                                    QtyTuker.ToString + "," + AmountTuker.ToString
                                                                End If
                                                            End If
                                                        End If
                                                    Next

                                                    '   b. tambahkan amount pada out item
                                                    rowItemD.amount += AmountTukerTotal
                                                    rowItem.Out_Amount += AmountTukerTotal
                                                    rowItem.Ending_Amount -= AmountTukerTotal

                                                    '   c. ganti remarks baru pada out item
                                                    rowItemD.remarks = rowItemD.remarks.Replace(RefRemarks, RemarksBaru)
                                                    '4. update nilainya ke db asli
                                                    '   a. amount dan remarks baru untuk out item detail
                                                    DTAItemD.UpdateQuery(rowItemD.amount, rowItemD.qty, _
                                                    Math.Round(rowItemD.amount / rowItemD.qty, DefPrecIdr, MidpointRounding.AwayFromZero), rowItemD.primary)
                                                    '   b. update qty dan amount untuk item end item 
                                                    For sd As Integer = TBLItemD_End.Rows.Count - 1 To 0 Step -1
                                                        rowItemD_End = TBLItemD_End.Rows(sd)
                                                        With rowItemD_End
                                                            If rowItemD_End.qty <> 0 Then
                                                                'update nilainya ke db
                                                                DTAItemD.UpdateQuery(.amount, .qty, Math.Round(.amount / .qty, _
                                                                DefPrecIdr, MidpointRounding.AwayFromZero), .primary)
                                                            Else
                                                                'hapus ke db
                                                                DTAItemD.DeleteQuery(.primary)
                                                            End If
                                                        End With
                                                    Next
                                                    '   c. end amount dan out amount untuk item closing
                                                    DTAItem.UpdateQuery(rowItem.Out_Amount, rowItem.Ending_Amount, rowItem.Primary)
                                                End If
                                            Next
                                        End If
                                    End If
                                Next

                                'sampe sini tandain sudah direpair 
                                rowDGV1.Cells(Me.ColItem_StatusRepair.Index).Value = Boolean.TrueString 
                            Else
                                rowDGV1.DefaultCellStyle.ForeColor = Color.Red
                            End If
                        Else
                            'MsgBox("")
                        End If
                    Next
                Next
                Me.dgvDetailItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class