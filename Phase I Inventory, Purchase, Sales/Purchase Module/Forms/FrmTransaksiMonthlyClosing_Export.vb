Option Explicit On
Option Strict Off

Imports MySql.Data.MySqlClient
Imports MySql.Data
Public Class FrmTransaksiMonthlyClosing_Export
    Public PrimaryClosing As Double = -1
    Public ClosingDate As Date = Now
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim DefPrec As Integer = 0
    Dim AppendCountRecap As Integer = 0
    Dim AppendCountRecapWithDetail As Integer = 0
    Dim AppendCountDetail As Integer = 0
    Dim ProgRecap As Integer = 0
    Dim ProgRecapWithDetail As Integer = 0
    Dim ProgRecapDetail As Integer = 0
    Dim FileExtStr_Recap As String = ""
    Dim FileExtStr_RecapWdetail As String = ""
    Dim FileExtStr_Detail As String = ""

    Dim TblQC As New DseTransaksiMonthlyClosing_Export.DataTable1DataTable
    Dim rowQC As DseTransaksiMonthlyClosing_Export.DataTable1Row

    Private Sub ExportProcess()
        Dim dta As New DseTransaksiMonthlyClosing_ExportTableAdapters.tbl_closing_itemTableAdapter
        Dim tbl As New DseTransaksiMonthlyClosing_Export.tbl_closing_itemDataTable
        Dim row As DseTransaksiMonthlyClosing_Export.tbl_closing_itemRow
        Dim dtaDet As New DseTransaksiMonthlyClosing_ExportTableAdapters.tbl_closing_item_detailTableAdapter
        Dim tblDet As New DseTransaksiMonthlyClosing_Export.tbl_closing_item_detailDataTable
        Dim rowDet As DseTransaksiMonthlyClosing_Export.tbl_closing_item_detailRow
        Dim sp As String = ","
        Dim spRplc As String = ","
        Dim StartT As Date = Now
        Dim str As String = ""

        Try
            Me.ObjCD.UpdateThreading()

            dta.Connection.ConnectionString = Core_Module.GetConstring
            dtaDet.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ChkRecap.Checked = True Or Me.ChkRecapWithDetail.Checked = True Then
                dta.Fill(tbl, PrimaryClosing)
            End If

            If Me.ChkRecap.Checked = True Then
                Me.Export_Recap(tbl, sp, spRplc, FileExtStr_Recap)
            End If
            If Me.ChkRecapWithDetail.Checked = True Then
                Me.Export_Recap_WithDetail(tbl, sp, spRplc, FileExtStr_RecapWdetail, PrimaryClosing)
            End If

            tbl.Dispose()

            If Me.ChkDetail.Checked = True Then
                Me.Export_Detail(sp, spRplc, FileExtStr_Detail, PrimaryClosing)
            End If

            'Dim TS As TimeSpan = Now - StartT
            'str = CStr(TS.Ticks / TimeSpan.TicksPerSecond) + " seconds"
            'MsgBox(str, MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExport.Click

        Try
            TblQC.Rows.Clear()

            If Me.ChkRecap.Checked = False And Me.ChkRecapWithDetail.Checked = False And Me.ChkDetail.Checked = False Then
                MsgBox("Please choose format to export first", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.GetFrmDia(Me.FileExtStr_Recap, Me.FileExtStr_RecapWdetail, Me.FileExtStr_Detail) = False Then
                Exit Sub
            End If

            DefPrec = Me.ObjCurr.GetDecimal_DefaultCurrency
            AppendCountRecap = 1000
            AppendCountRecapWithDetail = 1000
            AppendCountDetail = 1000
            ProgRecap = 100
            ProgRecapWithDetail = 100
            ProgRecapDetail = 500

            'Me.ExportProcess()
            'Exit Sub

            Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Try
            Me.Cursor = Cursors.WaitCursor

            Me.ProgressBar1.Visible = True
            Me.Label1.Text = "Preparing..."
            Me.Enabled = False

            Me.ExportProcess()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        Try
            Me.CekHasil()

            MsgBox("File Export Done", MsgBoxStyle.Information, "Information")

            Me.ProgressBar1.Visible = False
            Me.Label1.Text = ""
            Me.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
            Me.Close()
        End Try
    End Sub

    Private Sub CekHasil()
        'Dim i As New FrmTransaksiMonthlyClosing_Export_Exception

        'Try
        '    i.tbl = Me.TblQC
        '    i.ShowDialog()
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Function GetFrmDia(ByRef vFileExtStr_Recap As String, ByRef vFileExtStr_RecapWdetail As String, ByRef vFileExtStr_Detail As String) As Boolean
        Try
            FileExtStr_Recap = ""
            FileExtStr_RecapWdetail = ""
            FileExtStr_Detail = ""

            Me.SaveDia.FileName = Format(Me.ClosingDate, "yyyy-MM")
            Me.SaveDia.Title = "Please specify file"
            Me.SaveDia.Filter = "csv files (*.csv)|*.csv"
            'Me.SaveDia.Filter = "Excel Worksheets (*.xlsx)|*.xlsx|excel files (*.xls)|*.xls|csv files (*.csv)|*.csv"
            Me.SaveDia.FilterIndex = 1
            'Me.SaveDia.DefaultExt = "xls"
            Me.SaveDia.DefaultExt = "csv"
            'Me.SaveDia.FileNames = Format(Now, "yyyy-mm-dd")
            If Me.SaveDia.ShowDialog <> Windows.Forms.DialogResult.OK Or Me.SaveDia.FileName.Trim.Length = 0 Then
                Return False
            End If

            vFileExtStr_Recap = Me.SaveDia.FileName.ToString.Replace(".csv", "_recap.csv")   'System.IO.Path.GetExtension(Me.SaveDia.FileName)
            vFileExtStr_RecapWdetail = Me.SaveDia.FileName.ToString.Replace(".csv", "_recap_with_detail.csv") 'System.IO.Path.GetExtension(Me.SaveDia.FileName)
            vFileExtStr_Detail = Me.SaveDia.FileName.ToString.Replace(".csv", "_detail.csv") 'System.IO.Path.GetExtension(Me.SaveDia.FileName)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub ProsesProgressBar(ByVal vValue As Integer, ByRef vValueMove As Integer, ByVal vValueMax As Integer, ByVal vCountMove As Integer, ByVal vLblFormat As String)
        Dim AdaPergerakan As Boolean = False

        Try
            If vValue > vValueMove Then
                vValueMove += vCountMove
                If vValueMove > vValueMax Then
                    vValueMove = vValueMax
                End If

                AdaPergerakan = True
            Else
                If vValue >= vValueMax Then
                    vValueMove = vValueMax
                    AdaPergerakan = True
                End If
            End If

            If AdaPergerakan = True Then
                Me.ProgressBar1.Value = vValueMove
                Me.Label1.Text = vLblFormat & ": " & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & "% (" & Me.ObjCD.DigitGrouping(Me.ProgressBar1.Value, 0) & " of " & Me.ObjCD.DigitGrouping(Me.ProgressBar1.Maximum, 0) & ")"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Export_Recap(ByVal tbl As DseTransaksiMonthlyClosing_Export.tbl_closing_itemDataTable, _
    ByVal sp As String, ByVal spRplc As String, ByVal vFileExtStr As String)
        Dim row As DseTransaksiMonthlyClosing_Export.tbl_closing_itemRow
        Dim csvVal As String = ""
        Dim ValMove As Integer = 100

        Try
            If tbl.Rows.Count > 0 Then
                Me.ProgressBar1.Value = 0
                Me.ProgressBar1.Maximum = tbl.Rows.Count - 1

                csvVal = "Code No" & sp & "Description" & sp & "Unit" & sp & "Qty Begining" & sp & "Amount Begining" & sp & "Qty In" & sp & _
                "Amount In" & sp & "Qty Out" & sp & "Amount Out" & sp & "Qty End" & sp & "Amount End" & sp & "Inventory Account Code" & sp & _
                "Inventory Account Description" & sp & "Sales Account Code" & sp & "Sales Account Description" & sp & "Sales Return Account Code" & sp & _
                "Sales Return Account Description" & sp & "Cost of Sales Account Code" & sp & "Cost of Sales Account Description" & sp & _
                "Adjusment Account Code" & sp & "Adjustment Account Description" & sp & "Stock Taking Account Code" & sp & "Stock Taking Account Description" & sp & _
                "Category Code" & sp & "Category Description" & sp & "Sub Brand Code" & sp & "Sub Brand Description" & sp & "Brand Code" & sp & _
                "Brand Description" & sp & "Group Code" & sp & "Group Description"

                Me.CekQtyAndAmount(True, False, "Recap")
                My.Computer.FileSystem.WriteAllText(vFileExtStr, "", False, System.Text.Encoding.ASCII)
                tbl.DefaultView.RowFilter = Nothing
                tbl.DefaultView.Sort = "[product code] ASC"
                For i As Integer = 0 To tbl.DefaultView.Count - 1
                    Me.ProsesProgressBar(i, ValMove, Me.ProgressBar1.Maximum, Me.ProgRecap, "Export format recap")

                    row = tbl.DefaultView(i).Row
                    With row
                        csvVal += vbNewLine & .product_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .product_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .unit_code.ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.Begin_Qty, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.Begin_Amount, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.In_Qty, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.In_Amount, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.Out_Qty, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.Out_Amount, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.Ending_Qty, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.Ending_Amount, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += .inventory_acc_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .inventory_acc_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .sales_acc_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .sales_acc_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .sales_return_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .sales_return_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .cogs_acc_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .cogs_acc_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .adjustment_acc_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .adjusment_acc_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .stock_taking_acc_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .stock_taking_acc_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .category_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .category_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .sub_brand_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .sub_brand_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .brand_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .brand_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .group_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .group_desc.ToString.Replace(sp, "")

                        If csvVal.Length >= AppendCountRecap Then
                            My.Computer.FileSystem.WriteAllText(vFileExtStr, csvVal, True, System.Text.Encoding.ASCII)
                            csvVal = ""
                        End If

                        Me.CekQtyAndAmount(False, False, "Recap", row)
                    End With
                Next

                If csvVal <> "" Then
                    My.Computer.FileSystem.WriteAllText(vFileExtStr, csvVal, True, System.Text.Encoding.ASCII)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Export_Recap_WithDetail(ByVal tbl As DseTransaksiMonthlyClosing_Export.tbl_closing_itemDataTable, _
    ByVal sp As String, ByVal spRplc As String, ByVal vFileExtStr As String, ByVal vPrimarClosing As Double)
        Dim row As DseTransaksiMonthlyClosing_Export.tbl_closing_itemRow
        Dim dtaDet As New DseTransaksiMonthlyClosing_ExportTableAdapters.tbl_closing_item_detailTableAdapter
        Dim tblDet As New DseTransaksiMonthlyClosing_Export.tbl_closing_item_detailDataTable
        Dim rowDet As DseTransaksiMonthlyClosing_Export.tbl_closing_item_detailRow
        Dim csvVal As String = ""
        Dim TipeItemDetail As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
        Dim TipeTranskasiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase
        Dim ValMove As Integer = 100
        Dim FileNameRecapWdetail As String = "_Recap_With_Detail"

        Try
            If tbl.Rows.Count > 0 Then
                Me.ProgressBar1.Value = 0
                Me.ProgressBar1.Maximum = tbl.Rows.Count - 1

                tbl.DefaultView.RowFilter = Nothing
                tbl.DefaultView.Sort = "[primary product] ASC"

                dtaDet.Connection.ConnectionString = Core_Module.GetConstring

                For i As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = 0 To 9
                    tblDet.Rows.Clear()

                    Me.Label1.Text = "Get " & i.ToString & " data (" & CInt(i + 1) & " of 10)"

                    If i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Or i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust _
                    Or i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Or i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust _
                    Or i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                        TipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                        'sudah diurutkann by primary product
                        dtaDet.Fill(tblDet, vPrimarClosing, i, TipeItemDetail)
                    ElseIf i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Or i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling _
                    Or i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Or i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales _
                    Or i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                        TipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
                        'sudah diurutkann by primary product
                        dtaDet.Fill(tblDet, vPrimarClosing, i, TipeItemDetail)
                    End If

                    If tblDet.Rows.Count > 0 Then
                        For ij As Integer = 0 To tbl.DefaultView.Count - 1
                            row = tbl.DefaultView(ij).Row
                            With row
                                While tblDet.Rows.Count > 0
                                    rowDet = tblDet.Rows(0)
                                    If .Primary_Product < rowDet.primary_code_no Then
                                        Exit While
                                    ElseIf .Primary_Product = rowDet.primary_code_no Then
                                        If i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                                            .qty_in_purchase += rowDet.qty
                                            .amount_in_purchase += rowDet.amount
                                        ElseIf i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                                            .qty_in_adjust += rowDet.qty
                                            .amount_in_adjust += rowDet.amount
                                        ElseIf i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                                            .qty_in_sr += rowDet.qty
                                            .amount_in_st += rowDet.amount
                                        ElseIf i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                                            .qty_in_taking += rowDet.qty
                                            .amount_in_taking += rowDet.amount
                                        ElseIf i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                                            .qty_in_bundling += rowDet.qty
                                            .amount_in_bundling += rowDet.amount
                                        ElseIf i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then
                                            .qty_out_adjust += rowDet.qty
                                            .amount_out_adjust += rowDet.amount
                                        ElseIf i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                                            .qty_out_bundling += rowDet.qty
                                            .amount_out_bundling += rowDet.amount
                                        ElseIf i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Then
                                            .qty_out_pr += rowDet.qty
                                            .amount_out_pr += rowDet.amount
                                        ElseIf i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then
                                            .qty_out_sales += rowDet.qty
                                            .amount_out_sales += rowDet.amount
                                        ElseIf i = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                                            .qty_out_taking += rowDet.qty
                                            .amount_out_taking += rowDet.amount
                                        End If

                                        tblDet.Rows.Remove(rowDet)
                                    Else
                                        tblDet.Rows.Remove(rowDet)
                                    End If
                                End While
                            End With
                        Next
                    End If
                Next

                csvVal = "Code No" & sp & "Description" & sp & "Unit" & sp & "Qty Begining" & sp & "Amount Begining" & sp & "Qty In (Purchase)" & sp & _
                "Amount In (Purchase)" & sp & "Qty In (Adjusment)" & sp & "Amount In (Adjusment)" & sp & "Qty In (Sales Return)" & sp & "Amount In (Sales Return)" & sp & _
                "Qty In (Stock Taking)" & sp & "Amount In (Stock Taking)" & sp & "Qty In (Bundling)" & sp & "Amount In (Bundling)" & sp & "Qty Out (Sales)" & sp & _
                "Amount Out (Sales)" & sp & "Qty Out (Adjusment)" & sp & "Amount Out (Adjusment)" & sp & "Qty Out (Purchase Return)" & sp & "Amount Out (Purchase Return)" & sp & _
                "Qty Out (Stock Taking)" & sp & "Amount Out (Stock Taking)" & sp & "Qty Out (Bundling)" & sp & "Amount Out (Bundling)" & sp & "Qty End" & sp & "Amount End" & sp & _
                "Inventory Account Code" & sp & "Inventory Account Description" & sp & "Sales Account Code" & sp & "Sales Account Description" & sp & "Sales Return Account Code" & sp & _
                "Sales Return Account Description" & sp & "Cost of Sales Account Code" & sp & "Cost of Sales Account Description" & sp & "Adjusment Account Code" & sp & _
                "Adjustment Account Description" & sp & "Stock Taking Account Code" & sp & "Stock Taking Account Description" & sp & "Category Code" & sp & "Category Description" & sp & _
                "Sub Brand Code" & sp & "Sub Brand Description" & sp & "Brand Code" & sp & "Brand Description" & sp & "Group Code" & sp & "Group Description"

                My.Computer.FileSystem.WriteAllText(vFileExtStr, "", False, System.Text.Encoding.ASCII)
                tbl.DefaultView.RowFilter = Nothing
                tbl.DefaultView.Sort = "[product code] ASC"
                Me.CekQtyAndAmount(True, False, "Recap w/ detail")
                For i As Integer = 0 To tbl.DefaultView.Count - 1
                    Me.ProsesProgressBar(i, ValMove, Me.ProgressBar1.Maximum, Me.ProgRecapWithDetail, "Export format recap w/ detail")

                    row = tbl.DefaultView(i).Row
                    With row
                        csvVal += vbNewLine & .product_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .product_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .unit_code.ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.Begin_Qty, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.Begin_Amount, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.qty_in_purchase, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.amount_in_purchase, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.qty_in_adjust, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.amount_in_adjust, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.qty_in_sr, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.amount_in_st, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.qty_in_taking, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.amount_in_taking, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.qty_in_bundling, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.amount_in_bundling, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.qty_out_sales, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.amount_out_sales, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.qty_out_adjust, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.amount_out_adjust, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.qty_out_pr, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.amount_out_pr, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.qty_out_taking, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.amount_out_taking, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.qty_out_bundling, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.amount_out_bundling, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.Ending_Qty, .product_decimal).ToString.Replace(spRplc, "") & sp
                        csvVal += Me.GetNumericPlusMinus(.Ending_Amount, DefPrec).ToString.Replace(spRplc, "") & sp
                        csvVal += .inventory_acc_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .inventory_acc_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .sales_acc_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .sales_acc_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .sales_return_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .sales_return_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .cogs_acc_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .cogs_acc_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .adjustment_acc_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .adjusment_acc_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .stock_taking_acc_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .stock_taking_acc_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .category_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .category_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .sub_brand_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .sub_brand_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .brand_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .brand_desc.ToString.Replace(spRplc, "") & sp
                        csvVal += .group_code.ToString.Replace(spRplc, "") & sp
                        csvVal += .group_desc.ToString.Replace(sp, "")

                        If csvVal.Length >= AppendCountRecapWithDetail Then
                            My.Computer.FileSystem.WriteAllText(vFileExtStr, csvVal, True, System.Text.Encoding.ASCII)
                            csvVal = ""
                        End If
                    End With

                    Me.CekQtyAndAmount(False, False, "Recap w/ detail", row)
                Next

                If csvVal <> "" Then
                    My.Computer.FileSystem.WriteAllText(vFileExtStr, csvVal, True, System.Text.Encoding.ASCII)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Export_Detail(ByVal sp As String, ByVal spRplc As String, ByVal vFileExtStr As String, ByVal vPrimarClosing As Double)
        Dim row As DseTransaksiMonthlyClosing_Export.tbl_closing_itemRow
        Dim dtaDet As New DseTransaksiMonthlyClosing_ExportTableAdapters.tbl_closing_item_detail2TableAdapter
        Dim tblDet As New DseTransaksiMonthlyClosing_Export.tbl_closing_item_detail2DataTable
        Dim rowDet, rowDetBef As DseTransaksiMonthlyClosing_Export.tbl_closing_item_detail2Row
        Dim dtaDoc As New DseTransaksiMonthlyClosing_ExportTableAdapters.tbl_document_noTableAdapter
        Dim tblDoc, tblDoc2 As New DseTransaksiMonthlyClosing_Export.tbl_document_noDataTable
        Dim rowDoc, rowDoc2 As DseTransaksiMonthlyClosing_Export.tbl_document_noRow
        Dim csvVal As String = ""
        Dim TipeItemDetail As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
        Dim TipeTranskasiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase
        Dim PrimaryTransaksiAsal As Double = -1
        Dim Qty As Decimal = 0
        Dim Amount As Decimal = 0
        Dim Detail As String()
        Dim KeteranganAsal As String()
        Dim Counter As Integer = 0
        Dim BatasAtas As Integer = 0
        Dim BatasBawah As Integer = 0
        Dim TotalRows As Integer = 0
        Dim KeepRows As Integer = 30000
        Dim ValMove As Integer = 500
        Dim FileNamedetail As String = "_Detail"

        Dim cmdTarget As New MySqlCommand
        Dim ConnTarget As New MySqlConnection(Core_Module.GetConstring)

        Dim dtaAkun As New DseTransaksiMonthlyClosing_ExportTableAdapters.tbl_closing_akunTableAdapter
        Dim tblAkun As New DseTransaksiMonthlyClosing_Export.tbl_closing_akunDataTable
        Dim rowAkun As DseTransaksiMonthlyClosing_Export.tbl_closing_akunRow

        Try
            dtaDet.Connection.ConnectionString = Core_Module.GetConstring
            dtaAkun.Connection.ConnectionString = Core_Module.GetConstring
            dtaDoc.Connection.ConnectionString = Core_Module.GetConstring

            'sudah diurutkann by primary product
            dtaDet.Fill(tblDet, vPrimarClosing)

            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = tblDet.Rows.Count - 1

            My.Computer.FileSystem.WriteAllText(vFileExtStr, "", False, System.Text.Encoding.ASCII)
ForAppendData:
            BatasBawah += KeepRows
            BatasAtas = BatasBawah - KeepRows
            TotalRows = tblDet.Rows.Count

            tblDet.DefaultView.RowFilter = ""
            tblDet.DefaultView.Sort = "[Product Code], [tipe item detail], [tgl transaksi asal], [tipe transaksi asal]"
            Me.Label1.Text = "Export format detail: Get product data"
            For i As Integer = tblDet.DefaultView.Count - 1 To 0 Step -1
                rowDet = tblDet.DefaultView(i).Row
                With rowDet
                    If BatasAtas > i Or BatasBawah <= i Then
                        tblDet.Rows.Remove(rowDet)
                    End If
                End With
            Next

            If tblDet.Rows.Count <= 0 Then
                Exit Sub
            End If

            'get account data -----------------------------------------------------------------------------------------------------------------------------
            Me.Label1.Text = "Export format detail: Get account data"
            dtaAkun.Fill(tblAkun, vPrimarClosing)
            For i As Integer = 0 To tblDet.Rows.Count - 1
                rowDet = tblDet.Rows(i)
                With rowDet
                    While tblAkun.DefaultView.Count > 0
                        rowAkun = tblAkun.DefaultView(0).Row
                        If .primary_code_no < rowAkun.Primary_Product Then
                            Exit While
                        ElseIf .primary_code_no = rowAkun.Primary_Product Then
                            .inventory_acc_code = rowAkun.inventory_acc_code
                            .inventory_acc_desc = rowAkun.inventory_acc_desc
                            .adjustment_acc_code = rowAkun.adjustment_acc_code
                            .adjusment_acc_desc = rowAkun.adjusment_acc_desc
                            .stock_taking_acc_code = rowAkun.stock_taking_acc_code
                            .stock_taking_acc_desc = rowAkun.stock_taking_acc_desc
                            .cogs_acc_code = rowAkun.cogs_acc_code
                            .cogs_acc_desc = rowAkun.cogs_acc_desc
                            .sales_acc_code = rowAkun.sales_acc_code
                            .sales_acc_desc = rowAkun.sales_acc_desc
                            .sales_return_code = rowAkun.sales_return_code
                            .sales_return_desc = rowAkun.sales_return_desc
                            Exit While
                        Else
                            tblAkun.Rows.Remove(rowAkun)
                        End If
                    End While
                End With
            Next
            tblAkun.Rows.Clear()
            'get account data -----------------------------------------------------------------------------------------------------------------------------

            'get document no data -----------------------------------------------------------------------------------------------------------------------------
            Me.Label1.Text = "Export format detail: Get document data"
            dtaDoc.Connection.ConnectionString = Core_Module.GetConstring
            tblDet.DefaultView.RowFilter = "[tipe transaksi asal] <> -1"
            tblDet.DefaultView.Sort = "[tipe transaksi asal] ASC, [primary transaksi asal] ASC"
            For i As Integer = 0 To tblDet.DefaultView.Count - 1
                rowDet = tblDet.DefaultView(i).Row
                With rowDet
                    If i = 0 Then
                        'fill aop no
                        TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase '0
                        dtaDoc.Fill(tblDoc, vPrimarClosing, TipeTranskasiAsal) 'AOP
                    End If

                    If .tipe_transaksi_asal <> TipeTranskasiAsal Then
                        TipeTranskasiAsal = .tipe_transaksi_asal

                        If TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then '1
                            dtaDoc.Fill_01_SR(tblDoc, vPrimarClosing, TipeTranskasiAsal) 'SALES RETURN
                        ElseIf TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Or _
                        TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then '2 or 8
                            dtaDoc.Fill_02_InOut(tblDoc, vPrimarClosing, TipeTranskasiAsal)
                        ElseIf TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Or _
                        TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then '3 or 7
                            dtaDoc.Fill_03_Taking(tblDoc, vPrimarClosing, TipeTranskasiAsal)
                        ElseIf TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Or _
                        TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then '4 or 9
                            dtaDoc.Fill_04_Bundling(tblDoc, vPrimarClosing, TipeTranskasiAsal)
                        ElseIf TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then '5
                            dtaDoc.Fill_05_Invoice(tblDoc, vPrimarClosing, TipeTranskasiAsal) 'SALES INVOICE
                        ElseIf TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Then '6
                            dtaDoc.Fill_06_PR(tblDoc, vPrimarClosing, TipeTranskasiAsal) 'PURCHASE RETURN
                        End If
                    End If

                    While tblDoc.Rows.Count > 0
                        rowDoc = tblDoc.Rows(0)
                        If .primary_transaksi_asal < rowDoc.primary_transaksi_asal Then
                            Exit While
                        ElseIf .primary_transaksi_asal = rowDoc.primary_transaksi_asal Then
                            .document_no = rowDoc.No_Transaksi

                            Exit While
                        Else
                            tblDoc.Rows.Remove(rowDoc)
                        End If
                    End While
                End With
            Next
            tblDoc.Rows.Clear()
            'get document no data -----------------------------------------------------------------------------------------------------------------------------

            'translate remarks -----------------------------------------------------------------------------------------------------------------------------
            Me.Label1.Text = "Export format detail: Get translate remarks data"
            For i As Integer = 0 To tblDet.DefaultView.Count - 1
                rowDet = tblDet.DefaultView(i).Row
                With rowDet
                    If .primary = 9751555 Then
                        .primary = 9751555
                    End If
                    If String.IsNullOrEmpty(.remarks) = False Then
                        KeteranganAsal = .remarks.Split(";")
                        For Each Keterangan As String In KeteranganAsal
                            Detail = Keterangan.Split(",")
                            If Detail.Length - 1 >= 0 Then
                                If IsNumeric(Detail(0)) = True Then
                                    TipeTranskasiAsal = Detail(0)
                                Else
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                            If Detail.Length - 1 >= 1 Then
                                If IsNumeric(Detail(0)) = True Then
                                    PrimaryTransaksiAsal = Detail(1)
                                Else
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                            If Detail.Length - 1 >= 2 Then
                                If IsNumeric(Detail(0)) = True Then
                                    Qty = Detail(2)
                                Else
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                            If Detail.Length - 1 >= 3 Then
                                If IsNumeric(Detail(0)) = True Then
                                    Amount = Detail(3)
                                Else
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If

                            rowDoc2 = tblDoc2.Newtbl_document_noRow
                            rowDoc2.primary = .primary
                            rowDoc2.tipe_transaksi_asal = TipeTranskasiAsal
                            rowDoc2.primary_transaksi_asal = PrimaryTransaksiAsal
                            rowDoc2.No_Transaksi = ""
                            rowDoc2.qty = Qty
                            rowDoc2.amount = Amount
                            tblDoc2.Rows.Add(rowDoc2)
                        Next
                    End If
                End With
            Next

            tblDoc2.DefaultView.RowFilter = Nothing
            tblDoc2.DefaultView.Sort = "[tipe transaksi asal] ASC, [primary transaksi asal] ASC"
            For i As Integer = 0 To tblDoc2.DefaultView.Count - 1
                rowDoc2 = tblDoc2.DefaultView(i).Row
                With rowDoc2
                    If .primary = 9751555 Then
                        .primary = 9751555
                    End If
                    If i = 0 Then
                        'fill aop no
                        TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase '0
                        dtaDoc.Fill(tblDoc, vPrimarClosing, TipeTranskasiAsal) 'AOP
                    End If

                    If .tipe_transaksi_asal <> TipeTranskasiAsal Then
                        TipeTranskasiAsal = .tipe_transaksi_asal

                        If TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then '1
                            dtaDoc.Fill_01_SR(tblDoc, vPrimarClosing, TipeTranskasiAsal) 'SALES RETURN
                        ElseIf TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Or _
                        TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then '2 or 8
                            dtaDoc.Fill_02_InOut(tblDoc, vPrimarClosing, TipeTranskasiAsal)
                        ElseIf TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Or _
                        TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then '3 or 7
                            dtaDoc.Fill_03_Taking(tblDoc, vPrimarClosing, TipeTranskasiAsal)
                        ElseIf TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Or _
                        TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then '4 or 9
                            dtaDoc.Fill_04_Bundling(tblDoc, vPrimarClosing, TipeTranskasiAsal)
                        ElseIf TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then '5
                            dtaDoc.Fill_05_Invoice(tblDoc, vPrimarClosing, TipeTranskasiAsal) 'SALES INVOICE
                        ElseIf TipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Then '6
                            dtaDoc.Fill_06_PR(tblDoc, vPrimarClosing, TipeTranskasiAsal) 'PURCHASE RETURN
                        End If
                    End If

                    While tblDoc.Rows.Count > 0
                        rowDoc = tblDoc.Rows(0)
                        If .primary_transaksi_asal < rowDoc.primary_transaksi_asal Then
                            Exit While
                        ElseIf .primary_transaksi_asal = rowDoc.primary_transaksi_asal Then
                            .No_Transaksi = rowDoc.No_Transaksi

                            Exit While
                        Else
                            tblDoc.Rows.Remove(rowDoc)
                        End If
                    End While
                End With
            Next
            tblDoc.Rows.Clear()

            tblDet.DefaultView.RowFilter = Nothing
            tblDet.DefaultView.Sort = "[primary] ASC"
            tblDoc2.DefaultView.RowFilter = Nothing
            tblDoc2.DefaultView.Sort = "[primary] ASC"

            For i As Integer = 0 To tblDet.DefaultView.Count - 1
                rowDet = tblDet.DefaultView(i).Row
                With rowDet
                    .remarks = ""
                    While tblDoc2.DefaultView.Count > 0
                        rowDoc2 = tblDoc2.DefaultView(0).Row
                        If .primary < rowDoc2.primary Then
                            Exit While
                        ElseIf .primary = rowDoc2.primary Then
                            If .remarks <> "" Then
                                .remarks += " + " '+ vbCrLf
                            End If
                            .remarks += Me.GetNumericPlusMinus(rowDoc2.amount, DefPrec) + " (" + rowDoc2.No_Transaksi + " - " + Me.GetNumericPlusMinus(rowDoc2.qty, rowDet.product_dec) + " " + rowDet.Unit_Code + ")"

                            tblDoc2.Rows.Remove(rowDoc2)
                        Else
                            tblDoc2.Rows.Remove(rowDoc2)
                        End If
                    End While
                End With
            Next
            tblDoc2.Rows.Clear()
            'translate remarks -----------------------------------------------------------------------------------------------------------------------------

            'posisikan urutan dan isikan qty begin, amount begin, qty end, amount end -----------------------------------------------------------------------------------------------------------------------------
            Me.Label1.Text = "Export format detail: Get qty && amount data"

            'new
            tblDet.DefaultView.RowFilter = "[tipe item detail] = " & CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem) & _
            " AND [tipe item detail] = " & CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem)
            tblDet.DefaultView.RowFilter = ""
            For i As Integer = 0 To tblDet.DefaultView.Count - 1
                rowDet = tblDet.DefaultView(i).Row
                If rowDet.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem Then
                    rowDet.tgl_transaksi_asal = "1/1/1950"
                ElseIf rowDet.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem Then
                    rowDet.tgl_transaksi_asal = "1/1/2200"
                End If
            Next
            '-------

            tblDet.DefaultView.RowFilter = ""
            tblDet.DefaultView.Sort = "[Product Code], [tgl transaksi asal], [tipe item detail], [tipe transaksi asal], [primary transaksi asal]"
            Dim LastQtyBalance As Decimal = 0
            Dim LastAmountBalance As Decimal = 0
            For i As Integer = 0 To tblDet.DefaultView.Count - 1
                rowDet = tblDet.DefaultView(i).Row
                With rowDet
                    If i = 0 Then
                        LastQtyBalance = 0
                        LastAmountBalance = 0
                    Else
                        rowDetBef = tblDet.DefaultView(i - 1).Row
                        If .primary_code_no = rowDetBef.primary_code_no Then
                            LastQtyBalance = rowDetBef.qty_end
                            LastAmountBalance = rowDetBef.amount_end
                        Else
                            LastQtyBalance = 0
                            LastAmountBalance = 0
                        End If
                    End If

                    If .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem Or _
                    .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem Then
                        .qty = .qty * -1
                        .amount = .amount * -1
                    End If

                    .qty_end = .qty + LastQtyBalance
                    .amount_end = .amount + LastAmountBalance
                End With
            Next
            'posisikan urutan dan isikan qty begin, amount begin, qty end, amount end -----------------------------------------------------------------------------------------------------------------------------

            csvVal = "Code No" & sp & "Description" & sp & "Unit" & sp & "Date" & sp & "Tipe Document" & sp & " Tipe Item" & sp & "Document No" & sp & _
            "Qty Begining" & sp & "Amount Begining" & sp & "Overhead" & sp & "Qty End" & sp & "Amount End" & sp & "Return Price" & sp & "Remarks" & sp & _
            "Inventory Account Code" & sp & "Inventory Account Description" & sp & "Sales Account Code" & sp & "Sales Account Description" & sp & "Sales Return Account Code" & sp & _
            "Sales Return Account Description" & sp & "Cost of Sales Account Code" & sp & "Cost of Sales Account Description" & sp & "Adjusment Account Code" & sp & _
            "Adjustment Account Description" & sp & "Stock Taking Account Code" & sp & "Stock Taking Account Description" & sp & "Category Code" & sp & "Category Description" & sp & _
            "Sub Brand Code" & sp & "Sub Brand Description" & sp & "Brand Code" & sp & "Brand Description" & sp & "Group Code" & sp & "Group Description"

            Me.CekQtyAndAmount(True, True, "Detail")
            For i As Integer = 0 To tblDet.DefaultView.Count - 1
                Me.ProsesProgressBar(i + BatasAtas, ValMove, Me.ProgressBar1.Maximum, Me.ProgRecapDetail, "Export format detail")

                rowDet = tblDet.DefaultView(i).Row
                With rowDet
                    csvVal += vbNewLine & .product_code.ToString.Replace(spRplc, "") & sp
                    csvVal += .product_desc.ToString.Replace(spRplc, "") & sp
                    csvVal += .Unit_Code.ToString.Replace(spRplc, "") & sp
                    csvVal += Me.GetFormatTanggal(.tgl_transaksi_asal).ToString.Replace(spRplc, "") & sp
                    csvVal += Me.GetTransaksiAsal(.tipe_transaksi_asal).ToString.Replace(spRplc, "") & sp
                    csvVal += Me.GetTipeItemDetail(.tipe_item_detail).ToString.Replace(spRplc, "") & sp
                    csvVal += .document_no.ToString.Replace(spRplc, "") & sp
                    csvVal += Me.GetNumericPlusMinus(.qty, .product_dec).ToString.Replace(spRplc, "") & sp
                    csvVal += Me.GetNumericPlusMinus(.amount, DefPrec).ToString.Replace(spRplc, "") & sp
                    csvVal += Me.GetNumericPlusMinus(.Overhead, DefPrec).ToString.Replace(spRplc, "") & sp
                    csvVal += Me.GetNumericPlusMinus(.qty_end, .product_dec).ToString.Replace(spRplc, "") & sp
                    csvVal += Me.GetNumericPlusMinus(.amount_end, DefPrec).ToString.Replace(spRplc, "") & sp
                    csvVal += Me.GetNumericPlusMinus(.return_price, DefPrec).ToString.Replace(spRplc, "") & sp
                    csvVal += .remarks.ToString.Replace(spRplc, "") & sp
                    csvVal += .inventory_acc_code.ToString.Replace(spRplc, "") & sp
                    csvVal += .inventory_acc_desc.ToString.Replace(spRplc, "") & sp
                    csvVal += .sales_acc_code.ToString.Replace(spRplc, "") & sp
                    csvVal += .sales_acc_desc.ToString.Replace(spRplc, "") & sp
                    csvVal += .sales_return_code.ToString.Replace(spRplc, "") & sp
                    csvVal += .sales_return_desc.ToString.Replace(spRplc, "") & sp
                    csvVal += .cogs_acc_code.ToString.Replace(spRplc, "") & sp
                    csvVal += .cogs_acc_desc.ToString.Replace(spRplc, "") & sp
                    csvVal += .adjustment_acc_code.ToString.Replace(spRplc, "") & sp
                    csvVal += .adjusment_acc_desc.ToString.Replace(spRplc, "") & sp
                    csvVal += .stock_taking_acc_code.ToString.Replace(spRplc, "") & sp
                    csvVal += .stock_taking_acc_desc.ToString.Replace(spRplc, "") & sp
                    csvVal += .category_code.ToString.Replace(spRplc, "") & sp
                    csvVal += .category_desc.ToString.Replace(spRplc, "") & sp
                    csvVal += .sub_brand_code.ToString.Replace(spRplc, "") & sp
                    csvVal += .sub_brand_desc.ToString.Replace(spRplc, "") & sp
                    csvVal += .brand_code.ToString.Replace(spRplc, "") & sp
                    csvVal += .brand_desc.ToString.Replace(spRplc, "") & sp
                    csvVal += .group_code.ToString.Replace(spRplc, "") & sp
                    csvVal += .group_desc.ToString.Replace(sp, "")

                    If csvVal.Length >= AppendCountDetail Then
                        My.Computer.FileSystem.WriteAllText(vFileExtStr, csvVal, True, System.Text.Encoding.ASCII)
                        csvVal = ""
                    End If
                End With

                Me.CekQtyAndAmount(False, True, "Detail", Nothing, rowDet)
            Next

            If csvVal <> "" Then
                My.Computer.FileSystem.WriteAllText(vFileExtStr, csvVal, True, System.Text.Encoding.ASCII)
            End If

            If BatasAtas <= TotalRows Then
                dtaDet.Fill(tblDet, vPrimarClosing)
                GoTo ForAppendData
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTipeItemDetail(ByVal vTipeItemDetail As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail) As String
        Dim retval As String = ""

        Try
            If vTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem Then
                retval = "Begining Balance"
            ElseIf vTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem Then
                retval = "In"
            ElseIf vTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem Then
                retval = "Out"
            ElseIf vTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem Then
                retval = "Ending Balance"
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return retval
        End Try
    End Function

    Private Function GetTransaksiAsal(ByVal vTipeTrans As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal) As String
        Dim retval As String = ""

        Try
            If vTipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                retval = "In Purchase"
            ElseIf vTipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                retval = "In Sales Return"
            ElseIf vTipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                retval = "In Adjust"
            ElseIf vTipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                retval = "In Stock Take"
            ElseIf vTipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                retval = "Out Bundling"
            ElseIf vTipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then
                retval = "Out Sales"
            ElseIf vTipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Then
                retval = "Out Purchase Return"
            ElseIf vTipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                retval = "Out Stock Take"
            ElseIf vTipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then
                retval = "Out Adjust"
            ElseIf vTipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                retval = "In Bundling"
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return retval
        End Try
    End Function

    Private Function GetNumericPlusMinus(ByVal vNumeric As Decimal, ByVal vDec As Integer) As String
        Dim retval As String = ""

        Try
            If vNumeric < 0 Then 'jika minus
                retval = "(" & Me.ObjCD.DigitGrouping(vNumeric * -1, vDec).ToString & ")"
            Else : retval = Me.ObjCD.DigitGrouping(vNumeric, vDec).ToString
            End If

            retval = retval.ToString.Replace(",", "")

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return retval
        End Try
    End Function

    Private Function GetFormatTanggal(ByVal vDate As Date) As String

        Try
            If vDate <= #1/1/1950# Or vDate >= #1/1/2200# Then
                Return ""
            End If

            Return Format(vDate, "dd MMM yyyy")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiMonthlyClosing_Export_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ProgressBar1.Visible = False
            Me.Label1.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CekQtyAndAmount(ByVal vAdd As Boolean, ByVal vIsFormatDetail As Boolean, ByVal vFormat As String, _
    Optional ByVal vRow As DseTransaksiMonthlyClosing_Export.tbl_closing_itemRow = Nothing, _
    Optional ByVal vRow2 As DseTransaksiMonthlyClosing_Export.tbl_closing_item_detail2Row = Nothing)
        Try
            If vAdd = True Then
                rowQC = TblQC.NewDataTable1Row
                rowQC.format = vFormat
                TblQC.Rows.Add(rowQC)
                Exit Sub
            End If

            If TblQC.Rows.Count = 0 Then
                rowQC.format = vFormat
                rowQC = TblQC.NewDataTable1Row
                TblQC.Rows.Add(rowQC)
            End If

            rowQC = TblQC.Rows(TblQC.Rows.Count - 1)
            With rowQC
                If vIsFormatDetail = False Then
                    .begin_qty += vRow.Begin_Qty
                    .begin_amount += vRow.Begin_Amount
                    .in_qty += vRow.In_Qty
                    .in_amount += vRow.In_Amount
                    .out_qty += vRow.Out_Qty
                    .out_amount += vRow.Out_Amount
                    .ending_qty += vRow.Ending_Qty
                    .ending_amount += vRow.Ending_Amount
                    .qty_in_purchase += vRow.qty_in_purchase
                    .amount_in_purchase += vRow.amount_in_purchase
                    .qty_in_adjust += vRow.qty_in_adjust
                    .amount_in_adjust += vRow.amount_in_adjust
                    .qty_in_sr += vRow.qty_in_sr
                    .amount_in_st += vRow.amount_in_st
                    .qty_in_bundling += vRow.qty_in_bundling
                    .amount_in_bundling += vRow.amount_in_bundling
                    .qty_in_taking += vRow.qty_in_taking
                    .amount_in_taking += vRow.amount_in_taking
                    .qty_out_sales += vRow.qty_out_sales
                    .amount_out_sales += vRow.amount_out_sales
                    .qty_out_adjust += vRow.qty_out_adjust
                    .amount_out_adjust += vRow.amount_out_adjust
                    .qty_out_pr += vRow.qty_out_pr
                    .amount_out_pr += vRow.amount_out_pr
                    .qty_out_bundling += vRow.qty_out_bundling
                    .amount_out_bundling += vRow.amount_out_bundling
                    .qty_out_taking += vRow.qty_out_taking
                    .amount_out_taking += vRow.amount_out_taking
                Else
                    '.begin_qty += vRow2.qty_begin
                    '.begin_amount += vRow2.amount_begin
                    'If vRow2.qty > 0 Then
                    '    .in_qty += vRow2.qty
                    'Else
                    '    .out_qty += vRow2.qty * -1
                    'End If
                    'If vRow2.amount > 0 Then
                    '    .in_amount += vRow2.amount
                    'Else
                    '    .out_amount += vRow2.amount * -1
                    'End If
                    '.ending_qty += vRow2.qty 'vRow2.qty_end
                    '.ending_amount += vRow2.amount 'vRow2.amount_end

                    If vRow2.tipe_item_detail = 0 Then
                        .begin_qty += vRow2.qty
                        .begin_amount += vRow2.amount
                    ElseIf vRow2.tipe_item_detail = 1 Then
                        .in_qty += vRow2.qty
                        .in_amount += vRow2.amount
                    ElseIf vRow2.tipe_item_detail = 2 Then
                        .out_qty += vRow2.qty
                        .out_amount += vRow2.amount
                    ElseIf vRow2.tipe_item_detail = 3 Then
                        .ending_qty += vRow2.qty
                        .ending_amount += vRow2.amount
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class