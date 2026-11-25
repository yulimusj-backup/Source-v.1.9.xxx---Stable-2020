Option Explicit On
Option Strict Off

Public Class FrmCekSiklikBundling
    Dim LangsungClose As Boolean = False

    Public Enum EnumTipeTranskasiAsalClosing As Integer
        InPurchase '0
        InSR '1
        InAdjust '2
        InStockTakeAdjust '3
        OutBundling '4
        OutSales '5
        OutPR '6
        OutStockTakeAdjust '7
        OutAdjust '8
        InBundling '9
        OverheadCost '10
    End Enum

    Public Enum EnumJenisTransaksiAsalSC As Integer
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
    End Enum

    Dim ListPrimarySiklik As New List(Of String)
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim DTASC As New DseCekSiklikBundlingTableAdapters.tbltransaksi_stock_cardTableAdapter
    Dim TBLSC1 As New DseCekSiklikBundling.tbltransaksi_stock_cardDataTable
    Dim TBLSC2 As New DseCekSiklikBundling.tbltransaksi_stock_cardDataTable
    Public CekFromClosing As Boolean
    Public CheckPeriod As Date

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetBB(ByVal vPrimaryProduct As Double, ByVal vDTASC As DseCekSiklikBundlingTableAdapters.tbltransaksi_stock_cardTableAdapter, _
    ByVal vTglSampai As Date) As Decimal
        Dim RetVal As Nullable(Of Decimal)

        Try
            RetVal = vDTASC.ScalarQuery_GetBB(vPrimaryProduct, vTglSampai)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinBaris()
        Dim rowSC, rowSC1 As DseCekSiklikBundling.tbltransaksi_stock_cardRow

        Try
            For Each rowSC In Me.TBLSC1.Rows
                rowSC1 = TBLSC2.Newtbltransaksi_stock_cardRow
                For ij As Integer = 0 To TBLSC1.Columns.Count - 1
                    rowSC1.Item(ij) = rowSC.Item(ij)
                Next
                TBLSC2.Rows.Add(rowSC1)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Me.CekData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Get_In_Valid_Idx(ByVal vIdxAcuan As Integer, ByVal vTBLSC As DseCekSiklikBundling.tbltransaksi_stock_cardDataTable) As Integer
        Dim rowSC As DseCekSiklikBundling.tbltransaksi_stock_cardRow

        Try
            'mencari qty in, yang valid untuk perhitungan cost
            'kriteria valid, bs dapet nilainya langsung
            '1. semua transaksi in kecuali in bundling
            '2. transaksi in bundling yang tidak in out untuk primary product ini
            vIdxAcuan += 1
            While vIdxAcuan <= vTBLSC.DefaultView.Count - 1
                rowSC = vTBLSC.DefaultView(vIdxAcuan).Row
                If rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.InAdjust Or _
                rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.InPurchase Or _
                rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.InSR Or _
                rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.InStockTakeAdjust Then
                    Return vIdxAcuan
                ElseIf rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.InBundling Then
                    If rowSC.Sudah_Hitung <> 1 Then
                        If Me.ApakahBundlingInOut(rowSC.Primary_Transaksi_Asal, vTBLSC) = False Then
                            Return vIdxAcuan
                        End If
                    End If
                End If
                vIdxAcuan += 1
            End While

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Function ApakahBundlingInOut(ByVal vPrimaryTransaksiBundling As Double, ByVal vTBLSC As DseCekSiklikBundling.tbltransaksi_stock_cardDataTable) As Boolean
        Dim In_, Out_ As Boolean
        Dim rowSC As DseCekSiklikBundling.tbltransaksi_stock_cardRow

        Try
            'mengembalikan nilai true jika ada item in dan out yang sama
            'khusus item out, hnay masukkan yang belum dihitung saja
            In_ = False
            Out_ = False

            For Each rowSC In vTBLSC.Rows
                If (rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.InBundling Or _
                rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.OutBundling) And rowSC.Primary_Transaksi_Asal = vPrimaryTransaksiBundling Then
                    If rowSC._In > 0 Then
                        In_ = True
                    ElseIf rowSC.Out > 0 Then
                        If rowSC.Sudah_Hitung <> 1 Then
                            Out_ = True
                        End If
                    End If
                End If
            Next

            If In_ = True And Out_ = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Dim rowDGV As DataGridViewRow
        Dim Kombinasi As String = ""
        Dim PrimaryProduct As Double = -1
        Dim NoTransaksi As String = ""

        Try
            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView1.Rows.Clear()

            If Me.ListPrimarySiklik.Count = 0 Then
                'jika semua item bisa transaksi, maka msgbox ok!
                Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                If Me.CekFromClosing = True Then
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.LangsungClose = True
                    Me.Close()
                    Exit Sub
                Else
                    MsgBox("Data checked", MsgBoxStyle.Information, "Information")
                End If
            Else
                Dim DTASC As New DseCekSiklikBundlingTableAdapters.tbltransaksi_stock_cardTableAdapter

                'jika ada item yang tidak bisa transaksi, maka tampilkan ke dalam dgv, dan msgbox warning
                For Each Kombinasi In Me.ListPrimarySiklik
                    PrimaryProduct = Kombinasi.Substring(0, Kombinasi.IndexOf("-"))
                    NoTransaksi = Kombinasi.Substring(Kombinasi.IndexOf("-") + 1)
                    Me.DataGridView1.Rows.Add()

                    rowDGV = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1)
                    rowDGV.Cells(Me.ColCodeNo.Index).Value = Me.ObjProd.GetProductCode(PrimaryProduct)
                    rowDGV.Cells(Me.ColDesc.Index).Value = Me.ObjProd.GetDescription(PrimaryProduct)
                    rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = PrimaryProduct
                    rowDGV.Cells(Me.ColTransactionNo.Index).Value = NoTransaksi
                Next
                Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                DTASC.Dispose()
                Me.DataGridView1.ClearSelection()

                If Me.CekFromClosing = True Then
                    Me.btnDetemineManual.Visible = True
                    Me.btnDetemineManual.TabStop = True
                    Me.btnRepair.Visible = True
                    Me.btnRepair.TabStop = True
                    Me.ProgressBar1.Visible = False
                    Me.LblProgress.Visible = False

                    Me.DataGridView1.Top = Me.btnRepair.Top + Me.btnRepair.Height + 5
                    Me.DataGridView1.Height = Me.DataGridView2.Top - 5 - Me.DataGridView1.Top

                    MsgBox("Cyclic bundling detected", MsgBoxStyle.Information, "Information")
                    Me.ControlBox = True
                Else
                    MsgBox("Cyclic bundling detected", MsgBoxStyle.Information, "Information")
                End If

                Me.DataGridView1.ClearSelection()
                If Me.DataGridView1.RowCount > 0 Then
                    Me.DataGridView1.Item(0, 0).Selected = True
                Else
                    If Me.CekFromClosing = True And Me.DataGridView1.RowCount = 0 Then
                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Me.LangsungClose = True
                        Me.Close()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmCekSiklikBundling_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.DataGridView1.RowCount > 0 Then
                If Me.LangsungClose = False Then
                    'If MsgBox("Repair data before start monthly closing?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                    Me.LangsungClose = True
                    'Else
                    '    e.Cancel = True
                    'End If
                End If
            Else
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.LangsungClose = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmCekSiklikBundling_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.btnDetemineManual.Text = "Determine cost manually"
            Me.dtpPeriod.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
            Me.DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.DataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Me.DataGridView1.AlternatingRowsDefaultCellStyle.BackColor

            If Me.CekFromClosing = True Then
                Me.LblPeriod.Visible = False
                Me.dtpPeriod.Visible = False
                Me.dtpPeriod.TabStop = False
                Me.btnStart.Visible = False
                Me.btnStart.TabStop = False
                Me.ControlBox = False



                Me.ProgressBar1.Top = Me.dtpPeriod.Top
                Me.ProgressBar1.Left = Me.DataGridView1.Left
                Me.ProgressBar1.Width = Me.DataGridView1.Width

                Me.LblProgress.Top = Me.dtpPeriod.Top + Me.dtpPeriod.Height + 5
                Me.LblProgress.Left = Me.DataGridView1.Left

                'Me.btnDetemineManual.Top = Me.LblProgress.Top + Me.LblProgress.Height + 5
                'Me.btnRepair.Top = Me.btnDetemineManual.Top

                Me.DataGridView1.Top = Me.LblProgress.Top + Me.LblProgress.Height + 5
                Me.DataGridView1.Height = Me.DataGridView2.Top - 5 - Me.DataGridView1.Top

                Me.dtpPeriod.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.CheckPeriod)
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker1.RunWorkerAsync()
                'Me.CekData()


            End If

            'Me.DataGridView1.ClearSelection()
            'If Me.DataGridView1.RowCount > 0 Then
            '    Me.DataGridView1.Item(0, 0).Selected = True
            'End If

            Me.btnDetemineManual.Visible = False
            Me.btnDetemineManual.TabStop = False
            Me.btnRepair.Visible = False
            Me.btnRepair.TabStop = False
            Me.btnStart.Visible = False
            Me.btnStart.TabStop = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CekData()
        Dim DTACek As New DseCekSiklikBundlingTableAdapters.DataTable2TableAdapter
        Dim TBLCek As New DseCekSiklikBundling.DataTable2DataTable
        Dim rowCek As DseCekSiklikBundling.DataTable2Row
        Dim rowSC, rowIn1 As DseCekSiklikBundling.tbltransaksi_stock_cardRow
        Dim Idx1 As Integer = 0
        Dim TipeSC As EnumJenisTransaksiAsalSC = EnumJenisTransaksiAsalSC.Stock_Bundling
        Dim Idx As Integer = 0
        Dim BB As Decimal = 0
        Dim QtyInNext As Decimal = 0

        Try
            Me.ObjCD.UpdateThreading()

            'kosongkan list
            Me.ListPrimarySiklik.Clear()
            Me.TBLSC1.Rows.Clear()
            Me.TBLSC2.Rows.Clear()

            'kumpulkan data code no yang berpotensi siklik bundling
            Me.LblProgress.Text = "Loading data..."
            DTACek.Fill(TBLCek, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value))

            'load stock card
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBLCek.Rows.Count
            For Each rowCek In TBLCek.Rows
                Me.ProgressBar1.Value += 1
                Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - " + Me.ObjProd.GetProductCode(rowCek.Primary_Product)

                DTASC.Fill(TBLSC1, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), rowCek.Primary_Product, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value))
                'ganti tipe transaksi asal berdasarkan monthly closing
                For Each rowSC In TBLSC1.Rows
                    TipeSC = rowSC.Type_Transaksi_Asal
                    If TipeSC = EnumJenisTransaksiAsalSC.Stock_Bundling And rowSC._In > 0 Then
                        'bundling aft
                        rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.InBundling
                    ElseIf TipeSC = EnumJenisTransaksiAsalSC.Stock_Bundling And rowSC.Out > 0 Then
                        'bundling bef
                        rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.OutBundling
                    ElseIf TipeSC = EnumJenisTransaksiAsalSC.Stock_Taking And rowSC._In > 0 Then
                        rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.InStockTakeAdjust
                    ElseIf TipeSC = EnumJenisTransaksiAsalSC.Stock_Taking And rowSC.Out > 0 Then
                        rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.OutStockTakeAdjust
                    ElseIf TipeSC = EnumJenisTransaksiAsalSC.StockIn_Adjust Then
                        rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.InAdjust
                    ElseIf TipeSC = EnumJenisTransaksiAsalSC.StockIn_Purchase Then
                        rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.InPurchase
                    ElseIf TipeSC = EnumJenisTransaksiAsalSC.StockIn_SalesReturn Then
                        rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.InSR
                    ElseIf TipeSC = EnumJenisTransaksiAsalSC.StockOut_Adjust Then
                        rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.OutAdjust
                    ElseIf TipeSC = EnumJenisTransaksiAsalSC.StockOut_PurchaseReturn Then
                        rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.OutPR
                    ElseIf TipeSC = EnumJenisTransaksiAsalSC.StockOut_Sales Then
                        rowSC.Type_Transaksi_Asal = EnumTipeTranskasiAsalClosing.OutSales
                    Else
                        MsgBox("Logic error")
                    End If
                Next

                'urutkan
                TBLSC1.DefaultView.Sort = "[Date], [Type Transaksi Asal], [Primary Transaksi Asal]"

                'cek 1 per 1 transaksi, pastikan ok
                Idx = 0
                BB = Me.GetBB(rowCek.Primary_Product, DTASC, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value) - TimeSpan.FromDays(1))
                While Idx <= TBLSC1.DefaultView.Count - 1
                    rowSC = TBLSC1.DefaultView(Idx).Row
                    If rowSC.Sudah_Hitung = 1 Then
                        'MsgBox("")
                        Idx += 1
                    Else
                        If rowSC._In - rowSC.Out + BB >= 0 Then
                            rowSC.Balance = rowSC._In - rowSC.Out + BB
                            BB = rowSC.Balance
                            rowSC.Sudah_Hitung = 1
                            Idx += 1
                        Else
                            'cari in valid untuk transaksi selanjutnya
                            QtyInNext = 0
                            Idx1 = Idx
                            Idx1 = Me.Get_In_Valid_Idx(Idx1, TBLSC1)
                            If Idx1 = -1 Then
                                'cyclic
                                ListPrimarySiklik.Add(rowCek.Primary_Product.ToString + "-" + rowSC.Document_No)
                                Me.SalinBaris()
                                Exit While
                            Else
                                rowIn1 = TBLSC1.DefaultView(Idx1).Row
                                QtyInNext += rowIn1._In
                                rowIn1.Sudah_Hitung = 1
                                While rowSC._In - rowSC.Out + BB + QtyInNext < 0 Or Idx = -1
                                    Idx1 = Me.Get_In_Valid_Idx(Idx1, TBLSC1)
                                    If Idx1 <> -1 Then
                                        rowIn1 = TBLSC1.DefaultView(Idx1).Row
                                        QtyInNext += rowIn1._In
                                        rowIn1.Sudah_Hitung = 1
                                    Else
                                        'sudah tidak ada lagi yang valid, exit while
                                        Exit While
                                    End If
                                End While
                                If rowSC._In - rowSC.Out + BB + QtyInNext < 0 Or Idx = -1 Then
                                    'cyclic
                                    ListPrimarySiklik.Add(rowCek.Primary_Product.ToString + "-" + rowSC.Document_No)
                                    Me.SalinBaris()
                                    Exit While
                                Else
                                    rowSC.Balance = rowSC._In - rowSC.Out + BB + QtyInNext
                                    BB = rowSC.Balance
                                    rowSC.Sudah_Hitung = 1
                                    Idx += 1
                                End If
                            End If
                        End If
                    End If
                End While
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTACek.Dispose()
            TBLCek.Dispose()
        End Try
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim rowSC As DseCekSiklikBundling.tbltransaksi_stock_cardRow
        Dim DefPrec As Integer = 0
        Dim BB As Decimal = 0
        Dim DTASC As New DseCekSiklikBundlingTableAdapters.tbltransaksi_stock_cardTableAdapter

        Try
            Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView2.Rows.Clear()

            If Me.DataGridView1.SelectedCells.Count > 0 Then
                'isikan tabel cek ke dalam dgv 2
                rowDGV = Me.DataGridView1.Rows(Me.DataGridView1.SelectedCells(0).RowIndex)
                If IsNothing(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value) = False Then
                    TBLSC2.DefaultView.Sort = "[Date], [Type Transaksi Asal], [Primary Transaksi Asal]"
                    TBLSC2.DefaultView.RowFilter = "[Primary Product] = " + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString
                    DefPrec = Me.ObjProd.GetAngkaDecimal(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value)

                    'tambahkan data begining balance
                    BB = Me.GetBB(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value, DTASC, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value) - TimeSpan.FromDays(1))
                    Me.DataGridView2.Rows.Add()
                    rowDGV1 = Me.DataGridView2.Rows(Me.DataGridView2.Rows.Count - 1)
                    If BB = 0 Then
                        rowDGV1.Cells(Me.ColBalance.Index).Value = "-"
                    Else
                        rowDGV1.Cells(Me.ColBalance.Index).Value = Me.ObjCD.DigitGrouping(BB, DefPrec)
                    End If

                    rowDGV1.Cells(Me.ColDate.Index).Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
                    rowDGV1.Cells(Me.ColDocumentNo.Index).Value = ""
                    rowDGV1.Cells(Me.ColQtyIn.Index).Value = "-"
                    rowDGV1.Cells(Me.ColQtyOut.Index).Value = "-"
                    rowDGV1.Cells(Me.ColTransDesc.Index).Value = "Begining Balance"
                    rowDGV1.Cells(Me.ColSudahHitung.Index).Value = 1

                    For Each DV As DataRowView In TBLSC2.DefaultView
                        rowSC = DV.Row
                        Me.DataGridView2.Rows.Add()
                        rowDGV1 = Me.DataGridView2.Rows(Me.DataGridView2.Rows.Count - 1)
                        If rowSC.Balance = 0 Then
                            rowDGV1.Cells(Me.ColBalance.Index).Value = "-"
                        Else
                            rowDGV1.Cells(Me.ColBalance.Index).Value = Me.ObjCD.DigitGrouping(rowSC.Balance, DefPrec)
                        End If
                        rowDGV1.Cells(Me.ColDate.Index).Value = rowSC._Date
                        rowDGV1.Cells(Me.ColDocumentNo.Index).Value = rowSC.Document_No
                        If rowSC._In = 0 Then
                            rowDGV1.Cells(Me.ColQtyIn.Index).Value = "-"
                        Else
                            rowDGV1.Cells(Me.ColQtyIn.Index).Value = Me.ObjCD.DigitGrouping(rowSC._In, DefPrec)
                        End If
                        If rowSC.Out = 0 Then
                            rowDGV1.Cells(Me.ColQtyOut.Index).Value = "-"
                        Else
                            rowDGV1.Cells(Me.ColQtyOut.Index).Value = Me.ObjCD.DigitGrouping(rowSC.Out, DefPrec)
                        End If
                        rowDGV1.Cells(Me.ColTransDesc.Index).Value = rowSC.Description
                        rowDGV1.Cells(Me.ColSudahHitung.Index).Value = rowSC.Sudah_Hitung
                        If rowSC.Sudah_Hitung = 0 Then
                            rowDGV1.DefaultCellStyle.ForeColor = Color.Red
                        End If
                    Next
                End If
            End If
            If Me.ApakahPerluMakeUp = True Then
                Me.MakeUp(DefPrec)
            End If
            'pindahkan data in yang sudah hitung tapi balancenya 0, sehingga tampilan lebih enak dilihat
            Me.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"

            Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASC.Dispose()
        End Try
    End Sub

    Private Sub MakeUp(ByVal vDefPrec As Integer)
        Dim rowDGV, rowDGV1, rowDGV2, rowDGV3 As DataGridViewRow
        Dim Balance1, Balance, QtyIn1, QtyOut1, QtyIn2 As Decimal

        Try
            For ij As Integer = 0 To Me.DataGridView2.Rows.Count - 2
                rowDGV = Me.DataGridView2.Rows(ij)
                If rowDGV.Cells(Me.ColSudahHitung.Index).Value = 1 Then
                    rowDGV1 = Me.DataGridView2.Rows(rowDGV.Index + 1)
                    If IsNumeric(rowDGV.Cells(Me.ColBalance.Index).Value) = False Then
                        Balance = 0
                    Else
                        Balance = CDec(rowDGV.Cells(Me.ColBalance.Index).Value)
                    End If
                    If IsNumeric(rowDGV1.Cells(Me.ColBalance.Index).Value) = True Then
                        Balance1 = CDec(rowDGV1.Cells(Me.ColBalance.Index).Value)
                    Else
                        Balance1 = 0
                    End If
                    If IsNumeric(rowDGV1.Cells(Me.ColQtyIn.Index).Value) = True Then
                        QtyIn1 = CDec(rowDGV1.Cells(Me.ColQtyIn.Index).Value)
                    Else
                        QtyIn1 = 0
                    End If
                    If IsNumeric(rowDGV1.Cells(Me.ColQtyOut.Index).Value) = True Then
                        QtyOut1 = CDec(rowDGV1.Cells(Me.ColQtyOut.Index).Value)
                    Else
                        QtyOut1 = 0
                    End If

                    If Balance1 <> Balance + QtyIn1 - QtyOut1 Then
                        'pindahkan qty in pertama yang sudah dihitung yang ditemui
                        For jk As Integer = ij + 1 To Me.DataGridView2.Rows.Count - 1
                            rowDGV2 = Me.DataGridView2.Rows(jk)
                            If IsNumeric(rowDGV2.Cells(Me.ColQtyIn.Index).Value) = True Then
                                QtyIn2 = rowDGV2.Cells(Me.ColQtyIn.Index).Value
                            Else
                                QtyIn2 = 0
                            End If
                            If QtyIn2 > 0 And rowDGV2.Cells(Me.ColSudahHitung.Index).Value = 1 Then
                                Me.DataGridView2.Rows.InsertCopy(jk, ij + 1)
                                rowDGV3 = Me.DataGridView2.Rows(ij + 1)
                                For xx As Integer = 0 To Me.DataGridView2.ColumnCount - 1
                                    rowDGV3.Cells(xx).Value = rowDGV2.Cells(xx).Value
                                Next
                                Me.DataGridView2.Rows.RemoveAt(jk + 1)

                                'hitung balance pada baris baru
                                rowDGV = Me.DataGridView2.Rows(ij)
                                rowDGV1 = Me.DataGridView2.Rows(rowDGV.Index + 1)
                                If IsNumeric(rowDGV1.Cells(Me.ColQtyIn.Index).Value) = True Then
                                    QtyIn1 = CDec(rowDGV1.Cells(Me.ColQtyIn.Index).Value)
                                Else
                                    QtyIn1 = 0
                                End If
                                If IsNumeric(rowDGV1.Cells(Me.ColQtyOut.Index).Value) = True Then
                                    QtyOut1 = CDec(rowDGV1.Cells(Me.ColQtyOut.Index).Value)
                                Else
                                    QtyOut1 = 0
                                End If
                                If IsNumeric(rowDGV.Cells(Me.ColBalance.Index).Value) = False Then
                                    Balance = 0
                                Else
                                    Balance = CDec(rowDGV.Cells(Me.ColBalance.Index).Value)
                                End If
                                Balance1 = Balance + QtyIn1 - QtyOut1
                                If Balance1 = 0 Then
                                    rowDGV1.Cells(Me.ColBalance.Index).Value = "-"
                                Else
                                    rowDGV1.Cells(Me.ColBalance.Index).Value = Me.ObjCD.DigitGrouping(Balance1, vDefPrec)
                                End If
                                Exit For
                            End If
                        Next
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahPerluMakeUp() As Boolean
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim Balance1, Balance, QtyIn1, QtyOut1 As Decimal

        Try
            For Each rowDGV In Me.DataGridView2.Rows
                If rowDGV.Cells(Me.ColSudahHitung.Index).Value = 1 Then
                    If rowDGV.Index <> Me.DataGridView2.Rows.Count - 1 Then
                        rowDGV1 = Me.DataGridView2.Rows(rowDGV.Index + 1)

                        If IsNumeric(rowDGV.Cells(Me.ColBalance.Index).Value) = False Then
                            Balance = 0
                        Else
                            Balance = CDec(rowDGV.Cells(Me.ColBalance.Index).Value)
                        End If
                        If IsNumeric(rowDGV1.Cells(Me.ColBalance.Index).Value) = True Then
                            Balance1 = CDec(rowDGV1.Cells(Me.ColBalance.Index).Value)
                        Else
                            Balance1 = 0
                        End If
                        If IsNumeric(rowDGV1.Cells(Me.ColQtyIn.Index).Value) = True Then
                            QtyIn1 = CDec(rowDGV1.Cells(Me.ColQtyIn.Index).Value)
                        Else
                            QtyIn1 = 0
                        End If
                        If IsNumeric(rowDGV1.Cells(Me.ColQtyOut.Index).Value) = True Then
                            QtyOut1 = CDec(rowDGV1.Cells(Me.ColQtyOut.Index).Value)
                        Else
                            QtyOut1 = 0
                        End If

                        If Balance1 <> Balance + QtyIn1 - QtyOut1 Then
                            Return True
                        End If
                    End If
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnRepair_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRepair.Click
        Me.Close()
    End Sub

    Private Sub btnDetemineManual_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDetemineManual.Click
        Try
            If MsgBox("Determine cost manually later?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.LangsungClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class