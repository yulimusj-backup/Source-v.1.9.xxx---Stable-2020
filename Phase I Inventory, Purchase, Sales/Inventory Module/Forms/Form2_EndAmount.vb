Option Explicit On
Option Strict Off

Public Class Form2_EndAmount
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Buffer As String = ""
    Dim rowIndex As Integer = -1
    Dim ObjAkun As New Core_Module.ClsMasterAkun1
    Dim DefPrecCur As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim LoadFromAccessFinish As Boolean = False
    Dim TipeProgress As EnumTipeProgress = EnumTipeProgress.LoadSelisihAmount
    Dim AccountCode As String = ""
    Dim AccountDesc As String = ""
    Dim PrimaryAcc As Double = -1
    Dim PrimaryProduct As Double = -1
    Dim AmountRef As Decimal = 0
    Dim AmountExcel As Decimal = 0
    Dim CodeNo, Desc, UnitCode As String
    Dim QtyRef, QtyExcel, UnitPrice As Decimal
    Dim DecimalPoint As Integer = 0
    Dim SudahLoad_SelQty, SudahLoad_SelAmount, SudahLoad_SelSC, SudahLoad_DetailHasil, SudahLoad_DetailAcc As Boolean
    Dim TBLAcc As New DataSet2_EndAmount.tbltrans_custom_sch_cost1DataTable
    Dim TBLStockInOut As New DataSet2_EndAmount.tbltransaksi_stockinout_itemDataTable
    Dim PrimaryAccToCheck As Double = -1

    Enum EnumTipeProgress As Integer
        LoadSelisihAmount
        LoadSelisihQty
        LoadSelisihSC
        LoadHasilDetail
        LoadHasilDetailAcc
    End Enum

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Try
            If MsgBox("Load data from access db?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker2.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim DTA As New DataSet2_EndAmountTableAdapters.tbltrans_custom_end_amountTableAdapter

        Try
            If MsgBox("Save detail data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                'hapus data detail
                DTA.Delete()

                'simpan data db
                Me.Simpan_Item_UsingScript()

                Me.prgBar.Value = 0
                Me.lblProgress.Text = "Ready"

                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Simpan_Item_UsingScript() As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()
            Me.prgBar.Value = 0
            Me.prgBar.Maximum = Me.dgvHasil_Detail.Rows.Count

            While StartRowIndex < Me.dgvHasil_Detail.Rows.Count
                Me.prgBar.Value = StartRowIndex
                Me.lblProgress.Text = "Saving on database - data: " + StartRowIndex.ToString + " of " + Me.dgvHasil_Detail.Rows.Count.ToString

                sql = Me.GetInsertString_Item(StartRowIndex, BatchQty)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Me.lblProgress.Text = ""
            Me.prgBar.Value = 0

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_Item(ByVal vStartRowIndex As Double, ByVal BatchQty As Double ) As String
        Dim retval As String = ""
        Dim rowDGV As DataGridViewRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_custom_end_amount` (`primary product`,`period`,`amount`,`primary stock in purchase`,`Qty`,`unit price`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If Me.dgvHasil_Detail.Rows.Count > vStartRowIndex + i Then
                    rowDGV = Me.dgvHasil_Detail.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If Me.dgvHasil_Detail.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With rowDGV
                        retval = retval + vbCrLf + "(" + CDbl(.Cells(Me.ColHasil_PrimaryCodeNo.Index).Value).ToString + "," _
                        + "'" + Format(Me.dtpImportPeriod.Value.Date, "yyyy-MM-dd") + "'," + CDec(.Cells(Me.ColHasil_Amount.Index).Value).ToString + "," _
                        + CDbl(Me.TxtPrimaryStockInAdjust.Text).ToString + "," + CDec(.Cells(Me.ColHasil_Qty.Index).Value).ToString + "," _
                        + CDec(.Cells(Me.ColHasil_UnitCost.Index).Value).ToString + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Sub Form2_EndAmount_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim DocNo As String = "RESET2017-5"
        Dim PrimaryStockIn As Double = -1
        Dim DTAHeader As New DataSet2TableAdapters.tbltransaksi_stockinout1TableAdapter
        Dim DTA As New DataSet2_EndAmountTableAdapters.tbltrans_custom_sch_cost1TableAdapter
        Dim TBL As New DataSet2_EndAmount.tbltrans_custom_sch_cost1DataTable
        Dim row As DataSet2_EndAmount.tbltrans_custom_sch_cost1Row
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvHasil_Acc.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.dgvHasil_Acc.AlternatingRowsDefaultCellStyle.BackColor
            Me.dgvSelisihAmount.AlternatingRowsDefaultCellStyle.BackColor = Me.dgvHasil_Acc.AlternatingRowsDefaultCellStyle.BackColor
            Me.dgvSelisihQty.AlternatingRowsDefaultCellStyle.BackColor = Me.dgvHasil_Acc.AlternatingRowsDefaultCellStyle.BackColor

            Me.DefPrecCur = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.dtpImportPeriod.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpImportPeriod.Value)
            If DTAHeader.ScalarQuery_GetPrimary(DocNo).HasValue = True Then
                PrimaryStockIn = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value
                Me.TxtPrimaryStockInAdjust.Text = PrimaryStockIn
                Me.TxtStockInAdjustNo.Text = DocNo
            Else
                MsgBox("Please specify stock in-out adjust first.", MsgBoxStyle.Information, "Information")
                Me.Close()
            End If

            'load data akun
            DTA.Fill(TBL)
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AllowUserToAddRows = False
            For Each row In TBL.Rows
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                rowDGV.Cells(Me.ColAcc_AccDesc.Index).Value = row.Description
                rowDGV.Cells(Me.ColAcc_AccountCode.Index).Value = row.Account_Code
                rowDGV.Cells(Me.ColAcc_Amount.Index).Value = Me.ObjCD.DigitGrouping(row.amount, Me.DefPrecCur)
                rowDGV.Cells(Me.ColAcc_Period.Index).Value = Format(row.sch, "dd MMM yyyy")
                rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value = row.primary_akun
            Next
            Me.dgvItem.AllowUserToAddRows = True
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Me.dtpImportPeriod.Value = #12/1/2017#
            Me.lblProgress.Text = "Ready"
            Me.dgvSelisihAmount.AllowUserToAddRows = False
            Me.dgvSelisihQty.AllowUserToAddRows = False
            Me.dgvSelisihStock.AllowUserToAddRows = False
            Me.dgvHasil_Acc.AllowUserToAddRows = False
            Me.dgvHasil_Detail.AllowUserToAddRows = False

            Me.dgvHasil_Detail.AlternatingRowsDefaultCellStyle.BackColor = Me.dgvHasil_Acc.AlternatingRowsDefaultCellStyle.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnLoadFromDB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLoadFromDB.Click
        Dim DTA As New DataSet2_EndAmountTableAdapters.tbltransaksi_stockinout_itemTableAdapter 

        Try
            If MsgBox("Load data from db on period: " + Format(Me.dtpImportPeriod.Value.Date, "dd MMM yyyy"), MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                DTA.FillBy_DB(Me.TBLStockInOut, Me.dtpImportPeriod.Value.Date)
                Me.LoadDGV_HasilDetail()
                Me.LoadDGV_SelAccount()
                MsgBox("Data loaded", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.RowIndex = Me.dgvItem.RowCount - 1 Then
                If e.ColumnIndex <> Me.ColAcc_AccountCode.Index Then
                    e.Cancel = True
                    Exit Sub
                Else
                    'col account
                    If e.ColumnIndex = Me.ColAcc_AccountCode.Index Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
            Else
                If e.ColumnIndex <> Me.ColAcc_Period.Index And _
                e.ColumnIndex <> Me.ColAcc_Amount.Index And _
                e.ColumnIndex <> Me.ColAcc_AccountCode.Index Then
                    e.Cancel = True
                    Exit Sub
                Else
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColAcc_AccountCode.Index Then
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                Or Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                    'kalo kode akun dihapus, maka hapus satu bari semuanya
                    If e.RowIndex <> Me.dgvItem.RowCount - 1 Then
                        Me.rowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                        '  Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                    End If
                    Exit Sub
                Else
                    ''data account number
                    If e.ColumnIndex = Me.ColAcc_AccountCode.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Me.dgvItem.Item(Me.ColAcc_PrimaryAkun.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvItem.Item(Me.ColAcc_AccDesc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.dgvItem.Item(Me.ColAcc_PrimaryAkun.Index, e.RowIndex).Value)
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New Core_Module.FrmMasterAkun_Cari1
                            ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            If ij.mPrimaryAkun <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                                Me.dgvItem.Item(Me.ColAcc_AccDesc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                                Me.dgvItem.Item(Me.ColAcc_PrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun
                            ElseIf ij.mPrimaryAkun = -1 Then
                                '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                    Me.rowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                    ' Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                    End If
                End If
            ElseIf e.ColumnIndex = Me.ColAcc_Amount.Index Then
                Me.EndEdit_Amount(sender, e)
            ElseIf e.ColumnIndex = Me.ColAcc_Period.Index Then
                Me.EndEdit_Period(sender, e)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEdit_Period(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            Else
                'pastikan tanggal
                If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
                'set tgl pertama pada bulan ybs + format tampilan
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEdit_Amount(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            Else
                'pastikan angka positif
                If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'format dgv
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrecCur)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0

        Try
            'set nilai default baris baru
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value = -1
            rowDGV.Cells(Me.ColAcc_AccountCode.Index).Value = ""
            rowDGV.Cells(Me.ColAcc_AccDesc.Index).Value = ""
            rowDGV.Cells(Me.ColAcc_Amount.Index).Value = 0
            rowDGV.Cells(Me.ColAcc_Period.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.dgvItem.Rows.RemoveAt(Me.rowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSaveAkun_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveAkun.Click
        Dim rowDGV As DataGridViewRow
        Dim ListPeriod As New List(Of Date)
        Dim ListPrimaryAkun As New List(Of Double)
        Dim DTA As New DataSet2_EndAmountTableAdapters.tbltrans_custom_sch_costTableAdapter

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index <> Me.dgvItem.RowCount - 1 Then
                    'pastikan semua akun amountnya > 0
                    If Me.ObjCD.CekNonNegatif(rowDGV.Cells(Me.ColAcc_Amount.Index).Value) = False Then
                        Me.dgvItem.ClearSelection()
                        rowDGV.Selected = True
                        MsgBox("Please specify positive amount on: " + rowDGV.Cells(Me.ColAcc_AccountCode.Index).Value.ToString, MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    'pastikan semua akun period-nya date
                    If IsDate(rowDGV.Cells(Me.ColAcc_Period.Index).Value) = False Then
                        Me.dgvItem.ClearSelection()
                        rowDGV.Selected = True
                        MsgBox("Please specify valid period on: " + rowDGV.Cells(Me.ColAcc_AccountCode.Index).Value.ToString, MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If
            Next

            'pastikan untuk setiap period akunnya unik
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index <> Me.dgvItem.RowCount - 1 Then
                    If ListPeriod.Contains(CDate(rowDGV.Cells(Me.ColAcc_Period.Index).Value)) = False Then
                        ListPeriod.Add(CDate(rowDGV.Cells(Me.ColAcc_Period.Index).Value))
                    End If
                End If
            Next
            For Each Period As Date In ListPeriod
                ListPrimaryAkun.Clear()
                For Each rowDGV In Me.dgvItem.Rows
                    If rowDGV.Index <> Me.dgvItem.RowCount - 1 Then
                        If CDate(rowDGV.Cells(Me.ColAcc_Period.Index).Value) = Period Then
                            If ListPrimaryAkun.Contains(rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value) = True Then
                                Me.dgvItem.ClearSelection()
                                rowDGV.Selected = True
                                MsgBox("Please specify unique account on: " + rowDGV.Cells(Me.ColAcc_Period.Index).Value.ToString, MsgBoxStyle.Information, "Information")
                                Exit Sub
                            Else
                                ListPrimaryAkun.Add(rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value)
                            End If
                        End If
                    End If
                Next
            Next

            If MsgBox("Save account data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                'hapus data di tabel
                DTA.Delete()

                'simpan data
                For Each rowDGV In Me.dgvItem.Rows
                    If rowDGV.Index <> Me.dgvItem.RowCount - 1 Then
                        DTA.Insert(rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value, CDate(rowDGV.Cells(Me.ColAcc_Period.Index).Value), _
                        CDec(rowDGV.Cells(Me.ColAcc_Amount.Index).Value))
                    End If
                Next

                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Dim DTAExcel As New CustomSMESS.DseEndStockTableAdapters.TblEndStockTableAdapter
        Dim TBLExcel As New CustomSMESS.DseEndStock.TblEndStockDataTable
        Dim rowExcel As CustomSMESS.DseEndStock.TblEndStockRow
        Dim TBLAmount As New DataSet2_EndAmount.tbltrans_custom_end_amountDataTable
        Dim rowAmount As DataSet2_EndAmount.tbltrans_custom_end_amountRow
        Dim DTAProduct As New DataSet2_EndAmountTableAdapters.tblmaster_productTableAdapter
        Dim TBLProduct As New DataSet2_EndAmount.tblmaster_productDataTable
        Dim rowProduct As DataSet2_EndAmount.tblmaster_productRow
        Dim DTAAcc As New DataSet2_EndAmountTableAdapters.tbltrans_custom_sch_cost1TableAdapter
        Dim rowAcc As DataSet2_EndAmount.tbltrans_custom_sch_cost1Row
        Dim Amount As Decimal = 0
        Dim Jml As Decimal = 0
        Dim DV As Data.DataRowView
        Dim DTAStockInOut As New DataSet2_EndAmountTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim rowStockInOut As DataSet2_EndAmount.tbltransaksi_stockinout_itemRow
        Dim ListPrimaryAkun As New List(Of Double)
        Dim DTASC As New DataSet2_EndAmountTableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim TBLSC As New DataSet2_EndAmount.tbltransaksi_stock_cardDataTable
        Dim rowSC, rowSCDummy As DataSet2_EndAmount.tbltransaksi_stock_cardRow
        Dim TBLSCDummy As New DataSet2_EndAmount.tbltransaksi_stock_cardDataTable

        Try
            Me.TBLAcc.Rows.Clear()
            Me.TBLStockInOut.Rows.Clear()

            Me.LoadFromAccessFinish = False
            'load data dari access
            Me.lblProgress.Text = "Loading raw data"
            DTAExcel.Fill(TBLExcel)
            DTAProduct.Fill(TBLProduct)

            'translate ke data table, yang dipindahkan hanya yang amounnya > 0
            Me.prgBar.Value = 0
            Me.prgBar.Maximum = TBLExcel.Rows.Count
            TBLAmount.Rows.Clear()
            For Each rowExcel In TBLExcel.Rows
                Me.prgBar.Value += 1
                If Math.IEEERemainder(Me.prgBar.Value, 10) = 0 Or Me.prgBar.Value = Me.prgBar.Maximum Then
                    Me.lblProgress.Text = "Extracting excel - data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString + " - " + rowExcel.Product_Code
                End If

                If rowExcel.Amount > 0 Then
                    rowAmount = TBLAmount.Newtbltrans_custom_end_amountRow
                    rowAmount.amount = Math.Round(rowExcel.Amount, Me.DefPrecCur, MidpointRounding.AwayFromZero)
                    rowAmount.period = Me.dtpImportPeriod.Value.Date
                    rowAmount.primary = -1
                    TBLProduct.DefaultView.RowFilter = "[Product Code] = '" + rowExcel.Product_Code + "'"
                    rowProduct = TBLProduct.DefaultView(0).Row
                    rowAmount.primary_product = rowProduct.Primary
                    rowAmount.primary_stock_in_purchase = Me.TxtPrimaryStockInAdjust.Text
                    rowAmount.Qty = rowExcel.Qty
                    rowAmount.unit_price = Math.Round(rowAmount.amount / rowAmount.Qty, Me.DefPrecCur, MidpointRounding.AwayFromZero)
                    rowAmount.primary_akun_inv = rowProduct.Primary_Inventory_Account
                    TBLAmount.Rows.Add(rowAmount)
                End If
            Next
            TBLProduct.Rows.Clear()
            TBLProduct.Dispose()
            TBLExcel.Rows.Clear()
            TBLExcel.Dispose()

            'bandingkan amount per account dengan db, load ke dalam dgv jika ada yang masih salah
            Me.lblProgress.Text = "Preparing for amount data check"
            Me.prgBar.Value = 0
            DTAAcc.FillBy_Period(TBLAcc, Me.dtpImportPeriod.Value.Date)
            Me.prgBar.Maximum = TBLAcc.Rows.Count
            For Each rowAcc In TBLAcc.Rows
                Me.prgBar.Value += 1
                Me.lblProgress.Text = "Collecting amount data - " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString + " - Account: " + rowAcc.Account_Code
                Jml = 0
                TBLAmount.DefaultView.RowFilter = "[primary akun inv] = " + rowAcc.primary_akun.ToString
                For Each DV In TBLAmount.DefaultView
                    rowAmount = DV.Row
                    Jml += rowAmount.amount
                Next
                rowAcc.amount_excel = Jml
            Next

            Me.SudahLoad_SelAmount = False
            Me.TipeProgress = EnumTipeProgress.LoadSelisihAmount
            Me.lblProgress.Text = "Comparing amount data"
            Threading.Thread.Sleep(1000)
            Me.BackgroundWorker2.ReportProgress(0)
            While Me.SudahLoad_SelAmount = False
                Threading.Thread.Sleep(10000)
                Me.TBLAcc.Rows.Clear()
            End While

            'banding qty per product dengan db, load ke dgv bila ada yang salah
            'cek juga, pastikan tidak ada unit price = 0
            Me.prgBar.Value = 0
            Me.lblProgress.Text = "Preparing for qty data check"
            DTAStockInOut.Fill(TBLStockInOut, Me.TxtPrimaryStockInAdjust.Text)
            'hanya load data inventory yang ada dalam primary akun yang diberikan
            For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                If rowDGV.Index <> Me.dgvItem.RowCount - 1 Then
                    If CDate(rowDGV.Cells(Me.ColAcc_Period.Index).Value) = Me.dtpImportPeriod.Value.Date Then
                        ListPrimaryAkun.Add(rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value)
                    End If
                End If
            Next
            For ij As Double = TBLStockInOut.Rows.Count - 1 To 0 Step -1
                rowStockInOut = TBLStockInOut.Rows(ij)
                If ListPrimaryAkun.Contains(rowStockInOut.Primary_Inventory_Account) = False Then
                    TBLStockInOut.Rows.RemoveAt(ij)
                End If
            Next
            For Each rowStockInOut In TBLStockInOut.Rows
                TBLAmount.DefaultView.RowFilter = "[primary product] = " + rowStockInOut.Primary_Product.ToString
                If TBLAmount.DefaultView.Count = 0 Then
                    rowStockInOut.Qty_Excel = 0
                    rowStockInOut.Unit_Price = 0
                Else
                    Jml = 0
                    Amount = 0
                    For Each DV In TBLAmount.DefaultView
                        rowAmount = DV.Row
                        Jml += rowAmount.Qty
                        Amount += rowAmount.amount
                    Next
                    rowStockInOut.Qty_Excel = Jml
                    rowStockInOut.Amount = Math.Round(Amount, Me.DefPrecCur, MidpointRounding.AwayFromZero)
                    rowStockInOut.Unit_Price = Math.Round(rowStockInOut.Amount / rowStockInOut.Qty, Me.DefPrecCur, MidpointRounding.AwayFromZero)
                End If
            Next

            Me.SudahLoad_SelQty = False
            Me.lblProgress.Text = "Load selisih stock"
            Threading.Thread.Sleep(1000)
            Me.TipeProgress = EnumTipeProgress.LoadSelisihQty
            Me.BackgroundWorker2.ReportProgress(0)
            While Me.SudahLoad_SelQty = False
                Threading.Thread.Sleep(10000)
            End While

            'cek qty sc vs qty primary stock in
            'a. load data sc
            TBLSC.Rows.Clear()
            Me.lblProgress.Text = "Collecting stock card data"
            Me.prgBar.Value = 0
            Me.prgBar.Maximum = ListPrimaryAkun.Count
            For Each PrimaryAkun As Double In ListPrimaryAkun
                Me.prgBar.Value += 1
                Me.lblProgress.Text = "Collecting stock card data - data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString
                TBLSCDummy.Rows.Clear()
                DTASC.Fill(TBLSCDummy, PrimaryAkun, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpImportPeriod.Value.Date))
                For Each rowSCDummy In TBLSCDummy.Rows
                    rowSC = TBLSC.Newtbltransaksi_stock_cardRow
                    For ij As Integer = 0 To TBLSC.Columns.Count - 1
                        rowSC.Item(ij) = rowSCDummy.Item(ij)
                    Next
                    TBLSC.Rows.Add(rowSC)
                Next
            Next
            'b. bandingkan qty stock in out dengan qty stock card
            For Each rowStockInOut In TBLStockInOut.Rows
                TBLSC.DefaultView.RowFilter = "[Primary Product] = " + rowStockInOut.Primary_Product.ToString
                If TBLSC.DefaultView.Count = 0 Then
                    rowStockInOut.Qty_SC = 0
                Else
                    rowSC = TBLSC.DefaultView(0).Row
                    rowStockInOut.Qty_SC = rowSC.Qty
                End If
            Next
            'c. load ke dalam dgv selisih stock card
            Me.SudahLoad_SelSC = False
            Me.lblProgress.Text = "Load selisih stock card"
            Threading.Thread.Sleep(1000)
            Me.TipeProgress = EnumTipeProgress.LoadSelisihSC
            Me.BackgroundWorker2.ReportProgress(0)
            While Me.SudahLoad_SelSC = False
                Threading.Thread.Sleep(10000)
            End While

            'jika semua ok, sampe sini load ke dalam dgv
            If Me.dgvSelisihAmount.RowCount = 0 And Me.dgvSelisihQty.RowCount = 0 And Me.dgvSelisihStock.RowCount = 0 Then
                Me.prgBar.Value = 0
                Me.lblProgress.Text = "Loading result detail"
                Threading.Thread.Sleep(100)
                Me.SudahLoad_DetailHasil = False
                Me.TipeProgress = EnumTipeProgress.LoadHasilDetail
                Me.BackgroundWorker2.ReportProgress(0)
                While Me.SudahLoad_DetailHasil = False
                    Threading.Thread.Sleep(10000)
                End While

                Me.prgBar.Value = 0
                Me.lblProgress.Text = "Loading account result difference"
                Threading.Thread.Sleep(100)
                Me.SudahLoad_DetailAcc = False
                Me.TipeProgress = EnumTipeProgress.LoadHasilDetailAcc
                Me.BackgroundWorker2.ReportProgress(0)
                While Me.SudahLoad_DetailAcc = False
                    Threading.Thread.Sleep(10000)
                End While

                Me.LoadFromAccessFinish = True
            Else
                MsgBox("Please repair data first", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker2.ProgressChanged
        Try
            If Me.TipeProgress = EnumTipeProgress.LoadSelisihAmount Then
                Me.LoadDGV_SelAmount()
            ElseIf Me.TipeProgress = EnumTipeProgress.LoadSelisihQty Then
                Me.LoadDGV_SelQty()
            ElseIf Me.TipeProgress = EnumTipeProgress.LoadSelisihSC Then
                Me.LoadDGV_SelSC()
            ElseIf Me.TipeProgress = EnumTipeProgress.LoadHasilDetail Then
                Me.LoadDGV_HasilDetail()
            ElseIf Me.TipeProgress = EnumTipeProgress.LoadHasilDetailAcc Then
                Me.LoadDGV_SelAccount()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGV_SelAccount()
        Dim DTAAcc As New DataSet2_EndAmountTableAdapters.tbltrans_custom_sch_cost1TableAdapter
        Dim rowAcc As DataSet2_EndAmount.tbltrans_custom_sch_cost1Row
        Dim rowDGV As DataGridViewRow
        Dim Jml As Decimal = 0

        Try
            Me.prgBar.Style = ProgressBarStyle.Marquee
            DTAAcc.Connection.ConnectionString = Core_Module.GetConstring

            Me.lblProgress.Text = "Loading account result difference"
            Me.prgBar.Value = 0
            DTAAcc.FillBy_Period(TBLAcc, Me.dtpImportPeriod.Value.Date)
            Me.prgBar.Maximum = TBLAcc.Rows.Count

            For Each rowAcc In TBLAcc.Rows
                Me.prgBar.Value += 1
                Me.lblProgress.Text = "Collecting amount data - " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString + " - Account: " + rowAcc.Account_Code
                Jml = 0
                For Each rowDGV In Me.dgvHasil_Detail.Rows
                    If rowDGV.Cells(Me.ColHasil_PrimaryAkun.Index).Value = rowAcc.primary_akun Then
                        Jml += CDec(rowDGV.Cells(Me.ColHasil_Amount.Index).Value)
                    End If
                Next
                rowAcc.amount_excel = Jml
            Next

            Me.dgvHasil_Acc.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvHasil_Acc.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvHasil_Acc.Rows.Clear()
            For Each rowAcc In TBLAcc.Rows
                Me.dgvHasil_Acc.Rows.Add()
                rowDGV = Me.dgvHasil_Acc.Rows(Me.dgvHasil_Acc.Rows.Count - 1)
                rowDGV.Cells(Me.ColHasilAcc_AccDesc.Index).Value = rowAcc.Description
                rowDGV.Cells(Me.ColHasilAcc_AccCode.Index).Value = rowAcc.Account_Code
                rowDGV.Cells(Me.ColHasilAcc_AmountExcel.Index).Value = Me.ObjCD.DigitGrouping(rowAcc.amount_excel, Me.DefPrecCur)
                rowDGV.Cells(Me.ColHasilAcc_AmountRef.Index).Value = Me.ObjCD.DigitGrouping(rowAcc.amount, Me.DefPrecCur)
                rowDGV.Cells(Me.ColHasilAcc_PrimaryAccount.Index).Value = rowAcc.primary_akun
                rowDGV.Cells(Me.ColHasilAcc_Difference.Index).Value = Me.ObjCD.DigitGrouping(rowAcc.amount - rowAcc.amount_excel, Me.DefPrecCur)
            Next
            Me.dgvHasil_Acc.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvHasil_Acc.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.prgBar.Style = ProgressBarStyle.Continuous
            SudahLoad_DetailAcc = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGV_SelSC()
        Dim rowDGV As DataGridViewRow
        Dim rowStockInOut As DataSet2_EndAmount.tbltransaksi_stockinout_itemRow

        Try
            Me.dgvSelisihStock.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvSelisihStock.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvSelisihStock.Rows.Clear()
            For Each rowStockInOut In TBLStockInOut.Rows
                If rowStockInOut.Qty <> rowStockInOut.Qty_SC Then
                    Me.dgvSelisihStock.Rows.Add()
                    rowDGV = Me.dgvSelisihStock.Rows(Me.dgvSelisihStock.Rows.Count - 1)

                    rowDGV.Cells(Me.ColSelStock_CodeNo.Index).Value = rowStockInOut.Product_Code
                    rowDGV.Cells(Me.ColSelStock_Desc.Index).Value = rowStockInOut.Description
                    rowDGV.Cells(Me.ColSelStock_PrimaryCodeNo.Index).Value = rowStockInOut.Primary_Product
                    rowDGV.Cells(Me.ColSelStock_QtySC.Index).Value = Me.ObjCD.DigitGrouping(rowStockInOut.Qty_SC, rowStockInOut._Decimal)
                    rowDGV.Cells(Me.ColSelStock_QtyCustomStock.Index).Value = Me.ObjCD.DigitGrouping(rowStockInOut.Qty, rowStockInOut._Decimal)
                    rowDGV.Cells(Me.ColSelStock_Unit.Index).Value = rowStockInOut.Unit_Code
                End If
            Next
            Me.dgvSelisihStock.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvSelisihStock.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.SudahLoad_SelSC = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGV_HasilDetail()
        Dim rowDGV As DataGridViewRow
        Dim RowStockInOut As DataSet2_EndAmount.tbltransaksi_stockinout_itemRow

        Try
            Me.prgBar.Style = ProgressBarStyle.Marquee
            Me.dgvHasil_Detail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvHasil_Detail.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvHasil_Detail.Rows.Clear()

            For Each RowStockInOut In TBLStockInOut.Rows
                Me.dgvHasil_Detail.Rows.Add()
                rowDGV = Me.dgvHasil_Detail.Rows(Me.dgvHasil_Detail.Rows.Count - 1)

                rowDGV.Cells(Me.ColHasil_CodeNo.Index).Value = RowStockInOut.Product_Code
                rowDGV.Cells(Me.ColHasil_Desc.Index).Value = RowStockInOut.Description
                rowDGV.Cells(Me.ColHasil_PrimaryCodeNo.Index).Value = RowStockInOut.Primary_Product
                rowDGV.Cells(Me.ColHasil_Qty.Index).Value = Me.ObjCD.DigitGrouping(RowStockInOut.Qty, RowStockInOut._Decimal)
                rowDGV.Cells(Me.ColHasil_Unit.Index).Value = RowStockInOut.Unit_Code
                rowDGV.Cells(Me.ColHasil_UnitCost.Index).Value = Me.ObjCD.DigitGrouping(RowStockInOut.Unit_Price, Me.DefPrecCur)
                rowDGV.Cells(Me.ColHasil_AccCode.Index).Value = RowStockInOut.Account_Code
                rowDGV.Cells(Me.ColHasil_AccDesc.Index).Value = RowStockInOut.Account_Desc
                rowDGV.Cells(Me.ColHasil_Amount.Index).Value = Me.ObjCD.DigitGrouping(RowStockInOut.Amount, Me.DefPrecCur)
                rowDGV.Cells(Me.ColHasil_PrimaryAkun.Index).Value = RowStockInOut.Primary_Inventory_Account
                rowDGV.Cells(Me.ColHasil_No.Index).Value = Me.dgvHasil_Detail.Rows.Count
            Next

            Me.dgvHasil_Detail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvHasil_Detail.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.prgBar.Style = ProgressBarStyle.Continuous
            Me.SudahLoad_DetailHasil = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGV_SelQty()
        Dim NST As Decimal = 0

        Try
            Me.prgBar.Style = ProgressBarStyle.Marquee
            Dim rowDGV As DataGridViewRow
            Dim rowStockInOut As DataSet2_EndAmount.tbltransaksi_stockinout_itemRow

            Me.dgvSelisihQty.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvSelisihQty.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvSelisihQty.Rows.Clear()
            For Each rowStockInOut In TBLStockInOut.Rows
                'cek nst, takutnya ada keluar sejumlah satuan terkecil, dan tidak ada nilainya pas perhitungan closing
                NST = Math.Round((rowStockInOut.Amount / rowStockInOut.Qty) / (10 ^ rowStockInOut._Decimal), Me.DefPrecCur, MidpointRounding.AwayFromZero)

                If rowStockInOut.Qty <> rowStockInOut.Qty_Excel Or rowStockInOut.Amount <= 0 Or rowStockInOut.Unit_Price <= 0 Or NST = 0 Then
                    Me.dgvSelisihQty.Rows.Add()
                    rowDGV = Me.dgvSelisihQty.Rows(Me.dgvSelisihQty.Rows.Count - 1)

                    rowDGV.Cells(Me.ColSelQty_CodeNo.Index).Value = rowStockInOut.Product_Code
                    rowDGV.Cells(Me.ColSelQty_Desc.Index).Value = rowStockInOut.Description
                    rowDGV.Cells(Me.ColSelQty_PrimaryProduct.Index).Value = rowStockInOut.Primary_Product
                    rowDGV.Cells(Me.ColSelQty_QtyExcel.Index).Value = Me.ObjCD.DigitGrouping(rowStockInOut.Qty_Excel, rowStockInOut._Decimal)
                    rowDGV.Cells(Me.ColSelQty_QtyRef.Index).Value = Me.ObjCD.DigitGrouping(rowStockInOut.Qty, rowStockInOut._Decimal)
                    rowDGV.Cells(Me.ColSelQty_Unit.Index).Value = rowStockInOut.Unit_Code
                    rowDGV.Cells(Me.ColSelQty_Amount.Index).Value = Me.ObjCD.DigitGrouping(rowStockInOut.Amount, Me.DefPrecCur)
                    rowDGV.Cells(Me.ColSelQty_UnitPrice.Index).Value = Me.ObjCD.DigitGrouping(rowStockInOut.Unit_Price, Me.DefPrecCur)
                    rowDGV.Cells(Me.ColSelQty_DefPrec.Index).Value = rowStockInOut._Decimal
                    If NST = 0 Then
                        rowDGV.Cells(Me.ColSelQty_NST.Index).Value = Boolean.TrueString
                    Else
                        rowDGV.Cells(Me.ColSelQty_NST.Index).Value = Boolean.FalseString
                    End If
                End If
            Next
            Me.dgvSelisihQty.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvSelisihQty.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.SudahLoad_SelQty = True
            Me.prgBar.Style = ProgressBarStyle.Continuous

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGV_SelAmount()
        Dim rowAcc As DataSet2_EndAmount.tbltrans_custom_sch_cost1Row
        Dim rowDGV As DataGridViewRow

        Try
            Me.prgBar.Style = ProgressBarStyle.Marquee
            Me.dgvSelisihAmount.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvSelisihAmount.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvSelisihAmount.Rows.Clear()
            Me.prgBar.Value = 0

            While TBLAcc.Rows.Count > 0
                Me.prgBar.Value += 1
                rowAcc = TBLAcc.Rows(0)

                If rowAcc.amount <> rowAcc.amount_excel Then
                    Me.dgvSelisihAmount.Rows.Add()
                    rowDGV = Me.dgvSelisihAmount.Rows(Me.dgvSelisihAmount.Rows.Count - 1)

                    rowDGV.Cells(Me.ColSelAmt_AccCode.Index).Value = rowAcc.Account_Code
                    rowDGV.Cells(Me.ColSelAmt_AccDesc.Index).Value = rowAcc.Description
                    rowDGV.Cells(Me.ColSelAmt_AmountExcel.Index).Value = Me.ObjCD.DigitGrouping(rowAcc.amount_excel, Me.DefPrecCur)
                    rowDGV.Cells(Me.ColSelAmt_AmountRef.Index).Value = Me.ObjCD.DigitGrouping(rowAcc.amount, Me.DefPrecCur)
                    rowDGV.Cells(Me.ColSelAmt_PrimaryAkun.Index).Value = rowAcc.primary_akun
                    rowDGV.Cells(Me.ColSelAmt_Difference.Index).Value = Me.ObjCD.DigitGrouping(rowAcc.amount - rowAcc.amount_excel, Me.DefPrecCur)
                End If
                TBLAcc.Rows.RemoveAt(0)
            End While
            Me.dgvSelisihAmount.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvSelisihAmount.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.SudahLoad_SelAmount = True
            Me.prgBar.Style = ProgressBarStyle.Continuous
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.lblProgress.Text = "Ready"
            Me.prgBar.Value = 0
            If Me.LoadFromAccessFinish = True Then
                Me.LoadFromAccessFinish = False
                MsgBox("Data loaded.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If MsgBox("Start truncate amount on access db?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker3.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Dim DTAExcel As New CustomSMESS.DseEndStockTableAdapters.TblEndStockTableAdapter
        Dim TBLExcel As New CustomSMESS.DseEndStock.TblEndStockDataTable
        Dim rowExcel As CustomSMESS.DseEndStock.TblEndStockRow
        Try
            Me.ObjCD.UpdateThreading()

            'load data access
            DTAExcel.Fill(TBLExcel)
            Me.prgBar.Value = 0
            Me.prgBar.Maximum = TBLExcel.Rows.Count
            For Each rowExcel In TBLExcel.Rows
                Me.prgBar.Value += 1
                If Me.prgBar.Value = Me.prgBar.Maximum Or Math.IEEERemainder(Me.prgBar.Value, 10) = 0 Then
                    Me.lblProgress.Text = "Truncating amount - data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString
                End If
                rowExcel.Amount = Math.Round(rowExcel.Amount, Me.DefPrecCur)
                DTAExcel.UpdateQuery_SetAmount(rowExcel.Amount, rowExcel.ID)
            Next
            'untuk setip baris truncate amount dan update ke db

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            Me.lblProgress.Text = "Ready"
            Me.prgBar.Value = 0
            MsgBox("Data truncated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClearAccDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAccDB.Click
        Try
            If MsgBox("Clear unused rows in access database?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker5.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnJointAccDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnJointAccDB.Click
        Try
            If MsgBox("Joint 2 database?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker6.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnRepairNST_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRepairNST.Click
        Try
            If MsgBox("Start repair amount regarding NST?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker4.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Dim DTAAcc As New CustomSMESS.DseEndStockTableAdapters.TblEndStockTableAdapter
        Dim TBLAcc As New CustomSMESS.DseEndStock.TblEndStockDataTable
        Dim rowAcc As CustomSMESS.DseEndStock.TblEndStockRow
        Dim rowDGV As DataGridViewRow
        Dim DV As Data.DataRowView

        Try
            Me.lblProgress.Text = "Preparing data..."
            Me.ObjCD.UpdateThreading()
            'load data access
            DTAAcc.Fill(TBLAcc)

            'untuk setiap baris dalam dgv selisih
            'update nilai amount
            Me.prgBar.Value = 0
            Me.prgBar.Maximum = Me.dgvSelisihQty.Rows.Count
            For Each rowDGV In Me.dgvSelisihQty.Rows
                Me.prgBar.Value += 1
                Me.lblProgress.Text = "Updating data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString + " - " + rowDGV.Cells(Me.ColSelQty_CodeNo.Index).Value.ToString

                If rowDGV.Cells(Me.ColSelQty_NST.Index).Value = Boolean.TrueString Then
                    TBLAcc.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(rowDGV.Cells(Me.ColSelQty_CodeNo.Index).Value.ToString) + "'"
                    For Each DV In TBLAcc.DefaultView
                        rowAcc = DV.Row
                        DTAAcc.UpdateQuery_SetAmount(rowAcc.Amount * 10 ^ CInt(rowDGV.Cells(Me.ColSelQty_DefPrec.Index).Value), rowAcc.ID)
                    Next
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            Me.lblProgress.Text = "Ready"
            Me.prgBar.Value = 0
            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker5.DoWork
        Dim DTAAcc As New CustomSMESS.DseEndStockTableAdapters.TblEndStockTableAdapter
        Dim TBLAcc As New CustomSMESS.DseEndStock.TblEndStockDataTable
        Dim TBLAccNew As New CustomSMESS.DseEndStock.TblEndStockDataTable
        Dim rowAcc, rowAccNew As CustomSMESS.DseEndStock.TblEndStockRow
        Dim DTAProduct As New DataSet2_EndAmountTableAdapters.tblmaster_productTableAdapter
        Dim TBLProduct As New DataSet2_EndAmount.tblmaster_productDataTable
        Dim rowProduct As DataSet2_EndAmount.tblmaster_productRow
        Dim DTAAkun As New DataSet2_EndAmountTableAdapters.tbltrans_custom_sch_costTableAdapter
        Dim TBLAkun As New DataSet2_EndAmount.tbltrans_custom_sch_costDataTable
        Dim rowAkun As DataSet2_EndAmount.tbltrans_custom_sch_costRow
        Dim ListPrimaryAkun As New List(Of Double)

        Try
            ObjCD.UpdateThreading()

            Me.lblProgress.Text = "Loading master product"
            DTAProduct.Fill(TBLProduct)

            'menghapus semua item dalam acc db yang akunnya tidak di definisikan pada dgv akun
            '1. isikan list primary akun
            Me.lblProgress.Text = "Loading accout data"
            DTAAkun.Fill(TBLAkun)
            For Each rowAkun In TBLAkun.Rows
                ListPrimaryAkun.Add(rowAkun.primary_akun)
            Next

            '2. load data access
            Me.lblProgress.Text = "Loading access database"
            DTAAcc.Fill(TBLAcc)

            '3. untuk setiap baris data acc, pindahkan baris yang primary akunnya ada di dalam list
            Me.lblProgress.Text = "Filtering data"
            Me.prgBar.Value = 0
            Me.prgBar.Maximum = TBLAcc.Rows.Count
            TBLAccNew.Rows.Clear()
            For Each rowAcc In TBLAcc.Rows
                Me.prgBar.Value += 1
                If Me.prgBar.Value > 25 Then
                    If Math.IEEERemainder(Me.prgBar.Value, 100) = 0 Then
                        Me.lblProgress.Text = "Filtering data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString + " - " + rowAcc.Product_Code
                    ElseIf Me.prgBar.Value = Me.prgBar.Maximum Then
                        Me.lblProgress.Text = "Filtering data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString + " - " + rowAcc.Product_Code
                    End If
                Else
                    Me.lblProgress.Text = "Filtering data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString + " - " + rowAcc.Product_Code
                End If

                For Each rowProduct In TBLProduct.Rows
                    If rowProduct.Product_Code = rowAcc.Product_Code Then
                        If ListPrimaryAkun.Contains(rowProduct.Primary_Inventory_Account) = True Then
                            rowAccNew = TBLAccNew.NewTblEndStockRow
                            For gh As Integer = 0 To TBLAcc.Columns.Count - 1
                                rowAccNew.Item(gh) = rowAcc.Item(gh)
                            Next
                            TBLAccNew.Rows.Add(rowAccNew)
                        End If
                    End If
                Next

                'TBLProduct.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(rowAcc.Product_Code) + "'"
                'If TBLProduct.DefaultView.Count = 1 Then
                '    rowProduct = TBLProduct.DefaultView(0).Row
                'Else
                '    MsgBox("Logic error: product not found.", MsgBoxStyle.Information, "Information")
                'End If
            Next

            '4. hapus tabel asal
            Me.lblProgress.Text = "Deleting access rows"
            DTAAcc.DeleteQuery()

            '5. isikan tabel baru
            Me.lblProgress.Text = "Filtering data"
            Me.prgBar.Value = 0
            Me.prgBar.Maximum = TBLAccNew.Rows.Count
            For Each rowAccNew In TBLAccNew
                Me.prgBar.Value += 1

                If Me.prgBar.Value > 25 Then
                    If Math.IEEERemainder(Me.prgBar.Value, 100) = 0 Then
                        Me.lblProgress.Text = "Inserting data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString + " - " + rowAccNew.Product_Code
                    ElseIf Me.prgBar.Value = Me.prgBar.Maximum Then
                        Me.lblProgress.Text = "Inserting data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString + " - " + rowAccNew.Product_Code
                    End If
                Else
                    Me.lblProgress.Text = "Inserting data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString + " - " + rowAccNew.Product_Code
                End If

                With rowAccNew
                    If .IsWarehouse_CodeNull = True Then
                        .Warehouse_Code = ""
                        .Warehouse_Desc = ""
                    End If
                    DTAAcc.InsertQuery(.Desc, .No, .Product_Code, .Qty, .Unit, .Qty_M3_STD, .Unit_Price, .Warehouse_Code, .Warehouse_Desc, .Amount, .Acc_Code, .Acc_Desc)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAAcc.Dispose()
            TBLAcc.Dispose()
            TBLAccNew.Dispose()
            DTAProduct.Dispose()
            TBLProduct.Dispose()
            DTAAkun.Dispose()
            TBLAkun.Dispose()
            ListPrimaryAkun = Nothing
        End Try
    End Sub

    Private Sub BackgroundWorker5_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted
        Try
            Me.lblProgress.Text = "Ready"
            Me.prgBar.Value = 0
            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker6_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker6.DoWork
        Dim DTA_P As New CustomSMESS.DseEndStockTableAdapters.TblEndStock1TableAdapter
        Dim TBL_P As New CustomSMESS.DseEndStock.TblEndStock1DataTable
        Dim row_P As CustomSMESS.DseEndStock.TblEndStock1Row
        Dim DTA As New CustomSMESS.DseEndStockTableAdapters.TblEndStockTableAdapter

        Try
            Me.ObjCD.UpdateThreading()

            'load data end stock pensil
            Me.lblProgress.Text = "Loading data"
            DTA_P.Fill(TBL_P)

            'untuk setiap baris, tambahkan ke database slat
            Me.lblProgress.Text = "Inserting data"
            Me.prgBar.Value = 0
            Me.prgBar.Maximum = TBL_P.Rows.Count
            For Each row_P In TBL_P.Rows
                Me.prgBar.Value += 1

                If Me.prgBar.Value > 25 Then
                    If Math.IEEERemainder(Me.prgBar.Value, 100) = 0 Then
                        Me.lblProgress.Text = "Inserting data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString + " - " + row_P.Product_Code
                    ElseIf Me.prgBar.Value = Me.prgBar.Maximum Then
                        Me.lblProgress.Text = "Inserting data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString + " - " + row_P.Product_Code
                    End If
                Else
                    Me.lblProgress.Text = "Inserting data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString + " - " + row_P.Product_Code
                End If

                With row_P
                    If .IsWarehouse_CodeNull = True Then
                        .Warehouse_Code = ""
                        .Warehouse_Desc = ""
                    End If
                    DTA.InsertQuery(.Desc, .No, .Product_Code, .Qty, .Unit, .Qty_M3_STD, .Unit_Price, .Warehouse_Code, .Warehouse_Desc, .Amount, .Acc_Code, .Acc_Desc)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker6_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker6.RunWorkerCompleted
        Try
            Me.lblProgress.Text = "Ready"
            Me.prgBar.Value = 0
            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class