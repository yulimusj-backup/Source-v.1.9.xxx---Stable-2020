Option Explicit On
Option Strict Off

Public Class Form2_Bundling
    'form ini membut bundling per account
    '1. user mendefinisikan: primary account sumber, primary account target, primary location bundling, tanggal bundling1, dan tanggal bundling2
    '   primary account sumber adalah yang akan di nol-kan qtynya. primary account target adalah yang akan dijadikan penampung sementara
    '   tgl bundling 1, semua product dari primary account sumber di bundling ke product dengan primary account target
    '2. memastikan minimal ada 1 code no yang ada productnya pada tanggal bundling1
    '2.b memastikan tidak ada yang minus pada tanggal bundling 1
    '2.c pastikan primary acc sumber, primary acc target, primary loc ada isinya
    '2.d pastikan tgl bundling 2 > tgl bundling 1
    '3. mengumpulkan stock per tgl bundling 1 untuk setiap product yang primary accountnya diberikan
    '4. buat bundling1 pada tgl bundling1
    '5. buat bundling2 pada tgl bundling2

    Dim Doc1, Doc2 As String
    Dim DTABundling As New DataSet2_BundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
    Dim DTAStock As New DataSet2_BundlingTableAdapters.DataTable1TableAdapter
    Dim TipeProgress As EnumTipeProgress = EnumTipeProgress.Msgbox
    Dim MsgboxMessage As String = ""
    Dim MsgboxStyle As Microsoft.VisualBasic.MsgBoxStyle = MsgBoxStyle.Information
    Dim MsgboxTitle As String = ""
    Dim Finish As Boolean = False
    Dim TBLNegatifSumber As New DataSet2_Bundling.DataTable1DataTable
    Dim TBLNegatifTarget As New DataSet2_Bundling.DataTable1DataTable
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjAkun As New Core_Module.ClsMasterAkun1
    Dim ObjLoc As New Core_Module.clsMasterStorageArea

    Enum EnumTipeProgress As Integer
        Msgbox
        LoadStockNegatif
    End Enum

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Dim Primary1, Primary2 As Double

        Try
            Me.Doc1 = "Bundling1"
            Me.Doc2 = "Bundling2"
            Primary1 = -1
            Primary2 = -2

            '2.c pastikan primary acc sumber, primary acc target, primary loc ada isinya
            If Me.TxtPrimaryAccSumber.Text = "-1" Or Me.TxtPrimaryAccSumber.Text = "" Then
                MsgBox("Please specify: account sumber", MsgboxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.TxtPrimaryAccTarget.Text = "-1" Or Me.TxtPrimaryAccTarget.Text = "" Then
                MsgBox("Please specify: account target", MsgboxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.TxtPrimaryLoc.Text = "-1" Or Me.TxtPrimaryLoc.Text = "" Then
                MsgBox("Please specify: location", MsgboxStyle.Information, "Information")
                Exit Sub
            End If

            '2.d pastikan tgl bundling 2 > tgl bundling 1
            If Me.dtpBundling2.Value.Date <= Me.dtpBundling1.Value.Date Then
                MsgBox("Please specify bundling date2 > bundling date1", MsgboxStyle.Information, "Information")
            End If

            '2.e hapus bundling1 dan bundling2 dari tabel bundling dan dari tabel sc
            If DTABundling.ScalarQuery_GetPrimary(Me.Doc1).HasValue = True Then
                Primary1 = DTABundling.ScalarQuery_GetPrimary(Me.Doc1).Value
                Me.HapusBundling(Primary1)
            End If
            If DTABundling.ScalarQuery_GetPrimary(Me.Doc2).HasValue = True Then
                Primary2 = DTABundling.ScalarQuery_GetPrimary(Me.Doc2).Value
                Me.HapusBundling(Primary2)
            End If

            If MsgBox("Start creating bundling?", MsgboxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.KosongkanDGV()

                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HapusBundling(ByVal vPrimary As Double)
        Dim TipeSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Bundling
        Dim DTASC As New DataSet2TableAdapters.tbltransaksi_stock_card1TableAdapter

        Try
            '1. hapus bundling + itemnya
            Me.DTABundling.Delete(vPrimary)
            '2. hapus item sc
            DTASC.Delete(vPrimary, CInt(TipeSC))
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim TBLStockSumber As New DataSet2_Bundling.DataTable1DataTable
        Dim rowStockSumber, rowNegatif As DataSet2_Bundling.DataTable1Row
        Dim TBLStockTarget As New DataSet2_Bundling.DataTable1DataTable
        Dim rowStockTarget As DataSet2_Bundling.DataTable1Row
        Dim AdaNegatif As Boolean = False

        Try
            Me.ObjCD.UpdateThreading()
            Me.Finish = False
            Me.TBLNegatifSumber.Rows.Clear()
            Me.TBLNegatifTarget.Rows.Clear()

            'load data stock
            Me.LblProgress.Text = "Loading data"
            Me.DTAStock.Fill(TBLStockSumber, Me.TxtPrimaryAccSumber.Text, Me.dtpBundling1.Value.Date)
            Me.DTAStock.Fill(TBLStockTarget, Me.TxtPrimaryAccTarget.Text, Me.dtpBundling1.Value.Date)

            '2. memastikan minimal ada 1 code no yang ada productnya pada tanggal bundling1
            Me.LblProgress.Text = "Data checking"
            If TBLStockSumber.Rows.Count = 0 Then
                Me.TipeProgress = EnumTipeProgress.Msgbox
                Me.MsgboxMessage = "No stock on account sumber"
                Me.MsgboxStyle = Microsoft.VisualBasic.MsgBoxStyle.Information
                Me.MsgboxTitle = "Information"
                Me.BackgroundWorker1.ReportProgress(0)
                Exit Sub
            End If
            If TBLStockTarget.Rows.Count = 0 Then
                Me.TipeProgress = EnumTipeProgress.Msgbox
                Me.MsgboxMessage = "No stock on account target"
                Me.MsgboxStyle = Microsoft.VisualBasic.MsgBoxStyle.Information
                Me.MsgboxTitle = "Information"
                Me.BackgroundWorker1.ReportProgress(0)
                Exit Sub
            End If

            '2.b memastikan tidak ada yang minus pada tanggal bundling 1
            For Each rowStockSumber In TBLStockSumber.Rows
                If rowStockSumber.Qty <= 0 Then
                    rowNegatif = Me.TBLNegatifSumber.NewDataTable1Row
                    For gh As Integer = 0 To Me.TBLNegatifSumber.Columns.Count - 1
                        rowNegatif.Item(gh) = rowStockSumber.Item(gh)
                    Next
                    Me.TBLNegatifSumber.Rows.Add(rowNegatif)
                    AdaNegatif = True
                End If
            Next
            For Each rowStockTarget In TBLStockTarget.Rows
                If rowStockTarget.Qty <= 0 Then
                    rowNegatif = Me.TBLNegatifTarget.NewDataTable1Row
                    For gh As Integer = 0 To Me.TBLNegatifTarget.Columns.Count - 1
                        rowNegatif.Item(gh) = rowStockTarget.Item(gh)
                    Next
                    Me.TBLNegatifTarget.Rows.Add(rowNegatif)
                    AdaNegatif = True
                End If
            Next
            If AdaNegatif = True Then
                Me.TipeProgress = EnumTipeProgress.LoadStockNegatif
                Me.BackgroundWorker1.ReportProgress(0)
                Exit Sub
            End If

            '2.c klo item stock sumber + item stock target > 10000, maka harus pikirkan metode lain, karena akan bermasalah menentukan % costnya
            ' klo persentase cost minimum adalah 0.01%, maka total item after maksimum adalah 100%.
            If TBLStockSumber.Rows.Count + TBLStockTarget.Rows.Count > 10000 Then
                Me.TipeProgress = EnumTipeProgress.Msgbox
                Me.MsgboxMessage = "Item stock target + item stock sumber > 10000 item, gunakan metode lainnya"
                Me.MsgboxStyle = Microsoft.VisualBasic.MsgBoxStyle.Information
                Me.MsgboxTitle = "Information"
                Me.TipeProgress = EnumTipeProgress.Msgbox
                Me.BackgroundWorker1.ReportProgress(0)
                Exit Sub
            End If

            '4. buat bundling1 pada tgl bundling1
            Me.LblProgress.Text = "Creating bundling1"
            Me.BuatBundling1(TBLStockSumber, TBLStockTarget)
            '5. buat bundling2 pada tgl bundling2
            Me.LblProgress.Text = "Creating bundling2"
            Me.BuatBundling2(TBLStockSumber, TBLStockTarget)

            'sampe sini, tandain bahwa proses sudah selesai dengan baik
            Me.Finish = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Sub BuatBundling1(ByVal vTBLStockSumber As DataSet2_Bundling.DataTable1DataTable, ByVal vTBLStockTarget As DataSet2_Bundling.DataTable1DataTable)
        Dim vTBLHeader As New DataSet2_Bundling.tbltransaksi_stockbundlingDataTable
        Dim rowHeader As DataSet2_Bundling.tbltransaksi_stockbundlingRow
        Dim vTBLItemAfter As New DataSet2_Bundling.tbltransaksi_stockbundling_item_afterDataTable
        Dim vTBLItemBefore As New DataSet2_Bundling.tbltransaksi_stockbundling_item_beforeDataTable

        Try
            'isikan data header
            rowHeader = vTBLHeader.Newtbltransaksi_stockbundlingRow
            With rowHeader
                .Created_By = "USER"
                .Created_Date = Now
                .down_time_actual = 0
                .end_time_actual = #1/1/1950#
                .end_time_planned = #1/1/1950#
                .good_percentage = 0
                .identifier_jo = ""
                .identifier_wo = ""
                .Modified_By = "USER"
                .Modified_Date = Now
                .No_Transaksi = Me.Doc1
                .Primary = -1
                .primary_mo = -1
                .primary_process = -1
                .primary_work_unit = -1
                .primary_workshift = -1
                .prod__index_actual = 0
                .prod__index_planned = 0
                .Remarks = ""
                .setup_time_actual = 0
                .setup_time_planned = 0
                .start_time_actual = #1/1/1950#
                .start_time_planned = #1/1/1950#
                .Tanggal_Transaksi = Me.dtpBundling1.Value.Date
                .Timestamp = Now
                .Tipe_Bundling = 1
                .total_good = 0
                .total_output = 0
                .total_produce = 0
                .total_reject = 0
                .Urutan = 0
                .work_hour_actual = 0
                .work_hour_planned = 0
            End With
            vTBLHeader.Rows.Add(rowHeader)

            'isikan data item after
            Me.TambahkanItemAfter(vTBLStockTarget, vTBLItemAfter)
            Me.SetCost(vTBLItemAfter)

            'isikan data item before
            Me.TambahkanItemBefore(vTBLStockSumber, vTBLItemBefore)
            Me.TambahkanItemBefore(vTBLStockTarget, vTBLItemBefore)

            'simpan data bundling
            Me.SimpanBundling(vTBLHeader, vTBLItemAfter, vTBLItemBefore)
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Sub SetCost(ByVal vTBLItemAfter As DataSet2_Bundling.tbltransaksi_stockbundling_item_afterDataTable)
        Dim rowItemAfter As DataSet2_Bundling.tbltransaksi_stockbundling_item_afterRow
        Dim Sisa As Decimal = 100

        Try
            For ij As Double = 0 To vTBLItemAfter.Rows.Count - 1
                rowItemAfter = vTBLItemAfter.Rows(ij)
                If ij = vTBLItemAfter.Rows.Count - 1 Then
                    rowItemAfter.Faktor_Pemberat = Sisa
                Else
                    rowItemAfter.Faktor_Pemberat = 0.01
                    Sisa -= 0.01
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Sub TambahkanItemAfter(ByVal vTBLStock As DataSet2_Bundling.DataTable1DataTable, _
    ByVal vTBLItemAfter As DataSet2_Bundling.tbltransaksi_stockbundling_item_afterDataTable)
        Dim rowItemAfter As DataSet2_Bundling.tbltransaksi_stockbundling_item_afterRow
        Dim rowStock As DataSet2_Bundling.DataTable1Row

        Try
            For Each rowStock In vTBLStock.Rows
                rowItemAfter = vTBLItemAfter.Newtbltransaksi_stockbundling_item_afterRow
                With rowItemAfter
                    If rowStock.Batch_No_Controlled = 0 Then
                        .Batch_No = ""
                    Else
                        .Batch_No = "CLOSING"
                    End If
                    If rowStock.Expiry_Date_Controlled = False Then
                        .Expiry_Date = #1/1/1950#
                    Else
                        .Expiry_Date = #1/1/2020#
                    End If
                    .Faktor_Pemberat = 0
                    .Primary = -1
                    .Primary_Product = rowStock.Primary_Product
                    .Primary_Secondary_Units = -1
                    If rowStock.Location_Controlled = 0 Then
                        .Primary_Storage_Area = -1
                    Else
                        .Primary_Storage_Area = Me.TxtPrimaryLoc.Text
                    End If
                    .Primary_Transaksi = -1
                    .Primary_Units = rowStock.Primary_Unit
                    .Qty = rowStock.Qty
                    .Remarks = ""
                    .Secondary_Qty = -1
                End With
                vTBLItemAfter.Rows.Add(rowItemAfter)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Sub TambahkanItemBefore(ByVal vTBLStock As DataSet2_Bundling.DataTable1DataTable, _
    ByVal vTBLItemBefore As DataSet2_Bundling.tbltransaksi_stockbundling_item_beforeDataTable)
        Dim rowItemBefore As DataSet2_Bundling.tbltransaksi_stockbundling_item_beforeRow
        Dim rowStock As DataSet2_Bundling.DataTable1Row

        Try
            For Each rowStock In vTBLStock.Rows
                rowItemBefore = vTBLItemBefore.Newtbltransaksi_stockbundling_item_beforeRow
                With rowItemBefore
                    If rowStock.Batch_No_Controlled = 0 Then
                        .Batch_No = ""
                    Else
                        .Batch_No = "CLOSING"
                    End If
                    If rowStock.Expiry_Date_Controlled = False Then
                        .Expiry_Date = #1/1/1950#
                    Else
                        .Expiry_Date = #1/1/2020#
                    End If
                    .Primary = -1
                    .Primary_Product = rowStock.Primary_Product
                    .Primary_Secondary_Units = -1
                    If rowStock.Location_Controlled = 0 Then
                        .Primary_Storage_Area = -1
                    Else
                        .Primary_Storage_Area = Me.TxtPrimaryLoc.Text
                    End If
                    .Primary_Transaksi = -1
                    .Primary_Units = rowStock.Primary_Unit
                    .Qty = rowStock.Qty
                    .Remarks = ""
                    .Secondary_Qty = -1
                End With
                vTBLItemBefore.Rows.Add(rowItemBefore)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Sub BuatBundling2(ByVal vTBLStockSumber As DataSet2_Bundling.DataTable1DataTable, ByVal vTBLStockTarget As DataSet2_Bundling.DataTable1DataTable)
        Dim vTBLHeader As New DataSet2_Bundling.tbltransaksi_stockbundlingDataTable
        Dim rowHeader As DataSet2_Bundling.tbltransaksi_stockbundlingRow
        Dim vTBLItemAfter As New DataSet2_Bundling.tbltransaksi_stockbundling_item_afterDataTable
        Dim vTBLItemBefore As New DataSet2_Bundling.tbltransaksi_stockbundling_item_beforeDataTable

        Try
            'isikan data header
            rowHeader = vTBLHeader.Newtbltransaksi_stockbundlingRow
            With rowHeader
                .Created_By = "USER"
                .Created_Date = Now
                .down_time_actual = 0
                .end_time_actual = #1/1/1950#
                .end_time_planned = #1/1/1950#
                .good_percentage = 0
                .identifier_jo = ""
                .identifier_wo = ""
                .Modified_By = "USER"
                .Modified_Date = Now
                .No_Transaksi = Me.Doc2
                .Primary = -1
                .primary_mo = -1
                .primary_process = -1
                .primary_work_unit = -1
                .primary_workshift = -1
                .prod__index_actual = 0
                .prod__index_planned = 0
                .Remarks = ""
                .setup_time_actual = 0
                .setup_time_planned = 0
                .start_time_actual = #1/1/1950#
                .start_time_planned = #1/1/1950#
                .Tanggal_Transaksi = Me.dtpBundling2.Value.Date
                .Timestamp = Now
                .Tipe_Bundling = 1
                .total_good = 0
                .total_output = 0
                .total_produce = 0
                .total_reject = 0
                .Urutan = 0
                .work_hour_actual = 0
                .work_hour_planned = 0
            End With
            vTBLHeader.Rows.Add(rowHeader)

            'isikan data item after
            Me.TambahkanItemAfter(vTBLStockTarget, vTBLItemAfter)
            Me.TambahkanItemAfter(vTBLStockSumber, vTBLItemAfter)
            Me.SetCost(vTBLItemAfter)

            'isikan data item before
            Me.TambahkanItemBefore(vTBLStockTarget, vTBLItemBefore)

            'simpan data bundling
            Me.SimpanBundling(vTBLHeader, vTBLItemAfter, vTBLItemBefore)
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Sub SimpanBundling(ByVal vTBLHeader As DataSet2_Bundling.tbltransaksi_stockbundlingDataTable, ByVal vTBLItemAfter As DataSet2_Bundling.tbltransaksi_stockbundling_item_afterDataTable, _
    ByVal vTBLItemBefore As DataSet2_Bundling.tbltransaksi_stockbundling_item_beforeDataTable)
        Dim RowHeader As DataSet2_Bundling.tbltransaksi_stockbundlingRow
        Dim PrimaryTrans As Double = -1

        Try
            RowHeader = vTBLHeader.Rows(0)
            With RowHeader
                Me.DTABundling.Insert(.No_Transaksi, .Urutan, .Tanggal_Transaksi, .Remarks, .Created_By, .Created_Date, .Modified_By, .Modified_Date, _
                .Timestamp, .Tipe_Bundling, .primary_work_unit, .primary_process, .start_time_planned, .end_time_planned, _
                .setup_time_planned, .work_hour_planned, .prod__index_planned, .start_time_actual, .end_time_actual, .setup_time_actual, _
                .work_hour_actual, .down_time_actual, .prod__index_actual, .total_produce, .total_good, .total_output, .total_reject, _
                .good_percentage, .primary_workshift, .identifier_wo, .identifier_jo, .primary_mo)

                PrimaryTrans = DTABundling.ScalarQuery_GetPrimary(.No_Transaksi).Value
                Me.Simpan_ItemBefore_UsingScript(vTBLItemBefore, PrimaryTrans, .No_Transaksi, .Tanggal_Transaksi)
                Me.Simpan_ItemAfter_UsingScript(vTBLItemAfter, PrimaryTrans, .No_Transaksi, .Tanggal_Transaksi)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Simpan_ItemBefore_UsingScript(ByVal vTBL As DataSet2_Bundling.tbltransaksi_stockbundling_item_beforeDataTable, ByVal vPrimaryTrans As Double, _
    ByVal vDoc As String, ByVal vTgl As Date) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = vTBL.Rows.Count

            While StartRowIndex < vTBL.Rows.Count
                Me.ProgressBar1.Value = StartRowIndex
                Me.LblProgress.Text = "Saving on database - data: " + StartRowIndex.ToString + " of " + vTBL.Rows.Count.ToString

                sql = Me.GetInsertString_ItemBefore(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
                script.Query = sql
                'execute script
                script.Execute()

                sql = Me.GetInsertStringSC_ItemBefore(StartRowIndex, BatchQty, vTBL, vPrimaryTrans, ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Bundling, _
                vDoc, ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out, vTgl)
                script.Query = sql
                'execute script
                script.Execute()

                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Me.LblProgress.Text = ""
            Me.ProgressBar1.Value = 0

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertStringSC_ItemBefore(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
   ByVal vTBL As DataSet2_Bundling.tbltransaksi_stockbundling_item_beforeDataTable, ByVal vPrimaryTrans As Double, ByVal vTipeSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal, _
   ByVal vDocNo As String, ByVal vInOut As ClsTransaksiStockCard.EnumJenisStockCard, ByVal vTglTrans As Date) As String
        Dim retval As String = ""
        Dim row As DataSet2_Bundling.tbltransaksi_stockbundling_item_beforeRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltransaksi_stock_card` (`Primary Transaksi Asal`,`Type Transaksi Asal`,`Primary Product`,`Document No`,`In`,`Out`,`Description`,`InOut`" + _
            ",`Date`,`Expiry`,`primary location`,`Batch Number`,`Remarks`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTrans).ToString + "," + CInt(vTipeSC).ToString + "," + CDbl(.Primary_Product).ToString + ",'" + vDocNo + "'," _
                            + "0," + CDec(Math.Abs(.Qty)).ToString + ",'Stock Bundling'," + CInt(vInOut).ToString + ",'" + Format(vTglTrans, "yyyy-MM-dd") + "'," + "'" + _
                            Format(CDate(.Expiry_Date), "yyyy-MM-dd") + "'," + CDbl(.Primary_Storage_Area).ToString + "," + "'" + Me.ObjCD.FormatSpecialChar_MySQL(.Batch_No) + "'," _
                            + "''" + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function GetInsertStringSC_ItemAfter(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
  ByVal vTBL As DataSet2_Bundling.tbltransaksi_stockbundling_item_afterDataTable, ByVal vPrimaryTrans As Double, ByVal vTipeSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal, _
  ByVal vDocNo As String, ByVal vInOut As ClsTransaksiStockCard.EnumJenisStockCard, ByVal vTglTrans As Date) As String
        Dim retval As String = ""
        Dim row As DataSet2_Bundling.tbltransaksi_stockbundling_item_afterRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltransaksi_stock_card` (`Primary Transaksi Asal`,`Type Transaksi Asal`,`Primary Product`,`Document No`,`In`,`Out`,`Description`,`InOut`" + _
            ",`Date`,`Expiry`,`primary location`,`Batch Number`,`Remarks`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTrans).ToString + "," + CInt(vTipeSC).ToString + "," + CDbl(.Primary_Product).ToString + ",'" + vDocNo + "'," _
                            + CDec(Math.Abs(.Qty)).ToString + ",0" + ",'Stock Bundling'," + CInt(vInOut).ToString + ",'" + Format(vTglTrans, "yyyy-MM-dd") + "'," + "'" + _
                            Format(CDate(.Expiry_Date), "yyyy-MM-dd") + "'," + CDbl(.Primary_Storage_Area).ToString + "," + "'" + Me.ObjCD.FormatSpecialChar_MySQL(.Batch_No) + "'," _
                            + "''" + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemBefore(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DataSet2_Bundling.tbltransaksi_stockbundling_item_beforeDataTable, ByVal vPrimaryTrans As Double) As String
        Dim retval As String = ""
        Dim row As DataSet2_Bundling.tbltransaksi_stockbundling_item_beforeRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltransaksi_stockbundling_item_before` (`Primary Transaksi`,`Primary Product`,`Qty`,`Primary Units`," + _
            "`Secondary Qty`,`Primary Secondary Units`,`Batch No`,`Expiry Date`,`Primary Storage Area`,`Remarks`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + vPrimaryTrans.ToString + "," + CDbl(.Primary_Product).ToString + "," _
                        + CDec(.Qty).ToString + "," + CDbl(.Primary_Units).ToString + ",-1,-1,'" + Me.ObjCD.FormatSpecialChar_MySQL(.Batch_No) + "'," _
                        + "'" + Format(CDate(.Expiry_Date), "yyyy-MM-dd") + "'," + CDbl(.Primary_Storage_Area).ToString + ",'')" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function Simpan_ItemAfter_UsingScript(ByVal vTBL As DataSet2_Bundling.tbltransaksi_stockbundling_item_afterDataTable, ByVal vPrimaryTrans As Double, _
    ByVal vDoc As String, ByVal vTgl As Date) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = vTBL.Rows.Count

            While StartRowIndex < vTBL.Rows.Count
                Me.ProgressBar1.Value = StartRowIndex
                Me.LblProgress.Text = "Saving on database - data: " + StartRowIndex.ToString + " of " + vTBL.Rows.Count.ToString

                sql = Me.GetInsertString_ItemAfter(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
                script.Query = sql
                'execute script
                script.Execute()

                sql = Me.GetInsertStringSC_ItemAfter(StartRowIndex, BatchQty, vTBL, vPrimaryTrans, ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Bundling, _
                vDoc, ClsTransaksiStockCard.EnumJenisStockCard.Stock_In, vTgl)
                script.Query = sql
                'execute script
                script.Execute()

                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Me.LblProgress.Text = ""
            Me.ProgressBar1.Value = 0

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemAfter(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DataSet2_Bundling.tbltransaksi_stockbundling_item_afterDataTable, ByVal vPrimaryTrans As Double) As String
        Dim retval As String = ""
        Dim row As DataSet2_Bundling.tbltransaksi_stockbundling_item_afterRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltransaksi_stockbundling_item_after` (`Primary Transaksi`,`Primary Product`,`Qty`,`Primary Units`,`Secondary Qty`,`Primary Secondary Units`," + _
            "`Batch No`,`Expiry Date`,`Primary Storage Area`,`Remarks`,`Faktor Pemberat`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + vPrimaryTrans.ToString + "," + CDbl(.Primary_Product).ToString + "," _
                        + CDec(.Qty).ToString + "," + CDbl(.Primary_Units).ToString + ",-1,-1,'" + Me.ObjCD.FormatSpecialChar_MySQL(.Batch_No) + "'," _
                        + "'" + Format(CDate(.Expiry_Date), "yyyy-MM-dd") + "'," + CDbl(.Primary_Storage_Area).ToString + ",''," + CDec(.Faktor_Pemberat).ToString + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            If Me.TipeProgress = EnumTipeProgress.Msgbox Then
                Me.ShowMsgbox()
            ElseIf Me.TipeProgress = EnumTipeProgress.LoadStockNegatif Then
                Me.LoadStockNegatif()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Sub KosongkanDGV()
        Try
            Me.dgvSumber.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvSumber.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvSumber.Rows.Clear()
            Me.dgvSumber.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvSumber.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvTarget.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvTarget.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvTarget.Rows.Clear()
            Me.dgvTarget.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvTarget.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Sub LoadStockNegatif()
        Dim rowDGV As DataGridViewRow
        Dim rowNegatif As DataSet2_Bundling.DataTable1Row

        Try
            Me.KosongkanDGV()
            'stock sumber
            Me.dgvSumber.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvSumber.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each rowNegatif In Me.TBLNegatifSumber.Rows
                Me.dgvSumber.Rows.Add()
                rowDGV = Me.dgvSumber.Rows(Me.dgvSumber.Rows.Count - 1)
                rowDGV.Cells(Me.ColSumber_CodeNo.Index).Value = rowNegatif.Product_Code
                rowDGV.Cells(Me.ColSumber_Desc.Index).Value = rowNegatif.Description
                rowDGV.Cells(Me.ColSumber_No.Index).Value = Me.dgvSumber.Rows.Count
                rowDGV.Cells(Me.ColSumber_Qty.Index).Value = Me.ObjCD.DigitGrouping(rowNegatif.Qty, rowNegatif._Decimal)
                rowDGV.Cells(Me.ColSumber_Unit.Index).Value = rowNegatif.Unit_Code
            Next
            Me.dgvSumber.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvSumber.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            'stock target
            Me.dgvTarget.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvTarget.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each rowNegatif In Me.TBLNegatifTarget.Rows
                Me.dgvTarget.Rows.Add()
                rowDGV = Me.dgvTarget.Rows(Me.dgvTarget.Rows.Count - 1)
                rowDGV.Cells(Me.ColTarget_CodeNo.Index).Value = rowNegatif.Product_Code
                rowDGV.Cells(Me.ColTarget_Desc.Index).Value = rowNegatif.Description
                rowDGV.Cells(Me.ColTarget_No.Index).Value = Me.dgvTarget.Rows.Count
                rowDGV.Cells(Me.ColTarget_Qty.Index).Value = Me.ObjCD.DigitGrouping(rowNegatif.Qty, rowNegatif._Decimal)
                rowDGV.Cells(Me.ColTarget_Unit.Index).Value = rowNegatif.Unit_Code
            Next
            Me.dgvTarget.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvTarget.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            MsgBox("Negative stock detected, please fix first.", Microsoft.VisualBasic.MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Sub ShowMsgbox()
        Try
            MsgBox(Me.MsgboxMessage, Me.MsgboxStyle, Me.MsgboxTitle)
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.Finish = True Then
                MsgBox("Bundling completed. The document no is: " + Me.Doc1 + " and " + Me.Doc2, MsgboxStyle.Information, "Information")
                Me.ProgressBar1.Value = 0
                Me.LblProgress.Text = "Done"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Form2_Bundling_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.TxtPrimaryAccSumber.Text = "423"
            Me.TxtAccSumberCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryAccSumber.Text)
            Me.TxtAccSumberDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccSumber.Text)
            Me.TxtPrimaryAccTarget.Text = "428"
            Me.TxtAccTargetCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryAccTarget.Text)
            Me.TxtAccTargetDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccTarget.Text)
            Me.TxtPrimaryLoc.Text = "1"
            Me.TxtLocCode.Text = Me.ObjLoc.GetStorageAreaCode(Me.TxtPrimaryLoc.Text)
            Me.TxtLocDesc.Text = Me.ObjLoc.GetDescription(Me.TxtPrimaryLoc.Text)
            Me.dtpBundling1.Value = #12/30/2016#
            Me.dtpBundling2.Value = #1/1/2017#
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDeleteBundling_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteBundling.Click
        Dim Primary1, Primary2 As Double

        Try
            If MsgBox("Delete bundling?", Microsoft.VisualBasic.MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.Doc1 = "Bundling1"
                Me.Doc2 = "Bundling2"
                Primary1 = -1
                Primary2 = -2

                '2.e hapus bundling1 dan bundling2 dari tabel bundling dan dari tabel sc
                If DTABundling.ScalarQuery_GetPrimary(Me.Doc1).HasValue = True Then
                    Primary1 = DTABundling.ScalarQuery_GetPrimary(Me.Doc1).Value
                    Me.HapusBundling(Primary1)
                End If
                If DTABundling.ScalarQuery_GetPrimary(Me.Doc2).HasValue = True Then
                    Primary2 = DTABundling.ScalarQuery_GetPrimary(Me.Doc2).Value
                    Me.HapusBundling(Primary2)
                End If

                MsgBox("Bundling deleted", Microsoft.VisualBasic.MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDeleteStock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteStock.Click
        Dim JmlDelete As Double = 0
        Dim DTA As New DataSet2_BundlingTableAdapters.tbltrans_custom_stockTableAdapter
        Dim TBL As New DataSet2_Bundling.tbltrans_custom_stockDataTable
        Dim row As DataSet2_Bundling.tbltrans_custom_stockRow

        Try
            If MsgBox("Delete stock item? Account: " + Me.TxtAccSumberCode.Text, Microsoft.VisualBasic.MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                DTA.Fill(TBL, Me.TxtPrimaryAccSumber.Text)
                For Each row In TBL.Rows
                    DTA.Delete(row.Primary)
                    JmlDelete += 1
                Next

                MsgBox("Deleted: " + Me.ObjCD.BuangNolDibelakangDecimal(JmlDelete), Microsoft.VisualBasic.MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgboxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindAccSumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAccSumber.Click

    End Sub
End Class