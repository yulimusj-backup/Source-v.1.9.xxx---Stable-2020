Option Explicit On
Option Strict Off

Public Class FrmTransaksiStockCard2_UpdateMonthlyClosing

    Dim ObjCD As New Core_Module.ClsCekData
    Dim Succeed As Boolean

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If MsgBox("Are you sure?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                Exit Sub
            End If
            FrmTransaksiStockCard2_UpdateMonthlyClosing.CheckForIllegalCrossThreadCalls = False
            Me.Cursor = Cursors.WaitCursor
            Me.Button1.Enabled = False
            Me.ControlBox = False
            Me.Succeed = False
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim DtaMonthly As New DseTransaksiStockCard2_UpdateMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim TblMonthly As New DseTransaksiStockCard2_UpdateMonthlyClosing.tbltrans_purchase_closingDataTable
        Dim RowMonthly As DseTransaksiStockCard2_UpdateMonthlyClosing.tbltrans_purchase_closingRow

        Try
            Me.ObjCD.UpdateThreading()
            DtaMonthly.Connection.ConnectionString = Core_Module.GetConstring
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = DtaMonthly.Fill(TblMonthly)
            For Each RowMonthly In TblMonthly
                Me.ProgressBar1.Value += 1
                Me.Label1.Text = "Loading data " & Me.ProgressBar1.Value & " of " & Me.ProgressBar1.Maximum & ", document no: " & RowMonthly.No_Transaksi & "..."
                Dim TblProductQty As New Inventory_Module.DseLaporanStockDetail2.TblCetakDataTable
                Dim FrmLaporanStockDetail2_ As New FrmLaporanStockDetail2
                FrmLaporanStockDetail2_.LoadForMonthlyClosing = True
                DtaMonthly.DeleteQuery(RowMonthly.Primary)
                DtaMonthly.UpdateQuery(0, RowMonthly.Primary)
                If FrmLaporanStockDetail2_.FromCutOff_BeginingBalance(TblProductQty, Me.ObjCD.GetLasDateOfTheMonth(RowMonthly.Closing_Period)) Then
                    Me.Label1.Text = "Saving data " & Me.ProgressBar1.Value & " of " & Me.ProgressBar1.Maximum & ", document no: " & RowMonthly.No_Transaksi & "..."
                    If Me.SimpanUsingScript_ItemProductQty(TblProductQty, RowMonthly.Primary) Then
                        DtaMonthly.UpdateQuery(1, RowMonthly.Primary)
                    End If
                End If
            Next
            Succeed = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function SimpanUsingScript_ItemProductQty(ByVal vTBL As DseLaporanStockDetail2.TblCetakDataTable, _
            ByVal vPrimaryTransaksi As Double) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemProductQty(StartRowIndex, BatchQty, vTBL, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemProductQty(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vTBLItemDetail As DseLaporanStockDetail2.TblCetakDataTable, _
            ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As Inventory_Module.DseLaporanStockDetail2.TblCetakRow

        Try
            retval = "INSERT INTO `tbltrans_purchase_closing_item_productqty` (`Primary Transaksi`,`Primary Product`,`Qty`,`Lot No`,`Expiry Date`,`Primary Storage Area`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBLItemDetail.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBLItemDetail.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTBLItemDetail.Rows(vStartRowIndex + i)
                    If row.Lot_No <> "" Then
                        If row.Lot_No = "-" Then
                            row.Lot_No = ""
                        End If
                        If row.Lot_No.Contains("\") = True Then
                            row.Lot_No = row.Lot_No.Replace("\", "\\")
                        End If
                        If row.Lot_No.Contains(vbCrLf) = True Then
                            row.Lot_No = row.Lot_No.Replace(vbCrLf, "\r\n")
                        End If
                        If row.Lot_No.Contains("'") = True Then
                            row.Lot_No = row.Lot_No.Replace("'", "''")
                        End If
                    End If
                    retval = retval & vbCrLf & "(" & vPrimaryTransaksi.ToString & "," _
                    & CDbl(row.Primary_Product).ToString & "," _
                    & CDec(row.QtyDec).ToString & "," _
                    & "'" & row.Lot_No & "'," _
                    & "'" & Format(row.Expiry_Date_Date, "yyyy-MM-dd") & "'" & "," _
                    & CDbl(row.Primary_Location).ToString & ")" & Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.Cursor = Cursors.Default
            Me.Button1.Enabled = True
            Me.ControlBox = True
            Me.Label1.Text = "Ready..."
            Me.ProgressBar1.Value = 0
            If Me.Succeed Then
                MsgBox("Finished", MsgBoxStyle.Information, "Information")
            Else
                MsgBox("Something went wrong while update", MsgBoxStyle.Critical, "Error")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class