Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_SF_SetDeliveryPurcReq
    Dim LangsungTutup As Boolean = False
    Dim vrowItemDelivery As DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryRow
    Dim vTBLItemDelivery As New DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryDataTable
    Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
    Dim DefPrec As Integer = 0
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjPriorityLevel As New Core_Module.ClsMasterPriorityLevel
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Buffer As String = ""
    Dim mDeleteIndex As Integer = -1
    Dim mCekPriorityIndex As Integer = 0
    Dim ObjPurcReq As New ClsTransaksiPR

    Public vRowPurcReqBRef As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
    Public BAckColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public VIdentifier As String = ""
    Public vPrimaryCodeNo As Double = -1
    Public vReadOnly As Boolean = False

    Private Sub FrmTrans_MaterialPlanningDashboard_SF_SetDeliveryPurcReq_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.vReadOnly = False Then
                If Me.LangsungTutup = False Then
                    If MsgBox("Close dan discard data changes?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_SF_SetDeliveryPurcReq_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As DataGridViewRow
        Dim ListPriorityLevel As New List(Of String)

        Try
            Me.DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryCodeNo)
            Me.ObjPriorityLevel.GetListDescPriorityLevel(ListPriorityLevel)
            Me.ColPriorityLevel.Items.Clear()
            For Each Desc As String In ListPriorityLevel
                Me.ColPriorityLevel.Items.Add(Desc)
            Next

            Me.dgvDetail.AllowUserToAddRows = False
            Me.vTBLItemDelivery.Rows.Clear()
            Me.DTA.FillBy_ASC(Me.vTBLItemDelivery, Me.VIdentifier)
            If Me.vTBLItemDelivery.Rows.Count > 0 Then
                For Each Me.vrowItemDelivery In Me.vTBLItemDelivery.Rows
                    If vrowItemDelivery.identifier = Me.VIdentifier Then
                        Me.dgvDetail.Rows.Add()
                        rowDGV = Me.dgvDetail.Rows(Me.dgvDetail.Rows.Count - 1)
                        rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvDetail.Rows.Count
                        rowDGV.Cells(Me.ColPrimaryPriority.Index).Value = vrowItemDelivery.delivery_priority
                        rowDGV.Cells(Me.ColPriorityLevel.Index).Value = Me.ObjPriorityLevel.GetDesc(vrowItemDelivery.delivery_priority)
                        rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(vrowItemDelivery.delivery_qty, Me.DefPrec)
                        rowDGV.Cells(Me.ColRemarks.Index).Value = vrowItemDelivery.remarks
                        rowDGV.Cells(Me.ColDeliveryDate.Index).Value = Format(vrowItemDelivery.delivery_date, "dd MMM yyyy")
                    End If
                Next
            Else
                'tidak ada item delivery
                Me.dgvDetail.Rows.Add()
                rowDGV = Me.dgvDetail.Rows(Me.dgvDetail.Rows.Count - 1)
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvDetail.Rows.Count
                rowDGV.Cells(Me.ColPrimaryPriority.Index).Value = Me.vRowPurcReqBRef.primary_priority_level
                rowDGV.Cells(Me.ColPriorityLevel.Index).Value = Me.ObjPriorityLevel.GetDesc(Me.vRowPurcReqBRef.primary_priority_level)
                rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjPurcReq.GetJmlByIdentifier(Me.VIdentifier), Me.DefPrec)
                rowDGV.Cells(Me.ColRemarks.Index).Value = ""
                rowDGV.Cells(Me.ColDeliveryDate.Index).Value = Format(Me.vRowPurcReqBRef.promise_to_deliver, "dd MMM yyyy")
            End If

            If Me.vReadOnly = False Then
                Me.dgvDetail.AllowUserToAddRows = True
            Else
                Me.dgvDetail.AllowUserToDeleteRows = False
                Me.dgvDetail.AllowUserToAddRows = False
                Me.dgvDetail.ReadOnly = True
                Me.btnOK.Visible = False
                Me.btnOK.TabStop = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPriorityLevel = Nothing
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim JmlOrder, JmlDGV As Decimal

        Try
            'pastikan jumlah pengiriman totalnya sama dengan jumlah dgv
            JmlOrder = Me.ObjPurcReq.GetJmlByIdentifier(Me.VIdentifier)
            JmlDGV = Me.GetJmlTotal
            If JmlOrder <> JmlDGV Then
                MsgBox("Inconsistent delivery qty. Total delivery qty must be: " + Me.ObjCD.DigitGrouping(JmlOrder, Me.DefPrec), MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'pastikan qty tiap baris >0
            For Each rowdgv As DataGridViewRow In Me.dgvDetail.Rows
                If CDec(rowdgv.Cells(Me.ColQty.Index).Value) <= 0 And rowdgv.Index <> Me.dgvDetail.RowCount - 1 Then
                    MsgBox("Please specify positive qty data for row: " + (rowdgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            Next

            If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.LangsungTutup = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If Me.vReadOnly = False Then
                If MsgBox("Close dan discard data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                    LangsungTutup = True
                    Me.Close()
                End If
            Else
                LangsungTutup = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDetail_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvDetail.CellBeginEdit
        Try
            If Me.vReadOnly = True Then
                e.Cancel = True
            End If
            If e.RowIndex = Me.dgvDetail.Rows.Count - 1 Then
                'kalo baris terakhir, maka yang boleh di edit hanya date saja
                If e.ColumnIndex = Me.ColDeliveryDate.Index Then
                    Me.Buffer = Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColDeliveryDate.Index Or _
                e.ColumnIndex = Me.ColPriorityLevel.Index Or _
                e.ColumnIndex = Me.ColQty.Index Or _
                e.ColumnIndex = Me.ColRemarks.Index Then
                    Me.Buffer = Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJmlTotal() As Decimal
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV As DataGridViewRow In Me.dgvDetail.Rows
                RetVal += CDec(rowDGV.Cells(Me.ColQty.Index).Value)
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvDetail_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDetail.CellEndEdit
        Dim JmlOrder As Decimal = 0
        Dim JmlDGV As Decimal = 0
        Dim JmlMax As Decimal = 0
        Dim IndexDikurangi As Integer = 0

        Try
            If IsNothing(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                        Or Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    'Remarks
                    Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Exit Sub
                Else
                    'untuk kolom debet kredit, langsung di balikin ke nilai semula
                    Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            Else
                If Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                    Exit Sub
                End If

                If e.ColumnIndex = Me.ColRemarks.Index Then
                    'pastikan lengthnya tidak lebih dari 250 karakter
                    If Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                ElseIf e.ColumnIndex = Me.ColQty.Index Then
                    'pastikan nilai qty >0
                    If Me.ObjCD.CekBilanganPositif(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify positive number.", MsgBoxStyle.Information, "Information")
                        Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'digit grouping sesuai defprec
                    Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                ElseIf e.ColumnIndex = Me.ColPriorityLevel.Index Then
                    Me.mCekPriorityIndex = e.RowIndex
                    Me.BackgroundWorker2.RunWorkerAsync()
                ElseIf e.ColumnIndex = Me.ColDeliveryDate.Index Then
                    '0. pastikan formatnya tanggal
                    If IsDate(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify a valid date.", MsgBoxStyle.Information, "Information")
                        Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        If e.RowIndex = Me.dgvDetail.Rows.Count - 2 And e.RowIndex <> 0 Then
                            'jika data baru, hapuskan baris ybs
                            Me.mDeleteIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                        Exit Sub
                    End If
                    '1. pastikan tanggalnya lebih besar dari tanggal baris di atasnya
                    If e.RowIndex > 0 Then
                        If CDate(Me.dgvDetail.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value) <= CDate(Me.dgvDetail.Item(Me.ColDeliveryDate.Index, e.RowIndex - 1).Value) Then
                            MsgBox("Minimum delivery date is: " + Format(CDate(Me.dgvDetail.Item(Me.ColDeliveryDate.Index, e.RowIndex - 1).Value) + TimeSpan.FromDays(1), "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If e.RowIndex = Me.dgvDetail.Rows.Count - 2 Then
                                'jika data baru, hapuskan baris ybs
                                Me.mDeleteIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Exit Sub
                        End If
                    Else
                        '2. pastikan tanggalnya minimal sama dengan tanggal so
                        If Me.ObjPurcReq.GetTanggalTransaksi_ByIdentifier(Me.VIdentifier) > CDate(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            If MsgBox("Delivery date is smaller than pr date, continue?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                                Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                If e.RowIndex = Me.dgvDetail.Rows.Count - 2 And e.RowIndex <> 0 Then
                                    'jika data baru, hapuskan baris ybs
                                    Me.mDeleteIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                    If e.RowIndex < Me.dgvDetail.Rows.Count - 2 Then
                        '3. pastikan tanggalnya lebih kecil dari tanggal di baris bawahnya
                        If CDate(Me.dgvDetail.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value) >= CDate(Me.dgvDetail.Item(Me.ColDeliveryDate.Index, e.RowIndex + 1).Value) Then
                            MsgBox("Maximum delivery date is: " + Format(CDate(Me.dgvDetail.Item(Me.ColDeliveryDate.Index, e.RowIndex + 1).Value) - TimeSpan.FromDays(1), "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If

                    'format ulang tanggalnya
                    Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDetail_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvDetail.DataError
        'escape function, ketika rows added pada saat load form, saat fungsi yang mengisikan data kolom bertipe combobox pada kolom priority desc belum diisikan
    End Sub

    Private Sub dgvDetail_RowsAdded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvDetail.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvDetail.Rows(Me.dgvDetail.Rows.Count - 1)
            rowDGV.Cells(Me.ColDeliveryDate.Index).Value = ""
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvDetail.Rows.Count
            If rowDGV.Index = 0 Then
                rowDGV.Cells(Me.ColPrimaryPriority.Index).Value = Me.ObjPriorityLevel.GetPrimaryDefault
            Else
                rowDGV.Cells(Me.ColPrimaryPriority.Index).Value = Me.dgvDetail.Item(Me.ColPrimaryPriority.Index, rowDGV.Index - 1).Value
            End If
            rowDGV.Cells(Me.ColPriorityLevel.Index).Value = Me.ObjPriorityLevel.GetDesc(rowDGV.Cells(Me.ColPrimaryPriority.Index).Value)
            rowDGV.Cells(Me.ColRemarks.Index).Value = ""
            rowDGV.Cells(Me.ColQty.Index).Value = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDetail_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvDetail.RowsRemoved
        ObjCD.SetNumber(Me.dgvDetail)
    End Sub

    Private Sub dgvDetail_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvDetail.UserDeletingRow
        Try
            'jika sisa 1 baris yang valid, maka baris tersebut tidak boleh dihapuskan
            If Me.dgvDetail.Rows.Count = 2 Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Threading.Thread.Sleep(10)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.dgvDetail.Rows.RemoveAt(Me.mDeleteIndex)
            Me.mDeleteIndex = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        'Dim Urutan, UrutanDummy As Integer

        Try
            'pengecekan priority tidak jadi, karena memungkinkan untuk tanggal kedua yang high priority, walaupun tanggal pertama tidak high priority
            If Me.mCekPriorityIndex = -1 Then
                Exit Sub
            End If
            'pastikan priority level selalu lebih kecil atau sama dengan index priority level di atas nya
            'UrutanDummy = 0
            'Urutan = 0
            'If Me.mCekPriorityIndex > 0 Then
            '    UrutanDummy = Me.ObjPriorityLevel.GetUrutan(Me.dgvDetail.Item(Me.ColPrimaryPriority.Index, Me.mCekPriorityIndex - 1).Value)
            '    Urutan = Me.ObjPriorityLevel.GetUrutan(Me.ObjPriorityLevel.GetPrimary_ByDesc(Me.dgvDetail.Item(Me.ColPriorityLevel.Index, Me.mCekPriorityIndex).Value))
            'End If
            'If UrutanDummy > Urutan Then
            '    MsgBox("Maximum priority is " + Me.dgvDetail.Item(Me.ColPriorityLevel.Index, Me.mCekPriorityIndex - 1).Value)
            '    Me.dgvDetail.Item(Me.ColPriorityLevel.Index, Me.mCekPriorityIndex).Value = Me.Buffer
            '    Exit Sub
            'End If
            'pastikan priority level selalu lebih besar dari bawahnya
            'UrutanDummy = 0
            'Urutan = 0
            'If Me.mCekPriorityIndex < Me.dgvDetail.Rows.Count - 2 Then
            '    UrutanDummy = Me.ObjPriorityLevel.GetUrutan(Me.dgvDetail.Item(Me.ColPrimaryPriority.Index, Me.mCekPriorityIndex + 1).Value)
            '    Urutan = Me.ObjPriorityLevel.GetUrutan(Me.ObjPriorityLevel.GetPrimary_ByDesc(Me.dgvDetail.Item(Me.ColPriorityLevel.Index, Me.mCekPriorityIndex).Value))
            'End If
            'If UrutanDummy < Urutan Then
            '    MsgBox("Minimum priority is " + Me.dgvDetail.Item(Me.ColPriorityLevel.Index, Me.mCekPriorityIndex + 1).Value)
            '    Me.dgvDetail.Item(Me.ColPriorityLevel.Index, Me.mCekPriorityIndex).Value = Me.Buffer
            '    Exit Sub
            'End If
            'set primary
            Me.dgvDetail.Item(Me.ColPrimaryPriority.Index, Me.mCekPriorityIndex).Value = Me.ObjPriorityLevel.GetPrimary_ByDesc(Me.dgvDetail.Item(Me.ColPriorityLevel.Index, Me.mCekPriorityIndex).Value)
            Me.mCekPriorityIndex = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDetail_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDetail.CellContentClick
        Dim totalOrder As Decimal = 0
        Dim qtyOnDgvExCurrentRow As Decimal = 0
        Dim qtyRemaining As Decimal = 0

        Try
            If e.ColumnIndex = Me.ColPlusPlus.Index And e.RowIndex <> -1 Then
                If Me.dgvDetail.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value <> "" Then
                    totalOrder = Me.ObjPurcReq.GetJmlByIdentifier(Me.VIdentifier)
                    qtyOnDgvExCurrentRow = Me.GetJmlTotal - CDec(Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value)
                    qtyRemaining = totalOrder - qtyOnDgvExCurrentRow
                    If qtyRemaining <= 0 Then
                        Return
                    End If
                    Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(qtyRemaining, Me.DefPrec)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class