Option Explicit On
Option Strict Off

Public Class FrmTransaksiSO_SetDelivery
    Dim LangsungTutup As Boolean = False
    Dim vrowItemDelivery As DseTransaksiSO.tbltrans_so_itemdeliveryRow
    Dim vTBLItemDelivery As DseTransaksiSO.tbltrans_so_itemdeliveryDataTable
    Dim DefPrec As Integer = 0
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjPriorityLevel As New Core_Module.ClsMasterPriorityLevel
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Buffer As String = ""
    Dim mDeleteIndex As Integer = -1
    Dim mCekPriorityIndex As Integer = 0
    Dim UnitCode As String = ""
    Dim UnitCodeMulti As String = ""
    Dim Ratio_Str As String = ""
    Dim ObjUnit As New Core_Module.ClsMasterUnit

    Public FrmRef As FrmTransaksiSO
    Public VIdentifier As String = ""
    Public vPrimaryCodeNo As Double = -1
    Public vReadOnly As Boolean = False

    Public vPrimaryUnitMulti As Double = -1
    Public vPrimaryUnitDefault As Double = -1
    Public vRatioReference As Double = 1
    Public vRatioDefault As Double = 1
    Public vDefPrecMulti As Integer = 0
    Public vDefPrecDefault As Integer = 0
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub FrmTransaksiSO_SetDelivery_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub FrmTransaksiSO_SetDelivery_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As DataGridViewRow
        Dim ListPriorityLevel As New List(Of String)
        Dim dv As DataRowView
        Dim ColWidth As Double = 0

        Try
            Me.ShowIcon = False
            UnitCode = Me.ObjUnit.GetUnitCode(Me.vPrimaryUnitDefault)
            UnitCodeMulti = Me.ObjUnit.GetUnitCode(Me.vPrimaryUnitMulti)
            Ratio_Str = Me.ObjCD.DigitGrouping(Me.vRatioReference, Me.vDefPrecMulti) & " " & UnitCodeMulti & " = " _
            & Me.ObjCD.DigitGrouping(Me.vRatioDefault, Me.vDefPrecDefault) & " " & UnitCode

            Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryCodeNo)
            Me.ObjPriorityLevel.GetListDescPriorityLevel(ListPriorityLevel)
            Me.ColPriorityLevel.Items.Clear()
            For Each Desc As String In ListPriorityLevel
                Me.ColPriorityLevel.Items.Add(Desc)
            Next

            Me.dgvDetail.AllowUserToAddRows = False
            Me.vTBLItemDelivery = Me.FrmRef.vTBLItemDelivery
            Me.vTBLItemDelivery.DefaultView.RowFilter = Nothing
            Me.vTBLItemDelivery.DefaultView.Sort = "[delivery date]"
            For Each dv In Me.vTBLItemDelivery.DefaultView
                Me.vrowItemDelivery = dv.Row
                If vrowItemDelivery.identifier = Me.VIdentifier Then
                    Me.dgvDetail.Rows.Add()
                    rowDGV = Me.dgvDetail.Rows(Me.dgvDetail.Rows.Count - 1)
                    rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvDetail.Rows.Count
                    rowDGV.Cells(Me.ColPrimaryPriority.Index).Value = vrowItemDelivery.delivery_priority
                    rowDGV.Cells(Me.ColPriorityLevel.Index).Value = Me.ObjPriorityLevel.GetDesc(vrowItemDelivery.delivery_priority)
                    rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(vrowItemDelivery.delivery_qty, Me.DefPrec)
                    rowDGV.Cells(Me.ColRemarks.Index).Value = vrowItemDelivery.remarks
                    rowDGV.Cells(Me.ColDeliveryDate.Index).Value = Format(vrowItemDelivery.delivery_date, "dd MMM yyyy")
                    rowDGV.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(vrowItemDelivery.delivery_qty_multi, Me.vDefPrecMulti)
                End If
            Next

            If Me.vReadOnly = False Then
                Me.dgvDetail.AllowUserToAddRows = True
            Else
                Me.dgvDetail.AllowUserToDeleteRows = False
                Me.dgvDetail.AllowUserToAddRows = False
                Me.dgvDetail.ReadOnly = True
                Me.btnOK.Visible = False
                Me.btnOK.TabStop = False
            End If

            If Me.vPrimaryUnitDefault = Me.vPrimaryUnitMulti Then
                Me.ColRatio.Visible = False
                Me.ColQty.Visible = False
                Me.ColUnit.Visible = False
            Else
                Me.ColRatio.Visible = True
                Me.ColQty.Visible = True
                Me.ColUnit.Visible = True
            End If

            For Each col As DataGridViewColumn In Me.dgvDetail.Columns
                If col.Visible = True Then
                    ColWidth += col.Width
                End If
            Next
            Me.Size = New Point(ColWidth + 30 + Me.dgvDetail.RowHeadersWidth + Me.dgvDetail.Left, Me.Size.Height)
            Me.CenterToScreen()

            Me.dgvDetail.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
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
            JmlOrder = Me.GetJmlOrder(Me.vPrimaryCodeNo)
            JmlDGV = Me.GetJmlTotal
            If JmlOrder <> JmlDGV Then
                MsgBox("Inconsistent delivery qty. Total delivery qty must be: " + Me.ObjCD.DigitGrouping(JmlOrder, Me.vDefPrecMulti) + " " + Me.UnitCodeMulti, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'pastikan jumlah pengiriman total qty default sama dengan jumlah dgv
            JmlOrder = Me.GetJmlOrder_Default(Me.vPrimaryCodeNo)
            JmlDGV = Me.GetJmlTotal_Default
            If JmlOrder <> JmlDGV Then
                MsgBox("Inconsistent delivery qty. Total delivery qty must be: " + Me.ObjCD.DigitGrouping(JmlOrder, Me.vDefPrecDefault) + " " + Me.UnitCode, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'pastikan qty tiap baris >0
            For Each rowdgv As DataGridViewRow In Me.dgvDetail.Rows
                If CDec(rowdgv.Cells(Me.ColQtyMulti.Index).Value) <= 0 And rowdgv.Index <> Me.dgvDetail.RowCount - 1 Then
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
                e.ColumnIndex = Me.ColQtyMulti.Index Or _
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

    Private Function GetJmlOrder_Default(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV In Me.FrmRef.dgvItem.Rows
                If rowDGV.Cells(Me.FrmRef.ColPrimaryCodeNo.Index).Value = Me.vPrimaryCodeNo Then
                    RetVal += rowDGV.Cells(Me.FrmRef.ColQty.Index).Value
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlOrder(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV In Me.FrmRef.dgvItem.Rows
                If rowDGV.Cells(Me.FrmRef.ColPrimaryCodeNo.Index).Value = Me.vPrimaryCodeNo Then
                    RetVal += rowDGV.Cells(Me.FrmRef.ColQtyMulti.Index).Value
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlTotal() As Decimal
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV As DataGridViewRow In Me.dgvDetail.Rows
                RetVal += CDec(rowDGV.Cells(Me.ColQtyMulti.Index).Value)
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlTotal_Default() As Decimal
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
                ElseIf e.ColumnIndex = Me.ColQtyMulti.Index Then
                    'pastikan nilai qty >0
                    If Me.ObjCD.CekBilanganPositif(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify positive number.", MsgBoxStyle.Information, "Information")
                        Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'digit grouping sesuai defprec
                    Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value, Me.vDefPrecMulti)
                    'pastikan total nilai qty tidak lebih besar dari qty order
                    'If Me.dgvDetail.RowCount = 2 Then
                    '    'hanya berlaku jika jumlah baris yang valid hanya 1
                    '    JmlOrder = Me.GetJmlOrder(Me.vPrimaryCodeNo)
                    '    JmlDGV = Me.GetJmlTotal
                    '    If JmlOrder < JmlDGV Then
                    '        JmlMax = JmlOrder - (JmlDGV - CDec(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value))
                    '        MsgBox("Maximum value your can enter is: " + Me.ObjCD.DigitGrouping(JmlMax, Me.DefPrec), MsgBoxStyle.Information, "Information")
                    '        Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    '        Exit Sub
                    '    End If
                    'ElseIf Me.dgvDetail.RowCount = 3 Then
                    '    'jika jumlah baris valid ada 2, maka saling menyeimbangkan dengan catatan jumlah maksimum per baris harus lebih kecil dari total order
                    '    JmlOrder = Me.GetJmlOrder(Me.vPrimaryCodeNo)
                    '    If JmlOrder <= CDec(Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value) Then
                    '        MsgBox("Please specify value below: " + Me.ObjCD.DigitGrouping(JmlOrder, Me.DefPrec), MsgBoxStyle.Information, "Information")
                    '        Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    '        Exit Sub
                    '    End If
                    '    If e.RowIndex = 0 Then
                    '        IndexDikurangi = 1
                    '    ElseIf e.RowIndex = 1 Then
                    '        IndexDikurangi = 0
                    '    End If
                    '    Me.dgvDetail.Item(Me.ColQty.Index, IndexDikurangi).Value = Me.ObjCD.DigitGrouping(JmlOrder - _
                    '    CDec(Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value), DefPrec)
                    'Else
                    '    'jika lebih dari 2 baris valid, maka abaikan saja, nanti akan di cek pada button ok
                    'End If
                    Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) * _
                    CDec(Me.vRatioDefault) / CDec(Me.vRatioReference), Me.DefPrec)
                    If CDec(Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value) <= 0 Then
                        MsgBox("Qty restored to previous value, because it make default qty=0.", MsgBoxStyle.Information, "Information")
                        Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) * _
                        CDec(Me.vRatioDefault) / CDec(Me.vRatioReference), Me.DefPrec)
                        Exit Sub
                    End If
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
                        If Me.FrmRef.dtpDate.Value.Date > CDate(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            If MsgBox("Delivery date is smaller than so date, continue?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
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
            rowDGV.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping("0", Me.vDefPrecMulti)
            rowDGV.Cells(Me.ColUnit.Index).Value = Me.UnitCode
            rowDGV.Cells(Me.ColUnitMulti.Index).Value = Me.UnitCodeMulti
            rowDGV.Cells(Me.ColRatio.Index).Value = Me.Ratio_Str
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
            Threading.Thread.Sleep(10)
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
                    totalOrder = Me.GetJmlOrder(Me.vPrimaryCodeNo)
                    qtyOnDgvExCurrentRow = Me.GetJmlTotal - CDec(Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value)
                    qtyRemaining = totalOrder - qtyOnDgvExCurrentRow
                    If qtyRemaining <= 0 Then
                        Return
                    End If
                    Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(qtyRemaining, Me.vDefPrecMulti)
                    Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) * _
                    Me.vRatioDefault / Me.vRatioReference, Me.vDefPrecDefault)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class