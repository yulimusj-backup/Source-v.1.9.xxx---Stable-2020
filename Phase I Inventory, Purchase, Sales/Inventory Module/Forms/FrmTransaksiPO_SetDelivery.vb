Option Explicit On
Option Strict Off

Public Class FrmTransaksiPO_SetDelivery
    Dim LangsungTutup As Boolean = False
    Dim vrowItemDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryRow
    Dim vTBLItemDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryDataTable
    Dim DefPrec As Integer = 0
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjPriorityLevel As New Core_Module.ClsMasterPriorityLevel
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjUnit As New Core_Module.ClsMasterUnit
    Dim Buffer As String = ""
    Dim Buffer1 As String = ""
    Dim mDeleteIndex As Integer = -1
    Dim mCekPriorityIndex As Integer = 0

    Public FrmRef As FrmTransaksiPO
    Public VIdentifier As String = ""
    Public vPrimaryCodeNo As Double = -1
    Public vReadOnly As Boolean = False

    ''@rosadinnn
    ''@MUaT
    Public vPrimaryUnitMulti As Double = -1
    Dim DefPrecMulti As Integer = 0
    Public vRatioDefault As Double = 1
    Public vRatioReference As Double = 1
    Public vPrimaryUnitDefault As Double = -1
    Public DefPrecDefault As Integer = 0
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim UnitCode As String = ""
    Dim UnitCodeMulti As String = ""
    Dim Ratio_Str As String = ""

    Private Sub FrmTransaksiPO_SetDelivery_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub FrmTransaksiPO_SetDelivery_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As DataGridViewRow
        Dim ListPriorityLevel As New List(Of String)
        Dim ColWidth As Double = 0

        Try
            Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryCodeNo)
            UnitCode = Me.ObjUnit.GetUnitCode(Me.vPrimaryUnitDefault)
            UnitCodeMulti = Me.ObjUnit.GetUnitCode(Me.vPrimaryUnitMulti)

            ''@rosadinnn
            ''@MUaT
            Dim ObjPO As New ClsTransaksiPO
            Me.DefPrecMulti = ObjProduct.GetAngkaDecimal(Me.vPrimaryCodeNo, Me.vPrimaryUnitMulti)
            Me.DefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryCodeNo)
            Ratio_Str = Me.ObjCD.DigitGrouping(Me.vRatioReference, Me.DefPrecMulti) & " " & UnitCodeMulti & " = " _
            & Me.ObjCD.DigitGrouping(Me.vRatioDefault, Me.DefPrecDefault) & " " & UnitCode

            Me.dgvDetail.AllowUserToAddRows = False
            Me.vTBLItemDelivery = Me.FrmRef.vTBLItemDelivery
            For Each Me.vrowItemDelivery In Me.vTBLItemDelivery.Rows
                If vrowItemDelivery.identifier = Me.VIdentifier Then
                    Me.dgvDetail.Rows.Add()
                    rowDGV = Me.dgvDetail.Rows(Me.dgvDetail.Rows.Count - 1)
                    rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvDetail.Rows.Count
                    rowDGV.Cells(Me.ColDelay.Index).Value = vrowItemDelivery.delay
                    rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(vrowItemDelivery.delivery_qty, Me.DefPrec)
                    rowDGV.Cells(Me.ColRemarks.Index).Value = vrowItemDelivery.remarks
                    rowDGV.Cells(Me.ColDeliveryDate.Index).Value = Format(vrowItemDelivery.delivery_date, "dd MMM yyyy")
                    rowDGV.Cells(Me.ColUnit.Index).Value = Me.UnitCode
                    rowDGV.Cells(Me.ColUnitMulti.Index).Value = Me.UnitCodeMulti
                    rowDGV.Cells(Me.ColRatio.Index).Value = Me.Ratio_Str

                    ''@rosadinnn
                    ''MUaT
                    rowDGV.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(vrowItemDelivery.delivery_qty_multi, Me.DefPrecMulti)

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
                ''@rosadinnn
                ''@MUaT
                ''ganti defprec
                MsgBox("Inconsistent delivery qty. Total delivery qty must be: " + Me.ObjCD.DigitGrouping(JmlOrder, Me.DefPrecMulti) + " " + Me.UnitCodeMulti, MsgBoxStyle.Information, "Information")

                Exit Sub
            End If

            JmlOrder = Me.GetJmlOrder_Default(Me.vPrimaryCodeNo)
            JmlDGV = Me.GetJmlTotal_Default
            If JmlOrder <> JmlDGV Then
                ''@rosadinnn
                ''@MUaT
                ''ganti defprec
                MsgBox("Inconsistent delivery qty. Total delivery qty must be: " + Me.ObjCD.DigitGrouping(JmlOrder, Me.DefPrecDefault) + " " + Me.UnitCode, MsgBoxStyle.Information, "Information")

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

                ''@rosadinnn
                ''@MUaT
                ''@ganti qty yg bisa diedit jd qty multi
                If e.ColumnIndex = Me.ColDeliveryDate.Index Or _
                e.ColumnIndex = Me.ColDelay.Index Or _
                e.ColumnIndex = Me.ColQtyMulti.Index Or _
                e.ColumnIndex = Me.ColRemarks.Index Then
                    Me.Buffer = Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value
                    If e.ColumnIndex = Me.ColQtyMulti.Index Then
                        Me.Buffer1 = Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value
                    Else
                        Me.Buffer1 = ""
                    End If
                Else
                    e.Cancel = True
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJmlOrder(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0
        Try
            For Each rowDGV In Me.FrmRef.dgvItem.Rows
                If rowDGV.Cells(Me.FrmRef.ColPrimaryCodeNo.Index).Value = Me.vPrimaryCodeNo Then

                    ''@rosadinnn
                    ''@MUaT
                    ''@update berkaitan qty multi
                    RetVal += rowDGV.Cells(Me.FrmRef.ColQtyMulti.Index).Value

                End If
            Next
            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlOrder_Default(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0
        Try
            For Each rowDGV In Me.FrmRef.dgvItem.Rows
                If rowDGV.Cells(Me.FrmRef.ColPrimaryCodeNo.Index).Value = Me.vPrimaryCodeNo Then

                    ''@rosadinnn
                    ''@MUaT
                    ''@update berkaitan qty multi
                    RetVal += rowDGV.Cells(Me.FrmRef.ColQty.Index).Value

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

                ''@rosadinnn
                ''@MUaT
                ''@update berkaitan qty multi
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

                ''@rosadinnn
                ''@MUaT
                ''@update berkaitan qty multi
                RetVal += CDec(rowDGV.Cells(Me.ColQty.Index).Value)

            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvDetail_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDetail.CellContentClick
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
                    Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(qtyRemaining, Me.DefPrecMulti)

                    ''@rosadinnn
                    ''@MUaT
                    Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) * _
                                                                        Me.vRatioDefault / Me.vRatioReference, Me.DefPrec)

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDetail_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDetail.CellEndEdit
        Dim JmlOrder As Decimal = 0
        Dim JmlDGV As Decimal = 0
        Dim JmlMax As Decimal = 0
        Dim IndexDikurangi As Integer = 0

        Try
            If IsNothing(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
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

                    ''@rosadinnn
                    ''@MUaT
                    ''@perubahan kolom yg bisa diedit
                ElseIf e.ColumnIndex = Me.ColQtyMulti.Index Then
                    'pastikan nilai qty > 0
                    If Me.ObjCD.CekBilanganPositif(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify positive number.", MsgBoxStyle.Information, "Information")
                        Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.Buffer1
                        Exit Sub
                    End If

                    ''@rosadinnn
                    ''@MUaT
                    ''update defprec
                    'digit grouping sesuai defprecmulti
                    Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrecMulti)

                    ''@rosadinnn
                    ''@MUaT-> pastikan tidak bikin qty multi tidak 0
                    Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) * _
                                                                        CDec(Me.vRatioDefault) / CDec(Me.vRatioReference), Me.DefPrec)

                    If CDec(Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value) = 0 Then
                        MsgBox("Qty restored to previous value, because qty you entered make default qty=0.", MsgBoxStyle.Information, "Information")
                        Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.Buffer1
                    End If

                    'pastikan total nilai qty tidak lebih besar dari qty order
                    If Me.dgvDetail.RowCount = 2 Then
                        'hanya berlaku jika jumlah baris yang valid hanya 1
                        JmlOrder = Me.GetJmlOrder(Me.vPrimaryCodeNo)
                        JmlDGV = Me.GetJmlTotal
                        If JmlOrder < JmlDGV Then
                            JmlMax = JmlOrder - (JmlDGV - CDec(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value))
                            MsgBox("Maximum value your can enter is: " + Me.ObjCD.DigitGrouping(JmlMax, Me.DefPrecMulti), MsgBoxStyle.Information, "Information")
                            Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.Buffer1
                            Exit Sub
                        End If
                    ElseIf Me.dgvDetail.RowCount = 3 Then
                        'jika jumlah baris valid ada 2, maka saling menyeimbangkan dengan catatan jumlah maksimum per baris harus lebih kecil dari total order
                        JmlOrder = Me.GetJmlOrder(Me.vPrimaryCodeNo)

                        ''@rosadinnn
                        ''@MUaT
                        If JmlOrder <= CDec(Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) Then
                            MsgBox("Please specify value below: " + Me.ObjCD.DigitGrouping(JmlOrder, Me.DefPrecMulti), MsgBoxStyle.Information, "Information")
                            Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.Buffer1
                            Exit Sub
                        End If
                        If e.RowIndex = 0 Then
                            IndexDikurangi = 1
                        ElseIf e.RowIndex = 1 Then
                            IndexDikurangi = 0
                        End If
                        Me.dgvDetail.Item(Me.ColQtyMulti.Index, IndexDikurangi).Value = Me.ObjCD.DigitGrouping(JmlOrder - _
                        CDec(Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value), DefPrecMulti)

                        ''@rosadinnn
                        ''@MUaT
                        Me.dgvDetail.Item(Me.ColQty.Index, IndexDikurangi).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvDetail.Item(Me.ColQtyMulti.Index, IndexDikurangi).Value) * _
                                                                        CDec(Me.vRatioDefault) / CDec(Me.vRatioReference), Me.DefPrec)


                    Else
                        'jika lebih dari 2 baris valid, maka abaikan saja, nanti akan di cek pada button ok
                    End If
                ElseIf e.ColumnIndex = Me.ColDelay.Index Then
                    'pastikan nilai qty >0
                    If Me.ObjCD.CekNonNegatif(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify non-negative number.", MsgBoxStyle.Information, "Information")
                        Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value, 0)
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
                        '2. pastikan tanggalnya minimal sama dengan tanggal PO
                        If Me.FrmRef.dtpDate.Value.Date > CDate(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            MsgBox("Minimum delivery date is: " + Format(Me.FrmRef.dtpDate.Value.Date, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If e.RowIndex = Me.dgvDetail.Rows.Count - 2 And e.RowIndex <> 0 Then
                                'jika data baru, hapuskan baris ybs
                                Me.mDeleteIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Exit Sub
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
                rowDGV.Cells(Me.ColDelay.Index).Value = 0
            Else
                rowDGV.Cells(Me.ColDelay.Index).Value = Me.dgvDetail.Item(Me.ColDelay.Index, rowDGV.Index - 1).Value
            End If
            rowDGV.Cells(Me.ColRemarks.Index).Value = ""
            rowDGV.Cells(Me.ColQty.Index).Value = 0

            ''@rosadinnn
            ''@MUaT
            rowDGV.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping("0", Me.DefPrecMulti)
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
End Class