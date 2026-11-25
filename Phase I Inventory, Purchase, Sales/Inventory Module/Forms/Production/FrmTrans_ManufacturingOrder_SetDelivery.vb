Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_SetDelivery

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public FrmRef As FrmTrans_ManufacturingOrder
    Public VIdentifierJO As String = ""
    Public vPrimaryCodeNo As Double = -1
    Public vPrimaryUnit As Double = -1
    Public vReadOnly As Boolean = False

    Dim LangsungTutup As Boolean = False
    Dim vrowItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryRow
    Dim vTBLItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryDataTable
    Dim DefPrec As Integer = 0
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjPriorityLevel As New Core_Module.ClsMasterPriorityLevel
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjUnit As New Core_Module.ClsMasterUnit
    Dim Buffer As String = ""
    Dim Buffer1 As String = ""
    Dim mDeleteIndex As Integer = -1
    Dim mCekPriorityIndex As Integer = 0
    Dim UnitCode As String = ""

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
            Me.ColQty.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryCodeNo)
            Me.UnitCode = Me.ObjUnit.GetUnitCode(Me.vPrimaryUnit)
            Me.dgvDetail.AllowUserToAddRows = False
            Me.vTBLItemDelivery = Me.FrmRef.TBLItemDelivery
            For Each Me.vrowItemDelivery In Me.vTBLItemDelivery.Rows
                If vrowItemDelivery.identifier_jo = Me.VIdentifierJO Then
                    Me.dgvDetail.Rows.Add()
                    rowDGV = Me.dgvDetail.Rows(Me.dgvDetail.Rows.Count - 1)
                    rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvDetail.Rows.Count
                    rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(vrowItemDelivery.delivery_qty, Me.DefPrec)
                    rowDGV.Cells(Me.ColUnit.Index).Value = Me.UnitCode
                    rowDGV.Cells(Me.ColRemarks.Index).Value = vrowItemDelivery.remarks
                    rowDGV.Cells(Me.ColDeliveryDate.Index).Value = Format(vrowItemDelivery.delivery_date, "dd MMM yyyy")
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
            JmlOrder = Me.GetJmlOrder(Me.vPrimaryCodeNo)
            JmlDGV = Me.GetJmlTotal
            If JmlOrder <> JmlDGV Then
                MsgBox("Inconsistent delivery qty. Total delivery qty must be: " + Me.ObjCD.DigitGrouping(JmlOrder, Me.DefPrec) + " " + Me.UnitCode, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
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
                If e.ColumnIndex = Me.ColDeliveryDate.Index Then
                    Me.Buffer = Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColDeliveryDate.Index Or e.ColumnIndex = Me.ColQty.Index Or e.ColumnIndex = Me.ColRemarks.Index Then
                    Me.Buffer = Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value
                    If e.ColumnIndex = Me.ColQty.Index Then
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
            For Each rowDGV In Me.FrmRef.DgvItem.Rows
                If rowDGV.Cells(Me.FrmRef.ColItem_Identifier.Index).Value = Me.VIdentifierJO And rowDGV.Cells(Me.FrmRef.ColItem_PrimaryProduct.Index).Value = Me.vPrimaryCodeNo Then
                    RetVal += rowDGV.Cells(Me.FrmRef.ColItem_QtyTarget.Index).Value
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

    Private Sub dgvDetail_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDetail.CellEndEdit
        Dim JmlOrder As Decimal = 0
        Dim JmlDGV As Decimal = 0
        Dim JmlMax As Decimal = 0
        Dim IndexDikurangi As Integer = 0
        Try
            If IsNothing(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Exit Sub
                Else
                    Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            Else
                If Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                    Exit Sub
                End If
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    If Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                ElseIf e.ColumnIndex = Me.ColQty.Index Then
                    If Me.ObjCD.CekBilanganPositif(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify positive number.", MsgBoxStyle.Information, "Information")
                        Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.Buffer1
                        Exit Sub
                    End If
                    Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                    If Me.dgvDetail.RowCount = 2 Then
                        JmlOrder = Me.GetJmlOrder(Me.vPrimaryCodeNo)
                        JmlDGV = Me.GetJmlTotal
                        If JmlOrder < JmlDGV Then
                            JmlMax = JmlOrder - (JmlDGV - CDec(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value))
                            MsgBox("Maximum value your can enter is: " + Me.ObjCD.DigitGrouping(JmlMax, Me.DefPrec), MsgBoxStyle.Information, "Information")
                            Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.Buffer1
                            Exit Sub
                        End If
                    ElseIf Me.dgvDetail.RowCount = 3 Then
                        JmlOrder = Me.GetJmlOrder(Me.vPrimaryCodeNo)
                        If JmlOrder <= CDec(Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value) Then
                            MsgBox("Please specify value below: " + Me.ObjCD.DigitGrouping(JmlOrder, Me.DefPrec), MsgBoxStyle.Information, "Information")
                            Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.Buffer1
                            Exit Sub
                        End If
                        If e.RowIndex = 0 Then
                            IndexDikurangi = 1
                        ElseIf e.RowIndex = 1 Then
                            IndexDikurangi = 0
                        End If
                        Me.dgvDetail.Item(Me.ColQty.Index, IndexDikurangi).Value = Me.ObjCD.DigitGrouping(JmlOrder - CDec(Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value), DefPrec)
                    Else
                    End If
                ElseIf e.ColumnIndex = Me.ColDeliveryDate.Index Then
                    If IsDate(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify a valid date.", MsgBoxStyle.Information, "Information")
                        Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        If e.RowIndex = Me.dgvDetail.Rows.Count - 2 And e.RowIndex <> 0 Then
                            Me.mDeleteIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                        Exit Sub
                    End If
                    If e.RowIndex > 0 Then
                        If CDate(Me.dgvDetail.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value) <= CDate(Me.dgvDetail.Item(Me.ColDeliveryDate.Index, e.RowIndex - 1).Value) Then
                            MsgBox("Minimum delivery date is: " + Format(CDate(Me.dgvDetail.Item(Me.ColDeliveryDate.Index, e.RowIndex - 1).Value) + TimeSpan.FromDays(1), "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If e.RowIndex = Me.dgvDetail.Rows.Count - 2 Then
                                Me.mDeleteIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Exit Sub
                        End If
                    Else
                        If Me.FrmRef.DtpTargetFinishDate.Value.Date > CDate(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            MsgBox("Minimum delivery date is: " + Format(Me.FrmRef.DtpTargetFinishDate.Value.Date, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If e.RowIndex = Me.dgvDetail.Rows.Count - 2 And e.RowIndex <> 0 Then
                                Me.mDeleteIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Exit Sub
                        End If
                    End If
                    If e.RowIndex < Me.dgvDetail.Rows.Count - 2 Then
                        If CDate(Me.dgvDetail.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value) >= CDate(Me.dgvDetail.Item(Me.ColDeliveryDate.Index, e.RowIndex + 1).Value) Then
                            MsgBox("Maximum delivery date is: " + Format(CDate(Me.dgvDetail.Item(Me.ColDeliveryDate.Index, e.RowIndex + 1).Value) - TimeSpan.FromDays(1), "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
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
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvDetail.Rows.Count
            rowDGV.Cells(Me.ColDeliveryDate.Index).Value = ""
            rowDGV.Cells(Me.ColQty.Index).Value = 0
            rowDGV.Cells(Me.ColUnit.Index).Value = Me.UnitCode
            rowDGV.Cells(Me.ColRemarks.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDetail_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvDetail.RowsRemoved
        ObjCD.SetNumber(Me.dgvDetail)
    End Sub

    Private Sub dgvDetail_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvDetail.UserDeletingRow
        Try
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

    Public Sub New(ByVal vBackColorAll As Color, ByVal vFrm As FrmTrans_ManufacturingOrder, ByVal vIdentifierJO As String, ByVal vPrimaryCodeNo As Double, ByVal vPrimaryUnit As Double, ByVal vReadOnly As Boolean)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.BackColorAll = vBackColorAll
        Me.FrmRef = vFrm
        Me.VIdentifierJO = vIdentifierJO
        Me.vPrimaryCodeNo = vPrimaryCodeNo
        Me.vPrimaryUnit = vPrimaryUnit
        Me.vReadOnly = vReadOnly
    End Sub
End Class