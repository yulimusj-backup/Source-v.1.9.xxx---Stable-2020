Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_ExplainPrOB_Delivery
    Public ReadOnly1 As Boolean = False
    Dim LangsungTutup As Boolean = False
    Dim vrowItemDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryRow
    Dim vTBLItemDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable
    Dim DefPrec As Integer = 0
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjPriorityLevel As New Core_Module.ClsMasterPriorityLevel
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjPrOB As New ClsTransaksiPurchaseRequest
    Dim Buffer As String = ""
    Dim mDeleteIndex As Integer = -1
    Dim mCekPriorityIndex As Integer = 0
    Dim UnitCode As String = ""
    Dim ObjUnit As New Core_Module.ClsMasterUnit

    Public FrmRef As FrmTrans_MaterialPlanningDashboard_ExplainPrOB
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
    Public vQtySOBDefault As Decimal = 0
    Public vQtySOBMulti As Decimal = 0
    Public vSODate As Date = #1/1/1950#

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
        'Dim dv As Data.DataRowView
        Dim ColWidth As Double = 0

        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim TBL As New DseTransaksiMaterialPlanning2.tbltrans_so_itemdeliveryDataTable
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_so_itemdeliveryRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_identifier_so(TBL, Me.VIdentifier)

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
            For mn As Integer = TBL.Rows.Count - 1 To 0 Step -1
                row = TBL.Rows(mn)
                If row.identifier = Me.VIdentifier Then
                    Me.dgvDetail.Rows.Add()
                    rowDGV = Me.dgvDetail.Rows(Me.dgvDetail.Rows.Count - 1)
                    rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvDetail.Rows.Count
                    rowDGV.Cells(Me.ColPrimaryPriority.Index).Value = row.delivery_priority
                    rowDGV.Cells(Me.ColPriorityLevel.Index).Value = Me.ObjPriorityLevel.GetDesc(row.delivery_priority)
                    rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row.delivery_qty, Me.DefPrec)
                    rowDGV.Cells(Me.ColRemarks.Index).Value = row.remarks
                    rowDGV.Cells(Me.ColDeliveryDate.Index).Value = Format(row.delivery_date, "dd MMM yyyy")
                    rowDGV.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(row.delivery_qty_multi, Me.vDefPrecMulti)
                End If
            Next

            'isikan data balance ke dalam dgv
            Me.TxtSOBMulti.Text = Me.ObjCD.DigitGrouping(Me.vQtySOBMulti, Me.vDefPrecMulti)
            Me.TxtSOB.Text = Me.ObjCD.DigitGrouping(Me.vQtySOBDefault, Me.vDefPrecDefault)
            Me.TxtOrderQtyMulti.Text = Me.ObjCD.DigitGrouping(Me.ObjPrOB.GetQtyMulti(Me.VIdentifier), Me.vDefPrecMulti)
            Me.TxtOrderQty.Text = Me.ObjCD.DigitGrouping(Me.ObjPrOB.GetQty(Me.VIdentifier), Me.vDefPrecDefault)
            Me.vQtySOBMulti = Me.TxtOrderQtyMulti.Text
            Me.vQtySOBDefault = Me.TxtOrderQty.Text
            Me.IsikanDataBalance()

            If Me.vReadOnly = False Then
                Me.dgvDetail.AllowUserToAddRows = True
            Else
                Me.dgvDetail.AllowUserToDeleteRows = False
                Me.dgvDetail.AllowUserToAddRows = False
                Me.dgvDetail.ReadOnly = True
                Me.btnOK.Visible = False
                Me.btnOK.TabStop = False
                Me.btnCancel.Text = "Close"
            End If
            Me.CancelButton = Me.btnCancel

            Me.ColRatio.Visible = False
            If Me.vPrimaryUnitDefault = Me.vPrimaryUnitMulti Then
                Me.ColQty.Visible = False
                Me.ColUnit.Visible = False
                'Me.dgvDetail.Columns.Insert(Me.ColQtyMulti.Index + 1, Me.ColQtyBalanceMulti)
                'Me.ColQtyBalanceMulti.Index = Me.ColQtyMulti.Index + 1
                Me.ColQtyBalanceMulti.Visible = True
                Me.ColQtyBalanceDefault.Visible = False
            Else
                Me.ColQty.Visible = True
                Me.ColUnit.Visible = True
                Me.ColQtyBalanceMulti.Visible = True
                Me.ColQtyBalanceDefault.Visible = True
            End If

            Me.LblRatio.Text = Ratio_Str
            Me.LblUnitCode.Text = UnitCode
            Me.LblUnitCodeOrderBalance.Text = UnitCode
            Me.LblOrderBalanceUnitMulti.Text = UnitCodeMulti + " = "
            Me.LblOrderUnitMulti.Text = UnitCodeMulti + " = "

            Me.LblOrderUnitMulti.Left = Me.TxtOrderQtyMulti.Left + Me.TxtOrderQtyMulti.Width + 5
            Me.TxtOrderQty.Left = Me.LblOrderUnitMulti.Left + Me.LblOrderUnitMulti.Width + 5
            Me.LblUnitCode.Left = Me.TxtOrderQty.Left + Me.TxtOrderQty.Width + 5
            Me.LblOrderBalanceUnitMulti.Left = Me.TxtSOBMulti.Left + Me.TxtSOBMulti.Width + 5
            Me.TxtSOB.Left = Me.LblOrderBalanceUnitMulti.Left + Me.LblOrderBalanceUnitMulti.Width + 5
            Me.LblUnitCodeOrderBalance.Left = Me.TxtSOB.Left + Me.TxtSOB.Width + 5
            If UnitCodeMulti = UnitCode Then
                Me.LblOrderBalanceUnitMulti.Text = Me.LblOrderBalanceUnitMulti.Text.Replace("=", "")
                Me.LblOrderUnitMulti.Text = Me.LblOrderUnitMulti.Text.Replace("=", "")
                Me.TxtOrderQty.Visible = False
                Me.TxtOrderQty.TabStop = False
                Me.LblUnitCode.Visible = False
                Me.TxtSOB.Visible = False
                Me.TxtSOB.TabStop = False
                Me.LblUnitCodeOrderBalance.Visible = False
                Me.LblRatio.Visible = False
                Me.dgvDetail.Location = New Point(Me.dgvDetail.Location.X, Me.TxtSOBMulti.Top + Me.TxtSOB.Height + 5)
                Me.dgvDetail.Height = Me.btnCancel.Top - 5 - Me.dgvDetail.Location.Y
            End If
            Me.dgvDetail.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPriorityLevel = Nothing
        End Try
    End Sub

    Private Sub IsikanDataBalance()
        Dim rowDGV As DataGridViewRow
        Dim QtyAllocated As Decimal = 0
        Dim QtySOB As Decimal = Me.TxtSOBMulti.Text
        Dim QtyBalance, QtyBalanceMulti As Decimal
        Dim QtyDeliveryMulti As Decimal = 0

        Try
            For ij As Integer = Me.dgvDetail.RowCount - 1 To 0 Step -1
                rowDGV = Me.dgvDetail.Rows(ij)
                If IsNothing(rowDGV.Cells(Me.ColUnit.Index).Value) = False Then
                    QtyDeliveryMulti = rowDGV.Cells(Me.ColQtyMulti.Index).Value
                    QtyBalance = 0
                    QtyBalanceMulti = 0

                    If QtyDeliveryMulti <= QtySOB - QtyAllocated Then
                        'alokasikan semua untuk balancenya
                        QtyBalanceMulti = QtyDeliveryMulti
                        QtyBalance = Math.Round((QtyBalanceMulti * vRatioDefault) / Me.vRatioReference, Me.vDefPrecDefault, MidpointRounding.AwayFromZero)
                        rowDGV.Cells(Me.ColQtyBalanceMulti.Index).Value = Me.ObjCD.DigitGrouping(QtyBalanceMulti, Me.vDefPrecMulti)
                        rowDGV.Cells(Me.ColQtyBalanceDefault.Index).Value = Me.ObjCD.DigitGrouping(QtyBalance, Me.vDefPrecDefault)

                        QtyAllocated += QtyBalanceMulti
                    Else
                        'alokasikan sisanya untuk balance
                        QtyBalanceMulti = QtySOB - QtyAllocated
                        QtyBalance = Math.Round((QtyBalanceMulti * vRatioDefault) / Me.vRatioReference, Me.vDefPrecDefault, MidpointRounding.AwayFromZero)
                        rowDGV.Cells(Me.ColQtyBalanceMulti.Index).Value = Me.ObjCD.DigitGrouping(QtyBalanceMulti, Me.vDefPrecMulti)
                        rowDGV.Cells(Me.ColQtyBalanceDefault.Index).Value = Me.ObjCD.DigitGrouping(QtyBalance, Me.vDefPrecDefault)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadData()
        Dim rowDGV As DataGridViewRow
        Dim ListPriorityLevel As New List(Of String)
        Dim dv As Data.DataRowView
        Dim ColWidth As Double = 0

        Try
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
            Me.vTBLItemDelivery = Me.FrmRef.vDT_ItemSOBDelivery
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
                Me.btnCancel.Text = "Close"
            End If
            Me.CancelButton = Me.btnCancel

            Me.ColRatio.Visible = False
            If Me.vPrimaryUnitDefault = Me.vPrimaryUnitMulti Then
                Me.ColQty.Visible = False
                Me.ColUnit.Visible = False
            Else
                Me.ColQty.Visible = True
                Me.ColUnit.Visible = True
            End If

            Me.LblRatio.Text = Ratio_Str
            Me.LblUnitCode.Text = UnitCode
            Me.LblUnitCodeOrderBalance.Text = UnitCode
            Me.LblOrderBalanceUnitMulti.Text = UnitCodeMulti + " = "
            Me.LblOrderUnitMulti.Text = UnitCodeMulti + " = "
            Me.TxtSOBMulti.Text = Me.ObjCD.DigitGrouping(Me.vQtySOBMulti, Me.vDefPrecMulti)
            Me.TxtSOB.Text = Me.ObjCD.DigitGrouping(Me.vQtySOBDefault, Me.vDefPrecDefault)
            Me.TxtOrderQtyMulti.Text = Me.ObjCD.DigitGrouping(Me.ObjPrOB.GetQtyMulti(Me.VIdentifier), Me.vDefPrecMulti)
            Me.TxtOrderQty.Text = Me.ObjCD.DigitGrouping(Me.ObjPrOB.GetQty(Me.VIdentifier), Me.vDefPrecDefault)

            Me.LblOrderUnitMulti.Left = Me.TxtOrderQtyMulti.Left + Me.TxtOrderQtyMulti.Width + 5
            Me.TxtOrderQty.Left = Me.LblOrderUnitMulti.Left + Me.LblOrderUnitMulti.Width + 5
            Me.LblUnitCode.Left = Me.TxtOrderQty.Left + Me.TxtOrderQty.Width + 5
            Me.LblOrderBalanceUnitMulti.Left = Me.TxtSOBMulti.Left + Me.TxtSOBMulti.Width + 5
            Me.TxtSOB.Left = Me.LblOrderBalanceUnitMulti.Left + Me.LblOrderBalanceUnitMulti.Width + 5
            Me.LblUnitCodeOrderBalance.Left = Me.TxtSOB.Left + Me.TxtSOB.Width + 5
            If UnitCodeMulti = UnitCode Then
                Me.LblOrderBalanceUnitMulti.Text = Me.LblOrderBalanceUnitMulti.Text.Replace("=", "")
                Me.LblOrderUnitMulti.Text = Me.LblOrderUnitMulti.Text.Replace("=", "")
                Me.TxtOrderQty.Visible = False
                Me.TxtOrderQty.TabStop = False
                Me.LblUnitCode.Visible = False
                Me.TxtSOB.Visible = False
                Me.TxtSOB.TabStop = False
                Me.LblUnitCodeOrderBalance.Visible = False
                Me.LblRatio.Visible = False
                Me.dgvDetail.Location = New Point(Me.dgvDetail.Location.X, Me.TxtSOBMulti.Top + Me.TxtSOB.Height + 5)
                Me.dgvDetail.Height = Me.btnCancel.Top - 5 - Me.dgvDetail.Location.Y
            End If
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
            'pastikan jumlah order totalnya sama dengan jumlah dgv
            JmlOrder = CDec(Me.TxtOrderQtyMulti.Text)
            JmlDGV = Me.GetJmlTotalOrder_Multi
            If JmlOrder <> JmlDGV Then
                MsgBox("Inconsistent delivery qty. Total order qty must be: " + Me.ObjCD.DigitGrouping(JmlOrder, Me.vDefPrecMulti) + " " + Me.UnitCodeMulti, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'pastikan jumlah order total qty default sama dengan jumlah dgv
            JmlOrder = Me.TxtOrderQty.Text
            JmlDGV = Me.GetJmlTotalOrder_Default
            If JmlOrder <> JmlDGV Then
                MsgBox("Inconsistent delivery qty. Total order qty must be: " + Me.ObjCD.DigitGrouping(JmlOrder, Me.vDefPrecDefault) + " " + Me.UnitCode, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'pastikan jumlah sob multi nya sama dengan dgv

            'pastikan jumlah sob-nya sama dengan dgv

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

    'Private Function GetJmlOrder_Default(ByVal vPrimaryCodeNo As Double) As Decimal
    '    Dim rowDGV As DataGridViewRow
    '    Dim RetVal As Decimal = 0

    '    Try
    '        For Each rowDGV In Me.FrmRef.dgvItem.Rows
    '            If rowDGV.Cells(Me.FrmRef.ColPrimaryCodeNo.Index).Value = Me.vPrimaryCodeNo Then
    '                RetVal += rowDGV.Cells(Me.FrmRef.ColQty.Index).Value
    '            End If
    '        Next

    '        Return RetVal
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    'Private Function GetJmlOrder(ByVal vPrimaryCodeNo As Double) As Decimal
    '    Dim rowDGV As DataGridViewRow
    '    Dim RetVal As Decimal = 0

    '    Try
    '        For Each rowDGV In Me.FrmRef.dgvItem.Rows
    '            If rowDGV.Cells(Me.FrmRef.ColPrimaryCodeNo.Index).Value = Me.vPrimaryCodeNo Then
    '                RetVal += rowDGV.Cells(Me.FrmRef.ColQtyMulti.Index).Value
    '            End If
    '        Next

    '        Return RetVal
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Private Function GetJmlTotalOrder_Multi() As Decimal
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

    Private Function GetJmlTotalOrder_Default() As Decimal
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

    Private Function GetJmlTotalOrderBalance_Multi() As Decimal
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV As DataGridViewRow In Me.dgvDetail.Rows
                RetVal += CDec(rowDGV.Cells(Me.ColQtyBalanceMulti.Index).Value)
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlTotalOrderBalance_Default() As Decimal
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV As DataGridViewRow In Me.dgvDetail.Rows
                RetVal += CDec(rowDGV.Cells(Me.ColQtyBalanceDefault.Index).Value)
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

                    Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) * _
                    CDec(Me.vRatioDefault) / CDec(Me.vRatioReference), Me.DefPrec)
                    If CDec(Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value) <= 0 Then
                        MsgBox("Qty restored to previous value, because it make default qty=0.", MsgBoxStyle.Information, "Information")
                        Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) * _
                        CDec(Me.vRatioDefault) / CDec(Me.vRatioReference), Me.DefPrec)
                        Exit Sub
                    End If

                    Me.IsikanDataBalance()
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
                        If Me.vSODate.Date > CDate(Me.dgvDetail.Item(e.ColumnIndex, e.RowIndex).Value) Then
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
                    totalOrder = Me.vQtySOBMulti
                    qtyOnDgvExCurrentRow = Me.GetJmlTotalOrder_Multi - CDec(Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value)
                    qtyRemaining = totalOrder - qtyOnDgvExCurrentRow
                    If qtyRemaining <= 0 Then
                        Return
                    End If
                    Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(qtyRemaining, Me.vDefPrecMulti)
                    Me.dgvDetail.Item(Me.ColQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvDetail.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) * _
                    Me.vRatioDefault / Me.vRatioReference, Me.vDefPrecDefault)
                    Me.IsikanDataBalance()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class