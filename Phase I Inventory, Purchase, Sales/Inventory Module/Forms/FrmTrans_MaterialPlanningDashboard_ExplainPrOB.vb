Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_ExplainPrOB
    Public vPrimaryProduct As Double = -1
    Public UnitCode As String = ""
    Public CodeNo As String = ""
    Public vDT_ItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable
    Public vDT_ItemPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public AdaPerubahan As Boolean = False
    Public ReadOnly1 As Boolean = False
    Public UserCode As String = ""

    Dim DefPrec As Integer = 0
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjRefParty As New Core_Module.ClsMasterRefParty
    Dim Buffer As String = ""
    Dim ObjPR As New ClsTransaksiPR
    Dim ObjDeliveryPriority As New Core_Module.ClsMasterPriorityLevel
    Dim rowItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqRow
    Dim ObjUser As New Core_Module.ClsUser

    Private Sub FrmTrans_MaterialPlanningDashboard_ExplainPrOB_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DV As Data.DataRowView
        Dim rowDGV As DataGridViewRow
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqRow
        Dim RowItemPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryRow

        Try
            Me.ShowIcon = False
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryProduct)
            Me.vDT_ItemPrOB.DefaultView.RowFilter = "[primary product] = " + Me.vPrimaryProduct.ToString
            Me.vDT_ItemPrOB.DefaultView.Sort = "[Tanggal Transaksi], [No Transaksi]"
            For Each DV In Me.vDT_ItemPrOB.DefaultView
                row = DV.Row
                Me.dgvData.Rows.Add()
                rowDGV = Me.dgvData.Rows(Me.dgvData.Rows.Count - 1)
                rowDGV.Cells(Me.ColRemarks.Index).Value = row.remarks
                rowDGV.Cells(Me.ColDate.Index).Value = row.Tanggal_Transaksi
                rowDGV.Cells(Me.ColOrderQty.Index).Value = ObjCD.DigitGrouping(row.qty, DefPrec)
                rowDGV.Cells(Me.ColPRNumber.Index).Value = row.No_Transaksi
                rowDGV.Cells(Me.ColRefParty.Index).Value = Me.ObjRefParty.GetCode(row.primary_ref_rec) + "-" + Me.ObjRefParty.GetDesc(row.primary_ref_rec)
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvData.RowCount
                rowDGV.Cells(Me.ColSpareQty.Index).Value = ObjCD.DigitGrouping(row.MP_Spare_Qty, DefPrec)
                rowDGV.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(row.MP_Spare_Qty + row.qty, DefPrec)
                rowDGV.Cells(Me.ColPrimaryPR.Index).Value = row.primary_pr_rec
                rowDGV.Cells(Me.ColUnit.Index).Value = Me.UnitCode
                rowDGV.Cells(Me.ColIdentifier.Index).Value = row.identifier_pr_item
                If row.delivery_date_cont = 1 Then
                    rowDGV.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString
                    Me.vDT_ItemPrOBDelivery.DefaultView.RowFilter = "identifier = '" + row.identifier_pr_item + "'"
                    If Me.vDT_ItemPrOBDelivery.DefaultView.Count = 1 Then
                        RowItemPrOBDelivery = Me.vDT_ItemPrOBDelivery.DefaultView(0).Row
                        rowDGV.Cells(Me.ColDeliveryDate.Index).Value = Format(RowItemPrOBDelivery.delivery_date, "dd MMM yyyy")
                    ElseIf Me.vDT_ItemPrOBDelivery.DefaultView.Count > 1 Then
                        rowDGV.Cells(Me.ColDeliveryDate.Index).Value = "Multiple"
                    Else
                        MsgBox("Logic error: inconsistent delivery detail data.", MsgBoxStyle.Information, "Information")
                    End If
                Else
                    rowDGV.Cells(Me.ColDeliveryCont.Index).Value = Boolean.FalseString
                    rowDGV.Cells(Me.ColDeliveryDate.Index).Value = "-"
                End If

            Next

            Me.dgvData.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.Text = Me.Text + " - " + Me.CodeNo '+ " - " + Me.UnitCode
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub dgvData_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvData.CellBeginEdit
        Try
            If Me.ReadOnly1 = False Then
                If e.ColumnIndex = Me.ColTotalQty.Index Or e.ColumnIndex = Me.ColSpareQty.Index Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_EditSpareQty") = True Then
                        Me.Buffer = Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        MsgBox("You don't have access to edit this data.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                    End If
                ElseIf e.ColumnIndex = Me.ColDeliveryCont.Index Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_SF_Override_PReqETD") = True Then
                        Me.Buffer = Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        MsgBox("You don't have access to edit this data.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvData_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvData.CellEndEdit
        Try
            If IsNothing(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            Else
                If e.ColumnIndex = Me.ColTotalQty.Index Or e.ColumnIndex = Me.ColSpareQty.Index Then
                    'pastikan angka
                    If IsNumeric(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify a number.", MsgBoxStyle.Information, "Information")
                        Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'digit grouping
                    Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)

                    If e.ColumnIndex = Me.ColSpareQty.Index Then
                        'pastikan non negatif
                        If Me.ObjCD.CekNonNegatif(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Please specify a non negative number.", MsgBoxStyle.Information, "Information")
                            Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If

                        'hitung ulang qty total
                        Me.dgvData.Item(Me.ColTotalQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) + _
                        CDec(Me.dgvData.Item(Me.ColOrderQty.Index, e.RowIndex).Value), Me.DefPrec)
                    ElseIf e.ColumnIndex = Me.ColTotalQty.Index Then
                        'pastikan positif
                        If Me.ObjCD.CekBilanganPositif(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Please specify a positive number.", MsgBoxStyle.Information, "Information")
                            Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If

                        'pastikan lebih besar atau sama dengan kolom qty
                        If CDec(Me.dgvData.Item(Me.ColOrderQty.Index, e.RowIndex).Value) > CDec(Me.dgvData.Item(Me.ColTotalQty.Index, e.RowIndex).Value) Then
                            MsgBox("Please specify a number larger than " + Me.dgvData.Item(Me.ColOrderQty.Index, e.RowIndex).Value.ToString + " " + Me.dgvData.Item(Me.ColUnit.Index, e.RowIndex).Value.ToString + ".", MsgBoxStyle.Information, "Information")
                            Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If

                        'hitung spare
                        Me.dgvData.Item(Me.ColSpareQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) - _
                                          CDec(Me.dgvData.Item(Me.ColOrderQty.Index, e.RowIndex).Value), Me.DefPrec)
                    End If
                    'simpan ke database
                    Me.ObjPR.UpdateQtySpare(Me.dgvData.Item(Me.ColPrimaryPR.Index, e.RowIndex).Value, Me.vPrimaryProduct, _
                    Me.dgvData.Item(Me.ColSpareQty.Index, e.RowIndex).Value)
                    'update ke row item pr ob
                    Me.vDT_ItemPrOB.DefaultView.RowFilter = "[primary product] = " + Me.vPrimaryProduct.ToString + _
                    " AND [primary pr rec] = " + Me.dgvData.Item(Me.ColPrimaryPR.Index, e.RowIndex).Value.ToString
                    If Me.vDT_ItemPrOB.DefaultView.Count > 0 Then
                        rowItemPrOB = Me.vDT_ItemPrOB.DefaultView(0).Row
                        rowItemPrOB.MP_Spare_Qty = CDec(Me.dgvData.Item(Me.ColSpareQty.Index, e.RowIndex).Value)
                    Else
                        MsgBox("Logic error: row not found. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    End If
                    Me.AdaPerubahan = True
                ElseIf e.ColumnIndex = Me.ColDeliveryCont.Index Then
                    If CBool(Me.Buffer) = CBool(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) Then
                        Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        Me.EndEditDeliveryDateCont(sender, e, Me.dgvData.Rows(e.RowIndex))
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahMasihValid(ByVal vIdentifier As String, ByVal vPrimaryRefParty As Double, ByVal vPrimaryPR As Double, _
    ByVal vQtyOB As Decimal, ByVal vCodeNo As String, ByVal vPRNo As String, ByVal vDeliveryDateCont As Boolean) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim TBL As New DseTransaksiMaterialPlanning2.tbltrans_purchase_req_itemDataTable
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_purchase_req_itemRow
        Dim DTAItemDelivery As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Dim TBLItemDelivery As New DseTransaksiMaterialPlanning2.tbltrans_purchase_req_itemdeliveryDataTable
        Dim RowItemDelivery As DseTransaksiMaterialPlanning2.tbltrans_purchase_req_itemdeliveryRow
        Dim rowItemMPPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'sudah tidak valid jika
            '1. kombinasi primary p.req, primary product, identifier pr item, dan primary ref party sudah tidak ada
            DTA.Fill(TBL, vIdentifier)
            If TBL.Rows.Count = 0 Then
                MsgBox("Invalid purchase request order balance data: we can't find data for purchase request item with the following combination: " + vbCrLf + "Code No : " + _
                vCodeNo + " - P.Req Number: " + vPRNo + " - idenfitier: " + vIdentifier + ".", MsgBoxStyle.Information, "Information")
                Return False
            Else
                row = TBL.Rows(0)
                If row.primary_ref_party <> vPrimaryRefParty Then
                    MsgBox("Invalid purchase request order balance data: reference party does not match.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If row.primary_code_no <> vPrimaryProduct Then
                    MsgBox("Invalid purchase request order balance data: we can't find data for purchase request item with the following combination: " + vbCrLf + "Code No : " + _
                    vCodeNo + " - P.Req Number: " + vPRNo + " - idenfitier: " + vIdentifier + ".", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                '2. qty prob untuk identifier tersebut tidak cocok dengan data mp
                If row.OB <> vQtyOB Then
                    MsgBox("Invalid purchase request order balance data: order balance qty does not match.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                '3. delivery date control tidak sama
                If row.delivery_date_cont = 1 And vDeliveryDateCont = False Then
                    MsgBox("Invalid purchase request order balance data: delivery date control data does not match.", MsgBoxStyle.Information, "Information")
                    Return False
                ElseIf row.delivery_date_cont = 0 And vDeliveryDateCont = True Then
                    MsgBox("Invalid purchase request order balance data: delivery date control data does not match.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                '4. setiap baris dalam detail delivery mp, harus ada padanannya di detail delivery sob 
                '   -> tglnya sama
                DTAItemDelivery.FillBy_IdentifierPRItem(TBLItemDelivery, vIdentifier)
                Me.vDT_ItemPrOBDelivery.DefaultView.RowFilter = "identifier = '" + vIdentifier + "'"
                Me.vDT_ItemPrOBDelivery.DefaultView.Sort = "[delivery date] DESC"
                TBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                For ij As Integer = 0 To Me.vDT_ItemPrOBDelivery.DefaultView.Count - 1
                    RowItemDelivery = TBLItemDelivery.DefaultView(ij).Row
                    rowItemMPPrOBDelivery = Me.vDT_ItemPrOBDelivery.DefaultView(ij).Row
                    If RowItemDelivery.delivery_date <> rowItemMPPrOBDelivery.delivery_date Then
                        MsgBox("Invalid purchase request order balance data: detail delivery date does not match.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                Next
                '   -> pada setiap baris item so delivery berpadanan dengan item delivery so mp, kecuali baris delivery tanggal pertama, nilainya harus sama
                If Me.vDT_ItemPrOBDelivery.DefaultView.Count > 1 Then
                    For ij As Integer = 0 To Me.vDT_ItemPrOBDelivery.DefaultView.Count - 2
                        RowItemDelivery = TBLItemDelivery.DefaultView(ij).Row
                        rowItemMPPrOBDelivery = Me.vDT_ItemPrOBDelivery.DefaultView(ij).Row
                        If RowItemDelivery.delivery_qty <> rowItemMPPrOBDelivery.delivery_qty Then
                            MsgBox("Invalid purchase request order balance data: detail delivery qty does not match.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    Next
                ElseIf Me.vDT_ItemPrOBDelivery.DefaultView.Count = 1 Then
                    RowItemDelivery = TBLItemDelivery.DefaultView(0).Row
                    rowItemMPPrOBDelivery = Me.vDT_ItemPrOBDelivery.DefaultView(0).Row
                    If RowItemDelivery.delivery_qty <> rowItemMPPrOBDelivery.delivery_qty Then
                        MsgBox("Invalid purchase request order balance data: total delivery qty does not match.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub EndEditDeliveryDateCont(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
   ByVal row As DataGridViewRow)
        Dim ij As Double = 0
        Dim rowD As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryRow
        Dim rowItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqRow
        Dim TglAcuan As Date = #1/1/1950#
        Dim DefPrec As Integer = 0
        Dim TBLDelivery As New DseTransaksiMaterialPlanning2.tbltrans_purchase_req_itemdeliveryDataTable
        Dim rowDelivery As DseTransaksiMaterialPlanning2.tbltrans_purchase_req_itemdeliveryRow

        Try
            With row
                If row.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                    'jika jadi true maka delivery date jadi tanggal transaksi
                    TglAcuan = Me.ObjPR.GetTglDeliveryMinimum(row.Cells(Me.ColPrimaryPR.Index).Value)
                    If TglAcuan = #1/1/1950# Then
                        TglAcuan = Me.ObjPR.GetTanggalTransaksi(row.Cells(Me.ColPrimaryPR.Index).Value)
                    End If
                    row.Cells(Me.ColDeliveryDate.Index).Value = Format(TglAcuan, "dd MMM yyyy")
                    'tambahkan ke dalam item delivery date
                    rowD = Me.vDT_ItemPrOBDelivery.Newtbltrans_material_planning_itempurch_req_deliveryRow
                    With rowD
                        .delivery_date = TglAcuan
                        .delivery_priority = Me.ObjDeliveryPriority.GetPrimaryDefault
                        .delivery_qty = Me.dgvData.Item(Me.ColOrderQty.Index, e.RowIndex).Value
                        .identifier = Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value
                        .primary = -1
                        .primary_transaksi = -1
                        .primary_purc_req = Me.dgvData.Item(Me.ColPrimaryPR.Index, e.RowIndex).Value
                        .remarks = ""
                    End With
                    Me.vDT_ItemPrOBDelivery.Rows.Add(rowD)
                    'update ke dalam tabel sob item
                    For Each rowItemPrOB In Me.vDT_ItemPrOB.Rows
                        If rowItemPrOB.identifier_pr_item = Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value Then
                            rowItemPrOB.delivery_date_cont = 1
                            Exit For
                        End If
                    Next
                    'update ke dalam db
                    '1. set delivery date control
                    Me.ObjPR.SetDeliveryDateControl(True, TglAcuan, Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                    '2. tambahkan item delivery date
                    rowDelivery = TBLDelivery.Newtbltrans_purchase_req_itemdeliveryRow
                    With rowDelivery
                        .delivery_date = TglAcuan
                        .delivery_priority = Me.ObjDeliveryPriority.GetPrimaryDefault
                        .delivery_qty = Me.ObjPR.GetJmlOrder(Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                        .identifier = Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value
                        .primary = -1
                        .primary_transaksi = Me.dgvData.Item(Me.ColPrimaryPR.Index, e.RowIndex).Value
                        .remarks = ""
                    End With
                    TBLDelivery.Rows.Add(rowDelivery)
                    Me.ObjPR.HapusItemDetailDelivery(Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                    Me.ObjPR.SimpanItemDetailDelivery(TBLDelivery)
                Else
                    'jika jadi false maka delivery date jadi "-"
                    row.Cells(Me.ColDeliveryDate.Index).Value = "-"
                    'buang item delivery
                    While ij < Me.vDT_ItemPrOBDelivery.Rows.Count
                        rowD = Me.vDT_ItemPrOBDelivery.Rows(ij)
                        If rowD.identifier = Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value Then
                            Me.vDT_ItemPrOBDelivery.Rows.Remove(rowD)
                        Else
                            ij += 1
                        End If
                    End While
                    'update ke dalam tabel sob item
                    For Each rowItemPrOB In Me.vDT_ItemPrOB.Rows
                        If rowItemPrOB.identifier_pr_item = Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value Then
                            rowItemPrOB.delivery_date_cont = 0
                            Exit For
                        End If
                    Next
                    'update ke dalam db
                    '1. set delivery date control
                    Me.ObjPR.SetDeliveryDateControl(False, #1/1/1950#, Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                    '2. hapus item delivery date
                    Me.ObjPR.HapusItemDetailDelivery(Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class