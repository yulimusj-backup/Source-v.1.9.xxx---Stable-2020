Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_ExplainSOB
    Public vPrimaryProduct As Double = -1
    Public vDT_ItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable
    Public vDT_ItemSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryDataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UnitCode As String = ""
    Public CodeNo As String = ""
    Public AdaPerubahan As Boolean = False
    Public ReadOnly1 As Boolean = False
    Public UserCode As String = ""
    Public rowProduct As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row

    Dim rowItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobRow
    Dim DefPrec As Integer = 0
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCust As New Core_Module.clsMasterCustomer
    Dim Buffer As String = ""
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjDeliveryPriority As New Core_Module.ClsMasterPriorityLevel

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_ExplainSOB_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DV As Data.DataRowView
        Dim rowDGV As DataGridViewRow
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobRow
        Dim AdaDeliveryDate As Boolean = False
        Dim rowItemSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryRow
        Dim AdaCustomerPONo As Boolean = False

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryProduct)
            Me.vDT_ItemSOB.DefaultView.RowFilter = "[Primary Product] = " + Me.vPrimaryProduct.ToString
            Me.vDT_ItemSOB.DefaultView.Sort = "[Tanggal Transaksi], [No Transaksi]"
            Me.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            For Each DV In Me.vDT_ItemSOB.DefaultView
                row = DV.Row
                Me.dgvData.Rows.Add()
                rowDGV = Me.dgvData.Rows(Me.dgvData.Rows.Count - 1)
                rowDGV.Cells(Me.ColCustPONumber.Index).Value = Me.ObjSO.GetCustPO(row.primary_so_rec)
                If rowDGV.Cells(Me.ColCustPONumber.Index).Value <> "" Then
                    AdaCustomerPONo = True
                End If
                rowDGV.Cells(Me.ColDate.Index).Value = row.Tanggal_Transaksi
                rowDGV.Cells(Me.ColOrderQty.Index).Value = ObjCD.DigitGrouping(row.qty, DefPrec)
                rowDGV.Cells(Me.ColSONumber.Index).Value = row.No_Transaksi
                rowDGV.Cells(Me.ColCustomer.Index).Value = Me.ObjCust.GetCustName(row.primary_customer_rec)
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvData.RowCount
                rowDGV.Cells(Me.ColSpareQty.Index).Value = ObjCD.DigitGrouping(row.MP_Spare_Qty, DefPrec)
                rowDGV.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(row.MP_Spare_Qty + row.qty, DefPrec)
                rowDGV.Cells(Me.ColPrimarySO.Index).Value = row.primary_so_rec
                rowDGV.Cells(Me.ColUnit.Index).Value = Me.UnitCode
                rowDGV.Cells(Me.ColPrimaryUnit.Index).Value = Me.rowProduct.primary_unit
                rowDGV.Cells(Me.ColPrimaryUnitMulti.Index).Value = row.primary_unit_multi
                rowDGV.Cells(Me.ColRatioDefault.Index).Value = row.ratio_default
                rowDGV.Cells(Me.ColRatioReference.Index).Value = row.ratio_reference
                rowDGV.Cells(Me.ColOrderQtyMulti.Index).Value = row.qty_sob_multi
                rowDGV.Cells(Me.ColPrimaryCust.Index).Value = row.primary_customer_rec
                rowDGV.Cells(Me.ColIdentifier.Index).Value = row.identifier_so

                If row.delivery_date_cont = 1 Then
                    rowDGV.Cells(Me.ColDeliveryDateCont.Index).Value = Boolean.TrueString
                    AdaDeliveryDate = True
                    Me.vDT_ItemSOBDelivery.DefaultView.RowFilter = "identifier = '" + row.identifier_so + "'"
                    If Me.vDT_ItemSOBDelivery.DefaultView.Count = 1 Then
                        rowItemSOBDelivery = Me.vDT_ItemSOBDelivery.DefaultView(0).Row
                        rowDGV.Cells(Me.ColDeliveryDate.Index).Value = Format(rowItemSOBDelivery.delivery_date, "dd MMM yyyy")
                    Else
                        rowDGV.Cells(Me.ColDeliveryDate.Index).Value = "Multiple"
                    End If
                Else
                    rowDGV.Cells(Me.ColDeliveryDateCont.Index).Value = Boolean.FalseString
                    rowDGV.Cells(Me.ColDeliveryDate.Index).Value = "-"
                End If
                rowDGV.Cells(Me.ColIdentifier.Index).Value = row.identifier_so
            Next

            If AdaDeliveryDate = True Then
                Me.ColDeliveryDate.Visible = True
            Else
                Me.ColDeliveryDate.Visible = False
            End If
            If AdaCustomerPONo = True Then
                Me.ColCustPONumber.Visible = True
            Else
                Me.ColCustPONumber.Visible = False
            End If
            Me.dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvData.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.Text = Me.Text + " - " + Me.CodeNo '+ " - " + Me.UnitCode
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvData_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvData.CellBeginEdit
        Dim rowDGV As DataGridViewRow

        Try
            If Me.ReadOnly1 = False Then
                If e.RowIndex <> -1 Then
                    rowDGV = Me.dgvData.Rows(e.RowIndex)

                    If Me.ApakahSOBMasihValid(rowDGV.Cells(Me.ColPrimarySO.Index).Value, Me.vPrimaryProduct, rowDGV.Cells(Me.ColIdentifier.Index).Value, _
                    rowDGV.Cells(Me.ColOrderQtyMulti.Index).Value, rowDGV.Cells(Me.ColOrderQty.Index).Value, rowDGV.Cells(Me.ColRatioDefault.Index).Value, _
                    rowDGV.Cells(Me.ColRatioReference.Index).Value, rowDGV.Cells(Me.ColPrimaryCust.Index).Value, rowDGV.Cells(Me.ColSONumber.Index).Value, _
                    Me.ObjProduct.GetProductCode(Me.vPrimaryProduct), Me.rowProduct.primary_unit, rowDGV.Cells(Me.ColPrimaryUnitMulti.Index).Value, _
                    CBool(rowDGV.Cells(Me.ColDeliveryDateCont.Index).Value)) = False Then
                        e.Cancel = True
                        Exit Sub
                    End If

                    If e.ColumnIndex = Me.ColTotalQty.Index Or e.ColumnIndex = Me.ColSpareQty.Index Then
                        If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_EditSpareQty") = True Then
                            Me.Buffer = Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            MsgBox("You don't have access to edit this data.", MsgBoxStyle.Information, "Information")
                            e.Cancel = True
                        End If
                    ElseIf e.ColumnIndex = Me.ColDeliveryDateCont.Index = True Then
                        If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_SF_Override_SalesETD") = True Then
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
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvData_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvData.CellClick
        Dim ij As New FrmTrans_MaterialPlanningDashboard_ExplainSOBDelivery
        Dim DI As Integer = 0
        Dim rowD As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryRow
        Dim JmlDelivery As Integer = 0
        Dim DelDateD As Date = Today
        Dim rowDGV As DataGridViewRow
        Dim TBLDelivery As New DseTransaksiMaterialPlanning2.tbltrans_so_itemdeliveryDataTable
        Dim rowDelivery As DseTransaksiMaterialPlanning2.tbltrans_so_itemdeliveryRow

        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColDeliveryDate.Index Then
                If Me.dgvData.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value <> "-" And _
                Me.dgvData.Item(Me.ColDeliveryDateCont.Index, e.RowIndex).Value = Boolean.TrueString Then
                    rowDGV = Me.dgvData.Rows(e.RowIndex)
                    If Me.ApakahSOBMasihValid(rowDGV.Cells(Me.ColPrimarySO.Index).Value, Me.vPrimaryProduct, rowDGV.Cells(Me.ColIdentifier.Index).Value, _
                    rowDGV.Cells(Me.ColOrderQtyMulti.Index).Value, rowDGV.Cells(Me.ColOrderQty.Index).Value, rowDGV.Cells(Me.ColRatioDefault.Index).Value, _
                    rowDGV.Cells(Me.ColRatioReference.Index).Value, rowDGV.Cells(Me.ColPrimaryCust.Index).Value, rowDGV.Cells(Me.ColSONumber.Index).Value, _
                    Me.ObjProduct.GetProductCode(Me.vPrimaryProduct), Me.rowProduct.primary_unit, rowDGV.Cells(Me.ColPrimaryUnitMulti.Index).Value, _
                    CBool(rowDGV.Cells(Me.ColDeliveryDateCont.Index).Value)) = False Then
                        Exit Sub
                    End If

                    ij.vPrimaryUnitMulti = Me.dgvData.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value
                    ij.vPrimaryUnitDefault = Me.dgvData.Item(Me.ColPrimaryUnit.Index, e.RowIndex).Value
                    ij.vRatioReference = Me.dgvData.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                    ij.vRatioDefault = Me.dgvData.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                    ij.vDefPrecMulti = Me.ObjProduct.GetAngkaDecimal(vPrimaryProduct, _
                    Me.dgvData.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value)
                    ij.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryProduct)

                    ij.vPrimaryCodeNo = Me.vPrimaryProduct
                    ij.VIdentifier = Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value
                    ij.vQtySOBDefault = Me.dgvData.Item(Me.ColOrderQty.Index, e.RowIndex).Value
                    ij.vQtySOBMulti = Me.dgvData.Item(Me.ColOrderQtyMulti.Index, e.RowIndex).Value
                    ij.FrmRef = Me
                    If Me.ReadOnly1 = True Then
                        ij.vReadOnly = True
                        ij.ShowDialog()
                        Exit Sub
                    End If
                    ij.vSODate = Me.dgvData.Item(Me.ColDate.Index, e.RowIndex).Value

                    If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                        'hapuskan semua item delivery yang identifiernya di edit
                        DI = 0
                        While DI <= Me.vDT_ItemSOBDelivery.Rows.Count - 1
                            rowD = Me.vDT_ItemSOBDelivery.Rows(DI)
                            If rowD.identifier = ij.VIdentifier Then
                                Me.vDT_ItemSOBDelivery.Rows.Remove(rowD)
                            Else
                                DI += 1
                            End If
                        End While

                        'masukkan data item delivery baru -> disini harus dimodif
                        'MsgBox("belum")
                        ij.dgvDetail.AllowUserToAddRows = False
                        For Each rowDGV In ij.dgvDetail.Rows
                            If rowDGV.Cells(ij.ColQtyBalanceMulti.Index).Value <> 0 Then
                                rowD = Me.vDT_ItemSOBDelivery.Newtbltrans_material_planning_itemsob_deliveryRow
                                With rowD
                                    .delivery_date = CDate(rowDGV.Cells(ij.ColDeliveryDate.Index).Value)
                                    .delivery_priority = rowDGV.Cells(ij.ColPrimaryPriority.Index).Value
                                    .delivery_qty = rowDGV.Cells(ij.ColQtyBalanceDefault.Index).Value
                                    .identifier = ij.VIdentifier
                                    .primary = -1
                                    .primary_transaksi = -1
                                    .primary_so = Me.dgvData.Item(Me.ColPrimarySO.Index, e.RowIndex).Value
                                    .remarks = rowDGV.Cells(ij.ColRemarks.Index).Value
                                    .delivery_qty_multi = rowDGV.Cells(ij.ColQtyBalanceMulti.Index).Value
                                End With
                                Me.vDT_ItemSOBDelivery.Rows.Add(rowD)
                                JmlDelivery += 1
                                DelDateD = CDate(rowDGV.Cells(ij.ColDeliveryDate.Index).Value)
                            End If
                        Next

                        'simpan ke dalam db
                        TBLDelivery.Rows.Clear()
                        For Each rowDGV In ij.dgvDetail.Rows
                            rowDelivery = TBLDelivery.Newtbltrans_so_itemdeliveryRow
                            With rowDelivery
                                .delivery_date = CDate(rowDGV.Cells(ij.ColDeliveryDate.Index).Value)
                                .delivery_priority = rowDGV.Cells(ij.ColPrimaryPriority.Index).Value
                                .delivery_qty = rowDGV.Cells(ij.ColQty.Index).Value
                                .delivery_qty_multi = rowDGV.Cells(ij.ColQtyMulti.Index).Value
                                .identifier = ij.VIdentifier
                                .primary = -1
                                .primary_transaksi = Me.dgvData.Item(Me.ColPrimarySO.Index, e.RowIndex).Value
                                .remarks = rowDGV.Cells(ij.ColRemarks.Index).Value
                                'If row.Cells(Me.ColPrimaryUnit.Index).Value = row.Cells(Me.ColPrimaryUnitMulti.Index).Value Then
                                '    .delivery_qty_multi = Me.dgvData.Item(Me.ColOrderQty.Index, e.RowIndex).Value
                                'Else
                                '    .delivery_qty_multi = Me.dgvData.Item(Me.ColOrderQtyMulti.Index, e.RowIndex).Value
                                'End If
                            End With
                            TBLDelivery.Rows.Add(rowDelivery)
                        Next

                        Me.ObjSO.HapusItemDetailDelivery(Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                        Me.ObjSO.SimpanItemDetailDelivery(TBLDelivery)

                        If JmlDelivery > 1 Then
                            Me.dgvData.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value = "Multiple"
                            Me.ObjSO.SetDeliveryDateControl(True, #1/1/1950#, Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                        ElseIf JmlDelivery = 1 Then
                            Me.dgvData.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value = Format(DelDateD, "dd MMM yyyy")
                            Me.ObjSO.SetDeliveryDateControl(True, DelDateD, Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                        End If
                    End If
                End If
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
                If e.ColumnIndex = Me.ColSpareQty.Index Or e.ColumnIndex = Me.ColTotalQty.Index Then
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
                    Me.ObjSO.UpdateQtySpare(Me.dgvData.Item(Me.ColPrimarySO.Index, e.RowIndex).Value, Me.vPrimaryProduct, _
                    Me.dgvData.Item(Me.ColSpareQty.Index, e.RowIndex).Value)
                    'update ke row item pr ob
                    Me.vDT_ItemSOB.DefaultView.RowFilter = "[primary product] = " + Me.vPrimaryProduct.ToString + _
                    " AND [primary so rec] = " + Me.dgvData.Item(Me.ColPrimarySO.Index, e.RowIndex).Value.ToString
                    If Me.vDT_ItemSOB.DefaultView.Count > 0 Then
                        rowItemSOB = Me.vDT_ItemSOB.DefaultView(0).Row
                        rowItemSOB.MP_Spare_Qty = CDec(Me.dgvData.Item(Me.ColSpareQty.Index, e.RowIndex).Value)
                    Else
                        MsgBox("Logic error: row not found. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    End If
                    Me.AdaPerubahan = True
                ElseIf e.ColumnIndex = Me.ColDeliveryDateCont.Index Then
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

    Private Function ApakahSOBMasihValid(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, ByVal vIdentifierSO As String, ByVal vQTYSOBMulti As Decimal, _
    ByVal vQtySOB As Decimal, ByVal vRatioDefault As Decimal, ByVal vRatioReference As Decimal, ByVal vPrimaryCustomer As Double, ByVal vSONumber As String, _
    ByVal vCodeNo As String, ByVal vPrimaryUnit As Double, ByVal vPrimaryUnitMulti As Double, ByVal vDeliveryDateCont As Boolean) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_so_itemTableAdapter
        Dim TBL As New DseTransaksiMaterialPlanning2.tbltrans_so_itemDataTable
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_so_itemRow
        Dim DTAItemDelivery As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim TBLItemDelivery As New DseTransaksiMaterialPlanning2.tbltrans_so_itemdeliveryDataTable
        Dim RowItemDelivery As DseTransaksiMaterialPlanning2.tbltrans_so_itemdeliveryRow
        Dim rowItemMPSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryRow
        
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'sudah tidak valid jika
            '1. kombinasi primary so, primary product, primary unit multi, primary customer dan identifier so sudah tidak ada
            DTA.Fill(TBL, vPrimarySO, vPrimaryProduct, vIdentifierSO)
            If TBL.Rows.Count = 0 Then
                MsgBox("Invalid sales order: we can't find data for sales order item with the following combination: " + vbCrLf + "Code No : " + _
                vCodeNo + " - SO Number: " + vSONumber + " - idenfitier: " + vIdentifierSO, MsgBoxStyle.Information, "Information")
                Return False
            End If

            '2. qty sob untuk identifier tersebut tidak cocok dengan data mp
            row = TBL.Rows(0)
            If row.SOB <> vQtySOB Or row.SOB_Multi <> vQTYSOBMulti Then
                MsgBox("Invalid sales order balance data: balance qty does not match.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '3. ratio default dan ratio multi untuk item so tersebut tidak cocok dengan data di mp
            If row.Ratio_Default <> vRatioDefault Or row.Ratio_Reference <> vRatioReference Then
                MsgBox("Invalid sales order balance data: ratio data does not match.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '4 unit code/ multi uni code tidak sama
            If row.Primary_Units <> vPrimaryUnit Or row.Primary_Units_Multi <> vPrimaryUnitMulti Then
                MsgBox("Invalid sales order balance data: unit code or multi unit code does not match.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '5 customer tidak sama
            If row.Primary_Customer <> vPrimaryCustomer Then
                MsgBox("Invalid sales order balance data: customer code does not match.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '6. delivery date control tidak sama
            If row.Delivery_Date_Cont = True And vDeliveryDateCont = False Then
                MsgBox("Invalid sales order balance data: delivery date control does not match.", MsgBoxStyle.Information, "Information")
                Return False
            ElseIf row.Delivery_Date_Cont = False And vDeliveryDateCont = True Then
                MsgBox("Invalid sales order balance data: delivery date control does not match.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '7. setiap baris dalam detail delivery mp, harus ada padanannya di detail delivery sob 
            '   -> jumlah barisnya sama -> g jadi, karena kemungkinan ini valid
            DTAItemDelivery.FillBy_identifier_so(TBLItemDelivery, vIdentifierSO)
            Me.vDT_ItemSOBDelivery.DefaultView.RowFilter = "identifier = '" + vIdentifierSO + "'"
            Me.vDT_ItemSOBDelivery.DefaultView.Sort = "[delivery date] DESC"
            TBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"
            'If Me.vDT_ItemSOBDelivery.DefaultView.Count <> TBLItemDelivery.Rows.Count Then
            '    MsgBox("Invalid sales order balance data: detail delivery row count does not match.", MsgBoxStyle.Information, "Information")
            '    Return False
            'End If
            '   -> tglnya sama
            For ij As Integer = 0 To Me.vDT_ItemSOBDelivery.DefaultView.Count - 1
                RowItemDelivery = TBLItemDelivery.DefaultView(ij).Row
                rowItemMPSOBDelivery = Me.vDT_ItemSOBDelivery.DefaultView(ij).Row
                If RowItemDelivery.delivery_date <> rowItemMPSOBDelivery.delivery_date Then
                    MsgBox("Invalid sales order balance data: detail delivery date does not match.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Next
            '   -> pada setiap baris item so delivery berpadanan dengan item delivery so mp, kecuali baris delivery tanggal pertama, nilainya harus sama
            If Me.vDT_ItemSOBDelivery.DefaultView.Count > 1 Then
                For ij As Integer = 0 To Me.vDT_ItemSOBDelivery.DefaultView.Count - 2
                    RowItemDelivery = TBLItemDelivery.DefaultView(ij).Row
                    rowItemMPSOBDelivery = Me.vDT_ItemSOBDelivery.DefaultView(ij).Row
                    If RowItemDelivery.delivery_qty <> rowItemMPSOBDelivery.delivery_qty Then
                        MsgBox("Invalid sales order balance data: detail delivery qty does not match.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                Next
            ElseIf Me.vDT_ItemSOBDelivery.DefaultView.Count = 1 Then
                RowItemDelivery = TBLItemDelivery.DefaultView(0).Row
                rowItemMPSOBDelivery = Me.vDT_ItemSOBDelivery.DefaultView(0).Row
                If RowItemDelivery.delivery_qty < rowItemMPSOBDelivery.delivery_qty Then
                    MsgBox("Invalid sales order balance data: total delivery qty does not match.", MsgBoxStyle.Information, "Information")
                    Return False
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
        Dim rowD As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryRow
        Dim rowItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobRow
        Dim TglAcuan As Date = #1/1/1950#
        Dim DefPrec As Integer = 0
        Dim TBLDelivery As New DseTransaksiMaterialPlanning2.tbltrans_so_itemdeliveryDataTable
        Dim rowDelivery As DseTransaksiMaterialPlanning2.tbltrans_so_itemdeliveryRow

        Try
            With row
                If row.Cells(Me.ColDeliveryDateCont.Index).Value = Boolean.TrueString Then
                    'jika jadi true maka delivery date jadi tanggal transaksi
                    TglAcuan = Me.ObjSO.GetDeliveryDateMinimum(row.Cells(Me.ColPrimarySO.Index).Value)
                    If TglAcuan = #1/1/1950# Then
                        TglAcuan = Me.ObjSO.GetTanggalTransaksi(row.Cells(Me.ColIdentifier.Index).Value)
                    End If
                    row.Cells(Me.ColDeliveryDate.Index).Value = Format(TglAcuan, "dd MMM yyyy")
                    'tambahkan ke dalam item delivery date
                    rowD = Me.vDT_ItemSOBDelivery.Newtbltrans_material_planning_itemsob_deliveryRow
                    With rowD
                        .delivery_date = TglAcuan
                        .delivery_priority = Me.ObjDeliveryPriority.GetPrimaryDefault
                        .delivery_qty = Me.dgvData.Item(Me.ColOrderQty.Index, e.RowIndex).Value
                        .identifier = Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value
                        .primary = -1
                        .primary_transaksi = -1
                        .primary_so = Me.dgvData.Item(Me.ColPrimarySO.Index, e.RowIndex).Value
                        .remarks = ""
                        If row.Cells(Me.ColPrimaryUnit.Index).Value = row.Cells(Me.ColPrimaryUnitMulti.Index).Value Then
                            .delivery_qty_multi = Me.dgvData.Item(Me.ColOrderQty.Index, e.RowIndex).Value
                        Else
                            .delivery_qty_multi = Me.dgvData.Item(Me.ColOrderQtyMulti.Index, e.RowIndex).Value
                        End If
                    End With
                    Me.vDT_ItemSOBDelivery.Rows.Add(rowD)
                    'update ke dalam tabel sob item
                    For Each rowItemSOB In Me.vDT_ItemSOB.Rows
                        If rowItemSOB.identifier_so = Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value Then
                            rowItemSOB.delivery_date_cont = 1
                            Exit For
                        End If
                    Next
                    'update ke dalam db
                    '1. set delivery date control
                    Me.ObjSO.SetDeliveryDateControl(True, TglAcuan, Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                    '2. tambahkan item delivery date
                    rowDelivery = TBLDelivery.Newtbltrans_so_itemdeliveryRow
                    With rowDelivery
                        .delivery_date = TglAcuan
                        .delivery_priority = Me.ObjDeliveryPriority.GetPrimaryDefault
                        .delivery_qty = Me.ObjSO.GetQty(Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                        .delivery_qty_multi = Me.ObjSO.GetQtyMulti(Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                        .identifier = Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value
                        .primary = -1
                        .primary_transaksi =Me.dgvData.Item(Me.ColPrimarySO.Index, e.RowIndex).Value
                        .remarks = ""
                        'If row.Cells(Me.ColPrimaryUnit.Index).Value = row.Cells(Me.ColPrimaryUnitMulti.Index).Value Then
                        '    .delivery_qty_multi = Me.dgvData.Item(Me.ColOrderQty.Index, e.RowIndex).Value
                        'Else
                        '    .delivery_qty_multi = Me.dgvData.Item(Me.ColOrderQtyMulti.Index, e.RowIndex).Value
                        'End If
                    End With
                    TBLDelivery.Rows.Add(rowDelivery)
                    Me.ObjSO.HapusItemDetailDelivery(Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                    Me.ObjSO.SimpanItemDetailDelivery(TBLDelivery)
                Else
                    'jika jadi false maka delivery date jadi "-"
                    row.Cells(Me.ColDeliveryDate.Index).Value = "-"
                    'buang item delivery
                    While ij < Me.vDT_ItemSOBDelivery.Rows.Count
                        rowD = Me.vDT_ItemSOBDelivery.Rows(ij)
                        If rowD.identifier = Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value Then
                            Me.vDT_ItemSOBDelivery.Rows.Remove(rowD)
                        Else
                            ij += 1
                        End If
                    End While
                    'update ke dalam tabel sob item
                    For Each rowItemSOB In Me.vDT_ItemSOB.Rows
                        If rowItemSOB.identifier_so = Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value Then
                            rowItemSOB.delivery_date_cont = 0
                            Exit For
                        End If
                    Next
                    'update ke dalam db
                    '1. set delivery date control
                    Me.ObjSO.SetDeliveryDateControl(False, #1/1/1950#, Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                    '2. hapus item delivery date
                    Me.ObjSO.HapusItemDetailDelivery(Me.dgvData.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class