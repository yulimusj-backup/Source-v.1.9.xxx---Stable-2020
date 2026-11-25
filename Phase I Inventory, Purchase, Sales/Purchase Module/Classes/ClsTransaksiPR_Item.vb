Option Explicit On
Option Strict Off
Public Class ClsTransaksiPR_Item
    Public Sub UpdateMPSpareQty(ByVal vPrimaryTransaksi As Double, ByVal vFrm As FrmTransaksiPR)
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_item1TableAdapter
        Dim TBL As New DseTransaksiPR.tbltrans_purchase_req_item1DataTable
        Dim row As DseTransaksiPR.tbltrans_purchase_req_item1Row
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'fill item pr
            DTA.Fill(TBL, vPrimaryTransaksi)

            'update dgv -> 0
            For Each rowDGV In vFrm.DgvItem.Rows
                rowDGV.Cells(vFrm.ColItem_QtyMPSpare.Index).Value = 0
                For Each row In TBL.Rows
                    If row.primary_code_no = rowDGV.Cells(vFrm.ColItem_PrimaryProduct.Index).Value Then
                        rowDGV.Cells(vFrm.ColItem_QtyMPSpare.Index).Value = row.MP_Spare_Qty
                        Exit For
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function IsiDGVItem(ByVal vPrimaryTransaksi As Double, ByVal vDGVItem As DataGridView, _
                               ByVal i As FrmTransaksiPR, ByVal vObjCD As Core_Module.ClsCekData, _
                               ByVal vDT_ItemDelivery As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryDataTable) As Boolean

        Dim DTAItem As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim TBLItem As New DseTransaksiPR.tbltrans_purchase_req_itemDataTable
        Dim rowItem As DseTransaksiPR.tbltrans_purchase_req_itemRow
        Dim rowDGV As DataGridViewRow
        Dim DTAItemDelivery As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter

        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim ObjUnit As New Core_Module.ClsMasterUnit
        Dim ObjPR As New ClsTransaksiPR

        Try
            '   update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemDelivery.Connection.ConnectionString = DTAItem.Connection.ConnectionString

            '   autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            '   isikan data item

            vDGVItem.Rows.Clear()

            If DTAItem.FillBy_Primary(TBLItem, vPrimaryTransaksi) > 0 Then
                For Each rowItem In TBLItem

                    vDGVItem.Rows.Add()
                    vDGVItem.AllowUserToAddRows = False
                    rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)

                    Dim DefPrec As Integer = ObjProduct.GetAngkaDecimal(rowItem.primary_code_no)
                    Dim DeliveryDate_Str, Status As String

                    If CBool(rowItem.delivery_date_cont) = True Then
                        If rowItem.delivery_date = #1/1/1950# Then
                            DeliveryDate_Str = "Multiple"
                        Else
                            DeliveryDate_Str = Format(rowItem.delivery_date, "dd MMM yyyy")
                        End If
                    Else
                        DeliveryDate_Str = "-"
                    End If

                    If rowItem.status = 0 Then
                        Status = "PR Pending"
                    ElseIf rowItem.status = 1 Then
                        Status = "Purchase On Progress"
                    Else
                        Status = "Closed"
                    End If

                    With rowDGV
                        .Cells(i.ColItem_No.Index).Value = ""
                        .Cells(i.ColItem_PrimaryPR.Index).Value = rowItem.primary
                        .Cells(i.ColItem_PrimaryProduct.Index).Value = rowItem.primary_code_no
                        .Cells(i.ColItem_CodeNo.Index).Value = ObjProduct.GetProductCode(rowItem.primary_code_no)
                        .Cells(i.ColItem_Desc.Index).Value = ObjProduct.GetDescription(rowItem.primary_code_no)
                        .Cells(i.ColItem_Qty.Index).Value = ObjCD.DigitGrouping(rowItem.qty, DefPrec)
                        .Cells(i.ColItem_Unit.Index).Value = ObjUnit.GetUnitCode(rowItem.primary_unit)
                        .Cells(i.ColItem_DeliveryDateControl.Index).Value = CBool(rowItem.delivery_date_cont)
                        .Cells(i.ColItem_DeliveryDate.Index).Value = DeliveryDate_Str
                        .Cells(i.ColItem_Remarks.Index).Value = rowItem.remarks
                        If rowItem.track_using = 1 Then
                            .Cells(i.ColItem_TrackUsing.Index).Value = Boolean.TrueString
                        Else
                            .Cells(i.ColItem_TrackUsing.Index).Value = Boolean.FalseString
                        End If
                        .Cells(i.ColItem_OB.Index).Value = rowItem.OB
                        .Cells(i.ColItem_Status.Index).Value = Status
                        .Cells(i.ColItem_Identifier.Index).Value = rowItem.identifier
                    End With
                Next
            End If

            vDT_ItemDelivery.Rows.Clear()
            DTAItemDelivery.FillBy_Primary(vDT_ItemDelivery, vPrimaryTransaksi)
            vDGVItem.AllowUserToAddRows = False

            '   setnumber
            vObjCD.SetNumber(vDGVItem)

            '   autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            '   nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            DTAItemDelivery.Dispose()
            TBLItem = Nothing
        End Try
    End Function

    '''isidgv khusus PR cancel 
    Public Function IsiDGVItem2(ByVal vPrimaryTransaksi As Double, ByVal vDGVItem As DataGridView, _
                                ByVal i As FrmTransaksiPR_Cancel, ByVal vObjCD As Core_Module.ClsCekData) As Boolean

        Dim DTAItem As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim TBLItem As New DseTransaksiPR.tbltrans_purchase_req_itemDataTable
        Dim rowItem As DseTransaksiPR.tbltrans_purchase_req_itemRow
        Dim rowDGV As DataGridViewRow

        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim ObjUnit As New Core_Module.ClsMasterUnit
        Dim ObjPR As New ClsTransaksiPR

        Try
            '   update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            '   autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            '   isikan data item

            vDGVItem.Rows.Clear()

            If DTAItem.FillBy_Primary(TBLItem, vPrimaryTransaksi) > 0 Then
                For Each rowItem In TBLItem
                    If rowItem.OB > 0 Then
                        vDGVItem.Rows.Add()
                        vDGVItem.AllowUserToAddRows = False
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)

                        Dim DefPrec As Integer = ObjProduct.GetAngkaDecimal(rowItem.primary_code_no)

                        With rowDGV
                            .Cells(i.ColItem_No.Index).Value = ""
                            .Cells(i.ColItem_PrimaryCode.Index).Value = rowItem.primary_code_no
                            .Cells(i.ColItem_CodeNo.Index).Value = ObjProduct.GetProductCode(rowItem.primary_code_no)
                            .Cells(i.ColItem_Desc.Index).Value = ObjProduct.GetDescription(rowItem.primary_code_no)
                            .Cells(i.ColItem_Unit.Index).Value = ObjUnit.GetUnitCode(rowItem.primary_unit)
                            .Cells(i.ColItem_Qty.Index).Value = ObjCD.DigitGrouping(rowItem.OB, DefPrec)
                            .Cells(i.ColItem_Remarks.Index).Value = ""
                        End With
                    End If
                Next
            End If

            vDGVItem.AllowUserToAddRows = False

            '   setnumber
            vObjCD.SetNumber(vDGVItem)

            '   autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            '   nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem = Nothing
        End Try
    End Function

    Public Function GetNextIdentifier(ByVal vDGVItem As DataGridView, ByVal vPrimary As String, _
   ByVal Frm As FrmTransaksiPR) As String
        Dim rowDGV As DataGridViewRow
        Dim MaxVal As Double = 0
        Dim Index As Integer = -1
        Try
            For Each rowDGV In vDGVItem.Rows
                If rowDGV.Cells(Frm.ColItem_Identifier.Index).Value <> "" Then
                    Index = rowDGV.Cells(Frm.ColItem_Identifier.Index).Value.ToString.LastIndexOf("-")
                    If Index <> -1 Then
                        If MaxVal < CDbl(rowDGV.Cells(Frm.ColItem_Identifier.Index).Value.ToString.Substring(Index + 1)) Then
                            MaxVal = rowDGV.Cells(Frm.ColItem_Identifier.Index).Value.ToString.Substring(Index + 1)
                        End If
                    End If
                End If
            Next
            Return vPrimary + "-" + (MaxVal + 1).ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "-1-0"
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimaryTransaksi As Double, _
                                ByVal vDGVItem As DataGridView, _
                                ByVal frm As FrmTransaksiPR, _
                                ByVal vDT_ItemDelivery As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryDataTable) As Boolean
        Dim DTAItem As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim row As DataGridViewRow

        Dim DTAItemDelivery As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Dim rowItemDelivery As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryRow

        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim ObjUnit As New Core_Module.ClsMasterUnit
        Dim ObjPR As New ClsTransaksiPR

        Dim vPrimaryProduct As Double
        Dim vPrimaryUnit As Double
        Dim DelDateCont As Byte
        Dim DelDate As Date
        Dim Status As Integer
        Dim TrackUsing As Integer = 0

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemDelivery.Connection.ConnectionString = DTAItem.Connection.ConnectionString

            '--------------------------tambahan utk Identifier------------------
            Dim IdentifierLama As String = ""
            Dim Identifier As String = ""
            Dim counter As Integer = 0
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGVItem.AllowUserToAddRows = False

            'set identifier
            For Each row In vDGVItem.Rows
                With row
                    'cek apakah data identifier perlu diubah
                    IdentifierLama = row.Cells(frm.ColItem_Identifier.Index).Value.ToString
                    If IdentifierLama.StartsWith(vPrimaryTransaksi.ToString) = False Then
                        'nilai start index = 1 agar tanda minus tidak diperhitungkan
                        counter = IdentifierLama.Substring(IdentifierLama.IndexOf("-", 1) + 1)
                        Identifier = vPrimaryTransaksi.ToString + "-" + counter.ToString
                        'ganti data identifier lama dengan yang baru, untuk item
                        .Cells(frm.ColItem_Identifier.Index).Value = Identifier
                    Else
                        Identifier = IdentifierLama
                    End If
                End With
            Next

            For Each row In vDGVItem.Rows
                vDGVItem.AllowUserToAddRows = False

                vPrimaryProduct = ObjProduct.GetPrimary(row.Cells(frm.ColItem_CodeNo.Index).Value)
                vPrimaryUnit = ObjUnit.GetPrimary(row.Cells(frm.ColItem_Unit.Index).Value)

                If row.Cells(frm.ColItem_DeliveryDateControl.Index).Value = Boolean.TrueString Then
                    DelDateCont = 1
                    If IsDate(row.Cells(frm.ColItem_DeliveryDate.Index).Value) = True Then
                        DelDate = row.Cells(frm.ColItem_DeliveryDate.Index).Value
                    Else
                        DelDate = #1/1/1950#
                    End If
                Else
                    DelDateCont = 0
                    DelDate = #1/1/1950#
                End If

                If row.Cells(frm.ColItem_Status.Index).Value = "PR Pending" Then
                    Status = 0
                ElseIf row.Cells(frm.ColItem_Status.Index).Value = "Purchase On Progress" Then
                    Status = 1
                Else
                    Status = 2
                End If
                If row.Cells(frm.ColItem_TrackUsing.Index).Value = Boolean.TrueString Then
                    TrackUsing = 1
                Else
                    TrackUsing = 0
                End If
                With row
                    DTAItem.Insert(vPrimaryTransaksi, vPrimaryProduct, _
                                    .Cells(frm.ColItem_Qty.Index).Value, vPrimaryUnit, _
                                    .Cells(frm.ColItem_Remarks.Index).Value, DelDateCont, DelDate, _
                                    .Cells(frm.ColItem_Identifier.Index).Value, Status, TrackUsing, CDec(.Cells(frm.ColItem_QtyMPSpare.Index).Value))
                End With
            Next

            For Each rowItemDelivery In vDT_ItemDelivery.Rows
                With rowItemDelivery
                    If .identifier.StartsWith("-1") Then
                        .identifier = vPrimaryTransaksi.ToString + .identifier.Substring(2)
                    End If
                    DTAItemDelivery.Insert(vPrimaryTransaksi, .identifier, .delivery_qty, .delivery_date, .remarks, .delivery_priority)
                End With
            Next

            If frm.ChkRFP.Checked = True Then
                ObjPR.SetOrderBalance1(frm.TxtPrimaryPR.Text, ClsTransaksiPR.EnumReadyForPurchase.Yes)
            Else
                ObjPR.SetOrderBalance1(frm.TxtPrimaryPR.Text, ClsTransaksiPR.EnumReadyForPurchase.No)
            End If

            vDGVItem.AllowUserToAddRows = True
            vDGVItem.AllowUserToAddRows = True
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            vDT_ItemDelivery.Dispose()
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double) As Boolean

        Dim DTAItem As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim DTA_ItemDelivery As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter

        Try

            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTA_ItemDelivery.Connection.ConnectionString = DTAItem.Connection.ConnectionString

            DTAItem.Delete(vPrimaryTransaksi)
            DTA_ItemDelivery.Delete(vPrimaryTransaksi)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            DTA_ItemDelivery.Dispose()
        End Try
    End Function

    'Public Function IsiDGVItem(ByVal vPrimaryTransaksi As Double, ByVal vDGVItem As DataGridView, _
    '                            ByVal i As FrmTransaksiPR, ByVal vObjCD As Core_Module.ClsCekData, _
    '                            ByVal vDT_ItemDelivery As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryDataTable) As Boolean

    '    Dim DTAItem As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
    '    Dim TBLItem As New DseTransaksiPR.tbltrans_purchase_req_itemDataTable
    '    Dim rowItem As DseTransaksiPR.tbltrans_purchase_req_itemRow
    '    Dim rowDGV As DataGridViewRow
    '    Dim DTAItemDelivery As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter

    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Dim ObjProduct As New Core_Module.clsMasterProduct
    '    Dim ObjUnit As New Core_Module.ClsMasterUnit
    '    Dim ObjPR As New ClsTransaksiPR

    '    Try
    '        '   update connection string
    '        DTAItem.Connection.ConnectionString = Core_Module.GetConstring
    '        DTAItemDelivery.Connection.ConnectionString = DTAItem.Connection.ConnectionString

    '        '   autosize dgv
    '        vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '        vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
    '        Windows.Forms.Cursor.Current = Cursors.WaitCursor

    '        '   isikan data item

    '        vDGVItem.Rows.Clear()

    '        If DTAItem.FillBy_Primary(TBLItem, vPrimaryTransaksi) > 0 Then
    '            For Each rowItem In TBLItem

    '                vDGVItem.Rows.Add()
    '                vDGVItem.AllowUserToAddRows = False
    '                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)

    '                Dim DefPrec As Integer = ObjProduct.GetAngkaDecimal(rowItem.primary_code_no)
    '                Dim DeliveryDate_Str, Status As String

    '                If CBool(rowItem.delivery_date_cont) = True Then
    '                    If rowItem.delivery_date = #1/1/1950# Then
    '                        DeliveryDate_Str = "Multiple"
    '                    Else
    '                        DeliveryDate_Str = Format(rowItem.delivery_date, "dd MMM yyyy")
    '                    End If
    '                Else
    '                    DeliveryDate_Str = "-"
    '                End If

    '                If rowItem.status = 0 Then
    '                    Status = "PR Pending"
    '                ElseIf rowItem.status = 1 Then
    '                    Status = "Purchase On Progress"
    '                Else
    '                    Status = "Closed"
    '                End If

    '                With rowDGV
    '                    .Cells(i.ColItem_No.Index).Value = ""
    '                    .Cells(i.ColItem_PrimaryPR.Index).Value = rowItem.primary
    '                    .Cells(i.ColItem_PrimaryProduct.Index).Value = rowItem.primary_code_no
    '                    .Cells(i.ColItem_CodeNo.Index).Value = ObjProduct.GetProductCode(rowItem.primary_code_no)
    '                    .Cells(i.ColItem_Desc.Index).Value = ObjProduct.GetDescription(rowItem.primary_code_no)
    '                    .Cells(i.ColItem_Unit.Index).Value = ObjUnit.GetUnitCode(rowItem.primary_unit)
    '                    .Cells(i.ColItem_Qty.Index).Value = ObjCD.DigitGrouping(rowItem.qty, DefPrec)
    '                    .Cells(i.ColItem_DeliveryDateControl.Index).Value = CBool(rowItem.delivery_date_cont)
    '                    .Cells(i.ColItem_DeliveryDate.Index).Value = DeliveryDate_Str
    '                    .Cells(i.ColItem_OB.Index).Value = rowItem.OB
    '                    .Cells(i.ColItem_Status.Index).Value = Status
    '                    .Cells(i.ColItem_Remarks.Index).Value = rowItem.remarks
    '                    .Cells(i.ColItem_Identifier.Index).Value = rowItem.identifier
    '                    If rowItem.track_using = 1 Then
    '                        .Cells(i.ColItem_TrackUsing.Index).Value = Boolean.TrueString
    '                    Else
    '                        .Cells(i.ColItem_TrackUsing.Index).Value = Boolean.FalseString
    '                    End If
    '                End With
    '            Next
    '        End If

    '        vDT_ItemDelivery.Rows.Clear()
    '        DTAItemDelivery.FillBy_Primary(vDT_ItemDelivery, vPrimaryTransaksi)
    '        vDGVItem.AllowUserToAddRows = False

    '        '   setnumber
    '        vObjCD.SetNumber(vDGVItem)

    '        '   autosize dgv
    '        vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '        vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

    '        '   nilai balikan true
    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTAItem.Dispose()
    '        DTAItemDelivery.Dispose()
    '        TBLItem = Nothing
    '    End Try
    'End Function

    ''''isidgv khusus PR cancel 
    'Public Function IsiDGVItem2(ByVal vPrimaryTransaksi As Double, ByVal vDGVItem As DataGridView, _
    '                            ByVal i As FrmTransaksiPR_Cancel, ByVal vObjCD As Core_Module.ClsCekData) As Boolean

    '    Dim DTAItem As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
    '    Dim TBLItem As New DseTransaksiPR.tbltrans_purchase_req_itemDataTable
    '    Dim rowItem As DseTransaksiPR.tbltrans_purchase_req_itemRow
    '    Dim rowDGV As DataGridViewRow

    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Dim ObjProduct As New Core_Module.clsMasterProduct
    '    Dim ObjUnit As New Core_Module.ClsMasterUnit
    '    Dim ObjPR As New ClsTransaksiPR

    '    Try
    '        '   update connection string
    '        DTAItem.Connection.ConnectionString = Core_Module.GetConstring

    '        '   autosize dgv
    '        vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '        vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

    '        '   isikan data item

    '        vDGVItem.Rows.Clear()

    '        If DTAItem.FillBy_Primary(TBLItem, vPrimaryTransaksi) > 0 Then
    '            For Each rowItem In TBLItem
    '                If rowItem.OB > 0 Then
    '                    vDGVItem.Rows.Add()
    '                    vDGVItem.AllowUserToAddRows = False
    '                    rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)

    '                    Dim DefPrec As Integer = ObjProduct.GetAngkaDecimal(rowItem.primary_code_no)

    '                    With rowDGV
    '                        .Cells(i.ColItem_No.Index).Value = ""
    '                        .Cells(i.ColItem_PrimaryCode.Index).Value = rowItem.primary_code_no
    '                        .Cells(i.ColItem_CodeNo.Index).Value = ObjProduct.GetProductCode(rowItem.primary_code_no)
    '                        .Cells(i.ColItem_Desc.Index).Value = ObjProduct.GetDescription(rowItem.primary_code_no)
    '                        .Cells(i.ColItem_Unit.Index).Value = ObjUnit.GetUnitCode(rowItem.primary_unit)
    '                        .Cells(i.ColItem_Qty.Index).Value = ObjCD.DigitGrouping(rowItem.OB, DefPrec)
    '                        .Cells(i.ColItem_Remarks.Index).Value = ""
    '                    End With
    '                End If
    '            Next
    '        End If

    '        vDGVItem.AllowUserToAddRows = False

    '        '   setnumber
    '        vObjCD.SetNumber(vDGVItem)

    '        '   autosize dgv
    '        vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '        vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

    '        '   nilai balikan true
    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTAItem.Dispose()
    '        TBLItem = Nothing
    '    End Try
    'End Function
End Class
