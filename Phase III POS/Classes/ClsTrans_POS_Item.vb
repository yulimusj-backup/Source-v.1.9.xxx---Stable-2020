Option Explicit On
Option Strict Off
Imports Core_Module

Public Class ClsTrans_POS_Item
    Private mPrimaryAutoNumber As Double = -1

    Public Property PrimaryAutoNumber() As Double
        Get
            Return mPrimaryAutoNumber
        End Get
        Set(ByVal value As Double)
            mPrimaryAutoNumber = value
        End Set
    End Property

    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307"
        End Try
    End Function

    Public Function GetJmlByPrimaryProduct(ByVal vPrimaryProduct As Double) As Decimal
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPrimaryDisc1(ByVal vPrimaryDiskon As Double) As Decimal
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryDisc1(vPrimaryDiskon).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryDisc1(vPrimaryDiskon).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPrimaryDisc2(ByVal vPrimaryDiskon As Double) As Decimal
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryDisc2(vPrimaryDiskon).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryDisc2(vPrimaryDiskon).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPrimaryDisc3(ByVal vPrimaryDiskon As Double) As Decimal
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryDisc3(vPrimaryDiskon).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryDisc3(vPrimaryDiskon).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTrans_POS, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseTrans_POSTableAdapters.DataTable2TableAdapter
        Dim TBLItemP As New DseTrans_POS.DataTable2DataTable
        Dim rowItemP As DseTrans_POS.DataTable2Row
        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'isikan data item
            DTAItemP.Fill(TBLItemP, vPrimary)
            vDGVItem.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColCodeNo.Index).Value = rowItemP.Product_Code
                    .Cells(i.ColDesc.Index).Value = rowItemP.Description
                    .Cells(i.ColLocation.Index).Value = rowItemP.Storage_Area_Code
                    .Cells(i.ColBatchNumber.Index).Value = rowItemP.batch_number
                    .Cells(i.ColExpiry.Index).Value = Format(rowItemP.expiry_date, "dd MMM yyyy")
                    .Cells(i.ColQty.Index).Value = vObjCD.DigitGrouping(rowItemP.qty, rowItemP._Decimal)
                    .Cells(i.ColUnit.Index).Value = rowItemP.Unit_Code
                    .Cells(i.ColUnitPrice.Index).Value = vObjCD.DigitGrouping(rowItemP.unit_price, rowItemP.Default_Precision)
                    .Cells(i.ColDisc_Amount.Index).Value = vObjCD.DigitGrouping(rowItemP.discount, rowItemP.Default_Precision)
                    .Cells(i.ColNettPrice.Index).Value = vObjCD.DigitGrouping(rowItemP.nett_price, rowItemP.Default_Precision)
                    .Cells(i.ColSubTotal.Index).Value = vObjCD.DigitGrouping(rowItemP.sub_total, rowItemP.Default_Precision)
                    .Cells(i.ColPrimaryCodeNo.Index).Value = rowItemP.primary_product
                    .Cells(i.ColPrimaryUnit.Index).Value = rowItemP.primary_unit
                    .Cells(i.ColPrimaryDisc1.Index).Value = rowItemP.primary_disc_1
                    .Cells(i.ColPrimaryDisc2.Index).Value = rowItemP.primary_disc_2
                    .Cells(i.ColPrimaryDisc3.Index).Value = rowItemP.primary_disc_3
                    .Cells(i.ColDiscAmount1.Index).Value = rowItemP.disc_amount_1
                    .Cells(i.ColDiscAmount2.Index).Value = rowItemP.disc_amount_2
                    .Cells(i.ColDiscAmount3.Index).Value = rowItemP.disc_amount_3
                    .Cells(i.ColDiscString1.Index).Value = rowItemP.disc_string_1
                    .Cells(i.ColDiscString2.Index).Value = rowItemP.disc_string_2
                    .Cells(i.ColDiscString3.Index).Value = rowItemP.disc_string_3
                    .Cells(i.ColDiscString.Index).Value = rowItemP.disc_string
                    .Cells(i.ColDisc_Amount.Index).Value = vObjCD.DigitGrouping(rowItemP.discount, rowItemP.Default_Precision)
                    .Cells(i.ColPrimaryLocation.Index).Value = rowItemP.primary_location
                    .Cells(i.ColPrimaryAutoNumber.Index).Value = rowItemP.primary
                End With
            Next

            vObjCD.SetNumber(vDGVItem)

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemP.Dispose()
            TBLItemP.Dispose()
        End Try
    End Function

    Public Function IsiDGVPOSSession(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTrans_Session_OpenClose, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseTrans_POSTableAdapters.DataTable2TableAdapter
        Dim TBLItemP As New DseTrans_POS.DataTable2DataTable
        Dim rowItemP As DseTrans_POS.DataTable2Row
        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'isikan data item
            DTAItemP.Fill(TBLItemP, vPrimary)
            vDGVItem.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColCodeNo.Index).Value = rowItemP.Product_Code
                    .Cells(i.ColDesc.Index).Value = rowItemP.Description
                    .Cells(i.ColLocation.Index).Value = rowItemP.Storage_Area_Code
                    .Cells(i.ColBatchNumber.Index).Value = rowItemP.batch_number
                    .Cells(i.ColExpiry.Index).Value = Format(rowItemP.expiry_date, "dd MMM yyyy")
                    .Cells(i.ColQty.Index).Value = vObjCD.DigitGrouping(rowItemP.qty, rowItemP._Decimal)
                    .Cells(i.ColUnit.Index).Value = rowItemP.Unit_Code
                    .Cells(i.ColUnitPrice.Index).Value = vObjCD.DigitGrouping(rowItemP.unit_price, rowItemP.Default_Precision)
                    .Cells(i.ColDisc_Amount.Index).Value = vObjCD.DigitGrouping(rowItemP.discount, rowItemP.Default_Precision)
                    .Cells(i.ColNettPrice.Index).Value = vObjCD.DigitGrouping(rowItemP.nett_price, rowItemP.Default_Precision)
                    .Cells(i.ColSubTotal.Index).Value = vObjCD.DigitGrouping(rowItemP.sub_total, rowItemP.Default_Precision)
                    .Cells(i.ColPrimaryCodeNo.Index).Value = rowItemP.primary_product
                    .Cells(i.ColPrimaryUnit.Index).Value = rowItemP.primary_unit
                    .Cells(i.ColPrimaryDisc1.Index).Value = rowItemP.primary_disc_1
                    .Cells(i.ColPrimaryDisc2.Index).Value = rowItemP.primary_disc_2
                    .Cells(i.ColPrimaryDisc3.Index).Value = rowItemP.primary_disc_3
                    .Cells(i.ColDiscAmount1.Index).Value = rowItemP.disc_amount_1
                    .Cells(i.ColDiscAmount2.Index).Value = rowItemP.disc_amount_2
                    .Cells(i.ColDiscAmount3.Index).Value = rowItemP.disc_amount_3
                    .Cells(i.ColDiscString1.Index).Value = rowItemP.disc_string_1
                    .Cells(i.ColDiscString2.Index).Value = rowItemP.disc_string_2
                    .Cells(i.ColDiscString3.Index).Value = rowItemP.disc_string_3
                    .Cells(i.ColDiscString.Index).Value = rowItemP.disc_string
                    .Cells(i.ColDisc_Amount.Index).Value = vObjCD.DigitGrouping(rowItemP.discount, rowItemP.Default_Precision)
                    .Cells(i.ColPrimaryLocation.Index).Value = rowItemP.primary_location
                    .Cells(i.ColPrimaryAutoNumber.Index).Value = rowItemP.primary
                End With
            Next

            vObjCD.SetNumber(vDGVItem)

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemP.Dispose()
            TBLItemP.Dispose()
        End Try
    End Function

    Public Function HapusItemFullTransaksi(ByVal vPrimary As Double) As Boolean
        Dim DTAP As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Dim DTAB As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_item_paymentTableAdapter
        Try
            DTAP.Connection.ConnectionString = Me.GetConstring
            DTAB.Connection.ConnectionString = Me.GetConstring
            DTAP.Delete(vPrimary)
            DTAB.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAP.Dispose()
            DTAB.Dispose()
        End Try
    End Function

    Public Function HapusItemPerBaris(ByVal vPrimaryAutoNumber As Double) As Boolean
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.DeleteQuery_PrimaryAutoNumber(vPrimaryAutoNumber)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimary As Double, ByVal vDGVItemProduct As DataGridView, _
    ByVal i As FrmTrans_POS) As Boolean
        Dim DTAItemP As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Try
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            For Each rowDGV In vDGVItemProduct.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColPrimaryDisc1.Index).Value) = True Then
                    rowDGV.Cells(i.ColPrimaryDisc1.Index).Value = "-1"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColPrimaryDisc2.Index).Value) = True Then
                    rowDGV.Cells(i.ColPrimaryDisc2.Index).Value = "-1"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColPrimaryDisc3.Index).Value) = True Then
                    rowDGV.Cells(i.ColPrimaryDisc3.Index).Value = "-1"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscAmount1.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscAmount1.Index).Value = "0"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscAmount2.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscAmount2.Index).Value = "0"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscAmount3.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscAmount3.Index).Value = "0"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscString1.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscString1.Index).Value = ""
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscString2.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscString2.Index).Value = ""
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscString3.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscString3.Index).Value = ""
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColExpiry.Index).Value) = True Then
                    rowDGV.Cells(i.ColExpiry.Index).Value = "1950/01/01"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColPrimaryLocation.Index).Value) = True Then
                    rowDGV.Cells(i.ColPrimaryLocation.Index).Value = "-1"
                End If
                With rowDGV
                    DTAItemP.Insert(vPrimary, .Cells(i.ColPrimaryCodeNo.Index).Value, .Cells(i.ColQty.Index).Value, .Cells(i.ColUnitPrice.Index).Value, _
                    .Cells(i.ColDiscString.Index).Value, .Cells(i.ColDisc_Amount.Index).Value, .Cells(i.ColNettPrice.Index).Value, .Cells(i.ColSubTotal.Index).Value, _
                    .Cells(i.ColPrimaryDisc1.Index).Value, .Cells(i.ColDiscString1.Index).Value, .Cells(i.ColDiscAmount1.Index).Value, _
                    .Cells(i.ColPrimaryDisc2.Index).Value, .Cells(i.ColDiscString2.Index).Value, .Cells(i.ColDiscAmount2.Index).Value, _
                    .Cells(i.ColPrimaryDisc3.Index).Value, .Cells(i.ColDiscString3.Index).Value, .Cells(i.ColDiscAmount3.Index).Value, _
                    .Cells(i.ColPrimaryLocation.Index).Value, .Cells(i.ColBatchNumber.Index).Value, CDate(.Cells(i.ColExpiry.Index).Value), _
                    .Cells(i.ColPrimaryUnit.Index).Value)
                End With
            Next
            'Karena cuma satu Item dan itu pertama, maka lakukan get primary biasa
            Me.PrimaryAutoNumber = Me.GetPrimaryAutoNumberByPrimaryTransaksi(vPrimary)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemP.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetPrimaryAutoNumberByPrimaryTransaksi(ByVal vPrimaryTrans As Double) As Double
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimaryAutoNumberByPrimaryTrans(vPrimaryTrans).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryAutoNumberByPrimaryTrans(vPrimaryTrans).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanItemSatuSatu(ByVal vPrimary As Double, ByVal vDGVItemProduct As DataGridView, _
    ByVal RowIndexDGV As Integer, ByVal i As FrmTrans_POS) As Boolean
        Dim DTAItemP As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Dim ObjCD As New ClsCekData
        Try
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            If String.IsNullOrEmpty(vDGVItemProduct.Item(i.ColPrimaryDisc1.Index, RowIndexDGV).Value) = True Then
                vDGVItemProduct.Item(i.ColPrimaryDisc1.Index, RowIndexDGV).Value = "-1"
            End If
            If String.IsNullOrEmpty(vDGVItemProduct.Item(i.ColPrimaryDisc2.Index, RowIndexDGV).Value) = True Then
                vDGVItemProduct.Item(i.ColPrimaryDisc2.Index, RowIndexDGV).Value = "-1"
            End If
            If String.IsNullOrEmpty(vDGVItemProduct.Item(i.ColPrimaryDisc3.Index, RowIndexDGV).Value) = True Then
                vDGVItemProduct.Item(i.ColPrimaryDisc3.Index, RowIndexDGV).Value = "-1"
            End If
            If String.IsNullOrEmpty(vDGVItemProduct.Item(i.ColDiscAmount1.Index, RowIndexDGV).Value) = True Then
                vDGVItemProduct.Item(i.ColDiscAmount1.Index, RowIndexDGV).Value = "0"
            End If
            If String.IsNullOrEmpty(vDGVItemProduct.Item(i.ColDiscAmount2.Index, RowIndexDGV).Value) = True Then
                vDGVItemProduct.Item(i.ColDiscAmount2.Index, RowIndexDGV).Value = "0"
            End If
            If String.IsNullOrEmpty(vDGVItemProduct.Item(i.ColDiscAmount3.Index, RowIndexDGV).Value) = True Then
                vDGVItemProduct.Item(i.ColDiscAmount3.Index, RowIndexDGV).Value = "0"
            End If
            If String.IsNullOrEmpty(vDGVItemProduct.Item(i.ColDiscString1.Index, RowIndexDGV).Value) = True Then
                vDGVItemProduct.Item(i.ColDiscString1.Index, RowIndexDGV).Value = ""
            End If
            If String.IsNullOrEmpty(vDGVItemProduct.Item(i.ColDiscString2.Index, RowIndexDGV).Value) = True Then
                vDGVItemProduct.Item(i.ColDiscString2.Index, RowIndexDGV).Value = ""
            End If
            If String.IsNullOrEmpty(vDGVItemProduct.Item(i.ColDiscString3.Index, RowIndexDGV).Value) = True Then
                vDGVItemProduct.Item(i.ColDiscString3.Index, RowIndexDGV).Value = ""
            End If
            If String.IsNullOrEmpty(vDGVItemProduct.Item(i.ColExpiry.Index, RowIndexDGV).Value) = True Then
                vDGVItemProduct.Item(i.ColExpiry.Index, RowIndexDGV).Value = "1950/01/01"
            End If
            If String.IsNullOrEmpty(vDGVItemProduct.Item(i.ColPrimaryLocation.Index, RowIndexDGV).Value) = True Then
                vDGVItemProduct.Item(i.ColPrimaryLocation.Index, RowIndexDGV).Value = "-1"
            End If
            DTAItemP.Insert(vPrimary, vDGVItemProduct.Item(i.ColPrimaryCodeNo.Index, RowIndexDGV).Value, vDGVItemProduct.Item(i.ColQty.Index, RowIndexDGV).Value, _
            vDGVItemProduct.Item(i.ColUnitPrice.Index, RowIndexDGV).Value, vDGVItemProduct.Item(i.ColDiscString.Index, RowIndexDGV).Value, _
            vDGVItemProduct.Item(i.ColDisc_Amount.Index, RowIndexDGV).Value, vDGVItemProduct.Item(i.ColNettPrice.Index, RowIndexDGV).Value, _
            vDGVItemProduct.Item(i.ColSubTotal.Index, RowIndexDGV).Value, vDGVItemProduct.Item(i.ColPrimaryDisc1.Index, RowIndexDGV).Value, _
            vDGVItemProduct.Item(i.ColDiscString1.Index, RowIndexDGV).Value, vDGVItemProduct.Item(i.ColDiscAmount1.Index, RowIndexDGV).Value, _
            vDGVItemProduct.Item(i.ColPrimaryDisc2.Index, RowIndexDGV).Value, vDGVItemProduct.Item(i.ColDiscString2.Index, RowIndexDGV).Value, _
            vDGVItemProduct.Item(i.ColDiscAmount2.Index, RowIndexDGV).Value, vDGVItemProduct.Item(i.ColPrimaryDisc3.Index, RowIndexDGV).Value, _
            vDGVItemProduct.Item(i.ColDiscString3.Index, RowIndexDGV).Value, vDGVItemProduct.Item(i.ColDiscAmount3.Index, RowIndexDGV).Value, _
            vDGVItemProduct.Item(i.ColPrimaryLocation.Index, RowIndexDGV).Value, vDGVItemProduct.Item(i.ColBatchNumber.Index, RowIndexDGV).Value, _
            CDate(vDGVItemProduct.Item(i.ColExpiry.Index, RowIndexDGV).Value), vDGVItemProduct.Item(i.ColPrimaryUnit.Index, RowIndexDGV).Value)
            '-----
            Me.PrimaryAutoNumber = Me.GetPrimaryAutoNumber(vPrimary)

            'Sampai sini oke
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemP.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetPrimaryAutoNumber(ByVal vPrimaryTransaksi As Double) As Double
        Dim DTAItem As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Me.GetConstring
            If DTAItem.ScalarQuery_GetMaxPrimaryAutoNumber(vPrimaryTransaksi).HasValue Then
                Return DTAItem.ScalarQuery_GetMaxPrimaryAutoNumber(vPrimaryTransaksi).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanItemPaymentMultiple(ByVal vPrimary As Double, ByVal vDGVItemProduct As DataGridView, _
    ByVal i As FrmTrans_POS_Payment) As Boolean
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Dim DTAItemB As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_item_paymentTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTAItemB.Connection.ConnectionString = Me.GetConstring

            For Each rowDGV In vDGVItemProduct.Rows
                With rowDGV
                    If String.IsNullOrEmpty(.Cells(i.ColApprovalCode.Index).Value) = True Then
                        .Cells(i.ColApprovalCode.Index).Value = ""
                    End If
                    If String.IsNullOrEmpty(.Cells(i.ColAccountNumber.Index).Value) = True Then
                        .Cells(i.ColAccountNumber.Index).Value = ""
                    End If
                    DTAItemB.Insert(vPrimary, .Cells(i.ColPrimaryPayment.Index).Value, .Cells(i.ColAccountNumber.Index).Value, .Cells(i.ColApprovalCode.Index).Value, _
                    .Cells(i.ColRemarks.Index).Value, .Cells(i.ColPaymentAmount.Index).Value, .Cells(i.ColPaymentBankCharge.Index).Value, _
                    0, 0, 0, -1, -1, -1)
                End With
            Next

            'Update Is payment = 1 (multi_payment)
            DTA.UpdateQuery_IsPayment(1, vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAItemB.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function SimpanItemPaymentSingle(ByVal vPrimary As Double, ByVal i As FrmTrans_POS_Payment) As Boolean
        Dim DTAItemB As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_item_paymentTableAdapter
        Dim ObjCD As New ClsCekData
        Try
            DTAItemB.Connection.ConnectionString = Me.GetConstring

            If String.IsNullOrEmpty(i.TxtApprovalCode.Text) = True Then
                i.TxtApprovalCode.Text = ""
            End If
            If String.IsNullOrEmpty(i.TxtAccountNumber.Text) = True Then
                i.TxtAccountNumber.Text = ""
            End If
            DTAItemB.Insert(vPrimary, i.TxtPrimaryPayment.Text, i.TxtAccountNumber.Text, i.TxtApprovalCode.Text, _
            i.TxtRemarks.Text, i.TxtTotalPayment.Text, i.TxtBankCharge.Text, _
            0, 0, 0, -1, -1, -1)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemB.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function InsertDataPOSItem(ByVal vPrimary As Double, ByVal vPrimaryProd As Double, _
    ByVal vQty As Decimal, ByVal vUnitPrice As Decimal, ByVal vDiscString As String, ByVal vDisc_Amount As Decimal, _
    ByVal vNettPrice As Decimal, ByVal vSubTotal As Decimal, ByVal vPrimaryDisc1 As Double, ByVal vDiscString1 As String, _
    ByVal vDiscAmount1 As Decimal, ByVal vPrimaryDisc2 As Double, ByVal vDiscString2 As String, _
    ByVal vDiscAmount2 As Decimal, ByVal vPrimaryDisc3 As Double, ByVal vPrimaryDiscString3 As Double, _
    ByVal vDiscAmount3 As Decimal, ByVal vPrimaryLocation As Double, ByVal vBatchNumber As String, _
    ByVal vExpiryDate As Date, ByVal vPrimaryUnit As Double) As Boolean
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            DTA.Insert(vPrimary, vPrimaryProd, vQty, vUnitPrice, vDiscString, vDisc_Amount, _
            vNettPrice, vSubTotal, vPrimaryDisc1, vDiscString1, vDiscAmount1, vPrimaryDisc2, vDiscString2, _
            vDiscAmount2, vPrimaryDisc3, vPrimaryDiscString3, vDiscAmount3, vPrimaryLocation, _
            vBatchNumber, vExpiryDate, vPrimaryUnit)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPrimaryTransaksi(ByVal vPrimaryTransaksi As Double) As Double
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryTransaksi(vPrimaryTransaksi).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryTransaksi(vPrimaryTransaksi).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Public Sub UpdateDataButtonInput(ByVal vPrimaryAutoNumber As Double, ByVal qty As Decimal, _
    'ByVal discString As String, ByVal DiscAmount As Decimal, _
    'ByVal NettPrice As Decimal, ByVal subTotal As Decimal, _
    'ByVal DiscAmount1 As Decimal, ByVal DiscAmount2 As Decimal, _
    'ByVal DiscAmount3 As Decimal)
    'Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_itemTableAdapter
    'Try
    'DTA.Connection.ConnectionString = Me.GetConstring
    'DTA.UpdateQuery_UntukInputDataItem(qty, discString, DiscAmount, NettPrice, subTotal, _
    'DiscAmount1, DiscAmount2, DiscAmount3, vPrimaryAutoNumber)
    'Catch ex As Exception
    'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    'Finally
    'DTA.Dispose()
    'End Try
    'End Sub

    Public Sub UpdateDataButtonInputPlusDiskon(ByVal vPrimaryAutoNumber As Double, ByVal qty As Decimal, _
    ByVal discString As String, ByVal DiscAmount As Decimal, _
    ByVal NettPrice As Decimal, ByVal subTotal As Decimal, _
    ByVal DiscAmount1 As Decimal, ByVal DiscAmount2 As Decimal, ByVal DiscAmount3 As Decimal, _
    ByVal vPrimaryDisc1 As Double, ByVal vPrimaryDisc2 As Double, ByVal vPrimaryDisc3 As Double, _
    ByVal vDiscString1 As String, ByVal vDiscString2 As String, ByVal vDiscString3 As String)
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.UpdateQuery_UntukButtonInputPlusDiskon(qty, discString, DiscAmount, NettPrice, subTotal, _
            DiscAmount1, DiscAmount2, DiscAmount3, vPrimaryDisc1, vPrimaryDisc2, vPrimaryDisc3, _
            vDiscString1, vDiscString2, vDiscString3, vPrimaryAutoNumber)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub


End Class
