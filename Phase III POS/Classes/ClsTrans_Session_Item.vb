Imports Core_Module
Public Class ClsTrans_Session_Item

    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307"
        End Try
    End Function

    Public Function IsiDGVPendingPOS(ByVal vPrimarySession As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTrans_Session, ByVal vObjCD As ClsCekData) As Boolean
        Dim Obj As New ClsTrans_Session
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseTrans_SessionTableAdapters.tbltrans_pos_cabangTableAdapter
        Dim TBLItemP As New DseTrans_Session.tbltrans_pos_cabangDataTable
        Dim rowItemP As DseTrans_Session.tbltrans_pos_cabangRow
        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'i.Txt_TransDetail_TotalTransQty.Text = vObjCD.DigitGrouping(Obj.GetTotalTransaksiPOS(vPrimarySession), 0, ClsCekData.EnumTipeRounding.Middlepoint)
            'i.TxtSalesItemQty.Text = vObjCD.DigitGrouping(Obj.GetTotalQtyPOS(vPrimarySession), 0, ClsCekData.EnumTipeRounding.Middlepoint)

            'isikan data item
            DTAItemP.FillBy_PrimarySessionNoPaid(TBLItemP, vPrimarySession)
            vDGVItem.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.Col_PD_No.Index).Value = ""
                    .Cells(i.Col_PD_TransNo.Index).Value = rowItemP.no_trans
                    .Cells(i.Col_PD_SalesAmount.Index).Value = vObjCD.DigitGrouping(rowItemP.total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PD_Discount.Index).Value = rowItemP.discount_string
                    .Cells(i.Col_PD_DiscountAmount.Index).Value = vObjCD.DigitGrouping(rowItemP.discount, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PD_SubTotal.Index).Value = vObjCD.DigitGrouping(rowItemP.sub_total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PD_Tax.Index).Value = vObjCD.DigitGrouping(rowItemP.tax_percentage, 2)
                    .Cells(i.Col_PD_TaxAmount.Index).Value = vObjCD.DigitGrouping(rowItemP.tax_amount, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PD_Rounding.Index).Value = vObjCD.DigitGrouping(rowItemP.rounding, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PD_BankCharge.Index).Value = vObjCD.DigitGrouping(rowItemP.bank_charge, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PD_GrandTotal.Index).Value = vObjCD.DigitGrouping(rowItemP.grand_total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PD_TotalPayment.Index).Value = vObjCD.DigitGrouping(rowItemP.total_payment, i.DefPrec, i.RoundingTipe)
                    'sementara status transfer = false 
                    .Cells(i.Col_PD_Transfered.Index).Value = False
                    .Cells(i.Col_PD_PrimaryPOS.Index).Value = rowItemP.primary
                    .Cells(i.Col_PD_PrimarySession.Index).Value = rowItemP.primary_session
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
            Obj = Nothing
        End Try
    End Function

    Public Function IsiDGVPendingPOSDataTransfer(ByVal vPrimarySession As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTrans_Data_Transfer, ByVal vObjCD As ClsCekData) As Boolean
        Dim Obj As New ClsTrans_Session
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseTrans_SessionTableAdapters.tbltrans_pos_cabangTableAdapter
        Dim TBLItemP As New DseTrans_Session.tbltrans_pos_cabangDataTable
        Dim rowItemP As DseTrans_Session.tbltrans_pos_cabangRow
        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'i.Txt_TransDetail_TotalTransQty.Text = vObjCD.DigitGrouping(Obj.GetTotalTransaksiPOS(vPrimarySession), 0, ClsCekData.EnumTipeRounding.Middlepoint)
            'i.TxtSalesItemQty.Text = vObjCD.DigitGrouping(Obj.GetTotalQtyPOS(vPrimarySession), 0, ClsCekData.EnumTipeRounding.Middlepoint)

            'isikan data item
            DTAItemP.FillBy_PrimarySessionNoPaid(TBLItemP, vPrimarySession)
            vDGVItem.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.Col_PD_No.Index).Value = ""
                    .Cells(i.Col_PD_TransNo.Index).Value = rowItemP.no_trans
                    .Cells(i.Col_PD_SalesAmount.Index).Value = vObjCD.DigitGrouping(rowItemP.total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PD_Discount.Index).Value = rowItemP.discount_string
                    .Cells(i.Col_PD_DiscountAmount.Index).Value = vObjCD.DigitGrouping(rowItemP.discount, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PD_SubTotal.Index).Value = vObjCD.DigitGrouping(rowItemP.sub_total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PD_Tax.Index).Value = vObjCD.DigitGrouping(rowItemP.tax_percentage, 2)
                    .Cells(i.Col_PD_TaxAmount.Index).Value = vObjCD.DigitGrouping(rowItemP.tax_amount, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PD_Rounding.Index).Value = vObjCD.DigitGrouping(rowItemP.rounding, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PD_BankCharge.Index).Value = vObjCD.DigitGrouping(rowItemP.bank_charge, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PD_GrandTotal.Index).Value = vObjCD.DigitGrouping(rowItemP.grand_total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PD_TotalPayment.Index).Value = vObjCD.DigitGrouping(rowItemP.total_payment, i.DefPrec, i.RoundingTipe)
                    'sementara status transfer = false 
                    .Cells(i.Col_PD_Transfered.Index).Value = False
                    .Cells(i.Col_PD_PrimaryPOS.Index).Value = rowItemP.primary
                    .Cells(i.Col_PD_PrimarySession.Index).Value = rowItemP.primary_session
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
            Obj = Nothing
        End Try
    End Function

    Public Function IsiDGVPayout(ByVal vPrimarySession As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTrans_Session, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseTrans_SessionTableAdapters.DataTable1TableAdapter
        Dim TBLItemP As New DseTrans_Session.DataTable1DataTable
        Dim rowItemP As DseTrans_Session.DataTable1Row
        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item
            DTAItemP.Fill(TBLItemP, vPrimarySession)
            vDGVItem.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                If rowItemP.tipe_pay_out_drop_cash = ClsMast_PayoutDropCash.EnumTipe.Drop_Cash Then
                    rowItemP.Tipe_Payout_Desc = "Drop cash"
                Else
                    rowItemP.Tipe_Payout_Desc = "Payout"
                End If
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.Col_PayOutD_No.Index).Value = ""
                    .Cells(i.Col_PayOutD_Tipe.Index).Value = rowItemP.Tipe_Payout_Desc
                    .Cells(i.Col_PayoutD_Code.Index).Value = rowItemP.code
                    .Cells(i.Col_PayoutD_Account.Index).Value = rowItemP.Account_Code
                    .Cells(i.Col_PayoutD_Desc.Index).Value = rowItemP.Description
                    .Cells(i.Col_PayOutD_Amount.Index).Value = vObjCD.DigitGrouping(rowItemP.amount, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PayOutD_PrimarySession.Index).Value = rowItemP.primary_transaksi
                    .Cells(i.Col_PayOutD_PrimaryPayOutDropCash.Index).Value = rowItemP.primary_pay_out_drop_cash
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

    Public Function IsiDGVSucceeded(ByVal vPrimarySession As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTrans_Session, ByVal vObjCD As ClsCekData) As Boolean
        Dim Obj As New ClsTrans_Session
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseTrans_SessionTableAdapters.tbltrans_pos_cabangTableAdapter
        Dim TBLItemP As New DseTrans_Session.tbltrans_pos_cabangDataTable
        Dim rowItemP As DseTrans_Session.tbltrans_pos_cabangRow
        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            i.Txt_TransDetail_TotalTransQty.Text = vObjCD.DigitGrouping(Obj.GetTotalTransaksiPOS(vPrimarySession), 0, ClsCekData.EnumTipeRounding.Middlepoint)
            i.TxtSalesItemQty.Text = vObjCD.DigitGrouping(Obj.GetTotalQtyPOS(vPrimarySession), 0, ClsCekData.EnumTipeRounding.Middlepoint)

            'isikan data item
            DTAItemP.FillBy_PrimarySessionPaid(TBLItemP, vPrimarySession)
            vDGVItem.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.Col_TransD_No.Index).Value = ""
                    .Cells(i.Col_TransD_NoTransaksi.Index).Value = rowItemP.no_trans
                    .Cells(i.Col_TransD_SalesAmount.Index).Value = vObjCD.DigitGrouping(rowItemP.total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_TransD_Discount.Index).Value = rowItemP.discount_string
                    .Cells(i.Col_TransD_DiscountAmount.Index).Value = vObjCD.DigitGrouping(rowItemP.discount, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_TransD_SubTotal.Index).Value = vObjCD.DigitGrouping(rowItemP.sub_total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_TransD_Tax.Index).Value = vObjCD.DigitGrouping(rowItemP.tax_percentage, 2)
                    .Cells(i.Col_TransD_TaxAmount.Index).Value = vObjCD.DigitGrouping(rowItemP.tax_amount, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_TransD_Rounding.Index).Value = vObjCD.DigitGrouping(rowItemP.rounding, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_TransD_BankCharge.Index).Value = vObjCD.DigitGrouping(rowItemP.bank_charge, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_TransD_GrandTotal.Index).Value = vObjCD.DigitGrouping(rowItemP.grand_total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_TransD_Payment.Index).Value = vObjCD.DigitGrouping(rowItemP.total_payment, i.DefPrec, i.RoundingTipe)
                    'sementara status transfer = false 
                    .Cells(i.Col_TransD_StatusTransfer.Index).Value = False
                    .Cells(i.Col_TransD_PrimaryPOS.Index).Value = rowItemP.primary
                    .Cells(i.Col_TransD_PrimarySession.Index).Value = rowItemP.primary_session
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
            Obj = Nothing
        End Try
    End Function

    Public Function IsiDGVCancelledDetails(ByVal vPrimarySession As Double, ByVal vDGVItem As DataGridView, _
      ByVal i As FrmTrans_Session, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseTrans_SessionTableAdapters.tbltrans_session_cabang_itemvoidTableAdapter
        Dim TBLItemP As New DseTrans_Session.tbltrans_session_cabang_itemvoidDataTable
        Dim rowItemP As DseTrans_Session.tbltrans_session_cabang_itemvoidRow
        Dim ObjPOS As New ClsTrans_POS
        Dim ObjUser As New ClsUser

        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item
            DTAItemP.FillBy_PrimarySession(TBLItemP, vPrimarySession)
            vDGVItem.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.Col_CancelledItem_No.Index).Value = ""
                    .Cells(i.Col_CancelledItem_TransNo.Index).Value = ObjPOS.GetNoPOS(rowItemP.primary_transaksi_pos)
                    .Cells(i.Col_CancelledItem_UnitPrice.Index).Value = vObjCD.DigitGrouping(CDec(rowItemP.unit_price), i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledItem_Discount.Index).Value = rowItemP.disc_string
                    .Cells(i.Col_CancelledItem_Disc_Amount.Index).Value = vObjCD.DigitGrouping(rowItemP.disc, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledItem_NettPrice.Index).Value = vObjCD.DigitGrouping(rowItemP.nett_price, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledItem_Qty.Index).Value = vObjCD.DigitGrouping(rowItemP.qty, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledItem_SubTotal.Index).Value = vObjCD.DigitGrouping(rowItemP.sub_total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledItem_CancellationReason.Index).Value = rowItemP.Reason
                    .Cells(i.Col_CancelledItem_Authorization.Index).Value = ObjUser.GetKodeUser(rowItemP.Primary_User)
                    .Cells(i.Col_CancelledItem_PrimaryUser_Autho.Index).Value = rowItemP.Primary_User
                    .Cells(i.Col_CancelledItem_PrimarySession.Index).Value = rowItemP.primary_session
                    .Cells(i.Col_CancelledItem_PrimaryTransksiPOS.Index).Value = rowItemP.primary_transaksi_pos
                    .Cells(i.Col_CancelledItem_DescBarang.Index).Value = rowItemP.Description
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
            ObjPOS = Nothing
            ObjUser = Nothing
        End Try
    End Function

    Public Function IsiDGVCancelled(ByVal vPrimarySession As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTrans_Session, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseTrans_SessionTableAdapters.DataTable2TableAdapter
        Dim TBLItemP As New DseTrans_Session.DataTable2DataTable
        Dim rowItemP As DseTrans_Session.DataTable2Row
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabang_itemvoidTableAdapter
        Dim TBL As New DseTrans_Session.tbltrans_session_cabang_itemvoidDataTable
        Dim Row As DseTrans_Session.tbltrans_session_cabang_itemvoidRow
        Dim ObjPOS As New ClsTrans_POS
        Dim ObjUser As New ClsUser
        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item 
            DTAItemP.Fill(TBLItemP, vPrimarySession)
            vDGVItem.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                DTA.FillBy_PrimaryPOSNSession(TBL, vPrimarySession, rowItemP.primary)
                If TBL.Rows.Count > 0 Then
                    For Each Row In TBL.Rows
                        If rowItemP.Reason = "" Then
                            rowItemP.Reason = Row.Reason
                        Else
                            rowItemP.Reason += ", " + Row.Reason.ToString
                        End If
                    Next
                End If
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.Col_CancelledTrans_No.Index).Value = ""
                    .Cells(i.Col_CancelledTrans_TransNumber.Index).Value = rowItemP.no_trans
                    .Cells(i.Col_CancelledTrans_SalesAmount.Index).Value = vObjCD.DigitGrouping(CDec(rowItemP.total), i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledTrans_Discount.Index).Value = vObjCD.DigitGrouping(rowItemP.discount, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledTrans_NettSales.Index).Value = vObjCD.DigitGrouping(rowItemP.sub_total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledTrans_Reason.Index).Value = rowItemP.Reason
                    .Cells(i.Col_CancelledTrans_Autho.Index).Value = rowItemP.Kode_User
                    .Cells(i.Col_CancelledTrans_PrimaryPOS.Index).Value = rowItemP.primary
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
            DTA.Dispose()
            TBL.Dispose()
            ObjPOS = Nothing
            ObjUser = Nothing
        End Try
    End Function

    Public Function IsiDGVPaymentDetails(ByVal vPrimaryPayment As Double, _
    ByVal i As FrmTrans_Session, ByVal vObjCD As ClsCekData, _
    ByVal TBLItem As DseTrans_Session.DGVPaymentDetailsDataTable) As Boolean
        Dim Filter As String = ""
        Dim rowDGV As DataGridViewRow
        Dim RowView As DataRowView
        Dim RowItem As DseTrans_Session.DGVPaymentDetailsRow
        Try
            'autosize dgv
            i.dgv_PaymentD.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'isikan data item
            i.dgv_PaymentD.Rows.Clear()
            Filter = "[Primary Payment] = " + vPrimaryPayment.ToString
            TBLItem.DefaultView.RowFilter = Filter
            For Each RowView In TBLItem.DefaultView
                RowItem = RowView.Row
                i.dgv_PaymentD.Rows.Add()
                rowDGV = i.dgv_PaymentD.Rows(i.dgv_PaymentD.Rows.Count - 1)
                With rowDGV
                    .Cells(i.Col_PaymentD_No.Index).Value = ""
                    .Cells(i.Col_PaymentD_TransNo.Index).Value = RowItem.No_POS
                    .Cells(i.Col_PaymentD_PaymentMethod.Index).Value = RowItem.Payment_Code
                    .Cells(i.Col_PaymentD_AccountCode.Index).Value = RowItem.Account_Number
                    .Cells(i.Col_PaymentD_Approval_Code.Index).Value = RowItem.Approval_Code
                    .Cells(i.Col_PaymentD_Amount.Index).Value = vObjCD.DigitGrouping(RowItem.Payment_Amount, i.DefPrec, i.RoundingTipe)
                End With
            Next

            '-----
            vObjCD.SetNumber(i.dgv_PaymentD)
            i.dgv_PaymentD.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiDGVPaymentHeader(ByVal vPrimarySession As Double, _
    ByVal i As FrmTrans_Session, ByVal vObjCD As ClsCekData, _
    ByVal TBLItem As DseTrans_Session.DGVPaymentDetailsDataTable) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAPay As New DseTrans_SessionTableAdapters.tbltrans_pos_cabang_item_paymentTableAdapter
        Dim TBLPay As New DseTrans_Session.tbltrans_pos_cabang_item_paymentDataTable
        Dim RowPay As DseTrans_Session.tbltrans_pos_cabang_item_paymentRow

        Dim ObjPOS As New ClsTrans_POS
        Dim ObjPayment As New ClsMast_PaymentOption
        Dim RowItem As DseTrans_Session.DGVPaymentDetailsRow
        Dim TBLHeader As New DseTrans_Session.DGVPaymentHeaderDataTable
        Dim RowHeader As DseTrans_Session.DGVPaymentHeaderRow
        Dim Filter As String = ""
        Dim ListPrimaryPayment As New List(Of Double)
        Dim PrimaryPayment As Double
        Dim RowView As DataRowView
        Dim SumTotalPayment As Decimal = 0
        Dim SumTotalTransPOS As Double = 0

        Try
            DTAPay.Connection.ConnectionString = Me.GetConstring
            'autosize dgv
            i.dgv_Payment_Sum.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTAPay.FillBy_PrimarySession(TBLPay, vPrimarySession)
            'Data Detail
            TBLItem.Rows.Clear()
            For Each RowPay In TBLPay.Rows
                If ListPrimaryPayment.Contains(RowPay.primary_payment) = False Then
                    ListPrimaryPayment.Add(RowPay.primary_payment)
                End If
                RowItem = TBLItem.NewDGVPaymentDetailsRow
                With RowItem
                    .Primary_Session = vPrimarySession
                    .Primary_POS = RowPay.primary_transaksi
                    .Primary_Payment = RowPay.primary_payment
                    If RowPay.primary_payment = -1 Then
                        .Payment_Code = "CASH"
                    Else
                        .Payment_Code = ObjPayment.GetCode(RowPay.primary_payment)
                    End If
                    .No_POS = ObjPOS.GetNoPOS(RowPay.primary_transaksi)
                    If String.IsNullOrEmpty(RowPay.account_number) = False Then
                        .Account_Number = RowPay.account_number
                    Else
                        .Account_Number = ""
                    End If
                    If String.IsNullOrEmpty(RowPay.approval_code) = False Then
                        .Approval_Code = RowPay.approval_code
                    Else
                        .Approval_Code = ""
                    End If
                    .Payment_Amount = RowPay.amount
                End With
                TBLItem.Rows.Add(RowItem)
            Next

            'Data Header
            For Each PrimaryPayment In ListPrimaryPayment
                SumTotalPayment = 0
                SumTotalTransPOS = 0
                Filter = "[Primary Payment] = " + PrimaryPayment.ToString + ""
                TBLItem.DefaultView.RowFilter = Filter
                For Each RowView In TBLItem.DefaultView
                    RowItem = RowView.Row
                    SumTotalPayment += RowItem.Payment_Amount
                Next
                SumTotalTransPOS = TBLItem.DefaultView.Count
                RowHeader = TBLHeader.NewDGVPaymentHeaderRow
                With RowHeader
                    .Primary_Session = vPrimarySession
                    .Primary_Payment = PrimaryPayment
                    If PrimaryPayment = -1 Then
                        .Payment_Code = "CASH"
                    Else
                        .Payment_Code = ObjPayment.GetCode(PrimaryPayment)
                    End If
                    .Account_Code = ObjPayment.GetBankAccountNumber(PrimaryPayment)
                    .Payment_Amount = SumTotalPayment
                    .Transaksi_Qty = SumTotalTransPOS
                End With
                TBLHeader.Rows.Add(RowHeader)
            Next

            'Masukan ke DGV
            i.dgv_Payment_Sum.Rows.Clear()
            For Each RowHeader In TBLHeader.Rows
                i.dgv_Payment_Sum.Rows.Add()
                rowDGV = i.dgv_Payment_Sum.Rows(i.dgv_Payment_Sum.Rows.Count - 1)
                With rowDGV
                    .Cells(i.Col_PaymentS_No.Index).Value = ""
                    .Cells(i.Col_PaymentS_PaymentMethod.Index).Value = RowHeader.Payment_Code
                    .Cells(i.Col_PaymentS_AccountCode.Index).Value = RowHeader.Account_Code
                    .Cells(i.Col_PaymentS_Amount.Index).Value = vObjCD.DigitGrouping(RowHeader.Payment_Amount, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PaymentS_TransactionQty.Index).Value = RowHeader.Transaksi_Qty
                    .Cells(i.Col_PaymentS_PrimaryPayment.Index).Value = RowHeader.Primary_Payment
                End With
            Next
            For Each PrimaryPayment In ListPrimaryPayment
                Me.IsiDGVPaymentDetails(PrimaryPayment, i, vObjCD, TBLItem)
                Exit For
            Next

            vObjCD.SetNumber(i.dgv_Payment_Sum)
            i.dgv_Payment_Sum.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPay.Dispose()
            TBLPay.Dispose()
            TBLHeader.Dispose()
            ObjPOS = Nothing
            ObjPayment = Nothing
            ListPrimaryPayment = Nothing
        End Try
    End Function

    'Isi DGV Untuk Form Data Transfer
    Public Function IsiDGVPayoutDataTransfer(ByVal vPrimarySession As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTrans_Data_Transfer, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseTrans_SessionTableAdapters.DataTable1TableAdapter
        Dim TBLItemP As New DseTrans_Session.DataTable1DataTable
        Dim rowItemP As DseTrans_Session.DataTable1Row
        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item
            DTAItemP.Fill(TBLItemP, vPrimarySession)
            vDGVItem.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                If rowItemP.tipe_pay_out_drop_cash = ClsMast_PayoutDropCash.EnumTipe.Drop_Cash Then
                    rowItemP.Tipe_Payout_Desc = "Drop cash"
                Else
                    rowItemP.Tipe_Payout_Desc = "Payout"
                End If
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.Col_PayOutD_No.Index).Value = ""
                    .Cells(i.Col_PayOutD_Tipe.Index).Value = rowItemP.Tipe_Payout_Desc
                    .Cells(i.Col_PayoutD_Code.Index).Value = rowItemP.code
                    .Cells(i.Col_PayoutD_Account.Index).Value = rowItemP.Account_Code
                    .Cells(i.Col_PayoutD_Desc.Index).Value = rowItemP.Description
                    .Cells(i.Col_PayOutD_Amount.Index).Value = vObjCD.DigitGrouping(rowItemP.amount, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PayOutD_PrimarySession.Index).Value = rowItemP.primary_transaksi
                    .Cells(i.Col_PayOutD_PrimaryPayOutDropCash.Index).Value = rowItemP.primary_pay_out_drop_cash
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

    Public Function IsiDGVSucceededDataTransfer(ByVal vPrimarySession As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTrans_Data_Transfer, ByVal vObjCD As ClsCekData) As Boolean
        Dim Obj As New ClsTrans_Session
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseTrans_SessionTableAdapters.tbltrans_pos_cabangTableAdapter
        Dim TBLItemP As New DseTrans_Session.tbltrans_pos_cabangDataTable
        Dim rowItemP As DseTrans_Session.tbltrans_pos_cabangRow
        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            i.Txt_TransDetail_TotalTransQty.Text = vObjCD.DigitGrouping(Obj.GetTotalTransaksiPOS(vPrimarySession), 0, ClsCekData.EnumTipeRounding.Middlepoint)
            i.TxtSalesItemQty.Text = vObjCD.DigitGrouping(Obj.GetTotalQtyPOS(vPrimarySession), 0, ClsCekData.EnumTipeRounding.Middlepoint)

            'isikan data item
            DTAItemP.FillBy_PrimarySessionPaid(TBLItemP, vPrimarySession)
            vDGVItem.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.Col_TransD_No.Index).Value = ""
                    .Cells(i.Col_TransD_NoTransaksi.Index).Value = rowItemP.no_trans
                    .Cells(i.Col_TransD_SalesAmount.Index).Value = vObjCD.DigitGrouping(rowItemP.total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_TransD_Discount.Index).Value = rowItemP.discount_string
                    .Cells(i.Col_TransD_DiscountAmount.Index).Value = vObjCD.DigitGrouping(rowItemP.discount, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_TransD_SubTotal.Index).Value = vObjCD.DigitGrouping(rowItemP.sub_total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_TransD_Tax.Index).Value = vObjCD.DigitGrouping(rowItemP.tax_percentage, 2)
                    .Cells(i.Col_TransD_TaxAmount.Index).Value = vObjCD.DigitGrouping(rowItemP.tax_amount, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_TransD_Rounding.Index).Value = vObjCD.DigitGrouping(rowItemP.rounding, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_TransD_BankCharge.Index).Value = vObjCD.DigitGrouping(rowItemP.bank_charge, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_TransD_GrandTotal.Index).Value = vObjCD.DigitGrouping(rowItemP.grand_total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_TransD_Payment.Index).Value = vObjCD.DigitGrouping(rowItemP.total_payment, i.DefPrec, i.RoundingTipe)
                    'sementara status transfer = false 
                    .Cells(i.Col_TransD_StatusTransfer.Index).Value = False
                    .Cells(i.Col_TransD_PrimaryPOS.Index).Value = rowItemP.primary
                    .Cells(i.Col_TransD_PrimarySession.Index).Value = rowItemP.primary_session
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
            Obj = Nothing
        End Try
    End Function

    Public Function IsiDGVCancelledDetailsDataTransfer(ByVal vPrimarySession As Double, ByVal vDGVItem As DataGridView, _
      ByVal i As FrmTrans_Data_Transfer, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseTrans_SessionTableAdapters.tbltrans_session_cabang_itemvoidTableAdapter
        Dim TBLItemP As New DseTrans_Session.tbltrans_session_cabang_itemvoidDataTable
        Dim rowItemP As DseTrans_Session.tbltrans_session_cabang_itemvoidRow
        Dim ObjPOS As New ClsTrans_POS
        Dim ObjUser As New ClsUser
        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item
            DTAItemP.FillBy_PrimarySession(TBLItemP, vPrimarySession)
            vDGVItem.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.Col_CancelledItem_No.Index).Value = ""
                    .Cells(i.Col_CancelledItem_TransNo.Index).Value = ObjPOS.GetNoPOS(rowItemP.primary_transaksi_pos)
                    .Cells(i.Col_CancelledItem_UnitPrice.Index).Value = vObjCD.DigitGrouping(CDec(rowItemP.unit_price), i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledItem_Discount.Index).Value = rowItemP.disc_string
                    .Cells(i.Col_CancelledItem_Disc_Amount.Index).Value = vObjCD.DigitGrouping(rowItemP.disc, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledItem_NettPrice.Index).Value = vObjCD.DigitGrouping(rowItemP.nett_price, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledItem_Qty.Index).Value = vObjCD.DigitGrouping(rowItemP.qty, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledItem_SubTotal.Index).Value = vObjCD.DigitGrouping(rowItemP.sub_total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledItem_CancellationReason.Index).Value = rowItemP.Reason
                    .Cells(i.Col_CancelledItem_Authorization.Index).Value = ObjUser.GetKodeUser(rowItemP.Primary_User)
                    .Cells(i.Col_CancelledItem_PrimaryUser_Autho.Index).Value = rowItemP.Primary_User
                    .Cells(i.Col_CancelledItem_PrimarySession.Index).Value = rowItemP.primary_session
                    .Cells(i.Col_CancelledItem_PrimaryTransksiPOS.Index).Value = rowItemP.primary_transaksi_pos
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
            ObjPOS = Nothing
            ObjUser = Nothing
        End Try
    End Function

    Public Function IsiDGVCancelledDataTransfer(ByVal vPrimarySession As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTrans_Data_Transfer, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseTrans_SessionTableAdapters.DataTable2TableAdapter
        Dim TBLItemP As New DseTrans_Session.DataTable2DataTable
        Dim rowItemP As DseTrans_Session.DataTable2Row
        Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabang_itemvoidTableAdapter
        Dim TBL As New DseTrans_Session.tbltrans_session_cabang_itemvoidDataTable
        Dim Row As DseTrans_Session.tbltrans_session_cabang_itemvoidRow
        Dim ObjPOS As New ClsTrans_POS
        Dim ObjUser As New ClsUser
        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item 
            DTAItemP.Fill(TBLItemP, vPrimarySession)
            vDGVItem.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                DTA.FillBy_PrimaryPOSNSession(TBL, vPrimarySession, rowItemP.primary)
                If TBL.Rows.Count > 0 Then
                    For Each Row In TBL.Rows
                        If rowItemP.Reason = "" Then
                            rowItemP.Reason = Row.Reason
                        Else
                            rowItemP.Reason += ", " + Row.Reason.ToString
                        End If
                    Next
                End If
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.Col_CancelledTrans_No.Index).Value = ""
                    .Cells(i.Col_CancelledTrans_TransNumber.Index).Value = rowItemP.no_trans
                    .Cells(i.Col_CancelledTrans_SalesAmount.Index).Value = vObjCD.DigitGrouping(CDec(rowItemP.total), i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledTrans_Discount.Index).Value = vObjCD.DigitGrouping(rowItemP.discount, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledTrans_NettSales.Index).Value = vObjCD.DigitGrouping(rowItemP.sub_total, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_CancelledTrans_Reason.Index).Value = rowItemP.Reason
                    .Cells(i.Col_CancelledTrans_Autho.Index).Value = rowItemP.Kode_User
                    .Cells(i.Col_CancelledTrans_PrimaryPOS.Index).Value = rowItemP.primary
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
            DTA.Dispose()
            TBL.Dispose()
            ObjPOS = Nothing
            ObjUser = Nothing
        End Try
    End Function

    Public Function IsiDGVPaymentDetailsDataTransfer(ByVal vPrimaryPayment As Double, _
    ByVal i As FrmTrans_Data_Transfer, ByVal vObjCD As ClsCekData, _
    ByVal TBLItem As DseTrans_Session.DGVPaymentDetailsDataTable) As Boolean
        Dim Filter As String = ""
        Dim rowDGV As DataGridViewRow
        Dim RowView As DataRowView
        Dim RowItem As DseTrans_Session.DGVPaymentDetailsRow
        Try
            'autosize dgv
            i.dgv_PaymentD.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'isikan data item
            i.dgv_PaymentD.Rows.Clear()
            Filter = "[Primary Payment] = " + vPrimaryPayment.ToString
            TBLItem.DefaultView.RowFilter = Filter
            For Each RowView In TBLItem.DefaultView
                RowItem = RowView.Row
                i.dgv_PaymentD.Rows.Add()
                rowDGV = i.dgv_PaymentD.Rows(i.dgv_PaymentD.Rows.Count - 1)
                With rowDGV
                    .Cells(i.Col_PaymentD_No.Index).Value = ""
                    .Cells(i.Col_PaymentD_TransNo.Index).Value = RowItem.No_POS
                    .Cells(i.Col_PaymentD_PaymentMethod.Index).Value = RowItem.Payment_Code
                    .Cells(i.Col_PaymentD_AccountCode.Index).Value = RowItem.Account_Number
                    .Cells(i.Col_PaymentD_Approval_Code.Index).Value = RowItem.Approval_Code
                    .Cells(i.Col_PaymentD_Amount.Index).Value = vObjCD.DigitGrouping(RowItem.Payment_Amount, i.DefPrec, i.RoundingTipe)
                End With
            Next

            '-----
            vObjCD.SetNumber(i.dgv_PaymentD)
            i.dgv_PaymentD.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiDGVPaymentHeaderDataTransfer(ByVal vPrimarySession As Double, _
    ByVal i As FrmTrans_Data_Transfer, ByVal vObjCD As ClsCekData, _
    ByVal TBLItem As DseTrans_Session.DGVPaymentDetailsDataTable) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAPay As New DseTrans_SessionTableAdapters.tbltrans_pos_cabang_item_paymentTableAdapter
        Dim TBLPay As New DseTrans_Session.tbltrans_pos_cabang_item_paymentDataTable
        Dim RowPay As DseTrans_Session.tbltrans_pos_cabang_item_paymentRow

        Dim ObjPOS As New ClsTrans_POS
        Dim ObjPayment As New ClsMast_PaymentOption
        Dim RowItem As DseTrans_Session.DGVPaymentDetailsRow
        Dim TBLHeader As New DseTrans_Session.DGVPaymentHeaderDataTable
        Dim RowHeader As DseTrans_Session.DGVPaymentHeaderRow
        Dim Filter As String = ""
        Dim ListPrimaryPayment As New List(Of Double)
        Dim PrimaryPayment As Double
        Dim RowView As DataRowView
        Dim SumTotalPayment As Decimal = 0
        Dim SumTotalTransPOS As Double = 0

        Try
            DTAPay.Connection.ConnectionString = Me.GetConstring
            'autosize dgv
            i.dgv_Payment_Sum.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTAPay.FillBy_PrimarySession(TBLPay, vPrimarySession)
            'Data Detail
            TBLItem.Rows.Clear()
            For Each RowPay In TBLPay.Rows
                If ListPrimaryPayment.Contains(RowPay.primary_payment) = False Then
                    ListPrimaryPayment.Add(RowPay.primary_payment)
                End If
                RowItem = TBLItem.NewDGVPaymentDetailsRow
                With RowItem
                    .Primary_Session = vPrimarySession
                    .Primary_POS = RowPay.primary_transaksi
                    .Primary_Payment = RowPay.primary_payment
                    If RowPay.primary_payment = -1 Then
                        .Payment_Code = "CASH"
                    Else
                        .Payment_Code = ObjPayment.GetCode(RowPay.primary_payment)
                    End If
                    .No_POS = ObjPOS.GetNoPOS(RowPay.primary_transaksi)
                    If String.IsNullOrEmpty(RowPay.account_number) = False Then
                        .Account_Number = RowPay.account_number
                    Else
                        .Account_Number = ""
                    End If
                    If String.IsNullOrEmpty(RowPay.approval_code) = False Then
                        .Approval_Code = RowPay.approval_code
                    Else
                        .Approval_Code = ""
                    End If
                    .Payment_Amount = RowPay.amount
                End With
                TBLItem.Rows.Add(RowItem)
            Next

            'Data Header
            For Each PrimaryPayment In ListPrimaryPayment
                SumTotalPayment = 0
                SumTotalTransPOS = 0
                Filter = "[Primary Payment] = " + PrimaryPayment.ToString + ""
                TBLItem.DefaultView.RowFilter = Filter
                For Each RowView In TBLItem.DefaultView
                    RowItem = RowView.Row
                    SumTotalPayment += RowItem.Payment_Amount
                Next
                SumTotalTransPOS = TBLItem.DefaultView.Count
                RowHeader = TBLHeader.NewDGVPaymentHeaderRow
                With RowHeader
                    .Primary_Session = vPrimarySession
                    .Primary_Payment = PrimaryPayment
                    If PrimaryPayment = -1 Then
                        .Payment_Code = "CASH"
                    Else
                        .Payment_Code = ObjPayment.GetCode(PrimaryPayment)
                    End If
                    .Account_Code = ObjPayment.GetBankAccountNumber(PrimaryPayment)
                    .Payment_Amount = SumTotalPayment
                    .Transaksi_Qty = SumTotalTransPOS
                End With
                TBLHeader.Rows.Add(RowHeader)
            Next

            'Masukan ke DGV
            i.dgv_Payment_Sum.Rows.Clear()
            For Each RowHeader In TBLHeader.Rows
                i.dgv_Payment_Sum.Rows.Add()
                rowDGV = i.dgv_Payment_Sum.Rows(i.dgv_Payment_Sum.Rows.Count - 1)
                With rowDGV
                    .Cells(i.Col_PaymentS_No.Index).Value = ""
                    .Cells(i.Col_PaymentS_PaymentMethod.Index).Value = RowHeader.Payment_Code
                    .Cells(i.Col_PaymentS_AccountCode.Index).Value = RowHeader.Account_Code
                    .Cells(i.Col_PaymentS_Amount.Index).Value = vObjCD.DigitGrouping(RowHeader.Payment_Amount, i.DefPrec, i.RoundingTipe)
                    .Cells(i.Col_PaymentS_TransactionQty.Index).Value = RowHeader.Transaksi_Qty
                    .Cells(i.Col_PaymentS_PrimaryPayment.Index).Value = RowHeader.Primary_Payment
                End With
            Next
            For Each PrimaryPayment In ListPrimaryPayment
                Me.IsiDGVPaymentDetailsDataTransfer(PrimaryPayment, i, vObjCD, TBLItem)
                Exit For
            Next

            vObjCD.SetNumber(i.dgv_Payment_Sum)
            i.dgv_Payment_Sum.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPay.Dispose()
            TBLPay.Dispose()
            TBLHeader.Dispose()
            ObjPOS = Nothing
            ObjPayment = Nothing
            ListPrimaryPayment = Nothing
        End Try
    End Function
End Class
