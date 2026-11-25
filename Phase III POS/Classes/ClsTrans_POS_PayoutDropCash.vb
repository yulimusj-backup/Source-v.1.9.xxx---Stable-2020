Option Explicit On
Option Strict Off
Imports Core_Module

Public Class ClsTrans_POS_PayoutDropCash

    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307"
        End Try
    End Function

    Public Function GetJmlByPrimaryPayoutDropcash(ByVal vPrimaryPD As Double) As Decimal
        Dim DTA As New DseTrans_POS_PayoutDropCashTableAdapters.tbltrans_session_cabang_item_payout_dropcashTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryPayoutDropCash(vPrimaryPD).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryPayoutDropCash(vPrimaryPD).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal vPrimarySession As Double, ByVal vDGVItem As DataGridView, _
       ByVal i As FrmTrans_POS_PayoutDropCash_Input, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseTrans_POS_PayoutDropCashTableAdapters.tbltrans_session_cabang_item_payout_dropcashTableAdapter
        Dim TBLItemP As New DseTrans_POS_PayoutDropCash.tbltrans_session_cabang_item_payout_dropcashDataTable
        Dim rowItemP As DseTrans_POS_PayoutDropCash.tbltrans_session_cabang_item_payout_dropcashRow
        Dim ObjPay As New ClsMast_PayoutDropCash
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
                    .Cells(i.ColNo.Index).Value = ""
                    If rowItemP.tipe_pay_out_drop_cash = ClsMast_PayoutDropCash.EnumTipe.Drop_Cash Then
                        .Cells(i.ColTipe.Index).Value = "Drop Cash"
                    Else
                        .Cells(i.ColTipe.Index).Value = "Payout"
                    End If
                    .Cells(i.ColCode.Index).Value = ObjPay.GetCode(rowItemP.primary_pay_out_drop_cash)
                    .Cells(i.ColDesc.Index).Value = ObjPay.GetDesc(rowItemP.primary_pay_out_drop_cash)
                    .Cells(i.ColAmount.Index).Value = vObjCD.DigitGrouping(rowItemP.amount, i.Parent_Form.DefPrec)
                    .Cells(i.ColRemarks.Index).Value = rowItemP.remarks
                    .Cells(i.ColPrimaryPayout.Index).Value = rowItemP.primary_pay_out_drop_cash
                    .Cells(i.ColPrimarySession.Index).Value = rowItemP.primary_transaksi
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
            ObjPay = Nothing
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimarySession As Double, ByVal vDGVItemProduct As DataGridView, _
    ByVal i As FrmTrans_POS_PayoutDropCash_Input) As Boolean
        Dim DTAItemP As New DseTrans_POS_PayoutDropCashTableAdapters.tbltrans_session_cabang_item_payout_dropcashTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Dim vTipe As ClsMast_PayoutDropCash.EnumTipe
        Try
            DTAItemP.Connection.ConnectionString = Me.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'Hapus Data lama by primary Session
            DTAItemP.Delete(vPrimarySession)

            'masukan data baru 
            For Each rowDGV In vDGVItemProduct.Rows
                With rowDGV
                    If .Cells(i.ColTipe.Index).Value = "Drop Cash" Then
                        vTipe = ClsMast_PayoutDropCash.EnumTipe.Drop_Cash
                    Else
                        vTipe = ClsMast_PayoutDropCash.EnumTipe.Payout
                    End If
                    DTAItemP.Insert(vPrimarySession, vTipe, .Cells(i.ColPrimaryPayout.Index).Value, _
                    .Cells(i.ColAmount.Index).Value, .Cells(i.ColRemarks.Index).Value)
                End With
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemP.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function HapusItemPerBaris(ByVal vPrimaryAutoNumber As Double) As Boolean
        Dim DTA As New DseTrans_POS_PayoutDropCashTableAdapters.tbltrans_session_cabang_item_payout_dropcashTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.DeleteQuery_PrimaryAutoNumber(vPrimaryAutoNumber)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
