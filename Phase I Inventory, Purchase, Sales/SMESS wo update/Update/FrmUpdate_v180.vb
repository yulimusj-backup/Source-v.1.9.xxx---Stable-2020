Option Explicit On
Option Strict Off

Partial Public Class FrmUpdate

    Private Sub btn_v180_multiunit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_v180_multiunit.Click
        Try
            If MsgBox("Start multi unit update?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BgwMultiunit.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwMultiunit_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwMultiunit.DoWork
        Try
            Me.Update_v180_Multiunit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwMultiunit_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwMultiunit.RunWorkerCompleted
        Try
            Me.lbl_v180_multiunit.Text = "Data Updated"
            MsgBox("Update completed.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_v180_Multiunit()
        Try
            ''tbltrans_po_item
            ''0. qty multi
            ''1. primary unit multi
            ''2. pob multi
            ''3. git multi
            Dim DtaPoItem As New DseUpdate_v180TableAdapters.tbltrans_po_itemTableAdapter
            Dim TblPoItem As New DseUpdate_v180.tbltrans_po_itemDataTable
            Dim RowPoItem As DseUpdate_v180.tbltrans_po_itemRow
            'DtaPoItem.Connection.ConnectionString  =
            DtaPoItem.Fill(TblPoItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblPoItem.Rows.Count
            For Each RowPoItem In TblPoItem
                With RowPoItem
                    Me.prg_v180_multiunit.Value += 1
                    Me.lbl_v180_multiunit.Text = "Updating Data PO Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString
                    DtaPoItem.UpdateQuery(.Qty, .Primary_Units, .POB, .GIT, .Primary)
                End With
            Next

            ''tbltrans_purchasereturn_item
            ''0. total qty multi
            ''1. primary unit multi
            Dim DtaPrItem As New DseUpdate_v180TableAdapters.tbltrans_purchasereturn_itemTableAdapter
            Dim TblPrItem As New DseUpdate_v180.tbltrans_purchasereturn_itemDataTable
            Dim RowPrItem As DseUpdate_v180.tbltrans_purchasereturn_itemRow
            'DtaPrItem.Connection.ConnectionString  =
            DtaPrItem.Fill(TblPrItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblPrItem.Rows.Count
            For Each RowPrItem In TblPrItem
                With RowPrItem
                    Me.prg_v180_multiunit.Value += 1
                    Me.lbl_v180_multiunit.Text = "Updating Data PR Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString
                    DtaPrItem.UpdateQuery(.Total_Qty, .Primary_Units, .Primary)
                End With
            Next

            ''tbltrans_poc_item
            ''0. total qty multi
            ''1. primary unit multi
            Dim DtaPocItem As New DseUpdate_v180TableAdapters.tbltrans_poc_itemTableAdapter
            Dim TblPocItem As New DseUpdate_v180.tbltrans_poc_itemDataTable
            Dim RowPocItem As DseUpdate_v180.tbltrans_poc_itemRow
            'DtaPocItem.Connection.ConnectionString  =
            DtaPocItem.Fill(TblPocItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblPocItem.Rows.Count
            For Each RowPocItem In TblPocItem
                With RowPocItem
                    Me.prg_v180_multiunit.Value += 1
                    Me.lbl_v180_multiunit.Text = "Updating Data POC Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString
                    DtaPocItem.UpdateQuery(.Total_Qty, .Primary_Units, .Primary)
                End With
            Next

            ''tbltrans_packinglist_item
            ''0. total qty multi
            ''1. primary unit multi
            Dim DtaPlItem As New DseUpdate_v180TableAdapters.tbltrans_packinglist_itemTableAdapter
            Dim TblPlItem As New DseUpdate_v180.tbltrans_packinglist_itemDataTable
            Dim RowPlItem As DseUpdate_v180.tbltrans_packinglist_itemRow
            'DtaPlItem.Connection.ConnectionString  =
            DtaPlItem.Fill(TblPlItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblPlItem.Rows.Count
            For Each RowPlItem In TblPlItem
                With RowPlItem
                    Me.prg_v180_multiunit.Value += 1
                    Me.lbl_v180_multiunit.Text = "Updating Data Packing List Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString
                    DtaPlItem.UpdateQuery(.Total_Qty, .Primary_Units, .Primary)
                End With
            Next

            ''tbltrans_grn_item
            ''0. total qty multi
            ''1. primary unit multi
            Dim DtaGrnItem As New DseUpdate_v180TableAdapters.tbltrans_grn_itemTableAdapter
            Dim TblGrnItem As New DseUpdate_v180.tbltrans_grn_itemDataTable
            Dim RowGrnItem As DseUpdate_v180.tbltrans_grn_itemRow
            'DtaGrnItem.Connection.ConnectionString  =
            DtaGrnItem.Fill(TblGrnItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblGrnItem.Rows.Count
            For Each RowGrnItem In TblGrnItem
                With RowGrnItem
                    Me.prg_v180_multiunit.Value += 1
                    Me.lbl_v180_multiunit.Text = "Updating Data GRN Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString
                    DtaGrnItem.UpdateQuery(.Total_Qty, .Primary_Units, .Primary)
                End With
            Next

            ''tbltrans_po_itemdelivery
            ''0. delivery qty multi
            Dim DtaDqItem As New DseUpdate_v180TableAdapters.tbltrans_po_itemdeliveryTableAdapter
            Dim TblDqItem As New DseUpdate_v180.tbltrans_po_itemdeliveryDataTable
            Dim RowDqItem As DseUpdate_v180.tbltrans_po_itemdeliveryRow
            'DtaDqItem.Connection.ConnectionString  =
            DtaDqItem.Fill(TblDqItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblDqItem.Rows.Count
            For Each RowDqItem In TblDqItem
                With RowDqItem
                    Me.prg_v180_multiunit.Value += 1
                    Me.lbl_v180_multiunit.Text = "Updating Data PO Delivery Date Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString
                    DtaDqItem.UpdateQuery(.delivery_qty, .primary)
                End With
            Next

            ''tbltransaksi_stockinout_item
            ''0. primary unit multi
            ''1. qty multi
            Dim DtaSiOItem As New DseUpdate_v180TableAdapters.tbltransaksi_stockinout_itemTableAdapter
            Dim TblSiOItem As New DseUpdate_v180.tbltransaksi_stockinout_itemDataTable
            Dim RowSiOItem As DseUpdate_v180.tbltransaksi_stockinout_itemRow
            'DtaSiOItem.Connection.ConnectionString  =
            DtaSiOItem.Fill(TblSiOItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblSiOItem.Rows.Count
            For Each RowSiOItem In TblSiOItem
                With RowSiOItem
                    Me.prg_v180_multiunit.Value += 1
                    Me.lbl_v180_multiunit.Text = "Updating Data Stock in Out Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString
                    DtaSiOItem.UpdateQuery(.Qty, .Primary_Units, .Primary)
                End With
            Next

            ''tbltrans_aop_item
            ''0. primary unit multi
            ''1. qty multi
            Dim DtaAopItem As New DseUpdate_v180TableAdapters.tbltrans_aop_itemTableAdapter
            Dim TblAopItem As New DseUpdate_v180.tbltrans_aop_itemDataTable
            Dim RowAopItem As DseUpdate_v180.tbltrans_aop_itemRow
            'DtaAopItem.Connection.ConnectionString  =
            DtaAopItem.Fill(TblAopItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblAopItem.Rows.Count
            For Each RowAopItem In TblAopItem
                With RowAopItem
                    Me.prg_v180_multiunit.Value += 1
                    Me.lbl_v180_multiunit.Text = "Updating Data AOP Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString
                    DtaAopItem.UpdateQuery(.Total_Qty, .Primary_Units, .Primary)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
