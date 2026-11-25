Option Explicit On
Option Strict Off

Partial Public Class FrmUpdate
    Public Enum EnumTipeLeave As Integer
        Addition
        Deduction
    End Enum

    Public Enum EnumTipeLeave2 As Integer
        MonthlyLeaveBySystem
        LongLeaveBySystem
        Add_Ded_ByUser
    End Enum

    Private Sub btn_v180_master_cuti_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_v180_master_cuti.Click
        If MsgBox("Update identifier item cuti?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
            Control.CheckForIllegalCrossThreadCalls = False
            Me.bgw180_MasterCuti.RunWorkerAsync()
        End If
    End Sub

    Private Sub bgw180_MasterCuti_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw180_MasterCuti.DoWork
        Dim DTAItem As New DseUpdate_v180TableAdapters.tblmaster_cuti_itemTableAdapter
        Dim TBLItem As New DseUpdate_v180.tblmaster_cuti_itemDataTable
        Dim rowItem As DseUpdate_v180.tblmaster_cuti_itemRow
        Dim PrimaryTransaksiDummy As Double = -1
        Dim DTAItem1 As New DseUpdate_v180TableAdapters.tblmaster_cuti_item_descTableAdapter
        Dim counter As Double = 1
        Dim Identifier As String = ""
        Dim Tipe As EnumTipeLeave = EnumTipeLeave.Addition
        Dim Tipe1 As EnumTipeLeave2 = EnumTipeLeave2.Add_Ded_ByUser

        Try
            'isikan data cuti item
            DTAItem.Fill(TBLItem)
            Me.prg_v180_master_cuti.Value = 0
            Me.prg_v180_master_cuti.Maximum = TBLItem.Rows.Count

            'kosongkan data cuti item desc
            DTAItem1.Delete()

            'untuk setiap cuti item
            For Each rowItem In TBLItem.Rows
                Me.prg_v180_master_cuti.Value += 1
                Me.lbl_v180_master_cuti.Text = "Data " + Me.prg_v180_master_cuti.Value.ToString + " of " + Me.prg_v180_master_cuti.Maximum.ToString

                '1. isikan identifier
                If rowItem.primary_leave <> PrimaryTransaksiDummy Then
                    'transaksi baru
                    counter = 1
                    PrimaryTransaksiDummy = rowItem.primary_leave
                Else
                    counter += 1
                End If
                Identifier = rowItem.primary_leave.ToString + "-" + counter.ToString
                DTAItem.UpdateQuery(Identifier, rowItem.primary)
                '2. berikan keterangan jika ada add cuti
                Tipe = EnumTipeLeave.Addition
                If rowItem.leave_addition > 0 Then
                    DTAItem1.Insert(rowItem.primary_leave, Identifier, CInt(Tipe), rowItem.period, "leave addition", rowItem.leave_addition, CInt(Tipe1))
                End If
                '3. berikan keterangan jika ada ded cuti
                Tipe = EnumTipeLeave.Deduction
                If rowItem.leave_deduction > 0 Then
                    DTAItem1.Insert(rowItem.primary_leave, Identifier, CInt(Tipe), rowItem.period, "leave deduction", rowItem.leave_deduction, CInt(Tipe1))
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            DTAItem1.Dispose()
        End Try
    End Sub

    Private Sub bgw180_MasterCuti_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw180_MasterCuti.RunWorkerCompleted
        Me.prg_v180_master_cuti.Value = 0
        Me.lbl_v180_master_cuti.Text = ""
        MsgBox("Data updated.", MsgBoxStyle.Information, "Information")
    End Sub

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

    Private Sub Update_v180_MultiAcc()
        '-------------------------------------------------------------------------------------------------------------------------
        'tbltrans_aop
        'tbltrans_purchasereturn
        'tbltrans_salesinvoice
        'tbltrans_Salesreturn
        'tbltrans_historicalpurchasebalancemanual
        'tbltrans_historicalpurchasereturnbalancemanual
        'tbltrans_historicalsales
        'tbltrans_historicalsalesReturn
        'tbltrans_jurnal
        '-------------------------------------------------------------------------------------------------------------------------
        Dim ObjAkun As New Core_Module.ClsMasterAkun1
        Dim ObjCurr As New Core_Module.clsMasterCurrency
        Dim objSetting As New Core_Module.ClsTransaksiSetting
        Dim PrimaryCurrBase As Double = Core_Module.PrimaryBaseCurrency
        Dim PrimaryAkunAPBase As Double = -1
        Dim PrimaryAkunARBase As Double = -1
        Dim PrimaryAkunDPAPBase As Double = objSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                                            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP)

        Dim PrimaryAkunDPARBase As Double = objSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                                            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Invoice_DP)
        Try
            '-------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------
            'tbltrans_aop
            '-------------------------------------------------------------------------------------------------------------------------
            Dim DtaAOP As New DseUpdate_v180TableAdapters.tbltrans_aopTableAdapter
            Dim TblAOP As New DseUpdate_v180.tbltrans_aopDataTable
            Dim RowAOP As DseUpdate_v180.tbltrans_aopRow
            DtaAOP.Connection.ConnectionString = Core_Module.GetConstring
            DtaAOP.Fill(TblAOP)
            Me.prg_v180_multiAcc.Value = 0
            Me.prg_v180_multiAcc.Maximum = TblAOP.Rows.Count
            For Each RowAOP In TblAOP.Rows
                With RowAOP
                    Me.prg_v180_multiAcc.Value += 1
                    Me.lbl_v180_multiAcc.Text = "Updating Data AOP " + Me.prg_v180_multiAcc.Value.ToString + " of " + Me.prg_v180_multiAcc.Maximum.ToString
                    If .Primary_Curr_Tax = -1 Then
                        PrimaryAkunAPBase = ObjCurr.GetPrimaryAkunAP(.Primary_Currency)
                        If .Tipe_AOP = 0 Then
                            DtaAOP.UpdateQuery(PrimaryCurrBase, PrimaryAkunAPBase, PrimaryAkunDPAPBase, .Primary)
                        ElseIf .Tipe_AOP = 1 Then
                            If .Primary_Akun_DP <> -1 Then
                                DtaAOP.UpdateQuery(PrimaryCurrBase, PrimaryAkunAPBase, PrimaryAkunDPAPBase, .Primary)
                            Else
                                DtaAOP.UpdateQuery(PrimaryCurrBase, PrimaryAkunAPBase, -1, .Primary)
                            End If
                        End If
                    End If
                End With
            Next
            '-------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------
            'tbltrans_purchasereturn
            '-------------------------------------------------------------------------------------------------------------------------
            Dim DtaPReturn As New DseUpdate_v180TableAdapters.tbltrans_purchasereturnTableAdapter
            Dim TblPReturn As New DseUpdate_v180.tbltrans_purchasereturnDataTable
            Dim RowPReturn As DseUpdate_v180.tbltrans_purchasereturnRow
            DtaPReturn.Connection.ConnectionString = Core_Module.GetConstring
            DtaPReturn.Fill(TblPReturn)
            Me.prg_v180_multiAcc.Value = 0
            Me.prg_v180_multiAcc.Maximum = TblPReturn.Rows.Count
            For Each RowPReturn In TblPReturn.Rows
                With RowPReturn
                    Me.prg_v180_multiAcc.Value += 1
                    Me.lbl_v180_multiAcc.Text = "Updating Data Purchase Return " + Me.prg_v180_multiAcc.Value.ToString + " of " + Me.prg_v180_multiAcc.Maximum.ToString
                    If .Primary_Curr_Tax = -1 Then
                        PrimaryAkunAPBase = ObjCurr.GetPrimaryAkunAP(.Primary_Currency)
                        DtaPReturn.UpdateQuery(PrimaryAkunAPBase, PrimaryCurrBase, .Primary)
                    End If
                End With
            Next
            '-------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------
            'tbltrans_salesinvoice
            '-------------------------------------------------------------------------------------------------------------------------
            Dim DtaInvoice As New DseUpdate_v180TableAdapters.tbltrans_salesinvoiceTableAdapter
            Dim TblInvoice As New DseUpdate_v180.tbltrans_salesinvoiceDataTable
            Dim RowInvoice As DseUpdate_v180.tbltrans_salesinvoiceRow
            DtaInvoice.Connection.ConnectionString = Core_Module.GetConstring
            DtaInvoice.Fill(TblInvoice)
            Me.prg_v180_multiAcc.Value = 0
            Me.prg_v180_multiAcc.Maximum = TblInvoice.Rows.Count
            For Each RowInvoice In TblInvoice.Rows
                With RowInvoice
                    Me.prg_v180_multiAcc.Value += 1
                    Me.lbl_v180_multiAcc.Text = "Updating Data Invoice " + Me.prg_v180_multiAcc.Value.ToString + " of " + Me.prg_v180_multiAcc.Maximum.ToString
                    If .Primary_Curr_Tax = -1 Then
                        PrimaryAkunARBase = ObjCurr.GetPrimaryAkunAR(.Primary_Currency)
                        If .Tipe_Invoice = 0 Then
                            DtaInvoice.UpdateQuery(PrimaryAkunARBase, PrimaryCurrBase, PrimaryAkunDPARBase, 0, PrimaryCurrBase, .Primary)
                        ElseIf .Tipe_Invoice = 1 Then
                            If .Primary_Akun_DP <> -1 Then
                                DtaInvoice.UpdateQuery(PrimaryAkunARBase, PrimaryCurrBase, PrimaryAkunDPARBase, 0, PrimaryCurrBase, .Primary)
                            Else
                                DtaInvoice.UpdateQuery(PrimaryAkunARBase, PrimaryCurrBase, -1, 0, PrimaryCurrBase, .Primary)
                            End If
                        End If
                    End If
                End With
            Next
            '-------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------
            'tbltrans_Salesreturn
            '-------------------------------------------------------------------------------------------------------------------------
            Dim DtaSReturn As New DseUpdate_v180TableAdapters.tbltrans_salesreturnTableAdapter
            Dim TblSReturn As New DseUpdate_v180.tbltrans_salesreturnDataTable
            Dim RowSReturn As DseUpdate_v180.tbltrans_salesreturnRow
            DtaSReturn.Connection.ConnectionString = Core_Module.GetConstring
            DtaSReturn.Fill(TblSReturn)
            Me.prg_v180_multiAcc.Value = 0
            Me.prg_v180_multiAcc.Maximum = TblSReturn.Rows.Count
            For Each RowSReturn In TblSReturn.Rows
                With RowSReturn
                    Me.prg_v180_multiAcc.Value += 1
                    Me.lbl_v180_multiAcc.Text = "Updating Data Sales Return " + Me.prg_v180_multiAcc.Value.ToString + " of " + Me.prg_v180_multiAcc.Maximum.ToString
                    If .Primary_Curr_Tax = -1 Then
                        PrimaryAkunAPBase = ObjCurr.GetPrimaryAkunAP(.Primary_Currency)
                        DtaSReturn.UpdateQuery(PrimaryAkunARBase, PrimaryCurrBase, .Primary)
                    End If
                End With
            Next
            '-------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------
            'tbltrans_historicalpurchasebalancemanual
            '-------------------------------------------------------------------------------------------------------------------------
            Dim DtaHPurchase As New DseUpdate_v180TableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
            Dim TblHPurchase As New DseUpdate_v180.tbltrans_historicalpurchasebalancemanualDataTable
            Dim RowHPurchase As DseUpdate_v180.tbltrans_historicalpurchasebalancemanualRow
            DtaHPurchase.Connection.ConnectionString = Core_Module.GetConstring
            DtaHPurchase.Fill(TblHPurchase)
            Me.prg_v180_multiAcc.Value = 0
            Me.prg_v180_multiAcc.Maximum = TblHPurchase.Rows.Count
            For Each RowHPurchase In TblHPurchase.Rows
                With RowHPurchase
                    Me.prg_v180_multiAcc.Value += 1
                    Me.lbl_v180_multiAcc.Text = "Updating Data Historical Purchase Balance " + Me.prg_v180_multiAcc.Value.ToString + " of " + Me.prg_v180_multiAcc.Maximum.ToString
                    If .primary_ap_account = -1 Then
                        PrimaryAkunAPBase = ObjCurr.GetPrimaryAkunAP(.primary_currency)
                        DtaHPurchase.UpdateQuery(PrimaryAkunAPBase, .Primary)
                    End If
                End With
            Next
            '-------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------
            'tbltrans_historicalpurchasereturnbalancemanual
            '-------------------------------------------------------------------------------------------------------------------------
            Dim DtaHPReturn As New DseUpdate_v180TableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
            Dim TblHPReturn As New DseUpdate_v180.tbltrans_historicalpurchasereturnbalancemanualDataTable
            Dim RowHPReturn As DseUpdate_v180.tbltrans_historicalpurchasereturnbalancemanualRow
            DtaHPReturn.Connection.ConnectionString = Core_Module.GetConstring
            DtaHPReturn.Fill(TblHPReturn)
            Me.prg_v180_multiAcc.Value = 0
            Me.prg_v180_multiAcc.Maximum = TblHPReturn.Rows.Count
            For Each RowHPReturn In TblHPReturn.Rows
                With RowHPReturn
                    Me.prg_v180_multiAcc.Value += 1
                    Me.lbl_v180_multiAcc.Text = "Updating Data Historical Purchase Return Balance " + Me.prg_v180_multiAcc.Value.ToString + " of " + Me.prg_v180_multiAcc.Maximum.ToString
                    If .primary_ap_account = -1 Then
                        PrimaryAkunAPBase = ObjCurr.GetPrimaryAkunAP(.primary_currency)
                        DtaHPReturn.UpdateQuery(PrimaryAkunAPBase, .Primary)
                    End If
                End With
            Next
            '-------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------
            'tbltrans_historicalsales
            '-------------------------------------------------------------------------------------------------------------------------
            Dim DtaHSales As New DseUpdate_v180TableAdapters.tbltrans_historicalsalesTableAdapter
            Dim TblHSales As New DseUpdate_v180.tbltrans_historicalsalesDataTable
            Dim RowHSales As DseUpdate_v180.tbltrans_historicalsalesRow
            DtaHSales.Connection.ConnectionString = Core_Module.GetConstring
            DtaHSales.Fill(TblHSales)
            Me.prg_v180_multiAcc.Value = 0
            Me.prg_v180_multiAcc.Maximum = TblHSales.Rows.Count
            For Each RowHSales In TblHSales.Rows
                With RowHSales
                    Me.prg_v180_multiAcc.Value += 1
                    Me.lbl_v180_multiAcc.Text = "Updating Data Historical Sales Balance " + Me.prg_v180_multiAcc.Value.ToString + " of " + Me.prg_v180_multiAcc.Maximum.ToString
                    If .primary_ar_account = -1 Then
                        PrimaryAkunARBase = ObjCurr.GetPrimaryAkunAR(.primary_currency)
                        DtaHSales.UpdateQuery(PrimaryAkunARBase, .Primary)
                    End If
                End With
            Next
            '-------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------
            'tbltrans_historicalsalesReturn
            '-------------------------------------------------------------------------------------------------------------------------
            Dim DtaHSReturn As New DseUpdate_v180TableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
            Dim TblHSReturn As New DseUpdate_v180.tbltrans_historicalsalesreturn_manualDataTable
            Dim RowHSReturn As DseUpdate_v180.tbltrans_historicalsalesreturn_manualRow
            DtaHSReturn.Connection.ConnectionString = Core_Module.GetConstring
            DtaHSReturn.Fill(TblHSReturn)
            Me.prg_v180_multiAcc.Value = 0
            Me.prg_v180_multiAcc.Maximum = TblHSReturn.Rows.Count
            For Each RowHSReturn In TblHSReturn.Rows
                With RowHSReturn
                    Me.prg_v180_multiAcc.Value += 1
                    Me.lbl_v180_multiAcc.Text = "Updating Data Historical Sales Return Balance " + Me.prg_v180_multiAcc.Value.ToString + " of " + Me.prg_v180_multiAcc.Maximum.ToString
                    If .primary_ar_account = -1 Then
                        PrimaryAkunARBase = ObjCurr.GetPrimaryAkunAR(.primary_currency)
                        DtaHSReturn.UpdateQuery(PrimaryAkunARBase, .Primary)
                    End If
                End With
            Next
            '-------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------
            'tbltrans_historicalsalesReturn
            '-------------------------------------------------------------------------------------------------------------------------
            Dim DtaJurnal As New DseUpdate_v180TableAdapters.tbltrans_jurnalTableAdapter
            Dim TblJurnal As New DseUpdate_v180.tbltrans_jurnalDataTable
            Dim RowJurnal As DseUpdate_v180.tbltrans_jurnalRow
            DtaJurnal.UpdateQuery("")

            'DtaHSReturn.Connection.ConnectionString = Core_Module.GetConstring
            'DtaHSReturn.Fill(TblHSReturn)
            'Me.prg_v180_multiAcc.Value = 0
            'Me.prg_v180_multiAcc.Maximum = TblHSReturn.Rows.Count
            'For Each RowHSReturn In TblHSReturn.Rows
            '    With RowHSReturn
            '        Me.prg_v180_multiAcc.Value += 1
            '        Me.lbl_v180_multiAcc.Text = "Updating Data Historical Sales Return Balance " + Me.prg_v180_multiAcc.Value.ToString + " of " + Me.prg_v180_multiAcc.Maximum.ToString
            '        If .primary_ar_account = -1 Then
            '            PrimaryAkunARBase = ObjCurr.GetPrimaryAkunAR(.primary_currency)
            '            DtaHSReturn.UpdateQuery(PrimaryAkunARBase, .Primary)
            '        End If
            '    End With
            'Next
            '-------------------------------------------------------------------------------------------------------------------------
            DtaAOP.Dispose()
            TblAOP.Dispose()
            DtaPReturn.Dispose()
            TblPReturn.Dispose()
            DtaInvoice.Dispose()
            TblInvoice.Dispose()
            DtaSReturn.Dispose()
            TblSReturn.Dispose()
            DtaHPurchase.Dispose()
            TblHPurchase.Dispose()
            DtaHPReturn.Dispose()
            TblHPReturn.Dispose()
            DtaHSales.Dispose()
            TblHSales.Dispose()
            DtaHSReturn.Dispose()
            TblHSReturn.Dispose()
            DtaJurnal.Dispose()
            TblJurnal.Dispose()
            ObjAkun = Nothing
            ObjCurr = Nothing
            objSetting = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_v180_Multiunit()
        Dim StartPrimary, EndPrimary As Double
        Dim StartIndex As Double = 0
        Dim BatchQty As Double = 10000

        Try
            '============================================================================================================================================
            '============================================================================================================================================
            'Jalur Purchase
            '============================================================================================================================================
            '============================================================================================================================================
            ''tbltrans_po_item
            ''0. qty multi
            ''1. primary unit multi
            ''2. pob multi
            ''3. git multi
            Dim DtaPoItem As New DseUpdate_v180TableAdapters.tbltrans_po_itemTableAdapter
            Dim TblPoItem As New DseUpdate_v180.tbltrans_po_itemDataTable
            Dim RowPoItem As DseUpdate_v180.tbltrans_po_itemRow
            DtaPoItem.Connection.ConnectionString = Core_Module.GetConstring
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
            DtaPrItem.Connection.ConnectionString = Core_Module.GetConstring
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
            DtaPocItem.Connection.ConnectionString = Core_Module.GetConstring
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
            DtaPlItem.Connection.ConnectionString = Core_Module.GetConstring
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
            DtaGrnItem.Connection.ConnectionString = Core_Module.GetConstring
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
            DtaDqItem.Connection.ConnectionString = Core_Module.GetConstring
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
            DtaSiOItem.Connection.ConnectionString = Core_Module.GetConstring
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
            DtaAopItem.Connection.ConnectionString = Core_Module.GetConstring
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

            'tbltrans_material_planning_data_item
            Dim DTA_MP As New DseUpdate_v180TableAdapters.tbltrans_material_planning_itemTableAdapter
            Dim TBL_MP As New DseUpdate_v180.tbltrans_material_planning_itemDataTable
            Dim row_MP, rowMP1 As DseUpdate_v180.tbltrans_material_planning_itemRow
            DTA_MP.Connection.ConnectionString = Core_Module.GetConstring
            DTA_MP.Fill(TBL_MP)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TBL_MP.Rows.Count
            StartIndex = 0
            While StartIndex < TBL_MP.Rows.Count - 1
                Me.prg_v180_multiunit.Value = StartIndex + 1
                Me.lbl_v180_multiunit.Text = "Updating Data Material Planning Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString

                row_MP = TBL_MP.Rows(StartIndex)
                If StartIndex + BatchQty <= TBL_MP.Rows.Count - 1 Then
                    rowMP1 = TBL_MP.Rows(StartIndex + BatchQty)
                Else
                    rowMP1 = TBL_MP.Rows(TBL_MP.Rows.Count - 1)
                End If
                StartPrimary = row_MP.primary
                EndPrimary = rowMP1.primary
                DTA_MP.UpdateQuery2(1, 0, 0, StartPrimary, EndPrimary)
                StartIndex += BatchQty
            End While

            ''============================================================================================================================================
            ''============================================================================================================================================
            ''Jalur Sales
            ''============================================================================================================================================
            ''============================================================================================================================================
            Dim DtaquoItem As New DseUpdate_v180TableAdapters.tbltrans_quotation_itemTableAdapter
            Dim TblquoItem As New DseUpdate_v180.tbltrans_quotation_itemDataTable
            Dim RowquoItem, RowquoItem1 As DseUpdate_v180.tbltrans_quotation_itemRow
            DtaquoItem.Connection.ConnectionString = Core_Module.GetConstring
            'DtaPoItem.Connection.ConnectionString  =
            DtaquoItem.Fill(TblquoItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblquoItem.Rows.Count
            StartIndex = 0
            While StartIndex < TblquoItem.Rows.Count - 1
                Me.prg_v180_multiunit.Value = StartIndex + 1
                Me.lbl_v180_multiunit.Text = "Updating Data Quotation Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString

                RowquoItem = TblquoItem.Rows(StartIndex)
                If StartIndex + BatchQty <= TblquoItem.Rows.Count - 1 Then
                    RowquoItem1 = TblquoItem.Rows(StartIndex + BatchQty)
                Else
                    RowquoItem1 = TblquoItem.Rows(TblquoItem.Rows.Count - 1)
                End If

                StartPrimary = RowquoItem.Primary
                EndPrimary = RowquoItem1.Primary
                DtaquoItem.UpdateQuery1(StartPrimary, EndPrimary)
                StartIndex += BatchQty
            End While

            Dim DtaSOItem As New DseUpdate_v180TableAdapters.tbltrans_so_itemTableAdapter
            Dim TblSOItem As New DseUpdate_v180.tbltrans_so_itemDataTable
            Dim RowSOItem, RowSOItem1 As DseUpdate_v180.tbltrans_so_itemRow
            DtaSOItem.Connection.ConnectionString = Core_Module.GetConstring
            'DtaPoItem.Connection.ConnectionString  =
            DtaSOItem.Fill(TblSOItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblSOItem.Rows.Count
            StartIndex = 0
            While StartIndex < TblSOItem.Rows.Count - 1
                Me.prg_v180_multiunit.Value = StartIndex + 1
                Me.lbl_v180_multiunit.Text = "Updating Data Sales Order Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString

                RowSOItem = TblSOItem.Rows(StartIndex)
                If StartIndex + BatchQty <= TblSOItem.Rows.Count - 1 Then
                    RowSOItem1 = TblSOItem.Rows(StartIndex + BatchQty)
                Else
                    RowSOItem1 = TblSOItem.Rows(TblSOItem.Rows.Count - 1)
                End If

                StartPrimary = RowSOItem.Primary
                EndPrimary = RowSOItem1.Primary
                DtaSOItem.UpdateQuery1(StartPrimary, EndPrimary)
                StartIndex += BatchQty
            End While

            Dim DtaSODItem As New DseUpdate_v180TableAdapters.tbltrans_so_itemdeliveryTableAdapter
            Dim TblSODItem As New DseUpdate_v180.tbltrans_so_itemdeliveryDataTable
            Dim RowSODItem, RowSODItem1 As DseUpdate_v180.tbltrans_so_itemdeliveryRow
            DtaSODItem.Connection.ConnectionString = Core_Module.GetConstring
            'DtaPoItem.Connection.ConnectionString  =
            DtaSODItem.Fill(TblSODItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblSODItem.Rows.Count
            StartIndex = 0
            While StartIndex < TblSODItem.Rows.Count - 1
                Me.prg_v180_multiunit.Value = StartIndex + 1
                Me.lbl_v180_multiunit.Text = "Updating Data Sales Order Delivery Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString

                RowSODItem = TblSODItem.Rows(StartIndex)
                If StartIndex + BatchQty <= TblSODItem.Rows.Count - 1 Then
                    RowSODItem1 = TblSODItem.Rows(StartIndex + BatchQty)
                Else
                    RowSODItem1 = TblSODItem.Rows(TblSODItem.Rows.Count - 1)
                End If

                StartPrimary = RowSODItem.primary
                EndPrimary = RowSODItem1.primary
                DtaSODItem.UpdateQuery1(StartPrimary, EndPrimary)
                StartIndex += BatchQty
            End While

            Dim DtaSOCItem As New DseUpdate_v180TableAdapters.tbltrans_soc_itemTableAdapter
            Dim TblSOCItem As New DseUpdate_v180.tbltrans_soc_itemDataTable
            Dim RowSOCItem, RowSOCItem1 As DseUpdate_v180.tbltrans_soc_itemRow
            DtaSOCItem.Connection.ConnectionString = Core_Module.GetConstring
            'DtaPoItem.Connection.ConnectionString  =
            DtaSOCItem.Fill(TblSOCItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblSOCItem.Rows.Count
            StartIndex = 0
            While StartIndex < TblSOCItem.Rows.Count - 1
                Me.prg_v180_multiunit.Value = StartIndex + 1
                Me.lbl_v180_multiunit.Text = "Updating Data Sales Order cancelation Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString

                RowSOCItem = TblSOCItem.Rows(StartIndex)
                If StartIndex + BatchQty <= TblSOCItem.Rows.Count - 1 Then
                    RowSOCItem1 = TblSOCItem.Rows(StartIndex + BatchQty)
                Else
                    RowSOCItem1 = TblSOCItem.Rows(TblSOCItem.Rows.Count - 1)
                End If

                StartPrimary = RowSOCItem.Primary
                EndPrimary = RowSOCItem1.Primary
                DtaSOCItem.UpdateQuery1(StartPrimary, EndPrimary)
                StartIndex += BatchQty
            End While

            Dim DtaDOItem As New DseUpdate_v180TableAdapters.tbltrans_do_itemTableAdapter
            Dim TblDOItem As New DseUpdate_v180.tbltrans_do_itemDataTable
            Dim RowDOItem, RowDOItem1 As DseUpdate_v180.tbltrans_do_itemRow
            DtaDOItem.Connection.ConnectionString = Core_Module.GetConstring
            'DtaPoItem.Connection.ConnectionString  =
            DtaDOItem.Fill(TblDOItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblDOItem.Rows.Count
            StartIndex = 0
            While StartIndex < TblDOItem.Rows.Count - 1
                Me.prg_v180_multiunit.Value = StartIndex + 1
                Me.lbl_v180_multiunit.Text = "Updating Data DO Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString

                RowDOItem = TblDOItem.Rows(StartIndex)
                If StartIndex + BatchQty <= TblDOItem.Rows.Count - 1 Then
                    RowDOItem1 = TblDOItem.Rows(StartIndex + BatchQty)
                Else
                    RowDOItem1 = TblDOItem.Rows(TblDOItem.Rows.Count - 1)
                End If

                StartPrimary = RowDOItem.Primary
                EndPrimary = RowDOItem1.Primary
                DtaDOItem.UpdateQuery1(StartPrimary, EndPrimary)
                StartIndex += BatchQty
            End While

            Dim DtaInvItem As New DseUpdate_v180TableAdapters.tbltrans_salesinvoice_itemTableAdapter
            Dim TblInvItem As New DseUpdate_v180.tbltrans_salesinvoice_itemDataTable
            Dim RowInvItem, RowInvItem1 As DseUpdate_v180.tbltrans_salesinvoice_itemRow
            DtaInvItem.Connection.ConnectionString = Core_Module.GetConstring
            'DtaPoItem.Connection.ConnectionString  =
            DtaInvItem.Fill(TblInvItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblInvItem.Rows.Count
            StartIndex = 0
            While StartIndex < TblInvItem.Rows.Count - 1
                Me.prg_v180_multiunit.Value = StartIndex + 1
                Me.lbl_v180_multiunit.Text = "Updating Data invoice Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString

                RowInvItem = TblInvItem.Rows(StartIndex)
                If StartIndex + BatchQty <= TblInvItem.Rows.Count - 1 Then
                    RowInvItem1 = TblInvItem.Rows(StartIndex + BatchQty)
                Else
                    RowInvItem1 = TblInvItem.Rows(TblInvItem.Rows.Count - 1)
                End If

                StartPrimary = RowInvItem.Primary
                EndPrimary = RowInvItem1.Primary
                DtaInvItem.UpdateQuery1(StartPrimary, EndPrimary)
                StartIndex += BatchQty
            End While

            Dim DtaSRItem As New DseUpdate_v180TableAdapters.tbltrans_salesreturn_itemTableAdapter
            Dim TblSRItem As New DseUpdate_v180.tbltrans_salesreturn_itemDataTable
            Dim RowSRItem, RowSRItem1 As DseUpdate_v180.tbltrans_salesreturn_itemRow
            DtaSRItem.Connection.ConnectionString = Core_Module.GetConstring
            'DtaPoItem.Connection.ConnectionString  =
            DtaSRItem.Fill(TblSRItem)
            Me.prg_v180_multiunit.Value = 0
            Me.prg_v180_multiunit.Maximum = TblSRItem.Rows.Count
            StartIndex = 0
            While StartIndex < TblSRItem.Rows.Count - 1
                Me.prg_v180_multiunit.Value = StartIndex + 1
                Me.lbl_v180_multiunit.Text = "Updating Data Sales Return Item " + Me.prg_v180_multiunit.Value.ToString + " of " + Me.prg_v180_multiunit.Maximum.ToString

                RowSRItem = TblSRItem.Rows(StartIndex)
                If StartIndex + BatchQty <= TblSRItem.Rows.Count - 1 Then
                    RowSRItem1 = TblSRItem.Rows(StartIndex + BatchQty)
                Else
                    RowSRItem1 = TblSRItem.Rows(TblSRItem.Rows.Count - 1)
                End If

                StartPrimary = RowSRItem.Primary
                EndPrimary = RowSRItem1.Primary
                DtaSRItem.UpdateQuery1(StartPrimary, EndPrimary)
                StartIndex += BatchQty
            End While

            DTA_MP.Dispose()
            TBL_MP.Dispose()
            DtaSRItem.Dispose()
            TblSRItem.Dispose()
            DtaInvItem.Dispose()
            TblInvItem.Dispose()
            DtaDOItem.Dispose()
            TblDOItem.Dispose()
            DtaSOItem.Dispose()
            TblSOItem.Dispose()
            DtaSODItem.Dispose()
            TblSODItem.Dispose()
            DtaSOCItem.Dispose()
            TblSOCItem.Dispose()
            DtaquoItem.Dispose()
            TblquoItem.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message + ": " + Me.prg_v180_multiunit.Value.ToString, MsgBoxStyle.Critical, "Error")

        End Try
    End Sub

    Private Sub btn_v180_multiAcc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_v180_multiAcc.Click
        Try
            If MsgBox("Start multi account update?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BgwMultiAcc.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwMultiAcc_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwMultiAcc.DoWork

        Try
            Me.Update_v180_MultiAcc()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwMultiAcc_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwMultiAcc.RunWorkerCompleted

        Try
            Me.lbl_v180_multiAcc.Text = "Data Updated"
            MsgBox("Update completed.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwPR_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwPR.DoWork

        Try
            Me.Update_v180_PR()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwPR_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwPR.RunWorkerCompleted

        Try
            Me.lbl_v180_PR.Text = "Data Updated"
            MsgBox("Update completed.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn_v180_PR_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_v180_PR.Click
        Try
            If MsgBox("Start Purchase Request update?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BgwPR.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_v180_PR()
        '-------------------------------------------------------------------------------------------------------------------------
        'tbltrans_aop
        'tbltrans_purchasereturn
        'tbltrans_salesinvoice
        'tbltrans_Salesreturn
        'tbltrans_historicalpurchasebalancemanual
        'tbltrans_historicalpurchasereturnbalancemanual
        'tbltrans_historicalsales
        'tbltrans_historicalsalesReturn
        'tbltrans_jurnal
        '-------------------------------------------------------------------------------------------------------------------------
        Dim ObjAkun As New Core_Module.ClsMasterAkun1
        Dim ObjCurr As New Core_Module.clsMasterCurrency
        Dim objSetting As New Core_Module.ClsTransaksiSetting
        Dim PrimaryCurrBase As Double = Core_Module.PrimaryBaseCurrency
        Dim PrimaryAkunAPBase As Double = -1
        Dim PrimaryAkunARBase As Double = -1
        Try
            '-------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------
            'tbltrans_aop
            '-------------------------------------------------------------------------------------------------------------------------
            Dim DtaAOP As New DseUpdate_v180TableAdapters.tbltrans_purchase_req_itemTableAdapter
            Dim TblAOP As New DseUpdate_v180.tbltrans_purchase_req_itemDataTable
            Dim RowAOP As DseUpdate_v180.tbltrans_purchase_req_itemRow
            Dim DtaC As New DseUpdate_v180TableAdapters.CancelTableAdapter
            Dim TblC As New DseUpdate_v180.CancelDataTable
            Dim RowC As DseUpdate_v180.CancelRow
            Dim DtaP As New DseUpdate_v180TableAdapters.PickupTableAdapter
            Dim TblP As New DseUpdate_v180.PickupDataTable
            Dim RowP As DseUpdate_v180.PickupRow
            DtaAOP.Connection.ConnectionString = Core_Module.GetConstring
            DtaC.Connection.ConnectionString = Core_Module.GetConstring
            DtaP.Connection.ConnectionString = Core_Module.GetConstring
            DtaAOP.Fill(TblAOP)
            DtaC.Fill(TblC)
            DtaP.Fill(TblP)

            Me.prg_v180_PR.Value = 0
            Me.prg_v180_PR.Maximum = TblAOP.Rows.Count
            For Each RowAOP In TblAOP.Rows
                With RowAOP
                    Me.prg_v180_PR.Value += 1
                    Me.lbl_v180_PR.Text = "Updating Data PR " + Me.prg_v180_PR.Value.ToString + " of " + Me.prg_v180_PR.Maximum.ToString
                    TblC.DefaultView.RowFilter = Nothing
                    TblC.DefaultView.RowFilter = "[Primary PR] = " & .primary_transaksi & " And [Primary Code No] = " & .primary_code_no
                    If TblC.DefaultView.Count > 0 Then
                        Dim decCancel As Decimal = 0
                        Dim dv As Data.DataRowView
                        For Each dv In TblC.DefaultView
                            RowC = dv.Row
                            If IsNumeric(RowC.Total) Then
                                decCancel = CDec(RowC.Total)
                            End If
                        Next
                        .Qty_Cancel = decCancel
                    End If

                    TblP.DefaultView.RowFilter = Nothing
                    TblP.DefaultView.RowFilter = "[Primary PR] = " & .primary_transaksi & " And [Primary Code No] = " & .primary_code_no
                    If TblP.DefaultView.Count > 0 Then
                        Dim decCancel As Decimal = 0
                        Dim dv As Data.DataRowView
                        For Each dv In TblP.DefaultView
                            RowP = dv.Row
                            If IsNumeric(RowP.Total) Then
                                decCancel = CDec(RowP.Total)
                            End If
                        Next
                        .Qty_Pick_Up = decCancel
                    End If


                    DtaAOP.UpdateQuery(.Qty_Pick_Up, .Qty_Cancel, .primary)
                End With
            Next

            DtaAOP.Dispose()
            TblAOP.Dispose()
            DtaP.Dispose()
            TblP.Dispose()
            DtaC.Dispose()
            TblC.Dispose()

            ObjAkun = Nothing
            ObjCurr = Nothing
            objSetting = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn_v180_SC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_v180_SC.Click
        Try
            If MsgBox("Update stock card?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.Bgw_180_SC.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Bgw_180_SC_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles Bgw_180_SC.DoWork
        Me.UpdateStockCard_v180()
    End Sub

    Private Sub Bgw_180_SC_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles Bgw_180_SC.RunWorkerCompleted
        MsgBox("Update finish")
    End Sub

    Private Sub UpdateStockCard_v180()
        Dim sql As String = ""

        Try
            'update data
            '0. hapus data, kecuali yang berasal dari stock taking
            Me.HapusStockCard_v160()
            '1. stock in out
            Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut
            ObjStockInOut.UpdateStockCard_TransaksiStockInOut(Me.prg_v180_SC, Me.lbl_v180_SC)
            ObjStockInOut = Nothing
            '2. bundling
            Dim ObjStockBundling As New Inventory_Module.ClsTransaksiStockBundling
            ObjStockBundling.UpdateStockCard_TransaksiStockBundling(Me.prg_v180_SC, Me.lbl_v180_SC)
            ObjStockBundling = Nothing

            Me.lbl_v180_SC.Text = "done"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwSalary_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwSalary.DoWork
        Try
            Me.Update_v180_Salary()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwSalary_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwSalary.RunWorkerCompleted
        Try
            Me.lbl_v180_Salary.Text = "Data Updated"
            MsgBox("Update completed.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn_v180_Salary_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_v180_Salary.Click
        Try
            If MsgBox("Start Salary Component update?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BgwSalary.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_v180_Salary()
        '-------------------------------------------------------------------------------------------------------------------------
        'tbltrans_aop
        'tbltrans_purchasereturn
        'tbltrans_salesinvoice
        'tbltrans_Salesreturn
        'tbltrans_historicalpurchasebalancemanual
        'tbltrans_historicalpurchasereturnbalancemanual
        'tbltrans_historicalsales
        'tbltrans_historicalsalesReturn
        'tbltrans_jurnal
        '-------------------------------------------------------------------------------------------------------------------------
        Try
            '-------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------
            'tbltrans_aop
            '-------------------------------------------------------------------------------------------------------------------------
            Dim DtaSC As New DseUpdate_v180TableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
            Dim TblSC As New DseUpdate_v180.tblmaster_salary_formula_salarycomponentDataTable
            Dim RowSC As DseUpdate_v180.tblmaster_salary_formula_salarycomponentRow

            Dim DtaEM As New DseUpdate_v180TableAdapters.tblmaster_salary_formula_component_detailemployeeTableAdapter
            Dim TblEM As New DseUpdate_v180.tblmaster_salary_formula_component_detailemployeeDataTable
            Dim RowEM As DseUpdate_v180.tblmaster_salary_formula_component_detailemployeeRow

            DtaSC.Connection.ConnectionString = Core_Module.GetConstring
            DtaEM.Connection.ConnectionString = Core_Module.GetConstring

            DtaSC.Fill(TblSC)
            DtaEM.Fill(TblEM)

            Me.prg_v180_Salary.Value = 0
            Me.prg_v180_Salary.Maximum = TblSC.Rows.Count + TblEM.Rows.Count
            For Each RowSC In TblSC.Rows
                With RowSC
                    Me.prg_v180_Salary.Value += 1
                    Me.lbl_v180_Salary.Text = "Updating Data Salary Component " + Me.prg_v180_Salary.Value.ToString + " of " + Me.prg_v180_Salary.Maximum.ToString
                    .Created_by = .created_by_H
                    .Created_Date = .created_date_H
                    .Modified_by = .modified_by_H
                    .Modified_Date = .modified_date_H
                    DtaSC.UpdateQuery(.Created_by, .Created_Date, .Modified_by, .Modified_Date, .Primary)
                End With
            Next

            For Each RowEM In TblEM.Rows
                With RowEM
                    Me.prg_v180_Salary.Value += 1
                    Me.lbl_v180_Salary.Text = "Updating Data Salary Component " + Me.prg_v180_Salary.Value.ToString + " of " + Me.prg_v180_Salary.Maximum.ToString
                    .Created_by = .created_by_H
                    .Created_Date = .created_date_H
                    .Modified_by = .modified_by_H
                    .Modified_Date = .modified_date_H
                    DtaEM.UpdateQuery(.Created_by, .Created_Date, .Modified_by, .Modified_Date, .Primary)
                End With
            Next

            DtaSC.Dispose()
            TblSC.Dispose()
            DtaEM.Dispose()
            TblEM.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn_v180_DetailClosing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_v180_DetailClosing.Click
        Dim DTA1 As New DseUpdate_v180TableAdapters.tbltrans_purchase_closing_item_detailTableAdapter

        Try
            'tipe item 5 -> -1
            DTA1.UpdateQuery(-5, 5)
            'tipe item 6 -> 5
            DTA1.UpdateQuery(5, 6)
            'tipe item -1 -> 6
            DTA1.UpdateQuery(6, -5)

            'tipe item 2 -> -1
            DTA1.UpdateQuery(-5, 2)
            'tipe item 3 -> 2
            DTA1.UpdateQuery(2, 3)
            'tipe item 4 -> 3
            DTA1.UpdateQuery(3, 4)
            'tipe item 5 -> 4
            DTA1.UpdateQuery(4, 5)
            'tipe item -1 -> 5
            DTA1.UpdateQuery(5, -5)

            MsgBox("Selesai")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn_v180_DetailItemJurnal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_v180_DetailItemJurnal.Click
        Dim DTA1 As New DseUpdate_v180TableAdapters.tbltrans_purchase_closing_item_akunTableAdapter

        Try
            'tipe item 5 -> -1
            DTA1.UpdateQuery(-5, 5)
            'tipe item 6 -> 5
            DTA1.UpdateQuery(5, 6)
            'tipe item -1 -> 6
            DTA1.UpdateQuery(6, -5)

            'tipe item 2 -> -1
            DTA1.UpdateQuery(-5, 2)
            'tipe item 3 -> 2
            DTA1.UpdateQuery(2, 3)
            'tipe item 4 -> 3
            DTA1.UpdateQuery(3, 4)
            'tipe item 5 -> 4
            DTA1.UpdateQuery(4, 5)
            'tipe item -1 -> 5
            DTA1.UpdateQuery(5, -5)

            MsgBox("Selesai")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn_v180_ProsesBundling_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_v180_ProsesBundling.Click
        Dim ij As New FrmUpdate_v180_UpdateProsesBundling

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button13a_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        MsgBox("Pastikan akun DP dan akun Hutangnya tepat.", MsgBoxStyle.Information, "Information")
    End Sub

    Private Sub BgwSubtotal_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwSubtotal.DoWork
        Try
            Me.Update_v180_Subtotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwSubtotal_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwSubtotal.RunWorkerCompleted
        Try
            Me.lbl_v180_Subtotal.Text = "Data Updated"
            MsgBox("Update completed.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn_v180_Subtotal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_v180_Subtotal.Click
        Try
            If MsgBox("Start Invoice subtotal update?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BgwSubtotal.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_v180_Subtotal()
        '-------------------------------------------------------------------------------------------------------------------------
        'tbltrans_aop
        'tbltrans_purchasereturn
        'tbltrans_salesinvoice
        'tbltrans_Salesreturn
        'tbltrans_historicalpurchasebalancemanual
        'tbltrans_historicalpurchasereturnbalancemanual
        'tbltrans_historicalsales
        'tbltrans_historicalsalesReturn
        'tbltrans_jurnal
        '-------------------------------------------------------------------------------------------------------------------------
        Dim Objcd As New Core_Module.ClsCekData
        Dim Objcurr As New Core_Module.clsMasterCurrency
        Try
            '-------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------
            'tbltrans_aop
            '-------------------------------------------------------------------------------------------------------------------------
            Dim DtaSC As New DseUpdate_v180TableAdapters.Update_SubtotalTableAdapter
            Dim TblSC As New DseUpdate_v180.Update_SubtotalDataTable
            Dim RowSC As DseUpdate_v180.Update_SubtotalRow

            DtaSC.Connection.ConnectionString = Core_Module.GetConstring

            DtaSC.Fill(TblSC)

            Me.prg_v180_Subtotal.Value = 0
            Me.prg_v180_Subtotal.Maximum = TblSC.Rows.Count
            For Each RowSC In TblSC.Rows
                With RowSC
                    Me.prg_v180_Subtotal.Value += 1
                    Me.lbl_v180_Subtotal.Text = "Updating Data Subtotal Component " + Me.prg_v180_Subtotal.Value.ToString + " of " + Me.prg_v180_Subtotal.Maximum.ToString
                    Dim Subtotal As Decimal = 0
                    Subtotal = CDec(.Sub_Total_Before_DP + .ongkos_kirim - .Down_Payment)
                    Dim strSubtotal As String = 0
                    strSubtotal = Objcd.DigitGrouping(Subtotal, Objcurr.GetDefaultPrecission(.Primary_Currency), Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                    If IsNumeric(strSubtotal) Then
                        Subtotal = CDec(strSubtotal)
                    End If
                    DtaSC.UpdateQuery(Subtotal, .Primary)
                End With
            Next



            DtaSC.Dispose()
            TblSC.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn_v180_BundlingDate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_v180_BundlingDate.Click
        Dim DTA As New DseUpdate_v180TableAdapters.tbltrans_custom_update_bundlingTableAdapter

        Try
            If MsgBox("Start update?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                DTA.UpdateQuery_SetDateRef(#1/1/1950#)
                MsgBox("Data updated", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub btn_v180_JournalPaymentMethod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_v180_JournalPaymentMethod.Click
        Try
            MsgBox("Pastikan kolom payment method pada journal = ''", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn_v180_QuotationLimit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_v180_QuotationLimit.Click
        Try
            MsgBox("Pastikan kembali nilai limit quotation lvl 2 -> 50,000,000")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn_v180_MasterCust_TrimCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_v180_MasterCust_TrimCode.Click
        Try
            If MsgBox("Start customer code update?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.Bgw_MasterCust.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Bgw_MasterCust_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles Bgw_MasterCust.DoWork
        Try

            Me.Update_v180_MasterCust_TrimCode()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Bgw_MasterCust_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles Bgw_MasterCust.RunWorkerCompleted
        Try
            Me.lbl_v180_MasterCust_TrimCode.Text = "Data Updated"
            MsgBox("Update completed.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_v180_MasterCust_TrimCode()
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseUpdate_v180TableAdapters.tblmaster_customerTableAdapter
        Dim TBL As New DseUpdate_v180.tblmaster_customerDataTable
        Dim row As DseUpdate_v180.tblmaster_customerRow

        Try
            ObjCD.UpdateThreading()
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.Fill(TBL)
            Me.prg_v180_MasterCust_TrimCode.Value = 0
            Me.prg_v180_MasterCust_TrimCode.Maximum = TBL.Rows.Count

            For Each row In TBL.Rows
                DTA.UpdateQuery(row.Customer_Code.Trim, row.Primary)
                Me.prg_v180_MasterCust_TrimCode.Value += 1
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub btn_v180_UpdateJurnalItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_v180_UpdateJurnalItem2.Click
        Dim DTA1 As New DseUpdate_v180TableAdapters.tbltrans_purchase_closing_item_akunTableAdapter
        Dim DTA2 As New DseUpdate_v180TableAdapters.tbltrans_purchase_closing_item_detailTableAdapter

        Try
            'tipe item 5 -> -5
            DTA1.UpdateQuery(-5, 5)
            DTA2.UpdateQuery(-5, 5)
            'tipe item '6->5
            DTA1.UpdateQuery(5, 6)
            DTA2.UpdateQuery(5, 6)
            'tipe item 7->6
            DTA1.UpdateQuery(6, 7)
            DTA2.UpdateQuery(6, 7)
            'tipe item 8->7
            DTA1.UpdateQuery(7, 8)
            DTA2.UpdateQuery(7, 8)
            'tipe item 9->8
            DTA1.UpdateQuery(8, 9)
            DTA2.UpdateQuery(8, 9)
            'tipe item  -5 ->9
            DTA1.UpdateQuery(9, -5)
            DTA2.UpdateQuery(9, -5)

            'tipe item 4 -> -5
            DTA1.UpdateQuery(-5, 4)
            DTA2.UpdateQuery(-5, 4)
            'tipe item '5->4
            DTA1.UpdateQuery(4, 5)
            DTA2.UpdateQuery(4, 5)
            'tipe item 6->5
            DTA1.UpdateQuery(5, 6)
            DTA2.UpdateQuery(5, 6)
            'tipe item 7->6
            DTA1.UpdateQuery(6, 7)
            DTA2.UpdateQuery(6, 7)
            'tipe item 8->7
            DTA1.UpdateQuery(7, 8)
            DTA2.UpdateQuery(7, 8)
            'tipe item  -5 ->8
            DTA1.UpdateQuery(8, -5)
            DTA2.UpdateQuery(8, -5)

            MsgBox("Selesai")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            DTA2.Dispose()
        End Try
    End Sub

    Private Sub btn_v180_UpdateDiscAmount1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_v180_UpdateDiscAmount1.Click
        Dim DTAAOP As New DseUpdate_v180TableAdapters.tbltrans_aop_itemTableAdapter
        Dim DTAPR As New DseUpdate_v180TableAdapters.tbltrans_purchasereturn_itemTableAdapter

        Try
            DTAAOP.Connection.ConnectionString = Core_Module.GetConstring
            DTAPR.Connection.ConnectionString = DTAAOP.Connection.ConnectionString

            DTAAOP.UpdateQuery_SetDiscAmount()
            DTAPR.UpdateQuery_SetDiscAmount()

            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAAOP.Dispose()
            DTAPR.Dispose()
        End Try
    End Sub

    Private Sub btn_v180_UpdateSalary_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_v180_UpdateSalary.Click
        Try
            If MsgBox("Update tbltrans_salary: period start, period end, period count, primary dept, is resign?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.Bgw_v180_UpdateSalary.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Bgw_v180_UpdateSalary_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles Bgw_v180_UpdateSalary.DoWork
        Dim DTASalary As New DseUpdate_v180TableAdapters.tbltrans_salaryTableAdapter
        Dim TBLSalary As New DseUpdate_v180.tbltrans_salaryDataTable
        Dim rowSalary As DseUpdate_v180.tbltrans_salaryRow
        Dim DTAMasterKaryawan As New DseUpdate_v180TableAdapters.tblmaster_karyawanTableAdapter
        Dim TBLMasterKaryawan As New DseUpdate_v180.tblmaster_karyawanDataTable
        Dim rowMasterKaryawan As DseUpdate_v180.tblmaster_karyawanRow
        Dim StartWorkingDate, ResignDate As Date
        Dim PeriodCount, PeriodStart, PeriodEnd, IsResign As Integer
        Dim vResign As Boolean

        Try
            Me.ObjCD.UpdateThreading()
            DTASalary.Connection.ConnectionString = Core_Module.GetConstring
            DTAMasterKaryawan.Connection.ConnectionString = DTASalary.Connection.ConnectionString

            DTAMasterKaryawan.Fill(TBLMasterKaryawan)
            DTASalary.Fill(TBLSalary, 3)

            Me.prg_v180_UpdateSalary.Value = 0
            Me.prg_v180_UpdateSalary.Maximum = TBLMasterKaryawan.Rows.Count

            While TBLMasterKaryawan.Rows.Count > 0 And TBLSalary.Rows.Count > 0
                rowMasterKaryawan = TBLMasterKaryawan.Rows(0)
                rowSalary = TBLSalary.Rows(0)

                If rowMasterKaryawan.Primary = rowSalary.Primary_Employee Then
                    StartWorkingDate = rowMasterKaryawan.first_tax_period
                    ResignDate = rowMasterKaryawan.Resign_Date

                    If ResignDate >= rowSalary.Start_Date.Date And ResignDate <= rowSalary.End_Date Then
                        vResign = True
                        IsResign = 1
                    Else
                        vResign = False
                        IsResign = 0
                    End If

                    If StartWorkingDate.Date.Year < Me.ObjCD.GetFirstDateofTheMonth(rowSalary.End_Date.Date).Date.Year Then
                        If ResignDate = #1/1/1950# Then
                            PeriodCount = 12
                            PeriodStart = 1
                            PeriodEnd = 12
                        Else
                            If ResignDate.Date <= rowSalary.End_Date.Date And ResignDate >= rowSalary.Start_Date.Date Then
                                PeriodCount = ResignDate.Month
                                PeriodStart = 1
                                PeriodEnd = PeriodCount
                            Else
                                PeriodCount = 12
                                PeriodStart = 1
                                PeriodEnd = 12
                            End If
                        End If
                    ElseIf StartWorkingDate.Date.Year = Me.ObjCD.GetFirstDateofTheMonth(rowSalary.End_Date.Date).Date.Year Then
                        If ResignDate = #1/1/1950# Then
                            If StartWorkingDate.Month = 1 Then
                                PeriodCount = 12
                            ElseIf StartWorkingDate.Month = 2 Then
                                PeriodCount = 11
                            ElseIf StartWorkingDate.Month = 3 Then
                                PeriodCount = 10
                            ElseIf StartWorkingDate.Month = 4 Then
                                PeriodCount = 9
                            ElseIf StartWorkingDate.Month = 5 Then
                                PeriodCount = 8
                            ElseIf StartWorkingDate.Month = 6 Then
                                PeriodCount = 7
                            ElseIf StartWorkingDate.Month = 7 Then
                                PeriodCount = 6
                            ElseIf StartWorkingDate.Month = 8 Then
                                PeriodCount = 5
                            ElseIf StartWorkingDate.Month = 9 Then
                                PeriodCount = 4
                            ElseIf StartWorkingDate.Month = 10 Then
                                PeriodCount = 3
                            ElseIf StartWorkingDate.Month = 11 Then
                                PeriodCount = 2
                            ElseIf StartWorkingDate.Month = 12 Then
                                PeriodCount = 1
                            End If
                            PeriodStart = StartWorkingDate.Month
                            PeriodEnd = 12
                        Else
                            If ResignDate.Date <= rowSalary.End_Date.Date And ResignDate >= rowSalary.Start_Date.Date Then
                                PeriodCount = ResignDate.Month - StartWorkingDate.Month + 1
                                PeriodStart = StartWorkingDate.Month
                                PeriodEnd = ResignDate.Month
                            Else
                                If StartWorkingDate.Month = 1 Then
                                    PeriodCount = 12
                                ElseIf StartWorkingDate.Month = 2 Then
                                    PeriodCount = 11
                                ElseIf StartWorkingDate.Month = 3 Then
                                    PeriodCount = 10
                                ElseIf StartWorkingDate.Month = 4 Then
                                    PeriodCount = 9
                                ElseIf StartWorkingDate.Month = 5 Then
                                    PeriodCount = 8
                                ElseIf StartWorkingDate.Month = 6 Then
                                    PeriodCount = 7
                                ElseIf StartWorkingDate.Month = 7 Then
                                    PeriodCount = 6
                                ElseIf StartWorkingDate.Month = 8 Then
                                    PeriodCount = 5
                                ElseIf StartWorkingDate.Month = 9 Then
                                    PeriodCount = 4
                                ElseIf StartWorkingDate.Month = 10 Then
                                    PeriodCount = 3
                                ElseIf StartWorkingDate.Month = 11 Then
                                    PeriodCount = 2
                                ElseIf StartWorkingDate.Month = 12 Then
                                    PeriodCount = 1
                                End If
                                PeriodStart = StartWorkingDate.Month
                                PeriodEnd = 12
                            End If
                        End If
                    Else
                        PeriodCount = 12
                        PeriodStart = 1
                        PeriodEnd = 12
                    End If


                    DTASalary.UpdateQuery(PeriodStart, PeriodEnd, PeriodCount, rowMasterKaryawan.primary_departemen, IsResign, rowSalary.Primary)
                    TBLSalary.Rows.Remove(rowSalary)
                ElseIf rowMasterKaryawan.Primary > rowSalary.Primary_Employee Then
                    TBLSalary.Rows.Remove(rowSalary)
                ElseIf rowMasterKaryawan.Primary < rowSalary.Primary_Employee Then
                    TBLMasterKaryawan.Rows.Remove(rowMasterKaryawan)
                    Me.prg_v180_UpdateSalary.Value += 1
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASalary.Dispose()
            TBLSalary.Dispose()
            DTAMasterKaryawan.Dispose()
            TBLMasterKaryawan.Dispose()
        End Try
    End Sub

    Private Sub Bgw_v180_UpdateSalary_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles Bgw_v180_UpdateSalary.RunWorkerCompleted
        Try
            MsgBox("Update Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdate2018_ItemProdQtyMonthlyClosing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate2018_ItemProdQtyMonthlyClosing.Click

    End Sub
End Class
