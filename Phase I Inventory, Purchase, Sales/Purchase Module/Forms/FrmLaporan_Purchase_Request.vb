Option Explicit On
Option Strict Off
Public Class FrmLaporan_Purchase_Request
    Public UserCode As String = ""
    Public BAckColorAll As Color = Core_Module.GetSMESSBackColorAll
    Private ObjCd As New Core_Module.ClsCekData
    Private ObjPR As New ClsTransaksiPR
    Private ObjUser As New Core_Module.ClsUser
    Private ObjCancel As New ClsTransaksiPR_Cancel
    Private ObjPickupOrder As New ClsTransaksiPR_PickUpOrder
    Private ObjPO As New ClsTransaksiPO

    Dim DTAPickUP As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_order_item2TableAdapter
    Dim TBLPickUp As New DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_order_item2DataTable
    Dim rowPickUp As DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_order_item2Row
    Dim DTACancel As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancel_item1TableAdapter
    Dim TBLCancel As New DseTransaksiPR_Cancel.tbltrans_purchase_req_cancel_item1DataTable
    Dim rowCancel As DseTransaksiPR_Cancel.tbltrans_purchase_req_cancel_item1Row
    Dim buffer As String = ""
    Dim DummyIndex As Integer = -1
    Dim objItem As New Inventory_Module.ClsTransaksiMaterialPlanningGroupItem
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1
    Dim ObjCategory As New Core_Module.ClsMasterCategory
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjSettingProd As New Core_Module.ClsTransaksiSetting_ProdGroup
    Dim EraseRowIndex As Integer = 0
    Dim ObjLicense As New Core_Module.ClsLicense

    '''Product Code
    Private Sub TxtProductCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtProductCode.KeyDown
        Dim ObjProduct As New Core_Module.clsMasterProduct

        Try
            If e.KeyValue = 13 Then
                If Me.TxtProductCode.Text <> "" Then
                    If ObjProduct.ApakahSudahAda(Me.TxtProductCode.Text) = True Then
                        Me.TxtPrimaryProduct.Text = ObjProduct.GetPrimary(Me.TxtProductCode.Text)
                        Me.TxtProductCode.Text = ObjProduct.GetProductCode(Me.TxtPrimaryProduct.Text)
                        Me.TxtDescProduct.Text = ObjProduct.GetDescription(Me.TxtPrimaryProduct.Text)
                        Me.TxtRequestBy.Focus()
                    Else
                        Me.btnFindProduct.PerformClick()
                    End If
                Else
                    Me.btnFindProduct.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtProductCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductCode.Leave
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim hj As New Core_Module.frmCariMasterProduct
        hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.TxtProductCode.Text <> "" Then
                If ObjProduct.ApakahSudahAda(Me.TxtProductCode.Text) = True Then
                    Me.TxtPrimaryProduct.Text = ObjProduct.GetPrimary(Me.TxtProductCode.Text)
                    Me.TxtDescProduct.Text = ObjProduct.GetDescription(Me.TxtPrimaryProduct.Text)
                ElseIf ObjProduct.ApakahSudahAda(Me.TxtProductCode.Text) = False Then
                    If Me.TxtProductCode.Text <> "" Then
                        hj.txtCari.Text = Me.TxtProductCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryProduct > 0 Then
                        Me.TxtPrimaryProduct.Text = hj.PrimaryProduct
                        Me.TxtProductCode.Text = ObjProduct.GetProductCode(Me.TxtPrimaryProduct.Text)
                        Me.TxtDescProduct.Text = ObjProduct.GetDescription(Me.TxtPrimaryProduct.Text)
                    Else
                        Me.TxtPrimaryProduct.Text = "-1"
                        Me.TxtProductCode.Clear()
                        Me.TxtDescProduct.Clear()
                        Me.TxtProductCode.Focus()
                        Me.ChkAllProduct.Checked = False
                    End If
                End If
            Else
                Me.TxtPrimaryProduct.Text = "-1"
                Me.TxtProductCode.Clear()
                Me.TxtDescProduct.Clear()
                Me.ChkAllProduct.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindProduct.Click
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim hj As New Core_Module.frmCariMasterProduct
        Try
            hj.BackColorAll = Me.BAckColorAll
            If Me.TxtProductCode.Text <> "" Then
                hj.txtCari.Text = Me.TxtProductCode.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryProduct > 0 Then
                Me.TxtPrimaryProduct.Text = hj.PrimaryProduct
                Me.TxtProductCode.Text = ObjProduct.GetProductCode(Me.TxtPrimaryProduct.Text)
                Me.TxtDescProduct.Text = ObjProduct.GetDescription(Me.TxtPrimaryProduct.Text)
                Me.TxtRequestBy.Focus()
            Else
                Me.TxtPrimaryProduct.Text = "-1"
                Me.TxtProductCode.Clear()
                Me.TxtDescProduct.Clear()
                Me.TxtProductCode.Focus()
                Me.ChkAllProduct.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    '''RefParty
    Private Sub TxtRequestBy_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRequestBy.KeyDown
        Dim ObjRefParty As New Core_Module.ClsMasterRefParty
        Try
            If e.KeyValue = 13 Then
                If Me.TxtRequestBy.Text <> "" Then
                    If ObjRefParty.ApakahSudahAda(Me.TxtRequestBy.Text) = True Then
                        Me.TxtPrimaryRefParty.Text = ObjRefParty.GetPrimary(Me.TxtRequestBy.Text)
                        Me.TxtDescRefParty.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                        Me.TxtPRNo.Focus()
                    Else
                        Me.BtnFindRefParty.PerformClick()
                    End If
                Else
                    Me.BtnFindRefParty.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRequestBy_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRequestBy.Leave
        Dim ObjRefParty As New Core_Module.ClsMasterRefParty
        Dim hj As New Core_Module.FrmMasterRefParty_Cari
        hj.Aktif = True
        hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.TxtRequestBy.Text <> "" Then
                If ObjRefParty.ApakahSudahAda(Me.TxtRequestBy.Text) = True Then
                    Me.TxtPrimaryRefParty.Text = ObjRefParty.GetPrimary(Me.TxtRequestBy.Text)
                    Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                    Me.TxtDescRefParty.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                    Me.TxtPRNo.Focus()
                ElseIf ObjRefParty.ApakahSudahAda(Me.TxtRequestBy.Text) = False Then
                    If Me.TxtRequestBy.Text <> "" Then
                        hj.txtCari.Text = Me.TxtRequestBy.Text
                    End If

                    hj.ShowDialog()
                    If hj.PrimaryRefParty > 0 Then
                        Me.TxtPrimaryRefParty.Text = hj.PrimaryRefParty
                        Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                        Me.TxtDescRefParty.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                        Me.TxtPRNo.Focus()
                    Else
                        Me.TxtPrimaryRefParty.Text = "-1"
                        Me.TxtRequestBy.Clear()
                        Me.TxtDescRefParty.Clear()
                        Me.TxtRequestBy.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryRefParty.Text = "-1"
                Me.TxtRequestBy.Clear()
                Me.TxtDescRefParty.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindRefParty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindRefParty.Click
        Dim ObjRefParty As New Core_Module.ClsMasterRefParty
        Dim hj As New Core_Module.FrmMasterRefParty_Cari
        hj.Aktif = True
        Try
            hj.BackColorAll = Me.BAckColorAll
            If Me.TxtRequestBy.Text <> "" Then
                hj.txtCari.Text = Me.TxtRequestBy.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryRefParty > 0 Then
                Me.TxtPrimaryRefParty.Text = hj.PrimaryRefParty
                Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                Me.TxtDescRefParty.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                Me.TxtPRNo.Focus()
            Else
                Me.TxtPrimaryRefParty.Text = "-1"
                Me.TxtRequestBy.Clear()
                Me.TxtDescRefParty.Clear()
                Me.TxtRequestBy.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    '''Got Lost Focus
    Private Sub TxtRequestBy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRequestBy.GotFocus, TxtProductCode.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Txt.Enabled = True Then
                    Txt.BackColor = Me.BAckColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRequestBy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRequestBy.LostFocus, TxtProductCode.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Txt.Enabled = True Then
                    Txt.BackColor = Me.TxtPrimaryProduct.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '''PR
    Private Sub TxtPRNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPRNo.KeyDown
        Dim ObjPR As New ClsTransaksiPR
        Try
            If e.KeyValue = 13 Then
                If Me.TxtPRNo.Text <> "" Then
                    If ObjPR.ApakahSudahAda(Me.TxtPRNo.Text) = True Then
                        Me.TxtPrimaryPR.Text = ObjPR.GetPrimary(Me.TxtPRNo.Text)
                        Me.CboPilih.Focus()
                    Else
                        Me.btnFindPR.PerformClick()
                    End If
                Else
                    Me.btnFindPR.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPRNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPRNo.Leave
        Dim ObjPR As New ClsTransaksiPR
        Dim hj As New FrmTransaksiPR_Cari
        hj.BackColorAll = Me.BAckColorAll

        Try
            If Me.TxtPRNo.Text <> "" Then
                If ObjPR.ApakahSudahAda(Me.TxtPRNo.Text) = True Then
                    Me.TxtPrimaryPR.Text = ObjPR.GetPrimary(Me.TxtPRNo.Text)
                    Me.TxtPRNo.Text = ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                    Me.CboPilih.Focus()
                ElseIf ObjPR.ApakahSudahAda(Me.TxtPRNo.Text) = False Then
                    If Me.TxtPRNo.Text <> "" Then
                        hj.txtCari.Text = Me.TxtPRNo.Text
                    End If

                    hj.ShowDialog()
                    If hj.PrimaryPR > 0 Then
                        Me.TxtPrimaryPR.Text = hj.PrimaryPR
                        Me.TxtPRNo.Text = ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                        Me.CboPilih.Focus()
                    Else
                        Me.TxtPrimaryPR.Text = "-1"
                        Me.TxtPRNo.Clear()
                        Me.TxtPRNo.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryPR.Text = "-1"
                Me.TxtPRNo.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnFindPR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPR.Click
        Dim ObjPR As New ClsTransaksiPR
        Dim hj As New FrmTransaksiPR_Cari

        Try
            hj.BackColorAll = Me.BAckColorAll
            If Me.TxtPRNo.Text <> "" Then
                hj.txtCari.Text = Me.TxtPRNo.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryPR > 0 Then
                Me.TxtPrimaryPR.Text = hj.PrimaryPR
                Me.TxtPRNo.Text = ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                Me.CboPilih.Focus()
            Else
                Me.TxtPrimaryPR.Text = "-1"
                Me.TxtPRNo.Clear()
                Me.TxtPRNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    '''Got Lost Focus
    Private Sub TxtPRNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPRNo.GotFocus, TxtProductCode.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Txt.Enabled = True Then
                    Txt.BackColor = Me.BAckColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPRNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPRNo.LostFocus, TxtProductCode.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Txt.Enabled = True Then
                    Txt.BackColor = Me.TxtPrimaryProduct.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.DtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.DtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.DtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.DtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.DtpFrom.Value.Date > Me.DtpTo.Value.Date Then
                MsgBox("Starting period must be smaller than ending period", MsgBoxStyle.Information, "Information")
                Me.DtpFrom.Focus()
                Exit Sub
            Else
                Me.Cetakan(Me.DtpFrom.Value, Me.DtpTo.Value, Me.TxtPrimaryRefParty.Text, Me.TxtPrimaryProduct.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporan_Purchase_Request_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting_ProdGroup

        Try
            ObjSetting.SimpanData(Me.UserCode, Core_Module.ClsTransaksiSetting_ProdGroup.EnumTipeTransaksi.LaporanPurchaseRequest, _
            Me.dgvItemProductGroup, Me.ColProductGroup_No.Index, _
            Me.ColProductGroup_Level.Index, Me.ColProductGroup_CodeNo.Index, Me.ColProductGroup_Desc.Index, _
            Me.ColProductGroup_Remarks.Index, Me.ColProductGroup_PrimaryAcuan.Index)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSetting = Nothing
        End Try
    End Sub

    Private Sub FrmLaporan_Purchase_Request_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.WindowState = FormWindowState.Maximized
        Try
            Me.CboPilih.SelectedIndex = 0
            Me.DtpFrom.Select()
            Me.DtpFrom.Value = Now
            Me.DtpTo.Value = Now
            Me.dgvItemProductGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.IsiDGV_ItemGroup()
            Me.chkShowEstimateDelivery.Checked = True
            Me.chkShowEstimateDelivery.Visible = False
            Me.chkShowEstimateDelivery.TabStop = False
            Me.chkShowPONumber.Checked = True
            Me.chkShowPONumber.Visible = False
            Me.chkShowPONumber.TabStop = False
            Me.dgvItemProductGroup.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkAllProduct_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkAllProduct.CheckedChanged
        Try
            If Me.TxtProductCode.Text = "" Then
                Me.ChkAllProduct.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkAllProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkAllProduct.Click
        Try
            If Me.TxtProductCode.Text = "" Then
                MsgBox("Product code cannot empty string", MsgBoxStyle.Information, "Information")
                Me.TxtProductCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    'Private Sub TxtProductCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductCode.GotFocus
    '    Dim Txt As TextBox
    '    Try
    '        Txt = TryCast(sender, TextBox)
    '        If IsNothing(Txt) = False Then
    '            If Me.ChkAllProduct.Checked = True Then
    '                Txt.BackColor = Me.TxtDescProduct.BackColor
    '            Else
    '                Txt.BackColor = Me.BAckColorAll
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub TxtProductCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductCode.LostFocus
    '    Dim Txt As TextBox
    '    Try
    '        Txt = TryCast(sender, TextBox)
    '        If IsNothing(Txt) = False Then
    '            If Me.ChkAllProduct.Checked = True Then
    '                Txt.BackColor = Me.TxtDescProduct.BackColor
    '            Else
    '                Txt.BackColor = Me.TxtPrimaryProduct.BackColor
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub ChkAllProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkAllProduct.Click
    '    Try
    '        If Me.ChkAllProduct.Checked = True Then
    '            Me.TxtProductCode.ReadOnly = True
    '            Me.TxtProductCode.BackColor = Me.TxtDescProduct.BackColor
    '            Me.btnFindProduct.Enabled = False
    '            Me.TxtProductCode.TabStop = False
    '        Else
    '            Me.TxtProductCode.ReadOnly = False
    '            Me.TxtProductCode.BackColor = Me.TxtPrimaryProduct.BackColor
    '            Me.btnFindProduct.Enabled = True
    '            Me.TxtProductCode.TabStop = True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Function GetJmlPickUp(ByVal vPrimaryPR As Double, ByVal vPrimaryCodeNo As Double) As Decimal
        Try
            Me.TBLPickUp.DefaultView.RowFilter = "[primary pr] = " + vPrimaryPR.ToString + " AND [primary code no] = " + vPrimaryCodeNo.ToString
            If Me.TBLPickUp.DefaultView.Count = 0 Then
                Return 0
            Else
                Me.rowPickUp = TBLPickUp.DefaultView(0).Row
                Return Me.rowPickUp.qty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlCancel(ByVal vPrimaryPR As Double, ByVal vPrimaryCodeNo As Double) As Decimal
        Try
            Me.TBLCancel.DefaultView.RowFilter = "[primary pr] = " + vPrimaryPR.ToString + " AND [primary code no] = " + vPrimaryCodeNo.ToString
            If Me.TBLCancel.DefaultView.Count = 0 Then
                Return 0
            Else
                Me.rowCancel = TBLCancel.DefaultView(0).Row
                Return Me.rowCancel.qty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Public Sub Cetakan(ByVal StartDate As Date, ByVal EndDate As Date, ByVal PrimaryRefParty As Double, ByVal PrimaryProduct As Double)
        Dim ReportDetail As New CrpPurchase_Request_Report
        Dim DS As New DsePurchaseRequestReport
        Dim DTACetak_Item As New DsePurchaseRequestReportTableAdapters.TblCetak_ItemTableAdapter
        Dim TBLCetak_Item As DsePurchaseRequestReport.TblCetak_ItemDataTable = DS.TblCetak_Item
        Dim TBLCetak_ItemD As New DsePurchaseRequestReport.TblCetak_ItemDataTable
        Dim RowCetak_Item, RowCetak_ItemD As DsePurchaseRequestReport.TblCetak_ItemRow
        Dim TBLItem As DsePurchaseRequestReport.TBLItemDataTable = DS.TBLItem
        Dim ROWItem As DsePurchaseRequestReport.TBLItemRow

        Dim paramStartDate As Date
        Dim paramEndDate As Date
        Dim PrimaryTransaksi As Double = -1
        Dim index As Integer = 0
        Dim Filter As String = ""
        Dim StatusPR As ClsTransaksiPR.EnumReadyForPurchase = ClsTransaksiPR.EnumReadyForPurchase.Yes

        Dim DV1 As DataRowView
        Dim DV2 As DataRowView
        Dim ListPrimary As New List(Of Double)
        Dim ObjSF As New Inventory_Module.ClsTransaksiShippingForecast
        Dim vTBLItemDelivery As New Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable
        Dim vRowItemDelivery, vRowItemDelivery1 As Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim vTBLItemPOB As New Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable
        Dim vTBLItemStock As New Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable
        Dim DTACetak_Delivery As New DsePurchaseRequestReportTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Dim TBLCetak_Delivery As New DsePurchaseRequestReport.tbltrans_purchase_req_itemdeliveryDataTable
        Dim rowCetak_Delivery As DsePurchaseRequestReport.tbltrans_purchase_req_itemdeliveryRow
        Dim TipeProductGroup As Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            TBLItem.Rows.Clear()
            TBLCetak_Item.Rows.Clear()
            TBLCetak_ItemD.Rows.Clear()

            DTACetak_Item.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTACancel.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTAPickUP.Connection.ConnectionString = Core_Module.GetConstring
            DTACetak_Delivery.Connection.ConnectionString = Core_Module.GetConstring
            'fill data
            paramStartDate = DtpFrom.Value.Date
            paramEndDate = DtpTo.Value.Date

            'Me.DTACancel.Fill(Me.TBLCancel, paramEndDate, paramStartDate - TimeSpan.FromDays(31))
            'Me.DTAPickUP.Fill(Me.TBLPickUp, paramEndDate, paramStartDate - TimeSpan.FromDays(31))

            If Me.dgvItemProductGroup.RowCount > 1 Then
                'ada filter berdasarkan product group
                If Me.TxtPrimaryPR.Text <> "" And Me.TxtPrimaryPR.Text <> "-1" Then
                    DTACetak_Item.FillBy_PrimaryPR(TBLCetak_ItemD, paramEndDate, paramStartDate, Me.TxtPrimaryPR.Text)
                ElseIf Me.TxtPrimaryRefParty.Text <> "" And Me.TxtPrimaryRefParty.Text <> "-1" Then
                    DTACetak_Item.FillBy_PrimaryRefParty(TBLCetak_ItemD, paramEndDate, paramStartDate, Me.TxtPrimaryRefParty.Text)
                ElseIf Me.TxtPrimaryProduct.Text <> "" And Me.TxtPrimaryProduct.Text <> "-1" Then
                    DTACetak_Item.FillBy_PrimaryCodeNo(TBLCetak_ItemD, paramEndDate, paramStartDate, Me.TxtPrimaryProduct.Text)
                Else
                    DTACetak_Item.Fill(TBLCetak_ItemD, paramEndDate, paramStartDate)
                End If

                'product group, dkk
                For Each rowDGV As DataGridViewRow In Me.dgvItemProductGroup.Rows
                    If rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
                        TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)

                        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                            TBLCetak_ItemD.DefaultView.RowFilter = "[Primary Product Group] = " + rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                            TBLCetak_ItemD.DefaultView.RowFilter = "[Primary Brand] = " + rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                            TBLCetak_ItemD.DefaultView.RowFilter = "[Primary Sub Brand] = " + rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                            TBLCetak_ItemD.DefaultView.RowFilter = "[Primary Category] = " + rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        Else
                            TBLCetak_ItemD.DefaultView.RowFilter = "[primary code no] = " + rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        End If

                        While TBLCetak_ItemD.DefaultView.Count > 0
                            RowCetak_ItemD = TBLCetak_ItemD.DefaultView(0).Row
                            RowCetak_Item = TBLCetak_Item.NewTblCetak_ItemRow()
                            For ijk As Integer = 0 To TBLCetak_Item.Columns.Count - 1
                                RowCetak_Item.Item(ijk) = RowCetak_ItemD.Item(ijk)
                            Next
                            TBLCetak_Item.Rows.Add(RowCetak_Item)
                            TBLCetak_ItemD.Rows.Remove(RowCetak_ItemD)
                        End While
                    End If
                Next
            Else
                'tidak ada filter product group
                If Me.TxtPrimaryPR.Text <> "" And Me.TxtPrimaryPR.Text <> "-1" Then
                    DTACetak_Item.FillBy_PrimaryPR(TBLCetak_Item, paramEndDate, paramStartDate, Me.TxtPrimaryPR.Text)
                ElseIf Me.TxtPrimaryRefParty.Text <> "" And Me.TxtPrimaryRefParty.Text <> "-1" Then
                    DTACetak_Item.FillBy_PrimaryRefParty(TBLCetak_Item, paramEndDate, paramStartDate, Me.TxtPrimaryRefParty.Text)
                ElseIf Me.TxtPrimaryProduct.Text <> "" And Me.TxtPrimaryProduct.Text <> "-1" Then
                    DTACetak_Item.FillBy_PrimaryCodeNo(TBLCetak_Item, paramEndDate, paramStartDate, Me.TxtPrimaryProduct.Text)
                Else
                    DTACetak_Item.Fill(TBLCetak_Item, paramEndDate, paramStartDate)
                End If
            End If


            DTACetak_Delivery.Fill(TBLCetak_Delivery, paramEndDate, paramStartDate)

            'bentuk filter

            'filtering data
            '0. primary pr, sudah di atas
            '1. primary ref party -> just in case double filter
            TBLCetak_Item.DefaultView.RowFilter = ""
            If Me.TxtPrimaryRefParty.Text <> "-1" And Me.TxtPrimaryRefParty.Text <> "" Then
                Filter = "[primary ref party] = " + Me.TxtPrimaryRefParty.Text
            End If

            '2. kode product -> just in case double filter
            If Me.TxtPrimaryProduct.Text <> "-1" And Me.TxtPrimaryProduct.Text <> "" Then
                If Filter = "" Then
                    Filter = "[primary code no] = " + Me.TxtPrimaryProduct.Text
                Else
                    Filter = Filter + " AND [primary code no] = " + Me.TxtPrimaryProduct.Text
                End If
            End If

            '3. status pr 
            If Me.CboPilih.SelectedIndex = 1 Then
                'ready for purchase only
                StatusPR = ClsTransaksiPR.EnumReadyForPurchase.Yes
                If Filter = "" Then
                    Filter = "[ready for purchase] = " + CInt(StatusPR).ToString
                Else
                    Filter = Filter + " AND [ready for purchase] = " + CInt(StatusPR).ToString
                End If
            ElseIf Me.CboPilih.SelectedIndex = 2 Then
                StatusPR = ClsTransaksiPR.EnumReadyForPurchase.No
                If Filter = "" Then
                    Filter = "[ready for purchase] = " + CInt(StatusPR).ToString
                Else
                    Filter = Filter + " AND [ready for purchase] = " + CInt(StatusPR).ToString
                End If
            End If


            If Me.ChkOB.Checked = True Then
                If Filter = "" Then
                    Filter = "OB > 0"
                Else
                    Filter = Filter + " AND OB > 0"
                End If
            End If


            Dim FilterUser As String = ""
            'Created By
            If Me.DgvUser.Rows.Count > 0 Then
                For Each rowDGV1 As DataGridViewRow In Me.DgvUser.Rows
                    If Not IsDBNull(rowDGV1.Cells(Me.ColCreatedBy.Index).Value) AndAlso rowDGV1.Cells(Me.ColCreatedBy.Index).Value <> "" AndAlso Not String.IsNullOrEmpty(rowDGV1.Cells(Me.ColCreatedBy.Index).Value) Then
                        If FilterUser = "" Then
                            FilterUser &= " [created by] = '" & rowDGV1.Cells(Me.ColCreatedBy.Index).Value.ToString & "'"
                        Else
                            FilterUser &= " Or [created by] = '" & rowDGV1.Cells(Me.ColCreatedBy.Index).Value.ToString & "'"
                        End If
                    End If
                Next
            End If


            If Filter = "" Then
                If FilterUser <> "" Then
                    Filter &= "(" & FilterUser & ") "
                End If
            Else
                If FilterUser <> "" Then
                    Filter &= " AND (" & FilterUser & ") "
                End If
            End If
           


            'start filter, hanya pindahkan barang yg sesuai kriteria
            TBLCetak_Item.DefaultView.RowFilter = Filter
            TBLCetak_Item.DefaultView.Sort = "tanggal, [no transaksi]"
            If TBLCetak_Item.DefaultView.Count > 0 Then
                For Each DV2 In TBLCetak_Item.DefaultView
                    RowCetak_Item = DV2.Row
                    '''isi item
                    With RowCetak_Item
                        ROWItem = TBLItem.NewTBLItemRow
                        ROWItem.identifier = .identifier
                        ROWItem.Primary_Transaksi = .primary_transaksi
                        ROWItem.Primary_Product = .primary_code_no
                        ROWItem.Product_Code = .Product_Code
                        ROWItem.Description = .Description
                        ROWItem._Decimal = ._Decimal
                        ROWItem.Ordered = ObjCd.DigitGrouping(.ordered, ROWItem._Decimal)
                        ROWItem.PickUp = ObjCd.DigitGrouping(.qty_pick_up, ROWItem._Decimal)
                        ROWItem.Cancel = ObjCd.DigitGrouping(.qty_cancel, ROWItem._Decimal)
                        'ROWItem.PickUp = ObjCd.DigitGrouping(ObjPR.GetQtyPickUp(ROWItem.Primary_Transaksi, ROWItem.Primary_Product), ROWItem._Decimal)
                        'ROWItem.Cancel = ObjCd.DigitGrouping(ObjPR.GetQtyCancel(ROWItem.Primary_Transaksi, ROWItem.Primary_Product), ROWItem._Decimal)
                        'ROWItem.PickUp = ObjCd.DigitGrouping(Me.GetJmlPickUp(ROWItem.Primary_Transaksi, ROWItem.Primary_Product), ROWItem._Decimal)
                        'ROWItem.Cancel = ObjCd.DigitGrouping(Me.GetJmlCancel(ROWItem.Primary_Transaksi, ROWItem.Primary_Product), ROWItem._Decimal)
                        ROWItem.OB = ObjCd.DigitGrouping(.OB, ROWItem._Decimal)
                        ROWItem.Unit_Code = .Unit_Code
                        If .status = ClsTransaksiPR.EnumStatus.PR_Pending Then
                            ROWItem.Status = "PR Pending"
                        ElseIf .status = ClsTransaksiPR.EnumStatus.Purchase_On_Progress Then
                            ROWItem.Status = "On Progress"
                        Else
                            ROWItem.Status = "Closed"
                        End If
                        ROWItem.Remarks = .remarks

                        ROWItem.Request_By = .ref_party_code + " - " + .ref_party_name
                        ROWItem.Remarks_Header = .remarks_header
                        If .Kode_User = "" Then
                            ROWItem.User_Authorize = ""
                        Else
                            ROWItem.User_Authorize = .Kode_User + " on " + Format(.time_authorize_by, "dd MMM yyyy HH:mm")
                        End If
                        ROWItem._Date = Format(.tanggal, "dd MMM yyyy")
                        ROWItem.PR_No = .no_transaksi

                        TBLItem.AddTBLItemRow(ROWItem)

                        If RowCetak_Item.delivery_date_cont = ClsTransaksiPR.EnumDeliveryDate.Cheklist Then
                            If .delivery_date = #1/1/1950# Then
                                TBLCetak_Delivery.DefaultView.RowFilter = "[identifier] = '" & RowCetak_Item.identifier.ToString & "'"
                                TBLCetak_Delivery.DefaultView.Sort = "[Delivery Date] ASC"
                                rowCetak_Delivery = TBLCetak_Delivery.DefaultView(TBLCetak_Delivery.DefaultView.Count - 1).Row
                                ROWItem.Delivery_Date = Format(rowCetak_Delivery.delivery_date, "dd MMM yyyy")
                            Else
                                ROWItem.Delivery_Date += Format(RowCetak_Item.delivery_date, "dd MMM yyyy")
                            End If
                        Else
                            ROWItem.Delivery_Date += "-"
                        End If
                    End With
                Next
            End If

            'isikan availability
            '0. isikan list primary
            ListPrimary.Clear()
            For Each ROWItem In TBLItem.Rows
                If ListPrimary.Contains(ROWItem.Primary_Product) = False Then
                    ListPrimary.Add(ROWItem.Primary_Product)
                End If
            Next
            '1. hitung shipping forecast untuk barang yg ditampilkan saja
            If TBLItem.Rows.Count > 0 Then
                ObjSF.HitungShippingForecast(vTBLItemDelivery, vTBLItemPOB, vTBLItemStock, -1, _
                Core_Module.GetSMESSBackColorAll, -1, ListPrimary)
            End If
            '2. untuk goods, isikan datanya apa adanya berdasarkan shipping forecast
            For Each ROWItem In TBLItem.Rows
                vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + ROWItem.identifier + _
                "' AND tipe = " + CInt(Inventory_Module.ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req).ToString
                If ROWItem.Status = "Closed" Then
                    ROWItem.availability = "-"
                Else
                    ROWItem.availability = "Waiting"
                    If vTBLItemDelivery.DefaultView.Count > 0 Then
                        vRowItemDelivery = vTBLItemDelivery.DefaultView(0).Row
                        If vTBLItemDelivery.DefaultView.Count > 1 Then
                            For ij As Integer = 1 To vTBLItemDelivery.DefaultView.Count - 1
                                vRowItemDelivery1 = vTBLItemDelivery.DefaultView(ij).Row
                                vRowItemDelivery.qty_sob += vRowItemDelivery1.qty_sob
                                vRowItemDelivery.qty_available += vRowItemDelivery1.qty_available
                                vRowItemDelivery.qty_ready_stock += vRowItemDelivery1.qty_ready_stock
                                If vRowItemDelivery.Isprimary_po_refNull = True Then
                                    vRowItemDelivery.primary_po_ref = ""
                                End If
                                If vRowItemDelivery1.Isprimary_po_refNull = True Then
                                    vRowItemDelivery1.primary_po_ref = ""
                                End If
                                vRowItemDelivery.primary_po_ref = vRowItemDelivery.primary_po_ref + "-" + vRowItemDelivery1.primary_po_ref
                            Next
                        End If
                        If vRowItemDelivery.qty_available > 0 Then
                            If vRowItemDelivery.qty_sob <= vRowItemDelivery.qty_available Then
                                If vRowItemDelivery.qty_ready_stock >= vRowItemDelivery.qty_available Then
                                    ROWItem.availability = "Ready"
                                ElseIf vRowItemDelivery.qty_ready_stock > 0 Then
                                    ROWItem.availability = "Partial"
                                Else
                                    ROWItem.availability = "Waiting"
                                End If
                            Else
                                If vRowItemDelivery.qty_ready_stock >= vRowItemDelivery.qty_available Then
                                    ROWItem.availability = "Partial"
                                ElseIf vRowItemDelivery.qty_ready_stock > 0 Then
                                    ROWItem.availability = "Partial"
                                Else
                                    ROWItem.availability = "Waiting"
                                End If
                            End If
                        Else
                            ROWItem.availability = "n/a"
                        End If
                        If vRowItemDelivery.Isprimary_po_refNull = False Then
                            If vRowItemDelivery.primary_po_ref <> "" Then
                                ROWItem.availability = Me.Translate_RefPO(vRowItemDelivery.primary_po_ref, vRowItemDelivery, ROWItem.availability, vRowItemDelivery._Decimal)
                            End If
                        End If
                    End If
                End If
            Next

            '3. untuk service, hanya tampilkan data berdasarkan POB

            ReportDetail.SetDataSource(DS)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_StartDate.ParameterFieldName, paramStartDate)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_EndDate.ParameterFieldName, paramEndDate)
            Me.crv.ReportSource = ReportDetail
            Me.crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTACetak_Item.Dispose()
            TBLCetak_ItemD.Dispose()
            TBLCetak_Item.Dispose()
            TBLItem.Dispose()
            ListPrimary = Nothing
            ObjSF = Nothing
            vTBLItemDelivery.Dispose()
            vTBLItemPOB.Dispose()
            vTBLItemStock.Dispose()
            DTACetak_Delivery.Dispose()
            TBLCetak_Delivery.Dispose()
        End Try
    End Sub

    'Public Sub Cetakan(ByVal StartDate As Date, ByVal EndDate As Date, ByVal PrimaryRefParty As Double, ByVal PrimaryProduct As Double)
    '    Dim ReportDetail As New CrpPurchase_Request_Report
    '    Dim DS As New DsePurchaseRequestReport
    '    Dim DTACetak_Item As New DsePurchaseRequestReportTableAdapters.TblCetak_ItemTableAdapter
    '    Dim TBLCetak_Item As DsePurchaseRequestReport.TblCetak_ItemDataTable = DS.TblCetak_Item
    '    Dim TBLCetak_ItemD As New DsePurchaseRequestReport.TblCetak_ItemDataTable
    '    Dim RowCetak_Item, RowCetak_ItemD As DsePurchaseRequestReport.TblCetak_ItemRow
    '    Dim TBLItem As DsePurchaseRequestReport.TBLItemDataTable = DS.TBLItem
    '    Dim ROWItem As DsePurchaseRequestReport.TBLItemRow

    '    Dim paramStartDate As Date
    '    Dim paramEndDate As Date
    '    Dim PrimaryTransaksi As Double = -1
    '    Dim index As Integer = 0
    '    Dim Filter As String = ""
    '    Dim StatusPR As ClsTransaksiPR.EnumReadyForPurchase = ClsTransaksiPR.EnumReadyForPurchase.Yes

    '    Dim DV1 As DataRowView
    '    Dim DV2 As DataRowView
    '    Dim ListPrimary As New List(Of Double)
    '    Dim ObjSF As New Inventory_Module.ClsTransaksiShippingForecast
    '    Dim vTBLItemDelivery As New Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable
    '    Dim vRowItemDelivery, vRowItemDelivery1 As Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
    '    Dim vTBLItemPOB As New Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable
    '    Dim vTBLItemStock As New Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable
    '    Dim DTACetak_Delivery As New DsePurchaseRequestReportTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
    '    Dim TBLCetak_Delivery As New DsePurchaseRequestReport.tbltrans_purchase_req_itemdeliveryDataTable
    '    Dim rowCetak_Delivery As DsePurchaseRequestReport.tbltrans_purchase_req_itemdeliveryRow
    '    Dim TipeProductGroup As Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup

    '    Try
    '        Windows.Forms.Cursor.Current = Cursors.WaitCursor

    '        TBLItem.Rows.Clear()
    '        TBLCetak_Item.Rows.Clear()
    '        TBLCetak_ItemD.Rows.Clear()

    '        DTACetak_Item.Connection.ConnectionString = Core_Module.GetConstring
    '        Me.DTACancel.Connection.ConnectionString = Core_Module.GetConstring
    '        Me.DTAPickUP.Connection.ConnectionString = Core_Module.GetConstring
    '        DTACetak_Delivery.Connection.ConnectionString = Core_Module.GetConstring
    '        'fill data
    '        paramStartDate = DtpFrom.Value.Date
    '        paramEndDate = DtpTo.Value.Date

    '        Me.DTACancel.Fill(Me.TBLCancel, paramEndDate, paramStartDate - TimeSpan.FromDays(31))
    '        Me.DTAPickUP.Fill(Me.TBLPickUp, paramEndDate, paramStartDate - TimeSpan.FromDays(31))

    '        If Me.dgvItemProductGroup.RowCount > 1 Then
    '            'ada filter berdasarkan product group
    '            If Me.TxtPrimaryPR.Text <> "" And Me.TxtPrimaryPR.Text <> "-1" Then
    '                DTACetak_Item.FillBy_PrimaryPR(TBLCetak_ItemD, paramEndDate, paramStartDate, Me.TxtPrimaryPR.Text)
    '            ElseIf Me.TxtPrimaryRefParty.Text <> "" And Me.TxtPrimaryRefParty.Text <> "-1" Then
    '                DTACetak_Item.FillBy_PrimaryRefParty(TBLCetak_ItemD, paramEndDate, paramStartDate, Me.TxtPrimaryRefParty.Text)
    '            ElseIf Me.TxtPrimaryProduct.Text <> "" And Me.TxtPrimaryProduct.Text <> "-1" Then
    '                DTACetak_Item.FillBy_PrimaryCodeNo(TBLCetak_ItemD, paramEndDate, paramStartDate, Me.TxtPrimaryProduct.Text)
    '            Else
    '                DTACetak_Item.Fill(TBLCetak_ItemD, paramEndDate, paramStartDate)
    '            End If

    '            'product group, dkk
    '            For Each rowDGV As DataGridViewRow In Me.dgvItemProductGroup.Rows
    '                If rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
    '                    TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)

    '                    If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
    '                        TBLCetak_ItemD.DefaultView.RowFilter = "[Primary Product Group] = " + rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
    '                    ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
    '                        TBLCetak_ItemD.DefaultView.RowFilter = "[Primary Brand] = " + rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
    '                    ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
    '                        TBLCetak_ItemD.DefaultView.RowFilter = "[Primary Sub Brand] = " + rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
    '                    ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
    '                        TBLCetak_ItemD.DefaultView.RowFilter = "[Primary Category] = " + rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
    '                    Else
    '                        TBLCetak_ItemD.DefaultView.RowFilter = "[primary code no] = " + rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
    '                    End If

    '                    While TBLCetak_ItemD.DefaultView.Count > 0
    '                        RowCetak_ItemD = TBLCetak_ItemD.DefaultView(0).Row
    '                        RowCetak_Item = TBLCetak_Item.NewTblCetak_ItemRow()
    '                        For ijk As Integer = 0 To TBLCetak_Item.Columns.Count - 1
    '                            RowCetak_Item.Item(ijk) = RowCetak_ItemD.Item(ijk)
    '                        Next
    '                        TBLCetak_Item.Rows.Add(RowCetak_Item)
    '                        TBLCetak_ItemD.Rows.Remove(RowCetak_ItemD)
    '                    End While
    '                End If
    '            Next
    '        Else
    '            'tidak ada filter product group
    '            If Me.TxtPrimaryPR.Text <> "" And Me.TxtPrimaryPR.Text <> "-1" Then
    '                DTACetak_Item.FillBy_PrimaryPR(TBLCetak_Item, paramEndDate, paramStartDate, Me.TxtPrimaryPR.Text)
    '            ElseIf Me.TxtPrimaryRefParty.Text <> "" And Me.TxtPrimaryRefParty.Text <> "-1" Then
    '                DTACetak_Item.FillBy_PrimaryRefParty(TBLCetak_Item, paramEndDate, paramStartDate, Me.TxtPrimaryRefParty.Text)
    '            ElseIf Me.TxtPrimaryProduct.Text <> "" And Me.TxtPrimaryProduct.Text <> "-1" Then
    '                DTACetak_Item.FillBy_PrimaryCodeNo(TBLCetak_Item, paramEndDate, paramStartDate, Me.TxtPrimaryProduct.Text)
    '            Else
    '                DTACetak_Item.Fill(TBLCetak_Item, paramEndDate, paramStartDate)
    '            End If
    '        End If


    '        DTACetak_Delivery.Fill(TBLCetak_Delivery, paramEndDate, paramStartDate)

    '        'bentuk filter

    '        'filtering data
    '        '0. primary pr, sudah di atas
    '        '1. primary ref party -> just in case double filter
    '        TBLCetak_Item.DefaultView.RowFilter = ""
    '        If Me.TxtPrimaryRefParty.Text <> "-1" And Me.TxtPrimaryRefParty.Text <> "" Then
    '            Filter = "[primary ref party] = " + Me.TxtPrimaryRefParty.Text
    '        End If

    '        '2. kode product -> just in case double filter
    '        If Me.TxtPrimaryProduct.Text <> "-1" And Me.TxtPrimaryProduct.Text <> "" Then
    '            If Filter = "" Then
    '                Filter = "[primary code no] = " + Me.TxtPrimaryProduct.Text
    '            Else
    '                Filter = Filter + " AND [primary code no] = " + Me.TxtPrimaryProduct.Text
    '            End If
    '        End If

    '        '3. status pr 
    '        If Me.CboPilih.SelectedIndex = 1 Then
    '            'ready for purchase only
    '            StatusPR = ClsTransaksiPR.EnumReadyForPurchase.Yes
    '            If Filter = "" Then
    '                Filter = "[ready for purchase] = " + CInt(StatusPR).ToString
    '            Else
    '                Filter = Filter + " AND [ready for purchase] = " + CInt(StatusPR).ToString
    '            End If
    '        ElseIf Me.CboPilih.SelectedIndex = 2 Then
    '            StatusPR = ClsTransaksiPR.EnumReadyForPurchase.No
    '            If Filter = "" Then
    '                Filter = "[ready for purchase] = " + CInt(StatusPR).ToString
    '            Else
    '                Filter = Filter + " AND [ready for purchase] = " + CInt(StatusPR).ToString
    '            End If
    '        End If


    '        If Me.ChkOB.Checked = True Then
    '            If Filter = "" Then
    '                Filter = "OB > 0"
    '            Else
    '                Filter = Filter + " AND OB > 0"
    '            End If
    '        End If


    '        'start filter, hanya pindahkan barang yg sesuai kriteria
    '        TBLCetak_Item.DefaultView.RowFilter = Filter
    '        TBLCetak_Item.DefaultView.Sort = "tanggal, [no transaksi]"
    '        If TBLCetak_Item.DefaultView.Count > 0 Then
    '            For Each DV2 In TBLCetak_Item.DefaultView
    '                RowCetak_Item = DV2.Row
    '                '''isi item
    '                With RowCetak_Item
    '                    ROWItem = TBLItem.NewTBLItemRow
    '                    ROWItem.identifier = .identifier
    '                    ROWItem.Primary_Transaksi = .primary_transaksi
    '                    ROWItem.Primary_Product = .primary_code_no
    '                    ROWItem.Product_Code = .Product_Code
    '                    ROWItem.Description = .Description
    '                    ROWItem._Decimal = ._Decimal
    '                    ROWItem.Ordered = ObjCd.DigitGrouping(.ordered, ROWItem._Decimal)
    '                    ROWItem.PickUp = ObjCd.DigitGrouping(.qty_pick_up, ROWItem._Decimal)
    '                    ROWItem.Cancel = ObjCd.DigitGrouping(.qty_cancel, ROWItem._Decimal)
    '                    'ROWItem.PickUp = ObjCd.DigitGrouping(ObjPR.GetQtyPickUp(ROWItem.Primary_Transaksi, ROWItem.Primary_Product), ROWItem._Decimal)
    '                    'ROWItem.Cancel = ObjCd.DigitGrouping(ObjPR.GetQtyCancel(ROWItem.Primary_Transaksi, ROWItem.Primary_Product), ROWItem._Decimal)
    '                    'ROWItem.PickUp = ObjCd.DigitGrouping(Me.GetJmlPickUp(ROWItem.Primary_Transaksi, ROWItem.Primary_Product), ROWItem._Decimal)
    '                    'ROWItem.Cancel = ObjCd.DigitGrouping(Me.GetJmlCancel(ROWItem.Primary_Transaksi, ROWItem.Primary_Product), ROWItem._Decimal)
    '                    ROWItem.OB = ObjCd.DigitGrouping(.OB, ROWItem._Decimal)
    '                    ROWItem.Unit_Code = .Unit_Code
    '                    If .status = ClsTransaksiPR.EnumStatus.PR_Pending Then
    '                        ROWItem.Status = "PR Pending"
    '                    ElseIf .status = ClsTransaksiPR.EnumStatus.Purchase_On_Progress Then
    '                        ROWItem.Status = "Purchase On Progress"
    '                    Else
    '                        ROWItem.Status = "Closed"
    '                    End If
    '                    ROWItem.Remarks = .remarks

    '                    ROWItem.Request_By = .ref_party_code + " - " + .ref_party_name
    '                    ROWItem.Remarks_Header = .remarks_header
    '                    If .Kode_User = "" Then
    '                        ROWItem.User_Authorize = ""
    '                    Else
    '                        ROWItem.User_Authorize = .Kode_User + " on " + Format(.time_authorize_by, "dd MMM yyyy HH:mm")
    '                    End If
    '                    ROWItem._Date = Format(.tanggal, "dd MMM yyyy")
    '                    ROWItem.PR_No = .no_transaksi

    '                    TBLItem.AddTBLItemRow(ROWItem)

    '                    If RowCetak_Item.delivery_date_cont = ClsTransaksiPR.EnumDeliveryDate.Cheklist Then
    '                        If .delivery_date = #1/1/1950# Then
    '                            TBLCetak_Delivery.DefaultView.RowFilter = "[identifier] = '" & RowCetak_Item.identifier.ToString & "'"
    '                            TBLCetak_Delivery.DefaultView.Sort = "[Delivery Date] ASC"
    '                            rowCetak_Delivery = TBLCetak_Delivery.DefaultView(TBLCetak_Delivery.DefaultView.Count - 1).Row
    '                            ROWItem.Delivery_Date = Format(rowCetak_Delivery.delivery_date, "dd MMM yyyy")
    '                        Else
    '                            ROWItem.Delivery_Date += Format(RowCetak_Item.delivery_date, "dd MMM yyyy")
    '                        End If
    '                    Else
    '                        ROWItem.Delivery_Date += "-"
    '                    End If
    '                End With
    '            Next
    '        End If

    '        'isikan availability
    '        '0. isikan list primary
    '        ListPrimary.Clear()
    '        For Each ROWItem In TBLItem.Rows
    '            If ListPrimary.Contains(ROWItem.Primary_Product) = False Then
    '                ListPrimary.Add(ROWItem.Primary_Product)
    '            End If
    '        Next
    '        '1. hitung shipping forecast untuk barang yg ditampilkan saja
    '        ObjSF.HitungShippingForecast(vTBLItemDelivery, vTBLItemPOB, vTBLItemStock, -1, _
    '        Core_Module.GetSMESSBackColorAll, -1, ListPrimary)
    '        '2. untuk goods, isikan datanya apa adanya berdasarkan shipping forecast
    '        For Each ROWItem In TBLItem.Rows
    '            vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + ROWItem.identifier + _
    '            "' AND tipe = " + CInt(Inventory_Module.ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req).ToString
    '            If ROWItem.Status = "Closed" Then
    '                ROWItem.availability = "-"
    '            Else
    '                ROWItem.availability = "Waiting"
    '                If vTBLItemDelivery.DefaultView.Count > 0 Then
    '                    vRowItemDelivery = vTBLItemDelivery.DefaultView(0).Row
    '                    If vTBLItemDelivery.DefaultView.Count > 1 Then
    '                        For ij As Integer = 1 To vTBLItemDelivery.DefaultView.Count - 1
    '                            vRowItemDelivery1 = vTBLItemDelivery.DefaultView(ij).Row
    '                            vRowItemDelivery.qty_sob += vRowItemDelivery1.qty_sob
    '                            vRowItemDelivery.qty_available += vRowItemDelivery1.qty_available
    '                            vRowItemDelivery.qty_ready_stock += vRowItemDelivery1.qty_ready_stock
    '                            If vRowItemDelivery.Isprimary_po_refNull = True Then
    '                                vRowItemDelivery.primary_po_ref = ""
    '                            End If
    '                            If vRowItemDelivery1.Isprimary_po_refNull = True Then
    '                                vRowItemDelivery1.primary_po_ref = ""
    '                            End If
    '                            vRowItemDelivery.primary_po_ref = vRowItemDelivery.primary_po_ref + "-" + vRowItemDelivery1.primary_po_ref
    '                        Next
    '                    End If
    '                    If vRowItemDelivery.qty_available > 0 Then
    '                        If vRowItemDelivery.qty_sob <= vRowItemDelivery.qty_available Then
    '                            If vRowItemDelivery.qty_ready_stock >= vRowItemDelivery.qty_available Then
    '                                ROWItem.availability = "Ready"
    '                            ElseIf vRowItemDelivery.qty_ready_stock > 0 Then
    '                                ROWItem.availability = "Partial"
    '                            Else
    '                                ROWItem.availability = "Waiting"
    '                            End If
    '                        Else
    '                            If vRowItemDelivery.qty_ready_stock >= vRowItemDelivery.qty_available Then
    '                                ROWItem.availability = "Partial"
    '                            ElseIf vRowItemDelivery.qty_ready_stock > 0 Then
    '                                ROWItem.availability = "Partial"
    '                            Else
    '                                ROWItem.availability = "Waiting"
    '                            End If
    '                        End If
    '                    Else
    '                        ROWItem.availability = "n/a"
    '                    End If
    '                    If vRowItemDelivery.Isprimary_po_refNull = False Then
    '                        If vRowItemDelivery.primary_po_ref <> "" Then
    '                            ROWItem.availability = Me.Translate_RefPO(vRowItemDelivery.primary_po_ref, vRowItemDelivery, ROWItem.availability, vRowItemDelivery._Decimal)
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Next

    '        '3. untuk service, hanya tampilkan data berdasarkan POB

    '        ReportDetail.SetDataSource(DS)
    '        ReportDetail.SetParameterValue(ReportDetail.Parameter_StartDate.ParameterFieldName, paramStartDate)
    '        ReportDetail.SetParameterValue(ReportDetail.Parameter_EndDate.ParameterFieldName, paramEndDate)
    '        Me.crv.ReportSource = ReportDetail
    '        Me.crv.Show()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DS.Dispose()
    '        DTACetak_Item.Dispose()
    '        TBLCetak_ItemD.Dispose()
    '        TBLCetak_Item.Dispose()
    '        TBLItem.Dispose()
    '        ListPrimary = Nothing
    '        ObjSF = Nothing
    '        vTBLItemDelivery.Dispose()
    '        vTBLItemPOB.Dispose()
    '        vTBLItemStock.Dispose()
    '        DTACetak_Delivery.Dispose()
    '        TBLCetak_Delivery.Dispose()
    '    End Try
    'End Sub

    Private Function Translate_RefPO(ByVal vPrimaryPORef As String, ByVal vRowItemDelivery As _
    Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow, ByVal vStatus As String, ByVal vDefPrec As Integer) As String
        Dim RetVal As String = ""
        Dim Pecahan As String()
        Dim Pecahan1 As String()
        Dim PONumber As String = ""
        Dim RefDate As Date = #1/1/1950#
        Dim Qty As Decimal = 0
        Dim TBL As New DsePurchaseRequestReport.DataTable1DataTable
        Dim row As DsePurchaseRequestReport.DataTable1Row

        Try
            If vStatus = "Ready" Then
                RetVal = ""
            ElseIf vStatus = "Partial" Or vStatus = "Waiting" Then
                'ready
                If vRowItemDelivery.qty_ready_stock > 0 Then
                    If RetVal = "" Then
                        RetVal = Me.ObjCd.DigitGrouping(vRowItemDelivery.qty_ready_stock, vDefPrec) + " (Ready)"
                    Else
                        RetVal = RetVal + vbCrLf + Me.ObjCd.DigitGrouping(vRowItemDelivery.qty_ready_stock, vDefPrec) + "(Ready)"
                    End If
                End If
                'pob
                Pecahan = vPrimaryPORef.Split("-")
                For Each Fragmen As String In Pecahan
                    If Fragmen <> "" Then
                        Pecahan1 = Fragmen.Split(";")
                        PONumber = Me.ObjPO.GetNoTransaksiPO(Pecahan1(0))
                        RefDate = CDate(Pecahan1(1))
                        Qty = CDec(Pecahan1(2))
                        TBL.DefaultView.RowFilter = "[PO Number] = '" + Me.ObjCd.FormatSpecialChar_DataFilter(PONumber) + _
                        "' AND [Ref Date] = #" + Format(RefDate, "MM/dd/yyyy") + "#"
                        If TBL.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            row = TBL.NewDataTable1Row

                            row.PO_Number = PONumber
                            row.Ref_Date = RefDate
                            row.Qty = Qty

                            TBL.Rows.Add(row)
                        Else
                            row = TBL.DefaultView(0).Row
                            row.Qty += Qty
                        End If
                    End If
                Next

                For Each row In TBL.Rows
                    PONumber = row.PO_Number
                    RefDate = row.Ref_Date
                    Qty = row.Qty

                    If RefDate = #1/1/2200# Then
                        If Me.chkShowPONumber.Checked = True And Me.chkShowEstimateDelivery.Checked = True Then
                            If RetVal = "" Then
                                RetVal = Me.ObjCd.DigitGrouping(Qty, vDefPrec) + vbCrLf + PONumber + ":n/a"
                            Else
                                RetVal = RetVal + vbCrLf + Me.ObjCd.DigitGrouping(Qty, vDefPrec) + vbCrLf + PONumber + ":n/a"
                            End If
                        ElseIf Me.chkShowPONumber.Checked = False And Me.chkShowEstimateDelivery.Checked = True Then
                            If RetVal = "" Then
                                RetVal = Me.ObjCd.DigitGrouping(Qty, vDefPrec) + " (n/a)"
                            Else
                                RetVal = RetVal + vbCrLf + Me.ObjCd.DigitGrouping(Qty, vDefPrec) + " (n/a)"
                            End If
                        ElseIf Me.chkShowPONumber.Checked = True And Me.chkShowEstimateDelivery.Checked = False Then
                            If RetVal = "" Then
                                RetVal = Me.ObjCd.DigitGrouping(Qty, vDefPrec) + " (" + PONumber + ")"
                            Else
                                RetVal = RetVal + vbCrLf + Me.ObjCd.DigitGrouping(Qty, vDefPrec) + " (" + PONumber + ")"
                            End If
                        End If
                    Else
                        If Me.chkShowPONumber.Checked = True And Me.chkShowEstimateDelivery.Checked = True Then
                            If RetVal = "" Then
                                RetVal = Me.ObjCd.DigitGrouping(Qty, vDefPrec) + vbCrLf + PONumber + ":" + Format(RefDate, "dd/MM/yy")
                            Else
                                RetVal = RetVal + vbCrLf + Me.ObjCd.DigitGrouping(Qty, vDefPrec) + vbCrLf + PONumber + ":" + Format(RefDate, "dd/MM/yy")
                            End If
                        ElseIf Me.chkShowPONumber.Checked = True And Me.chkShowEstimateDelivery.Checked = False Then
                            If RetVal = "" Then
                                RetVal = Me.ObjCd.DigitGrouping(Qty, vDefPrec) + " (" + PONumber + ")"
                            Else
                                RetVal = RetVal + vbCrLf + Me.ObjCd.DigitGrouping(Qty, vDefPrec) + " (" + PONumber + ")"
                            End If
                        ElseIf Me.chkShowPONumber.Checked = False And Me.chkShowEstimateDelivery.Checked = True Then
                            If RetVal = "" Then
                                RetVal = Me.ObjCd.DigitGrouping(Qty, vDefPrec) + " (" + Format(RefDate, "dd/MM/yy") + ")"
                            Else
                                RetVal = RetVal + vbCrLf + Me.ObjCd.DigitGrouping(Qty, vDefPrec) + "(" + Format(RefDate, "dd/MM/yy") + ")"
                            End If
                        End If
                    End If
                Next

                'n/a
                If vRowItemDelivery.qty_sob - vRowItemDelivery.qty_available > 0 Then
                    If RetVal = "" Then
                        RetVal = Me.ObjCd.DigitGrouping(vRowItemDelivery.qty_sob - vRowItemDelivery.qty_available, vDefPrec) + " (n/a)"
                    Else
                        RetVal = RetVal + vbCrLf + Me.ObjCd.DigitGrouping(vRowItemDelivery.qty_sob - vRowItemDelivery.qty_available, vDefPrec) + " (n/a)"
                    End If
                End If
            End If

            If Me.chkShowEstimateDelivery.Checked = True Or Me.chkShowPONumber.Checked = True Then
                If vStatus = "Partial" Or vStatus = "Waiting" Then
                    Return RetVal
                Else
                    Return vStatus
                End If
            Else
                Return vStatus
            End If
        Catch ex As Exception
            MsgBox(ex.Message + "-" + vPrimaryPORef.ToString, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            TBL = Nothing
        End Try
    End Function

    Private Sub CboPilih_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboPilih.SelectedIndexChanged
        Try
            If Me.CboPilih.SelectedIndex = 2 Then
                If Me.ChkOB.Checked = True Then
                    Me.ChkOB.Checked = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkOB_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkOB.CheckedChanged
        Try
            If Me.CboPilih.SelectedIndex = 2 Then
                Me.ChkOB.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkOB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkOB.Click
        Try
            If Me.CboPilih.SelectedIndex = 2 Then
                MsgBox("If you pr show with balance then cannot change unauthorize.", MsgBoxStyle.Information, "Information")
                Me.CboPilih.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemProductGroup.CellBeginEdit
        Dim rowDGV As DataGridViewRow

        Try
            'yang bisa diedit:
            '1. jika baris baru
            If e.RowIndex = Me.dgvItemProductGroup.Rows.Count - 1 Then
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'pastikan baris diatasnya sudah beres dahulu
                    If e.RowIndex > 0 Then
                        rowDGV = Me.dgvItemProductGroup.Rows(e.RowIndex - 1)
                        If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value.ToString.Trim = "" Then
                            MsgBox("Please specify " + rowDGV.Cells(Me.ColProductGroup_Level.Index).Value.ToString + _
                            " code on the previous row.", MsgBoxStyle.Information, "Information")
                            Me.dgvItemProductGroup.ClearSelection()
                            rowDGV.Selected = True
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = -1 Then
                    '2. jika bukan baris baru dan primary =-1
                    'yang bisa diedit hanya code product group dan level
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    '3. jika bukan baris baru dan primary <> -1
                    'yang bisa diedit: code, remarks
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or _
                    e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            If Me.DummyIndex <> -1 Then
                Me.dgvItemProductGroup.Rows.RemoveAt(Me.DummyIndex)
                DummyIndex = -1
                Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemProductGroup.CellEndEdit
        Dim Level As Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand

        Try
            If IsNothing(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                       Or Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'Remarks
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    'kalo kode product group dihapus, maka hapus satu bari semuanya
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                    Exit Sub
                End If
            Else
                'level
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'do nothing

                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If

                    'cek ke master
                    Level = Me.objItem.GetProductLevel(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
                    If Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                        Me.EndEditProductGroup(e, Me.ProductGroup_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                        Me.EndEditProductGroup(e, Me.Brand_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                        Me.EndEditProductGroup(e, Me.SubBrand_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                        Me.EndEditCategory(e)
                    Else
                        Me.EndEditProduct(e)
                    End If
                ElseIf e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'remarks
                    'pastikan ditrim dan  max 250 karakter
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'tidak boleh ada karakter aneh2
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Then
                        MsgBox("Please avoid using the following character: ', *, %, #.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditProductGroup(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal vProductGroup_Level As Integer)
        Try
            If Me.ObjProductGroup.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level) = True Then
                '1. kode product group ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product group yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.FrmCariMasterProductGroup
                ij.ProductGroupLevel = vProductGroup_Level
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProductGroup <> -1 Then
                    '2.a. product group yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProductGroup.GetGroupCode(ij.PrimaryProductGroup)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(ij.PrimaryProductGroup)
                ElseIf ij.PrimaryProductGroup = -1 Then
                    '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If
            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.buffer, vProductGroup_Level)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditCategory(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjCategory.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode category ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode category yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterCategory
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryCategory <> -1 Then
                    '2.a. category yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryCategory
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                ElseIf ij.PrimaryCategory = -1 Then
                    '2.b. kode category yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex_ItemProductGroup(ByVal vCode As String, ByVal vRefIndex As Integer, ByVal vProductGroupLevelStr As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = vCode And _
                rowDGV.Index <> vRefIndex And _
                rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = vProductGroupLevelStr Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub EndEditProduct(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjProduct.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode product ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterProduct
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProduct <> -1 Then
                    '2.a. product yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProduct
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                ElseIf ij.PrimaryProduct = -1 Then
                    '2.b. kode product yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItemProductGroup.DataError
        'do nothing -> dibuat untuk handle data error event, jangan dihapuskan
    End Sub

    Private Sub dgvItemProductGroup_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemProductGroup.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItemProductGroup.Rows(Me.dgvItemProductGroup.Rows.Count - 1)
            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = Me.objItem.GetProductLevelStr(Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup)
            rowDGV.Cells(Me.ColProductGroup_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value = -1
            rowDGV.Cells(Me.ColProductGroup_Remarks.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemProductGroup.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiDGV_ItemGroup()
        Dim ObjSettingPG As New Core_Module.ClsTransaksiSetting_ProdGroup

        Try
            ObjSettingPG.IsikanDGV(Me.UserCode, Core_Module.ClsTransaksiSetting_ProdGroup.EnumTipeTransaksi.LaporanPurchaseRequest, Me.dgvItemProductGroup, _
            Me.ColProductGroup_No.Index, Me.ColProductGroup_Level.Index, Me.ColProductGroup_CodeNo.Index, Me.ColProductGroup_Desc.Index, _
            Me.ColProductGroup_Remarks.Index, Me.ColProductGroup_PrimaryAcuan.Index)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSettingPG = Nothing
        End Try
    End Sub


    Private Sub DgvUser_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvUser.CellBeginEdit
        Try
            If Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value <> -1 Then
                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                    Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                    Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vKodeUser As String, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvUser.Rows.Count - 2
                RowDGV = Me.DgvUser.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCreatedBy.Index).Value = vKodeUser Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvUser_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvUser.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvUser.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim KodeUser As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCreatedBy.Index Then
                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjUser.ApakahSudahAda(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            KodeUser = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            index = Me.GetIndexItem(KodeUser, e.RowIndex)
                            If index < 0 Then
                                Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value)
                                Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value
                                Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwUser.RunWorkerAsync()
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterUser
                            ij.TxtCari.Text = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.backColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.Kodeuser <> "" Then
                                index = Me.GetIndexItem(ij.Kodeuser, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ij.Kodeuser
                                    Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                    Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(ij.Kodeuser)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwUser.RunWorkerAsync()
                                        Else
                                            Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Exit Sub
                                    Else
                                        Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ij.Kodeuser
                                        Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                        Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(ij.Kodeuser)
                                    End If
                                End If
                            ElseIf ij.Kodeuser = "" Then
                                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        If Me.DgvUser.Rows.Count - 1 <> e.RowIndex Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwUser.RunWorkerAsync()
                                        End If
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                Else
                                    Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.DgvUser)
                    End If
                Else
                    Me.EraseRowIndex = e.RowIndex
                    Me.BgwUser.RunWorkerAsync()
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvUser.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            ' If Me.TxtNoPO.Enabled = True Then
            rowDGV = Me.DgvUser.Rows(Me.DgvUser.RowCount - 1)
            With rowDGV
                ' .Cells(Me.ColN.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColCreatedBy.Index).Value = ""
                .Cells(Me.ColCreatedBy_Name.Index).Value = ""
                .Cells(Me.ColCreatedBy_Primary.Index).Value = "-1"
            End With
            ' End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUser_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwUser.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUser_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwUser.RunWorkerCompleted
        Try
            Me.DgvUser.Rows.RemoveAt(Me.EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class