Public Class FrmTrans_MaterialPlanningDashboard
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public _Command As EnumCommand = EnumCommand.NOTHING_
    Public _PrimaryMP As Double = -1

    Dim Buffer As String = ""
    Dim Splitter1_SplitterDistance As Integer = 0
    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjUser As New Core_Module.ClsUser

    Public Enum EnumCommand As Integer
        NOTHING_
        NEW_
        EDIT
        DELETE
        VIEW
    End Enum

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiReq_Purchase_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim dummy As String = ""

        Try
            'connection string untuk product muti unit (2 DTA) di set dalam fungsi isi cbo multi uni

            For Each ColItem As DataGridViewColumn In Me.dgvMP.Columns
                ColItem.SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            Me.dgvMP.ReadOnly = False
            Me.Splitter1_SplitterDistance = 326

            'BTS: semua kolom yang berhubungan dengan qty -> visible false
            Me.ColMP_QtyGIT.Visible = False
            Me.ColMP_QtyInquiry.Visible = False
            Me.ColMP_QtyLeadTime.Visible = False
            Me.ColMP_QtyOnHand.Visible = False
            Me.ColMP_QtyOrder.Visible = True
            Me.ColMP_QtyOrderMaximum.Visible = False
            Me.ColMP_QtyOrderOptimum.Visible = False
            Me.ColMP_QtyPendingDO.Visible = False
            Me.ColMP_QtyPOB.Visible = False
            Me.ColMP_QtyRequired.Visible = True
            Me.ColMP_QtySafetyStock.Visible = False
            Me.ColMP_QtySOB.Visible = False
            Me.ColMP_QtyOrderMinimum.Visible = False
            Me.ColMP_OrderStrategy.Visible = False

            'BTS: semua kolom yang berhubungan dengan purchase -> visible false
            Me.ColMP_PreferedSupplier.Visible = False
            Me.ColMP_NettPrice.Visible = False
            Me.ColMP_Disc.Visible = False
            Me.ColMP_Currency.Visible = False
            Me.ColMP_UnitPrice.Visible = False
            Me.ColMP_SubTotal.Visible = False
            Me.ColMP_UnitPrice.Visible = False

            Me.cboFilter_ProductGroupLevel.SelectedIndex = 0
            Me.SetFilter()

            'material planning
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvMP.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.NonAktifkanDetail_SF()
            Me.dgvSF.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.UpdateKolomSF()

            Me.Timer1.Enabled = True
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP") = False Then
                Me.gboMP.Enabled = False
                Me.btnMP_Delete.Enabled = False
                Me.btnMP_Edit.Enabled = False
                Me.btnMP_Find.Enabled = False
                Me.btnMP_New.Enabled = False
                Me.btnMP_Print.Enabled = False
                Me.TabControl1.SelectedIndex = 1
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_EditParameter") = False Then
                Me.LinkLabel2.Enabled = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_SF") = False Then
                Me.gboSF.Enabled = False
                Me.btnSF_Delete.Enabled = False
                Me.btnSF_Edit.Enabled = False
                Me.btnSF_Find.Enabled = False
                Me.btnSF_New.Enabled = False
                Me.btnSF_Print.Enabled = False
            End If
            If Me.vPrimaryStockInPurchaseReference = -1 Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_SF_DOAll") = False Then
                    Me.SetFilter()
                    Me.TxtFilter_UserCreateSO.Text = Me.UserCode
                    Me.TxtFilter_UserCreateSO.Enabled = False
                    Me.TxtFilter_NamaUserCreateSO.Text = Me.ObjUser.GetNama(Me.UserCode)
                    Me.btnFilter_FindUser.Enabled = False
                End If
            End If

            Me.cboAvailability.SelectedIndex = 0

            'rnew
            Me.CboFilter_ProcurementStrategy.SelectedIndex = 0

            If Me.vPrimaryStockInPurchaseReference <> -1 Then
                Me.TxtSF_StockInPurchaseNo.Visible = True
                Me.LblSF_StockInPurchaseNo.Visible = True
                Me.LinkLabel1.Visible = False

                Me.TabControl1.SelectedIndex = 1
                Me.SF_Baru()

                Me.TxtSF_StockInPurchaseNo.Text = Me.ObjStockInPurchase.GetNoTransaksi(Me.vPrimaryStockInPurchaseReference)

                Me.HitungShippingForecast()
                Me.LoadDGV_SF()
                Me.FollowUp()
                Me.btnSF_OK.Enabled = False
                Me.btnSF_Cancel.Enabled = False
                Me.ControlBox = False
            Else
                Me.TxtSF_StockInPurchaseNo.Visible = False
                Me.LblSF_StockInPurchaseNo.Visible = False
                Me.LinkLabel1.Visible = True
            End If

            'load setting PrOB > POB
            dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanningGrouping_PrOBLargerThanPOB, Me.UserCode)
            If dummy = "" Or dummy = "0" Then
                Me.chkFilter_PrOBLargerThanPOB.Checked = False
            Else
                Me.chkFilter_PrOBLargerThanPOB.Checked = True
            End If
            Me.llbMP_AddToPO.Text = "Add to PO/MO"
            Me.llbMP_CreatePO.Text = "Create PO/MO"
            Me.llbMP_CreatePO.Left = Me.TabControl2.Left + Me.TabControl2.Width - 5 - Me.llbMP_CreatePO.Width
            Me.llbMP_AddToPO.Left = Me.llbMP_CreatePO.Left - 5 - Me.llbMP_AddToPO.Width
            Me.llbMP_CekAll.Left = Me.llbMP_AddToPO.Left - 15 - Me.llbMP_CekAll.Width
            ' Me.ColMP_PrimaryUnitMulti.Visible = True
            If Me._Command = EnumCommand.VIEW And Me._PrimaryMP <> -1 Then
                Me.ViewMPDirectly()
            End If
            Me.SetDPM()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ViewMPDirectly()
        Try
            Me._Command = EnumCommand.NOTHING_
            If Me.TxtMP_Remarks.Enabled Then
                If Me.SaveRequired = True Then
                    If MsgBox("You've just been following up one or more item, you have to save this transaction." & vbNewLine _
                    & "Do you want to save this transaction and load MP No " & Me.ObjMP.GetDocumentNo(Me._PrimaryMP) & "?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                    If Me.ObjMP.SimpanPerubahan(Me.vDT, Me.UserCode, Me, Me.vDT_ItemGIT, Me.vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, _
                    vDT_ItemPrOB, vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, Me.vDT_ItemFollowUp, Me.vDT_ItemSOBDelivery, Me.vDT_ItemPrOBDelivery) Then
                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Me.SaveRequired = False
                        Me.NonAktifkanDetail()
                    End If
                Else
                    If MsgBox("Are you sure to discard this data and load MP No " & Me.ObjMP.GetDocumentNo(Me._PrimaryMP) & "?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
            End If
            Me.Cursor = Cursors.WaitCursor
            If Me.ObjMP.CariDetail(Me._PrimaryMP, Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, _
                vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, Me.vDT_ItemFollowUp, Me.vDT_ItemSOBDelivery, Me.vDT_ItemPrOBDelivery) = True Then
                Me.SalinKeField()
                Me.btnMP_Edit.Focus()
            Else
                Me.Close()
            End If
            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvMP.DataError
        Try
            'bagian ini harus ada untuk mencegah error ketika load form
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView2_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs)
        Try
            'bagian ini harus ada ketika melakukan load form untuk mencegah error

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SplitContainer1_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles SplitContainer1.SizeChanged
        Try
            If Me.Splitter1_SplitterDistance <> 0 Then
                Me.SplitContainer1.SplitterDistance = Me.Splitter1_SplitterDistance
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Me.SetFilter()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetFilter()
        Try
            If Me.SplitContainer1.Panel1Collapsed = True Then
                Me.LinkLabel1.Text = "<<< Hide Filter"
                Me.SplitContainer1.Panel1Collapsed = False
                Me.SplitContainer1.SplitterDistance = Me.Splitter1_SplitterDistance
                Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
                Me.TxtFilter_PrimaryPreferedSupplier.Text = "-1"
                Me.TxtFilter_PrimaryProductGroup.Text = "-1"
                Me.TxtFilter_PrimaryCustomer.Text = "-1"

                Me.SetFilterCustSupp()
            Else
                Me.LinkLabel1.Text = ">>> Show Filter"
                Me.SplitContainer1.Panel1Collapsed = True
            End If
            Me.TabControl1.Refresh()
            Me.dgvMP.ScrollBars = ScrollBars.Both
            Me.dgvMP.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetFilterCustSupp()
        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                Me.TxtFilter_PreferedSupplier.Visible = True
                Me.btnFilter_FindPreferedSupplier.Visible = True
                Me.TxtFilter_PreferedSupplier.TabStop = True
                Me.TxtFilter_PreferedSupplierDesc.Visible = True
                Me.LblFilter_PreferedSupplier.Visible = True

                Me.TxtFilter_CustCode.Visible = False
                Me.TxtFilter_CustCode.TabStop = False
                Me.TxtFilter_CustName.Visible = False
                Me.btnFilter_FindCust.Visible = False
                Me.btnFilter_FindCust.TabStop = False
                Me.LblFilter_Cust.Visible = False

                Me.TxtFilter_UserCreateSO.Visible = False
                Me.TxtFilter_UserCreateSO.TabStop = False
                Me.LblFilter_UserCreateSO.Visible = False
                Me.btnFilter_FindUser.Visible = False
                Me.btnFilter_FindUser.TabStop = False
                Me.TxtFilter_NamaUserCreateSO.TabStop = False
                Me.TxtFilter_NamaUserCreateSO.Visible = False

                Me.GboFilter_SODate.Visible = False
                Me.LblFilter_ProcurementStrategy.Visible = True
                Me.CboFilter_ProcurementStrategy.Visible = True

                Me.LblRefParty.Visible = False
                Me.Txtfilter_RefPartyCode.Visible = False
                Me.Txtfilter_RefPartyCode.TabStop = False
                Me.Txtfilter_RefPartyName.Visible = False
                Me.Txtfilter_RefPartyName.TabStop = False
                Me.btnFilter_FindRefParty.Visible = False
                Me.btnFilter_FindRefParty.TabStop = False
                Me.GboFilter_PurcReqDate.Visible = False

                Me.btnFilter_FindPurchReq.Top = CboFilter_ProcurementStrategy.Top + CboFilter_ProcurementStrategy.Height + 5
                Me.TxtFilter_PurcReqNo.Top = CboFilter_ProcurementStrategy.Top + CboFilter_ProcurementStrategy.Height + 5
                Me.LblFilter_PurcReq.Top = CboFilter_ProcurementStrategy.Top + CboFilter_ProcurementStrategy.Height + 5
                Me.LblFilter_PurcReq.Visible = True
                Me.TxtFilter_PurcReqNo.Visible = True
                Me.TxtFilter_PurcReqNo.TabStop = True
                Me.btnFilter_FindPurchReq.Visible = True
                Me.btnFilter_FindPurchReq.TabStop = True

                Me.GboFilter_MODate.Visible = True
                Me.GboFilter_MODate.Top = Me.btnFilter_FindPurchReq.Top + Me.btnFilter_FindPurchReq.Height + 10

                Me.TxtFilter_MONo.Top = Me.GboFilter_MODate.Top + Me.GboFilter_MODate.Height + 5
                Me.btnFilter_FindMO.Top = Me.GboFilter_MODate.Top + Me.GboFilter_MODate.Height + 5
                Me.LblFilter_FindMO.Top = Me.GboFilter_MODate.Top + Me.GboFilter_MODate.Height + 5
                Me.LblFilter_FindMO.Text = "MO No"
                Me.TxtFilter_MONo.Visible = True
                Me.btnFilter_FindMO.Visible = True
                Me.LblFilter_FindMO.Visible = True

                Me.llbMP_Filter.Top = Me.TxtFilter_MONo.Top + Me.TxtFilter_MONo.Height + 5
            Else
                Me.LblFilter_PurcReq.Visible = False
                Me.TxtFilter_PurcReqNo.Visible = False
                Me.TxtFilter_PurcReqNo.TabStop = False
                Me.btnFilter_FindPurchReq.Visible = False
                Me.btnFilter_FindPurchReq.TabStop = False

                Me.TxtFilter_PreferedSupplier.Visible = False
                Me.btnFilter_FindPreferedSupplier.Visible = False
                Me.btnFilter_FindPreferedSupplier.TabStop = False
                Me.TxtFilter_PreferedSupplier.TabStop = False
                Me.TxtFilter_PreferedSupplierDesc.Visible = False
                Me.LblFilter_PreferedSupplier.Visible = False

                Me.TxtFilter_CustCode.Visible = True
                Me.TxtFilter_CustCode.TabStop = True
                Me.TxtFilter_CustName.Visible = True
                Me.btnFilter_FindCust.Visible = True
                Me.LblFilter_Cust.Visible = True

                Me.TxtFilter_UserCreateSO.Visible = True
                Me.TxtFilter_UserCreateSO.TabStop = True
                Me.LblFilter_UserCreateSO.Visible = True
                Me.btnFilter_FindUser.Visible = True
                Me.TxtFilter_NamaUserCreateSO.Visible = True

                Me.GboFilter_SODate.Visible = True
                Me.GboFilter_MODate.Visible = False
                Me.LblFilter_ProcurementStrategy.Visible = False
                Me.CboFilter_ProcurementStrategy.Visible = False

                Me.LblRefParty.Visible = True
                Me.Txtfilter_RefPartyCode.Visible = True
                Me.Txtfilter_RefPartyCode.TabStop = True
                Me.Txtfilter_RefPartyName.Visible = True
                Me.btnFilter_FindRefParty.Visible = True
                Me.GboFilter_PurcReqDate.Visible = True
                Me.TxtFilter_MONo.Visible = False
                Me.btnFilter_FindMO.Visible = False
                Me.LblFilter_FindMO.Visible = False

                Me.llbMP_Filter.Top = Me.GboFilter_PurcReqDate.Top + Me.GboFilter_PurcReqDate.Height + 5
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            Dim i As New FrmTrans_MaterialPlanningDashboard_Data
            i.ShowDialog()
            i.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                If Me.vPrimaryStockInPurchaseReference <> -1 Then
                    Me.TabControl1.SelectedIndex = 1
                    Exit Sub
                End If

                Me.LinkLabel2.Visible = True
                Me.llbColumnSettings.Visible = False
                Me.LinkLabel3.Visible = True
            Else
                Me.LinkLabel2.Visible = False
                Me.llbColumnSettings.Visible = True
                Me.LinkLabel3.Visible = False
            End If

            Me.SetFilterCustSupp()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            If Me.vPrimaryStockInPurchaseReference = -1 Then
                Me.LinkLabel1.Visible = True
                Me.LinkLabel1.BringToFront()

                Me.LinkLabel1.Location = New Point(Me.LinkLabel1.Location.X, Me.llbColumnSettings.Location.Y)
            End If

            Me.Timer1.Enabled = False

            Me.btnClose.SendToBack()
            Me.btnClose.Location = New Point(100, 100)
            Me.SplitContainer1.SplitterDistance = 326

            Me.ColMP_ProcurementStrategy.SortMode = DataGridViewColumnSortMode.Automatic
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAuthDetail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAuthDetail.LinkClicked
        Dim frm As New FrmTrans_MaterialPlanningDashboard_AuthorizeDetail
        Dim rowDgv As DataGridViewRow
        Dim dta As New DseTrans_MaterialPlanningDashboard_AuthorizeTableAdapters.tbltrans_material_planning2TableAdapter
        Dim tbl As New DseTrans_MaterialPlanningDashboard_Authorize.tbltrans_material_planning2DataTable
        Dim row As DseTrans_MaterialPlanningDashboard_Authorize.tbltrans_material_planning2Row
        Try
            If IsNothing(Me.ObjMP) = False Then
                If Me.ObjMP.Primary > -1 Then
                    If Me.ObjMP.GetAuthLevelNeeded(Me.ObjMP.Primary) = 0 Then
                        MsgBox("This MP no authorization needed.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    dta.Connection.ConnectionString = Core_Module.GetConstring
                    If dta.Fill(tbl, Me.ObjMP.Primary) > 0 Then
                        row = tbl.Rows(0)
                        rowDgv = frm.dgvItem.Rows(frm.dgvItem.Rows.Add())
                        rowDgv.Cells(frm.colNo.Index).Value = 1
                        rowDgv.Cells(frm.colAuth.Index).Value = "Level 1"
                        rowDgv.Cells(frm.colUserCode.Index).Value = row.user_auth_1
                        If row.authorize_lv1_time > #1/1/1950# Then
                            rowDgv.Cells(frm.colTime.Index).Value = Format(row.authorize_lv1_time, "dd MMM yyyy HH.mm")
                        Else
                            rowDgv.Cells(frm.colTime.Index).Value = "-"
                        End If
                        rowDgv = frm.dgvItem.Rows(frm.dgvItem.Rows.Add())
                        rowDgv.Cells(frm.colNo.Index).Value = 2
                        rowDgv.Cells(frm.colAuth.Index).Value = "Level 2"
                        rowDgv.Cells(frm.colUserCode.Index).Value = row.user_auth_2
                        If row.authorize_lv2_time > #1/1/1950# Then
                            rowDgv.Cells(frm.colTime.Index).Value = Format(row.authorize_lv2_time, "dd MMM yyyy HH.mm")
                        Else
                            rowDgv.Cells(frm.colTime.Index).Value = "-"
                        End If
                        rowDgv = frm.dgvItem.Rows(frm.dgvItem.Rows.Add())
                        rowDgv.Cells(frm.colNo.Index).Value = 3
                        rowDgv.Cells(frm.colAuth.Index).Value = "Level 3"
                        rowDgv.Cells(frm.colUserCode.Index).Value = row.user_auth_3
                        If row.authorize_lv3_time > #1/1/1950# Then
                            rowDgv.Cells(frm.colTime.Index).Value = Format(row.authorize_lv3_time, "dd MMM yyyy HH.mm")
                        Else
                            rowDgv.Cells(frm.colTime.Index).Value = "-"
                        End If
                        frm.ShowDialog()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Try
            If Me._Command = EnumCommand.VIEW And Me._PrimaryMP <> -1 Then
                Me.ViewMPDirectly()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class