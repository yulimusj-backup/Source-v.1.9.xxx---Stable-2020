Option Explicit On
Option Strict Off

Imports Core_Module
Partial Public Class FrmTrans_MaterialPlanningDashboard
    Dim baruSF As Boolean
    Dim ObjSF As New ClsTransaksiShippingForecast
    Dim ObjSFItem As New ClsTransaksiShippingForecast_Item
    Dim ObjPriorityLevel As New Core_Module.ClsMasterPriorityLevel
    Dim vTBLItemStock As New DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable
    Dim vTBLItemPOB As New DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable
    Public vTBLItemDelivery As New DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable
    Dim mSortColIndex As Integer = -1
    Dim mSortASC As Boolean = False
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim ObjRefParty As New Core_Module.ClsMasterRefParty
    Dim ObjDO As New ClsTransaksiDO
    Dim ObjPickUpOrder As New ClsTransaksiPR_PickUpOrder
    Dim TipeItemSO_Str As String = "SO"
    Dim TipeItemPurchReq_Str As String = "Purc. Req"
    Public vPrimaryStockInPurchaseReference As Double = -1
    Public vFromStockIn As Boolean = False
    Public TBL_ProductQty1 As New DseTransaksiShippingForecast.tblmaster_productqtyDataTable

    Dim BufferCekBoxSF As String = ""

    Private Sub AktifkanDetail_SF()
        Try
            Me.ObjCD.SetBackColor(Me.gboSF.Controls, Me.TxtSF_Primary.BackColor, False)
            ModCekData.KendalikanControl(Me.gboSF.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.TabPage3.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnSF_Cancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.dgvSF.ReadOnly = False

            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel1.Controls, Me.TxtSF_Primary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel2.Controls, Me.TxtSF_Primary.BackColor, False)

            If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._SF) = True Then
                Me.TxtSF_DocNo.ReadOnly = True
                Me.TxtSF_DocNo.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtSF_DocNo.TabStop = False
                Me.dtpSF.Focus()
            Else
                Me.TxtSF_DocNo.ReadOnly = False
                Me.TxtSF_DocNo.BackColor = Me.TxtSF_Primary.BackColor
                Me.TxtSF_DocNo.TabStop = True
                Me.TxtSF_DocNo.Focus()
            End If

            Me.llbSF_Calculate.Enabled = True
            If Me.baruSF = True Then
                Me.llbSF_CreateDO.Enabled = True
                'Me.llbSF_AddToDO.Enabled = True
            Else
                Me.llbSF_CreateDO.Enabled = False
                'Me.llbSF_AddToDO.Enabled = False
            End If
            Me.llbSF_CekAll.Enabled = True

            Me.cboAvailability.Enabled = True
            Me.chkSF_LateOnly.Enabled = True

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_SF_CreateDO") = False Then
                Me.llbSF_CreateDO.Enabled = False
                'Me.llbSF_AddToDO.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail_SF()
        Try
            ModCekData.KendalikanControl(Me.gboSF.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.TabPage3.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnClose

            Me.dgvSF.AllowUserToDeleteRows = False
            Me.dgvSF.ReadOnly = True
            Me.ObjCD.SetBackColor(Me.gboSF.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel2.Controls, Me.BackColor, False)

            Me.llbSF_Calculate.Enabled = False
            Me.llbSF_CreateDO.Enabled = False
            'Me.llbSF_AddToDO.Enabled = False
            Me.llbSF_CekAll.Enabled = False

            Me.chkSF_LateOnly.Enabled = False

            Me.cboAvailability.Enabled = False
            Me.chkSF_LateOnly.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField_SF()
        Try
            ModCekData.KosongkanControl(Me.gboSF.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer6.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer6.Panel2.Controls, True, True, True)

            Me.TxtSF_Primary.Text = "-1"
            Me.tslSF_CreatedBy.Text = "Created By: "
            Me.tslSF_CreatedDate.Text = "Created Date: "
            Me.tslSF_ModifiedBy.Text = "Modified By: "
            Me.tslSF_ModifiedDate.Text = "Modified Date: "

            Me.dgvSF.Rows.Clear()
            Me.vTBLItemStock.Rows.Clear()
            Me.vTBLItemPOB.Rows.Clear()
            Me.vTBLItemDelivery.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar_SF() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/1950#
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpSF.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("document no")
            KoleksiTextBox.Add(Me.TxtSF_DocNo)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'minimal harus punya 1 data item
            If Me.vTBLItemDelivery.Rows.Count = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._SF) = True Then
                If Me.baruSF = False Then
                    If Me.ObjSF.TanggalTransaksi.Month = Me.dtpSF.Value.Month And _
                                       Me.ObjSF.TanggalTransaksi.Year = Me.dtpSF.Value.Year Then
                        Me.TxtSF_DocNo.Text = Me.ObjSF.NoTransaksi
                    Else
                        If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                        MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                            Me.dtpSF.Value = Me.ObjSF.TanggalTransaksi
                            Me.TxtSF_DocNo.Text = Me.ObjSF.NoTransaksi
                            Return False
                        End If
                    End If
                End If
            End If

            'kalo saSFe cek berkaitan dengan TimeStamp apakah masih benar
            If Me.baruSF = False Then
                If Me.ObjSF.TimeStamp <> Me.ObjSF.GetTimestamp(Me.ObjSF.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjSF.CariDetail(Me.ObjSF.Primary, Me.vTBLItemStock, Me.vTBLItemPOB, Me.vTBLItemDelivery)
                        Me.SalinKeField_SF()
                    End If
                    Return False
                End If
            End If

            'kalo saSFe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            DTA.Dispose()
        End Try
    End Function

    Private Sub SalinKeObjek_SF()
        Try
            'salin data header
            With Me.ObjSF
                .NoTransaksi = Me.TxtSF_DocNo.Text
                .TanggalTransaksi = Me.dtpSF.Value
                .Remarks = Me.TxtSF_Remarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetStatusCek_SF()
        Dim rowDGV As DataGridViewRow
        Dim CekSemua As Boolean = True

        Try
            If Me.dgvSF.RowCount = 0 Then
                Me.llbSF_CekAll.Visible = False
            Else
                Me.llbSF_CekAll.Visible = True
                For Each rowDGV In Me.dgvSF.Rows
                    If rowDGV.Cells(Me.ColSF_CreateDO.Index).Value = Boolean.FalseString Then
                        CekSemua = False
                    End If

                    If CekSemua = True Then
                        Me.llbSF_CekAll.Text = "Uncheck All"
                    Else
                        Me.llbSF_CekAll.Text = "Check All"
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField_SF()
        Try
            With Me.ObjSF
                Me.TxtSF_DocNo.Text = .NoTransaksi
                Me.dtpSF.Value = .TanggalTransaksi
                Me.TxtSF_Remarks.Text = .Remarks
                Me.TxtSF_Primary.Text = .Primary

                Me.tslSF_CreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslSF_CreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslSF_ModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslSF_ModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
            Me.LoadDGV_SF()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSF_OK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSF_OK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar_SF() = True Then
                Me.SalinKeObjek_SF()
                If baruSF = True Then
                    'konfirmasi diSimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjSF.SimpanBaru(Me.UserCode, Me, Me.vTBLItemPOB, Me.vTBLItemDelivery, Me.vTBLItemStock) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.ObjSF.CariDetail(Me.ObjSF.Primary, Me.vTBLItemStock, Me.vTBLItemPOB, Me.vTBLItemDelivery)
                            Me.SalinKeField_SF()
                            Me.NonAktifkanDetail_SF()
                            Me.btnSF_Find.Focus()
                        End If
                    End If
                Else
                    'konfirmasi diSimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjSF.SimpanPerubahan(Me.UserCode, Me, vTBLItemPOB, vTBLItemDelivery, Me.vTBLItemStock) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField_SF()
                            Me.NonAktifkanDetail_SF()
                            Me.btnSF_Edit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSF_Cancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtSF_Remarks.Text = "" And Me.vTBLItemDelivery.Rows.Count = 0 And Me.dtpSF.Value.Date = Today Then
                Me.NonAktifkanDetail_SF()
                Me.btnSF_New.Focus()
                Me.KosongkanField_SF()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If Me.vPrimaryStockInPurchaseReference = -1 Then
                    If hasil = MsgBoxResult.Yes Then
                        If baruSF = True Then
                            Me.KosongkanField_SF()
                        Else
                            Me.ObjSF.CariDetail(Me.TxtSF_Primary.Text, Me.vTBLItemStock, Me.vTBLItemPOB, vTBLItemDelivery)
                            Me.SalinKeField_SF()
                        End If
                    Else
                        Exit Sub
                    End If
                    Me.NonAktifkanDetail_SF()
                    Me.btnSF_New.Focus()
                Else
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSF_New_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSF_New.Click
        Try
            Me.SF_Baru()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SF_Baru()
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baruSF = True
            Me.AktifkanDetail_SF()
            Me.KosongkanField_SF()
            'Me.dtpSF.Focus()
            Me.ObjSF.ResetObject(Me.vTBLItemPOB, Me.vTBLItemDelivery, Me.vTBLItemStock)
            Me.TxtSF_DocNo.Text = Me.ObjSF.GetNoBaru1(Me.dtpSF.Value)
            Me.tslSF_CreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslSF_CreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslSF_ModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslSF_ModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.SetStatusCek_SF()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub btnSF_Find_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSF_Find.Click
        Dim hj As New FrmTrans_ShippingForecastDashboard_SF_Find

        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()

            If Not hj.PrimaryTransaksi = -1 Then
                If Me.ObjSF.CariDetail(hj.PrimaryTransaksi, Me.vTBLItemStock, vTBLItemPOB, vTBLItemDelivery) = True Then
                    Me.SalinKeField_SF()
                    Me.btnSF_Edit.Focus()
                Else
                    Me.btnSF_Find.Focus()
                End If
                Me.btnSF_Find.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnSF_Edit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSF_Edit.Click
        Try
            If Me.TxtSF_DocNo.Text = "" Then
                Me.btnSF_Find.PerformClick()
            End If

            If Me.TxtSF_DocNo.Text <> "" Then
                If Me.ObjSF.ApakahBolehEdit(Me.TxtSF_Primary.Text, Me.UserCode, True, True) = True Then
                    If Me.ObjSF.TimeStamp <> Me.ObjSF.GetTimestamp(Me.ObjSF.Primary) Then
                        Me.ObjSF.CariDetail(Me.ObjSF.Primary, Me.vTBLItemStock, vTBLItemPOB, vTBLItemDelivery)
                    End If
                    Me.baruSF = False
                    Me.AktifkanDetail_SF()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSF_Delete.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.TxtSF_DocNo.Text = "" Then
                Me.btnSF_Find.PerformClick()
            End If

            If Me.TxtSF_DocNo.Text <> "" Then
                If Me.ObjSF.ApakahBolehHapus(Me.TxtSF_Primary.Text, Me.UserCode, True, True) = False Then
                    Me.btnSF_Find.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjSF.Hapus(Me.TxtSF_Primary.Text, Me.UserCode) = True Then
                            Me.KosongkanField_SF()
                            Me.btnSF_New.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub TxtSF_Remarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSF_Remarks.GotFocus
        Try
            Me.TxtSF_Remarks.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSF_Remarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSF_Remarks.LostFocus
        Try
            Me.TxtSF_Remarks.BackColor = Me.TxtSF_Primary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpSF_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpSF.Leave
        Try
            If Me.TxtSF_Remarks.Enabled = True Then
                Me.BackgroundWorker4.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            If Me.TxtSF_Remarks.Enabled = True Then
                If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._SF) = True Then
                    If Me.baruSF = False Then
                        If Me.ObjSF.TanggalTransaksi.Month = Me.dtpSF.Value.Month And Me.ObjSF.TanggalTransaksi.Year = Me.dtpSF.Value.Year Then
                            'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                            Me.TxtSF_DocNo.Text = Me.ObjSF.NoTransaksi
                        Else
                            'jika pindah bulan maka akan ganti no
                            Me.TxtSF_DocNo.Text = Me.ObjSF.GetNoBaru1(Me.dtpSF.Value)
                        End If
                    Else
                        Me.TxtSF_DocNo.Text = Me.ObjSF.GetNoBaru1(Me.dtpSF.Value)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSF_DocNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSF_DocNo.GotFocus
        Try
            If Me.TxtSF_Remarks.Enabled = True And Me.TxtSF_DocNo.Enabled = True Then
                If Me.TxtSF_DocNo.ReadOnly = True Then
                    Me.TxtSF_DocNo.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtSF_DocNo.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSF_DocNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSF_DocNo.LostFocus
        Try
            If Me.TxtSF_Remarks.Enabled = True And Me.TxtSF_DocNo.Enabled = True Then
                If Me.TxtSF_DocNo.ReadOnly = True Then
                    Me.TxtSF_DocNo.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtSF_DocNo.BackColor = Me.TxtSF_Primary.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbSF_Calculate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSF_Calculate.LinkClicked
        Try
            If Me.TxtSF_Remarks.Enabled = True Then
                If MsgBox("Calculate shipping forecast?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.HitungShippingForecast()
                    Me.LoadDGV_SF()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMax_Stock() As Double
        Dim rowStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockRow
        Dim Retval As Double = 0

        Try
            For Each rowStock In Me.vTBLItemStock.Rows
                If Retval < rowStock.primary Then
                    Retval = rowStock.primary
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub HitungShippingForecast()
        Dim DTA_ProductQty As New DseTransaksiShippingForecastTableAdapters.tblmaster_productqtyTableAdapter
        Dim TBL_ProductQty As New DseTransaksiShippingForecast.tblmaster_productqtyDataTable
        Dim row_ProductQty, row_ProductQty1 As DseTransaksiShippingForecast.tblmaster_productqtyRow
        Dim rowStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockRow
        Dim rowItemPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim CekOtorisasiPO As Boolean = False
        Dim DTAPOItem As New DseTransaksiShippingForecastTableAdapters.tbltrans_po_itemTableAdapter
        Dim TBLPOItem As New DseTransaksiShippingForecast.tbltrans_po_itemDataTable
        Dim rowPOItem As DseTransaksiShippingForecast.tbltrans_po_itemRow
        Dim rowDGV As DataGridViewRow
        Dim TBL_POB As New DseTransaksiMaterialPlanning2.DataTable3DataTable
        Dim row_POB As DseTransaksiMaterialPlanning2.DataTable3Row
        Dim ListSudah As New List(Of Double)
        Dim PrimaryPOMax, PrimaryPOMin As Double
        Dim DTAItemDeliveryPO As New DseTransaksiShippingForecastTableAdapters.tbltrans_po_itemdeliveryTableAdapter
        Dim TBLItemDeliveryPO As New DseTransaksiShippingForecast.tbltrans_po_itemdeliveryDataTable
        Dim rowItemDeliveryPO As DseTransaksiShippingForecast.tbltrans_po_itemdeliveryRow
        Dim DTAPackingList As New DseTransaksiShippingForecastTableAdapters.tblpackinglist_itemTableAdapter
        Dim TBLPackingList As New DseTransaksiShippingForecast.tblpackinglist_itemDataTable
        Dim rowPackingList As DseTransaksiShippingForecast.tblpackinglist_itemRow
        Dim DTASOItem As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemTableAdapter
        Dim TBLSOItem As New DseTransaksiShippingForecast.tbltrans_so_itemDataTable
        Dim rowSOItem As DseTransaksiShippingForecast.tbltrans_so_itemRow
        Dim DTASOItemDelivery As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim TBLSOItemDelivery As New DseTransaksiShippingForecast.tbltrans_so_itemdeliveryDataTable
        Dim rowSOItemDelivery As DseTransaksiShippingForecast.tbltrans_so_itemdeliveryRow
        Dim rowItemSO_SF As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim PrimaryPriorityLevelDefault As Double = Me.ObjPriorityLevel.GetPrimaryDefault
        Dim UrutanPriodityLevelDefault As Integer = Me.ObjPriorityLevel.GetUrutan(PrimaryPriorityLevelDefault)

        Dim ListPrimaryCodeNo As New List(Of Double)
        Dim DTAPendingDO As New DseTransaksiShippingForecastTableAdapters.tbltrans_do_itemTableAdapter
        Dim TBLPendingDO As New DseTransaksiShippingForecast.tbltrans_do_itemDataTable
        Dim rowPendingDO As DseTransaksiShippingForecast.tbltrans_do_itemRow

        Dim DTABooking As New DseTransaksiShippingForecastTableAdapters.tbltrans_booking_invitemTableAdapter
        Dim TBLBooking As New DseTransaksiShippingForecast.tbltrans_booking_invitemDataTable
        Dim rowBooking As DseTransaksiShippingForecast.tbltrans_booking_invitemRow
        Dim IndexAcuan As Integer = 0
        Dim QtyPOReady As Decimal
        Dim QtyReadyStock, QtyReadyStock1 As Decimal
        Dim IndexItemDelivery As Integer = 0
        Dim IndexItemPOB As Integer = 0
        Dim Dummy As Decimal = 0

        Dim DTAPurchReqItem As New DseTransaksiShippingForecastTableAdapters.tbltrans_purch_req_itemTableAdapter
        Dim TBLPurchReqItem As New DseTransaksiShippingForecast.tbltrans_purch_req_itemDataTable
        Dim rowPurchReqItem As DseTransaksiShippingForecast.tbltrans_purch_req_itemRow
        Dim DTAPurchReqItemDelivery As New DseTransaksiShippingForecastTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Dim TBLPurchReqItemDelivery As New DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryDataTable
        Dim rowPurchReqItemDelivery As DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryRow
        Dim sdf As Double = 0
        Dim ListPrimaryCodeNo_StockIn As New List(Of Double)

        Dim DTAGroup As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_item_adv_exclusionTableAdapter
        Dim TBLGroup As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionDataTable
        Dim ijkl As New FrmTrans_MaterialPlanningDashboard_Adv_Exclusion
        Dim LanjutQty As Boolean = False

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'tidak mungkin follow up only
            Me.chkFilter_FollowUpOnly.Checked = False

            DTA_ProductQty.Connection.ConnectionString = Core_Module.GetConstring
            DTAPOItem.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTAPendingDO.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTAItemDeliveryPO.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTAPackingList.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTASOItemDelivery.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTABooking.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTAPurchReqItem.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTAPurchReqItemDelivery.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTAGroup.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString

            'kosongkan data awal
            Me.vTBLItemDelivery.Rows.Clear()
            Me.vTBLItemPOB.Rows.Clear()
            Me.vTBLItemStock.Rows.Clear()

            'load data list primary code no yang diperhitungkan 
            'jika merujuk pada stock in purchase
            If Me.vPrimaryStockInPurchaseReference <> -1 Then
                ListPrimaryCodeNo_StockIn.Clear()
                For Each row_ProductQty1 In Me.TBL_ProductQty1.Rows
                    If ListPrimaryCodeNo_StockIn.Contains(row_ProductQty1.Primary_Product) = False Then
                        ListPrimaryCodeNo_StockIn.Add(row_ProductQty1.Primary_Product)
                    End If
                Next
            End If

            'load data stock
            If Me.vFromStockIn = False Then
                'perhitunngan normal
                DTA_ProductQty.Fill(TBL_ProductQty)
                For Each row_ProductQty In TBL_ProductQty.Rows
                    rowStock = Me.vTBLItemStock.Newtbltrans_shipping_forecast_itemstockRow

                    With rowStock
                        .primary = Me.GetPrimaryMax_Stock + 1
                        .primary_code_no = row_ProductQty.Primary_Product
                        .primary_transaksi = Me.TxtSF_Primary.Text
                        .primary_unit = row_ProductQty.Primary_Unit
                        .qty = row_ProductQty.Qty
                    End With
                    Me.vTBLItemStock.Rows.Add(rowStock)
                Next
                'kurangkan data stock dengan pending do
                DTAPendingDO.Fill(TBLPendingDO)
                If Me.vPrimaryStockInPurchaseReference <> -1 Then
                    For gh As Integer = TBLPendingDO.Rows.Count - 1 To 0 Step -1
                        rowPendingDO = TBLPendingDO.Rows(gh)
                        If ListPrimaryCodeNo_StockIn.Contains(rowPendingDO.Primary_Product) = False Then
                            TBLPendingDO.Rows.Remove(rowPendingDO)
                        End If
                    Next
                End If
                For Each rowPendingDO In TBLPendingDO.Rows
                    Me.vTBLItemStock.DefaultView.RowFilter = "[primary code no] = " + rowPendingDO.Primary_Product.ToString
                    If Me.vTBLItemStock.DefaultView.Count > 0 Then
                        rowStock = Me.vTBLItemStock.DefaultView(0).Row
                        rowStock.qty -= rowPendingDO.Total_Qty
                        If rowStock.qty <= 0 Then
                            Me.vTBLItemStock.Rows.Remove(rowStock)
                        End If
                    End If
                Next
            Else
                'jika dari stock in, qty on hand = qty stock in
                TBL_ProductQty.Rows.Clear()
                For Each row_ProductQty1 In Me.TBL_ProductQty1.Rows
                    row_ProductQty = TBL_ProductQty.Newtblmaster_productqtyRow
                    For jk As Integer = 0 To TBL_ProductQty.Columns.Count - 1
                        row_ProductQty.Item(jk) = row_ProductQty1.Item(jk)
                    Next
                    TBL_ProductQty.Rows.Add(row_ProductQty)
                Next
                For Each row_ProductQty In TBL_ProductQty.Rows
                    rowStock = Me.vTBLItemStock.Newtbltrans_shipping_forecast_itemstockRow

                    With rowStock
                        .primary = Me.GetPrimaryMax_Stock + 1
                        .primary_code_no = row_ProductQty.Primary_Product
                        .primary_transaksi = Me.TxtSF_Primary.Text
                        .primary_unit = row_ProductQty.Primary_Unit
                        .qty = row_ProductQty.Qty
                    End With
                    Me.vTBLItemStock.Rows.Add(rowStock)
                Next
            End If


            'load data pob
            'untuk pob, tampilkan form peringatan bila ada po yang otorisasinya tidak sempurna
            If Me.vFromStockIn = False Then
                'hanya me-load data pob, jika bukan dari stock in purchase
                'jika dari stock in purchase hanya butuh data on hand yang merepresentasikan data stock in purchase
                DTAPOItem.Fill(TBLPOItem)
                CekOtorisasiPO = False
                If Me.vPrimaryStockInPurchaseReference = -1 Then
                    For Each rowPOItem In TBLPOItem.Rows
                        If rowPOItem.Authorization_Level <> rowPOItem.Authorization_Level_Needed Then
                            CekOtorisasiPO = True
                            Exit For
                        End If
                    Next
                End If

                If CekOtorisasiPO = True Then
                    'pindahkan ke dalam data tabel yang akan digunakan untuk pengecekan otorisasi
                    '1 po hanya dipindahkan 1x saja
                    For Each rowPOItem In TBLPOItem.Rows
                        If ListSudah.Contains(rowPOItem.Primary_Transaksi) = False Then
                            ListSudah.Add(rowPOItem.Primary_Transaksi)
                            row_POB = TBL_POB.NewDataTable3Row
                            With row_POB
                                .Authorization_Level = rowPOItem.Authorization_Level
                                .Authorization_Level_Needed = rowPOItem.Authorization_Level_Needed
                                .No_Transaksi = rowPOItem.PO_Number
                                .POB = rowPOItem.POB
                                .Primary_PO = rowPOItem.Primary_Transaksi
                                .Primary_Product = rowPOItem.Primary_Product
                                .Primary_Supplier = rowPOItem.Primary_Supplier
                                .Supplier = rowPOItem.Supplier
                                .Tanggal_Transaksi = rowPOItem.PO_Date
                            End With
                            TBL_POB.Rows.Add(row_POB)
                        End If
                    Next
                    FrmCek.vDT_ItemPOB = TBL_POB
                    FrmCek.BackColorAll = Me.BackColorAll
                    'hanya minta persetujuan jika user ybs, memiliki hak akses atas perhitungan material planning
                    'jika tidak memiliki hak akses, maka tidak ada kepentingan form dimunculkan, ambil datanya langsung saja
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP") = True Then
                        FrmCek.ShowDialog()
                    End If

                    'buang item po yang berasal dari transaksi yang tidak includekan
                    For Each rowDGV In FrmCek.dgvItem.Rows
                        If rowDGV.Cells(FrmCek.ColInclude.Index).Value = Boolean.FalseString Then
                            TBLPOItem.DefaultView.RowFilter = "[Primary Transaksi] = " + rowDGV.Cells(FrmCek.ColPrimaryPO.Index).Value.ToString
                            While TBLPOItem.DefaultView.Count > 0
                                rowPOItem = TBLPOItem.DefaultView(0).Row
                                TBLPOItem.Removetbltrans_po_itemRow(rowPOItem)
                            End While
                        End If
                    Next
                End If

                'pindahkan ke dalam item pob
                'cari primary po max dan min
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                'product exclusion
                If Me.vPrimaryStockInPurchaseReference = -1 Then
                    ijkl.UserCode = Me.UserCode
                    ijkl.ShowWarehouseExclusion = False
                    ijkl.TipeTransaksiExclusion = ClsTransaksiMaterialPlanning.EnumTipeTransaksiExclusion.ShippingForecast
                    ijkl.ShowDialog()
                    DTAGroup.Fill(TBLGroup, Me.UserCode, CInt(ClsTransaksiMaterialPlanning.EnumTipeTransaksiExclusion.ShippingForecast))
                End If
                Me.BuangItemExclude_POB(TBLPOItem, TBLGroup, ListPrimaryCodeNo_StockIn)

                If TBLPOItem.Rows.Count > 0 Then
                    PrimaryPOMin = Double.MaxValue
                    PrimaryPOMax = 0
                    For Each rowPOItem In TBLPOItem.Rows
                        If PrimaryPOMax < rowPOItem.Primary_Transaksi Then
                            PrimaryPOMax = rowPOItem.Primary_Transaksi
                        End If
                        If PrimaryPOMin > rowPOItem.Primary_Transaksi Then
                            PrimaryPOMin = rowPOItem.Primary_Transaksi
                        End If
                    Next
                    'load item delivery
                    DTAItemDeliveryPO.Fill(TBLItemDeliveryPO, PrimaryPOMax, PrimaryPOMin)
                    TBLItemDeliveryPO.DefaultView.Sort = "[delivery date] DESC"
                    'isikan ke dalam tabel item pob
                    For Each rowPOItem In TBLPOItem.Rows
                        sdf += 1
                        TBLItemDeliveryPO.DefaultView.RowFilter = "identifier = '" + rowPOItem.identifier + "'"
                        If TBLItemDeliveryPO.DefaultView.Count = 0 Then
                            'tidak ada delivery control -> set tanggal ke 1 jan 2200
                            rowItemPOB = Me.vTBLItemPOB.Newtbltrans_shipping_forecast_item_pobRow
                            With rowItemPOB
                                .delivery_delay = 0
                                .packing_list_date = #1/1/1950#
                                .packing_list_no_rec = ""
                                .po_date_rec = rowPOItem.PO_Date
                                .po_number_rec = rowPOItem.PO_Number
                                .primary = Me.GetPrimaryMax_ItemPOB + 1
                                .primary_code_no = rowPOItem.Primary_Product
                                .primary_packing_list = -1
                                .primary_po = rowPOItem.Primary_Transaksi
                                .primary_supplier = rowPOItem.Primary_Supplier
                                .primary_supplier_rec = rowPOItem.Primary_Supplier
                                .primary_transaksi = Me.TxtSF_Primary.Text
                                .promise_to_deliver = #1/1/2200#
                                .qty = rowPOItem.POB
                                .remarks = ""
                                .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.POB)
                                .identifier = rowPOItem.identifier
                            End With
                            Me.vTBLItemPOB.Rows.Add(rowItemPOB)
                        Else
                            'ambil item delivery control mulai dari yang paling besar promise to deliverynya
                            While rowPOItem.POB > 0
                                rowItemDeliveryPO = TBLItemDeliveryPO.DefaultView(0).Row
                                rowItemPOB = Me.vTBLItemPOB.Newtbltrans_shipping_forecast_item_pobRow
                                With rowItemPOB
                                    .delivery_delay = rowItemDeliveryPO.delay
                                    .packing_list_date = #1/1/1950#
                                    .packing_list_no_rec = ""
                                    .po_date_rec = rowPOItem.PO_Date
                                    .po_number_rec = rowPOItem.PO_Number
                                    .primary = Me.GetPrimaryMax_ItemPOB + 1
                                    .primary_code_no = rowPOItem.Primary_Product
                                    .primary_packing_list = -1
                                    .primary_po = rowPOItem.Primary_Transaksi
                                    .primary_supplier = rowPOItem.Primary_Supplier
                                    .primary_supplier_rec = rowPOItem.Primary_Supplier
                                    .primary_transaksi = Me.TxtSF_Primary.Text
                                    .promise_to_deliver = rowItemDeliveryPO.delivery_date
                                    If rowPOItem.POB > rowItemDeliveryPO.delivery_qty Then
                                        .qty = rowItemDeliveryPO.delivery_qty
                                        rowPOItem.POB -= rowItemDeliveryPO.delivery_qty
                                        rowItemDeliveryPO.delivery_qty = 0
                                    Else
                                        .qty = rowPOItem.POB
                                        rowPOItem.POB = 0
                                    End If
                                    .remarks = ""
                                    .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.POB)
                                    .identifier = rowPOItem.identifier
                                End With
                                Me.vTBLItemPOB.Rows.Add(rowItemPOB)

                                'persiapan untuk looping selanjutnya
                                If rowItemDeliveryPO.delivery_qty = 0 Then
                                    TBLItemDeliveryPO.Rows.Remove(rowItemDeliveryPO)
                                End If
                            End While
                        End If
                    Next
                End If

                'load data git
                DTAPackingList.Fill(TBLPackingList)
                Me.BuangItemExclude_GIT(TBLPackingList, TBLGroup, ListPrimaryCodeNo_StockIn)
                'isikan data delay
                If TBLPackingList.Rows.Count > 0 Then
                    PrimaryPOMax = 0
                    PrimaryPOMin = Double.MaxValue
                    For Each rowPackingList In TBLPackingList.Rows
                        If PrimaryPOMax < rowPackingList.Primary_PO Then
                            PrimaryPOMax = rowPackingList.Primary_PO
                        End If
                        If PrimaryPOMin > rowPackingList.Primary_PO Then
                            PrimaryPOMin = rowPackingList.Primary_PO
                        End If
                    Next

                    DTAItemDeliveryPO.Fill(TBLItemDeliveryPO, PrimaryPOMax, PrimaryPOMin)
                    TBLPOItem.Rows.Clear()
                    For Each rowPackingList In TBLPackingList.Rows
                        TBLItemDeliveryPO.DefaultView.RowFilter = "[primary transaksi] = " + rowPackingList.Primary_PO.ToString + " AND [Primary Product] = " + rowPackingList.Primary_Product.ToString
                        If TBLItemDeliveryPO.DefaultView.Count > 0 Then
                            rowItemDeliveryPO = TBLItemDeliveryPO.DefaultView(0).Row
                            rowPackingList.delay = rowItemDeliveryPO.delay
                            rowPackingList.identifier = rowItemDeliveryPO.identifier
                        Else
                            'tidak ada janji kirim delay di anggap 0
                            rowPackingList.delay = 0
                            If TBLPOItem.Rows.Count = 0 Then
                                'load ulang data po item jika kosong
                                DTAPOItem.FillBy_PrimaryPOMax_Min(TBLPOItem, PrimaryPOMax, PrimaryPOMin)
                            End If
                            TBLPOItem.DefaultView.RowFilter = "[primary transaksi] = " + rowPackingList.Primary_PO.ToString + " AND [Primary Product] = " + rowPackingList.Primary_Product.ToString
                            rowPOItem = TBLPOItem.DefaultView(0).Row
                            rowPackingList.identifier = rowPOItem.identifier
                        End If
                    Next
                    'pindahkan ke dalam item packinglist
                    For Each rowPackingList In TBLPackingList.Rows
                        rowItemPOB = Me.vTBLItemPOB.Newtbltrans_shipping_forecast_item_pobRow
                        With rowItemPOB
                            .delivery_delay = rowPackingList.delay
                            .packing_list_date = rowPackingList.Tanggal_Transaksi
                            .packing_list_no_rec = rowPackingList.No_Transaksi
                            'untuk po date rec, diisikan dengan tanggal packing list, untuk pengurutan data ketika akan menentukan item delivery
                            .po_date_rec = #1/1/1950#
                            'untuk po number, diisikan dengan supplier packing list number, untuk record jika suatu saat diperlukan
                            .po_number_rec = rowPackingList.Supplier_Packing_List_No
                            .primary = Me.GetPrimaryMax_ItemPOB + 1
                            .primary_code_no = rowPackingList.Primary_Product
                            .primary_packing_list = rowPackingList.Primary_Packing_List
                            .primary_po = rowPackingList.Primary_PO
                            .primary_supplier = rowPackingList.Primary_Supplier
                            .primary_supplier_rec = rowPackingList.Primary_Supplier
                            .primary_transaksi = Me.TxtSF_Primary.Text
                            .promise_to_deliver = rowPackingList.Estimated_Arrival_Date
                            .qty = rowPackingList.Total_Qty
                            .remarks = ""
                            .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.GIT)
                            .identifier = rowPackingList.identifier
                        End With
                        Me.vTBLItemPOB.Rows.Add(rowItemPOB)
                    Next
                End If
            End If

            'load data sob
            If Me.vFromStockIn = False Then
                'hanya include sob, jika menghitung bukan dari stock in purchase
                DTASOItem.Fill(TBLSOItem)
                Me.BuangItemExclude_SOB(TBLSOItem, TBLGroup, ListPrimaryCodeNo_StockIn)
                If TBLSOItem.Rows.Count > 0 Then
                    PrimaryPOMax = 0
                    PrimaryPOMax = Double.MinValue
                    For Each rowSOItem In TBLSOItem.Rows
                        If rowSOItem.Primary_Transaksi > PrimaryPOMax Then
                            PrimaryPOMax = rowSOItem.Primary_Transaksi
                        End If
                        If PrimaryPOMin > rowSOItem.Primary_Transaksi Then
                            PrimaryPOMin = rowSOItem.Primary_Transaksi
                        End If
                    Next
                    DTASOItemDelivery.FillBy_PrimarySOMaxMin(TBLSOItemDelivery, PrimaryPOMin, PrimaryPOMax)

                    TBLSOItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                    For Each rowSOItem In TBLSOItem.Rows
                        TBLSOItemDelivery.DefaultView.RowFilter = "identifier = '" + rowSOItem.identifier + "'"

                        If TBLSOItemDelivery.DefaultView.Count = 0 Then
                            'tambahkan ke dalam item so sesuai sob
                            rowItemSO_SF = Me.vTBLItemDelivery.Newtbltrans_shipping_forecast_itemdeliveryRow
                            With rowItemSO_SF
                                .create_do = 0
                                'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery
                                'dan 0 untuk qty available
                                .qty_available = 0
                                .qty_ready_stock = 0
                                .estimate_delivery = #1/1/2200#
                                .primary = Me.GetPrimaryMax_ItemSOB + 1
                                .primary_code_no = rowSOItem.Primary_Product
                                .primary_cust = rowSOItem.Primary_Customer
                                .primary_cust_rec = rowSOItem.Primary_Customer
                                .primary_priority_level = PrimaryPriorityLevelDefault
                                .primary_so = rowSOItem.Primary_Transaksi
                                .primary_transaksi = Me.TxtSF_Primary.Text
                                .primary_unit = rowSOItem.Primary_Units
                                'tidak ada delivery control -> set tanggal ke 1 jan 2200
                                .promise_to_deliver = #1/1/2200#
                                .qty_sob = rowSOItem.SOB
                                .remark = ""
                                .so_created_by = rowSOItem.Created_By
                                .so_date_rec = rowSOItem.Tanggal_Transaksi
                                .so_number_rec = rowSOItem.No_Transaksi
                                .urutan_priority_level = UrutanPriodityLevelDefault
                                .customer_name = rowSOItem.Customer_Name
                                .Product_Code = rowSOItem.Product_Code
                                .Description = rowSOItem.Description
                                .so_number = rowSOItem.No_Transaksi
                                .so_date = rowSOItem.Tanggal_Transaksi
                                .Cust_PO_No = rowSOItem.Cust_PO_No
                                ._Decimal = rowSOItem._Decimal
                                .Unit_Code = rowSOItem.Unit_Code
                                .Primary_Category = rowSOItem.Primary_Category
                                .Primary_Sub_Brand = rowSOItem.Primary_Sub_Brand
                                .Primary_Brand = rowSOItem.Primary_Brand
                                .primary_product_group = rowSOItem.Primary_Product_Group
                                .identifier = rowSOItem.identifier
                                .booked = 0
                                .no_booking = ""
                                .qty_booked = 0
                                .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB)
                                .tipe_product = rowSOItem.Type_Product
                            End With
                            Me.vTBLItemDelivery.Rows.Add(rowItemSO_SF)
                        Else
                            'ambil item delivery control mulai dari yang paling besar promise to deliverynya
                            While rowSOItem.SOB > 0

                                rowSOItemDelivery = TBLSOItemDelivery.DefaultView(0).Row
                                rowItemSO_SF = Me.vTBLItemDelivery.Newtbltrans_shipping_forecast_itemdeliveryRow
                                With rowItemSO_SF
                                    .create_do = 0
                                    'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery dan 0 untuk qty availabel
                                    .estimate_delivery = #1/1/2200#
                                    .qty_available = 0
                                    .qty_ready_stock = 0
                                    .primary = Me.GetPrimaryMax_ItemSOB + 1
                                    .primary_code_no = rowSOItem.Primary_Product
                                    .primary_cust = rowSOItem.Primary_Customer
                                    .primary_cust_rec = rowSOItem.Primary_Customer
                                    .primary_priority_level = rowSOItemDelivery.delivery_priority
                                    .urutan_priority_level = rowSOItemDelivery.urutan
                                    .primary_so = rowSOItem.Primary_Transaksi
                                    .primary_transaksi = Me.TxtSF_Primary.Text
                                    .primary_unit = rowSOItem.Primary_Units
                                    'tidak ada delivery control -> set tanggal ke 1 jan 2200
                                    .promise_to_deliver = rowSOItemDelivery.delivery_date
                                    If rowSOItem.SOB > rowSOItemDelivery.delivery_qty Then
                                        .qty_sob = rowSOItemDelivery.delivery_qty
                                        rowSOItem.SOB -= rowSOItemDelivery.delivery_qty
                                        rowSOItemDelivery.delivery_qty = 0
                                    Else
                                        .qty_sob = rowSOItem.SOB
                                        rowSOItem.SOB = 0
                                    End If
                                    .remark = ""
                                    .so_created_by = rowSOItem.Created_By
                                    .so_date_rec = rowSOItem.Tanggal_Transaksi
                                    .so_number_rec = rowSOItem.No_Transaksi
                                    .customer_name = rowSOItem.Customer_Name
                                    .Product_Code = rowSOItem.Product_Code
                                    .Description = rowSOItem.Description
                                    .so_number = rowSOItem.No_Transaksi
                                    .so_date = rowSOItem.Tanggal_Transaksi
                                    .Cust_PO_No = rowSOItem.Cust_PO_No
                                    ._Decimal = rowSOItem._Decimal
                                    .Unit_Code = rowSOItem.Unit_Code
                                    .Primary_Category = rowSOItem.Primary_Category
                                    .Primary_Sub_Brand = rowSOItem.Primary_Sub_Brand
                                    .Primary_Brand = rowSOItem.Primary_Brand
                                    .primary_product_group = rowSOItem.Primary_Product_Group
                                    .identifier = rowSOItem.identifier
                                    .booked = 0
                                    .no_booking = ""
                                    .qty_booked = 0
                                    .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB)
                                    .tipe_product = rowSOItem.Type_Product
                                    If rowSOItemDelivery.delivery_qty = 0 Then
                                        TBLSOItemDelivery.Rows.Remove(rowSOItemDelivery)
                                    End If
                                End With
                                Me.vTBLItemDelivery.Rows.Add(rowItemSO_SF)
                            End While
                        End If
                    Next
                End If

                'tandain item sob yang ada inventory bookingnya
                DTABooking.Fill(TBLBooking)
                Me.BuangItemExclude_BookingInv(TBLBooking, TBLGroup, ListPrimaryCodeNo_StockIn)
                For Each rowBooking In TBLBooking.Rows
                    Me.vTBLItemDelivery.DefaultView.RowFilter = "[primary code no] = " + rowBooking.Primary_Product.ToString + " AND [primary so] = " + rowBooking.Primary_SO.ToString
                    Me.vTBLItemDelivery.DefaultView.Sort = "[promise to deliver] ASC, [urutan priority level] ASC, [so date rec] ASC, [primary so] ASC"
                    IndexAcuan = 0
                    While IndexAcuan < Me.vTBLItemDelivery.DefaultView.Count And rowBooking.Qty > 0
                        rowItemSO_SF = Me.vTBLItemDelivery.DefaultView(IndexAcuan).Row

                        If rowItemSO_SF.qty_sob - rowItemSO_SF.qty_booked >= rowBooking.Qty Then
                            rowItemSO_SF.qty_booked += rowBooking.Qty
                            rowBooking.Qty = 0
                        Else
                            rowBooking.Qty -= (rowItemSO_SF.qty_sob - rowItemSO_SF.qty_booked)
                            rowItemSO_SF.qty_booked = rowItemSO_SF.qty_sob
                            IndexAcuan += 1
                        End If

                        rowItemSO_SF.booked = 1
                        If rowItemSO_SF.no_booking = "" Then
                            rowItemSO_SF.no_booking = rowBooking.No_Transaksi
                        Else
                            If rowItemSO_SF.no_booking.Contains(rowBooking.No_Transaksi) = False Then
                                rowItemSO_SF.no_booking = rowItemSO_SF.no_booking + ", " + rowBooking.No_Transaksi
                            End If
                        End If
                    End While
                Next
            End If

            'tambahkan item sob dengan purchase request
            DTAPurchReqItem.Fill(TBLPurchReqItem)
            Me.BuangItemExclude_PurcReq(TBLPurchReqItem, TBLGroup, ListPrimaryCodeNo_StockIn)
            If TBLPurchReqItem.Rows.Count > 0 Then
                PrimaryPOMax = 0
                PrimaryPOMax = Double.MinValue
                For Each rowPurchReqItem In TBLPurchReqItem.Rows
                    If rowPurchReqItem.primary_transaksi > PrimaryPOMax Then
                        PrimaryPOMax = rowPurchReqItem.primary_transaksi
                    End If
                    If PrimaryPOMin > rowPurchReqItem.primary_transaksi Then
                        PrimaryPOMin = rowPurchReqItem.primary_transaksi
                    End If
                Next
                DTAPurchReqItemDelivery.Fill(TBLPurchReqItemDelivery, PrimaryPOMax, PrimaryPOMin)

                TBLPurchReqItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                For Each rowPurchReqItem In TBLPurchReqItem.Rows
                    TBLPurchReqItemDelivery.DefaultView.RowFilter = "identifier = '" + rowPurchReqItem.identifier + "'"

                    If TBLPurchReqItemDelivery.DefaultView.Count = 0 Then
                        'tambahkan ke dalam item sob sesuai Purc Req OB
                        rowItemSO_SF = Me.vTBLItemDelivery.Newtbltrans_shipping_forecast_itemdeliveryRow
                        With rowItemSO_SF
                            .create_do = 0
                            'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery
                            'dan 0 untuk qty available
                            .qty_available = 0
                            .qty_ready_stock = 0
                            .estimate_delivery = #1/1/2200#
                            .primary = Me.GetPrimaryMax_ItemSOB + 1
                            .primary_code_no = rowPurchReqItem.primary_code_no
                            .primary_cust = rowPurchReqItem.primary_ref_party
                            .primary_cust_rec = rowPurchReqItem.primary_ref_party
                            .primary_priority_level = PrimaryPriorityLevelDefault
                            .primary_so = rowPurchReqItem.primary_transaksi
                            .primary_transaksi = Me.TxtSF_Primary.Text
                            .primary_unit = rowPurchReqItem.primary_unit
                            'tidak ada delivery control -> set tanggal ke 1 jan 2200
                            .promise_to_deliver = #1/1/2200#
                            .qty_sob = rowPurchReqItem.OB
                            .remark = rowPurchReqItem.remarks_item
                            .so_created_by = rowPurchReqItem.created_by
                            .so_date_rec = rowPurchReqItem.tanggal
                            .so_number_rec = rowPurchReqItem.no_transaksi
                            .urutan_priority_level = UrutanPriodityLevelDefault
                            .customer_name = rowPurchReqItem.Ref_Party_Desc
                            .Product_Code = rowPurchReqItem.Product_Code
                            .Description = rowPurchReqItem.Description
                            .so_number = rowPurchReqItem.no_transaksi
                            .so_date = rowPurchReqItem.tanggal
                            .Cust_PO_No = ""
                            ._Decimal = rowPurchReqItem._Decimal
                            .Unit_Code = rowPurchReqItem.Unit_Code
                            .Primary_Category = rowPurchReqItem.Primary_Category
                            .Primary_Sub_Brand = rowPurchReqItem.Primary_Sub_Brand
                            .Primary_Brand = rowPurchReqItem.Primary_Brand
                            .primary_product_group = rowPurchReqItem.Primary_Product_Group
                            .identifier = rowPurchReqItem.identifier
                            .booked = 0
                            .no_booking = ""
                            .qty_booked = 0
                            .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req)
                            .tipe_product = rowPurchReqItem.Type
                        End With
                        Me.vTBLItemDelivery.Rows.Add(rowItemSO_SF)
                    Else
                        'ambil item delivery control mulai dari yang paling besar promise to deliverynya
                        While rowPurchReqItem.OB > 0
                            rowPurchReqItemDelivery = TBLPurchReqItemDelivery.DefaultView(0).Row
                            rowItemSO_SF = Me.vTBLItemDelivery.Newtbltrans_shipping_forecast_itemdeliveryRow
                            With rowItemSO_SF
                                .create_do = 0
                                'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery dan 0 untuk qty availabel
                                .estimate_delivery = #1/1/2200#
                                .qty_available = 0
                                .qty_ready_stock = 0
                                .primary = Me.GetPrimaryMax_ItemSOB + 1
                                .primary_code_no = rowPurchReqItem.primary_code_no
                                .primary_cust = rowPurchReqItem.primary_ref_party
                                .primary_cust_rec = rowPurchReqItem.primary_ref_party
                                .primary_priority_level = rowPurchReqItemDelivery.delivery_priority
                                .urutan_priority_level = rowPurchReqItemDelivery.urutan
                                .primary_so = rowPurchReqItem.primary_transaksi
                                .primary_transaksi = Me.TxtSF_Primary.Text
                                .primary_unit = rowPurchReqItem.primary_unit
                                'tidak ada delivery control -> set tanggal ke 1 jan 2200
                                .promise_to_deliver = rowPurchReqItemDelivery.delivery_date
                                If rowPurchReqItem.OB > rowPurchReqItemDelivery.delivery_qty Then
                                    .qty_sob = rowPurchReqItemDelivery.delivery_qty
                                    rowPurchReqItem.OB -= rowPurchReqItemDelivery.delivery_qty
                                    rowPurchReqItemDelivery.delivery_qty = 0
                                Else
                                    .qty_sob = rowPurchReqItem.OB
                                    rowPurchReqItem.OB = 0
                                End If
                                .remark = rowPurchReqItem.remarks_item
                                .so_created_by = rowPurchReqItem.created_by
                                .so_date_rec = rowPurchReqItem.tanggal
                                .so_number_rec = rowPurchReqItem.no_transaksi
                                .customer_name = rowPurchReqItem.Ref_Party_Desc
                                .Product_Code = rowPurchReqItem.Product_Code
                                .Description = rowPurchReqItem.Description
                                .so_number = rowPurchReqItem.no_transaksi
                                .so_date = rowPurchReqItem.tanggal
                                .Cust_PO_No = ""
                                ._Decimal = rowPurchReqItem._Decimal
                                .Unit_Code = rowPurchReqItem.Unit_Code
                                .Primary_Category = rowPurchReqItem.Primary_Category
                                .Primary_Sub_Brand = rowPurchReqItem.Primary_Sub_Brand
                                .Primary_Brand = rowPurchReqItem.Primary_Brand
                                .primary_product_group = rowPurchReqItem.Primary_Product_Group
                                .identifier = rowPurchReqItem.identifier
                                .booked = 0
                                .no_booking = ""
                                .qty_booked = 0
                                .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req)
                                .tipe_product = rowPurchReqItem.Type
                                If rowPurchReqItemDelivery.delivery_qty = 0 Then
                                    TBLPurchReqItemDelivery.Rows.Remove(rowPurchReqItemDelivery)
                                End If
                            End With
                            Me.vTBLItemDelivery.Rows.Add(rowItemSO_SF)
                        End While
                    End If
                Next
            End If

            'pasangkan data sob dengan item stock dan pob, tentukan estimate delivery
            'delivery qty hanya bisa tidak sama dengan qty sob, bila barang tersedia sebagian dan belum ada order.
            'pasangkan terlebih dahulu yang ada inventory bookingnya
            'baru kemudia sisanya
            ListPrimaryCodeNo.Clear()
            For Each rowItemSO_SF In vTBLItemDelivery
                If rowItemSO_SF.tipe_product = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                    If Me.vFromStockIn = False Then
                        'untuk service, jika bukan dari stock in purchase,
                        'maka selalu dianggap waiting
                        rowItemSO_SF.qty_available = 0
                        rowItemSO_SF.qty_ready_stock = 0
                        'waiting, maka tandai estimate delivery menjadi 1 jan 2200
                        rowItemSO_SF.estimate_delivery = #1/1/2200#
                    Else
                        'jika dari stock in purchase, maka hitung normal, tetapi hanya memperhitungkan sejumlah stock in saja
                        'yang direpresentasikan dala qty on hand
                        If ListPrimaryCodeNo.Contains(rowItemSO_SF.primary_code_no) = False Then
                            ListPrimaryCodeNo.Add(rowItemSO_SF.primary_code_no)
                        End If
                    End If
                Else
                    If ListPrimaryCodeNo.Contains(rowItemSO_SF.primary_code_no) = False Then
                        ListPrimaryCodeNo.Add(rowItemSO_SF.primary_code_no)
                    End If
                End If
            Next
            For Each rowItemPOB In Me.vTBLItemPOB.Rows
                'set parameter item pob yang akan digunakan pada perhitungan delivery date
                rowItemPOB.available_date = rowItemPOB.promise_to_deliver + TimeSpan.FromDays(rowItemPOB.delivery_delay)
                rowItemPOB.qty_used = 0
            Next
            For Each rowStock In Me.vTBLItemStock.Rows
                rowStock.qty_used = 0
            Next

            For Each PrimaryCodeNo As Double In ListPrimaryCodeNo
                'filter
                Me.vTBLItemStock.DefaultView.RowFilter = "[primary code no] = " + PrimaryCodeNo.ToString
                Me.vTBLItemPOB.DefaultView.RowFilter = "[primary code no] = " + PrimaryCodeNo.ToString
                Me.vTBLItemPOB.DefaultView.Sort = "[available date] ASC, [promise to deliver] ASC, [po date rec] ASC, [primary po] ASC"
                Me.vTBLItemDelivery.DefaultView.RowFilter = "[primary code no] = " + PrimaryCodeNo.ToString
                Me.vTBLItemDelivery.DefaultView.Sort = "[booked] DESC, [promise to deliver] ASC, [urutan priority level] ASC, [so date rec] ASC, [primary so] ASC"

                'kembalikan parameter ke nilai awal
                IndexItemDelivery = 0
                QtyReadyStock = 0

                'pasangkan dengan yang ready stock dahulu
                If Me.vTBLItemStock.DefaultView.Count > 0 Then
                    rowStock = Me.vTBLItemStock.DefaultView(0).Row
                    While rowStock.qty - rowStock.qty_used > 0 And vTBLItemDelivery.DefaultView.Count > IndexItemDelivery
                        rowItemSO_SF = Me.vTBLItemDelivery.DefaultView(IndexItemDelivery).Row
                        If Me.vFromStockIn = False Then
                            'jika bukan dari stock in, maka pr-ob tidak memperhitungkan on hand, hanya sob yang
                            'memperhitungkan on hand
                            If rowItemSO_SF.tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req) Then
                                IndexItemDelivery += 1
                            Else
                                If rowStock.qty - rowStock.qty_used >= rowItemSO_SF.qty_sob Then
                                    rowStock.qty_used += rowItemSO_SF.qty_sob
                                    'karena ready stock, maka qty availabel = qty sob
                                    rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                                    rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                                    'karena ready stock, maka tandai estimate delivery menjadi ASAP -> 1 jan 1950
                                    rowItemSO_SF.estimate_delivery = #1/1/1950#
                                    'naikan index item delivery, cek item delivery selanjutnya
                                    IndexItemDelivery += 1
                                Else
                                    QtyReadyStock = rowStock.qty - rowStock.qty_used
                                    rowStock.qty_used = rowStock.qty
                                    Exit While
                                End If
                            End If
                        Else
                            'jika dari stock in, maka semua memperhitungkan on hand
                            If rowStock.qty - rowStock.qty_used >= rowItemSO_SF.qty_sob Then
                                rowStock.qty_used += rowItemSO_SF.qty_sob
                                'karena ready stock, maka qty availabel = qty sob
                                rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                                rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                                'karena ready stock, maka tandai estimate delivery menjadi ASAP -> 1 jan 1950
                                rowItemSO_SF.estimate_delivery = #1/1/1950#
                                'naikan index item delivery, cek item delivery selanjutnya
                                IndexItemDelivery += 1
                            Else
                                QtyReadyStock = rowStock.qty - rowStock.qty_used
                                rowStock.qty_used = rowStock.qty
                                Exit While
                            End If
                        End If

                    End While
                End If

                'pasangkan dengan item po, hanya jika bukan dari stock in
                QtyPOReady = 0
                If Me.vFromStockIn = False Then
                    'editted by iwan: 2016-09-08, ini bikin masalah
                    'IndexItemDelivery = 0 '-> ulang lagi dari yang paling atas, hanya saja untuk yang rowItemSO_SF.qty_available < rowItemSO_SF.qty_sob
                    If Me.vTBLItemPOB.DefaultView.Count > 0 Then
                        IndexItemPOB = 0
                        While IndexItemPOB < vTBLItemPOB.DefaultView.Count And IndexItemDelivery < vTBLItemDelivery.DefaultView.Count
                            rowItemPOB = Me.vTBLItemPOB.DefaultView(IndexItemPOB).Row
                            rowItemSO_SF = Me.vTBLItemDelivery.DefaultView(IndexItemDelivery).Row
                            'jika pr-ob, maka tidak boleh meperhitungkan ready stock
                            If rowItemSO_SF.tipe <> CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req) Then
                                QtyReadyStock1 = QtyReadyStock
                            Else
                                QtyReadyStock1 = 0
                            End If
                            If QtyReadyStock1 + QtyPOReady + rowItemPOB.qty - rowItemPOB.qty_used >= rowItemSO_SF.qty_sob Then
                                'jika item yang tersedia lebih besar dari item delivery, maka 
                                'set delivery date + qty available delivery + pindah ke baris item delivery selanjutnya
                                rowItemSO_SF.estimate_delivery = rowItemPOB.available_date
                                rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                                IndexItemDelivery += 1

                                'kurangkan qty on hand dan qty po ready
                                Dummy = rowItemSO_SF.qty_sob
                                If QtyReadyStock1 >= Dummy Then
                                    rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                                    QtyReadyStock1 -= Dummy
                                    Dummy = 0
                                Else
                                    Dummy -= QtyReadyStock1
                                    rowItemSO_SF.qty_ready_stock = QtyReadyStock1
                                    QtyReadyStock1 = 0

                                    'jika ada masih ada sisa dummy, maka kurangkan dengan qty po ready (dari pob baris sebelumnya)
                                    If QtyPOReady >= Dummy Then
                                        QtyPOReady -= Dummy
                                        Dummy = 0
                                    Else
                                        Dummy -= QtyPOReady
                                        QtyPOReady = 0
                                    End If
                                End If

                                'tambahkan qty used, pada pob baris ybs
                                rowItemPOB.qty_used += Dummy

                                'jika semua qty sudah habis dipakai, maka pindah ke baris item pob selanjutnya
                                If rowItemPOB.qty = rowItemPOB.qty_used Then
                                    IndexItemPOB += 1
                                End If
                            Else
                                'jika qty item delivery lebih besar dari qty po ready + qty po,ada 2 kemungkinan:
                                If IndexItemPOB + 1 < vTBLItemPOB.DefaultView.Count Then
                                    '1. tambahkan qty po dari baris selanjut, karena dari pob, maka qty ready stock tidak di tambahkan
                                    QtyPOReady += rowItemPOB.qty - rowItemPOB.qty_used
                                    rowItemPOB.qty_used = rowItemPOB.qty
                                    IndexItemPOB += 1
                                Else
                                    '2. tidak ada baris selanjutnya, maka qty availble hanya sejumlah yang ada,
                                    'estimate delivery tidak dikehui kapan, tapi available qty ada
                                    rowItemSO_SF.qty_available = rowItemPOB.qty - rowItemPOB.qty_used + QtyPOReady + QtyReadyStock1
                                    rowItemSO_SF.estimate_delivery = #1/1/2200#
                                    'perhitungan ready stock
                                    If QtyReadyStock1 >= rowItemSO_SF.qty_sob Then
                                        rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                                        QtyReadyStock1 -= rowItemSO_SF.qty_sob
                                    Else
                                        rowItemSO_SF.qty_ready_stock = QtyReadyStock1
                                        QtyReadyStock1 = 0
                                    End If
                                    'index item pob tetap dinaikan, untuk menghentikan looping perhitungan
                                    IndexItemPOB += 1
                                    'tandai bahwa semua item pob sudah dipergunakan
                                    rowItemPOB.qty_used = rowItemPOB.qty
                                    QtyPOReady = 0
                                End If
                            End If
                            'kembalikan posisi qty ready stock
                            If rowItemSO_SF.tipe <> CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req) Then
                                QtyReadyStock = QtyReadyStock1
                            End If
                        End While
                    End If
                End If

                'tambahkan qty po ready sisa ke dalam item sob/pr-ob terakhir
                If QtyPOReady > 0 Or QtyReadyStock > 0 Then
                    rowItemSO_SF = Me.vTBLItemDelivery.DefaultView(IndexItemDelivery).Row
                    rowItemSO_SF.qty_available = QtyPOReady + QtyReadyStock
                    rowItemSO_SF.qty_ready_stock = QtyReadyStock
                    rowItemSO_SF.estimate_delivery = #1/1/2200#
                End If
                If vFromStockIn = True Then
                    If QtyReadyStock > 0 Then
                        rowItemSO_SF = Me.vTBLItemDelivery.DefaultView(IndexItemDelivery).Row
                        rowItemSO_SF.qty_available = QtyReadyStock
                        rowItemSO_SF.qty_ready_stock = QtyReadyStock
                        rowItemSO_SF.estimate_delivery = #1/1/2200#
                    End If
                End If
            Next

            Me.llbSF_CekAll.Text = "Check All"
        Catch ex As Exception
            MsgBox(ex.Message + ": sdf - " + sdf.ToString, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryCodeNo = Nothing
            DTA_ProductQty.Dispose()
            TBL_ProductQty.Dispose()
            DTAPOItem.Dispose()
            TBLPOItem.Dispose()
            TBL_POB.Dispose()
            ListSudah = Nothing
            DTAItemDeliveryPO.Dispose()
            TBLItemDeliveryPO.Dispose()
            DTAPackingList.Dispose()
            TBLPackingList.Dispose()
            DTASOItem.Dispose()
            TBLSOItem.Dispose()
            DTASOItemDelivery.Dispose()
            TBLSOItemDelivery.Dispose()
            DTAPendingDO.Dispose()
            TBLPendingDO.Dispose()
            DTABooking.Dispose()
            TBLBooking.Dispose()
            DTAPurchReqItem.Dispose()
            TBLPurchReqItem.Dispose()
            DTAPurchReqItemDelivery.Dispose()
            TBLPurchReqItemDelivery.Dispose()
            ListPrimaryCodeNo_StockIn = Nothing

            DTAGroup.Dispose()
            TBLGroup.Dispose()
            ijkl.Dispose()
        End Try
    End Sub

    Private Sub BuangItemExclude_POB(ByVal TBLPOItem As DseTransaksiShippingForecast.tbltrans_po_itemDataTable, _
    ByVal TBLGroup As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionDataTable, _
    ByVal ListPrimaryCodeNo_StockIn As List(Of Double))
        Dim rowPOItem As DseTransaksiShippingForecast.tbltrans_po_itemRow
        Dim ProductLevel As ClsTransaksiMaterialPlanningGroup.ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand
        Dim Filter As String = ""
        Dim TipeExclusion As ClsTransaksiMaterialPlanning.EnumTipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude
        Dim TblDummy As New DseTransaksiShippingForecast.tbltrans_po_itemDataTable
        Dim rowDummy As DseTransaksiShippingForecast.tbltrans_po_itemRow
        Dim rowGroup As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionRow

        Try
            If TBLGroup.Rows.Count > 0 Then
                rowGroup = TBLGroup.Rows(0)
                TipeExclusion = rowGroup.tipe_exclusion

                If TipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude Then
                    For Each rowGroup In TBLGroup.Rows
                        ProductLevel = rowGroup.level
                        With rowGroup
                            If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                                Filter = "[Primary Brand] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                                Filter = "[Primary Category] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                                Filter = "[Primary Product] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                                Filter = "[Primary Product Group] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                                Filter = "[Primary Sub Brand] = " + .primary_acuan.ToString
                            End If

                            TBLPOItem.DefaultView.RowFilter = Filter
                            While TBLPOItem.DefaultView.Count > 0
                                rowPOItem = TBLPOItem.DefaultView(0).Row
                                TBLPOItem.Rows.Remove(rowPOItem)
                            End While
                        End With
                    Next
                Else
                    TblDummy.Rows.Clear()

                    'pindahkan ke tabel dummy
                    For Each rowGroup In TBLGroup.Rows
                        ProductLevel = rowGroup.level
                        With rowGroup
                            If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                                Filter = "[Primary Brand] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                                Filter = "[Primary Category] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                                Filter = "[Primary Product] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                                Filter = "[Primary Product Group] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                                Filter = "[Primary Sub Brand] = " + .primary_acuan.ToString
                            End If

                            TBLPOItem.DefaultView.RowFilter = Filter
                            While TBLPOItem.DefaultView.Count > 0
                                rowPOItem = TBLPOItem.DefaultView(0).Row

                                rowDummy = TblDummy.Newtbltrans_po_itemRow
                                For hj As Integer = 0 To TblDummy.Columns.Count - 1
                                    rowDummy.Item(hj) = rowPOItem.Item(hj)
                                Next
                                TblDummy.Rows.Add(rowDummy)

                                TBLPOItem.Rows.Remove(rowPOItem)
                            End While
                        End With
                    Next
                    'kosongkan tabel asal
                    TBLPOItem.Rows.Clear()

                    'pindahkan kembali ke tabel asal
                    For Each rowDummy In TblDummy.Rows
                        rowPOItem = TBLPOItem.Newtbltrans_po_itemRow
                        For hj As Integer = 0 To TblDummy.Columns.Count - 1
                            rowPOItem.Item(hj) = rowDummy.Item(hj)
                        Next
                        TBLPOItem.Rows.Add(rowPOItem)
                    Next
                End If
            End If

            If ListPrimaryCodeNo_StockIn.Count > 0 Then
                For ij As Double = TBLPOItem.Rows.Count - 1 To 0 Step -1
                    rowPOItem = TBLPOItem.Rows(ij)
                    If ListPrimaryCodeNo_StockIn.Contains(rowPOItem.Primary_Product) = False Then
                        TBLPOItem.Rows.Remove(rowPOItem)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblDummy = Nothing
        End Try
    End Sub

    Private Sub BuangItemExclude_GIT(ByVal TBLGITItem As DseTransaksiShippingForecast.tblpackinglist_itemDataTable, _
    ByVal TBLGroup As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionDataTable, _
    ByVal ListPrimaryCodeNo_StockIn As List(Of Double))
        Dim rowGITItem As DseTransaksiShippingForecast.tblpackinglist_itemRow
        Dim ProductLevel As ClsTransaksiMaterialPlanningGroup.ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand
        Dim Filter As String = ""
        Dim TipeExclusion As ClsTransaksiMaterialPlanning.EnumTipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude
        Dim TblDummy As New DseTransaksiShippingForecast.tblpackinglist_itemDataTable
        Dim rowDummy As DseTransaksiShippingForecast.tblpackinglist_itemRow
        Dim rowGroup As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionRow

        Try
            If TBLGroup.Rows.Count > 0 Then
                rowGroup = TBLGroup.Rows(0)
                TipeExclusion = rowGroup.tipe_exclusion

                If TipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude Then
                    For Each rowGroup In TBLGroup.Rows
                        ProductLevel = rowGroup.level
                        With rowGroup
                            If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                                Filter = "[Primary Brand] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                                Filter = "[Primary Category] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                                Filter = "[Primary Product] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                                Filter = "[Primary Product Group] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                                Filter = "[Primary Sub Brand] = " + .primary_acuan.ToString
                            End If

                            TBLGITItem.DefaultView.RowFilter = Filter
                            While TBLGITItem.DefaultView.Count > 0
                                rowGITItem = TBLGITItem.DefaultView(0).Row
                                TBLGITItem.Rows.Remove(rowGITItem)
                            End While
                        End With
                    Next
                Else
                    TblDummy.Rows.Clear()

                    'pindahkan ke tabel dummy
                    For Each rowGroup In TBLGroup.Rows
                        ProductLevel = rowGroup.level
                        With rowGroup
                            If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                                Filter = "[Primary Brand] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                                Filter = "[Primary Category] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                                Filter = "[Primary Product] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                                Filter = "[Primary Product Group] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                                Filter = "[Primary Sub Brand] = " + .primary_acuan.ToString
                            End If

                            TBLGITItem.DefaultView.RowFilter = Filter
                            While TBLGITItem.DefaultView.Count > 0
                                rowGITItem = TBLGITItem.DefaultView(0).Row

                                rowDummy = TblDummy.Newtblpackinglist_itemRow
                                For hj As Integer = 0 To TblDummy.Columns.Count - 1
                                    rowDummy.Item(hj) = rowGITItem.Item(hj)
                                Next
                                TblDummy.Rows.Add(rowDummy)

                                TBLGITItem.Rows.Remove(rowGITItem)
                            End While
                        End With
                    Next
                    'kosongkan tabel asal
                    TBLGITItem.Rows.Clear()

                    'pindahkan kembali ke tabel asal
                    For Each rowDummy In TblDummy.Rows
                        rowGITItem = TBLGITItem.Newtblpackinglist_itemRow
                        For hj As Integer = 0 To TblDummy.Columns.Count - 1
                            rowGITItem.Item(hj) = rowDummy.Item(hj)
                        Next
                        TBLGITItem.Rows.Add(rowGITItem)
                    Next
                End If
            End If

            If ListPrimaryCodeNo_StockIn.Count > 0 Then
                For ij As Double = TBLGITItem.Rows.Count - 1 To 0 Step -1
                    rowGITItem = TBLGITItem.Rows(ij)
                    If ListPrimaryCodeNo_StockIn.Contains(rowGITItem.Primary_Product) = False Then
                        TBLGITItem.Rows.Remove(rowGITItem)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblDummy = Nothing
        End Try
    End Sub

    Private Sub BuangItemExclude_SOB(ByVal TBLSO_Item As DseTransaksiShippingForecast.tbltrans_so_itemDataTable, _
   ByVal TBLGroup As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionDataTable, _
    ByVal ListPrimaryCodeNo_StockIn As List(Of Double))
        Dim rowSO_Item As DseTransaksiShippingForecast.tbltrans_so_itemRow
        Dim ProductLevel As ClsTransaksiMaterialPlanningGroup.ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand
        Dim Filter As String = ""
        Dim TipeExclusion As ClsTransaksiMaterialPlanning.EnumTipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude
        Dim TblDummy As New DseTransaksiShippingForecast.tbltrans_so_itemDataTable
        Dim rowDummy As DseTransaksiShippingForecast.tbltrans_so_itemRow
        Dim rowGroup As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionRow

        Try
            If TBLGroup.Rows.Count > 0 Then
                rowGroup = TBLGroup.Rows(0)
                TipeExclusion = rowGroup.tipe_exclusion

                If TipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude Then
                    For Each rowGroup In TBLGroup.Rows
                        ProductLevel = rowGroup.level
                        With rowGroup
                            If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                                Filter = "[Primary Brand] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                                Filter = "[Primary Category] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                                Filter = "[Primary Product] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                                Filter = "[Primary Product Group] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                                Filter = "[Primary Sub Brand] = " + .primary_acuan.ToString
                            End If

                            TBLSO_Item.DefaultView.RowFilter = Filter
                            While TBLSO_Item.DefaultView.Count > 0
                                rowSO_Item = TBLSO_Item.DefaultView(0).Row
                                TBLSO_Item.Rows.Remove(rowSO_Item)
                            End While
                        End With
                    Next
                Else
                    TblDummy.Rows.Clear()

                    'pindahkan ke tabel dummy
                    For Each rowGroup In TBLGroup.Rows
                        ProductLevel = rowGroup.level
                        With rowGroup
                            If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                                Filter = "[Primary Brand] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                                Filter = "[Primary Category] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                                Filter = "[Primary Product] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                                Filter = "[Primary Product Group] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                                Filter = "[Primary Sub Brand] = " + .primary_acuan.ToString
                            End If

                            TBLSO_Item.DefaultView.RowFilter = Filter
                            While TBLSO_Item.DefaultView.Count > 0
                                rowSO_Item = TBLSO_Item.DefaultView(0).Row

                                rowDummy = TblDummy.Newtbltrans_so_itemRow
                                For hj As Integer = 0 To TblDummy.Columns.Count - 1
                                    rowDummy.Item(hj) = rowSO_Item.Item(hj)
                                Next
                                TblDummy.Rows.Add(rowDummy)

                                TBLSO_Item.Rows.Remove(rowSO_Item)
                            End While
                        End With
                    Next
                    'kosongkan tabel asal
                    TBLSO_Item.Rows.Clear()

                    'pindahkan kembali ke tabel asal
                    For Each rowDummy In TblDummy.Rows
                        rowSO_Item = TBLSO_Item.Newtbltrans_so_itemRow
                        For hj As Integer = 0 To TblDummy.Columns.Count - 1
                            rowSO_Item.Item(hj) = rowDummy.Item(hj)
                        Next
                        TBLSO_Item.Rows.Add(rowSO_Item)
                    Next
                End If
            End If

            If ListPrimaryCodeNo_StockIn.Count > 0 Then
                For ij As Double = TBLSO_Item.Rows.Count - 1 To 0 Step -1
                    rowSO_Item = TBLSO_Item.Rows(ij)
                    If ListPrimaryCodeNo_StockIn.Contains(rowSO_Item.Primary_Product) = False Then
                        TBLSO_Item.Rows.Remove(rowSO_Item)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblDummy = Nothing
        End Try
    End Sub

    Private Sub FollowUp()
        Dim DTAItemStockIn As New DseTransaksiShippingForecastTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TBLItemStockIn As New DseTransaksiShippingForecast.tbltransaksi_stockinout_itemDataTable
        Dim rowItemStockIn As DseTransaksiShippingForecast.tbltransaksi_stockinout_itemRow
        Dim IndexDGV As Integer = 0
        Dim rowDel As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow

        Try
            DTAItemStockIn.Connection.ConnectionString = Core_Module.GetConstring
            'check list item yang ready sejumlah item finish goods pada stock in purchase
            DTAItemStockIn.Fill(TBLItemStockIn, Me.vPrimaryStockInPurchaseReference)
            For Each rowItemStockIn In TBLItemStockIn.Rows
                For Each rowDGV As DataGridViewRow In Me.dgvSF.Rows
                    If rowItemStockIn.Qty <= 0 Then
                        Exit For
                    Else
                        If rowItemStockIn.Primary_Product = rowDGV.Cells(Me.ColSF_PrimaryCodeNo.Index).Value Then
                            If rowDGV.Cells(Me.ColSF_DeliveryStatus.Index).Value.ToString.ToLower = "ready" Or rowDGV.Cells(Me.ColSF_DeliveryStatus.Index).Value.ToString.ToLower = "partial" Then
                                rowDGV.Cells(Me.ColSF_CreateDO.Index).Value = Boolean.TrueString
                                Me.vTBLItemDelivery.DefaultView.RowFilter = "primary = " + rowDGV.Cells(Me.ColSF_Primary.Index).Value.ToString
                                rowDel = Me.vTBLItemDelivery.DefaultView(0).Row
                                rowDel.create_do = 1

                                If rowDGV.Cells(Me.ColSF_QtyReadyStock.Index).Value = 0 Then
                                    rowItemStockIn.Qty -= rowDGV.Cells(Me.ColSF_QtySOB.Index).Value
                                Else
                                    rowItemStockIn.Qty -= rowDGV.Cells(Me.ColSF_QtyReadyStock.Index).Value
                                End If
                            End If
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemStockIn.Dispose()
            TBLItemStockIn.Dispose()
        End Try
    End Sub

    Private Sub BuangItemExclude_BookingInv(ByVal TBLBooking_Item As DseTransaksiShippingForecast.tbltrans_booking_invitemDataTable, _
   ByVal TBLGroup As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionDataTable, _
    ByVal ListPrimaryCodeNo_StockIn As List(Of Double))
        Dim rowBooking_Item As DseTransaksiShippingForecast.tbltrans_booking_invitemRow
        Dim ProductLevel As ClsTransaksiMaterialPlanningGroup.ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand
        Dim Filter As String = ""
        Dim TipeExclusion As ClsTransaksiMaterialPlanning.EnumTipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude
        Dim TblDummy As New DseTransaksiShippingForecast.tbltrans_booking_invitemDataTable
        Dim rowDummy As DseTransaksiShippingForecast.tbltrans_booking_invitemRow
        Dim rowGroup As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionRow

        Try
            If TBLGroup.Rows.Count > 0 Then
                rowGroup = TBLGroup.Rows(0)
                TipeExclusion = rowGroup.tipe_exclusion

                If TipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude Then
                    For Each rowGroup In TBLGroup.Rows
                        ProductLevel = rowGroup.level
                        With rowGroup
                            If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                                Filter = "[Primary Brand] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                                Filter = "[Primary Category] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                                Filter = "[Primary Product] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                                Filter = "[Primary Product Group] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                                Filter = "[Primary Sub Brand] = " + .primary_acuan.ToString
                            End If

                            TBLBooking_Item.DefaultView.RowFilter = Filter
                            While TBLBooking_Item.DefaultView.Count > 0
                                rowBooking_Item = TBLBooking_Item.DefaultView(0).Row
                                TBLBooking_Item.Rows.Remove(rowBooking_Item)
                            End While
                        End With
                    Next
                Else
                    TblDummy.Rows.Clear()

                    'pindahkan ke tabel dummy
                    For Each rowGroup In TBLGroup.Rows
                        ProductLevel = rowGroup.level
                        With rowGroup
                            If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                                Filter = "[Primary Brand] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                                Filter = "[Primary Category] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                                Filter = "[Primary Product] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                                Filter = "[Primary Product Group] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                                Filter = "[Primary Sub Brand] = " + .primary_acuan.ToString
                            End If

                            TBLBooking_Item.DefaultView.RowFilter = Filter
                            While TBLBooking_Item.DefaultView.Count > 0
                                rowBooking_Item = TBLBooking_Item.DefaultView(0).Row

                                rowDummy = TblDummy.Newtbltrans_booking_invitemRow
                                For hj As Integer = 0 To TblDummy.Columns.Count - 1
                                    rowDummy.Item(hj) = rowBooking_Item.Item(hj)
                                Next
                                TblDummy.Rows.Add(rowDummy)

                                TBLBooking_Item.Rows.Remove(rowBooking_Item)
                            End While
                        End With
                    Next
                    'kosongkan tabel asal
                    TBLBooking_Item.Rows.Clear()

                    'pindahkan kembali ke tabel asal
                    For Each rowDummy In TblDummy.Rows
                        rowBooking_Item = TBLBooking_Item.Newtbltrans_booking_invitemRow
                        For hj As Integer = 0 To TblDummy.Columns.Count - 1
                            rowBooking_Item.Item(hj) = rowDummy.Item(hj)
                        Next
                        TBLBooking_Item.Rows.Add(rowBooking_Item)
                    Next
                End If
            End If

            If ListPrimaryCodeNo_StockIn.Count > 0 Then
                For ij As Double = TBLBooking_Item.Rows.Count - 1 To 0 Step -1
                    rowBooking_Item = TBLBooking_Item.Rows(ij)
                    If ListPrimaryCodeNo_StockIn.Contains(rowBooking_Item.Primary_Product) = False Then
                        TBLBooking_Item.Rows.Remove(rowBooking_Item)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblDummy = Nothing
        End Try
    End Sub

    Private Sub BuangItemExclude_PurcReq(ByVal TBLPurch_req_item As DseTransaksiShippingForecast.tbltrans_purch_req_itemDataTable, _
  ByVal TBLGroup As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionDataTable, _
    ByVal ListPrimaryCodeNo_StockIn As List(Of Double))
        Dim rowPurch_req_item As DseTransaksiShippingForecast.tbltrans_purch_req_itemRow
        Dim ProductLevel As ClsTransaksiMaterialPlanningGroup.ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand
        Dim Filter As String = ""
        Dim TipeExclusion As ClsTransaksiMaterialPlanning.EnumTipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude
        Dim TblDummy As New DseTransaksiShippingForecast.tbltrans_purch_req_itemDataTable
        Dim rowDummy As DseTransaksiShippingForecast.tbltrans_purch_req_itemRow
        Dim rowGroup As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionRow

        Try
            If TBLGroup.Rows.Count > 0 Then
                rowGroup = TBLGroup.Rows(0)
                TipeExclusion = rowGroup.tipe_exclusion

                If TipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude Then
                    For Each rowGroup In TBLGroup.Rows
                        ProductLevel = rowGroup.level
                        With rowGroup
                            If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                                Filter = "[Primary Brand] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                                Filter = "[Primary Category] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                                Filter = "[primary code no] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                                Filter = "[Primary Product Group] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                                Filter = "[Primary Sub Brand] = " + .primary_acuan.ToString
                            End If

                            TBLPurch_req_item.DefaultView.RowFilter = Filter
                            While TBLPurch_req_item.DefaultView.Count > 0
                                rowPurch_req_item = TBLPurch_req_item.DefaultView(0).Row
                                TBLPurch_req_item.Rows.Remove(rowPurch_req_item)
                            End While
                        End With
                    Next
                Else
                    TblDummy.Rows.Clear()

                    'pindahkan ke tabel dummy
                    For Each rowGroup In TBLGroup.Rows
                        ProductLevel = rowGroup.level
                        With rowGroup
                            If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                                Filter = "[Primary Brand] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                                Filter = "[Primary Category] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                                Filter = "[primary code no] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                                Filter = "[Primary Product Group] = " + .primary_acuan.ToString
                            ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                                Filter = "[Primary Sub Brand] = " + .primary_acuan.ToString
                            End If

                            TBLPurch_req_item.DefaultView.RowFilter = Filter
                            While TBLPurch_req_item.DefaultView.Count > 0
                                rowPurch_req_item = TBLPurch_req_item.DefaultView(0).Row

                                rowDummy = TblDummy.Newtbltrans_purch_req_itemRow
                                For hj As Integer = 0 To TblDummy.Columns.Count - 1
                                    rowDummy.Item(hj) = rowPurch_req_item.Item(hj)
                                Next
                                TblDummy.Rows.Add(rowDummy)

                                TBLPurch_req_item.Rows.Remove(rowPurch_req_item)
                            End While
                        End With
                    Next
                    'kosongkan tabel asal
                    TBLPurch_req_item.Rows.Clear()

                    'pindahkan kembali ke tabel asal
                    For Each rowDummy In TblDummy.Rows
                        rowPurch_req_item = TBLPurch_req_item.Newtbltrans_purch_req_itemRow
                        For hj As Integer = 0 To TblDummy.Columns.Count - 1
                            rowPurch_req_item.Item(hj) = rowDummy.Item(hj)
                        Next
                        TBLPurch_req_item.Rows.Add(rowPurch_req_item)
                    Next
                End If
            End If

            If ListPrimaryCodeNo_StockIn.Count > 0 Then
                For ij As Double = TBLPurch_req_item.Rows.Count - 1 To 0 Step -1
                    rowPurch_req_item = TBLPurch_req_item.Rows(ij)
                    If ListPrimaryCodeNo_StockIn.Contains(rowPurch_req_item.primary_code_no) = False Then
                        TBLPurch_req_item.Rows.Remove(rowPurch_req_item)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblDummy = Nothing
        End Try
    End Sub

    Private Sub PasangkanItemSOB(ByVal vBookedOnly As Boolean, ByVal ListPrimaryCodeNo As List(Of Double))
        Dim QtyPOReady As Decimal
        Dim QtyReadyStock As Decimal
        Dim IndexItemDelivery As Integer = 0
        Dim IndexItemPOB As Integer = 0
        Dim rowStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockRow
        Dim rowItemSO_SF As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim rowItemPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim Dummy As Decimal = 0

        Try
            For Each PrimaryCodeNo As Double In ListPrimaryCodeNo
                'filter
                Me.vTBLItemStock.DefaultView.RowFilter = "[primary code no] = " + PrimaryCodeNo.ToString
                Me.vTBLItemPOB.DefaultView.RowFilter = "[primary code no] = " + PrimaryCodeNo.ToString
                Me.vTBLItemPOB.DefaultView.Sort = "[available date] ASC, [promise to deliver] ASC, [po date rec] ASC, [primary po] ASC"
                If vBookedOnly = True Then
                    Me.vTBLItemDelivery.DefaultView.RowFilter = "[primary code no] = " + PrimaryCodeNo.ToString + " AND booked = 1"
                Else
                    Me.vTBLItemDelivery.DefaultView.RowFilter = "[primary code no] = " + PrimaryCodeNo.ToString
                End If
                Me.vTBLItemDelivery.DefaultView.Sort = "[promise to deliver] ASC, [booked] DESC, [urutan priority level] ASC, [so date rec] ASC, [primary so] ASC"

                'kembalikan parameter ke nilai awal
                IndexItemDelivery = 0
                QtyPOReady = 0

                'pasangkan dengan yang ready stock dahulu
                If Me.vTBLItemStock.DefaultView.Count > 0 Then
                    rowStock = Me.vTBLItemStock.DefaultView(0).Row
                    While rowStock.qty - rowStock.qty_used > 0 And vTBLItemDelivery.DefaultView.Count > IndexItemDelivery
                        rowItemSO_SF = Me.vTBLItemDelivery.DefaultView(IndexItemDelivery).Row
                        If rowStock.qty - rowStock.qty_used >= rowItemSO_SF.qty_sob Then
                            rowStock.qty_used += rowItemSO_SF.qty_sob
                            'karena ready stock, maka qty availabel = qty sob
                            rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                            rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                            'karena ready stock, maka tandai estimate delivery menjadi ASAP -> 1 jan 1950
                            rowItemSO_SF.estimate_delivery = #1/1/1950#
                            'naikan index item delivery, cek item delivery selanjutnya
                            IndexItemDelivery += 1
                        Else
                            QtyPOReady = rowStock.qty - rowStock.qty_used
                            QtyReadyStock = QtyPOReady
                            rowStock.qty_used = rowStock.qty
                            Exit While
                        End If
                    End While
                End If

                'pasangkan dengan item po
                If Me.vTBLItemPOB.DefaultView.Count > 0 Then
                    IndexItemPOB = 0
                    While IndexItemPOB < vTBLItemPOB.DefaultView.Count And IndexItemDelivery < vTBLItemDelivery.DefaultView.Count
                        rowItemPOB = Me.vTBLItemPOB.DefaultView(IndexItemPOB).Row
                        rowItemSO_SF = Me.vTBLItemDelivery.DefaultView(IndexItemDelivery).Row
                        If QtyPOReady + rowItemPOB.qty - rowItemPOB.qty_used >= rowItemSO_SF.qty_sob Then
                            'jika item yang tersedia lebih besar dari item delivery, maka 
                            'set delivery date + qty available delivery + pindah ke baris item delivery selanjutnya
                            rowItemSO_SF.estimate_delivery = rowItemPOB.available_date
                            rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                            IndexItemDelivery += 1

                            'kurangkan qty po ready
                            Dummy = rowItemSO_SF.qty_sob
                            If QtyPOReady >= Dummy Then
                                QtyPOReady -= Dummy
                                Dummy = 0
                            Else
                                Dummy -= QtyPOReady
                                QtyPOReady = 0
                            End If

                            'tambahkan qty used
                            rowItemPOB.qty_used += Dummy

                            'jika semua qty sudah habis dipakai, maka pindah ke baris item pob selanjutnya
                            If rowItemPOB.qty = rowItemPOB.qty_used Then
                                IndexItemPOB += 1
                            End If

                            'perhitungan qty ready stock 
                            If QtyReadyStock >= rowItemSO_SF.qty_sob Then
                                rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                                QtyReadyStock -= rowItemSO_SF.qty_sob
                            Else
                                rowItemSO_SF.qty_ready_stock = QtyReadyStock
                                QtyReadyStock = 0
                            End If
                        Else
                            'jika qty item delivery lebih besar dari qty po ready + qty po,ada 2 kemungkinan:
                            If IndexItemPOB + 1 < vTBLItemPOB.DefaultView.Count Then
                                '1. tambahkan qty po dari baris selanjut, karena dari pob, maka qty ready stock tidak di tambahkan
                                QtyPOReady += rowItemPOB.qty - rowItemPOB.qty_used
                                rowItemPOB.qty_used = rowItemPOB.qty
                                IndexItemPOB += 1
                            Else
                                '2. tidak ada baris selanjutnya, maka qty availble hanya sejumlah yang ada,
                                'estimate delivery tidak dikehui kapan, tapi available qty ada
                                rowItemSO_SF.qty_available = rowItemPOB.qty - rowItemPOB.qty_used + QtyPOReady
                                rowItemSO_SF.estimate_delivery = #1/1/2200#
                                'perhitungan ready stock
                                If QtyReadyStock >= rowItemSO_SF.qty_sob Then
                                    rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                                    QtyReadyStock -= rowItemSO_SF.qty_sob
                                Else
                                    rowItemSO_SF.qty_ready_stock = QtyReadyStock
                                    QtyReadyStock = 0
                                End If
                                'index item pob tetap dinaikan, untuk menghentikan looping perhitungan
                                IndexItemPOB += 1
                            End If
                        End If
                    End While
                End If

                'tambahkan qty po ready sisa ke dalam item sob terakhir
                If QtyPOReady > 0 Then
                    rowItemSO_SF = Me.vTBLItemDelivery.DefaultView(IndexItemDelivery).Row
                    rowItemSO_SF.qty_available = QtyPOReady
                    rowItemSO_SF.qty_ready_stock = QtyReadyStock
                    rowItemSO_SF.estimate_delivery = #1/1/2200#
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMax_ItemSOB() As Double
        Dim rowsob As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim Retval As Double = 0

        Try
            For Each rowsob In Me.vTBLItemDelivery.Rows
                If Retval < rowsob.primary Then
                    Retval = rowsob.primary
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryMax_ItemPOB() As Double
        Dim rowpob As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim Retval As Double = 0

        Try
            For Each rowpob In Me.vTBLItemPOB.Rows
                If Retval < rowpob.primary Then
                    Retval = rowpob.primary
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetFilterSF() As String
        Dim Filter As String = ""

        Try
            If Me.SplitContainer1.Panel1Collapsed = True Then
                If Me.vPrimaryStockInPurchaseReference = -1 Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_SF_DOAll") = False Then
                        If Filter = "" Then
                            Filter = "[so created by] = '" + Me.TxtFilter_UserCreateSO.Text + "'"
                        Else
                            Filter = Filter + " AND [so created by] = '" + Me.TxtFilter_UserCreateSO.Text + "'"
                        End If
                    Else
                        Filter = ""
                    End If
                End If
            Else
                Me.TxtFilter_CodeNo.Text = Me.TxtFilter_CodeNo.Text.Trim
                If Me.TxtFilter_CodeNo.Text <> "" Then
                    Filter = "[Product Code] like '" + Me.ObjCD.FormatSpecialChar_DataFilter(Me.TxtFilter_CodeNo.Text) + "*'"
                End If
                Me.TxtFilter_Desc.Text = Me.TxtFilter_Desc.Text.Trim
                If Me.TxtFilter_Desc.Text <> "" Then
                    If Filter = "" Then
                        Filter = "[Description] like '" + Me.ObjCD.FormatSpecialChar_DataFilter(Me.TxtFilter_Desc.Text) + "*'"
                    Else
                        Filter = Filter + " AND [Description] like '" + Me.ObjCD.FormatSpecialChar_DataFilter(Me.TxtFilter_Desc.Text) + "*'"
                    End If
                End If
                'If Me.TxtFilter_PrimaryCustomer.Text <> "-1" And Me.TxtFilter_PrimaryCustomer.Text <> "-1" Then
                '    If Filter = "" Then
                '        Filter = "[primary cust] = " + Me.TxtFilter_PrimaryCustomer.Text
                '    Else
                '        Filter = Filter + " AND [primary cust] = " + Me.TxtFilter_PrimaryCustomer.Text
                '    End If
                'End If
                If Me.vPrimaryStockInPurchaseReference = -1 Then
                    If Me.TxtFilter_UserCreateSO.Text <> "" Then
                        If Filter = "" Then
                            Filter = "[so created by] = '" + Me.TxtFilter_UserCreateSO.Text + "'"
                        Else
                            Filter = Filter + " AND [so created by] = '" + Me.TxtFilter_UserCreateSO.Text + "'"
                        End If
                    End If
                End If
                If Me.TxtFilter_ProductGroupCode.Text <> "" Then
                    If Me.cboFilter_ProductGroupLevel.SelectedIndex = 0 Then
                        'product group
                        If Filter = "" Then
                            Filter = "[primary product group] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        Else
                            Filter = Filter + " AND [primary product group] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        End If
                    ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 1 Then
                        'brand
                        If Filter = "" Then
                            Filter = "[Primary Brand] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        Else
                            Filter = Filter + " AND [Primary Brand] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        End If
                    ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 2 Then
                        'sub band
                        If Filter = "" Then
                            Filter = "[Primary Sub Brand] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        Else
                            Filter = Filter + " AND [Primary Sub Brand] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        End If
                    ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 3 Then
                        'category
                        If Filter = "" Then
                            Filter = "[Primary Category] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        Else
                            Filter = Filter + " AND [Primary Category] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        End If
                    End If
                End If

                'rnew
                'If Me.chkSODateFilter.Checked = True Then
                '    If Filter = "" Then
                '        Filter = "[so date] >= #" + Format(Me.DtpFilter_From.Value, "yyyy-MM-dd") + "# AND [so date] <= #" + Format(Me.DtpFilter_To.Value, "yyyy-MM-dd") + "#"
                '    Else
                '        Filter = Filter + " AND [so date] >= #" + Format(Me.DtpFilter_From.Value, "yyyy-MM-dd") + "# AND [so date] <= #" + Format(Me.DtpFilter_To.Value, "yyyy-MM-dd") + "#"
                '    End If
                'End If
            End If

            Return Filter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub LoadDGV_SF()
        Dim Filter As String = ""
        Dim DV As Data.DataRowView
        Dim rowDGV As DataGridViewRow
        Dim Masuk As Boolean = True
        Dim rowItemSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim Lanjut As Boolean = False

        Try
            Filter = Me.GetFilterSF

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'secara default disort berdasarkan: estimate delivery, customer code, product code, so date, so number
            Me.vTBLItemDelivery.DefaultView.Sort = "[promise to deliver] ASC, [customer name] ASC, [Product Code] ASC, [so date rec] ASC, [so number rec] ASC"
            Me.vTBLItemDelivery.DefaultView.RowFilter = Filter

            Me.dgvSF.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvSF.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvSF.Rows.Clear()

            Me.ColSF_SODate.DefaultCellStyle.Format = "dd MMM yyyy"

            For Each DV In Me.vTBLItemDelivery.DefaultView
                Lanjut = True
                rowItemSOB = DV.Row
                If Me.chkSF_LateOnly.Checked = True Then
                    If rowItemSOB.promise_to_deliver = #1/1/2200# Then
                        Lanjut = False
                    Else
                        If rowItemSOB.promise_to_deliver.Date >= rowItemSOB.estimate_delivery.Date Then
                            If rowItemSOB.qty_available < rowItemSOB.qty_sob Then
                                Lanjut = True
                            Else
                                Lanjut = False
                            End If
                        Else
                            Lanjut = True
                        End If
                    End If
                Else
                    Lanjut = True
                End If
                'goods/service
                If Lanjut = True Then
                    If Me.chkFilter_Goods.Checked = False And rowItemSOB.tipe_product = CInt(Core_Module.clsMasterProduct.EnumTipeProduct.Goods) Then
                        Lanjut = False
                    End If
                End If
                If Lanjut = True Then
                    If Me.chkFilter_Service.Checked = False And rowItemSOB.tipe_product = CInt(Core_Module.clsMasterProduct.EnumTipeProduct.Services) Then
                        Lanjut = False
                    End If
                End If
                If Lanjut = True Then
                    If Me.chkFilter_FollowUpOnly.Checked = True And rowItemSOB.create_do <> 1 Then
                        Lanjut = False
                    End If
                End If
                If Lanjut = True Then
                    If Me.cboAvailability.SelectedIndex = 1 Then
                        'hanya yg available : qty available > qty sob dan tanggal available > dtp available
                        If rowItemSOB.qty_available >= rowItemSOB.qty_sob Then
                            If Me.dtpAvailable.Checked = False Then
                                Lanjut = True
                            Else
                                If rowItemSOB.estimate_delivery.Date <= Me.dtpAvailable.Value.Date Then
                                    Lanjut = True
                                Else
                                    Lanjut = False
                                End If
                            End If
                        Else
                            Lanjut = False
                        End If
                    ElseIf Me.cboAvailability.SelectedIndex = 2 Then
                        If rowItemSOB.qty_available < rowItemSOB.qty_sob Then
                            Lanjut = True
                        Else
                            Lanjut = False
                        End If
                    End If

                    'If Me.chkSF_AvailableOnly.Checked = True Then
                    '    If rowItemSOB.qty_available > 0 Then
                    '        Lanjut = True
                    '    Else
                    '        Lanjut = False
                    '    End If
                    'Else
                    '    Lanjut = True
                    'End If
                End If

                'berlaku untuk item sob
                If Lanjut = True Then
                    If rowItemSOB.tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB) Then
                        'primary customer
                        If Me.TxtFilter_PrimaryCustomer.Text <> -1 Then
                            If rowItemSOB.primary_cust <> Me.TxtFilter_PrimaryCustomer.Text Then
                                Lanjut = False
                            End If
                        End If
                        'so date
                        If Lanjut = True Then
                            If Me.chkSODateFilter.Checked = True Then
                                If rowItemSOB.so_date.Date < Me.DtpFilter_From.Value.Date Or rowItemSOB.so_date.Date > Me.DtpFilter_To.Value.Date Then
                                    Lanjut = False
                                End If
                            End If
                        End If
                        'item sob
                        If Lanjut = True Then
                            If Me.chkFilter_SOBItem.Checked = False Then
                                Lanjut = False
                            End If
                        End If
                    ElseIf rowItemSOB.tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req) Then
                        'primary ref party
                        If Me.Txtfilter_PrimaryRefParty.Text <> -1 Then
                            If rowItemSOB.primary_cust <> Me.Txtfilter_PrimaryRefParty.Text Then
                                Lanjut = False
                            End If
                        End If
                        'pr. date
                        If Lanjut = True Then
                            If Me.chkFilter_PurhReqDate.Checked = True Then
                                If rowItemSOB.so_date.Date < Me.DtpFilter_PurcReq_From.Value.Date Or rowItemSOB.so_date.Date > Me.DtpFilter_PurchReq_To.Value.Date Then
                                    Lanjut = False
                                End If
                            End If
                        End If
                        'item pr
                        If Lanjut = True Then
                            If Me.chkFilter_PrOBItem.Checked = False Then
                                Lanjut = False
                            End If
                        End If
                    End If
                End If

                If Lanjut = True Then
                    Me.dgvSF.Rows.Add()
                    rowDGV = Me.dgvSF.Rows(Me.dgvSF.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColSF_Primary.Index).Value = rowItemSOB.primary
                        .Cells(Me.ColSF_CodeNo.Index).Value = rowItemSOB.Product_Code
                        If rowItemSOB.create_do = 1 Then
                            .Cells(Me.ColSF_CreateDO.Index).Value = Boolean.TrueString
                        Else
                            .Cells(Me.ColSF_CreateDO.Index).Value = Boolean.FalseString
                        End If
                        .Cells(Me.ColSF_Customer.Index).Value = rowItemSOB.customer_name
                        .Cells(Me.ColSF_CustomerPONumber.Index).Value = rowItemSOB.Cust_PO_No
                        .Cells(Me.ColSF_EstimateDelivery.Index).Value = (rowItemSOB.estimate_delivery.Date - #1/1/1950#).Days
                        .Cells(Me.ColSF_No.Index).Value = Me.dgvSF.Rows.Count
                        .Cells(Me.ColSF_PrimaryCodeNo.Index).Value = rowItemSOB.primary_code_no
                        .Cells(Me.ColSF_PrimarySO.Index).Value = rowItemSOB.primary_so
                        .Cells(Me.ColSF_PrimaryUnit.Index).Value = rowItemSOB.primary_unit
                        .Cells(Me.ColSF_PromiseToDeliver.Index).Value = (rowItemSOB.promise_to_deliver.Date - #1/1/1950#).Days
                        .Cells(Me.ColSF_QtyAvailable.Index).Value = ObjCD.DigitGrouping(rowItemSOB.qty_available, rowItemSOB._Decimal)
                        .Cells(Me.ColSF_QtySOB.Index).Value = ObjCD.DigitGrouping(rowItemSOB.qty_sob, rowItemSOB._Decimal)
                        .Cells(Me.ColSF_SODate.Index).Value = rowItemSOB.so_date
                        .Cells(Me.ColSF_Unit.Index).Value = rowItemSOB.Unit_Code
                        .Cells(Me.ColSF_Desc.Index).Value = rowItemSOB.Description
                        If rowItemSOB.estimate_delivery.Date = #1/1/1950# Then
                            .Cells(Me.ColSF_EstimateDelivery_Str.Index).Value = "Today"
                        ElseIf rowItemSOB.estimate_delivery.Date = #1/1/2200# Then
                            .Cells(Me.ColSF_EstimateDelivery_Str.Index).Value = "n/a"
                        Else
                            .Cells(Me.ColSF_EstimateDelivery_Str.Index).Value = Format(rowItemSOB.estimate_delivery.Date, "dd MMM yyyy")
                        End If
                        If rowItemSOB.promise_to_deliver.Date = #1/1/1950# Then
                            .Cells(Me.ColSF_PromiseToDeliver_Str.Index).Value = "Today"
                        ElseIf rowItemSOB.promise_to_deliver = #1/1/2200# Then
                            .Cells(Me.ColSF_PromiseToDeliver_Str.Index).Value = "n/a"
                        Else
                            .Cells(Me.ColSF_PromiseToDeliver_Str.Index).Value = Format(rowItemSOB.promise_to_deliver.Date, "dd MMM yyyy")
                        End If
                        .Cells(Me.ColSF_SONumber.Index).Value = rowItemSOB.so_number
                        .Cells(Me.ColSF_PrimaryCust.Index).Value = rowItemSOB.primary_cust

                        If rowItemSOB.qty_available > 0 Then
                            If rowItemSOB.qty_sob <= rowItemSOB.qty_available Then
                                If rowItemSOB.qty_ready_stock >= rowItemSOB.qty_available Then
                                    .Cells(Me.ColSF_DeliveryStatus.Index).Value = "Ready"
                                ElseIf rowItemSOB.qty_ready_stock > 0 Then
                                    .Cells(Me.ColSF_DeliveryStatus.Index).Value = "Partial"
                                Else
                                    .Cells(Me.ColSF_DeliveryStatus.Index).Value = "Waiting"
                                End If
                            Else
                                If rowItemSOB.qty_ready_stock >= rowItemSOB.qty_available Then
                                    .Cells(Me.ColSF_DeliveryStatus.Index).Value = "Partial"
                                ElseIf rowItemSOB.qty_ready_stock > 0 Then
                                    .Cells(Me.ColSF_DeliveryStatus.Index).Value = "Partial"
                                Else
                                    .Cells(Me.ColSF_DeliveryStatus.Index).Value = "Waiting"
                                End If
                            End If
                        Else
                            .Cells(Me.ColSF_DeliveryStatus.Index).Value = "Waiting"
                        End If
                        .Cells(Me.ColSF_QtyReadyStock.Index).Value = rowItemSOB.qty_ready_stock
                        .Cells(Me.ColSF_TypeInt.Index).Value = rowItemSOB.tipe
                        If rowItemSOB.tipe = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB Then
                            .Cells(Me.ColSF_Type.Index).Value = Me.TipeItemSO_Str
                        Else
                            .Cells(Me.ColSF_Type.Index).Value = Me.TipeItemPurchReq_Str
                        End If
                        .Cells(Me.ColSF_Remarks.Index).Value = rowItemSOB.remark
                        If rowItemSOB.tipe_product = Core_Module.clsMasterProduct.EnumTipeProduct.Services And Me.vFromStockIn = False Then
                            .Cells(Me.ColSF_QtyAvailable.Index).Value = "-"
                        End If
                    End With

                    If rowItemSOB.promise_to_deliver.Date < rowItemSOB.estimate_delivery.Date Or _
                    rowItemSOB.qty_available < rowItemSOB.qty_sob Then
                        rowDGV.DefaultCellStyle.ForeColor = Color.Red
                    End If
                End If
            Next

            Me.dgvSF.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            If Me.dgvSF.Rows.Count > 0 Then
                Me.llbSF_CekAll.Visible = True
            Else
                Me.llbSF_CekAll.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkSF_LateOnly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkSF_LateOnly.CheckedChanged
        Try
            Me.LoadDGV_SF()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSF_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSF.ColumnHeaderMouseClick
        If e.ColumnIndex = Me.ColSF_PromiseToDeliver_Str.Index Then
            Me.mSortColIndex = Me.ColSF_PromiseToDeliver_Str.Index
            If Me.mSortASC = True Then
                Me.mSortASC = False
                Me.dgvSF.Sort(Me.ColSF_PromiseToDeliver, System.ComponentModel.ListSortDirection.Descending)
            Else
                Me.mSortASC = True
                Me.dgvSF.Sort(Me.ColSF_PromiseToDeliver, System.ComponentModel.ListSortDirection.Ascending)
            End If
        ElseIf e.ColumnIndex = Me.ColSF_EstimateDelivery_Str.Index Then
            Me.mSortColIndex = Me.ColSF_EstimateDelivery_Str.Index
            If Me.mSortASC = True Then
                Me.mSortASC = False
                Me.dgvSF.Sort(Me.ColSF_EstimateDelivery, System.ComponentModel.ListSortDirection.Descending)
            Else
                Me.mSortASC = True
                Me.dgvSF.Sort(Me.ColSF_EstimateDelivery, System.ComponentModel.ListSortDirection.Ascending)
            End If
        Else
            Me.mSortASC = False
            Me.mSortColIndex = -1
        End If
    End Sub

    Private Sub llbSF_CekAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSF_CekAll.LinkClicked
        Dim RefVal As Boolean = False
        Dim ModifiInt As Integer = 0
        Dim rowDel As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If Me.llbSF_CekAll.Text = "Check All" Then
                RefVal = True
                Me.llbSF_CekAll.Text = "Uncheck All"
                ModifiInt = 1
            Else
                RefVal = False
                Me.llbSF_CekAll.Text = "Check All"
                ModifiInt = 0
            End If

            Me.dgvSF.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            If ModifiInt = 0 Then
                For Each rowDGV As DataGridViewRow In Me.dgvSF.Rows
                    rowDGV.Cells(Me.ColSF_CreateDO.Index).Value = Boolean.FalseString
                    Me.vTBLItemDelivery.DefaultView.RowFilter = "primary = " + rowDGV.Cells(Me.ColSF_Primary.Index).Value.ToString
                    rowDel = Me.vTBLItemDelivery.DefaultView(0).Row
                    rowDel.create_do = 0
                Next
            Else
                For Each rowDGV As DataGridViewRow In Me.dgvSF.Rows
                    'cek semua item yang siap kirim: memiliki ready stock + estimate deliverynya lebih kecil dari tanggal acuan
                    If rowDGV.Cells(Me.ColSF_DeliveryStatus.Index).Value.ToString.ToLower = "ready" Then
                        rowDGV.Cells(Me.ColSF_CreateDO.Index).Value = Boolean.TrueString
                        Me.vTBLItemDelivery.DefaultView.RowFilter = "primary = " + rowDGV.Cells(Me.ColSF_Primary.Index).Value.ToString
                        rowDel = Me.vTBLItemDelivery.DefaultView(0).Row
                        rowDel.create_do = 1
                    Else
                        If CDec(rowDGV.Cells(Me.ColSF_QtyReadyStock.Index).Value) > 0 And _
                        CDate(#1/1/1950# + TimeSpan.FromDays(rowDGV.Cells(Me.ColSF_EstimateDelivery.Index).Value)) <= Me.dtpSF.Value.Date Then
                            rowDGV.Cells(Me.ColSF_CreateDO.Index).Value = Boolean.TrueString
                            Me.vTBLItemDelivery.DefaultView.RowFilter = "primary = " + rowDGV.Cells(Me.ColSF_Primary.Index).Value.ToString
                            rowDel = Me.vTBLItemDelivery.DefaultView(0).Row
                            rowDel.create_do = 1
                        End If
                    End If
                Next
            End If
            Me.dgvSF.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSF_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSF.CellBeginEdit
        Dim ShowMsgbox As Boolean = False

        Try
            'yang bisa diedit hanya kolom create do
            If e.ColumnIndex = Me.ColSF_CreateDO.Index Or e.ColumnIndex = Me.ColSF_Remarks.Index Then
                If e.ColumnIndex = Me.ColSF_CreateDO.Index Then
                    If Me.BufferCekBoxSF <> Me.dgvSF.Item(Me.ColSF_CodeNo.Index, e.RowIndex).Value Then
                        Me.BufferCekBoxSF = Me.dgvSF.Item(Me.ColSF_CodeNo.Index, e.RowIndex).Value
                    Else
                        ShowMsgbox = True
                    End If

                    If Me.dgvSF.Item(Me.ColSF_QtyReadyStock.Index, e.RowIndex).Value <= 0 Then
                        If ShowMsgbox = True Then
                            MsgBox("Ready stock qty = 0, you cannot create do form this item", MsgBoxStyle.Information, "Information")
                        End If
                        e.Cancel = True
                    End If
                Else
                    Me.Buffer = Me.dgvSF.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSF_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSF.CellEndEdit
        Dim rowDel As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow

        Try
            If e.ColumnIndex = Me.ColSF_CreateDO.Index Then
                'update ke dalam item deliverynya
                Me.vTBLItemDelivery.DefaultView.RowFilter = "primary = " + Me.dgvSF.Item(Me.ColSF_Primary.Index, e.RowIndex).Value.ToString
                rowDel = Me.vTBLItemDelivery.DefaultView(0).Row
                If Me.dgvSF.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString Then
                    rowDel.create_do = 0
                Else
                    rowDel.create_do = 1
                End If
            ElseIf e.ColumnIndex = Me.ColSF_Remarks.Index Then
                If IsNothing(Me.dgvSF.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.dgvSF.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Else
                    If Me.dgvSF.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvSF.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvSF.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'update ke dalam item deliverynya
                    Me.vTBLItemDelivery.DefaultView.RowFilter = "primary = " + Me.dgvSF.Item(Me.ColSF_Primary.Index, e.RowIndex).Value.ToString
                    rowDel = Me.vTBLItemDelivery.DefaultView(0).Row
                    rowDel.remark = Me.dgvSF.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSF_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSF.CellMouseClick
        Try
            If e.ColumnIndex <> -1 Then
                If e.RowIndex <> -1 Then
                    If e.ColumnIndex = Me.ColSF_EstimateDelivery_Str.Index Then
                        Dim ij As New FrmTrans_MaterialPlanningDashboard_SF_DetailDelivery
                        If Me.TxtSF_Remarks.Enabled = False Then
                            ij.vReadOnly = True
                        Else
                            ij.vReadOnly = False
                        End If
                        ij.UserCode = Me.UserCode
                        ij.vTBLItemDelivery = Me.vTBLItemDelivery
                        ij.vTBLItemPOB = Me.vTBLItemPOB
                        ij.vTBLItemStock = Me.vTBLItemStock
                        ij.vPrimaryTarget = Me.dgvSF.Item(Me.ColSF_Primary.Index, e.RowIndex).Value
                        ij.vPrimaryCodeNoTarget = Me.dgvSF.Item(Me.ColSF_PrimaryCodeNo.Index, e.RowIndex).Value
                        ij.vFromStockIn = Me.vFromStockIn
                        ij.vPrimaryStockInPurchaseReference = Me.vPrimaryStockInPurchaseReference
                        ij.TBL_ProductQty1 = Me.TBL_ProductQty1

                        ij.ShowDialog()
                        ij.Dispose()
                        If ij.DataChanged = True Then
                            Me.LoadDGV_SF()
                            Me.FollowUp()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbColumnSettings_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbColumnSettings.LinkClicked
        Dim ij As New FrmTrans_MaterialPlanningDashboard_SF_SetVisible

        Try
            'tampilkan form untuk setting data
            ij.vUserName = Me.UserCode
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                'update kolom
                Me.UpdateKolomSF()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub UpdateKolomSF()
        Dim Dummy As String = ""

        Try
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowSONumber, _
            Me.UserCode)
            If Dummy = "" Then
                Me.ColSF_SONumber.Visible = False
            Else
                If Dummy = 1 Then
                    Me.ColSF_SONumber.Visible = True
                Else
                    Me.ColSF_SONumber.Visible = False
                End If
            End If

            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowProductDesc, _
            Me.UserCode)
            If Dummy = "" Then
                Me.ColSF_Desc.Visible = False
            Else
                If Dummy = 1 Then
                    Me.ColSF_Desc.Visible = True
                Else
                    Me.ColSF_Desc.Visible = False
                End If
            End If

            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowCustPO, _
            Me.UserCode)
            If Dummy = "" Then
                Me.ColSF_CustomerPONumber.Visible = False
            Else
                If Dummy = 1 Then
                    Me.ColSF_CustomerPONumber.Visible = True
                Else
                    Me.ColSF_CustomerPONumber.Visible = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboAvailability_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAvailability.SelectedIndexChanged
        Try
            If Me.cboAvailability.SelectedIndex = 1 Then
                Me.dtpAvailable.Enabled = True
            Else
                Me.dtpAvailable.Enabled = False
            End If
            Me.LoadDGV_SF()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbSF_CreateDO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbSF_CreateDO.Click
        Dim ij As New FrmTrans_MaterialPlanningDashboard_SF_AddDO
        Dim AdaCreatedDO As Boolean = False
        Dim JmlDO As Double = 0
        Dim JmlPR As Double = 0
        Dim hasil As MsgBoxResult

        Try
            For Each rowdgv As DataGridViewRow In Me.dgvSF.Rows
                If rowdgv.Cells(Me.ColSF_CreateDO.Index).Value = Boolean.TrueString And rowdgv.Cells(Me.ColSF_TypeInt.Index).Value = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB Then
                    AdaCreatedDO = True
                    Exit For
                End If
            Next

            JmlDO = Me.ObjDO.GetJml("%", EnumSearchCriteria.ByNoTransaksi)
            If AdaCreatedDO = True Then
                ij.vFrmRef = Me
                ij.vdgvRef = Me.dgvSF
                ij.vUserCode = Me.UserCode
                ij.BackColorAll = Me.BackColorAll

                hasil = MsgBox("Create DO by SO?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    ij.vDObySO = True
                End If

                ij.ShowDialog()
                'jika jml do bertambah, maka harus recalculate ulang
            End If

            AdaCreatedDO = False
            For Each rowdgv As DataGridViewRow In Me.dgvSF.Rows
                If rowdgv.Cells(Me.ColSF_CreateDO.Index).Value = Boolean.TrueString And _
                rowdgv.Cells(Me.ColSF_TypeInt.Index).Value = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req Then
                    AdaCreatedDO = True
                    Exit For
                End If
            Next
            JmlPR = Me.ObjPickUpOrder.GetJml("%", EnumSearchCriteria.ByNoTransaksi)
            If AdaCreatedDO = True Then
                Me.CreatePickUpOrder()
            End If

            If Me.vPrimaryStockInPurchaseReference = -1 Then
                If JmlDO <> Me.ObjDO.GetJml("%", EnumSearchCriteria.ByNoTransaksi) Or _
                          JmlPR <> Me.ObjPickUpOrder.GetJml("%", EnumSearchCriteria.ByNoTransaksi) Then
                    MsgBox("Recalculating shipping forecast.", MsgBoxStyle.Information, "Information")
                    Me.HitungShippingForecast()
                    Me.LoadDGV_SF()
                End If
            Else
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub CreatePickUpOrder()
        Dim i As New FrmTransaksiPR_PickUpOrder
        Dim rowDGV, rowDGVPickUpOrder As DataGridViewRow
        Dim AdaModify As Boolean = False
        Dim ListPrimaryRefParty As New List(Of Double)
        Dim Dummy As String = ""
        Dim Index1 As Integer = -1
        Dim IdentifierBaru As String = ""
        Dim ij As Integer = 0
        Dim vRowItemDelivery As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow

        Try
            For Each rowDGV In Me.dgvSF.Rows
                If rowDGV.Cells(Me.ColSF_CreateDO.Index).Value = Boolean.TrueString And _
                rowDGV.Cells(Me.ColSF_TypeInt.Index).Value = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req Then
                    AdaModify = True
                    Exit For
                End If
            Next

            If AdaModify = True Then
                '1. kelompokkan berdasarkan primary ref party
                For Each rowDGV In Me.dgvSF.Rows
                    If rowDGV.Cells(Me.ColSF_CreateDO.Index).Value = Boolean.TrueString And _
                    rowDGV.Cells(Me.ColSF_TypeInt.Index).Value = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req Then
                        If ListPrimaryRefParty.Contains(rowDGV.Cells(Me.ColSF_PrimaryCust.Index).Value) = False Then
                            ListPrimaryRefParty.Add(rowDGV.Cells(Me.ColSF_PrimaryCust.Index).Value)
                        End If
                    End If
                Next

                '2. untuk setiap primary RefParty, buat pick up order baru
                For Each vPrimaryRefParty As Double In ListPrimaryRefParty
                    ij += 1
                    i = New FrmTransaksiPR_PickUpOrder
                    i.BAckColorAll = Me.BackColorAll
                    i.UserCode = Me.UserCode
                    i.vReadyToSave = False
                    i.Show()
                    i.Location = New Point((i.Width * -1) - 1, (i.Height * -1) - 1)
                    i.NewPickUpOrder()
                    i.SetRefParty(vPrimaryRefParty)

                    i.lblProgress.Text = "Data " + ij.ToString + " of " + ListPrimaryRefParty.Count.ToString
                    i.prgAll.Maximum = ListPrimaryRefParty.Count
                    i.prgAll.Value = ij
                    i.lblProgress.Visible = True
                    i.btnStopProcess.Visible = True
                    i.prgAll.Visible = True
                    i.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    i.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                    i.TxtPrimaryStockInPurchase.Text = Me.vPrimaryStockInPurchaseReference
                    i.TxtStockInPurchaseNo.text = Me.ObjStockInPurchase.GetNoTransaksi(Me.vPrimaryStockInPurchaseReference)

                    'isikan data item
                    For Each rowDGV In Me.dgvSF.Rows
                        If rowDGV.Cells(Me.ColSF_CreateDO.Index).Value = Boolean.TrueString And _
                        rowDGV.Cells(Me.ColSF_TypeInt.Index).Value = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req Then
                            If rowDGV.Cells(Me.ColSF_PrimaryCust.Index).Value = vPrimaryRefParty Then
                                Index1 = i.GetIndex(rowDGV.Cells(Me.ColSF_PrimarySO.Index).Value, rowDGV.Cells(Me.ColSF_PrimaryCodeNo.Index).Value)
                                vTBLItemDelivery.DefaultView.RowFilter = "primary = " + rowDGV.Cells(Me.ColSF_Primary.Index).Value.ToString
                                If vTBLItemDelivery.DefaultView.Count = 0 Then
                                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                                    Exit Sub
                                End If
                                vRowItemDelivery = vTBLItemDelivery.DefaultView(0).Row

                                'cari data yang sama persis
                                If Index1 <> -1 Then
                                    'tambahkan nilai qty
                                    Dim row As DataGridViewRow = i.DgvItem.Rows(Index1)
                                    row.Cells(i.ColItem_Qty.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(i.ColItem_Qty.Index).Value) + _
                                    CDec(rowDGV.Cells(Me.ColSF_QtyAvailable.Index).Value), vRowItemDelivery._Decimal)
                                Else
                                    'tambahkan baris baru
                                    'cegah jangan sampe primary IPO kosong sama sekali
                                    i.DgvItem.Rows.Add()
                                    rowDGVPickUpOrder = i.DgvItem.Rows(i.DgvItem.Rows.Count - 1)
                                    With rowDGVPickUpOrder
                                        .Cells(i.ColItem_PrimaryPR.Index).Value = vRowItemDelivery.primary_so
                                        .Cells(i.ColItem_PrimaryCodeNo.Index).Value = vRowItemDelivery.primary_code_no
                                        .Cells(i.ColItem_PRNo.Index).Value = rowDGV.Cells(Me.ColSF_SONumber.Index).Value
                                        .Cells(i.ColItem_CodeNo.Index).Value = rowDGV.Cells(Me.ColSF_CodeNo.Index).Value
                                        .Cells(i.ColItem_Desc.Index).Value = rowDGV.Cells(Me.ColSF_Desc.Index).Value
                                        .Cells(i.ColItem_Qty.Index).Value = ObjCD.DigitGrouping(vRowItemDelivery.qty_available, vRowItemDelivery._Decimal)
                                        .Cells(i.ColItem_Unit.Index).Value = rowDGV.Cells(Me.ColSF_Unit.Index).Value
                                        .Cells(i.ColItem_Remarks.Index).Value = rowDGV.Cells(Me.ColSF_Remarks.Index).Value
                                        .Cells(i.ColItem_No.Index).Value = i.DgvItem.RowCount
                                        .Cells(i.ColItem_PrimaryUnit.Index).Value = Me.ObjProd.GetPrimaryUnit(vRowItemDelivery.primary_code_no)
                                    End With
                                End If
                            End If
                        End If
                    Next

                    'fokuskan ke baris terakhir
                    While i.DgvItem.Rows(i.DgvItem.Rows.Count - 1).Displayed = False _
                    Or i.DgvItem.DisplayedRowCount(True) <> i.DgvItem.DisplayedRowCount(False)
                        i.DgvItem.FirstDisplayedScrollingRowIndex = i.DgvItem.FirstDisplayedScrollingRowIndex + 1
                    End While
                    i.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

                    'i.Location = i.BeginingLocation
                    i.Hide()
                    i.vReadyToSave = True

                    If i.ShowDialog() = Windows.Forms.DialogResult.Abort Then
                        i.Dispose()
                        Exit For
                    End If
                    i.Dispose()
                Next
            End If

            MsgBox("Pick up order created.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryRefParty = Nothing
        End Try
    End Sub

    Private Sub dgvSF_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSF.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvSF, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpAvailable_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpAvailable.Enter
        Try
            Me.Buffer = Format(Me.dtpAvailable.Value.Date, "dd MMM yyyy") + " - " + Me.dtpAvailable.Checked.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpAvailable_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpAvailable.Leave
        Try
            If Me.Buffer <> Format(Me.dtpAvailable.Value.Date, "dd MMM yyyy") + " - " + Me.dtpAvailable.Checked.ToString Then
                Me.BackgroundWorker5.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker5.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted
        Try
            Me.LoadDGV_SF()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFilter_CustCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFilter_CustCode.KeyDown
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtFilter_CustCode.Text) = True Then
                    Me.TxtFilter_PrimaryCustomer.Text = Me.ObjCustomer.GetPrimary(Me.TxtFilter_CustCode.Text)
                    Me.TxtFilter_CustName.Text = Me.ObjCustomer.GetCustName(Me.TxtFilter_PrimaryCustomer.Text)
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.TxtFilter_CustCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtFilter_CustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtFilter_PrimaryCustomer.Text = i.PrimaryCustomer
                        Me.TxtFilter_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtFilter_PrimaryCustomer.Text)
                        Me.TxtFilter_CustName.Text = Me.ObjCustomer.GetCustName(Me.TxtFilter_PrimaryCustomer.Text)
                    Else
                        Me.TxtFilter_CustCode.Clear()
                        Me.TxtFilter_CustName.Clear()
                        Me.TxtFilter_PrimaryCustomer.Text = "-1"
                        Me.TxtFilter_CustCode.Focus()
                    End If
                End If

                Me.LoadUlang()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtFilter_CustCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFilter_CustCode.Leave
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtFilter_CustCode.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtFilter_CustCode.Text) Then
                    Dim PrimaryBaru As Decimal = Me.ObjCustomer.GetPrimary(Me.TxtFilter_CustCode.Text)
                    If PrimaryBaru <> Me.TxtFilter_PrimaryCustomer.Text Then
                        Me.TxtFilter_PrimaryCustomer.Text = PrimaryBaru
                        Me.TxtFilter_CustName.Text = Me.ObjCustomer.GetCustName(Me.TxtFilter_PrimaryCustomer.Text)
                    End If
                Else
                    If Me.TxtFilter_CustCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtFilter_CustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtFilter_PrimaryCustomer.Text = i.PrimaryCustomer
                        Me.TxtFilter_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtFilter_PrimaryCustomer.Text)
                        Me.TxtFilter_CustName.Text = Me.ObjCustomer.GetCustName(Me.TxtFilter_PrimaryCustomer.Text)
                    Else
                        Me.TxtFilter_CustCode.Clear()
                        Me.TxtFilter_CustName.Clear()
                        Me.TxtFilter_PrimaryCustomer.Text = "-1"
                        Me.TxtFilter_CustCode.Focus()
                    End If
                End If
            Else
                Me.TxtFilter_CustCode.Clear()
                Me.TxtFilter_CustName.Clear()
                Me.TxtFilter_PrimaryCustomer.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFilter_FindCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter_FindCust.Click
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtFilter_CustCode.Text <> "" Then
                i.txtCari.Text = Me.TxtFilter_CustCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.TxtFilter_PrimaryCustomer.Text = i.PrimaryCustomer
                Me.TxtFilter_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtFilter_PrimaryCustomer.Text)
                Me.TxtFilter_CustName.Text = Me.ObjCustomer.GetCustName(Me.TxtFilter_PrimaryCustomer.Text)
            Else
                Me.TxtFilter_CustCode.Clear()
                Me.TxtFilter_CustName.Clear()
                Me.TxtFilter_PrimaryCustomer.Text = "-1"
                Me.TxtFilter_CustCode.Focus()
            End If

            Me.LoadUlang()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtFilter_UserCreateSO_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFilter_UserCreateSO.KeyDown
        Dim i As New FrmCariMasterUser
        i.backColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjUser.ApakahSudahAda(Me.TxtFilter_UserCreateSO.Text) = True Then
                    Me.TxtFilter_NamaUserCreateSO.Text = Me.ObjUser.GetNama(Me.TxtFilter_UserCreateSO.Text)
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.TxtFilter_UserCreateSO.Text <> "" Then
                        i.TxtCari.Text = Me.TxtFilter_UserCreateSO.Text
                    End If

                    i.ShowDialog()
                    If i.Kodeuser <> "" Then
                        Me.TxtFilter_UserCreateSO.Text = i.Kodeuser
                        Me.TxtFilter_NamaUserCreateSO.Text = Me.ObjUser.GetNama(Me.TxtFilter_UserCreateSO.Text)
                    Else
                        Me.TxtFilter_UserCreateSO.Clear()
                        Me.TxtFilter_NamaUserCreateSO.Clear()
                    End If
                End If

                Me.LoadUlang()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtFilter_UserCreateSO_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFilter_UserCreateSO.Leave
        Dim i As New FrmCariMasterUser
        i.backColorAll = Me.BackColorAll

        Try
            If Me.TxtFilter_UserCreateSO.Text <> "" Then
                If Me.ObjUser.ApakahSudahAda(Me.TxtFilter_UserCreateSO.Text) Then
                    Me.TxtFilter_NamaUserCreateSO.Text = Me.ObjUser.GetNama(Me.TxtFilter_UserCreateSO.Text)
                Else
                    If Me.TxtFilter_UserCreateSO.Text <> "" Then
                        i.TxtCari.Text = Me.TxtFilter_UserCreateSO.Text
                    End If

                    i.ShowDialog()
                    If i.Kodeuser <> "" Then
                        Me.TxtFilter_UserCreateSO.Text = i.Kodeuser
                        Me.TxtFilter_NamaUserCreateSO.Text = Me.ObjUser.GetNama(Me.TxtFilter_UserCreateSO.Text)
                    Else
                        Me.TxtFilter_UserCreateSO.Clear()
                        Me.TxtFilter_NamaUserCreateSO.Clear()
                    End If
                End If
            Else
                Me.TxtFilter_UserCreateSO.Clear()
                Me.TxtFilter_NamaUserCreateSO.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFilter_FindUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter_FindUser.Click
        Dim i As New FrmCariMasterUser
        i.backColorAll = Me.BackColorAll

        Try
            If Me.TxtFilter_UserCreateSO.Text <> "" Then
                i.TxtCari.Text = Me.TxtFilter_UserCreateSO.Text
            End If

            i.ShowDialog()
            If i.Kodeuser <> "" Then
                Me.TxtFilter_UserCreateSO.Text = i.Kodeuser
                Me.TxtFilter_NamaUserCreateSO.Text = Me.ObjUser.GetNama(Me.TxtFilter_UserCreateSO.Text)
            Else
                Me.TxtFilter_UserCreateSO.Clear()
                Me.TxtFilter_NamaUserCreateSO.Clear()
            End If

            Me.LoadUlang()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub GotFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFilter_UserCreateSO.GotFocus, _
    TxtFilter_CodeNo.GotFocus, TxtFilter_CustCode.GotFocus, TxtFilter_Desc.GotFocus, TxtFilter_UserCreateSO.GotFocus, _
    TxtFilter_PreferedSupplier.GotFocus, TxtFilter_ProductGroupCode.GotFocus, Txtfilter_RefPartyCode.GotFocus, TxtFilter_PurcReqNo.GotFocus
        Dim txt As TextBox = TryCast(sender, TextBox)
        txt.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtFilter_CodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFilter_CodeNo.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.LoadUlang()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFilter_Desc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFilter_Desc.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.LoadUlang()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub lostFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFilter_UserCreateSO.LostFocus, _
    TxtFilter_CodeNo.LostFocus, TxtFilter_CustCode.LostFocus, TxtFilter_Desc.LostFocus, TxtFilter_UserCreateSO.LostFocus, _
    TxtFilter_PreferedSupplier.LostFocus, TxtFilter_ProductGroupCode.LostFocus, Txtfilter_RefPartyCode.LostFocus, TxtFilter_PurcReqNo.LostFocus
        Dim txt As TextBox = TryCast(sender, TextBox)
        txt.BackColor = Color.White
    End Sub

    Private Sub btnSF_Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSF_Print.Click
        Dim i As New FrmTrans_MaterialPlanningDashboard_SF_Print

        Try
            If Me.TxtSF_DocNo.Text = "" Then
                Me.btnSF_Find.PerformClick()
            End If

            If Me.TxtSF_DocNo.Text <> "" Then
                i.vFilter = Me.GetFilterSF
                i.vLateOnly = Me.chkSF_LateOnly.Checked
                i.vPrimaryTransaksi = Me.TxtSF_Primary.Text
                i.vAvDateChecked = Me.dtpAvailable.Checked
                i.vAvDate = Me.dtpAvailable.Value
                If Me.cboAvailability.SelectedIndex = 0 Then
                    i.vEnumPrint = ClsTransaksiShippingForecast.EnumPrintAvailability.All
                ElseIf Me.cboAvailability.SelectedIndex = 1 Then
                    i.vEnumPrint = ClsTransaksiShippingForecast.EnumPrintAvailability.AvailableOnly
                Else
                    i.vEnumPrint = ClsTransaksiShippingForecast.EnumPrintAvailability.NotAvailableOnly
                End If
                i.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkSODateFilter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSODateFilter.CheckedChanged
        Try
            If Me.chkSODateFilter.Checked = True Then
                Me.DtpFilter_From.Enabled = True
                Me.DtpFilter_To.Enabled = True
                Me.Label7.Enabled = True
                Me.Label21.Enabled = True
            Else
                Me.DtpFilter_From.Enabled = False
                Me.DtpFilter_To.Enabled = False
                Me.Label7.Enabled = False
                Me.Label21.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpFilter_To_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DtpFilter_To.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Me.DtpFilter_From.Focus()
                Me.BackgroundWorker6.RunWorkerAsync()
                Me.DtpFilter_To.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpFilter_From_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DtpFilter_From.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Me.llbMP_Filter.Focus()
                Me.BackgroundWorker6.RunWorkerAsync()
                Me.DtpFilter_From.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkFilter_PurhReqFilter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFilter_PurhReqDate.CheckedChanged
        Try
            If Me.chkFilter_PurhReqDate.Checked = True Then
                Me.DtpFilter_PurcReq_From.Enabled = True
                Me.DtpFilter_PurchReq_To.Enabled = True
                Me.Label34.Enabled = True
                Me.Label35.Enabled = True
            Else
                Me.DtpFilter_PurcReq_From.Enabled = False
                Me.DtpFilter_PurchReq_To.Enabled = False
                Me.Label34.Enabled = False
                Me.Label35.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpFilter_PurchReq_To_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DtpFilter_PurchReq_To.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Me.DtpFilter_PurcReq_From.Focus()
                Me.BackgroundWorker6.RunWorkerAsync()
                Me.DtpFilter_PurchReq_To.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpFilter_PurcReq_From_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DtpFilter_PurcReq_From.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Me.llbMP_Filter.Focus()
                Me.BackgroundWorker6.RunWorkerAsync()
                Me.DtpFilter_PurcReq_From.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker6_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker6.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker6_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker6.RunWorkerCompleted
        Try
            Me.LoadUlang()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkSODateFilter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles chkSODateFilter.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Me.BackgroundWorker6.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txtfilter_RefPartyCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Txtfilter_RefPartyCode.KeyDown
        Dim i As New FrmMasterRefParty_Cari
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjRefParty.ApakahSudahAda(Me.Txtfilter_RefPartyCode.Text, ClsMasterRefParty.EnumAktif.Yes) = True Then
                    Me.Txtfilter_PrimaryRefParty.Text = Me.ObjRefParty.GetPrimary(Me.Txtfilter_RefPartyCode.Text)
                    Me.Txtfilter_RefPartyName.Text = Me.ObjRefParty.GetDesc(Me.Txtfilter_PrimaryRefParty.Text)
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.Txtfilter_RefPartyCode.Text <> "" Then
                        i.txtCari.Text = Me.Txtfilter_RefPartyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryRefParty > 0 Then
                        Me.Txtfilter_PrimaryRefParty.Text = i.PrimaryRefParty
                        Me.Txtfilter_RefPartyCode.Text = Me.ObjRefParty.GetCode(Me.Txtfilter_PrimaryRefParty.Text)
                        Me.Txtfilter_RefPartyName.Text = Me.ObjRefParty.GetDesc(Me.Txtfilter_PrimaryRefParty.Text)
                    Else
                        Me.Txtfilter_RefPartyCode.Clear()
                        Me.Txtfilter_RefPartyName.Clear()
                        Me.Txtfilter_PrimaryRefParty.Text = "-1"
                        Me.Txtfilter_RefPartyCode.Focus()
                    End If
                End If

                Me.LoadUlang()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub Txtfilter_RefPartyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txtfilter_RefPartyCode.Leave
        Dim i As New FrmMasterRefParty_Cari
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.Txtfilter_RefPartyCode.Text <> "" Then
                If Me.ObjRefParty.ApakahSudahAda(Me.Txtfilter_RefPartyCode.Text, ClsMasterRefParty.EnumAktif.Yes) = True Then
                    Dim PrimaryBaru As Decimal = Me.ObjRefParty.GetPrimary(Me.Txtfilter_RefPartyCode.Text)
                    If PrimaryBaru <> Me.Txtfilter_PrimaryRefParty.Text Then
                        Me.Txtfilter_PrimaryRefParty.Text = PrimaryBaru
                        Me.Txtfilter_RefPartyName.Text = Me.ObjRefParty.GetDesc(Me.Txtfilter_PrimaryRefParty.Text)
                    End If
                Else
                    If Me.Txtfilter_RefPartyCode.Text <> "" Then
                        i.txtCari.Text = Me.Txtfilter_RefPartyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryRefParty > 0 Then
                        Me.Txtfilter_PrimaryRefParty.Text = i.PrimaryRefParty
                        Me.Txtfilter_RefPartyCode.Text = Me.ObjRefParty.GetCode(Me.Txtfilter_PrimaryRefParty.Text)
                        Me.Txtfilter_RefPartyName.Text = Me.ObjRefParty.GetDesc(Me.Txtfilter_PrimaryRefParty.Text)
                    Else
                        Me.Txtfilter_RefPartyCode.Clear()
                        Me.Txtfilter_RefPartyName.Clear()
                        Me.Txtfilter_PrimaryRefParty.Text = "-1"
                        Me.Txtfilter_RefPartyCode.Focus()
                    End If
                End If
            Else
                Me.Txtfilter_RefPartyCode.Clear()
                Me.Txtfilter_RefPartyName.Clear()
                Me.Txtfilter_PrimaryRefParty.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFilter_FindRefParty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter_FindRefParty.Click
        Dim i As New FrmMasterRefParty_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.Txtfilter_RefPartyCode.Text <> "" Then
                i.txtCari.Text = Me.Txtfilter_RefPartyCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryRefParty > 0 Then
                Me.Txtfilter_PrimaryRefParty.Text = i.PrimaryRefParty
                Me.Txtfilter_RefPartyCode.Text = Me.ObjRefParty.GetCode(Me.Txtfilter_PrimaryRefParty.Text)
                Me.Txtfilter_RefPartyName.Text = Me.ObjRefParty.GetDesc(Me.Txtfilter_PrimaryRefParty.Text)
            Else
                Me.Txtfilter_RefPartyCode.Clear()
                Me.Txtfilter_RefPartyName.Clear()
                Me.Txtfilter_PrimaryRefParty.Text = "-1"
                Me.Txtfilter_RefPartyCode.Focus()
            End If

            Me.LoadUlang()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub
End Class
