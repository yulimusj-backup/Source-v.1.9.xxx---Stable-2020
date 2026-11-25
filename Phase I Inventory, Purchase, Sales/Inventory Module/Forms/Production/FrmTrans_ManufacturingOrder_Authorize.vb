Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_Authorize
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim objSetting As New Core_Module.ClsTransaksiSettingUser
    Dim objCD As New Core_Module.ClsCekData
    Dim objUser As New Core_Module.ClsUser
    Dim objAuthMO As New ClsTrans_ManufacturingOrder
    Dim TBLItemBOM As New DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
    Dim TBLProcessOrder As New DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable
    Dim TBLItemWO As New DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
    Dim TBLItemRM As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
    Dim TBLItemRMAdd As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addDataTable
    Dim TblItemWIP As New DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable
    Dim TBLJO_Temp As New DseTrans_ManufacturingOrder.JO_TempDataTable
    Dim EraseRowIndex As Integer = 0
    Dim buffer As String = ""
    Public val As New List(Of Double)
    Dim frm As New FrmTrans_ManufacturingOrder

    Public TBLItemDelivery As New DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryDataTable
    Dim RowItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryRow

    Private Sub FrmTrans_ManufacturingOrder_Authorize_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            Me.TboAuthorize_Click()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_Authorize_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim frm As New FrmTrans_ManufacturingOrder
        Try
            ''ok-rosadin-r5b
            Me.DGVRM.AllowUserToDeleteRows = False
            Me.TxtNoMO.CharacterCasing = CharacterCasing.Upper
            Me.txtCodeNo.CharacterCasing = CharacterCasing.Upper

            Me.ColRM_GPActual.Visible = False
            Me.ColRM_GPTarget.Visible = False
            Me.ColRM_PickUpLocation.Visible = False
            Me.ColRM_StandByLocation.Visible = False
            Me.ReadOnlyDetail()
            Me.objAuthMO.IsiDGVAuthorize(TxtPrimaryMO.Text, Me.UserCode, Me.DgvOtorisasi, Me.DgvCancelOtorisasi, Me.DGVAllotorisasi, Me)
            If DgvOtorisasi.RowCount > 0 Then
                If Me.objAuthMO.CariDetail(Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, 0).Value, Me.DgvItemJO, TBLItemBOM, TBLProcessOrder, _
                TBLItemWO, TBLItemRM, frm, objCD, Me.TBLItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, TBLJO_Temp) = True Then
                    Me.SalinKeField()
                End If
            End If
            Me.DGVRM.ReadOnly = True
            Me.setColumnVisibility()
            Me.cboStatus.Left = Me.TxtNoMO.Left
            Me.cboStatus.Width = Me.TxtNoMO.Width
            Me.TabPage4.Text = "All MO"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub setColumnVisibility()
        Dim ShowCodeNo, ShowDesc, ShowProcessCode, ShowProcessDesc As String

        Try
            ShowCodeNo = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowCodeNo, Me.UserCode)
            ShowDesc = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowDescription, Me.UserCode)
            ShowProcessCode = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowProcessCode, Me.UserCode)
            ShowProcessDesc = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowProcessDesc, Me.UserCode)

            If ShowCodeNo = "" And ShowDesc = "" And ShowProcessCode = "" And ShowProcessDesc = "" Then
                ShowCodeNo = "1"
                ShowDesc = "1"
                ShowProcessCode = "1"
                ShowProcessDesc = "1"
                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowCodeNo, ShowCodeNo, Me.UserCode)
                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                         Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowDescription, ShowDesc, Me.UserCode)
                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowProcessCode, ShowProcessCode, Me.UserCode)
                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowProcessDesc, ShowProcessDesc, Me.UserCode)
            End If
            If ShowCodeNo = "1" Then
                Me.ColWO2_CodeNo.Visible = True
            Else
                Me.ColWO2_CodeNo.Visible = False
            End If
            If ShowDesc = "1" Then
                Me.ColWO2_ProductDesc.Visible = True
            Else
                Me.ColWO2_ProductDesc.Visible = False
            End If
            If ShowProcessCode = "1" Then
                Me.ColWO2_ProcessCode.Visible = True
            Else
                Me.ColWO2_ProcessCode.Visible = False
            End If
            If ShowProcessDesc = "1" Then
                Me.ColWO2_ProcessDesc.Visible = True
            Else
                Me.ColWO2_ProcessDesc.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ReadOnlyDetail()
        Try

            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer6.Panel1.Controls, True, True, True, True, True, False, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer6.Panel2.Controls, True, True, True, True, True, False, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, False, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.GroupBox2.Controls, True, True, True, True, True, False, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, False, True, True)

            Me.DgvOtorisasi.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvCancelOtorisasi.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVAllotorisasi.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvUser.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVItemJO.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.TGV_WO.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVRM.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll


            Me.CancelButton = Me.btnExit
            Me.cboStatus.Enabled = False
            Me.cboStatus.SelectedIndex = 2

            Me.ColRM_StandByLocation.Visible = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub KosongkanField()
        Try
            Me.txtDetailNoMO.Clear()
            Me.dtpTanggalMo.Value = Today
            Me.txtRemarks.Text = ""

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.TGV_WO.Nodes.Clear()
            Me.TGV_WO.Rows.Clear()
            Me.TBLItemBOM.Rows.Clear()
            Me.TBLItemWO.Rows.Clear()
            Me.TBLItemRM.Rows.Clear()
            Me.TBLProcessOrder.Rows.Clear()
            Me.TBLJO_Temp.Rows.Clear()
            Me.DGVItemJO.Rows.Clear()
            Me.DGVRM.Rows.Clear()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Try
            With Me.objAuthMO

                Me.txtPrimary.Text = .Primary
                Me.dtpTanggalMo.Value = .TanggalTransaksi
                Me.txtPrimaryAuth.Text = .PrimaryUser_Auth
                If .IsReady_ForScheduling = ClsTrans_ManufacturingOrder.EnumPilihSchedule.Schedule_Aktif Then
                    Me.ChkStatScheduling.Checked = True
                ElseIf .IsReady_ForScheduling = ClsTrans_ManufacturingOrder.EnumPilihSchedule.Schedule_tdk_Aktif Then
                    Me.ChkStatScheduling.Checked = False
                End If
                Me.txtDetailNoMO.Text = .NoTransaksi
                Me.txtRemarks.Text = .Remarks

                If Me.DGVItemJO.Rows.Count > 0 Then
                    Dim r As DataGridViewCell = Me.DGVItemJO.SelectedCells(0)
                    r.Selected = True
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TboAuthorize_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TboAuthorize.Click
        Try
            Me.TboAuthorize_Click()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub ChkMODate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkMODate.CheckedChanged
        If Me.ChkMODate.Checked = True Then
            Me.DtpFrom.Enabled = True
            Me.DtpTo.Enabled = True
        Else
            Me.DtpFrom.Enabled = False
            Me.DtpTo.Enabled = False
        End If
    End Sub

    Private Sub DgvUser_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvUser.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCreatedBy_Name.Index Then
                e.Cancel = True
            Else
                If e.ColumnIndex <> Me.ColCreatedBy.Index Then
                    If Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = "-1" Then
                        e.Cancel = True
                    End If
                Else
                    buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvUser.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvUser.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryUser As Double = -1
        Dim RowIdx As Integer = e.RowIndex

        Try
            With row

                If IsNothing(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColCreatedBy.Index Then
                        Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value.ToString.Replace("'", Me.buffer).Replace("*", Me.buffer).Replace("#", Me.buffer).Replace("@", Me.buffer)
                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.objUser.ApakahSudahAda(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryUser = Me.objUser.GetPrimary(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemUser(DgvUser.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex)
                            If index < 0 Then
                                Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = PrimaryUser
                                Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.objUser.GetUserCode(PrimaryUser)
                                Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(PrimaryUser)

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
                            PrimaryUser = Me.objUser.GetPrimary(ij.Kodeuser)

                            If PrimaryUser <> -1 Then
                                Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = PrimaryUser
                                Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.objUser.GetUserCode(PrimaryUser)
                                Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(PrimaryUser)

                                index = Me.GetIndexItemUser(DgvUser.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex)

                                If index > -1 Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    If String.IsNullOrEmpty(Me.buffer) = False Then
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.objUser.GetPrimary(Me.buffer)
                                        Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value)
                                        'Me.dgvUser.Item(Me.ColIdetifierHeader.Index, e.RowIndex).Value = Me.dgvManager.Item(Me.ColIdentifier.Index, Me.dgvManager.CurrentRow.Index).Value
                                        Exit Sub
                                    Else
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwUser.RunWorkerAsync()
                                    End If
                                Else
                                    Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.objUser.GetUserCode(PrimaryUser)
                                    Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(PrimaryUser)

                                End If
                            ElseIf PrimaryUser = -1 Then
                                If Me.DgvUser.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwUser.RunWorkerAsync()
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.objCD.SetNumber(Me.DgvUser)
                                    Exit Sub
                                End If
                                If Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwUser.RunWorkerAsync()
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.objCD.SetNumber(Me.DgvUser)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.objCD.SetNumber(Me.DgvUser)
                    End If
                Else
                    Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BgwUser.RunWorkerAsync()
                    Else
                        Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.buffer
                    End If
                    Me.objCD.SetNumber(Me.DgvUser)
                    Exit Sub
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemUser(ByVal vUsercode As String, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Try
            'pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.dgvUser.Rows.Count - 1
                RowDGV = Me.dgvUser.Rows(ij)
                If RowDGV.Cells(Me.ColCreatedBy.Index).Value <> "" Then
                    If vRowIdxFocus <> ij Then
                        If RowDGV.Cells(Me.ColCreatedBy.Index).Value = vUsercode Then
                            Return Me.DgvUser.Rows.IndexOf(RowDGV)
                        End If
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BgwUser_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwUser.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUser_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwUser.RunWorkerCompleted
        Try
            If DgvUser.RowCount > 1 Then
                Me.DgvUser.Rows.RemoveAt(EraseRowIndex)
                Me.objCD.SetNumber(Me.DgvUser)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''ok-rosadin-r5b
    Private Sub TxtNoMO_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoMO.GotFocus
        Try
            If Me.TxtNoMO.Enabled Then
                If Not Me.TxtNoMO.ReadOnly Then
                    Me.TxtNoMO.BackColor = Me.BackColorAll
                End If
            Else
                Me.TxtNoMO.BackColor = Me.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNoMO_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtNoMO.KeyDown
        Dim objCode As New FrmTrans_ManufacturingOrder_Cari

        Try
            If e.KeyValue = 13 Then
                If Me.TxtNoMO.Text <> "" Then
                    If objAuthMO.Apakahsudahada(Me.TxtNoMO.Text) = True Then
                        Me.TxtPrimaryMO.Text = objAuthMO.GetPrimary(Me.TxtNoMO.Text)
                        Me.TxtNoMO.Text = objAuthMO.GetNoTransaksi(Me.TxtPrimaryMO.Text)
                        Me.BtnFindMO.Focus()
                    Else
                        Me.BtnFindMO.PerformClick()
                    End If
                Else
                    Me.BtnFindMO.PerformClick()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
        End Try
    End Sub

    Private Sub TxtNoMO_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoMO.Leave
        Dim i As New FrmTrans_ManufacturingOrder_Cari

        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtNoMO.Text <> "" Then
                If objAuthMO.Apakahsudahada(Me.TxtNoMO.Text) = True Then
                    Me.TxtPrimaryMO.Text = objAuthMO.GetPrimary(Me.TxtNoMO.Text)
                    Me.TxtNoMO.Text = objAuthMO.GetNoTransaksi(Me.TxtPrimaryMO.Text)
                Else
                    i.txtCari.Text = Me.TxtNoMO.Text
                    i.ShowDialog()
                    If i.Primary > 0 Then
                        Me.TxtPrimaryMO.Text = i.Primary
                        Me.TxtNoMO.Text = objAuthMO.GetNoTransaksi(Me.TxtPrimaryMO.Text)
                    Else
                        Me.TxtNoMO.Clear()
                        Me.TxtPrimaryMO.Text = "-1"
                        Me.BtnFindMO.Focus()
                    End If
                End If
            Else
                Me.TxtNoMO.Clear()
                Me.TxtPrimaryMO.Text = "-1"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()

        End Try
    End Sub

    Private Sub BtnFindMO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindMO.Click
        Dim i As New FrmTrans_ManufacturingOrder_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtNoMO.Text <> "" Then
                i.txtCari.Text = Me.TxtNoMO.Text
            End If
            i.ShowDialog()
            If i.Primary > 0 Then
                Me.TxtPrimaryMO.Text = i.Primary
                Me.TxtNoMO.Text = objAuthMO.GetNoTransaksi(Me.TxtPrimaryMO.Text)
                Me.ChkMODate.Focus()
            Else
                Me.TxtNoMO.Clear()
                Me.TxtPrimaryMO.Text = "-1"
                Me.BtnFindMO.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub BtnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFilter.Click
        Try
            objAuthMO.IsiDGVAuthorize(TxtPrimaryMO.Text, Me.UserCode, Me.DgvOtorisasi, Me.DgvCancelOtorisasi, Me.DGVAllotorisasi, Me)
            If Me.TboAuthorize.SelectedIndex = 0 Then
                If DgvOtorisasi.RowCount > 0 Then
                    If Me.objAuthMO.CariDetail(Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, 0).Value, Me.DgvItemJO, TBLItemBOM, TBLProcessOrder, _
                            TBLItemWO, TBLItemRM, frm, objCD, Me.TBLItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, TBLJO_Temp) = True Then
                        Me.SalinKeField()
                    End If
                End If
            ElseIf Me.TboAuthorize.SelectedIndex = 1 Then
                If DgvCancelOtorisasi.RowCount > 0 Then
                    If Me.objAuthMO.CariDetail(Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, 0).Value, Me.DgvItemJO, TBLItemBOM, TBLProcessOrder, _
                            TBLItemWO, TBLItemRM, frm, objCD, Me.TBLItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, TBLJO_Temp) = True Then
                        Me.SalinKeField()
                    End If
                End If
            ElseIf Me.TboAuthorize.SelectedIndex = 2 Then
                If DGVAllotorisasi.RowCount > 0 Then
                    If Me.objAuthMO.CariDetail(Me.DGVAllotorisasi.Item(Me.col3_Primary.Index, 0).Value, Me.DgvItemJO, TBLItemBOM, TBLProcessOrder, _
                            TBLItemWO, TBLItemRM, frm, objCD, Me.TBLItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, TBLJO_Temp) = True Then
                        Me.SalinKeField()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbCheckAll_Authorize_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbCheckAll_Authorize.LinkClicked
        Try
            Dim i As Integer
            If Me.TboAuthorize.SelectedIndex = 0 Then
                For i = 0 To Me.DgvOtorisasi.RowCount - 1
                    Me.DgvOtorisasi.Item(Me.ColDgvAuth_Authorize.Index, i).Value = True
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub LlbUncheckAll_Athorize_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUncheckAll_Athorize.LinkClicked
        Try
            Dim i As Integer
            If Me.TboAuthorize.SelectedIndex = 0 Then
                For i = 0 To Me.DgvOtorisasi.RowCount - 1
                    Me.DgvOtorisasi.Item(Me.ColDgvAuth_Authorize.Index, i).Value = False
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub

    Private Sub LlbCheckAll_Cancel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbCheckAll_Cancel.LinkClicked
        Try
            Dim i As Integer
            If Me.TboAuthorize.SelectedIndex = 1 Then
                For i = 0 To Me.DgvCancelOtorisasi.RowCount - 1
                    Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Cancel.Index, i).Value = True
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub LlbUncheckAll_Cancel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUncheckAll_Cancel.LinkClicked
        Try
            Dim i As Integer
            If Me.TboAuthorize.SelectedIndex = 1 Then
                For i = 0 To Me.DgvCancelOtorisasi.RowCount - 1
                    Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Cancel.Index, i).Value = False
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub LlbCheckAll_All_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbCheckAll_All.LinkClicked
        Try
            Dim i As Integer
            If Me.TboAuthorize.SelectedIndex = 2 Then
                For i = 0 To Me.DGVAllotorisasi.RowCount - 1
                    Me.DGVAllotorisasi.Item(Me.col3_Print.Index, i).Value = True
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbUncheckAll_All_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUncheckAll_All.LinkClicked
        Try
            Dim i As Integer
            If Me.TboAuthorize.SelectedIndex = 2 Then
                For i = 0 To Me.DGVAllotorisasi.RowCount - 1
                    Me.DGVAllotorisasi.Item(Me.col3_Print.Index, i).Value = False
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub tsbAuthorize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbAuthorize.Click
        Dim i As Integer
        Dim dp As New Core_Module.DialogPassword
        Dim hasil As MsgBoxResult
        Dim frm As New FrmTrans_SPL
        Dim kosong As Boolean = True

        Dim objMO As New ClsTrans_ManufacturingOrder
        Dim listOfFail As New List(Of String)
        Dim frmMO As New FrmTrans_ManufacturingOrder

        Try
            If Me.TboAuthorize.SelectedIndex = 0 Then
                Me.DgvOtorisasi.EndEdit()
                For x As Integer = 0 To Me.DgvOtorisasi.RowCount - 1
                    If Me.DgvOtorisasi.Item(Me.ColDgvAuth_Authorize.Index, x).Value = True Then
                        kosong = False
                        Exit For
                    End If
                Next
            ElseIf Me.TboAuthorize.SelectedIndex = 1 Then
                Me.DgvCancelOtorisasi.EndEdit()
                For x As Integer = 0 To Me.DgvCancelOtorisasi.RowCount - 1
                    If Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Cancel.Index, x).Value = True Then
                        kosong = False
                        Exit For
                    End If
                Next
            End If

            If kosong = False Then
                hasil = MsgBox("" & Me.tsbAuthorize.Text & " ?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.DgvUser.Focus()
                    dp.txtUserName.Text = Me.UserCode
                    dp.ShowDialog()

                    If dp.DialogResult = Windows.Forms.DialogResult.OK Then
                        Me.KosongkanField()
                        If Me.TboAuthorize.SelectedIndex = 0 Then
                            For i = 0 To Me.DgvOtorisasi.RowCount - 1
                                If Me.DgvOtorisasi.Item(Me.ColDgvAuth_Authorize.Index, i).Value = True Then
                                    Me.objAuthMO.Authorize(Me.objUser.GetPrimary(Me.UserCode), Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, i).Value, 1, Me.UserCode)
                                End If
                            Next
                            objAuthMO.IsiDGVAuthorize(TxtPrimaryMO.Text, Me.UserCode, Me.DgvOtorisasi, Me.DgvCancelOtorisasi, Me.DGVAllotorisasi, Me)
                            If DgvOtorisasi.RowCount > 0 Then
                                If Me.objAuthMO.CariDetail(Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, Me.DgvOtorisasi.CurrentRow.Index).Value, Me.DgvItemJO, TBLItemBOM, TBLProcessOrder, _
                                    TBLItemWO, TBLItemRM, frmMO, objCD, Me.TBLItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, TBLJO_Temp) = True Then
                                    Me.SalinKeField()
                                End If
                            End If
                        ElseIf Me.TboAuthorize.SelectedIndex = 1 Then
                            For i = 0 To Me.DgvCancelOtorisasi.RowCount - 1
                                If Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Cancel.Index, i).Value = True Then
                                    If objMO.GetClosedJOCount(Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, i).Value) = 0 Then
                                        If objMO.GetScheduledMOCount(Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, i).Value) = 0 Then
                                            Me.objAuthMO.CancelAuthorize(Now, -1, Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, i).Value, 0, Me.UserCode)
                                        Else
                                            listOfFail.Add(Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_MONo.Index, i).Value)
                                        End If
                                    Else
                                        listOfFail.Add(Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_MONo.Index, i).Value)
                                    End If
                                End If
                            Next
                            objAuthMO.IsiDGVAuthorize(TxtPrimaryMO.Text, Me.UserCode, Me.DgvCancelOtorisasi, Me.DgvCancelOtorisasi, Me.DGVAllotorisasi, Me)
                            If DgvCancelOtorisasi.RowCount > 0 Then
                                If Me.objAuthMO.CariDetail(Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, Me.DgvCancelOtorisasi.CurrentRow.Index).Value, Me.DgvItemJO, TBLItemBOM, TBLProcessOrder, _
                                    TBLItemWO, TBLItemRM, frmMO, objCD, Me.TBLItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, TBLJO_Temp) = True Then
                                    Me.SalinKeField()
                                End If
                            End If
                        End If

                        If listOfFail.Count > 0 Then
                            Dim s0 As String = ""
                            For Each s1 As String In listOfFail
                                s0 += s1 & ","
                            Next
                            s0 = s0.Remove(s0.Length - 1, 1)
                            MsgBox("The following MO failed to process, because of some JO is closed or scheduled." & vbNewLine & s0, MsgBoxStyle.Information, "Information")
                        Else
                            MsgBox("" & Me.tsbAuthorize.Text & " Succeed", MsgBoxStyle.Information, "Information")
                        End If

                    End If
                End If
            Else
                MsgBox("There must be at least one manufacturing order that will be " & Me.tsbAuthorize.Text, MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub TboAuthorize_Click()
        Dim frm As New FrmTrans_ManufacturingOrder
        Try
            Me.cboStatus.Enabled = False
            Me.ChkMODate.Checked = False
            Me.DtpFrom.Value = Now
            Me.DtpTo.Value = Now
            Me.tsbAuthorize.Visible = True
            Me.TsbPrintAuth.Visible = False
            Me.KosongkanField()
            If Me.TboAuthorize.SelectedIndex = 0 Then
                Me.cboStatus.SelectedIndex = 2
                Me.tsbAuthorize.Text = "Authorized"
                objAuthMO.IsiDGVAuthorize(Me.TxtPrimaryMO.Text, Me.UserCode, Me.DgvOtorisasi, Me.DgvCancelOtorisasi, Me.DGVAllotorisasi, Me)
                If DgvOtorisasi.RowCount > 0 Then
                    If Me.objAuthMO.CariDetail(Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, 0).Value, Me.DgvItemJO, TBLItemBOM, TBLProcessOrder, _
                    TBLItemWO, TBLItemRM, frm, objCD, Me.TBLItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, TBLJO_Temp) = True Then
                        Me.SalinKeField()
                    End If
                End If
            ElseIf Me.TboAuthorize.SelectedIndex = 1 Then
                Me.cboStatus.SelectedIndex = 1
                Me.tsbAuthorize.Text = "Un-Authorized"
                objAuthMO.IsiDGVAuthorize(Me.TxtPrimaryMO.Text, Me.UserCode, Me.DgvOtorisasi, Me.DgvCancelOtorisasi, Me.DGVAllotorisasi, Me)
                If DgvCancelOtorisasi.RowCount > 0 Then
                    If Me.objAuthMO.CariDetail(Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, 0).Value, Me.DgvItemJO, TBLItemBOM, TBLProcessOrder, _
                    TBLItemWO, TBLItemRM, frm, objCD, Me.TBLItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, TBLJO_Temp) = True Then
                        Me.SalinKeField()
                    End If
                End If
            Else
                Me.cboStatus.SelectedIndex = 0
                Me.cboStatus.Enabled = True
                Me.ChkMODate.Checked = True
                Me.DtpFrom.Value = Now.AddMonths(-3)
                Me.tsbAuthorize.Visible = False
                Me.TsbPrintAuth.Visible = True
                objAuthMO.IsiDGVAuthorize(Me.TxtPrimaryMO.Text, Me.UserCode, Me.DgvOtorisasi, Me.DgvCancelOtorisasi, Me.DGVAllotorisasi, Me)
                If DGVAllotorisasi.RowCount > 0 Then
                    If Me.objAuthMO.CariDetail(Me.DGVAllotorisasi.Item(Me.col3_Primary.Index, 0).Value, Me.DgvItemJO, TBLItemBOM, TBLProcessOrder, _
                    TBLItemWO, TBLItemRM, frm, objCD, Me.TBLItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, TBLJO_Temp) = True Then
                        Me.SalinKeField()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemJO_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItemJO.SelectionChanged
        Dim f As New FrmTrans_ManufacturingOrder
        Try
            Me.DGVRM.Rows.Clear()

            If Me.DgvItemJO.SelectedCells.Count > 0 Then
                If String.IsNullOrEmpty(Me.DgvItemJO.Item(Me.ColItem_PrimaryProduct.Index, Me.DgvItemJO.SelectedCells.Item(0).RowIndex).Value) = False Then

                    If TabControl3.SelectedIndex = 0 Then
                        f.LoadDgvItemWO_TreeView(Me.DgvItemJO.Item(Me.ColItem_Identifier.Index, Me.DgvItemJO.SelectedCells.Item(0).RowIndex).Value, Me.TBLItemWO, _
                        Me.TGV_WO, f)
                    ElseIf TabControl3.SelectedIndex = 1 Then
                        f.LoadRMDetail(Me.DgvItemJO.Item(Me.ColItem_Identifier.Index, Me.DgvItemJO.SelectedCells.Item(0).RowIndex).Value, Me.TBLItemWO, _
                        Me.TBLItemRM, Me.DGVRM, f)
                    End If
                    f.Dispose()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TGV_WO_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles TGV_WO.CellBeginEdit
        Try
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TGV_WO_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles TGV_WO.CellContentClick
        Dim f As New FrmTrans_ManufacturingOrder
        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColWO2_ShowDetail.Index Then
                    Dim i As New FrmTrans_ManufacturingOrder_WO
                    i.vMONo = Me.txtDetailNoMO.Text
                    i.vJONo = objAuthMO.GetJONumber_InDgvJO(Me.TGV_WO.Item(Me.ColWO2_IdentifierJO.Index, e.RowIndex).Value, Me.DgvItemJO, f)
                    i.vIdentifierWO = Me.TGV_WO.Item(Me.ColWO2_IdentifierItem.Index, e.RowIndex).Value
                    i.vIdentifierjO = Me.TGV_WO.Item(Me.ColWO2_IdentifierJO.Index, e.RowIndex).Value
                    i.vTblItemWO = Me.TBLItemWO
                    i.vTblItemRM = Me.TBLItemRM
                    i.UserCode = Me.UserCode
                    'If Me.btnOK.Enabled = True Then
                    '    i.vBolehEdit = True
                    'End If
                    i.ShowDialog()
                    i.Dispose()

                    'Load ulang Tgv WO / Dgv RM
                    If Me.DgvItemJO.SelectedCells.Count > 0 Then
                        If String.IsNullOrEmpty(Me.DgvItemJO.Item(Me.ColItem_PrimaryProduct.Index, Me.DgvItemJO.SelectedCells.Item(0).RowIndex).Value) = False Then
                            If TabControl3.SelectedIndex = 0 Then
                                f.LoadDgvItemWO_TreeView(Me.DgvItemJO.Item(Me.ColItem_Identifier.Index, Me.DgvItemJO.SelectedCells.Item(0).RowIndex).Value, Me.TBLItemWO, _
                        Me.TGV_WO, f)
                            ElseIf TabControl3.SelectedIndex = 1 Then
                                f.LoadRMDetail(Me.DgvItemJO.Item(Me.ColItem_Identifier.Index, Me.DgvItemJO.SelectedCells.Item(0).RowIndex).Value, Me.TBLItemWO, _
                                Me.TBLItemRM, Me.DGVRM, f)
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl3.Click
        Dim f As New FrmTrans_ManufacturingOrder
        If Me.DgvItemJO.SelectedCells.Count > 0 Then
            If String.IsNullOrEmpty(Me.DgvItemJO.Item(Me.ColItem_PrimaryProduct.Index, Me.DgvItemJO.SelectedCells.Item(0).RowIndex).Value) = False Then
                If TabControl3.SelectedIndex = 0 Then
                    f.LoadDgvItemWO_TreeView(Me.DgvItemJO.Item(Me.ColItem_Identifier.Index, Me.DgvItemJO.SelectedCells.Item(0).RowIndex).Value, Me.TBLItemWO, _
                    Me.TGV_WO, f)
                ElseIf TabControl3.SelectedIndex = 1 Then
                    f.LoadRMDetail(Me.DgvItemJO.Item(Me.ColItem_Identifier.Index, Me.DgvItemJO.SelectedCells.Item(0).RowIndex).Value, Me.TBLItemWO, _
            Me.TBLItemRM, Me.DGVRM, f)
                End If
            End If
        End If
    End Sub

    Private Sub DgvOtorisasi_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvOtorisasi.CellContentClick
        'Dim frm As New FrmTrans_ManufacturingOrder
        'Me.KosongkanField()
        'Try
        '    If DgvOtorisasi.RowCount > 0 Then
        '        If Me.objAuthMO.CariDetail(Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, Me.DgvOtorisasi.CurrentRow.Index).Value, Me.DgvItemJO, TBLItemBOM, TBLProcessOrder, _
        '        TBLItemWO, TBLItemRM, frm, objCD, TBLJO_Temp) = True Then
        '            Me.SalinKeField()
        '        End If
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub DgvCancelOtorisasi_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvCancelOtorisasi.CellContentClick
        'Dim frm As New FrmTrans_ManufacturingOrder
        'Me.KosongkanField()
        'Try
        '    If DgvCancelOtorisasi.RowCount > 0 Then
        '        If Me.objAuthMO.CariDetail(Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, Me.DgvCancelOtorisasi.CurrentRow.Index).Value, Me.DgvItemJO, TBLItemBOM, TBLProcessOrder, _
        '        TBLItemWO, TBLItemRM, frm, objCD, TBLJO_Temp) = True Then
        '            Me.SalinKeField()
        '        End If
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub DGVAllotorisasi_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVAllotorisasi.CellContentClick
        'Dim frm As New FrmTrans_ManufacturingOrder
        'Me.KosongkanField()
        'Try
        '    If DGVAllotorisasi.RowCount > 0 Then
        '        If Me.objAuthMO.CariDetail(Me.DGVAllotorisasi.Item(Me.col3_Primary.Index, Me.DGVAllotorisasi.CurrentRow.Index).Value, Me.DgvItemJO, TBLItemBOM, TBLProcessOrder, _
        '        TBLItemWO, TBLItemRM, frm, objCD, TBLJO_Temp) = True Then
        '            Me.SalinKeField()
        '        End If
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub TsbPrintAuth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TsbPrintAuth.Click

        Dim frmTransMO_Print As New FrmTrans_ManufacturingOrder_Print
        Dim kosong As Boolean = True
        Try
            Me.DGVAllotorisasi.EndEdit()
            For x As Integer = 0 To Me.DGVAllotorisasi.RowCount - 1
                If Me.DGVAllotorisasi.Item(Me.col3_Print.Index, x).Value = True Then
                    kosong = False
                    Exit For
                End If
            Next

            If kosong = False Then
                frmTransMO_Print.val.Clear()
                Me.val.Clear()
                GetListPrimary()
                Dim i As New FrmTrans_ManufacturingOrder_PrintOption
                If i.ShowDialog <> Windows.Forms.DialogResult.OK Then
                    Exit Sub
                End If
                frmTransMO_Print.PaperSize = i.PaperSize
                frmTransMO_Print.val = Me.val
                frmTransMO_Print.pilihan = True
                frmTransMO_Print.PrimaryMO = Me.txtPrimary.Text
                frmTransMO_Print.ShowDialog()
            Else
                MsgBox("Please select the Manufacturing Order which will be in print", MsgBoxStyle.Information)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frmTransMO_Print.Dispose()
        End Try
    End Sub

    Private Function GetListPrimary() As List(Of Double)

        Try
            For Each dr As DataGridViewRow In Me.DGVAllotorisasi.Rows
                If dr.Cells(Me.col3_Primary.Index).Value <> -1 Then
                    If dr.Cells(Me.col3_Print.Index).Value = True Then
                        val.Add(dr.Cells(Me.col3_Primary.Index).Value)
                    End If
                End If
            Next
            Return val
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return val
        End Try
    End Function

    Private Sub tsbExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemJO_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvItemJO.DataError

    End Sub

    Private Sub DGVRM_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DGVRM.DataError

    End Sub

    Private Sub TGV_WO_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles TGV_WO.DataError

    End Sub

    ''ok-rosadin-r5b
    Private Sub TxtNoMO_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoMO.LostFocus
        Try
            If Me.TxtNoMO.Enabled Then
                If Not Me.TxtNoMO.ReadOnly Then
                    Me.TxtNoMO.BackColor = Color.White
                End If
            Else
                Me.TxtNoMO.BackColor = Me.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.GotFocus
        Try
            If Me.txtCodeNo.Enabled Then
                If Not Me.txtCodeNo.ReadOnly Then
                    Me.txtCodeNo.BackColor = Me.BackColorAll
                End If
            Else
                Me.txtCodeNo.BackColor = Me.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''ok-rosadin-r5b
    Dim objProduct As New Core_Module.clsMasterProduct
    Private Sub txtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCodeNo.KeyDown
        Try
            If e.KeyCode = 13 Then
                If Me.objProduct.ApakahSudahAda(Me.txtCodeNo.Text) Then
                    Me.txtPrimaryCodeNo.Text = Me.objProduct.GetPrimary(Me.txtCodeNo.Text)
                    Me.BtnFilter.Focus()
                Else
                    Me.btnFindProductCode.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''ok-rosaidn-r5b
    Private Sub txtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.Leave
        Try
            If Me.txtCodeNo.Text.Trim <> "" Then
                If Me.objProduct.ApakahSudahAda(Me.txtCodeNo.Text) Then
                    Me.txtPrimaryCodeNo.Text = Me.objProduct.GetPrimary(Me.txtCodeNo.Text)
                Else
                    Me.btnFindProductCode.PerformClick()
                End If
            Else
                Me.txtPrimaryCodeNo.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.LostFocus
        Try
            If Me.txtCodeNo.Enabled Then
                If Not Me.txtCodeNo.ReadOnly Then
                    Me.txtCodeNo.BackColor = Color.White
                End If
            Else
                Me.txtCodeNo.BackColor = Me.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''ok-rosaidn-r5b
    Private Sub btnFindProductCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindProductCode.Click
        Dim frmFindCode As New Core_Module.frmCariMasterProduct
        Try
            frmFindCode.txtCari.Text = Me.txtCodeNo.Text
            frmFindCode.ShowDialog()
            If frmFindCode.PrimaryProduct > 0 Then
                Me.txtPrimaryCodeNo.Text = frmFindCode.PrimaryProduct
                Me.txtCodeNo.Text = Me.objProduct.GetProductCode(frmFindCode.PrimaryProduct)
                Me.BtnFilter.Focus()
            Else
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtCodeNo.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVRM_Sorted1(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVRM.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DGVRM)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvOtorisasi_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvOtorisasi.SelectionChanged
        Try
            If Me.DgvOtorisasi.SelectedCells.Count > 0 And Me.TboAuthorize.SelectedIndex = 0 Then
                If IsNothing(Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, Me.DgvOtorisasi.SelectedCells(0).RowIndex).Value) Then
                    Exit Sub
                End If
                If Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, Me.DgvOtorisasi.SelectedCells(0).RowIndex).Value.ToString <> Me.TxtPrimaryMO.Text Then
                    If Me.objAuthMO.CariDetail(Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, Me.DgvOtorisasi.CurrentRow.Index).Value, Me.DgvItemJO, _
                    TBLItemBOM, TBLProcessOrder, TBLItemWO, _
                    TBLItemRM, frm, objCD, Me.TBLItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, TBLJO_Temp) = True Then
                        Me.SalinKeField()
                    End If
                End If
            Else
                Me.KosongkanField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCancelOtorisasi_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCancelOtorisasi.SelectionChanged
        Try
            If Me.DgvCancelOtorisasi.SelectedCells.Count > 0 And Me.TboAuthorize.SelectedIndex = 1 Then
                If IsNothing(Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, Me.DgvCancelOtorisasi.SelectedCells(0).RowIndex).Value) Then
                    Exit Sub
                End If
                If Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, Me.DgvCancelOtorisasi.SelectedCells(0).RowIndex).Value.ToString <> Me.TxtPrimaryMO.Text Then
                    If Me.objAuthMO.CariDetail(Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, Me.DgvCancelOtorisasi.CurrentRow.Index).Value, Me.DgvItemJO, _
                    TBLItemBOM, TBLProcessOrder, TBLItemWO, TBLItemRM, frm, objCD, Me.TBLItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, TBLJO_Temp) = True Then
                        Me.SalinKeField()
                    End If
                End If
            Else
                Me.KosongkanField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAllotorisasi_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVAllotorisasi.SelectionChanged
        Try
            If Me.DGVAllotorisasi.SelectedCells.Count > 0 And Me.TboAuthorize.SelectedIndex = 2 Then
                If IsNothing(Me.DGVAllotorisasi.Item(Me.ColDgvUnAuth_Primary.Index, Me.DGVAllotorisasi.SelectedCells(0).RowIndex).Value) Then
                    Exit Sub
                End If
                If Me.DGVAllotorisasi.Item(Me.ColDgvUnAuth_Primary.Index, Me.DGVAllotorisasi.SelectedCells(0).RowIndex).Value.ToString <> Me.TxtPrimaryMO.Text Then
                    If Me.objAuthMO.CariDetail(Me.DGVAllotorisasi.Item(Me.col3_Primary.Index, Me.DGVAllotorisasi.CurrentRow.Index).Value, Me.DgvItemJO, _
                    TBLItemBOM, TBLProcessOrder, TBLItemWO, TBLItemRM, frm, objCD, Me.TBLItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, TBLJO_Temp) = True Then
                        Me.SalinKeField()
                    End If
                End If
            Else
                Me.KosongkanField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
 
End Class