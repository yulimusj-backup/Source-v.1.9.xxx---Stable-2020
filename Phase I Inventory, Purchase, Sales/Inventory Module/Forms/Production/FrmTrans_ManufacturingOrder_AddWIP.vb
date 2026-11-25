Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_AddWIP
    Public vIdentifierJO As String = ""
    Public vTblItemWO_Original As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
    Public vTblItemRM_Original As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
    Public vTblItemWO As New DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
    Public vTblItemRM As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
    Public vTblItemBoM_Original As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vUserCode As String = ""
    Public vJONO As String = ""
    Public vMONo As String = ""
    Public vMODate As Date = Today
    Public vPrimaryMO As Double = -1

    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjMO As New ClsTrans_ManufacturingOrder
    Dim ObjMP As New ClsTransaksiMaterialPlanning_Calc
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser
    Dim LangsungClose As Boolean = False
    Dim ListPrimaryProduct As New List(Of Double)
    Dim Buffer As String = ""
    Dim DefaultBackColor1 As Color = Color.White
    Dim ijkl As New FrmTrans_MaterialPlanningDashboard_Adv_Exclusion2

    Dim vDT As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_item2DataTable
    Dim vDT_ItemGIT As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable
    Dim vDT_ItemDO As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable
    Dim vDT_ItemPOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable
    Dim vDT_ItemSOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable
    Dim vDT_ItemPrOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable
    Dim vDT_ItemSettingLocation As New DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable
    Dim vDT_ItemQty As New DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable
    Dim vDT_ItemMOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable
    Dim vDT_ItemRMReq As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable

    Private Sub LoadDGV()
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim rowD As DataGridViewRow
        Dim ObjProses As New Core_Module.ClsMasterProses
        Dim vDecProd_Item As Integer = 0

        Try
            Me.Panel1.Visible = True

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGVWIP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGVWIP.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            DGVWIP.AllowUserToResizeColumns = False
            DGVWIP.AllowUserToResizeRows = False

            If DGVWIP.Rows.Count > 0 Then
                DGVWIP.Rows.Clear()
            End If

            'hitung qty available
            ijkl.UserCode = vUserCode
            ijkl.TipeTransaksiExclusion = ClsTransaksiMaterialPlanning.EnumTipeTransaksiExclusion.MaterialPlanning
            ijkl.vListPrimaryCodeNo = Me.ListPrimaryProduct
            'ijkl.TopLevel = True
            ijkl.ShowDialog()

            Me.ControlBox = False
            Me.btnCancel.Enabled = False
            Me.btnOK.Enabled = False

            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjProses = Nothing
        End Try
    End Sub

    Private Sub SetRowVisible()
        Dim rowD As DataGridViewRow
        Dim Dummy As Decimal = 0

        Try
            If Me.chkQtyAvailableOnly.Checked = False Then
                For Each rowD In Me.DGVWIP.Rows
                    rowD.Visible = True
                    If Math.IEEERemainder(CDbl(rowD.Index), 2) = 0 Then
                        'genap
                        rowD.DefaultCellStyle.BackColor = Me.DefaultBackColor1
                    Else
                        rowD.DefaultCellStyle.BackColor = Me.BackColorAll
                    End If
                Next
                Me.ObjCD.SetNumber(DGVWIP)
                'Me.DGVWIP.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Else
                For Each rowD In Me.DGVWIP.Rows
                    If CDec(rowD.Cells(Me.ColWIP_QtyAvailable.Index).Value) > 0 Or CDec(rowD.Cells(Me.ColWIP_QtyWIP.Index).Value) > 0 Then
                        rowD.Visible = True
                    Else
                        rowD.Visible = False
                    End If
                Next
                Me.ObjCD.SetNumber(DGVWIP)


                Me.DGVWIP.AlternatingRowsDefaultCellStyle.BackColor = Me.DGVWIP.DefaultCellStyle.BackColor
                For ij As Integer = 0 To Me.DGVWIP.Rows.Count - 1
                    rowD = Me.DGVWIP.Rows(ij)
                    If rowD.Visible = True Then
                        Dummy = rowD.Cells(Me.ColWIP_No.Index).Value
                        If Math.IEEERemainder(Dummy, 2) = 0 Then
                            'genap
                            rowD.DefaultCellStyle.BackColor = Me.BackColorAll
                        Else
                            rowD.DefaultCellStyle.BackColor = Me.DefaultBackColor1
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetQtyAvailable(ByVal vPrimaryProduct As Double) As Decimal
        Dim vRowDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item2Row

        Try
            For Each vRowDT In Me.vDT.Rows
                If vRowDT.primary_code_no = vPrimaryProduct Then
                    Return vRowDT.qty_required
                End If
            Next

            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FrmTrans_ManufacturingOrder_AddWIP_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.LangsungClose = False Then
                If MsgBox("Close form and discard data change?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_AddWIP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        
    End Sub

    Private Sub TxtMP_QtySOB_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtySOB.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainSOB
        Dim rowDGV As DataGridViewRow

        Try
            If Me.DGVWIP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.DGVWIP.Rows(Me.DGVWIP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtySOB.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColWIP_PrimaryCodeNo.Index).Value
                    i.vDT_ItemSOB = Me.vDT_ItemSOB
                    i.BackColorAll = Me.BackColorAll

                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyPrOB_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyPrOB.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainPrOB
        Dim rowDGV As DataGridViewRow

        Try
            If Me.DGVWIP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.DGVWIP.Rows(Me.DGVWIP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyPrOB.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColWIP_PrimaryCodeNo.Index).Value
                    i.CodeNo = rowDGV.Cells(Me.ColWIP_CodeNo.Index).Value
                    i.UnitCode = rowDGV.Cells(Me.ColWIP_UnitCode.Index).Value

                    i.vDT_ItemPrOB = Me.vDT_ItemPrOB
                    i.BackColorAll = Me.BackColorAll

                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyProductionBalance_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyMOB.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainMOB
        Dim rowDGV As DataGridViewRow

        Try
            If Me.DGVWIP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.DGVWIP.Rows(Me.DGVWIP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyMOB.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColWIP_PrimaryCodeNo.Index).Value
                    i.vDT_ItemMOB = Me.vDT_ItemMOB
                    i.BackColorAll = Me.BackColorAll

                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyRMReq_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyRMReq.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainRMReq
        Dim rowDGV As DataGridViewRow

        Try
            If Me.DGVWIP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.DGVWIP.Rows(Me.DGVWIP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyRMReq.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColWIP_PrimaryCodeNo.Index).Value
                    i.vDT_ItemRMReq = Me.vDT_ItemRMReq
                    i.BackColorAll = Me.BackColorAll
                    i.UserCode = Me.vUserCode

                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyPendingDO_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyPendingDO.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainPendingDO
        Dim rowDGV As DataGridViewRow

        Try
            If Me.DGVWIP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.DGVWIP.Rows(Me.DGVWIP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyPendingDO.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColWIP_PrimaryCodeNo.Index).Value
                    i.vDT_ItemDO = Me.vDT_ItemDO
                    i.BackColorAll = Me.BackColorAll

                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyPOB_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyPOB.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainPOB
        Dim rowDGV As DataGridViewRow

        Try
            If Me.DGVWIP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.DGVWIP.Rows(Me.DGVWIP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyPOB.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColWIP_PrimaryCodeNo.Index).Value
                    i.vDT_ItemPOB = Me.vDT_ItemPOB
                    i.BackColorAll = Me.BackColorAll

                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyGIT_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyGIT.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainGIT
        Dim rowDGV As DataGridViewRow

        Try
            If Me.DGVWIP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.DGVWIP.Rows(Me.DGVWIP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyGIT.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColWIP_PrimaryCodeNo.Index).Value
                    i.vDT_ItemGIT = Me.vDT_ItemGIT
                    i.BackColorAll = Me.BackColorAll

                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyOnHand_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyOnHand.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainQtyOnHand
        Dim rowDGV As DataGridViewRow

        Try
            If Me.DGVWIP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.DGVWIP.Rows(Me.DGVWIP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyOnHand.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColWIP_PrimaryCodeNo.Index).Value
                    i.UnitCode = rowDGV.Cells(Me.ColWIP_UnitCode.Index).Value
                    i.DefPrec = rowDGV.Cells(Me.ColWIP_DefPrecUnit.Index).Value
                    i.vDT_ItemQty = Me.vDT_ItemQty
                    i.vDT_ItemSettingLocation = Me.vDT_ItemSettingLocation
                    i.vBackColorAll = Me.BackColorAll

                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i = Nothing
        End Try
    End Sub

    Private Sub TxtMP_QtyAvailable_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyAvailable.MouseClick
        'Dim Dummy As String = ""

        'Try
        '    If Me.DGVWIP.SelectedCells.Count > 0 Then
        '        Dummy = "Qty Available = On Hand + Goods In Transit + Purchase Balance + Production Balance  - Purch. Request - Sales Balance - Pending DO - Raw Material Req."
        '        Dummy = Dummy + vbCrLf + Me.TxtMP_QtyOnHand.Text + " + " + Me.TxtMP_QtyGIT.Text + " + " + Me.TxtMP_QtyPOB.Text + " + " + Me.TxtMP_QtyMOB.Text + _
        '        " - " + Me.TxtMP_QtyPrOB.Text + " - " + Me.TxtMP_QtySOB.Text + " - " + Me.TxtMP_QtyPendingDO.Text + " - " + Me.TxtMP_QtyRMReq.Text + " - "
        '        Dummy = Dummy + vbCrLf + Me.TxtMP_QtyAvailable.Text + " " + Me.DGVWIP.Item(Me.ColWIP_UnitCode.Index, Me.DGVWIP.SelectedCells(0).RowIndex).Value.ToString
        '        MsgBox(Dummy, MsgBoxStyle.Information, "Information")
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
        Dim ij As New FrmTrans_ManufacturingOrder_AddWIP1
        Dim DefPrec As Integer = 0

        Try
            If Me.DGVWIP.SelectedCells.Count > 0 Then
                ij.TxtMP_QtyAvailable.Text = Me.TxtMP_QtyAvailable.Text
                ij.TxtMP_QtyGIT.Text = Me.TxtMP_QtyGIT.Text
                ij.TxtMP_QtyMOB.Text = Me.TxtMP_QtyMOB.Text
                ij.TxtMP_QtyOnHand.Text = Me.TxtMP_QtyOnHand.Text
                ij.TxtMP_QtyPendingDO.Text = Me.TxtMP_QtyPendingDO.Text
                ij.TxtMP_QtyPOB.Text = Me.TxtMP_QtyPOB.Text
                ij.TxtMP_QtyPrOB.Text = Me.TxtMP_QtyPrOB.Text
                ij.TxtMP_QtyRMReq.Text = Me.TxtMP_QtyRMReq.Text
                ij.TxtMP_QtySOB.Text = Me.TxtMP_QtySOB.Text
                DefPrec = Me.DGVWIP.Item(Me.ColWIP_DefPrecUnit.Index, Me.DGVWIP.SelectedCells(0).RowIndex).Value
                ij.TextBox1.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtMP_QtyOnHand.Text) + CDec(Me.TxtMP_QtyGIT.Text) + CDec(Me.TxtMP_QtyMOB.Text) + CDec(Me.TxtMP_QtyPOB.Text), DefPrec)
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVWIP_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVWIP.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColWIP_QtyWIP.Index Then
                Me.Buffer = Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value
                Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value)
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVWIP_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVWIP.CellEndEdit
        'Dim vRow_ItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqRow
        'Dim vRow As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item2Row
        Dim QtyRMReq As Decimal = 0
        Dim vRowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim vRowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim vRowItemBoM_Original As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim Selisih As Decimal = 0

        Try
            If IsNothing(Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            Else
                If e.ColumnIndex = Me.ColWIP_QtyWIP.Index Then
                    'kalo nilainya sama dengan lama, maka tidak usah susah2 jalan coding di bawah
                    If CDec(Me.Buffer) = CDec(Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value) Then
                        Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'pastikan non negatif
                    If Me.ObjCD.CekNonNegatif(Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'digit grouping
                    Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value, _
                    Me.DGVWIP.Item(Me.ColWIP_DefPrecUnit.Index, e.RowIndex).Value)

                    'jika qty available < 0, maka konfirmasi ulang
                    Selisih = CDec(Me.Buffer) - CDec(Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value)
                    If CDec(Me.DGVWIP.Item(Me.ColWIP_QtyAvailable.Index, e.RowIndex).Value) + Selisih < 0 Then
                        If MsgBox("Qty WIP is larger than Qty Available, this could mean new MO is needed. Continue anyway?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                            Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If

                    'pastikan qty wip <= qty required -> qty produksi minimal 0, ga boleh nol
                    If CDec(Me.DGVWIP.Item(Me.ColWIP_QtyWIP.Index, e.RowIndex).Value) > _
                    CDec(Me.DGVWIP.Item(Me.ColWIP_QtyReq.Index, e.RowIndex).Value) Then
                        MsgBox("Qty WIP must be smaller than Qty Required.", MsgBoxStyle.Information, "Information")
                        Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'pada vDT_ItemRMReq, update baris ybs, hapus baris, atau buar baris baru, tergantung qty end edit 
                    If Me.UpdateItemRMReq(Me.DGVWIP.Item(Me.ColWIP_IdentifierWO.Index, e.RowIndex).Value, _
                    Me.DGVWIP.Item(Me.ColWIP_PrimaryCodeNo.Index, e.RowIndex).Value, _
                    Me.DGVWIP.Item(Me.ColWIP_QtyWIP.Index, e.RowIndex).Value, _
                    Me.DGVWIP.Item(Me.ColWIP_PrimaryProses.Index, e.RowIndex).Value, _
                    Me.DGVWIP.Item(Me.ColWIP_WONo.Index, e.RowIndex).Value) = False Then
                        Me.DGVWIP.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'pada vdt, update qty rm req, qty available, hitung ulang qty required
                    Me.UpdateQtyAvailable(Me.DGVWIP.Item(Me.ColWIP_PrimaryCodeNo.Index, e.RowIndex).Value)

                    'update qty required pada tabel item rm
                    For Each vRowItemRM In Me.vTblItemRM.Rows
                        If vRowItemRM.Identifier_WO = Me.DGVWIP.Item(Me.ColWIP_IdentifierWO.Index, e.RowIndex).Value And _
                        vRowItemRM.Identifier_JO = Me.DGVWIP.Item(Me.ColWIP_IdentifierJO.Index, e.RowIndex).Value And _
                        vRowItemRM.Primary_Product_Code_RM = Me.DGVWIP.Item(Me.ColWIP_PrimaryCodeNo.Index, e.RowIndex).Value Then
                            vRowItemRM.WIP_Qty = Me.DGVWIP.Item(Me.ColWIP_QtyWIP.Index, e.RowIndex).Value
                            vRowItemRM.Balance_Qty = vRowItemRM.Required_Qty + vRowItemRM.Additional_Qty - vRowItemRM.WIP_Qty - vRowItemRM.Used_Qty - vRowItemRM.Replace_Qty

                            'update qty wip, qty target produce, qty target result, qty balance produce, qty balance result, qty balance mp pada tabel item wo
                            'yang memproduksi wip yg kita alokasikan
                            For Each vRowItemWO In Me.vTblItemWO.Rows
                                If vRowItemWO.Identifier = Me.DGVWIP.Item(Me.ColWIP_IdentifierWO.Index, e.RowIndex).Value And _
                                vRowItemWO.Identifier_JO = Me.vIdentifierJO Then
                                    For Each vRowItemBoM_Original In Me.vTblItemBoM_Original.Rows
                                        If vRowItemBoM_Original.Primary_Product_RM = Me.DGVWIP.Item(Me.ColWIP_PrimaryCodeNo.Index, e.RowIndex).Value And _
                                        vRowItemBoM_Original.Identifier_WO = Me.DGVWIP.Item(Me.ColWIP_IdentifierWO.Index, e.RowIndex).Value Then
                                            IsiTblItemWORM(vRowItemWO.Level, vRowItemRM.Primary_Product_Code_RM, vRowItemRM.Balance_Qty, vTblItemWO, vTblItemRM, _
                                            Me.vTblItemBoM_Original, vIdentifierJO, vRowItemWO.Identifier, vRowItemBoM_Original.Identifier_BOM)
                                            Exit For
                                        End If
                                    Next
                                    Exit For
                                End If
                            Next

                            Exit For
                        End If
                    Next

                    'hitung ulang qty produksi
                    'Me.DGVWIP.Item(Me.ColWIP_QtyProduce.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DGVWIP.Item(Me.ColWIP_QtyReq.Index, e.RowIndex).Value - _
                    'Me.DGVWIP.Item(Me.ColWIP_QtyWIP.Index, e.RowIndex).Value, Me.DGVWIP.Item(Me.ColWIP_DefPrecUnit.Index, e.RowIndex).Value)

                    'update qty available, untuk semua row DGV yang primary code no-nya sama
                    Me.UpdateQty()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateQtyAvailable(ByVal vPrimaryRM As Double)
        Dim vRow_ItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqRow
        Dim vRow As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item2Row
        Dim QtyRMReq As Decimal = 0

        Try
            For Each vRow In Me.vDT.Rows
                If vRow.primary_code_no = vPrimaryRM Then
                    QtyRMReq = 0
                    For Each vRow_ItemRMReq In Me.vDT_ItemRMReq.Rows
                        If vRow_ItemRMReq.primary_product = vRow.primary_code_no Then
                            QtyRMReq += vRow_ItemRMReq.qty
                        End If
                    Next

                    vRow.qty_rm_req = QtyRMReq
                    vRow.qty_required = vRow.qty_on_hand + vRow.qty_pob + vRow.qty_mob + vRow.qty_git - vRow.qty_inquiry _
                    - vRow.qty_sob - vRow.qty_prob - vRow.qty_pending_do - vRow.qty_rm_req

                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function UpdateItemRMReq(ByVal vIdentiferWO As String, ByVal vPrimaryCodeNoWIP As Double, ByVal vQtyWIP As Decimal, _
    ByVal vPrimaryProses As Double, ByVal vWONo As String) As Boolean
        Dim vRow_ItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqRow
        Try
            Me.vDT_ItemRMReq.DefaultView.RowFilter = "[identifier jo rec] = '" + Me.vIdentifierJO + _
                                   "' AND [identifier wo rec] = '" + vIdentiferWO + _
                                   "' AND [primary product] = " + vPrimaryCodeNoWIP.ToString + _
                                   " AND [is wip] = 1"
            If CDec(vQtyWIP) = 0 Then
                'pastikan hapus baris item rm req, yang bersesuai dengan baris tersebut
                If Me.vDT_ItemRMReq.DefaultView.Count <> 0 And Me.vDT_ItemRMReq.DefaultView.Count <> 1 Then
                    MsgBox("Logic error: wrong asumtion. We're asuming unique row for the filter above, multiple rows found. Please contact your vendor if you see this message.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                Else
                    While Me.vDT_ItemRMReq.DefaultView.Count > 0
                        Me.vDT_ItemRMReq.Rows.Remove(Me.vDT_ItemRMReq.DefaultView(0).Row)
                    End While
                End If
            Else
                If Me.vDT_ItemRMReq.DefaultView.Count = 1 Then
                    '1. jika ada baris item rm, maka edit qty-nya sesuai dengan qty wip
                    vRow_ItemRMReq = Me.vDT_ItemRMReq.DefaultView(0).Row
                    vRow_ItemRMReq.qty = CDec(vQtyWIP)
                ElseIf Me.vDT_ItemRMReq.DefaultView.Count = 0 Then
                    '2. jika tidak ada baris item rm-nya, maka bikin baris baru
                    vRow_ItemRMReq = Me.vDT_ItemRMReq.Newtbltrans_material_planning_item_rm_reqRow
                    With vRow_ItemRMReq
                        .identifier_jo_rec = Me.vIdentifierJO
                        .identifier_wo_rec = vIdentiferWO
                        .is_wip = 1
                        .jo_no_rec = Me.vJONO
                        .mo_date_rec = Me.vMODate
                        .mo_no_rec = Me.vMONo
                        .primary = -1
                        .primary_mo_rec = Me.vPrimaryMO
                        .primary_product = vPrimaryCodeNoWIP
                        .primary_proses = vPrimaryProses
                        .primary_transaksi = -1
                        .qty = vQtyWIP
                        .wo_no_rec = vWONo
                        .MP_Spare_Qty = 0
                        .required_date = #1/1/1950#
                    End With
                    Me.vDT_ItemRMReq.Rows.Add(vRow_ItemRMReq)
                Else
                    MsgBox("Logic error: wrong asumtion. We're asuming unique row for the filter above, multiple rows found. Please contact your vendor if you see this message.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub IsiTblItemWORM(ByVal vLevel As Integer, ByVal vPrimaryProduct As Double, ByVal vQtyTarget As Decimal, _
    ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, _
    ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vIdentifierJO As String, ByVal vIdentifierParent As String, _
    ByVal vIdentifierBOM_Parent As String)
        Dim vTblItemBoM_ As New DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
        Dim rowItemBOM_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim dec_Product As Integer = 0

        Try
            vTblItemBoM_ = vTblItemBOM.Copy

            vTblItemBoM_.DefaultView.RowFilter = "[primary product] = " & vPrimaryProduct.ToString & " AND [Identifier BOM Parent] = '" & _
            vIdentifierBOM_Parent.ToString & "' AND [Identifier JO] = '" & vIdentifierJO & "'"
            If vTblItemBoM_.DefaultView.Count > 0 Then
                dec_Product = Me.ObjProduct.GetAngkaDecimal(vPrimaryProduct)
                rowItemBOM_ = vTblItemBoM_.DefaultView(0).Row
                vTblItemWO.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString & " AND [Identifier] = '" _
                & rowItemBOM_.Identifier_WO.ToString & "' AND [Identifier JO] = '" & vIdentifierJO & "'"
                If vTblItemWO.DefaultView.Count = 1 Then
                    rowItemWO = vTblItemWO.DefaultView(0).Row
                    With rowItemWO
                        .Primary_Process = rowItemBOM_.Primary_Process
                        .Qty_Target_Result = Me.ObjCD.RoundUp1(vQtyTarget, dec_Product)
                        .Qty_Target_Prod = Me.ObjCD.RoundUp1(vQtyTarget / (rowItemBOM_.Good_Percentage_Process / 100), dec_Product)
                        .Plan_Good_Perct = rowItemBOM_.Good_Percentage_Process
                        .Qty_Projected_Result = Me.ObjCD.RoundDown1((.Qty_Target_Prod + .Qty_Target_Prod_Add - .Qty_Actual_Prod) * (.Plan_Good_Perct / 100) + .Qty_Actual_Result, dec_Product)
                        .Actual_Good_Perct = 0
                        .Balance_Produce = Me.ObjCD.RoundUp1(.Qty_Target_Prod - .Qty_Actual_Prod, dec_Product)
                        .Balance_Result = Me.ObjCD.RoundUp1(.Qty_Target_Result - .Qty_Actual_Result, dec_Product)
                        .Qty_Balance_MP = Me.ObjCD.RoundUp1(.Qty_Target_Result - .Qty_Actual_Result, dec_Product)
                    End With
                Else
                    MsgBox("Logic error: wrong table filter. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                For i As Integer = 0 To vTblItemBoM_.DefaultView.Count - 1
                    rowItemBOM_ = vTblItemBoM_.DefaultView(i).Row
                    vTblItemRM.DefaultView.RowFilter = "[Primary Product Code RM] = " & rowItemBOM_.Primary_Product_RM.ToString & _
                    " AND [Identifier WO] = '" & rowItemWO.Identifier.ToString & "' AND [Identifier JO] = '" & vIdentifierJO & "'"
                    If vTblItemRM.DefaultView.Count = 1 Then
                        rowItemRM = vTblItemRM.DefaultView(0).Row
                        With rowItemRM
                            .Required_Qty = Me.ObjCD.RoundUp1((rowItemWO.Qty_Target_Prod / rowItemBOM_.Qty_Ref_Process * rowItemBOM_.Qty_RM) / (rowItemBOM_.Good_Percentage_RM / 100), rowItemBOM_.Decimal_RM)
                            .Allocated_Qty = Me.ObjCD.RoundUp1(.Required_Qty, rowItemBOM_.Decimal_RM)
                            .Balance_Qty = Me.ObjCD.RoundUp1(.Allocated_Qty + .Additional_Qty - .Used_Qty - .WIP_Qty - .Replace_Qty, rowItemBOM_.Decimal_RM)
                            .Target_Qty = Me.ObjCD.RoundUp1((rowItemWO.Qty_Target_Prod / rowItemBOM_.Qty_Ref_Process) * rowItemBOM_.Qty_RM, rowItemBOM_.Decimal_RM)

                            If rowItemRM.is_wip = 1 Then
                                Me.UpdateItemRMReq(rowItemRM.Identifier_WO, rowItemRM.Primary_Product_Code_RM, rowItemRM.WIP_Qty, _
                                rowItemRM.Primary_Product_Code_RM, Me.GetWONumber(rowItemRM.Identifier_WO))
                                Me.UpdateQtyAvailable(rowItemRM.Primary_Product_Code_RM)
                            End If
                        End With
                    Else
                        MsgBox("Logic error: wrong table filter. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    IsiTblItemWORM(rowItemWO.Level, rowItemRM.Primary_Product_Code_RM, rowItemRM.Balance_Qty, vTblItemWO, vTblItemRM, _
                    vTblItemBOM, vIdentifierJO, rowItemWO.Identifier, rowItemBOM_.Identifier_BOM)
                Next

            End If
        Catch ex As Exception
            MsgBox(ex.Message & "," & vLevel & "," & vPrimaryProduct & "," & vQtyTarget)
        Finally
            vTblItemBoM_.Dispose()
        End Try
    End Sub

    Private Function GetWONumber(ByVal vIdentifierWO As String) As String
        Dim vRowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow

        Try
            For Each vRowItemWO In Me.vTblItemWO.Rows
                If vRowItemWO.Identifier = vIdentifierWO And _
                vRowItemWO.Identifier_JO = Me.vIdentifierJO Then
                    Return vRowItemWO.WO_Number
                End If
            Next
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub UpdateQty()
        Dim vRow As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item2Row
        Dim rowDGV As DataGridViewRow
        Dim vRowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim vDecProd_Item As Integer = 0

        Try
            Me.DGVWIP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'untuk semua baris dalam dgv, cari item rm yang bersesuaian, kemudian update nilai2nya
            For Each rowDGV In Me.DGVWIP.Rows
                'qty required, qty produce, qty wip, qty available
                vDecProd_Item = rowDGV.Cells(Me.ColWIP_DefPrecUnit.Index).Value
                For Each vRowItemRM In Me.vTblItemRM.Rows
                    If vRowItemRM.Primary_Product_Code_RM = rowDGV.Cells(Me.ColWIP_PrimaryCodeNo.Index).Value And _
                    vRowItemRM.Identifier_WO = rowDGV.Cells(Me.ColWIP_IdentifierWO.Index).Value And _
                    vRowItemRM.Identifier_JO = Me.vIdentifierJO Then
                        With vRowItemRM
                            rowDGV.Cells(Me.ColWIP_QtyProduce.Index).Value = Me.ObjCD.DigitGrouping(.Required_Qty + .Additional_Qty - .WIP_Qty, vDecProd_Item)
                            rowDGV.Cells(Me.ColWIP_QtyReq.Index).Value = Me.ObjCD.DigitGrouping(.Required_Qty + .Additional_Qty, vDecProd_Item)
                            rowDGV.Cells(Me.ColWIP_QtyWIP.Index).Value = Me.ObjCD.DigitGrouping(.WIP_Qty, vDecProd_Item)

                            For Each vRow In Me.vDT.Rows
                                If vRow.primary_code_no = rowDGV.Cells(Me.ColWIP_PrimaryCodeNo.Index).Value Then
                                    rowDGV.Cells(Me.ColWIP_QtyAvailable.Index).Value = Me.ObjCD.DigitGrouping(vRow.qty_required, vDecProd_Item)
                                    vRowItemRM.Available_Qty = Me.ObjCD.DigitGrouping(vRow.qty_required, vDecProd_Item)
                                    Exit For
                                End If
                            Next
                        End With

                        Exit For
                    End If
                Next
            Next
            Me.DGVWIP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVWIP_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVWIP.SelectionChanged
        Dim vRowDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item2Row
        Dim rowDGV As DataGridViewRow

        Try
            If Me.DGVWIP.SelectedCells.Count = 0 Then
                Me.Kosongkan_TxtQty()
            Else
                rowDGV = Me.DGVWIP.Rows(Me.DGVWIP.SelectedCells(0).RowIndex)
                For Each vRowDT In Me.vDT.Rows
                    If vRowDT.primary_code_no = rowDGV.Cells(Me.ColWIP_PrimaryCodeNo.Index).Value Then
                        Me.TxtMP_QtyAvailable.Text = Me.ObjCD.DigitGrouping(vRowDT.qty_required, rowDGV.Cells(Me.ColWIP_DefPrecUnit.Index).Value)
                        Me.TxtMP_QtyGIT.Text = Me.ObjCD.DigitGrouping(vRowDT.qty_git, rowDGV.Cells(Me.ColWIP_DefPrecUnit.Index).Value)
                        Me.TxtMP_QtyMOB.Text = Me.ObjCD.DigitGrouping(vRowDT.qty_mob, rowDGV.Cells(Me.ColWIP_DefPrecUnit.Index).Value)
                        Me.TxtMP_QtyOnHand.Text = Me.ObjCD.DigitGrouping(vRowDT.qty_on_hand, rowDGV.Cells(Me.ColWIP_DefPrecUnit.Index).Value)
                        Me.TxtMP_QtyPendingDO.Text = Me.ObjCD.DigitGrouping(vRowDT.qty_pending_do, rowDGV.Cells(Me.ColWIP_DefPrecUnit.Index).Value)
                        Me.TxtMP_QtyPOB.Text = Me.ObjCD.DigitGrouping(vRowDT.qty_pob, rowDGV.Cells(Me.ColWIP_DefPrecUnit.Index).Value)
                        Me.TxtMP_QtyPrOB.Text = Me.ObjCD.DigitGrouping(vRowDT.qty_prob, rowDGV.Cells(Me.ColWIP_DefPrecUnit.Index).Value)
                        Me.TxtMP_QtyRMReq.Text = Me.ObjCD.DigitGrouping(vRowDT.qty_rm_req, rowDGV.Cells(Me.ColWIP_DefPrecUnit.Index).Value)
                        Me.TxtMP_QtySafetyStock.Text = Me.ObjCD.DigitGrouping(vRowDT.qty_safety_stock, rowDGV.Cells(Me.ColWIP_DefPrecUnit.Index).Value)
                        Me.TxtMP_QtySOB.Text = Me.ObjCD.DigitGrouping(vRowDT.qty_sob, rowDGV.Cells(Me.ColWIP_DefPrecUnit.Index).Value)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Kosongkan_TxtQty()
        Try
            Me.TxtMP_QtyAvailable.Text = Me.ObjCD.DigitGrouping(0, 2)
            Me.TxtMP_QtyGIT.Text = Me.ObjCD.DigitGrouping(0, 2)
            Me.TxtMP_QtyMOB.Text = Me.ObjCD.DigitGrouping(0, 2)
            Me.TxtMP_QtyOnHand.Text = Me.ObjCD.DigitGrouping(0, 2)
            Me.TxtMP_QtyPendingDO.Text = Me.ObjCD.DigitGrouping(0, 2)
            Me.TxtMP_QtyPOB.Text = Me.ObjCD.DigitGrouping(0, 2)
            Me.TxtMP_QtyPrOB.Text = Me.ObjCD.DigitGrouping(0, 2)
            Me.TxtMP_QtyRMReq.Text = Me.ObjCD.DigitGrouping(0, 2)
            Me.TxtMP_QtySafetyStock.Text = Me.ObjCD.DigitGrouping(0, 2)
            Me.TxtMP_QtySOB.Text = Me.ObjCD.DigitGrouping(0, 2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkQtyAvailableOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQtyAvailableOnly.CheckedChanged
        Try
            Me.SetRowVisible()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim rowDGV As DataGridViewRow

        Try
            'pastikan tidak ada qty wip yang lebih besar dr qty required
            For Each rowDGV In Me.DGVWIP.Rows
                If CDec(rowDGV.Cells(Me.ColWIP_QtyReq.Index).Value) < CDec(rowDGV.Cells(Me.ColWIP_QtyWIP.Index).Value) Then
                    MsgBox("Qty WIP > Qty Required in WO: " + rowDGV.Cells(Me.ColWIP_WONo.Index).Value.ToString + ". Please fix.", MsgBoxStyle.Information, "Information")
                    Me.DGVWIP.ClearSelection()
                    rowDGV.Selected = True
                    Exit Sub
                End If
            Next

            If MsgBox("Accept data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.LangsungClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyAvailable_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtMP_QtyAvailable.TextChanged

    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim vRowItemRM, vRowItemRM_Original As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim vRowItemWO, vRowItemWO_Original As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow

        Try
            Me.Timer1.Enabled = False

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.DefaultBackColor1 = Me.DGVWIP.DefaultCellStyle.BackColor
            Me.MinimizeBox = False
            Me.MaximizeBox = True

            'pindahkan ke tabel antara
            vTblItemRM.Rows.Clear()
            vTblItemWO.Rows.Clear()
            For Each vRowItemRM_Original In Me.vTblItemRM_Original
                If vRowItemRM_Original.Identifier_JO = Me.vIdentifierJO Then
                    vRowItemRM = Me.vTblItemRM.Newtbltrans_mo_item_rmRow
                    For mn As Integer = 0 To Me.vTblItemRM.Columns.Count - 1
                        vRowItemRM.Item(mn) = vRowItemRM_Original.Item(mn)
                    Next
                    Me.vTblItemRM.Rows.Add(vRowItemRM)
                End If
            Next
            For Each vRowItemWO_Original In Me.vTblItemWO_Original.Rows
                If vRowItemWO_Original.Identifier_JO = Me.vIdentifierJO Then
                    vRowItemWO = Me.vTblItemWO.Newtbltrans_mo_item_woRow
                    For mn As Integer = 0 To Me.vTblItemWO.Columns.Count - 1
                        vRowItemWO.Item(mn) = vRowItemWO_Original.Item(mn)
                    Next
                    Me.vTblItemWO.Rows.Add(vRowItemWO)
                End If
            Next
            vTblItemRM.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "' AND [is wip] = 1"
            vTblItemRM.DefaultView.Sort = "[Primary Product Code RM] ASC"
            Me.ListPrimaryProduct.Clear()
            If vTblItemRM.DefaultView.Count > 0 Then
                For Each DV As Data.DataRowView In Me.vTblItemRM.DefaultView
                    vRowItemRM = DV.Row
                    ListPrimaryProduct.Add(vRowItemRM.Primary_Product_Code_RM)
                Next
            End If

            'Me.DGVWIP.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.LoadDGV()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_AddWIP_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Try
            Me.Timer1.Interval = 1000
            Me.Timer1.Enabled = True
            Control.CheckForIllegalCrossThreadCalls = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow

        Try
            Me.ObjCD.UpdateThreading()

            vTblItemRM.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "' AND [is wip] = 1"
            vTblItemRM.DefaultView.Sort = "[Primary Product Code RM] ASC"

            ObjMP.HitungQtyAvailable(Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, Me.vDT_ItemPOB, Me.vDT_ItemSOB, Me.vDT_ItemPrOB, Me.vDT_ItemSettingLocation, _
            Me.vDT_ItemQty, Me.vDT_ItemMOB, Me.vDT_ItemRMReq, Me.ObjSetting, Me.BackColorAll, Me.vUserCode, Me.ObjCD, _
            Me.ObjSettingUser, Me.ListPrimaryProduct, Me.vTblItemWO, Me.vIdentifierJO, Me.vJONO, Me.vMONo, Me.vMODate, Me.vPrimaryMO)
            If vTblItemRM.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblItemRM.DefaultView
                    rowItemRM = dv.Row
                    rowItemRM.Available_Qty = Me.GetQtyAvailable(rowItemRM.Primary_Product_Code_RM)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim rowD As DataGridViewRow
        Dim ObjProses As New Core_Module.ClsMasterProses
        Dim vDecProd_Item As Integer = 0

        Try
            Me.Panel1.Visible = False

            vTblItemRM.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "' AND [is wip] = 1"
            vTblItemRM.DefaultView.Sort = "[Primary Product Code RM] ASC"
            If vTblItemRM.DefaultView.Count > 0 Then
                DGVWIP.AllowUserToAddRows = False
                For Each dv As Data.DataRowView In vTblItemRM.DefaultView
                    rowItemRM = dv.Row
                    With rowItemRM
                        DGVWIP.Rows.Add()
                        rowD = DGVWIP.Rows(DGVWIP.RowCount - 1)
                        rowD.Cells(Me.ColWIP_No.Index).Value = DGVWIP.Rows.Count.ToString
                        vDecProd_Item = ObjProduct.GetAngkaDecimal(.Primary_Product_Code_RM)
                        rowD.Cells(Me.ColWIP_PrimaryCodeNo.Index).Value = .Primary_Product_Code_RM
                        vTblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" & .Identifier_JO.ToString & "' AND [Identifier] = '" & .Identifier_WO & "'"
                        If vTblItemWO.DefaultView.Count > 0 Then
                            rowItemWO = vTblItemWO.DefaultView(0).Row
                            rowD.Cells(Me.ColWIP_PrimaryProses.Index).Value = rowItemWO.Primary_Process
                            rowD.Cells(Me.ColWIP_Proses.Index).Value = ObjProses.GetCode(rowItemWO.Primary_Process)
                        End If
                        rowD.Cells(Me.ColWIP_PrimaryUnit.Index).Value = ObjProduct.GetPrimaryUnit(.Primary_Product_Code_RM)
                        rowD.Cells(Me.ColWIP_WONo.Index).Value = ObjMO.GetWONumber_InTbl(.Identifier_WO, vTblItemWO, .Identifier_JO)
                        rowD.Cells(Me.ColWIP_CodeNo.Index).Value = ObjProduct.GetProductCode(.Primary_Product_Code_RM)
                        rowD.Cells(Me.ColWIP_Desc.Index).Value = ObjProduct.GetDescription(.Primary_Product_Code_RM)
                        rowD.Cells(Me.ColWIP_UnitCode.Index).Value = ObjProduct.GetUnitCode(.Primary_Product_Code_RM)
                        rowD.Cells(Me.ColWIP_QtyProduce.Index).Value = Me.ObjCD.DigitGrouping(.Required_Qty + .Additional_Qty - .WIP_Qty, vDecProd_Item)
                        rowD.Cells(Me.ColWIP_QtyReq.Index).Value = Me.ObjCD.DigitGrouping(.Required_Qty + .Additional_Qty, vDecProd_Item)
                        rowD.Cells(Me.ColWIP_QtyAvailable.Index).Value = Me.ObjCD.DigitGrouping(.Available_Qty, vDecProd_Item)
                        rowD.Cells(Me.ColWIP_QtyWIP.Index).Value = Me.ObjCD.DigitGrouping(.WIP_Qty, vDecProd_Item)
                        rowD.Cells(Me.ColWIP_IdentifierWO.Index).Value = .Identifier_WO
                        rowD.Cells(Me.ColWIP_IdentifierJO.Index).Value = vIdentifierJO

                        rowD.Cells(Me.ColWIP_Remarks.Index).Value = ""
                        rowD.Cells(Me.ColWIP_DefPrecUnit.Index).Value = vDecProd_Item

SkipRow:
                    End With
                Next
            End If

            DGVWIP.Sort(DGVWIP.Columns(Me.ColWIP_WONo.Index), System.ComponentModel.ListSortDirection.Descending)

            Me.SetRowVisible()

            DGVWIP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'DGVWIP.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Me.DGVWIP.ClearSelection()
            For ij As Integer = 0 To Me.DGVWIP.Rows.Count - 1
                If Me.DGVWIP.Rows(ij).Visible = True Then
                    Me.DGVWIP.Item(0, ij).Selected = True
                End If
            Next

            Me.ControlBox = True
            Me.btnCancel.Enabled = True
            Me.btnOK.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjProses = Nothing
        End Try
    End Sub
End Class