Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_WO
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vIdentifierWO As String = ""
    Public vMONo As String = ""
    Public vJONo As String = ""
    Public vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
    Public vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
    Public vBolehEdit As Boolean = False
    Public vIdentifierJO As String = ""

    Dim objProd As New Core_Module.clsMasterProduct
    Dim objProses As New Core_Module.ClsMasterProses
    Dim objWU As New Core_Module.ClsMasterWorkUnit
    Dim objLoc As New Core_Module.clsMasterStorageArea
    Dim objCD As New Core_Module.ClsCekData
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
    Dim ObjMO As New ClsTrans_ManufacturingOrder
    Dim ObjMO_Item As New ClsTrans_ManufacturingOrder_Item

    Dim vDecProd As Double = -1
    Dim txtSelect As TextBox
    Dim Buffer As String = ""
    Dim Buff_Date As Date
    Dim Buff_DocNo As String = ""
    Dim FromOtherButton As Boolean = False
    Dim CancelClose As Boolean = False

    Private Sub FrmTrans_ManufacturingOrder_WO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Hasil As MsgBoxResult

        Try
            If vBolehEdit = True Then
                If Me.FromOtherButton = False Then
                    Hasil = MsgBox("Accept data changes before closing form?", MsgBoxStyle.YesNoCancel, "Confirm")
                    If Hasil = MsgBoxResult.Yes Then
                        If Me.ApakahDataBenar = True Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            Dim vFrm As New FrmTrans_ManufacturingOrder
                            Me.UpdateItemWORM()
                            vFrm.Dispose()
                            Me.DialogResult = Windows.Forms.DialogResult.OK
                            Me.FromOtherButton = True
                            Me.Close()
                        End If
                    ElseIf Hasil = MsgBoxResult.No Then
                        Me.FromOtherButton = True
                        Me.Close()
                    ElseIf Hasil = MsgBoxResult.Cancel Then
                        e.Cancel = True
                    End If
                Else
                    If Me.CancelClose = True Then
                        Me.CancelClose = False
                        e.Cancel = True
                    End If
                    Me.FromOtherButton = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_WO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.DataGridView3.Rows.Add("Qty", 0, 0, "Unit")
        'Me.DataGridView3.Rows.Add("Good Percentage", 0.0, 0.0, "%")

        Try
            If Me.vIdentifierJO = "" Then
                MsgBox("Logic error: no identifier JO specified. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            ColRM_StandByLocation.Visible = False
            ColRM_PickUpLocation.Visible = False
            ColRM_GPTarget.Visible = False
            ColRM_GPActual.Visible = False
            ColRM_QtyAllocated.Visible = False

            Me.DgvItemRM.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvItemBundling.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If vBolehEdit = False Then
                Me.NonAktifkanDetail()
                Me.DgvItemRM.ReadOnly = True
            End If

            If ObjMO.GetJmlBundling(vIdentifierWO, vIdentifierJO) > 0 Then
                Me.DtpDate.Enabled = False
            End If

            Me.TxtPlanProdRate.ReadOnly = True
            Me.SalinKeField()

            Buff_DocNo = Me.TxtDocNo.Text
            Buff_Date = Me.DtpDate.Value

            Me.Timer1.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahDataBenar() As Boolean
        Try
            'product harus memiliki item bom, jadi minimal harus ada 1 baris item di DgvItemBOM
            'If Me.TxtPrimaryWorkUnit.Text = "-1" Then
            '    MsgBox("Work Unit tidak boleh kosong", MsgBoxStyle.Information, "Information")
            '    Me.TxtWorkUnitCode.Focus()
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub NonAktifkanDetail()
        Try
            Me.objCD.KendalikanControl(Me.GboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.objCD.KendalikanControl(Me.GroupBox2.Controls, True, False, True, False, True, False, True, False)
            Me.objCD.KendalikanControl(Me.GroupBox3.Controls, True, False, True, False, True, False, True, False)
            Me.objCD.KendalikanControl(Me.GroupBox4.Controls, True, False, True, False, True, False, True, False)
            Me.objCD.KendalikanControl(Me.GroupBox5.Controls, True, False, True, False, True, False, True, False)
            Me.objCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.objCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.objCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.BtnFindWorkUnit.Enabled = False
            Me.DgvItemRM.TabStop = False
            Me.DgvItemBundling.TabStop = True
            Me.SplitContainer1.TabStop = True
            Me.objCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.objCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)
            Me.BtnOK.Visible = False
            Me.BtnCancel.Text = "Close"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function GetWorkUnitCode(ByVal vIdentifierWO As String) As String
        Dim dta As New DseTrans_ManufacturingOrderTableAdapters.DataTable1TableAdapter
        Dim tbl As New DseTrans_ManufacturingOrder.DataTable1DataTable
        Dim row As DseTrans_ManufacturingOrder.DataTable1Row
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Fill(tbl, vIdentifierWO) > 0 Then
                If tbl.Rows.Count = 1 Then
                    row = tbl.Rows(0)
                    Return row.work_unit_code
                Else
                    Return "Multiple"
                End If
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub SalinKeField()
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim rowD_RM As DataGridViewRow
        Dim vDecProd_Item As Integer = 0

        Try
            Me.vTblItemWO.DefaultView.RowFilter = "[Identifier] = '" & vIdentifierWO & "'"
            If Me.vTblItemWO.DefaultView.Count > 0 Then
                rowItemWO = Me.vTblItemWO.DefaultView(0).Row
                With rowItemWO
                    Me.DtpDate.Value = ._Date
                    Me.TxtDocNo.Text = .WO_Number
                    Me.TxtMONo.Text = vMONo
                    Me.TxtJONo.Text = vJONo
                    Me.TxtProcessCode.Text = objProses.GetCode(.Primary_Process)
                    Me.TxtProcessDesc.Text = objProses.GetDesc(.Primary_Process)

                    vDecProd = objProd.GetAngkaDecimal(.Primary_Product)

                    'Qty Processed (Units)
                    Me.TxtTargetResult.Text = Me.objCD.DigitGrouping(.Qty_Target_Result, vDecProd)
                    Me.TxtTargetProduce.Text = Me.objCD.DigitGrouping(.Qty_Target_Prod, vDecProd)
                    Me.TxtTargetProduceAdd.Text = Me.objCD.DigitGrouping(.Qty_Target_Prod_Add, vDecProd)
                    Me.TxtActualResult.Text = Me.objCD.DigitGrouping(.Qty_Actual_Result, vDecProd)
                    Me.TxtActualProduce.Text = Me.objCD.DigitGrouping(.Qty_Actual_Prod, vDecProd)
                    Me.TxtProjectedResult.Text = Me.objCD.DigitGrouping(.Qty_Projected_Result, vDecProd)
                    If .Qty_Target_Prod = 0 Then
                        Me.TxtProgress.Text = Me.objCD.DigitGrouping(0, 2)
                    Else
                        Me.TxtProgress.Text = Me.objCD.DigitGrouping((.Qty_Actual_Prod / .Qty_Target_Prod) * 100, 2)
                    End If

                    'Good Percentage
                    Me.TxtTargetPerc.Text = Me.objCD.DigitGrouping(.Plan_Good_Perct, 2)
                    Me.TxtActualPerc.Text = Me.objCD.DigitGrouping(.Actual_Good_Perct, 2)

                    'Work Unit
                    Me.TxtPrimaryWorkUnit.Text = .Primary_Work_Unit
                    Me.TxtWorkUnitCode.Text = objWU.GetCode(.Primary_Work_Unit)
                    Me.TxtWorkUnitDesc.Text = objWU.GetDesc(.Primary_Work_Unit)
                    Me.txtScheduledWorkUnit.Text = Me.GetWorkUnitCode(.Identifier)

                    'Plan
                    Me.DtpPlanStart.Value = .Plan_Start_Time
                    Me.DtpPlanEnd.Value = .Plan_End_Time
                    Me.TxtPlanWorkHour.Text = Me.objCD.Convert_ToString(.Plan_Work_Hour)
                    Me.TxtPlanSetupTime.Text = Me.objCD.Convert_ToString(.Plan_Setup_Time)
                    Me.TxtPlanTotal.Text = Me.objCD.Convert_ToString(.Plan_Total_Time)
                    Me.TxtPlanProdRate.Text = Me.objCD.DigitGrouping(.Plan_Prod_Rate, vDecProd)

                    'Actual
                    Me.DtpActualStart.Value = .Actual_Start_Time
                    Me.DtpActualEnd.Value = .Actual_End_Time
                    Me.TxtActualWorkHour.Text = Me.objCD.Convert_ToString(.Actual_Work_Hour)
                    Me.TxtActualSetupTime.Text = Me.objCD.Convert_ToString(.Actual_Setup_Time)
                    Me.TxtActualTotal.Text = Me.objCD.Convert_ToString(.Actual_Total_Time)
                    Me.TxtActualDownTime.Text = Me.objCD.Convert_ToString(.Actual_Down_Time)
                    Me.TxtActualProdRate.Text = Me.objCD.DigitGrouping(.Actual_Prod_Rate, vDecProd, Core_Module.ClsCekData.EnumTipeRounding.Down)

                    Me.tslCreatedBy.Text = "Created By: " + .Created_By
                    Me.tslCreatedDate.Text = "Created Date: " + Format(.Created_Date, "dd-MMM-yyyy") + "-" + .Created_Date.ToShortTimeString
                    Me.tslModifiedBy.Text = "Modified By: " + .Modified_By
                    Me.tslModifiedDate.Text = "Modified Date: " + Format(.Modified_Date, "dd-MMM-yyyy") + "-" + .Modified_Date.ToShortTimeString
                End With
            End If

            'isi dgv RM
            Me.DgvItemRM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvItemRM.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.vTblItemRM.DefaultView.RowFilter = "[Identifier WO] = '" & vIdentifierWO & "'"
            If Me.vTblItemRM.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In Me.vTblItemRM.DefaultView
                    rowItemRM = dv.Row
                    Me.DgvItemRM.Rows.Add()
                    rowD_RM = Me.DgvItemRM.Rows(Me.DgvItemRM.RowCount - 1)
                    With rowItemRM
                        vDecProd_Item = objProd.GetAngkaDecimal(.Primary_Product_Code_RM)
                        rowD_RM.Cells(Me.ColRM_No.Index).Value = Me.DgvItemRM.Rows.Count.ToString
                        rowD_RM.Cells(Me.ColRM_CodeNo.Index).Value = objProd.GetProductCode(.Primary_Product_Code_RM)
                        rowD_RM.Cells(Me.ColRM_Desc.Index).Value = objProd.GetDescription(.Primary_Product_Code_RM)
                        rowD_RM.Cells(Me.ColRM_GPActual.Index).Value = Me.objCD.DigitGrouping(.GP_Actual, 2)
                        rowD_RM.Cells(Me.ColRM_GPTarget.Index).Value = Me.objCD.DigitGrouping(.GP_Target, 2)
                        rowD_RM.Cells(Me.ColRM_Target.Index).Value = Me.objCD.DigitGrouping(.Target_Qty, vDecProd_Item)
                        rowD_RM.Cells(Me.ColRM_Unit.Index).Value = objProd.GetUnitCode(.Primary_Product_Code_RM)
                        Me.vTblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" & .Identifier_JO.ToString & "' AND [Identifier] = '" & .Identifier_WO & "'"
                        If Me.vTblItemWO.DefaultView.Count > 0 Then
                            rowItemWO = Me.vTblItemWO.DefaultView(0).Row
                            rowD_RM.Cells(Me.ColRM_WONumber.Index).Value = rowItemWO.WO_Number
                            rowD_RM.Cells(Me.ColRM_PrimaryProcess.Index).Value = rowItemWO.Primary_Process
                            rowD_RM.Cells(Me.ColRM_Process.Index).Value = objProses.GetCode(rowItemWO.Primary_Process)
                        End If
                        rowD_RM.Cells(Me.ColRM_QtyAdditional.Index).Value = Me.objCD.DigitGrouping(.Additional_Qty, vDecProd_Item)
                        rowD_RM.Cells(Me.ColRM_QtyAllocated.Index).Value = Me.objCD.DigitGrouping(.Allocated_Qty, vDecProd_Item)
                        rowD_RM.Cells(Me.ColRM_QtyBalance.Index).Value = Me.objCD.DigitGrouping(.Balance_Qty, vDecProd_Item)
                        rowD_RM.Cells(Me.ColRM_QtyWIP.Index).Value = Me.objCD.DigitGrouping(.WIP_Qty, vDecProd_Item)
                        rowD_RM.Cells(Me.ColRM_QtyRequired.Index).Value = Me.objCD.DigitGrouping(.Required_Qty, vDecProd_Item)
                        rowD_RM.Cells(Me.ColRM_Target.Index).Value = Me.objCD.DigitGrouping(.Target_Qty, vDecProd_Item)
                        rowD_RM.Cells(Me.ColRM_QtyUsed.Index).Value = Me.objCD.DigitGrouping(.Used_Qty, vDecProd_Item)
                        rowD_RM.Cells(Me.ColRM_Remarks.Index).Value = .Remarks
                        rowD_RM.Cells(Me.ColRM_StandByLocation.Index).Value = objLoc.GetStorageAreaCode(.Primary_StandBy_Location)
                        rowD_RM.Cells(Me.ColRM_PickUpLocation.Index).Value = objLoc.GetStorageAreaCode(.Primary_PickUp_Location)
                        rowD_RM.Cells(Me.ColRM_PrimaryStandByLocation.Index).Value = .Primary_StandBy_Location
                        rowD_RM.Cells(Me.ColRM_PrimaryPickUpLocation.Index).Value = .Primary_PickUp_Location
                        rowD_RM.Cells(Me.ColRM_PrimaryCodeNo.Index).Value = .Primary_Product_Code_RM
                        rowD_RM.Cells(Me.ColRM_PrimaryUnit.Index).Value = objProd.GetPrimaryUnit(.Primary_Product_Code_RM)
                        rowD_RM.Cells(Me.ColRM_IdentifierItem.Index).Value = "-1-1" 'blabla
                        rowD_RM.Cells(Me.ColRM_IdentifierWO.Index).Value = .Identifier_WO
                    End With
                Next
            End If
            Me.DgvItemRM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'isi dgv bundling
            ObjMO_Item.IsiDgvItemBundling(vIdentifierWO, Me.DgvItemBundling, Me, vIdentifierJO)

            Me.dgvWorkUnitDetail.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Dim r0 As DataGridViewRow
            Dim d As New DseTrans_ManufacturingOrderTableAdapters.DataTable2TableAdapter
            Dim t As New DseTrans_ManufacturingOrder.DataTable2DataTable
            Dim r As DseTrans_ManufacturingOrder.DataTable2Row
            d.Connection.ConnectionString = Core_Module.GetConstring
            d.Fill(t, vIdentifierWO)
            Me.dgvWorkUnitDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvWorkUnitDetail.Rows.Clear()
            For Each r In t
                r0 = Me.dgvWorkUnitDetail.Rows(Me.dgvWorkUnitDetail.Rows.Add)
                With r0
                    If r.plan_end_time > #1/1/1950# Then
                        .Cells(Me.ColWU_Finish.Index).Value = Format(r.plan_end_time, "dd MMM yyyy HH.mm")
                    End If
                    .Cells(Me.ColWU_No.Index).Value = Me.dgvWorkUnitDetail.Rows.Count
                    .Cells(Me.ColWU_Progress.Index).Value = Me.objCD.DigitGrouping(r.progress, 2)
                    .Cells(Me.ColWU_ActualProduceQty.Index).Value = Me.objCD.DigitGrouping(r.qty_produce, r._Decimal)
                    .Cells(Me.ColWU_Rate.Index).Value = Me.objCD.DigitGrouping(r.plan_prod_rate, r._Decimal)
                    If r.plan_start_time > #1/1/1950# Then
                        .Cells(Me.ColWU_Start.Index).Value = Format(r.plan_start_time, "dd MMM yyyy HH.mm")
                    End If
                    .Cells(Me.ColWU_Unit.Index).Value = r.Unit_Code
                    .Cells(Me.ColWU_WH.Index).Value = Me.objCD.Convert_ToString(r.plan_work_duration)
                    .Cells(Me.ColWU_WorkUnit.Index).Value = r.work_unit_code
                End With
            Next
            Me.dgvWorkUnitDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Dim hasil As MsgBoxResult

        Try
            Me.FromOtherButton = False
            If Me.ApakahDataBenar() = True Then
                hasil = MsgBox("Close form and accept data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Dim vFrm As New FrmTrans_ManufacturingOrder
                    Me.UpdateItemWORM()
                    vFrm.Dispose()
                    Me.FromOtherButton = True
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Dim hasil As MsgBoxResult

        Try
            Me.FromOtherButton = False
            If Me.vBolehEdit = True Then
                hasil = MsgBox("Close form and discard data changes?", MsgBoxStyle.YesNo, "Confirmation")
                If hasil = MsgBoxResult.Yes Then
                    Me.FromOtherButton = True
                    Me.Close()
                Else
                    Me.FromOtherButton = True
                    Me.CancelClose = True
                End If
            Else
                Me.Close()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateItemWORM()
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim rowD As DataGridViewRow

        Try
            vTblItemWO.DefaultView.RowFilter = "[Identifier] = '" & vIdentifierWO & "'"
            If vTblItemWO.DefaultView.Count > 0 Then
                rowItemWO = vTblItemWO.DefaultView(0).Row
                With rowItemWO
                    .WO_Number = Me.TxtDocNo.Text
                    ._Date = Me.DtpDate.Value
                    .Primary_Work_Unit = Me.TxtPrimaryWorkUnit.Text
                    .Plan_Start_Time = Me.DtpPlanStart.Value
                    .Plan_End_Time = Me.DtpPlanEnd.Value
                    .Plan_Setup_Time = Me.objCD.Convert_ToMinutes(Me.TxtPlanSetupTime.Text)
                    .Plan_Work_Hour = Me.objCD.Convert_ToMinutes(Me.TxtPlanWorkHour.Text)
                    .Plan_Delay_Time = 0
                    .Plan_Total_Time = Me.objCD.Convert_ToMinutes(Me.TxtPlanTotal.Text)
                    .Plan_Prod_Rate = Me.TxtPlanProdRate.Text
                    .Actual_Start_Time = Me.DtpActualStart.Value
                    .Actual_End_Time = Me.DtpActualEnd.Value
                    .Actual_Setup_Time = Me.objCD.Convert_ToMinutes(Me.TxtActualSetupTime.Text)
                    .Actual_Work_Hour = Me.objCD.Convert_ToMinutes(Me.TxtActualWorkHour.Text)
                    .Actual_Total_Time = Me.objCD.Convert_ToMinutes(Me.TxtActualTotal.Text)
                    .Actual_Down_Time = Me.objCD.Convert_ToMinutes(Me.TxtActualDownTime.Text)
                    .Actual_Prod_Rate = Me.TxtActualProdRate.Text
                    .Actual_Delay_Time = 0
                    .Plan_Good_Perct = Me.TxtTargetPerc.Text
                    .Actual_Good_Perct = Me.TxtActualPerc.Text
                    .Qty_Target_Result = Me.TxtTargetResult.Text
                    .Qty_Target_Prod = Me.TxtTargetProduce.Text
                    .Qty_Target_Prod_Add = Me.TxtTargetProduceAdd.Text
                    .Qty_Actual_Prod = Me.TxtActualProduce.Text
                    .Qty_Actual_Result = Me.TxtActualResult.Text
                    .Qty_Projected_Result = Me.TxtProjectedResult.Text
                    .Modified_By = UserCode
                    .Modified_Date = Now
                    .Timestamp = Now
                    '.Balance_Produce = Me.txt
                    '.Balance_Result = Me.objCD.RoundUp1(.Qty_Target_Result - .Qty_Actual_Result, rowItemBOM_.Decimal_RM)
                End With

                For i As Integer = 0 To Me.DgvItemRM.Rows.Count - 1
                    rowD = Me.DgvItemRM.Rows(i)
                    vTblItemRM.DefaultView.RowFilter = "[Primary Product Code RM] = " & rowD.Cells(Me.ColRM_PrimaryCodeNo.Index).Value.ToString & _
                    " AND [Identifier WO] = '" & rowD.Cells(Me.ColRM_IdentifierWO.Index).Value & "'"
                    If vTblItemRM.DefaultView.Count > 0 Then
                        rowItemRM = vTblItemRM.DefaultView(0).Row
                        With rowItemRM
                            '.Allocated_Qty = rowD.Cells(Me.ColRM_QtyAllocated.Index).Value
                            .Additional_Qty = CDec(rowD.Cells(Me.ColRM_QtyAdditional.Index).Value)
                            .Balance_Qty = CDec(rowD.Cells(Me.ColRM_QtyAllocated.Index).Value) + CDec(rowD.Cells(Me.ColRM_QtyAdditional.Index).Value) - _
                                           CDec(rowD.Cells(Me.ColRM_QtyUsed.Index).Value)
                            .Remarks = rowD.Cells(Me.ColRM_Remarks.Index).Value
                            .Primary_StandBy_Location = CDbl(rowD.Cells(Me.ColRM_PrimaryStandByLocation.Index).Value)
                            .Primary_PickUp_Location = CDbl(rowD.Cells(Me.ColRM_PrimaryPickUpLocation.Index).Value)
                        End With
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemRM_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItemRM.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColRM_QtyAdditional.Index Or e.ColumnIndex = Me.ColRM_Remarks.Index Then
                Me.Buffer = Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemRM_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItemRM.CellEndEdit
        Dim vDecProd_Item As Decimal = 0

        Try
            vDecProd_Item = objProd.GetAngkaDecimal(Me.DgvItemRM.Item(Me.ColRM_PrimaryCodeNo.Index, e.RowIndex).Value)
            Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "").Replace("#", "").Replace("*", "").Replace("`", "").Replace("$", "").Replace("%", "")
            If e.ColumnIndex = Me.ColRM_QtyAllocated.Index Then
                'If IsNumeric(Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                '    If IsNumeric(Me.Buffer) = True Then
                '        Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.Buffer, 0)
                '    Else
                '        Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value = 0
                '    End If
                'Else
                '    If Me.objCD.CekBilanganPositif(Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                '        MsgBox("Please enter positive number for qty produce.", MsgBoxStyle.Information, _
                '        "Information")
                '        Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.Buffer, 0)
                '        Exit Sub
                '    Else
                '        Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value, 0)
                '    End If
                'End If
            ElseIf e.ColumnIndex = Me.ColRM_QtyAdditional.Index Then
                If IsNumeric(Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If IsNumeric(Me.Buffer) = True Then
                        Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.Buffer, vDecProd_Item)
                    Else
                        Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value = 0
                    End If
                Else
                    '    If Me.ObjCD.CekBilanganPositif(Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    '        MsgBox("Please enter positive number for qty produce.", MsgBoxStyle.Information, _
                    '        "Information")
                    '        Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.Buffer, 0)
                    '        Exit Sub
                    '    Else
                    Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.DgvItemRM.Item(e.ColumnIndex, e.RowIndex).Value, vDecProd_Item)
                    '    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemRM_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItemRM.Sorted
        Try
            Me.objCD.SetNumber(Me.DgvItemRM)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemBundling_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItemBundling.CellBeginEdit
        Try
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemBundling_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItemBundling.CellContentClick
        Try
            If e.RowIndex <> -1 Then
                If Me.DgvItemBundling.Item(Me.dgvBundling_PrimaryBundling.Index, e.RowIndex).Value <> -1 Then
                    If e.ColumnIndex = Me.dgvBundling_DocNo.Index Then
                        Dim i As New FrmTrans_StockBundling
                        i.PrimaryBundling = CDbl(Me.DgvItemBundling.Item(Me.dgvBundling_PrimaryBundling.Index, e.RowIndex).Value)
                        i.ShowDialog()
                        i.dispose()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemBundling_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItemBundling.Sorted
        Try
            Me.objCD.SetNumber(Me.DgvItemBundling)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPlanWorkHour_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPlanWorkHour.Enter
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.TxtTargetProduceAdd.Enabled = True Then
                    If Txt.Enabled = True Then
                        Me.Buffer = Txt.Text
                        Me.txtSelect = Txt
                        Me.BackgroundWorker1.RunWorkerAsync()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPlanWorkHour_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPlanWorkHour.KeyDown
        Try
            ' Me.TxtPlanProdRate.Text = CDec(Me.TxtTargetProduce.Text) / Me.objCD.Convert_ToMinutes(Me.TxtPlanWorkHour.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPlanWorkHour_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPlanWorkHour.Leave
        Dim Txt As TextBox

        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.BtnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If String.IsNullOrEmpty(Txt.Text) = False Then
                            If Me.objCD.Is_Time_String(Txt.Text) = True Then
                                If Me.objCD.Convert_ToMinutes(Txt.Text) > 1439 Then
                                    MsgBox("Schedule in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                                    Txt.Text = Me.Buffer
                                End If
                            Else
                                MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                                Txt.Text = Me.Buffer
                            End If
                        Else
                            Txt.Text = "00.00"
                        End If
                    End If
                End If
            End If

            If Me.objCD.Convert_ToMinutes(Me.TxtPlanWorkHour.Text) > 0 Then
                Me.TxtPlanProdRate.Text = Me.objCD.RoundDown1(CDec(Me.TxtTargetProduce.Text) / (Me.objCD.Convert_ToMinutes(Me.TxtPlanWorkHour.Text) / 60), vDecProd)
            Else
                Me.TxtPlanProdRate.Text = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.txtSelect.SelectAll()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTargetProduceAdd_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTargetProduceAdd.Enter
        Try
            Me.objCD.Number_Enter(Me.TxtTargetProduceAdd, sender, e)
            Me.TxtProjectedResult.Text = Me.objCD.RoundDown1((CDec(Me.TxtTargetProduce.Text) + CDec(Me.TxtTargetProduceAdd.Text) - CDec(Me.TxtActualProduce.Text)) * (CDec(Me.TxtTargetPerc.Text) / 100) + CDec(Me.TxtActualResult.Text), vDecProd)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTargetProduceAdd_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTargetProduceAdd.KeyDown
        Try
            Me.objCD.Number_KeyDown(Me.TxtTargetProduceAdd, sender, e, 0)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTargetProduceAdd_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTargetProduceAdd.KeyUp
        Try
            Me.objCD.Number_KeyUp(Me.TxtTargetProduceAdd, sender, e, False)
            Me.TxtProjectedResult.Text = Me.objCD.RoundDown1((CDec(Me.TxtTargetProduce.Text) + CDec(Me.TxtTargetProduceAdd.Text) - CDec(Me.TxtActualProduce.Text)) * (CDec(Me.TxtTargetPerc.Text) / 100) + CDec(Me.TxtActualResult.Text), vDecProd)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTargetProduceAdd_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTargetProduceAdd.Leave
        Try
            Me.objCD.Number_Leave(Me.TxtTargetProduceAdd, sender, e, vDecProd, Me.objCD)
            Me.TxtProjectedResult.Text = Me.objCD.RoundDown1((CDec(Me.TxtTargetProduce.Text) + CDec(Me.TxtTargetProduceAdd.Text) - CDec(Me.TxtActualProduce.Text)) * (CDec(Me.TxtTargetPerc.Text) / 100) + CDec(Me.TxtActualResult.Text), vDecProd)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.ValueChanged
        Dim urutan As Integer = 0

        Try
            If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.WO) = False Then
                Exit Sub
            End If

            Dim f As New FrmTrans_ManufacturingOrder
            If Me.DtpDate.Value.Date = Me.Buff_Date.Date Then
                Me.TxtDocNo.Text = Me.Buff_DocNo
            Else
                Me.TxtDocNo.Text = Me.ObjMO.GetNoBaru1(Me.DtpDate.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.WO, _
                f.getMaxNoUrutWO(Me.vTblItemWO), Nothing, -1, Me.vTblItemWO)
            End If
            f.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Public Function getMaxNoUrutWO() As Integer
    '    Dim urutan As Integer = 0
    '    Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow

    '    Try
    '        Me.vTblItemWO.DefaultView.RowFilter = Nothing
    '        Me.vTblItemWO.DefaultView.Sort = "[Urutan] Desc"
    '        If Me.vTblItemWO.DefaultView.Count > 0 Then
    '            rowItemWO = Me.vTblItemWO.DefaultView(0).Row
    '            If IsNumeric(rowItemWO.Urutan) = True Then
    '                urutan = rowItemWO.Urutan
    '            End If
    '        End If

    '        'If vFilter <> "" Then
    '        '    Me.vTblItemWO.DefaultView.RowFilter = vFilter
    '        'End If
    '        Return urutan
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Private Sub TxtWorkUnitCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtWorkUnitCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtWorkUnitCode.Text <> "" Then
                    If Me.objWU.ApakahSudahAda(Me.TxtWorkUnitCode.Text) = True Then
                        Me.TxtPrimaryWorkUnit.Text = Me.objWU.GetPrimary(Me.TxtWorkUnitCode.Text)
                        Me.TxtWorkUnitCode.Text = Me.objWU.GetCode(Me.TxtPrimaryWorkUnit.Text)
                        Me.TxtWorkUnitDesc.Text = Me.objWU.GetDesc(Me.TxtPrimaryWorkUnit.Text)
                        Me.TxtPlanWorkHour.Focus()
                    Else
                        Me.BtnFindWorkUnit.PerformClick()
                    End If
                Else
                    Me.BtnFindWorkUnit.PerformClick()
                End If
                Me.UpdateLocation(CDbl(Me.TxtPrimaryWorkUnit.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWorkUnitCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWorkUnitCode.Leave
        Dim hj As New Core_Module.FrmMasterWorkUnit_Cari
        hj.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtWorkUnitCode.Text <> "" Then
                If Me.objWU.ApakahSudahAda(Me.TxtWorkUnitCode.Text) = True Then
                    Me.TxtPrimaryWorkUnit.Text = Me.objWU.GetPrimary(Me.TxtWorkUnitCode.Text)
                    Me.TxtWorkUnitCode.Text = Me.objWU.GetCode(Me.TxtPrimaryWorkUnit.Text)
                    Me.TxtWorkUnitDesc.Text = Me.objWU.GetDesc(Me.TxtPrimaryWorkUnit.Text)
                    Me.TxtPlanWorkHour.Focus()
                Else
                    If Me.TxtWorkUnitCode.Text <> "" Then
                        hj.txtCari.Text = Me.TxtWorkUnitCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryWorkUnit > 0 Then
                        Me.TxtPrimaryWorkUnit.Text = hj.PrimaryWorkUnit
                        Me.TxtWorkUnitCode.Text = Me.objWU.GetCode(hj.PrimaryWorkUnit)
                        Me.TxtWorkUnitDesc.Text = Me.objWU.GetDesc(hj.PrimaryWorkUnit)
                        Me.TxtPlanWorkHour.Focus()
                    Else
                        Me.TxtPrimaryWorkUnit.Text = "-1"
                        Me.TxtWorkUnitCode.Clear()
                        Me.TxtWorkUnitDesc.Clear()
                        Me.TxtWorkUnitCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryWorkUnit.Text = "-1"
                Me.TxtWorkUnitCode.Clear()
                Me.TxtWorkUnitDesc.Clear()
                'Me.TxtRemarks.Focus()
            End If

            Me.UpdateLocation(CDbl(Me.TxtPrimaryWorkUnit.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindWorkUnit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindWorkUnit.Click
        Try
            Dim hj As New Core_Module.FrmMasterWorkUnit_Cari
            Try
                hj.BackColorAll = Me.BackColorAll
                If Me.TxtWorkUnitCode.Text <> "" Then
                    hj.txtCari.Text = Me.TxtWorkUnitCode.Text
                End If

                hj.ShowDialog()
                If hj.PrimaryWorkUnit > 0 Then
                    Me.TxtPrimaryWorkUnit.Text = hj.PrimaryWorkUnit
                    Me.TxtWorkUnitCode.Text = Me.objWU.GetCode(hj.PrimaryWorkUnit)
                    Me.TxtWorkUnitDesc.Text = Me.objWU.GetDesc(hj.PrimaryWorkUnit)
                    Me.TxtPlanWorkHour.Focus()
                Else
                    Me.TxtPrimaryWorkUnit.Text = "-1"
                    Me.TxtWorkUnitCode.Clear()
                    Me.TxtWorkUnitDesc.Clear()
                    Me.TxtWorkUnitCode.Focus()
                End If

                Me.UpdateLocation(CDbl(Me.TxtPrimaryWorkUnit.Text))
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Finally
                hj.Dispose()
            End Try
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWorkUnitCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWorkUnitCode.GotFocus, TxtTargetProduceAdd.GotFocus, TxtPlanWorkHour.GotFocus, TxtPlanProdRate.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.BtnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtDocNo.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWorkUnitCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWorkUnitCode.LostFocus, TxtTargetProduceAdd.LostFocus, TxtPlanWorkHour.LostFocus, TxtPlanProdRate.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.BtnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtDocNo.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimaryWorkUnit.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateLocation(ByVal vPrimaryWorkUnit As Double)
        Dim PrimaryStandByLocation, PrimaryPickUpLocation As Double
        Dim StandByLocation, PickUpLocation As String
        Dim rowD As DataGridViewRow

        Try
            PrimaryStandByLocation = objWU.GetPrimaryInputLocation(vPrimaryWorkUnit)
            PrimaryPickUpLocation = objWU.GetPrimaryOutputLocation(vPrimaryWorkUnit)
            StandByLocation = objLoc.GetStorageAreaCode(PrimaryStandByLocation)
            PickUpLocation = objLoc.GetStorageAreaCode(PrimaryPickUpLocation)

            For i As Integer = 1 To Me.DgvItemRM.RowCount
                rowD = Me.DgvItemRM.Rows(Me.DgvItemRM.RowCount - i)
                rowD.Cells(Me.ColRM_PrimaryStandByLocation.Index).Value = PrimaryStandByLocation
                rowD.Cells(Me.ColRM_PrimaryPickUpLocation.Index).Value = PrimaryPickUpLocation
                rowD.Cells(Me.ColRM_StandByLocation.Index).Value = StandByLocation
                rowD.Cells(Me.ColRM_PickUpLocation.Index).Value = PickUpLocation
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Me.Timer1.Enabled = False
            Me.DgvItemRM.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class