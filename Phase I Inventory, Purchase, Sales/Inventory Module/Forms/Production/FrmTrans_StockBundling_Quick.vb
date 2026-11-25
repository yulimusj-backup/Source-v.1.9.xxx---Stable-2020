Option Explicit On
Option Strict Off

Public Class FrmTrans_StockBundling_Quick
    Dim TblItemWO As New DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
    Dim ObjMOItem As New ClsTrans_ManufacturingOrder_Item
    Dim ObjMO As New ClsTrans_ManufacturingOrder
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjProses As New Core_Module.ClsMasterProses
    Dim ObjCD As New Core_Module.ClsCekData
    Dim TBLWorkUnit As New DseTrans_ManufacturingOrder.DataTable9DataTable
    Dim DTAWorkUnit As New DseTrans_ManufacturingOrderTableAdapters.DataTable9TableAdapter
    Dim rowWorkUnit As DseTrans_ManufacturingOrder.DataTable9Row
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser
    Dim ObjSch As New ClsTrans_Schedule
    Dim ObjWU As New Core_Module.ClsMasterWorkUnit
    Dim _Stop As Boolean = False

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = "" 

    Private Sub LoadDGV()
        Try
            'load data item wo
            Me.TblItemWO.Rows.Clear()
            Me.dgvItemWO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemWO.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItemWO.Rows.Clear()

            If Me.ObjMO.IsScheduled(Me.TxtIdentifierJO.Text.Substring(0, Me.TxtIdentifierJO.Text.IndexOf("-")), Me.TxtIdentifierJO.Text) = False Then
                MsgBox("Please specified scheduled job order", MsgBoxStyle.Information, "Information")
                Me.TxtIdentifierJO.Text = ""
                Me.TxtJONumber.Text = ""
                Me.TxtPrimaryMO.Text = "-1"
                Me.TxtMONumber.Text = ""

                Me.KosongkanDGV()

                Exit Sub
            End If
            Me.ObjMOItem.IsiDataItemWO(Me.TxtIdentifierJO.Text, Me.TblItemWO)
            Me.DTAWorkUnit.Fill(Me.TBLWorkUnit, Me.TxtIdentifierJO.Text)

            'tampilkan ke dalam dgv
            Me.LoadDgvItemWO_TableView(Me.TxtIdentifierJO.Text, Me.TblItemWO, Me.dgvItemWO)

            Me.dgvItemWO.ReadOnly = False
            Me.dgvItemWO.AllowUserToAddRows = False
            Me.dgvItemWO.AllowUserToDeleteRows = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            If Me.dgvItemWO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None Then
                Me.dgvItemWO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_DetailWO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.DTAWorkUnit.Connection.ConnectionString = Core_Module.GetConstring
            Me.dgvItemWO.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.ColWOTbl_Remarks.Visible = False
            Me.SetColumns()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub LoadDgvItemWO_TableView(ByVal vIdentifierJO As String, ByVal vDataTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vDGV As DataGridView)
        Dim DataTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
        Dim RowTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowDGV As DataGridViewRow
        Dim TBLDummy As New DseTrans_ManufacturingOrder.tbl_codenoDataTable
        Dim rowDummy As DseTrans_ManufacturingOrder.tbl_codenoRow
        Dim TBLDummy_Process As New DseTrans_ManufacturingOrder.tbl_processDataTable
        Dim rowDummy_Process As DseTrans_ManufacturingOrder.tbl_processRow
        'Dim Status_ As String = ""
        Dim Filter As String = ""

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGV.Rows.Clear()

            DataTable_ = vDataTable_.Copy
            Filter = "[Identifier JO] = '" & vIdentifierJO + "'"
            Me.TxtCodeNo.Text = Me.TxtCodeNo.Text.Trim
            If Me.TxtCodeNo.Text <> "" Then
                Filter = Filter + " AND [Product Code] LIKE '" + Me.TxtCodeNo.Text + "%'"
            End If
            Me.TxtDesc.Text = Me.TxtDesc.Text.Trim
            If Me.TxtDesc.Text <> "" Then
                Filter = Filter + " AND [Description] LIKE '" + Me.TxtDesc.Text + "%'"
            End If
            If Me.txtPrimaryProses.Text <> "-1" And Me.txtPrimaryProses.Text <> "" Then
                Filter = Filter + " AND [Primary Process] =" + Me.txtPrimaryProses.Text + ""
            End If
            vDataTable_.DefaultView.RowFilter = Filter
            vDataTable_.DefaultView.Sort = "Level DESC, Primary"

            If vDataTable_.DefaultView.Count > 0 Then
                For Each DV As Data.DataRowView In vDataTable_.DefaultView
                    RowTable_ = DV.Row
                    If RowTable_.Qty_Target_Prod > 0 Then
                        vDGV.Rows.Add()
                        rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                        'TBLDummy.DefaultView.RowFilter = "[Primary Product] = " + RowTable_.Primary_Product.ToString
                        'If TBLDummy.DefaultView.Count = 0 Then
                        '    rowDummy = TBLDummy.Newtbl_codenoRow
                        '    rowDummy.Primary_Product = RowTable_.Primary_Product
                        '    rowDummy.Code_No = Me.ObjProduct.GetProductCode(RowTable_.Primary_Product)
                        '    rowDummy.Desc = Me.ObjProduct.GetDescription(RowTable_.Primary_Product)
                        '    TBLDummy.Rows.Add(rowDummy)
                        'Else
                        '    rowDummy = TBLDummy.DefaultView(0).Row
                        'End If
                        'rowDGV.Cells(Me.ColWOTbl_CodeNo.Index).Value = rowDummy.Code_No
                        'rowDGV.Cells(Me.ColWOTbl_Desc.Index).Value = rowDummy.Desc
                        rowDGV.Cells(Me.ColWOTbl_CodeNo.Index).Value = RowTable_.Product_Code
                        rowDGV.Cells(Me.ColWOTbl_Desc.Index).Value = RowTable_.Description
                        rowDGV.Cells(Me.ColWOTbl_IdentifierJO.Index).Value = RowTable_.Identifier_JO
                        rowDGV.Cells(Me.ColWOTbl_IdentifierWO.Index).Value = RowTable_.Identifier
                        rowDGV.Cells(Me.ColWOTbl_No.Index).Value = rowDGV.Index + 1
                        rowDGV.Cells(Me.ColWOTbl_PrimaryProcess.Index).Value = RowTable_.Primary_Process
                        TBLDummy_Process.DefaultView.RowFilter = "[Primary Process] = " + RowTable_.Primary_Process.ToString
                        If TBLDummy_Process.DefaultView.Count = 0 Then
                            rowDummy_Process = TBLDummy_Process.Newtbl_processRow
                            rowDummy_Process.Primary_Process = RowTable_.Primary_Process
                            rowDummy_Process.Code = Me.ObjProses.GetCode(RowTable_.Primary_Process)
                            rowDummy_Process.Desc = Me.ObjProses.GetDesc(rowDummy_Process.Primary_Process)
                            TBLDummy_Process.Rows.Add(rowDummy_Process)
                        Else
                            rowDummy_Process = TBLDummy_Process.DefaultView(0).Row
                        End If
                        rowDGV.Cells(Me.ColWOTbl_ProcessCode.Index).Value = rowDummy_Process.Code
                        rowDGV.Cells(Me.ColWOTbl_ProcessDesc.Index).Value = rowDummy_Process.Desc
                        rowDGV.Cells(Me.ColWOTbl_Progress.Index).Value = Me.ObjCD.DigitGrouping((RowTable_.Qty_Actual_Prod / RowTable_.Qty_Target_Prod) * 100, 2)
                        rowDGV.Cells(Me.ColWOTbl_Remarks.Index).Value = RowTable_.Remarks
                        rowDGV.Cells(Me.ColWOTbl_WONO.Index).Value = RowTable_.WO_Number
                        Me.TBLWorkUnit.DefaultView.RowFilter = "[identifier wo] = '" + RowTable_.Identifier + "'"
                        If Me.TBLWorkUnit.DefaultView.Count = 0 Then
                            MsgBox("Logic error, work unit not found.", MsgBoxStyle.Information, "Information")
                        ElseIf Me.TBLWorkUnit.DefaultView.Count = 1 Then
                            Me.rowWorkUnit = Me.TBLWorkUnit.DefaultView(0).Row
                            rowDGV.Cells(Me.ColWOTbl_WorkUnit.Index).Value = Me.rowWorkUnit.work_unit_code
                        Else
                            rowDGV.Cells(Me.ColWOTbl_WorkUnit.Index).Value = "Multiple"
                        End If

                        rowDGV.Cells(Me.ColWOTbl_StartDate.Index).Value = Format(RowTable_.Plan_Start_Time, "dd MMM yyyy")
                        rowDGV.Cells(Me.ColWOTbl_StartTime.Index).Value = Format(RowTable_.Plan_Start_Time, "HH.mm")
                        rowDGV.Cells(Me.ColWOTbl_EndDate.Index).Value = Format(RowTable_.Plan_End_Time, "dd MMM yyyy")
                        rowDGV.Cells(Me.ColWOTbl_EndTime.Index).Value = Format(RowTable_.Plan_End_Time, "HH.mm")
                    End If
                Next
            End If

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJONumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJONumber.GotFocus
        Try
            If Me.TxtJONumber.Enabled = True Then
                If Me.TxtJONumber.ReadOnly = True Then
                    Me.TxtJONumber.BackColor = Me.TxtMONumber.BackColor
                Else
                    Me.TxtJONumber.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJONumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtJONumber.KeyDown
        Dim i As New FrmTrans_ManufacturingOrder_PrintJO_Cari
        Dim Jml As Long = 0
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjMO.GetJmlJO(Me.TxtJONumber.Text)
                If Jml = 1 Then
                    Me.TxtIdentifierJO.Text = Me.ObjMO.GetIdentifierJO(Me.TxtJONumber.Text)
                    Me.TxtPrimaryMO.Text = Me.TxtIdentifierJO.Text.Substring(0, Me.TxtIdentifierJO.Text.IndexOf("-"))
                    Me.TxtMONumber.Text = Me.ObjMO.GetNoTransaksi(Me.TxtPrimaryMO.Text)
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.TxtJONumber.Text <> "" Then
                        i.txtCari.Text = Me.TxtJONumber.Text.Trim
                    End If
                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()

                    If i.IdentifierJO <> "" Then
                        Me.TxtIdentifierJO.Text = i.IdentifierJO
                        Me.TxtJONumber.Text = Me.ObjMO.GetJONumber(Me.TxtIdentifierJO.Text)
                        Me.TxtPrimaryMO.Text = Me.TxtIdentifierJO.Text.Substring(0, Me.TxtIdentifierJO.Text.IndexOf("-"))
                        Me.TxtMONumber.Text = Me.ObjMO.GetNoTransaksi(Me.TxtPrimaryMO.Text)
                    Else
                        Me.TxtJONumber.Clear()
                        Me.TxtIdentifierJO.Text = ""
                        Me.TxtPrimaryMO.Text = "-1"
                        Me.TxtMONumber.Text = ""
                        Me.KosongkanDGV()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtJONumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJONumber.Leave
        Dim i As New FrmTrans_ManufacturingOrder_PrintJO_Cari

        Try
            If Me.TxtJONumber.Text <> "" Then
                Me.TxtJONumber.Text = Me.TxtJONumber.Text.Trim

                If Me.ObjMO.GetJmlJO(Me.TxtJONumber.Text) = 1 Then
                    Me.TxtIdentifierJO.Text = Me.ObjMO.GetIdentifierJO(Me.TxtJONumber.Text)
                    Me.TxtPrimaryMO.Text = Me.TxtIdentifierJO.Text.Substring(0, Me.TxtIdentifierJO.Text.IndexOf("-"))
                    Me.TxtMONumber.Text = Me.ObjMO.GetNoTransaksi(Me.TxtPrimaryMO.Text)
                Else
                    i.BackColorAll = Me.BackColorAll
                    If Me.TxtJONumber.Text <> "" Then
                        i.txtCari.Text = Me.TxtJONumber.Text
                    End If

                    i.ShowDialog()
                    If i.IdentifierJO <> "" Then
                        Me.TxtIdentifierJO.Text = i.IdentifierJO
                        Me.TxtJONumber.Text = Me.ObjMO.GetJONumber(Me.TxtIdentifierJO.Text)
                        Me.TxtPrimaryMO.Text = Me.TxtIdentifierJO.Text.Substring(0, Me.TxtIdentifierJO.Text.IndexOf("-"))
                        Me.TxtMONumber.Text = Me.ObjMO.GetNoTransaksi(Me.TxtPrimaryMO.Text)
                    Else
                        Me.TxtJONumber.Clear()
                        Me.TxtIdentifierJO.Text = ""
                        Me.TxtPrimaryMO.Text = "-1"
                        Me.TxtMONumber.Text = ""
                        Me.KosongkanDGV()
                    End If
                End If
            Else
                Me.TxtJONumber.Clear()
                Me.TxtIdentifierJO.Text = ""
                Me.TxtPrimaryMO.Text = "-1"
                Me.TxtMONumber.Text = ""
                Me.KosongkanDGV()
            End If

            If Me.TxtJONumber.Text <> "" Then
                Me.LoadDGV()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub KosongkanDGV()
        Try
            Me.dgvItemWO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemWO.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItemWO.Rows.Clear()
            Me.dgvItemWO.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Me.txtProcessCode.Text.Clone()
            Me.txtPrimaryProses.Text = "-1"
            Me.TxtCodeNo.Clear()
            Me.TxtDesc.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindJO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindJO.Click
        Dim i As New FrmTrans_ManufacturingOrder_PrintJO_Cari

        Try
            i.BackColorAll = Me.BackColorAll
            'kalo non invoice do only lemot!!!
            'i.Non_InvoiceDOOnly = True

            If Me.TxtJONumber.Text <> "" Then
                i.txtCari.Text = Me.TxtJONumber.Text.Trim
            End If

            i.ShowDialog()
            If i.IdentifierJO <> "" Then
                Me.TxtIdentifierJO.Text = i.IdentifierJO
                Me.TxtJONumber.Text = Me.ObjMO.GetJONumber(Me.TxtIdentifierJO.Text)
                Me.TxtPrimaryMO.Text = Me.TxtIdentifierJO.Text.Substring(0, Me.TxtIdentifierJO.Text.IndexOf("-"))
                Me.TxtMONumber.Text = Me.ObjMO.GetNoTransaksi(Me.TxtPrimaryMO.Text)

                Me.LoadDGV()
            Else
                Me.TxtJONumber.Clear()
                Me.TxtIdentifierJO.Text = ""
                Me.TxtPrimaryMO.Text = "-1"
                Me.TxtMONumber.Text = ""
                Me.KosongkanDGV()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtJONumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJONumber.LostFocus
        Try
            Me.TxtJONumber.BackColor = Me.TxtIdentifierJO.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemWO_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemWO.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColWOTbl_CreateLog.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemWO_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemWO.CellContentClick
        Dim rowDGV As DataGridViewRow

        Try
            If e.ColumnIndex <> -1 And e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColWOTbl_View.Index Then
                    rowDGV = Me.dgvItemWO.Rows(e.RowIndex)

                    If rowDGV.Cells(Me.ColWOTbl_WorkUnit.Index).Value = "Multiple" Then
                        Dim ij As New FrmTrans_StockBundling_Quick2
                        ij.IdentifierJO = rowDGV.Cells(Me.ColWOTbl_IdentifierJO.Index).Value
                        ij.IdentifierWO = rowDGV.Cells(Me.ColWOTbl_IdentifierWO.Index).Value

                        If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                            Dim rowWU As DseTrans_ManufacturingOrder.DataTable10Row
                            rowWU = ij.TBLWU.DefaultView(ij.cboWorkUnit.SelectedIndex).Row

                            Me.ShowLog(rowDGV.Cells(Me.ColWOTbl_IdentifierWO.Index).Value, _
                            rowDGV.Cells(Me.ColWOTbl_IdentifierJO.Index).Value, rowWU.primary_work_unit)
                        End If
                    Else
                        Me.ShowLog(rowDGV.Cells(Me.ColWOTbl_IdentifierWO.Index).Value, _
                        rowDGV.Cells(Me.ColWOTbl_IdentifierJO.Index).Value, Me.ObjWU.GetPrimary(rowDGV.Cells(Me.ColWOTbl_WorkUnit.Index).Value))
                    End If
                ElseIf e.ColumnIndex = Me.ColWOTbl_Edit.Index Then
                    rowDGV = Me.dgvItemWO.Rows(e.RowIndex)

                    If rowDGV.Cells(Me.ColWOTbl_WorkUnit.Index).Value = "Multiple" Then
                        Dim ij As New FrmTrans_StockBundling_Quick2
                        ij.IdentifierJO = rowDGV.Cells(Me.ColWOTbl_IdentifierJO.Index).Value
                        ij.IdentifierWO = rowDGV.Cells(Me.ColWOTbl_IdentifierWO.Index).Value

                        If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                            Dim rowWU As DseTrans_ManufacturingOrder.DataTable10Row
                            rowWU = ij.TBLWU.DefaultView(ij.cboWorkUnit.SelectedIndex).Row

                            Me.EditLog(rowDGV.Cells(Me.ColWOTbl_IdentifierWO.Index).Value, _
                            rowDGV.Cells(Me.ColWOTbl_IdentifierJO.Index).Value, rowWU.primary_work_unit, rowDGV.Index)
                        End If
                    Else
                        Me.EditLog(rowDGV.Cells(Me.ColWOTbl_IdentifierWO.Index).Value, _
                        rowDGV.Cells(Me.ColWOTbl_IdentifierJO.Index).Value, Me.ObjWU.GetPrimary(rowDGV.Cells(Me.ColWOTbl_WorkUnit.Index).Value), rowDGV.Index)
                    End If
                ElseIf e.ColumnIndex = Me.ColWOTbl_Delete.Index Then
                    rowDGV = Me.dgvItemWO.Rows(e.RowIndex)

                    If rowDGV.Cells(Me.ColWOTbl_WorkUnit.Index).Value = "Multiple" Then
                        Dim ij As New FrmTrans_StockBundling_Quick2
                        ij.IdentifierJO = rowDGV.Cells(Me.ColWOTbl_IdentifierJO.Index).Value
                        ij.IdentifierWO = rowDGV.Cells(Me.ColWOTbl_IdentifierWO.Index).Value

                        If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                            Dim rowWU As DseTrans_ManufacturingOrder.DataTable10Row
                            rowWU = ij.TBLWU.DefaultView(ij.cboWorkUnit.SelectedIndex).Row

                            Me.DeleteLog(rowDGV.Cells(Me.ColWOTbl_IdentifierWO.Index).Value, _
                            rowDGV.Cells(Me.ColWOTbl_IdentifierJO.Index).Value, rowWU.primary_work_unit, rowDGV.Index)
                        End If
                    Else
                        Me.DeleteLog(rowDGV.Cells(Me.ColWOTbl_IdentifierWO.Index).Value, _
                        rowDGV.Cells(Me.ColWOTbl_IdentifierJO.Index).Value, Me.ObjWU.GetPrimary(rowDGV.Cells(Me.ColWOTbl_WorkUnit.Index).Value), rowDGV.Index)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DeleteLog(ByVal IdentifierWO As String, ByVal IdentifierJO As String, ByVal PrimaryWorkUnit As Double, ByVal vrowIndex As Integer)
        Dim i As New Inventory_Module.FrmTrans_StockBundling

        Try
            Dim d As New DseTrans_ScheduleTableAdapters.DataTable5TableAdapter
            Dim t As New DseTrans_Schedule.DataTable5DataTable
            Dim r As DseTrans_Schedule.DataTable5Row
            d.Connection.ConnectionString = Core_Module.GetConstring
            d.Fill(t, ClsTrans_StockBundling.EnumTipeBundling.Produksi1, IdentifierWO, PrimaryWorkUnit, IdentifierJO)

            If t.Count > 1 Then
                Dim f As New FrmTrans_StockBundling_ChooseOne(t)
                f.ShowDialog()
                i._PrimaryBundling = f.PrimaryBundling
            ElseIf t.Count = 1 Then
                r = t.Rows(0)
                i._PrimaryBundling = r.Primary
            ElseIf t.Count = 0 Then
                MsgBox("This WO number don't have a log.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If i._PrimaryBundling <> -1 Then
                i.UserCode = Me.UserCode
                i.BackColorAll = Me.BackColorAll
                i._Command = FrmTrans_StockBundling.EnumCommand.DELETE

                i.ShowDialog()

                Me.dgvItemWO.Item(Me.ColWOTbl_Progress.Index, vrowIndex).Value = Me.ObjCD.DigitGrouping(Me.ObjSch.GetProgress(IdentifierJO, _
                     IdentifierWO, PrimaryWorkUnit), 2)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EditLog(ByVal IdentifierWO As String, ByVal IdentifierJO As String, ByVal PrimaryWorkUnit As Double, ByVal vrowIndex As Integer)
        Dim i As New Inventory_Module.FrmTrans_StockBundling
        Try
            If IdentifierWO <> "" And PrimaryWorkUnit <> -1 Then
                Dim d As New DseTrans_ScheduleTableAdapters.DataTable5TableAdapter
                Dim t As New DseTrans_Schedule.DataTable5DataTable
                Dim r As DseTrans_Schedule.DataTable5Row
                d.Connection.ConnectionString = Core_Module.GetConstring
                d.Fill(t, ClsTrans_StockBundling.EnumTipeBundling.Produksi1, IdentifierWO, PrimaryWorkUnit, IdentifierJO)

                If t.Count > 1 Then
                    Dim f As New FrmTrans_StockBundling_ChooseOne(t)
                    f.ShowDialog()
                    i._PrimaryBundling = f.PrimaryBundling
                ElseIf t.Count = 1 Then
                    r = t.Rows(0)
                    i._PrimaryBundling = r.Primary
                ElseIf t.Count = 0 Then
                    MsgBox("This WO number don't have a log.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If i._PrimaryBundling <> -1 Then
                    i.UserCode = Me.UserCode
                    i.BackColorAll = Me.BackColorAll
                    i._Command = FrmTrans_StockBundling.EnumCommand.EDIT

                    i.ShowDialog()

                    Me.dgvItemWO.Item(Me.ColWOTbl_Progress.Index, vrowIndex).Value = Me.ObjCD.DigitGrouping(Me.ObjSch.GetProgress(IdentifierJO, _
                     IdentifierWO, PrimaryWorkUnit), 2)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub ShowLog(ByVal IdentifierWO As String, ByVal IdentifierJO As String, ByVal PrimaryWorkUnit As Double)
        Dim i As New Inventory_Module.FrmTrans_StockBundling

        Try
            If IdentifierWO <> "" And PrimaryWorkUnit <> -1 Then
                Dim d As New DseTrans_ScheduleTableAdapters.DataTable5TableAdapter
                Dim t As New DseTrans_Schedule.DataTable5DataTable
                Dim r As DseTrans_Schedule.DataTable5Row
                d.Connection.ConnectionString = Core_Module.GetConstring
                d.Fill(t, ClsTrans_StockBundling.EnumTipeBundling.Produksi1, IdentifierWO, PrimaryWorkUnit, IdentifierJO)

                If t.Count > 1 Then
                    Dim f As New FrmTrans_StockBundling_ChooseOne(t)
                    f.ShowDialog()
                    i._PrimaryBundling = f.PrimaryBundling
                ElseIf t.Count = 1 Then
                    r = t.Rows(0)
                    i._PrimaryBundling = r.Primary
                ElseIf t.Count = 0 Then
                    MsgBox("This WO number don't have a log.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If i._PrimaryBundling <> -1 Then
                    Dim ij As Inventory_Module.FrmTrans_StockBundling
                    Dim j As Windows.Forms.Form
                    For Each j In Me.ParentForm.MdiChildren
                        ij = TryCast(j, Inventory_Module.FrmTrans_StockBundling)
                        If IsNothing(ij) = False Then
                            ''jika sudah terbuka sebelumnya
                            ij.UserCode = Me.UserCode
                            ij._PrimaryBundling = i._PrimaryBundling
                            ij._Command = FrmTrans_StockBundling.EnumCommand.VIEW
                            ij.BringToFront()
                            ij.Focus()
                            Exit Sub
                        End If
                    Next
                    ''jika tidak sedang terbuka
                    i.UserCode = Me.UserCode
                    i.BackColorAll = Core_Module.GetSMESSBackColorAll
                    i._Command = FrmTrans_StockBundling.EnumCommand.VIEW
                    i.MdiParent = Me.ParentForm
                    i.ShowIcon = False
                    i.Show()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub btnCreateLog_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCreateLog.Click
        Dim Jml As Integer = 0
        Dim rowdgv As DataGridViewRow
        Dim JmlTotal As Integer = 0

        Try
            Me._Stop = False
            For Each rowdgv In Me.dgvItemWO.Rows
                If rowdgv.Cells(Me.ColWOTbl_CreateLog.Index).Value = Boolean.TrueString Then
                    JmlTotal += 1
                End If
            Next
            For Each rowdgv In Me.dgvItemWO.Rows
                If rowdgv.Cells(Me.ColWOTbl_CreateLog.Index).Value = Boolean.TrueString Then
                    If Me._Stop = True Then
                        MsgBox("Process stopped.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Jml += 1
                    If rowdgv.Cells(Me.ColWOTbl_WorkUnit.Index).Value = "Multiple" Then
                        Dim ij As New FrmTrans_StockBundling_Quick2
                        ij.IdentifierJO = rowdgv.Cells(Me.ColWOTbl_IdentifierJO.Index).Value
                        ij.IdentifierWO = rowdgv.Cells(Me.ColWOTbl_IdentifierWO.Index).Value

                        If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                            Dim rowWU As DseTrans_ManufacturingOrder.DataTable10Row
                            rowWU = ij.TBLWU.DefaultView(ij.cboWorkUnit.SelectedIndex).Row

                            Me.CreateLog(rowdgv.Cells(Me.ColWOTbl_IdentifierWO.Index).Value, _
                            rowdgv.Cells(Me.ColWOTbl_IdentifierJO.Index).Value, rowWU.primary_work_unit, rowdgv.Index, Jml, JmlTotal)
                        End If
                    Else
                        Me.CreateLog(rowdgv.Cells(Me.ColWOTbl_IdentifierWO.Index).Value, _
                        rowdgv.Cells(Me.ColWOTbl_IdentifierJO.Index).Value, Me.ObjWU.GetPrimary(rowdgv.Cells(Me.ColWOTbl_WorkUnit.Index).Value), _
                        rowdgv.Index, Jml, JmlTotal)
                    End If

                    If Me._Stop = False Then
                        rowdgv.Cells(Me.ColWOTbl_CreateLog.Index).Value = Boolean.FalseString
                    End If
                End If
            Next

            If Jml = 0 Then
                MsgBox("Please specify at least 1 work order to create log", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbColumnSetting_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbColumnSetting.LinkClicked
        Dim ij As New FrmTrans_StockBundling_Quick1

        Try
            ij.UserCode = Me.UserCode
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.SetColumns()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetColumns()
        Dim Dummy As String = ""

        Try
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowCodeNo, Me.UserCode)
            If Dummy = "" Then
                Dummy = "1"
            End If
            If Dummy = "1" Then
                Me.ColWOTbl_CodeNo.Visible = True
            Else
                Me.ColWOTbl_CodeNo.Visible = False
            End If
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowDesc, Me.UserCode)
            If Dummy = "" Then
                Dummy = "1"
            End If
            If Dummy = "1" Then
                Me.ColWOTbl_Desc.Visible = True
            Else
                Me.ColWOTbl_Desc.Visible = False
            End If
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
           Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowProcessCode, Me.UserCode)
            If Dummy = "" Then
                Dummy = "1"
            End If
            If Dummy = "1" Then
                Me.ColWOTbl_ProcessCode.Visible = True
            Else
                Me.ColWOTbl_ProcessCode.Visible = False
            End If
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
          Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowProcessDesc, Me.UserCode)
            If Dummy = "" Then
                Dummy = "1"
            End If
            If Dummy = "1" Then
                Me.ColWOTbl_ProcessDesc.Visible = True
            Else
                Me.ColWOTbl_ProcessDesc.Visible = False
            End If
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
          Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowStartTime, Me.UserCode)
            If Dummy = "" Then
                Dummy = "1"
            End If
            If Dummy = "1" Then
                Me.ColWOTbl_StartDate.Visible = True
                Me.ColWOTbl_StartTime.Visible = True
            Else
                Me.ColWOTbl_StartDate.Visible = False
                Me.ColWOTbl_StartTime.Visible = False
            End If
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowEndTime, Me.UserCode)
            If Dummy = "" Then
                Dummy = "1"
            End If
            If Dummy = "1" Then
                Me.ColWOTbl_EndDate.Visible = True
                Me.ColWOTbl_EndTime.Visible = True
            Else
                Me.ColWOTbl_EndDate.Visible = False
                Me.ColWOTbl_EndTime.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreateLog(ByVal IdentifierWO As String, ByVal IdentifierJO As String, _
    ByVal PrimaryWorkUnit As Double, ByVal vRowIndex As Integer, ByVal vUrutan As Double, ByVal vTotalLog As Double)
        Dim i As New Inventory_Module.FrmTrans_StockBundling

        Try
            If IdentifierWO <> "" And PrimaryWorkUnit <> -1 Then
                i.UserCode = Me.UserCode
                i.BackColorAll = Me.BackColorAll
                i._IdentifierWO = IdentifierWO
                i._IdentifierJO = IdentifierJO
                i._PrimaryWorkUnit = PrimaryWorkUnit
                i.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Produksi1
                i._Command = FrmTrans_StockBundling.EnumCommand.NEW_

                i.prg1.Visible = True
                i.lblProgress.Visible = True
                i.lblProgress.Text = "Data " + Me.ObjCD.DigitGrouping(vUrutan, 0) + " of " + Me.ObjCD.DigitGrouping(vTotalLog, 0)
                i.btnFind.Visible = False
                i.btnFind.TabStop = False
                i.btnNew.Visible = False
                i.btnNew.TabStop = False
                i.btnEdit.Visible = False
                i.btnEdit.TabStop = False
                i.btnDelete.Visible = False
                i.btnDelete.TabStop = False
                i.btnExit.Visible = False
                i.btnExit.TabStop = False
                i.btnStop.Visible = True
                i.btnStop.TabStop = True

                i.ShowDialog()

                Me.dgvItemWO.Item(Me.ColWOTbl_Progress.Index, vRowIndex).Value = Me.ObjCD.DigitGrouping(Me.ObjSch.GetProgress(IdentifierJO, _
                IdentifierWO, PrimaryWorkUnit), 2)

                If i._Stop = True Then
                    Me._Stop = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemWO.Rows
                If CDec(rowDGV.Cells(Me.ColWOTbl_Progress.Index).Value) < 100 Then
                    rowDGV.Cells(Me.ColWOTbl_CreateLog.Index).Value = Boolean.TrueString
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemWO.Rows
                rowDGV.Cells(Me.ColWOTbl_CreateLog.Index).Value = Boolean.FalseString
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.GotFocus
        Try
            Me.TxtCodeNo.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        Try
            Me.TxtCodeNo.BackColor = Me.TxtPrimaryMO.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.GotFocus
        Try
            Me.TxtDesc.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.LostFocus
        Try
            Me.TxtDesc.BackColor = Me.TxtPrimaryMO.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtProcess_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProcessCode.GotFocus
        Try
            Me.txtProcessCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtProcess_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProcessCode.LostFocus
        Try
            Me.txtProcessCode.BackColor = Me.TxtPrimaryMO.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
        Dim i As New Core_Module.frmCariMasterProduct2
        Dim Jml As Long = 0
        Dim PrimaryProduct As Double = -1

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNo.Text)
                PrimaryProduct = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)

                If Jml = 1 And Me.ObjProduct.ApakahGoods(PrimaryProduct) = True Then
                    Me.TxtDesc.Text = Me.ObjProduct.GetDescription(PrimaryProduct) 
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then 
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(PrimaryProduct)
                        Me.TxtDesc.Text = Me.ObjProduct.GetDescription(PrimaryProduct)
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCodeNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindCodeNo.Click
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct2

        Try
            frmFindProduct.txtCari.Text = Me.TxtCodeNo.Text
            frmFindProduct.ShowDialog()
            If frmFindProduct.PrimaryProduct <> "" Then
                Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(frmFindProduct.PrimaryProduct)
                Me.TxtDesc.Text = Me.ObjProduct.GetDescription(frmFindProduct.PrimaryProduct)
                Me.TxtDesc.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilterProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterProcess.Click
        Try
            If Me.TxtIdentifierJO.Text <> "" Then
                Me.LoadDGV()
            Else
                Me.KosongkanDGV()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProcessCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProcessCode.KeyDown
        Dim i As New Core_Module.FrmMasterProsesCari
        Dim Obj As New Core_Module.ClsMasterProses

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 And Me.txtProcessCode.ReadOnly = False And Me.txtProcessCode.Enabled = True Then
                If Obj.ApakahSudahAda(Me.txtProcessCode.Text) = True Then 
                    Me.txtPrimaryProses.Text = Obj.GetPrimary(Me.txtProcessCode.Text)
                    Me.btnFilterProcess.Focus()
                Else
                    If Me.txtProcessCode.Text <> "" Then
                        i.txtCari.Text = Me.txtProcessCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimary > 0 Then 
                        Me.txtPrimaryProses.Text = i.mPrimary
                        Me.txtProcessCode.Text = Obj.GetCode(Me.txtPrimaryProses.Text)
                        Me.btnFilterProcess.Focus()
                    Else
                        Me.txtProcessCode.Clear()
                        Me.txtPrimaryProses.Text = "-1"
                        Me.txtProcessCode.Focus() 
                    End If
                End If 
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub txtProcessCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProcessCode.Leave
        Dim i As New Core_Module.FrmMasterProsesCari
        Dim obj As New Core_Module.ClsMasterProses

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.txtProcessCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.txtProcessCode.Text) = True Then
                    Me.txtPrimaryProses.Text = obj.GetPrimary(Me.txtProcessCode.Text)
                Else
                    If Me.txtProcessCode.Text <> "" Then
                        i.txtCari.Text = Me.txtProcessCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimary > 0 Then 
                        Me.txtPrimaryProses.Text = i.mPrimary
                        Me.txtProcessCode.Text = obj.GetCode(Me.txtPrimaryProses.Text)
                        Me.btnFilterProcess.Focus()
                    Else
                        Me.txtProcessCode.Clear()
                        Me.txtPrimaryProses.Text = "-1" 
                        Me.txtProcessCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryProses.Text = "-1" 
            End If 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindProses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindProcess.Click
        Dim i As New Core_Module.FrmMasterProsesCari
        Dim obj As New Core_Module.ClsMasterProses

        Try
            i.BackColorAll = Me.BackColorAll
            'i.cboCriteria.SelectedIndex = 0
            If Me.txtProcessCode.Text <> "" Then
                i.txtCari.Text = Me.txtProcessCode.Text
            End If

            i.ShowDialog()
            If i.mPrimary > 0 Then 
                Me.txtPrimaryProses.Text = i.mPrimary
                Me.txtProcessCode.Text = obj.GetCode(Me.txtPrimaryProses.Text)
                Me.btnFilterProcess.Focus()
            Else
                Me.txtProcessCode.Clear()
                Me.txtProcessCode.Focus() 
            End If 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub
End Class