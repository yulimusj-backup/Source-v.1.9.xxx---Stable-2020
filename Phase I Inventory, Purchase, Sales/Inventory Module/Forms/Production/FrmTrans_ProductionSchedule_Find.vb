''rosadin-r15

Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_Find

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TBLSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
    Public PrimaryMO As Double = -1
    Public IdentifierJO As String = ""
    Public IdentifierWO As String = ""
    Public PrimaryProcess As Double = -1
    Public PrimaryWorkunit As Double = -1
    Public GroupBy As FrmTrans_ProductionSchedule.enumGroup = FrmTrans_ProductionSchedule.enumGroup.BY_JO

    Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow
    Dim objMO As New ClsTrans_ManufacturingOrder
    Dim objProduct As New Core_Module.clsMasterProduct
    Dim objProcess As New Core_Module.ClsMasterProses
    Dim objWU As New Core_Module.ClsMasterWorkUnit
    Dim objCD As New Core_Module.ClsCekData

    Private Sub FrmTrans_ProductionSchedule_Find_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim TBLSchBarTemp As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim rowSchBarTemp As DseTrans_Schedule.tbl_sch_barRow
        Try
            TBLSchBarTemp = Me.TBLSchBar.Copy
            If TBLSchBar.Columns.Contains("Finish Good") = False Then
                Me.TBLSchBar.Columns.Add("Finish Good")
                Me.TBLSchBar.Columns("Finish Good").DataType = System.Type.GetType("System.String")
                Me.TBLSchBar.Columns("Finish Good").DefaultValue = ""
            End If
            If Me.GroupBy = FrmTrans_ProductionSchedule.enumGroup.BY_JO Then
                For Each Me.rowSchBar In Me.TBLSchBar
                    If Me.rowSchBar.is_header = 0 Then
                        TBLSchBarTemp.DefaultView.RowFilter = "[is header] = 1 and [identifier jo]='" & Me.rowSchBar.identifier_jo & "'"
                        For Each DV As Data.DataRowView In TBLSchBarTemp.DefaultView
                            rowSchBarTemp = DV.Row
                            If Not String.IsNullOrEmpty(rowSchBarTemp.product_code) Then
                                Me.rowSchBar("Finish Good") = rowSchBarTemp.product_code
                            End If
                        Next
                    End If
                Next
            Else
                Dim DtaJO As New DseTrans_ScheduleTableAdapters.DataTable8TableAdapter
                Dim TblJO As New DseTrans_Schedule.DataTable8DataTable
                Dim RowJO As DseTrans_Schedule.DataTable8Row
                DtaJO.Connection.ConnectionString = Core_Module.GetConstring
                DtaJO.Fill(TblJO)
                For Each Me.rowSchBar In Me.TBLSchBar
                    If Me.rowSchBar.is_header = 0 Then
                        TblJO.DefaultView.RowFilter = "[identifier] = '" & Me.rowSchBar.identifier_jo & "'"
                        For Each DV As Data.DataRowView In TblJO.DefaultView
                            RowJO = DV.Row
                            If Not String.IsNullOrEmpty(RowJO.Product_Code) Then
                                Me.rowSchBar("Finish Good") = RowJO.Product_Code
                            End If
                        Next
                    End If
                Next
            End If
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DialogResult = Windows.Forms.DialogResult.None
            Me.txtMO.Focus()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMO.GotFocus, txtJO.GotFocus, txtFinishGood.GotFocus, txtProcess.GotFocus, txtWO.GotFocus, txtWIP.GotFocus, txtWorkunit.GotFocus
        Dim TextBox_ As TextBox
        Try
            TextBox_ = TryCast(sender, TextBox)
            TextBox_.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMO.LostFocus, txtJO.LostFocus, txtFinishGood.LostFocus, txtProcess.LostFocus, txtWO.LostFocus, txtWIP.LostFocus, txtWorkunit.LostFocus
        Dim TextBox_ As TextBox
        Try
            TextBox_ = TryCast(sender, TextBox)
            TextBox_.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindMO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindMO.Click
        Dim FrmFindMO As New FrmTrans_ManufacturingOrder_Cari
        Try
            FrmFindMO.txtCari.Text = Me.txtMO.Text
            FrmFindMO.ShowDialog()
            If FrmFindMO.Primary <> -1 Then
                Me.txtMO.Text = Me.objMO.GetNoTransaksi(FrmFindMO.Primary)
                Me.txtJO.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindJO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindJO.Click
        Dim frmFindJO As New FrmTrans_ManufacturingOrder_PrintJO_Cari
        Try
            frmFindJO.txtCari.Text = Me.txtJO.Text
            frmFindJO.ShowDialog()
            If frmFindJO.IdentifierJO <> "" Then
                Me.txtJO.Text = Me.objMO.GetJONumber(frmFindJO.IdentifierJO)
                Me.txtFinishGood.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindFG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindFG.Click
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
            frmFindProduct.txtCari.Text = txtFinishGood.Text
            frmFindProduct.ShowDialog()
            If frmFindProduct.PrimaryProduct <> -1 Then
                Me.txtFinishGood.Text = Me.objProduct.GetProductCode(frmFindProduct.PrimaryProduct)
                Me.txtProcess.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindProcess.Click
        Dim frmFindProcess As New Core_Module.FrmMasterProsesCari
        Try
            frmFindProcess.txtCari.Text = txtProcess.Text
            frmFindProcess.ShowDialog()
            If frmFindProcess.mPrimary <> -1 Then
                Me.txtProcess.Text = Me.objProcess.GetCode(frmFindProcess.mPrimary)
                Me.txtWO.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindWO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindWO.Click
        Dim frmFindWO As New FrmTrans_ManufacturingOrder_WO_Cari
        Try
            frmFindWO.txtCari.Text = txtWO.Text
            frmFindWO.vStatus1 = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled
            frmFindWO.vStatus2 = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.On_Progress
            frmFindWO.vStatus3 = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Finish
            frmFindWO.ShowDialog()
            If frmFindWO.Identifier <> "" Then
                Me.txtWO.Text = Me.objMO.GetWONumber(frmFindWO.Identifier, frmFindWO.IdentifierJO)
                Me.txtWIP.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindWIP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindWIP.Click
        Dim frmFindWIP As New Core_Module.frmCariMasterProduct
        Try
            frmFindWIP.txtCari.Text = Me.txtWIP.Text
            frmFindWIP.ShowDialog()
            If frmFindWIP.PrimaryProduct <> -1 Then
                Me.txtWIP.Text = Me.objProduct.GetProductCode(frmFindWIP.PrimaryProduct)
                Me.txtWorkunit.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindWU_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindWU.Click
        Dim frmFindWU As New Core_Module.FrmMasterWorkUnit_Cari
        Try
            frmFindWU.txtCari.Text = txtWorkunit.Text
            frmFindWU.ShowDialog()
            If frmFindWU.PrimaryWorkUnit <> -1 Then
                Me.txtWorkunit.Text = Me.objWU.GetCode(frmFindWU.PrimaryWorkUnit)
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnGoto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim RowDgv As DataGridViewRow
        Try
            If Me.DgvItem.RowCount <= 0 Then
                MsgBox("No result found.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.DgvItem.SelectedRows.Count <= 0 Then
                MsgBox("No result selected.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            RowDgv = Me.DgvItem.Rows(Me.DgvItem.SelectedRows(0).Index)
            With RowDgv
                PrimaryMO = .Cells(Me.colPrimaryMO.Index).Value
                IdentifierJO = .Cells(Me.colIdentifierJO.Index).Value
                IdentifierWO = .Cells(Me.colIdentifierWO.Index).Value
                PrimaryProcess = .Cells(Me.colPrimaryProcess.Index).Value
                PrimaryWorkunit = .Cells(Me.colPrimaryWorkuunit.Index).Value
            End With
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim RowDgv As DataGridViewRow
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.TBLSchBar.DefaultView.RowFilter = "[is header] = 0 AND " & GetFilter()  'Me.objCD.FormatSpecialChar_DataFilter()
            Me.TBLSchBar.DefaultView.Sort = "[mo no], [jo no], [plan start time]"
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvItem.Rows.Clear()
            For Each DRV As Data.DataRowView In Me.TBLSchBar.DefaultView
                Me.rowSchBar = DRV.Row
                With Me.rowSchBar
                    RowDgv = Me.DgvItem.Rows(Me.DgvItem.Rows.Add())
                    RowDgv.Cells(Me.colNo.Index).Value = Me.DgvItem.Rows.Count
                    RowDgv.Cells(Me.colMO.Index).Value = .mo_no
                    RowDgv.Cells(Me.colJO.Index).Value = .jo_no
                    RowDgv.Cells(Me.ColFinishGood.Index).Value = Me.rowSchBar("Finish Good")
                    RowDgv.Cells(Me.colProcess.Index).Value = .process_code
                    RowDgv.Cells(Me.colWO.Index).Value = .wo_no
                    RowDgv.Cells(Me.colWIP.Index).Value = .product_code
                    RowDgv.Cells(Me.colWorkunit.Index).Value = .work_unit_code
                    RowDgv.Cells(Me.colPlanStartDate.Index).Value = Format(.plan_start_time, "dd MMM yyyy")
                    RowDgv.Cells(Me.colPlanStartTime.Index).Value = Format(.plan_start_time, "HH:mm")
                    RowDgv.Cells(Me.colPlanEndDate.Index).Value = Format(.plan_end_time, "dd MMM yyyy")
                    RowDgv.Cells(Me.colPlanEndTime.Index).Value = Format(.plan_end_time, "HH:mm")
                    RowDgv.Cells(Me.colPrimaryMO.Index).Value = .primary_mo
                    RowDgv.Cells(Me.colIdentifierJO.Index).Value = .identifier_jo
                    RowDgv.Cells(Me.colIdentifierWO.Index).Value = .identifier_wo
                    RowDgv.Cells(Me.colPrimaryProcess.Index).Value = .primary_process
                    RowDgv.Cells(Me.colPrimaryWorkuunit.Index).Value = .primary_work_unit
                End With
            Next
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Function GetFilter() As String
        Try
            If txtMO.Text = "" Then
                GetFilter = "[mo no] = [mo no]"
            Else
                GetFilter = "[mo no] LIKE '" & Me.objCD.FormatSpecialChar_DataFilter(txtMO.Text) & "%'"
            End If
            If txtJO.Text = "" Then
                GetFilter += " AND [jo no] = [jo no]"
            Else
                GetFilter += " AND [jo no] LIKE '" & Me.objCD.FormatSpecialChar_DataFilter(txtJO.Text) & "%'"
            End If
            'If Me.GroupBy = FrmTrans_ProductionSchedule.enumGroup.BY_JO Then
            If txtFinishGood.Text = "" Then
                GetFilter += " AND [finish good] = [finish good]"
            Else
                GetFilter += " AND [finish good] LIKE '" & Me.objCD.FormatSpecialChar_DataFilter(txtFinishGood.Text) & "%'"
            End If
            'End If
            If txtProcess.Text = "" Then
                GetFilter += " AND [process code] = [process code]"
            Else
                GetFilter += " AND [process code] LIKE '" & Me.objCD.FormatSpecialChar_DataFilter(txtProcess.Text) & "%'"
            End If
            If txtWO.Text = "" Then
                GetFilter += " AND [wo no] = [wo no]"
            Else
                GetFilter += " AND [wo no] LIKE '" & Me.objCD.FormatSpecialChar_DataFilter(txtWO.Text) & "%'"
            End If
            If txtWIP.Text = "" Then
                GetFilter += " AND [product code] = [product code]"
            Else
                GetFilter += " AND [product code] LIKE '" & Me.objCD.FormatSpecialChar_DataFilter(txtWIP.Text) & "%'"
            End If
            If txtWorkunit.Text = "" Then
                GetFilter += " AND [work unit code] = [work unit code]"
            Else
                GetFilter += " AND [work unit code] LIKE '" & Me.objCD.FormatSpecialChar_DataFilter(txtWorkunit.Text) & "%'"
            End If
            Return GetFilter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellContentClick
        Dim RowDgv As DataGridViewRow
        Try
            If e.ColumnIndex = Me.colGoto.Index Then
                RowDgv = Me.DgvItem.Rows(Me.DgvItem.SelectedRows(0).Index)
                With RowDgv
                    PrimaryMO = .Cells(Me.colPrimaryMO.Index).Value
                    IdentifierJO = .Cells(Me.colIdentifierJO.Index).Value
                    IdentifierWO = .Cells(Me.colIdentifierWO.Index).Value
                    PrimaryProcess = .Cells(Me.colPrimaryProcess.Index).Value
                    PrimaryWorkunit = .Cells(Me.colPrimaryWorkuunit.Index).Value
                End With
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.Sorted
        Dim objCd As New Core_Module.ClsCekData
        Try
            objCd.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class