''rosadin-r15
Option Explicit On
Option Strict Off
Public Class FrmTransMO_WorkShift_Allocate
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll()

    Public IsEdited As Boolean = False ''rosadin-r15
    Dim ObjSch As New ClsTrans_Schedule ''rosadin-r15

    Dim ObjWU As New Core_Module.ClsMasterWorkUnit
    Dim ObjWUG As New Core_Module.ClsMasterWorkUnitGroup
    Dim ObjWS As New Core_Module.ClsMasterWorkShift
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjTM As New ClsTransMO_WorkShift_Allocate
    Dim buffer As String = ""
    Dim buffer_from As String = ""
    Dim buffer_to As String = ""
    Dim EraseRowIndex As Integer = 0
    Dim EraseRowIndex_WUG As Integer = 0
    Dim IrowIdx As Integer = 0

    Dim tblmaster_workunit As New DseTransMO_WorkShift_Allocate1.tblmaster_workunitDataTable
    Dim rowmaster_workunit As DseTransMO_WorkShift_Allocate1.tblmaster_workunitRow

    Dim tblmaster_workshift As New DseTransMO_WorkShift_Allocate1.tblmaster_workshiftDataTable
    Dim rowmaster_workshift As DseTransMO_WorkShift_Allocate1.tblmaster_workshiftRow

    Dim tbltrans_mo_workshift_allocate As New DseTransMO_WorkShift_Allocate1.tbltrans_mo_workshift_allocateDataTable
    Dim rowtrans_mo_workshift_allocate As DseTransMO_WorkShift_Allocate1.tbltrans_mo_workshift_allocateRow

    Dim tblConverter As New DseTransMO_WorkShift_Allocate1.tbl_converterDataTable
    Dim rowConverter As DseTransMO_WorkShift_Allocate1.tbl_converterRow
    Dim rowConverter_Before As DseTransMO_WorkShift_Allocate1.tbl_converterRow
    Dim DTABundling As New DseTransMO_WorkShift_Allocate2TableAdapters.tbltransaksi_stockbundlingTableAdapter
    Dim DTASch As New DseTransMO_WorkShift_Allocate2TableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
    Dim DTA_OT As New DseTransMO_WorkShift_Allocate2TableAdapters.tbltrans_overtime_itemTableAdapter

    Private Sub FrmTransMO_WorkShift_Allocate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ''ok-rosadin-r5b
            Me.SplitContainer1.TabStop = False
            Me.SplitContainer2.TabStop = False
            Me.dgvWU.AllowUserToResizeColumns = False
            Me.dgvWU.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            Me.dgvWUG.AllowUserToResizeColumns = False
            Me.dgvWUG.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

            Me.dgvWU.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvWUG.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.LoadWorkShiftColumn()

            Me.DTABundling.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTASch.Connection.ConnectionString = Me.DTABundling.Connection.ConnectionString
            Me.DTA_OT.Connection.ConnectionString = Me.DTABundling.Connection.ConnectionString

            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Please make sure start date <= end date.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.LoadAllocate()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub LoadWorkShiftColumn()
        Dim dtamaster_workshift As New DseTransMO_WorkShift_Allocate1TableAdapters.tblmaster_workshiftTableAdapter
        Try
            Me.Cursor = Cursors.WaitCursor
            dtamaster_workshift.Connection.ConnectionString = Core_Module.GetConstring
            dtamaster_workshift.Fill(Me.tblmaster_workshift, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production)
            For Each Me.rowmaster_workshift In Me.tblmaster_workshift.Rows
                Dim column As New DataGridViewCheckBoxColumn
                column.HeaderText = Me.rowmaster_workshift.code
                column.Name = "Col" & Me.rowmaster_workshift.primary
                Me.dgvItem.Columns.Add(column)
            Next
            Me.tblmaster_workshift.DefaultView.RowFilter = ""
            Me.tblmaster_workshift.DefaultView.Sort = "code"
            Me.cbo_shift.DataSource = Me.tblmaster_workshift.DefaultView
            Me.cbo_shift.ValueMember = "primary"
            Me.cbo_shift.DisplayMember = "code"
            Me.cbo_shift.SelectedIndex = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub LoadAllocate()
        Dim dtamaster_workunit As New DseTransMO_WorkShift_Allocate1TableAdapters.tblmaster_workunitTableAdapter
        Dim dtatrans_mo_workshift_allocate As New DseTransMO_WorkShift_Allocate1TableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim currentDate As Date
        Dim currentIndex As Integer = 0

        Try
            Me.Cursor = Cursors.WaitCursor
            Me.dgvItem.Rows.Clear()
            Me.tblConverter.Rows.Clear()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dtamaster_workunit.Connection.ConnectionString = Core_Module.GetConstring
            dtatrans_mo_workshift_allocate.Connection.ConnectionString = Core_Module.GetConstring

            If (Me.rboWU.Checked And Me.dgvWU.Rows.Count <= 1) Or _
                (Me.rboWUG.Checked And Me.ClbWU.CheckedItems.Count = 0) Then
                ''fill all
                dtamaster_workunit.Fill(Me.tblmaster_workunit)
            ElseIf Me.rboWU.Checked And Me.dgvWU.Rows.Count > 1 Then
                ''fill from datagridview
                Me.tblmaster_workunit.Rows.Clear()
                For Each dri As DataGridViewRow In Me.dgvWU.Rows
                    If IsNothing(dri.Cells(Me.colWU_Primary.Index).Value) = False Then
                        If String.IsNullOrEmpty(dri.Cells(Me.colWU_Primary.Index).Value.ToString) = False Then
                            If dri.Cells(Me.colWU_Primary.Index).Value.ToString <> "" And dri.Cells(Me.colWU_Primary.Index).Value.ToString <> "-1" Then
                                Me.rowmaster_workunit = Me.tblmaster_workunit.Newtblmaster_workunitRow
                                Me.rowmaster_workunit.primary = dri.Cells(Me.colWU_Primary.Index).Value
                                Me.rowmaster_workunit.work_unit_code = dri.Cells(Me.colWU_Code.Index).Value
                                Me.rowmaster_workunit.description = dri.Cells(Me.colWU_Desc.Index).Value
                                Me.tblmaster_workunit.Rows.Add(Me.rowmaster_workunit)
                            End If
                        End If
                    End If
                Next
            ElseIf Me.rboWUG.Checked And Me.ClbWU.CheckedItems.Count > 0 Then
                ''fill from checklistbox checked
                Me.tblmaster_workunit.Rows.Clear()
                For c As Integer = 0 To Me.ClbWU.Items.Count - 1
                    If Me.ClbWU.GetItemCheckState(c) = CheckState.Checked Then
                        Me.rowmaster_workunit = Me.tblmaster_workunit.Newtblmaster_workunitRow
                        Me.rowmaster_workunit.primary = Me.ObjWU.GetPrimary(CStr(Me.ClbWU.Items(c)))
                        Me.rowmaster_workunit.work_unit_code = CStr(Me.ClbWU.Items(c))
                        Me.rowmaster_workunit.description = Me.ObjWU.GetDesc(Me.rowmaster_workunit.primary)
                        Me.tblmaster_workunit.Rows.Add(Me.rowmaster_workunit)
                    End If
                Next
            End If
            dtatrans_mo_workshift_allocate.Fill(Me.tbltrans_mo_workshift_allocate, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)

            ''hilangkan allocate yg workunitnya tidak ada dalam list filter
            Me.SortTblUnit(tblmaster_workunit, "[primary]")
            Me.SortTblAllocate(tbltrans_mo_workshift_allocate, "[primary work unit]")
            Dim indexLeft As Integer = 0
            Dim indexRight As Integer = 0
            While indexRight <= tbltrans_mo_workshift_allocate.Rows.Count - 1
                If indexLeft <= tblmaster_workunit.Rows.Count - 1 Then
                    Me.rowmaster_workunit = tblmaster_workunit.Rows(indexLeft)
                    Me.rowtrans_mo_workshift_allocate = Me.tbltrans_mo_workshift_allocate.Rows(indexRight)
                    If Me.rowmaster_workunit.primary > Me.rowtrans_mo_workshift_allocate.primary_work_unit Then
                        Me.tbltrans_mo_workshift_allocate.Rows.RemoveAt(indexRight)
                    ElseIf Me.rowmaster_workunit.primary = Me.rowtrans_mo_workshift_allocate.primary_work_unit Then
                        indexRight += 1
                    Else
                        indexLeft += 1
                    End If
                Else
                    Me.tbltrans_mo_workshift_allocate.Rows.RemoveAt(indexRight)
                End If
            End While

            Me.SortTblAllocate(tbltrans_mo_workshift_allocate, "[work unit code], [date], [primary workshift]")
            Me.SortTblUnit(tblmaster_workunit, "[work unit code]")

            ''covert workshift
            For Each Me.rowmaster_workunit In Me.tblmaster_workunit
                For i As Integer = 0 To DateDiff(DateInterval.Day, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                    currentDate = Me.dtpFrom.Value.Date.AddDays(i)
                    Me.rowConverter = Me.tblConverter.Newtbl_converterRow
                    Me.rowConverter.primary_work_unit = Me.rowmaster_workunit.primary
                    Me.rowConverter.work_unit_code = Me.rowmaster_workunit.work_unit_code
                    Me.rowConverter.work_unit_desc = Me.rowmaster_workunit.description
                    Me.rowConverter._date = currentDate
                    Me.rowConverter.primary_workshift_combine = ""
                    Me.tblConverter.Rows.Add(Me.rowConverter)
                    If currentIndex <= Me.tbltrans_mo_workshift_allocate.Count - 1 Then
                        Me.rowtrans_mo_workshift_allocate = Me.tbltrans_mo_workshift_allocate.Rows(currentIndex)
                        If Me.rowtrans_mo_workshift_allocate.primary_work_unit = Me.rowmaster_workunit.primary And Me.rowtrans_mo_workshift_allocate._date = currentDate Then
                            Me.rowConverter.primary_workshift_combine = Me.rowtrans_mo_workshift_allocate.primary_workshift.ToString
                            currentIndex += 1
                            If currentIndex <= Me.tbltrans_mo_workshift_allocate.Count - 1 Then
                                Me.rowtrans_mo_workshift_allocate = Me.tbltrans_mo_workshift_allocate.Rows(currentIndex)
                                While Me.rowtrans_mo_workshift_allocate.primary_work_unit = Me.rowmaster_workunit.primary And Me.rowtrans_mo_workshift_allocate._date = currentDate
                                    Me.rowConverter.primary_workshift_combine = Me.rowConverter.primary_workshift_combine + "-" + Me.rowtrans_mo_workshift_allocate.primary_workshift.ToString
                                    currentIndex += 1
                                    If currentIndex > Me.tbltrans_mo_workshift_allocate.Count - 1 Then
                                        Exit While
                                    End If
                                    Me.rowtrans_mo_workshift_allocate = Me.tbltrans_mo_workshift_allocate.Rows(currentIndex)
                                End While
                            End If
                        Else
                            Me.rowConverter.primary_workshift_combine = "-1"
                        End If
                    End If
                Next
            Next

            ''fill to datagridview
            Dim dr As DataGridViewRow
            For Each Me.rowConverter In Me.tblConverter.Rows
                If Me.tblConverter.Rows.IndexOf(Me.rowConverter) = 0 Then
                    dr = Me.dgvItem.Rows(Me.dgvItem.Rows.Add())
                    dr.Cells(Me.colNo.Index).Value = Me.dgvItem.Rows.Count.ToString
                    dr.Cells(Me.colWorkUnitCode.Index).Value = Me.rowConverter.work_unit_code
                    dr.Cells(Me.colWorkUnitDesc.Index).Value = Me.rowConverter.work_unit_desc
                    dr.Cells(Me.colTglDari.Index).Value = Format(Me.rowConverter._date, "dd MMM yyyy")
                    dr.Cells(Me.colTglSampai.Index).Value = Format(Me.rowConverter._date, "dd MMM yyyy")
                    dr.Cells(Me.colPrimaryWU.Index).Value = Me.rowConverter.primary_work_unit
                    Me.CheckWorkshift(Me.rowConverter.primary_workshift_combine, dr.Index)
                Else
                    If Me.rowConverter.primary_work_unit = Me.rowConverter_Before.primary_work_unit And _
                            Me.rowConverter.primary_workshift_combine = Me.rowConverter_Before.primary_workshift_combine Then
                        dr.Cells(Me.colTglSampai.Index).Value = Format(Me.rowConverter._date, "dd MMM yyyy")
                    Else
                        dr = Me.dgvItem.Rows(Me.dgvItem.Rows.Add())
                        dr.Cells(Me.colNo.Index).Value = Me.dgvItem.Rows.Count.ToString
                        dr.Cells(Me.colWorkUnitCode.Index).Value = Me.rowConverter.work_unit_code
                        dr.Cells(Me.colWorkUnitDesc.Index).Value = Me.rowConverter.work_unit_desc
                        dr.Cells(Me.colTglDari.Index).Value = Format(Me.rowConverter._date, "dd MMM yyyy")
                        dr.Cells(Me.colTglSampai.Index).Value = Format(Me.rowConverter._date, "dd MMM yyyy")
                        dr.Cells(Me.colPrimaryWU.Index).Value = Me.rowConverter.primary_work_unit
                        Me.CheckWorkshift(Me.rowConverter.primary_workshift_combine, dr.Index)
                    End If
                End If
                Me.rowConverter_Before = Me.rowConverter
            Next

            Me.TranslateDate()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub SortTblUnit(ByRef vTbl As DseTransMO_WorkShift_Allocate1.tblmaster_workunitDataTable, ByVal vSort As String)
        Dim vRow As DseTransMO_WorkShift_Allocate1.tblmaster_workunitRow
        Dim tblDummy As New DseTransMO_WorkShift_Allocate1.tblmaster_workunitDataTable
        Dim rowDummy As DseTransMO_WorkShift_Allocate1.tblmaster_workunitRow
        Try
            vTbl.DefaultView.RowFilter = ""
            vTbl.DefaultView.Sort = vSort
            For i As Integer = 0 To vTbl.DefaultView.Count - 1
                vRow = vTbl.DefaultView(i).Row
                rowDummy = tblDummy.NewRow
                For i2 As Integer = 0 To vTbl.Columns.Count - 1
                    rowDummy.Item(i2) = vRow.Item(i2)
                Next
                tblDummy.Rows.Add(rowDummy)
            Next
            vTbl.Rows.Clear()
            For Each rowDummy In tblDummy.Rows
                vRow = vTbl.NewRow
                For i2 As Integer = 0 To tblDummy.Columns.Count - 1
                    vRow.Item(i2) = rowDummy.Item(i2)
                Next
                vTbl.Rows.Add(vRow)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub SortTblAllocate(ByRef vTbl As DseTransMO_WorkShift_Allocate1.tbltrans_mo_workshift_allocateDataTable, ByVal vSort As String)
        Dim vRow As DseTransMO_WorkShift_Allocate1.tbltrans_mo_workshift_allocateRow
        Dim tblDummy As New DseTransMO_WorkShift_Allocate1.tbltrans_mo_workshift_allocateDataTable
        Dim rowDummy As DseTransMO_WorkShift_Allocate1.tbltrans_mo_workshift_allocateRow
        Try
            vTbl.DefaultView.RowFilter = ""
            vTbl.DefaultView.Sort = vSort
            For i As Integer = 0 To vTbl.DefaultView.Count - 1
                vRow = vTbl.DefaultView(i).Row
                rowDummy = tblDummy.NewRow
                For i2 As Integer = 0 To vTbl.Columns.Count - 1
                    rowDummy.Item(i2) = vRow.Item(i2)
                Next
                tblDummy.Rows.Add(rowDummy)
            Next
            vTbl.Rows.Clear()
            For Each rowDummy In tblDummy.Rows
                vRow = vTbl.NewRow
                For i2 As Integer = 0 To tblDummy.Columns.Count - 1
                    vRow.Item(i2) = rowDummy.Item(i2)
                Next
                vTbl.Rows.Add(vRow)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub CheckWorkshift(ByVal vStringCombine As String, ByVal vRowIndex As Integer)
        Dim array() As String
        Try
            If String.IsNullOrEmpty(vStringCombine) Then
                Exit Sub
            End If
            If vStringCombine.Trim = "" Or vStringCombine.Trim = "-1" Then
                Exit Sub
            End If
            array = vStringCombine.Split("-")
            For Each s As String In array
                If Me.dgvItem.Columns.Contains("Col" & s) Then
                    Me.dgvItem.Item("Col" & s, vRowIndex).Value = True
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Sub TranslateDate()
        Try
            Me.Cursor = Cursors.WaitCursor
            For Each r As DataGridViewRow In Me.dgvItem.Rows
                If r.Cells(Me.colTglDari.Index).Value <> r.Cells(Me.colTglSampai.Index).Value Then
                    r.Cells(Me.colDate.Index).Value = r.Cells(Me.colTglDari.Index).Value & " - " & r.Cells(Me.colTglSampai.Index).Value
                Else
                    r.Cells(Me.colDate.Index).Value = r.Cells(Me.colTglDari.Index).Value
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub TambahWorkUnit(ByVal vPrimaryGroup As Double)
        Dim dtaWorkUnit As New DseTransMO_WorkShift_Allocate1TableAdapters.tblmaster_workunitTableAdapter
        Try
            dtaWorkUnit.Connection.ConnectionString = Core_Module.GetConstring
            If Me.bufferPrimary = vPrimaryGroup Then
                Exit Sub
            End If
            Me.HapusWorkUnit(Me.bufferPrimary)
            dtaWorkUnit.FillBy_PrimaryGroup(Me.tblmaster_workunit, vPrimaryGroup)
            For Each Me.rowmaster_workunit In Me.tblmaster_workunit.Rows
                Me.ClbWU.Items.Add(Me.rowmaster_workunit.work_unit_code, False)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub HapusWorkUnit(ByVal vPrimaryGroup As Double)
        Dim dtaWorkUnit As New DseTransMO_WorkShift_Allocate1TableAdapters.tblmaster_workunitTableAdapter
        Try
            dtaWorkUnit.Connection.ConnectionString = Core_Module.GetConstring
            dtaWorkUnit.FillBy_PrimaryGroup(Me.tblmaster_workunit, vPrimaryGroup)
            For i As Integer = Me.ClbWU.Items.Count - 1 To 0 Step -1
                tblmaster_workunit.DefaultView.RowFilter = "[work unit code] = '" & Me.ClbWU.Items(i).ToString & "'"
                If tblmaster_workunit.DefaultView.Count > 0 Then
                    Me.ClbWU.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboWU_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rboWU.CheckedChanged
        Try
            If Me.rboWU.Checked Then
                Me.dgvWU.ReadOnly = False
                Me.dgvWUG.ReadOnly = True
                Me.ClbWU.Enabled = False
                Me.dgvWUG.Rows.Clear()
                Me.ClbWU.Items.Clear()
                Me.dgvWUG.AllowUserToAddRows = False
                Me.dgvWU.AllowUserToAddRows = True
                Me.llbCheckAll.Enabled = False
                Me.llbUncheckAll.Enabled = False
            Else
                Me.dgvWU.ReadOnly = True
                Me.dgvWUG.ReadOnly = False
                Me.ClbWU.Enabled = True
                Me.dgvWU.Rows.Clear()
                Me.dgvWUG.AllowUserToAddRows = True
                Me.dgvWU.AllowUserToAddRows = False
                Me.llbCheckAll.Enabled = True
                Me.llbUncheckAll.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub dgvWU_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvWU.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colWU_Code.Index Then
                Me.buffer = Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWU_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWU.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvWU.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim PrimaryWU As Double = -1
        Dim RowIdx As Integer = e.RowIndex

        Try
            With row
                If IsNothing(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.colWU_Code.Index Then
                        Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.ObjWU.ApakahSudahAda(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value, True) = True Then
                            PrimaryWU = Me.ObjWU.GetPrimary(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItem(PrimaryWU, e.RowIndex)
                            If index < 0 Then
                                Me.dgvWU.Item(Me.colWU_Primary.Index, e.RowIndex).Value = PrimaryWU.ToString
                                Me.dgvWU.Item(Me.colWU_Code.Index, e.RowIndex).Value = Me.ObjWU.GetCode(PrimaryWU)
                                Me.dgvWU.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.ObjWU.GetDesc(PrimaryWU)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvWU.Item(Me.colWU_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvWU)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterWorkUnit_Cari
                            ij.txtCari.Text = Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.Status = "Only Active"
                            ij.ShowDialog()
                            If ij.PrimaryWorkUnit <> -1 Then
                                index = Me.GetIndexItem(ij.PrimaryWorkUnit, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvWU.Item(Me.colWU_Primary.Index, e.RowIndex).Value = ij.PrimaryWorkUnit
                                    Me.dgvWU.Item(Me.colWU_Code.Index, e.RowIndex).Value = Me.ObjWU.GetCode(ij.PrimaryWorkUnit)
                                    Me.dgvWU.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.ObjWU.GetDesc(ij.PrimaryWorkUnit)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvWU)
                                        Exit Sub
                                    Else
                                        Me.dgvWU.Item(Me.colWU_Primary.Index, e.RowIndex).Value = ij.PrimaryWorkUnit
                                        Me.dgvWU.Item(Me.colWU_Code.Index, e.RowIndex).Value = Me.ObjWU.GetCode(ij.PrimaryWorkUnit)
                                        Me.dgvWU.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.ObjWU.GetDesc(ij.PrimaryWorkUnit)
                                    End If
                                End If
                            ElseIf ij.PrimaryWorkUnit = -1 Then
                                If Me.dgvWU.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvWU)
                                    Exit Sub
                                End If
                                If Me.dgvWU.Item(Me.colWU_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvWU)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvWU)
                    End If
                Else
                    Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWU_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvWU.RowsAdded
        Try
            Me.ObjCD.SetNumber(Me.dgvWU, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWU_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvWU.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvWU, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWU_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvWU.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvWU)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Me.ObjCD.UpdateThreading()
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.dgvWU.Rows.RemoveAt(EraseRowIndex)
            Me.ObjCD.SetNumber(Me.dgvWU, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vPrimary As Double, _
      ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvWU.Rows.Count - 2
                RowDGV = Me.dgvWU.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.colWU_Primary.Index).Value = vPrimary Then
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

    Dim bufferPrimary As Double = -1
    Private Sub dgvWUGG_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvWUG.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colWUG_Code.Index Then
                Me.buffer = Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value
                Me.bufferPrimary = Me.dgvWUG(Me.colWUG_Primary.Index, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWUG_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWUG.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvWUG.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim PrimaryWUG As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim dV As New DataGridViewRow
        Try
            With row
                If IsNothing(Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.colWUG_Code.Index Then
                        Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.ObjWUG.Apakahsudahada(Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryWUG = Me.ObjWUG.GetPrimary(Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItem_WUG(PrimaryWUG, e.RowIndex)
                            If index < 0 Then
                                Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value = PrimaryWUG
                                Me.dgvWUG.Item(Me.colWUG_Code.Index, e.RowIndex).Value = Me.ObjWUG.GetCode(PrimaryWUG)
                                Me.dgvWUG.Item(Me.colWUG_Desc.Index, e.RowIndex).Value = Me.ObjWUG.GetDesc(PrimaryWUG)
                                .Selected = True
                                Me.TambahWorkUnit(PrimaryWUG)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvWUG.Item(Me.colWUG_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex_WUG = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvWUG)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterWorkUnitGroup_Cari
                            ij.txtCari.Text = Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryWorkUnitGroup <> -1 Then
                                index = Me.GetIndexItem_WUG(ij.PrimaryWorkUnitGroup, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value = ij.PrimaryWorkUnitGroup
                                    Me.dgvWUG.Item(Me.colWUG_Code.Index, e.RowIndex).Value = Me.ObjWUG.GetCode(ij.PrimaryWorkUnitGroup)
                                    Me.dgvWUG.Item(Me.colWUG_Desc.Index, e.RowIndex).Value = Me.ObjWUG.GetDesc(ij.PrimaryWorkUnitGroup)
                                    Me.dgvWUG.Rows(dgvWUG.Rows.Count - 1).Selected = True
                                    Me.TambahWorkUnit(ij.PrimaryWorkUnitGroup)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex_WUG = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvWUG)
                                        Exit Sub
                                    Else
                                        Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value = ij.PrimaryWorkUnitGroup
                                        Me.dgvWUG.Item(Me.colWUG_Code.Index, e.RowIndex).Value = Me.ObjWUG.GetCode(ij.PrimaryWorkUnitGroup)
                                        Me.dgvWUG.Item(Me.colWUG_Desc.Index, e.RowIndex).Value = Me.ObjWUG.GetDesc(ij.PrimaryWorkUnitGroup)
                                        Me.dgvWUG.Rows(dgvWUG.Rows.Count - 1).Selected = True
                                        Me.TambahWorkUnit(ij.PrimaryWorkUnitGroup)
                                    End If
                                End If
                            ElseIf ij.PrimaryWorkUnitGroup = -1 Then
                                If Me.dgvWUG.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex_WUG = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvWUG)
                                    Exit Sub
                                End If
                                If Me.dgvWUG.Item(Me.colWUG_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex_WUG = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvWUG)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvWUG)
                    End If
                Else
                    Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
  
    Private Sub dgvWUG_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvWUG.RowsAdded
        Try
            Me.ObjCD.SetNumber(Me.dgvWUG, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWUG_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvWUG.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvWUG, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWUG_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvWUG.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvWUG, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Me.ObjCD.UpdateThreading()
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.dgvWUG.Rows.RemoveAt(EraseRowIndex_WUG)
            Me.ObjCD.SetNumber(Me.dgvWUG, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem_WUG(ByVal vPrimary As Double, _
      ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvWUG.Rows.Count - 2
                RowDGV = Me.dgvWUG.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.colWUG_Primary.Index).Value = vPrimary Then
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

    Private Sub dgvWUG_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvWUG.UserDeletingRow
        Try
            Me.HapusWorkUnit(Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.Row.Index).Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Try
            For i As Integer = 0 To Me.ClbWU.Items.Count - 1
                Me.ClbWU.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Try
            For i As Integer = 0 To Me.ClbWU.Items.Count - 1
                Me.ClbWU.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUnAllocate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUnAllocate.LinkClicked
        Dim FirmPlanDate As Date = Today
        Dim dummyWU As String = ""
        Dim rowDGV As DataGridViewRow
        Dim idx As Integer = 0
        Dim PrimaryWS As Double = -1
        Dim ListPrimaryWU As New List(Of Double)
        Dim DTAAllocate As New DseTransMO_WorkShift_Allocate2TableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim IdxAcuan As Integer = 0

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            PrimaryWS = Me.ObjWS.GetPrimary(Me.cbo_shift.Text, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production)
            DTAAllocate.Connection.ConnectionString = Core_Module.GetConstring

            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Please make sure start date <= end date.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.cbo_shift.Text = "" Then
                MsgBox("Please specify workshift.", MsgBoxStyle.Information, "Information")
                Me.cbo_shift.Focus()
                Exit Sub
            End If

            '0. pastikan ada work unit yang dimaksud
            If Me.dgvItem.Rows.Count = 0 Then
                MsgBox("Please specify at least  work unit to allocate", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '1. tidak boleh sebelum firm plan
            FirmPlanDate = Me.ObjTM.GetFirmPlan
            If Me.dtpFrom.Value.Date <= FirmPlanDate.Date Then
                MsgBox("Please specify date after firm plan: " + Format(FirmPlanDate, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'tandai kolom workshift yang akan di un allocate
            For ij As Integer = 7 To Me.dgvItem.Columns.Count - 1
                If Me.ObjWS.GetPrimary(Me.dgvItem.Columns(ij).HeaderText, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production) = PrimaryWS Then
                    IdxAcuan = ij
                    Exit For
                End If
            Next

            '2. tidak boleh kalo sudah ada bundling
            dummyWU = ""
            For Each rowDGV In Me.dgvItem.Rows
                If CBool(rowDGV.Cells(IdxAcuan).Value) = Boolean.TrueString Then
                    If ListPrimaryWU.Contains(rowDGV.Cells(Me.colPrimaryWU.Index).Value) = False Then
                        ListPrimaryWU.Add(rowDGV.Cells(Me.colPrimaryWU.Index).Value)
                        If Me.ObjTM.GetJmlBundling(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, rowDGV.Cells(Me.colPrimaryWU.Index).Value, Me.DTABundling) > 0 Then
                            idx += 1
                            If dummyWU = "" Then
                                dummyWU = idx.ToString + ". " + rowDGV.Cells(Me.colWorkUnitCode.Index).Value.ToString
                            Else
                                dummyWU = dummyWU + vbCrLf + idx.ToString + ". " + rowDGV.Cells(Me.colWorkUnitCode.Index).Value.ToString
                            End If
                        End If
                    End If
                End If
            Next
            If dummyWU <> "" Then
                MsgBox("There is production log on period: " + Format(Me.dtpFrom.Value, "dd MMM yyyy") + " to " + _
                Format(Me.dtpTo.Value, "dd MMM yyyy") + " for the following work unit:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '2.b. tidak boleh kalo sudah ada overtime
            dummyWU = ""
            ListPrimaryWU.Clear()
            For Each rowDGV In Me.dgvItem.Rows
                If CBool(rowDGV.Cells(IdxAcuan).Value) = Boolean.TrueString Then
                    If ListPrimaryWU.Contains(rowDGV.Cells(Me.colPrimaryWU.Index).Value) = False Then
                        ListPrimaryWU.Add(rowDGV.Cells(Me.colPrimaryWU.Index).Value)
                        If Me.ObjTM.GetJmlOT(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, rowDGV.Cells(Me.colPrimaryWU.Index).Value, Me.DTA_OT) > 0 Then
                            idx += 1
                            If dummyWU = "" Then
                                dummyWU = idx.ToString + ". " + rowDGV.Cells(Me.colWorkUnitCode.Index).Value.ToString
                            Else
                                dummyWU = dummyWU + vbCrLf + idx.ToString + ". " + rowDGV.Cells(Me.colWorkUnitCode.Index).Value.ToString
                            End If
                        End If
                    End If
                End If
            Next
            If dummyWU <> "" Then
                MsgBox("There is overtime on period: " + Format(Me.dtpFrom.Value, "dd MMM yyyy") + " to " + _
                Format(Me.dtpTo.Value, "dd MMM yyyy") + " for the following work unit:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '3. tidak boleh kalo ada workshift yang tabrakan -> ini tidak perlu, karena unallocate
            dummyWU = ""
            'If Me.ObjTM.ApakahAmanDrJadwalTabrakan(PrimaryWS, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, ListPrimaryWU, dummyWU) = False Then
            '    MsgBox("Conflicting worktime for the following work unit:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
            '    Exit Sub
            'End If

            ''rosadin-r15
            ''4. tidak boleh kalo ada jadwal-sementara
            ''MsgBox("belum :")
            ''MsgBox("belum: harusnya bisa geser schedule aja.")
            'dummyWU = ""
            'idx = 0
            'ListPrimaryWU.Clear()
            'For Each rowDGV In Me.dgvItem.Rows
            '    If CBool(rowDGV.Cells(IdxAcuan).Value) = Boolean.TrueString Then
            '        If ListPrimaryWU.Contains(rowDGV.Cells(Me.colPrimaryWU.Index).Value) = False Then
            '            ListPrimaryWU.Add(rowDGV.Cells(Me.colPrimaryWU.Index).Value)
            '            If Me.ObjTM.GetJmlSch(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, rowDGV.Cells(Me.colPrimaryWU.Index).Value, Me.DTASch) > 0 Then
            '                idx += 1
            '                If dummyWU = "" Then
            '                    dummyWU = idx.ToString + ". " + rowDGV.Cells(Me.colWorkUnitCode.Index).Value.ToString
            '                Else
            '                    dummyWU = dummyWU + vbCrLf + idx.ToString + ". " + rowDGV.Cells(Me.colWorkUnitCode.Index).Value.ToString
            '                End If
            '            End If
            '        End If
            '    End If
            'Next
            'If dummyWU <> "" Then
            '    MsgBox("There is production schedule on period: " + Format(Me.dtpFrom.Value, "dd MMM yyyy") + " to " + _
            '    Format(Me.dtpTo.Value, "dd MMM yyyy") + " for the following work unit:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
            '    Exit Sub
            'End If
            Dim vTBLSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
            Dim vTBLSpaceBeforeDeleteSpace As New DseTrans_Schedule.tbl_sch_barDataTable
            Dim TBLUpdate As New DseTrans_OT_Item.DataTable2DataTable
            Dim objSchOT As New ClsTrans_Schedule_OT
            ObjSch.LoadScheduleBar_WOItemOnly(#1/1/1950#, #12/31/2200#, vTBLSchBar)
            ObjSch.LoadScheduleBar_TanpaScheduledItem(ObjSch.GetFirmPlan, #12/31/2200#, vTBLSpaceBeforeDeleteSpace)
            For Each PrimaryWorkunit As Double In ListPrimaryWU
                objSchOT.RecalculateAfterDeleteAllocation(PrimaryWorkunit, Me.dtpFrom.Value.Date, vTBLSpaceBeforeDeleteSpace, vTBLSchBar, TBLUpdate, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, PrimaryWS)
            Next
            If Not objSchOT.EnoughAllocation(TBLUpdate) Then
                Exit Sub
            End If
            If objSchOT.AnyClosedJO(TBLUpdate) Then
                Exit Sub
            End If

            Dim message As String = ""
            If objSchOT.AnyProducedJO(TBLUpdate, message) Then
                MsgBox("Cannot un-allocate workshift, there are log transaction for the following JO: " & vbNewLine & message & ".", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If MsgBox("Start unallocating workshift?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                Exit Sub
            End If
            objSchOT.UpdateScheduleChange(TBLUpdate)
            Me.IsEdited = True

            For Each PrimaryWU As Double In ListPrimaryWU
                DTAAllocate.DeleteQuery(PrimaryWU, PrimaryWS, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
            Next
            Me.LoadAllocate()
            MsgBox("Data unallocated.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryWU = Nothing
            DTAAllocate.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub llbAllocate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAllocate.LinkClicked
        Dim FirmPlanDate As Date = Today
        Dim dummyWU As String = ""
        Dim rowDGV As DataGridViewRow
        Dim idx As Integer = 0
        Dim PrimaryWS As Double = -1
        Dim ListPrimaryWU As New List(Of Double)
        Dim DTAAllocate As New DseTransMO_WorkShift_Allocate2TableAdapters.tbltrans_mo_workshift_allocateTableAdapter

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTAAllocate.Connection.ConnectionString = Core_Module.GetConstring

            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Please make sure start date <= end date.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.cbo_shift.Text = "" Then
                MsgBox("Please specify workshift.", MsgBoxStyle.Information, "Information")
                Me.cbo_shift.Focus()
                Exit Sub
            End If

            '0. pastikan ada work unit yang dimaksud
            If Me.dgvItem.Rows.Count = 0 Then
                MsgBox("Please specify at least  work unit to allocate", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '1. tidak boleh sebelum firm plan
            FirmPlanDate = Me.ObjTM.GetFirmPlan
            If Me.dtpFrom.Value.Date <= FirmPlanDate.Date Then
                MsgBox("Please specify date after firm plan: " + Format(FirmPlanDate, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '2. tidak boleh kalo sudah ada bundling
            dummyWU = ""
            For Each rowDGV In Me.dgvItem.Rows
                If ListPrimaryWU.Contains(rowDGV.Cells(Me.colPrimaryWU.Index).Value) = False Then
                    ListPrimaryWU.Add(rowDGV.Cells(Me.colPrimaryWU.Index).Value)
                    If Me.ObjTM.GetJmlBundling(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, rowDGV.Cells(Me.colPrimaryWU.Index).Value, Me.DTABundling) > 0 Then
                        idx += 1
                        If dummyWU = "" Then
                            dummyWU = idx.ToString + ". " + rowDGV.Cells(Me.colWorkUnitCode.Index).Value.ToString
                        Else
                            dummyWU = dummyWU + vbCrLf + idx.ToString + ". " + rowDGV.Cells(Me.colWorkUnitCode.Index).Value.ToString
                        End If
                    End If
                End If
            Next
            If dummyWU <> "" Then
                MsgBox("There is production log on period: " + Format(Me.dtpFrom.Value, "dd MMM yyyy") + " to " + _
                Format(Me.dtpTo.Value, "dd MMM yyyy") + " for the following work unit:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '2. tidak boleh kalo sudah ada overtime
            dummyWU = ""
            ListPrimaryWU.Clear()
            For Each rowDGV In Me.dgvItem.Rows
                If ListPrimaryWU.Contains(rowDGV.Cells(Me.colPrimaryWU.Index).Value) = False Then
                    ListPrimaryWU.Add(rowDGV.Cells(Me.colPrimaryWU.Index).Value)
                    If Me.ObjTM.GetJmlOT(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, rowDGV.Cells(Me.colPrimaryWU.Index).Value, Me.DTA_OT) > 0 Then
                        idx += 1
                        If dummyWU = "" Then
                            dummyWU = idx.ToString + ". " + rowDGV.Cells(Me.colWorkUnitCode.Index).Value.ToString
                        Else
                            dummyWU = dummyWU + vbCrLf + idx.ToString + ". " + rowDGV.Cells(Me.colWorkUnitCode.Index).Value.ToString
                        End If
                    End If
                End If
            Next
            If dummyWU <> "" Then
                MsgBox("There is overtime on period: " + Format(Me.dtpFrom.Value, "dd MMM yyyy") + " to " + _
                Format(Me.dtpTo.Value, "dd MMM yyyy") + " for the following work unit:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '3. tidak boleh kalo ada workshift yang tabrakan -> ini tidak perlu, karena unallocate
            dummyWU = ""
            PrimaryWS = Me.ObjWS.GetPrimary(Me.cbo_shift.Text, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production)
            If Me.ObjTM.ApakahAmanDrJadwalTabrakan(PrimaryWS, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, ListPrimaryWU, dummyWU) = False Then
                MsgBox("Conflicting worktime for the following work shift:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            ''rosadin-r15
            ''4. tidak boleh kalo ada jadwal-sementara
            ''MsgBox("belum :")
            ''MsgBox("belum: harusnya bisa geser schedule aja.")
            'dummyWU = ""
            'idx = 0
            'ListPrimaryWU.Clear()
            'For Each rowDGV In Me.dgvItem.Rows
            '    If ListPrimaryWU.Contains(rowDGV.Cells(Me.colPrimaryWU.Index).Value) = False Then
            '        ListPrimaryWU.Add(rowDGV.Cells(Me.colPrimaryWU.Index).Value)
            '        If Me.ObjTM.GetJmlSch(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, rowDGV.Cells(Me.colPrimaryWU.Index).Value, Me.DTASch) > 0 Then
            '            idx += 1
            '            If dummyWU = "" Then
            '                dummyWU = idx.ToString + ". " + rowDGV.Cells(Me.colWorkUnitCode.Index).Value.ToString
            '            Else
            '                dummyWU = dummyWU + vbCrLf + idx.ToString + ". " + rowDGV.Cells(Me.colWorkUnitCode.Index).Value.ToString
            '            End If
            '        End If
            '    End If
            'Next
            'If dummyWU <> "" Then
            '    MsgBox("There is production schedule on period: " + Format(Me.dtpFrom.Value, "dd MMM yyyy") + " to " + _
            '    Format(Me.dtpTo.Value, "dd MMM yyyy") + " for the following work unit:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
            '    Exit Sub
            'End If
            'MsgBox("tanya rosadin") '-> ini kenapa?
            Dim objSchOT As New ClsTrans_Schedule_OT
            Dim vTBLSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
            Dim vTBLSpaceBeforeAllocate As New DseTrans_Schedule.tbl_sch_barDataTable
            'ObjSch.LoadScheduleBar_WOItemOnly(#1/1/1950#, #12/31/2200#, vTBLSchBar)
            'ObjSch.LoadScheduleBar_TanpaScheduledItem(ObjSch.GetFirmPlan, #12/31/2200#, vTBLSpaceBeforeAllocate)
            ObjSch.LoadScheduleBar_WOItemOnly(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, vTBLSchBar) 'modif ama iwan: 2018-06-29
            ObjSch.LoadScheduleBar_TanpaScheduledItem(ObjSch.GetFirmPlan, Me.dtpTo.Value.Date, vTBLSpaceBeforeAllocate)
            'If objSchOT.AnyClosedJO(vTBLSchBar, Me.dtpFrom.Value.Date, ListPrimaryWU) Then
            '    Exit Sub
            'End If

            Dim message As String = ""
            If objSchOT.AnyProducedJO(vTBLSchBar, Me.dtpFrom.Value.Date, ListPrimaryWU, message, Me.dtpTo.Value.Date) Then
                MsgBox("Cannot allocate workshift, there are log transaction for the following JO: " & vbNewLine & message & ".", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'mulai unallocate
            If MsgBox("Start allocating workshift?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                Exit Sub
            End If
            For Each PrimaryWU As Double In ListPrimaryWU
                DTAAllocate.DeleteQuery(PrimaryWU, PrimaryWS, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
                Me.Simpan_UsingScript(PrimaryWS, PrimaryWU, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
            Next
            Me.LoadAllocate()
            MsgBox("Data allocated.", MsgBoxStyle.Information, "Information")

            For Each PrimaryWU As Double In ListPrimaryWU
                If Me.ObjSch.ApakahAda_PlanSchedule(PrimaryWU, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date) = True Then
                    objSchOT.RecalculateAfterAddSpace(PrimaryWU, Me.dtpFrom.Value.Date, vTBLSpaceBeforeAllocate, vTBLSchBar)
                End If
            Next
            Me.IsEdited = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryWU = Nothing
            DTAAllocate.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Function Simpan_UsingScript(ByVal vPrimaryWS As Double, ByVal vPrimaryWU As Double, ByVal vStartDate As Date, ByVal vEndDate As Date) As Boolean
        Dim conStr As String = My.Settings.Core_DbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        'Dim BatchQty As Double = 500
        'Dim StartRowIndex As Double = 0
        Dim StartDateDummy As Date = #1/1/1950#
        Dim EndDateDummy As Date = #1/1/1950#

        Try
            'open connection
            Conn.Open()

            StartDateDummy = vStartDate.Date
            While StartDateDummy <= vEndDate
                If (StartDateDummy + TimeSpan.FromDays(30)).Date > vEndDate.Date Then
                    EndDateDummy = vEndDate.Date
                Else
                    EndDateDummy = (StartDateDummy + TimeSpan.FromDays(30)).Date
                End If
                sql = Me.GetInsertString(vPrimaryWS, vPrimaryWU, StartDateDummy, EndDateDummy)
                script.Query = sql
                'execute script
                script.Execute()

                StartDateDummy = (EndDateDummy + TimeSpan.FromDays(1)).Date
            End While
            'StartRowIndex += BatchQty

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString(ByVal vPrimaryWS As Double, ByVal vPrimaryWU As Double, ByVal vStartDate As Date, ByVal vEndDate As Date) As String
        Dim retval As String = ""
        'Dim row As DataGridViewRow
        Dim Delimiter As String = ","
        Dim Idx As Integer = 0
        Dim CurDate As Date = Today

        Try
            retval = "INSERT INTO `tbltrans_mo_workshift_allocate` (`date`,`primary work unit`,`primary workshift`,`created by`,`created date`) VALUES "
            While vStartDate.Date + TimeSpan.FromDays(Idx) <= vEndDate
                CurDate = vStartDate.Date + TimeSpan.FromDays(Idx)
                If CurDate.Date = vEndDate.Date Then
                    Delimiter = ";"
                End If
                retval = retval + vbCrLf + "('" + Format(CurDate, "yyyy-MM-dd") + "'," + vPrimaryWU.ToString + "," _
                   + vPrimaryWS.ToString + "," _
                   + "'" + Me.UserCode + "'," _
                   + "'" + Format(Now, "yyyy-MM-dd HH:mm:ss") + "'" _
                   + ")" + Delimiter

                Idx += 1
            End While
            
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub dtpFrom_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpFrom.Enter
        Try
            Me.buffer_from = Me.dtpFrom.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpFrom_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpFrom.Leave
        Try
            If CDate(Me.buffer_from).Date <> Me.dtpFrom.Value.Date Then
                Me.dgvItem.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpTo_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpTo.Enter
        Try
            Me.buffer_to = Me.dtpTo.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpTo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpTo.Leave
        Try
            If CDate(Me.buffer_to) <> Me.dtpTo.Value.Date Then
                Me.dgvItem.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex > 6 Then
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.buffer = Boolean.FalseString
                Else
                    Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            If e.ColumnIndex > 6 Then
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If CBool(Me.buffer) <> CBool(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString) Then
                        If CBool(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            'allocate
                            If Me.AllocateWS(Me.ObjWS.GetPrimary(Me.dgvItem.Columns(e.ColumnIndex).HeaderText, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production), _
                            Me.dgvItem.Item(Me.colPrimaryWU.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.colTglDari.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.colTglSampai.Index, e.RowIndex).Value) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CBool(Me.buffer)
                            End If
                        ElseIf CBool(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            'unallocate
                            If Me.UnAllocateWS(Me.ObjWS.GetPrimary(Me.dgvItem.Columns(e.ColumnIndex).HeaderText, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production), _
                            Me.dgvItem.Item(Me.colPrimaryWU.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.colTglDari.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.colTglSampai.Index, e.RowIndex).Value) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CBool(Me.buffer)
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function AllocateWS(ByVal vPrimaryWS As Double, ByVal vPrimaryWU As Double, ByVal vStartDate As Date, ByVal vEndDate As Date) As Boolean
        Dim FirmPlanDate As Date = Today
        Dim dummyWU As String = ""
        Dim ListPrimaryWU As New List(Of Double)
        Dim DTAAllocate As New DseTransMO_WorkShift_Allocate2TableAdapters.tbltrans_mo_workshift_allocateTableAdapter

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTAAllocate.Connection.ConnectionString = Core_Module.GetConstring

            If vStartDate.Date > vEndDate.Date Then
                MsgBox("Please make sure start date <= end date.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '1. tidak boleh sebelum firm plan
            FirmPlanDate = Me.ObjTM.GetFirmPlan
            If vStartDate.Date <= FirmPlanDate.Date Then
                MsgBox("Please specify date after firm plan: " + Format(FirmPlanDate, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            ''2. tidak boleh kalo sudah ada bundling 
            'dummyWU = ""
            'If ListPrimaryWU.Contains(vPrimaryWU) = False Then
            '    ListPrimaryWU.Add(vPrimaryWU)
            '    If Me.ObjTM.GetJmlBundling(vStartDate.Date, vEndDate.Date, vPrimaryWU, Me.DTABundling) > 0 Then
            '        If dummyWU = "" Then
            '            dummyWU = "1. " + Me.ObjWU.GetCode(vPrimaryWU)
            '        Else
            '            dummyWU = dummyWU + vbCrLf + "1. " + Me.ObjWU.GetCode(vPrimaryWU)
            '        End If
            '    End If
            'End If
            'If dummyWU <> "" Then
            '    MsgBox("There is production log on period: " + Format(vStartDate, "dd MMM yyyy") + " to " + _
            '    Format(vEndDate, "dd MMM yyyy") + " for the following work unit:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            '2.b. tidak boleh klo sudah ada ot
            dummyWU = ""
            ListPrimaryWU.Clear()
            If ListPrimaryWU.Contains(vPrimaryWU) = False Then
                ListPrimaryWU.Add(vPrimaryWU)
                If Me.ObjTM.GetJmlOT(vStartDate.Date, vEndDate.Date, vPrimaryWU, Me.DTA_OT) > 0 Then
                    If dummyWU = "" Then
                        dummyWU = "1. " + Me.ObjWU.GetCode(vPrimaryWU)
                    Else
                        dummyWU = dummyWU + vbCrLf + "1. " + Me.ObjWU.GetCode(vPrimaryWU)
                    End If
                End If
            End If
            If dummyWU <> "" Then
                MsgBox("There is overtime on period: " + Format(vStartDate, "dd MMM yyyy") + " to " + _
                Format(vEndDate, "dd MMM yyyy") + " for the following work unit:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
                Return False
            End If

            '3. tidak boleh kalo ada workshift yang tabrakan -> ini tidak perlu, karena unallocate
            dummyWU = ""
            If Me.ObjTM.ApakahAmanDrJadwalTabrakan(vPrimaryWS, vStartDate.Date, vEndDate.Date, ListPrimaryWU, dummyWU) = False Then
                MsgBox("Conflicting worktime for the following work shift:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
                Return False
            End If

            ''rosadin-r15
            ''4. tidak boleh kalo ada jadwal-sementara
            ''MsgBox("belum :")
            ''MsgBox("belum: harusnya bisa geser schedule aja.")
            'dummyWU = ""
            'ListPrimaryWU.Clear()
            'If ListPrimaryWU.Contains(vPrimaryWU) = False Then
            '    ListPrimaryWU.Add(vPrimaryWU)
            '    If Me.ObjTM.GetJmlSch(vStartDate.Date, vEndDate.Date, vPrimaryWU, Me.DTASch) > 0 Then
            '        If dummyWU = "" Then
            '            dummyWU = "1. " + Me.ObjWU.GetCode(vPrimaryWU)
            '        Else
            '            dummyWU = dummyWU + vbCrLf + "1. " + Me.ObjWU.GetCode(vPrimaryWU)
            '        End If
            '    End If
            'End If
            'If dummyWU <> "" Then
            '    MsgBox("There is production schedule on period: " + Format(vStartDate, "dd MMM yyyy") + " to " + _
            '    Format(vEndDate, "dd MMM yyyy") + " for the following work unit:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
            '    Return False
            'End If
            Dim objSchOT As New ClsTrans_Schedule_OT
            Dim vTBLSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
            Dim vTBLSpaceBeforeAllocate As New DseTrans_Schedule.tbl_sch_barDataTable
            ObjSch.LoadScheduleBar_WOItemOnly(#1/1/1950#, #12/31/2200#, vTBLSchBar)
            ObjSch.LoadScheduleBar_TanpaScheduledItem(ObjSch.GetFirmPlan, #12/31/2200#, vTBLSpaceBeforeAllocate)
            If objSchOT.AnyClosedJO(vTBLSchBar, vStartDate.Date, ListPrimaryWU) Then
                Exit Function
            End If

            Dim message As String = ""
            If objSchOT.AnyProducedJO(vTBLSchBar, vStartDate.Date, ListPrimaryWU, message) Then
                MsgBox("Cannot allocate workshift, there are log transaction for the following JO: " & vbNewLine & message & ".", MsgBoxStyle.Information, "Information")
                Exit Function
            End If

            'mulai allocate
            For Each PrimaryWU As Double In ListPrimaryWU
                DTAAllocate.DeleteQuery(PrimaryWU, vPrimaryWS, vEndDate.Date, vStartDate.Date)
                Me.Simpan_UsingScript(vPrimaryWS, PrimaryWU, vStartDate.Date, vEndDate.Date)
            Next

            ''rosadin-r15
            For Each PrimaryWU As Double In ListPrimaryWU
                objSchOT.RecalculateAfterAddSpace(PrimaryWU, vStartDate.Date, vTBLSpaceBeforeAllocate, vTBLSchBar)
            Next
            Me.IsEdited = True

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryWU = Nothing
            DTAAllocate.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Private Function UnAllocateWS(ByVal vPrimaryWS As Double, ByVal vPrimaryWU As Double, ByVal vStartDate As Date, ByVal vEndDate As Date) As Boolean
        Dim FirmPlanDate As Date = Today
        Dim dummyWU As String = ""
        Dim ListPrimaryWU As New List(Of Double)
        Dim DTAAllocate As New DseTransMO_WorkShift_Allocate2TableAdapters.tbltrans_mo_workshift_allocateTableAdapter

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTAAllocate.Connection.ConnectionString = Core_Module.GetConstring

            If vStartDate.Date > vEndDate.Date Then
                MsgBox("Please make sure start date <= end date.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '1. tidak boleh sebelum firm plan
            FirmPlanDate = Me.ObjTM.GetFirmPlan
            If vStartDate.Date <= FirmPlanDate.Date Then
                MsgBox("Please specify date after firm plan: " + Format(FirmPlanDate, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            ''2. tidak boleh kalo sudah ada bundling
            'dummyWU = ""
            'If ListPrimaryWU.Contains(vPrimaryWU) = False Then
            '    ListPrimaryWU.Add(vPrimaryWU)
            '    If Me.ObjTM.GetJmlBundling(vStartDate.Date, vEndDate.Date, vPrimaryWU, Me.DTABundling) > 0 Then
            '        If dummyWU = "" Then
            '            dummyWU = "1. " + Me.ObjWU.GetCode(vPrimaryWU)
            '        Else
            '            dummyWU = dummyWU + vbCrLf + "1. " + Me.ObjWU.GetCode(vPrimaryWU)
            '        End If
            '    End If
            'End If
            'If dummyWU <> "" Then
            '    MsgBox("There is production log on period: " + Format(vStartDate, "dd MMM yyyy") + " to " + _
            '    Format(vEndDate, "dd MMM yyyy") + " for the following work unit:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            '2.b. tidak boleh ada overtime
            dummyWU = ""
            ListPrimaryWU.Clear()
            If ListPrimaryWU.Contains(vPrimaryWU) = False Then
                ListPrimaryWU.Add(vPrimaryWU)
                If Me.ObjTM.GetJmlOT(vStartDate.Date, vEndDate.Date, vPrimaryWU, Me.DTA_OT) > 0 Then
                    If dummyWU = "" Then
                        dummyWU = "1. " + Me.ObjWU.GetCode(vPrimaryWU)
                    Else
                        dummyWU = dummyWU + vbCrLf + "1. " + Me.ObjWU.GetCode(vPrimaryWU)
                    End If
                End If
            End If
            If dummyWU <> "" Then
                MsgBox("There is overtime on period: " + Format(vStartDate, "dd MMM yyyy") + " to " + _
                Format(vEndDate, "dd MMM yyyy") + " for the following work unit:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
                Return False
            End If

            '3. tidak boleh kalo ada workshift yang tabrakan -> ini tidak perlu, karena unallocate
            dummyWU = ""
            'If Me.ObjTM.ApakahAmanDrJadwalTabrakan(vPrimaryWS, vStartDate.Date, vEndDate.Date, ListPrimaryWU, dummyWU) = False Then
            '    MsgBox("Conflicting worktime for the following work shift:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            '4. tidak boleh kalo ada jadwal-sementara
            'MsgBox("belum :")
            'MsgBox("belum: harusnya bisa geser schedule aja.")

            ''rosadin-r15
            'dummyWU = ""
            'ListPrimaryWU.Clear()
            'If ListPrimaryWU.Contains(vPrimaryWU) = False Then
            '    ListPrimaryWU.Add(vPrimaryWU)
            '    If Me.ObjTM.GetJmlSch(vStartDate.Date, vEndDate.Date, vPrimaryWU, Me.DTASch) > 0 Then
            '        If dummyWU = "" Then
            '            dummyWU = "1. " + Me.ObjWU.GetCode(vPrimaryWU)
            '        Else
            '            dummyWU = dummyWU + vbCrLf + "1. " + Me.ObjWU.GetCode(vPrimaryWU)
            '        End If
            '    End If
            'End If
            'If dummyWU <> "" Then
            '    MsgBox("There is production schedule on period: " + Format(vStartDate, "dd MMM yyyy") + " to " + _
            '    Format(vEndDate, "dd MMM yyyy") + " for the following work unit:" + vbCrLf + dummyWU, MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            Dim vTBLSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
            Dim vTBLSpaceBeforeDeleteSpace As New DseTrans_Schedule.tbl_sch_barDataTable
            Dim TBLUpdate As New DseTrans_OT_Item.DataTable2DataTable
            Dim objSchOT As New ClsTrans_Schedule_OT
            ObjSch.LoadScheduleBar_WOItemOnly(#1/1/1950#, #12/31/2200#, vTBLSchBar)
            ObjSch.LoadScheduleBar_TanpaScheduledItem(ObjSch.GetFirmPlan, #12/31/2200#, vTBLSpaceBeforeDeleteSpace)
            For Each PrimaryWorkunit As Double In ListPrimaryWU
                objSchOT.RecalculateAfterDeleteAllocation(PrimaryWorkunit, vStartDate.Date, vTBLSpaceBeforeDeleteSpace, vTBLSchBar, TBLUpdate, vStartDate.Date, vEndDate.Date, vPrimaryWS)
            Next
            If Not objSchOT.EnoughAllocation(TBLUpdate) Then
                Exit Function
            End If
            If objSchOT.AnyClosedJO(TBLUpdate) Then
                Exit Function
            End If

            Dim message As String = ""
            If objSchOT.AnyProducedJO(TBLUpdate, message) Then
                MsgBox("Cannot un-allocate workshift, there are log transaction for the following JO: " & vbNewLine & message & ".", MsgBoxStyle.Information, "Information")
                Exit Function
            End If

            objSchOT.UpdateScheduleChange(TBLUpdate)
            Me.IsEdited = True

            'mulai unallocate
            For Each PrimaryWU As Double In ListPrimaryWU
                DTAAllocate.DeleteQuery(PrimaryWU, vPrimaryWS, vEndDate.Date, vStartDate.Date)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryWU = Nothing
            DTAAllocate.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class

'#Region "old code"
'Option Explicit On
'Option Strict Off
'Public Class FrmTransMO_WorkShift_Allocate
'    Dim i As Integer = -1
'    Dim c As Boolean = False
'    Dim vInsertQuery, vDeleteQuery As String
'    Dim vMaxQuery As Integer = 500
'    Public UserCode As String = "USER"
'    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
'    Dim TBL_WU As New DseTransMO_WorkShift_Allocate.tblmaster_workunitDataTable
'    Dim row_WU As DseTransMO_WorkShift_Allocate.tblmaster_workunitRow
'    Dim TBL_WS As New DseTransMO_WorkShift_Allocate.tblmaster_workshiftDataTable
'    Dim row_WS As DseTransMO_WorkShift_Allocate.tblmaster_workshiftRow
'    Dim tblLog As New Core_Module.DseTransMO_WorkShift_Allocate.tbltrans_mo_workshift_allocateDataTable
'    Dim tblLog_ As New Core_Module.DseTransMO_WorkShift_Allocate.tbltrans_mo_workshift_allocateDataTable
'    Dim rowLog As Core_Module.DseTransMO_WorkShift_Allocate.tbltrans_mo_workshift_allocateRow
'    Dim rowLog_ As Core_Module.DseTransMO_WorkShift_Allocate.tbltrans_mo_workshift_allocateRow
'    Dim TBL_WU_Dummy As New DseTransMO_WorkShift_Allocate.tblmaster_workunitDataTable
'    Dim row_WU_Dummy As DseTransMO_WorkShift_Allocate.tblmaster_workunitRow
'    Dim TBL_WU_Dummy1 As New DseTransMO_WorkShift_Allocate.tblmaster_workunitDataTable
'    Dim row_WU_Dummy1 As DseTransMO_WorkShift_Allocate.tblmaster_workunitRow
'    Dim tblInsertDB As New DseTransMO_WorkShift_Allocate.tblinsertDBDataTable
'    Dim ObjWU As New Core_Module.ClsMasterWorkUnit
'    Dim ObjWUG As New Core_Module.ClsMasterWorkUnitGroup
'    Dim ObjWS As New Core_Module.ClsMasterWorkShift
'    Dim ObjCD As New Core_Module.ClsCekData
'    Dim ObjTM As New Core_Module.ClsTransMO_WorkShift_Allocate
'    Dim buffer As String = ""
'    Dim EraseRowIndex As Integer = 0
'    Dim EraseRowIndex_WUG As Integer = 0
'    Dim IrowIdx As Integer = 0
'    Dim Jcelledit As Boolean = False
'    Dim Jrowindex As Integer = 0
'    Dim Jbuffercodews As String = ""
'    Dim Jbufferprimaryws As String = "-1"
'    Dim allocate_selesai As Boolean = True
'    Dim list_primarywu As New List(Of Double)
'    Dim TBL As New DseTransMO_WorkShift_Allocate.tblmaster_workunitDataTable
'    Dim row As DseTransMO_WorkShift_Allocate.tblmaster_workunitRow
'    Dim TBL_ As New DseTransMO_WorkShift_Allocate.tblmaster_workunitDataTable
'    Dim row_ As DseTransMO_WorkShift_Allocate.tblmaster_workunitRow
'    Dim primary_wug As Double = -1

'    Private Sub FrmTransMO_WorkShift_Allocate_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
'        Try
'            'If Me.Jcelledit = True Then
'            '    If Me.dgvItem.Item(Me.Jbuffercodews, Me.Jrowindex).Value = False Then
'            '        Me.UpdateDataAllocate1(Me.dgvItem.Item(Me.colTglDari.Index, Me.Jrowindex).Value, Me.dgvItem.Item(Me.colTglDari.Index, Me.Jrowindex).Value, _
'            '        Me.dgvItem.Item(Me.colPrimaryWU.Index, Me.Jrowindex).Value, _
'            '           Jbufferprimaryws, True)
'            '    ElseIf Me.dgvItem.Item(Me.Jbuffercodews, Me.Jrowindex).Value = True Then
'            '        Me.UpdateDataAllocate1(Me.dgvItem.Item(Me.colTglDari.Index, Me.Jrowindex).Value, Me.dgvItem.Item(Me.colTglDari.Index, Me.Jrowindex).Value, _
'            '        Me.dgvItem.Item(Me.colPrimaryWU.Index, Me.Jrowindex).Value, _
'            '           Jbufferprimaryws, False)
'            '    End If
'            'End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub FrmTransMO_WorkShift_Allocate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'        FrmTransMO_WorkShift_Allocate.CheckForIllegalCrossThreadCalls = False
'        Try
'            Me.SetDGV()
'            Me.ObjTM.GetTBL_WU(Me.TBL_WU, -1)
'            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
'            Me.dgvWU.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
'            Me.dgvWUG.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
'            Me.tblInsertDB.Rows.Clear()
'            Me.TBL_WU_Dummy.Rows.Clear()
'            Me.TBL_WU_Dummy1.Rows.Clear()
'            'Me.dgvWUG.AllowUserToAddRows = False

'            Me.btnFilter.Select()
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub dgvWU_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvWU.CellBeginEdit
'        Try
'            If e.ColumnIndex = Me.colWU_Code.Index Then
'                Me.buffer = Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value
'            Else
'                e.Cancel = True
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub dgvWU_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWU.CellEndEdit
'        Dim row As DataGridViewRow = Me.dgvWU.Rows(e.RowIndex)
'        Dim index As Integer = -1
'        Dim PrimaryWU As Double = -1
'        Dim RowIdx As Integer = e.RowIndex

'        Try
'            With row
'                If IsNothing(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
'                    If e.ColumnIndex = Me.colWU_Code.Index Then
'                        Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
'                        If Me.ObjWU.ApakahSudahAda(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value, True) = True Then
'                            PrimaryWU = Me.ObjWU.GetPrimary(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value)
'                            index = Me.GetIndexItem(PrimaryWU, e.RowIndex)
'                            If index < 0 Then
'                                Me.dgvWU.Item(Me.colWU_Primary.Index, e.RowIndex).Value = PrimaryWU.ToString
'                                Me.dgvWU.Item(Me.colWU_Code.Index, e.RowIndex).Value = Me.ObjWU.GetCode(PrimaryWU)
'                                Me.dgvWU.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.ObjWU.GetDesc(PrimaryWU)
'                            Else
'                                If e.RowIndex <> index Then
'                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
'                                    Me.dgvWU.Item(Me.colWU_Code.Index, e.RowIndex).Value = ""
'                                    If String.IsNullOrEmpty(Me.buffer) = True Then
'                                        Me.EraseRowIndex = e.RowIndex
'                                        Me.BackgroundWorker1.RunWorkerAsync()
'                                    Else
'                                        Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
'                                    End If
'                                    Me.ObjCD.SetNumber(Me.dgvWU)
'                                    Exit Sub
'                                End If
'                            End If
'                        Else

'                            Dim ij As New Core_Module.FrmMasterWorkUnit_Cari
'                            ij.txtCari.Text = Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value
'                            ij.BackColorAll = Me.BackColorAll
'                            ij.Status = "Only Active"
'                            ij.ShowDialog()
'                            If ij.PrimaryWorkUnit <> -1 Then
'                                index = Me.GetIndexItem(ij.PrimaryWorkUnit, e.RowIndex)
'                                If index < 0 Then
'                                    Me.dgvWU.Item(Me.colWU_Primary.Index, e.RowIndex).Value = ij.PrimaryWorkUnit
'                                    Me.dgvWU.Item(Me.colWU_Code.Index, e.RowIndex).Value = Me.ObjWU.GetCode(ij.PrimaryWorkUnit)
'                                    Me.dgvWU.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.ObjWU.GetDesc(ij.PrimaryWorkUnit)
'                                Else
'                                    If e.RowIndex <> index Then
'                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
'                                        If String.IsNullOrEmpty(Me.buffer) = True Then
'                                            Me.EraseRowIndex = e.RowIndex
'                                            Me.BackgroundWorker1.RunWorkerAsync()
'                                        Else
'                                            Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
'                                        End If
'                                        Me.ObjCD.SetNumber(Me.dgvWU)
'                                        Exit Sub
'                                    Else
'                                        Me.dgvWU.Item(Me.colWU_Primary.Index, e.RowIndex).Value = ij.PrimaryWorkUnit
'                                        Me.dgvWU.Item(Me.colWU_Code.Index, e.RowIndex).Value = Me.ObjWU.GetCode(ij.PrimaryWorkUnit)
'                                        Me.dgvWU.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.ObjWU.GetDesc(ij.PrimaryWorkUnit)
'                                    End If
'                                End If
'                            ElseIf ij.PrimaryWorkUnit = -1 Then
'                                If Me.dgvWU.Rows.Count - 1 <> e.RowIndex Then
'                                    If String.IsNullOrEmpty(Me.buffer) = True Then
'                                        Me.EraseRowIndex = e.RowIndex
'                                        Me.BackgroundWorker1.RunWorkerAsync()
'                                    Else
'                                        Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
'                                    End If
'                                    Me.ObjCD.SetNumber(Me.dgvWU)
'                                    Exit Sub
'                                End If
'                                If Me.dgvWU.Item(Me.colWU_Code.Index, e.RowIndex).Value <> "" Then
'                                    If String.IsNullOrEmpty(Me.buffer) = True Then
'                                        Me.EraseRowIndex = e.RowIndex
'                                        Me.BackgroundWorker1.RunWorkerAsync()
'                                    Else
'                                        Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
'                                    End If
'                                    Me.ObjCD.SetNumber(Me.dgvWU)
'                                    Exit Sub
'                                End If
'                            End If
'                            ij = Nothing
'                        End If
'                        Me.ObjCD.SetNumber(Me.dgvWU)
'                    End If
'                Else
'                    Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
'                End If
'            End With
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub dgvWU_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvWU.RowsAdded
'        Try
'            Me.ObjCD.SetNumber(Me.dgvWU, False)
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub dgvWU_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvWU.RowsRemoved
'        Try
'            Me.ObjCD.SetNumber(Me.dgvWU, False)
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub dgvWU_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvWU.Sorted
'        Try
'            Me.ObjCD.SetNumber(Me.dgvWU)
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
'        Me.ObjCD.UpdateThreading()
'        Try
'            Threading.Thread.Sleep(1)
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
'        Try
'            Me.dgvWU.Rows.RemoveAt(EraseRowIndex)
'            Me.ObjCD.SetNumber(Me.dgvWU, False)
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Function GetIndexItem(ByVal vPrimary As Double, _
'      ByVal vRowIdxFocus As Integer) As Double
'        Dim RowDGV As DataGridViewRow
'        Dim Retval As Double = -1
'        Try
'            For ij As Integer = 0 To Me.dgvWU.Rows.Count - 2
'                RowDGV = Me.dgvWU.Rows(ij)
'                If vRowIdxFocus <> ij Then
'                    If RowDGV.Cells(Me.colWU_Primary.Index).Value = vPrimary Then
'                        Retval = ij
'                        Exit For
'                    End If
'                End If
'            Next
'            Return Retval
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Function

'    Private Sub dgvWUGG_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvWUG.CellBeginEdit
'        Try
'            If e.ColumnIndex = Me.colWUG_Code.Index Then
'                Me.buffer = Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value
'            Else
'                e.Cancel = True
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub dgvWUG_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWUG.CellEndEdit
'        Dim row As DataGridViewRow = Me.dgvWUG.Rows(e.RowIndex)
'        Dim index As Integer = -1
'        Dim PrimaryWUG As Double = -1
'        Dim RowIdx As Integer = e.RowIndex
'        Dim dV As New DataGridViewRow
'        Try
'            With row
'                If IsNothing(Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
'                    If e.ColumnIndex = Me.colWUG_Code.Index Then
'                        Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
'                        If Me.ObjWUG.Apakahsudahada(Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
'                            PrimaryWUG = Me.ObjWUG.GetPrimary(Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value)
'                            index = Me.GetIndexItem_WUG(PrimaryWUG, e.RowIndex)
'                            If index < 0 Then
'                                Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value = PrimaryWUG
'                                Me.dgvWUG.Item(Me.colWUG_Code.Index, e.RowIndex).Value = Me.ObjWUG.GetCode(PrimaryWUG)
'                                Me.dgvWUG.Item(Me.colWUG_Desc.Index, e.RowIndex).Value = Me.ObjWUG.GetDesc(PrimaryWUG)
'                                .Selected = True
'                            Else
'                                If e.RowIndex <> index Then
'                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
'                                    Me.dgvWUG.Item(Me.colWUG_Code.Index, e.RowIndex).Value = ""
'                                    If String.IsNullOrEmpty(Me.buffer) = True Then
'                                        Me.EraseRowIndex_WUG = e.RowIndex
'                                        Me.BackgroundWorker2.RunWorkerAsync()
'                                    Else
'                                        Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
'                                    End If
'                                    Me.ObjCD.SetNumber(Me.dgvWUG)
'                                    Exit Sub
'                                End If
'                            End If
'                        Else

'                            Dim ij As New Core_Module.FrmMasterWorkUnitGroup_Cari
'                            ij.txtCari.Text = Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value
'                            ij.BackColorAll = Me.BackColorAll
'                            ij.ShowDialog()
'                            If ij.PrimaryWorkUnitGroup <> -1 Then
'                                index = Me.GetIndexItem_WUG(ij.PrimaryWorkUnitGroup, e.RowIndex)
'                                If index < 0 Then
'                                    Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value = ij.PrimaryWorkUnitGroup
'                                    Me.dgvWUG.Item(Me.colWUG_Code.Index, e.RowIndex).Value = Me.ObjWUG.GetCode(ij.PrimaryWorkUnitGroup)
'                                    Me.dgvWUG.Item(Me.colWUG_Desc.Index, e.RowIndex).Value = Me.ObjWUG.GetDesc(ij.PrimaryWorkUnitGroup)
'                                    Me.dgvWUG.Rows(dgvWUG.Rows.Count - 1).Selected = True
'                                Else
'                                    If e.RowIndex <> index Then
'                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
'                                        If String.IsNullOrEmpty(Me.buffer) = True Then
'                                            Me.EraseRowIndex_WUG = e.RowIndex
'                                            Me.BackgroundWorker2.RunWorkerAsync()
'                                        Else
'                                            Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
'                                        End If
'                                        Me.ObjCD.SetNumber(Me.dgvWUG)
'                                        Exit Sub
'                                    Else
'                                        Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value = ij.PrimaryWorkUnitGroup
'                                        Me.dgvWUG.Item(Me.colWUG_Code.Index, e.RowIndex).Value = Me.ObjWUG.GetCode(ij.PrimaryWorkUnitGroup)
'                                        Me.dgvWUG.Item(Me.colWUG_Desc.Index, e.RowIndex).Value = Me.ObjWUG.GetDesc(ij.PrimaryWorkUnitGroup)
'                                        Me.dgvWUG.Rows(dgvWUG.Rows.Count - 1).Selected = True
'                                    End If
'                                End If
'                            ElseIf ij.PrimaryWorkUnitGroup = -1 Then
'                                If Me.dgvWUG.Rows.Count - 1 <> e.RowIndex Then
'                                    If String.IsNullOrEmpty(Me.buffer) = True Then
'                                        Me.EraseRowIndex_WUG = e.RowIndex
'                                        Me.BackgroundWorker2.RunWorkerAsync()
'                                    Else
'                                        Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
'                                    End If
'                                    Me.ObjCD.SetNumber(Me.dgvWUG)
'                                    Exit Sub
'                                End If
'                                If Me.dgvWUG.Item(Me.colWUG_Code.Index, e.RowIndex).Value <> "" Then
'                                    If String.IsNullOrEmpty(Me.buffer) = True Then
'                                        Me.EraseRowIndex_WUG = e.RowIndex
'                                        Me.BackgroundWorker2.RunWorkerAsync()
'                                    Else
'                                        Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
'                                    End If
'                                    Me.ObjCD.SetNumber(Me.dgvWUG)
'                                    Exit Sub
'                                End If
'                            End If
'                            ij = Nothing
'                        End If
'                        Me.ObjCD.SetNumber(Me.dgvWUG)
'                    End If
'                Else
'                    Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
'                End If
'            End With
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub dgvWUG_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvWUG.RowsAdded
'        Try
'            Me.ObjCD.SetNumber(Me.dgvWUG, False)
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub dgvWUG_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvWUG.SelectionChanged
'        Try
'            If dgvWUG.SelectedCells.Count > 0 Then
'                IrowIdx = Me.dgvWUG.SelectedCells(0).RowIndex
'                If Me.dgvWUG.Item(Me.colWUG_Code.Index, IrowIdx).Value <> "" Then
'                    isillb_wug(Me.dgvWUG.Item(Me.colWUG_Primary.Index, IrowIdx).Value)
'                Else
'                    Me.TBL_.Rows.Clear()
'                End If
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub dgvWUG_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvWUG.Sorted
'        Try
'            Me.ObjCD.SetNumber(Me.dgvWUG, False)
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
'        Me.ObjCD.UpdateThreading()
'        Try
'            Threading.Thread.Sleep(1)
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
'        Try
'            Me.dgvWUG.Rows.RemoveAt(EraseRowIndex_WUG)
'            Me.ObjCD.SetNumber(Me.dgvWUG, False)
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Function GetIndexItem_WUG(ByVal vPrimary As Double, _
'      ByVal vRowIdxFocus As Integer) As Double
'        Dim RowDGV As DataGridViewRow
'        Dim Retval As Double = -1
'        Try
'            For ij As Integer = 0 To Me.dgvWUG.Rows.Count - 2
'                RowDGV = Me.dgvWUG.Rows(ij)
'                If vRowIdxFocus <> ij Then
'                    If RowDGV.Cells(Me.colWUG_Primary.Index).Value = vPrimary Then
'                        Retval = ij
'                        Exit For
'                    End If
'                End If
'            Next
'            Return Retval
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Function

'    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
'        Try
'            For i As Integer = 0 To Me.ClbWU.Items.Count - 1
'                Me.ClbWU.SetItemCheckState(i, CheckState.Unchecked)
'            Next

'            For i As Integer = 0 To Me.ClbWU.Items.Count - 1
'                Me.ClbWU.SetItemCheckState(i, CheckState.Checked)
'            Next
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub llbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
'        Try
'            For i As Integer = 0 To Me.ClbWU.Items.Count - 1
'                Me.ClbWU.SetItemCheckState(i, CheckState.Unchecked)
'            Next
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub rboWU_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rboWU.CheckedChanged
'        Try
'            If rboWU.Checked = True Then
'                ClbWU.DataSource = Nothing
'                Me.TBL_WU_Dummy.Rows.Clear()
'                Me.TBL_WU_Dummy1.Rows.Clear()
'                'Me.dgvWUG.ClearSelection()
'                'Me.dgvWUG.AllowUserToAddRows = False
'                'Me.dgvWU.AllowUserToAddRows = True
'                Me.ClbWU.Items.Clear()
'                Me.dgvWUG.Rows.Clear()
'                Me.dgvWU.ReadOnly = False
'                Me.dgvWUG.ReadOnly = True
'                Me.llbCheckAll.Enabled = False
'                Me.llbUncheckAll.Enabled = False
'                Me.dgvWU.AllowUserToDeleteRows = True
'                Me.dgvWUG.AllowUserToDeleteRows = False
'                Me.ClbWU.Enabled = False
'                Me.dgvWU.TabStop = True
'                Me.dgvWUG.TabStop = False
'                Me.list_primarywu.Clear()
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub rboWUG_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rboWUG.CheckedChanged
'        Try
'            If rboWUG.Checked = True Then
'                Me.dgvWU.Rows.Clear()
'                'Me.dgvWU.AllowUserToAddRows = False
'                'Me.dgvWUG.AllowUserToAddRows = True
'                Me.dgvWU.ReadOnly = True
'                Me.dgvWUG.ReadOnly = False
'                Me.dgvWU.AllowUserToDeleteRows = False
'                Me.dgvWUG.AllowUserToDeleteRows = True
'                Me.llbCheckAll.Enabled = True
'                Me.llbUncheckAll.Enabled = True
'                Me.ClbWU.Enabled = True
'                Me.dgvWU.TabStop = False
'                Me.dgvWUG.TabStop = True
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub SetDGV()
'        Dim col As DataGridViewColumn
'        Try
'            Me.ObjTM.GetTBL_WS(Me.TBL_WS)
'            For Each row_WS In TBL_WS
'                col = New DataGridViewCheckBoxColumn
'                col.Name = row_WS.code
'                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
'                col.SortMode = DataGridViewColumnSortMode.NotSortable
'                col.HeaderText = row_WS.code.Substring(0, 1) + row_WS.code.Substring(1, row_WS.code.Length - 1).ToLower
'                Me.dgvItem.Columns.Add(col)
'            Next
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Public Sub FilterData_WorkUnit()
'        Try
'            Me.TBL_WU_Dummy1.Rows.Clear()
'            If Me.rboWU.Checked = True Then
'                If Me.dgvWU.Rows.Count > 0 Then
'                    If String.IsNullOrEmpty(Me.dgvWU.Item(Me.colWUG_Code.Index, 0).Value) = True Then
'                        Exit Sub
'                    End If

'                    TBL_WU.DefaultView.RowFilter = Me.GetFilter(Me.dgvWU, "primary", Me.colWU_Primary)
'                    For vInt As Integer = 0 To Me.TBL_WU.DefaultView.Count - 1
'                        row_WU = Me.TBL_WU.DefaultView(vInt).Row
'                        row_WU_Dummy1 = TBL_WU_Dummy1.Newtblmaster_workunitRow
'                        row_WU_Dummy1.primary = row_WU.primary
'                        row_WU_Dummy1.primary_work_unit_group = row_WU.primary_work_unit_group
'                        row_WU_Dummy1.work_unit_code = row_WU.work_unit_code
'                        row_WU_Dummy1.description = row_WU.description
'                        TBL_WU_Dummy1.Addtblmaster_workunitRow(row_WU_Dummy1)
'                    Next
'                End If
'            Else
'                isitbldummywu_byprimarywug()
'                For Each row_WU_Dummy In TBL_WU_Dummy
'                    row_WU_Dummy1 = TBL_WU_Dummy1.Newtblmaster_workunitRow
'                    row_WU_Dummy1.primary = row_WU_Dummy.primary
'                    row_WU_Dummy1.primary_work_unit_group = row_WU_Dummy.primary_work_unit_group
'                    row_WU_Dummy1.work_unit_code = row_WU_Dummy.work_unit_code
'                    row_WU_Dummy1.description = row_WU_Dummy.description
'                    TBL_WU_Dummy1.Addtblmaster_workunitRow(row_WU_Dummy1)
'                Next
'                If TBL_WU_Dummy.Rows.Count = 0 Then TBL_WU_Dummy1.Rows.Clear()
'            End If
'            TBL_WU.DefaultView.RowFilter = Nothing
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Public Function GetFilter(ByVal vDgvItem As DataGridView, ByVal Primary_Filter As String, ByVal vCol_Primary As DataGridViewTextBoxColumn) As String
'        Dim Retval_ As String = ""
'        Dim row_ As DataGridViewRow
'        Try
'            vDgvItem.AllowUserToAddRows = False
'            If vDgvItem.RowCount <> 0 Then
'                For Each row_ In vDgvItem.Rows
'                    If row_.Index = 0 Then
'                        Retval_ &= "(" & "[" & Primary_Filter & "] =" & row_.Cells(vCol_Primary.Index).Value
'                    Else
'                        If row_.Index = vDgvItem.RowCount - 1 Then
'                            Retval_ &= " OR " & "[" & Primary_Filter & "] =" & row_.Cells(vCol_Primary.Index).Value & ")"
'                        Else
'                            Retval_ &= " OR " & "[" & Primary_Filter & "] =" & row_.Cells(vCol_Primary.Index).Value
'                        End If
'                    End If
'                Next
'                If vDgvItem.RowCount = 1 Then
'                    Retval_ &= ")"
'                End If
'            End If
'            vDgvItem.AllowUserToAddRows = True
'            Return Retval_
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'            Return ""
'        End Try
'    End Function

'    Dim buffer_chk As Boolean = False
'    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
'        Try
'            If e.ColumnIndex = Me.colWorkUnitCode.Index Or e.ColumnIndex = Me.colWorkUnitDesc.Index Or e.ColumnIndex = Me.colDate.Index Then
'                e.Cancel = True
'            Else
'                For Each row_WS In TBL_WS.Rows
'                    If e.ColumnIndex = dgvItem.Columns.Item(row_WS.code).Index Then
'                        Jbuffercodews = row_WS.code
'                        Jbufferprimaryws = row_WS.primary
'                        buffer_chk = Me.dgvItem.Item(row_WS.code, e.RowIndex).Value
'                        Exit For
'                    End If
'                Next
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Public Function UpdateDataAllocate1(ByVal vTglDari As Date, ByVal vTglSampai As Date, ByVal vPrimaryWU As String, ByVal vPrimaryWS As String, Optional ByVal Save As Boolean = True) As Boolean
'        Dim List_WU As New List(Of Double)
'        Try
'            If Save = True Then
'                Me.ObjTM.DeleteAllocate(vTglDari, vTglSampai, CDbl(vPrimaryWU), CDbl(vPrimaryWS), True, False)
'                If Me.ObjTM.InsertAllocate(vTglDari, vTglSampai, Me.UserCode, vPrimaryWS, vPrimaryWU, Me.tblInsertDB) = False Then
'                    Me.dgvItem.Item(Me.ObjWS.GetCodeByPrimary(vPrimaryWS), Jrowindex).Value = False
'                End If
'            Else
'                If Me.ObjTM.DeleteAllocate(vTglDari, vTglSampai, CDbl(vPrimaryWU), CDbl(vPrimaryWS), True, False) = False Then
'                    Me.dgvItem.Item(Me.ObjWS.GetCodeByPrimary(vPrimaryWS), Jrowindex).Value = True
'                End If
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'            Return False
'        Finally
'        End Try
'    End Function

'    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
'        Try
'            If e.ColumnIndex = Me.colWorkUnitCode.Index Or e.ColumnIndex = Me.colWorkUnitDesc.Index Or e.ColumnIndex = Me.colDate.Index Or _
'            e.ColumnIndex = Me.colTglDari.Index Or e.ColumnIndex = Me.colTglSampai.Index Then
'                Exit Sub
'            Else
'                Jcelledit = True
'                Jrowindex = e.RowIndex
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
'        Dim vTgldari As Date
'        Dim vTglsampai As Date
'        Dim vPrimaryWU As String = "-1"
'        Try
'            If e.ColumnIndex = Me.colWorkUnitCode.Index Or e.ColumnIndex = Me.colWorkUnitDesc.Index Or e.ColumnIndex = Me.colDate.Index Or _
'            e.ColumnIndex = Me.colTglDari.Index Or e.ColumnIndex = Me.colTglSampai.Index Then
'                Exit Sub
'            Else
'                Me.Jcelledit = False
'                vTgldari = Convert.ToDateTime(Me.dgvItem.Item(Me.colTglDari.Index, e.RowIndex).Value)
'                vTglsampai = CDate(Me.dgvItem.Item(Me.colTglSampai.Index, e.RowIndex).Value)
'                vPrimaryWU = Me.dgvItem.Item(Me.colPrimaryWU.Index, e.RowIndex).Value
'                If e.ColumnIndex = dgvItem.Columns.Item(Jbuffercodews).Index Then
'                    If Me.dgvItem.Item(Jbuffercodews, e.RowIndex).Value = buffer_chk Then Exit Sub

'                    If Me.dgvItem.Item(Jbuffercodews, e.RowIndex).Value = True Then
'                        Me.Cursor = Cursors.WaitCursor
'                        Me.UpdateDataAllocate1(vTgldari, vTglsampai, vPrimaryWU, Jbufferprimaryws, True)
'                    Else
'                        Me.Cursor = Cursors.WaitCursor
'                        Me.UpdateDataAllocate1(vTgldari, vTglsampai, vPrimaryWU, Jbufferprimaryws, False)
'                    End If
'                End If
'            End If

'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'            Me.Cursor = Cursors.Default
'        Finally
'            Me.Cursor = Cursors.Default
'            Me.dgvItem.Refresh()
'            Me.allocate_selesai = True
'        End Try
'    End Sub

'    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
'        Me.Close()
'    End Sub

'    Public Function isi_DGV(ByVal tblWU As Core_Module.DseTransMO_WorkShift_Allocate.tblmaster_workunitDataTable, _
'    ByVal rowWU As Core_Module.DseTransMO_WorkShift_Allocate.tblmaster_workunitRow) As Boolean
'        Dim dtaLog As New DseTransMO_WorkShift_AllocateTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
'        Dim date_ As Date
'        Dim dr As DataGridViewRow
'        Dim listCol As New List(Of String)
'        Dim List_WU As New List(Of String)
'        Try
'            dtaLog.Connection.ConnectionString = Core_Module.GetConstring

'            Me.dgvItem.Rows.Clear()
'            Me.tblInsertDB.Rows.Clear()
'            dtaLog.Connection.ConnectionString = Core_Module.GetConstring
'            dtaLog.FillBy_TgldariDanTglSampai(Me.tblLog, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
'            'tblLog_.Rows.Clear()

'            'isikan tabel dgn data "date original" /hari untuk disimpan ke db berdasarkan tgl awal dan akhir di form serta dari db sblmnya jika sudah ada
'            If Me.TBL_WU_Dummy1.Rows.Count = 0 Then
'                Me.isi_tblInsertDB(Me.tblLog, Me.tblInsertDB, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TBL_WU)
'            Else
'                Me.isi_tblInsertDB(Me.tblLog, Me.tblInsertDB, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TBL_WU_Dummy1)
'            End If

'            For Each Me.rowLog In Me.tblLog
'                Me.rowLog_ = Me.tblLog_.Newtbltrans_mo_workshift_allocateRow
'                For i As Integer = 0 To Me.tblLog.Columns.Count - 1
'                    Me.rowLog_.Item(i) = Me.rowLog.Item(i)
'                Next
'                Me.tblLog_.Rows.Add(Me.rowLog_)
'            Next

'            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
'            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

'            For Each rowWU In tblWU
'                For i As Integer = 0 To DateDiff(DateInterval.Day, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
'                    date_ = Me.dtpFrom.Value.Date.AddDays(i)
'                    If Me.IsDifferentCombination(rowWU.primary, date_) = True Then
'                        listCol.Clear()
'                        For Each Me.row_WS In Me.TBL_WS
'                            listCol.Add(GetProperty(rowWU.primary, date_, Me.row_WS.primary).ToString)
'                        Next
'                        If GetIndex(rowWU.primary, listCol) = -1 Then
'                            dgvItem.Rows.Add()
'                            dr = dgvItem.Rows(dgvItem.Rows.Count - 1)
'                            dr.Cells(colNo.Index).Value = dgvItem.Rows.Count.ToString
'                            dr.Cells(colWorkUnitCode.Index).Value = rowWU.work_unit_code.ToUpper
'                            dr.Cells(colWorkUnitDesc.Index).Value = rowWU.description
'                            dr.Cells(colTglDari.Index).Value = Format(date_, "MM/dd/yyyy")
'                            dr.Cells(colTglSampai.Index).Value = Format(date_, "MM/dd/yyyy")
'                            dr.Cells(colDate.Index).Value = Format(date_, "dd-MMM-yy")
'                            dr.Cells(colPrimaryWU.Index).Value = rowWU.primary
'                            For Each row_WS In TBL_WS
'                                Dim Cstring As String = ""
'                                If row_WS.code.Length > 1 Then
'                                    Cstring = row_WS.code.Substring(0, 1) + row_WS.code.Substring(1, row_WS.code.Length - 1).ToLower
'                                Else
'                                    Cstring = row_WS.code
'                                End If
'                                dr.Cells(Cstring).Value = GetProperty(rowWU.primary, date_, row_WS.primary)
'                            Next
'                        Else
'                            dr = dgvItem.Rows(dgvItem.Rows.Count - 1)
'                            dr.Cells(colTglSampai.Index).Value = Format(date_, "MM/dd/yyyy")
'                            dr.Cells(colDate.Index).Value = Format(CDate(dr.Cells(colTglDari.Index).Value), "dd-MMM-yy") & " to " & Format(date_, "dd-MMM-yy")
'                        End If
'                    Else

'                        dr = dgvItem.Rows(dgvItem.Rows.Count - 1)
'                        dr.Cells(colTglSampai.Index).Value = Format(date_, "MM/dd/yyyy")
'                        dr.Cells(colDate.Index).Value = Format(CDate(dr.Cells(colTglDari.Index).Value), "dd-MMM-yy") & " to " & Format(date_, "dd-MMM-yy")
'                    End If
'                Next
'            Next

'            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
'            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        Finally
'            dtaLog.Dispose()
'        End Try
'    End Function

'    Private Function GetIndex(ByVal vPrimaryWorkUnit As Double, ByVal vListOfWS As List(Of String)) As Integer
'        Dim dr As DataGridViewRow
'        Try
'            If Me.dgvItem.Rows.Count > 0 Then
'                dr = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
'                If dr.Cells(Me.colPrimaryWU.Index).Value = vPrimaryWorkUnit Then
'                    For Each Me.row_WS In Me.TBL_WS
'                        If dr.Cells(Me.row_WS.code).Value <> vListOfWS.Item(Me.TBL_WS.Rows.IndexOf(Me.row_WS)) Then
'                            Return -1
'                        End If
'                    Next
'                    Return Me.dgvItem.Rows.IndexOf(dr)
'                End If
'                Return -1
'            Else
'                Return -1
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Function

'    Private Function GetProperty(ByVal vPrimaryWorkUnit As Double, ByVal vDate As Date, ByVal vPrimaryWorkShift As Double) As Boolean
'        Try
'            Me.tblLog.DefaultView.RowFilter = "[primary work unit]=" & vPrimaryWorkUnit.ToString & " and [date] = #" & Format(vDate, "MM/dd/yyyy") & "# and [primary workshift] = " & vPrimaryWorkShift.ToString & ""
'            If tblLog.DefaultView.Count > 0 Then
'                Return True
'            End If
'            Return False
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Function

'    Private Function IsDifferentCombination(ByVal vPrimaryWorkUnit As Double, ByVal vDate As Date) As Boolean
'        Try
'            Me.tblLog_.DefaultView.RowFilter = "[primary work unit] = " & vPrimaryWorkUnit.ToString & " and [date] = #" & Format(vDate.AddDays(-1), "MM/dd/yyyy") & "#"
'            Me.tblLog.DefaultView.RowFilter = "[primary work unit] = " & vPrimaryWorkUnit.ToString & " and [date] = #" & Format(vDate, "MM/dd/yyyy") & "#"
'            If Me.tblLog_.DefaultView.Count = Me.tblLog.DefaultView.Count Then
'                If tblLog.DefaultView.Count = 0 Then
'                    Return True
'                End If
'                Me.tblLog_.DefaultView.Sort = "[primary workshift]"
'                Me.tblLog.DefaultView.Sort = "[primary workshift]"
'                For i As Integer = 0 To tblLog_.DefaultView.Count - 1
'                    Me.rowLog_ = Me.tblLog_.DefaultView(i).Row
'                    Me.rowLog = Me.tblLog.DefaultView(i).Row
'                    If Me.rowLog_.primary_workshift <> Me.rowLog.primary_workshift Then
'                        Return True
'                    End If
'                Next
'                Return False
'            End If
'            Return True
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Function

'    Private Function isi_tblInsertDB(ByVal tblAllocate As DseTransMO_WorkShift_Allocate.tbltrans_mo_workshift_allocateDataTable, _
'ByRef tblInsertDB As DseTransMO_WorkShift_Allocate.tblinsertDBDataTable, ByVal vtglDari As Date, ByVal vtglSampai As Date, _
'ByVal tblWU As DseTransMO_WorkShift_Allocate.tblmaster_workunitDataTable) As Boolean
'        Dim rowAllocate As DseTransMO_WorkShift_Allocate.tbltrans_mo_workshift_allocateRow
'        Dim rowInsertDB As DseTransMO_WorkShift_Allocate.tblinsertDBRow
'        Dim rowWU As DseTransMO_WorkShift_Allocate.tblmaster_workunitRow
'        Dim date_ As Date
'        Try
'            If tblWU.Rows.Count > 0 Then
'                For Each rowWU In tblWU
'                    For int As Integer = 0 To DateDiff(DateInterval.Day, vtglDari, vtglSampai)
'                        date_ = vtglDari.Date.AddDays(int)

'                        tblAllocate.DefaultView.Sort = "[work unit code] ASC,[date] ASC"
'                        tblAllocate.DefaultView.RowFilter = ""
'                        tblAllocate.DefaultView.RowFilter = "[date]=#" & Format(date_.Date, "MM/dd/yyyy") & "#" _
'                        & " AND [primary work unit]=" & rowWU.primary.ToString
'                        If tblAllocate.DefaultView.Count > 0 Then
'                            For count As Integer = 0 To tblAllocate.DefaultView.Count - 1
'                                rowAllocate = tblAllocate.DefaultView(count).Row
'                                rowInsertDB = tblInsertDB.NewtblinsertDBRow
'                                With rowInsertDB
'                                    .primary = rowAllocate.primary
'                                    .primary_work_shift = rowAllocate.primary_workshift
'                                    .primary_work_unit = rowAllocate.primary_work_unit
'                                    .work_shift_code = rowAllocate.work_shift_code
'                                    .work_shift_desc = rowAllocate.work_shift_desc
'                                    .work_unit_code = rowAllocate.work_unit_code
'                                    .work_unit_desc = rowAllocate.description
'                                    ._date = Format(rowAllocate._date, "MM/dd/yyyy").ToString
'                                    .tgl_dari = rowAllocate._date
'                                    .tgl_sampai = rowAllocate._date
'                                End With
'                                tblInsertDB.AddtblinsertDBRow(rowInsertDB)
'                            Next
'                        Else
'                            rowInsertDB = tblInsertDB.NewtblinsertDBRow
'                            With rowInsertDB
'                                .primary = -1
'                                .primary_work_shift = -1
'                                .primary_work_unit = rowWU.primary.ToString
'                                .work_shift_code = ""
'                                .work_shift_desc = ""
'                                .work_unit_code = rowWU.work_unit_code
'                                .work_unit_desc = rowWU.description
'                                ._date = Format(date_.Date, "MM/dd/yyyy".ToString)
'                                .tgl_dari = date_
'                                .tgl_sampai = date_
'                            End With
'                            tblInsertDB.AddtblinsertDBRow(rowInsertDB)
'                        End If
'                    Next
'                Next

'                For Each rowWU In tblWU
'                    For int As Integer = 0 To DateDiff(DateInterval.Day, vtglDari, vtglSampai)
'                        date_ = vtglDari.Date.AddDays(int)
'                        tblInsertDB.DefaultView.RowFilter = ""
'                        tblInsertDB.DefaultView.RowFilter = "[date]='" & Format(date_, "MM/dd/yyyy") & "'" & "AND [primary work unit]=" & rowWU.primary
'                        If tblInsertDB.DefaultView.Count = 0 Then
'                            rowInsertDB = tblInsertDB.NewtblinsertDBRow
'                            With rowInsertDB
'                                .primary = -1
'                                .primary_work_shift = -1
'                                .primary_work_unit = rowWU.primary.ToString
'                                .work_shift_code = ""
'                                .work_shift_desc = ""
'                                .work_unit_code = rowWU.work_unit_code
'                                .work_unit_desc = rowWU.description
'                                ._date = Format(date_.Date, "MM/dd/yyyy".ToString)
'                                .tgl_dari = date_
'                                .tgl_sampai = date_
'                            End With
'                            tblInsertDB.AddtblinsertDBRow(rowInsertDB)
'                        End If
'                    Next
'                Next
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Function

'    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError
'    End Sub

'    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
'        Try
'            Me.ObjCD.SetNumber(Me.dgvItem, False)
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub iscbo_cheklist()
'        Try
'            Me.cbo_shift.DataSource = Nothing
'            Me.cbo_shift.Items.Clear()
'            If TBL_WS.Count > 0 Then
'                Me.cbo_shift.DataSource = TBL_WS.DefaultView
'                Me.cbo_shift.DisplayMember = "code"
'                Me.cbo_shift.SelectedIndex = 0
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
'        Try
'            If Me.ProgressBar1.Visible = False And Me.cbo_shift.SelectedIndex <> -1 And Me.dgvItem.RowCount > 0 Then
'                Me.ProgressBar1.Visible = True
'                Me.BackgroundWorker3.RunWorkerAsync()
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub ExecuteQuery(ByVal vQueryDelete As String, ByVal vQueryInsert As String, Optional ByVal deleted_only As Boolean = False)
'        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(Core_Module.GetConstring)
'        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, vQueryDelete)
'        Try
'            Conn.Open()
'            If deleted_only = False Then
'                script.Query = vQueryDelete
'                script.Execute()
'                script.Query = vInsertQuery
'                script.Execute()
'            Else
'                script.Query = vQueryDelete
'                script.Execute()
'            End If
'            Conn.Close()
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub CheckAll(ByVal vAllocate As Boolean, ByVal vColName As String)
'        Try
'            For Each DataGridViewRow_ As DataGridViewRow In Me.dgvItem.Rows
'                DataGridViewRow_.Cells(vColName).Value = vAllocate
'            Next
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Dim vmsg_allocate As String
'    Dim vmsgHapus As String
'    Dim vjmlbatas_msgallocate As Integer = 0

'    Private Sub GenerateQuery(ByRef vQueryDelete As String, ByRef vQueryInsert As String)
'        Dim primaryWorkShift As Double = Me.ObjWS.GetPrimary(Me.cbo_shift.Text)
'        Dim counter As Integer = 0

'        vQueryDelete = ""
'        vInsertQuery = ""
'        vmsg_allocate = ""
'        vjmlbatas_msgallocate = 0
'        Dim vlist_codewu As New List(Of String)
'        Dim vlist_codewu2 As New List(Of String)
'        Dim jml_pjgmsg As Integer = 0
'        Try
'            'buat rangka scrip pertama u/ insert
'            vQueryInsert = "INSERT INTO tbltrans_mo_workshift_allocate (`date`,`primary work unit`,`primary workshift`,`created by`,`created date`) VALUES "
'            Me.ObjTM.isitbl_spl()
'            Me.ObjTM.isitbl_bundling(False, #1/1/1950#, #1/1/1950#)
'            For Each DataGridViewRow_ As DataGridViewRow In Me.dgvItem.Rows
'                With DataGridViewRow_
'                    '
'                    Me.ObjTM.DeleteAllocate(.Cells(Me.colTglDari.Index).Value, .Cells(Me.colTglSampai.Index).Value, .Cells(Me.colPrimaryWU.Index).Value, primaryWorkShift, False, True)
'                    'jika maximum yang akan di allocatenya < dari batas
'                    If counter < vMaxQuery Then 'skala per baris di dgv
'                        For i As Integer = 0 To DateDiff(DateInterval.Day, CDate(.Cells(Me.colTglDari.Index).Value), CDate(.Cells(Me.colTglSampai.Index).Value))
'                            'jika data tidak bentrok dgn data yang lain
'                            If Me.ObjTM.apakahschedule_allocatebenar(.Cells(Me.colPrimaryWU.Index).Value, Format(CDate(.Cells(Me.colTglDari.Index).Value).AddDays(i), "yyyy-MM-dd"), primaryWorkShift, True) = True Then
'                                If counter < vMaxQuery Then 'skala per tanggal
'                                    counter += 1
'                                Else
'                                    'jika > dari batas maka buat rangka scrip baru
'                                    counter = 1
'                                    vQueryInsert &= "INSERT INTO tbltrans_mo_workshift_allocate (`date`,`primary work unit`,`primary workshift`,`created by`,`created date`) VALUES "
'                                End If
'                                'masukan ke rangka scripnya
'                                vQueryInsert &= "('" & Format(CDate(.Cells(Me.colTglDari.Index).Value).AddDays(i), "yyyy-MM-dd") & "'," & .Cells(Me.colPrimaryWU.Index).Value & "," & primaryWorkShift.ToString & ",'" & Me.UserCode & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "'),"
'                                'jika sudah sampai batas per tanggal kasih tanda ; selesai
'                                If counter = vMaxQuery Then
'                                    vQueryInsert = vQueryInsert.Substring(0, vQueryInsert.Length - 1)
'                                    vQueryInsert &= ";" & vbNewLine
'                                End If
'                            Else
'                                'jika data tidak benar
'                                If vjmlbatas_msgallocate < 5 Then
'                                    If vlist_codewu.Contains(.Cells(Me.colWorkUnitCode.Index).Value) = True Then
'                                        If jml_pjgmsg < 5 Then
'                                            vmsg_allocate &= ", " & Format(CDate(.Cells(Me.colTglDari.Index).Value).AddDays(i), "dd MMM yyyy")
'                                            jml_pjgmsg += 1
'                                        Else
'                                            If vmsg_allocate.Substring(vmsg_allocate.Length - 1, 1) <> "," Then
'                                                vmsg_allocate &= " etc.,"
'                                                vlist_codewu2.Add(.Cells(Me.colWorkUnitCode.Index).Value)
'                                            End If

'                                        End If
'                                    Else
'                                        vlist_codewu.Add(.Cells(Me.colWorkUnitCode.Index).Value)
'                                        vmsg_allocate &= vbNewLine & .Cells(Me.colWorkUnitCode.Index).Value & " :  " & Format(CDate(.Cells(Me.colTglDari.Index).Value).AddDays(i), "dd MMM yyyy")
'                                        vjmlbatas_msgallocate += 1
'                                        jml_pjgmsg = 0
'                                    End If

'                                Else
'                                    If vmsg_allocate.Substring(vmsg_allocate.Length - 1, 1) <> "." Then
'                                        vmsg_allocate &= vbNewLine & "many more.."
'                                    End If
'                                End If
'                            End If
'                        Next
'                    Else
'                        For i As Integer = 0 To DateDiff(DateInterval.Day, CDate(.Cells(Me.colTglDari.Index).Value), CDate(.Cells(Me.colTglSampai.Index).Value))
'                            counter = 0
'                            vQueryInsert &= "INSERT INTO tbltrans_mo_workshift_allocate (`date`,`primary work unit`,`primary workshift`,`created by`,`created date`) VALUES "

'                            If Me.ObjTM.apakahschedule_allocatebenar(.Cells(Me.colPrimaryWU.Index).Value, Format(CDate(.Cells(Me.colTglDari.Index).Value).AddDays(i), "yyyy-MM-dd hh:mm:ss"), primaryWorkShift, True) = True Then
'                                If counter < vMaxQuery Then ''SKALA per TANGGAL
'                                    counter += 1
'                                Else
'                                    counter = 1
'                                    vQueryInsert &= "INSERT INTO tbltrans_mo_workshift_allocate (`date`,`primary work unit`,`primary workshift`,`created by`,`created date`) VALUES "
'                                End If
'                                vQueryInsert &= "('" & Format(CDate(.Cells(Me.colTglDari.Index).Value).AddDays(i), "yyyy-MM-dd") & "'," & .Cells(Me.colPrimaryWU.Index).Value & "," & primaryWorkShift.ToString & ",'" & Me.UserCode & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "'),"
'                                If counter = vMaxQuery Then
'                                    vQueryInsert = vQueryInsert.Substring(0, vQueryInsert.Length - 1)
'                                    vQueryInsert &= ";" & vbNewLine
'                                End If
'                            Else
'                                'jika data tidak benar
'                                If vjmlbatas_msgallocate < 5 Then
'                                    If vlist_codewu.Contains(.Cells(Me.colWorkUnitCode.Index).Value) = True Then

'                                        If jml_pjgmsg < 5 Then
'                                            vmsg_allocate &= ", " & Format(CDate(.Cells(Me.colTglDari.Index).Value).AddDays(i), "dd MMM yyyy")
'                                            jml_pjgmsg += 1
'                                        Else
'                                            If vmsg_allocate.Substring(vmsg_allocate.Length - 1, 1) <> "," Then
'                                                vmsg_allocate &= " etc.,"
'                                                vlist_codewu2.Add(.Cells(Me.colWorkUnitCode.Index).Value)
'                                            End If

'                                        End If

'                                    Else
'                                        vlist_codewu.Add(.Cells(Me.colWorkUnitCode.Index).Value)
'                                        vmsg_allocate &= vbNewLine & .Cells(Me.colWorkUnitCode.Index).Value & " :  " & Format(CDate(.Cells(Me.colTglDari.Index).Value).AddDays(i), "dd MMM yyyy")
'                                        vjmlbatas_msgallocate += 1
'                                        jml_pjgmsg = 0
'                                    End If
'                                Else
'                                    If vmsg_allocate.Substring(vmsg_allocate.Length - 1, 1) <> "." Then
'                                        vmsg_allocate &= vbNewLine & "many more.."
'                                    End If
'                                End If
'                            End If
'                        Next
'                    End If
'                End With
'            Next

'            If counter < vMaxQuery Then
'                vQueryInsert = vQueryInsert.Substring(0, vQueryInsert.Length - 1)
'                vQueryInsert &= ";" & vbNewLine
'            End If

'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub generatequery_delete(ByVal vprimarywu As Double, ByVal vprimaryws As Double, ByVal vtglmulai As Date, ByVal vtglselesai As Date, ByVal vDeleteQuery As String)
'        vDeleteQuery = ""
'        Dim dummy_date As Date
'        Try
'            For i As Integer = 0 To DateDiff(DateInterval.Day, vtglmulai.Date, vtglselesai.Date)
'                dummy_date = vtglmulai.Date.AddDays(i)
'                If Me.ObjTM.apakahsudahada_displ(vprimarywu, vprimaryws, dummy_date.Date) = False Then
'                    vDeleteQuery &= dummy_date.ToString & " "
'                End If
'            Next
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
'        Try
'            Me.ObjCD.UpdateThreading()
'            Me.ProgressBar1.Visible = True
'            Me.ProgressBar1.Maximum = 100
'            Me.ProgressBar1.Value = 50
'            Me.Cursor = Cursors.WaitCursor
'            Me.ObjTM.vlist_codewu.Clear()
'            Me.ObjTM.vmsg_unallocate = ""
'            Me.ObjTM.vjmlbatas_msgunallocate = 1
'            GenerateQuery(Me.vDeleteQuery, Me.vInsertQuery)
'            Dim pjg As Integer = vInsertQuery.Length
'            If pjg > 130 Then
'                ExecuteQuery(Me.vDeleteQuery, Me.vInsertQuery, False)
'            End If

'            If vmsg_allocate <> "" Then
'                MsgBox("Can't allocate this schedule already exist." & vbNewLine & vmsg_allocate, MsgBoxStyle.Information, "Information")
'            End If

'            Me.dgvWU.AllowUserToAddRows = False
'            Me.dgvWUG.AllowUserToAddRows = False
'            If Me.dgvWU.RowCount = 0 And Me.dgvWUG.RowCount = 0 Then
'                Me.isi_DGV(Me.TBL_WU, Me.row_WU)
'            Else
'                Me.isi_DGV(Me.TBL_WU_Dummy1, Me.row_WU_Dummy1)
'            End If
'            Me.dgvWU.AllowUserToAddRows = True
'            Me.dgvWUG.AllowUserToAddRows = True

'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        Finally

'        End Try
'    End Sub

'    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
'        Try
'            Me.ProgressBar1.Value = 0
'            Me.ProgressBar1.Visible = False
'            Me.Cursor = Cursors.Default
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
'        Try
'            If Me.ProgressBar1.Visible = False And Me.cbo_shift.SelectedIndex <> -1 And Me.dgvItem.RowCount > 0 Then
'                Me.ProgressBar1.Visible = True
'                Me.BackgroundWorker4.RunWorkerAsync()
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub BackgroundWorker4_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
'        Try
'            Me.ObjCD.UpdateThreading()
'            Me.ProgressBar1.Maximum = 100
'            Me.ProgressBar1.Value = 50
'            Me.Cursor = Cursors.WaitCursor
'            Me.ObjTM.vlist_codewu.Clear()
'            Me.ObjTM.vmsg_unallocate = ""
'            Me.ObjTM.vjmlbatas_msgunallocate = 1

'            GenerateQuery(Me.vDeleteQuery, Me.vInsertQuery)
'            ExecuteQuery(Me.vDeleteQuery, Me.vInsertQuery, True)

'            If Me.ObjTM.vmsg_unallocate <> "" Then
'                MsgBox("Can't unallocate this data referred other transaction." & vbNewLine & ObjTM.vmsg_unallocate, _
'                MsgBoxStyle.Information, "Information")
'            End If

'            Me.dgvWU.AllowUserToAddRows = False
'            Me.dgvWUG.AllowUserToAddRows = False
'            If Me.dgvWU.RowCount = 0 And Me.dgvWUG.RowCount = 0 Then
'                Me.isi_DGV(Me.TBL_WU, Me.row_WU)
'            Else
'                Me.isi_DGV(Me.TBL_WU_Dummy1, Me.row_WU_Dummy1)
'            End If
'            Me.dgvWU.AllowUserToAddRows = True
'            Me.dgvWUG.AllowUserToAddRows = True
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        Finally
'        End Try
'    End Sub

'    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
'        Try
'            Me.ProgressBar1.Value = 0
'            Me.ProgressBar1.Visible = False
'            Me.Cursor = Cursors.Default
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
'        Try
'            If Me.ProgressBar1.Visible = False And Me.allocate_selesai = True And Me.Cursor = Cursors.Default Then
'                Me.ProgressBar1.Visible = True
'                Me.Cursor = Cursors.WaitCursor
'                Me.BackgroundWorker5.RunWorkerAsync()
'            End If

'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub BackgroundWorker5_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker5.DoWork
'        'Try
'        '    Me.ObjCD.UpdateThreading()
'        '    Me.ProgressBar1.Maximum = 100
'        '    Me.ProgressBar1.Value = 50
'        '    Me.Cursor = Cursors.WaitCursor
'        '    If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
'        '        MsgBox("Period from should not be greater than period to.", MsgBoxStyle.Information, "Information")
'        '        Me.dgvItem.Rows.Clear()
'        '        Me.tblInsertDB.Clear()
'        '        Me.TBL_WU_Dummy1.Clear()
'        '        Me.tblLog.Rows.Clear()
'        '        Me.tblLog_.Rows.Clear()
'        '        Me.cbo_shift.SelectedIndex = -1
'        '        Me.cbo_shift.DataSource = Nothing
'        '        Me.cbo_shift.Items.Clear()
'        '        Me.Cursor = Cursors.Default
'        '        Exit Sub
'        '    End If

'        '    FilterData_WorkUnit()
'        '    Me.dgvWU.AllowUserToAddRows = False
'        '    Me.dgvWUG.AllowUserToAddRows = False
'        '    If Me.dgvWU.RowCount = 0 And Me.dgvWUG.RowCount = 0 Then
'        '        Me.isi_DGV(Me.TBL_WU, Me.row_WU)
'        '    Else
'        '        Me.isi_DGV(Me.TBL_WU_Dummy1, Me.row_WU_Dummy1)
'        '    End If

'        '    Me.iscbo_cheklist()
'        '    Me.dgvWU.AllowUserToAddRows = True
'        '    Me.dgvWUG.AllowUserToAddRows = True
'        'Catch ex As Exception
'        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        'Finally
'        'End Try
'    End Sub

'    Private Sub BackgroundWorker5_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted
'        Try
'            Me.ObjCD.UpdateThreading()
'            Me.ProgressBar1.Maximum = 100
'            Me.ProgressBar1.Value = 50
'            Me.Cursor = Cursors.WaitCursor
'            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
'                MsgBox("Period from should not be greater than period to.", MsgBoxStyle.Information, "Information")
'                Me.dgvItem.Rows.Clear()
'                Me.tblInsertDB.Clear()
'                Me.TBL_WU_Dummy1.Clear()
'                Me.tblLog.Rows.Clear()
'                Me.tblLog_.Rows.Clear()
'                Me.cbo_shift.SelectedIndex = -1
'                Me.cbo_shift.DataSource = Nothing
'                Me.cbo_shift.Items.Clear()
'                Me.Cursor = Cursors.Default
'                Exit Sub
'            End If

'            FilterData_WorkUnit()
'            Me.dgvWU.AllowUserToAddRows = False
'            Me.dgvWUG.AllowUserToAddRows = False
'            If Me.dgvWU.RowCount = 0 And Me.dgvWUG.RowCount = 0 Then
'                Me.isi_DGV(Me.TBL_WU, Me.row_WU)
'            Else
'                Me.isi_DGV(Me.TBL_WU_Dummy1, Me.row_WU_Dummy1)
'            End If

'            Me.iscbo_cheklist()
'            Me.dgvWU.AllowUserToAddRows = True
'            Me.dgvWUG.AllowUserToAddRows = True
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        Finally
'        End Try

'        Try
'            Me.ProgressBar1.Value = 0
'            Me.ProgressBar1.Visible = False
'            Me.Cursor = Cursors.Default
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Public Function isillb_wug(ByVal vprimary_wug As Double) As Boolean
'        Try
'            primary_wug = vprimary_wug
'            TBL_.Rows.Clear()
'            TBL_WU.DefaultView.Sort = "[work unit code] ASC"
'            TBL_WU.DefaultView.RowFilter = "[primary work unit group]=" & vprimary_wug
'            For i As Integer = 0 To TBL_WU.DefaultView.Count - 1
'                row_WU = TBL_WU.DefaultView(i).Row
'                TBL_.DefaultView.Sort = "[work unit code] ASC"
'                TBL_.DefaultView.RowFilter = "[primary]=" & row_WU.primary
'                With row
'                    If TBL_.DefaultView.Count = 0 Then
'                        row_ = TBL_.Newtblmaster_workunitRow()
'                        row_.primary = row_WU.primary
'                        row_.primary_work_unit_group = row_WU.primary_work_unit_group
'                        row_.work_unit_code = row_WU.work_unit_code
'                        row_.description = row_WU.description
'                        TBL_.Addtblmaster_workunitRow(row_)
'                    End If
'                End With
'            Next

'            TBL_WU.DefaultView.RowFilter = Nothing
'            TBL_.DefaultView.RowFilter = Nothing
'            Me.ClbWU.DataSource = TBL_.DefaultView
'            Me.ClbWU.DisplayMember = "work unit code"
'            Isi_chelist()
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Function

'    Private Sub ClbWU_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles ClbWU.ItemCheck
'        Try
'            If ClbWU.GetItemCheckState(e.Index) = CheckState.Checked Then
'                Me.remove_cheklist()
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub ClbWU_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClbWU.Leave
'        Try
'            Me.isitbldummywu_byprimarywug()
'            Me.Isi_chelist()
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Public Function isitbldummywu_byprimarywug() As Boolean
'        Try
'            For Each DV As DataRowView In ClbWU.CheckedItems
'                row_WU = DV.Row
'                TBL_WU_Dummy.DefaultView.RowFilter = "[primary]=" & row_WU.primary
'                If TBL_WU_Dummy.DefaultView.Count = 0 Then
'                    row_WU_Dummy = TBL_WU_Dummy.Newtblmaster_workunitRow()
'                    row_WU_Dummy.primary = row_WU.primary
'                    row_WU_Dummy.primary_work_unit_group = row_WU.primary_work_unit_group
'                    row_WU_Dummy.work_unit_code = row_WU.work_unit_code
'                    row_WU_Dummy.description = row_WU.description
'                    TBL_WU_Dummy.Addtblmaster_workunitRow(row_WU_Dummy)
'                End If
'            Next
'            Me.TBL_WU_Dummy.DefaultView.RowFilter = Nothing
'            Return True
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Function

'    Public Function Isi_chelist() As Boolean
'        Try
'            TBL_.DefaultView.RowFilter = "[primary work unit group]=" & primary_wug
'            If TBL_.DefaultView.Count > 0 Then
'                For i As Integer = 0 To TBL_.DefaultView.Count - 1
'                    row_ = TBL_.DefaultView(i).Row
'                    TBL_WU_Dummy.DefaultView.RowFilter = "[primary]=" & row_.primary
'                    For ij As Integer = 0 To TBL_WU_Dummy.DefaultView.Count - 1
'                        row_WU_Dummy = TBL_WU_Dummy.DefaultView(ij).Row
'                        If row_.primary = row_WU_Dummy.primary Then
'                            Me.ClbWU.SetItemCheckState(i, CheckState.Checked)
'                        Else
'                            Me.ClbWU.SetItemCheckState(i, CheckState.Unchecked)
'                        End If
'                    Next
'                Next
'            End If
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Function

'    Public Sub remove_cheklist()
'        TBL_.DefaultView.RowFilter = Nothing
'        For i As Integer = 0 To TBL_.DefaultView.Count - 1
'            row_ = TBL_.DefaultView(i).Row
'            For ij As Integer = TBL_WU_Dummy.Rows.Count - 1 To 0 Step -1
'                row_WU_Dummy = TBL_WU_Dummy.Rows(ij)
'                If row_.primary = row_WU_Dummy.primary Then
'                    TBL_WU_Dummy.Rows.Remove(row_WU_Dummy)
'                End If
'            Next
'        Next
'    End Sub

'    Private Sub dgvWUG_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvWUG.UserDeletedRow
'        Try
'            Me.ObjCD.SetNumber(Me.dgvWUG, False)
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub

'    Private Sub dgvWUG_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvWUG.UserDeletingRow
'        Dim vprimary_wug As Double
'        Try
'            vprimary_wug = Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.Row.Index).Value
'            TBL_.DefaultView.RowFilter = Nothing
'            TBL_.DefaultView.RowFilter = "[primary work unit group]=" & vprimary_wug.ToString
'            For i As Integer = 0 To TBL_.DefaultView.Count - 1
'                row_ = TBL_.DefaultView(i).Row
'                For ij As Integer = TBL_WU_Dummy.Rows.Count - 1 To 0 Step -1
'                    row_WU_Dummy = TBL_WU_Dummy.Rows(ij)
'                    If row_.primary = row_WU_Dummy.primary Then
'                        TBL_WU_Dummy.Rows.Remove(row_WU_Dummy)
'                    End If
'                Next
'            Next
'        Catch ex As Exception
'            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
'        End Try
'    End Sub
'End Class
'#End Region


