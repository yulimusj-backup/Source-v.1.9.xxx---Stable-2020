Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_NewSch4

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public Baru As Boolean = False
    Public LoadFromSch As Boolean = False
    Public PrimaryConfig As Double = -1
    Public PrimaryProduct As Double = -1
    Public IdentifierJO As String = ""

    Public TblWorkUnitSch As DseTrans_ProductionSchedule_NewSch3.DataTable9DataTable
    Public RowWorkUnitSch As DseTrans_ProductionSchedule_NewSch3.DataTable9Row

    Dim mDecimal As Integer = 0
    Dim ObjCd As New Core_Module.ClsCekData
    Dim tblProduct As New DseTrans_ProductionSchedule_NewSch3.DataTable7DataTable
    Dim rowProduct As DseTrans_ProductionSchedule_NewSch3.DataTable7Row
    Dim tblProcess As New DseTrans_ProductionSchedule_NewSch3.DataTable5DataTable
    Dim rowProcess As DseTrans_ProductionSchedule_NewSch3.DataTable5Row
    Dim tblWorkUnit As New DseTrans_ProductionSchedule_NewSch3.DataTable6DataTable
    Dim rowWorkUnit As DseTrans_ProductionSchedule_NewSch3.DataTable6Row
    Dim tblHeader As New DseTrans_ProductionSchedule_NewSch3.DataTable1DataTable
    Dim rowHeader As DseTrans_ProductionSchedule_NewSch3.DataTable1Row
    Dim Buffer As Object
    Dim objProduct As New Core_Module.clsMasterProduct
    Dim objProcess As New Core_Module.ClsMasterProses
    Dim CancelClose As Boolean = False

    Function ApakahDataBenar() As Boolean
        Dim dta As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable1TableAdapter
        Dim value As New Nullable(Of Long)
        Try
            If Me.txtCode.Text.Trim.Length <= 0 Then
                MsgBox("Predefined setting code cannot be an empty string.", MsgBoxStyle.Information, "Information")
                Me.txtCode.Focus()
                Return False
            End If
            dta.Connection.ConnectionString = Core_Module.GetConstring
            value = dta.ScalarQuery_GetJmlByCodeAndNotPrimary(Me.txtCode.Text, Me.PrimaryConfig, Me.PrimaryProduct)
            If value.HasValue() Then
                If value.Value > 0 Then
                    MsgBox("Predefined setting code " & Me.txtCode.Text & " already exist for code no " & Me.objProduct.GetProductCode(Me.PrimaryProduct) & ".", MsgBoxStyle.Information, "Information")
                    Me.txtCode.Focus()
                    Return False
                End If
            End If
            For Each dr As DataGridViewRow In Me.DgvWUAllocate_Process.Rows
                Me.tblWorkUnit.DefaultView.RowFilter = "[identifier process] = '" & dr.Cells(Me.ColAll_IdentifierProcess.Index).Value & "'"
                If Me.tblWorkUnit.DefaultView.Count <= 0 Then
                    MsgBox("Please select at least 1 work unit for process " & dr.Cells(Me.ColAll_Process_Code.Index).Value & " code no " & dr.Cells(Me.ColAll_Process_CodeNo.Index).Value, MsgBoxStyle.Information, "Information")
                    Me.DgvWUAllocate_Process.ClearSelection()
                    dr.Selected = True
                    Return False
                End If
            Next

            Dim listOfInActiveWorkUnit As New List(Of String)
            Dim InActiveWorkUnit As String = ""
            Dim objWorkUnit As New Core_Module.ClsMasterWorkUnit
            For Each rowWorkUnit In tblWorkUnit
                If objWorkUnit.ApakahSudahAdaOnlyActive(rowWorkUnit.work_unit_code) = False Then
                    If listOfInActiveWorkUnit.Contains(rowWorkUnit.work_unit_code) = False Then
                        listOfInActiveWorkUnit.Add(rowWorkUnit.work_unit_code)
                    End If
                End If
            Next
            If listOfInActiveWorkUnit.Count > 0 Then
                For Each wu As String In listOfInActiveWorkUnit
                    InActiveWorkUnit += vbNewLine & wu
                Next
                MsgBox("The following work units is not active, please check again" & InActiveWorkUnit, MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Sub FillBoM(ByVal vPrimaryProduct As Double, ByRef vTbl As DseTrans_ProductionSchedule_NewSch3.DataTable8DataTable)
        Dim dtaProductBoM As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable8TableAdapter
        Dim tblProductBoM As New DseTrans_ProductionSchedule_NewSch3.DataTable8DataTable
        Dim rowProductBoM As DseTrans_ProductionSchedule_NewSch3.DataTable8Row
        Dim vRow As DseTrans_ProductionSchedule_NewSch3.DataTable8Row
        Try
            dtaProductBoM.Connection.ConnectionString = Core_Module.GetConstring
            dtaProductBoM.Fill(tblProductBoM, vPrimaryProduct)
            For Each rowProductBoM In tblProductBoM
                vRow = vTbl.NewDataTable8Row
                For i As Integer = 0 To vTbl.Columns.Count - 1
                    vRow.Item(i) = rowProductBoM.Item(i)
                Next
                vTbl.Rows.Add(vRow)
                Me.FillBoM(rowProductBoM.primary_product_raw, vTbl)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dtaProductBoM.Dispose()
        End Try
    End Sub

    Function GetIndex(ByVal vPrimaryIndex As Double) As Integer
        Try
            For Each row As DataGridViewRow In Me.DgvWUAllocate_WU.Rows
                If vPrimaryIndex = row.Cells(Me.CoAll_WU_PrimaryWU.Index).Value Then
                    Return row.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Function GetPrimary(ByVal vCode As String, ByVal vPrimaryProductOutput As Double) As Double
        Dim dta As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable1TableAdapter
        Dim val As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            val = dta.ScalarQuery_GetPrimary(vCode, vPrimaryProductOutput)
            If val.HasValue Then
                Return val.Value
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Function GetPrimary(ByVal vCode As String, ByVal vPrimaryProductOutput As Double, ByVal vPrimaryMasterLine As Double) As Double
        Dim dta As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable1TableAdapter
        Dim val As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            val = dta.ScalarQuery_GetPrimary2(vCode, vPrimaryProductOutput, vPrimaryMasterLine)
            If val.HasValue Then
                Return val.Value
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Function GetPrimaryProcess(ByVal vPrimaryProduct As Double) As Double
        Dim dta As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable8TableAdapter
        Dim val As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            val = dta.ScalarQuery_GetPrimaryProcess(vPrimaryProduct)
            If val.HasValue Then
                Return val.Value
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Sub ReSortTblProcess(ByRef vTblProcess As DseTrans_ProductionSchedule_NewSch3.DataTable5DataTable)
        Dim tblProcessDummy As New DseTrans_ProductionSchedule_NewSch3.DataTable5DataTable
        Dim rowProcessDummy As DseTrans_ProductionSchedule_NewSch3.DataTable5Row
        Try
            For i As Integer = vTblProcess.Rows.Count - 1 To 0 Step -1
                rowProcess = vTblProcess.Rows(i)
                rowProcessDummy = tblProcessDummy.NewDataTable5Row
                For i2 As Integer = 0 To vTblProcess.Columns.Count - 1
                    rowProcessDummy.Item(i2) = rowProcess.Item(i2)
                Next
                tblProcessDummy.Rows.Add(rowProcessDummy)
            Next
            vTblProcess.Rows.Clear()
            For Each rowProcessDummy In tblProcessDummy
                rowProcess = vTblProcess.NewDataTable5Row
                For i As Integer = 0 To vTblProcess.Columns.Count - 1
                    rowProcess.Item(i) = rowProcessDummy.Item(i)
                Next
                vTblProcess.Rows.Add(rowProcess)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub LoadNewConfig(ByVal vPrimaryProduct As Double)
        Dim DTAItemWO As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo1TableAdapter
        Dim vTBLItemWO As New DseTrans_Schedule.tbltrans_mo_item_wo1DataTable
        Dim vRowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim dtaProduct As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable7TableAdapter
        Dim tblProductBoM As New DseTrans_ProductionSchedule_NewSch3.DataTable8DataTable
        Dim Identifier As Integer = 0
        Dim PrimaryProcess As Double = -1
        Dim dr As DataGridViewRow
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvWUAllocate_WU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            tblProduct.Rows.Clear()
            dtaProduct.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemWO.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemWO.Fill2(vTBLItemWO, Me.IdentifierJO) '2017-08-22: diubah ama iwan, include semua wo dalam jo ybs, walaupun wo-nya qty produce = 0, supaya bisa berlaku umum untuk produk ini
            For Each vRowItemWO In vTBLItemWO
                Me.rowProcess = Me.tblProcess.NewDataTable5Row
                With Me.rowProcess
                    .primary_master = -1
                    .primary_process = vRowItemWO.Primary_Process
                    .primary_product = vRowItemWO.Primary_Product
                    Identifier += 1
                    .identifier = "-1-" & Identifier.ToString
                    .Process_Code = Me.objProcess.GetCode(.primary_process)
                    .Product_Code = Me.objProduct.GetProductCode(.primary_product)
                    .Product_Desc = Me.objProduct.GetDescription(.primary_product)
                    ._Decimal = Me.objProduct.GetAngkaDecimal(.primary_product)
                End With
                Me.tblProcess.Rows.Add(Me.rowProcess)
            Next

            Dim listOfPrimaryProduct As New List(Of Double)
            For Each rowProcess In tblProcess
                dr = DgvWUAllocate_Process.Rows(DgvWUAllocate_Process.Rows.Add())
                dr.Cells(Me.ColAll_Process_Code.Index).Value = rowProcess.Process_Code
                dr.Cells(Me.ColAll_Process_CodeNo.Index).Value = rowProcess.Product_Code
                dr.Cells(Me.ColAll_Process_Desc.Index).Value = rowProcess.Product_Desc
                dr.Cells(Me.ColAll_Process_No.Index).Value = Me.DgvWUAllocate_Process.Rows.Count
                dr.Cells(Me.ColAll_Process_PrimaryCodeNo.Index).Value = rowProcess.primary_product
                dr.Cells(Me.ColAll_Process_PrimaryProcess.Index).Value = rowProcess.primary_process
                dr.Cells(Me.ColAll_IdentifierProcess.Index).Value = rowProcess.identifier

                If listOfPrimaryProduct.Contains(rowProcess.primary_product) = False Then
                    dtaProduct.ClearBeforeFill = False
                    dtaProduct.Fill(tblProduct, rowProcess.primary_product)
                    dtaProduct.ClearBeforeFill = True
                    listOfPrimaryProduct.Add(rowProcess.primary_product)
                End If

            Next
            tblProduct.DefaultView.RowFilter = "[primary product] = " & Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_PrimaryCodeNo.Index, 0).Value.ToString
            If tblProduct.DefaultView.Count > 0 Then
                For i As Integer = 0 To tblProduct.DefaultView.Count - 1
                    Me.rowProduct = Me.tblProduct.DefaultView(i).Row
                    dr = DgvWUAllocate_WU.Rows(Me.DgvWUAllocate_WU.Rows.Add)
                    dr.Cells(Me.CoAll_WU_Code.Index).Value = rowProduct.work_unit_code
                    dr.Cells(Me.CoAll_WU_Desc.Index).Value = rowProduct.description
                    dr.Cells(Me.CoAll_WU_No.Index).Value = Me.ObjCd.DigitGrouping(Me.DgvWUAllocate_WU.Rows.Count, 0)
                    dr.Cells(Me.CoAll_WU_ProdRate.Index).Value = Me.ObjCd.DigitGrouping(rowProduct.rate_per_hour, rowProcess._Decimal)
                    dr.Cells(Me.CoAll_WU_Efficiency.Index).Value = Me.ObjCd.DigitGrouping(rowProduct.efficiency, 2)
                    dr.Cells(Me.CoAll_WU_PlanRate.Index).Value = Me.ObjCd.DigitGrouping(rowProduct.rate_per_hour * rowProduct.efficiency / 100, rowProcess._Decimal)
                    dr.Cells(Me.CoAll_WU_PrimaryWU.Index).Value = rowProduct.primary_work_unit
                    dr.Cells(Me.CoAll_WU_Allocate.Index).Value = False
                Next
            End If
            Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvWUAllocate_WU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub LoadSavedConfig(ByVal vPrimaryConfig As Double)
        Dim dtaProduct As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable7TableAdapter
        Dim dtaHeader As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable1TableAdapter
        Dim dtaProcess As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable5TableAdapter
        Dim dtaWorkUnit As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable6TableAdapter
        Dim dr As DataGridViewRow
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvWUAllocate_WU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dtaProduct.Connection.ConnectionString = Core_Module.GetConstring
            dtaHeader.Connection.ConnectionString = Core_Module.GetConstring
            tblProduct.Rows.Clear()
            If dtaHeader.FillBy_Primary(tblHeader, vPrimaryConfig) > 0 Then
                rowHeader = tblHeader.Rows(0)
                Me.txtCode.Text = rowHeader.code
                Me.txtPrimary.Text = vPrimaryConfig
                dtaProcess.Connection.ConnectionString = Core_Module.GetConstring
                Dim listOfPrimaryProduct As New List(Of Double)
                If dtaProcess.FillBy_Primary(tblProcess, vPrimaryConfig) > 0 Then
                    For Each rowProcess In tblProcess
                        dr = DgvWUAllocate_Process.Rows(DgvWUAllocate_Process.Rows.Add())
                        dr.Cells(Me.ColAll_Process_Code.Index).Value = rowProcess.Process_Code
                        dr.Cells(Me.ColAll_Process_CodeNo.Index).Value = rowProcess.Product_Code
                        dr.Cells(Me.ColAll_Process_Desc.Index).Value = rowProcess.Product_Desc
                        dr.Cells(Me.ColAll_Process_No.Index).Value = Me.DgvWUAllocate_Process.Rows.Count
                        dr.Cells(Me.ColAll_Process_PrimaryCodeNo.Index).Value = rowProcess.primary_product
                        dr.Cells(Me.ColAll_Process_PrimaryProcess.Index).Value = rowProcess.primary_process
                        dr.Cells(Me.ColAll_IdentifierProcess.Index).Value = rowProcess.identifier
                        If listOfPrimaryProduct.Contains(rowProcess.primary_product) = False Then
                            listOfPrimaryProduct.Add(rowProcess.primary_product)
                            dtaProduct.ClearBeforeFill = False
                            dtaProduct.Fill(tblProduct, rowProcess.primary_product)
                            dtaProduct.ClearBeforeFill = True
                        End If
                    Next
                    tblProduct.DefaultView.RowFilter = "[primary product] = " & Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_PrimaryCodeNo.Index, 0).Value.ToString
                    If tblProduct.DefaultView.Count > 0 Then
                        For i As Integer = 0 To tblProduct.DefaultView.Count - 1
                            Me.rowProduct = Me.tblProduct.DefaultView(i).Row
                            dr = DgvWUAllocate_WU.Rows(Me.DgvWUAllocate_WU.Rows.Add)
                            dr.Cells(Me.CoAll_WU_Code.Index).Value = rowProduct.work_unit_code
                            dr.Cells(Me.CoAll_WU_Desc.Index).Value = rowProduct.description
                            dr.Cells(Me.CoAll_WU_No.Index).Value = Me.ObjCd.DigitGrouping(Me.DgvWUAllocate_WU.Rows.Count, 0)
                            dr.Cells(Me.CoAll_WU_ProdRate.Index).Value = Me.ObjCd.DigitGrouping(rowProduct.rate_per_hour, rowProcess._Decimal)
                            dr.Cells(Me.CoAll_WU_Efficiency.Index).Value = Me.ObjCd.DigitGrouping(rowProduct.efficiency, 2)
                            dr.Cells(Me.CoAll_WU_PlanRate.Index).Value = Me.ObjCd.DigitGrouping(rowProduct.rate_per_hour * rowProduct.efficiency / 100, rowProcess._Decimal)
                            dr.Cells(Me.CoAll_WU_PrimaryWU.Index).Value = rowProduct.primary_work_unit
                            dr.Cells(Me.CoAll_WU_Allocate.Index).Value = False
                        Next
                    End If
                    dtaWorkUnit.Connection.ConnectionString = Core_Module.GetConstring
                    If dtaWorkUnit.FillBy_Primary(tblWorkUnit, vPrimaryConfig) > 0 Then
                        Me.tblWorkUnit.DefaultView.RowFilter = "[primary master] = " & Me.rowHeader.primary.ToString & " AND [identifier process] = '" & Me.DgvWUAllocate_Process.Item(Me.ColAll_IdentifierProcess.Index, 0).Value.ToString & "'"
                        Me.tblWorkUnit.DefaultView.Sort = "priority"
                        Dim index As Integer = -1
                        For i As Integer = 0 To tblWorkUnit.DefaultView.Count - 1
                            rowWorkUnit = tblWorkUnit.DefaultView(i).Row
                            index = Me.GetIndex(rowWorkUnit.primary_work_unit)
                            If index = -1 Then
                                dr = DgvWUAllocate_WU.Rows(Me.DgvWUAllocate_WU.Rows.Add)
                                dr.Cells(Me.CoAll_WU_Code.Index).Value = rowWorkUnit.work_unit_code
                                dr.Cells(Me.CoAll_WU_Desc.Index).Value = rowWorkUnit.description
                                dr.Cells(Me.CoAll_WU_No.Index).Value = Me.ObjCd.DigitGrouping(Me.DgvWUAllocate_WU.Rows.Count, 0)
                                dr.Cells(Me.CoAll_WU_ProdRate.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.rate, rowProcess._Decimal)
                                dr.Cells(Me.CoAll_WU_Efficiency.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.effeciency, 2)
                                dr.Cells(Me.CoAll_WU_PlanRate.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.plan_rate, rowProcess._Decimal)
                                dr.Cells(Me.CoAll_WU_PrimaryWU.Index).Value = rowWorkUnit.primary_work_unit
                                dr.Cells(Me.CoAll_WU_Allocate.Index).Value = True
                            Else
                                dr = Me.DgvWUAllocate_WU.Rows(index)
                                dr.Cells(Me.CoAll_WU_Allocate.Index).Value = True
                                dr.Cells(Me.CoAll_WU_ProdRate.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.rate, rowProcess._Decimal)
                                dr.Cells(Me.CoAll_WU_Efficiency.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.effeciency, 2)
                                dr.Cells(Me.CoAll_WU_PlanRate.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.plan_rate, rowProcess._Decimal)
                            End If
                        Next
                    End If
                End If
            Else
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
            Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvWUAllocate_WU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub Save()
        Dim dtaHeader As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable1TableAdapter
        Dim dtaItemProcess As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable2TableAdapter
        Dim dtaItemWorkUnit As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable3TableAdapter
        Try
            dtaHeader.Connection.ConnectionString = Core_Module.GetConstring
            dtaItemProcess.Connection.ConnectionString = Core_Module.GetConstring
            dtaItemWorkUnit.Connection.ConnectionString = Core_Module.GetConstring
            dtaHeader.InsertQuery(Me.txtCode.Text.Trim, Me.PrimaryProduct, Me.UserCode, Now, Me.UserCode, Now)
            Me.PrimaryConfig = Me.GetPrimary(Me.txtCode.Text, Me.PrimaryProduct)

            For Each rowProcess In Me.tblProcess
                If rowProcess.identifier.StartsWith("-1") Then
                    rowProcess.identifier = Me.PrimaryConfig.ToString & rowProcess.identifier.Substring(2, rowProcess.identifier.Length - 2)
                End If
                dtaItemProcess.InsertQuery(Me.PrimaryConfig, rowProcess.primary_process, rowProcess.primary_product, rowProcess.identifier)
            Next
            For Each rowWorkUnit In tblWorkUnit
                If rowWorkUnit.identifier_process.StartsWith("-1") Then
                    rowWorkUnit.identifier_process = Me.PrimaryConfig.ToString & rowWorkUnit.identifier_process.Substring(2, rowWorkUnit.identifier_process.Length - 2)
                End If
                dtaItemWorkUnit.InsertQuery(Me.PrimaryConfig, rowWorkUnit.identifier_process, rowWorkUnit.priority, rowWorkUnit.primary_work_unit, _
                         Me.ObjCd.DigitGrouping(rowWorkUnit.rate, Me.mDecimal), Me.ObjCd.DigitGrouping(rowWorkUnit.effeciency, 2), Me.ObjCd.DigitGrouping(rowWorkUnit.plan_rate, Me.mDecimal))
            Next
            MsgBox("New Configuration saved.", MsgBoxStyle.Information, "Information")
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_NewSch4_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.CancelClose = True Then
                Me.CancelClose = False
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_NewSch4_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DgvWUAllocate_WU.ReadOnly = False
            Me.SplitContainer1.TabStop = False
            If Me.Baru Then
                Me.Text = "Add new work unit configuration for: " & Me.objProduct.GetProductCode(Me.PrimaryProduct)
                If Me.LoadFromSch = False Then
                    Me.LoadNewConfig(Me.PrimaryProduct)
                Else

                End If
            Else
                Me.Text = "Edit work unit configuration for: " & Me.objProduct.GetProductCode(Me.PrimaryProduct)
                If Me.LoadFromSch = False Then
                    Me.LoadSavedConfig(Me.PrimaryConfig)
                Else

                End If
            End If
            Me.DgvWUAllocate_Process.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvWUAllocate_WU.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.mDecimal = Me.objProduct.GetAngkaDecimal(Me.PrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWUAllocate_WU_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvWUAllocate_WU.CellBeginEdit
        Try
            If e.ColumnIndex = Me.CoAll_WU_Allocate.Index Or e.ColumnIndex = Me.CoAll_WU_ProdRate.Index Or e.ColumnIndex = Me.CoAll_WU_Efficiency.Index Then
                If e.ColumnIndex <> Me.CoAll_WU_Allocate.Index And Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_Allocate.Index, e.RowIndex).Value = False Then
                    e.Cancel = True
                    Exit Sub
                End If
                Me.Buffer = Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWUAllocate_WU_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvWUAllocate_WU.CellEndEdit
        Try
            If Not IsNothing(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) Then
                If String.IsNullOrEmpty(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.CoAll_WU_ProdRate.Index Then
                        If IsNumeric(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If CDec(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                            Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCd.DigitGrouping(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value, Me.mDecimal)
                        Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PlanRate.Index, e.RowIndex).Value = Me.ObjCd.DigitGrouping(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_ProdRate.Index, e.RowIndex).Value * Me.DgvWUAllocate_WU(Me.CoAll_WU_Efficiency.Index, e.RowIndex).Value / 100, Me.mDecimal)
                        Me.tblWorkUnit.DefaultView.RowFilter = "[primary work unit] = " & Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryWU.Index, e.RowIndex).Value & " and [identifier process] = '" & Me.DgvWUAllocate_Process.Item(Me.ColAll_IdentifierProcess.Index, Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex).Value.ToString & "'"
                        If Me.tblWorkUnit.DefaultView.Count > 0 Then
                            Me.rowWorkUnit = Me.tblWorkUnit.DefaultView(0).Row
                            Me.rowWorkUnit.rate = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_ProdRate.Index, e.RowIndex).Value
                            Me.rowWorkUnit.plan_rate = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PlanRate.Index, e.RowIndex).Value
                        End If
                    End If
                    If e.ColumnIndex = Me.CoAll_WU_Efficiency.Index Then
                        If IsNumeric(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCd.DigitGrouping(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        If CDec(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Or CDec(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                            Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PlanRate.Index, e.RowIndex).Value = Me.ObjCd.DigitGrouping(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_ProdRate.Index, e.RowIndex).Value * Me.DgvWUAllocate_WU(Me.CoAll_WU_Efficiency.Index, e.RowIndex).Value / 100, Me.mDecimal)
                        Me.tblWorkUnit.DefaultView.RowFilter = "[primary work unit] = " & Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryWU.Index, e.RowIndex).Value & " and [identifier process] = '" & Me.DgvWUAllocate_Process.Item(Me.ColAll_IdentifierProcess.Index, Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex).Value.ToString & "'"
                        If Me.tblWorkUnit.DefaultView.Count > 0 Then
                            Me.rowWorkUnit = Me.tblWorkUnit.DefaultView(0).Row
                            Me.rowWorkUnit.rate = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_ProdRate.Index, e.RowIndex).Value
                            Me.rowWorkUnit.plan_rate = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PlanRate.Index, e.RowIndex).Value
                        End If
                    End If

                    For i As Integer = tblWorkUnit.Rows.Count - 1 To 0 Step -1
                        rowWorkUnit = tblWorkUnit.Rows(i)
                        If rowWorkUnit.identifier_process = Me.DgvWUAllocate_Process.Item(Me.ColAll_IdentifierProcess.Index, Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex).Value.ToString Then
                            tblWorkUnit.Rows.RemoveAt(i)
                        End If
                    Next
                    Dim priority As Integer = 0
                    For Each row As DataGridViewRow In DgvWUAllocate_WU.Rows
                        If row.Cells(Me.CoAll_WU_Allocate.Index).Value = True Then
                            Me.rowWorkUnit = Me.tblWorkUnit.NewDataTable6Row
                            Me.rowWorkUnit.primary_master = Me.PrimaryConfig
                            Me.rowWorkUnit.identifier_process = Me.DgvWUAllocate_Process.Item(Me.ColAll_IdentifierProcess.Index, Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex).Value.ToString
                            Me.rowWorkUnit.primary_work_unit = row.Cells(Me.CoAll_WU_PrimaryWU.Index).Value
                            Me.rowWorkUnit.work_unit_code = row.Cells(Me.CoAll_WU_Code.Index).Value
                            Me.rowWorkUnit.description = row.Cells(Me.CoAll_WU_Desc.Index).Value
                            priority += 1
                            Me.rowWorkUnit.priority = priority
                            Me.rowWorkUnit.rate = row.Cells(Me.CoAll_WU_ProdRate.Index).Value
                            Me.rowWorkUnit.effeciency = row.Cells(Me.CoAll_WU_Efficiency.Index).Value
                            Me.rowWorkUnit.plan_rate = row.Cells(Me.CoAll_WU_PlanRate.Index).Value
                            Me.tblWorkUnit.Rows.Add(Me.rowWorkUnit)
                        End If
                    Next

                Else
                    Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            Else
                Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.ApakahDataBenar() = True Then
                If MsgBox("Save configuration?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    If Me.Baru = False Then
                        Me._Update()
                    Else
                        Me.Save()
                    End If
                End If
                Me.txtCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.GotFocus
        Try
            Me.txtCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.LostFocus
        Try
            Me.txtCode.BackColor = Me.txtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub _Update()
        Dim dtaHeader As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable1TableAdapter
        Dim dtaItemProcess As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable2TableAdapter
        Dim dtaItemWorkUnit As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable3TableAdapter
        Try
            dtaHeader.Connection.ConnectionString = Core_Module.GetConstring
            dtaItemProcess.Connection.ConnectionString = Core_Module.GetConstring
            dtaItemWorkUnit.Connection.ConnectionString = Core_Module.GetConstring
            dtaHeader.UpdateQuery(Me.txtCode.Text.Trim, Me.UserCode, Now, Me.PrimaryConfig)
            dtaItemProcess.DeleteQuery(Me.PrimaryConfig)
            dtaItemWorkUnit.DeleteQuery(Me.PrimaryConfig)
            For Each rowProcess In Me.tblProcess
                If rowProcess.identifier.StartsWith("-1") Then
                    rowProcess.identifier = Me.PrimaryConfig.ToString & rowProcess.identifier.Substring(2, rowProcess.identifier.Length - 2)
                End If
                dtaItemProcess.InsertQuery(Me.PrimaryConfig, rowProcess.primary_process, rowProcess.primary_product, rowProcess.identifier)
            Next
            For Each rowWorkUnit In tblWorkUnit
                If rowWorkUnit.identifier_process.StartsWith("-1") Then
                    rowWorkUnit.identifier_process = Me.PrimaryConfig.ToString & rowWorkUnit.identifier_process.Substring(2, rowWorkUnit.identifier_process.Length - 2)
                End If
                dtaItemWorkUnit.InsertQuery(Me.PrimaryConfig, rowWorkUnit.identifier_process, rowWorkUnit.priority, rowWorkUnit.primary_work_unit, _
                        Me.ObjCd.DigitGrouping(rowWorkUnit.rate, Me.mDecimal), Me.ObjCd.DigitGrouping(rowWorkUnit.effeciency, 2), Me.ObjCd.DigitGrouping(rowWorkUnit.plan_rate, Me.mDecimal))
            Next
            MsgBox("Configuration updated.", MsgBoxStyle.Information, "Information")
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Are you sure to cancel and discard data changes?", MsgBoxStyle.YesNo, "Conformation") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            Else
                CancelClose = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWUAllocate_Process_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvWUAllocate_Process.CellContentClick
        'Dim listOfPrimaryWorkUnit As New List(Of Double)
        Dim PrimaryProcess As Double = -1
        Dim objWU As New Core_Module.ClsMasterWorkUnit
        Dim counter As Integer = 0
        Dim TBLDummy As New Inventory_Module.DseTrans_StockBundling1.DataTable1DataTable
        Dim rowDummy As Inventory_Module.DseTrans_StockBundling1.DataTable1Row

        Try
            If e.ColumnIndex = Me.ColAll_SetRef.Index Then
                For Each r As DataGridViewRow In Me.DgvWUAllocate_WU.Rows
                    If r.Cells(Me.CoAll_WU_Allocate.Index).Value = Boolean.TrueString Then
                        'listOfPrimaryWorkUnit.Add(r.Cells(Me.CoAll_WU_PrimaryWU.Index).Value) ''catat daftar WU yang dijadikan ref
                        rowDummy = TBLDummy.NewDataTable1Row

                        rowDummy.efficiency = r.Cells(Me.CoAll_WU_Efficiency.Index).Value
                        rowDummy.plan_rate = r.Cells(Me.CoAll_WU_PlanRate.Index).Value
                        rowDummy.primary_work_unit = r.Cells(Me.CoAll_WU_PrimaryWU.Index).Value
                        rowDummy.rate = r.Cells(Me.CoAll_WU_ProdRate.Index).Value

                        TBLDummy.Rows.Add(rowDummy)
                    End If
                Next
                If TBLDummy.Rows.Count = 0 Then
                    MsgBox("Process as reference must have at least 1 work unit selected.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                'If listOfPrimaryWorkUnit.Count = 0 Then
                '    MsgBox("Process as reference must have at least 1 work unit selected.", MsgBoxStyle.Information, "Information")
                '    Exit Sub
                'End If
                If MsgBox("Are you sure to update other product with " & Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_CodeNo.Index, e.RowIndex).Value & " as reference?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If
                Me.Cursor = Cursors.WaitCursor
                PrimaryProcess = Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_PrimaryProcess.Index, e.RowIndex).Value
                Dim ok As Boolean = True
                For Each r As DataGridViewRow In Me.DgvWUAllocate_Process.Rows ''update 1-1
                    If r.Index <> e.RowIndex Then ''tak usah cek baris yang dijadikan reference
                        If r.Cells(Me.ColAll_Process_PrimaryProcess.Index).Value = PrimaryProcess Then ''hanya update yang process nya sama
                            ok = True
                            ''cek apakah product yang mau di eksekusi memiliki work unit yang sama dengan product yang dijakdikan referense
                            For Each rowDummy In TBLDummy.Rows
                                Me.tblProduct.DefaultView.RowFilter = "[primary product] = " & r.Cells(Me.ColAll_Process_PrimaryCodeNo.Index).Value.ToString & " AND [primary work unit] = " & rowDummy.primary_work_unit.ToString
                                If Me.tblProduct.DefaultView.Count = 0 Then
                                    MsgBox("Product code: " & r.Cells(Me.ColAll_Process_CodeNo.Index).Value & " don't have work unit: " & objWU.GetCode(rowDummy.primary_work_unit) & ".", MsgBoxStyle.Information, "Information")
                                    ok = False ''tidak eksekusi product ini
                                    Exit For
                                End If
                            Next
                            'For Each primaryWU As Double In listOfPrimaryWorkUnit
                            '    Me.tblProduct.DefaultView.RowFilter = "[primary product] = " & r.Cells(Me.ColAll_Process_PrimaryCodeNo.Index).Value.ToString & " AND [primary work unit] = " & primaryWU.ToString
                            '    If Me.tblProduct.DefaultView.Count = 0 Then
                            '        MsgBox("Product code: " & r.Cells(Me.ColAll_Process_CodeNo.Index).Value & " don't have work unit: " & objWU.GetCode(primaryWU) & ".", MsgBoxStyle.Information, "Information")
                            '        ok = False ''tidak eksekusi product ini
                            '        Exit For
                            '    End If
                            'Next

                            If ok Then
                                ''hapus semua work unit pada product yg mau di eksekusi
                                For i As Integer = tblWorkUnit.Rows.Count - 1 To 0 Step -1
                                    rowWorkUnit = tblWorkUnit.Rows(i)
                                    If rowWorkUnit.identifier_process = r.Cells(Me.ColAll_IdentifierProcess.Index).Value Then
                                        tblWorkUnit.Rows.RemoveAt(i)
                                    End If
                                Next
                                Dim priority As Integer = 0
                                Me.tblProduct.DefaultView.RowFilter = "[primary product] = " & r.Cells(Me.ColAll_Process_PrimaryCodeNo.Index).Value.ToString
                                For i As Integer = 0 To Me.tblProduct.DefaultView.Count - 1
                                    Me.rowProduct = Me.tblProduct.DefaultView(i).Row
                                    For Each rowDummy In TBLDummy.Rows
                                        If Me.rowProduct.primary_work_unit = rowDummy.primary_work_unit Then
                                            Me.rowWorkUnit = Me.tblWorkUnit.NewDataTable6Row
                                            Me.rowWorkUnit.primary_master = Me.PrimaryConfig
                                            Me.rowWorkUnit.identifier_process = r.Cells(Me.ColAll_IdentifierProcess.Index).Value
                                            Me.rowWorkUnit.primary_work_unit = rowProduct.primary_work_unit
                                            Me.rowWorkUnit.work_unit_code = rowProduct.work_unit_code
                                            Me.rowWorkUnit.description = rowProduct.description
                                            priority += 1
                                            Me.rowWorkUnit.priority = priority
                                            Me.rowWorkUnit.rate = rowDummy.rate
                                            Me.rowWorkUnit.effeciency = rowDummy.efficiency
                                            Me.rowWorkUnit.plan_rate = rowDummy.plan_rate
                                            Me.tblWorkUnit.Rows.Add(Me.rowWorkUnit)
                                        End If
                                    Next
                                    'If listOfPrimaryWorkUnit.Contains(Me.rowProduct.primary_work_unit) Then
                                    '    Me.rowWorkUnit = Me.tblWorkUnit.NewDataTable6Row
                                    '    Me.rowWorkUnit.primary_master = Me.PrimaryConfig
                                    '    Me.rowWorkUnit.identifier_process = r.Cells(Me.ColAll_IdentifierProcess.Index).Value
                                    '    Me.rowWorkUnit.primary_work_unit = rowProduct.primary_work_unit
                                    '    Me.rowWorkUnit.work_unit_code = rowProduct.work_unit_code
                                    '    Me.rowWorkUnit.description = rowProduct.description
                                    '    priority += 1
                                    '    Me.rowWorkUnit.priority = priority
                                    '    Me.rowWorkUnit.rate = rowProduct.rate_per_hour
                                    '    Me.rowWorkUnit.effeciency = rowProduct.efficiency
                                    '    Me.rowWorkUnit.plan_rate = rowProduct.rate_per_hour * rowProduct.efficiency / 100
                                    '    Me.tblWorkUnit.Rows.Add(Me.rowWorkUnit)
                                    'End If
                                Next
                                counter += 1
                            End If

                        End If
                    End If
                Next
                If counter > 0 Then
                    MsgBox("Updated: " & counter.ToString & " other product.", MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("No other product updated.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub DgvWUAllocate_Process_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvWUAllocate_Process.SelectionChanged
        Dim dr As DataGridViewRow
        Dim dtaProduct As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable7TableAdapter
        Try
            If DgvWUAllocate_Process.SelectedCells.Count <= 0 Then
                Exit Sub
            End If
            If Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex > -1 Then
                Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvWUAllocate_WU.Rows.Clear()

                If IsNothing(Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_PrimaryCodeNo.Index, Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex).Value) Then
                    Exit Sub
                End If
                tblProduct.DefaultView.RowFilter = "[primary product] = " & Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_PrimaryCodeNo.Index, Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex).Value
                If tblProduct.DefaultView.Count > 0 Then
                    For i As Integer = 0 To tblProduct.DefaultView.Count - 1
                        Me.rowProduct = Me.tblProduct.DefaultView(i).Row
                        dr = DgvWUAllocate_WU.Rows(Me.DgvWUAllocate_WU.Rows.Add)
                        dr.Cells(Me.CoAll_WU_Code.Index).Value = rowProduct.work_unit_code
                        dr.Cells(Me.CoAll_WU_Desc.Index).Value = rowProduct.description
                        dr.Cells(Me.CoAll_WU_No.Index).Value = Me.ObjCd.DigitGrouping(Me.DgvWUAllocate_WU.Rows.Count, 0)
                        dr.Cells(Me.CoAll_WU_ProdRate.Index).Value = Me.ObjCd.DigitGrouping(rowProduct.rate_per_hour, rowProcess._Decimal)
                        dr.Cells(Me.CoAll_WU_Efficiency.Index).Value = Me.ObjCd.DigitGrouping(rowProduct.efficiency, 2)
                        dr.Cells(Me.CoAll_WU_PlanRate.Index).Value = Me.ObjCd.DigitGrouping(rowProduct.rate_per_hour * rowProduct.efficiency / 100, rowProcess._Decimal)
                        dr.Cells(Me.CoAll_WU_PrimaryWU.Index).Value = rowProduct.primary_work_unit
                        dr.Cells(Me.CoAll_WU_Allocate.Index).Value = False
                    Next
                End If
                Dim primaryConfig As Double = -1
                If Me.tblHeader.Count > 0 Then
                    primaryConfig = Me.rowHeader.primary
                End If
                Me.tblWorkUnit.DefaultView.RowFilter = "[primary master] = " & primaryConfig.ToString & " AND [identifier process] = '" & Me.DgvWUAllocate_Process.Item(Me.ColAll_IdentifierProcess.Index, Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex).Value & "'"
                Me.tblWorkUnit.DefaultView.Sort = "priority"
                Dim index As Integer = -1
                If tblWorkUnit.DefaultView.Count > 0 Then
                    For i As Integer = 0 To tblWorkUnit.DefaultView.Count - 1
                        rowWorkUnit = tblWorkUnit.DefaultView(i).Row
                        index = Me.GetIndex(rowWorkUnit.primary_work_unit)
                        If index = -1 Then
                            dr = DgvWUAllocate_WU.Rows(Me.DgvWUAllocate_WU.Rows.Add)
                            dr.Cells(Me.CoAll_WU_Code.Index).Value = rowWorkUnit.work_unit_code
                            dr.Cells(Me.CoAll_WU_Desc.Index).Value = rowWorkUnit.description
                            dr.Cells(Me.CoAll_WU_No.Index).Value = Me.ObjCd.DigitGrouping(Me.DgvWUAllocate_WU.RowCount, 0)
                            dr.Cells(Me.CoAll_WU_ProdRate.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.rate, rowProcess._Decimal)
                            dr.Cells(Me.CoAll_WU_Efficiency.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.effeciency, 2)
                            dr.Cells(Me.CoAll_WU_PlanRate.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.plan_rate, rowProcess._Decimal)
                            dr.Cells(Me.CoAll_WU_PrimaryWU.Index).Value = rowWorkUnit.primary_work_unit
                            dr.Cells(Me.CoAll_WU_Allocate.Index).Value = True
                        Else
                            dr = Me.DgvWUAllocate_WU.Rows(index)
                            dr.Cells(Me.CoAll_WU_Allocate.Index).Value = True
                            dr.Cells(Me.CoAll_WU_ProdRate.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.rate, rowProcess._Decimal)
                            dr.Cells(Me.CoAll_WU_Efficiency.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.effeciency, 2)
                            dr.Cells(Me.CoAll_WU_PlanRate.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.plan_rate, rowProcess._Decimal)
                        End If
                    Next
                End If

                Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Try
            Me.CheckAll(True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Try
            Me.CheckAll(False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub CheckAll(ByVal vCheck As Boolean)
        Try
            Me.Cursor = Cursors.WaitCursor
            For i As Integer = tblWorkUnit.Rows.Count - 1 To 0 Step -1
                rowWorkUnit = tblWorkUnit.Rows(i)
                If rowWorkUnit.identifier_process = Me.DgvWUAllocate_Process.Item(Me.ColAll_IdentifierProcess.Index, Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex).Value.ToString Then
                    tblWorkUnit.Rows.RemoveAt(i)
                End If
            Next
            Dim priority As Integer = 0
            For Each row As DataGridViewRow In DgvWUAllocate_WU.Rows
                row.Cells(Me.CoAll_WU_Allocate.Index).Value = vCheck
                If vCheck Then
                    Me.rowWorkUnit = Me.tblWorkUnit.NewDataTable6Row
                    Me.rowWorkUnit.primary_master = Me.PrimaryConfig
                    Me.rowWorkUnit.identifier_process = Me.DgvWUAllocate_Process.Item(Me.ColAll_IdentifierProcess.Index, Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex).Value.ToString
                    Me.rowWorkUnit.primary_work_unit = row.Cells(Me.CoAll_WU_PrimaryWU.Index).Value
                    Me.rowWorkUnit.work_unit_code = row.Cells(Me.CoAll_WU_Code.Index).Value
                    Me.rowWorkUnit.description = row.Cells(Me.CoAll_WU_Desc.Index).Value
                    priority += 1
                    Me.rowWorkUnit.priority = priority
                    Me.rowWorkUnit.rate = row.Cells(Me.CoAll_WU_ProdRate.Index).Value
                    Me.rowWorkUnit.effeciency = row.Cells(Me.CoAll_WU_Efficiency.Index).Value
                    Me.rowWorkUnit.plan_rate = row.Cells(Me.CoAll_WU_PlanRate.Index).Value
                    Me.tblWorkUnit.Rows.Add(Me.rowWorkUnit)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
End Class