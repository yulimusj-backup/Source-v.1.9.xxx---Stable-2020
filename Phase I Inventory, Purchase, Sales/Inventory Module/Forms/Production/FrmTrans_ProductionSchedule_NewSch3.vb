Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_NewSch3

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public IdentifierJO As String = ""
    Public PrimaryConfig As Double = -1
    Public Stopped As Boolean = False
    Dim primaryProduct As Double = -1
    Dim ObjCd As New Core_Module.ClsCekData
    Dim tblProcess As New DseTrans_ProductionSchedule_NewSch3.DataTable5DataTable
    Dim rowProcess As DseTrans_ProductionSchedule_NewSch3.DataTable5Row
    Dim tblWorkUnit As New DseTrans_ProductionSchedule_NewSch3.DataTable6DataTable
    Dim rowWorkUnit As DseTrans_ProductionSchedule_NewSch3.DataTable6Row
    Dim tblHeader As New DseTrans_ProductionSchedule_NewSch3.DataTable1DataTable
    Dim rowHeader As DseTrans_ProductionSchedule_NewSch3.DataTable1Row
    Dim tblProduct As New DseTrans_ProductionSchedule_NewSch3.DataTable7DataTable
    Dim rowProduct As DseTrans_ProductionSchedule_NewSch3.DataTable7Row
    Dim CancelClose As Boolean = False
    Dim ObjSch As New ClsTrans_Schedule
    Dim DTAAcuan As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch3TableAdapter
    Dim TBLAcuan As New DseTrans_Schedule.tbltrans_mo_item_wo_item_sch3DataTable
    Dim rowAcuan As DseTrans_Schedule.tbltrans_mo_item_wo_item_sch3Row
    Dim vDummyRowIndex As Integer = -1

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim listOfInActiveWorkUnit As New List(Of String)
        Dim InActiveWorkUnit As String = ""
        Dim objWorkUnit As New Core_Module.ClsMasterWorkUnit
        Try
            Me.Cursor = Cursors.WaitCursor
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
                Me.llbEdit.Focus()
                CancelClose = True
                Exit Sub
            End If

            If Me.rboManual.Checked And Me.cboWorkUnitConfig.Items.Count = 0 Then
                MsgBox("Please add a configuration work unit first or choose automatic optimization.", MsgBoxStyle.Information, "Information")
                Me.llbAdd.Focus()
                CancelClose = True
                Exit Sub
            End If

            If Me.IsValidWorkUnit = False Then
                CancelClose = True
                Exit Sub
            End If

            If Me.rboManual.Checked And Me.cboWorkUnitConfig.SelectedIndex > -1 Then
                'iwan: 2018-09-25
                'pastikan setiap product memiliki item work unit yang dirujuk pada pre-configured wu
                'jika tidak maka bisa logic error pada saat perhitungan jadwal baru
                If Me.ApakahWorkUnitLengkap = False Then
                    Exit Sub
                End If

                Me.PrimaryConfig = Me.rowHeader.primary
            ElseIf Me.rboAutomatic.Checked Or (Me.rboManual.Checked And Me.cboWorkUnitConfig.SelectedIndex <= -1) Then
                Me.PrimaryConfig = -1
            End If
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function ApakahWorkUnitLengkap() As Boolean
        'Return False
        Dim DTA_WU_Preconfigured As New DseTrans_Schedule2TableAdapters.tbltrans_mo_master_predefine_workunit_itemworkunitTableAdapter
        Dim TBL_WU_Preconfigured As New DseTrans_Schedule2.tbltrans_mo_master_predefine_workunit_itemworkunitDataTable
        Dim row_WU_Preconfigured As DseTrans_Schedule2.tbltrans_mo_master_predefine_workunit_itemworkunitRow
        Dim DTA_WU_Master As New DseTrans_Schedule2TableAdapters.tblmaster_product_item_workunitTableAdapter
        Dim TBL_WU_Master As New DseTrans_Schedule2.tblmaster_product_item_workunitDataTable
        Dim row_WU_Master As DseTrans_Schedule2.tblmaster_product_item_workunitRow
        Dim PrimaryMax, PrimaryMin As Double
        Dim Idx As Integer = 0 
        Dim Idx1 As Integer = 0
        Dim Add As Integer = 0
        Dim Ketemu As Boolean = False

        Try
            DTA_WU_Preconfigured.Connection.ConnectionString = Core_Module.GetConstring
            DTA_WU_Master.Connection.ConnectionString = DTA_WU_Preconfigured.Connection.ConnectionString
            PrimaryMax = Double.MinValue
            PrimaryMin = Double.MaxValue

            'isikan data tabel work unit dari preconfigured
            DTA_WU_Preconfigured.Fill(TBL_WU_Preconfigured, Me.rowHeader.primary)

            If TBL_WU_Preconfigured.Rows.Count > 0 Then
                'cari primary product max min
                For Each row_WU_Preconfigured In TBL_WU_Preconfigured.Rows
                    If PrimaryMax < row_WU_Preconfigured.primary_product Then
                        PrimaryMax = row_WU_Preconfigured.primary_product
                    End If
                    If PrimaryMin > row_WU_Preconfigured.primary_product Then
                        PrimaryMin = row_WU_Preconfigured.primary_product
                    End If
                Next

                'load data dari table master
                DTA_WU_Master.Fill(TBL_WU_Master, PrimaryMax, PrimaryMin)

                'untuk setiap baris work unit, harus ada di dalam master productnya
                While Idx <= TBL_WU_Preconfigured.Rows.Count - 1
                    row_WU_Preconfigured = TBL_WU_Preconfigured.Rows(Idx) 
                    Add = 0
                    Ketemu = False

                    While Idx1 + Add <= TBL_WU_Master.Rows.Count - 1
                        row_WU_Master = TBL_WU_Master.Rows(Idx1 + Add)
                        If row_WU_Preconfigured.primary_product > row_WU_Master.primary_product Then
                            'cek baris selanjutnya pada master product
                            Idx1 += 1
                            Add = 0
                        ElseIf row_WU_Preconfigured.primary_product < row_WU_Master.primary_product Then
                            'work unit tsb tidak ditemukan pada master product
                            Ketemu = False
                            Exit While
                        Else
                            If row_WU_Preconfigured.primary_work_unit = row_WU_Master.primary_work_unit Then
                                TBL_WU_Preconfigured.Rows.Remove(row_WU_Preconfigured)
                                Ketemu = True 
                                Exit While
                            Else
                                Add += 1
                            End If
                        End If
                    End While

                    If Ketemu = False Then
                        Idx += 1 
                    End If
                End While

                If TBL_WU_Preconfigured.Rows.Count > 0 Then
                    Dim ij As New FrmTrans_ProductionSchedule_NewSch3_Exception
                    ij.FrmParent = Me
                    ij.TBLRef = TBL_WU_Preconfigured

                    ij.ShowDialog()

                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA_WU_Preconfigured.Dispose()
            TBL_WU_Preconfigured.Rows.Clear()
            TBL_WU_Preconfigured.Dispose()
            DTA_WU_Master.Dispose()
            TBL_WU_Master.Dispose()
            TBL_WU_Master.Rows.Clear()
        End Try
    End Function

    Function IsValidWorkUnit() As Boolean
        Dim tblConfig As New DseTrans_ProductionSchedule_NewSch3.DataTable12DataTable
        Dim rowConfig As DseTrans_ProductionSchedule_NewSch3.DataTable12Row

        Dim dtaP As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable13TableAdapter
        Dim tblP As New DseTrans_ProductionSchedule_NewSch3.DataTable13DataTable
        Dim rowP As DseTrans_ProductionSchedule_NewSch3.DataTable13Row

        Try
            For Each Me.rowProcess In Me.tblProcess.Rows
                Me.tblWorkUnit.DefaultView.RowFilter = "[identifier process] = '" & Me.rowProcess.identifier & "' and [primary master] = " & Me.PrimaryConfig.ToString
                If Me.tblWorkUnit.DefaultView.Count > 0 Then
                    For i As Integer = 0 To Me.tblWorkUnit.DefaultView.Count - 1
                        Me.rowWorkUnit = Me.tblWorkUnit.DefaultView(i).Row
                        rowConfig = tblConfig.NewDataTable12Row
                        rowConfig.identifier = Me.rowWorkUnit.identifier_process
                        rowConfig.primary_product = Me.rowProcess.primary_product
                        rowConfig.product_code = Me.rowProcess.Product_Code
                        rowConfig.primary_work_unit = Me.rowWorkUnit.primary_work_unit
                        rowConfig.work_unit_code = Me.rowWorkUnit.work_unit_code
                        tblConfig.Rows.Add(rowConfig)
                    Next
                End If
            Next

            dtaP.Connection.ConnectionString = Core_Module.GetConstring
            dtaP.ClearBeforeFill = False
            tblP.Rows.Clear()
            For Each rowConfig In tblConfig
                dtaP.Fill(tblP, rowConfig.primary_product)
            Next
            For Each rowConfig In tblConfig
                tblP.DefaultView.RowFilter = "[primary product] = " & rowConfig.primary_product.ToString & " and [primary work unit] = " & rowConfig.primary_work_unit
                If tblP.DefaultView.Count = 0 Then
                    MsgBox("Invalid work unit '" & rowConfig.work_unit_code & "' for code no '" & rowConfig.product_code & "'", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub FrmTrans_ProductionSchedule_NewSch3_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.CancelClose = True Then
                e.Cancel = True
                Me.CancelClose = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_NewSch3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim objProduct As New Core_Module.clsMasterProduct
        Try
            Me.CoAll_WU_Allocate.Visible = False
            Me.DgvWUAllocate_Process.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvWUAllocate_WU.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.primaryProduct = Me.GetPrimaryProductByJO
            Me.rboManual.Checked = True
            Me.Text = Me.Text & " for: " & objProduct.GetProductCode(Me.primaryProduct)
            DTAAcuan.Connection.ConnectionString = Core_Module.GetConstring
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboManual_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rboManual.CheckedChanged
        Try
            If Me.rboManual.Checked = True Then
                Me.ReloadSavedConfig()
            Else
                Me.cboWorkUnitConfig.Enabled = False
                Me.llbAdd.Enabled = False
                Me.llbDelete.Enabled = False
                Me.llbEdit.Enabled = False
                Me.cboWorkUnitConfig.SelectedIndex = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub ReloadSavedConfig()
        Dim dtaHeader As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable1TableAdapter
        Dim dtaProcess As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable5TableAdapter
        Dim dtaWorkUnit As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable6TableAdapter
        Dim dtaProcessCurrentJO As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable4TableAdapter
        Dim tblProcessCurrentJO As New DseTrans_ProductionSchedule_NewSch3.DataTable4DataTable
        Dim dtaProcessOtherJO As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable2TableAdapter
        Dim tblProcessOtherJO As New DseTrans_ProductionSchedule_NewSch3.DataTable2DataTable
        Dim rowProcessOtherJO As DseTrans_ProductionSchedule_NewSch3.DataTable2Row
        Dim listOfPrimaryMasterExist As New List(Of Double)
        Dim listOfPrimaryMasterCorrect As New List(Of Double)
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.cboWorkUnitConfig.Enabled = True
            Me.llbAdd.Enabled = True
            Me.llbDelete.Enabled = True
            Me.llbEdit.Enabled = True

            dtaHeader.Connection.ConnectionString = Core_Module.GetConstring
            dtaProcess.Connection.ConnectionString = Core_Module.GetConstring
            dtaWorkUnit.Connection.ConnectionString = Core_Module.GetConstring
            dtaProcessCurrentJO.Connection.ConnectionString = Core_Module.GetConstring
            dtaProcessOtherJO.Connection.ConnectionString = Core_Module.GetConstring
            tblHeader.Rows.Clear()

            dtaProcess.Fill(tblProcess, Me.primaryProduct)
            dtaWorkUnit.Fill(tblWorkUnit, Me.primaryProduct)
            dtaProcessCurrentJO.Fill(tblProcessCurrentJO, Me.IdentifierJO)
            dtaProcessOtherJO.Fill(tblProcessOtherJO, Me.primaryProduct)

            For Each rowProcessOtherJO In tblProcessOtherJO
                If listOfPrimaryMasterExist.Contains(rowProcessOtherJO.primary_master) = False Then listOfPrimaryMasterExist.Add(rowProcessOtherJO.primary_master)
            Next

            For Each primaryMasterExist As Double In listOfPrimaryMasterExist
                tblProcessOtherJO.DefaultView.RowFilter = "[primary master] = " & primaryMasterExist.ToString
                If tblProcessOtherJO.DefaultView.Count = tblProcessCurrentJO.Count Then
                    Dim Correct As Boolean = True
                    For i As Integer = 0 To tblProcessOtherJO.DefaultView.Count - 1
                        rowProcessOtherJO = tblProcessOtherJO.DefaultView(i).Row
                        tblProcessCurrentJO.DefaultView.RowFilter = "[Primary Process] = " & rowProcessOtherJO.primary_process.ToString & " AND [Primary Product] = " & rowProcessOtherJO.primary_product.ToString
                        If tblProcessCurrentJO.DefaultView.Count = 0 Then Correct = False
                    Next
                    If Correct And listOfPrimaryMasterCorrect.Contains(primaryMasterExist) = False Then listOfPrimaryMasterCorrect.Add(primaryMasterExist)
                End If
            Next

            ''rosadin
            For i As Integer = tblWorkUnit.Rows.Count - 1 To 0 Step -1
                rowWorkUnit = tblWorkUnit.Rows(i)
                If Not listOfPrimaryMasterCorrect.Contains(rowWorkUnit.primary_master) Then
                    tblWorkUnit.Rows.RemoveAt(i)
                End If
            Next

            dtaHeader.ClearBeforeFill = False
            For Each primaryMasterCorrect As Double In listOfPrimaryMasterCorrect
                dtaHeader.FillBy_Primary(tblHeader, primaryMasterCorrect)
            Next
            dtaHeader.ClearBeforeFill = True

            tblHeader.DefaultView.RowFilter = ""
            tblHeader.DefaultView.Sort = "code"
            Me.cboWorkUnitConfig.DataSource = tblHeader.DefaultView
            Me.cboWorkUnitConfig.DisplayMember = "code"
            Me.cboWorkUnitConfig.ValueMember = "primary"
            If Me.cboWorkUnitConfig.Items.Count > 0 Then
                Me.cboWorkUnitConfig.SelectedIndex = -1
                Me.cboWorkUnitConfig.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Function GetPrimaryProductByJO() As Double
        Dim dta As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable1TableAdapter
        Dim n As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            n = dta.ScalarQuery_GetProductByJO(Me.IdentifierJO)
            If n.HasValue Then
                Return n.Value
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub cboWorkUnitConfig_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboWorkUnitConfig.SelectedIndexChanged
        Dim dr As DataGridViewRow
        Dim FirmPlanDate As Date = Today
        Dim rowDGV As DataGridViewRow
        Dim DummyStart As Date = Today

        Try
            Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvWUAllocate_WU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvWUAllocate_Process.Rows.Clear()
            Me.DgvWUAllocate_WU.Rows.Clear()
            Me.tblProduct.Rows.Clear()
            If Me.cboWorkUnitConfig.Items.Count > 0 And Me.cboWorkUnitConfig.SelectedIndex > -1 Then
                Me.rowHeader = Me.tblHeader.DefaultView(Me.cboWorkUnitConfig.SelectedIndex).Row
                Me.tblProcess.DefaultView.RowFilter = "[primary master] = " & rowHeader.primary.ToString
                For i As Integer = 0 To tblProcess.DefaultView.Count - 1
                    rowProcess = tblProcess.DefaultView(i).Row
                    dr = DgvWUAllocate_Process.Rows(DgvWUAllocate_Process.Rows.Add())
                    dr.Cells(Me.ColAll_Process_Code.Index).Value = rowProcess.Process_Code
                    dr.Cells(Me.ColAll_Process_CodeNo.Index).Value = rowProcess.Product_Code
                    dr.Cells(Me.ColAll_Process_Desc.Index).Value = rowProcess.Product_Desc
                    dr.Cells(Me.ColAll_Process_No.Index).Value = Me.DgvWUAllocate_Process.Rows.Count
                    dr.Cells(Me.ColAll_Process_PrimaryCodeNo.Index).Value = rowProcess.primary_product
                    dr.Cells(Me.ColAll_Process_PrimaryProcess.Index).Value = rowProcess.primary_process
                    dr.Cells(Me.ColAll_IdentifierProcess.Index).Value = rowProcess.identifier
                Next
                Me.tblWorkUnit.DefaultView.RowFilter = "[primary master] = " & Me.rowHeader.primary.ToString & " AND [identifier process] = '" & Me.DgvWUAllocate_Process.Item(Me.ColAll_IdentifierProcess.Index, 0).Value.ToString & "'"
                Me.tblWorkUnit.DefaultView.Sort = "priority"
                For i As Integer = 0 To tblWorkUnit.DefaultView.Count - 1
                    rowWorkUnit = tblWorkUnit.DefaultView(i).Row
                    dr = DgvWUAllocate_WU.Rows(Me.DgvWUAllocate_WU.Rows.Add)
                    dr.Cells(Me.CoAll_WU_Code.Index).Value = rowWorkUnit.work_unit_code
                    dr.Cells(Me.CoAll_WU_Desc.Index).Value = rowWorkUnit.description
                    dr.Cells(Me.CoAll_WU_No.Index).Value = Me.ObjCd.DigitGrouping(Me.DgvWUAllocate_WU.RowCount, 0)
                    dr.Cells(Me.CoAll_WU_ProdRate.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.rate, rowProcess._Decimal)
                    dr.Cells(Me.CoAll_WU_Efficiency.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.effeciency, 2)
                    dr.Cells(Me.CoAll_WU_PlanRate.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.plan_rate, rowProcess._Decimal)
                    dr.Cells(Me.CoAll_WU_PrimaryWU.Index).Value = rowWorkUnit.primary_work_unit
                    dr.Cells(Me.CoAll_WU_Allocate.Index).Value = True
                Next

                'isikan start time
                If Me.DgvWUAllocate_WU.RowCount > 0 Then
                    FirmPlanDate = Me.ObjSch.GetFirmPlan
                    For Each rowDGV In Me.DgvWUAllocate_WU.Rows
                        'cari tanggal end terakhir pada sch bar, untuk wu ybs
                        DummyStart = Me.GetStartTime(rowDGV.Cells(Me.CoAll_WU_PrimaryWU.Index).Value, FirmPlanDate)

                        If DummyStart <> #1/1/1950# Then
                            'bandingkan dengan tanggal terakhir, proses sebelumnya
                            rowDGV.Cells(Me.CoAll_WU_StartDate.Index).Value = Format(DummyStart, "dd MMM yyyy HH.mm")
                        Else
                            'jika ada alokasi, maka masukkan jam mulai pertama pada tgl ybs
                            DummyStart = Me.ObjSch.GetFirstStartTime(FirmPlanDate + TimeSpan.FromDays(1), rowDGV.Cells(Me.CoAll_WU_PrimaryWU.Index).Value)
                            If DummyStart <> #1/1/1950# Then
                                rowDGV.Cells(Me.CoAll_WU_StartDate.Index).Value = Format(DummyStart, "dd MMM yyyy HH.mm")
                            Else
                                rowDGV.Cells(Me.CoAll_WU_StartDate.Index).Value = Format(FirmPlanDate, "dd MMM yyyy HH.mm")
                            End If
                        End If
                    Next
                End If
            End If
            Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvWUAllocate_WU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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

    Private Sub llbDelete_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbDelete.LinkClicked
        Dim dta As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable1TableAdapter
        Try
            If Me.cboWorkUnitConfig.Items.Count <= 0 Then
                Exit Sub
            End If
            If Me.rowHeader.primary > -1 Then

                If Me.ApakahBolehHapus(Me.rowHeader.primary) = False Then
                    MsgBox("This configuration is already use in schedule.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If MsgBox("Are you sure to delete this config?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    dta.Connection.ConnectionString = Core_Module.GetConstring
                    If dta.DeleteQuery(Me.rowHeader.primary) > 0 Then
                        MsgBox("Predefined work unit config deleted.", MsgBoxStyle.Information, "Information")
                    Else
                        MsgBox("Predefined work unit config failed to delete.", MsgBoxStyle.Information, "Information")
                    End If
                    Me.ReloadSavedConfig()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahBolehHapus(ByVal vPrimaryConfig As Double) As Boolean
        Dim d As New DseMasterWorkUnitConfigurationTableAdapters.DataTable1TableAdapter
        Dim v As New Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            v = d.ScalarQuery_GetCountInJO(vPrimaryConfig)
            If v.HasValue Then
                If v.Value > 0 Then
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub llbEdit_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbEdit.LinkClicked
        Dim f As New FrmTrans_ProductionSchedule_NewSch4
        Try
            If Me.cboWorkUnitConfig.Items.Count <= 0 Then
                Exit Sub
            End If
            If Me.rowHeader.primary > -1 Then
                f.PrimaryConfig = Me.rowHeader.primary
                f.PrimaryProduct = Me.primaryProduct
                f.UserCode = Me.UserCode
                f.BackColorAll = Me.BackColorAll
                f.Baru = False
                f.IdentifierJO = Me.IdentifierJO
                f.ShowDialog()
                Me.ReloadSavedConfig()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAdd_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAdd.LinkClicked
        Dim f As New FrmTrans_ProductionSchedule_NewSch4
        Try
            f.PrimaryConfig = -1
            f.PrimaryProduct = Me.primaryProduct
            f.UserCode = Me.UserCode
            f.BackColorAll = Me.BackColorAll
            f.Baru = True
            f.IdentifierJO = Me.IdentifierJO
            f.ShowDialog()
            Me.ReloadSavedConfig()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetStartTime(ByVal vPrimaryWU As Double, ByVal vFirmPlanned As Date) As Date
        Try
            If Me.TBLAcuan.Rows.Count = 0 Then
                Me.DTAAcuan.Fill(Me.TBLAcuan, (vFirmPlanned - TimeSpan.FromDays(2)).Date)
            End If

            Me.TBLAcuan.DefaultView.RowFilter = "[primary work unit] = " + vPrimaryWU.ToString
            If Me.TBLAcuan.DefaultView.Count = 0 Then
                Return #1/1/1950#
            Else
                Me.rowAcuan = Me.TBLAcuan.DefaultView(0).Row
                Return Me.rowAcuan.plan_end_time
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvWUAllocate_Process_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvWUAllocate_Process.SelectionChanged
        Dim dr As DataGridViewRow
        Dim FirmPlanDate As Date = Today
        Dim rowDGV As DataGridViewRow
        Dim DummyStart As Date = Today

        Try
            If DgvWUAllocate_Process.SelectedCells.Count <= 0 Then
                Me.vDummyRowIndex = -1
                Exit Sub
            End If
            If Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex > -1 Then
                If Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex = Me.vDummyRowIndex Then
                    Exit Sub
                Else
                    Me.vDummyRowIndex = Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex
                End If

                Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvWUAllocate_WU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvWUAllocate_WU.Rows.Clear()
                If IsNothing(Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_PrimaryCodeNo.Index, Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex).Value) Then
                    Exit Sub
                End If
                Me.tblWorkUnit.DefaultView.RowFilter = "[primary master] = " & Me.rowHeader.primary.ToString & " AND [identifier process] = '" & Me.DgvWUAllocate_Process.Item(Me.ColAll_IdentifierProcess.Index, Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex).Value & "'"
                Me.tblWorkUnit.DefaultView.Sort = "priority"
                For i As Integer = 0 To tblWorkUnit.DefaultView.Count - 1
                    rowWorkUnit = tblWorkUnit.DefaultView(i).Row
                    dr = DgvWUAllocate_WU.Rows(Me.DgvWUAllocate_WU.Rows.Add)
                    dr.Cells(Me.CoAll_WU_Code.Index).Value = rowWorkUnit.work_unit_code
                    dr.Cells(Me.CoAll_WU_Desc.Index).Value = rowWorkUnit.description
                    dr.Cells(Me.CoAll_WU_No.Index).Value = Me.DgvWUAllocate_WU.RowCount
                    dr.Cells(Me.CoAll_WU_ProdRate.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.rate, rowProcess._Decimal)
                    dr.Cells(Me.CoAll_WU_Efficiency.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.effeciency, 2)
                    dr.Cells(Me.CoAll_WU_PlanRate.Index).Value = Me.ObjCd.DigitGrouping(rowWorkUnit.plan_rate, rowProcess._Decimal)
                    dr.Cells(Me.CoAll_WU_PrimaryWU.Index).Value = rowWorkUnit.primary_work_unit
                    dr.Cells(Me.CoAll_WU_StartDate.Index).Value = ""

                    dr.Cells(Me.CoAll_WU_Allocate.Index).Value = True
                Next

                'isikan start time
                FirmPlanDate = Me.ObjSch.GetFirmPlan
                For Each rowDGV In Me.DgvWUAllocate_WU.Rows
                    'cari tanggal end terakhir pada sch bar, untuk wu ybs
                    DummyStart = Me.GetStartTime(rowDGV.Cells(Me.CoAll_WU_PrimaryWU.Index).Value, FirmPlanDate)

                    If DummyStart <> #1/1/1950# Then
                        'bandingkan dengan tanggal terakhir, proses sebelumnya
                        rowDGV.Cells(Me.CoAll_WU_StartDate.Index).Value = Format(DummyStart, "dd MMM yyyy HH.mm")
                    Else
                        'jika ada alokasi, maka masukkan jam mulai pertama pada tgl ybs
                        DummyStart = Me.ObjSch.GetFirstStartTime(FirmPlanDate + TimeSpan.FromDays(1), rowDGV.Cells(Me.CoAll_WU_PrimaryWU.Index).Value)
                        If DummyStart <> #1/1/1950# Then
                            rowDGV.Cells(Me.CoAll_WU_StartDate.Index).Value = Format(DummyStart, "dd MMM yyyy HH.mm")
                        Else
                            rowDGV.Cells(Me.CoAll_WU_StartDate.Index).Value = Format(FirmPlanDate, "dd MMM yyyy HH.mm")
                        End If
                    End If
                Next

                Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.DgvWUAllocate_WU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStop.Click
        If MsgBox("Stop adding new schedule?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
            Me.Stopped = True
            Me.Close()
        End If
    End Sub

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim d As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable1TableAdapter
        Dim s As String = ""
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            s = d.ScalarQuery_GetMasterLineCode(vPrimary)
            If String.IsNullOrEmpty(s) = False Then
                Return s
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
End Class