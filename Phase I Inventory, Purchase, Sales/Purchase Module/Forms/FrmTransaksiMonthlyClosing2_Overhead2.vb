Option Explicit On
Option Strict Off

Public Class FrmTransaksiMonthlyClosing2_Overhead2
    Public ClosingPeriod As Date = Today
    Public vTBLCostGroup_Account As New DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountDataTable
    Public vTBLCostGroup_Scope As New DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeDataTable
    Public vTBLCostGroup_CodeNo As New DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noDataTable
    Public vTBLCostGroup_CodeNo2 As New DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_no2DataTable
    Public UserCode As String = ""

    Dim mDgv As DataGridView
    Dim mEraseRowIndex As Integer = -1
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Buffer As String = ""
    Dim BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjUnit As New Core_Module.ClsMasterUnit
    Dim ObjAkun As New AccModule.ClsMasterAkun
    Dim ObjProcess As New Core_Module.ClsMasterProses
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjClosing As New ClsTransaksiMonthlyClosing
    Dim ObjCLosingItem As New ClsTransaksiMonthlyClosingItem
    Dim CekBeforeCLose As Boolean = True

    Private Sub BgwDelCostGroup_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwDelCostGroup.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDelCostGroup_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwDelCostGroup.RunWorkerCompleted
        Try
            Me.dgvItemCost_CostGroup.Rows.RemoveAt(Me.mEraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDelCostAcc_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwDelCostAcc.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDelCostAcc_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwDelCostAcc.RunWorkerCompleted
        Try
            Me.dgvItemCost_CostAccountAllocation.Rows.RemoveAt(Me.mEraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDelCostScope_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwDelCostScope.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDelCostScope_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwDelCostScope.RunWorkerCompleted
        Try
            Me.dgvItemCost_Scope.Rows.RemoveAt(Me.mEraseRowIndex)
            Me.ObjCD.SetNumber(Me.dgvItemCost_Scope)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemCost_CostGroup.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvItemCost_CostGroup.Item(Me.ColCostGroup_GroupDesc.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColCostGroup_GroupDesc.Index Or e.ColumnIndex = Me.ColCostGroup_Remarks.Index Or _
                e.ColumnIndex = Me.ColCostGroup_RefUnit.Index Then
                    Me.Buffer = Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColCostGroup_GroupDesc.Index Then
                    Me.Buffer = Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemCostGroup(ByVal vDesc As String, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1

        Try
            For ij As Integer = 0 To Me.dgvItemCost_CostGroup.Rows.Count - 2
                RowDGV = Me.dgvItemCost_CostGroup.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value = vDesc Then
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

    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemCost_CostGroup.CellEndEdit
        Dim rowdgv As DataGridViewRow = Me.dgvItemCost_CostGroup.Rows(e.RowIndex)
        Dim index As Integer = -1

        Try
            If IsNothing(Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColCostGroup_GroupDesc.Index Then
                    Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    index = Me.GetIndexItemCostGroup(Me.dgvItemCost_CostGroup.Item(Me.ColCostGroup_GroupDesc.Index, e.RowIndex).Value, e.RowIndex)
                    If index > 0 Then
                        MsgBox("Description duplication, please change description.", MsgBoxStyle.Information, "Information")
                        Me.mDgv = Me.dgvItemCost_CostGroup
                        Me.mEraseRowIndex = rowdgv.Index
                        Me.BgwDelCostGroup.RunWorkerAsync()
                        Exit Sub
                    End If
                    'jika yang sedang ditampilkan adalah recap cost group-account atau recap cost group-scope
                    'maka tampilkan kembali 
                    If Me.TabControl2.SelectedTab.Name = Me.TabPage8.Name Then
                        Me.TampilkanOverhead_AccountAllocationRecap()
                    End If
                ElseIf e.ColumnIndex = Me.ColCostGroup_RefUnit.Index Then
                    ''data unit 
                    Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ObjUnit.ApakahSudahAda(Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode unit yang diketik ada --> load data primary unit
                        Me.dgvItemCost_CostGroup.Item(Me.ColCostGroup_PrimaryRefUnit.Index, e.RowIndex).Value = Me.ObjUnit.GetPrimary(Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                    Else
                        '2. kode unit yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New Core_Module.FrmCariMasterUnit
                        ij.BackColorAll = Me.BackColorAll
                        ij.txtCari.Text = Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()

                        If ij.PrimaryUnit <> -1 Then
                            '2.a. kode unit yang dicari ada --> load data primary unit + kode unit
                            Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjUnit.GetUnitCode(ij.PrimaryUnit)
                            Me.dgvItemCost_CostGroup.Item(Me.ColCostGroup_PrimaryRefUnit.Index, e.RowIndex).Value = ij.PrimaryUnit
                        ElseIf ij.PrimaryUnit = -1 Then
                            '2.b. kode unit yang dicari tidak ada --> kembalikan ke kondisi buffer
                            Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        ij = Nothing
                    End If
                End If
            Else
                If e.ColumnIndex = Me.ColCostGroup_Remarks.Index Then
                    Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColCostGroup_GroupDesc.Index Then
                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                        If Me.dgvItemCost_CostGroup.Rows.Count - 1 <> e.RowIndex Then
                            Me.mEraseRowIndex = e.RowIndex
                            Me.BgwDelCostGroup.RunWorkerAsync()
                        End If
                    Else
                        Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                ElseIf e.ColumnIndex = Me.ColCostGroup_RefUnit.Index Then
                    Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Me.dgvItemCost_CostGroup.Item(Me.ColCostGroup_PrimaryRefUnit.Index, e.RowIndex).Value = -1
                Else
                    Me.dgvItemCost_CostGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemCost_CostGroup.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvItemCost_CostGroup.TabStop = True

            rowDGV = Me.dgvItemCost_CostGroup.Rows(Me.dgvItemCost_CostGroup.RowCount - 1)
            With rowDGV
                .Cells(Me.ColCostGroup_Identifier.Index).Value = Me.GetNextIdentifierCostGroup(Me.dgvItemCost_CostGroup, -1, Me)
                .Cells(Me.ColCostGroup_AllocationStyle.Index).Value = "By Process Output"
                .Cells(Me.ColCostGroup_RefUnit.Index).Value = ""
                .Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value = -1
                .Cells(Me.ColCostGroup_OverheadAmount.Index).Value = Me.ObjCD.DigitGrouping(0, Me.ObjCurrency.GetDecimal_DefaultCurrency)
                .Cells(Me.ColCostGroup_Remarks.Index).Value = ""
            End With
            Me.btnUP.Enabled = True
            Me.btnDown.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemCost_CostGroup.RowsRemoved
        Try
            If Me.dgvItemCost_CostGroup.Rows.Count = 0 Then
                Me.dgvItemCost_CostGroup.TabStop = False

                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemCost_CostGroup.SelectionChanged
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.dgvItemCost_CostGroup.SelectedRows.Count > 0 And Me.dgvItemCost_CostGroup.Rows.Count > 1 Then
                If dgvItemCost_CostGroup.Rows.Count - 1 <= 1 Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                Else
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                End If
            Else
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If

            If dgvItemCost_CostGroup.Rows.Count - 1 = -1 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Exit Sub
            End If

            Me.TampilkanOverheadDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub TampilkanOverheadDetail()
        Try
            If Me.TabControl2.SelectedTab.Name = Me.TabPage5.Name Then
                Me.TampilkanOverhead_AccountAllocation()
            ElseIf Me.TabControl2.SelectedTab.Name = Me.TabPage6.Name Then
                Me.TampilkanOverhead_Scope()
            ElseIf Me.TabControl2.SelectedTab.Name = Me.TabPage8.Name Then
                Me.TampilkanOverhead_AccountAllocationRecap()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanCodeNo_n_Desc_n_UnitCode_n_ProcessCode()
        Dim ListPrimary As New List(Of Double)
        Dim ListPrimaryUnit As New List(Of Double)
        Dim ListPrimaryProcess As New List(Of Double)
        Dim TBLProduct As New DseTransaksiMonthlyClosing_Overhead.tblmaster_productDataTable
        Dim rowProduct As DseTransaksiMonthlyClosing_Overhead.tblmaster_productRow
        Dim DTAProduct As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tblmaster_productTableAdapter
        Dim rowCodeNo As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noRow
        Dim DTAUnit As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tblmaster_master_unitTableAdapter
        Dim TBLUnit As New DseTransaksiMonthlyClosing_Overhead.tblmaster_master_unitDataTable
        Dim rowUnit As DseTransaksiMonthlyClosing_Overhead.tblmaster_master_unitRow
        Dim DTAProcess As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tblmaster_prosesTableAdapter
        Dim TBLProcess As New DseTransaksiMonthlyClosing_Overhead.tblmaster_prosesDataTable
        Dim rowProcess As DseTransaksiMonthlyClosing_Overhead.tblmaster_prosesRow

        Try
            DTAProduct.Connection.ConnectionString = Core_Module.GetConstring
            Me.vTBLCostGroup_CodeNo.DefaultView.RowFilter = "[Code No] = ''"
            For Each DV As DataRowView In Me.vTBLCostGroup_CodeNo.DefaultView
                rowCodeNo = DV.Row
                If ListPrimary.Contains(rowCodeNo.primary_code_no) = False Then
                    ListPrimary.Add(rowCodeNo.primary_code_no)
                End If
                If ListPrimaryUnit.Contains(rowCodeNo.primary_unit) = False Then
                    ListPrimaryUnit.Add(rowCodeNo.primary_unit)
                End If
                If ListPrimaryProcess.Contains(rowCodeNo.primary_process) = False Then
                    ListPrimaryProcess.Add(rowCodeNo.primary_process)
                End If
            Next

            If ListPrimary.Count > 0 Then
                DTAProduct.Fill(TBLProduct)
                For Each PrimaryCodeNo As Double In ListPrimary
                    Me.vTBLCostGroup_CodeNo.DefaultView.RowFilter = "[primary code no] = " + PrimaryCodeNo.ToString
                    TBLProduct.DefaultView.RowFilter = "Primary = " + PrimaryCodeNo.ToString
                    rowProduct = TBLProduct.DefaultView(0).Row
                    For Each DV As DataRowView In Me.vTBLCostGroup_CodeNo.DefaultView
                        rowCodeNo = DV.Row
                        rowCodeNo.Code_No = rowProduct.Product_Code
                        rowCodeNo.Desc = rowProduct.Description
                    Next
                Next
            End If

            If ListPrimaryUnit.Count > 0 Then
                DTAUnit.Fill(TBLUnit)
                For Each PrimaryUnit As Double In ListPrimaryUnit
                    Me.vTBLCostGroup_CodeNo.DefaultView.RowFilter = "[primary unit] = " + PrimaryUnit.ToString
                    TBLUnit.DefaultView.RowFilter = "Primary = " + PrimaryUnit.ToString
                    rowUnit = TBLUnit.DefaultView(0).Row
                    For Each DV As DataRowView In Me.vTBLCostGroup_CodeNo.DefaultView
                        rowCodeNo = DV.Row
                        rowCodeNo.Unit_Code = rowUnit.Unit_Code
                    Next
                Next
            End If

            If ListPrimaryProcess.Count > 0 Then
                DTAProcess.Fill(TBLProcess)
                For Each PrimaryProcess As Double In ListPrimaryProcess
                    Me.vTBLCostGroup_CodeNo.DefaultView.RowFilter = "[primary process] = " + PrimaryProcess.ToString
                    TBLProcess.DefaultView.RowFilter = "Primary = " + PrimaryProcess.ToString
                    rowProcess = TBLProcess.DefaultView(0).Row
                    For Each DV As DataRowView In Me.vTBLCostGroup_CodeNo.DefaultView
                        rowCodeNo = DV.Row
                        rowCodeNo.Process_Code = rowProcess.Code
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLProduct.Dispose()
            DTAProduct.Dispose()
            ListPrimary = Nothing
            DTAUnit.Dispose()
            TBLUnit.Dispose()
            DTAProcess.Dispose()
            TBLProcess.Dispose()
        End Try
    End Sub

    Private Sub TampilkanOverhead_AccountAllocationRecap()
        Try
            If Me.rboGroup_Account.Checked = True Then
                Me.TampilkanOverhead_AccountAllocationRecap1()
            Else
                Me.TampilkanOverhead_AccountAllocationRecap2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TampilkanOverhead_AccountAllocationRecap1()
        'tampilkan cost group-account
        Dim Identifier As String = ""
        Dim rowAccount As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim rowDGV, rowDGVCostGroup As DataGridViewRow
        Dim DefPrec As Integer = 0

        Try
            dgvItemCost_AllocationRecap1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvItemCost_AllocationRecap1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            dgvItemCost_AllocationRecap1.AllowUserToAddRows = False

            Me.dgvItemCost_AllocationRecap1.Rows.Clear()
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

            For Each rowDGVCostGroup In Me.dgvItemCost_CostGroup.Rows
                If String.IsNullOrEmpty(rowDGVCostGroup.Cells(Me.ColCostGroup_GroupDesc.Index).Value) = False Then
                    Identifier = rowDGVCostGroup.Cells(Me.ColCostGroup_Identifier.Index).Value

                    'tampilkan isi account allocation
                    Me.vTBLCostGroup_Account.DefaultView.RowFilter = "identifier = '" + Identifier + "'"
                    For Each DV As DataRowView In Me.vTBLCostGroup_Account.DefaultView
                        rowAccount = DV.Row
                        Me.dgvItemCost_AllocationRecap1.Rows.Add()
                        rowDGV = Me.dgvItemCost_AllocationRecap1.Rows(Me.dgvItemCost_AllocationRecap1.Rows.Count - 1)

                        rowDGV.Cells(Me.ColAlloRecap1_CostGroup.Index).Value = rowDGVCostGroup.Cells(Me.ColCostGroup_GroupDesc.Index).Value
                        rowDGV.Cells(Me.ColAlloRecap1_AccCode.Index).Value = Me.ObjAkun.GetAccountCode(rowAccount.primary_akun)
                        rowDGV.Cells(Me.ColAlloRecap1_AccDesc.Index).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(rowAccount.primary_akun)
                        rowDGV.Cells(Me.ColAlloRecap1_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping(rowAccount.allocated_amount, DefPrec)
                        rowDGV.Cells(Me.ColAlloRecap1_AllocatedPerct.Index).Value = Me.ObjCD.DigitGrouping(rowAccount.perc_allocated, 2)
                        rowDGV.Cells(Me.ColAlloRecap1_Balance.Index).Value = Me.ObjCD.DigitGrouping(rowAccount.balance, DefPrec)
                        rowDGV.Cells(Me.ColAlloRecap1_Identifier.Index).Value = Identifier
                        rowDGV.Cells(Me.ColAlloRecap1_PrimaryAkun.Index).Value = rowAccount.primary_akun
                        If rowAccount.ignore_negative = 1 Then
                            rowDGV.Cells(Me.ColAlloRecap1_IngoreNegativeBalance.Index).Value = Boolean.TrueString
                        Else
                            rowDGV.Cells(Me.ColAlloRecap1_IngoreNegativeBalance.Index).Value = Boolean.FalseString
                        End If
                    Next
                End If
            Next
            dgvItemCost_AllocationRecap1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TampilkanOverhead_AccountAllocationRecap2()
        'tampilkan cost group - scope
        Dim rowScope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow
        Dim rowDGV, rowDGVCostGroup As DataGridViewRow
        Dim Identifier As String = ""

        Try
            Me.dgvItemCost_AllocationRecap2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItemCost_AllocationRecap2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvItemCost_AllocationRecap2.AllowUserToAddRows = False

            Me.dgvItemCost_AllocationRecap2.Rows.Clear()
            For Each rowDGVCostGroup In Me.dgvItemCost_CostGroup.Rows
                If String.IsNullOrEmpty(rowDGVCostGroup.Cells(Me.ColCostGroup_GroupDesc.Index).Value) = False Then
                    Identifier = rowDGVCostGroup.Cells(Me.ColCostGroup_Identifier.Index).Value

                    'tampilkan isi account scope
                    Me.vTBLCostGroup_Scope.DefaultView.RowFilter = "identifier = '" + Identifier + "'"
                    For Each DV As DataRowView In Me.vTBLCostGroup_Scope.DefaultView
                        rowScope = DV.Row
                        Me.dgvItemCost_AllocationRecap2.Rows.Add()
                        rowDGV = Me.dgvItemCost_AllocationRecap2.Rows(Me.dgvItemCost_AllocationRecap2.Rows.Count - 1)
                        rowDGV.Cells(Me.ColAlloRecap2_CostGroup.Index).Value = rowDGVCostGroup.Cells(Me.ColCostGroup_GroupDesc.Index).Value
                        rowDGV.Cells(Me.ColAlloRecap2_Identifier.Index).Value = Identifier
                        rowDGV.Cells(Me.ColAlloRecap2_PrimaryProcess.Index).Value = rowScope.primary_process
                        rowDGV.Cells(Me.ColAlloRecap2_ProcessCode.Index).Value = Me.ObjProcess.GetCode(rowScope.primary_process)
                        rowDGV.Cells(Me.ColAlloRecap2_ProcessDesc.Index).Value = Me.ObjProcess.GetDesc(rowScope.primary_process)
                    Next
                End If
            Next

            Me.dgvItemCost_AllocationRecap2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TampilkanOverhead_AccountAllocation()
        Dim Identifier As String = ""
        Dim rowAccount As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim rowDGV As DataGridViewRow
        Dim DefPrec As Integer = 0
        Dim RefIndex As Integer = 0

        Try
            Me.dgvItemCost_CostAccountAllocation.AllowUserToAddRows = False
            Me.dgvItemCost_CostAccountAllocation.Rows.Clear()

            RefIndex = Me.dgvItemCost_CostGroup.Rows.Count - 1

            If Me.dgvItemCost_CostGroup.SelectedCells.Count > 0 Then
                If Me.dgvItemCost_CostGroup.SelectedCells(0).RowIndex <> RefIndex Then
                    DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
                    Identifier = Me.dgvItemCost_CostGroup.Item(Me.ColCostGroup_Identifier.Index, Me.dgvItemCost_CostGroup.SelectedCells(0).RowIndex).Value

                    'tampilkan isi account allocation
                    Me.vTBLCostGroup_Account.DefaultView.RowFilter = "identifier = '" + Identifier + "'"
                    For Each DV As DataRowView In Me.vTBLCostGroup_Account.DefaultView
                        rowAccount = DV.Row
                        Me.dgvItemCost_CostAccountAllocation.Rows.Add()
                        rowDGV = Me.dgvItemCost_CostAccountAllocation.Rows(Me.dgvItemCost_CostAccountAllocation.Rows.Count - 1)
                        rowDGV.Cells(Me.ColAccAll_AccountCode.Index).Value = Me.ObjAkun.GetAccountCode(rowAccount.primary_akun)
                        rowDGV.Cells(Me.ColAccAll_AccountDesc.Index).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(rowAccount.primary_akun)
                        rowDGV.Cells(Me.ColAccAll_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping(rowAccount.allocated_amount, DefPrec)
                        rowDGV.Cells(Me.ColAccAll_AllocatedPerct.Index).Value = Me.ObjCD.DigitGrouping(rowAccount.perc_allocated, 2)
                        rowDGV.Cells(Me.ColAccAll_EndBalance.Index).Value = Me.ObjCD.DigitGrouping(rowAccount.balance, DefPrec)
                        rowDGV.Cells(Me.ColAccAll_Identifier.Index).Value = Identifier
                        rowDGV.Cells(Me.ColAccAll_PrimaryAkun.Index).Value = rowAccount.primary_akun
                        rowDGV.Cells(Me.ColAccAll_Remarks.Index).Value = rowAccount.remarks
                        If rowAccount.ignore_negative = 1 Then
                            rowDGV.Cells(Me.ColAccAll_IgnoreNegative.Index).Value = Boolean.TrueString
                        Else
                            rowDGV.Cells(Me.ColAccAll_IgnoreNegative.Index).Value = Boolean.FalseString
                        End If
                    Next

                    'cek apakah boleh add new row
                    Me.dgvItemCost_CostAccountAllocation.AllowUserToAddRows = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TampilkanOverhead_Scope()
        Dim Identifier As String = ""
        Dim rowScope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow
        Dim rowDGV As DataGridViewRow
        Dim RefIndex As Integer = -1

        Try
            RefIndex = Me.dgvItemCost_CostGroup.Rows.Count - 1

            Me.dgvItemCost_Scope.AllowUserToAddRows = False
            Me.dgvItemCost_Scope.Rows.Clear()

            If Me.dgvItemCost_CostGroup.SelectedCells.Count > 0 Then
                If Me.dgvItemCost_CostGroup.SelectedCells(0).RowIndex <> RefIndex Then
                    Identifier = Me.dgvItemCost_CostGroup.Item(Me.ColCostGroup_Identifier.Index, Me.dgvItemCost_CostGroup.SelectedCells(0).RowIndex).Value

                    'tampilkan isi account scope
                    Me.vTBLCostGroup_Scope.DefaultView.RowFilter = "identifier = '" + Identifier + "'"
                    For Each DV As DataRowView In Me.vTBLCostGroup_Scope.DefaultView
                        rowScope = DV.Row
                        Me.dgvItemCost_Scope.Rows.Add()
                        rowDGV = Me.dgvItemCost_Scope.Rows(Me.dgvItemCost_Scope.Rows.Count - 1)
                        rowDGV.Cells(Me.ColAllScope_Identifier.Index).Value = Identifier
                        rowDGV.Cells(Me.ColAllScope_No.Index).Value = rowDGV.Index + 1
                        rowDGV.Cells(Me.ColAllScope_PrimaryProcess.Index).Value = rowScope.primary_process
                        rowDGV.Cells(Me.ColAllScope_ProcessCode.Index).Value = Me.ObjProcess.GetCode(rowScope.primary_process)
                        rowDGV.Cells(Me.ColAllScope_ProcessDesc.Index).Value = Me.ObjProcess.GetDesc(rowScope.primary_process)
                        rowDGV.Cells(Me.ColAllScope_Remarks.Index).Value = rowScope.remarks
                    Next

                    'cek apakah boleh add new row
                    Me.dgvItemCost_Scope.AllowUserToAddRows = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItemCost_CostGroup.UserDeletingRow
        Dim vRowCostAccount As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim vRowCostScope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow
        Dim vRowCostCodeNo As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noRow

        Try
            Me.vTBLCostGroup_Account.DefaultView.RowFilter = "[Identifier] = '" & Me.dgvItemCost_CostGroup.Item(Me.ColCostGroup_Identifier.Index, e.Row.Index).Value.ToString & "'"
            Me.vTBLCostGroup_Scope.DefaultView.RowFilter = "[Identifier] = '" & Me.dgvItemCost_CostGroup.Item(Me.ColCostGroup_Identifier.Index, e.Row.Index).Value.ToString & "'"
            Me.vTBLCostGroup_CodeNo.DefaultView.RowFilter = "[Identifier] = '" & Me.dgvItemCost_CostGroup.Item(Me.ColCostGroup_Identifier.Index, e.Row.Index).Value.ToString & "'"
            While Me.vTBLCostGroup_Account.DefaultView.Count > 0
                vRowCostAccount = Me.vTBLCostGroup_Account.DefaultView(0).Row
                Me.vTBLCostGroup_Account.Rows.Remove(vRowCostAccount)
            End While
            While Me.vTBLCostGroup_Scope.DefaultView.Count > 0
                vRowCostScope = Me.vTBLCostGroup_Scope.DefaultView(0).Row
                Me.vTBLCostGroup_Scope.Rows.Remove(vRowCostScope)
            End While
            While Me.vTBLCostGroup_CodeNo.DefaultView.Count > 0
                vRowCostCodeNo = Me.vTBLCostGroup_CodeNo.DefaultView(0).Row
                Me.vTBLCostGroup_CodeNo.Rows.Remove(vRowCostCodeNo)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetNextIdentifierCostGroup(ByVal vDGVItem As DataGridView, _
    ByVal vPrimaryTransaksi As String, ByVal Frm As FrmTransaksiMonthlyClosing2_Overhead2) As String
        Dim rowDGV As DataGridViewRow
        Dim MaxVal As Double = 0
        Dim Index As Integer = -1
        Try
            For Each rowDGV In vDGVItem.Rows
                If rowDGV.Cells(Frm.ColCostGroup_Identifier.Index).Value <> "" Then
                    Index = rowDGV.Cells(Frm.ColCostGroup_Identifier.Index).Value.ToString.LastIndexOf("-")
                    If Index <> -1 Then
                        If MaxVal < CDbl(rowDGV.Cells(Frm.ColCostGroup_Identifier.Index).Value.ToString.Substring(Index + 1)) Then
                            MaxVal = rowDGV.Cells(Frm.ColCostGroup_Identifier.Index).Value.ToString.Substring(Index + 1)
                        End If
                    End If
                End If
            Next
            Return vPrimaryTransaksi + "-" + (MaxVal + 1).ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "-1-0"
        End Try
    End Function

    Private Sub ShowRowProcess(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            If row.Index < Me.dgvItemCost_CostGroup.FirstDisplayedScrollingRowIndex Then
                While row.Displayed = False
                    Me.dgvItemCost_CostGroup.FirstDisplayedScrollingRowIndex = Me.dgvItemCost_CostGroup.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItemCost_CostGroup.FirstDisplayedScrollingRowIndex Then
                While Me.dgvItemCost_CostGroup.Rows(row.Index).Displayed = False
                    Me.dgvItemCost_CostGroup.FirstDisplayedScrollingRowIndex = Me.dgvItemCost_CostGroup.FirstDisplayedScrollingRowIndex + 1
                End While

                If Me.dgvItemCost_CostGroup.DisplayedRowCount(True) <> Me.dgvItemCost_CostGroup.DisplayedRowCount(False) Then
                    Me.dgvItemCost_CostGroup.FirstDisplayedScrollingRowIndex = Me.dgvItemCost_CostGroup.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.dgvItemCost_CostGroup.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemCost_CostGroup.SelectedRows(0)

                If row.Index = 0 Then
                    Me.dgvItemCost_CostGroup.Focus()
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                    Exit Sub
                End If
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItemCost_CostGroup.Rows.Remove(row)
                    Me.dgvItemCost_CostGroup.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemCost_CostGroup.Focus()
                    Me.dgvItemCost_CostGroup.ClearSelection()
                    row.Selected = True
                    If row.Displayed = False Then
                        Me.ShowRowProcess(row, True)
                    End If
                End If
                Me.dgvItemCost_CostGroup.AllowUserToAddRows = False
                Me.dgvItemCost_CostGroup.AllowUserToAddRows = True

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvItemCost_CostGroup.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemCost_CostGroup.SelectedRows(0)

                If row.Index = Me.dgvItemCost_CostGroup.RowCount - 2 Then
                    Me.dgvItemCost_CostGroup.Focus()
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                    Exit Sub
                End If
                If row.Index <> Me.dgvItemCost_CostGroup.Rows.Count - 2 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItemCost_CostGroup.Rows.Remove(row)
                    Me.dgvItemCost_CostGroup.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemCost_CostGroup.Focus()
                    Me.dgvItemCost_CostGroup.ClearSelection()
                    row.Selected = True
                    Me.ShowRowProcess(row, False)
                End If
            End If
            Me.dgvItemCost_CostGroup.AllowUserToAddRows = False
            Me.dgvItemCost_CostGroup.AllowUserToAddRows = True

            Me.btnUP.Enabled = True
            Me.btnDown.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowRowProcessAccount(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            If row.Index < Me.dgvItemCost_CostAccountAllocation.FirstDisplayedScrollingRowIndex Then
                While row.Displayed = False
                    Me.dgvItemCost_CostAccountAllocation.FirstDisplayedScrollingRowIndex = Me.dgvItemCost_CostAccountAllocation.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItemCost_CostAccountAllocation.FirstDisplayedScrollingRowIndex Then
                While Me.dgvItemCost_CostAccountAllocation.Rows(row.Index).Displayed = False
                    Me.dgvItemCost_CostAccountAllocation.FirstDisplayedScrollingRowIndex = Me.dgvItemCost_CostAccountAllocation.FirstDisplayedScrollingRowIndex + 1
                End While
                If Me.dgvItemCost_CostAccountAllocation.DisplayedRowCount(True) <> Me.dgvItemCost_CostAccountAllocation.DisplayedRowCount(False) Then
                    Me.dgvItemCost_CostAccountAllocation.FirstDisplayedScrollingRowIndex = Me.dgvItemCost_CostAccountAllocation.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnUpAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUpAccount.Click
        Dim rowDGV As DataGridViewRow
        Dim rowD As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim vRowCostAccount As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow

        Try
            If Me.dgvItemCost_CostAccountAllocation.SelectedRows.Count > 0 Then
                Me.BtnUpAccount.Enabled = False
                Me.BtnDownAccount.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemCost_CostAccountAllocation.SelectedRows(0)

                If row.Index = 0 Then
                    Me.dgvItemCost_CostAccountAllocation.Focus()
                    Me.BtnUpAccount.Enabled = True
                    Me.BtnDownAccount.Enabled = True
                    Exit Sub
                End If
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItemCost_CostAccountAllocation.Rows.Remove(row)
                    Me.dgvItemCost_CostAccountAllocation.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemCost_CostAccountAllocation.Focus()
                    Me.dgvItemCost_CostAccountAllocation.ClearSelection()
                    row.Selected = True
                    If row.Displayed = False Then
                        Me.ShowRowProcessAccount(row, True)
                    End If

                    Me.vTBLCostGroup_Account.DefaultView.RowFilter = "[identifier] = '" & _
                    Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_Identifier.Index, 0).Value.ToString & "'"
                    While Me.vTBLCostGroup_Account.DefaultView.Count > 0
                        vRowCostAccount = Me.vTBLCostGroup_Account.DefaultView(0).Row
                        Me.vTBLCostGroup_Account.Rows.Remove(vRowCostAccount)
                    End While
                    If Me.dgvItemCost_CostAccountAllocation.RowCount > 0 Then
                        For i As Integer = 0 To Me.dgvItemCost_CostAccountAllocation.Rows.Count - 2
                            rowDGV = Me.dgvItemCost_CostAccountAllocation.Rows(i)
                            rowD = Me.vTBLCostGroup_Account.Newtbltrans_purchase_closing_item_cost_group_accountRow
                            With rowD
                                .primary = "-1"
                                .primary_transaksi = -1
                                .primary_akun = rowDGV.Cells(Me.ColAccAll_PrimaryAkun.Index).Value
                                .identifier = rowDGV.Cells(Me.ColAccAll_Identifier.Index).Value
                                .balance = rowDGV.Cells(Me.ColAccAll_EndBalance.Index).Value
                                .perc_allocated = rowDGV.Cells(Me.ColAccAll_AllocatedPerct.Index).Value
                                .allocated_amount = rowDGV.Cells(Me.ColAccAll_AllocatedAmount.Index).Value
                                If rowDGV.Cells(Me.ColAccAll_Remarks.Index).Value = "" Then
                                    rowDGV.Cells(Me.ColAccAll_Remarks.Index).Value = ""
                                End If
                                .remarks = rowDGV.Cells(Me.ColAccAll_Remarks.Index).Value
                            End With
                            Me.vTBLCostGroup_Account.Rows.Add(rowD)
                        Next
                    End If
                    Me.dgvItemCost_CostAccountAllocation.AllowUserToAddRows = False
                    Me.dgvItemCost_CostAccountAllocation.AllowUserToAddRows = True
                End If

                Me.BtnUpAccount.Enabled = True
                Me.BtnDownAccount.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnDownAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDownAccount.Click
        Dim rowDGV As DataGridViewRow
        Dim rowD As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim vRowCostAccount As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow

        Try
            If Me.dgvItemCost_CostAccountAllocation.SelectedRows.Count > 0 Then
                Me.BtnUpAccount.Enabled = False
                Me.BtnDownAccount.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemCost_CostAccountAllocation.SelectedRows(0)

                If row.Index = Me.dgvItemCost_CostAccountAllocation.RowCount - 2 Then
                    Me.dgvItemCost_CostAccountAllocation.Focus()
                    Me.BtnUpAccount.Enabled = True
                    Me.BtnDownAccount.Enabled = True
                    Exit Sub
                End If
                If row.Index <> Me.dgvItemCost_CostAccountAllocation.Rows.Count - 2 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItemCost_CostAccountAllocation.Rows.Remove(row)
                    Me.dgvItemCost_CostAccountAllocation.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemCost_CostAccountAllocation.Focus()
                    Me.dgvItemCost_CostAccountAllocation.ClearSelection()
                    row.Selected = True
                    Me.ShowRowProcessAccount(row, False)

                    Me.vTBLCostGroup_Account.DefaultView.RowFilter = "[identifier] = '" & Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_Identifier.Index, 0).Value.ToString & "'"
                    While Me.vTBLCostGroup_Account.DefaultView.Count > 0
                        vRowCostAccount = Me.vTBLCostGroup_Account.DefaultView(0).Row
                        Me.vTBLCostGroup_Account.Rows.Remove(vRowCostAccount)
                    End While
                    If Me.dgvItemCost_CostAccountAllocation.RowCount > 0 Then
                        For i As Integer = 0 To Me.dgvItemCost_CostAccountAllocation.Rows.Count - 2
                            rowDGV = Me.dgvItemCost_CostAccountAllocation.Rows(i)
                            rowD = Me.vTBLCostGroup_Account.Newtbltrans_purchase_closing_item_cost_group_accountRow
                            With rowD
                                .primary = "-1"
                                .primary_transaksi = -1
                                .primary_akun = rowDGV.Cells(Me.ColAccAll_PrimaryAkun.Index).Value
                                .identifier = rowDGV.Cells(Me.ColAccAll_Identifier.Index).Value
                                .balance = rowDGV.Cells(Me.ColAccAll_EndBalance.Index).Value
                                .perc_allocated = rowDGV.Cells(Me.ColAccAll_AllocatedPerct.Index).Value
                                .allocated_amount = rowDGV.Cells(Me.ColAccAll_AllocatedAmount.Index).Value
                                If rowDGV.Cells(Me.ColAccAll_Remarks.Index).Value = "" Then
                                    rowDGV.Cells(Me.ColAccAll_Remarks.Index).Value = ""
                                End If
                                .remarks = rowDGV.Cells(Me.ColAccAll_Remarks.Index).Value
                            End With
                            Me.vTBLCostGroup_Account.Rows.Add(rowD)
                        Next
                    End If
                    Me.dgvItemCost_CostAccountAllocation.AllowUserToAddRows = False
                    Me.dgvItemCost_CostAccountAllocation.AllowUserToAddRows = True
                Else
                    row.Selected = True
                End If

                Me.BtnUpAccount.Enabled = True
                Me.BtnDownAccount.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowRowProcessScope(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            If row.Index < Me.dgvItemCost_Scope.FirstDisplayedScrollingRowIndex Then
                While row.Displayed = False
                    Me.dgvItemCost_Scope.FirstDisplayedScrollingRowIndex = Me.dgvItemCost_Scope.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItemCost_Scope.FirstDisplayedScrollingRowIndex Then
                While Me.dgvItemCost_Scope.Rows(row.Index).Displayed = False
                    Me.dgvItemCost_Scope.FirstDisplayedScrollingRowIndex = Me.dgvItemCost_Scope.FirstDisplayedScrollingRowIndex + 1
                End While

                If Me.dgvItemCost_Scope.DisplayedRowCount(True) <> Me.dgvItemCost_Scope.DisplayedRowCount(False) Then
                    Me.dgvItemCost_Scope.FirstDisplayedScrollingRowIndex = Me.dgvItemCost_Scope.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnUpScope_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUpScope.Click
        Dim rowDGV As DataGridViewRow
        Dim rowD As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow
        Dim vRowCostScope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow

        Try
            If Me.dgvItemCost_Scope.SelectedRows.Count > 0 Then
                Me.BtnUpScope.Enabled = False
                Me.BtnDownScope.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemCost_Scope.SelectedRows(0)

                If row.Index = 0 Then
                    Me.dgvItemCost_Scope.Focus()
                    Me.BtnUpScope.Enabled = True
                    Me.BtnDownScope.Enabled = True
                    Exit Sub
                End If
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItemCost_Scope.Rows.Remove(row)
                    Me.dgvItemCost_Scope.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemCost_Scope.Focus()
                    Me.dgvItemCost_Scope.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItemCost_Scope)
                    If row.Displayed = False Then
                        Me.ShowRowProcessScope(row, True)
                    End If

                    Me.vTBLCostGroup_Scope.DefaultView.RowFilter = "[identifier] = '" & Me.dgvItemCost_Scope.Item(Me.ColAllScope_Identifier.Index, 0).Value.ToString & "'"
                    While Me.vTBLCostGroup_Scope.DefaultView.Count > 0
                        vRowCostScope = Me.vTBLCostGroup_Scope.DefaultView(0).Row
                        Me.vTBLCostGroup_Scope.Rows.Remove(vRowCostScope)
                    End While
                    If Me.dgvItemCost_Scope.RowCount > 0 Then
                        For i As Integer = 0 To Me.dgvItemCost_Scope.Rows.Count - 2
                            rowDGV = Me.dgvItemCost_Scope.Rows(i)
                            rowD = Me.vTBLCostGroup_Scope.Newtbltrans_purchase_closing_item_cost_group_scopeRow
                            With rowD
                                .primary = "-1"
                                .primary_transaksi = -1
                                .primary_process = rowDGV.Cells(Me.ColAllScope_PrimaryProcess.Index).Value
                                .identifier = rowDGV.Cells(Me.ColAllScope_Identifier.Index).Value
                                If rowDGV.Cells(Me.ColAllScope_Remarks.Index).Value = "" Then
                                    rowDGV.Cells(Me.ColAllScope_Remarks.Index).Value = ""
                                End If
                                .remarks = rowDGV.Cells(Me.ColAllScope_Remarks.Index).Value
                            End With
                            Me.vTBLCostGroup_Scope.Rows.Add(rowD)
                        Next
                    End If
                    Me.dgvItemCost_Scope.AllowUserToAddRows = False
                    Me.dgvItemCost_Scope.AllowUserToAddRows = True
                End If

                Me.BtnUpScope.Enabled = True
                Me.BtnDownScope.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnDownScope_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDownScope.Click
        Dim rowDGV As DataGridViewRow
        Dim rowD As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow
        Dim vRowCostScope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow

        Try
            If Me.dgvItemCost_Scope.SelectedRows.Count > 0 Then
                Me.BtnUpScope.Enabled = False
                Me.BtnDownScope.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemCost_Scope.SelectedRows(0)

                If row.Index = Me.dgvItemCost_Scope.RowCount - 2 Then
                    Me.dgvItemCost_Scope.Focus()
                    Me.BtnUpScope.Enabled = True
                    Me.BtnDownScope.Enabled = True
                    Exit Sub
                End If
                If row.Index <> Me.dgvItemCost_Scope.Rows.Count - 2 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItemCost_Scope.Rows.Remove(row)
                    Me.dgvItemCost_Scope.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemCost_Scope.Focus()
                    Me.dgvItemCost_Scope.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItemCost_Scope)
                    Me.ShowRowProcessScope(row, False)

                    Me.vTBLCostGroup_Scope.DefaultView.RowFilter = "[identifier] = '" & Me.dgvItemCost_Scope.Item(Me.ColAllScope_Identifier.Index, 0).Value.ToString & "'"
                    While Me.vTBLCostGroup_Scope.DefaultView.Count > 0
                        vRowCostScope = Me.vTBLCostGroup_Scope.DefaultView(0).Row
                        Me.vTBLCostGroup_Scope.Rows.Remove(vRowCostScope)
                    End While
                    If Me.dgvItemCost_Scope.RowCount > 0 Then
                        For i As Integer = 0 To Me.dgvItemCost_Scope.Rows.Count - 2
                            rowDGV = Me.dgvItemCost_Scope.Rows(i)
                            rowD = Me.vTBLCostGroup_Scope.Newtbltrans_purchase_closing_item_cost_group_scopeRow
                            With rowD
                                .primary = "-1"
                                .primary_transaksi = -1
                                .primary_process = rowDGV.Cells(Me.ColAllScope_PrimaryProcess.Index).Value
                                .identifier = rowDGV.Cells(Me.ColAllScope_Identifier.Index).Value
                                If rowDGV.Cells(Me.ColAllScope_Remarks.Index).Value = "" Then
                                    rowDGV.Cells(Me.ColAllScope_Remarks.Index).Value = ""
                                End If
                                .remarks = rowDGV.Cells(Me.ColAllScope_Remarks.Index).Value
                            End With
                            Me.vTBLCostGroup_Scope.Rows.Add(rowD)
                        Next
                    End If
                    Me.dgvItemCost_Scope.AllowUserToAddRows = False
                    Me.dgvItemCost_Scope.AllowUserToAddRows = True
                End If

                Me.BtnUpScope.Enabled = True
                Me.BtnDownScope.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCostAccount_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemCost_CostAccountAllocation.CellBeginEdit
        Try
            'ketika mulai edit item account, minimal harus ada 1 cell pada cost group yang dipilih
            'dan itu bukan baris terakhir
            If Me.dgvItemCost_CostGroup.SelectedCells.Count = 0 Then
                MsgBox("Please specify a valid cost group first.", MsgBoxStyle.Information, "Information")
                e.Cancel = True
            Else
                If Me.dgvItemCost_CostGroup.SelectedCells(0).RowIndex = Me.dgvItemCost_CostGroup.Rows.Count - 1 Then
                    MsgBox("Please specify a valid cost group first.", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                End If
            End If

            If String.IsNullOrEmpty(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AccountCode.Index, e.RowIndex).Value) = False Then
                If (e.ColumnIndex = Me.ColAccAll_AccountCode.Index Or e.ColumnIndex = Me.ColAccAll_AllocatedPerct.Index _
                Or e.ColumnIndex = Me.ColAccAll_IgnoreNegative.Index _
                Or e.ColumnIndex = Me.ColAccAll_Remarks.Index) Then
                    Me.Buffer = Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value
                    'buang nol dibelakang decimal
                    If e.ColumnIndex = Me.ColAccAll_AllocatedPerct.Index Then
                        Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
                    End If
                ElseIf e.ColumnIndex = Me.ColAccAll_AllocatedAmount.Index Then
                    'jika ignore negative dan end balance negative, maka kolom amount tidak bisa di edit
                    If Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_IgnoreNegative.Index, e.RowIndex).Value = Boolean.TrueString And _
                    CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_EndBalance.Index, e.RowIndex).Value) < 0 Then
                        e.Cancel = True
                    Else
                        Me.Buffer = Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value
                        'buang nol dibelakang decimal
                        Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColAccAll_AccountCode.Index Then
                    Me.Buffer = Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
            If (e.ColumnIndex = Me.ColAccAll_AccountDesc.Index Or e.ColumnIndex = Me.ColAccAll_EndBalance.Index) Then
                e.Cancel = True
            End If

            If Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_EndBalance.Index, e.RowIndex).Value = 0 Then
                'If e.ColumnIndex = Me.ColAccAll_AllocatedPerct.Index Or e.ColumnIndex = Me.ColAccAll_AllocatedAmount.Index Then
                If e.ColumnIndex = Me.ColAccAll_AllocatedAmount.Index Then
                    e.Cancel = True
                    Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemCostAccount(ByVal vPrimaryAkun As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvItemCost_CostAccountAllocation.Rows.Count - 2
                RowDGV = Me.dgvItemCost_CostAccountAllocation.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColAccAll_PrimaryAkun.Index).Value = vPrimaryAkun Then
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

    Private Sub dgvCostAccount_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemCost_CostAccountAllocation.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvItemCost_CostAccountAllocation.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim PrimaryAkun As Double
        Dim PrimaryJurnal As Double
        Dim TipeAkun As Integer
        Dim Identifier As String

        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim DsBalance As New AccModule.DseTransaksiJournal
        Dim DtaBalance As New AccModule.DseTransaksiJournalTableAdapters.DataTableBalance1aTableAdapter
        Dim TblBalance As AccModule.DseTransaksiJournal.DataTableBalance1aDataTable = DsBalance.DataTableBalance1a
        Dim RowBalance As AccModule.DseTransaksiJournal.DataTableBalance1aRow
        Dim CostGroup_Balance As Decimal

        Dim rowDGV As DataGridViewRow = Me.dgvItemCost_CostAccountAllocation.Rows(e.RowIndex)
        Dim rowD As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim vRowCostAccount As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim DefPrec As Integer = 0
        Dim TotalAllocated_Perct As Decimal = 0
        Dim TotalAllocated_Amount As Decimal = 0
        Dim DV As DataRowView

        Try
            Identifier = Me.dgvItemCost_CostGroup.Item(Me.ColCostGroup_Identifier.Index, Me.dgvItemCost_CostGroup.SelectedCells(0).RowIndex).Value
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

            With row
                If IsNothing(Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColAccAll_AccountCode.Index Then
                        PrimaryJurnal = -1

                        If Me.Buffer <> "" Then
                            Me.vTBLCostGroup_Account.DefaultView.RowFilter = "[identifier] = '" & Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_Identifier.Index, e.RowIndex).Value.ToString & _
                            "' AND [primary akun] = " & Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_PrimaryAkun.Index, e.RowIndex).Value & ""
                            While Me.vTBLCostGroup_Account.DefaultView.Count > 0
                                vRowCostAccount = Me.vTBLCostGroup_Account.DefaultView(0).Row
                                Me.vTBLCostGroup_Account.Rows.Remove(vRowCostAccount)
                            End While
                        End If
                        If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryAkun = Me.ObjAkun.GetPrimary(Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value)
                            TipeAkun = Me.ObjAkun.GetTipeAkun(PrimaryAkun)
                            If AccModule.ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales = TipeAkun _
                                   Or AccModule.ClsMasterAkun.EnumTipeAkun.Expense = TipeAkun _
                                   Or AccModule.ClsMasterAkun.EnumTipeAkun.Other_Expense = TipeAkun _
                                   Or AccModule.ClsMasterAkun.EnumTipeAkun.Income = TipeAkun _
                                   Or AccModule.ClsMasterAkun.EnumTipeAkun.Other_Income = TipeAkun Then
                            Else
                                MsgBox("Please choose one of the following account type: Cost Of Goods Sales, Expense, Other Expense, Income, Or Other Income.", MsgBoxStyle.Information, "Information")
                                Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Me.mDgv = Me.dgvItemCost_CostAccountAllocation
                                Me.mEraseRowIndex = row.Index
                                Me.BgwDelCostAcc.RunWorkerAsync()
                                Exit Sub
                            End If

                            index = Me.GetIndexItemCostAccount(PrimaryAkun, e.RowIndex)
                            If index < 0 Then
                                Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_PrimaryAkun.Index, e.RowIndex).Value = PrimaryAkun
                                Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_Identifier.Index, e.RowIndex).Value = Identifier
                                Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AccountCode.Index, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(PrimaryAkun)
                                Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AccountDesc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAkun)
                                'Hitung Balance(Debit - Credit) kecuali primary jurnalnya sendiri, pada periode tertentu berdasarkan primary akun terpilih.
                                DtaBalance.FillBy_TglDariSampai_PrimaryAkun_NotPrimaryJurnal(TblBalance, PrimaryJurnal, _
                                Me.ObjCD.GetFirstDateofTheMonth(Me.ClosingPeriod), Me.ObjCD.GetLasDateOfTheMonth(Me.ClosingPeriod), PrimaryAkun)
                                CostGroup_Balance = 0
                                For Each RowBalance In TblBalance
                                    With RowBalance
                                        .Primary = .Primary
                                        If .Primary <> PrimaryJurnal Then
                                            CostGroup_Balance += .Debet - .Kredit
                                        End If
                                    End With
                                Next
                                Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_EndBalance.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CostGroup_Balance, DefPrec)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                                    Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    If Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                        Me.mDgv = Me.dgvItemCost_CostAccountAllocation
                                        Me.mEraseRowIndex = row.Index
                                        Me.BgwDelCostAcc.RunWorkerAsync()
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmMasterAkun_Cari1
                            ij.BackColorAll = Me.BackColorAll
                            ij.txtCari.Text = Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            index = Me.GetIndexItemCostAccount(ij.mPrimaryAkun, e.RowIndex)
                            TipeAkun = Me.ObjAkun.GetTipeAkun(ij.mPrimaryAkun)
                            If ij.mPrimaryAkun = -1 Then
                                Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Me.mDgv = Me.dgvItemCost_CostAccountAllocation
                                If Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                    Me.mEraseRowIndex = row.Index
                                    Me.BgwDelCostAcc.RunWorkerAsync()
                                End If
                                Exit Sub
                            Else
                                If AccModule.ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales = TipeAkun _
                                   Or AccModule.ClsMasterAkun.EnumTipeAkun.Expense = TipeAkun _
                                   Or AccModule.ClsMasterAkun.EnumTipeAkun.Other_Expense = TipeAkun _
                                   Or AccModule.ClsMasterAkun.EnumTipeAkun.Income = TipeAkun _
                                   Or AccModule.ClsMasterAkun.EnumTipeAkun.Other_Income = TipeAkun Then
                                Else
                                    MsgBox("Please choose one of the following account type: Cost Of Goods Sales, Expense, Other Expense, Income, Or Other Income.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Me.mDgv = Me.dgvItemCost_CostAccountAllocation
                                    Me.mEraseRowIndex = row.Index
                                    Me.BgwDelCostAcc.RunWorkerAsync()
                                    Exit Sub
                                End If

                                If index < 0 Then
                                    Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_PrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                    Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_Identifier.Index, e.RowIndex).Value = Identifier
                                    Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AccountCode.Index, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                                    Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AccountDesc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                                    'Hitung Balance(Debit - Credit) kecuali primary jurnalnya sendiri, pada periode tertentu berdasarkan primary akun terpilih.
                                    DtaBalance.FillBy_TglDariSampai_PrimaryAkun_NotPrimaryJurnal(TblBalance, PrimaryJurnal, _
                                    Me.ObjCD.GetFirstDateofTheMonth(Me.ClosingPeriod), Me.ObjCD.GetLasDateOfTheMonth(Me.ClosingPeriod), ij.mPrimaryAkun)
                                    CostGroup_Balance = 0
                                    For Each RowBalance In TblBalance
                                        With RowBalance
                                            .Primary = .Primary
                                            If .Primary <> PrimaryJurnal Then
                                                CostGroup_Balance += .Debet - .Kredit
                                            End If
                                        End With
                                    Next
                                    Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_EndBalance.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CostGroup_Balance, DefPrec)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                                        Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        If Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                            Me.mDgv = Me.dgvItemCost_CostAccountAllocation
                                            Me.mEraseRowIndex = row.Index
                                            Me.BgwDelCostAcc.RunWorkerAsync()
                                        End If
                                        Exit Sub
                                    End If
                                End If
                            End If
                            ij = Nothing
                        End If
                    End If

                    If e.ColumnIndex = Me.ColAccAll_AllocatedPerct.Index Or e.ColumnIndex = Me.ColAccAll_AllocatedAmount.Index Or _
                    e.ColumnIndex = Me.ColAccAll_IgnoreNegative.Index Then
                        If IsNumeric(Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Else
                            Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If

                        If .Cells(Me.ColAccAll_EndBalance.Index).Value = 0 And e.ColumnIndex = Me.ColAccAll_AllocatedAmount.Index Then
                            Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, 2)
                            Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedAmount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                        Else
                            TotalAllocated_Perct = 0
                            TotalAllocated_Amount = 0
                            Me.vTBLCostGroup_Account.DefaultView.RowFilter = "[primary akun] = " & Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_PrimaryAkun.Index, e.RowIndex).Value.ToString & ""
                            For Each DV In Me.vTBLCostGroup_Account.DefaultView
                                vRowCostAccount = DV.Row
                                If vRowCostAccount.identifier <> Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_Identifier.Index, _
                                e.RowIndex).Value.ToString Then
                                    TotalAllocated_Perct += vRowCostAccount.perc_allocated
                                    'disini tidak menggunakan allocated amount secara langsung, karena jika ignore negative dan balance-nya negative
                                    'maka allocated amount selalu 0, yang akan mengakibatkan perhitungannya ngaco
                                    TotalAllocated_Amount += Math.Round(vRowCostAccount.perc_allocated * vRowCostAccount.balance / 100, DefPrec, MidpointRounding.AwayFromZero)
                                End If
                            Next

                            If e.ColumnIndex = Me.ColAccAll_AllocatedPerct.Index Then
                                If Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, e.RowIndex).Value > 0 Then
                                    If Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, e.RowIndex).Value <= 100 Then
                                        If TotalAllocated_Perct + CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, _
                                        e.RowIndex).Value) > 100 Then
                                            'cek total allocated, pastikan tidak bisa > 100%
                                            MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(100 - TotalAllocated_Perct, 2) + " %", MsgBoxStyle.Information, "Information")
                                            Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, e.RowIndex).Value = Me.Buffer
                                            Exit Sub
                                        ElseIf TotalAllocated_Perct + CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, _
                                        e.RowIndex).Value) = 100 Then
                                            'jika allocated = 100%, maka dihitung berdasarkan sisanya
                                            row.Cells(Me.ColAccAll_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAccAll_EndBalance.Index).Value) - TotalAllocated_Amount, DefPrec)
                                            Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(100 - TotalAllocated_Perct, 2)
                                        Else
                                            'jika allocated < 100, maka dihitung berdasarkan persentase
                                            row.Cells(Me.ColAccAll_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping((row.Cells(Me.ColAccAll_EndBalance.Index).Value * row.Cells(Me.ColAccAll_AllocatedPerct.Index).Value) / 100, DefPrec)
                                            Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                                        End If
                                    Else
                                        MsgBox("Allocated % must be <= 100%", MsgBoxStyle.Information, "Information")
                                        Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, e.RowIndex).Value = Me.Buffer
                                        Exit Sub
                                    End If
                                Else
                                    Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            ElseIf e.ColumnIndex = Me.ColAccAll_IgnoreNegative.Index Then
                                If Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.TrueString Then
                                    'jika true, jika balance < 0, maka set allocated amount = 0
                                    If CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_EndBalance.Index, e.RowIndex).Value) < 0 Then
                                        Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedAmount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                                    End If
                                Else
                                    'hitung ulang nilai allocated amount berdasarkan allocated %
                                    If TotalAllocated_Perct + CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, _
                                    e.RowIndex).Value) = 100 Then
                                        'jika allocated = 100%, maka dihitung berdasarkan sisanya
                                        row.Cells(Me.ColAccAll_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAccAll_EndBalance.Index).Value) - TotalAllocated_Amount, DefPrec)
                                    Else
                                        'jika allocated < 100, maka dihitung berdasarkan persentase
                                        row.Cells(Me.ColAccAll_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping((row.Cells(Me.ColAccAll_EndBalance.Index).Value * row.Cells(Me.ColAccAll_AllocatedPerct.Index).Value) / 100, DefPrec)
                                    End If
                                End If
                            ElseIf e.ColumnIndex = Me.ColAccAll_AllocatedAmount.Index Then
                                If CDec(row.Cells(Me.ColAccAll_EndBalance.Index).Value) > 0 Then
                                    If CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedAmount.Index, e.RowIndex).Value) <= CDec(row.Cells(Me.ColAccAll_EndBalance.Index).Value) Then
                                        If TotalAllocated_Amount + CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedAmount.Index, _
                                        e.RowIndex).Value) > CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_EndBalance.Index, _
                                        e.RowIndex).Value) Then
                                            'cek total allocated, pastikan tidak bisa > 100%
                                            MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAccAll_EndBalance.Index).Value) - TotalAllocated_Amount, DefPrec), MsgBoxStyle.Information, "Information")
                                            Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, e.RowIndex).Value = Me.Buffer
                                            Exit Sub
                                        ElseIf TotalAllocated_Amount + CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedAmount.Index, _
                                        e.RowIndex).Value) = CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_EndBalance.Index, _
                                        e.RowIndex).Value) Then
                                            'jika allocated = 100%, maka dihitung berdasarkan sisanya
                                            row.Cells(Me.ColAccAll_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAccAll_EndBalance.Index).Value) - TotalAllocated_Amount, DefPrec)
                                            Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(100 - TotalAllocated_Perct, 2)
                                        Else
                                            'jika allocated < 100 %, maka dihitung berdasarkan persentase biasa aj
                                            row.Cells(Me.ColAccAll_AllocatedPerct.Index).Value = Me.ObjCD.DigitGrouping((row.Cells(Me.ColAccAll_AllocatedAmount.Index).Value * 100) / row.Cells(Me.ColAccAll_EndBalance.Index).Value, 2)
                                            Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedAmount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value, DefPrec)
                                        End If
                                    Else
                                        MsgBox("Allocated amount must be <= balance", MsgBoxStyle.Information, "Information")
                                        Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedAmount.Index, e.RowIndex).Value = Me.Buffer
                                        Exit Sub
                                    End If
                                ElseIf CDec(row.Cells(Me.ColAccAll_EndBalance.Index).Value) < 0 Then
                                    If CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedAmount.Index, e.RowIndex).Value) < 0 Then
                                        If CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedAmount.Index, e.RowIndex).Value) >= CDec(row.Cells(Me.ColAccAll_EndBalance.Index).Value) Then
                                            If TotalAllocated_Amount + CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedAmount.Index, _
                                            e.RowIndex).Value) < CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_EndBalance.Index, _
                                            e.RowIndex).Value) Then
                                                'cek total allocated, pastikan tidak bisa > 100%
                                                MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAccAll_EndBalance.Index).Value) - TotalAllocated_Amount, DefPrec), MsgBoxStyle.Information, "Information")
                                                Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, e.RowIndex).Value = Me.Buffer
                                                Exit Sub
                                            ElseIf TotalAllocated_Amount + CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedAmount.Index, _
                                            e.RowIndex).Value) = CDec(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_EndBalance.Index, _
                                            e.RowIndex).Value) Then
                                                'jika allocated = 100%, maka dihitung berdasarkan sisanya
                                                row.Cells(Me.ColAccAll_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAccAll_EndBalance.Index).Value) - TotalAllocated_Amount, DefPrec)
                                                Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedPerct.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(100 - TotalAllocated_Perct, 2)
                                            Else
                                                'jika allocated < 100 %, maka dihitung berdasarkan persentase biasa aj
                                                row.Cells(Me.ColAccAll_AllocatedPerct.Index).Value = Me.ObjCD.DigitGrouping((row.Cells(Me.ColAccAll_AllocatedAmount.Index).Value * 100) / row.Cells(Me.ColAccAll_EndBalance.Index).Value, 2)
                                                Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedAmount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value, DefPrec)
                                            End If
                                        Else
                                            MsgBox("Allocated amount must be >= balance.", MsgBoxStyle.Information, "Information")
                                            Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedAmount.Index, e.RowIndex).Value = Me.Buffer
                                            Exit Sub
                                        End If
                                    Else
                                        MsgBox("Please enter negative number.", MsgBoxStyle.Information, "Information")
                                        Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_AllocatedAmount.Index, e.RowIndex).Value = Me.Buffer
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If

                    'Simpan Ke DataTable cost Account tutup
                Else
                    If e.ColumnIndex = Me.ColAccAll_Remarks.Index Then
                        Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    ElseIf e.ColumnIndex = Me.ColAccAll_AccountCode.Index Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            If Me.dgvItemCost_CostAccountAllocation.Rows.Count - 1 <> e.RowIndex Then
                                Me.mEraseRowIndex = e.RowIndex
                                Me.BgwDelCostAcc.RunWorkerAsync()
                            End If
                        Else
                            Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        Exit Sub
                    Else
                        Me.dgvItemCost_CostAccountAllocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If

                'Simpan Ke Datatable cost Account
                If e.ColumnIndex = Me.ColAccAll_AccountCode.Index Or e.ColumnIndex = Me.ColAccAll_AllocatedPerct.Index _
                Or e.ColumnIndex = Me.ColAccAll_AllocatedAmount.Index _
                Or e.ColumnIndex = Me.ColAccAll_Remarks.Index _
                Or e.ColumnIndex = Me.ColAccAll_IgnoreNegative.Index Then
                    If IsNothing(row.Cells(Me.ColAccAll_PrimaryAkun.Index).Value) = False Then
                        Me.vTBLCostGroup_Account.DefaultView.RowFilter = "[identifier] = '" & Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_Identifier.Index, e.RowIndex).Value.ToString & _
                        "' AND [primary akun] = " & Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_PrimaryAkun.Index, e.RowIndex).Value & ""
                        If Me.vTBLCostGroup_Account.DefaultView.Count > 0 Then
                            vRowCostAccount = Me.vTBLCostGroup_Account.DefaultView(0).Row
                            With vRowCostAccount
                                .primary_transaksi = -1
                                .primary_akun = rowDGV.Cells(Me.ColAccAll_PrimaryAkun.Index).Value
                                .identifier = rowDGV.Cells(Me.ColAccAll_Identifier.Index).Value
                                .balance = rowDGV.Cells(Me.ColAccAll_EndBalance.Index).Value
                                .perc_allocated = rowDGV.Cells(Me.ColAccAll_AllocatedPerct.Index).Value
                                .allocated_amount = rowDGV.Cells(Me.ColAccAll_AllocatedAmount.Index).Value
                                If rowDGV.Cells(Me.ColAccAll_IgnoreNegative.Index).Value = Boolean.TrueString Then
                                    .ignore_negative = 1
                                Else
                                    .ignore_negative = 0
                                End If
                                If Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_Remarks.Index, e.RowIndex).Value = "" Then
                                    rowDGV.Cells(Me.ColAccAll_Remarks.Index).Value = ""
                                End If
                                .remarks = rowDGV.Cells(Me.ColAccAll_Remarks.Index).Value
                            End With
                        Else
                            rowD = Me.vTBLCostGroup_Account.Newtbltrans_purchase_closing_item_cost_group_accountRow
                            With rowD
                                .primary = Me.GetPrimaryMax_AllocatedAcc + 1
                                .primary_transaksi = -1
                                .primary_akun = rowDGV.Cells(Me.ColAccAll_PrimaryAkun.Index).Value
                                .identifier = rowDGV.Cells(Me.ColAccAll_Identifier.Index).Value
                                .balance = rowDGV.Cells(Me.ColAccAll_EndBalance.Index).Value
                                .perc_allocated = rowDGV.Cells(Me.ColAccAll_AllocatedPerct.Index).Value
                                .allocated_amount = rowDGV.Cells(Me.ColAccAll_AllocatedAmount.Index).Value
                                If Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_Remarks.Index, e.RowIndex).Value = "" Then
                                    rowDGV.Cells(Me.ColAccAll_Remarks.Index).Value = ""
                                End If
                                .remarks = rowDGV.Cells(Me.ColAccAll_Remarks.Index).Value
                                If rowDGV.Cells(Me.ColAccAll_IgnoreNegative.Index).Value = Boolean.TrueString Then
                                    .ignore_negative = 1
                                Else
                                    .ignore_negative = 0
                                End If
                            End With
                            Me.vTBLCostGroup_Account.Rows.Add(rowD)
                        End If


                        'update ke dgv cost group
                        Me.Update_Cost_Group_Overhead_Amount(Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_Identifier.Index, _
                        e.RowIndex).Value.ToString)
                    End If
                Else
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DsBalance.Dispose()
            DtaBalance.Dispose()
            TblBalance.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Private Sub dgvCostAccount_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemCost_CostAccountAllocation.RowsAdded
        Dim rowDGV2 As DataGridViewRow
        Dim beforeIndexIdentifier As Integer = e.RowIndex - 1
        Dim beforeindexRefUnit As Integer = e.RowIndex - 2
        Dim PrimaryRefUnit As Double = -1
        Dim identifier As String = ""
        Dim DefPrec As Integer = 0

        Try
            Me.dgvItemCost_CostAccountAllocation.TabStop = True
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

            If Me.dgvItemCost_CostGroup.SelectedCells.Count > 0 Then
                identifier = Me.dgvItemCost_CostGroup.Item(Me.ColCostGroup_Identifier.Index, Me.dgvItemCost_CostGroup.SelectedCells(0).RowIndex).Value
            Else
                identifier = ""
            End If

            'If beforeIndexIdentifier >= 0 Then
            rowDGV2 = Me.dgvItemCost_CostAccountAllocation.Rows(e.RowIndex)
            With rowDGV2
                'identifier = Me.dgvItemCost_CostAccountAllocation.Item(Me.ColAccAll_Identifier.Index, beforeIndexIdentifier).Value
                .Cells(Me.ColAccAll_Identifier.Index).Value = identifier
                .Cells(Me.ColAccAll_IgnoreNegative.Index).Value = Boolean.FalseString
                .Cells(Me.ColAccAll_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                .Cells(Me.ColAccAll_AllocatedPerct.Index).Value = Me.ObjCD.DigitGrouping(0, 2)
                .Cells(Me.ColAccAll_EndBalance.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
            End With
            'End If
            Me.BtnUpAccount.Enabled = True
            Me.BtnDownAccount.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCostAccount_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemCost_CostAccountAllocation.RowsRemoved
        Try
            If Me.dgvItemCost_CostAccountAllocation.Rows.Count = 0 Then
                Me.dgvItemCost_CostAccountAllocation.TabStop = False
                Me.BtnUpAccount.Enabled = False
                Me.BtnDownAccount.Enabled = False

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCostAccount_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemCost_CostAccountAllocation.SelectionChanged
        Try
            If Me.dgvItemCost_CostAccountAllocation.SelectedRows.Count > 0 And Me.dgvItemCost_CostAccountAllocation.Rows.Count > 1 Then
                Me.BtnUpAccount.Enabled = True
                Me.BtnDownAccount.Enabled = True
                Dim row As DataGridViewRow = Me.dgvItemCost_CostAccountAllocation.SelectedRows(0)
                If row.Index = Me.dgvItemCost_CostAccountAllocation.RowCount - 1 Then
                    row.Selected = False
                    Exit Sub
                End If
            Else
                Me.BtnUpAccount.Enabled = False
                Me.BtnDownAccount.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemCost_CostAccountAllocation_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemCost_CostAccountAllocation.UserDeletedRow
        Try
            Me.Update_Cost_Group_Overhead_Amount(Me.Buffer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCostAccount_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItemCost_CostAccountAllocation.UserDeletingRow
        Dim vRowCostAccount As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItemCost_CostAccountAllocation.Rows(e.Row.Index)
            vTBLCostGroup_Account.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(Me.ColAccAll_Identifier.Index).Value.ToString + "' AND [primary akun] = " + rowDGV.Cells(Me.ColAccAll_PrimaryAkun.Index).Value.ToString

            While Me.vTBLCostGroup_Account.DefaultView.Count > 0
                vRowCostAccount = Me.vTBLCostGroup_Account.DefaultView(0).Row
                Me.vTBLCostGroup_Account.Rows.Remove(vRowCostAccount)
            End While
            Me.Buffer = rowDGV.Cells(Me.ColAccAll_Identifier.Index).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl2.SelectedIndexChanged
        Try
            Me.TampilkanOverheadDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemCost_Scope.CellBeginEdit
        Try
            'ketika mulai edit item account, minimal harus ada 1 cell pada cost group yang dipilih
            'dan itu bukan baris terakhir
            If Me.dgvItemCost_CostGroup.SelectedCells.Count = 0 Then
                MsgBox("Please specify a valid cost group first.", MsgBoxStyle.Information, "Information")
                e.Cancel = True
            Else
                If Me.dgvItemCost_CostGroup.SelectedCells(0).RowIndex = Me.dgvItemCost_CostGroup.Rows.Count - 1 Then
                    MsgBox("Please specify a valid cost group first.", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                End If
            End If

            If String.IsNullOrEmpty(Me.dgvItemCost_Scope.Item(Me.ColAllScope_ProcessCode.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColAllScope_ProcessCode.Index Or e.ColumnIndex = Me.ColAllScope_Remarks.Index Then
                    Me.Buffer = Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColAllScope_ProcessCode.Index Then
                    Me.Buffer = Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem_CostScope(ByVal vPrimaryProcess As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1

        Try
            For ij As Integer = 0 To Me.dgvItemCost_Scope.Rows.Count - 2
                RowDGV = Me.dgvItemCost_Scope.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColAllScope_PrimaryProcess.Index).Value = vPrimaryProcess Then
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

    Private Sub dgvItemProductGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemCost_Scope.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvItemCost_Scope.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim PrimaryProcess As Double = -1

        Dim rowDGV As DataGridViewRow = Me.dgvItemCost_Scope.Rows(e.RowIndex)
        Dim rowD As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow
        Dim vRowCostScope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow
        Dim count As New Integer
        Dim Identifier As String = ""

        Try
            Identifier = Me.dgvItemCost_CostGroup.Item(Me.ColCostGroup_Identifier.Index, Me.dgvItemCost_CostGroup.SelectedCells(0).RowIndex).Value

            With row
                If IsNothing(Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColAllScope_ProcessCode.Index Then
                        ''data process 
                        Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjProcess.ApakahSudahAda(Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            'pastikan tidak ada kode process yang double
                            PrimaryProcess = Me.ObjProcess.GetPrimary(Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItem_CostScope(PrimaryProcess, e.RowIndex)
                            If index = -1 Then
                                'kode process yang diketik ada --> load data primary process + desc
                                Me.dgvItemCost_Scope.Item(Me.ColAllScope_PrimaryProcess.Index, e.RowIndex).Value = PrimaryProcess
                                Me.dgvItemCost_Scope.Item(Me.ColAllScope_ProcessDesc.Index, e.RowIndex).Value = Me.ObjProcess.GetDesc(Me.dgvItemCost_Scope.Item(Me.ColAllScope_PrimaryProcess.Index, e.RowIndex).Value)
                                Me.dgvItemCost_Scope.Item(Me.ColAllScope_Identifier.Index, e.RowIndex).Value = Identifier
                            Else
                                MsgBox("Process code duplication, please choose other process.", MsgBoxStyle.Information, "Information")

                                'kode process yang diketikkan sudah ada sebelumnya
                                If Me.Buffer = "" Then
                                    'hapuskan baris
                                    mEraseRowIndex = e.RowIndex
                                    Me.BgwDelCostScope.RunWorkerAsync()
                                Else
                                    Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                                Exit Sub
                            End If
                        Else
                            '2. kode process yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New Core_Module.FrmMasterProsesCari
                            ij.BackColorAll = Me.BackColorAll
                            ij.txtCari.Text = Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()

                            If ij.mPrimary <> -1 Then
                                index = Me.GetIndexItem_CostScope(ij.mPrimary, e.RowIndex)
                                If index = -1 Then
                                    'kode process yang dicari tidak double --> load data primary process + kode process + desc
                                    Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProcess.GetCode(ij.mPrimary)
                                    Me.dgvItemCost_Scope.Item(Me.ColAllScope_PrimaryProcess.Index, e.RowIndex).Value = ij.mPrimary
                                    Me.dgvItemCost_Scope.Item(Me.ColAllScope_ProcessDesc.Index, e.RowIndex).Value = Me.ObjProcess.GetDesc(ij.mPrimary)
                                    Me.dgvItemCost_Scope.Item(Me.ColAllScope_Identifier.Index, e.RowIndex).Value = Identifier
                                Else
                                    'kode process yang dicari double
                                    MsgBox("Process code duplication, please choose other process.", MsgBoxStyle.Information, "Information")

                                    If Me.Buffer = "" Then
                                        'hapuskan baris
                                        mEraseRowIndex = e.RowIndex
                                        Me.BgwDelCostScope.RunWorkerAsync()
                                    Else
                                        Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Exit Sub
                                End If
                            ElseIf ij.mPrimary = -1 Then
                                'kode process yang dicari tidak ada --> kembalikan ke kondisi buffer
                                If Me.Buffer = "" Then
                                    'hapuskan baris
                                    mEraseRowIndex = e.RowIndex
                                    Me.BgwDelCostScope.RunWorkerAsync()
                                Else
                                    Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                                Exit Sub
                            End If
                            ij = Nothing
                        End If
                    ElseIf e.ColumnIndex = Me.ColAllScope_Remarks.Index Then
                        'terima aj apa adanya, batasan lenght sudah ada disetting kolom
                    End If
                Else
                    If e.ColumnIndex = Me.ColAllScope_Remarks.Index Then
                        Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    ElseIf e.ColumnIndex = Me.ColAllScope_ProcessCode.Index Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            If Me.dgvItemCost_Scope.Rows.Count - 1 <> e.RowIndex Then
                                Me.mEraseRowIndex = e.RowIndex
                                Me.BgwDelCostScope.RunWorkerAsync()
                                Exit Sub
                            End If
                        Else
                            Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    Else
                        Me.dgvItemCost_Scope.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If

                'Simpan Ke Datatable cost Scope
                If e.ColumnIndex = Me.ColAllScope_ProcessCode.Index _
                Or e.ColumnIndex = Me.ColAllScope_Remarks.Index Then
                    If IsNothing(row.Cells(Me.ColAllScope_PrimaryProcess.Index).Value) = False Then
                        Me.vTBLCostGroup_Scope.DefaultView.RowFilter = "[identifier] = '" & Me.dgvItemCost_Scope.Item(Me.ColAllScope_Identifier.Index, e.RowIndex).Value.ToString & _
                        "' AND [primary process] = " & Me.dgvItemCost_Scope.Item(Me.ColAllScope_PrimaryProcess.Index, e.RowIndex).Value & ""

                        If Me.vTBLCostGroup_Scope.DefaultView.Count > 0 Then
                            vRowCostScope = Me.vTBLCostGroup_Scope.DefaultView(0).Row
                            With vRowCostScope
                                .primary_transaksi = -1
                                .primary_process = rowDGV.Cells(Me.ColAllScope_PrimaryProcess.Index).Value
                                .identifier = rowDGV.Cells(Me.ColAllScope_Identifier.Index).Value
                                .remarks = rowDGV.Cells(Me.ColAllScope_Remarks.Index).Value
                            End With
                            Exit Sub
                        Else
                            rowD = Me.vTBLCostGroup_Scope.Newtbltrans_purchase_closing_item_cost_group_scopeRow
                            With rowD
                                .primary = Me.GetPrimaryMax_AllocatedScope + 1
                                .primary_transaksi = -1
                                .primary_process = rowDGV.Cells(Me.ColAllScope_PrimaryProcess.Index).Value
                                .identifier = rowDGV.Cells(Me.ColAllScope_Identifier.Index).Value
                                .remarks = rowDGV.Cells(Me.ColAllScope_Remarks.Index).Value
                            End With
                            Me.vTBLCostGroup_Scope.Rows.Add(rowD)
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemCost_Scope.RowsAdded
        Dim rowDGV As DataGridViewRow
        Dim beforeIndexIdentifier As Integer = e.RowIndex - 1
        Dim identifier As String = ""

        Try
            Me.dgvItemCost_Scope.TabStop = True

            rowDGV = Me.dgvItemCost_Scope.Rows(Me.dgvItemCost_Scope.RowCount - 1)
            With rowDGV
                rowDGV.Cells(Me.ColAllScope_No.Index).Value = rowDGV.Index + 1
                If String.IsNullOrEmpty(rowDGV.Cells(Me.ColAllScope_ProcessCode.Index).Value) = True Then
                    If beforeIndexIdentifier >= 0 Then
                        identifier = Me.dgvItemCost_Scope.Item(Me.ColAllScope_Identifier.Index, beforeIndexIdentifier).Value
                        .Cells(Me.ColAllScope_Identifier.Index).Value = identifier
                    End If
                End If
                .Cells(Me.ColAllScope_Remarks.Index).Value = ""
            End With
            Me.BtnUpScope.Enabled = True
            Me.BtnDownScope.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemCost_Scope.RowsRemoved
        Try
            If Me.dgvItemCost_Scope.Rows.Count = 0 Then
                Me.dgvItemCost_Scope.TabStop = False
                Me.BtnUpScope.Enabled = False
                Me.BtnDownScope.Enabled = False
            End If
            Me.ObjCD.SetNumber(Me.dgvItemCost_Scope)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemCost_Scope.SelectionChanged
        Try
            If Me.dgvItemCost_Scope.SelectedRows.Count > 0 And Me.dgvItemCost_Scope.Rows.Count > 1 Then
                Me.BtnUpScope.Enabled = True
                Me.BtnDownScope.Enabled = True
                Dim row As DataGridViewRow = Me.dgvItemCost_Scope.SelectedRows(0)
                If row.Index = Me.dgvItemCost_Scope.RowCount - 1 Then
                    row.Selected = False
                    Exit Sub
                End If
            Else
                Me.BtnUpScope.Enabled = False
                Me.BtnDownScope.Enabled = False
            End If
            Me.ObjCD.SetNumber(Me.dgvItemCost_Scope)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemCost_Scope.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItemCost_Scope)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemCost_Scope_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItemCost_Scope.UserDeletingRow
        Dim vRowScope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItemCost_Scope.Rows(e.Row.Index)
            Me.vTBLCostGroup_Scope.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(Me.ColAllScope_Identifier.Index).Value.ToString + _
            "' AND [primary process] = " + rowDGV.Cells(Me.ColAllScope_PrimaryProcess.Index).Value.ToString

            While Me.vTBLCostGroup_Scope.DefaultView.Count > 0
                vRowScope = Me.vTBLCostGroup_Scope.DefaultView(0).Row
                Me.vTBLCostGroup_Scope.Rows.Remove(vRowScope)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboGroup_Account_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboGroup_Account.CheckedChanged
        Try
            If Me.rboGroup_Account.Checked = True Then
                Me.dgvItemCost_AllocationRecap1.BringToFront()
            Else
                Me.dgvItemCost_AllocationRecap2.BringToFront()
            End If
            Me.TampilkanOverhead_AccountAllocationRecap()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemCost_AllocationRecap1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemCost_AllocationRecap1.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColAlloRecap1_AllocatedPerct.Index Or e.ColumnIndex = Me.ColAlloRecap1_AllocatedAmount.Index Then
                If Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_Balance.Index, e.RowIndex).Value = 0 And e.ColumnIndex = Me.ColAlloRecap1_AllocatedAmount.Index Then
                    e.Cancel = True
                ElseIf Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_IngoreNegativeBalance.Index, e.RowIndex).Value = Boolean.TrueString And _
                CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_Balance.Index, e.RowIndex).Value) < 0 Then
                    e.Cancel = True
                Else
                    Me.Buffer = Me.dgvItemCost_AllocationRecap1.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            ElseIf e.ColumnIndex = Me.ColAlloRecap1_IngoreNegativeBalance.Index Then
                Me.Buffer = Me.dgvItemCost_AllocationRecap1.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemCost_AllocationRecap1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemCost_AllocationRecap1.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvItemCost_AllocationRecap1.Rows(e.RowIndex)
        Dim rowDGV As DataGridViewRow = Me.dgvItemCost_AllocationRecap1.Rows(e.RowIndex)
        Dim vRowCostAccount As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim DefPrec As Integer = 0
        'Dim Identifier As String = ""
        Dim TotalAllocated_Perct As Decimal = 0
        Dim TotalAllocated_Amount As Decimal = 0
        Dim DV As DataRowView

        Try
            'Identifier = Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_Identifier.Index, e.RowIndex).Value
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

            With row
                If IsNothing(Me.dgvItemCost_AllocationRecap1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColAlloRecap1_AllocatedPerct.Index Or e.ColumnIndex = Me.ColAlloRecap1_AllocatedAmount.Index Or _
                    e.ColumnIndex = Me.ColAlloRecap1_IngoreNegativeBalance.Index Then
                        If IsNumeric(Me.dgvItemCost_AllocationRecap1.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Else
                            Me.dgvItemCost_AllocationRecap1.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If .Cells(Me.ColAlloRecap1_Balance.Index).Value = 0 And e.ColumnIndex = Me.ColAlloRecap1_AllocatedAmount.Index Then
                            Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedPerct.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, 2)
                            Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                        Else
                            TotalAllocated_Perct = 0
                            TotalAllocated_Amount = 0
                            Me.vTBLCostGroup_Account.DefaultView.RowFilter = "[primary akun] = " & Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_PrimaryAkun.Index, e.RowIndex).Value.ToString & ""
                            For Each DV In Me.vTBLCostGroup_Account.DefaultView
                                vRowCostAccount = DV.Row
                                If vRowCostAccount.identifier <> Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_Identifier.Index, _
                                e.RowIndex).Value.ToString Then
                                    TotalAllocated_Perct += vRowCostAccount.perc_allocated
                                    'disini tidak menggunakan allocated amount secara langsung, karena jika ignore negative dan balance-nya negative
                                    'maka allocated amount selalu 0, yang akan mengakibatkan perhitungannya ngaco
                                    TotalAllocated_Amount += Math.Round(vRowCostAccount.perc_allocated * vRowCostAccount.balance / 100, DefPrec, MidpointRounding.AwayFromZero)
                                End If
                            Next

                            If e.ColumnIndex = Me.ColAlloRecap1_AllocatedPerct.Index Then
                                If Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedPerct.Index, e.RowIndex).Value > 0 Then
                                    If Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedPerct.Index, e.RowIndex).Value <= 100 Then
                                        If TotalAllocated_Perct + CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedPerct.Index, _
                                        e.RowIndex).Value) > 100 Then
                                            'cek total allocated, pastikan tidak bisa > 100%
                                            MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(100 - TotalAllocated_Perct, 2) + " %", MsgBoxStyle.Information, "Information")
                                            Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedPerct.Index, e.RowIndex).Value = Me.Buffer
                                            Exit Sub
                                        ElseIf TotalAllocated_Perct + CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedPerct.Index, _
                                        e.RowIndex).Value) = 100 Then
                                            'jika allocated = 100%, maka dihitung berdasarkan sisanya
                                            row.Cells(Me.ColAlloRecap1_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAlloRecap1_Balance.Index).Value) - TotalAllocated_Amount, DefPrec)
                                            Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedPerct.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(100 - TotalAllocated_Perct, 2)
                                        Else
                                            'jika allocated < 100, maka dihitung berdasarkan persentase
                                            row.Cells(Me.ColAlloRecap1_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping((row.Cells(Me.ColAlloRecap1_Balance.Index).Value * row.Cells(Me.ColAlloRecap1_AllocatedPerct.Index).Value) / 100, DefPrec)
                                            Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedPerct.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemCost_AllocationRecap1.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                                        End If
                                    Else
                                        MsgBox("Allocated % must be <= 100%", MsgBoxStyle.Information, "Information")
                                        Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedPerct.Index, e.RowIndex).Value = Me.Buffer
                                        Exit Sub
                                    End If
                                Else
                                    Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedPerct.Index, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            ElseIf e.ColumnIndex = Me.ColAlloRecap1_IngoreNegativeBalance.Index Then
                                If Me.dgvItemCost_AllocationRecap1.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.TrueString Then
                                    'jika true, jika balance < 0, maka set allocated amount = 0
                                    If CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_Balance.Index, e.RowIndex).Value) < 0 Then
                                        Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                                    End If
                                Else
                                    'hitung ulang nilai allocated amount berdasarkan allocated %
                                    If TotalAllocated_Perct + CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedPerct.Index, _
                                    e.RowIndex).Value) = 100 Then
                                        'jika allocated = 100%, maka dihitung berdasarkan sisanya
                                        row.Cells(Me.ColAlloRecap1_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAlloRecap1_Balance.Index).Value) - TotalAllocated_Amount, DefPrec)
                                    Else
                                        'jika allocated < 100, maka dihitung berdasarkan persentase
                                        row.Cells(Me.ColAlloRecap1_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping((row.Cells(Me.ColAlloRecap1_Balance.Index).Value * row.Cells(Me.ColAlloRecap1_AllocatedPerct.Index).Value) / 100, DefPrec)
                                    End If
                                End If
                            ElseIf e.ColumnIndex = Me.ColAlloRecap1_AllocatedAmount.Index Then
                                If CDec(row.Cells(Me.ColAlloRecap1_Balance.Index).Value) > 0 Then
                                    If CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, e.RowIndex).Value) <= CDec(row.Cells(Me.ColAlloRecap1_Balance.Index).Value) Then
                                        If TotalAllocated_Amount + CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, _
                                        e.RowIndex).Value) > CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_Balance.Index, _
                                        e.RowIndex).Value) Then
                                            'cek total allocated, pastikan tidak bisa > 100%
                                            MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAlloRecap1_Balance.Index).Value) - TotalAllocated_Amount, DefPrec), MsgBoxStyle.Information, "Information")
                                            Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, e.RowIndex).Value = Me.Buffer
                                            Exit Sub
                                        ElseIf TotalAllocated_Amount + CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, _
                                        e.RowIndex).Value) = CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_Balance.Index, _
                                        e.RowIndex).Value) Then
                                            'jika allocated = 100%, maka dihitung berdasarkan sisanya
                                            row.Cells(Me.ColAlloRecap1_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAlloRecap1_Balance.Index).Value) - TotalAllocated_Amount, DefPrec)
                                            Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedPerct.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(100 - TotalAllocated_Perct, 2)
                                        Else
                                            'jika allocated < 100 %, maka dihitung berdasarkan persentase biasa aj
                                            row.Cells(Me.ColAlloRecap1_AllocatedPerct.Index).Value = Me.ObjCD.DigitGrouping((row.Cells(Me.ColAlloRecap1_AllocatedAmount.Index).Value * 100) / row.Cells(Me.ColAlloRecap1_Balance.Index).Value, 2)
                                            Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemCost_AllocationRecap1.Item(e.ColumnIndex, e.RowIndex).Value, DefPrec)
                                        End If
                                    Else
                                        MsgBox("Allocated amount must be <= balance", MsgBoxStyle.Information, "Information")
                                        Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, e.RowIndex).Value = Me.Buffer
                                        Exit Sub
                                    End If
                                ElseIf CDec(row.Cells(Me.ColAlloRecap1_Balance.Index).Value) < 0 Then
                                    If CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, e.RowIndex).Value) < 0 Then
                                        If CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, e.RowIndex).Value) >= CDec(row.Cells(Me.ColAlloRecap1_Balance.Index).Value) Then
                                            If TotalAllocated_Amount + CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, _
                                            e.RowIndex).Value) < CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_Balance.Index, _
                                            e.RowIndex).Value) Then
                                                'cek total allocated, pastikan tidak bisa > 100%
                                                MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAlloRecap1_Balance.Index).Value) - TotalAllocated_Amount, DefPrec), MsgBoxStyle.Information, "Information")
                                                Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, e.RowIndex).Value = Me.Buffer
                                                Exit Sub
                                            ElseIf TotalAllocated_Amount + CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, _
                                            e.RowIndex).Value) = CDec(Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_Balance.Index, _
                                            e.RowIndex).Value) Then
                                                'jika allocated = 100%, maka dihitung berdasarkan sisanya
                                                row.Cells(Me.ColAlloRecap1_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAlloRecap1_Balance.Index).Value) - TotalAllocated_Amount, DefPrec)
                                                Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedPerct.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(100 - TotalAllocated_Perct, 2)
                                            Else
                                                'jika allocated < 100 %, maka dihitung berdasarkan persentase biasa aj
                                                row.Cells(Me.ColAlloRecap1_AllocatedPerct.Index).Value = Me.ObjCD.DigitGrouping((row.Cells(Me.ColAlloRecap1_AllocatedAmount.Index).Value * 100) / row.Cells(Me.ColAlloRecap1_Balance.Index).Value, 2)
                                                Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemCost_AllocationRecap1.Item(e.ColumnIndex, e.RowIndex).Value, DefPrec)
                                            End If
                                        Else
                                            MsgBox("Allocated amount must be >= balance.", MsgBoxStyle.Information, "Information")
                                            Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, e.RowIndex).Value = Me.Buffer
                                            Exit Sub
                                        End If
                                    Else
                                        MsgBox("Please enter negative number.", MsgBoxStyle.Information, "Information")
                                        Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_AllocatedAmount.Index, e.RowIndex).Value = Me.Buffer
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                    'Simpan Ke DataTable cost Account tutup
                Else
                    Me.dgvItemCost_AllocationRecap1.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'Simpan Ke Datatable cost Account
                If e.ColumnIndex = Me.ColAlloRecap1_AllocatedPerct.Index _
                Or e.ColumnIndex = Me.ColAlloRecap1_AllocatedAmount.Index _
                Or e.ColumnIndex = Me.ColAlloRecap1_IngoreNegativeBalance.Index Then
                    If IsNothing(row.Cells(Me.ColAlloRecap1_PrimaryAkun.Index).Value) = False Then
                        Me.vTBLCostGroup_Account.DefaultView.RowFilter = "[identifier] = '" & Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_Identifier.Index, e.RowIndex).Value.ToString & _
                        "' AND [primary akun] = " & Me.dgvItemCost_AllocationRecap1.Item(Me.ColAlloRecap1_PrimaryAkun.Index, e.RowIndex).Value & ""
                        If Me.vTBLCostGroup_Account.DefaultView.Count > 0 Then
                            vRowCostAccount = Me.vTBLCostGroup_Account.DefaultView(0).Row
                            With vRowCostAccount
                                .primary_transaksi = -1
                                .primary_akun = rowDGV.Cells(Me.ColAlloRecap1_PrimaryAkun.Index).Value
                                .identifier = rowDGV.Cells(Me.ColAlloRecap1_Identifier.Index).Value
                                .balance = rowDGV.Cells(Me.ColAlloRecap1_Balance.Index).Value
                                .perc_allocated = rowDGV.Cells(Me.ColAlloRecap1_AllocatedPerct.Index).Value
                                .allocated_amount = rowDGV.Cells(Me.ColAlloRecap1_AllocatedAmount.Index).Value
                                If rowDGV.Cells(Me.ColAlloRecap1_IngoreNegativeBalance.Index).Value = Boolean.TrueString Then
                                    .ignore_negative = 1
                                Else
                                    .ignore_negative = 0
                                End If
                            End With
                            'Exit Sub
                        Else
                            MsgBox("Logic error, please contact yout vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If

                        'update ke dgv cost
                        Me.Update_Cost_Group_Overhead_Amount(rowDGV.Cells(Me.ColAlloRecap1_Identifier.Index).Value)
                    End If
                Else
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetCostGroup(ByVal vIdentifier As String) As String
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemCost_CostGroup.Rows
                If rowDGV.Cells(Me.ColCostGroup_Identifier.Index).Value = vIdentifier Then
                    Return rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value
                End If
            Next

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub RefreshBalance_Overhead()
        Dim DtaBalance As New AccModule.DseTransaksiJournalTableAdapters.DataTableBalance1aTableAdapter
        Dim TblBalance As New AccModule.DseTransaksiJournal.DataTableBalance1aDataTable
        Dim RowBalance As AccModule.DseTransaksiJournal.DataTableBalance1aRow
        Dim ListPrimaryAkun As New List(Of Double)
        Dim rowAcc As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim PrimaryJurnal As Double = -1
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim CostGroup_Balance As Decimal = 0
        Dim AllocatedAmount As Decimal = 0
        Dim DefPrec As Integer = 0
        Dim TotalPerct As Decimal = 0

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DtaBalance.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = DtaBalance.Connection.ConnectionString
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

            PrimaryJurnal = -1

            For Each rowAcc In Me.vTBLCostGroup_Account
                If ListPrimaryAkun.Contains(rowAcc.primary_akun) = False Then
                    ListPrimaryAkun.Add(rowAcc.primary_akun)
                End If
            Next

            For Each PrimaryAkun As Double In ListPrimaryAkun
                'masukkan nilai amount per kode akun
                CostGroup_Balance = 0
                TotalPerct = 0
                AllocatedAmount = 0

                DtaBalance.FillBy_TglDariSampai_PrimaryAkun_NotPrimaryJurnal(TblBalance, PrimaryJurnal, _
                Me.ObjCD.GetFirstDateofTheMonth(Me.ClosingPeriod), Me.ObjCD.GetLasDateOfTheMonth(Me.ClosingPeriod), PrimaryAkun)
                For Each RowBalance In TblBalance
                    With RowBalance
                        If .Primary <> PrimaryJurnal Then
                            CostGroup_Balance += .Debet - .Kredit
                        End If
                    End With
                Next

                Me.vTBLCostGroup_Account.DefaultView.RowFilter = "[primary akun] = " + PrimaryAkun.ToString
                Me.vTBLCostGroup_Account.DefaultView.Sort = "primary"

                For Each DV As DataRowView In Me.vTBLCostGroup_Account.DefaultView
                    rowAcc = DV.Row
                    TotalPerct += rowAcc.perc_allocated
                    rowAcc.balance = CostGroup_Balance
                Next
                For ij As Integer = 0 To Me.vTBLCostGroup_Account.DefaultView.Count - 1
                    rowAcc = Me.vTBLCostGroup_Account.DefaultView(ij).Row
                    If ij = vTBLCostGroup_Account.DefaultView.Count - 1 And TotalPerct = 100 Then
                        rowAcc.allocated_amount = rowAcc.balance - AllocatedAmount
                    Else
                        rowAcc.allocated_amount = Math.Round(rowAcc.perc_allocated * rowAcc.balance / 100, DefPrec, MidpointRounding.AwayFromZero)
                        AllocatedAmount += rowAcc.allocated_amount
                    End If
                Next
                For ij As Integer = 0 To Me.vTBLCostGroup_Account.DefaultView.Count - 1
                    rowAcc = Me.vTBLCostGroup_Account.DefaultView(ij).Row
                    If rowAcc.ignore_negative = 1 And rowAcc.balance < 0 Then
                        rowAcc.allocated_amount = 0
                    End If
                Next
            Next
            Me.Update_Cost_Group_Overhead_Amount()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaBalance.Dispose()
            TblBalance.Dispose()
            ListPrimaryAkun = Nothing
            DTA.Dispose()
        End Try
    End Sub

    Private Sub llbRefreshBalance_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbRefreshBalance.LinkClicked

        Try
            If MsgBox("Recalculate account balance?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                Exit Sub
            End If
            Me.RefreshBalance_Overhead()
            Me.TampilkanOverheadDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMax_AllocatedAcc() As Double
        Dim rowAcc As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim RetVal As Double = 0

        Try
            For Each rowAcc In Me.vTBLCostGroup_Account.Rows
                If rowAcc.primary > RetVal Then
                    RetVal = rowAcc.primary
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryMax_AllocatedScope() As Double
        Dim rowScope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow
        Dim RetVal As Double = 0

        Try
            For Each rowScope In Me.vTBLCostGroup_Scope.Rows
                If rowScope.primary > RetVal Then
                    RetVal = rowScope.primary
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryMax_AllocatedCodeNo() As Double
        Dim rowCodeNo As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noRow
        Dim RetVal As Double = 0

        Try
            For Each rowCodeNo In Me.vTBLCostGroup_CodeNo.Rows
                If rowCodeNo.primary > RetVal Then
                    RetVal = rowCodeNo.primary
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub RecalculateOverhead()
    '    Dim vRow As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noRow
    '    Dim vRow2 As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_no2Row
    '    Dim vRowAccount As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
    '    Dim vRowScope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow
    '    Dim TotalAmount As Decimal = 0
    '    Dim TotalQty As Decimal = 0
    '    Dim TotalAmount_Allocated As Decimal = 0
    '    Dim Trans_Amount_Allocated As Decimal = 0
    '    Dim Trans_Amount_Allocated_D As Decimal = 0
    '    Dim Trans_Amount_Allocated_D2 As Decimal = 0
    '    Dim Trans_Amount_Allocated_D3 As Decimal = 0
    '    Dim Trans_Qty As Decimal = 0

    '    Dim DTABundling_Aft As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tblbundling_item_afterTableAdapter
    '    Dim TBLBundling_Aft As New DseTransaksiMonthlyClosing_Overhead.tblbundling_item_afterDataTable
    '    Dim rowBundling_Aft As DseTransaksiMonthlyClosing_Overhead.tblbundling_item_afterRow
    '    Dim DTA_Multi As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tblmaster_product_multi_unitTableAdapter
    '    Dim TBL_Multi As New DseTransaksiMonthlyClosing_Overhead.tblmaster_product_multi_unitDataTable
    '    Dim TBL_Multi_01 As New DseTransaksiMonthlyClosing_Overhead.tblmaster_product_multi_unitDataTable
    '    Dim row_Multi, row_Multi01 As DseTransaksiMonthlyClosing_Overhead.tblmaster_product_multi_unitRow
    '    Dim rowDGV As DataGridViewRow
    '    Dim DV As DataRowView
    '    Dim Filter1 As String = ""
    '    Dim DefPrec As Integer = 0
    '    Dim TBLCek As New DseTransaksiMonthlyClosing_Overhead.tblCek_MultiUnitDataTable
    '    Dim rowCek As DseTransaksiMonthlyClosing_Overhead.tblCek_MultiUnitRow
    '    Dim FrmCek As New FrmTransaksiMonthlyClosing2_CekMultiUnit
    '    Dim ListPrimaryBundling As New List(Of Double)
    '    Dim DTABundling As New DseTransaksiMonthlyClosingTableAdapters.DataTable12TableAdapter
    '    Dim TBLBundling As New DseTransaksiMonthlyClosing.DataTable12DataTable
    '    Dim TBLDummy As New DseTransaksiMonthlyClosing_Overhead.tblDummyDataTable
    '    Dim rowDummy As DseTransaksiMonthlyClosing_Overhead.tblDummyRow

    '    Try
    '        DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
    '        'kosongkan data overhead di data tabel
    '        Me.vTBLCostGroup_CodeNo.Rows.Clear()
    '        Me.vTBLCostGroup_CodeNo2.Rows.Clear()

    '        'kumpulkan semua data overhead
    '        DTABundling.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_Multi.Connection.ConnectionString = DTABundling.Connection.ConnectionString

    '        DTABundling_Aft.Fill(TBLBundling_Aft, Me.ObjCD.GetLasDateOfTheMonth(Me.ClosingPeriod), _
    '        Me.ObjCD.GetFirstDateofTheMonth(Me.ClosingPeriod))
    '        DTA_Multi.Fill(TBL_Multi)
    '        TBL_Multi_01.Rows.Clear()

    '        'cek apakah ada item product yang tidak memiliki konversi ke dalam unit yang dibutuhkan
    '        TBLCek.Rows.Clear()
    '        For Each rowDGV In Me.dgvItemCost_CostGroup.Rows
    '            If String.IsNullOrEmpty(rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value) = False Then
    '                If rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value <> -1 Then
    '                    Me.vTBLCostGroup_Scope.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(Me.ColCostGroup_Identifier.Index).Value.ToString + "'"
    '                    Filter1 = ""
    '                    For Each DV In Me.vTBLCostGroup_Scope.DefaultView
    '                        vRowScope = DV.Row
    '                        If Filter1 = "" Then
    '                            Filter1 = "primary process = " + vRowScope.primary_process.ToString
    '                        Else
    '                            Filter1 = Filter1 + " OR " + "primary process = " + vRowScope.primary_process.ToString
    '                        End If
    '                    Next

    '                    For Each DV In TBLBundling_Aft.DefaultView
    '                        rowBundling_Aft = DV.Row
    '                        With rowBundling_Aft
    '                            If .Primary_Units <> rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value Then
    '                                'pertama cari dahulu ke tabel multi-01, lebih sedikit isinya
    '                                TBL_Multi_01.DefaultView.RowFilter = "[Primary Product] = " + .Primary_Product.ToString + " AND [Primary Unit Multi] = " + rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value.ToString
    '                                If TBL_Multi_01.DefaultView.Count = 0 Then
    '                                    TBL_Multi.DefaultView.RowFilter = "[Primary Product] = " + .Primary_Product.ToString + " AND [Primary Unit Multi] = " + rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value.ToString
    '                                    If TBL_Multi.DefaultView.Count = 0 Then
    '                                        'masukkan ke dalam data table untuk keluarkan form error untuk menandakan product apa saja yang belum memiliki multi unit
    '                                        TBLCek.DefaultView.RowFilter = "[Cost Group] = '" + rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value.ToString + _
    '                                        "' AND [Product Code] = '" + rowBundling_Aft.Product_Code + "' AND [Primary Ref Unit] = " + rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value.ToString
    '                                        If TBLCek.DefaultView.Count = 0 Then
    '                                            'tambahkan baris baru
    '                                            rowCek = TBLCek.NewtblCek_MultiUnitRow
    '                                            rowCek.Cost_Group = rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value
    '                                            rowCek.Primary_Ref_Unit = rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value.ToString
    '                                            rowCek.Product_Code = rowBundling_Aft.Product_Code
    '                                            rowCek.Product_Desc = rowBundling_Aft.Description
    '                                            rowCek.Ref_Unit = rowDGV.Cells(Me.ColCostGroup_RefUnit.Index).Value.ToString
    '                                            TBLCek.Rows.Add(rowCek)
    '                                        End If
    '                                    Else
    '                                        row_Multi = TBL_Multi.DefaultView(0).Row
    '                                        'tambahkan baris dari tabel multi -> tabel multi 01, sehingga lain kali hitungannya lebih cepat
    '                                        row_Multi01 = TBL_Multi_01.Newtblmaster_product_multi_unitRow
    '                                        For ij As Integer = 0 To TBL_Multi_01.Columns.Count - 1
    '                                            row_Multi01.Item(ij) = row_Multi(ij)
    '                                        Next
    '                                        TBL_Multi_01.Rows.Add(row_Multi01)
    '                                    End If
    '                                End If
    '                            End If
    '                        End With
    '                    Next
    '                End If
    '            End If
    '        Next
    '        If TBLCek.Rows.Count > 0 Then
    '            FrmCek.TBLCek = TBLCek
    '            FrmCek.BackColorAll = Me.BackColorAll
    '            FrmCek.ShowDialog()
    '            Exit Sub
    '        End If

    '        'untuk setiap kelompok cost
    '        For Each rowDGV In Me.dgvItemCost_CostGroup.Rows
    '            TotalAmount = 0
    '            TotalQty = 0
    '            If String.IsNullOrEmpty(rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value) = False Then
    '                'a. kumpulkan allocated amount total
    '                TBLDummy.Rows.Clear()
    '                Me.vTBLCostGroup_Account.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(Me.ColCostGroup_Identifier.Index).Value.ToString + "'"
    '                For Each DV In Me.vTBLCostGroup_Account.DefaultView
    '                    vRowAccount = DV.Row
    '                    TotalAmount += vRowAccount.allocated_amount
    '                    'a.1 isikan tabel dummy, untuk faktor kali item code no2
    '                    rowDummy = TBLDummy.NewtblDummyRow
    '                    rowDummy.Primary_Akun = vRowAccount.primary_akun
    '                    rowDummy.Allocated_Amount = vRowAccount.allocated_amount
    '                    TBLDummy.Rows.Add(rowDummy)
    '                Next

    '                'b. isikan qty1 dan primary unit 1,jika dibutuhkan, konversi ke dalam satuan yang digunakan
    '                Me.vTBLCostGroup_Scope.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(Me.ColCostGroup_Identifier.Index).Value.ToString + "'"
    '                Filter1 = ""
    '                If Me.vTBLCostGroup_Scope.DefaultView.Count = 0 Then
    '                    Filter1 = "[primary process] = -2"
    '                Else
    '                    For Each DV In Me.vTBLCostGroup_Scope.DefaultView
    '                        vRowScope = DV.Row
    '                        If Filter1 = "" Then
    '                            Filter1 = "[primary process] = " + vRowScope.primary_process.ToString
    '                        Else
    '                            Filter1 = Filter1 + " OR [primary process] = " + vRowScope.primary_process.ToString
    '                        End If
    '                    Next
    '                End If
    '                TBLBundling_Aft.DefaultView.RowFilter = Filter1
    '                TBLBundling_Aft.DefaultView.RowFilter = Filter1
    '                TBLBundling_Aft.DefaultView.Sort = "[Tanggal Transaksi], [Bundling No], [Primary Product]"

    '                If rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value <> -1 Then
    '                    'cara baca tabel multi, jika unit default btg, dan unit multi-nya m3
    '                    'qty default . primary unit default = qty ratio . primary multi unit
    '                    '100 . btg = 5.435 m3
    '                    For Each DV In TBLBundling_Aft.DefaultView
    '                        rowBundling_Aft = DV.Row
    '                        With rowBundling_Aft
    '                            If .Primary_Units <> rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value Then
    '                                'pertama cari dahulu ke tabel multi-01, lebih sedikit isinya
    '                                TBL_Multi_01.DefaultView.RowFilter = "[Primary Product] = " + .Primary_Product.ToString + " AND [Primary Unit Multi] = " + rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value.ToString
    '                                If TBL_Multi_01.DefaultView.Count = 0 Then
    '                                    TBL_Multi.DefaultView.RowFilter = "[Primary Product] = " + .Primary_Product.ToString + " AND [Primary Unit Multi] = " + rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value.ToString
    '                                    If TBL_Multi.DefaultView.Count = 0 Then
    '                                        MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
    '                                        Exit Sub
    '                                    Else
    '                                        row_Multi = TBL_Multi.DefaultView(0).Row
    '                                        .Qty1 = Math.Round(.Qty * row_Multi.Qty_Rasio / row_Multi.Qty_Default, row_Multi.Def_Prec_Multi, MidpointRounding.AwayFromZero)
    '                                        'tambahkan baris dari tabel multi -> tabel multi 01, sehingga lain kali hitungannya lebih cepat
    '                                        row_Multi01 = TBL_Multi_01.Newtblmaster_product_multi_unitRow
    '                                        For ij As Integer = 0 To TBL_Multi_01.Columns.Count - 1
    '                                            row_Multi01.Item(ij) = row_Multi(ij)
    '                                        Next
    '                                        TBL_Multi_01.Rows.Add(row_Multi01)
    '                                    End If
    '                                Else
    '                                    row_Multi = TBL_Multi_01.DefaultView(0).Row
    '                                    .Qty1 = Math.Round(.Qty * row_Multi.Qty_Rasio / row_Multi.Qty_Default, row_Multi.Def_Prec_Multi, MidpointRounding.AwayFromZero)
    '                                End If
    '                                .Primary_Unit_1 = rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value
    '                                .Decimal_Qty = row_Multi.Def_Prec_Multi
    '                            Else
    '                                'qty1=qty, primary unit1= primary unit, decimal qty tdk berubah
    '                                rowBundling_Aft.Qty1 = rowBundling_Aft.Qty
    '                                rowBundling_Aft.Primary_Unit_1 = rowBundling_Aft.Primary_Units
    '                            End If
    '                        End With
    '                    Next
    '                Else
    '                    'set qty1=qty dan primary unit1= primary unit
    '                    'decimal qty tidak di ubah
    '                    For Each DV In TBLBundling_Aft.DefaultView
    '                        rowBundling_Aft = DV.Row
    '                        rowBundling_Aft.Qty1 = rowBundling_Aft.Qty
    '                        rowBundling_Aft.Primary_Unit_1 = rowBundling_Aft.Primary_Units
    '                    Next
    '                End If

    '                'c. kumpulkan semua item after bundling yang masuk ke dalam prosesnya, hitung total qty
    '                TotalQty = 0
    '                For Each DV In TBLBundling_Aft.DefaultView
    '                    rowBundling_Aft = DV.Row
    '                    TotalQty += rowBundling_Aft.Qty1
    '                Next
    '                rowDGV.Cells(Me.ColCostGroup_TotalQty.Index).Value = TotalQty

    '                'd. hitung per item-nya dan masukkan ke dalam data table
    '                ListPrimaryBundling.Clear()
    '                TotalAmount_Allocated = 0
    '                For ij As Double = 0 To TBLBundling_Aft.DefaultView.Count - 1
    '                    rowBundling_Aft = TBLBundling_Aft.DefaultView(ij).Row
    '                    If ListPrimaryBundling.Contains(rowBundling_Aft.Primary_Transaksi) = False Then
    '                        ListPrimaryBundling.Add(rowBundling_Aft.Primary_Transaksi)
    '                    End If
    '                Next

    '                For ij As Integer = 0 To ListPrimaryBundling.Count - 1
    '                    TBLBundling_Aft.DefaultView.RowFilter = "[Primary Transaksi] = " + ListPrimaryBundling.Item(ij).ToString
    '                    'hitung qty pada bundling ybs
    '                    Trans_Qty = 0
    '                    For Each DV In TBLBundling_Aft.DefaultView
    '                        rowBundling_Aft = DV.Row
    '                        Trans_Qty += rowBundling_Aft.Qty1
    '                    Next

    '                    'hitung amount untuk bundling ybs
    '                    If ij = ListPrimaryBundling.Count - 1 Then
    '                        Trans_Amount_Allocated = TotalAmount - TotalAmount_Allocated
    '                    Else
    '                        Trans_Amount_Allocated = Math.Round(TotalAmount * Trans_Qty / TotalQty, DefPrec, MidpointRounding.AwayFromZero)
    '                        TotalAmount_Allocated += Trans_Amount_Allocated
    '                    End If

    '                    'hitung overhead untuk masing2 item overhead
    '                    Trans_Amount_Allocated_D = 0
    '                    For jk As Double = 0 To TBLBundling_Aft.DefaultView.Count - 1
    '                        rowBundling_Aft = TBLBundling_Aft.DefaultView(jk).Row
    '                        vRow = Me.vTBLCostGroup_CodeNo.Newtbltrans_purchase_closing_item_cost_group_code_noRow

    '                        With vRow
    '                            If jk < TBLBundling_Aft.DefaultView.Count - 1 Then
    '                                .allocated_amount = Math.Round(Trans_Amount_Allocated * rowBundling_Aft.Faktor_Pemberat / 100, DefPrec, MidpointRounding.AwayFromZero)
    '                                Trans_Amount_Allocated_D += .allocated_amount
    '                            Else
    '                                .allocated_amount = Trans_Amount_Allocated - Trans_Amount_Allocated_D
    '                            End If

    '                            .identifier = rowDGV.Cells(Me.ColCostGroup_Identifier.Index).Value.ToString
    '                            .primary = Me.GetPrimaryMax_AllocatedCodeNo + 1
    '                            .primary_code_no = rowBundling_Aft.Primary_Product
    '                            .primary_transaksi = -1
    '                            .primary_unit = rowBundling_Aft.Primary_Unit_1
    '                            .primary_bundling = rowBundling_Aft.Primary_Transaksi
    '                            .qty_ref = rowBundling_Aft.Qty1
    '                            .decimal_qty = rowBundling_Aft.Decimal_Qty
    '                            .Bundling_No = rowBundling_Aft.Bundling_No
    '                            .primary_process = rowBundling_Aft.primary_process
    '                            .Tanggal_Transaksi = rowBundling_Aft.Tanggal_Transaksi
    '                            .remark = ""
    '                            .cost_factor = rowBundling_Aft.Faktor_Pemberat
    '                        End With

    '                        Me.vTBLCostGroup_CodeNo.Rows.Add(vRow)

    '                        'masukkan item untuk item code no2
    '                        Trans_Amount_Allocated_D2 = 0
    '                        Trans_Amount_Allocated_D3 = 0
    '                        TBLDummy.DefaultView.RowFilter = ""
    '                        TBLDummy.DefaultView.Sort = "[Allocated Amount] DESC"
    '                        For wk As Integer = 0 To TBLDummy.DefaultView.Count
    '                            rowDummy = TBLDummy.DefaultView(wk).Row
    '                            If wk = TBLDummy.DefaultView.Count - 1 Then
    '                                Trans_Amount_Allocated_D2 = vRow.allocated_amount - Trans_Amount_Allocated_D3
    '                            Else
    '                                Trans_Amount_Allocated_D2 = Math.Round(vRow.allocated_amount * rowDummy.Allocated_Amount / TotalAmount, _
    '                                DefPrec, MidpointRounding.AwayFromZero)
    '                                Trans_Amount_Allocated_D3 += Trans_Amount_Allocated_D2
    '                            End If

    '                            Me.vTBLCostGroup_CodeNo2.DefaultView.RowFilter = "[primary account] = " + rowDummy.Primary_Akun.ToString + " AND [primary process] = " + vRow.primary_process.ToString
    '                            If Me.vTBLCostGroup_CodeNo2.DefaultView.Count = 0 Then
    '                                vRow2 = Me.vTBLCostGroup_CodeNo2.Newtbltrans_purchase_closing_item_cost_group_code_no2Row
    '                                With vRow2
    '                                    .allocated_amount = Trans_Amount_Allocated_D2
    '                                    .primary = Me.vTBLCostGroup_CodeNo2.Rows.Count + 1
    '                                    .primary_account = rowDummy.Primary_Akun
    '                                    .primary_process = vRow.primary_process
    '                                    .primary_transaksi = -1
    '                                End With
    '                                Me.vTBLCostGroup_CodeNo2.Rows.Add(vRow2)
    '                            Else
    '                                vRow2 = Me.vTBLCostGroup_CodeNo2.DefaultView(0).Row
    '                                vRow2.allocated_amount += Trans_Amount_Allocated_D2
    '                            End If
    '                        Next
    '                    Next
    '                Next
    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTABundling.Dispose()
    '        TBLBundling.Dispose()
    '        DTA_Multi.Dispose()
    '        TBL_Multi.Dispose()
    '        TBLCek.Dispose()
    '        FrmCek.Dispose()
    '        ListPrimaryBundling = Nothing
    '        TBLDummy.Dispose()
    '    End Try
    'End Sub

    Private Sub Update_Cost_Group_Overhead_Amount(Optional ByVal vIdentifier As String = "")
        Dim DefPrec As Integer = 0
        Dim Lanjut As Boolean = False
        Dim vRow As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim Amount As Decimal = 0

        Try
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            For Each rowDGV As DataGridViewRow In Me.dgvItemCost_CostGroup.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value) = False Then
                    If vIdentifier = "" Then
                        'langsung kerjain
                        Lanjut = True
                    Else
                        'cek apakah identifier sama
                        If rowDGV.Cells(Me.ColCostGroup_Identifier.Index).Value = vIdentifier Then
                            Lanjut = True
                        Else
                            Lanjut = False
                        End If
                    End If

                    If Lanjut = True Then
                        Amount = 0
                        Me.vTBLCostGroup_Account.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(Me.ColCostGroup_Identifier.Index).Value.ToString + "'"
                        For Each DV As DataRowView In Me.vTBLCostGroup_Account.DefaultView
                            vRow = DV.Row
                            Amount += vRow.allocated_amount
                        Next

                        rowDGV.Cells(Me.ColCostGroup_OverheadAmount.Index).Value = Me.ObjCD.DigitGrouping(Amount, DefPrec)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail_Overhead()
        Try
            'aul
            Me.dgvItemCost_CostGroup.ReadOnly = False
            Me.dgvItemCost_CostAccountAllocation.ReadOnly = False
            Me.dgvItemCost_Scope.ReadOnly = False
            Me.dgvItemCost_AllocationRecap1.ReadOnly = False
            Me.dgvItemCost_CostGroup.AllowUserToAddRows = True
            Me.dgvItemCost_CostAccountAllocation.AllowUserToAddRows = True
            Me.dgvItemCost_Scope.AllowUserToAddRows = True

            Me.btnUP.Enabled = False
            Me.BtnUpAccount.Enabled = False
            Me.BtnUpScope.Enabled = False
            Me.btnDown.Enabled = False
            Me.BtnDownAccount.Enabled = False
            Me.BtnDownScope.Enabled = False

            Me.llbRefreshBalance.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField_Overhead()
        Try
            Me.dgvItemCost_CostGroup.Rows.Clear()
            Me.dgvItemCost_CostAccountAllocation.Rows.Clear()
            Me.dgvItemCost_Scope.Rows.Clear()
            Me.dgvItemCost_AllocationRecap1.Rows.Clear()

            Me.vTBLCostGroup_Account.Rows.Clear()
            Me.vTBLCostGroup_CodeNo.Rows.Clear()
            Me.vTBLCostGroup_Scope.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail_Overhead()
        Try
            Me.dgvItemCost_CostGroup.ReadOnly = True
            Me.dgvItemCost_CostAccountAllocation.ReadOnly = True
            Me.dgvItemCost_Scope.ReadOnly = True
            Me.dgvItemCost_AllocationRecap1.ReadOnly = True

            Me.dgvItemCost_CostGroup.AllowUserToAddRows = False
            Me.dgvItemCost_CostAccountAllocation.AllowUserToAddRows = False
            Me.dgvItemCost_Scope.AllowUserToAddRows = False

            Me.btnUP.Enabled = False
            Me.BtnUpAccount.Enabled = False
            Me.BtnUpScope.Enabled = False
            Me.btnDown.Enabled = False
            Me.BtnDownAccount.Enabled = False
            Me.BtnDownScope.Enabled = False

            Me.llbRefreshBalance.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinDataOverheadBulanLalu()
        Dim PrimaryClosingB4 As Double = -1
        Dim rowDGV As DataGridViewRow
        Dim Len As Integer = 0
        Dim vRowAcc As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim vRowScope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow

        Try
            PrimaryClosingB4 = Me.ObjClosing.GetPrimary(Me.ObjCD.GetFirstDateofTheMonth(Me.ClosingPeriod - TimeSpan.FromDays(1)))
            If PrimaryClosingB4 <> -1 Then
                Me.ObjCLosingItem.IsiDGV_Overhead(PrimaryClosingB4, Me.dgvItemCost_CostGroup, _
                Me.vTBLCostGroup_Account, Me.vTBLCostGroup_Scope, Me.vTBLCostGroup_CodeNo, ColCostGroup_AllocationStyle.Index, ColCostGroup_GroupDesc.Index, _
                ColCostGroup_Identifier.Index, ColCostGroup_OverheadAmount.Index, ColCostGroup_PrimaryRefUnit.Index, _
                ColCostGroup_Remarks.Index, ColCostGroup_TotalQty.Index, ColCostGroup_RefUnit.Index, _
                ObjCD, Me.ObjCurrency.GetDecimal_DefaultCurrency, _
                Core_Module.GetConstring, Me.vTBLCostGroup_CodeNo2, False)

                'tandain ulang identifier + Primary Transaksi
                Len = PrimaryClosingB4.ToString.Length
                For Each rowDGV In Me.dgvItemCost_CostGroup.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value) = False Then
                        rowDGV.Cells(Me.ColCostGroup_Identifier.Index).Value = "-1" + rowDGV.Cells(Me.ColCostGroup_Identifier.Index).Value.ToString.Substring(Len)
                    End If
                Next
                For Each vRowAcc In vTBLCostGroup_Account.Rows
                    With vRowAcc
                        .identifier = "-1" + .identifier.Substring(Len)
                        .primary_transaksi = -1
                    End With
                Next
                For Each vRowScope In vTBLCostGroup_Scope.Rows
                    With vRowScope
                        .identifier = "-1" + .identifier.Substring(Len)
                        .primary_transaksi = -1
                    End With
                Next
                Me.vTBLCostGroup_CodeNo.Rows.Clear()

                'tarik kembali nilai overhead yang di allokasikan untuk bulan ybs
                Me.RefreshBalance_Overhead()
                'hitung ulang nilai overhead untuk setiap transaksi
                'Me.RecalculateOverhead() -> ga perlu hitung ulang karena disini tidak ada tab-nya
                Me.TampilkanOverheadDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiMonthlyClosing2_Overhead2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.CekBeforeCLose = True Then
                If MsgBox("Cancel data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiMonthlyClosing2_Overhead2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dgvItemCost_AllocationRecap1.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemCost_AllocationRecap2.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemCost_CostAccountAllocation.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemCost_CostGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemCost_Scope.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.SalinDataOverheadBulanLalu()
            'Me.ColCostGroup_Identifier.Visible = True
            'Me.ColAllScope_Identifier.Visible = True
            'Me.ColAlloRecap2_Identifier.Visible = True
            'Me.ColAlloRecap1_Identifier.Visible = True
            'Me.ColAccAll_Identifier.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            'cek kebenaran data
            If Me.ApakahDataBenar_Overhead = True Then
                'konfirmasi ulang
                If MsgBox("Accept overhead cost data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.RefreshBalance_Overhead()
                    Me.TampilkanOverheadDetail()

                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.CekBeforeCLose = False
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar_Overhead() As Boolean
        Dim rowDGV As DataGridViewRow
        Dim Dummy As String = ""
        Dim counter As Double = 0
        Dim ListPrimaryAkun As New List(Of Double)
        Dim vRowAkun As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim EndBalance, AllocatedAmount As Decimal
        Dim DTAMasterProses As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tblmaster_proses1TableAdapter
        Dim TBLMasterProses As New DseTransaksiMonthlyClosing_Overhead.tblmaster_proses1DataTable
        Dim rowMasterProses As DseTransaksiMonthlyClosing_Overhead.tblmaster_proses1Row
        Dim vRowCostGroup_Scope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow
        Dim vRowCostGroup_Account As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim TBLDummy As New DseTransaksiMonthlyClosing_Overhead.tblmaster_akunDataTable
        Dim rowDummy As DseTransaksiMonthlyClosing_Overhead.tblmaster_akunRow
        Dim TBLCek As New DseTransaksiMonthlyClosing_Overhead.tblCek_MultiUnitDataTable
        Dim rowCek As DseTransaksiMonthlyClosing_Overhead.tblCek_MultiUnitRow
        Dim FrmCek As New FrmTransaksiMonthlyClosing2_CekMultiUnit
        Dim Filter1 As String = ""
        Dim DTABundling_Aft As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tblbundling_item_afterTableAdapter
        Dim TBLBundling_Aft As New DseTransaksiMonthlyClosing_Overhead.tblbundling_item_afterDataTable
        Dim rowBundling_Aft As DseTransaksiMonthlyClosing_Overhead.tblbundling_item_afterRow
        Dim DTA_Multi As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tblmaster_product_multi_unitTableAdapter
        Dim TBL_Multi As New DseTransaksiMonthlyClosing_Overhead.tblmaster_product_multi_unitDataTable
        Dim TBL_Multi_01 As New DseTransaksiMonthlyClosing_Overhead.tblmaster_product_multi_unitDataTable
        Dim row_Multi, row_Multi01 As DseTransaksiMonthlyClosing_Overhead.tblmaster_product_multi_unitRow
        Dim DV As DataRowView
        Dim vRowScope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTAMasterProses.Connection.ConnectionString = Core_Module.GetConstring
            DTABundling_Aft.Connection.ConnectionString = DTAMasterProses.Connection.ConnectionString
            DTA_Multi.Connection.ConnectionString = DTAMasterProses.Connection.ConnectionString

            Dummy = ""
            For Each rowDGV In Me.dgvItemCost_CostGroup.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value) = False Then
                    'setiap cost group minimal memiliki 1 proses dan 1 akun
                    Me.vTBLCostGroup_Account.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(Me.ColCostGroup_Identifier.Index).Value + "'"
                    Me.vTBLCostGroup_Scope.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(Me.ColCostGroup_Identifier.Index).Value + "'"
                    If Me.vTBLCostGroup_Scope.DefaultView.Count = 0 Or Me.vTBLCostGroup_Account.DefaultView.Count = 0 Then
                        If Dummy = "" Then
                            Dummy = rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value.ToString
                        Else
                            Dummy = Dummy + ", " + rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value.ToString
                        End If
                    End If
                End If
            Next
            If Dummy <> "" Then
                MsgBox("The following cost group is invalid: " + Dummy + _
                ". Please provide at least 1 account and 1 process for each cost group.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'tidak boleh ada item allocated account yang allocated perct dan allocated amount-nya 0
            Dummy = ""
            For Each row As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow In Me.vTBLCostGroup_Account.Rows
                If row.balance <> 0 Then
                    If row.allocated_amount = 0 And row.perc_allocated = 0 Then
                        counter += 1
                        If Dummy = "" Then
                            Dummy = counter.ToString + ". Cost group: " + Me.GetCostGroup(row.identifier) + " - Account Code: " + Me.ObjAkun.GetAccountCode(row.primary_akun)
                        Else
                            Dummy = Dummy + vbCrLf + counter.ToString + ". Cost group: " + Me.GetCostGroup(row.identifier) + " - Account Code: " + Me.ObjAkun.GetAccountCode(row.primary_akun)
                        End If
                    End If
                End If
            Next
            If Dummy <> "" Then
                MsgBox("Please specify allocated amount form the following account: " + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
                Return False
            End If

            'tidak boleh ada account yang allocated amount > balance
            Dummy = ""
            counter = 0
            For Each vRowAkun In Me.vTBLCostGroup_Account.Rows
                If ListPrimaryAkun.Contains(vRowAkun.primary_akun) = False Then
                    ListPrimaryAkun.Add(vRowAkun.primary_akun)
                End If
            Next
            For Each PrimaryAkun As Double In ListPrimaryAkun
                Me.vTBLCostGroup_Account.DefaultView.RowFilter = "[primary akun] = " + PrimaryAkun.ToString
                'ambil data end balance
                vRowAkun = Me.vTBLCostGroup_Account.DefaultView(0).Row
                EndBalance = vRowAkun.balance
                AllocatedAmount = 0

                'totalkan allocated amount
                For Each DV In Me.vTBLCostGroup_Account.DefaultView
                    vRowAkun = DV.Row
                    AllocatedAmount += vRowAkun.allocated_amount
                Next

                'tidak allocated amount > balance, maka masukkan ke dummy
                If EndBalance > 0 Then
                    If AllocatedAmount > EndBalance Then
                        counter += 1
                        If Dummy = "" Then
                            Dummy = counter.ToString + ". " + Me.ObjAkun.GetAccountCode(PrimaryAkun)
                        Else
                            Dummy = Dummy + vbCrLf + counter.ToString + ". " + Me.ObjAkun.GetAccountCode(PrimaryAkun)
                        End If
                    End If
                Else
                    If AllocatedAmount < EndBalance Then
                        counter += 1
                        If Dummy = "" Then
                            Dummy = counter.ToString + ". " + Me.ObjAkun.GetAccountCode(PrimaryAkun)
                        Else
                            Dummy = Dummy + vbCrLf + counter.ToString + ". " + Me.ObjAkun.GetAccountCode(PrimaryAkun)
                        End If
                    End If
                End If

            Next
            If Dummy <> "" Then
                MsgBox("The allocated amount > end balance for the following account, please fix:" + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
                Return False
            End If

            'jika ada kode proses pada periode tersebut yang terjadi, tetapi tidak dapat overhead
            Dummy = ""
            counter = 0
            DTAMasterProses.Fill(TBLMasterProses, Me.ObjCD.GetLasDateOfTheMonth(Me.ClosingPeriod), _
            Me.ObjCD.GetFirstDateofTheMonth(Me.ClosingPeriod))
            For Each vRowCostGroup_Scope In Me.vTBLCostGroup_Scope.Rows
                TBLMasterProses.DefaultView.RowFilter = "[primary process] = " + vRowCostGroup_Scope.primary_process.ToString

                While TBLMasterProses.DefaultView.Count > 0
                    TBLMasterProses.Rows.Remove(TBLMasterProses.DefaultView(0).Row)
                End While
            Next
            If TBLMasterProses.Rows.Count > 0 Then
                For Each rowMasterProses In TBLMasterProses.Rows
                    counter += 1
                    If Dummy = "" Then
                        Dummy = counter.ToString + ". " + rowMasterProses.Code + " - " + rowMasterProses.Desc
                    Else
                        Dummy = Dummy + vbCrLf + counter.ToString + ". " + rowMasterProses.Code + " - " + rowMasterProses.Desc
                    End If
                Next

                If MsgBox("The following process doesn't have any overhead, continue process anyway?" + vbCrLf + Dummy, MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Return False
                End If
            End If

            'jika ada kode akun yang allocated perct-nya tidak 100%
            TBLDummy.Rows.Clear()
            Dummy = ""
            counter = 0
            For Each vRowCostGroup_Account In Me.vTBLCostGroup_Account.Rows
                TBLDummy.DefaultView.RowFilter = "[primary akun] = " + vRowCostGroup_Account.primary_akun.ToString
                If TBLDummy.DefaultView.Count = 0 Then
                    rowDummy = TBLDummy.Newtblmaster_akunRow
                    With rowDummy
                        .primary_akun = vRowCostGroup_Account.primary_akun
                        .code = Me.ObjAkun.GetAccountCode(.primary_akun)
                        .desc = Me.ObjAkun.GetDescription(.code)
                        .perct = vRowCostGroup_Account.perc_allocated
                    End With
                    TBLDummy.Rows.Add(rowDummy)
                Else
                    rowDummy = TBLDummy.DefaultView(0).Row
                    rowDummy.perct += vRowCostGroup_Account.perc_allocated
                End If
            Next
            For Each rowDummy In TBLDummy.Rows
                If rowDummy.perct <> 100 Then
                    counter += 1
                    If Dummy = "" Then
                        Dummy = counter.ToString + ". " + rowDummy.code + " - " + rowDummy.desc
                    Else
                        Dummy = Dummy + vbCrLf + counter.ToString + ". " + rowDummy.code + " - " + rowDummy.desc
                    End If
                End If
            Next
            If Dummy <> "" Then
                If MsgBox("The following account allocated percentage is not 100%." + vbCrLf + Dummy + vbCrLf + vbCrLf + _
                "Continue process anyway?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Return False
                End If
            End If

            'cek apakah ada item product yang tidak memiliki konversi ke dalam unit yang dibutuhkan
            TBLCek.Rows.Clear()
            TBL_Multi_01.Rows.Clear()
            TBLBundling_Aft.Rows.Clear()
            TBLCek.Rows.Clear()

            DTABundling_Aft.Fill(TBLBundling_Aft, Me.ObjCD.GetLasDateOfTheMonth(Me.ClosingPeriod), _
            Me.ObjCD.GetFirstDateofTheMonth(Me.ClosingPeriod))
            DTA_Multi.Fill(TBL_Multi)

            For Each rowDGV In Me.dgvItemCost_CostGroup.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value) = False Then
                    If rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value <> -1 Then
                        Me.vTBLCostGroup_Scope.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(Me.ColCostGroup_Identifier.Index).Value.ToString + "'"
                        Filter1 = ""
                        For Each DV In Me.vTBLCostGroup_Scope.DefaultView
                            vRowScope = DV.Row
                            If Filter1 = "" Then
                                Filter1 = "[primary process] = " + vRowScope.primary_process.ToString
                            Else
                                Filter1 = Filter1 + " OR " + "[primary process] = " + vRowScope.primary_process.ToString
                            End If
                        Next
                        TBLBundling_Aft.DefaultView.RowFilter = Filter1

                        For Each DV In TBLBundling_Aft.DefaultView
                            rowBundling_Aft = DV.Row
                            With rowBundling_Aft
                                If .Primary_Units <> rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value Then
                                    'pertama cari dahulu ke tabel multi-01, lebih sedikit isinya
                                    TBL_Multi_01.DefaultView.RowFilter = "[Primary Product] = " + .Primary_Product.ToString + " AND [Primary Unit Multi] = " + rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value.ToString
                                    If TBL_Multi_01.DefaultView.Count = 0 Then
                                        TBL_Multi.DefaultView.RowFilter = "[Primary Product] = " + .Primary_Product.ToString + " AND [Primary Unit Multi] = " + rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value.ToString
                                        If TBL_Multi.DefaultView.Count = 0 Then
                                            'masukkan ke dalam data table untuk keluarkan form error untuk menandakan product apa saja yang belum memiliki multi unit
                                            TBLCek.DefaultView.RowFilter = "[Cost Group] = '" + rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value.ToString + _
                                            "' AND [Product Code] = '" + rowBundling_Aft.Product_Code + "' AND [Primary Ref Unit] = " + rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value.ToString
                                            If TBLCek.DefaultView.Count = 0 Then
                                                'tambahkan baris baru
                                                rowCek = TBLCek.NewtblCek_MultiUnitRow
                                                rowCek.Cost_Group = rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value
                                                rowCek.Primary_Ref_Unit = rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value.ToString
                                                rowCek.Product_Code = rowBundling_Aft.Product_Code
                                                rowCek.Product_Desc = rowBundling_Aft.Description
                                                rowCek.Ref_Unit = rowDGV.Cells(Me.ColCostGroup_RefUnit.Index).Value.ToString
                                                TBLCek.Rows.Add(rowCek)

                                                'tambahkan baris dari tabel multi -> tabel multi 01, sehingga lain kali hitungannya lebih cepat
                                                row_Multi01 = TBL_Multi_01.Newtblmaster_product_multi_unitRow
                                                With row_Multi01
                                                    .Def_Prec_Default = 0
                                                    .Def_Prec_Multi = 0
                                                    .Primary_Product = rowBundling_Aft.Primary_Product
                                                    .Primary_Unit_Default = 1
                                                    .Primary_Unit_Multi = rowDGV.Cells(Me.ColCostGroup_PrimaryRefUnit.Index).Value
                                                    .Qty_Default = 1
                                                    .Qty_Rasio = 1
                                                End With
                                                TBL_Multi_01.Rows.Add(row_Multi01)
                                            End If
                                        Else
                                            row_Multi = TBL_Multi.DefaultView(0).Row
                                            'tambahkan baris dari tabel multi -> tabel multi 01, sehingga lain kali hitungannya lebih cepat
                                            row_Multi01 = TBL_Multi_01.Newtblmaster_product_multi_unitRow
                                            For ij As Integer = 0 To TBL_Multi_01.Columns.Count - 1
                                                row_Multi01.Item(ij) = row_Multi(ij)
                                            Next
                                            TBL_Multi_01.Rows.Add(row_Multi01)
                                        End If
                                    End If
                                End If
                            End With
                        Next
                    End If
                End If
            Next

            If TBLCek.Rows.Count > 0 Then
                FrmCek.TBLCek = TBLCek
                FrmCek.RefPeriod = Me.ClosingPeriod
                FrmCek.usercode = Me.UserCode
                FrmCek.BackColorAll = Me.BackColorAll

                FrmCek.ShowDialog()
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryAkun = Nothing
            DTAMasterProses.Dispose()
            TBLMasterProses.Dispose()
            TBLDummy.Dispose()
            TBLCek.Dispose()
            FrmCek.Dispose()
            DTABundling_Aft.Dispose()
            TBLBundling_Aft.Dispose()
            DTA_Multi.Dispose()
            TBL_Multi.Dispose()
            TBL_Multi_01.Dispose()
        End Try
    End Function
End Class