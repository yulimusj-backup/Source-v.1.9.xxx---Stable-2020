Option Explicit On
Option Strict Off

Public Class FrmLaporanProd_RMAllocation
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim objCd As New Core_Module.ClsCekData
    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim objProduct As New Core_Module.clsMasterProduct
    Dim objProcess As New Core_Module.ClsMasterProses
    Dim objWS As New Core_Module.ClsMasterWorkShift
    Dim objWorkUnit As New Core_Module.ClsMasterWorkUnit
    Dim objWorkUnitGroup As New Core_Module.ClsMasterWorkUnitGroup
    Dim tblWU As New DseLaporanProd_PlanningReport.tblmaster_workunitDataTable
    Dim tblWU2 As New DseLaporanProd_PlanningReport.tblmaster_workunitDataTable
    Dim rowWU As DseLaporanProd_PlanningReport.tblmaster_workunitRow
    Dim rowWU2 As DseLaporanProd_PlanningReport.tblmaster_workunitRow

    Dim buffer As Object
    Dim mIndex As Integer = -1
    Dim mIndex2 As Integer = -1

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmLaporanProd_PlanningReport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim dtaWU As New DseLaporanProd_PlanningReportTableAdapters.tblmaster_workunitTableAdapter
        Try
            dtaWU.Connection.ConnectionString = Core_Module.GetConstring
            dtaWU.Fill(Me.tblWU)
            Me.dgvProduct.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvProcess.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvWorkUnit.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvWUG.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvProduct.CellBeginEdit
        Try
            If Me.colProduct_GroupLevel.Index Or Me.colProduct_Code.Index Then
                buffer = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvProduct.CellEndEdit

        Try
            If IsNothing(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.colProduct_GroupLevel.Index Then
                        cellEndEditGroupLevel(sender, e)
                    End If
                    If e.ColumnIndex = Me.colProduct_Code.Index Then
                        cellEndEditCode(sender, e)
                    End If
                Else
                    If IsNothing(Me.buffer) = False Then
                        If String.IsNullOrEmpty(buffer) = False Then
                            Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                    Me.mIndex = e.RowIndex
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            Else
                If IsNothing(Me.buffer) = False Then
                    If String.IsNullOrEmpty(buffer) = False Then
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                Me.mIndex = e.RowIndex
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cellEndEditGroupLevel(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindGroup As New Core_Module.FrmCariMasterProductGroup
        Dim frmFindCategory As New Core_Module.frmCariMasterCategory
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
            Select Case Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value
                Case "Group"
                    frmFindGroup.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 3
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                Case "Brand"
                    frmFindGroup.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 2
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                Case "Sub Brand"
                    frmFindGroup.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 1
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                Case "Category"
                    frmFindCategory.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindCategory.ShowDialog()
                    If frmFindCategory.PrimaryCategory > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindCategory.PrimaryCategory
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objCategory.GetCategoryCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetBrandDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                Case "Product"
                    frmFindProduct.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindProduct.ShowDialog()
                    If frmFindProduct.PrimaryProduct > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindProduct.PrimaryProduct
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objProduct.GetProductCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
            End Select
            If Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndexProduct(e.RowIndex, Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value, Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value & " " & Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.mIndex = e.RowIndex
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cellEndEditCode(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindGroup As New Core_Module.FrmCariMasterProductGroup
        Dim frmFindCategory As New Core_Module.frmCariMasterCategory
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
            Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
            Select Case Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value
                Case "Group"
                    If Me.objGroup.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 3) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 3)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 3
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.buffer) = False Then
                                If String.IsNullOrEmpty(buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                            Me.mIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
                Case "Brand"
                    If Me.objGroup.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 2) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 2
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.buffer) = False Then
                                If String.IsNullOrEmpty(buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                            Me.mIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
                Case "Sub Brand"
                    If Me.objGroup.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 1) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 1)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 1
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.buffer) = False Then
                                If String.IsNullOrEmpty(buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                            Me.mIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
                Case "Category"
                    If Me.objCategory.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objCategory.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetBrandDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindCategory.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindCategory.ShowDialog()
                        If frmFindCategory.PrimaryCategory > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindCategory.PrimaryCategory
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objCategory.GetCategoryCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetBrandDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.buffer) = False Then
                                If String.IsNullOrEmpty(buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                            Me.mIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
                Case "Product"
                    If Me.objProduct.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objProduct.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindProduct.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindProduct.ShowDialog()
                        If frmFindProduct.PrimaryProduct > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindProduct.PrimaryProduct
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objProduct.GetProductCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.buffer) = False Then
                                If String.IsNullOrEmpty(buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                            Me.mIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
            End Select
            If Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndexProduct(e.RowIndex, Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value, Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value & " " & Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.mIndex = e.RowIndex
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getIndexProduct(ByVal vRowException As Integer, ByVal vGroupLevel As String, ByVal vPrimary As Double) As Integer
        Try
            For Each dr As DataGridViewRow In Me.dgvProduct.Rows
                With dr
                    If Me.dgvProduct.Rows.IndexOf(dr) <> vRowException Then
                        If dr.Cells(Me.colProduct_Primary.Index).Value <> -1 Then
                            If dr.Cells(Me.colProduct_GroupLevel.Index).Value = vGroupLevel And dr.Cells(Me.colProduct_Primary.Index).Value = vPrimary Then
                                Return Me.dgvProduct.Rows.IndexOf(dr)
                            End If
                        End If
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvProduct_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvProduct.DataError

    End Sub

    Private Sub dgvProduct_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvProduct.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvProduct.Rows(Me.dgvProduct.Rows.Count - 1)
            dr.Cells(Me.colProduct_No.Index).Value = Me.dgvProduct.Rows.Count
            dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Group"
            dr.Cells(Me.colProduct_Desc.Index).Value = ""
            dr.Cells(Me.colProduct_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvProduct.RowsRemoved
        Try
            Me.objCd.SetNumber(Me.dgvProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvProduct.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.dgvProduct.Rows.Count > 1 Then
                Me.dgvProduct.Rows.RemoveAt(Me.mIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProcess_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvProcess.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colProcess_Code.Index Then
                Me.buffer = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProcess_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvProcess.CellEndEdit
        Dim frmFindProcess As New Core_Module.FrmMasterProsesCari
        Try
            If IsNothing(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.colProcess_Code.Index Then
                        If Me.objProcess.ApakahSudahAda(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value = Me.objProcess.GetPrimary(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvProcess.Item(Me.colProcess_Desc.Index, e.RowIndex).Value = Me.objProcess.GetDesc(Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value)
                            If getIndexProcess(e.RowIndex, Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value) <> -1 Then
                                MsgBox("Process " & Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                If IsNothing(Me.buffer) = False Then
                                    If String.IsNullOrEmpty(buffer) = False Then
                                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value = Me.objProcess.GetPrimary(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value)
                                        Me.dgvProcess.Item(Me.colProcess_Desc.Index, e.RowIndex).Value = Me.objProcess.GetDesc(Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value)
                                        Exit Sub
                                    End If
                                End If
                                Me.mIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            End If
                        Else
                            frmFindProcess.txtCari.Text = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value
                            frmFindProcess.ShowDialog()
                            If frmFindProcess.mPrimary > 0 Then
                                Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value = frmFindProcess.mPrimary
                                Me.dgvProcess.Item(Me.colProcess_Code.Index, e.RowIndex).Value = Me.objProcess.GetCode(Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value)
                                Me.dgvProcess.Item(Me.colProcess_Desc.Index, e.RowIndex).Value = Me.objProcess.GetDesc(Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value)
                                If getIndexProcess(e.RowIndex, Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value) <> -1 Then
                                    MsgBox("Process " & Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                    If IsNothing(Me.buffer) = False Then
                                        If String.IsNullOrEmpty(buffer) = False Then
                                            Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                            Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value = Me.objProcess.GetPrimary(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value)
                                            Me.dgvProcess.Item(Me.colProcess_Desc.Index, e.RowIndex).Value = Me.objProcess.GetDesc(Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value)
                                            Exit Sub
                                        End If
                                    End If
                                    Me.mIndex = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                End If
                            Else
                                If IsNothing(Me.buffer) = False Then
                                    If String.IsNullOrEmpty(buffer) = False Then
                                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                                Me.mIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            End If
                        End If
                    End If
                Else
                    If IsNothing(Me.buffer) = False Then
                        If String.IsNullOrEmpty(buffer) = False Then
                            Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                    Me.mIndex = e.RowIndex
                    Me.BackgroundWorker2.RunWorkerAsync()
                End If
            Else
                If IsNothing(Me.buffer) = False Then
                    If String.IsNullOrEmpty(buffer) = False Then
                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                Me.mIndex = e.RowIndex
                Me.BackgroundWorker2.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getIndexProcess(ByVal vRowException As Integer, ByVal vPrimary As Double) As Integer
        Try
            For Each dr As DataGridViewRow In Me.dgvProcess.Rows
                With dr
                    If Me.dgvProcess.Rows.IndexOf(dr) <> vRowException Then
                        If .Cells(Me.colProcess_Primary.Index).Value <> -1 Then
                            If .Cells(Me.colProcess_Primary.Index).Value = vPrimary Then
                                Return Me.dgvProcess.Rows.IndexOf(dr)
                            End If
                        End If
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvProcess_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvProcess.DataError

    End Sub

    Private Sub dgvProcess_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvProcess.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvProcess.Rows(Me.dgvProcess.Rows.Count - 1)
            dr.Cells(Me.colProcess_No.Index).Value = Me.dgvProcess.Rows.Count
            dr.Cells(Me.colProcess_Code.Index).Value = ""
            dr.Cells(Me.colProcess_Desc.Index).Value = ""
            dr.Cells(Me.colProcess_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProcess_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvProcess.RowsRemoved
        Try
            Me.objCd.SetNumber(Me.dgvProcess)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProcess_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvProcess.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvProcess)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If Me.dgvProcess.Rows.Count > 1 Then
                Me.dgvProcess.Rows.RemoveAt(Me.mIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWorkUnit_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvWorkUnit.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colWU_Code.Index Then
                Me.buffer = Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWorkUnit_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWorkUnit.CellEndEdit
        Dim frm As New Core_Module.FrmMasterWorkUnit_Cari
        Try
            If IsNothing(Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.buffer = Me.dgvWorkUnit(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.colWU_Code.Index Then
                        If Me.objWorkUnit.ApakahSudahAdaOnlyActive(Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value = Me.objWorkUnit.GetPrimary(Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvWorkUnit.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.objWorkUnit.GetDesc(Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value)
                            If getIndexWorkUnit(e.RowIndex, Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value) <> -1 Then
                                MsgBox("Work Unit " & Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                If IsNothing(Me.buffer) = False Then
                                    If String.IsNullOrEmpty(buffer) = False Then
                                        Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value = Me.objWorkUnit.GetPrimary(Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value)
                                        Me.dgvWorkUnit.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.objWorkUnit.GetDesc(Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value)
                                        Exit Sub
                                    End If
                                End If
                                Me.mIndex = e.RowIndex
                                Me.BackgroundWorker4.RunWorkerAsync()
                            End If
                        Else
                            frm.txtCari.Text = Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value
                            frm.ShowDialog()
                            If frm.PrimaryWorkUnit > 0 Then
                                Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value = frm.PrimaryWorkUnit
                                Me.dgvWorkUnit.Item(Me.colWU_Code.Index, e.RowIndex).Value = Me.objWorkUnit.GetCode(Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value)
                                Me.dgvWorkUnit.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.objWorkUnit.GetDesc(Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value)
                                If getIndexWorkUnit(e.RowIndex, Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value) <> -1 Then
                                    MsgBox("Work Unit " & Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                    If IsNothing(Me.buffer) = False Then
                                        If String.IsNullOrEmpty(buffer) = False Then
                                            Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                            Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value = Me.objWorkUnit.GetPrimary(Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value)
                                            Me.dgvWorkUnit.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.objWorkUnit.GetDesc(Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value)
                                            Exit Sub
                                        End If
                                    End If
                                    Me.mIndex = e.RowIndex
                                    Me.BackgroundWorker4.RunWorkerAsync()
                                End If
                            Else
                                If IsNothing(Me.buffer) = False Then
                                    If String.IsNullOrEmpty(buffer) = False Then
                                        Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                                Me.mIndex = e.RowIndex
                                Me.BackgroundWorker4.RunWorkerAsync()
                            End If
                        End If
                    End If
                Else
                    If IsNothing(Me.buffer) = False Then
                        If String.IsNullOrEmpty(buffer) = False Then
                            Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                    Me.mIndex = e.RowIndex
                    Me.BackgroundWorker4.RunWorkerAsync()
                End If
            Else
                If IsNothing(Me.buffer) = False Then
                    If String.IsNullOrEmpty(buffer) = False Then
                        Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                Me.mIndex = e.RowIndex
                Me.BackgroundWorker4.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getIndexWorkUnit(ByVal vRowException As Integer, ByVal vPrimary As Double) As Integer
        Try
            For Each dr As DataGridViewRow In Me.dgvWorkUnit.Rows
                With dr
                    If Me.dgvWorkUnit.Rows.IndexOf(dr) <> vRowException Then
                        If .Cells(Me.colWU_Primary.Index).Value <> -1 Then
                            If .Cells(Me.colWU_Primary.Index).Value = vPrimary Then
                                Return Me.dgvWorkUnit.Rows.IndexOf(dr)
                            End If
                        End If
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvWorkUnit_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvWorkUnit.DataError

    End Sub

    Private Sub dgvWorkUnit_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvWorkUnit.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvWorkUnit.Rows(Me.dgvWorkUnit.Rows.Count - 1)
            dr.Cells(Me.colWU_No.Index).Value = Me.dgvWorkUnit.Rows.Count
            dr.Cells(Me.colWU_Code.Index).Value = ""
            dr.Cells(Me.colWU_Desc.Index).Value = ""
            dr.Cells(Me.colWU_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWorkUnit_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvWorkUnit.RowsRemoved
        Try
            Me.objCd.SetNumber(Me.dgvWorkUnit)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWorkUnit_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvWorkUnit.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvWorkUnit)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            If Me.dgvWorkUnit.Rows.Count > 1 Then
                Me.dgvWorkUnit.Rows.RemoveAt(Me.mIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboWU_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboWU.CheckedChanged
        Try
            If Me.rboWU.Checked = True Then
                Me.dgvWorkUnit.ReadOnly = False
                Me.dgvWorkUnit.AllowUserToAddRows = True

                Me.rboWUG.Checked = False
                Me.dgvWUG.Rows.Clear()
                Me.dgvWUG.ReadOnly = True
                Me.dgvWUG.AllowUserToAddRows = False

                Me.chkWorkUnit.Items.Clear()
                Me.chkWorkUnit.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboWUG_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rboWUG.CheckedChanged
        Try
            If Me.rboWUG.Checked = True Then
                Me.dgvWorkUnit.Rows.Clear()
                Me.dgvWorkUnit.ReadOnly = True
                Me.dgvWorkUnit.AllowUserToAddRows = False

                Me.rboWU.Checked = False
                Me.dgvWUG.ReadOnly = False
                Me.dgvWUG.AllowUserToAddRows = True

                Me.chkWorkUnit.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWUG_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvWUG.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colWU_Code.Index Then
                Me.buffer = Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub addToList(ByVal vPrimaryWUG As Double, ByVal vAdd As Boolean)
        Try
            Me.tblWU.DefaultView.RowFilter = "[primary work unit group] = " & vPrimaryWUG.ToString
            If vAdd = True Then
                For i As Integer = 0 To Me.tblWU.DefaultView.Count - 1
                    Me.rowWU = Me.tblWU.DefaultView(i).Row
                    Me.rowWU2 = Me.tblWU2.Newtblmaster_workunitRow
                    For i2 As Integer = 0 To Me.tblWU.Columns.Count - 1
                        Me.rowWU2.Item(i2) = Me.rowWU.Item(i2)
                    Next
                    Me.tblWU2.Rows.Add(Me.rowWU2)
                Next
            Else
                Me.tblWU2.DefaultView.RowFilter = "[primary work unit group] = " & vPrimaryWUG.ToString
                For i As Integer = Me.tblWU2.DefaultView.Count - 1 To 0 Step -1
                    Me.rowWU2 = Me.tblWU2.DefaultView(i).Row
                    Me.tblWU2.Rows.Remove(Me.rowWU2)
                Next
            End If
            Me.chkWorkUnit.Items.Clear()
            For Each Me.rowWU2 In Me.tblWU2
                If Me.rowWU2.Show = 0 Then
                    Me.chkWorkUnit.Items.Add(Me.rowWU2.work_unit_code, False)
                Else
                    Me.chkWorkUnit.Items.Add(Me.rowWU2.work_unit_code, True)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWUG_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWUG.CellEndEdit
        Dim frm As New Core_Module.FrmMasterWorkUnitGroup_Cari
        Try
            If IsNothing(Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.buffer = Me.dgvWUG(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.colWUG_Code.Index Then
                        If Me.objWorkUnitGroup.Apakahsudahada(Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value = Me.objWorkUnitGroup.GetPrimary(Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvWUG.Item(Me.colWUG_Desc.Index, e.RowIndex).Value = Me.objWorkUnitGroup.GetDesc(Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value)
                            If getIndexWorkUnitGroup(e.RowIndex, Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value) <> -1 Then
                                MsgBox("Work Unit Group" & Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                If IsNothing(Me.buffer) = False Then
                                    If String.IsNullOrEmpty(buffer) = False Then
                                        Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value = Me.objWorkUnitGroup.GetPrimary(Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value)
                                        Me.dgvWUG.Item(Me.colWUG_Desc.Index, e.RowIndex).Value = Me.objWorkUnitGroup.GetDesc(Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value)
                                        Exit Sub
                                    End If
                                End If
                                Me.mIndex = e.RowIndex
                                Me.BackgroundWorker5.RunWorkerAsync()
                                Exit Sub
                            End If
                            Me.addToList(Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value, True)
                        Else
                            frm.txtCari.Text = Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value
                            frm.ShowDialog()
                            If frm.PrimaryWorkUnitGroup > 0 Then
                                Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value = frm.PrimaryWorkUnitGroup
                                Me.dgvWUG.Item(Me.colWUG_Code.Index, e.RowIndex).Value = Me.objWorkUnitGroup.GetCode(Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value)
                                Me.dgvWUG.Item(Me.colWUG_Desc.Index, e.RowIndex).Value = Me.objWorkUnitGroup.GetDesc(Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value)
                                If getIndexWorkUnitGroup(e.RowIndex, Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value) <> -1 Then
                                    MsgBox("Work Unit Group" & Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                    If IsNothing(Me.buffer) = False Then
                                        If String.IsNullOrEmpty(buffer) = False Then
                                            Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                            Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value = Me.objWorkUnitGroup.GetPrimary(Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value)
                                            Me.dgvWUG.Item(Me.colWUG_Desc.Index, e.RowIndex).Value = Me.objWorkUnitGroup.GetDesc(Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value)
                                            Exit Sub
                                        End If
                                    End If
                                    Me.mIndex = e.RowIndex
                                    Me.BackgroundWorker5.RunWorkerAsync()
                                    Exit Sub
                                End If
                                Me.addToList(Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.RowIndex).Value, True)
                            Else
                                If IsNothing(Me.buffer) = False Then
                                    If String.IsNullOrEmpty(buffer) = False Then
                                        Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                                Me.mIndex = e.RowIndex
                                Me.BackgroundWorker5.RunWorkerAsync()
                            End If
                        End If
                    End If
                Else
                    If IsNothing(Me.buffer) = False Then
                        If String.IsNullOrEmpty(buffer) = False Then
                            Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                    Me.mIndex = e.RowIndex
                    Me.BackgroundWorker5.RunWorkerAsync()
                End If
            Else
                If IsNothing(Me.buffer) = False Then
                    If String.IsNullOrEmpty(buffer) = False Then
                        Me.dgvWUG.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                Me.mIndex = e.RowIndex
                Me.BackgroundWorker5.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getIndexWorkUnitGroup(ByVal vRowException As Integer, ByVal vPrimary As Double) As Integer
        Try
            For Each dr As DataGridViewRow In Me.dgvWUG.Rows
                With dr
                    If Me.dgvWUG.Rows.IndexOf(dr) <> vRowException Then
                        If .Cells(Me.colWUG_Primary.Index).Value <> -1 Then
                            If .Cells(Me.colWUG_Primary.Index).Value = vPrimary Then
                                Return Me.dgvWUG.Rows.IndexOf(dr)
                            End If
                        End If
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvWUG_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvWUG.DataError

    End Sub

    Private Sub dgvWUG_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvWUG.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvWUG.Rows(Me.dgvWUG.Rows.Count - 1)
            dr.Cells(Me.colWUG_No.Index).Value = Me.dgvWUG.Rows.Count
            dr.Cells(Me.colWUG_Code.Index).Value = ""
            dr.Cells(Me.colWUG_Desc.Index).Value = ""
            dr.Cells(Me.colWUG_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWUG_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvWUG.RowsRemoved
        Try
            Me.objCd.SetNumber(Me.dgvWUG)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWUG_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvWUG.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvWUG)
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
            If Me.dgvWUG.Rows.Count > 1 Then
                Me.dgvWUG.Rows.RemoveAt(Me.mIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkWorkUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkWorkUnit.Leave
        Try
            For i As Integer = 0 To chkWorkUnit.Items.Count - 1
                Me.tblWU2.DefaultView.RowFilter = "[work unit code] = '" & Me.chkWorkUnit.Items.Item(i).ToString & "'"
                Me.rowWU2 = Me.tblWU2.DefaultView(0).Row
                If Me.chkWorkUnit.GetItemCheckState(i) = CheckState.Checked Then
                    Me.rowWU2.Show = 1
                Else
                    Me.rowWU2.Show = 0
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim primaryWUG As Double = -1

    Private Sub dgvWUG_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvWUG.UserDeletedRow
        Try
            Me.addToList(Me.primaryWUG, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWUG_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvWUG.UserDeletingRow
        Try
            If IsNothing(Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.Row.Index).Value) = False Then
                Me.primaryWUG = Me.dgvWUG.Item(Me.colWUG_Primary.Index, e.Row.Index).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FilterByProduct(ByRef vTbl As Data.DataTable)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim filter As String = ""
        Try
            tblDummy = vTbl.Copy
            For Each dr As DataGridViewRow In Me.dgvProduct.Rows
                With dr
                    If dr.Cells(Me.colProduct_Primary.Index).Value <> -1 Then
                        If dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Group" Then
                            If filter = "" Then
                                filter = "[Primary Group] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Group] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                        If dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Brand" Then
                            If filter = "" Then
                                filter = "[Primary Brand] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Brand] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                        If dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Sub Brand" Then
                            If filter = "" Then
                                filter = "[Primary Sub Brand] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Sub Brand] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                        If dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Category" Then
                            If filter = "" Then
                                filter = "[Primary Category] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Category] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                        If dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Product" Then
                            If filter = "" Then
                                filter = "[Primary Product] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Product] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                    End If
                End With
            Next
            tblDummy.DefaultView.RowFilter = filter
            vTbl.Rows.Clear()
            For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                rowDummy = tblDummy.DefaultView(i).Row
                rowTbl = vTbl.NewRow
                For ii As Integer = 0 To tblDummy.Columns.Count - 1
                    rowTbl.Item(ii) = rowDummy.Item(ii)
                Next
                vTbl.Rows.Add(rowTbl)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FilterByProcess(ByRef vTbl As Data.DataTable)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim filter As String = ""
        Try
            tblDummy = vTbl.Copy
            For Each dr As DataGridViewRow In Me.dgvProcess.Rows
                With dr
                    If dr.Cells(Me.colProcess_Primary.Index).Value <> -1 Then
                        If filter = "" Then
                            filter = "[Primary Process] = " & .Cells(Me.colProcess_Primary.Index).Value.ToString
                        Else
                            filter += " OR [Primary Process] = " & .Cells(Me.colProcess_Primary.Index).Value.ToString
                        End If
                    End If
                End With
            Next
            tblDummy.DefaultView.RowFilter = filter
            vTbl.Rows.Clear()
            For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                rowDummy = tblDummy.DefaultView(i).Row
                rowTbl = vTbl.NewRow
                For ii As Integer = 0 To tblDummy.Columns.Count - 1
                    rowTbl.Item(ii) = rowDummy.Item(ii)
                Next
                vTbl.Rows.Add(rowTbl)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FilterByWorkshift(ByRef vTbl As Data.DataTable)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim filter As String = ""
        Try
            tblDummy = vTbl.Copy
            For Each dr As DataGridViewRow In Me.dgvProcess.Rows
                With dr
                    If dr.Cells(Me.colProcess_Primary.Index).Value <> -1 Then
                        If filter = "" Then
                            filter = "[Primary Work Shift] = " & .Cells(Me.colProcess_Primary.Index).Value.ToString
                        Else
                            filter += " OR [Primary Work Shift] = " & .Cells(Me.colProcess_Primary.Index).Value.ToString
                        End If
                    End If
                End With
            Next
            tblDummy.DefaultView.RowFilter = filter
            vTbl.Rows.Clear()
            For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                rowDummy = tblDummy.DefaultView(i).Row
                rowTbl = vTbl.NewRow
                For ii As Integer = 0 To tblDummy.Columns.Count - 1
                    rowTbl.Item(ii) = rowDummy.Item(ii)
                Next
                vTbl.Rows.Add(rowTbl)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FilterByWorkUnit(ByRef vTbl As Data.DataTable)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim filter As String = ""
        Try
            tblDummy = vTbl.Copy
            If Me.rboWU.Checked = True Then
                For Each dr As DataGridViewRow In Me.dgvWorkUnit.Rows
                    With dr
                        If dr.Cells(Me.colWU_Primary.Index).Value <> -1 Then
                            If filter = "" Then
                                filter = "[Primary Work Unit] = " & .Cells(Me.colWU_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Work Unit] = " & .Cells(Me.colWU_Primary.Index).Value.ToString
                            End If
                        End If
                    End With
                Next
            Else
                For i As Integer = 0 To Me.chkWorkUnit.Items.Count - 1
                    If Me.chkWorkUnit.GetItemCheckState(i) = CheckState.Checked Then
                        If filter = "" Then
                            filter = "[Work Unit Code] = '" & Me.chkWorkUnit.Items.Item(i).ToString & "'"
                        Else
                            filter += " OR [Work Unit Code] = '" & Me.chkWorkUnit.Items.Item(i).ToString & "'"
                        End If
                    End If
                Next
            End If

            tblDummy.DefaultView.RowFilter = filter
            vTbl.Rows.Clear()
            For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                rowDummy = tblDummy.DefaultView(i).Row
                rowTbl = vTbl.NewRow
                For ii As Integer = 0 To tblDummy.Columns.Count - 1
                    rowTbl.Item(ii) = rowDummy.Item(ii)
                Next
                vTbl.Rows.Add(rowTbl)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Try
            MsgBox("belum:")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class