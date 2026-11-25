Option Explicit On
Option Strict Off

Public Class FrmLaporanProd_Downtime

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim objCd As New Core_Module.ClsCekData
    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim objProduct As New Core_Module.clsMasterProduct
    Dim objProcess As New Core_Module.ClsMasterProses
    Dim objWorkUnit As New Core_Module.ClsMasterWorkUnit
    Dim objMO As New ClsTrans_ManufacturingOrder
    Dim buffer As Object
    Dim index As Integer = -1
    Dim ObjLicense As New Core_Module.ClsLicense

    Enum EnumGroupLevel As Integer
        GROUP
        BRAND
        SUB_BRAND
        CATEGORY
        PRODUCT
    End Enum

    Private Sub FrmLaporanProd_Downtime_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvProduct.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvProcess.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvMO.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvWorkUnit.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
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
                    Me.index = e.RowIndex
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            Else
                If IsNothing(Me.buffer) = False Then
                    If String.IsNullOrEmpty(buffer) = False Then
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                Me.index = e.RowIndex
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
                    Me.index = e.RowIndex
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
                            Me.index = e.RowIndex
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
                            Me.index = e.RowIndex
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
                            Me.index = e.RowIndex
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
                            Me.index = e.RowIndex
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
                            Me.index = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
            End Select
            If Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndexProduct(e.RowIndex, Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value, Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value & " " & Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.index = e.RowIndex
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
                Me.dgvProduct.Rows.RemoveAt(Me.index)
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
                                Me.index = e.RowIndex
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
                                    Me.index = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                End If
                            Else
                                If IsNothing(Me.buffer) = False Then
                                    If String.IsNullOrEmpty(buffer) = False Then
                                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                                Me.index = e.RowIndex
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
                    Me.index = e.RowIndex
                    Me.BackgroundWorker2.RunWorkerAsync()
                End If
            Else
                If IsNothing(Me.buffer) = False Then
                    If String.IsNullOrEmpty(buffer) = False Then
                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                Me.index = e.RowIndex
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
                Me.dgvProcess.Rows.RemoveAt(Me.index)
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
                                Me.index = e.RowIndex
                                Me.BackgroundWorker3.RunWorkerAsync()
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
                                    Me.index = e.RowIndex
                                    Me.BackgroundWorker3.RunWorkerAsync()
                                End If
                            Else
                                If IsNothing(Me.buffer) = False Then
                                    If String.IsNullOrEmpty(buffer) = False Then
                                        Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                                Me.index = e.RowIndex
                                Me.BackgroundWorker3.RunWorkerAsync()
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
                    Me.index = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                End If
            Else
                If IsNothing(Me.buffer) = False Then
                    If String.IsNullOrEmpty(buffer) = False Then
                        Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                Me.index = e.RowIndex
                Me.BackgroundWorker3.RunWorkerAsync()
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

    Private Sub BackgroundWorker3_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            If Me.dgvWorkUnit.Rows.Count > 1 Then
                Me.dgvWorkUnit.Rows.RemoveAt(Me.index)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvMO_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvMO.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colMO_DocNo.Index Then
                Me.buffer = Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvMO_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMO.CellEndEdit
        Dim frm As New FrmTrans_ManufacturingOrder_Cari
        Try
            If IsNothing(Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.buffer = Me.dgvMO(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.colMO_DocNo.Index Then
                        If Me.objMO.Apakahsudahada(Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            Me.dgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value = Me.objMO.GetPrimary(Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvMO.Item(Me.colMO_Date.Index, e.RowIndex).Value = Format(Me.objMO.GetDate(Me.dgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value), "dd MMM yyyy")
                            If getIndexMO(e.RowIndex, Me.dgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value) <> -1 Then
                                MsgBox("Manufacturing Order " & Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                If IsNothing(Me.buffer) = False Then
                                    If String.IsNullOrEmpty(buffer) = False Then
                                        Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Me.dgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value = Me.objMO.GetPrimary(Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value)
                                        Me.dgvMO.Item(Me.colMO_Date.Index, e.RowIndex).Value = Format(Me.objMO.GetDate(Me.dgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value), "dd MMM yyyy")
                                        Exit Sub
                                    End If
                                End If
                                Me.index = e.RowIndex
                                Me.BackgroundWorker4.RunWorkerAsync()
                            End If
                        Else
                            frm.txtCari.Text = Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value
                            frm.ShowDialog()
                            If frm.Primary > 0 Then
                                Me.dgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value = frm.Primary
                                Me.dgvMO.Item(Me.colMO_DocNo.Index, e.RowIndex).Value = Me.objMO.GetNoTransaksi(Me.dgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value)
                                Me.dgvMO.Item(Me.colMO_Date.Index, e.RowIndex).Value = Format(Me.objMO.GetDate(Me.dgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value), "dd MMM yyyy")
                                If getIndexMO(e.RowIndex, Me.dgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value) <> -1 Then
                                    MsgBox("Manufacturing Order " & Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                    If IsNothing(Me.buffer) = False Then
                                        If String.IsNullOrEmpty(buffer) = False Then
                                            Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                            Me.dgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value = Me.objMO.GetPrimary(Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value)
                                            Me.dgvMO.Item(Me.colMO_Date.Index, e.RowIndex).Value = Format(Me.objMO.GetDate(Me.dgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value), "dd MMM yyyy")
                                            Exit Sub
                                        End If
                                    End If
                                    Me.index = e.RowIndex
                                    Me.BackgroundWorker4.RunWorkerAsync()
                                End If
                            Else
                                If IsNothing(Me.buffer) = False Then
                                    If String.IsNullOrEmpty(buffer) = False Then
                                        Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                                Me.index = e.RowIndex
                                Me.BackgroundWorker4.RunWorkerAsync()
                            End If
                        End If
                    End If
                Else
                    If IsNothing(Me.buffer) = False Then
                        If String.IsNullOrEmpty(buffer) = False Then
                            Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                    Me.index = e.RowIndex
                    Me.BackgroundWorker4.RunWorkerAsync()
                End If
            Else
                If IsNothing(Me.buffer) = False Then
                    If String.IsNullOrEmpty(buffer) = False Then
                        Me.dgvMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                Me.index = e.RowIndex
                Me.BackgroundWorker4.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getIndexMO(ByVal vRowException As Integer, ByVal vPrimary As Double) As Integer
        Try
            For Each dr As DataGridViewRow In Me.dgvMO.Rows
                With dr
                    If Me.dgvMO.Rows.IndexOf(dr) <> vRowException Then
                        If .Cells(Me.colMO_Primary.Index).Value <> -1 Then
                            If .Cells(Me.colMO_Primary.Index).Value = vPrimary Then
                                Return Me.dgvMO.Rows.IndexOf(dr)
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

    Private Sub dgvMO_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvMO.DataError

    End Sub

    Private Sub dgvMO_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvMO.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvMO.Rows(Me.dgvMO.Rows.Count - 1)
            dr.Cells(Me.colMO_No.Index).Value = Me.dgvMO.Rows.Count
            dr.Cells(Me.colMO_DocNo.Index).Value = ""
            dr.Cells(Me.colMO_Date.Index).Value = ""
            dr.Cells(Me.colMO_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvMO_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvMO.RowsRemoved
        Try
            Me.objCd.SetNumber(Me.dgvMO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvMO_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvMO.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvMO)
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
            If Me.dgvMO.Rows.Count > 1 Then
                Me.dgvMO.Rows.RemoveAt(Me.index)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FilterByProduct(ByRef vTbl As DseLaporanProd_Downtime.DataTable1DataTable)
        Dim tblDummy As New DseLaporanProd_Downtime.DataTable1DataTable
        Dim rowDummy As DseLaporanProd_Downtime.DataTable1Row
        Dim rowTbl As DseLaporanProd_Downtime.DataTable1Row
        Dim filter As String = ""
        Try
            For Each rowTbl In vTbl
                rowDummy = tblDummy.NewDataTable1Row
                For i As Integer = 0 To vTbl.Columns.Count - 1
                    rowDummy.Item(i) = rowTbl.Item(i)
                Next
                tblDummy.Rows.Add(rowDummy)
            Next
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
            If tblDummy.DefaultView.Count > 0 Then
                For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                    rowDummy = tblDummy.DefaultView(i).Row
                    rowTbl = vTbl.NewDataTable1Row
                    For ii As Integer = 0 To tblDummy.Columns.Count - 1
                        rowTbl.Item(ii) = rowDummy.Item(ii)
                    Next
                    vTbl.Rows.Add(rowTbl)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FilterByProcess(ByRef vTbl As DseLaporanProd_Downtime.DataTable1DataTable)
        Dim tblDummy As New DseLaporanProd_Downtime.DataTable1DataTable
        Dim rowDummy As DseLaporanProd_Downtime.DataTable1Row
        Dim rowTbl As DseLaporanProd_Downtime.DataTable1Row
        Dim filter As String = ""
        Try
            For Each rowTbl In vTbl
                rowDummy = tblDummy.NewDataTable1Row
                For i As Integer = 0 To vTbl.Columns.Count - 1
                    rowDummy.Item(i) = rowTbl.Item(i)
                Next
                tblDummy.Rows.Add(rowDummy)
            Next
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
            If tblDummy.DefaultView.Count > 0 Then
                For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                    rowDummy = tblDummy.DefaultView(i).Row
                    rowTbl = vTbl.NewDataTable1Row
                    For ii As Integer = 0 To tblDummy.Columns.Count - 1
                        rowTbl.Item(ii) = rowDummy.Item(ii)
                    Next
                    vTbl.Rows.Add(rowTbl)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FilterByMO(ByRef vTbl As DseLaporanProd_Downtime.DataTable1DataTable)
        Dim tblDummy As New DseLaporanProd_Downtime.DataTable1DataTable
        Dim rowDummy As DseLaporanProd_Downtime.DataTable1Row
        Dim rowTbl As DseLaporanProd_Downtime.DataTable1Row
        Dim filter As String = ""
        Try
            For Each rowTbl In vTbl
                rowDummy = tblDummy.NewDataTable1Row
                For i As Integer = 0 To vTbl.Columns.Count - 1
                    rowDummy.Item(i) = rowTbl.Item(i)
                Next
                tblDummy.Rows.Add(rowDummy)
            Next
            For Each dr As DataGridViewRow In Me.dgvMO.Rows
                With dr
                    If dr.Cells(Me.colMO_Primary.Index).Value <> -1 Then
                        If filter = "" Then
                            filter = "[Primary MO] = " & .Cells(Me.colMO_Primary.Index).Value.ToString
                        Else
                            filter += " OR [Primary MO] = " & .Cells(Me.colMO_Primary.Index).Value.ToString
                        End If
                    End If
                End With
            Next
            tblDummy.DefaultView.RowFilter = filter
            vTbl.Rows.Clear()
            If tblDummy.DefaultView.Count > 0 Then
                For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                    rowDummy = tblDummy.DefaultView(i).Row
                    rowTbl = vTbl.NewDataTable1Row
                    For ii As Integer = 0 To tblDummy.Columns.Count - 1
                        rowTbl.Item(ii) = rowDummy.Item(ii)
                    Next
                    vTbl.Rows.Add(rowTbl)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FilterByWU(ByRef vTbl As DseLaporanProd_Downtime.DataTable1DataTable)
        Dim tblDummy As New DseLaporanProd_Downtime.DataTable1DataTable
        Dim rowDummy As DseLaporanProd_Downtime.DataTable1Row
        Dim rowTbl As DseLaporanProd_Downtime.DataTable1Row
        Dim filter As String = ""
        Try
            For Each rowTbl In vTbl
                rowDummy = tblDummy.NewDataTable1Row
                For i As Integer = 0 To vTbl.Columns.Count - 1
                    rowDummy.Item(i) = rowTbl.Item(i)
                Next
                tblDummy.Rows.Add(rowDummy)
            Next
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
            tblDummy.DefaultView.RowFilter = filter
            vTbl.Rows.Clear()
            If tblDummy.DefaultView.Count > 0 Then
                For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                    rowDummy = tblDummy.DefaultView(i).Row
                    rowTbl = vTbl.NewDataTable1Row
                    For ii As Integer = 0 To tblDummy.Columns.Count - 1
                        rowTbl.Item(ii) = rowDummy.Item(ii)
                    Next
                    vTbl.Rows.Add(rowTbl)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim ds As New DseLaporanProd_Downtime
        Dim dta As New DseLaporanProd_DowntimeTableAdapters.DataTable1TableAdapter
        Dim tbl As DseLaporanProd_Downtime.DataTable1DataTable = ds.DataTable1
        Dim report As New CrpLaporanProd_Downtime
        Dim row As DseLaporanProd_Downtime.DataTable1Row
        Dim tblSubTotal As New DseLaporanProd_Downtime.DataTable2DataTable
        Dim rowSubTotal As DseLaporanProd_Downtime.DataTable2Row
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            Me.Cursor = Cursors.WaitCursor
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Fill(tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date) > 0 Then
                If dgvProduct.Rows.Count > 1 Then
                    FilterByProduct(tbl)
                End If
                If dgvProcess.Rows.Count > 1 Then
                    FilterByProcess(tbl)
                End If
                If dgvMO.Rows.Count > 1 Then
                    FilterByMO(tbl)
                End If
                If dgvWorkUnit.Rows.Count > 1 Then
                    FilterByWU(tbl)
                End If
                ''store on collection per DT
                For Each row In tbl
                    With row
                        .Down_Time_Duration_Str = Me.objCd.Convert_ToString(.Down_Time_Duration)
                        tblSubTotal.DefaultView.RowFilter = "[Primary DT] = " & .Primary_Down_Time.ToString
                        If tblSubTotal.DefaultView.Count = 0 Then
                            rowSubTotal = tblSubTotal.NewDataTable2Row
                            rowSubTotal.Primary_DT = .Primary_Down_Time
                            rowSubTotal.Duration = .Down_Time_Duration
                            tblSubTotal.Rows.Add(rowSubTotal)
                        Else
                            rowSubTotal = tblSubTotal.DefaultView(0).Row
                            rowSubTotal.Duration += .Down_Time_Duration
                        End If
                    End With
                Next
                ''load from collection
                For Each row In tbl
                    With row
                        tblSubTotal.DefaultView.RowFilter = "[Primary DT] = " & .Primary_Down_Time.ToString
                        If tblSubTotal.DefaultView.Count > 0 Then
                            rowSubTotal = tblSubTotal.DefaultView(0).Row
                            .Sub_Total_Per_DT = rowSubTotal.Duration
                            .Sub_Total_Per_DT_Str = Me.objCd.Convert_ToString(.Sub_Total_Per_DT)
                        End If
                    End With
                Next
                report.SetDataSource(ds)
                Me.crvReport.ReportSource = report
                Me.crvReport.Zoom(90)
                Me.crvReport.Refresh()
                Me.crvReport.Show()
            Else
                MsgBox("There is no data on this period.", MsgBoxStyle.Information, "Information")
                Me.crvReport.ReportSource = Nothing
                Me.crvReport.Refresh()
                Me.crvReport.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

End Class