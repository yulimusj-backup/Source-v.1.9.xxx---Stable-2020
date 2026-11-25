Option Explicit On
Option Strict Off

Public Class FrmLaporanProd_RejectReport
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

    Private Sub FrmLaporanProd_RejectReport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.cboFormat.SelectedIndex = 0
            Me.cboTop.SelectedIndex = 0
            Me.dgvProduct.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvProcess.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvMO.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
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

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
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

            If Me.cboTop.Text = 5 Then
                If Me.cboFormat.SelectedIndex = 0 Or Me.cboFormat.SelectedIndex = 2 Then
                    Simple0()
                    Exit Sub
                ElseIf Me.cboFormat.SelectedIndex = 1 Or Me.cboFormat.SelectedIndex = 3 Then
                    Simple1()
                    Exit Sub
                ElseIf Me.cboFormat.SelectedIndex = 4 Or Me.cboFormat.SelectedIndex = 5 Then
                    Simple1b()
                    Exit Sub
                End If
            End If
            MsgBox("On Progress", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Simple0()
        Dim ds As New DseLaporanProd_RejectReport
        Dim dta As New DseLaporanProd_RejectReportTableAdapters.DataTable1TableAdapter
        Dim tbl As New DseLaporanProd_RejectReport.DataTable1DataTable
        Dim tblPrint As DseLaporanProd_RejectReport.PrintDataTable = ds.Print
        Dim rowPrint As DseLaporanProd_RejectReport.PrintRow
        Dim report As New CrpLaporanProd_RejectReport0
        Dim counter As Integer = 1
        Dim prevCode As String = ""
        Dim currentCode As String = ""
        Dim sort As String = ""
        Dim paramCode As String = ""
        Dim paramDesc As String = ""
        Try
            Me.Cursor = Cursors.WaitCursor
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If Me.cboFormat.SelectedIndex = 0 Then
                dta.Fill_GroupByProduct(tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                sort = "[Product Code], [Reject Qty] DESC"
                paramCode = "PRODUCT CODE"
                paramDesc = "PRODUCT DESC"
            ElseIf Me.cboFormat.SelectedIndex = 2 Then
                dta.Fill_GroupByProcess(tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                sort = "[Process Code], [Reject Qty] DESC"
                paramCode = "PROCESS CODE"
                paramDesc = "PROCESS NAME"
            End If
            FilterByProduct(tbl)
            FilterByProcess(tbl)
            FilterByMO(tbl)
            If tbl.Rows.Count > 0 Then
                SortTBL(tbl, sort)
                For Each row As DseLaporanProd_RejectReport.DataTable1Row In tbl
                    If Me.cboFormat.SelectedIndex = 0 Then
                        currentCode = row.Product_Code
                        tblPrint.DefaultView.RowFilter = "[Code] = '" & row.Product_Code & "'"
                    ElseIf Me.cboFormat.SelectedIndex = 2 Then
                        currentCode = row.Process_Code
                        tblPrint.DefaultView.RowFilter = "[Code] = '" & row.Process_Code & "'"
                    End If
                    If currentCode = prevCode Then
                        counter += 1
                        rowPrint = tblPrint.DefaultView(0).Row
                        Select Case counter
                            Case 2
                                rowPrint.Reject_Code_2 = row.Reject_Code
                                rowPrint.Reject_Qty_2 = row.Reject_Qty
                            Case 3
                                rowPrint.Reject_Code_3 = row.Reject_Code
                                rowPrint.Reject_Qty_3 = row.Reject_Qty
                            Case 4
                                rowPrint.Reject_Code_4 = row.Reject_Code
                                rowPrint.Reject_Qty_4 = row.Reject_Qty
                            Case 5
                                rowPrint.Reject_Code_5 = row.Reject_Code
                                rowPrint.Reject_Qty_5 = row.Reject_Qty
                            Case Else
                                rowPrint.Others_Qty += row.Reject_Qty
                        End Select
                    Else
                        counter = 1
                        rowPrint = tblPrint.NewPrintRow
                        With rowPrint
                            If Me.cboFormat.SelectedIndex = 0 Then
                                .Code = row.Product_Code
                                .Desc = row.Product_Desc
                            ElseIf Me.cboFormat.SelectedIndex = 2 Then
                                .Code = row.Process_Code
                                .Desc = row.Process_Desc
                            End If
                            .Reject_Code_1 = row.Reject_Code
                            .Reject_Qty_1 = row.Reject_Qty
                        End With
                        tblPrint.Rows.Add(rowPrint)
                    End If
                    If Me.cboFormat.SelectedIndex = 0 Then
                        prevCode = row.Product_Code
                    ElseIf Me.cboFormat.SelectedIndex = 2 Then
                        prevCode = row.Process_Code
                    End If
                Next
                report.SetDataSource(ds)
                report.SetParameterValue(report.Parameter_Code.ParameterFieldName, paramCode)
                report.SetParameterValue(report.Parameter_Desc.ParameterFieldName, paramDesc)
                Me.crvReport.ReportSource = report
                Me.crvReport.Zoom(80)
                Me.crvReport.Refresh()
                Me.crvReport.Show()
            Else
                MsgBox("No Data on this period.", MsgBoxStyle.Information, "Information")
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

    Private Sub Simple1()
        Dim ds As New DseLaporanProd_RejectReport
        Dim dta As New DseLaporanProd_RejectReportTableAdapters.DataTable1TableAdapter
        Dim tbl As New DseLaporanProd_RejectReport.DataTable1DataTable
        Dim tblPrint As DseLaporanProd_RejectReport.PrintDataTable = ds.Print
        Dim rowPrint As DseLaporanProd_RejectReport.PrintRow
        Dim report As New CrpLaporanProd_RejectReport1
        Dim counter As Integer = 1
        Dim prevCode As String = ""
        Dim currentCode As String = ""
        Dim sort As String = ""
        Dim paramCode As String = ""
        Dim paramDesc As String = ""
        Dim lst As New List(Of String)
        Try
            Me.Cursor = Cursors.WaitCursor
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If Me.cboFormat.SelectedIndex = 1 Then
                dta.Fill_GroupByProduct(tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                sort = "[Product Code], [Reject Qty] DESC"
                paramCode = "PRODUCT CODE"
                paramDesc = "PRODUCT DESC"
            ElseIf Me.cboFormat.SelectedIndex = 3 Then
                dta.Fill_GroupByProcess(tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                sort = "[Process Code], [Reject Qty] DESC"
                paramCode = "PROCESS CODE"
                paramDesc = "PROCESS NAME"
            End If
            FilterByProduct(tbl)
            FilterByProcess(tbl)
            FilterByMO(tbl)
            lst = GetTop(Me.cboTop.Text, tbl)
            If tbl.Rows.Count > 0 Then
                SortTBL(tbl, sort)
                For Each row As DseLaporanProd_RejectReport.DataTable1Row In tbl
                    If Me.cboFormat.SelectedIndex = 1 Then
                        currentCode = row.Product_Code
                        tblPrint.DefaultView.RowFilter = "[Code] = '" & row.Product_Code & "'"
                    ElseIf Me.cboFormat.SelectedIndex = 3 Then
                        currentCode = row.Process_Code
                        tblPrint.DefaultView.RowFilter = "[Code] = '" & row.Process_Code & "'"
                    End If
                    If currentCode = prevCode Then
                        counter += 1
                        rowPrint = tblPrint.DefaultView(0).Row
                        Select Case lst.IndexOf(row.Reject_Code)
                            Case 0
                                rowPrint.Reject_Qty_1 += row.Reject_Qty
                            Case 1
                                rowPrint.Reject_Qty_2 += row.Reject_Qty
                            Case 2
                                rowPrint.Reject_Qty_3 += row.Reject_Qty
                            Case 3
                                rowPrint.Reject_Qty_4 += row.Reject_Qty
                            Case 4
                                rowPrint.Reject_Qty_5 += row.Reject_Qty
                            Case Else
                                rowPrint.Others_Qty += row.Reject_Qty
                        End Select
                    Else
                        counter = 1
                        rowPrint = tblPrint.NewPrintRow
                        With rowPrint
                            If Me.cboFormat.SelectedIndex = 1 Then
                                .Code = row.Product_Code
                                .Desc = row.Product_Desc
                            ElseIf Me.cboFormat.SelectedIndex = 3 Then
                                .Code = row.Process_Code
                                .Desc = row.Process_Desc
                            End If
                            If lst.Count > 0 Then
                                .Reject_Code_1 = lst.Item(0)
                            End If
                            If lst.Count > 1 Then
                                .Reject_Code_2 = lst.Item(1)
                            End If
                            If lst.Count > 2 Then
                                .Reject_Code_3 = lst.Item(2)
                            End If
                            If lst.Count > 3 Then
                                .Reject_Code_4 = lst.Item(3)
                            End If
                            If lst.Count > 4 Then
                                .Reject_Code_5 = lst.Item(4)
                            End If
                            Select Case lst.IndexOf(row.Reject_Code)
                                Case 0
                                    rowPrint.Reject_Qty_1 += row.Reject_Qty
                                Case 1
                                    rowPrint.Reject_Qty_2 += row.Reject_Qty
                                Case 2
                                    rowPrint.Reject_Qty_3 += row.Reject_Qty
                                Case 3
                                    rowPrint.Reject_Qty_4 += row.Reject_Qty
                                Case 4
                                    rowPrint.Reject_Qty_5 += row.Reject_Qty
                                Case Else
                                    rowPrint.Others_Qty += row.Reject_Qty
                            End Select
                        End With
                        tblPrint.Rows.Add(rowPrint)
                    End If
                    If Me.cboFormat.SelectedIndex = 1 Then
                        prevCode = row.Product_Code
                    ElseIf Me.cboFormat.SelectedIndex = 3 Then
                        prevCode = row.Process_Code
                    End If
                Next
                report.SetDataSource(ds)
                report.SetParameterValue(report.Parameter_Code.ParameterFieldName, paramCode)
                report.SetParameterValue(report.Parameter_Desc.ParameterFieldName, paramDesc)
                Me.crvReport.ReportSource = report
                Me.crvReport.Zoom(80)
                Me.crvReport.Refresh()
                Me.crvReport.Show()
            Else
                MsgBox("No Data on this period.", MsgBoxStyle.Information, "Information")
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

    Private Sub Simple1b()
        Dim ds As New DseLaporanProd_RejectReport
        Dim dta As New DseLaporanProd_RejectReportTableAdapters.DataTable1TableAdapter
        Dim tbl As New DseLaporanProd_RejectReport.DataTable1DataTable
        Dim tblPrint As DseLaporanProd_RejectReport.PrintDataTable = ds.Print
        Dim rowPrint As DseLaporanProd_RejectReport.PrintRow
        Dim report As New CrpLaporanProd_RejectReport1b
        Dim counter As Integer = 1
        Dim prevCode As String = ""
        Dim currentCode As String = ""
        Dim sort As String = ""
        Dim paramCode As String = ""
        Dim paramDesc As String = ""
        Dim paramCodeII As String = ""
        Dim paramDescII As String = ""
        Dim lst As New List(Of String)
        Try
            Me.Cursor = Cursors.WaitCursor
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill_GroupByProductProcess(tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
            If Me.cboFormat.SelectedIndex = 4 Then
                sort = "[Product Code], [Process Code], [Reject Qty] DESC"
                paramCode = "PRODUCT CODE"
                paramDesc = "PRODUCT DESC"
                paramDescII = "PROCESS"
            ElseIf Me.cboFormat.SelectedIndex = 5 Then
                sort = "[Process Code], [Product Code], [Reject Qty] DESC"
                paramCode = "PROCESS"
                paramDesc = "PRODUCT CODE"
                paramDescII = "PRODUCT DESC"
            End If
            FilterByProduct(tbl)
            FilterByProcess(tbl)
            FilterByMO(tbl)
            lst = GetTop(Me.cboTop.Text, tbl)
            If tbl.Rows.Count > 0 Then
                SortTBL(tbl, sort)
                For Each row As DseLaporanProd_RejectReport.DataTable1Row In tbl
                    If Me.cboFormat.SelectedIndex = 4 Then
                        currentCode = row.Product_Code & "-" & row.Process_Code
                        tblPrint.DefaultView.RowFilter = "[Code] = '" & row.Product_Code & "'"
                    ElseIf Me.cboFormat.SelectedIndex = 5 Then
                        currentCode = row.Process_Code & "-" & row.Product_Code
                        tblPrint.DefaultView.RowFilter = "[Code] = '" & row.Process_Desc & "'"
                    End If
                    If currentCode = prevCode Then
                        counter += 1
                        rowPrint = tblPrint.DefaultView(0).Row
                        Select Case lst.IndexOf(row.Reject_Code)
                            Case 0
                                rowPrint.Reject_Qty_1 += row.Reject_Qty
                            Case 1
                                rowPrint.Reject_Qty_2 += row.Reject_Qty
                            Case 2
                                rowPrint.Reject_Qty_3 += row.Reject_Qty
                            Case 3
                                rowPrint.Reject_Qty_4 += row.Reject_Qty
                            Case 4
                                rowPrint.Reject_Qty_5 += row.Reject_Qty
                            Case Else
                                rowPrint.Others_Qty += row.Reject_Qty
                        End Select
                    Else
                        counter = 1
                        rowPrint = tblPrint.NewPrintRow
                        With rowPrint
                            If Me.cboFormat.SelectedIndex = 4 Then
                                .Code = row.Product_Code
                                .Desc = row.Product_Desc
                                .Desc_II = row.Process_Desc
                            ElseIf Me.cboFormat.SelectedIndex = 5 Then
                                .Code = row.Process_Desc
                                .Desc = row.Product_Code
                                .Desc_II = row.Product_Desc
                            End If
                            If lst.Count > 0 Then
                                .Reject_Code_1 = lst.Item(0)
                            End If
                            If lst.Count > 1 Then
                                .Reject_Code_2 = lst.Item(1)
                            End If
                            If lst.Count > 2 Then
                                .Reject_Code_3 = lst.Item(2)
                            End If
                            If lst.Count > 3 Then
                                .Reject_Code_4 = lst.Item(3)
                            End If
                            If lst.Count > 4 Then
                                .Reject_Code_5 = lst.Item(4)
                            End If
                            Select Case lst.IndexOf(row.Reject_Code)
                                Case 0
                                    rowPrint.Reject_Qty_1 += row.Reject_Qty
                                Case 1
                                    rowPrint.Reject_Qty_2 += row.Reject_Qty
                                Case 2
                                    rowPrint.Reject_Qty_3 += row.Reject_Qty
                                Case 3
                                    rowPrint.Reject_Qty_4 += row.Reject_Qty
                                Case 4
                                    rowPrint.Reject_Qty_5 += row.Reject_Qty
                                Case Else
                                    rowPrint.Others_Qty += row.Reject_Qty
                            End Select
                        End With
                        tblPrint.Rows.Add(rowPrint)
                    End If
                    If Me.cboFormat.SelectedIndex = 4 Then
                        prevCode = row.Product_Code & "-" & row.Process_Code
                    ElseIf Me.cboFormat.SelectedIndex = 5 Then
                        prevCode = row.Process_Code & "-" & row.Product_Code
                    End If
                Next
                report.SetDataSource(ds)
                report.SetParameterValue(report.Parameter_Code.ParameterFieldName, paramCode)
                report.SetParameterValue(report.Parameter_Desc.ParameterFieldName, paramDesc)
                report.SetParameterValue(report.Parameter_CodeII.ParameterFieldName, paramCodeII)
                report.SetParameterValue(report.Parameter_DescII.ParameterFieldName, paramDescII)
                Me.crvReport.ReportSource = report
                Me.crvReport.Zoom(80)
                Me.crvReport.Refresh()
                Me.crvReport.Show()
            Else
                MsgBox("No Data on this period.", MsgBoxStyle.Information, "Information")
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

    Private Function GetTop(ByVal vTop As Integer, ByVal vTbl As DseLaporanProd_RejectReport.DataTable1DataTable) As List(Of String)
        Dim tblTop As New DseLaporanProd_RejectReport.Top_NDataTable
        Dim rowTop As DseLaporanProd_RejectReport.Top_NRow
        Dim lst As New List(Of String)
        Try
            For Each dr As DseLaporanProd_RejectReport.DataTable1Row In vTbl
                tblTop.DefaultView.RowFilter = "[Reject Code] = '" & dr.Reject_Code & "'"
                If tblTop.DefaultView.Count = 0 Then
                    rowTop = tblTop.NewTop_NRow
                    rowTop.Reject_Code = dr.Reject_Code
                    rowTop.Reject_Qty = dr.Reject_Qty
                    tblTop.Rows.Add(rowTop)
                Else
                    rowTop = tblTop.DefaultView(0).Row
                    rowTop.Reject_Qty += dr.Reject_Qty
                End If
            Next
            tblTop.DefaultView.RowFilter = ""
            tblTop.DefaultView.Sort = "[Reject Qty] DESC"
            For i As Integer = 0 To tblTop.DefaultView.Count - 1
                rowTop = tblTop.DefaultView(i).Row
                If lst.Count < vTop Then
                    lst.Add(rowTop.Reject_Code)
                End If
            Next
            Return lst
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return lst
        End Try
    End Function

    Private Sub SortTBL(ByRef vTbl As Data.DataTable, ByVal vSortString As String)
        Dim tbl As New Data.DataTable
        Dim row As Data.DataRow
        Dim vRow As Data.DataRow
        Try
            tbl = vTbl.Copy
            tbl.DefaultView.Sort = vSortString
            vTbl.Rows.Clear()
            For Each dataRowView_ As Data.DataRowView In tbl.DefaultView
                row = dataRowView_.Row
                vRow = vTbl.NewRow
                For i As Integer = 0 To vTbl.Columns.Count - 1
                    vRow.Item(i) = row.Item(i)
                Next
                vTbl.Rows.Add(vRow)
            Next
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

    Private Sub FilterByMO(ByRef vTbl As Data.DataTable)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim filter As String = ""
        Try
            tblDummy = vTbl.Copy
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
                    rowTbl = vTbl.NewRow
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

    Private Sub _GroupByProductResetPerProduct()
        'Dim dta As New DseLaporanProd_RejectReportTableAdapters.DataTable1TableAdapter
        'Dim ds As New DseLaporanProd_RejectReport
        'Dim tbl As DseLaporanProd_RejectReport.DataTable1DataTable = ds.DataTable1
        'Dim row As DseLaporanProd_RejectReport.DataTable1Row
        'Dim row_ As DseLaporanProd_RejectReport.DataTable1Row
        'Dim report As New CrpLaporanProd_RejectReport
        'Dim previousPrimaryProduct As Double = -1
        'Dim counter As Integer = 0
        'Try
        '    Me.Cursor = Cursors.WaitCursor
        '    dta.Connection.ConnectionString = Core_Module.GetConstring
        '    If dta.Fill(tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date) > 0 Then
        '        ''hanya sisakan yang top N, buat baris untuk others
        '        tbl.DefaultView.RowFilter = ""
        '        tbl.DefaultView.Sort = "[Product Code] ASC, [Reject Qty] DESC"
        '        For i As Integer = 0 To tbl.DefaultView.Count - 1
        '            row = tbl.DefaultView(i).Row
        '            If row.Primary_Product = previousPrimaryProduct Then
        '                counter += 1
        '            Else
        '                counter = 1
        '            End If
        '            row.Counter = counter
        '            row.Reject_Code = row.Counter.ToString & ". " & row.Reject_Code
        '            previousPrimaryProduct = row.Primary_Product
        '        Next
        '        For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
        '            row = tbl.Rows(i)
        '            If row.Counter > Me.txtShowTop.Text Then
        '                tbl.DefaultView.RowFilter = "[Primary Product] = " & row.Primary_Product.ToString & " AND [Counter] = " & (CInt(Me.txtShowTop.Text) + 1).ToString
        '                If tbl.DefaultView.Count > 0 Then
        '                    row_ = tbl.DefaultView(0).Row
        '                    row.Primary_Reject = -1
        '                    row.Reject_Code = "OTHERS"
        '                    row.Reject_Desc = "OTHERS"
        '                    If row.Counter <> CInt(Me.txtShowTop.Text) + 1 Then
        '                        row_.Reject_Qty += row.Reject_Qty
        '                    End If
        '                Else
        '                    row_ = tbl.NewDataTable1Row
        '                    For i2 As Integer = 0 To tbl.Columns.Count - 1
        '                        row_.Item(i2) = row.Item(i2)
        '                    Next
        '                    row.Primary_Reject = -1
        '                    row.Reject_Code = "OTHERS"
        '                    row.Reject_Desc = "OTHERS"
        '                    tbl.Rows.Add(row_)
        '                End If
        '                If row.Counter > CInt(Me.txtShowTop.Text) + 1 Then
        '                    tbl.Rows.RemoveAt(i)
        '                End If
        '            End If
        '        Next
        '        _ReSortTbl(tbl)
        '        report.SetDataSource(ds)
        '        Me.crvReport.ReportSource = report
        '        Me.crvReport.Refresh()
        '        Me.crvReport.Show()
        '    Else
        '        MsgBox("There is no data on this period.", MsgBoxStyle.Information, "Information")
        '        Me.crvReport.ReportSource = Nothing
        '        Me.crvReport.Refresh()
        '        Me.crvReport.Show()
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    Me.Cursor = Cursors.Default
        'End Try
    End Sub

    Private Sub _ReSortTbl(ByRef vTbl As DseLaporanProd_RejectReport.DataTable1DataTable)
        'Dim vTbl_ As New DseLaporanProd_RejectReport.DataTable1DataTable
        'Dim vRow, vRow_ As DseLaporanProd_RejectReport.DataTable1Row
        'Try
        '    For Each vRow In vTbl
        '        vRow_ = vTbl_.NewDataTable1Row
        '        For i As Integer = 0 To vTbl.Columns.Count - 1
        '            vRow_.Item(i) = vRow.Item(i)
        '        Next
        '        vTbl_.Rows.Add(vRow_)
        '    Next
        '    vTbl.Rows.Clear()
        '    vTbl_.DefaultView.RowFilter = ""
        '    vTbl_.DefaultView.Sort = "[Product Code] ASC, [Reject Qty] DESC"
        '    For i As Integer = 0 To vTbl_.DefaultView.Count - 1
        '        vRow_ = vTbl_.DefaultView(i).Row
        '        vRow = vTbl.NewDataTable1Row
        '        For i2 As Integer = 0 To vTbl.Columns.Count - 1
        '            vRow.Item(i2) = vRow_.Item(i2)
        '        Next
        '        vTbl.Rows.Add(vRow)
        '    Next
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally

        'End Try
    End Sub

End Class