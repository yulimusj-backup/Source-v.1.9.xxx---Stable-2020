Option Explicit On
Option Strict Off

Public Class FrmLaporanProd_RawMaterialUsage

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim objCd As New Core_Module.ClsCekData
    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim objProduct As New Core_Module.clsMasterProduct
    Dim objMO As New ClsTrans_ManufacturingOrder
    Dim buffer As Object
    Dim index As Integer = -1
    Dim ObjLicense As New Core_Module.ClsLicense

    Enum EnumFormat As Integer
        ProductRecap
        ProductDetail
        JORecap
        JODetail
        WORecap
        WODetail
    End Enum

    Private Sub FrmLaporanProd_RawMaterialUsage_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvProduct.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvMO.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.CboFormat.SelectedIndex = EnumFormat.ProductRecap
            Me.chkIncludeWIP.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkLogDateFrom_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkLogDateFrom.CheckedChanged
        Try
            If Me.ChkLogDateFrom.Checked Then
                Me.dtpFromBundlingDate.Enabled = True
            Else
                Me.dtpFromBundlingDate.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkMODate_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkMODate.CheckedChanged
        Try
            If Me.ChkMODate.Checked Then
                Me.dtpFromMODate.Enabled = True
                Me.dtpToMODate.Enabled = True
            Else
                Me.dtpFromMODate.Enabled = False
                Me.dtpToMODate.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Try
            If Me.ChkLogDateFrom.Checked And Me.dtpFromBundlingDate.Value.Date > Me.dtpToBundlingDate.Value.Date Then
                MsgBox("Invalid log date filter", MsgBoxStyle.Information, "Information")
                Return
            End If
            If Me.ChkMODate.Checked And Me.dtpFromMODate.Value.Date > Me.dtpToMODate.Value.Date Then
                MsgBox("Invalid MO date filter", MsgBoxStyle.Information, "Information")
                Return
            End If
            If (Me.CboFormat.SelectedIndex = EnumFormat.JORecap Or Me.CboFormat.SelectedIndex = EnumFormat.JODetail) And _
                    ((Me.ChkCodeFG.Checked = False And Me.ChkDescFG.Checked = False) Or (Me.ChkCodeRM.Checked = False And Me.ChkDescRM.Checked = False)) Then
                MsgBox("Select at least Show Code or Show Description for Finish Goods and Raw Material", MsgBoxStyle.Information, "Information")
                Return
            End If
            Dim _start As Date = Now
            If Me.CboFormat.SelectedIndex = EnumFormat.ProductRecap Or Me.CboFormat.SelectedIndex = EnumFormat.ProductDetail Or _
                Me.CboFormat.SelectedIndex = EnumFormat.JORecap Or Me.CboFormat.SelectedIndex = EnumFormat.JODetail Then
                LoadReport_ProductAndJO()
            ElseIf Me.CboFormat.SelectedIndex = EnumFormat.WORecap Or Me.CboFormat.SelectedIndex = EnumFormat.WODetail Then
                LoadReport_WO()
            End If
            Console.WriteLine((Now - _start).TotalSeconds)
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
                                filter = "[Product JO Primary] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Product JO Primary] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
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

    Private Sub CboFormat_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboFormat.SelectedIndexChanged
        Try
            If CboFormat.SelectedIndex <> EnumFormat.WORecap And CboFormat.SelectedIndex <> EnumFormat.WODetail Then
                Me.chkIncludeWIP.Enabled = False
                Me.chkIncludeWIP.Checked = False
            Else
                Me.chkIncludeWIP.Enabled = True
            End If
            If Me.CboFormat.SelectedIndex = EnumFormat.JORecap Or Me.CboFormat.SelectedIndex = EnumFormat.JODetail Then
                Me.GroupBox6.Enabled = True
                Me.GroupBox7.Enabled = True
            Else
                Me.GroupBox6.Enabled = False
                Me.GroupBox7.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    Sub LoadReport_ProductAndJO()
        Dim DtaLog As New DseLaporanProd_RawMaterialUsageTableAdapters.DataTable2TableAdapter
        Dim TblLog As New DseLaporanProd_RawMaterialUsage.DataTable2DataTable
        Dim RowLog As DseLaporanProd_RawMaterialUsage.DataTable2Row
        Dim DtaMO As New DseLaporanProd_RawMaterialUsageTableAdapters.DataTable3TableAdapter
        Dim TblMO As New DseLaporanProd_RawMaterialUsage.DataTable3DataTable
        Dim RowMO As DseLaporanProd_RawMaterialUsage.DataTable3Row
        Dim RowMO2 As DseLaporanProd_RawMaterialUsage.DataTable3Row
        Dim PrimaryMO_Min As Double = 9999999999
        Dim PrimaryMO_Max As Double = 0
        Dim IndexLog As Integer = 0
        Dim IndexMO As Integer = 0
        Dim Report As New CrpLaporanProd_RawMaterialUsage_Product
        Dim Report_JO As New CrpLaporanProd_RawMaterialUsage_JO
        Dim DS As New DseLaporanProd_RawMaterialUsage
        Dim TblPrint As DseLaporanProd_RawMaterialUsage.PrintDataTable = DS.Print
        Dim RowPrint As DseLaporanProd_RawMaterialUsage.PrintRow
        Dim LogPeriod As String = ""
        Dim MOPeriod As String = ""
        Dim ShowLogDetail As Boolean = False
        Try
            ''init
            Me.Cursor = Cursors.WaitCursor
            DtaLog.Connection.ConnectionString = Core_Module.GetConstring
            DtaMO.Connection.ConnectionString = Core_Module.GetConstring

            ''load log data
            If Me.ChkLogDateFrom.Checked Then
                DtaLog.FillBy_LogDate(TblLog, Me.dtpFromBundlingDate.Value.Date, Me.dtpToBundlingDate.Value.Date)
            Else
                DtaLog.FillBy_LogDateUntil(TblLog, Me.dtpToBundlingDate.Value.Date)
            End If
            If TblLog.Count > 0 Then
                RowLog = TblLog(0)
                PrimaryMO_Min = RowLog.MO_Primary
                RowLog = TblLog(TblLog.Count - 1)
                PrimaryMO_Max = RowLog.MO_Primary
            End If

            ''load MO data
            If Me.ChkMODate.Checked Then
                DtaMO.FillBy_MO_PrimaryMinMaxAndDate(TblMO, Me.dtpFromMODate.Value.Date, Me.dtpToMODate.Value.Date, PrimaryMO_Min, PrimaryMO_Max)
            Else
                DtaMO.FillBy_MO_PrimaryMinMax(TblMO, PrimaryMO_Min, PrimaryMO_Max)
            End If

            ''delete exlude MO (based on user filter)
            Me.FilterByMO(TblMO)

            ''delete exclude MO (based on product filter)
            Me.FilterByProduct(TblMO)

            ''delete exclude MO (based on log data)
            TblLog.DefaultView.RowFilter = ""
            TblLog.DefaultView.Sort = "[MO Primary]"
            TblMO.DefaultView.RowFilter = ""
            TblMO.DefaultView.Sort = "[MO Primary]"
            While IndexMO < TblMO.DefaultView.Count
                While IndexLog < TblLog.DefaultView.Count
                    RowMO = TblMO.DefaultView(IndexMO).Row
                    RowLog = TblLog.DefaultView(IndexLog).Row
                    If RowLog.MO_Primary < RowMO.MO_Primary Then
                        IndexLog += 1
                    ElseIf RowLog.MO_Primary = RowMO.MO_Primary Then
                        IndexMO += 1
                        Exit While
                    ElseIf RowLog.MO_Primary > RowMO.MO_Primary Then
                        TblMO.Rows.Remove(RowMO)
                        Exit While
                    End If
                End While
            End While

            ''calculate used qty
            IndexMO = 0
            IndexLog = 0
            TblMO.DefaultView.RowFilter = ""
            TblMO.DefaultView.Sort = "[MO Primary],[Product JO Primary], [Product RM Primary]"
            TblLog.DefaultView.RowFilter = ""
            TblLog.DefaultView.Sort = "[MO Primary],[Product JO Primary], [Product RM Primary]"
            While IndexMO < TblMO.DefaultView.Count
                RowMO = TblMO.DefaultView(IndexMO).Row
                While IndexLog < TblLog.DefaultView.Count
                    RowLog = TblLog.DefaultView(IndexLog).Row
                    If RowLog.MO_Primary < RowMO.MO_Primary Then
                        IndexLog += 1
                    ElseIf RowLog.MO_Primary = RowMO.MO_Primary Then
                        If RowLog.Product_JO_Primary < RowMO.Product_JO_Primary Then
                            IndexLog += 1
                        ElseIf RowLog.Product_JO_Primary = RowMO.Product_JO_Primary Then
                            If RowLog.Product_RM_Primary < RowMO.Product_RM_Primary Then
                                IndexLog += 1
                            ElseIf RowLog.Product_RM_Primary = RowMO.Product_RM_Primary Then
                                RowMO.Used_Qty += RowLog.Log_Qty
                                IndexLog += 1
                                RowLog.Include = True
                            ElseIf RowLog.Product_RM_Primary > RowMO.Product_RM_Primary Then
                                Exit While
                            End If
                        ElseIf RowLog.Product_JO_Primary > RowMO.Product_JO_Primary Then
                            Exit While
                        End If
                    ElseIf RowLog.MO_Primary > RowMO.MO_Primary Then
                        Exit While
                    End If
                End While
                IndexMO += 1
            End While

            ''remove 0 used qty
            If Me.ChkShowZeroUsed.Checked = False Then
                For i As Integer = TblMO.Count - 1 To 0 Step -1
                    RowMO = TblMO.Rows(i)
                    If RowMO.Used_Qty <= 0 Then
                        TblMO.Rows.RemoveAt(i)
                    End If
                Next
            End If

            ''group data 
            IndexMO = 0
            Dim LastJO_Identifier As String = ""
            Dim LastProductJO_Primary As Double = -1
            Dim LastProductRM_Primary As Double = -1
            TblMO.DefaultView.RowFilter = ""
            If Me.CboFormat.SelectedIndex = EnumFormat.ProductRecap Or Me.CboFormat.SelectedIndex = EnumFormat.ProductDetail Then
                TblMO.DefaultView.Sort = "[Product JO Primary], [Product RM Primary]"
            ElseIf Me.CboFormat.SelectedIndex = EnumFormat.JORecap Or Me.CboFormat.SelectedIndex = EnumFormat.JODetail Then
                TblMO.DefaultView.Sort = "[Identifier JO], [Product RM Primary]"
            End If
            While IndexMO < TblMO.DefaultView.Count
                RowMO = TblMO.DefaultView(IndexMO).Row
                If ((Me.CboFormat.SelectedIndex = EnumFormat.ProductRecap Or Me.CboFormat.SelectedIndex = EnumFormat.ProductDetail) And _
                RowMO.Product_JO_Primary = LastProductJO_Primary And RowMO.Product_RM_Primary = LastProductRM_Primary) Or _
                        ((Me.CboFormat.SelectedIndex = EnumFormat.JORecap Or Me.CboFormat.SelectedIndex = EnumFormat.JODetail) And _
                        RowMO.Identifier_JO = LastJO_Identifier And RowMO.Product_RM_Primary = LastProductRM_Primary) Then
                    RowMO2 = TblMO.DefaultView(IndexMO - 1).Row
                    RowMO2.Required_Qty += RowMO.Required_Qty
                    RowMO2.Additional_Qty += RowMO.Additional_Qty
                    RowMO2.Replace_Qty += RowMO.Replace_Qty
                    RowMO2.Used_Qty += RowMO.Used_Qty
                    LastJO_Identifier = RowMO.Identifier_JO
                    LastProductJO_Primary = RowMO.Product_JO_Primary
                    LastProductRM_Primary = RowMO.Product_RM_Primary
                    TblMO.Rows.Remove(RowMO)
                Else
                    LastJO_Identifier = RowMO.Identifier_JO
                    LastProductJO_Primary = RowMO.Product_JO_Primary
                    LastProductRM_Primary = RowMO.Product_RM_Primary
                    IndexMO += 1
                End If
            End While

            ''calculate balance
            For Each RowMO In TblMO
                If Me.ChkLogDateFrom.Checked = False Then
                    RowMO.Balance_Qty = RowMO.Required_Qty + RowMO.Additional_Qty - RowMO.Used_Qty - RowMO.Replace_Qty
                Else
                    RowMO.Balance_Qty = -1
                End If
            Next

            ''add log detail
            If Me.CboFormat.SelectedIndex = EnumFormat.ProductDetail Then
                ShowLogDetail = True
                TblLog.DefaultView.RowFilter = "[Include] = True"
                TblLog.DefaultView.Sort = "[Product JO Primary], [Product RM Primary]"
                TblMO.DefaultView.RowFilter = ""
                TblMO.DefaultView.Sort = "[Product JO Primary], [Product RM Primary]"
                IndexLog = 0
                IndexMO = 0
                While IndexLog < TblLog.DefaultView.Count
                    RowLog = TblLog.DefaultView(IndexLog).Row
                    While IndexMO < TblMO.DefaultView.Count
                        RowMO = TblMO.DefaultView(IndexMO).Row
                        If RowMO.Product_JO_Primary < RowLog.Product_JO_Primary Then
                            IndexMO += 1
                        ElseIf RowMO.Product_JO_Primary = RowLog.Product_JO_Primary Then
                            If RowMO.Product_RM_Primary < RowLog.Product_RM_Primary Then
                                IndexMO += 1
                            ElseIf RowMO.Product_RM_Primary = RowLog.Product_RM_Primary Then
                                IndexLog += 1
                                RowPrint = DS.Print.NewPrintRow
                                With RowPrint
                                    .Position_I = RowMO.Product_JO_Code
                                    .Position_II = RowMO.Product_JO_Description
                                    .RM_Code = RowMO.Product_RM_Code
                                    .RM_Desc = RowMO.Product_RM_Description
                                    .Unit = RowMO.Unit_Code_RM
                                    .Required_Qty = RowMO.Required_Qty
                                    .Additional_Qty = RowMO.Additional_Qty
                                    .Replaced_Qty = RowMO.Replace_Qty
                                    .Used_Qty = RowMO.Used_Qty
                                    .Balance_Qty = RowMO.Balance_Qty
                                    ._Decimal = RowMO._Product_RM__Decimal_
                                    .Bundling_No = RowLog.Log_No
                                    .Bundling_Date = RowLog.Log_Date
                                    .Bundling_Qty = RowLog.Log_Qty
                                End With
                                DS.Print.Rows.Add(RowPrint)
                                Exit While
                            ElseIf RowMO.Product_RM_Primary > RowLog.Product_RM_Primary Then
                                IndexLog += 1
                                Exit While
                            End If
                        ElseIf RowMO.Product_JO_Primary > RowLog.Product_JO_Primary Then
                            IndexLog += 1
                            Exit While
                        End If
                    End While
                End While
            ElseIf Me.CboFormat.SelectedIndex = EnumFormat.JODetail Then
                ShowLogDetail = True
                TblLog.DefaultView.RowFilter = "[Include] = True"
                TblLog.DefaultView.Sort = "[MO Primary], [Product JO Primary], [Product RM Primary]"
                TblMO.DefaultView.RowFilter = ""
                TblMO.DefaultView.Sort = "[MO Primary], [Product JO Primary], [Product RM Primary]"
                IndexLog = 0
                IndexMO = 0
                While IndexLog < TblLog.DefaultView.Count
                    RowLog = TblLog.DefaultView(IndexLog).Row
                    While IndexMO < TblMO.DefaultView.Count
                        RowMO = TblMO.DefaultView(IndexMO).Row
                        If RowMO.MO_Primary < RowLog.MO_Primary Then
                            IndexMO += 1
                        ElseIf RowMO.MO_Primary = RowLog.MO_Primary Then
                            If RowMO.Product_JO_Primary < RowLog.Product_JO_Primary Then
                                IndexMO += 1
                            ElseIf RowMO.Product_JO_Primary = RowLog.Product_JO_Primary Then
                                If RowMO.Product_RM_Primary < RowLog.Product_RM_Primary Then
                                    IndexMO += 1
                                ElseIf RowMO.Product_RM_Primary = RowLog.Product_RM_Primary Then
                                    IndexLog += 1
                                    RowPrint = DS.Print.NewPrintRow
                                    With RowPrint
                                        .JO_Number = RowMO.JO_Number
                                        .Position_I = RowMO.Product_JO_Code
                                        .Position_II = RowMO.Product_JO_Description
                                        .RM_Code = RowMO.Product_RM_Code
                                        .RM_Desc = RowMO.Product_RM_Description
                                        .Unit = RowMO.Unit_Code_RM
                                        .Required_Qty = RowMO.Required_Qty
                                        .Additional_Qty = RowMO.Additional_Qty
                                        .Replaced_Qty = RowMO.Replace_Qty
                                        .Used_Qty = RowMO.Used_Qty
                                        .Balance_Qty = RowMO.Balance_Qty
                                        ._Decimal = RowMO._Product_RM__Decimal_
                                        .Bundling_No = RowLog.Log_No
                                        .Bundling_Date = RowLog.Log_Date
                                        .Bundling_Qty = RowLog.Log_Qty
                                    End With
                                    DS.Print.Rows.Add(RowPrint)
                                    Exit While
                                ElseIf RowMO.Product_RM_Primary > RowLog.Product_RM_Primary Then
                                    IndexLog += 1
                                    Exit While
                                End If
                            ElseIf RowMO.Product_JO_Primary > RowLog.Product_JO_Primary Then
                                IndexLog += 1
                                Exit While
                            End If
                        ElseIf RowMO.MO_Primary > RowLog.MO_Primary Then
                            IndexLog += 1
                            Exit While
                        End If
                    End While
                End While
            End If

            ''copy to print source
            IndexMO = 0
            TblMO.DefaultView.RowFilter = ""
            TblMO.DefaultView.Sort = "[JO Number], [Product JO Code], [Product RM Code]"
            While IndexMO < TblMO.DefaultView.Count
                RowMO = TblMO.DefaultView(IndexMO).Row
                RowPrint = DS.Print.NewPrintRow
                With RowPrint
                    .JO_Number = RowMO.JO_Number
                    .Position_I = RowMO.Product_JO_Code
                    .Position_II = RowMO.Product_JO_Description
                    .RM_Code = RowMO.Product_RM_Code
                    .RM_Desc = RowMO.Product_RM_Description
                    .Unit = RowMO.Unit_Code_RM
                    .Required_Qty = RowMO.Required_Qty
                    .Additional_Qty = RowMO.Additional_Qty
                    .Replaced_Qty = RowMO.Replace_Qty
                    .Used_Qty = RowMO.Used_Qty
                    .Balance_Qty = RowMO.Balance_Qty
                    ._Decimal = RowMO._Product_RM__Decimal_
                End With
                DS.Print.Rows.Add(RowPrint)
                IndexMO += 1
            End While

            ''sort data
            If Me.CboFormat.SelectedIndex = EnumFormat.ProductRecap Or Me.CboFormat.SelectedIndex = EnumFormat.ProductDetail Then
                Me.SortTBL(DS.Print, "[Position I], [RM Code], [Bundling Date], [Bundling No]")
            ElseIf Me.CboFormat.SelectedIndex = EnumFormat.JORecap Or Me.CboFormat.SelectedIndex = EnumFormat.JODetail Then
                If Me.ChkCodeFG.Checked And Me.ChkCodeRM.Checked Then
                    Me.SortTBL(DS.Print, "[JO Number], [Position I], [RM Code], [Bundling Date], [Bundling No]")
                ElseIf Me.ChkCodeFG.Checked And Me.ChkCodeRM.Checked = False Then
                    Me.SortTBL(DS.Print, "[JO Number], [Position I], [RM Desc], [Bundling Date], [Bundling No]")
                ElseIf Me.ChkCodeFG.Checked = False And Me.ChkCodeRM.Checked Then
                    Me.SortTBL(DS.Print, "[JO Number], [Position II], [RM Code], [Bundling Date], [Bundling No]")
                ElseIf Me.ChkCodeFG.Checked = False And Me.ChkCodeRM.Checked = False Then
                    Me.SortTBL(DS.Print, "[JO Number], [Position II], [RM Desc], [Bundling Date], [Bundling No]")
                End If
            End If

            ''set report params
            If Me.ChkLogDateFrom.Checked = False Then
                LogPeriod = "Begining until " & Format(Me.dtpToBundlingDate.Value.Date, "dd MMM yyyy")
            Else
                LogPeriod = Format(Me.dtpFromBundlingDate.Value.Date, "dd MMM yyyy") & " to " & Format(Me.dtpToBundlingDate.Value.Date, "dd MMM yyyy")
            End If
            If Me.ChkMODate.Checked = False Then
                MOPeriod = "All period"
            Else
                MOPeriod = Format(Me.dtpFromMODate.Value.Date, "dd MMM yyyy") & " to " & Format(Me.dtpToMODate.Value.Date, "dd MMM yyyy")
            End If

            ''show report
            If Me.CboFormat.SelectedIndex = EnumFormat.ProductRecap Or Me.CboFormat.SelectedIndex = EnumFormat.ProductDetail Then
                Report.SetDataSource(DS)
                Report.SetParameterValue(Report.Parameter_Bundling_Period.ParameterFieldName, LogPeriod)
                Report.SetParameterValue(Report.Parameter_MO_Period.ParameterFieldName, MOPeriod)
                Report.SetParameterValue(Report.Parameter_vDetailFormat.ParameterFieldName, ShowLogDetail)
                Me.crv.ReportSource = Report
            ElseIf Me.CboFormat.SelectedIndex = EnumFormat.JORecap Or Me.CboFormat.SelectedIndex = EnumFormat.JODetail Then
                ''modify code-desc FG-RM
                For Each RowPrint In DS.Print
                    If Me.ChkCodeFG.Checked = False And Me.ChkDescFG.Checked Then
                        RowPrint.Position_I = RowPrint.Position_II
                    ElseIf Me.ChkCodeFG.Checked And Me.ChkDescFG.Checked Then
                        RowPrint.Position_I = RowPrint.Position_I & vbNewLine & RowPrint.Position_II
                    End If
                    If Me.ChkCodeRM.Checked = False And Me.ChkDescRM.Checked Then
                        RowPrint.RM_Code = RowPrint.RM_Desc
                    ElseIf Me.ChkCodeRM.Checked And Me.ChkDescRM.Checked Then
                        RowPrint.RM_Code = RowPrint.RM_Code & vbNewLine & RowPrint.RM_Desc
                    End If
                Next
                Report_JO.SetDataSource(DS)
                Report_JO.SetParameterValue(Report_JO.Parameter_Bundling_Period.ParameterFieldName, LogPeriod)
                Report_JO.SetParameterValue(Report_JO.Parameter_MO_Period.ParameterFieldName, MOPeriod)
                Report_JO.SetParameterValue(Report_JO.Parameter_vDetailFormat.ParameterFieldName, ShowLogDetail)
                Me.crv.ReportSource = Report_JO
            End If
            Me.crv.Zoom(85)
            Me.crv.Show()

            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Console.WriteLine(IndexMO & " " & IndexLog)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub LoadReport_WO()
        Dim DtaLog As New DseLaporanProd_RawMaterialUsageTableAdapters.DataTable6TableAdapter
        Dim TblLog As New DseLaporanProd_RawMaterialUsage.DataTable6DataTable
        Dim RowLog As DseLaporanProd_RawMaterialUsage.DataTable6Row
        Dim DtaMO As New DseLaporanProd_RawMaterialUsageTableAdapters.DataTable5TableAdapter
        Dim TblMO As New DseLaporanProd_RawMaterialUsage.DataTable5DataTable
        Dim RowMO As DseLaporanProd_RawMaterialUsage.DataTable5Row
        Dim RowMO2 As DseLaporanProd_RawMaterialUsage.DataTable5Row
        Dim PrimaryMO_Min As Double = 9999999999
        Dim PrimaryMO_Max As Double = 0
        Dim IndexLog As Integer = 0
        Dim IndexMO As Integer = 0
        Dim Report As New CrpLaporanProd_RawMaterialUsage_WO
        Dim DS As New DseLaporanProd_RawMaterialUsage
        Dim TblPrint As DseLaporanProd_RawMaterialUsage.PrintDataTable = DS.Print
        Dim RowPrint As DseLaporanProd_RawMaterialUsage.PrintRow
        Dim LogPeriod As String = ""
        Dim MOPeriod As String = ""
        Dim ShowLogDetail As Boolean = False
        Dim IsWIP As Integer = 0
        Dim IsWIP2 As Integer = 0
        Try
            ''init
            Me.Cursor = Cursors.WaitCursor
            If Me.chkIncludeWIP.Checked Then
                IsWIP2 = 1
            End If
            DtaLog.Connection.ConnectionString = Core_Module.GetConstring
            DtaMO.Connection.ConnectionString = Core_Module.GetConstring

            ''load log data
            If Me.ChkLogDateFrom.Checked Then
                DtaLog.FillBy_LogDate(TblLog, Me.dtpFromBundlingDate.Value.Date, Me.dtpToBundlingDate.Value.Date)
            Else
                DtaLog.FillBy_LogDateUntil(TblLog, Me.dtpToBundlingDate.Value.Date)
            End If
            If TblLog.Count > 0 Then
                RowLog = TblLog(0)
                PrimaryMO_Min = RowLog.MO_Primary
                RowLog = TblLog(TblLog.Count - 1)
                PrimaryMO_Max = RowLog.MO_Primary
            End If

            ''load MO data
            If Me.ChkMODate.Checked Then
                DtaMO.FillBy_MO_PrimaryMinMaxAndDate(TblMO, PrimaryMO_Min, IsWIP, IsWIP2, PrimaryMO_Max, Me.dtpFromMODate.Value.Date, Me.dtpToMODate.Value.Date)
            Else
                DtaMO.FillBy_MO_PrimaryMinMax(TblMO, PrimaryMO_Min, IsWIP, IsWIP2, PrimaryMO_Max)
            End If

            ''delete exlude MO (based on user filter)
            Me.FilterByMO(TblMO)

            ''delete exclude MO (based on product filter)
            Me.FilterByProduct(TblMO)

            ''delete exclude MO (based on log data)
            TblLog.DefaultView.RowFilter = ""
            TblLog.DefaultView.Sort = "[MO Primary]"
            TblMO.DefaultView.RowFilter = ""
            TblMO.DefaultView.Sort = "[MO Primary]"
            While IndexMO < TblMO.DefaultView.Count
                While IndexLog < TblLog.DefaultView.Count
                    RowMO = TblMO.DefaultView(IndexMO).Row
                    RowLog = TblLog.DefaultView(IndexLog).Row
                    If RowLog.MO_Primary < RowMO.MO_Primary Then
                        IndexLog += 1
                    ElseIf RowLog.MO_Primary = RowMO.MO_Primary Then
                        IndexMO += 1
                        Exit While
                    ElseIf RowLog.MO_Primary > RowMO.MO_Primary Then
                        TblMO.Rows.Remove(RowMO)
                        Exit While
                    End If
                End While
            End While

            ''calculate used qty
            IndexMO = 0
            IndexLog = 0
            TblMO.DefaultView.RowFilter = ""
            TblMO.DefaultView.Sort = "[MO Primary],[Product JO Primary], [Product WO Primary], [Product RM Primary]"
            TblLog.DefaultView.RowFilter = ""
            TblLog.DefaultView.Sort = "[MO Primary],[Product JO Primary], [Product WO Primary], [Product RM Primary]"
            While IndexMO < TblMO.DefaultView.Count
                RowMO = TblMO.DefaultView(IndexMO).Row
                While IndexLog < TblLog.DefaultView.Count
                    RowLog = TblLog.DefaultView(IndexLog).Row
                    If RowLog.MO_Primary < RowMO.MO_Primary Then
                        IndexLog += 1
                    ElseIf RowLog.MO_Primary = RowMO.MO_Primary Then
                        If RowLog.Product_JO_Primary < RowMO.Product_JO_Primary Then
                            IndexLog += 1
                        ElseIf RowLog.Product_JO_Primary = RowMO.Product_JO_Primary Then
                            If RowLog.Product_WO_Primary < RowMO.Product_WO_Primary Then
                                IndexLog += 1
                            ElseIf RowLog.Product_WO_Primary = RowMO.Product_WO_Primary Then
                                If RowLog.Product_RM_Primary < RowMO.Product_RM_Primary Then
                                    IndexLog += 1
                                ElseIf RowLog.Product_RM_Primary = RowMO.Product_RM_Primary Then
                                    RowMO.Used_Qty += RowLog.Log_Qty
                                    IndexLog += 1
                                    RowLog.Include = True
                                ElseIf RowLog.Product_RM_Primary > RowMO.Product_RM_Primary Then
                                    Exit While
                                End If
                            ElseIf RowLog.Product_WO_Primary > RowMO.Product_WO_Primary Then
                                Exit While
                            End If
                        ElseIf RowLog.Product_JO_Primary > RowMO.Product_JO_Primary Then
                            Exit While
                        End If
                    ElseIf RowLog.MO_Primary > RowMO.MO_Primary Then
                        Exit While
                    End If
                End While
                IndexMO += 1
            End While

            ''remove 0 used qty
            If Me.ChkShowZeroUsed.Checked = False Then
                For i As Integer = TblMO.Count - 1 To 0 Step -1
                    RowMO = TblMO.Rows(i)
                    If RowMO.Used_Qty <= 0 Then
                        TblMO.Rows.RemoveAt(i)
                    End If
                Next
            End If

            ''group data 
            IndexMO = 0
            Dim LastMO_Primary As Double = -1
            Dim LastProductJO_Primary As Double = -1
            Dim LastProductWO_Primary As Double = -1
            Dim LastProductRM_Primary As Double = -1
            TblMO.DefaultView.RowFilter = ""
            TblMO.DefaultView.Sort = "[MO Primary], [Product JO Primary], [Product WO Primary], [Product RM Primary]"
            While IndexMO < TblMO.DefaultView.Count
                RowMO = TblMO.DefaultView(IndexMO).Row
                If RowMO.MO_Primary = LastMO_Primary And RowMO.Product_JO_Primary = LastProductJO_Primary And _
                        RowMO.Product_WO_Primary = LastProductWO_Primary And RowMO.Product_RM_Primary = LastProductRM_Primary Then
                    RowMO2 = TblMO.DefaultView(IndexMO - 1).Row
                    RowMO2.Required_Qty += RowMO.Required_Qty
                    RowMO2.Additional_Qty += RowMO.Additional_Qty
                    RowMO2.Replace_Qty += RowMO.Replace_Qty
                    RowMO2.Used_Qty += RowMO.Used_Qty
                    LastMO_Primary = RowMO.MO_Primary
                    LastProductJO_Primary = RowMO.Product_JO_Primary
                    LastProductWO_Primary = RowMO.Product_WO_Primary
                    LastProductRM_Primary = RowMO.Product_RM_Primary
                    TblMO.Rows.Remove(RowMO)
                Else
                    LastMO_Primary = RowMO.MO_Primary
                    LastProductJO_Primary = RowMO.Product_JO_Primary
                    LastProductWO_Primary = RowMO.Product_WO_Primary
                    LastProductRM_Primary = RowMO.Product_RM_Primary
                    IndexMO += 1
                End If
            End While

            ''calculate balance
            For Each RowMO In TblMO
                If Me.ChkLogDateFrom.Checked = False Then
                    RowMO.Balance_Qty = RowMO.Required_Qty + RowMO.Additional_Qty - RowMO.Used_Qty - RowMO.Replace_Qty
                Else
                    RowMO.Balance_Qty = -1
                End If
            Next

            ''add log detail
            If Me.CboFormat.SelectedIndex = EnumFormat.WODetail Then
                ShowLogDetail = True
                TblLog.DefaultView.RowFilter = "[Include] = True"
                TblLog.DefaultView.Sort = "[MO Primary], [Product JO Primary], [Product WO Primary], [Product RM Primary]"
                TblMO.DefaultView.RowFilter = ""
                TblMO.DefaultView.Sort = "[MO Primary], [Product JO Primary], [Product WO Primary], [Product RM Primary]"
                IndexLog = 0
                IndexMO = 0
                While IndexLog < TblLog.DefaultView.Count
                    RowLog = TblLog.DefaultView(IndexLog).Row
                    While IndexMO < TblMO.DefaultView.Count
                        RowMO = TblMO.DefaultView(IndexMO).Row
                        If RowMO.MO_Primary < RowLog.MO_Primary Then
                            IndexMO += 1
                        ElseIf RowMO.MO_Primary = RowLog.MO_Primary Then
                            If RowMO.Product_JO_Primary < RowLog.Product_JO_Primary Then
                                IndexMO += 1
                            ElseIf RowMO.Product_JO_Primary = RowLog.Product_JO_Primary Then
                                If RowMO.Product_WO_Primary < RowLog.Product_WO_Primary Then
                                    IndexMO += 1
                                ElseIf RowMO.Product_WO_Primary = RowLog.Product_WO_Primary Then
                                    If RowMO.Product_RM_Primary < RowLog.Product_RM_Primary Then
                                        IndexMO += 1
                                    ElseIf RowMO.Product_RM_Primary = RowLog.Product_RM_Primary Then
                                        IndexLog += 1
                                        RowPrint = DS.Print.NewPrintRow
                                        With RowPrint
                                            .WO_Date = RowMO.WO_Date
                                            .WO_Number = RowMO.WO_Number
                                            .RM_Code = RowMO.Product_RM_Code
                                            .RM_Desc = RowMO.Product_RM_Description
                                            .Unit = RowMO.Product_RM_Unit_Code
                                            .Required_Qty = RowMO.Required_Qty
                                            .Additional_Qty = RowMO.Additional_Qty
                                            .Replaced_Qty = RowMO.Replace_Qty
                                            .Used_Qty = RowMO.Used_Qty
                                            .Balance_Qty = RowMO.Balance_Qty
                                            ._Decimal = RowMO.Product_RM_Precision
                                            .Bundling_No = RowLog.Log_No
                                            .Bundling_Date = RowLog.Log_Date
                                            .Bundling_Qty = RowLog.Log_Qty
                                        End With
                                        DS.Print.Rows.Add(RowPrint)
                                        Exit While
                                    ElseIf RowMO.Product_RM_Primary > RowLog.Product_RM_Primary Then
                                        IndexLog += 1
                                        Exit While
                                    End If
                                ElseIf RowMO.Product_WO_Primary > RowLog.Product_WO_Primary Then
                                    IndexLog += 1
                                    Exit While
                                End If
                            ElseIf RowMO.Product_JO_Primary > RowLog.Product_JO_Primary Then
                                IndexLog += 1
                                Exit While
                            End If
                        ElseIf RowMO.MO_Primary > RowLog.MO_Primary Then
                            IndexLog += 1
                            Exit While
                        End If
                    End While
                End While
            End If

            ''copy to print source
            IndexMO = 0
            TblMO.DefaultView.RowFilter = ""
            TblMO.DefaultView.Sort = "[WO Date], [WO Number], [Product RM Code]"
            While IndexMO < TblMO.DefaultView.Count
                RowMO = TblMO.DefaultView(IndexMO).Row
                RowPrint = DS.Print.NewPrintRow
                With RowPrint
                    .WO_Date = RowMO.WO_Date
                    .WO_Number = RowMO.WO_Number
                    .RM_Code = RowMO.Product_RM_Code
                    .RM_Desc = RowMO.Product_RM_Description
                    .Unit = RowMO.Product_RM_Unit_Code
                    .Required_Qty = RowMO.Required_Qty
                    .Additional_Qty = RowMO.Additional_Qty
                    .Replaced_Qty = RowMO.Replace_Qty
                    .Used_Qty = RowMO.Used_Qty
                    .Balance_Qty = RowMO.Balance_Qty
                    ._Decimal = RowMO.Product_RM_Precision
                End With
                DS.Print.Rows.Add(RowPrint)
                IndexMO += 1
            End While

            ''sort data
            Me.SortTBL(DS.Print, "[WO Date], [WO Number], [RM Code], [Bundling Date], [Bundling No]")

            ''set report params
            If Me.ChkLogDateFrom.Checked = False Then
                LogPeriod = "Begining until " & Format(Me.dtpToBundlingDate.Value.Date, "dd MMM yyyy")
            Else
                LogPeriod = Format(Me.dtpFromBundlingDate.Value.Date, "dd MMM yyyy") & " to " & Format(Me.dtpToBundlingDate.Value.Date, "dd MMM yyyy")
            End If
            If Me.ChkMODate.Checked = False Then
                MOPeriod = "All period"
            Else
                MOPeriod = Format(Me.dtpFromMODate.Value.Date, "dd MMM yyyy") & " to " & Format(Me.dtpToMODate.Value.Date, "dd MMM yyyy")
            End If

            ''show report
            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_Bundling_Period.ParameterFieldName, LogPeriod)
            Report.SetParameterValue(Report.Parameter_MO_Period.ParameterFieldName, MOPeriod)
            Report.SetParameterValue(Report.Parameter_vDetailFormat.ParameterFieldName, ShowLogDetail)
            Me.crv.ReportSource = Report
            Me.crv.Zoom(85)
            Me.crv.Show()

            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Console.WriteLine(IndexMO & " " & IndexLog)
        Finally
            Me.Cursor = Cursors.Default
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
                            filter = "[MO Primary] = " & .Cells(Me.colMO_Primary.Index).Value.ToString
                        Else
                            filter += " OR [MO Primary] = " & .Cells(Me.colMO_Primary.Index).Value.ToString
                        End If
                    End If
                End With
            Next
            If filter = "" Then
                Exit Sub
            End If
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
End Class