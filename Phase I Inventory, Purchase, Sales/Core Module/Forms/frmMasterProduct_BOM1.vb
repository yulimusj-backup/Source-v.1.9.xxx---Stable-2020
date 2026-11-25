Option Explicit On
Option Strict Off

Public Class frmMasterProduct_BOM1

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim objProduct As New Core_Module.clsMasterProduct
    Dim objProcess As New Core_Module.ClsMasterProses
    Dim frmProduct As New Core_Module.frmMasterProduct
    Dim objCD As New Core_Module.ClsCekData
    Dim buffer As Object
    Dim rowToDelete As Integer
    Dim rowDgv As DataGridViewRow
    Dim DV As Data.DataRowView

#Region "Tab2"
    Private Sub frmMasterProduct_BOM1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvItem1.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem2.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvProcess.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1ProcessCode.GotFocus, txt1ProcessDesc.GotFocus, txt1RMCode.GotFocus, txt2CodeNo.GotFocus
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            txt.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1ProcessCode.LostFocus, txt1ProcessDesc.LostFocus, txt1RMCode.LostFocus, txt2CodeNo.LostFocus
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            txt.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn2FindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2FindCodeNo.Click
        Dim frmFindCodeNo As New Core_Module.frmCariMasterProduct
        Try
            frmFindCodeNo.txtCari.Text = Me.txt2CodeNo.Text
            frmFindCodeNo.TipePencarian = frmCariMasterProduct.enumTipePencarian.Goods_Only
            frmFindCodeNo.TipeProduct = clsMasterProduct.EnumTipeProduct.Goods
            frmFindCodeNo.ShowDialog()
            If frmFindCodeNo.PrimaryProduct > -1 Then
                Me.txt2CodeNo.Text = objProduct.GetProductCode(frmFindCodeNo.PrimaryProduct)
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
        Try
            If e.ColumnIndex = Me.colProcess_Code.Index Then
                If IsNothing(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If String.IsNullOrEmpty(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim) = False Then
                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper.Trim
                        If Me.buffer = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value Then
                            Exit Sub
                        End If
                        If Me.objProcess.ApakahSudahAda(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If Me.IsExistProcess(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex) = True Then
                                MsgBox("Process code: " & Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                Me.BackToBufferOrDelete(e)
                                Exit Sub
                            End If
                            Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value = Me.objProcess.GetPrimary(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvProcess.Item(Me.colProcess_Desc.Index, e.RowIndex).Value = Me.objProcess.GetDesc(Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value)
                        Else
                            Dim frmFindProcess As New Core_Module.FrmMasterProsesCari
                            frmFindProcess.txtCari.Text = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value
                            'frmFindProcess.mStatusAktif = True
                            'frmFindProcess.mSearchAll = False
                            frmFindProcess.ShowDialog()
                            If frmFindProcess.mPrimary > -1 Then
                                If Me.IsExistProcess(Me.objProcess.GetCode(frmFindProcess.mPrimary), e.RowIndex) = True Then
                                    MsgBox("Process code: " & Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                    Me.BackToBufferOrDelete(e)
                                    Exit Sub
                                End If
                                Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value = frmFindProcess.mPrimary
                                Me.dgvProcess.Item(Me.colProcess_Code.Index, e.RowIndex).Value = Me.objProcess.GetCode(frmFindProcess.mPrimary)
                                Me.dgvProcess.Item(Me.colProcess_Desc.Index, e.RowIndex).Value = Me.objProcess.GetDesc(frmFindProcess.mPrimary)
                            Else
                                Me.BackToBufferOrDelete(e)
                                Exit Sub
                            End If
                        End If
                    Else
                        Me.BackToBufferOrDelete(e)
                        Exit Sub
                    End If
                Else
                    Me.BackToBufferOrDelete(e)
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub BackToBufferOrDelete(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If IsNothing(Me.buffer) = False Then
                If String.IsNullOrEmpty(Me.buffer.ToString.Trim) = False Then
                    Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    Me.rowToDelete = e.RowIndex
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            Else
                Me.rowToDelete = e.RowIndex
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.rowToDelete < Me.dgvProcess.RowCount - 1 Then
                Me.dgvProcess.Rows.RemoveAt(Me.rowToDelete)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProcess_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvProcess.DataError

    End Sub

    Private Sub dgvProcess_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvProcess.RowsAdded
        Try
            rowDgv = Me.dgvProcess.Rows(Me.dgvProcess.Rows.Count - 1)
            rowDgv.Cells(Me.colProcess_No.Index).Value = Me.dgvProcess.Rows.Count
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProcess_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvProcess.RowsRemoved
        Try
            Me.objCD.SetNumber(Me.dgvProcess)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProcess_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvProcess.Sorted
        Try
            Me.objCD.SetNumber(Me.dgvProcess)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsExistProcess(ByVal vProcessCode As String, ByVal vCurrentRow As Integer) As Boolean
        Try
            Me.dgvProcess.AllowUserToAddRows = False
            For Each Me.rowDgv In Me.dgvProcess.Rows
                If Me.dgvProcess.Rows.IndexOf(Me.rowDgv) <> vCurrentRow Then
                    If Me.rowDgv.Cells(Me.colProcess_Code.Index).Value.ToString.Trim.ToUpper = vProcessCode.Trim.ToUpper Then
                        Return True
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvProcess.AllowUserToAddRows = True
        End Try
    End Function

    Private Sub btn2Filter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2Filter.Click
        Dim dta As New DseMasterProduct_BOM1TableAdapters.DataTable1TableAdapter
        Dim tbl As New DseMasterProduct_BOM1.DataTable1DataTable
        Dim row As DseMasterProduct_BOM1.DataTable1Row
        Try
            Me.dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem2.Rows.Clear()
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Fill(tbl, Me.txt2CodeNo.Text.Trim & "%", Me.txt2Desc.Text.Trim & "%") > 0 Then
                Me.Cursor = Cursors.WaitCursor
                tbl.DefaultView.RowFilter = Me.GetProcessFilter
                If tbl.DefaultView.Count > 0 Then
                    For Each Me.DV In tbl.DefaultView
                        row = Me.DV.Row
                        Me.rowDgv = Me.dgvItem2.Rows(Me.dgvItem2.Rows.Add)
                        With Me.rowDgv
                            .Cells(Me.col2_GPProcess.Index).Value = Me.objCD.DigitGrouping(row.GP_Process, 2)
                            .Cells(Me.col2_Tolerance.Index).Value = Me.objCD.DigitGrouping(row.Tolerance_Process, 2)
                            .Cells(Me.col2CodeNo.Index).Value = row.Product_Code
                            .Cells(Me.col2Desc.Index).Value = row.Product_Desc
                            .Cells(Me.col2Modify.Index).Value = Boolean.FalseString
                            .Cells(Me.col2No.Index).Value = Me.dgvItem2.Rows.Count.ToString
                            .Cells(Me.col2PrimaryCodeNo.Index).Value = row.Primary_Product
                            .Cells(Me.col2PrimaryProcess.Index).Value = row.Primary_Process
                            .Cells(Me.col2PrimaryUnit.Index).Value = row.Primary_Unit
                            .Cells(Me.col2Process.Index).Value = row.Process_Code
                            .Cells(Me.col2QtyRef.Index).Value = Me.objCD.DigitGrouping(row.Qty_Ref, row._Default__Precision_)
                            .Cells(Me.col2Unit.Index).Value = row.Unit_Code
                            .Cells(Me.Col2Decimal.Index).Value = row._Default__Precision_
                        End With
                    Next
                Else
                    MsgBox("Product not found.", MsgBoxStyle.Information, "Information")
                End If
            Else
                MsgBox("Product not found.", MsgBoxStyle.Information, "Information")
            End If
            Me.dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function GetProcessFilter() As String
        GetProcessFilter = String.Empty
        Try
            Me.dgvProcess.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvProcess.AllowUserToAddRows = False
            If Me.dgvProcess.RowCount > 0 Then
                For Each Me.rowDgv In Me.dgvProcess.Rows
                    If String.IsNullOrEmpty(Me.rowDgv.Cells(Me.colProcess_Primary.Index).Value) = False Then
                        If String.IsNullOrEmpty(GetProcessFilter) Then
                            GetProcessFilter = "[Primary Process] = " & Me.rowDgv.Cells(Me.colProcess_Primary.Index).Value.ToString
                        Else
                            GetProcessFilter = GetProcessFilter & " OR [Primary Process] = " & Me.rowDgv.Cells(Me.colProcess_Primary.Index).Value.ToString
                        End If
                    End If
                Next
            End If
            Me.dgvProcess.AllowUserToAddRows = True
            Me.dgvProcess.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            If String.IsNullOrEmpty(GetProcessFilter) Then
                Return String.Empty
            Else
                Return "(" & GetProcessFilter & ")"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub llbCheckAll2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll2.LinkClicked
        Try
            Me.dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDgv In Me.dgvItem2.Rows
                Me.rowDgv.Cells(Me.col2Modify.Index).Value = Boolean.TrueString
            Next
            Me.dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUnCheckAll2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUnCheckAll2.LinkClicked
        Try
            Me.dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDgv In Me.dgvItem2.Rows
                Me.rowDgv.Cells(Me.col2Modify.Index).Value = Boolean.FalseString
            Next
            Me.dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem2_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem2.CellBeginEdit
        Try
            If e.ColumnIndex = Me.col2Modify.Index Or e.ColumnIndex = Me.col2QtyRef.Index Or e.ColumnIndex = Me.col2_GPProcess.Index Or e.ColumnIndex = Me.col2_Tolerance.Index Then
                Me.buffer = Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem2_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem2.CellEndEdit
        Try
            If e.ColumnIndex = Me.col2QtyRef.Index Then
                If Me.objCD.CekBilanganPositif(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value, Me.dgvItem2.Item(Me.Col2Decimal.Index, e.RowIndex).Value)
                If CDec(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                    Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                Me.UpdateQtyReference(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value, Me.dgvItem2.Item(Me.col2PrimaryCodeNo.Index, e.RowIndex).Value)
            End If
            If e.ColumnIndex = Me.col2_GPProcess.Index Then
                If Me.objCD.CekBilanganPositif(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                If CDec(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                    Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                If CDec(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                    Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                Me.dgvItem2.Item(Me.col2_Tolerance.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetTolerancePerc(Me.dgvItem2.Item(Me.col2_GPProcess.Index, e.RowIndex).Value), 2)
                Me.UpdateGP_AndTolerance_FG(Me.dgvItem2.Item(Me.col2_GPProcess.Index, e.RowIndex).Value, Me.dgvItem2.Item(Me.col2_Tolerance.Index, e.RowIndex).Value, Me.dgvItem2.Item(Me.col2PrimaryCodeNo.Index, e.RowIndex).Value)
            End If
            If e.ColumnIndex = Me.col2_Tolerance.Index Then
                If Me.objCD.CekNonNegatif(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                Me.dgvItem2.Item(Me.col2_GPProcess.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetGoodPercentagePerc(Me.dgvItem2.Item(Me.col2_Tolerance.Index, e.RowIndex).Value), 2)
                If CDec(Me.dgvItem2.Item(Me.col2_GPProcess.Index, e.RowIndex).Value) <= 0 Then
                    Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItem2.Item(Me.col2_GPProcess.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetGoodPercentagePerc(Me.dgvItem2.Item(Me.col2_Tolerance.Index, e.RowIndex).Value), 2)
                    Exit Sub
                End If
                Me.UpdateGP_AndTolerance_FG(Me.dgvItem2.Item(Me.col2_GPProcess.Index, e.RowIndex).Value, Me.dgvItem2.Item(Me.col2_Tolerance.Index, e.RowIndex).Value, Me.dgvItem2.Item(Me.col2PrimaryCodeNo.Index, e.RowIndex).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function UpdateQtyReference(ByVal vQtyRef As Decimal, ByVal vPrimaryProduct As Double) As Boolean
        Dim dta As New DseMasterProduct_BOM1TableAdapters.tblmaster_productTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.UpdateQuery_QtyRef(vQtyRef, vPrimaryProduct) > 0 Then
                If Me.UpdateModifiedProduct(Me.UserCode, vPrimaryProduct) = True Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function UpdateGP_AndTolerance_FG(ByVal vGP As Decimal, ByVal vTolerance As Decimal, ByVal vPrimaryProduct As Double) As Boolean
        Dim dta As New DseMasterProduct_BOM1TableAdapters.tblmaster_productTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.UpdateQuery_GP_AndTolerance(vGP, vTolerance, vTolerance, vPrimaryProduct) > 0 Then
                If Me.UpdateModifiedProduct(Me.UserCode, vPrimaryProduct) = True Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItem2_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem2.DataError

    End Sub

    Private Sub dgvItem2_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem2.Sorted
        Try
            Me.objCD.SetNumber(Me.dgvItem2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn2SetQtyRef_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2SetQtyRef.Click
        Dim frmSetQtyRef As New Core_Module.frmMasterProduct_BOM1_SetQtyReference
        Try
            If Me.GetSelectedCount(Me.dgvItem2, Me.col2Modify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Not Me.IsAllColumnSimilar(Me.dgvItem2, Me.col2Unit, Me.col2Modify) Then
                MsgBox("Unit code for all selected products must be same.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmSetQtyRef.Text = "Set Qty Ref"
            frmSetQtyRef.lblQtyReference.Text = "Qty reference"
            frmSetQtyRef.lblUnit.Text = Me.GetFirstSelectedValue(Me.dgvItem2, Me.col2Modify, Me.col2Unit)
            frmSetQtyRef.BackColorAll = Me.BackColorAll
            frmSetQtyRef.Decimal_ = Me.GetFirstSelectedValue(Me.dgvItem2, Me.col2Modify, Me.Col2Decimal)
            If frmSetQtyRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                For Each Me.rowDgv In Me.dgvItem2.Rows
                    If Me.rowDgv.Cells(Me.col2Modify.Index).Value = Boolean.TrueString Then
                        If Me.UpdateQtyReference(frmSetQtyRef.QtyReference, Me.rowDgv.Cells(Me.col2PrimaryCodeNo.Index).Value) = True Then
                            Me.rowDgv.Cells(Me.col2QtyRef.Index).Value = Me.objCD.DigitGrouping(frmSetQtyRef.QtyReference, Me.rowDgv.Cells(Me.Col2Decimal.Index).Value)
                        End If
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Function GetSelectedCount(ByVal vDgv As DataGridView, ByVal vCommitColumn As DataGridViewCheckBoxColumn) As Integer
        GetSelectedCount = 0
        Try
            For Each Me.rowDgv In vDgv.Rows
                If Me.rowDgv.Cells(vCommitColumn.Index).Value = Boolean.TrueString Then
                    GetSelectedCount += 1
                End If
            Next
            Return GetSelectedCount
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetSelectedCount
        End Try
    End Function

    Function IsAllColumnSimilar(ByVal vDgv As DataGridView, ByVal vColumn As DataGridViewTextBoxColumn, ByVal vCommitColumn As DataGridViewCheckBoxColumn) As Boolean
        Dim lastValue As Object = Nothing
        Dim trueFound As Integer = 0
        Try
            If vDgv.Rows.Count <= 1 Then
                Return True
            End If
            For Each Me.rowDgv In vDgv.Rows
                If Me.rowDgv.Cells(vCommitColumn.Index).Value = Boolean.TrueString Then
                    trueFound += 1
                    If trueFound > 1 And Me.rowDgv.Cells(vColumn.Index).Value <> lastValue Then
                        Return False
                    End If
                    lastValue = Me.rowDgv.Cells(vColumn.Index).Value
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Function GetFirstSelectedValue(ByVal vDgv As DataGridView, ByVal vModifyColumn As DataGridViewCheckBoxColumn, ByVal vColValueToGet As DataGridViewTextBoxColumn) As String
        GetFirstSelectedValue = String.Empty
        Try
            For Each Me.rowDgv In vDgv.Rows
                If Me.rowDgv.Cells(vModifyColumn.Index).Value = Boolean.TrueString Then
                    Return Me.rowDgv.Cells(vColValueToGet.Index).Value
                End If
            Next
            Return GetFirstSelectedValue
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetFirstSelectedValue
        End Try
    End Function

    Private Sub btn2SetGPProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2SetGPProcess.Click
        Dim frmSetQtyRef As New Core_Module.frmMasterProduct_BOM1_SetQtyReference
        Try
            If Me.GetSelectedCount(Me.dgvItem2, Me.col2Modify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmSetQtyRef.Text = "Set Good % Process"
            frmSetQtyRef.lblQtyReference.Text = "Good percentage"
            frmSetQtyRef.lblUnit.Text = "%"
            frmSetQtyRef.BackColorAll = Me.BackColorAll
            frmSetQtyRef.Decimal_ = 2
            frmSetQtyRef.Size = New System.Drawing.Size(frmSetQtyRef.btnOK.Left + frmSetQtyRef.btnOK.Width + 10, frmSetQtyRef.Height)
            If frmSetQtyRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                For Each Me.rowDgv In Me.dgvItem2.Rows
                    If Me.rowDgv.Cells(Me.col2Modify.Index).Value = Boolean.TrueString Then
                        If Me.UpdateGP_AndTolerance_FG(frmSetQtyRef.QtyReference, frmProduct.GetTolerancePerc(frmSetQtyRef.QtyReference), Me.rowDgv.Cells(Me.col2PrimaryCodeNo.Index).Value) = True Then
                            Me.rowDgv.Cells(Me.col2_GPProcess.Index).Value = Me.objCD.DigitGrouping(frmSetQtyRef.QtyReference, 2)
                            Me.rowDgv.Cells(Me.col2_Tolerance.Index).Value = Me.objCD.DigitGrouping(frmProduct.GetTolerancePerc(frmSetQtyRef.QtyReference), 2)
                        End If
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btn2AddRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2AddRM.Click
        Dim frmAddRM As New frmMasterProduct_BOM1_AddRM
        Try
            If Me.GetSelectedCount(Me.dgvItem2, Me.col2Modify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Not Me.IsAllColumnSimilar(Me.dgvItem2, Me.col2PrimaryProcess, Me.col2Modify) Then
                MsgBox("Process code for all selected products must be same.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            frmAddRM.UserCode = Me.UserCode
            frmAddRM.BackColorall = Me.BackColorAll
            If frmAddRM.ShowDialog = Windows.Forms.DialogResult.OK Then

                Dim listOfInvalid As New List(Of String)
                listOfInvalid = Me.GetAlreadyExistRM(frmAddRM)
                If listOfInvalid.Count > 0 Then
                    Dim existRM As String = String.Empty
                    For Each str As String In listOfInvalid
                        If String.IsNullOrEmpty(existRM) Then
                            existRM = str
                        Else
                            existRM = existRM & vbNewLine & str
                        End If
                    Next
                    MsgBox("The following raw material is already exist, add raw material process canceled." & vbNewLine & existRM, MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                listOfInvalid.Clear()
                listOfInvalid = Me.GetInfinitelyBoM(frmAddRM)
                If listOfInvalid.Count > 0 Then
                    Dim existRM As String = String.Empty
                    For Each str As String In listOfInvalid
                        If String.IsNullOrEmpty(existRM) Then
                            existRM = str
                        Else
                            existRM = existRM & vbNewLine & str
                        End If
                    Next
                    MsgBox("Adding the following RM will cause infinite recursive, process canceled." & vbNewLine & existRM, MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Me.Cursor = Cursors.WaitCursor
                For Each Me.rowDgv In Me.dgvItem2.Rows
                    If Me.rowDgv.Cells(Me.col2Modify.Index).Value = Boolean.TrueString Then
                        For Each rowDgvRM As DataGridViewRow In frmAddRM.dgvItemBOM.Rows
                            If rowDgvRM.Cells(frmAddRM.ColBOM_PrimaryProduct.Index).Value <> "-1" Then
                                Me.InsertRM(Me.rowDgv.Cells(Me.col2PrimaryCodeNo.Index).Value, rowDgvRM.Cells(frmAddRM.ColBOM_PrimaryProduct.Index).Value, _
                                rowDgvRM.Cells(frmAddRM.ColBOM_Qty.Index).Value, rowDgvRM.Cells(frmAddRM.ColBOM_GoodPercentage.Index).Value, _
                                rowDgvRM.Cells(frmAddRM.ColBOM_Tolerance.Index).Value, rowDgvRM.Cells(frmAddRM.ColBOM_Remarks.Index).Value)
                            End If
                        Next
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Function GetInfinitelyBoM(ByVal frmAddRM As frmMasterProduct_BOM1_AddRM) As List(Of String)
        Dim listOfBoM As New List(Of Double)
        Dim listOfChild As New List(Of Double)
        GetInfinitelyBoM = New List(Of String)
        Try
            For Each Me.rowDgv In frmAddRM.dgvItemBOM.Rows
                If Me.rowDgv.Cells(frmAddRM.ColBOM_PrimaryProduct.Index).Value <> "-1" Then
                    listOfBoM.Add(Me.rowDgv.Cells(frmAddRM.ColBOM_PrimaryProduct.Index).Value)
                End If
            Next
            For Each Me.rowDgv In Me.dgvItem2.Rows
                If Me.rowDgv.Cells(Me.col2Modify.Index).Value = Boolean.TrueString Then
                    listOfChild.Clear()
                    Me.GetChild(Me.rowDgv.Cells(Me.col2PrimaryCodeNo.Index).Value, listOfChild)
                    If listOfChild.Count > 0 Then
                        Dim e As String = ""
                        For Each d As Double In listOfBoM
                            If listOfChild.Contains(d) Then
                                e = objProduct.GetProductCode(d)
                                GetInfinitelyBoM.Add("Product code " & e & "  as raw material of " & Me.rowDgv.Cells(Me.col2CodeNo.Index).Value)
                            End If
                        Next
                    End If
                End If
            Next
            Return GetInfinitelyBoM
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetInfinitelyBoM
        End Try
    End Function

    Function GetInfinitelyBoM2(ByVal vPrimaryNewRM As Double) As List(Of String)
        Dim listOfBoM As New List(Of Double)
        Dim listOfChild As New List(Of Double)
        GetInfinitelyBoM2 = New List(Of String)
        Try
            Me.GetChild(Me.rowDgv.Cells(Me.ColPrimaryCodeNoRM.Index).Value, listOfChild)
            If listOfChild.Count > 0 Then
                If listOfChild.Contains(vPrimaryNewRM) Then
                    GetInfinitelyBoM2.Add("Product code " & objProduct.GetProductCode(vPrimaryNewRM) & "  as raw material of " & Me.rowDgv.Cells(Me.col2CodeNo.Index).Value)
                End If
            End If
            Return GetInfinitelyBoM2
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetInfinitelyBoM2
        End Try
    End Function

    Private Sub GetChild(ByVal vPrimaryProduct As Double, ByVal vListChild As List(Of Double))
        Dim DtaBoM As New dseMasterProduct_BOMTableAdapters.tblmaster_product_item_bomTableAdapter
        Dim TblBoM As New dseMasterProduct_BOM.tblmaster_product_item_bomDataTable
        Dim RowBoM As dseMasterProduct_BOM.tblmaster_product_item_bomRow

        Try
            DtaBoM.Connection.ConnectionString = Core_Module.GetConstring
            DtaBoM.FillBy_PrimaryParent(TblBoM, vPrimaryProduct)

            If vListChild.Contains(vPrimaryProduct) = False Then
                vListChild.Add(vPrimaryProduct)
            End If
            For Each RowBoM In TblBoM
                If vListChild.Contains(RowBoM.primary_product) = False Then
                    vListChild.Add(RowBoM.primary_product)
                End If
                Me.GetChild(RowBoM.primary_product, vListChild)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaBoM.Dispose()
            TblBoM.Dispose()
        End Try
    End Sub

    Function GetAlreadyExistRM(ByVal frmAddRM As frmMasterProduct_BOM1_AddRM) As List(Of String)
        Dim dta As New DseMasterProduct_BOM1TableAdapters.tblmaster_product_item_bomTableAdapter
        Dim tbl As New DseMasterProduct_BOM1.tblmaster_product_item_bomDataTable
        Dim row As DseMasterProduct_BOM1.tblmaster_product_item_bomRow
        GetAlreadyExistRM = New List(Of String)
        Try
            dta.ClearBeforeFill = False
            dta.Connection.ConnectionString = Core_Module.GetConstring
            tbl.Rows.Clear()
            For Each Me.rowDgv In Me.dgvItem2.Rows
                If Me.rowDgv.Cells(Me.col2Modify.Index).Value = Boolean.TrueString Then
                    dta.Fill(tbl, Me.rowDgv.Cells(Me.col2PrimaryCodeNo.Index).Value)
                End If
            Next
            For Each Me.rowDgv In frmAddRM.dgvItemBOM.Rows
                If Me.rowDgv.Cells(frmAddRM.ColBOM_PrimaryProduct.Index).Value <> "-1" Then
                    tbl.DefaultView.RowFilter = "[primary product raw] = " & Me.rowDgv.Cells(frmAddRM.ColBOM_PrimaryProduct.Index).Value.ToString
                    If tbl.DefaultView.Count > 0 Then
                        For Each DV As Data.DataRowView In tbl.DefaultView
                            row = DV.Row
                            GetAlreadyExistRM.Add(Me.rowDgv.Cells(frmAddRM.ColBOM_CodeNo.Index).Value & " as raw material of " & row.Finish_Good_Code)
                        Next
                    End If
                End If
            Next
            Return GetAlreadyExistRM
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetAlreadyExistRM
        End Try
    End Function

    Function InsertRM(ByVal vPrimaryProduct As Double, ByVal vPrimaryProductRM As Double, ByVal vQty As Decimal, ByVal vGP As Decimal, ByVal vTolerance As Decimal, _
    ByVal vRemarks As String) As Boolean
        Dim dta As New DseMasterProduct_BOM1TableAdapters.tblmaster_product_item_bomTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.InsertQuery(vPrimaryProduct, vPrimaryProductRM, vQty, vGP, vTolerance, vRemarks, vTolerance) > 0 Then
                If Me.UpdateModifiedProduct(Me.UserCode, vPrimaryProduct) = True Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
#End Region

#Region "Tab1"
    Private Sub btn1FindRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1FindRM.Click
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
            frmFindProduct.txtCari.Text = Me.txt1RMCode.Text
            frmFindProduct.TipePencarian = frmCariMasterProduct.enumTipePencarian.Goods_Only
            frmFindProduct.TipeProduct = clsMasterProduct.EnumTipeProduct.Goods
            frmFindProduct.ShowDialog()
            If frmFindProduct.PrimaryProduct > 0 Then
                Me.txt1PrimaryRM.Text = frmFindProduct.PrimaryProduct
                Me.txt1RMCode.Text = Me.objProduct.GetProductCode(frmFindProduct.PrimaryProduct)
                Me.txt1RMDesc.Text = Me.objProduct.GetDescription(frmFindProduct.PrimaryProduct)
                Me.txt1ProcessCode.Focus()
            Else
                Me.txt1PrimaryRM.Text = "-1"
                Me.txt1RMCode.Text = String.Empty
                Me.txt1RMDesc.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt1RMCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt1RMCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                Me.txt1RMCode.Text = Me.txt1RMCode.Text.Trim
                If Me.objProduct.ApakahSudahAda(Me.txt1RMCode.Text) Then
                    Me.txt1PrimaryRM.Text = Me.objProduct.GetPrimary(Me.txt1RMCode.Text)
                    Me.txt1RMDesc.Text = Me.objProduct.GetDescription(Me.txt1PrimaryRM.Text)
                    Me.txt1ProcessCode.Focus()
                Else
                    Me.btn1FindRM.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt1RMCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1RMCode.Leave
        Try
            Me.txt1RMCode.Text = Me.txt1RMCode.Text.Trim
            If Not String.IsNullOrEmpty(Me.txt1RMCode.Text) Then
                If Me.objProduct.ApakahSudahAda(Me.txt1RMCode.Text) Then
                    Me.txt1PrimaryRM.Text = Me.objProduct.GetPrimary(Me.txt1RMCode.Text)
                    Me.txt1RMDesc.Text = Me.objProduct.GetDescription(Me.txt1PrimaryRM.Text)
                Else
                    Me.btn1FindRM.PerformClick()
                End If
            Else
                Me.txt1PrimaryRM.Text = "-1"
                Me.txt1RMCode.Text = String.Empty
                Me.txt1RMDesc.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind1Process_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1FindProcess.Click
        Dim frmFindProcess As New Core_Module.FrmMasterProsesCari
        Try
            frmFindProcess.txtCari.Text = Me.txt1ProcessCode.Text
            frmFindProcess.ShowDialog()
            If frmFindProcess.mPrimary > 0 Then
                Me.txt1PrimaryProcess.Text = frmFindProcess.mPrimary
                Me.txt1ProcessCode.Text = Me.objProcess.GetCode(frmFindProcess.mPrimary)
                Me.txt1ProcessDesc.Text = Me.objProcess.GetDesc(frmFindProcess.mPrimary)
                Me.btn1Filter.Focus()
            Else
                Me.txt1PrimaryProcess.Text = "-1"
                Me.txt1ProcessCode.Text = String.Empty
                Me.txt1ProcessDesc.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt1ProcessCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt1ProcessCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                Me.txt1ProcessCode.Text = Me.txt1ProcessCode.Text.Trim
                If Me.objProcess.ApakahSudahAda(Me.txt1ProcessCode.Text) Then
                    Me.txt1PrimaryProcess.Text = Me.objProcess.GetPrimary(Me.txt1ProcessCode.Text)
                    Me.txt1ProcessDesc.Text = Me.objProcess.GetDesc(Me.txt1PrimaryProcess.Text)
                    Me.btn1Filter.Focus()
                Else
                    Me.btn1FindProcess.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt1ProcessCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1ProcessCode.Leave
        Try
            Me.txt1ProcessCode.Text = Me.txt1ProcessCode.Text.Trim
            If Not String.IsNullOrEmpty(Me.txt1ProcessCode.Text) Then
                If Me.objProcess.ApakahSudahAda(Me.txt1ProcessCode.Text) Then
                    Me.txt1PrimaryProcess.Text = Me.objProcess.GetPrimary(Me.txt1ProcessCode.Text)
                    Me.txt1ProcessDesc.Text = Me.objProcess.GetDesc(Me.txt1PrimaryProcess.Text)
                Else
                    Me.btn1FindProcess.PerformClick()
                End If
            Else
                Me.txt1PrimaryProcess.Text = "-1"
                Me.txt1ProcessCode.Text = String.Empty
                Me.txt1ProcessDesc.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn1Filter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1Filter.Click
        Dim dta As New DseMasterProduct_BOM1TableAdapters.DataTable2TableAdapter
        Dim tbl As New DseMasterProduct_BOM1.DataTable2DataTable
        Dim row As DseMasterProduct_BOM1.DataTable2Row

        Try
            If Me.txt1PrimaryRM.Text = "" Or Me.txt1PrimaryRM.Text = "-1" Then
                MsgBox("Specifiy the raw material code.", MsgBoxStyle.Information, "Information")
                Me.txt1RMCode.Focus()
                Exit Sub
            End If
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If Me.txt1PrimaryProcess.Text = "" Or Me.txt1PrimaryProcess.Text = "-1" Then
                dta.Fill(tbl, Me.txt1PrimaryRM.Text)
            Else
                dta.FillBy_PrimaryProcess(tbl, Me.txt1PrimaryRM.Text, Me.txt1PrimaryProcess.Text)
            End If
            Me.dgvItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem1.Rows.Clear()
            For Each row In tbl
                Me.rowDgv = Me.dgvItem1.Rows(Me.dgvItem1.Rows.Add())
                With Me.rowDgv
                    .Cells(Me.ColCodeNo.Index).Value = row.Product_Code_FG
                    .Cells(Me.ColDesc.Index).Value = row.Product_Desc_FG
                    .Cells(Me.ColGPProcess.Index).Value = Me.objCD.DigitGrouping(row.GP_FG, 2)
                    .Cells(Me.ColGPRM.Index).Value = Me.objCD.DigitGrouping(row.GP_RM, 2)
                    .Cells(Me.ColModify.Index).Value = Boolean.FalseString
                    .Cells(Me.ColNo.Index).Value = Me.dgvItem1.Rows.Count
                    .Cells(Me.ColPrimaryCodeNo.Index).Value = row.Primary_FG
                    .Cells(Me.colPrimaryProcess.Index).Value = row.Primary_Process_FG
                    .Cells(Me.colPrimaryUnit.Index).Value = row.Primary_Unit_FG
                    .Cells(Me.colPrimaryUnitRM.Index).Value = row.Primary_Unit_RM
                    .Cells(Me.ColProcess.Index).Value = row.Process_Code_FG
                    .Cells(Me.ColQtyRef.Index).Value = Me.objCD.DigitGrouping(row.Qty_Ref_FG, row.__Decimal__FG)
                    .Cells(Me.ColQtyRM.Index).Value = Me.objCD.DigitGrouping(row.Qty_RM, row.__Decimal__RM)
                    .Cells(Me.colToleranceProcess.Index).Value = Me.objCD.DigitGrouping(row.Tolerance_FG, 2)
                    .Cells(Me.colToleranceRM.Index).Value = Me.objCD.DigitGrouping(row.Tolerance_RM, 2)
                    .Cells(Me.ColUnit.Index).Value = row.Unit_Code_FG
                    .Cells(Me.ColUnitRM.Index).Value = row.Unit_Code_RM
                    .Cells(Me.ColDecimalFG.Index).Value = row.__Decimal__FG
                    .Cells(Me.ColDecimalRM.Index).Value = row.__Decimal__RM
                    .Cells(Me.ColPrimaryCodeNoRM.Index).Value = row.Primary_RM
                End With
            Next
            Me.dgvItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem1.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColQtyRef.Index Or e.ColumnIndex = Me.ColGPProcess.Index Or e.ColumnIndex = Me.colToleranceProcess.Index Or _
            e.ColumnIndex = Me.ColQtyRM.Index Or e.ColumnIndex = Me.ColGPRM.Index Or e.ColumnIndex = Me.colToleranceRM.Index Or e.ColumnIndex = Me.ColModify.Index Then
                Me.buffer = Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem1.CellEndEdit
        Try
            If IsNothing(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex <> Me.ColModify.IndeterminateValue Then
                    Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                End If
                If String.IsNullOrEmpty(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.buffer = Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If

                    ''qty ref FG
                    If e.ColumnIndex = Me.ColQtyRef.Index Then
                        If Me.objCD.CekBilanganPositif(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColDecimalFG.Index, e.RowIndex).Value)
                        If CDec(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.UpdateQtyReference(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                    End If

                    ''GP process
                    If e.ColumnIndex = Me.ColGPProcess.Index Then
                        If Me.objCD.CekBilanganPositif(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        If CDec(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        If CDec(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(Me.colToleranceProcess.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetTolerancePerc(Me.dgvItem1.Item(Me.ColGPProcess.Index, e.RowIndex).Value), 2)
                        Me.UpdateGP_AndTolerance_FG(Me.dgvItem1.Item(Me.ColGPProcess.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.colToleranceProcess.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                    End If

                    ''tolerance process
                    If e.ColumnIndex = Me.colToleranceProcess.Index Then
                        If Me.objCD.CekNonNegatif(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        Me.dgvItem1.Item(Me.ColGPProcess.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetGoodPercentagePerc(Me.dgvItem1.Item(Me.colToleranceProcess.Index, e.RowIndex).Value), 2)
                        If CDec(Me.dgvItem1.Item(Me.ColGPProcess.Index, e.RowIndex).Value) <= 0 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItem1.Item(Me.ColGPProcess.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetGoodPercentagePerc(Me.dgvItem1.Item(Me.colToleranceProcess.Index, e.RowIndex).Value), 2)
                            Exit Sub
                        End If
                        Me.UpdateGP_AndTolerance_FG(Me.dgvItem1.Item(Me.ColGPProcess.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.colToleranceProcess.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                    End If

                    ''qty rm 
                    If e.ColumnIndex = Me.ColQtyRM.Index Then
                        If Me.objCD.CekBilanganPositif(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value, Me.objProduct.GetAngkaDecimal(Me.dgvItem1.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                        If CDec(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        UpdateQtyRM(Me.dgvItem1.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColPrimaryCodeNoRM.Index, e.RowIndex).Value, Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If

                    ''gp rm
                    If e.ColumnIndex = Me.ColGPRM.Index Then
                        If Me.objCD.CekBilanganPositif(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        If CDec(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        If CDec(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(Me.colToleranceRM.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetTolerancePerc(Me.dgvItem1.Item(Me.ColGPRM.Index, e.RowIndex).Value), 2)
                        Me.UpdateGP_AndTolerance_RM(Me.dgvItem1.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColPrimaryCodeNoRM.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColGPRM.Index, e.RowIndex).Value.ToString, Me.dgvItem1.Item(Me.colToleranceRM.Index, e.RowIndex).Value)
                    End If

                    ''tolerance rm
                    If e.ColumnIndex = Me.colToleranceRM.Index Then
                        If Me.objCD.CekBilanganPositif(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        If CDec(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        Me.dgvItem1.Item(Me.ColGPRM.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetGoodPercentagePerc(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value), 2)
                        Me.UpdateGP_AndTolerance_RM(Me.dgvItem1.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColPrimaryCodeNoRM.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColGPRM.Index, e.RowIndex).Value.ToString, Me.dgvItem1.Item(Me.colToleranceRM.Index, e.RowIndex).Value)
                    End If

                Else
                    Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            Else
                Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem1.DataError

    End Sub

    Private Sub dgvItem1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem1.Sorted
        Try
            Me.objCD.SetNumber(dgvItem1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function UpdateQtyRM(ByVal vPrimaryProduct As Double, ByVal vPrimaryRM As Double, ByVal vQtyRm As Decimal) As Boolean
        Dim dta As New DseMasterProduct_BOM1TableAdapters.tblmaster_product_item_bomTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.UpdateQuery_QtyRM(vQtyRm, vPrimaryProduct, vPrimaryRM) > 0 Then
                If Me.UpdateModifiedProduct(Me.UserCode, vPrimaryProduct) = True Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Function UpdateGP_AndTolerance_RM(ByVal vPrimaryProduct As Double, ByVal vPrimaryRM As Double, ByVal vGP As Decimal, ByVal vTolerance As Decimal) As Boolean
        Dim dta As New DseMasterProduct_BOM1TableAdapters.tblmaster_product_item_bomTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.UpdateQuery_GP_AndTolerance(vGP, vTolerance, vPrimaryProduct, vPrimaryRM) > 0 Then
                If Me.UpdateModifiedProduct(Me.UserCode, vPrimaryProduct) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub llbCheckAll1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll1.LinkClicked
        Try
            Me.dgvItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDgv In Me.dgvItem1.Rows
                Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString
            Next
            Me.dgvItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll1.LinkClicked
        Try
            Me.dgvItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDgv In Me.dgvItem1.Rows
                Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.FalseString
            Next
            Me.dgvItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn1AddQtyRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1AddQtyRM.Click
        Dim frmSetQtyRef As New Core_Module.frmMasterProduct_BOM1_SetQtyReference
        Try
            If Me.GetSelectedCount(Me.dgvItem1, Me.ColModify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmSetQtyRef.Text = "Add Qty RM"
            frmSetQtyRef.lblQtyReference.Text = "Additional qty RM"
            frmSetQtyRef.lblUnit.Text = Me.GetFirstSelectedValue(Me.dgvItem1, Me.ColModify, Me.ColUnitRM)
            frmSetQtyRef.BackColorAll = Me.BackColorAll
            frmSetQtyRef.Decimal_ = Me.GetFirstSelectedValue(Me.dgvItem1, Me.ColModify, Me.ColDecimalRM)
            If frmSetQtyRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                For Each Me.rowDgv In Me.dgvItem1.Rows
                    If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        If Me.AddRM(rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value, rowDgv.Cells(Me.ColPrimaryCodeNoRM.Index).Value, frmSetQtyRef.QtyReference) = True Then
                            Me.rowDgv.Cells(Me.ColQtyRM.Index).Value = Me.objCD.DigitGrouping(CDec(Me.rowDgv.Cells(Me.ColQtyRM.Index).Value) + frmSetQtyRef.QtyReference, Me.rowDgv.Cells(Me.ColDecimalRM.Index).Value)
                        End If
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function AddRM(ByVal vPrimaryProduct As Double, ByVal vPrimaryRM As Double, ByVal vAdditionalRMQty As Decimal) As Boolean
        Dim dta As New DseMasterProduct_BOM1TableAdapters.tblmaster_product_item_bomTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.UpdateQuery_AddRM(vAdditionalRMQty, vPrimaryProduct, vPrimaryRM) <= 0 Then
                Return False
            Else
                If Me.UpdateModifiedProduct(Me.UserCode, vPrimaryProduct) = False Then
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function UpdateModifiedProduct(ByVal vUserCode As String, ByVal vPrimaryProduct As Double) As Boolean
        Dim dta As New DseMasterProduct_BOM1TableAdapters.tblmaster_productTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.UpdateQuery_ModifiedByDate(vUserCode, Now, vPrimaryProduct) > 0 Then
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub btn1SetQtyRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1SetQtyRM.Click
        Dim frmSetQtyRef As New Core_Module.frmMasterProduct_BOM1_SetQtyReference
        Try
            If Me.GetSelectedCount(Me.dgvItem1, Me.ColModify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmSetQtyRef.Text = "Set Qty RM"
            frmSetQtyRef.lblQtyReference.Text = "Qty raw material"
            frmSetQtyRef.lblUnit.Text = Me.GetFirstSelectedValue(Me.dgvItem1, Me.ColModify, Me.ColUnitRM)
            frmSetQtyRef.BackColorAll = Me.BackColorAll
            frmSetQtyRef.Decimal_ = Me.GetFirstSelectedValue(Me.dgvItem1, Me.ColModify, Me.ColDecimalRM)
            If frmSetQtyRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                For Each Me.rowDgv In Me.dgvItem1.Rows
                    If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        If Me.UpdateQtyRM(rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value, rowDgv.Cells(Me.ColPrimaryCodeNoRM.Index).Value, frmSetQtyRef.QtyReference) = True Then
                            Me.rowDgv.Cells(Me.ColQtyRM.Index).Value = Me.objCD.DigitGrouping(frmSetQtyRef.QtyReference, Me.rowDgv.Cells(Me.ColDecimalRM.Index).Value)
                        End If
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btn1AddQtyRMByPersen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1AddQtyRMByPersen.Click
        Dim frmSetQtyRef As New Core_Module.frmMasterProduct_BOM1_SetQtyReference
        Try
            If Me.GetSelectedCount(Me.dgvItem1, Me.ColModify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmSetQtyRef.Text = "Add Qty RM by %"
            frmSetQtyRef.lblQtyReference.Text = "Additional % RM"
            frmSetQtyRef.lblUnit.Text = "%"
            frmSetQtyRef.BackColorAll = Me.BackColorAll
            frmSetQtyRef.Decimal_ = Me.GetFirstSelectedValue(Me.dgvItem1, Me.ColModify, Me.ColDecimalRM)
            If frmSetQtyRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                Dim newVal As Decimal = 0
                For Each Me.rowDgv In Me.dgvItem1.Rows
                    If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        newVal = Me.objCD.DigitGrouping(Me.rowDgv.Cells(Me.ColQtyRM.Index).Value + (Me.rowDgv.Cells(Me.ColQtyRM.Index).Value * (frmSetQtyRef.QtyReference / 100)), Me.rowDgv.Cells(Me.ColDecimalRM.Index).Value)
                        If Me.UpdateQtyRM(rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value, rowDgv.Cells(Me.ColPrimaryCodeNoRM.Index).Value, newVal) = True Then
                            Me.rowDgv.Cells(Me.ColQtyRM.Index).Value = Me.objCD.DigitGrouping(newVal, Me.rowDgv.Cells(Me.ColDecimalRM.Index).Value)
                        End If
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btn1SetGPProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1SetGPProcess.Click
        Dim frmSetQtyRef As New Core_Module.frmMasterProduct_BOM1_SetQtyReference
        Try
            If Me.GetSelectedCount(Me.dgvItem1, Me.ColModify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmSetQtyRef.Text = "Set Good % Process"
            frmSetQtyRef.lblQtyReference.Text = "Good percentage"
            frmSetQtyRef.lblUnit.Text = "%"
            frmSetQtyRef.BackColorAll = Me.BackColorAll
            frmSetQtyRef.Decimal_ = 2
            frmSetQtyRef.Size = New System.Drawing.Size(frmSetQtyRef.btnOK.Left + frmSetQtyRef.btnOK.Width + 10, frmSetQtyRef.Height)
            If frmSetQtyRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                For Each Me.rowDgv In Me.dgvItem1.Rows
                    If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        If Me.UpdateGP_AndTolerance_FG(frmSetQtyRef.QtyReference, frmProduct.GetTolerancePerc(frmSetQtyRef.QtyReference), Me.rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value) = True Then
                            Me.rowDgv.Cells(Me.ColGPProcess.Index).Value = Me.objCD.DigitGrouping(frmSetQtyRef.QtyReference, 2)
                            Me.rowDgv.Cells(Me.colToleranceProcess.Index).Value = Me.objCD.DigitGrouping(frmProduct.GetTolerancePerc(frmSetQtyRef.QtyReference), 2)
                        End If
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btn1SetGPRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1SetGPRM.Click
        Dim frmSetQtyRef As New Core_Module.frmMasterProduct_BOM1_SetQtyReference
        Try
            If Me.GetSelectedCount(Me.dgvItem1, Me.ColModify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmSetQtyRef.Text = "Set Good % RM"
            frmSetQtyRef.lblQtyReference.Text = "Good % RM"
            frmSetQtyRef.lblUnit.Text = "%"
            frmSetQtyRef.BackColorAll = Me.BackColorAll
            frmSetQtyRef.Decimal_ = 2
            frmSetQtyRef.Size = New System.Drawing.Size(frmSetQtyRef.btnOK.Left + frmSetQtyRef.btnOK.Width + 10, frmSetQtyRef.Height)
            If frmSetQtyRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                For Each Me.rowDgv In Me.dgvItem1.Rows
                    If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        If Me.UpdateGP_AndTolerance_RM(Me.rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.rowDgv.Cells(Me.ColPrimaryCodeNoRM.Index).Value, frmSetQtyRef.QtyReference, frmProduct.GetTolerancePerc(frmSetQtyRef.QtyReference)) = True Then
                            Me.rowDgv.Cells(Me.ColGPRM.Index).Value = Me.objCD.DigitGrouping(frmSetQtyRef.QtyReference, 2)
                            Me.rowDgv.Cells(Me.colToleranceRM.Index).Value = Me.objCD.DigitGrouping(frmProduct.GetTolerancePerc(frmSetQtyRef.QtyReference), 2)
                        End If
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btn1Remove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1Remove.Click
        Try
            If Me.GetSelectedCount(Me.dgvItem1, Me.ColModify) = 0 Then
                MsgBox("Select at least one product to remove RM.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If MsgBox("Are you sure to remove raw material?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                For i As Integer = Me.dgvItem1.Rows.Count - 1 To 0 Step -1  'Each Me.rowDgv In Me.dgvItem1.Rows
                    Me.rowDgv = Me.dgvItem1.Rows(i)
                    If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        If Me.DeleteRM(Me.rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.rowDgv.Cells(Me.ColPrimaryCodeNoRM.Index).Value) = True Then
                            Me.dgvItem1.Rows.RemoveAt(i)
                        End If
                    End If
                Next
                Me.objCD.SetNumber(Me.dgvItem1)
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function DeleteRM(ByVal vPrimaryProduct As Double, ByVal vPrimaryRM As Double) As Boolean
        Dim dta As New DseMasterProduct_BOM1TableAdapters.tblmaster_product_item_bomTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.DeleteQuery(vPrimaryProduct, vPrimaryRM) > 0 Then
                If Me.UpdateModifiedProduct(Me.UserCode, vPrimaryProduct) = True Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btn1Replace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1Replace.Click
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
            If Me.GetSelectedCount(Me.dgvItem1, Me.ColModify) = 0 Then
                MsgBox("Select at least one product to replace RM.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            frmFindProduct.ShowDialog()
            If frmFindProduct.PrimaryProduct > 0 Then
                If Me.GetFirstSelectedValue(Me.dgvItem1, Me.ColModify, Me.ColPrimaryCodeNoRM) = frmFindProduct.PrimaryProduct Then
                    MsgBox("Select another raw material.", MsgBoxStyle.Information, "Information")
                    Me.txt1RMCode.Focus()
                    Exit Sub
                End If

                ''-----------------------------------------------------------------------------------------
                Dim dtaBoM As New DseMasterProduct_BOM1TableAdapters.tblmaster_product_item_bomTableAdapter
                Dim tblBoM As New DseMasterProduct_BOM1.tblmaster_product_item_bomDataTable
                Dim rowBoM As DseMasterProduct_BOM1.tblmaster_product_item_bomRow
                Dim alreadyExistRM As String = String.Empty
                dtaBoM.Connection.ConnectionString = Core_Module.GetConstring
                dtaBoM.ClearBeforeFill = False
                For Each Me.rowDgv In Me.dgvItem1.Rows
                    If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        dtaBoM.Fill(tblBoM, Me.rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value)
                    End If
                Next
                tblBoM.DefaultView.RowFilter = "[primary product raw] = " & frmFindProduct.PrimaryProduct.ToString
                If tblBoM.DefaultView.Count > 0 Then
                    For Each Me.DV In tblBoM.DefaultView
                        rowBoM = Me.DV.Row
                        If String.IsNullOrEmpty(alreadyExistRM) Then
                            alreadyExistRM = rowBoM.Finish_Good_Code
                        Else
                            alreadyExistRM = alreadyExistRM & vbNewLine & rowBoM.Finish_Good_Code
                        End If
                    Next
                End If
                If Not String.IsNullOrEmpty(alreadyExistRM) Then
                    MsgBox("The following WIP/FG already has " & Me.objProduct.GetProductCode(frmFindProduct.PrimaryProduct) & " as BoM" & vbNewLine & alreadyExistRM, MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                ''-----------------------------------------------------------------------------------------
                Dim listOfInvalid As New List(Of String)
                listOfInvalid = GetInfinitelyBoM2(frmFindProduct.PrimaryProduct)
                If listOfInvalid.Count > 0 Then
                    Dim existRM As String = String.Empty
                    For Each str As String In listOfInvalid
                        If String.IsNullOrEmpty(existRM) Then
                            existRM = str
                        Else
                            existRM = existRM & vbNewLine & str
                        End If
                    Next
                    MsgBox("Adding the following RM will cause infinite recursive, process canceled." & vbNewLine & existRM, MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If MsgBox("Are you sure to replace RM with " & Me.objProduct.GetProductCode(frmFindProduct.PrimaryProduct) & "?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    For Each Me.rowDgv In Me.dgvItem1.Rows
                        If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                            If UpdatePrimaryRM(Me.rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.rowDgv.Cells(Me.ColPrimaryCodeNoRM.Index).Value, frmFindProduct.PrimaryProduct) Then

                            End If
                        End If
                    Next
                    Me.btn1Filter.PerformClick()
                    MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Function UpdatePrimaryRM(ByVal vPrimaryProduct As Double, ByVal vPrimaryRM_Old As Double, ByVal vPrimaryRM_New As Double) As Boolean
        Dim dta As New DseMasterProduct_BOM1TableAdapters.tblmaster_product_item_bomTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.UpdateQuery_PrimaryRM(vPrimaryRM_New, vPrimaryRM_Old, vPrimaryProduct) > 0 Then
                If Me.UpdateModifiedProduct(Me.UserCode, vPrimaryProduct) = True Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region
   
End Class