Option Explicit On
Option Strict Off

Public Class frmMasterProduct_BOM1_AddRM

    Public BackColorall As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = String.Empty

    Dim objCD As New Core_Module.ClsCekData
    Dim frmProduct As New Core_Module.frmMasterProduct
    Dim buffer As String = String.Empty
    Dim objProduct As New Core_Module.clsMasterProduct
    Dim IndexRemove As Integer = 0

    Private Sub dgvItemBOM_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemBOM.CellBeginEdit
        Try
            If e.RowIndex = Me.dgvItemBOM.Rows.Count - 1 Then
                If e.ColumnIndex <> Me.ColBOM_CodeNo.Index Then
                    e.Cancel = True
                Else
                    Me.buffer = Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            Else
                If e.ColumnIndex = Me.ColBOM_CodeNo.Index Or _
                e.ColumnIndex = Me.ColBOM_GoodPercentage.Index Or _
                e.ColumnIndex = Me.ColBOM_Qty.Index Or _
                e.ColumnIndex = Me.ColBOM_Remarks.Index Or _
                e.ColumnIndex = Me.ColBOM_Tolerance.Index Then
                    Me.buffer = Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemBOM_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemBOM.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvItemBOM.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryCodeNo As Double

        Try
            With row
                If IsNothing(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If

                    If e.ColumnIndex = Me.ColBOM_CodeNo.Index Then
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.objProduct.ApakahSudahAda(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryCodeNo = Me.objProduct.GetPrimary(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexBOM(PrimaryCodeNo)
                            If index < 0 Then
                                Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value = Me.objProduct.GetPrimary(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvItemBOM.Item(Me.ColBOM_CodeNo.Index, e.RowIndex).Value = Me.objProduct.GetProductCode(PrimaryCodeNo)
                                Me.dgvItemBOM.Item(Me.ColBOM_Desc.Index, e.RowIndex).Value = Me.objProduct.GetDescription(PrimaryCodeNo)
                                Me.dgvItemBOM.Item(Me.ColBOM_Qty.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(1, Me.objProduct.GetAngkaDecimal(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value))
                                Me.dgvItemBOM.Item(Me.ColBOM_GoodPercentage.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(100, 2)
                                Me.dgvItemBOM.Item(Me.ColBOM_Tolerance.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(0, 2)
                                Me.dgvItemBOM.Item(Me.ColBOM_Unit.Index, e.RowIndex).Value = Me.objProduct.GetUnitCode(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication in product code, please change your product code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    If Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                        Me.IndexRemove = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If

                                    Exit Sub
                                Else
                                    Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.frmCariMasterProduct
                            ij.TipePencarian = frmCariMasterProduct.enumTipePencarian.Goods_Only
                            ij.txtCari.Text = Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            index = Me.GetIndexBOM(ij.PrimaryProduct)
                            If ij.PrimaryProduct = -1 Then
                                If Me.buffer = "" Then
                                    Me.IndexRemove = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                                Exit Sub
                            Else
                                If index < 0 Then
                                    Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value = ij.PrimaryProduct.ToString
                                    Me.dgvItemBOM.Item(Me.ColBOM_CodeNo.Index, e.RowIndex).Value = Me.objProduct.GetProductCode(ij.PrimaryProduct)
                                    Me.dgvItemBOM.Item(Me.ColBOM_Desc.Index, e.RowIndex).Value = Me.objProduct.GetDescription(ij.PrimaryProduct)
                                    Me.dgvItemBOM.Item(Me.ColBOM_Qty.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(1, Me.objProduct.GetAngkaDecimal(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value))
                                    Me.dgvItemBOM.Item(Me.ColBOM_GoodPercentage.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(100, 2)
                                    Me.dgvItemBOM.Item(Me.ColBOM_Tolerance.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(0, 2)
                                    Me.dgvItemBOM.Item(Me.ColBOM_Unit.Index, e.RowIndex).Value = Me.objProduct.GetUnitCode(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication in product code, please change your product code", MsgBoxStyle.Information, "Information")
                                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        If Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                            Me.IndexRemove = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        End If
                                        Exit Sub
                                    Else
                                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                            End If
                            ij = Nothing
                        End If
                    ElseIf e.ColumnIndex = Me.ColBOM_Qty.Index Then
                        If Me.objCD.CekBilanganPositif(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value, Me.objProduct.GetAngkaDecimal(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value))
                    ElseIf e.ColumnIndex = Me.ColBOM_GoodPercentage.Index Then
                        If Me.objCD.CekBilanganPositif(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        If CDec(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        Me.dgvItemBOM.Item(Me.ColBOM_Tolerance.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetTolerancePerc(Me.dgvItemBOM.Item(Me.ColBOM_GoodPercentage.Index, e.RowIndex).Value), 2)
                    ElseIf e.ColumnIndex = Me.ColBOM_Remarks.Index Then
                        If Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        End If
                    ElseIf e.ColumnIndex = Me.ColBOM_Tolerance.Index Then
                        If Me.objCD.CekBilanganPositif(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        If CDec(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        Me.dgvItemBOM.Item(Me.ColBOM_GoodPercentage.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetGoodPercentagePerc(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value), 2)
                    End If
                Else
                    If e.ColumnIndex = Me.ColBOM_CodeNo.Index Then
                        If Me.buffer = "" Then
                            Me.IndexRemove = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        Exit Sub
                    ElseIf e.ColumnIndex = Me.ColBOM_Qty.Index Or e.ColumnIndex = Me.ColBOM_GoodPercentage.Index Then
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    ElseIf e.ColumnIndex = Me.ColBOM_Remarks.Index Then
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexBOM(ByVal vPrimaryCodeNo As Double) As Double
        Try
            For Each rowDGV As DataGridViewRow In Me.dgvItemBOM.Rows
                If rowDGV.Index = Me.dgvItemBOM.Rows.Count - 1 Then
                    Return -1
                Else
                    If rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value = vPrimaryCodeNo Then
                        Return rowDGV.Index
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Dim PrimaryProduct As Double = -1

        Try
            PrimaryProduct = Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, Me.IndexRemove).Value
            Me.dgvItemBOM.Rows.RemoveAt(Me.IndexRemove)
            Me.IndexRemove = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterProduct_BOM1_AddRM_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DialogResult = Windows.Forms.DialogResult.None
            Me.dgvItemBOM.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorall
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

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.dgvItemBOM.Rows.Count = 1 Then
                MsgBox("Input at least 1 item raw material.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemBOM_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemBOM.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            rowDGV = Me.dgvItemBOM.Rows(Me.dgvItemBOM.Rows.Count - 1)
            With rowDGV
                .Cells(Me.ColBOM_CodeNo.Index).Value = ""
                .Cells(Me.ColBOM_Desc.Index).Value = ""
                .Cells(Me.ColBOM_GoodPercentage.Index).Value = 100.0
                .Cells(Me.ColBOM_PrimaryProduct.Index).Value = -1
                .Cells(Me.ColBOM_Qty.Index).Value = 0
                .Cells(Me.ColBOM_Remarks.Index).Value = ""
                .Cells(Me.ColBOM_Unit.Index).Value = "Unit"
                .Cells(Me.ColBOM_Tolerance.Index).Value = "0.00"
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 
End Class