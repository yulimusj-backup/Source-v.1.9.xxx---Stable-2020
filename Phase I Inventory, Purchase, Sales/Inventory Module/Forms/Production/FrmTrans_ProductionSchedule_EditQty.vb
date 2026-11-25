Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_EditQty
    Public mDefPrec As Integer = 0
    Public mJmlProduce As Decimal = 0
    Public mPrimaryProduct As Double = -1
    Public mUnitCode As String = ""
    Public mPrimaryPredefinedWU As Double = -1
    Public ListPrimaryWU As New List(Of Double)

    Dim ObjCD As New Core_Module.ClsCekData
    Dim Buffer As String = ""

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Cancel data change?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim rowDGV As DataGridViewRow
        Dim JmlProduce1 As Decimal = 0
        Dim ListKodeMesin As New List(Of String)
        Dim DummyStr As String = ""
        Dim DummyInt As Integer = 0

        Try
            'cek jumlahnya harus konsisten
            For Each rowDGV In Me.DgvWUAllocate_WU.Rows
                JmlProduce1 += CDec(rowDGV.Cells(Me.CoAll_WU_ProduceQty.Index).Value)
            Next
            If mJmlProduce <> JmlProduce1 Then
                MsgBox("Please make sure total produce qty: " + Me.ObjCD.DigitGrouping(Me.mJmlProduce, Me.mDefPrec) + " " + mUnitCode, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            'pastikan yang di cek box, ada qty produce nya
            For Each rowDGV In Me.DgvWUAllocate_WU.Rows
                If rowDGV.Cells(Me.CoAll_WU_Allocate.Index).Value = Boolean.TrueString And _
                CDec(rowDGV.Cells(Me.CoAll_WU_ProduceQty.Index).Value) = 0 Then
                    ListKodeMesin.Add(rowDGV.Cells(Me.CoAll_WU_Code.Index).Value)
                End If
            Next
            If ListKodeMesin.Count > 0 Then
                For Each KodeMesin As String In ListKodeMesin
                    DummyInt += 1

                    If DummyStr = "" Then
                        DummyStr = DummyInt.ToString + ". " + KodeMesin
                    Else
                        DummyStr = DummyStr + vbCrLf + DummyInt.ToString + ". " + KodeMesin
                    End If
                Next

                MsgBox("Please specify a positive number for qty produce data for the following work unit:" + vbCrLf + DummyStr, _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If MsgBox("Accept data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListKodeMesin = Nothing
        End Try
    End Sub

    Private Sub DgvWUAllocate_WU_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvWUAllocate_WU.CellBeginEdit
        Try
            If e.ColumnIndex = Me.CoAll_WU_Allocate.Index Then
                Me.Buffer = Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value
            ElseIf e.ColumnIndex = Me.CoAll_WU_ProduceQty.Index Then
                If Me.mPrimaryPredefinedWU = -1 Then
                    If Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_Allocate.Index, e.RowIndex).Value = Boolean.FalseString Then
                        e.Cancel = True
                    Else
                        Me.Buffer = Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    If Me.ListPrimaryWU.Contains(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryWU.Index, e.RowIndex).Value) = True Then
                        If Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_Allocate.Index, e.RowIndex).Value = Boolean.FalseString Then
                            e.Cancel = True
                        Else
                            Me.Buffer = Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        e.Cancel = True
                    End If
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWUAllocate_WU_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvWUAllocate_WU.CellEndEdit
        Try
            If IsNothing(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Buffer
            Else
                If e.ColumnIndex = Me.CoAll_WU_Allocate.Index Then
                    'jika jadi tidak alocate, maka hitung kembali defiasi
                    If Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_Allocate.Index, e.RowIndex).Value = Boolean.FalseString Then
                        Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_ProduceQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, Me.mDefPrec)
                        Me.SetDeviation()
                    End If
                ElseIf e.ColumnIndex = Me.CoAll_WU_ProduceQty.Index Then
                    'cek bilangan positif
                    If Me.ObjCD.CekBilanganPositif(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify a positive data.", MsgBoxStyle.Information, "Information")
                        Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Buffer
                        Exit Sub
                    End If

                    'digit grouping
                    Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value, Me.mDefPrec)

                    'set deviasi
                    Me.SetDeviation()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetDeviation()
        Dim rowDGV As DataGridViewRow
        Dim TotalProduce1 As Decimal = 0

        Try
            For Each rowDGV In Me.DgvWUAllocate_WU.Rows
                TotalProduce1 += CDec(rowDGV.Cells(Me.CoAll_WU_ProduceQty.Index).Value)
                Me.lblDeviasi.Text = "Deviation = " + Me.ObjCD.DigitGrouping((Me.mJmlProduce - TotalProduce1), Me.mDefPrec) + " " + Me.mUnitCode
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_EditQty_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.mPrimaryPredefinedWU <> -1 Then
                Me.CoAll_WU_Allocate.Visible = False
            End If
            Me.lblTotalQtyProduce.Text = "Total Qty Produce = " + Me.ObjCD.DigitGrouping(Me.mJmlProduce, Me.mDefPrec) + " " + Me.mUnitCode
            Me.SetDeviation()
            Me.DgvWUAllocate_WU.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class