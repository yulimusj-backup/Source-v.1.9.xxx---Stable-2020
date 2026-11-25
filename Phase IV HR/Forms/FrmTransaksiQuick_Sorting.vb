Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports System.Drawing
Imports system.io
Public Class FrmTransaksiQuick_Sorting
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TipeTransaksi As ClsTransaksiQuick_Sorting.EnumTipeTransaksi
    Public ShowUlang As Boolean = False
    'Public ListKolom As String = ""

    Dim PrimaryUser As Double = -1
    Dim Tbl As New DseTransaksiQuick_Sorting.tbltrans_setting_user_sortingDataTable
    Dim StrAsc As String = "Ascending"
    Dim StrDesc As String = "Descending"
    Dim StrUnSorted As String = "Unsorted"
    Dim Buffer As String = ""
    Dim FromBtnOK As Boolean = False
    Dim rowIdxUbah As Integer = -1

    Dim ObjUser As New Core_Module.ClsUser
    Dim Obj As New ClsTransaksiQuick_Sorting
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub FrmTransaksiQuick_Sorting_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.FromBtnOK = False Then
                If MsgBox("Save settings before close?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.Obj.SimpanSetting(Me.PrimaryUser, Me.TipeTransaksi, Me)
                    Me.ShowUlang = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.FromBtnOK = False
        End Try
    End Sub

    Private Sub FrmTransaksiQuick_Sorting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.PrimaryUser = Me.ObjUser.GetPrimary(Me.UserCode)

            Me.DgvItem.Rows.Add("Employee Code", StrAsc, 1)
            Me.DgvItem.Rows.Add("Employee NIK", StrAsc, 2)
            Me.DgvItem.Rows.Add("Department Code", StrAsc, 3)
            Me.DgvItem.Rows.Add("Job Position Code", StrAsc, 4)
            Me.DgvItem.Rows.Add("Employee Name", StrAsc, 5)
            Me.DgvItem.Rows.Add("Transaction No", StrAsc, 6)
            Me.DgvItem.Rows.Add("Transaction Date", StrAsc, 7)
            Me.DgvItem.Rows.Add("Period From", StrUnSorted, "-")
            Me.DgvItem.Rows.Add("Period To", StrUnSorted, "-")

            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.Obj.LoadSetting(Me.PrimaryUser, Me.TipeTransaksi, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click

        Try
            Me.FromBtnOK = True
            If MsgBox("Save settings?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                Me.Close()
                Exit Sub
            End If

            Me.Obj.SimpanSetting(Me.PrimaryUser, Me.TipeTransaksi, Me)
            Me.ShowUlang = True
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColDesc.Index Then
                e.Cancel = True
            ElseIf e.ColumnIndex = Me.ColSortType.Index Then
                Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                If IsNumeric(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Dim Row As DataGridViewRow = Me.DgvItem.Rows(e.RowIndex)
        Dim rd As DataGridViewRow
        Dim RowIdx As Integer = e.RowIndex
        Dim ijk As Integer = 1
        Dim MaxPriority As Integer = 0

        Try
            For Each rd In Me.DgvItem.Rows
                If Me.Obj.GetTipeSorting(rd.Cells(Me.ColSortType.Index).Value) <> ClsTransaksiQuick_Sorting.EnumTipeSorting.UnSorted Then
                    MaxPriority += 1
                End If
            Next

            If e.ColumnIndex = Me.ColSortType.Index Then
                If Me.Obj.GetTipeSorting(Row.Cells(e.ColumnIndex).Value) = ClsTransaksiQuick_Sorting.EnumTipeSorting.UnSorted Then
                    Dim PriorityBuff As Integer = CInt(Row.Cells(Me.ColPriority.Index).Value)
                    Row.Cells(Me.ColPriority.Index).Value = "-"
                    For Each rd In Me.DgvItem.Rows
                        If IsNumeric(rd.Cells(Me.ColPriority.Index).Value) = True Then
                            If CDec(PriorityBuff) < CDec(rd.Cells(Me.ColPriority.Index).Value) Then
                                rd.Cells(Me.ColPriority.Index).Value = CDec(rd.Cells(Me.ColPriority.Index).Value) - 1
                            End If
                        End If
                    Next
                Else
                    'max priority 
                    If IsNumeric(Row.Cells(Me.ColPriority.Index).Value) = False Then
                        If Row.Cells(e.ColumnIndex).Value <> Me.Buffer Then
                            Row.Cells(Me.ColPriority.Index).Value = MaxPriority
                        End If
                    End If
                End If
            ElseIf e.ColumnIndex = Me.ColPriority.Index Then
                'batasan qty positif
                If String.IsNullOrEmpty(Row.Cells(Me.ColPriority.Index).Value) = True Then
                    Row.Cells(Me.ColPriority.Index).Value = Me.Buffer
                    Exit Sub
                End If
                If Me.ObjCD.CekBilanganPositif(Row.Cells(Me.ColPriority.Index).Value) = False Then
                    Row.Cells(Me.ColPriority.Index).Value = Me.Buffer
                    Exit Sub
                End If
                If CDec(Row.Cells(Me.ColPriority.Index).Value) > MaxPriority Then
                    Row.Cells(Me.ColPriority.Index).Value = Me.Buffer
                Else
                    For Each rd In Me.DgvItem.Rows
                        If rd.Index <> RowIdx Then
                            If Row.Cells(Me.ColPriority.Index).Value = rd.Cells(Me.ColPriority.Index).Value Then
                                rd.Cells(Me.ColPriority.Index).Value = Me.Buffer
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.rowIdxUbah = e.RowIndex
        End Try
    End Sub
End Class