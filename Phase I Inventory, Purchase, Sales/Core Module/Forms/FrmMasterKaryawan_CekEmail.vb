Option Explicit On
Option Strict Off

Public Class FrmMasterKaryawan_CekEmail
    Public UserCode As String = ""
    Public ListPrimaryKaryawan As List(Of Double)
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawan3TableAdapter
    Dim TBL As New DseMasterKaryawan.tblmaster_karyawan3DataTable
    Dim row As DseMasterKaryawan.tblmaster_karyawan3Row
    Dim PrimaryKary As Double
    Dim IsClose As Boolean = False
    Dim buffer As String = ""

    Private Sub FrmMasterKaryawan_CekEmail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.dgvEmailFix.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Dim rowDGV As DataGridViewRow

        Try
            'load data karyawan
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)

            For Each PrimaryKary In ListPrimaryKaryawan
                TBL.DefaultView.RowFilter = "Primary = " + PrimaryKary.ToString
                If TBL.DefaultView.Count > 0 Then
                    row = TBL.DefaultView(0).Row
                    Me.dgvEmailFix.Rows.Add()
                    rowDGV = Me.dgvEmailFix.Rows(Me.dgvEmailFix.Rows.Count - 1)

                    With rowDGV
                        .Cells(Me.ColDepartment.Index).Value = row.Department
                        .Cells(Me.ColEmailAddress.Index).Value = ""
                        .Cells(Me.ColEmployeeCode.Index).Value = row.Code
                        .Cells(Me.ColEmployeeName.Index).Value = row.Full_Name
                        .Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                        .Cells(Me.ColPosition.Index).Value = row.Job_Position
                        .Cells(Me.ColPrimaryKaryawan.Index).Value = row.Primary
                    End With
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Dim rowDGV As DataGridViewRow
        Dim AdaSalah As Boolean = False

        Try
            For Each rowDGV In Me.dgvEmailFix.Rows
                If rowDGV.Cells(Me.ColEmailAddress.Index).Value = "" Then
                    AdaSalah = True
                    Exit For
                End If
            Next

            If AdaSalah = True Then
                If MsgBox("There's still invalid email address, ignore and close anyway?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.IsClose = True
                    Me.Close()
                Else
                    Me.IsClose = False
                End If
            Else
                Me.IsClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKaryawan_CekEmail_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim rowDGV As DataGridViewRow
        Dim AdaSalah As Boolean = False

        Try
            If Me.IsClose = False Then
                For Each rowDGV In Me.dgvEmailFix.Rows
                    If rowDGV.Cells(Me.ColEmailAddress.Index).Value = "" Then
                        AdaSalah = True
                        Exit For
                    End If
                Next

                If AdaSalah = True Then
                    If MsgBox("There's still invalid email address, ignore and close anyway?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmailFix_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvEmailFix.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColEmailAddress.Index Then
                Me.buffer = Me.dgvEmailFix.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmailFix_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvEmailFix.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColEmailAddress.Index Then
                'pastikan format emailnya tepat, jika tidak maka kembalikan nilainya ke buffer
                Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                If Expression.IsMatch(Me.dgvEmailFix.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("The email address is not valid.", MsgBoxStyle.Information, "Information")
                    Me.dgvEmailFix.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    Me.DTA.UpdateQuery_SET_Email(Me.dgvEmailFix.Item(e.ColumnIndex, e.RowIndex).Value, _
                    Me.dgvEmailFix.Item(Me.ColPrimaryKaryawan.Index, e.RowIndex).Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class