Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing
Public Class FrmTransaksiPayment_Income_Deduction
    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub dgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellClick
        Dim FrmExplain As New FrmTrans_Gaji_ShowExplaination
        Try
            If Me.dgvItem.Rows.Count <= 0 Then
                Exit Sub
            End If
            If e.RowIndex < 0 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColExplain.Index Then
                FrmExplain.TxtFormula.Text = Me.dgvItem.Item(Me.ColExplaination.Index, e.RowIndex).Value
                If Me.dgvItem.Item(Me.ColOverride.Index, e.RowIndex).Value = True Then
                    FrmExplain.TxtFormula.Text += vbCrLf + "Override Value = " + Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value.ToString
                    FrmExplain.TxtFormula.Text += vbCrLf + "Reason = " + Me.dgvItem.Item(Me.ColOverrideReason.Index, e.RowIndex).Value.ToString
                End If
                FrmExplain.ShowDialog()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmExplain.Dispose()
        End Try
    End Sub

    Public Sub SetNumber(ByVal dgv As DataGridView, Optional ByVal vSetAutosize As Boolean = True, _
    Optional ByVal BerikutYgdHide As Boolean = True)
        Dim NomorBaru As Integer = 1
        Dim Putih As Boolean = False
        Dim Row As DataGridViewRow
        Try
            If vSetAutosize = True Then
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            End If
            dgv.DefaultCellStyle.BackColor = Color.White
            If BerikutYgdHide = False Then
                For i As Integer = 1 To dgv.Rows.Count
                    Row = Me.dgvItem.Rows(i - 1)
                    If dgv.Rows(i - 1).Visible = True Then
                        dgv.Item(0, i - 1).Value = NomorBaru
                        NomorBaru += 1
                        If Putih = False Then
                            Row.DefaultCellStyle.BackColor = Color.White
                            Putih = True
                        Else
                            Row.DefaultCellStyle.BackColor = Me.BackColorAll
                            Putih = False
                        End If
                    End If
                Next
            Else
                For i As Integer = 1 To dgv.Rows.Count
                    Row = Me.dgvItem.Rows(i - 1)
                    dgv.Item(0, i - 1).Value = i
                    If Putih = False Then
                        Row.DefaultCellStyle.BackColor = Color.White
                        Putih = True
                    Else
                        Row.DefaultCellStyle.BackColor = Me.BackColorAll
                        Putih = False
                    End If
                Next
            End If
            If vSetAutosize = True Then
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkShowZeroAmount_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkShowZeroAmount.CheckedChanged
        Dim RowDGV As DataGridViewRow
        Dim ApakahBackColor As Boolean = False
        Try
            If Me.ChkShowZeroAmount.Checked = False Then
                If Me.dgvItem.Rows.Count > 0 Then
                    For Each RowDGV In Me.dgvItem.Rows
                        If CDec(RowDGV.Cells(Me.ColAmount.Index).Value) = 0 Then
                            Me.dgvItem.Rows(RowDGV.Index).Visible = False
                            If ApakahBackColor = False Then
                                RowDGV.DefaultCellStyle.BackColor = Color.White
                                ApakahBackColor = True
                            Else
                                RowDGV.DefaultCellStyle.BackColor = Me.BackColorAll
                                ApakahBackColor = False
                            End If
                        End If
                    Next
                End If
                Me.SetNumber(Me.dgvItem, True, False)
            Else
                If Me.dgvItem.Rows.Count > 0 Then
                    For Each RowDGV In Me.dgvItem.Rows
                        Me.dgvItem.Rows(RowDGV.Index).Visible = True
                        If ApakahBackColor = False Then
                            RowDGV.DefaultCellStyle.BackColor = Color.White
                            ApakahBackColor = True
                        Else
                            RowDGV.DefaultCellStyle.BackColor = Me.BackColorAll
                            ApakahBackColor = False
                        End If
                    Next
                End If
                Me.SetNumber(Me.dgvItem, True, True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPayment_Income_Deduction_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim RowDGV As DataGridViewRow
        Dim ApakahBackColor As Boolean = False
        Try
            Me.ShowIcon = False
            Me.ChkShowZeroAmount.Checked = False
            If Me.dgvItem.Rows.Count > 0 Then
                For Each RowDGV In Me.dgvItem.Rows
                    If CDec(RowDGV.Cells(Me.ColAmount.Index).Value) = 0 Then
                        Me.dgvItem.Rows(RowDGV.Index).Visible = False
                        If ApakahBackColor = False Then
                            RowDGV.DefaultCellStyle.BackColor = Color.White
                            ApakahBackColor = True
                        Else
                            RowDGV.DefaultCellStyle.BackColor = Me.BackColorAll
                            ApakahBackColor = False
                        End If
                    End If
                Next
            End If
            Me.SetNumber(Me.dgvItem, True, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class