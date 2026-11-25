Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_Exception
    Public FrmRef As New FrmTrans_ProductionSchedule_NewSch2
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub FrmTrans_ProductionSchedule_Exception_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV, rowDGV1 As DataGridViewRow

        Try
            Me.btnOK.Enabled = True
            For Each rowDGV In FrmRef.dgvException.Rows
                Me.dgvException.Rows.Add()
                rowDGV1 = Me.dgvException.Rows(Me.dgvException.Rows.Count - 1)
                For ij As Integer = 0 To Me.dgvException.ColumnCount - 1
                    rowDGV1.Cells(ij).Value = rowDGV.Cells(ij).Value
                Next
                If rowDGV1.Cells(Me.ColException_AllowIgnore.Index).Value = 1 Then
                    Me.btnOK.Enabled = False
                    rowDGV1.DefaultCellStyle.BackColor = Color.LightPink
                End If
            Next

            If Me.btnOK.Enabled = False Then
                Me.btnCancel.Text = "Close"
                Me.Label1.Text = "We found the following exception, the one in red must be resolved first."
            Else
                Me.dgvException.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            'If MsgBox("Continue save?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
            'End If
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
End Class