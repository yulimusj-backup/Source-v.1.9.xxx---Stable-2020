Option Explicit On
Option Strict Off

Public Class FrmTrans_Gaji_ShowExplaination

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransGaji_ShowExplaination_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ''Me.TxtFormula.Text = "Name : Basic Salary" + vbCrLf + vbTab + "Nominal = 2,000,000.00"
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class