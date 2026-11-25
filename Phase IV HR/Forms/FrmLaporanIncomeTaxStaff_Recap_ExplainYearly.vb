Option Explicit On
Option Strict Off

Public Class FrmLaporanIncomeTaxStaff_Recap_ExplainYearly

    Private Sub FrmLaporanIncomeTaxStaff_Recap_ExplainYearly_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.TxtExplain.DeselectAll()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class