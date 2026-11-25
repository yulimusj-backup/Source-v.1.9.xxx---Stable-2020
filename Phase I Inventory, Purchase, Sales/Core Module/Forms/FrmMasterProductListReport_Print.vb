''@rosadinnn
''@20150915 0125

Option Explicit On
Option Strict Off

Public Class FrmMasterProductListReport_Print

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class