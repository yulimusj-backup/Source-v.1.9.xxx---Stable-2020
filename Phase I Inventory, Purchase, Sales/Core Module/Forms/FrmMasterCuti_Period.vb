''@rosadinnn
''@20150910 0806
Option Explicit On
Option Strict Off
Public Class FrmMasterCuti_Period

    Public mPeriod As Date
    Dim ObjCd As New Core_Module.ClsCekData

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Try
            Me.DtpPeriod.Value = CDate(Me.mPeriod.Year.ToString & Format(Me.DtpPeriod.Value, "MMM dd"))
            Me.DtpPeriod.Value = Me.ObjCd.GetFirstDateofTheMonth(Me.DtpPeriod.Value)
            Me.DialogResult = Windows.Forms.DialogResult.OK
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class