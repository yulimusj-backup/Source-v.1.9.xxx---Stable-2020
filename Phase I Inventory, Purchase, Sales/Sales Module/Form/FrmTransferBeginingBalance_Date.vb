Public Class FrmTransferBeginingBalance_Date
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If Me.dtpFrom.Value.Date < Me.GetSMESSConvertionDate.Date Or Me.dtpTo.Value.Date < Me.GetSMESSConvertionDate.Date Then
                MsgBox("Period cannot less than SMESS convertion date: #" & Format(Me.GetSMESSConvertionDate, "dd MMMM yyyy") & "#.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Start date must be smaller than end date.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Err")
        End Try
    End Sub

    Private Function GetSMESSConvertionDate() As Date
        Dim ObjAcc As New Core_Module.ClsTranskasiAccClosing

        Try
            Return ObjAcc.GetSMESSConvertionDate
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Err")
        Finally
            ObjAcc = Nothing
        End Try
    End Function
End Class