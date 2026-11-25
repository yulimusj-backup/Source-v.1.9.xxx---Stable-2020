Public Class FrmTrans_ManufacturingOrder_RMReqDate
    Public MODate As Date = #1/1/1950#
    Dim LangsungClose As Boolean = False

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Try
            If Me.ApakahDataBenar = False Then
                Exit Sub
            Else
                If MsgBox("Set raw material required date as " + Format(Me.dtpRequired.Value.Date, "dd MMM yyyy") + "?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.LangsungClose = True
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Try
            If Me.MODate.Date > Me.dtpRequired.Value.Date Then
                MsgBox("Please specify required date with minimum value: " + Format(Me.MODate, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FrmTrans_ManufacturingOrder_RMReqDate_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.LangsungClose = False Then
                If Me.ApakahDataBenar = False Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Cancel data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.LangsungClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class