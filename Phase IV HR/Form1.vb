Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim frm1 As New FrmMasterAttendanceServer
        Dim frm2 As New Core_Module.FrmMasterKaryawan
        Dim frm3 As New Core_Module.FrmMasterWorkShift
        Dim frm4 As New FrmTrans_Attendance_DataImport
        Dim frm5 As New FrmTrans_Attendance
        Dim frm6 As New Core_Module.FrmMasterWorkShift_Allocate
        Dim frm7 As New FrmTrans_Attendance_Log
        Dim frm8 As New FrmTrans_Attendance_Actual_In_Report
        Dim FrmMasterLeave As New Core_Module.FrmMasterCuti

        Try
            If Me.RadioButton1.Checked = True Then
                frm1.UserCode = Me.TextBox1.Text
                frm1.ShowDialog()
            ElseIf Me.RadioButton2.Checked = True Then
                frm2.UserCode = Me.TextBox1.Text
                frm2.ShowDialog()
            ElseIf Me.RadioButton3.Checked = True Then
                frm3.UserCode = Me.TextBox1.Text
                frm3.ShowDialog()
            ElseIf Me.RadioButton4.Checked = True Then
                frm4.UserCode = Me.TextBox1.Text
                frm4.ShowDialog()
            ElseIf Me.RadioButton7.Checked = True Then
                frm7.UserCode = Me.TextBox1.Text
                frm7.ShowDialog()
            ElseIf Me.RadioButton8.Checked = True Then
                frm8.UserCode = Me.TextBox1.Text
                frm8.ShowDialog()
            ElseIf Me.rboMasterLeave.Checked = True Then
                FrmMasterLeave.UserCode = Me.TextBox1.Text
                FrmMasterLeave.ShowDialog()
            ElseIf Me.RadioButton5.Checked = True Then
                frm5.UserCode = Me.TextBox1.Text
                frm5.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub RadioButton8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton8.CheckedChanged

    End Sub
End Class