Public Class FrmMasterWorkshift_CopyData
    Public pilihan As Integer = 0
    Public Batal As Boolean = False
    Public tipeWorkshift As ClsMasterWorkShift.EnumTipeWorkshift = ClsMasterWorkShift.EnumTipeWorkshift.HR
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.RboWeeklyActivities.Checked Then
                pilihan = ClsMasterWorkShift.EnumPilihan.weekly
            ElseIf Me.RboHolidayList.Checked Then
                pilihan = ClsMasterWorkShift.EnumPilihan.holidaylist
            ElseIf Me.rboAttendanceSetting.Checked Then
                pilihan = ClsMasterWorkShift.EnumPilihan.AttendanceSetting
            ElseIf Me.RboBoth.Checked Then
                pilihan = ClsMasterWorkShift.EnumPilihan.All
            End If
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
        Me.Batal = True
    End Sub

    Private Sub FrmMasterWorkshift_CopyData_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            If Me.tipeWorkshift = ClsMasterWorkShift.EnumTipeWorkshift.Production Then
                Me.rboAttendanceSetting.Checked = False
                Me.rboAttendanceSetting.Enabled = False
            End If
            Me.RboWeeklyActivities.Checked = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub
End Class