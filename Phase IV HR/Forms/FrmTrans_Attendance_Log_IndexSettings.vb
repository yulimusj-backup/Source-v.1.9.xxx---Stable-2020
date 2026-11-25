Public Class FrmTrans_Attendance_Log_IndexSettings

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            year_index.Text = My.Settings.YEAR
            year_lenght.Text = My.Settings.YEAR_LENGTH
            month_index.Text = My.Settings.MONTH
            month_length.Text = My.Settings.MONTH_LENGTH
            day_index.Text = My.Settings.DATE_
            day_length.Text = My.Settings.DATE_LENGTH
            hour_index.Text = My.Settings.HOUR
            hour_length.Text = My.Settings.HOUR_LENGTH
            minute_index.Text = My.Settings.MINUTE
            minute_length.Text = My.Settings.MINUTE_LENGTH
            user_index.Text = My.Settings.USER_ID
            user_length.Text = My.Settings.USER_ID_LENGTH
            remove_zero.Checked = My.Settings.REMOVE_ZERO_BEFORE_UID
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            My.Settings.YEAR = year_index.Text
            My.Settings.YEAR_LENGTH = year_lenght.Text
            My.Settings.MONTH = month_index.Text
            My.Settings.MONTH_LENGTH = month_length.Text
            My.Settings.DATE_ = day_index.Text
            My.Settings.DATE_LENGTH = day_length.Text
            My.Settings.HOUR = hour_index.Text
            My.Settings.HOUR_LENGTH = hour_length.Text
            My.Settings.MINUTE = minute_index.Text
            My.Settings.MINUTE_LENGTH = minute_length.Text
            My.Settings.USER_ID = user_index.Text
            My.Settings.USER_ID_LENGTH = user_length.Text
            My.Settings.REMOVE_ZERO_BEFORE_UID = remove_zero.Checked
            My.Settings.Save()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class