Imports Core_Module
Public Class FrmLogin
    Dim ObjLC As New Core_Module.ClsLicense

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        'login
        'Me.UsernameTextBox.Text = Me.UsernameTextBox.Text.Trim
        'Me.PasswordTextBox.Text = "123"
        If Core_Module.ModUser.Login(Me.UsernameTextBox, Me.PasswordTextBox) = True Then

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            FrmMain.Show()
            Me.PasswordTextBox.Clear()
            Me.Hide()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End If
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub UsernameTextBox_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles UsernameTextBox.GotFocus
        Me.UsernameTextBox.BackColor = Color.MistyRose
    End Sub

    Private Sub UsernameTextBox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles UsernameTextBox.LostFocus
        Me.UsernameTextBox.BackColor = Me.PasswordTextBox.BackColor
    End Sub

    Private Sub PasswordTextBox_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles PasswordTextBox.GotFocus
        Me.PasswordTextBox.BackColor = Color.MistyRose
    End Sub

    Private Sub PasswordTextBox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles PasswordTextBox.LostFocus
        Me.PasswordTextBox.BackColor = Me.UsernameTextBox.BackColor
    End Sub

    Private Sub FrmLogin_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Text = "Login - SMESS - Ver. " + CStr(My.Application.Info.Version.Major) + "." + _
            CStr(My.Application.Info.Version.Minor) + "." + CStr(My.Application.Info.Version.Build)
        Me.Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Me.UsernameTextBox.Text = "USER"
            Me.PasswordTextBox.Text = "123"
            Me.Timer1.Enabled = False
            Me.OK.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
