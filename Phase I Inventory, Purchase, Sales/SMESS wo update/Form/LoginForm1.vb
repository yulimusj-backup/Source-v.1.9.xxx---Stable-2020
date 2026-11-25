Imports Core_Module

Public Class LoginForm1
    Dim Var_Drag As Boolean
    Dim Var_ClickAt As Point
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
        Me.UsernameTextBox.Text = Me.UsernameTextBox.Text.Trim
        Me.PasswordTextBox.Text = Me.PasswordTextBox.Text.Trim
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
        Me.UsernameTextBox.BackColor = Core_Module.GetSMESSColor
    End Sub

    Private Sub UsernameTextBox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles UsernameTextBox.LostFocus
        Me.UsernameTextBox.BackColor = Me.PasswordTextBox.BackColor
    End Sub

    Private Sub PasswordTextBox_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles PasswordTextBox.GotFocus
        Me.PasswordTextBox.BackColor = Core_Module.GetSMESSColor
    End Sub

    Private Sub PasswordTextBox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles PasswordTextBox.LostFocus
        Me.PasswordTextBox.BackColor = Me.UsernameTextBox.BackColor
    End Sub

    Private Sub FrmLogin_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Text = "Login - SMESS - Ver. " + My.Settings.db_version
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PictureBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseDown
        Try
            'when left cliek
            If e.Button = Windows.Forms.MouseButtons.Left = True Then
                Me.Var_Drag = True
                Me.Var_ClickAt = New Point(e.X, e.Y)
            Else
                Me.Var_Drag = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
        Try
            'when user drag
            If Me.Var_Drag = True Then
                Dim Var_POint As Point
                Var_POint = Me.PointToScreen(New Point(e.X, e.Y))
                Var_POint.Offset(-Me.Var_ClickAt.X, -Me.Var_ClickAt.Y)
                Me.Location = Var_POint
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PictureBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseUp
        Me.Var_Drag = False
    End Sub
End Class
