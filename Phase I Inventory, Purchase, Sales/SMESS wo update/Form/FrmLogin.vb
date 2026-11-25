Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmLogin
    Dim ObjLC As New Core_Module.ClsLicense
    Dim ObjMonitor As New Core_Module.ClsTransaksiMonitor

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Try
            'login
            Me.UsernameTextBox.Text = Me.UsernameTextBox.Text.Trim
            Me.PasswordTextBox.Text = Me.PasswordTextBox.Text.Trim

            If Core_Module.ModUser.Login(Me.UsernameTextBox, Me.PasswordTextBox) = True Then
                If Me.ObjMonitor.ApakahLicenseSMESSCukup(True) = False Then
                    MsgBox("Insufficient SMESS active license, please release some license.", _
                    MsgBoxStyle.Information, "Information")
                    Dim i As New Core_Module.FrmTransaksiMonitor
                    i.ShowDialog()
                    Exit Sub
                End If

                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                FrmMain.Show()
                Me.PasswordTextBox.Clear()
                Me.Hide()
                Windows.Forms.Cursor.Current = Cursors.Arrow
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub UsernameTextBox_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles UsernameTextBox.GotFocus
        Me.UsernameTextBox.BackColor = Color.FromArgb(0, 231, 249)
    End Sub

    Private Sub UsernameTextBox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles UsernameTextBox.LostFocus
        Me.UsernameTextBox.BackColor = Me.PasswordTextBox.BackColor
    End Sub

    Private Sub PasswordTextBox_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles PasswordTextBox.GotFocus
        Me.PasswordTextBox.BackColor = Color.FromArgb(0, 231, 249)
    End Sub

    Private Sub PasswordTextBox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles PasswordTextBox.LostFocus
        Me.PasswordTextBox.BackColor = Me.UsernameTextBox.BackColor
    End Sub

    Private Sub FrmLogin_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Label4.Width += 5
            Me.Label2.Width += 20
            'Me.Label4.SendToBack()

            Control.CheckForIllegalCrossThreadCalls = False
            FrmSplashScreen.ShowDialog()
            Me.Timer1.Enabled = False
            Me.Text = "Login - SMESS - Ver. " + My.Settings.db_version

            Me.UsernameTextBox.Focus()
            Me.LogoPictureBox.BackColor = Core_Module.GetSMESSColor
            Me.Label1.BackColor = Core_Module.GetSMESSColor
            Me.Label2.BackColor = Core_Module.GetSMESSColor
            Me.Label3.BackColor = Core_Module.GetSMESSColor
            Me.Label4.BackColor = Core_Module.GetSMESSColor

            Me.OK.BackColor = Core_Module.GetSMESSColor
            Me.Cancel.BackColor = Core_Module.GetSMESSColor
            Me.LineShape1.BorderColor = Core_Module.GetSMESSColor
            Me.LineShape2.BorderColor = Core_Module.GetSMESSColor
            Me.LineShape3.BorderColor = Core_Module.GetSMESSColor

            Me.LineShape3.X1 = Me.Width - 3
            Me.LineShape3.X2 = Me.Width - 3
            Me.LineShape2.Y1 = Me.Height - 3
            Me.LineShape2.Y2 = Me.Height - 3

            Me.PasswordTextBox.UseSystemPasswordChar = True
            Me.PasswordTextBox.PasswordChar = "*"
            Me.BringToFront()

            Me.Timer1.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub TxtKeyIn_UserCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Me.BackgroundWorker1.RunWorkerAsync()
    'End Sub

    ''Private Sub TxtKeyIn_UserCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtKeyIn_UserCode.KeyDown
    ''    Dim ft As New Font(Me.TxtKeyIn_UserCode.Font, FontStyle.Regular)

    ''    Try
    ''        If Me.UsernameTextBox.Text = "" And Me.TxtKeyIn_UserCode.Text.ToLower = "user name" Then
    ''            Me.TxtKeyIn_UserCode.Font = ft
    ''            Me.TxtKeyIn_UserCode.CharacterCasing = CharacterCasing.Upper
    ''            Me.TxtKeyIn_UserCode.Text = ""
    ''        End If
    ''    Catch ex As Exception
    ''        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    ''    End Try
    ''End Sub

    'Private Sub TxtKeyIn_UserCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Me.BackgroundWorker2.RunWorkerAsync()
    'End Sub

    'Private Sub TxtKeyIn_Password_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Dim ft As New Font(Me.TxtKeyIn_Password.Font, FontStyle.Regular)

    '    Try
    '        If Me.PasswordTextBox.Text = "" Then
    '            Me.TxtKeyIn_Password.Font = ft
    '            Me.TxtKeyIn_Password.UseSystemPasswordChar = True
    '            Me.TxtKeyIn_Password.Text = ""
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub TxtKeyIn_Password_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Dim ft As New Font(Me.TxtKeyIn_Password.Font, FontStyle.Italic)

    '    Try
    '        Me.TxtKeyIn_Password.Text = Me.TxtKeyIn_Password.Text.Trim

    '        Me.PasswordTextBox.Text = Me.TxtKeyIn_Password.Text

    '        If Me.PasswordTextBox.Text = "" Then
    '            Me.TxtKeyIn_Password.Font = ft
    '            Me.TxtKeyIn_Password.UseSystemPasswordChar = False
    '            Me.TxtKeyIn_Password.Text = "Password"
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
    '    Threading.Thread.Sleep(10)
    'End Sub

    'Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
    '    Dim ft As New Font(Me.TxtKeyIn_UserCode.Font, FontStyle.Regular)

    '    Try
    '        If Me.UsernameTextBox.Text = "" Then
    '            Me.TxtKeyIn_UserCode.Text = ""
    '            Me.TxtKeyIn_UserCode.Font = ft
    '            Me.TxtKeyIn_UserCode.CharacterCasing = CharacterCasing.Upper
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
    '    Threading.Thread.Sleep(10)
    'End Sub

    'Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
    '    Dim ft As New Font(Me.TxtKeyIn_UserCode.Font, FontStyle.Italic)

    '    Try
    '        Me.TxtKeyIn_UserCode.Text = Me.TxtKeyIn_UserCode.Text.Trim
    '        Me.UsernameTextBox.Text = Me.TxtKeyIn_UserCode.Text

    '        If Me.UsernameTextBox.Text = "" Then
    '            Me.TxtKeyIn_UserCode.Font = ft
    '            Me.TxtKeyIn_UserCode.CharacterCasing = CharacterCasing.Normal
    '            Me.TxtKeyIn_UserCode.Text = "User Name"
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.Timer1.Enabled = False
        Me.UsernameTextBox.Text = "USER"
        Me.PasswordTextBox.Text = "simbada69"
        Me.OK.PerformClick()

    End Sub
End Class
