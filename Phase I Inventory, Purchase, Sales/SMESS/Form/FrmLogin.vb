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
        Dim ObjIdentifier As New CustomSMESS.ClsIdentifier
        Dim VersionDB As String = ""
        Dim ObjVersionControl As New Core_Module.Clstrans_VersionControl
        Dim ObjVersion As New ClsVersionControl
        Dim LastVersionTxt As String = ""
        Dim DefaultVersion As String = "1.6.40"
        Dim ObjUpdate1 As New ClsUpdateProgram

        Try
            Me.Label4.Width += 5
            Me.Label2.Width += 20
            'Me.Label4.SendToBack()

            Control.CheckForIllegalCrossThreadCalls = False
            FrmSplashScreen.ShowDialog()
            Me.Timer1.Enabled = False
            'Me.Text = "Login - SMESS - Ver. " + My.Settings.db_version

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

            If Me.CekConnectionString() = False Then
                MsgBox("Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Me.Close()
            End If

            If ObjIdentifier.ApakahIdentifierDBCocok = False Then
                MsgBox("The custom module loaded is not for the current database, please contact your vendor if you see this message.", _
                MsgBoxStyle.Information, "Information")
                Me.Close()
                Exit Sub
            End If

            'update version pada db
            If My.Settings.a.Trim <> "" Then
                VersionDB = ObjVersionControl.GetVersion(My.Settings.a)
                LastVersionTxt = ObjVersion.GetVersion(DefaultVersion)
                If ObjUpdate1.ApakahVersionBerbeda(LastVersionTxt, VersionDB) = True Then
                    'jika true, maka version DB > last version
                    'update last version
                    'ini terjadi kalo abis update
                    ObjVersion.SetVersion(VersionDB)
                ElseIf ObjUpdate1.ApakahVersionBerbeda(VersionDB, LastVersionTxt) = True Then
                    'jika true, maka last version > version db
                    'update version db
                    'ini terjadi kalo lisence baru masuk atau setelah update
                    ObjVersionControl.SetVersion(My.Settings.a, LastVersionTxt)
                End If
            End If
            Me.Label7.Text = Format(Today, "yyyy")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjIdentifier = Nothing
            ObjVersionControl = Nothing
            ObjVersion = Nothing
            ObjUpdate1 = Nothing
        End Try
    End Sub

    Private Function CekConnectionString() As Boolean
        Dim Acuan As String = Core_Module.GetConstring.ToLower
        Dim Compare As String = ""

        Try
            Compare = CustomSMESS.GetConstring_CustomModule.ToLower
            'MsgBox(Acuan + vbCrLf + Core_Module.GetConstring_CoreModule1.ToLower)
            If Core_Module.GetConstring_CoreModule1.ToLower.StartsWith(Acuan) = False Then
                MsgBox("Invalid connection string on core module 1")
                MsgBox(Core_Module.GetConstring_CoreModule1.ToLower)
                Return False
            End If
            If Core_Module.GetConstring_CoreModule2.ToLower.StartsWith(Acuan) = False Then
                MsgBox("Invalid connection string on core module 2")
                Return False
            End If
            If AccModule.GetConstring_AccModule.ToLower.StartsWith(Acuan) = False Then
                MsgBox("Invalid connection string on acc module")
                Return False
            End If
            If Client_Update_Module.GetConstring_ClientUpdateModule.ToLower.StartsWith(Acuan) = False Then
                MsgBox("Invalid connection string on client update module")
                Return False
            End If
            If CustomSMESS.GetConstring_CustomModule.ToLower.StartsWith(Acuan) = False Then
                MsgBox("Invalid connection string on custom module")
                Return False
            End If
            If E_mail_Module.GetConstring_EmailModule.ToLower.StartsWith(Acuan) = False Then
                MsgBox("Invalid connection string on email module")
                Return False
            End If
            If HR.GetConstring_HRModule.ToLower.StartsWith(Acuan) = False Then
                MsgBox("Invalid connection string on HR module")
                Return False
            End If
            If Inventory_Module.GetConstring_Inventory_Module.ToLower.StartsWith(Acuan) = False Then
                MsgBox("Invalid connection string on inventory module")
                Return False
            End If
            If Purchase_Module.GetConstring_Purchase_Module.ToLower.StartsWith(Acuan) = False Then
                MsgBox("Invalid connection string on purchase module")
                Return False
            End If
            If Sales_Module.GetConstring_Sales_Module.ToLower.StartsWith(Acuan) = False Then
                MsgBox("Invalid connection string on sales module")
                Return False
            End If
            If SMESS.GetConstring_SMESS_Module.ToLower.StartsWith(Acuan) = False Then
                MsgBox("Invalid connection string on smess module")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

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
        Me.PasswordTextBox.Text = "kuroinu13114"
        Me.OK.PerformClick()

    End Sub
End Class
