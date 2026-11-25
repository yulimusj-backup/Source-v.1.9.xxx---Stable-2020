''@rosadinnn
''2015-04-15

Option Explicit On
Option Strict Off
Imports Microsoft.Win32
Imports System.Net
Public Class FrmEmailSettings_Test
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public EmailFrom As String = ""
    Public ReplyTo As String = ""
    Public UserName As String = ""
    Public Password As String
    Public SMTPServer As String = ""
    Public SMTPPort As String = ""
    Public CCto As String = ""
    Public Subject As String = ""
    Public Message As String = ""
    Public PrefixFilename As String = ""
    Public UseSSL As Boolean = False

    Dim countEmail As Integer = 0
    Dim sentEmail As Integer = 0
    Dim ListAdd As New List(Of String)
    Dim ListAddCc As New List(Of String)
    Dim ListAtt As New List(Of String)
    Private Sub FrmEmailSettings_Test_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            FrmEmailSettings_Test.CheckForIllegalCrossThreadCalls = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.GotFocus
        Try
            Me.TextBox1.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.LostFocus
        Try
            Me.TextBox1.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TextBox1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        Try
            If e.KeyCode = 13 Then
                Button1.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            ''validasi format email sederhana
            Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
            If Expression.IsMatch(TextBox1.Text) = False Then
                MsgBox("The email address is not valid.", MsgBoxStyle.Information, "Information")
                TextBox1.Focus()
                TextBox1.SelectAll()
                Return
            End If

            ''append alamat, cc, dan attachment
            countEmail = 1
            sentEmail = 0
            Me.ListAdd.Add(Me.TextBox1.Text)
            Me.ListAddCc.Add(Me.CCto)
            If My.Computer.FileSystem.FileExists(Application.StartupPath & "\SMESS _ ABOUT US.pdf") = False Then
                'copy dari my setting ke folder ybs
                MsgBox("File 'SMESS _ ABOUT US.pdf' on " & Application.StartupPath & " doesn't exist", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.ListAtt.Add(Application.StartupPath & "\SMESS _ ABOUT US.pdf")

            ''bgw kirim email
            SendEmail(Me.ListAdd, Me.ListAddCc, Me.Subject, Me.Message, Me.ListAtt, Me.EmailFrom, Me.ReplyTo, Me.UserName, Me.Password, Me.SMTPServer, _
            Me.SMTPPort, Me.UseSSL)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SendEmail(ByVal ListAddress As List(Of String), ByVal ListAddressCc As List(Of String), _
                ByVal vEmailSubject As String, ByVal vEmailMessage As String, ByVal vListFileName As List(Of String), _
                ByVal vMailFrom As String, ByVal vReplyTo As String, ByVal vUsername As String, ByVal vPassword As String, _
                ByVal vSMTPServer As String, ByVal vSMTPPort As String, ByVal vUseSSL As Boolean)
        Dim AddTo, AddToCc, vFileName As String
        Try

            Me.Cursor = Cursors.WaitCursor
            Me.Label2.Visible = True
            Me.Button1.Enabled = False
            Me.TextBox1.Enabled = False

            Dim MyMailMessage As New Net.Mail.MailMessage()
            For Each AddTo In ListAddress
                MyMailMessage.To.Add(New Mail.MailAddress(AddTo))
            Next
            For Each AddToCc In ListAddressCc
                MyMailMessage.CC.Add(New Mail.MailAddress(AddToCc))
            Next
            MyMailMessage.Subject = vEmailSubject
            MyMailMessage.Body = vEmailMessage
            MyMailMessage.From = New Mail.MailAddress(vMailFrom)
            MyMailMessage.ReplyTo = New Mail.MailAddress(vReplyTo)
            MyMailMessage.DeliveryNotificationOptions = Mail.DeliveryNotificationOptions.OnSuccess
            For Each vFileName In vListFileName
                If vFileName <> "" Then
                    Dim fileSize As Long
                    fileSize = My.Computer.FileSystem.GetFileInfo(vFileName).Length
                    If fileSize > 10485760 Then
                        MsgBox("Maximum allowable file size is 10 mb.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    MyMailMessage.Attachments.Add(New Mail.Attachment(vFileName))
                End If
            Next
            Dim SMTPServer As New Mail.SmtpClient
            SMTPServer.Credentials = New Net.NetworkCredential(vUsername, vPassword)
            SMTPServer.Host = vSMTPServer
            SMTPServer.Port = vSMTPPort
            SMTPServer.EnableSsl = vUseSSL
            Dim UserState As Object = MyMailMessage
            AddHandler SMTPServer.SendCompleted, AddressOf SMTPServer_OnCompleted
            SMTPServer.SendAsync(MyMailMessage, UserState)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SMTPServer_OnCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        Try
            Dim CompletedMailMessage As Mail.MailMessage = CType(e.UserState, Mail.MailMessage)
            Dim EmailTo As String = CompletedMailMessage.To.Item(0).ToString
            Dim EmailFrom As String = CompletedMailMessage.From.Address
            Dim EmailSubject As String = CompletedMailMessage.Subject
            If e.Cancelled Then
                MsgBox("SendAsync email to " + EmailTo.ToString + " regarding " + EmailSubject.ToString + " was canceled.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If e.Error Is Nothing Then
                Me.sentEmail += 1
                MsgBox("SendAsync email to " + EmailTo.ToString + " regarding " + EmailSubject.ToString + " was *SENT*.", MsgBoxStyle.Information, "Information")
            Else
                Me.sentEmail += 1
                MsgBox("SendAsync email to " + EmailTo.ToString + " regarding " + EmailSubject.ToString + " was *NOT SENT*.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.TextBox1.Enabled = True
            Me.Button1.Enabled = True
            Me.Label2.Visible = False
            Me.Cursor = Cursors.Default
        End Try
    End Sub
End Class