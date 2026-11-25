''@rosadinnn
''20150424

Option Explicit On
Option Strict Off

Imports Microsoft.Win32
Imports System.Net

Public Class FrmEmailSettings_LogReport_View
    Public TipeTrans As Core_Module.ClsEmailSettings.EnumTipeTransaksiEmail
    Public PrimaryTransaksiAcuan As Double = -1
    Public DocNo As String = ""
    Public PrimaryCustomer As Double = -1
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim objEmail As New ClsEmailSettings
    Dim objCd As New Core_Module.ClsCekData
    Dim objEmailSettings As New Core_Module.ClsEmailSettings
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim vListAddr As List(Of String) = New List(Of String)
    Dim vListAttc As List(Of String) = New List(Of String)
    Dim vReplyTo As String = ""
    Dim vUsername As String = ""
    Dim vPassword As String = ""
    Dim vSmtpServer As String = ""
    Dim vSmtpPort As String = ""
    Dim vUseSSL As Boolean = False

    Public vAttachmentFilename As String

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAccountCode.GotFocus, txtTo.GotFocus, txtCc.GotFocus, txtSubject.GotFocus, txtMessage.GotFocus
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            txt.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAccountCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                If txtAccountCode.Text <> "" Then
                    If Me.objEmail.ApakahSudahAda(Me.txtAccountCode.Text, Me.UserCode) = True Then
                        Me.txtPrimaryAccount.Text = Me.objEmail.GetPrimaryByCode(Me.txtAccountCode.Text)
                        Me.txtAccountCode.Text = Me.objEmail.GetCodeByPrimary(Me.txtPrimaryAccount.Text)
                        Me.txtAccountName.Text = Me.objEmail.GetDesc(Me.txtPrimaryAccount.Text)
                        Me.txtEmailFrom.Text = Me.objEmail.GetEmailFrom(Me.txtPrimaryAccount.Text)
                    Else
                        Me.btnFindAccount.PerformClick()
                    End If
                Else
                    Me.txtPrimaryAccount.Text = "-1"
                    Me.txtAccountCode.Text = ""
                    Me.txtAccountName.Text = ""
                    Me.txtEmailFrom.Text = ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAccountCode.Leave
        Try
            If txtAccountCode.Text <> "" Then
                If Me.objEmail.ApakahSudahAda(Me.txtAccountCode.Text, Me.UserCode) = True Then
                    Me.txtPrimaryAccount.Text = Me.objEmail.GetPrimaryByCode(Me.txtAccountCode.Text)
                    Me.txtAccountCode.Text = Me.objEmail.GetCodeByPrimary(Me.txtPrimaryAccount.Text)
                    Me.txtAccountName.Text = Me.objEmail.GetDesc(Me.txtPrimaryAccount.Text)
                    Me.txtEmailFrom.Text = Me.objEmail.GetEmailFrom(Me.txtPrimaryAccount.Text)
                Else
                    Me.btnFindAccount.PerformClick()
                End If
            Else
                Me.txtPrimaryAccount.Text = "-1"
                Me.txtAccountCode.Text = ""
                Me.txtAccountName.Text = ""
                Me.txtEmailFrom.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTo.Leave
        Try
            If txtTo.Text <> "" Then
                Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                If Expression.IsMatch(txtTo.Text) = False Then
                    MsgBox("The email address is not valid.", MsgBoxStyle.Information, "Information")
                    txtTo.Focus()
                    txtTo.SelectAll()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCc.Leave
        Try
            If txtCc.Text <> "" Then
                Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                If Expression.IsMatch(txtCc.Text) = False Then
                    MsgBox("The email address is not valid.", MsgBoxStyle.Information, "Information")
                    txtCc.Focus()
                    txtCc.SelectAll()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAccountCode.LostFocus, txtTo.LostFocus, txtCc.LostFocus, txtSubject.LostFocus, txtMessage.LostFocus
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            txt.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAccount.Click
        Dim frm As New Core_Module.FrmEmailSettings_Cari
        Try
            frm.txtCari.Text = Me.txtAccountCode.Text
            frm.IsAccountWithUserCode = True
            frm.UserCode = Me.UserCode
            frm.ShowDialog()
            If frm.primaryAccount > 0 Then
                Me.txtPrimaryAccount.Text = frm.primaryAccount
                Me.txtAccountCode.Text = Me.objEmail.GetCodeByPrimary(frm.primaryAccount)
                Me.txtAccountName.Text = Me.objEmail.GetDesc(frm.primaryAccount)
                Me.txtEmailFrom.Text = Me.objEmail.GetEmailFrom(frm.primaryAccount)
            Else
                Me.txtPrimaryAccount.Text = "-1"
                Me.txtAccountCode.Text = ""
                Me.txtAccountName.Text = ""
                Me.txtEmailFrom.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnResend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResend.Click
        Try
            If Me.txtPrimaryAccount.Text = "" Or Me.txtPrimaryAccount.Text = "-1" Then
                MsgBox("Email account cannot empty string", MsgBoxStyle.Information, "Information")
                Me.txtAccountCode.Focus()
                Exit Sub
            End If
            If Me.txtTo.Text.Trim = "" Then
                MsgBox("Receiver (to) cannot empty string", MsgBoxStyle.Information, "Information")
                Me.txtTo.Focus()
                Exit Sub
            End If
            If Me.txtSubject.Text.Trim = "" Then
                MsgBox("Subject cannot empty string", MsgBoxStyle.Information, "Information")
                Me.txtSubject.Focus()
                Exit Sub
            End If
            If Me.txtMessage.Text.Trim = "" Then
                MsgBox("Message cannot empty string", MsgBoxStyle.Information, "Information")
                Me.txtMessage.Focus()
                Exit Sub
            End If
            If MsgBox("Are you sure?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                Me.txtMessage.Focus()
                Exit Sub
            End If
            Me.objCd.KendalikanControl(Me.GroupBox1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.llbFile.Enabled = False
            Me.btnClose.Enabled = False
            Me.btnResend.Enabled = False
            Me.ControlBox = False
            Me.Cursor = Cursors.WaitCursor
            Me.vListAddr.Add(Me.txtTo.Text)
            Me.vListAttc.Add(vAttachmentFilename)
            Me.objEmailSettings.GetAllDetail(Me.txtPrimaryAccount.Text, Me.TipeTrans, "", vReplyTo, vUsername, vPassword, vSmtpServer, vSmtpPort, "", "", "", vUseSSL)
            Me.SendEmail(Me.vListAddr, Me.txtCc.Text, Me.txtSubject.Text, Me.txtMessage.Text, Me.vListAttc, Me.txtEmailFrom.Text, Me.vReplyTo, Me.vSmtpServer, Me.vSmtpPort, Me.vUsername, Me.vPassword)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#Region "email control"
    Public Sub SendEmail(ByVal vListAddress As List(Of String), ByVal vCc As String, _
            ByVal vSubject As String, ByVal vMessage As String, ByVal vListFileName As List(Of String), _
            ByVal vFrom As String, ByVal vReplyTo As String, ByVal vSmtpServer As String, ByVal vSmtpPort As String, ByVal vUsername As String, ByVal vPassword As String)
        Dim AddTo, vFileName As String
        Dim Index1 As Integer = 0
        Try
            Dim MyMailMessage As New Net.Mail.MailMessage()
            For Each AddTo In vListAddress
                MyMailMessage.To.Add(New Mail.MailAddress(AddTo))
            Next
            MyMailMessage.CC.Add(New Mail.MailAddress(vCc))
            MyMailMessage.Subject = vSubject
            MyMailMessage.Body = vMessage
            MyMailMessage.From = New Mail.MailAddress(vFrom)
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
            SMTPServer.Host = vSmtpServer
            SMTPServer.Port = vSmtpPort
            SMTPServer.EnableSsl = True
            Dim UserState As Object = MyMailMessage
            AddHandler SMTPServer.SendCompleted, AddressOf SMTPServer_OnCompleted
            SMTPServer.SendAsync(MyMailMessage, UserState)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.objCd.KendalikanControl(Me.GroupBox1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.llbFile.Enabled = True
            Me.btnClose.Enabled = True
            Me.btnResend.Enabled = True
            Me.ControlBox = True
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Public Sub SMTPServer_OnCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        Dim sent As Integer = 0
        Dim _listAddrTo As String = ""
        Try
            Dim CompletedMailMessage As Mail.MailMessage = CType(e.UserState, Mail.MailMessage)
            Dim EmailFrom As String = CompletedMailMessage.From.Address
            Dim EmailSubject As String = CompletedMailMessage.Subject
            If e.Cancelled Then
                MsgBox("SendAsync email from " + EmailFrom.ToString + " regarding " + EmailSubject.ToString + " was canceled.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If e.Error Is Nothing Then
                sent = 1
                MsgBox("Email sent", MsgBoxStyle.Information, "Information")
            Else
                sent = 0
                MsgBox("Email NOT sent", MsgBoxStyle.Exclamation, "Information")
            End If
            Me.objEmailSettings.SaveEmailLog(CDbl(Me.txtPrimaryAccount.Text), Me.TipeTrans, Me.PrimaryTransaksiAcuan, Me.DocNo, Me.txtEmailFrom.Text, Me.txtTo.Text & ";", Me.txtCc.Text, Me.txtSubject.Text, Me.txtMessage.Text, sent, Me.UserCode, Now, Me.PrimaryCustomer, vAttachmentFilename)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.objCd.KendalikanControl(Me.GroupBox1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.llbFile.Enabled = True
            Me.btnClose.Enabled = True
            Me.btnResend.Enabled = True
            Me.ControlBox = True
            Me.Cursor = Cursors.Default
        End Try
    End Sub
#End Region

    Private Sub FrmEmailSettings_LogReport_View_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.llbFile.Text = My.Computer.FileSystem.GetName(vAttachmentFilename)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub llbFile_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbFile.LinkClicked
        Try
            If My.Computer.FileSystem.FileExists(vAttachmentFilename) Then
                Process.Start(vAttachmentFilename)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class