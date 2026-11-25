Option Explicit On
Option Strict Off

Public Class FrmTransksiCreateEmail
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjEmail As New ClsTransEmail
    Dim TBL As New Core_Module.DSUser.tbluser1DataTable
    Dim ManualClosing As Boolean = False

    Public ReplYEmailPrimary As Double = -1
    Public Reply As Boolean = False
    Public ReplySucceed As Boolean = False
    Public Forward As Boolean = True
    Public UserCode As String = ""
    Public ReplySubject As String = ""
    Public ForwardSubject As String = ""
    Public ForwardEmailPrimary As Double = -1

    Private Function ApakahDataBenar() As Boolean
        Try
            'data salah bila:
            '1. tidak ada yang dikirimkan
            If Me.clbUser.CheckedItems.Count = 0 Then
                MsgBox("Please specify to whom you will send the email to.", MsgBoxStyle.Information, "Infomation")
                Return False
            End If
            '2. mengirimkan email kosong
            If Me.txtEmail.Text = "" Then
                MsgBox("Please specify the message in your email.", MsgBoxStyle.Information, "Information")
                Me.txtEmail.Focus()
                Return False
            End If
            'nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinDateKeObj()
        Try
            With Me.ObjEmail
                .Subject = Me.txtSubject.Text
                .Content = Me.txtEmail.Text
                .CreatedBy = Me.UserCode
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransksiCreateEmail_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If ManualClosing = False Then
                'konfirmasi
                If Me.txtEmail.Text = "" And Me.clbUser.CheckedItems.Count = 0 _
                And Me.txtSubject.Text = "" Then
                Else
                    Dim ij As MsgBoxResult = MsgBox("Send email before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                    If ij = MsgBoxResult.Yes Then
                        If Me.ApakahDataBenar = True Then
                            'simpan data
                            Me.SalinDateKeObj()
                            If Me.ObjEmail.SendEmail(Me.clbUser, Me.Reply, Me.ReplYEmailPrimary) = True Then
                                MsgBox("Email sent", MsgBoxStyle.Information, "Information")
                                If Me.Reply = True Then
                                    Me.ReplySucceed = True
                                Else
                                    Me.ReplySucceed = False
                                End If
                            End If
                        Else
                            e.Cancel = True
                        End If
                    ElseIf ij = MsgBoxResult.Cancel Then
                        'cancel simpan data
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransksiCreateEmail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.ObjUser.IsiListBoxUser(Me.clbUser, TBL, UserCode)
            If Me.Reply = True Then
                Dim From As String = Me.ObjEmail.GetCreatedBy(Me.ReplYEmailPrimary)
                Dim rowview As DataRowView
                Dim row As Core_Module.DSUser.tbluser1Row

                For i As Integer = 0 To Me.clbUser.Items.Count - 1
                    rowview = Me.clbUser.Items(i)
                    row = rowview.Row
                    If row.Kode_User = From Then
                        Me.clbUser.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
                    End If
                Next
                Me.txtSubject.Text = "Re: " + Me.ReplySubject
                Me.ReplySucceed = False
            ElseIf Me.Forward = True Then
                Me.txtSubject.Text = "Fwd: " + Me.ForwardSubject
                Me.ObjEmail.ShowContent(Me.txtEmail, Me.ForwardEmailPrimary)
            End If
            Me.ManualClosing = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Try
            For i As Integer = 0 To Me.clbUser.Items.Count - 1
                Me.clbUser.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Try
            For i As Integer = 0 To Me.clbUser.Items.Count - 1
                Me.clbUser.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSend.LinkClicked
        Try
            'cek apakah data benar
            If Me.ApakahDataBenar = True Then
                'konfirmasi
                If MsgBox("Send email?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    'simpan data & konfirmasi
                    Me.SalinDateKeObj()
                    If Me.ObjEmail.SendEmail(Me.clbUser, Me.Reply, Me.ReplYEmailPrimary) = True Then
                        MsgBox("Email sent", MsgBoxStyle.Information, "Information")
                        If Me.Reply = True Then
                            Me.ReplySucceed = True
                        Else
                            Me.ReplySucceed = False
                        End If
                        Me.ManualClosing = True
                        Me.Close()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbExit_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbExit.LinkClicked
        Try
            If Me.txtEmail.Text = "" And Me.txtSubject.Text = "" Then
                Me.ManualClosing = True
                Me.Close()
            Else
                'konfirmasi
                If MsgBox("Exit dan discard data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    'cancel
                    Me.ManualClosing = True
                    Me.Close()
                Else
                    'cancel - cancel
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class