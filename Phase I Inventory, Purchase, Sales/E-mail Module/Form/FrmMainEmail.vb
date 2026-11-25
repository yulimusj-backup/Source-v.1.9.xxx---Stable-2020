Option Explicit On
Option Strict Off

Public Class FrmMainEmail
    Public UserCode As String
    Dim ObjEmail As New ClsTransEmail
    Dim TBL As New DsTransaksiEmail.tbltrans_email1DataTable

    Private Sub llbCreate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCreate.LinkClicked
        Dim hj As New FrmTransksiCreateEmail
        Try
            hj.UserCode = Me.UserCode
            hj.Reply = False
            hj.Forward = False
            hj.ShowDialog()
            hj.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmail_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvEmail.RowsAdded
        Try
            If Me.dgvEmail.Rows.Count > 0 Then
                Dim row As Windows.Forms.DataGridViewRow
                Me.dgvEmail.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None

                If Me.dgvEmail.Columns.Count = 7 Then
                    For Each row In Me.dgvEmail.Rows
                        If row.Cells(3).Value = 0 Then
                            row.Cells("Image").Value = My.Resources.email_24
                        ElseIf row.Cells(3).Value = 2 Then
                            row.Cells("Image").Value = My.Resources.MessageQueuing
                        Else
                            row.Cells("Image").Value = My.Resources.page
                        End If
                    Next
                End If

                Me.dgvEmail.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiInbox()
        Try
            Me.gboItem.Text = "Inbox"
            Me.dgvEmail.ClearSelection()
            Me.ObjEmail.IsiDGV(Me.UserCode, ClsTransEmail.EnumFolder.Inbox, Me.dgvEmail, TBL)
            Dim i As New Object
            Dim e As New Windows.Forms.DataGridViewRowsAddedEventArgs(0, 0)
            Me.dgvEmail_RowsAdded(i, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiSentItem()
        Try
            Me.gboItem.Text = "Sent Items"
            Me.ObjEmail.IsiDGV(Me.UserCode, ClsTransEmail.EnumFolder.SentItem, Me.dgvEmail, TBL)
            Dim i As New Object
            Dim e As New Windows.Forms.DataGridViewRowsAddedEventArgs(0, 0)
            Me.dgvEmail_RowsAdded(i, e)
            Me.dgvEmail.ClearSelection()
            If Me.dgvEmail.RowCount > 0 Then
                Me.dgvEmail.Rows(0).Selected = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMainEmail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'isikan folder inbox untuk user ybs 
            If Me.UserCode <> "" Then
                Me.IsiInbox()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbInbox_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbInbox.LinkClicked
        Me.IsiInbox()
    End Sub

    Private Sub llbSentItem_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSentItem.LinkClicked
        Me.IsiSentItem()
    End Sub

    Private Sub llbclearInbox_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbclearInbox.LinkClicked
        Try
            If MsgBox("Clear your inbox?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.ObjEmail.HapusEmail(Me.UserCode, ClsTransEmail.EnumFolder.Inbox)
                Me.IsiInbox()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbClearSentItem_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbClearSentItem.LinkClicked
        Try
            If MsgBox("Clear your sent items folder?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.ObjEmail.HapusEmail(Me.UserCode, ClsTransEmail.EnumFolder.SentItem)
                Me.IsiInbox()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmail_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvEmail.SelectionChanged
        Try
            If Me.dgvEmail.SelectedRows.Count = 0 Then
                'tampilkan text box
                Me.txtEmail.Text = ""
                'aktifkan llb
                Me.llbDelete.Enabled = False
                Me.llbForward.Enabled = False
                Me.llbReply.Enabled = False
                Me.llbRespond.Enabled = False
                Me.llbRespond2.Enabled = False
            Else
                'tampilkan text box
                Me.ObjEmail.ShowContent(Me.txtEmail, _
                Me.dgvEmail.Item("Primary", Me.dgvEmail.SelectedRows(0).Index).Value)
                'aktifkan llb
                Me.llbDelete.Enabled = True
                Me.llbForward.Enabled = True
                If Me.gboItem.Text = "Inbox" Then
                    Me.llbReply.Enabled = True
                Else
                    Me.llbReply.Enabled = False
                End If
                Me.llbRespond.Enabled = True
                Me.llbRespond2.Enabled = True
                Dim Status As ClsTransEmail.EnumEmailStatus = Me.dgvEmail.Item("Status", _
                Me.dgvEmail.SelectedRows(0).Index).Value
                If Status = ClsTransEmail.EnumEmailStatus.NewEmail And Me.dgvEmail.Columns.Count = 7 Then
                    'ubah icon
                    Status = ClsTransEmail.EnumEmailStatus.AlreadyRead
                    Me.dgvEmail.Item("Status", Me.dgvEmail.SelectedRows(0).Index).Value = Status
                    Me.dgvEmail.Item("Image", Me.dgvEmail.SelectedRows(0).Index).Value = My.Resources.page
                    'ubah status
                    Me.ObjEmail.UpDateStatus(ClsTransEmail.EnumEmailStatus.AlreadyRead, _
                    Me.dgvEmail.Item("Primary", Me.dgvEmail.SelectedRows(0).Index).Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbDelete_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbDelete.LinkClicked
        Try
            'hapus data email yang sedang difokuskan pada dgv
            If Me.dgvEmail.SelectedRows.Count > 0 Then
                If MsgBox("Delete email?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.ObjEmail.HapusEmail(CDbl(Me.dgvEmail.Item("Primary", Me.dgvEmail.SelectedRows(0).Index).Value))
                    If Me.gboItem.Text = "Inbox" Then
                        Me.llbInbox_LinkClicked(sender, e)
                    Else
                        Me.llbSentItem_LinkClicked(sender, e)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbReply_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbReply.LinkClicked
        Try
            If Me.dgvEmail.SelectedRows.Count > 0 And Me.gboItem.Text = "Inbox" Then
                Dim hj As New FrmTransksiCreateEmail
                hj.UserCode = Me.UserCode
                hj.Reply = True
                hj.ReplYEmailPrimary = Me.dgvEmail.Item("Primary", Me.dgvEmail.SelectedRows(0).Index).Value
                hj.ReplySubject = Me.dgvEmail.Item("Subject", Me.dgvEmail.SelectedRows(0).Index).Value
                hj.ShowDialog()
                'ubah status sudah di reply
                Dim Status As ClsTransEmail.EnumEmailStatus = Me.dgvEmail.Item("Status", _
                Me.dgvEmail.SelectedRows(0).Index).Value
                If hj.ReplySucceed = True Then
                    'ubah icon
                    Status = ClsTransEmail.EnumEmailStatus.Replied
                    Me.dgvEmail.Item("Status", Me.dgvEmail.SelectedRows(0).Index).Value = Status
                    Me.dgvEmail.Item("Image", Me.dgvEmail.SelectedRows(0).Index).Value = My.Resources.MessageQueuing
                    'ubah status
                    Me.ObjEmail.UpDateStatus(ClsTransEmail.EnumEmailStatus.Replied, _
                    Me.dgvEmail.Item("Primary", Me.dgvEmail.SelectedRows(0).Index).Value)
                End If
                hj.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbForward_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbForward.LinkClicked
        Try
            If Me.dgvEmail.SelectedRows.Count > 0 Then
                Dim hj As New FrmTransksiCreateEmail
                hj.Reply = False
                hj.UserCode = Me.UserCode
                hj.Forward = True
                hj.ForwardEmailPrimary = Me.dgvEmail.Item("Primary", Me.dgvEmail.SelectedRows(0).Index).Value
                hj.ForwardSubject = Me.dgvEmail.Item("Subject", Me.dgvEmail.SelectedRows(0).Index).Value
                hj.ShowDialog()
                hj.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbClearAlreadyRead_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbClearAlreadyRead.LinkClicked
        Try
            If MsgBox("Clear your already read email?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.ObjEmail.HapusEmailSudahDibaca(Me.UserCode)
                Me.IsiInbox()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class