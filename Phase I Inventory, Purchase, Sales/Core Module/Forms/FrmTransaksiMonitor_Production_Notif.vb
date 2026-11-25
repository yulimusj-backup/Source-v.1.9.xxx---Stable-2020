Option Explicit On
Option Strict Off

Public Class FrmTransaksiMonitor_Production_Notif

    Dim objMySQL As New Core_Module.ClsMySQLReference
    Dim cancelClose As Boolean = True
    Public expiredTime As Date = Now
    Public userIdentifier As Double = -1
    Public identifierTarget As Double = -1
    Public requestType As ClsTransaksiMonitor_Production.enumRequestType = ClsTransaksiMonitor_Production.enumRequestType.SEND

    Private Sub btnReject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReject.Click
        Try
            If MsgBox("Are you sure to reject this request?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.ResponsRequest(ClsTransaksiMonitor_Production.enumAnswer.REJECT, Me.txtMessage.Text)
                Me.cancelClose = False
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnAccept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAccept.Click
        Try
            If MsgBox("Are you sure to accept this request?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.ResponsRequest(ClsTransaksiMonitor_Production.enumAnswer.ACCEPT, Me.txtMessage.Text)
                Me.cancelClose = False
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiMonitor_Production_Notif_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.cancelClose = True Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiMonitor_Production_Notif_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'Me.NotifyIcon1.Visible = True
            'Me.NotifyIcon1.Icon = Me.Icon
            'Me.NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
            'Me.NotifyIcon1.BalloonTipText = Me.TextBox1.Text
            'Me.NotifyIcon1.ShowBalloonTip(50000)

            Me.cancelClose = True
            Me.txtMessage.Clear()
            Me.txtMessage.Focus()
            Me.txtMessage.Select()
            Me.timeoutTimer.Start()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub timeoutTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles timeoutTimer.Tick
        Try
            If Me.isStillExistRequest(Me.userIdentifier, Me.identifierTarget, Me.requestType) = False Then 'Me.objMySQL.GetServerTime > Me.expiredTime
                Me.cancelClose = False
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ResponsRequest(ByVal vAnswer As ClsTransaksiMonitor_Production.enumAnswer, ByVal vAnswerNote As String) As Boolean
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_handshakeTableAdapter
        Try
            If String.IsNullOrEmpty(vAnswerNote) Then
                vAnswerNote = ""
            End If
            d.Connection.ConnectionString = Core_Module.GetConstring
            d.Update(vAnswer, Now, vAnswerNote)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function isStillExistRequest(ByVal vUserIdentifier As Double, ByVal vIdentifierTarget As Double, ByVal vReqType As ClsTransaksiMonitor_Production.enumRequestType) As Boolean
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_handshakeTableAdapter
        Dim n As New Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            n = d.ScalarQuery_CheckExistRequest(vUserIdentifier, vIdentifierTarget, vReqType)
            If n.HasValue Then
                If n.Value > 0 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub NotifyIcon1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        Try
            Me.WindowState = FormWindowState.Normal
            Me.NotifyIcon1.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class