Option Explicit On
Option Strict Off
Public Class FrmTransaksiMonitor_Production_Request
    Public userIdentifier As Double = 0
    Public identifierTarget As Double = 0
    Public requestType As ClsTransaksiMonitor_Production.enumRequestType = ClsTransaksiMonitor_Production.enumRequestType.TAKE
    Public userTarget As String = "-"
    Public answer As ClsTransaksiMonitor_Production.enumAnswer = ClsTransaksiMonitor_Production.enumAnswer.REJECT
    Public answerNote As String = ""
    Dim objMySql As New Core_Module.ClsMySQLReference
    Dim answerMonitorInterval As Double = 1000
    Dim timeoutInterval As Double = 5 * 60000
    Dim objCd As New Core_Module.ClsCekData
    Dim cancelClose As Boolean = False

    Private Function createRequest() As Boolean
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_handshakeTableAdapter
        Try
            If Me.isExistRequest = True Then
                Me.answerNote = "Request is already exist in current time, please try again until other request is finished."
                Return False
            End If
            d.Connection.ConnectionString = Core_Module.GetConstring
            If d.Insert(Me.userIdentifier, Me.identifierTarget, Me.requestType, ClsTransaksiMonitor_Production.enumAnswer.WAITING, _
                    Now, Me.objMySql.GetServerTime.AddMilliseconds(Me.timeoutInterval), "") <= 0 Then
                Me.answerNote = "Failed to create request."
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function isExistRequest() As Boolean
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_handshakeTableAdapter
        Dim n As New Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            n = d.ScalarQuery_GetJmlInRequest()
            If n.HasValue Then
                If n.Value > 0 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        End Try
    End Function

    Private Sub FrmTransaksiMonitor_Production_Request_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.cancelClose = True Then
                'Me.Button2.PerformClick()
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiMonitor_Production_Request_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        FrmTransaksiMonitor_Production_Request.CheckForIllegalCrossThreadCalls = False
        Try
            Me.Label1.Text = "Please wait until " & userTarget & " answer your request."
            If createRequest() = True Then
                'Me.BackgroundWorker1.RunWorkerAsync()
                Me.Bgw1()
            Else
                Me.answer = ClsTransaksiMonitor_Production.enumAnswer.FAIL_CREATE_REQUEST
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub timerAnswer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles timerAnswer.Tick
        Try
            checkAnswer()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub checkAnswer()
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_handshakeTableAdapter
        Dim t As New DseTransaksiMonitor_Production.tbltrans_prod_handshakeDataTable
        Dim r As DseTransaksiMonitor_Production.tbltrans_prod_handshakeRow
        Dim n As New Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            If d.FillBy_Request(t, Me.userIdentifier, Me.identifierTarget, Me.requestType) > 0 Then
                r = t.Rows(0)
                If r.answer = ClsTransaksiMonitor_Production.enumAnswer.ACCEPT Then
                    Me.answer = ClsTransaksiMonitor_Production.enumAnswer.ACCEPT
                    Me.answerNote = r.answer_note
                    'Me.BackgroundWorker1.CancelAsync()
                    Me.BgwStop()
                    Me.Close()
                ElseIf r.answer = ClsTransaksiMonitor_Production.enumAnswer.REJECT Then
                    Me.answer = ClsTransaksiMonitor_Production.enumAnswer.REJECT
                    Me.answerNote = r.answer_note
                    'Me.BackgroundWorker1.CancelAsync()
                    Me.BgwStop()
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Me.objCd.UpdateThreading()
            Me.timerAnswer.Interval = Me.answerMonitorInterval
            Me.timerAnswer.Start()

            Me.timeoutTimer.Interval = Me.timeoutInterval
            Me.timeoutTimer.Start()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Bgw1()
        Try
            Me.objCd.UpdateThreading()
            Me.timerAnswer.Interval = Me.answerMonitorInterval
            Me.timerAnswer.Start()

            Me.timeoutTimer.Interval = Me.timeoutInterval
            Me.timeoutTimer.Start()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If MsgBox("Are you sure to cancel request?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                'Me.BackgroundWorker1.CancelAsync()
                Me.BgwStop()
                Me.answer = ClsTransaksiMonitor_Production.enumAnswer.USER_CANCEL
                Me.cancelClose = False
                Me.Close()
            Else
                Me.cancelClose = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub clearRequest()
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_handshakeTableAdapter
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            d.Delete(Me.userIdentifier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub timeoutTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles timeoutTimer.Tick
        Try
            'Me.BackgroundWorker1.CancelAsync()
            Me.BgwStop()
            Me.answerNote = "Request timeout [5 minutes]."
            Me.answer = ClsTransaksiMonitor_Production.enumAnswer.TIMEOUT
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.timerAnswer.Stop()
            Me.clearRequest()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwStop()
        Try
            Me.timerAnswer.Stop()
            Me.clearRequest()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class