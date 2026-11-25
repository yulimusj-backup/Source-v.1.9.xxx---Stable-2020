Option Explicit On
Option Strict Off

Public Class FrmTransaksiMonitor_Production
    Public userIdentifier_1 As Double = 0
    Public userCode As String = ""
    Public UserType As Core_Module.ClsTransaksiMonitor_Production.enumUserType = ClsTransaksiMonitor_Production.enumUserType.READ_ONLY

    Dim refreshInterval As Double = 10 * 1000.0
    Dim objUser As New Core_Module.ClsUser
    Dim objMonitor As New ClsTransaksiMonitor_Production
    Dim BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub FrmTransaksiMonitor_Production_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.loadMonitor()
            Me.ProgressBar1.Maximum = Me.refreshInterval
            Me.Timer1.Start()
            Me.SetHakAkses()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.objMonitor.flushMonitor_Tick()
            Me.loadMonitor()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        Try
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As New FrmTransaksiMonitor_Production_Request
        Dim identifierTarget As Double = 0
        Dim userTarget As String = ""
        Dim UserTypeTarget As Core_Module.ClsTransaksiMonitor_Production.enumUserType = Core_Module.ClsTransaksiMonitor_Production.enumUserType.READ_ONLY

        Dim answer As ClsTransaksiMonitor_Production.enumAnswer = ClsTransaksiMonitor_Production.enumAnswer.WAITING
        Try
            If e.ColumnIndex <> Me.ColTake.Index And e.ColumnIndex <> Me.ColSend.Index Then
                Exit Sub
            End If
            identifierTarget = Me.DataGridView1.Item(Me.ColIdentifier.Index, e.RowIndex).Value
            userTarget = Me.DataGridView1.Item(Me.ColUser.Index, e.RowIndex).Value
            UserTypeTarget = Me.DataGridView1.Item(Me.ColUserType.Index, e.RowIndex).Value

            If identifierTarget = Me.userIdentifier_1 Then
                MsgBox("You cannot take or send control to yourself.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColTake.Index Then
                If Me.objMonitor.getUserControl(Me.userIdentifier_1, Me.UserType) = ClsTransaksiMonitor_Production.enumUserControl.FULL_CONTROL Then
                    MsgBox("You can only take contol if your status is 'Read Only'.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If Me.objMonitor.getUserControl(identifierTarget, Me.UserType) = ClsTransaksiMonitor_Production.enumUserControl.READ_ONLY Then
                    MsgBox("You can only take contol from user with 'Full Control' status.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                i.requestType = ClsTransaksiMonitor_Production.enumRequestType.TAKE
            End If
            If e.ColumnIndex = Me.ColSend.Index Then
                If Me.objMonitor.getUserControl(Me.userIdentifier_1, Me.UserType) = ClsTransaksiMonitor_Production.enumUserControl.READ_ONLY Then
                    MsgBox("You can only send contol if your status is 'Full Control'.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If Me.objMonitor.getUserControl(identifierTarget, UserTypeTarget) = ClsTransaksiMonitor_Production.enumUserControl.FULL_CONTROL Then
                    MsgBox("You can only send contol to user with 'Read Only' status.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If UserTypeTarget = ClsTransaksiMonitor_Production.enumUserType.READ_ONLY Then
                    MsgBox("You can only send contol to user with access to create production schedule.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                i.requestType = ClsTransaksiMonitor_Production.enumRequestType.SEND
            End If
            i.userIdentifier = Me.userIdentifier_1
            i.identifierTarget = identifierTarget
            i.userTarget = userTarget
            i.ShowDialog()
            If i.answer = ClsTransaksiMonitor_Production.enumAnswer.FAIL_CREATE_REQUEST Then
                If String.IsNullOrEmpty(i.answerNote) = False Then
                    MsgBox(i.answerNote, MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Fail to create request.", MsgBoxStyle.Information, "Information")
                End If
            ElseIf i.answer = ClsTransaksiMonitor_Production.enumAnswer.USER_CANCEL Then
                MsgBox("Request canceled by user.", MsgBoxStyle.Information, "Information")
            ElseIf i.answer = ClsTransaksiMonitor_Production.enumAnswer.TIMEOUT Then
                If String.IsNullOrEmpty(i.answerNote) = False Then
                    MsgBox(i.answerNote, MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Request timeout.", MsgBoxStyle.Information, "Information")
                End If
            ElseIf i.answer = ClsTransaksiMonitor_Production.enumAnswer.REJECT Then
                If String.IsNullOrEmpty(i.answerNote) = False Then
                    MsgBox("Your request was rejected, partner message: " & vbNewLine & i.answerNote, MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Your request was rejected.", MsgBoxStyle.Information, "Information")
                End If
            ElseIf i.answer = ClsTransaksiMonitor_Production.enumAnswer.ACCEPT Then
                Me.swithUserControl(Me.userIdentifier_1, identifierTarget)
                If String.IsNullOrEmpty(i.answerNote) = False Then
                    MsgBox("Your request was accepted, partner message: " & vbNewLine & i.answerNote, MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Your request was accepted.", MsgBoxStyle.Information, "Information")
                End If
                If i.requestType = ClsTransaksiMonitor_Production.enumRequestType.TAKE Then
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Me.Visible = True
        End Try
    End Sub

    Private Sub swithUserControl(ByVal vUserIdentifier As Double, ByVal vIdentifierTarget As Double)
        Dim vPrimaryUser As Double = -1
        Dim vPrimaryTarget As Double = -1
        Try
            vPrimaryUser = Me.objMonitor.getPrimary(vUserIdentifier)
            vPrimaryTarget = Me.objMonitor.getPrimary(vIdentifierTarget)
            Me.objMonitor.updatePrimary(-1, vUserIdentifier)
            Me.objMonitor.updatePrimary(vPrimaryUser, vIdentifierTarget)
            Me.objMonitor.updatePrimary(vPrimaryTarget, vUserIdentifier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub loadMonitor()
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_monitorTableAdapter
        Dim t As New DseTransaksiMonitor_Production.tbltrans_prod_monitorDataTable
        Dim r As DseTransaksiMonitor_Production.tbltrans_prod_monitorRow
        Dim dgvRow As DataGridViewRow

        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DataGridView1.Rows.Clear()
            If d.Fill(t) > 0 Then
                For Each r In t
                    Me.DataGridView1.Rows.Add()
                    dgvRow = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1)
                    With dgvRow
                        .Cells(Me.ColNo.Index).Value = Me.DataGridView1.Rows.Count.ToString
                        .Cells(Me.ColCoputerID.Index).Value = r.computer_name
                        .Cells(Me.ColIdentifier.Index).Value = r.identifier.ToString
                        .Cells(Me.ColIPAddress.Index).Value = r.ip_address
                        .Cells(Me.ColPrimary.Index).Value = r.primary
                        If Me.objMonitor.getUserControl(r.identifier, r.user_type) = ClsTransaksiMonitor_Production.enumUserControl.FULL_CONTROL Then
                            .Cells(Me.ColStatus.Index).Value = "Full Control"
                        Else
                            .Cells(Me.ColStatus.Index).Value = "Read Only"
                        End If
                        .Cells(Me.ColUser.Index).Value = r.user_code
                        If r.user_code = Me.userCode Then
                            .Cells(Me.ColUser.Index).Value = .Cells(Me.ColUser.Index).Value '& " [YOU]"
                        End If
                        .Cells(Me.ColUserType.Index).Value = r.user_type   ' objMonitor.GetUserTypeStr(r.user_type)
                    End With
                Next
            End If
            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.ProgressBar1.Value = 0
        End Try
    End Sub

    Private Sub SetHakAkses()
        Dim objUser As New Core_Module.ClsUser
        Try
            If objUser.ApakahMemilikiHakAkses(Me.userCode, "FrmTrans_ProductionSchedule_SendTakeControl") = False Then
                Me.ColSend.Visible = False
                Me.ColTake.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Me.ProgressBar1.Value += 100
            If Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                Me.ProgressBar1.Value = 0
                Me.loadMonitor()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class