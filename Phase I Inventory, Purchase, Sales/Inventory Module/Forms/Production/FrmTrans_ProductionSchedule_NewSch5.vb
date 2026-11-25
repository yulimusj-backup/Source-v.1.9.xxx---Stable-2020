Public Class FrmTrans_ProductionSchedule_NewSch5
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TblRef As New DseTrans_Schedule.TblDependencyDataTable
    Dim rowRef As DseTrans_Schedule.TblDependencyRow
    Dim ObjCD As New Core_Module.ClsCekData
    Dim LansungClose As Boolean = False
    Dim Buffer As String = ""

    Private Sub FrmTrans_ProductionSchedule_NewSch5_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.cboDependency.Items.Clear()
            Me.cboDependency.Items.Add("End-Start")
            Me.cboDependency.Items.Add("Start-Start")
            Me.cboDependency.Items.Add("Start-Start (Strict)")
            Me.cboDependency.SelectedIndex = 0

            Me.cboProcess.Items.Clear()
            For Each rowRef In TblRef.Rows
                Me.cboProcess.Items.Add(rowRef.Process_Code)
            Next
            Me.cboProcess.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDelay_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDelay.Enter
        Try
            Me.Buffer = Me.TxtDelay.Text

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDelay_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDelay.GotFocus
        Try
            If Me.TxtDelay.Enabled = True Then
                Me.TxtDelay.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDelay_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDelay.Leave
        Try
            Me.TxtDelay.Text = Me.TxtDelay.Text.Trim
            If Me.ObjCD.Is_Time_String(Me.TxtDelay.Text, True, True) = False Then
                Me.TxtDelay.Text = Me.Buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDelay_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDelay.LostFocus
        Try
            If Me.TxtDelay.Enabled = True Then
                Me.TxtDelay.BackColor = Me.TxtBackColor.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboProcess_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProcess.SelectedIndexChanged
        Try
            If Me.cboProcess.SelectedIndex = -1 Then
                Me.TxtDelay.Text = "00.00"
                Me.cboDependency.SelectedIndex = 0
            Else
                For Each rowRef In Me.TblRef.Rows
                    If rowRef.Process_Code = Me.cboProcess.Text Then
                        Me.cboDependency.SelectedIndex = rowRef.Tipe_Dependency
                        Me.TxtDelay.Text = Me.ObjCD.Convert_ToString(rowRef.Delay_Dependency)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_NewSch5_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.LansungClose = False Then
                If MsgBox("Cancel data change?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Else
                    e.Cancel = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If MsgBox("Accept data change?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.LansungClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class