Option Explicit On
Option Strict Off

Public Class FrmMasterWorkShift_AdvanceSettings

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""
    Public ReadOnly_ As Boolean = True
    Public ObjWS As New ClsMasterWorkShift
    Public ObjEmployee As New ClsMasterKaryawan
    Public CalledFromWS As Boolean = True
    Dim objCd As New Core_Module.ClsCekData

    Public disabledEarlyOT As Boolean = False
    Public disabledLate As Boolean = False
    Public disabledEarlyOut As Boolean = False
    Public disabledOT As Boolean = False

    Private Sub FrmMasterWorkShift_AdvanceSettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If CalledFromWS = True Then
                With Me.ObjWS
                    If .CalculateEarlyOTOnWeekend = 1 Then
                        Me.chkCalculateEarlyOTOnWeekend.Checked = True
                    Else
                        Me.chkCalculateEarlyOTOnWeekend.Checked = False
                    End If
                    If .CalculateLateOnWeekend = 1 Then
                        Me.chkCalculateLateOnWeekend.Checked = True
                    Else
                        Me.chkCalculateLateOnWeekend.Checked = False
                    End If
                    If .CalculateEarlyOutOnWeekend = 1 Then
                        Me.chkCalculateEarlyOutOnWeekend.Checked = True
                    Else
                        Me.chkCalculateEarlyOutOnWeekend.Checked = False
                    End If
                    If .CalculateOTOnWeekend = 1 Then
                        Me.chkCalculateOTOnWeekend.Checked = True
                    Else
                        Me.chkCalculateOTOnWeekend.Checked = False
                    End If
                    If .CalculateEarlyOTOnHoliday = 1 Then
                        Me.chkCalculateEarlyOTOnHoliday.Checked = True
                    Else
                        Me.chkCalculateEarlyOTOnHoliday.Checked = False
                    End If
                    If .CalculateLateOnHoliday = 1 Then
                        Me.chkCalculateLateOnHoliday.Checked = True
                    Else
                        Me.chkCalculateLateOnHoliday.Checked = False
                    End If
                    If .CalculateEarlyOutOnHoliday = 1 Then
                        Me.chkCalculateEarlyOutOnHoliday.Checked = True
                    Else
                        Me.chkCalculateEarlyOutOnHoliday.Checked = False
                    End If
                    If .CalculateOTOnHoliday = 1 Then
                        Me.chkCalculateOTOnHoliday.Checked = True
                    Else
                        Me.chkCalculateOTOnHoliday.Checked = False
                    End If
                End With
            Else
                With Me.ObjEmployee
                    If .CalculateEarlyOTOnWeekend = 1 Then
                        Me.chkCalculateEarlyOTOnWeekend.Checked = True
                    Else
                        Me.chkCalculateEarlyOTOnWeekend.Checked = False
                    End If
                    If .CalculateLateOnWeekend = 1 Then
                        Me.chkCalculateLateOnWeekend.Checked = True
                    Else
                        Me.chkCalculateLateOnWeekend.Checked = False
                    End If
                    If .CalculateEarlyOutOnWeekend = 1 Then
                        Me.chkCalculateEarlyOutOnWeekend.Checked = True
                    Else
                        Me.chkCalculateEarlyOutOnWeekend.Checked = False
                    End If
                    If .CalculateOTOnWeekend = 1 Then
                        Me.chkCalculateOTOnWeekend.Checked = True
                    Else
                        Me.chkCalculateOTOnWeekend.Checked = False
                    End If
                    If .CalculateEarlyOTOnHoliday = 1 Then
                        Me.chkCalculateEarlyOTOnHoliday.Checked = True
                    Else
                        Me.chkCalculateEarlyOTOnHoliday.Checked = False
                    End If
                    If .CalculateLateOnHoliday = 1 Then
                        Me.chkCalculateLateOnHoliday.Checked = True
                    Else
                        Me.chkCalculateLateOnHoliday.Checked = False
                    End If
                    If .CalculateEarlyOutOnHoliday = 1 Then
                        Me.chkCalculateEarlyOutOnHoliday.Checked = True
                    Else
                        Me.chkCalculateEarlyOutOnHoliday.Checked = False
                    End If
                    If .CalculateOTOnHoliday = 1 Then
                        Me.chkCalculateOTOnHoliday.Checked = True
                    Else
                        Me.chkCalculateOTOnHoliday.Checked = False
                    End If
                End With
            End If
            
            If Me.ReadOnly_ = True Then
                For Each control As Control In Me.Controls
                    If control.Name.StartsWith("chk") = True Or control.Name.StartsWith("btnOK") = True Then
                        control.Enabled = False
                    End If
                Next
            End If

            If Me.disabledEarlyOT = True Then
                Me.chkCalculateEarlyOTOnWeekend.Enabled = False
                Me.chkCalculateEarlyOTOnHoliday.Enabled = False
            End If
            If Me.disabledLate = True Then
                Me.chkCalculateLateOnWeekend.Enabled = False
                Me.chkCalculateLateOnHoliday.Enabled = False
            End If
            If Me.disabledEarlyOut = True Then
                Me.chkCalculateEarlyOutOnWeekend.Enabled = False
                Me.chkCalculateEarlyOutOnHoliday.Enabled = False
            End If
            If Me.disabledOT = True Then
                Me.chkCalculateOTOnWeekend.Enabled = False
                Me.chkCalculateOTOnHoliday.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterWorkShift_AdvanceSettings_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.ReadOnly_ = False And Me.DialogResult <> Windows.Forms.DialogResult.OK Then
                If MsgBox("Close and discard data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Else
                    e.Cancel = True
                End If
            ElseIf Me.ReadOnly_ = False Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
            Else
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.ReadOnly_ = False Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
            Else
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class