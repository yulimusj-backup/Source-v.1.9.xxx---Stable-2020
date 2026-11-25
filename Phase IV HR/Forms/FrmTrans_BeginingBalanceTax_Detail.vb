Option Explicit On
Option Strict Off

Public Class FrmTrans_BeginingBalanceTax_Detail

    Public TBL As New DseTrans_Tax_Beginning_Balance.TblSettingDetailDataTable
    Dim Row As DseTrans_Tax_Beginning_Balance.TblSettingDetailRow

    Private Sub FrmTrans_BeginingBalanceTax_Detail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.TBL.Rows.Count > 0 Then
                Row = TBL.Rows(0)
                If Row.NIK = 0 Then
                    Me.ChkShowNIK.Checked = False
                Else
                    Me.ChkShowNIK.Checked = True
                End If
                If Row.Dept = 0 Then
                    Me.ChkShowDepartment.Checked = False
                Else
                    Me.ChkShowDepartment.Checked = True
                End If
                If Row.Position = 0 Then
                    Me.ChkShowPosition.Checked = False
                Else
                    Me.ChkShowPosition.Checked = True
                End If
            End If
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkShowDepartment_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkShowDepartment.CheckedChanged
        Try
            Row = TBL.Rows(0)
            If Me.ChkShowDepartment.Checked = True Then
                Row.Dept = 1
            Else
                Row.Dept = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkShowNIK_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkShowNIK.CheckedChanged
        Try
            Row = TBL.Rows(0)
            If Me.ChkShowNIK.Checked = True Then
                Row.NIK = 1
            Else
                Row.NIK = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkShowPosition_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkShowPosition.CheckedChanged
        Try
            Row = TBL.Rows(0)
            If Me.ChkShowPosition.Checked = True Then
                Row.Position = 1
            Else
                Row.Position = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class