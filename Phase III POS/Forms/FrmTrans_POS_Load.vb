Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmTrans_POS_Payment_Load
    Public mCashierCode As String = ""
    Public mExit As Boolean
    Public mPrimarySession As Double = -1

    Dim ObjSession As New ClsTrans_Session
    Dim ObjPOS As New ClsTrans_POS
    Dim mrOK As Boolean = False
    Dim DTA As New DseTrans_SessionTableAdapters.tbltrans_session_cabang1_3TableAdapter
    Dim TBL As New DseTrans_Session.tbltrans_session_cabang1_3DataTable

    Private Sub FrmTrans_POS_Payment_Load_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Hasil As MsgBoxResult
        Try
            If mrOK = False Then
                Hasil = MsgBox("Do you want to close from POS transaction?", MsgBoxStyle.YesNo, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    Me.mExit = True
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_POS_Payment_Detail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.CBOSession_ID.DataSource = Nothing
            Me.CBOSession_ID.Items.Clear()
            ObjSession.IsiCBOSessionID(mCashierCode, Me.CBOSession_ID, TBL, DTA)
            Me.CBOSession_ID.SelectedIndex = 0
            Me.CBOSession_ID.Focus()
            Me.mExit = False
            Me.mrOK = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CBOSession_ID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBOSession_ID.SelectedIndexChanged
        Dim row As DseTrans_Session.tbltrans_session_cabang1_3Row
        Try
            If Me.CBOSession_ID.SelectedIndex = -1 Then
                Me.TxtPrimarySession.Text = "-1"
            Else
                row = Me.TBL.DefaultView(Me.CBOSession_ID.SelectedIndex).Row
                Me.TxtPrimarySession.Text = row.primary
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            mPrimarySession = Me.TxtPrimarySession.Text

            Me.mrOK = True
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class