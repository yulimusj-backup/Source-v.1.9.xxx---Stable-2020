Option Strict Off
Option Explicit On

Public Class FrmDetailAddress
    Public PrimaryCust As Double = -1
    Public vAddIdentifier As String = ""
    Public vAddIdentifierResult As String = ""

    Dim ObjCust As New clsMasterCustomer

    Private Sub FrmDetailAddress_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'iskan combo box
            Me.ObjCust.IsiCboAddress(Me.cboCustAddCode, Me.PrimaryCust)
            If Me.vAddIdentifier <> "" Then
                'tandai kode cust dan nama cust
                Me.Text = "Detail Address: " + Me.ObjCust.GetCustomerCode(Me.PrimaryCust) + "-" + Me.ObjCust.GetCustName(Me.PrimaryCust)
                'cari data item yang benar
                For i As Integer = 0 To Me.cboCustAddCode.Items.Count
                    If Me.vAddIdentifier = Me.ObjCust.GetIdentifierAdd(Me.PrimaryCust, _
                    Me.cboCustAddCode.Items(i)) Then
                        Me.cboCustAddCode.SelectedIndex = i
                        Exit Sub
                    End If
                Next
                'kalo ga ketemu data yang benar maka keluar
                If Me.cboCustAddCode.Items.Count > 0 Then
                    Me.cboCustAddCode.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCustAddCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCustAddCode.SelectedIndexChanged
        Try
            'ganti add identifier dan text box
            If Me.cboCustAddCode.SelectedIndex <> -1 Then
                Me.vAddIdentifierResult = Me.ObjCust.GetIdentifierAdd(Me.PrimaryCust, _
                Me.cboCustAddCode.Text)
                Me.ObjCust.IsiTxtAddress(Me.vAddIdentifierResult, Me.PrimaryCust, Me.txtAdd)
            Else
                Me.txtAdd.Text = ""
                Me.vAddIdentifierResult = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
End Class