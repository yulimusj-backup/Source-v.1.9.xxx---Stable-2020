Option Strict Off
Option Explicit On

Public Class FrmTransaksiPO_DetailShipTo
    Public vAddIdentifier As String = ""
    Public vAddIdentifierResult As String = ""

    Dim ObjCompany As New Core_Module.ClsMasterCompanyData
    Dim TBL As New Core_Module.DseMasterCompanyData.tbltrans_company_data_itemDataTable

    Private Sub FrmDetailAddress_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer = 0
        Dim row As Core_Module.DseMasterCompanyData.tbltrans_company_data_itemRow

        Try
            'iskan combo box
            Me.ObjCompany.IsiCboAdd(Me.TBL, Me.cboCustAddCode)
            If Me.vAddIdentifier <> "" Then
                'tandai kode cust dan nama cust
                Me.Text = "Detail Ship To Address"

                'cari data item yang benar
                While i <= Me.TBL.DefaultView.Count - 1
                    row = Me.TBL.DefaultView(i).Row
                    If row.Add_Identifier = Me.vAddIdentifier Then
                        Me.cboCustAddCode.SelectedIndex = i
                        Exit Sub
                    End If
                    i += 1
                End While

                'kalo ga ketemu data yang benar maka keluar
                If Me.cboCustAddCode.Items.Count > 0 Then
                    Me.cboCustAddCode.SelectedIndex = 0
                End If
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

    Private Sub cboCustAddCode_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCustAddCode.SelectedValueChanged
        Dim row As Core_Module.DseMasterCompanyData.tbltrans_company_data_itemRow

        Try
            'ganti add identifier dan text box
            If Me.cboCustAddCode.SelectedIndex <> -1 Then
                row = Me.TBL.DefaultView(Me.cboCustAddCode.SelectedIndex).Row
                Me.TxtIdentifier.Text = row.Add_Identifier
                Me.txtAdd.Text = Me.ObjCompany.GetAddress(Me.TxtIdentifier.Text)
            Else
                Me.TxtIdentifier.Text = ""
                Me.txtAdd.Text = ""
                Me.vAddIdentifierResult = ""
            End If
            Me.vAddIdentifierResult = Me.TxtIdentifier.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class