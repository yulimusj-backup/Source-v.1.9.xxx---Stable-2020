Option Explicit On
Option Strict Off

Public Class FrmTransaksiPurchaseReturn_DetailAddress
    Public SupplierAddressIdentifier As String = ""
    Public TBL As DseTransaksiPurchaseReturn.DataTable_Address1DataTable
    Dim row As DseTransaksiPurchaseReturn.DataTable_Address1Row

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiPurchaseReturn_DetailAddress_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.cboAddressCode.Items.Clear()
            For Each row In Me.TBL.Rows
                Me.cboAddressCode.Items.Add(row.Address_Code)
            Next

            For i As Integer = 0 To Me.cboAddressCode.Items.Count - 1
                row = Me.TBL.Rows(i)
                If Me.row.Identifier = Me.SupplierAddressIdentifier Then
                    Me.cboAddressCode.SelectedIndex = i
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboAddressCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAddressCode.SelectedIndexChanged
        Try
            If Me.cboAddressCode.SelectedIndex <> -1 Then
                For Each row In Me.TBL.Rows
                    If row.Address_Code = Me.cboAddressCode.Text Then
                        Me.TxtAddress.Text = row.Line_1 + vbCrLf + _
                        row.Line_2 + vbCrLf + _
                        row.Line_3 + vbCrLf + _
                        row.Line_4 + vbCrLf + _
                        row.Line_5 + vbCrLf
                    End If
                Next
            Else
                Me.TxtAddress.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class