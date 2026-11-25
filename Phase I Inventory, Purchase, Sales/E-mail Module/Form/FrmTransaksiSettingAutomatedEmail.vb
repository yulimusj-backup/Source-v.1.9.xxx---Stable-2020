Option Explicit On
Option Strict Off

Public Class FrmTransaksiSettingAutomatedEmail
    Dim ObjSetting As New ClsTransaksiEmailAutomated
    Dim TBLUser As New Core_Module.DSUser.tbluser1DataTable
    Dim TBLEmail As New DsTransaksiEmail.tbltrans_email_automatedDataTable

    Private Sub FrmTransaksiSettingAutomatedEmail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ObjUser As New Core_Module.ClsUser

        Try
            ObjUser.IsiListBoxUser(Me.clbUser, Me.TBLUser, "")
            Me.ObjSetting.IsiListBox(Me.ListBoxSubject, Me.TBLEmail)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUser = Nothing
        End Try
    End Sub

    Private Sub ListBoxSubject_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBoxSubject.SelectedIndexChanged
        Try
            If Me.ListBoxSubject.SelectedIndex <> -1 Then
                Dim dataview As Data.DataRowView = Me.ListBoxSubject.SelectedItems(0)
                Dim row As DsTransaksiEmail.tbltrans_email_automatedRow = dataview.Row
                Me.ObjSetting.CariDetail(row.Primary, Me.clbUser)
                Me.txtDescription.Text = Me.ObjSetting.Description
            Else
                Me.txtDescription.Text = ""
                While Me.clbUser.CheckedItems.Count > 0
                    Me.clbUser.SetItemCheckState(Me.clbUser.CheckedIndices(0), _
                    Windows.Forms.CheckState.Unchecked)
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Dim dataview As Data.DataRowView = Me.ListBoxSubject.SelectedItems(0)
                Dim row As DsTransaksiEmail.tbltrans_email_automatedRow = dataview.Row
                If Me.ObjSetting.SimpanData(row.Primary, Me.clbUser) = True Then
                    MsgBox("Data saved", MsgBoxStyle.Information, "Error")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
End Class