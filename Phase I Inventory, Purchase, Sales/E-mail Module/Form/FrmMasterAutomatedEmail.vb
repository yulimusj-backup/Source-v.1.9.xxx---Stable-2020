Option Explicit On
Option Strict Off

Public Class FrmMasterAutomatedEmail

    Private Sub Tbltrans_email_automatedBindingNavigatorSaveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Validate()
        Me.Tbltrans_email_automatedBindingSource.EndEdit()
        Me.Tbltrans_email_automatedTableAdapter.Update(Me.DsTransaksiEmail.tbltrans_email_automated)

    End Sub

    Private Sub FrmMasterAutomatedEmail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'DsTransaksiEmail.tbltrans_email_automated' table. You can move, or remove it, as needed.
        Me.Tbltrans_email_automatedTableAdapter.Fill(Me.DsTransaksiEmail.tbltrans_email_automated)
        'TODO: This line of code loads data into the 'DsTransaksiEmail.tbltrans_email_automated' table. You can move, or remove it, as needed.
        Me.Tbltrans_email_automatedTableAdapter.Fill(Me.DsTransaksiEmail.tbltrans_email_automated)
        'TODO: This line of code loads data into the 'DsTransaksiEmail.tbltrans_email_automated' table. You can move, or remove it, as needed.
        Me.Tbltrans_email_automatedTableAdapter.Fill(Me.DsTransaksiEmail.tbltrans_email_automated)

    End Sub

    Private Sub Tbltrans_email_automatedBindingNavigatorSaveItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Validate()
        Me.Tbltrans_email_automatedBindingSource.EndEdit()
        Me.Tbltrans_email_automatedTableAdapter.Update(Me.DsTransaksiEmail.tbltrans_email_automated)

    End Sub

    Private Sub Tbltrans_email_automatedBindingNavigatorSaveItem_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Tbltrans_email_automatedBindingNavigatorSaveItem.Click
        Me.Validate()
        Me.Tbltrans_email_automatedBindingSource.EndEdit()
        Me.Tbltrans_email_automatedTableAdapter.Update(Me.DsTransaksiEmail.tbltrans_email_automated)

    End Sub
End Class