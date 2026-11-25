Option Explicit On
Option Strict Off

Imports System.Windows.Forms

Public Class FrmTransaksiSalesInvoice_JenisFakturPajakStandar
    Public JenisFakturPajakStandar As FrmCetakFaktur.EnumJenisFakturStandar = FrmCetakFaktur.EnumJenisFakturStandar.HargaJual

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub rboHargaJual_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboHargaJual.CheckedChanged
        Try
            If Me.rboHargaJual.Checked = True Then
                Me.JenisFakturPajakStandar = FrmCetakFaktur.EnumJenisFakturStandar.HargaJual
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboPenggantian_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboPenggantian.CheckedChanged
        Try
            If Me.rboPenggantian.Checked = True Then
                Me.JenisFakturPajakStandar = FrmCetakFaktur.EnumJenisFakturStandar.Penggantian
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboTermin_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboTermin.CheckedChanged
        Try
            If Me.rboTermin.Checked = True Then
                Me.JenisFakturPajakStandar = FrmCetakFaktur.EnumJenisFakturStandar.Termin
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboUangMuka_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboUangMuka.CheckedChanged
        Try
            If Me.rboUangMuka.Checked = True Then
                Me.JenisFakturPajakStandar = FrmCetakFaktur.EnumJenisFakturStandar.UangMuka
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesInvoice_JenisFakturPajakStandar_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.rboUangMuka.Enabled = False
    End Sub
End Class
