Option Explicit On
Option Strict Off
Public Class FrmTrans_ManufacturingOrder_PrintOption

    Public PaperSize As EnumPaperSize = EnumPaperSize.A5
    Public Enum EnumPaperSize As Integer
        A5
        A4
    End Enum

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.PaperSize = Me.ComboBox1.SelectedIndex
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_PrintOption_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.AcceptButton = Me.Button1
        Me.CancelButton = Me.Button2
        Me.ComboBox1.SelectedIndex = EnumPaperSize.A5
    End Sub
End Class