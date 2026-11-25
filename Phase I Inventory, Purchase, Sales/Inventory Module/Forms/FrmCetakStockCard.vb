Option Explicit On
Option Strict Off

Public Class FrmCetakStockCard
    Dim Report As New crpCetakStockCard
    Dim ObjStockCard As New ClsTransaksiStockCard

    Public StartDate As Date = Now
    Public EndDate As Date = Now
    Public PrimaryProduct As Double = -1
    Public ProductCode As String = ""
    Public Units As String = ""
    Dim ObjProduct As New Core_Module.clsMasterProduct

    Private Sub FrmCetakStockCard_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.ObjStockCard.Print(Me.PrimaryProduct, Me.StartDate, Me.EndDate, Me.Report, _
            Me.crvStockCard, Me.ProductCode, "", Me.Units, Me.ObjProduct.GetAngkaDecimal(Me.PrimaryProduct), False, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class