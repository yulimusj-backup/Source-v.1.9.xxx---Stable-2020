Public Class FrmTrans_Export_eSPT_Msg
    Public PenghasilanBruto As Decimal = 0
    Public JmlKary As Double = 0

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency

    Private Sub FrmTrans_Export_eSPT_Msg_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.TxtJmlKary.Text = Me.ObjCD.DigitGrouping(Me.JmlKary, 0)
            Me.TxtPenghasilanBruto.Text = Me.ObjCD.DigitGrouping(Me.PenghasilanBruto, Me.ObjCur.GetDecimal_DefaultCurrency)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class