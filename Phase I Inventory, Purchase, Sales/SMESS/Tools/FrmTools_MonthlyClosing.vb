Public Class FrmTools_MonthlyClosing
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjAkun As New AccModule.ClsMasterAkun
    Dim BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjClosing As New Purchase_Module.ClsTransaksiMonthlyClosing

    Private Sub FrmTools_MonthlyClosing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.T1_dtpFrom.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.T1_dtpFrom.Value)
            Me.T2_dtp.Value = Me.T1_dtpFrom.Value
            Me.T3_dtpPeriod.Value = Me.T1_dtpFrom.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class