Public Class FrmTransaksiLoan_CetakJournal
    Public TanggalPayment As Date = #1/1/1950#
    Public PrimaryJournal As Double = -1
    Public ObjLoan As ClsTransaksiLoan
    Public ObjLoanPayment As ClsTransaksiLoanPayment
    Public ObjPayment As ClsTrans_Payment
    Dim ObjJurnal_Custom As New CustomSMESS.ClsTransaksiJournal
    Public vDgv As Windows.Forms.DataGridView
    Public vFrm As FrmTransaksiLoan_Payment
    Public cetak As EnumCetak
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    Enum EnumCetak
        Loan
        LoanPayment
        Payment
    End Enum

    Private Sub FrmTransaksiLoan_CetakJournal_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomPrint) = "1" Then
                Me.ObjJurnal_Custom.CetakJurnal(Me.PrimaryJournal, Me.crvReport)
            Else
                If Me.cetak = EnumCetak.Loan Then
                    Me.ObjLoan.CetakJournal(Me.PrimaryJournal, Me.crvReport)
                ElseIf Me.cetak = EnumCetak.LoanPayment Then
                    Me.ObjLoanPayment.CetakJournal(Me.PrimaryJournal, Me.crvReport, vDgv, vFrm)
                ElseIf Me.cetak = EnumCetak.Payment Then
                    Me.ObjPayment.CetakJournal(Me.PrimaryJournal, Me.crvReport, Me.TanggalPayment)
                End If
            End If
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
End Class