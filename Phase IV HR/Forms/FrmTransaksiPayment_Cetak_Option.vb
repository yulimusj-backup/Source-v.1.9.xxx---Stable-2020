Option Explicit On
Option Strict Off
Public Class FrmTransaksiPayment_Cetak_Option
    Public mPrimaryJournal As Double = -1
    Public mFrmTransPayment As FrmTransaksiPayment
    Public mPrimarySalaryMin As Double = -1
    Public mPrimarySalaryMax As Double = -1
    Public mNoSalary As String = ""
    Public mDefPrec As Double = 0
    Public mTanggalTransaksi As Date = #1/1/1950#
    Public mPeriodFrom As Date = #1/1/1950#
    Public mPeriodTo As Date = #1/1/1950#

    Dim mFrmCetakJournal As New FrmTransaksiLoan_CetakJournal
    Dim mFrmCetakPaymentDetail As New FrmTransaksiPayment_Cetak

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Obj As New ClsTrans_Payment
        Try
            If Me.RadioButton1.Checked = True Then
                mFrmCetakJournal.PrimaryJournal = Me.mPrimaryJournal
                mFrmCetakJournal.TanggalPayment = Me.mTanggalTransaksi
                mFrmCetakJournal.cetak = FrmTransaksiLoan_CetakJournal.EnumCetak.Payment
                mFrmCetakJournal.ObjPayment = Obj
                mFrmCetakJournal.ShowDialog()
            ElseIf Me.RadioButton2.Checked = True Then
                mFrmCetakPaymentDetail.mNoSalary = Me.mNoSalary
                mFrmCetakPaymentDetail.mTanggalTransaksi = Me.mTanggalTransaksi
                mFrmCetakPaymentDetail.mPeriodFrom = Me.mPeriodFrom
                mFrmCetakPaymentDetail.mPeriodTo = Me.mPeriodTo
                mFrmCetakPaymentDetail.mFrmTransPayment = Me.mFrmTransPayment
                mFrmCetakPaymentDetail.mPrimarySalaryMin = Me.mPrimarySalaryMin
                mFrmCetakPaymentDetail.mPrimarySalaryMax = Me.mPrimarySalaryMax
                mFrmCetakPaymentDetail.mDefPrec = Me.mDefPrec
                mFrmCetakPaymentDetail.ShowDialog()
            End If
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPayment_Cetak_Option_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.RadioButton1.Checked = True
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class