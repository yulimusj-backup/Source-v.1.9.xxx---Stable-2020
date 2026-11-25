''@rosadinnn
''@201506241151
Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class FrmMasterSalaryFormula_Options

    Public mPrimary As Double = -1
    Public mCode As String = ""

    Private Sub FrmTrans_Gaji_Cetak_IGI_Options_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.CboFormat.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        Dim FrmPrint As New FrmMasterSalaryFormula_Print
        Try
            If Me.CboFormat.SelectedIndex = 0 Then
                FrmPrint.mReportType = FrmMasterSalaryFormula_Print.EnumReportType.EMPLOYEE_LIST
            ElseIf Me.CboFormat.SelectedIndex = 1 Then
                FrmPrint.mReportType = FrmMasterSalaryFormula_Print.EnumReportType.SALARY_FORMULA
            ElseIf Me.CboFormat.SelectedIndex = 2 Then
                FrmPrint.mReportType = FrmMasterSalaryFormula_Print.EnumReportType.SALARY_PARAMS
            End If
            FrmPrint.mPrimary = Me.mPrimary
            FrmPrint.mCode = Me.mCode
            Me.Hide()
            FrmPrint.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmPrint.Dispose()
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class