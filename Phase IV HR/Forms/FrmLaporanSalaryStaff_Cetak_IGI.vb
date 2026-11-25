Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing
Public Class FrmLaporanSalaryStaff_Cetak_IGI
    Public mDariLaporanMonthlySalaryStaffIncome As Boolean = False

    Private Sub FrmLaporanSalaryStaff_Cetak_IGI_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.crvReport1.DisplayBackgroundEdge = True
            Me.crvReport1.DisplayStatusBar = True
            Me.crvReport1.DisplayToolbar = True
            Me.crvReport1.ShowExportButton = True
            Me.crvReport1.ShowGotoPageButton = True
            Me.crvReport1.ShowPageNavigateButtons = True
            Me.crvReport1.ShowPrintButton = True
            Me.crvReport1.ShowTextSearchButton = True
            Me.crvReport1.ShowZoomButton = True

            Me.crvReport2.DisplayBackgroundEdge = True
            Me.crvReport2.DisplayStatusBar = True
            Me.crvReport2.DisplayToolbar = True
            Me.crvReport2.ShowExportButton = True
            Me.crvReport2.ShowGotoPageButton = True
            Me.crvReport2.ShowPageNavigateButtons = True
            Me.crvReport2.ShowPrintButton = True
            Me.crvReport2.ShowTextSearchButton = True
            Me.crvReport2.ShowZoomButton = True

            Me.crvReport3.DisplayBackgroundEdge = True
            Me.crvReport3.DisplayStatusBar = True
            Me.crvReport3.DisplayToolbar = True
            Me.crvReport3.ShowExportButton = True
            Me.crvReport3.ShowGotoPageButton = True
            Me.crvReport3.ShowPageNavigateButtons = True
            Me.crvReport3.ShowPrintButton = True
            Me.crvReport3.ShowTextSearchButton = True
            Me.crvReport3.ShowZoomButton = True

            Me.crvReport4.DisplayBackgroundEdge = True
            Me.crvReport4.DisplayStatusBar = True
            Me.crvReport4.DisplayToolbar = True
            Me.crvReport4.ShowExportButton = True
            Me.crvReport4.ShowGotoPageButton = True
            Me.crvReport4.ShowPageNavigateButtons = True
            Me.crvReport4.ShowPrintButton = True
            Me.crvReport4.ShowTextSearchButton = True
            Me.crvReport4.ShowZoomButton = True

            Me.crvReport5.DisplayBackgroundEdge = True
            Me.crvReport5.DisplayStatusBar = True
            Me.crvReport5.DisplayToolbar = True
            Me.crvReport5.ShowExportButton = True
            Me.crvReport5.ShowGotoPageButton = True
            Me.crvReport5.ShowPageNavigateButtons = True
            Me.crvReport5.ShowPrintButton = True
            Me.crvReport5.ShowTextSearchButton = True
            Me.crvReport5.ShowZoomButton = True

            If Me.mDariLaporanMonthlySalaryStaffIncome = True Then
                Me.TabControl1.TabPages.Remove(Me.TabPage5)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class