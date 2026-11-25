Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class FrmTrans_Gaji_Cetak_IGI_Options
    Public mQuickSalary_Form As Boolean = False
    '--Single Setting--
    Public mPrimarySalary As Double = -1
    '--Quick Setting--
    Public mListPrimarySalary As New List(Of Double)
    Public mStartDate_Min As Date = Date.MaxValue
    Public mEndDate_Max As Date = Date.MinValue
    Public mTanggalTrans_Dari As Date = #1/1/1950#
    Public mTanggalTrans_Sampai As Date = #1/1/1950#
    Public mDgvsalaryQuick As DataGridView
    Public mIndexColoumn_StartDate As Integer = 0
    Public mIndexColoumn_EndDate As Integer = 0
    Public mIndexColoumn_PrimarySalary As Integer = 0

    Dim ij As New FrmTrans_Gaji_Cetak_IGI

    Private Sub FrmTrans_Gaji_Cetak_IGI_Options_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.CboFormat.SelectedIndex = 0
            If Me.mQuickSalary_Form = False Then
                'Parameter dari calculate salary
                ij.mPrimarySalary = Me.mPrimarySalary
                ij.mQuickSalary_Form = Me.mQuickSalary_Form
                ij.mTanggalTrans_Dari = Me.mTanggalTrans_Dari
                ij.mTanggalTrans_Sampai = Me.mTanggalTrans_Sampai
            Else
                'Parameter dari quick salary
                ij.mQuickSalary_Form = True
                ij.mTanggalTrans_Dari = Me.mTanggalTrans_Dari
                ij.mTanggalTrans_Sampai = Me.mTanggalTrans_Sampai
                ij.mStartDate_Min = Me.mStartDate_Min
                ij.mEndDate_Max = Me.mEndDate_Max
                ij.mDgvsalaryQuick = Me.mDgvsalaryQuick
                ij.mIndexColoumn_StartDate = Me.mIndexColoumn_StartDate
                ij.mIndexColoumn_EndDate = Me.mIndexColoumn_EndDate
                ij.mIndexColoumn_PrimarySalary = Me.mIndexColoumn_PrimarySalary
                ij.mListPrimarySalary = Me.mListPrimarySalary
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        Try
            If Me.CboFormat.SelectedIndex = 0 Then
                ij.mFormatCetakan = FrmTrans_Gaji_Cetak_IGI.EnumFormatCetakan.A5
            Else
                ij.mFormatCetakan = FrmTrans_Gaji_Cetak_IGI.EnumFormatCetakan.Continuous
            End If
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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