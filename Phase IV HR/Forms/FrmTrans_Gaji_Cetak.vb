Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class FrmTrans_Gaji_Cetak
    'Public Enum EnumFormatCetakan As Integer
    '    A5
    '    Continuous
    'End Enum
    Public mPrimary As Double = -1
    Public mDef_Prec As Double = 0
    Public UseCustom As Boolean = False
    Public QuickSalary As Boolean = False
    Public vStartDate As Date = #1/1/1950#
    Public vEndDate As Date = #1/1/1950#
    Public vListPrimarySalary As New List(Of Double)

    Dim Obj As New ClsTransaksiSalary
    Dim ObjSalaryCustom As New CustomSMESS.ClsTransaksiSalary

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_Gaji_Cetak_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ShowIcon = False
            If Me.UseCustom = False Then
                Me.Obj.Cetak(Me.crv1, Me.mPrimary, Me.mDef_Prec)
            Else
                If Me.QuickSalary = True Then
                    Me.ObjSalaryCustom.CetakTransaksiSalaryQuick(Me.vStartDate, Me.vEndDate, Me.vListPrimarySalary, Me.crv1)
                Else
                    Me.ObjSalaryCustom.CetakTransaksiSalary(Me.crv1, Me.mPrimary)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub CreatePDF()
        Try
            Me.ObjSalaryCustom.CetakTransaksiSalary(Me.crv1, Me.mPrimary, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class