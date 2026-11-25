Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing, Core_Module
Public Class FrmMasterWorkShift_Allocate_Data01

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TBL As New Core_Module.DseMasterWorkshift_Allocation.Output_DGVDataTable

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterWorkShift_Allocate_Data01_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DGVPayment.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVPayment.DataSource = Me.TBL.DefaultView
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class