Public Class FrmLaporanCommission
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim FormatLaporan As EnumFormat = EnumFormat.Simple
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Enum EnumFormat As Integer
        Simple
        With_Detail_Item
        With_Detail_Transaction
    End Enum

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub btnSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSettings.Click
        Dim ij As New Core_Module.FrmMasterCommission

        Try
            ij.UserCode = Me.UserCode
            ij.BackColorAll = Me.BackColorAll

            ij.ShowDialog()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub FrmLaporanCommission_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.cboFormat.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboFormat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormat.SelectedIndexChanged
        Try
            If Me.cboFormat.SelectedIndex = 0 Then
                Me.FormatLaporan = EnumFormat.Simple
            ElseIf Me.cboFormat.SelectedIndex = 1 Then
                Me.FormatLaporan = EnumFormat.With_Detail_Item
            Else
                Me.FormatLaporan = EnumFormat.With_Detail_Transaction
            End If

            If Me.FormatLaporan = EnumFormat.With_Detail_Transaction Then
                Me.llbSorting.Visible = True
            Else
                Me.llbSorting.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


End Class