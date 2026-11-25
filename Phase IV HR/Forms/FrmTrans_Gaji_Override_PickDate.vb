Public Class FrmTrans_Gaji_Override_PickDate
    Public IdentifierSalaryComponent As String = ""

    Dim LangsungClose As Boolean = False
    Dim ObjSalaryItem As New ClsTransaksiSalary_Item
    Dim ObjComp As New Core_Module.ClsMasterSalaryFormula_item

    Private Sub FrmTrans_Gaji_Override_PickDate_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.LangsungClose = False Then
                If MsgBox("Close form and cancel data copy?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_Override_PickDate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Try
            'cek apakah salary calculation pada tgl ybs yang menggunakan salary component tersebut
            If Me.ObjSalaryItem.GetJmlItem(Me.IdentifierSalaryComponent, Me.dtpDate.Value) = 0 Then
                MsgBox("There is no component: " + Me.ObjComp.GetDescComponent(Me.IdentifierSalaryComponent) + " on salary calculation dated " + _
                Format(Me.dtpDate.Value.Date, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'konfirmasi ulang
            If MsgBox("Start data copy from salary calculation on " + Format(Me.dtpDate.Value.Date, "dd MMM yyyy") + "?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.LangsungClose = True
                Me.DialogResult = Windows.Forms.DialogResult.OK
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class