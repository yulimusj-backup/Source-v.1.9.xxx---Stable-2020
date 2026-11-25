Public Class FrmTransaksiLicensing
    Dim ObjLicense As New ClsLicense

    Public BackColorAll As Color = Color.LightBlue

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiLicensing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.txtProductKey.Text = My.Settings.a
            Me.TxtLicenseKey.Text = My.Settings.b
            Me.dtpExpiry.Value = Me.ObjLicense.GetExpiryDate(My.Settings.b)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim EndDate As Date = Today

        Try
            If Me.TxtNewLicenseKey.Text.Trim <> "" Then
                If MsgBox("Enter new license key?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    'perbarui license key
                    If Me.ObjLicense.GetLicenseKey(Me.ObjLicense.GetExpiryDate(My.Settings.b) + TimeSpan.FromDays(1), _
                        My.Settings.a, Me.ObjLicense.GetTipeLicense(My.Settings.b) + 1) = Me.TxtNewLicenseKey.Text Then

                        My.Settings.b = Me.ObjLicense.GetLicenseKey(Me.ObjLicense.GetExpiryDate(My.Settings.b) + TimeSpan.FromDays(1), _
                        My.Settings.a, Me.ObjLicense.GetTipeLicense(My.Settings.b) + 1)

                        MsgBox("License Accepted", MsgBoxStyle.Information, "Information")

                        Me.Close()
                    Else
                        'jika sampai sini, maka ada 2 kemungkinan:
                        '1. license diperbarui dengan periode yang tidak standar
                        If IsDate(Me.ObjLicense.GetExpiryDate(Me.TxtNewLicenseKey.Text)) = True Then
                            EndDate = Me.ObjLicense.GetExpiryDate(Me.TxtNewLicenseKey.Text)
                            If Me.ObjLicense.GetLicenseKey(Me.ObjLicense.GetExpiryDate(My.Settings.b) + TimeSpan.FromDays(1), My.Settings.a, Me.ObjLicense.GetTipeLicense(Me.TxtNewLicenseKey.Text), EndDate) = Me.TxtNewLicenseKey.Text Then
                                My.Settings.b = Me.ObjLicense.GetLicenseKey(Me.ObjLicense.GetExpiryDate(My.Settings.b) + TimeSpan.FromDays(1), _
                                My.Settings.a, Me.ObjLicense.GetTipeLicense(My.Settings.b) + 1, EndDate)

                                MsgBox("License Accepted", MsgBoxStyle.Information, "Information")

                                Me.Close()
                                Exit Sub
                            End If
                        End If

                        '2. license invalid
                        MsgBox("Invalid license key", MsgBoxStyle.Information, "Information")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class