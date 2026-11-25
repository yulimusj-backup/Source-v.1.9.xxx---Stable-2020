Public Class FrmUpdateExpiryLicense
    Dim license_key_lama, license_key_baru As String
    Dim ExpiryBaru As Date = Today
    Dim ObjLicense As New ClsLicense

    Private Sub FrmUpdateExpiryLicense_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim d1, d2, m1, m2, y1, y2, y3, y4 As String
        Try
            'set expiry baru
            Me.ExpiryBaru = #12/31/2017#

            'ambil license key
            Me.license_key_lama = My.Settings.b
            d1 = Format(Me.ExpiryBaru, "ddMMyyyy").Substring(0, 1)
            d2 = Format(Me.ExpiryBaru, "ddMMyyyy").Substring(1, 1)
            m1 = Format(Me.ExpiryBaru, "ddMMyyyy").Substring(2, 1)
            m2 = Format(Me.ExpiryBaru, "ddMMyyyy").Substring(3, 1)
            y1 = Format(Me.ExpiryBaru, "ddMMyyyy").Substring(4, 1)
            y2 = Format(Me.ExpiryBaru, "ddMMyyyy").Substring(5, 1)
            y3 = Format(Me.ExpiryBaru, "ddMMyyyy").Substring(6, 1)
            y4 = Format(Me.ExpiryBaru, "ddMMyyyy").Substring(7, 1)

            'hitung license key baru
            license_key_baru = license_key_lama.Substring(0, 24) + d1
            license_key_baru = d2 + license_key_baru.Substring(1)
            license_key_baru = license_key_baru.Substring(0, 6) + m1 + license_key_baru.Substring(7)
            license_key_baru = license_key_baru.Substring(0, 12) + m2 + license_key_baru.Substring(13)
            license_key_baru = license_key_baru.Substring(0, 9) + y1 + license_key_baru.Substring(10)
            license_key_baru = license_key_baru.Substring(0, 21) + y2 + license_key_baru.Substring(22)
            license_key_baru = license_key_baru.Substring(0, 15) + y3 + license_key_baru.Substring(16)
            license_key_baru = license_key_baru.Substring(0, 18) + y4 + license_key_baru.Substring(19)

            'd1 = license_key_lama.Substring(24, 1)
            'd2 = license_key_lama.Substring(0, 1)
            'm1 = license_key_lama.Substring(6, 1)
            'm2 = license_key_lama.Substring(12, 1)
            'y1 = license_key_lama.Substring(9, 1)
            'y2 = license_key_lama.Substring(21, 1)
            'y3 = license_key_lama.Substring(15, 1)
            'y4 = license_key_lama.Substring(18, 1)

            'chekcing data
            If license_key_baru.Length <> license_key_lama.Length Then
                MsgBox("difference in length, please fix first", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.ExpiryBaru <> Me.ObjLicense.GetExpiryDate(Me.license_key_baru) Then
                MsgBox("target expiry date not achieved, please fix first", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'update license key ke semua module
            My.Settings.b = Me.license_key_baru
            My.Settings.Save()
            Sales_Module.SetLicense(My.Settings.b)
            Purchase_Module.SetLicense(My.Settings.b)
            Inventory_Module.SetLicense(My.Settings.b)
            Core_Module.SetLicense(My.Settings.b)
            AccModule.SetLicense(My.Settings.b)
            HR.ModGlobalVariable.SetLicense(My.Settings.b)
            My.Settings.Save()

            Me.Label1.Text = "Expiry set to:" + Format(Me.ExpiryBaru, "dd MMM yyyy")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class