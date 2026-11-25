Option Strict Off
Option Explicit On

Public Class FrmTransaksiLicensing
    Dim ObjLicense As New ClsLicense

    Public BackColorAll As Color = Color.LightBlue

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        If Me.txtProductKey.Text <> Me.ObjLicense.GetProductKey Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Else
            Me.DialogResult = Windows.Forms.DialogResult.Abort
        End If
        Me.Close()
    End Sub

    Private Sub FrmTransaksiLicensing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.txtProductKey.Text = Me.ObjLicense.GetProductKey
            Me.TxtLicenseKey.Text = My.Settings.b
            Me.dtpExpiry.Value = Me.ObjLicense.GetExpiryDate(My.Settings.b)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim EndDate As Date = Today
        Dim ObjVer As New Core_Module.Clstrans_VersionControl
        Dim OldLicenseType, NewLicenseType As ClsLicense.EnumTipeLicense

        Try
            'generate license key
            OldLicenseType = Me.ObjLicense.GetTipeLicense(Me.TxtLicenseKey.Text)
            If CInt(OldLicenseType) > CInt(ClsLicense.EnumTipeLicense.Normal9) Then
                OldLicenseType = ClsLicense.EnumTipeLicense.Normal9
            End If
            If OldLicenseType = ClsLicense.EnumTipeLicense.Normal9 Then
                NewLicenseType = ClsLicense.EnumTipeLicense.Normal1
            Else
                NewLicenseType = Me.ObjLicense.GetTipeLicense(Me.TxtLicenseKey.Text) + 1
            End If

            If Me.TxtNewLicenseKey.Text.Trim <> "" Then
                'cek apakah expiry dari end date dari license baru yang dimasukkan lebih besar dari tanggal hari ini
                If Today.Date > Me.ObjLicense.GetExpiryDate(Me.TxtNewLicenseKey.Text) Then
                    MsgBox("Expiry date for the license you enter is " + Format(Me.ObjLicense.GetExpiryDate(Me.TxtNewLicenseKey.Text), _
                    "dd MMM yyyy") + ". This license has expired. Please contact your vendor for a valid license.", MsgBoxStyle.Information, "Information")
                    Me.TxtNewLicenseKey.Clear()
                    Me.TxtNewLicenseKey.Focus()
                    Exit Sub
                End If

                If MsgBox("Enter new license key?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    'perbarui license key
                    If Me.ObjLicense.GetLicenseKey(Me.ObjLicense.GetExpiryDate(My.Settings.b) + TimeSpan.FromDays(1), _
                        Me.txtProductKey.Text, NewLicenseType) = Me.TxtNewLicenseKey.Text Then

                        My.Settings.b = Me.ObjLicense.GetLicenseKey(Me.ObjLicense.GetExpiryDate(My.Settings.b) + TimeSpan.FromDays(1), _
                        Me.txtProductKey.Text, NewLicenseType)
                        My.Settings.a = Me.txtProductKey.Text
                        My.Settings.Save()
                        ObjVer.UpdateNewLicenseKey(My.Settings.b, My.Settings.a)
                        Sales_Module.SetLicense(My.Settings.b)
                        Purchase_Module.SetLicense(My.Settings.b)
                        Inventory_Module.SetLicense(My.Settings.b)
                        Core_Module.SetLicense(My.Settings.b)
                        AccModule.SetLicense(My.Settings.b)
                        HR.ModGlobalVariable.SetLicense(My.Settings.b)

                        MsgBox("License Accepted", MsgBoxStyle.Information, "Information")

                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Me.Close()
                    Else
                        'jika sampai sini, maka ada 2 kemungkinan:
                        '1. license diperbarui dengan periode yang tidak standar
                        If IsDate(Me.ObjLicense.GetExpiryDate(Me.TxtNewLicenseKey.Text)) = True Then
                            EndDate = Me.ObjLicense.GetExpiryDate(Me.TxtNewLicenseKey.Text)
                            If Me.ObjLicense.GetLicenseKey(Me.ObjLicense.GetExpiryDate(My.Settings.b) + TimeSpan.FromDays(1), _
                            Me.txtProductKey.Text, NewLicenseType, EndDate) = Me.TxtNewLicenseKey.Text Then
                                My.Settings.b = Me.ObjLicense.GetLicenseKey(Me.ObjLicense.GetExpiryDate(My.Settings.b) + TimeSpan.FromDays(1), _
                                Me.txtProductKey.Text, NewLicenseType, EndDate)
                                My.Settings.a = Me.txtProductKey.Text

                                My.Settings.Save()
                                ObjVer.UpdateNewLicenseKey(My.Settings.b, My.Settings.a)
                                Sales_Module.SetLicense(My.Settings.b)
                                Purchase_Module.SetLicense(My.Settings.b)
                                Inventory_Module.SetLicense(My.Settings.b)
                                Core_Module.SetLicense(My.Settings.b)
                                AccModule.SetLicense(My.Settings.b)
                                HR.ModGlobalVariable.SetLicense(My.Settings.b)

                                MsgBox("License Accepted", MsgBoxStyle.Information, "Information")
                                Me.DialogResult = Windows.Forms.DialogResult.OK

                                Me.Close()
                                Exit Sub
                            End If
                        End If

                        '2. license invalid
                        MsgBox("Invalid license key", MsgBoxStyle.Information, "Information")
                    End If
                End If
            Else
                MsgBox("Please specify a valid license.", MsgBoxStyle.Information, "Information")
                Me.TxtNewLicenseKey.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class