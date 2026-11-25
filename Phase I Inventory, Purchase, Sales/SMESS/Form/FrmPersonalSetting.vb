Option Explicit On
Option Strict Off

Public Class FrmPersonalSetting
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser
    Dim LangsungClose As Boolean = False

    Private Sub FrmPersonalSetting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.SalinKeFieldDefaultFindBy()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldDefaultFindBy()
        Dim Dummy As String = ""

        Try
            'salin data print menggunakan custom logic
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MasterCustomer, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MasterCustomer_DefaultFindBy, Me.UserCode)
            If Dummy = "" Then
                Me.cboFindBy_Customer.SelectedIndex = 0
            Else
                Me.cboFindBy_Customer.SelectedIndex = Dummy
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MasterProduct, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MasterProduct_DefaultFindBy, Me.UserCode)
            If Dummy = "" Then
                Me.cboFindBy_Product.SelectedIndex = 0
            Else
                Me.cboFindBy_Product.SelectedIndex = Dummy
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MasterSupplier, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MasterSupplier_DefaultFindBy, Me.UserCode)
            If Dummy = "" Then
                Me.cboFindBy_Supplier.SelectedIndex = 0
            Else
                Me.cboFindBy_Supplier.SelectedIndex = Dummy
            End If
            'add employee
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MasterKaryawan, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MasterKaryawan_DefaultFindBy, Me.UserCode)
            If Dummy = "" Then
                Me.cboFindBy_Employee.SelectedIndex = 0
            Else
                Me.cboFindBy_Employee.SelectedIndex = Dummy
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.SimpanDataSettingFindBy()
                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                Me.LangsungClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingFindBy()
        Try
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MasterCustomer, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MasterCustomer_DefaultFindBy, Me.cboFindBy_Customer.SelectedIndex, Me.UserCode)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MasterProduct, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MasterProduct_DefaultFindBy, Me.cboFindBy_Product.SelectedIndex, Me.UserCode)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MasterSupplier, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MasterSupplier_DefaultFindBy, Me.cboFindBy_Supplier.SelectedIndex, Me.UserCode)
            'add employee
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MasterKaryawan, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MasterKaryawan_DefaultFindBy, Me.cboFindBy_Employee.SelectedIndex, Me.UserCode)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            If MsgBox("Close form and discard data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.LangsungClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmPersonalSetting_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Hasil As MsgBoxResult

        Try
            If Me.LangsungClose = False Then
                Hasil = MsgBox("Save data before closing form?", MsgBoxStyle.YesNoCancel, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    Me.SimpanDataSettingFindBy()
                    Me.LangsungClose = True
                    Me.Close()
                ElseIf Hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                ElseIf Hasil = MsgBoxResult.No Then
                    Me.LangsungClose = True
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class