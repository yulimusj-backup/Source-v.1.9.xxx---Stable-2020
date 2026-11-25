Option Explicit On
Option Strict Off

Public Class FrmTrans_SettingNPWP
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vPrimaryAcuan As Double = -1
    Public typeNPWP As Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Company
    Public vTipeFormMaster As ClsTrans_SettingNPWPNo.EnumFromMaster = ClsTrans_SettingNPWPNo.EnumFromMaster.Customer
    Public IsNPWPDigit As ClsTrans_SettingNPWPNo.EnumNPWPDigit = ClsTrans_SettingNPWPNo.EnumNPWPDigit._16Digit

    Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim Count As Integer = 0

    Private Function BlokSimbol(ByVal obj As TextBox, ByVal e As Windows.Forms.KeyPressEventArgs) As Boolean
        Try
            If e.KeyChar = "'" Or _
            e.KeyChar = "%" Or _
            e.KeyChar = "*" Or _
            e.KeyChar = "#" Then
                obj.ReadOnly = True
                obj.BackColor = Me.BackColorAll
            Else
                obj.ReadOnly = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtNPWPNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWPNo.GotFocus
        Try
            TxtNPWPNo.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNPWPNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtNPWPNo.KeyPress
        Try
            Me.BlokSimbol(Me.TxtNPWPNo, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNPWPNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWPNo.Leave
        Try
            'update npwp 16 digit.
            If Me.IsNPWPDigit = ClsTrans_SettingNPWPNo.EnumNPWPDigit._16Digit Then
                If Me.vTipeFormMaster = ClsTrans_SettingNPWPNo.EnumFromMaster.Customer Or _
                            Me.vTipeFormMaster = ClsTrans_SettingNPWPNo.EnumFromMaster.Karyawan Then
                    Me.TxtNIK.Text = Me.TxtNPWPNo.Text
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNPWPNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWPNo.LostFocus
        Try
            TxtNPWPNo.BackColor = Txt.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Try
            DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Dim Dialog As MsgBoxResult
        Try
            If Me.TxtNPWPNo.Text = "" Then
                MsgBox("NPWP No cannot empty string.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.IsNPWPDigit = ClsTrans_SettingNPWPNo.EnumNPWPDigit._15Digit Then
                If Me.ObjSetNPWP.ApakahSudah15Digit(Me.TxtNPWPNo.Text, True) = False Then
                    Exit Sub
                End If

            ElseIf Me.IsNPWPDigit = ClsTrans_SettingNPWPNo.EnumNPWPDigit._16Digit Then
                If Me.ObjSetNPWP.ApakahSudah16Digit(Me.TxtNPWPNo.Text, True) = False Then
                    Exit Sub
                End If

                If Me.vTipeFormMaster = ClsTrans_SettingNPWPNo.EnumFromMaster.Customer Or _
                   Me.vTipeFormMaster = ClsTrans_SettingNPWPNo.EnumFromMaster.Karyawan Then
                    If Me.TxtNIK.Text <> Me.TxtNPWPNo.Text Then
                        MsgBox("NPWP and NIK must be the same", MsgBoxStyle.Information, "Confirmation")
                        Exit Sub
                    End If
                End If
            End If


            Dialog = MsgBox("Do you want to update this data?", MsgBoxStyle.YesNo, "Confirmation")
            If Dialog = MsgBoxResult.Yes Then

                If Me.vTipeFormMaster = ClsTrans_SettingNPWPNo.EnumFromMaster.Customer Then
                    If Me.IsNPWPDigit = ClsTrans_SettingNPWPNo.EnumNPWPDigit._16Digit Then
                        If Me.CboTipeNPWP.SelectedIndex = Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP.Company Then
                            Me.ObjSetNPWP.UpdateNPWP_MasterCustomer(Me.TxtNPWPNo.Text, Me.TxtNPWPLama.Text, Me.CboTipeNPWP.SelectedIndex, vPrimaryAcuan)
                        Else

                            Me.ObjSetNPWP.UpdateNPWP_MasterCustomer(Me.TxtNPWPNo.Text, Me.TxtNPWPLama.Text, Me.CboTipeNPWP.SelectedIndex, Me.TxtNIK.Text, vPrimaryAcuan)
                        End If
                    Else
                        Me.ObjSetNPWP.UpdateNPWP_MasterCustomer(Me.TxtNPWPNo.Text, Me.CboTipeNPWP.SelectedIndex, vPrimaryAcuan)
                    End If
                   
                ElseIf Me.vTipeFormMaster = ClsTrans_SettingNPWPNo.EnumFromMaster.Supplier Then
                    If Me.IsNPWPDigit = ClsTrans_SettingNPWPNo.EnumNPWPDigit._16Digit Then
                        Me.ObjSetNPWP.UpdateNPWP_MasterSupplier(Me.TxtNPWPNo.Text, Me.TxtNPWPLama.Text, Me.CboTipeNPWP.SelectedIndex, vPrimaryAcuan)
                    Else
                        Me.ObjSetNPWP.UpdateNPWP_MasterSupplier(Me.TxtNPWPNo.Text, Me.CboTipeNPWP.SelectedIndex, vPrimaryAcuan)
                    End If

                ElseIf Me.vTipeFormMaster = ClsTrans_SettingNPWPNo.EnumFromMaster.Karyawan Then

                    If Me.IsNPWPDigit = ClsTrans_SettingNPWPNo.EnumNPWPDigit._16Digit Then
                        Me.ObjSetNPWP.UpdateNPWP_MasterKaryawan(Me.TxtNPWPNo.Text, Me.TxtNPWPLama.Text, ClsTrans_SettingNPWPNo.TipeNPWP.Individual, Me.TxtNIK.Text, vPrimaryAcuan)
                    Else
                        Me.ObjSetNPWP.UpdateNPWP_MasterKaryawan(Me.TxtNPWPNo.Text, ClsTrans_SettingNPWPNo.TipeNPWP.Individual, vPrimaryAcuan)
                    End If
                End If

                MsgBox("Successfully updated.", MsgBoxStyle.Information, "Information")
                DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_SettingNPWP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.IsNPWPDigit = ClsTrans_SettingNPWPNo.EnumNPWPDigit._16Digit Then
                Me.Label3.Text = Me.Label3.Text
            Else
                Me.Label3.Text = "Used 15 digit for NPWP No."
            End If

            Me.TxtNPWPNo.Text = Me.TxtNPWPLama.Text
            Me.CboTipeNPWP.SelectedIndex = Me.typeNPWP
            If Me.CboTipeNPWP.SelectedIndex = 0 Then
                Me.Label2.Visible = False
                Me.TxtNIK.Visible = False
            Else
                Me.Label2.Visible = True
                Me.TxtNIK.Visible = True
                Me.TxtNIK.ReadOnly = True
            End If

            If Me.vTipeFormMaster = ClsTrans_SettingNPWPNo.EnumFromMaster.Supplier Then
                Me.GroupBox1.Top = Me.Label3.Top + 15
                Me.GroupBox1.Height = Me.GroupBox1.Height - Me.TxtNPWPNo.Height
                Me.BtnCancel.Top = Me.GroupBox1.Bottom + 10
                Me.BtnOK.Top = Me.GroupBox1.Bottom + 10
                Me.Height = Me.Height - Me.BtnOK.Height - Me.BtnCancel.Height
                Me.TxtNIK.Visible = False
                Me.Label2.Visible = False
            ElseIf Me.vTipeFormMaster = ClsTrans_SettingNPWPNo.EnumFromMaster.Karyawan Then
                Me.Label2.Top = Me.Label1.Top
                Me.TxtNIK.Top = Me.TxtNPWPNo.Top
                Me.Label1.Top = Me.Label4.Top
                Me.TxtNPWPNo.Top = Me.CboTipeNPWP.Top

                Me.GroupBox1.Top = Me.Label3.Top + 15
                Me.GroupBox1.Height = Me.GroupBox1.Height - Me.CboTipeNPWP.Height
                Me.BtnCancel.Top = Me.GroupBox1.Bottom + 10
                Me.BtnOK.Top = Me.GroupBox1.Bottom + 10
                Me.Height = Me.Height - Me.BtnOK.Height - Me.BtnCancel.Height
                Me.CboTipeNPWP.Visible = False
                Me.Label4.Visible = False
                Me.Label2.Text = "ID Card No"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboTipeNPWP_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboTipeNPWP.SelectedIndexChanged
        Try

            If Me.vTipeFormMaster = ClsTrans_SettingNPWPNo.EnumFromMaster.Customer Then
                If Me.CboTipeNPWP.SelectedIndex = 0 Then
                    Me.TxtNIK.Text = ""
                    Me.TxtNIK.Visible = False
                    Me.Label2.Visible = False
                    Me.GroupBox1.Height = 92 - Me.TxtNPWPNo.Height
                Else
                    Me.TxtNIK.Text = Me.TxtNPWPNo.Text
                    Me.TxtNIK.Visible = True
                    Me.Label2.Visible = True
                    Me.TxtNIK.ReadOnly = True
                    If Me.Count > 0 Then 'kalau waktu awal ga usah di set.
                        Me.GroupBox1.Height = 72 + Me.TxtNPWPNo.Height
                    End If
                End If
            ElseIf Me.vTipeFormMaster = ClsTrans_SettingNPWPNo.EnumFromMaster.Karyawan Then
                Exit Sub
            Else
                Me.Label2.Visible = False
                Me.TxtNIK.Visible = False
            End If

            Me.Count += 1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class