Option Explicit On
Option Strict Off
Public Class FrmMasterKaryawan_Kompetensi_Upd_Opt
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public DataType As ClsMasterKompetensi.EnumDataType = ClsMasterKompetensi.EnumDataType.Alfabetic
    Public DefPrec As Decimal = 0
    Dim Ext As Boolean = True
    Dim ObjCD As New Core_Module.ClsCekData

    Private Function ApakahDataBenar() As Boolean
        Try
            If DataType = CInt(ClsMasterKompetensi.EnumDataType.Alfabetic) Then
                If Me.cboDefValue.SelectedIndex = -1 Then
                    MsgBox("Please selected data value.", MsgBoxStyle.Information, "Information")
                    Me.cboDefValue.Select()
                    Return False
                End If
            ElseIf DataType = CInt(ClsMasterKompetensi.EnumDataType.Number) Then
                If Me.TxtDefValueDigit.Text = 0 Or Me.TxtDefValueDigit.Text = "" Then
                    MsgBox("Please positive number for value data.", MsgBoxStyle.Information, "Informasi")
                    Me.TxtDefValueDigit.Select()
                    Return False
                End If
            ElseIf DataType = CInt(ClsMasterKompetensi.EnumDataType.Tanggal) Then
            ElseIf DataType = CInt(ClsMasterKompetensi.EnumDataType.Text) Then
                If Me.TxtDefValueText.Text = "" Or String.IsNullOrEmpty(Me.TxtDefValueText.Text) = True Then
                    MsgBox("Value can't empty string.", MsgBoxStyle.Information, "Information")
                    Me.TxtDefValueText.Select()
                    Return False
                Else
                    Me.TxtDefValueText.Text = Me.TxtDefValueText.Text.ToString.Replace("'", "").Replace("*", "").Replace("^", "").Replace("%", "").ToString.Trim
                End If
            ElseIf DataType = CInt(ClsMasterKompetensi.EnumDataType.YN) Then
                If Me.cboDefValue.SelectedIndex = -1 Then
                    MsgBox("Please selected data value.", MsgBoxStyle.Information, "Information")
                    Me.cboDefValue.Select()
                    Return False
                End If
            End If

            Me.txtRemarks.Text = Me.txtRemarks.Text.ToString.Replace("'", "").Replace("*", "").Replace("^", "").Replace("%", "").ToString.Trim
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.DialogResult = Windows.Forms.DialogResult.No
        Ext = True
        Me.Close()
    End Sub

    Private Sub BtnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Dim MsgResult As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = False Then
                Exit Sub
            End If
            MsgResult = MsgBox("Save this data?", MsgBoxStyle.OkCancel, "Confrm")
            If MsgResult = MsgBoxResult.Ok Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Ext = True
                Me.Close()
            Else
                Ext = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboValue_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub FrmMasterKaryawan_Kompetensi_Upd_Opt_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Ext = True Then
                'Me.Close()
            Else
                e.Cancel = True
            End If
            Ext = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKaryawan_Kompetensi_Upd_Opt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If DataType = CInt(ClsMasterKompetensi.EnumDataType.Number) Then
                Me.TxtDefValueDigit.BringToFront()
                TxtDefValueDigit.Select()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = True
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = False
                Me.cboAlfabetic.TabStop = False

                Me.TxtDefValueDigit.Visible = True
                Me.TxtDefValueText.Visible = False
                Me.cboDefValue.Visible = False
                Me.DtpDefValue.Visible = False
                Me.cboAlfabetic.Visible = False

                Exit Sub
            End If
            If DataType = CInt(ClsMasterKompetensi.EnumDataType.Tanggal) Then
                Me.DtpDefValue.BringToFront()
                DtpDefValue.Select()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = True
                Me.cboDefValue.TabStop = False
                Me.cboAlfabetic.TabStop = False

                Me.TxtDefValueDigit.Visible = False
                Me.TxtDefValueText.Visible = False
                Me.cboDefValue.Visible = False
                Me.DtpDefValue.Visible = True
                Me.cboAlfabetic.Visible = False

                Exit Sub
            End If
            If DataType = CInt(ClsMasterKompetensi.EnumDataType.Text) Then
                Me.TxtDefValueText.BringToFront()
                TxtDefValueText.Select()
                Me.TxtDefValueText.TabStop = True
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = False
                Me.cboAlfabetic.TabStop = False

                Me.TxtDefValueDigit.Visible = False
                Me.TxtDefValueText.Visible = True
                Me.cboDefValue.Visible = False
                Me.DtpDefValue.Visible = False
                Me.cboAlfabetic.Visible = False

                Exit Sub
            End If
            If DataType = CInt(ClsMasterKompetensi.EnumDataType.YN) Then
                Me.cboDefValue.BringToFront()
                cboDefValue.Select()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = True
                Me.cboAlfabetic.TabStop = False

                Me.TxtDefValueDigit.Visible = False
                Me.TxtDefValueText.Visible = False
                Me.cboDefValue.Visible = True
                Me.DtpDefValue.Visible = False
                Me.cboAlfabetic.Visible = False

                Exit Sub
            End If

            If DataType = CInt(ClsMasterKompetensi.EnumDataType.Alfabetic) Then
                Me.cboDefValue.BringToFront()
                cboDefValue.Select()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = True
                Me.cboAlfabetic.TabStop = True

                Me.TxtDefValueDigit.Visible = False
                Me.TxtDefValueText.Visible = False
                Me.cboDefValue.Visible = False
                Me.DtpDefValue.Visible = False
                Me.cboAlfabetic.Visible = True

                Exit Sub
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefValueDigit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueDigit.Click
        Try
            Me.ObjCD.Number_Click(TxtDefValueDigit)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefValueDigit_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueDigit.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtDefValueDigit, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefValueDigit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueDigit.GotFocus
        Me.TxtDefValueDigit.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtDefValueDigit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDefValueDigit.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDefValueDigit, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefValueDigit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueDigit.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtDefValueDigit, sender, e, Me.DefPrec, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefValueDigit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueDigit.LostFocus
        Me.TxtDefValueDigit.BackColor = TextBox1.BackColor
    End Sub

    Private Sub TxtDefValueText_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueText.GotFocus
        Me.TxtDefValueText.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtDefValueText_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueText.LostFocus
        Me.TxtDefValueText.BackColor = Me.TextBox1.BackColor
    End Sub

    Private Sub txtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.GotFocus
        Me.txtRemarks.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.LostFocus
        Me.txtRemarks.BackColor = Me.TextBox1.BackColor
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Try
            If Me.CheckBox1.Checked = True Then
                Me.DtpNextCertDate.Enabled = True
            Else
                Me.DtpNextCertDate.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class