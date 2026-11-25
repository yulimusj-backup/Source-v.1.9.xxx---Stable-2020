Public Class FrmTrans_POS_ConfirmPassword
    Public mCancel As Boolean = False
    Public mTipe As EnumTipe = EnumTipe.ItemDeletion
    Dim mrOk As Boolean = False
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjGeneralSet As New ClsTrans_POS_GeneralSetting
    Dim Dummy As String
    Dim counter As Double = 0
    Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbluserTableAdapter
    Dim TBL As New DseTrans_POS_GeneralSetting.tbluserDataTable

    Public Enum EnumTipe As Integer
        ItemDeletion
        WholeTransactionDeletion
        TransactionModification
        ModifyUnitPrice
        ModifiyDiscount
        AllowReceiptReprint
    End Enum

    Private Sub FrmTrans_POS_ConfirmPassword_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.mrOk = False Then
                Me.mCancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSO_DialogInvoiceNunggak_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'load orang - orang yang memiliki akses untuk melakukan oveeride invoice nunggka
            'orang-orang yang punya akses terhadap form payment
            ObjGeneralSet.IsiCBOConfirmPassword(Me.cboUser, DTA, TBL, Me.mTipe)
            While counter < Me.cboUser.Items.Count
                If Me.ObjUser.ApakahPunyaHakAkses(Me.cboUser.Items(counter), "FrmTrans_POS") = False Then
                    Me.cboUser.Items.RemoveAt(counter)
                Else
                    counter += 1
                End If
            End While
            If Me.cboUser.Items.Count > 0 Then
                Me.cboUser.SelectedIndex = 0
            End If
            Me.mCancel = False
            Me.mrOk = False
            Me.cboUser.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            'cek password
            If Core_Module.ModUser.Login(Me.TxtUserName, Me.TxtPassword) = True Then
                'jika benar maka dialog result = ok
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            Else
                Me.TxtPassword.Focus()
            End If
            Me.mCancel = False
            Me.mrOk = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboUser_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboUser.SelectedIndexChanged
        Try
            Me.TxtUserName.Text = Me.cboUser.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPassword.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.btnOK.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.mCancel = True
            Me.mrOk = False
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPassword_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPassword.GotFocus
        Try
            If Me.TxtPassword.Enabled = True Then
                If Me.TxtPassword.ReadOnly = False Then
                    Me.TxtPassword.BackColor = Color.LightGoldenrodYellow
                Else
                    Me.TxtPassword.BackColor = Me.TxtUserName.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPassword_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPassword.LostFocus
        Try
            If Me.TxtPassword.Enabled = True Then
                If Me.TxtPassword.ReadOnly = False Then
                    Me.TxtPassword.BackColor = Color.White
                Else
                    Me.TxtPassword.BackColor = Me.TxtUserName.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class