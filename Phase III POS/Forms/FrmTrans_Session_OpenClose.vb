Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmTrans_Session_OpenClose
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris
    Public mPrimarySession As Double = -1
    Public mParentFormSession As FrmTrans_Session
    Public mrOK As Boolean = False
    Public Obj As New ClsTrans_POS
    Public mFormPassword As FrmTrans_POS_ConfirmPassword

    Dim FrmPOS As New FrmTrans_POS
    Dim ObjCD As New ClsCekData
    Dim ObjSession As New ClsTrans_Session
    Dim ObjSessionItem As New ClsTrans_Session_Item
    Dim ObjPOS As New ClsTrans_POS
    Dim ObjPOSItem As New ClsTrans_POS_Item
    Dim ObjMember As New ClsMast_Member
    Dim ObjCur As New clsMasterCurrency
    Dim ObjVoid As New ClsTrans_POS_Void
    Dim ObjUser As New ClsUser
    Dim ObjGeneralSet As New ClsTrans_POS_GeneralSetting
    Dim TBL As New DseTrans_Session.tbltrans_pos_cabang1DataTable

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Session_OpenClose_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.mrOK = False
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.ObjSession.LoadDataPOSBuatSession(Me.ListBox1, Me.TBL, Me.mPrimarySession)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function HitungSubTotalDGV() As Decimal
        Dim RowDGV As DataGridViewRow
        Dim retval As Decimal = 0
        Try
            For Each RowDGV In Me.dgvItem.Rows
                retval += RowDGV.Cells(Me.ColSubTotal.Index).Value
            Next
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

   
    Private Sub btnPayment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPayment.Click
        Dim ij As New FrmTrans_POS_Payment
        Try
            If ObjPOS.CekApakahSudahPaid(Me.TxtPrimary.Text) = True Then
                MsgBox("This transaction has been paid", MsgBoxStyle.Information, "Information")
                Me.ObjSession.LoadDataPOSBuatSession(Me.ListBox1, Me.TBL, Me.mPrimarySession)
                Exit Sub
            End If

            If Me.dgvItem.Rows.Count > 0 Then
                'Load Form POS
                Me.FrmPOS.PrimaryCurr = ObjGeneralSet.GetPrimaryCurr()
                Me.FrmPOS.DefPrec = ObjCur.GetDefaultPrecission(ObjGeneralSet.GetPrimaryCurr())
                If ObjGeneralSet.GetAllowRounding = 0 Then
                    Me.FrmPOS.RoundingTipe = ClsCekData.EnumTipeRounding.Middlepoint
                    Me.FrmPOS.RoundingValue = 1
                ElseIf ObjGeneralSet.GetAllowRounding = 1 Then
                    Me.FrmPOS.RoundingTipe = ObjGeneralSet.GetRoundingTipe
                    If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.A_Seribu Then
                        Me.FrmPOS.RoundingValue = 1000
                    End If
                    If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.B_LimaRatus Then
                        Me.FrmPOS.RoundingValue = 500
                    End If
                    If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.C_Seratus Then
                        Me.FrmPOS.RoundingValue = 100
                    End If
                    If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.D_LimaPuluh Then
                        Me.FrmPOS.RoundingValue = 50
                    End If
                    If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.E_Sepuluh Then
                        Me.FrmPOS.RoundingValue = 10
                    End If
                    If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.F_Lima Then
                        Me.FrmPOS.RoundingValue = 5
                    End If
                    If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.G_Satu Then
                        Me.FrmPOS.RoundingValue = 1
                    End If
                    If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.H_NolKomaLima Then
                        Me.FrmPOS.RoundingValue = 0.5
                    End If
                    If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.I_NolKomaSatu Then
                        Me.FrmPOS.RoundingValue = 0.1
                    End If
                    If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.J_NolKomaNolLima Then
                        Me.FrmPOS.RoundingValue = 0.05
                    End If
                    If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.K_NolKomaNolSatu Then
                        Me.FrmPOS.RoundingValue = 0.01
                    End If
                    If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.L_NolKomaNolNolLima Then
                        Me.FrmPOS.RoundingValue = 0.005
                    End If
                    If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.M_NolKomaNolNolSatu Then
                        Me.FrmPOS.RoundingValue = 0.001
                    End If
                End If
                Me.FrmPOS.DefPrec = ObjCur.GetDecimal_DefaultCurrency_POS()
                'Btn Find
                Me.FrmPOS.TxtPrimarySession.Text = Me.TxtPrimarySession.Text
                Me.FrmPOS.NonAktifkanDetail()
                Me.FrmPOS.PbxViod.Visible = False
                Me.FrmPOS.PbxPaid.Visible = False
                Me.FrmPOS.Obj.ResetObject(Me.FrmPOS.dgvItem)
                Me.FrmPOS.KosongkanField()
                Me.FrmPOS.Obj.CariDetail(Me.TxtPrimary.Text, Me.FrmPOS.dgvItem, Me.ObjCD, Me.FrmPOS)
                Me.FrmPOS.SalinKeFieldUntukFind()
                Me.FrmPOS.btnClose.Enabled = True
                Me.FrmPOS.btnFind.Enabled = True
                Me.FrmPOS.btnNext.Enabled = False
                Me.FrmPOS.KeyPreview = True
                Me.FrmPOS.btnFind.Focus()
                Me.FrmPOS.UserCode = Me.UserCode
                'Payment details
                ij.ParentForm_1 = Me.FrmPOS
                ij.Obj = Me.FrmPOS.Obj
                ij.cboPaymentTipe.Focus()
                ij.ShowDialog()
                'Load Data
                'Me.ListBox1.Items.Clear()
                Me.TxtPrimary.Text = "-1"
                Me.TxtIdNumber.Text = ""
                Me.TxtPrimarySession.Text = Me.mPrimarySession
                Me.TxtPrimaryMember.Text = "-1"
                Me.TxtMember.Text = ""
                Me.dgvItem.Rows.Clear()
                Me.TxtDisplay.Clear()
                Me.ObjSession.LoadDataPOSBuatSession(Me.ListBox1, Me.TBL, Me.mPrimarySession)
                If Me.TBL.Rows.Count <= 0 Then
                    Me.mrOK = True
                    Me.btnClose.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim row As DseTrans_Session.tbltrans_pos_cabang1Row
        Try
            If Me.ListBox1.SelectedIndex = -1 Then
                Me.TxtPrimary.Text = "-1"
                Me.TxtIdNumber.Text = ""
                Me.TxtPrimarySession.Text = Me.mPrimarySession
                Me.TxtPrimaryMember.Text = "-1"
                Me.TxtMember.Text = ""
                Me.dgvItem.Rows.Clear()
            Else
                row = Me.TBL.DefaultView(Me.ListBox1.SelectedIndex).Row
                Me.SalinDataKeField(row.primary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinDataKeField(ByVal vPrimaryTransaksiPOS As Double)
        Try
            Me.TxtPrimary.Text = vPrimaryTransaksiPOS
            Me.TxtIdNumber.Text = ObjPOS.GetNoPOS(Me.TxtPrimary.Text)
            Me.TxtPrimarySession.Text = ObjPOS.GetPrimarySession(Me.TxtPrimary.Text)
            Me.TxtPrimaryMember.Text = Me.ObjPOS.GetPrimaryMember(Me.TxtPrimary.Text)
            Me.TxtMember.Text = Me.ObjMember.GetCode(Me.TxtPrimaryMember.Text)
            Me.ObjPOSItem.IsiDGVPOSSession(Me.TxtPrimary.Text, Me.dgvItem, Me, Me.ObjCD)
            Me.TxtDisplay.Text = "Total Sales: " + Me.ObjCur.GetCurrencyCode(Me.ObjGeneralSet.GetPrimaryCurr()).ToString + ". " + _
            Me.ObjCD.DigitGrouping(Me.HitungSubTotalDGV().ToString, 0, ClsCekData.EnumTipeRounding.Middlepoint)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Dim iy As New FrmTrans_POS_VoidQty1
        Dim Hasil As MsgBoxResult
        Dim row As DseTrans_Session.tbltrans_pos_cabang1Row
        Dim i As Integer = 0

        Try
            If Me.ListBox1.Items.Count > 0 Then
                Hasil = MsgBox("Do you want to void transaction ?", MsgBoxStyle.YesNo, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    'Sisa POS nya di Void
                    While i <= Me.TBL.Rows.Count - 1
                        row = Me.TBL.Rows(i)
                        If ObjPOS.GetIsVoid(row.primary) = 0 Then
                            With row
                                Me.ObjPOSItem.IsiDGVPOSSession(row.primary, Me.dgvItem, Me, Me.ObjCD)
                                Me.ObjPOS.UpdateIsVoid(row.primary, 1)
                                Me.ObjVoid.SimpanItem(row.primary, Me.mPrimarySession, Me.dgvItem, _
                                Me.FrmPOS, "", ObjUser.GetPrimary(Me.UserCode))
                            End With
                            i += 1
                        Else
                            Me.TBL.Rows.RemoveAt(i)
                        End If
                    End While
                    Me.btnClose.PerformClick()
                    Me.mrOK = True
                Else
                    Me.mrOK = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            iy.Dispose()
        End Try
    End Sub

    Private Sub btnVoid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVoid.Click
        Try
            If Me.TxtPrimary.Text <> "" And Me.TxtPrimary.Text <> "-1" Then
                Me.VoidTransaction(Me.TxtPrimary.Text, True, "")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnVoidAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVoidAll.Click
        Dim row As DseTrans_Session.tbltrans_pos_cabang1Row
        Dim iy As New FrmTrans_POS_VoidQty1

        Try
            'confirm
            If Me.ListBox1.Items.Count > 0 Then
                If MsgBox("Void all?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    iy.ShowDialog()
                    If iy.mrOk = True Then
                        'untuk setiap transaksi pos yg ada
                        For Each row In Me.TBL.Rows
                            '1. load datanya
                            Me.SalinDataKeField(row.primary)
                            '2. set void
                            Me.VoidTransaction(row.primary, False, iy.TxtReason.Text)
                        Next
                        'load ulang list box
                        Me.ObjSession.LoadDataPOSBuatSession(Me.ListBox1, Me.TBL, Me.mPrimarySession)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            iy.Dispose()
        End Try
    End Sub

    Private Sub VoidTransaction(ByVal vPrimaryPOS As Double, ByVal vConfirm As Boolean, ByVal vReason As String)
        Dim ij As New FrmTrans_POS_ConfirmPassword
        Dim iy As New FrmTrans_POS_VoidQty1
        Dim Hasil As MsgBoxResult
        Dim vUserCode As String = ""
        Dim vPrimaryUser As Double = -1

        Try
            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                MsgBox("Session ID : " + ObjSession.GetSessionID(Me.TxtPrimarySession.Text).ToString + " is closed." _
                , MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Obj.CekApakahSudahPaid(vPrimaryPOS) = True Then
                MsgBox("This transaction has been paid", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Obj.GetIsVoid(vPrimaryPOS) = 1 Then
                MsgBox("This transaction has been void", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If vPrimaryPOS <> "-1" Then
                If vConfirm = True Then
                    Hasil = MsgBox("Do you want to void transaction POS ?", MsgBoxStyle.YesNo, "Confirm")
                Else
                    Hasil = MsgBoxResult.Yes
                End If

                If Hasil = MsgBoxResult.Yes Then
                    'yang bisa dihapuskan ada transaksi dari session yang masih open
                    If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Open Then
                        'jika perlu konfirmasi, maka kelurkan form konfirmasi terlebih dahulu
                        If ObjGeneralSet.GetAllowUserToCancelWholeTransaction > 0 Then
                            If ObjGeneralSet.GetUseAuthorizationToCancelWholeTransaction > 0 Then
                                Me.mFormPassword = ij
                                Me.mFormPassword.cboUser.Focus()
                                Me.mFormPassword.ShowDialog()
                                If ij.mCancel = True Then
                                    Exit Sub
                                End If
                                vUserCode = ij.cboUser.Text
                                vPrimaryUser = ObjUser.GetPrimary(vUserCode)
                                ij.Dispose()
                            End If
                        Else
                            MsgBox("Transaction POS can't delete", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                        If vConfirm = True Then
                            iy.ShowDialog()
                        Else
                            iy.TxtReason.Text = vReason
                            iy.mrOk = True
                        End If

                        If iy.mrOk = True Then
                            'jika tidak perlu konfirmasi, maka langsung set properties "is void" dalam database menjadi 1
                            Me.Obj.UpdateIsVoid(vPrimaryPOS, 1)
                            Me.ObjVoid.SimpanItem(vPrimaryPOS, Me.TxtPrimarySession.Text, Me.dgvItem, Me, iy.TxtReason.Text, vPrimaryUser)
                            If vConfirm = True Then
                                Me.ObjSession.LoadDataPOSBuatSession(Me.ListBox1, Me.TBL, Me.mPrimarySession)
                            End If
                        Else
                            Exit Sub
                        End If
                    Else
                        MsgBox("Transaction POS can't delete because transaction session is closed", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            iy.Dispose()
        End Try
    End Sub
End Class