Option Explicit On
Option Strict Off

Public Class FrmMasterCurrency
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow

    Dim baru As Boolean
    Dim objCurrency As clsMasterCurrency
    Dim buffer As String = ""
    Dim ObjCD As New ClsCekData

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            Me.TxtPrimaryCurrency.Text = "-1"
            Me.txtDefaultPrecision.Text = "0"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("currency code")
            KoleksiTextBox.Add(Me.TxtCurrencyCode)
            KoleksiNamaData.Add("currency name")
            KoleksiTextBox.Add(Me.TxtCurrencyName)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()
            KoleksiNamaData.Add("default precision")
            KoleksiTextBox.Add(Me.txtDefaultPrecision)
            If ModCekData.CekNonNegatif(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.objCurrency.ApakahSudahAda(Me.TxtCurrencyCode.Text) = True Then
                    MsgBox("Duplication in currency code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtCurrencyCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.objCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                If ProductGroupLama <> Me.TxtCurrencyCode.Text Then
                    If Me.objCurrency.ApakahSudahAda(Me.TxtCurrencyCode.Text) = True Then
                        MsgBox("Duplication in currency code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtCurrencyCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objCurrency.GetTimeStamp(Me.TxtPrimaryCurrency.Text)
                If Me.objCurrency.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objCurrency.CariDetail(Me.objCurrency.Primary)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'maksimal default precision 4
            If CInt(Me.txtDefaultPrecision.Text) > 4 Then
                MsgBox("Maiximum value for default precision is 4.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'kalo nilai defaul precisionnya turun, maka harus dicek, apakah ada nilai yang dilanggar, sehingga menyebabkan sistem menjadi tidak stabil.
            'msgbox("belum")

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.objCurrency
                .CurrencyCode = Me.TxtCurrencyCode.Text
                .Description = Me.TxtCurrencyName.Text
                .DefaultPrecision = Me.txtDefaultPrecision.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objCurrency
                Me.TxtCurrencyCode.Text = .CurrencyCode
                Me.TxtCurrencyName.Text = .Description
                Me.TxtPrimaryCurrency.Text = .Primary
                Me.txtDefaultPrecision.Text = .DefaultPrecision
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.objCurrency = New clsMasterCurrency
            Me.baru = True
            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objCurrency.SimpanBaru(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objCurrency.SimpanPerubahan(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtCurrencyCode.Text = "" And Me.TxtCurrencyName.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtCurrencyCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariMasterCurrency

        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            If Not ij.PrimaryCurrency = -1 Then
                If Me.objCurrency.CariDetail(ij.PrimaryCurrency) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If Me.TxtCurrencyCode.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.TxtCurrencyCode.Text <> "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.TxtCurrencyCode.Focus()
            'jika IDR, tidak bisa diedit dan didelete
            If Me.TxtCurrencyCode.Text = "IDR" Then
                Me.TxtCurrencyCode.Enabled = False
                Me.TxtCurrencyName.Focus()
            End If
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtCurrencyCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtCurrencyCode.Text <> "" Then
                'khusus untuk IDR, tidak bisa dihapus, karena digunakan oleh sistem
                If Me.TxtCurrencyCode.Text = "IDR" Or Me.TxtPrimaryCurrency.Text = "1" Then
                    MsgBox("This the system default currency, you cannot delete this data", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If Me.objCurrency.ApakahBolehHapus(Me.TxtPrimaryCurrency.Text) = False Then
                    'MsgBox("Belum")
                    MsgBox("You cannot delete this data, because there are customer using this data.", _
                    MsgBoxStyle.Information, "Information")
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objCurrency.Hapus(Me.TxtPrimaryCurrency.Text) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objCurrency.SimpanBaru(Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objCurrency.SimpanPerubahan(Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDefaultPrecision_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDefaultPrecision.Enter
        Try
            Me.buffer = Me.txtDefaultPrecision.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDefaultPrecision_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDefaultPrecision.Leave
        Try
            If ModCekData.CekNonNegatif(Me.txtDefaultPrecision.Text) = False Then
                MsgBox("Default precision must a non-negative numeric data.", MsgBoxStyle.Information, _
                "Information")
                Me.txtDefaultPrecision.Text = Me.buffer
                Me.txtDefaultPrecision.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TxtCurrencyCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode.GotFocus
        TxtCurrencyCode.BackColor = BackColorAll
    End Sub

    Private Sub TxtCurrencyCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode.LostFocus
        TxtCurrencyCode.BackColor = Me.TxtPrimaryCurrency.BackColor
    End Sub

    Private Sub TxtCurrencyName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyName.GotFocus
        TxtCurrencyName.BackColor = BackColorAll
    End Sub

    Private Sub TxtCurrencyName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyName.LostFocus
        TxtCurrencyName.BackColor = Me.TxtPrimaryCurrency.BackColor
    End Sub

    Private Sub txtDefaultPrecision_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDefaultPrecision.GotFocus
        txtDefaultPrecision.BackColor = BackColorAll
    End Sub

    Private Sub txtDefaultPrecision_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDefaultPrecision.LostFocus
        txtDefaultPrecision.BackColor = Me.TxtPrimaryCurrency.BackColor
    End Sub
End Class