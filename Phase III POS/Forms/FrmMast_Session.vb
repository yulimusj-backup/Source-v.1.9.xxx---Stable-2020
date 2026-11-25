Option Explicit On
Option Strict Off
Imports Core_Module, AccModule

Public Class FrmMast_Session
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris

    Dim baru As Boolean
    Dim DefaultClock As DateTime = "1950/01/01" + " " + Now.Hour.ToString + ":" + Now.Minute.ToString + ":" + Now.Second.ToString
    Dim buffer As String = ""
    Dim ObjCD As New ClsCekData
    Dim Obj As New ClsMast_Session
    Dim ObjCashier As New ClsUser
    Dim ObjAcc As New ClsMasterAkun
    Dim ObjBranch As New ClsMast_Cabang
    Dim DefPrec As Double = 0

    Private Sub FrmMasterAirport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Obj = New ClsMast_Session
            Me.baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit
            Me.BackColorDefault()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)

            Me.TxtPrimary.Text = "-1"
            Me.TxtCash_Beginning.Text = "0"
            Me.dtpSessionPeriod.Value = DefaultClock
            Me.DtpSessionThrough.Value = DefaultClock
            Me.TxtPrimaryCashier.Text = "-1"
            Me.TxtPrimaryBranch.Text = "-1"
            Me.TxtPrimaryBeginingCashAccount.Text = "-1"
            Me.TxtPrimaryCashierCashAccount.Text = "-1"
            Me.TxtPrimaryBalanceCashAccount.Text = "-1"

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
        Dim ListPrimaryAkun As New List(Of Double)
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Reference Number")
            KoleksiTextBox.Add(Me.TxtReferenceName)

            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.Obj.ApakahSudahAda(Me.TxtReferenceName.Text) = True Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Duplication in reference name, please change your code", _
                        MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Duplikasi code, silahkan ubah reference name", _
                        MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtReferenceName.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.Obj.GetRefName(Me.TxtPrimary.Text)
                If TipeCodeLama <> Me.TxtReferenceName.Text Then
                    If Me.Obj.ApakahSudahAda(Me.TxtReferenceName.Text) = True Then
                        If Bahasa = EnumBahasa.Inggris Then
                            MsgBox("Duplication in code, please change your reference name", _
                            MsgBoxStyle.Information, "Information")
                        ElseIf Bahasa = EnumBahasa.Indonesia Then
                            MsgBox("Duplikasi code, silahkan ubah reference name", _
                            MsgBoxStyle.Information, "Informasi")
                        End If
                        Me.TxtReferenceName.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Obj.GetTimestamp(Me.TxtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If Bahasa = EnumBahasa.Inggris Then
                        If MsgBox("Other user have modified this data, reload latest data?", _
                        MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Me.Obj.CariDetail(Me.Obj.Primary, Me.ObjCD, Me)
                            Me.SalinKeField()
                        End If
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        If MsgBox("Beberapa data telah dimodifikasi, apakah mau mengambil kembali data terakhir?", _
                        MsgBoxStyle.YesNo, "Konfirm") = MsgBoxResult.Yes Then
                            Me.Obj.CariDetail(Me.Obj.Primary, Me.ObjCD, Me)
                            Me.SalinKeField()
                        End If
                    End If
                    Return False
                End If
            End If

            'Desc harus diisi
            If String.IsNullOrEmpty(Me.TxtReferenceName.Text) = True Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input reference name", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan reference name", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtReferenceName.Focus()
                Return False
            End If

            If String.IsNullOrEmpty(Me.TxtCashierCode.Text) = True Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input cashier name", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan cashier name", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtCashierCode.Focus()
                Return False
            End If

            If Me.TxtPrimaryBranch.Text = "-1" Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input branch code", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan branch code", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtBranchCode.Focus()
                Return False
            End If
            If Me.TxtCash_Beginning.Text = "0" Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input beginning cash", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan beginning cash", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtCash_Beginning.Focus()
                Return False
            End If
            If Me.TxtPrimaryBeginingCashAccount.Text = "-1" Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input beginning cash account", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan beginning cash account", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtBeginingCashAccountCode.Focus()
                Return False
            End If
            If Me.TxtPrimaryCashierCashAccount.Text = "-1" Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input cashier cash account", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan cashier cash account", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtCashierCashAccountCode.Focus()
                Return False
            End If
            If Me.TxtPrimaryBalanceCashAccount.Text = "-1" Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input balance cash account", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan balance cash account", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtBalanceCashAccountCode.Focus()
                Return False
            End If

            If String.IsNullOrEmpty(Me.TxtRemarks.Text) = False Then
                If ObjCD.CekDataKosong(Me.TxtRemarks.Text) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Data tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtRemarks.Focus()
                    Return False
                End If
            End If

            'Cek Bilangan Positif
            If ObjCD.CekBilanganPositif(Me.TxtCash_Beginning.Text) = False Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtCash_Beginning.Focus()
                Return False
            End If

            'Cek kode akun
            If ListPrimaryAkun.Contains(Me.TxtPrimaryCashierCashAccount.Text) = False Then
                ListPrimaryAkun.Add(Me.TxtPrimaryCashierCashAccount.Text)
            Else
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Cashier cash account code has been used in another account code.", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Cashier cash account code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtCashierCashAccountCode.Focus()
                Return False
            End If

            If ListPrimaryAkun.Contains(Me.TxtPrimaryBeginingCashAccount.Text) = False Then
                ''ListPrimaryAkun.Add(Me.TxtPrimaryBeginingCashAccount.Text)
            Else
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Beginning cash account code has been used in another account code.", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Beginning cash account code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtBeginingCashAccountCode.Focus()
                Return False
            End If

            If ListPrimaryAkun.Contains(Me.TxtPrimaryBalanceCashAccount.Text) = False Then
                ''ListPrimaryAkun.Add(Me.TxtPrimaryBalanceCashAccount.Text)
            Else
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Balance cash account code has been used in another account code.", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Balance cash account code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtBalanceCashAccountCode.Focus()
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            ListPrimaryAkun = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                .RefName = Me.TxtReferenceName.Text
                .SessionPeriod = "1950/01/01" + " " + Me.dtpSessionPeriod.Value.Hour.ToString + ":" + Me.dtpSessionPeriod.Value.Minute.ToString + ":" + Me.dtpSessionPeriod.Value.Second.ToString
                .SessionThrough = "1950/01/01" + " " + Me.DtpSessionThrough.Value.Hour.ToString + ":" + Me.DtpSessionThrough.Value.Minute.ToString + ":" + Me.DtpSessionThrough.Value.Second.ToString
                .Remarks = Me.TxtRemarks.Text
                .CashierCode = Me.TxtCashierCode.Text
                .PrimaryBranch = Me.TxtPrimaryBranch.Text
                .BeginningCash = Me.TxtCash_Beginning.Text
                .PrimaryBeginningCashAccount = Me.TxtPrimaryBeginingCashAccount.Text
                .PrimaryCashierCashAccount = Me.TxtPrimaryCashierCashAccount.Text
                .PrimaryBalanceCashAccount = Me.TxtPrimaryBalanceCashAccount.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtReferenceName.Text = .RefName
                Me.dtpSessionPeriod.Value = "01/01/1950 " + .SessionPeriod.Hour.ToString + ":" + .SessionPeriod.Minute.ToString + ":" + .SessionPeriod.Second.ToString
                Me.DtpSessionThrough.Value = "01/01/1950 " + .SessionThrough.Hour.ToString + ":" + .SessionThrough.Minute.ToString + ":" + .SessionThrough.Second.ToString
                Me.TxtRemarks.Text = .Remarks
                Me.TxtPrimaryCashier.Text = "-1"
                Me.TxtCashierCode.Text = .CashierCode
                Me.TxtCashierName.Text = ObjCashier.GetNama(.CashierCode)
                Me.TxtPrimaryBranch.Text = .PrimaryBranch
                Me.TxtBranchCode.Text = ObjBranch.GetCode(.PrimaryBranch)
                Me.TxtBranchName.Text = ObjBranch.GetDesc(.PrimaryBranch)
                Me.TxtCash_Beginning.Text = ObjCD.DigitGrouping(.BeginningCash, DefPrec)
                Me.TxtPrimaryBeginingCashAccount.Text = .PrimaryBeginningCashAccount
                Me.TxtBeginingCashAccountCode.Text = ObjAcc.GetAccountCode(.PrimaryBeginningCashAccount)
                Me.TxtBeginingCashAccountDesc.Text = ObjAcc.GetDescription(Me.TxtBeginingCashAccountCode.Text)
                Me.TxtPrimaryCashierCashAccount.Text = .PrimaryCashierCashAccount
                Me.TxtCashierCashAccountCode.Text = ObjAcc.GetAccountCode(.PrimaryCashierCashAccount)
                Me.TxtCashierCashAccountDesc.Text = ObjAcc.GetDescription(Me.TxtCashierCashAccountCode.Text)
                Me.TxtPrimaryBalanceCashAccount.Text = .PrimaryBalanceCashAccount
                Me.TxtBalanceCashAccountCode.Text = ObjAcc.GetAccountCode(.PrimaryBalanceCashAccount)
                Me.TxtBalanceCashAccountDesc.Text = ObjAcc.GetDescription(Me.TxtBalanceCashAccountCode.Text)

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Kontrol Button
    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, _
                        Me.TxtPrimary.Text, Bahasa, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, _
                        Me.TxtPrimary.Text, Bahasa, Me) = True Then
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
            If Not Me.TxtReferenceName.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                        Me.TxtReferenceName.Clear()
                    Else
                        Me.Obj.CariDetail(Me.TxtPrimary.Text, Me.ObjCD, Me)
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
            Me.Baru = True
            Obj.ResetObject()
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.TxtReferenceName.Focus()
            Me.TxtReferenceName.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ijk As New FrmMast_Session_Cari
        Try
            ijk.BackColorAll = Me.BackColorAll
            ijk.cboCriteria.Items.Clear()
            ijk.txtCari.Text = ""
            ijk.dgvHasil.DataSource = ""
            ijk.ShowDialog()

            If Not ijk.mPrimary = -1 Then
                If Me.Obj.CariDetail(ijk.mPrimary, Me.ObjCD, Me) = True Then
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
            ijk.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtReferenceName.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtReferenceName.Text <> "" Then
                If Me.Obj.ApakahBolehEdit(Me.TxtPrimary.Text, Me.UserCode, True, True, Bahasa) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.TxtReferenceName.Focus()
                    Me.TxtReferenceName.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtReferenceName.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtReferenceName.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.TxtPrimary.Text, Me.UserCode, True, True, _
                 Bahasa, Me) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.TxtPrimary.Text, Me.UserCode, True, True, _
                        Bahasa, Me) = True Then
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
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtReferenceName.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, _
                            Me.TxtPrimary.Text, Bahasa, Me) = True Then
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
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, _
                            Me.TxtPrimary.Text, Bahasa, Me) = True Then
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


    'Coding BackColor
    Private Sub TxtReferenceName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReferenceName.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtReferenceName.Enabled = True Then
                    If Me.TxtReferenceName.ReadOnly = False Then
                        TxtReferenceName.BackColor = Me.BackColorAll
                    Else
                        TxtReferenceName.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReferenceName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReferenceName.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtReferenceName.Enabled = True Then
                    If Me.TxtReferenceName.ReadOnly = False Then
                        TxtReferenceName.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtReferenceName.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCashierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCashierCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCashierCode.Enabled = True Then
                    If Me.TxtCashierCode.ReadOnly = False Then
                        TxtCashierCode.BackColor = Me.BackColorAll
                    Else
                        TxtCashierCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCashierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCashierCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCashierCode.Enabled = True Then
                    If Me.TxtCashierCode.ReadOnly = False Then
                        TxtCashierCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtCashierCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.BackColorAll
                Else
                    TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCity_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBranchCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBranchCode.Enabled = True Then
                    If Me.TxtBranchCode.ReadOnly = False Then
                        Me.TxtBranchCode.BackColor = Me.BackColorAll
                    Else
                        Me.TxtBranchCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCity_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBranchCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBranchCode.Enabled = True Then
                    If Me.TxtBranchCode.ReadOnly = False Then
                        Me.TxtBranchCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtBranchCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDescItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCash_Beginning.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCash_Beginning.Enabled = True Then
                    If Me.TxtCash_Beginning.ReadOnly = False Then
                        Me.TxtCash_Beginning.BackColor = Me.BackColorAll
                    Else
                        Me.TxtCash_Beginning.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDescItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCash_Beginning.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCash_Beginning.Enabled = True Then
                    If Me.TxtCash_Beginning.ReadOnly = False Then
                        Me.TxtCash_Beginning.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtCash_Beginning.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarksItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginingCashAccountCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBeginingCashAccountCode.Enabled = True Then
                    If Me.TxtBeginingCashAccountCode.ReadOnly = False Then
                        Me.TxtBeginingCashAccountCode.BackColor = Me.BackColorAll
                    Else
                        Me.TxtBeginingCashAccountCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarksItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginingCashAccountCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBeginingCashAccountCode.Enabled = True Then
                    If Me.TxtBeginingCashAccountCode.ReadOnly = False Then
                        Me.TxtBeginingCashAccountCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtBeginingCashAccountCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTerminalCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCashierCashAccountCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCashierCashAccountCode.Enabled = True Then
                    If Me.TxtCashierCashAccountCode.ReadOnly = False Then
                        Me.TxtCashierCashAccountCode.BackColor = Me.BackColorAll
                    Else
                        Me.TxtCashierCashAccountCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTerminalCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCashierCashAccountCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCashierCashAccountCode.Enabled = True Then
                    If Me.TxtCashierCashAccountCode.ReadOnly = False Then
                        Me.TxtCashierCashAccountCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtCashierCashAccountCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBalanceCashAccountCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBalanceCashAccountCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBalanceCashAccountCode.Enabled = True Then
                    If Me.TxtBalanceCashAccountCode.ReadOnly = False Then
                        Me.TxtBalanceCashAccountCode.BackColor = Me.BackColorAll
                    Else
                        Me.TxtBalanceCashAccountCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBalanceCashAccountCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBalanceCashAccountCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBalanceCashAccountCode.Enabled = True Then
                    If Me.TxtBalanceCashAccountCode.ReadOnly = False Then
                        Me.TxtBalanceCashAccountCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtBalanceCashAccountCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCashierCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtCashierCode.KeyDown
        Dim objCode As New ClsUser
        Dim i As New FrmCariMasterUser
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCashierCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtCashierCode.Text) = True Then
                        Me.TxtCashierName.Text = objCode.GetNama(Me.TxtCashierCode.Text)
                        Me.TxtBranchCode.Focus()
                    Else
                        Me.btnFindCashier.PerformClick()
                    End If
                Else
                    Me.btnFindCashier.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCashierCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtCashierCode.Leave
        Dim objCode As New ClsUser
        Dim i As New FrmCariMasterUser
        'i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCashierCode.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtCashierCode.Text) = True Then
                    Me.TxtCashierName.Text = objCode.GetNama(Me.TxtCashierCode.Text)
                Else
                    If Me.TxtCashierCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCashierCode.Text
                    End If

                    i.ShowDialog()
                    If String.IsNullOrEmpty(i.Kodeuser) = False Then
                        Me.TxtCashierCode.Text = i.Kodeuser
                        Me.TxtCashierName.Text = objCode.GetNama(Me.TxtCashierCode.Text)
                    Else
                        Me.TxtCashierCode.Clear()
                        Me.TxtCashierName.Clear()
                        Me.TxtPrimaryCashier.Text = "-1"
                        Me.TxtCashierCode.Focus()
                    End If
                End If
            Else
                Me.TxtCashierCode.Clear()
                Me.TxtCashierName.Clear()
                Me.TxtPrimaryCashier.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindCashier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCashier.Click
        Dim i As New FrmCariMasterUser
        Dim ObjCode As New ClsUser
        'i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCashierCode.Text <> "" Then
                i.TxtCari.Text = Me.TxtCashierCode.Text
            End If

            i.ShowDialog()
            If String.IsNullOrEmpty(i.Kodeuser) = False Then
                Me.TxtCashierCode.Text = i.Kodeuser
                Me.TxtCashierName.Text = ObjCode.GetNama(Me.TxtCashierCode.Text)
                Me.TxtBranchCode.Focus()
            Else
                Me.TxtCashierCode.Clear()
                Me.TxtCashierName.Clear()
                Me.TxtPrimaryCashier.Text = "-1"
                Me.TxtCashierCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtBranchCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtBranchCode.KeyDown
        Dim objCode As New ClsMast_Cabang
        Dim i As New FrmMast_Cabang_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.TxtBranchCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtBranchCode.Text) = True Then
                        Me.TxtPrimaryBranch.Text = objCode.GetPrimary(Me.TxtBranchCode.Text)
                        Me.TxtBranchCode.Text = objCode.GetCode(Me.TxtPrimaryBranch.Text)
                        Me.TxtBranchName.Text = objCode.GetDesc(Me.TxtPrimaryBranch.Text)
                        Me.TxtRemarks.Focus()
                    Else
                        Me.BtnFindBranchCode.PerformClick()
                    End If
                Else
                    Me.BtnFindBranchCode.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtBranchCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtBranchCode.Leave
        Dim objCode As New ClsMast_Cabang
        Dim i As New FrmMast_Cabang_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtBranchCode.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtBranchCode.Text) = True Then
                    Me.TxtPrimaryBranch.Text = objCode.GetPrimary(Me.TxtBranchCode.Text)
                    Me.TxtBranchCode.Text = objCode.GetCode(Me.TxtPrimaryBranch.Text)
                    Me.TxtBranchName.Text = objCode.GetDesc(Me.TxtPrimaryBranch.Text)
                Else
                    If Me.TxtBranchCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtBranchCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimary > 0 Then
                        Me.TxtPrimaryBranch.Text = i.mPrimary
                        Me.TxtBranchCode.Text = objCode.GetCode(Me.TxtPrimaryBranch.Text)
                        Me.TxtBranchName.Text = objCode.GetDesc(Me.TxtPrimaryBranch.Text)
                    Else
                        Me.TxtBranchCode.Clear()
                        Me.TxtBranchName.Clear()
                        Me.TxtPrimaryBranch.Text = "-1"
                        Me.TxtBranchCode.Focus()
                    End If
                End If
            Else
                Me.TxtBranchCode.Clear()
                Me.TxtBranchName.Clear()
                Me.TxtPrimaryBranch.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindBranchCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindBranchCode.Click
        Dim i As New FrmMast_Cabang_Cari
        Dim ObjCode As New ClsMast_Cabang
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtBranchCode.Text <> "" Then
                i.txtCari.Text = Me.TxtBranchCode.Text
            End If

            i.ShowDialog()
            If i.mPrimary > 0 Then
                Me.TxtPrimaryBranch.Text = i.mPrimary
                Me.TxtBranchCode.Text = ObjCode.GetCode(Me.TxtPrimaryBranch.Text)
                Me.TxtBranchName.Text = ObjCode.GetDesc(Me.TxtPrimaryBranch.Text)
                Me.TxtBranchCode.Focus()
            Else
                Me.TxtBranchCode.Clear()
                Me.TxtBranchName.Clear()
                Me.TxtPrimaryBranch.Text = "-1"
                Me.TxtBranchCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtBeginingCashAccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtBeginingCashAccountCode.KeyDown
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.TxtBeginingCashAccountCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtBeginingCashAccountCode.Text) = True Then
                        Me.TxtPrimaryBeginingCashAccount.Text = objCode.GetPrimary(Me.TxtBeginingCashAccountCode.Text)
                        Me.TxtBeginingCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryBeginingCashAccount.Text)
                        Me.TxtBeginingCashAccountDesc.Text = objCode.GetDescription(Me.TxtBeginingCashAccountCode.Text)
                        Me.TxtCashierCashAccountCode.Focus()
                    Else
                        Me.btnFindBeginingCashAccount.PerformClick()
                    End If
                Else
                    Me.btnFindBeginingCashAccount.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtBeginingCashAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtBeginingCashAccountCode.Leave
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtBeginingCashAccountCode.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtBeginingCashAccountCode.Text) = True Then
                    Me.TxtPrimaryBeginingCashAccount.Text = objCode.GetPrimary(Me.TxtBeginingCashAccountCode.Text)
                    Me.TxtBeginingCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryBeginingCashAccount.Text)
                    Me.TxtBeginingCashAccountDesc.Text = objCode.GetDescription(Me.TxtBeginingCashAccountCode.Text)
                Else
                    If Me.TxtBeginingCashAccountCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtBeginingCashAccountCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryBeginingCashAccount.Text = i.mPrimaryAkun
                        Me.TxtBeginingCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryBeginingCashAccount.Text)
                        Me.TxtBeginingCashAccountDesc.Text = objCode.GetDescription(Me.TxtBeginingCashAccountCode.Text)
                    Else
                        Me.TxtBeginingCashAccountCode.Clear()
                        Me.TxtBeginingCashAccountDesc.Clear()
                        Me.TxtPrimaryBeginingCashAccount.Text = "-1"
                        Me.TxtBeginingCashAccountCode.Focus()
                    End If
                End If
            Else
                Me.TxtBeginingCashAccountCode.Clear()
                Me.TxtBeginingCashAccountDesc.Clear()
                Me.TxtPrimaryBeginingCashAccount.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindBeginningCashAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindBeginingCashAccount.Click
        Dim i As New FrmMasterAkun_Cari
        Dim ObjCode As New ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtBeginingCashAccountCode.Text <> "" Then
                i.txtCari.Text = Me.TxtBeginingCashAccountCode.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryBeginingCashAccount.Text = i.mPrimaryAkun
                Me.TxtBeginingCashAccountCode.Text = ObjCode.GetAccountCode(Me.TxtPrimaryBeginingCashAccount.Text)
                Me.TxtBeginingCashAccountDesc.Text = ObjCode.GetDescription(Me.TxtBeginingCashAccountCode.Text)
                Me.TxtCashierCashAccountCode.Focus()
            Else
                Me.TxtBeginingCashAccountCode.Clear()
                Me.TxtBeginingCashAccountDesc.Clear()
                Me.TxtPrimaryBeginingCashAccount.Text = "-1"
                Me.TxtBeginingCashAccountCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCashierCashAccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtCashierCashAccountCode.KeyDown
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCashierCashAccountCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtCashierCashAccountCode.Text) = True Then
                        Me.TxtPrimaryCashierCashAccount.Text = objCode.GetPrimary(Me.TxtCashierCashAccountCode.Text)
                        Me.TxtCashierCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryCashierCashAccount.Text)
                        Me.TxtCashierCashAccountDesc.Text = objCode.GetDescription(Me.TxtCashierCashAccountCode.Text)
                        Me.TxtBalanceCashAccountCode.Focus()
                    Else
                        Me.btnFindCashierCashAccount.PerformClick()
                    End If
                Else
                    Me.btnFindCashierCashAccount.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCashierCashAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtCashierCashAccountCode.Leave
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCashierCashAccountCode.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtCashierCashAccountCode.Text) = True Then
                    Me.TxtPrimaryCashierCashAccount.Text = objCode.GetPrimary(Me.TxtCashierCashAccountCode.Text)
                    Me.TxtCashierCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryCashierCashAccount.Text)
                    Me.TxtCashierCashAccountDesc.Text = objCode.GetDescription(Me.TxtCashierCashAccountCode.Text)
                Else
                    If Me.TxtCashierCashAccountCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCashierCashAccountCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryCashierCashAccount.Text = i.mPrimaryAkun
                        Me.TxtCashierCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryCashierCashAccount.Text)
                        Me.TxtCashierCashAccountDesc.Text = objCode.GetDescription(Me.TxtCashierCashAccountCode.Text)
                    Else
                        Me.TxtCashierCashAccountCode.Clear()
                        Me.TxtCashierCashAccountDesc.Clear()
                        Me.TxtPrimaryCashierCashAccount.Text = "-1"
                        Me.TxtCashierCashAccountCode.Focus()
                    End If
                End If
            Else
                Me.TxtCashierCashAccountCode.Clear()
                Me.TxtCashierCashAccountDesc.Clear()
                Me.TxtPrimaryCashierCashAccount.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindCashierCashAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnFindCashierCashAccount.Click
        Dim i As New FrmMasterAkun_Cari
        Dim ObjCode As New ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCashierCashAccountCode.Text <> "" Then
                i.txtCari.Text = Me.TxtCashierCashAccountCode.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryCashierCashAccount.Text = i.mPrimaryAkun
                Me.TxtCashierCashAccountCode.Text = ObjCode.GetAccountCode(Me.TxtPrimaryCashierCashAccount.Text)
                Me.TxtCashierCashAccountDesc.Text = ObjCode.GetDescription(Me.TxtCashierCashAccountCode.Text)
                Me.TxtBalanceCashAccountCode.Focus()
            Else
                Me.TxtCashierCashAccountCode.Clear()
                Me.TxtCashierCashAccountDesc.Clear()
                Me.TxtPrimaryCashierCashAccount.Text = "-1"
                Me.TxtCashierCashAccountCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtBalanceCashAccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtBalanceCashAccountCode.KeyDown
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.TxtBalanceCashAccountCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtBalanceCashAccountCode.Text) = True Then
                        Me.TxtPrimaryBalanceCashAccount.Text = objCode.GetPrimary(Me.TxtBalanceCashAccountCode.Text)
                        Me.TxtBalanceCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryBalanceCashAccount.Text)
                        Me.TxtBalanceCashAccountDesc.Text = objCode.GetDescription(Me.TxtBalanceCashAccountCode.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.btnFindBalanceCashAccount.PerformClick()
                    End If
                Else
                    Me.btnFindBalanceCashAccount.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtBalanceCashAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtBalanceCashAccountCode.Leave
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtBalanceCashAccountCode.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtBalanceCashAccountCode.Text) = True Then
                    Me.TxtPrimaryBalanceCashAccount.Text = objCode.GetPrimary(Me.TxtBalanceCashAccountCode.Text)
                    Me.TxtBalanceCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryBalanceCashAccount.Text)
                    Me.TxtBalanceCashAccountDesc.Text = objCode.GetDescription(Me.TxtBalanceCashAccountCode.Text)
                Else
                    If Me.TxtBalanceCashAccountCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtBalanceCashAccountCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryBalanceCashAccount.Text = i.mPrimaryAkun
                        Me.TxtBalanceCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryBalanceCashAccount.Text)
                        Me.TxtBalanceCashAccountDesc.Text = objCode.GetDescription(Me.TxtBalanceCashAccountCode.Text)
                    Else
                        Me.TxtBalanceCashAccountCode.Clear()
                        Me.TxtBalanceCashAccountDesc.Clear()
                        Me.TxtPrimaryBalanceCashAccount.Text = "-1"
                        Me.TxtBalanceCashAccountCode.Focus()
                    End If
                End If
            Else
                Me.TxtBalanceCashAccountCode.Clear()
                Me.TxtBalanceCashAccountDesc.Clear()
                Me.TxtPrimaryBalanceCashAccount.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindBalanceCashAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnFindBalanceCashAccount.Click
        Dim i As New FrmMasterAkun_Cari
        Dim ObjCode As New ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtBalanceCashAccountCode.Text <> "" Then
                i.txtCari.Text = Me.TxtBalanceCashAccountCode.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryBalanceCashAccount.Text = i.mPrimaryAkun
                Me.TxtBalanceCashAccountCode.Text = ObjCode.GetAccountCode(Me.TxtPrimaryBalanceCashAccount.Text)
                Me.TxtBalanceCashAccountDesc.Text = ObjCode.GetDescription(Me.TxtBalanceCashAccountCode.Text)
                Me.btnOK.Focus()
            Else
                Me.TxtBalanceCashAccountCode.Clear()
                Me.TxtBalanceCashAccountDesc.Clear()
                Me.TxtPrimaryBalanceCashAccount.Text = "-1"
                Me.TxtBalanceCashAccountCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCash_Beginning_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCash_Beginning.Enter
        Try
            ObjCD.Number_Enter(Me.TxtCash_Beginning, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCash_Beginning_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCash_Beginning.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.TxtCash_Beginning, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCash_Beginning_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCash_Beginning.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.TxtCash_Beginning, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCash_Beginning_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCash_Beginning.Leave
        Try
            ObjCD.Number_Leave(Me.TxtCash_Beginning, sender, e, Me.DefPrec, ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class