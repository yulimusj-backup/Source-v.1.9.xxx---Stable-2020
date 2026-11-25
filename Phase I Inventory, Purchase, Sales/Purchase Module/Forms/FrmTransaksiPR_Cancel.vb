Option Explicit On
Option Strict Off

Public Class FrmTransaksiPR_Cancel

    Public UserCode As String = ""
    Public BAckColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vReadyForPurchase As ClsTransaksiPR.EnumReadyForPurchase

    Dim ObjPRCancel As New ClsTransaksiPR_Cancel
    Dim ObjPR As New ClsTransaksiPR
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjRefParty As New Core_Module.ClsMasterRefParty

    Dim baru As Boolean
    Dim Buffer As String = ""
    Dim DefPrec As Integer = 0
    Dim EraseRowIndex As Integer = 0

    Private Sub AktifkanDetail()
        Try
            Core_Module.ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.TLP1.Controls, True, True, True, True, True, True, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.TLP2.Controls, True, True, True, True, True, True, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryPR.BackColor, False)
            ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryPR.BackColor, False)
            ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryPR.BackColor, False)
            ObjCD.SetBackColor(Me.TLP1.Controls, Me.TxtPrimaryPR.BackColor, False)
            ObjCD.SetBackColor(Me.TLP2.Controls, Me.TxtPrimaryPR.BackColor, False)
            ObjCD.SetBackColor(Me.Controls, Me.TxtPrimaryPR.BackColor, False)
            Me.SplitContainer1.TabStop = False
            If Me.DgvItem.RowCount >= 1 Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            Else
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            Me.DgvItem.AllowUserToDeleteRows = True
            Me.DgvItem.ReadOnly = False
            If baru = True Then
                Me.DtpTglTransaksi.Value = Now
                Me.DtpTglTransaksi.Focus()
                Me.DgvItem.TabStop = True
            End If

            Me.KondisiPRNumber(False)
            Me.ObjCD.SetNumber(Me.DgvItem, False)
            Me.DtpTglTransaksi.Focus()
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Core_Module.ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Core_Module.ModCekData.KendalikanControl(Me.TLP1.Controls, True, False, True, False, True, False, True, False)
            Core_Module.ModCekData.KendalikanControl(Me.TLP2.Controls, True, False, True, False, True, False, True, False)
            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)

            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TLP1.Controls, TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TLP2.Controls, TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, TxtReadOnly.BackColor, False)
            Me.SplitContainer1.TabStop = True
            Me.DgvItem.ReadOnly = True
            Me.DgvItem.AllowUserToAddRows = False
            Me.DgvItem.AllowUserToDeleteRows = False

            Me.CancelButton = Me.btnExit
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ObjCD.KosongkanControl(Me.gboDetail.Controls, True, False, True)
            ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, False, True)
            ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, False, True)
            Me.DgvItem.Rows.Clear()
            Me.TxtPrimary.Text = "-1"
            Me.TxtPrimaryPR.Text = "-1"
            Me.TxtPrimaryRefParty.Text = "-1"

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

        Dim KoleksiNamaData1 As New Collection
        Dim KoleksiTextBox1 As New Collection

        Try
            '''cek data yang harus diisi
            KoleksiNamaData.Add("Document No")
            KoleksiTextBox.Add(Me.TxtNoTransaksi)
            KoleksiNamaData.Add("PR Number ")
            KoleksiTextBox.Add(Me.txtNoPR)
            KoleksiNamaData.Add("Reason")
            KoleksiTextBox.Add(Me.txtReason)

            If Core_Module.ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '''khusus untuk remarks saja
            KoleksiNamaData1.Add("Remarks")
            KoleksiTextBox1.Add(Me.TxtRemarks)

            If Me.CekDataRemarks(KoleksiTextBox1, KoleksiNamaData1) = False Then
                Return False
            End If

            '''cek data yang salah, double dalam kode
            If Me.baru = True Then
                'If ObjPRCancel.ApakahSudahAda(Me.TxtNoTransaksi.Text) = True Then
                '    MsgBox("Duplication in document no, please change your document number", _
                '    MsgBoxStyle.OkOnly, "Information")
                '    Me.DtpTglTransaksi.Focus()
                '    Return False
                'End If
            Else
                'Dim ProductGroupLama As String = Me.ObjPRCancel.GetNoTransaksi(Me.TxtPrimary.Text)
                'If ProductGroupLama <> Me.TxtNoTransaksi.Text Then
                '    If Me.ObjPRCancel.ApakahSudahAda(Me.TxtNoTransaksi.Text) = True Then
                '        MsgBox("Duplication in document no, please change your document number", _
                '        MsgBoxStyle.OkOnly, "Information")
                '        Me.DtpTglTransaksi.Focus()
                '        Return False
                '    End If
                'End If

                '''data salah jika nilai timestamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                '''proses edit
                Dim OldTimeStamp As Date = Me.ObjPRCancel.GetTimestamp(Me.TxtPrimary.Text)
                If Me.ObjPRCancel.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjPRCancel.CariDetail(Me.TxtPrimary.Text, Me.DgvItem, Me.ObjCD, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            '''Item tidak boleh kosong 
            If Me.DgvItem.Rows.Count = 1 Then
                For i As Integer = 0 To DgvItem.RowCount - 1
                    If Me.DgvItem.Item(Me.ColItem_CodeNo.Index, i).Value = "" Then
                        MsgBox("Please input product code", MsgBoxStyle.Information, "Information ")
                        Return False
                    End If
                Next
            End If

            If Me.ApakahDataItemBenar() = False Then
                Return False
            End If

            If Me.ObjPRCancel.ApakahTanggalSudahBenar(Me.TxtPrimaryPR.Text, Me.DtpTglTransaksi.Value) = False Then
                Me.DtpTglTransaksi.Focus()
                Return False
            End If

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
            With Me.ObjPRCancel
                .Primary = Me.TxtPrimary.Text
                .PrimaryPR = Me.TxtPrimaryPR.Text
                .Tanggal = Me.DtpTglTransaksi.Value
                .NoTransaksi = Me.TxtNoTransaksi.Text
                .Reason = Me.txtReason.Text
                .Remarks = Me.TxtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjPRCancel
                Me.TxtPrimary.Text = .Primary
                Me.TxtPrimaryPR.Text = .PrimaryPR
                Me.DtpTglTransaksi.Text = .Tanggal
                Me.TxtNoTransaksi.Text = .NoTransaksi
                Me.txtNoPR.Text = Me.ObjPR.GetNoTransaksi(.PrimaryPR)
                Me.TxtPrimaryRefParty.Text = Me.ObjPR.GetPrimaryRefParty(.PrimaryPR)
                Me.TxtRequestBy.Text = Me.ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                Me.txtDesc.Text = Me.ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                Me.txtReason.Text = .Reason
                Me.TxtRemarks.Text = .Remarks
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
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
                    '''disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjPRCancel.SimpanBaru(Me.DgvItem, Me.UserCode, Me, Me.baru, Me.TxtPrimary.Text, Me.TxtPrimaryPR.Text) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.BtnFind.Focus()
                        End If
                    End If
                Else
                    '''disimpan atau tidak
                    If Me.ObjPRCancel.ApakahBolehEdit(Me, Me.TxtPrimary.Text, Me.UserCode, True) = True Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.ObjPRCancel.SimpanPerubahan(Me.DgvItem, Me.UserCode, Me, Me.baru, Me.TxtPrimary.Text, Me.TxtPrimaryPR.Text) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            End If
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
            If Me.TxtRequestBy.Text <> "" Or Me.DgvItem.RowCount > 1 Or Me.txtNoPR.Text <> "" Or Me.TxtNoTransaksi.Text <> "" Then
                ''' batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjPRCancel.CariDetail(Me.TxtPrimary.Text, Me.DgvItem, Me.ObjCD, Me)
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

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFind.Click
        Dim hj As New FrmTransaksiPR_Cancel_Cari
        Try
            hj.BackColorAll = Me.BAckColorAll
            hj.ShowDialog()

            If Not hj.PrimaryPRCancel = -1 Then
                If Me.ObjPRCancel.CariDetail(hj.PrimaryPRCancel, Me.DgvItem, Me.ObjCD, Me) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.BtnFind.Focus()
                End If
                Me.BtnFind.Focus()
                Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.DtpTglTransaksi.Focus()

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.txtNoPR.Text <> "" And Me.TxtPrimary.Text <> "-1" Then
                If Me.ObjPRCancel.ApakahBolehEdit(Me, Me.TxtPrimary.Text, UserCode, True) = True Then
                    Me.baru = False
                    Me.ObjPRCancel.CariDetail(Me.TxtPrimary.Text, Me.DgvItem, Me.ObjCD, Me)
                    Me.AktifkanDetail()
                    Me.SalinKeField()
                    Me.DgvItem.ClearSelection()
                    Me.DtpTglTransaksi.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtPrimary.Text = "-1" Then
                Me.BtnFind.PerformClick()
            End If
            If Me.txtNoPR.Text <> "" And Me.TxtPrimary.Text <> "-1" Then
                If Me.ObjPRCancel.ApakahBolehHapus(Me.TxtPrimary.Text, UserCode, Me, True) = False Then
                    Me.BtnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjPRCancel.Hapus(Me.TxtPrimary.Text, Me.UserCode, Me) = True Then
                            Me.KosongkanField()
                            Me.NonAktifkanDetail()
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

    Private Sub FrmMasterLoad_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtNoTransaksi.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            ''' simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPRCancel.SimpanBaru(Me.DgvItem, Me.UserCode, Me, Me.baru, Me.TxtPrimary.Text, Me.TxtPrimaryPR.Text) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.BtnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            '''' simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPRCancel.SimpanPerubahan(Me.DgvItem, Me.UserCode, Me, Me.baru, Me.TxtPrimary.Text, Me.TxtPrimaryPR.Text) = True Then
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
        Finally
        End Try
    End Sub

    Private Sub TxtReason_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus, _
                                    txtNoPR.GotFocus, txtReason.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.BAckColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRequestBy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus, _
                                        txtNoPR.LostFocus, txtReason.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimary.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '''Cari Data Keydown & Leave & Click 
    Private Sub txtNoPR_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNoPR.KeyDown
        Dim objPR As New ClsTransaksiPR
        Dim jk As New ClsTransaksiPR_Item
        Try
            If Me.txtNoPR.ReadOnly = True Then Exit Sub
            If e.KeyValue = 13 Then
                If Me.txtNoPR.Text <> "" Then
                    If objPR.ApakahSudahAda4(ClsTransaksiPR.EnumReadyForPurchase.Yes, Me.txtNoPR.Text) = True Then
                        Me.TxtPrimaryPR.Text = objPR.GetPrimary(Me.txtNoPR.Text)
                        Me.txtNoPR.Text = objPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                        Me.TxtPrimaryRefParty.Text = objPR.GetPrimaryRefParty(Me.TxtPrimaryPR.Text)
                        Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                        Me.txtDesc.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                        Me.txtReason.Focus()
                    Else
                        Me.btnFindPR.PerformClick()
                    End If
                Else
                    Me.btnFindPR.PerformClick()
                End If
            End If

            If Me.TxtPrimaryPR.Text <> "-1" Then
                Me.ObjPRCancel.IsiDGV2(Me.TxtPrimaryPR.Text, Me.TxtPrimary.Text, Me.DgvItem, Me, False)
                'jk.IsiDGVItem2(Me.TxtPrimaryPR.Text, Me.DgvItem, Me, Me.ObjCD)
                Me.KondisiPRNumber(True)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoPR_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoPR.Leave
        Dim ObjPR As New ClsTransaksiPR
        Dim hj As New FrmTransaksiPR_Cari
        Dim jk As New ClsTransaksiPR_Item
        hj.vRFP = True
        hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.txtNoPR.ReadOnly = True Then Exit Sub

            If Me.txtNoPR.Text <> "" Then
                If ObjPR.ApakahSudahAda4(ClsTransaksiPR.EnumReadyForPurchase.Yes, Me.txtNoPR.Text) = True Then
                    Me.TxtPrimaryPR.Text = ObjPR.GetPrimary(Me.txtNoPR.Text)
                    Me.txtNoPR.Text = ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                    Me.TxtPrimaryRefParty.Text = ObjPR.GetPrimaryRefParty(Me.TxtPrimaryPR.Text)
                    Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                    Me.txtDesc.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                    Me.txtReason.Focus()
                Else
                    If Me.txtNoPR.Text <> "" Then
                        hj.txtCari.Text = Me.txtNoPR.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryPR > 0 Then
                        Me.TxtPrimaryPR.Text = hj.PrimaryPR
                        Me.txtNoPR.Text = ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                        Me.TxtPrimaryRefParty.Text = ObjPR.GetPrimaryRefParty(Me.TxtPrimaryPR.Text)
                        Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                        Me.txtDesc.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                        Me.txtReason.Focus()
                    Else
                        Me.TxtPrimaryPR.Text = "-1"
                        Me.txtNoPR.Clear()
                        Me.TxtRequestBy.Clear()
                        Me.txtDesc.Clear()
                        Me.txtNoPR.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryPR.Text = "-1"
                Me.txtNoPR.Clear()
                Me.TxtRequestBy.Clear()
                Me.txtDesc.Clear()
            End If
         
            If Me.TxtPrimaryPR.Text <> "-1" Then
                Me.ObjPRCancel.IsiDGV2(Me.TxtPrimaryPR.Text, Me.TxtPrimary.Text, Me.DgvItem, Me, False)
                'jk.IsiDGVItem2(Me.TxtPrimaryPR.Text, Me.DgvItem, Me, Me.ObjCD)
            End If
            If Me.txtNoPR.Text <> "" Then
                Me.KondisiPRNumber(True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindPR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPR.Click

        Dim ObjPR As New ClsTransaksiPR
        Dim hj As New FrmTransaksiPR_Cari
        Dim jk As New ClsTransaksiPR_Item
        hj.vRFP = True
        Try

            hj.BackColorAll = Me.BAckColorAll
            If Me.txtNoPR.Text <> "" Then
                hj.txtCari.Text = Me.txtNoPR.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryPR > 0 Then
                Me.TxtPrimaryPR.Text = hj.PrimaryPR
                Me.txtNoPR.Text = ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                Me.TxtPrimaryRefParty.Text = ObjPR.GetPrimaryRefParty(Me.TxtPrimaryPR.Text)
                Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                Me.txtDesc.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
            Else
                Me.TxtPrimaryPR.Text = "-1"
                Me.txtNoPR.Clear()
                Me.TxtRequestBy.Clear()
                Me.txtDesc.Clear()
                Me.txtNoPR.Focus()
            End If
            If Me.TxtPrimaryPR.Text <> "-1" Then
                Me.ObjPRCancel.IsiDGV2(Me.TxtPrimaryPR.Text, Me.TxtPrimary.Text, Me.DgvItem, Me, False)
                'jk.IsiDGVItem2(Me.TxtPrimaryPR.Text, Me.DgvItem, Me, Me.ObjCD)
            End If
            If Me.txtNoPR.Text <> "" Then
                Me.KondisiPRNumber(True)
            End If
            If Me.TxtPrimaryPR.Text <> "-1" Then
                Me.txtReason.Focus()
            Else
                Me.txtNoPR.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub DtpTglTransaksi_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DtpTglTransaksi.ValueChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.baru = False Then
                    If Me.ObjPRCancel.Tanggal.Month = Me.DtpTglTransaksi.Value.Month And _
                        Me.ObjPRCancel.Tanggal.Year = Me.DtpTglTransaksi.Value.Year Then
                        Me.TxtNoTransaksi.Text = Me.ObjPRCancel.NoTransaksi
                    Else
                        Me.TxtNoTransaksi.Text = Me.ObjPRCancel.GetNoBaru(Me.DtpTglTransaksi.Value)
                    End If
                Else
                    Me.TxtNoTransaksi.Text = Me.ObjPRCancel.GetNoBaru(Me.DtpTglTransaksi.Value)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Try

            If String.IsNullOrEmpty(Me.DgvItem(Me.ColItem_CodeNo.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColItem_CodeNo.Index Or e.ColumnIndex = Me.ColItem_Qty.Index Or _
                                e.ColumnIndex = Me.ColItem_Remarks.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    Me.DgvItem.Columns(Me.ColItem_CodeNo.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvItem.Columns(Me.ColItem_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvItem.Columns(Me.ColItem_Qty.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvItem.Columns(Me.ColItem_Unit.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    '''simpan nilai awal di buffer
                    Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColItem_CodeNo.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    Me.DgvItem.Columns(Me.ColItem_CodeNo.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvItem.Columns(Me.ColItem_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvItem.Columns(Me.ColItem_Qty.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvItem.Columns(Me.ColItem_Unit.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    '''simpan nilai awal di buffer
                    Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Try
            Dim row As DataGridViewRow = Me.DgvItem.Rows(e.RowIndex)
            With row
                If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColItem_Qty.Index Then
                        If IsNumeric(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            If IsNumeric(Me.Buffer) = True Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Else
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = 0
                            End If
                        End If
                    End If

                    If e.ColumnIndex <> Me.ColItem_Remarks.Index Then
                        If e.ColumnIndex = Me.ColItem_Qty.Index Then
                            Me.EndEditQty(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColItem_CodeNo.Index Then
                            Me.EndEditProductCode(sender, e)
                            Exit Sub
                        Else
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColItem_Remarks.Index Then
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
                If e.ColumnIndex = Me.ColItem_Remarks.Index Then
                    If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> "" Then
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Trim(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If
                End If
                If e.ColumnIndex = Me.ColItem_CodeNo.Index Then
                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    Else
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                    Me.ObjCD.SetNumber(Me.DgvItem)
                    Exit Sub
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            '''set header fungsi digunakan ketika edit code 
            Me.DgvItem.Columns(Me.ColItem_CodeNo.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvItem.Columns(Me.ColItem_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvItem.Columns(Me.ColItem_Qty.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvItem.Columns(Me.ColItem_Unit.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vPrimary As Double, _
       ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.DgvItem.Rows.Count - 2
                RowDGV = Me.DgvItem.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColItem_PrimaryCode.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.DgvItem.Rows.RemoveAt(EraseRowIndex)
            Me.ObjCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvItem.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.DgvItem, True)
            Me.KondisiPRNumber(False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

   
    Private Sub DgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            Me.DgvItem.TabStop = True
            If Me.TxtRemarks.Enabled = True Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            Else
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If

            If Me.TxtNoTransaksi.Enabled = True Then
                rowDGV = Me.DgvItem.Rows(Me.DgvItem.RowCount - 1)
                With rowDGV
                    .Cells(Me.ColItem_No.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColItem_PrimaryCode.Index).Value = "-1"
                    .Cells(Me.ColItem_CodeNo.Index).Value = ""
                    .Cells(Me.ColItem_Qty.Index).Value = 0
                    .Cells(Me.ColItem_Unit.Index).Value = ""
                    .Cells(Me.ColItem_Remarks.Index).Value = ""
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            rowDGV = Nothing
        End Try

    End Sub

    '''up & down
    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.SelectionChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.DgvItem.SelectedRows.Count > 0 Then
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                    Me.DgvItem.AllowUserToAddRows = False
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                    Me.DgvItem.AllowUserToAddRows = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowRowProcess(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.DgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.DgvItem.FirstDisplayedScrollingRowIndex = Me.DgvItem.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.DgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.DgvItem.Rows(row.Index).Displayed = False
                    Me.DgvItem.FirstDisplayedScrollingRowIndex = Me.DgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.DgvItem.DisplayedRowCount(True) <> Me.DgvItem.DisplayedRowCount(False) Then
                    Me.DgvItem.FirstDisplayedScrollingRowIndex = Me.DgvItem.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.DgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.DgvItem.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.DgvItem.Rows.Remove(row)
                    Me.DgvItem.Rows.InsertRange(indexbaru, row)
                    Me.DgvItem.Focus()
                    Me.DgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.DgvItem)
                    If row.Displayed = False Then
                        Me.ShowRowProcess(row, True)
                    End If
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try

            If Me.DgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.DgvItem.SelectedRows(0)

                If row.Index <> Me.DgvItem.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.DgvItem.Rows.Remove(row)
                    Me.DgvItem.Rows.InsertRange(indexbaru, row)
                    Me.DgvItem.Focus()
                    Me.DgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.DgvItem)
                    Me.ShowRowProcess(row, True)
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvItem.RowsRemoved
        Try
            If Me.DgvItem.Rows.Count = 0 Then
                Me.DgvItem.TabStop = False
                If Me.TxtRemarks.Enabled = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    '' fungsi untuk product code
    Private Sub EndEditProductCode(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.DgvItem.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryProduct As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            If Me.txtNoPR.Text <> "" Then
                With row
                    If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If e.ColumnIndex = Me.ColItem_CodeNo.Index Then
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                            PrimaryProduct = Me.ObjProduct.GetPrimary(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value)

                            '''cek di masterproduct
                            If Me.ObjProduct.ApakahSudahAda(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                '''cek di transaksi pr
                                If Me.ObjPR.ApakahSudahAda5(Me.TxtPrimaryPR.Text, PrimaryProduct) = True Then
                                    '''cek nilai OB gg boleh = 0
                                    If Me.ObjPR.ApakahSudahAda6(Me.TxtPrimaryPR.Text, PrimaryProduct) = True Then
                                        index = Me.GetIndexItem(PrimaryProduct, e.RowIndex)
                                        If index < 0 Then
                                            Me.DgvItem.Item(Me.ColItem_PrimaryCode.Index, e.RowIndex).Value = PrimaryProduct
                                            Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(PrimaryProduct)
                                            Me.DgvItem.Item(Me.ColItem_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(PrimaryProduct)
                                            Me.DgvItem.Item(Me.ColItem_Unit.Index, e.RowIndex).Value = Me.ObjProduct.GetUnitCode(PrimaryProduct)
                                            Me.DgvItem.Item(Me.ColItem_Qty.Index, e.RowIndex).Value = Me.ObjPR.GetOB(Me.TxtPrimaryPR.Text, PrimaryProduct)
                                        Else
                                            If e.RowIndex <> index Then
                                                MsgBox("Duplication code, please input another code.", MsgBoxStyle.OkOnly, "Information")
                                                Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = ""
                                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                    Me.EraseRowIndex = e.RowIndex
                                                    Me.BackgroundWorker1.RunWorkerAsync()
                                                Else
                                                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                                End If
                                                Me.ObjCD.SetNumber(Me.DgvItem)
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        If e.RowIndex <> index Then
                                            MsgBox("This product code qty value is zero please choose another code no.", MsgBoxStyle.OkOnly, "Information")
                                            Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = ""
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowIndex = e.RowIndex
                                                Me.BackgroundWorker1.RunWorkerAsync()
                                            Else
                                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            Me.ObjCD.SetNumber(Me.DgvItem)
                                            Exit Sub
                                        End If
                                    End If
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Product code is not listed on the transaction number of the purchase request", MsgBoxStyle.OkOnly, "Information")
                                        Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvItem)
                                        Exit Sub
                                    End If
                                End If
                            Else
                                '''kalo tidak ada di master product
                                Dim ij As New Core_Module.frmCariMasterProduct
                                ij.txtCari.Text = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BAckColorAll
                                ij.ShowDialog()
                                If ij.PrimaryProduct <> -1 Then
                                    index = Me.GetIndexItem(ij.PrimaryProduct, e.RowIndex)
                                    If index < 0 Then
                                        '''cek di transaksi pr
                                        If Me.ObjPR.ApakahSudahAda5(Me.TxtPrimaryPR.Text, ij.PrimaryProduct) = True Then
                                            '''cek nilai OB gg boleh = 0
                                            If Me.ObjPR.ApakahSudahAda6(Me.TxtPrimaryPR.Text, ij.PrimaryProduct) = True Then
                                                Me.DgvItem.Item(Me.ColItem_PrimaryCode.Index, e.RowIndex).Value = ij.PrimaryProduct
                                                Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                                                Me.DgvItem.Item(Me.ColItem_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                                                Me.DgvItem.Item(Me.ColItem_Qty.Index, e.RowIndex).Value = Me.ObjPR.GetOB(Me.TxtPrimaryPR.Text, ij.PrimaryProduct)
                                                Me.DgvItem.Item(Me.ColItem_Unit.Index, e.RowIndex).Value = Me.ObjProduct.GetUnitCode(ij.PrimaryProduct)
                                            Else
                                                If e.RowIndex <> index Then
                                                    MsgBox("This product code qty value is zero please find another", MsgBoxStyle.OkOnly, "Information")
                                                    Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = ""
                                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                        Me.EraseRowIndex = e.RowIndex
                                                        Me.BackgroundWorker1.RunWorkerAsync()
                                                    Else
                                                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                                    End If
                                                    Me.ObjCD.SetNumber(Me.DgvItem)
                                                    Exit Sub
                                                End If
                                            End If
                                        Else
                                            If e.RowIndex <> index Then
                                                MsgBox("Code no is not listed on the purchase request", MsgBoxStyle.OkOnly, "Information")
                                                Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = ""
                                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                    Me.EraseRowIndex = e.RowIndex
                                                    Me.BackgroundWorker1.RunWorkerAsync()
                                                Else
                                                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                                End If
                                                Me.ObjCD.SetNumber(Me.DgvItem)
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        If e.RowIndex <> index Then
                                            MsgBox("Duplication code, please input another code.", MsgBoxStyle.OkOnly, "Information")
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowIndex = e.RowIndex
                                                Me.BackgroundWorker1.RunWorkerAsync()
                                            Else
                                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            Me.ObjCD.SetNumber(Me.DgvItem)
                                            Exit Sub
                                        Else
                                            '''cek ditransaksi pr
                                            If Me.ObjPR.ApakahSudahAda5(Me.TxtPrimaryPR.Text, ij.PrimaryProduct) = True Then
                                                '''cek nilai OB gg boleh = 0
                                                If Me.ObjPR.ApakahSudahAda6(Me.TxtPrimaryPR.Text, ij.PrimaryProduct) = True Then
                                                    Me.DgvItem.Item(Me.ColItem_PrimaryCode.Index, e.RowIndex).Value = ij.PrimaryProduct
                                                    Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                                                    Me.DgvItem.Item(Me.ColItem_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                                                    Me.DgvItem.Item(Me.ColItem_Qty.Index, e.RowIndex).Value = Me.ObjPR.GetOB(Me.TxtPrimaryPR.Text, ij.PrimaryProduct)
                                                    Me.DgvItem.Item(Me.ColItem_Unit.Index, e.RowIndex).Value = Me.ObjProduct.GetUnitCode(ij.PrimaryProduct)
                                                Else
                                                    If e.RowIndex <> index Then
                                                        MsgBox("This product code qty value is zero please find another", MsgBoxStyle.OkOnly, "Information")
                                                        Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = ""
                                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                            Me.EraseRowIndex = e.RowIndex
                                                            Me.BackgroundWorker1.RunWorkerAsync()
                                                        Else
                                                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                                        End If
                                                        Me.ObjCD.SetNumber(Me.DgvItem)
                                                        Exit Sub
                                                    End If
                                                End If
                                            Else
                                                If e.RowIndex <> index Then
                                                    MsgBox("Product code is not listed on the transaction number of the purchase request", MsgBoxStyle.OkOnly, "Information")
                                                    Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = ""
                                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                        Me.EraseRowIndex = e.RowIndex
                                                        Me.BackgroundWorker1.RunWorkerAsync()
                                                    Else
                                                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                                    End If
                                                    Me.ObjCD.SetNumber(Me.DgvItem)
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                ElseIf ij.PrimaryProduct = -1 Then
                                    If Me.DgvItem.Rows.Count - 1 <> e.RowIndex Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvItem)
                                        Exit Sub
                                    End If
                                    If Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value <> "" Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvItem)
                                        Exit Sub
                                    End If
                                End If
                                ij = Nothing
                            End If
                            Me.ObjCD.SetNumber(Me.DgvItem)
                        End If
                    Else
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End With
            Else
                MsgBox("Please select the first number of purchase requests", MsgBoxStyle.OkOnly, "Information")
                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Me.DgvItem.AllowUserToAddRows = False
                Exit Sub
            End If

            If Me.txtNoPR.Text <> "" Then
                Me.KondisiPRNumber(False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQty(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
     ByVal row As DataGridViewRow)
        Dim vPrimaryCodeNo As Double = -1
        Dim vUnit As String = ""
        vPrimaryCodeNo = Me.ObjProduct.GetPrimary(Me.DgvItem(Me.ColItem_CodeNo.Index, e.RowIndex).Value)
        vUnit = Me.ObjProduct.GetUnitCode(vPrimaryCodeNo)
        Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(vPrimaryCodeNo)

        Try
            With row
                '''untuk nilai qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.OkOnly, _
                    "Information")
                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                Dim DGVQty As Decimal = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Dim OB As Decimal = Me.ObjPR.GetOB(Me.TxtPrimaryPR.Text, vPrimaryCodeNo)
                Dim ItemDB As Decimal = Me.ObjPRCancel.GetQty(Me.TxtPrimary.Text, vPrimaryCodeNo, Me.TxtPrimaryPR.Text)
                Dim Maximumnya As Decimal = 0
                Dim OB_StlhEdit As Decimal = 0

                OB_StlhEdit = OB + ItemDB - DGVQty
                If OB_StlhEdit < 0 Then
                    Maximumnya = DGVQty + OB_StlhEdit
                    MsgBox("Maximum value you can enter is" & " " & Me.ObjCD.DigitGrouping(Maximumnya, DefPrec) & _
                    " " & vUnit.ToLower & " ", MsgBoxStyle.OkOnly, "Information")
                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                '''digit grouping
                Me.DgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.DgvItem(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function HapusSpasi(ByVal Data As String) As String
        Try
            Do While Data.EndsWith(" ") And Data.Length > 0
                Data = Data.Remove(Data.Length - 1, 1)
            Loop
            Return Data
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function CekDataRemarks(ByVal KoleksiTextBox As Collection, ByVal KoleksiNamaData As Collection) As Boolean
        Dim i As Int16
        Dim data As String
        Dim Txt As Windows.Forms.TextBox

        Try
            Do While i < KoleksiTextBox.Count
                '''trim kanan kiri
                Txt = TryCast(KoleksiTextBox.Item(i + 1), Windows.Forms.TextBox)
                If IsNothing(Txt) = False Then
                    Txt.Text = Txt.Text.Trim
                End If
                '''alin data
                data = KoleksiTextBox.Item(i + 1).text
                data = HapusSpasi(data)

                '''cek berkaitan dengan karakter yang tidak diingikan --> ', *, %
                If data.Contains("'") = True Or data.Contains("*") = True Or data.Contains("%") = True Then
                    MsgBox("" + KoleksiNamaData.Item(i + 1) _
                        + " data can not contain the following characters: ' , % , or * .", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                i += 1
            Loop

            '''sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function


    Public Function ApakahAdaItem() As Boolean
        Try
            If Me.DgvItem.RowCount = 1 Then
                For i As Integer = 0 To DgvItem.RowCount - 1
                    If Me.DgvItem.Item(Me.ColItem_CodeNo.Index, i).Value <> "" Then
                        Return True
                    Else
                        Me.DgvItem.AllowUserToAddRows = False
                    End If
                Next
            ElseIf Me.DgvItem.RowCount > 1 Then
                Return True
            Else
                Return False
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Public Sub KondisiPRNumber(Optional ByVal ShowMsg As Boolean = True)
        Try
            If baru = False Then

                If ApakahAdaItem() = True Then
                    Me.txtNoPR.ReadOnly = True
                    Me.btnFindPR.Enabled = False
                    Me.txtNoPR.BackColor = Me.TxtNoTransaksi.BackColor
                    Me.txtNoPR.TabStop = False
                    Me.DgvItem.TabStop = True
                Else
                    If ShowMsg = True Then
                        MsgBox("Can't make transaction cancellation because qty is zero", MsgBoxStyle.OkOnly, "Information")
                    End If
                    Me.txtNoPR.Clear()
                    Me.TxtRequestBy.Clear()
                    Me.txtDesc.Clear()
                    Me.TxtPrimaryPR.Text = "-1"
                    Me.TxtPrimaryRefParty.Text = "-1"
                    Me.txtNoPR.ReadOnly = False
                    Me.txtNoPR.BackColor = Me.TxtRemarks.BackColor
                    Me.txtNoPR.TabStop = True
                    Me.DgvItem.TabStop = False
                    Me.btnFindPR.Enabled = True
                    Me.txtNoPR.Focus()
                End If
            Else
                If ApakahAdaItem() = True Then
                    Me.txtNoPR.ReadOnly = True
                    Me.btnFindPR.Enabled = False
                    Me.txtNoPR.BackColor = Me.TxtNoTransaksi.BackColor
                    Me.txtNoPR.TabStop = False
                    Me.DgvItem.TabStop = True
                Else
                    If ShowMsg = True Then
                        MsgBox("Can't make transaction cancellation because qty is zero", MsgBoxStyle.OkOnly, "Information")
                    End If
                    Me.TxtPrimaryPR.Text = "-1"
                    Me.TxtPrimaryRefParty.Text = "-1"
                    Me.txtNoPR.ReadOnly = False
                    Me.txtNoPR.TabStop = True
                    Me.DgvItem.TabStop = False
                    Me.btnFindPR.Enabled = True
                    Me.txtNoPR.Clear()
                    Me.TxtRequestBy.Clear()
                    Me.txtDesc.Clear()
                    Me.txtNoPR.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
 
    Private Function ApakahDataItemBenar() As Boolean
        Try
            'data item salah bila
            Dim rowDgv As DataGridViewRow

            For Each rowDgv In Me.DgvItem.Rows
                If rowDgv.Cells(Me.ColItem_CodeNo.Index).Value <> "" Then
                    If Trim(rowDgv.Cells(Me.ColItem_Qty.Index).Value) = "" Or _
                            Trim(rowDgv.Cells(Me.ColItem_Qty.Index).Value) <= 0 Then
                        Me.DgvItem.ClearSelection()
                        rowDgv.Selected = True
                        MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.OkOnly, _
                                            "Information")
                        Return False
                    End If
                End If
            Next

            If Me.ObjPRCancel.ApakahQtyBenar(Me.TxtPrimary.Text, Me.TxtPrimaryPR.Text, Me.DgvItem, Me) = False Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim i As New FrmTransaksiPR_Cancel_Cetak
        Try
            If Me.TxtPrimary.Text = -1 Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimary.Text <> -1 Then
                ObjPRCancel.CetakPR_Cancel(Me.TxtPrimary.Text, i.crvReport)
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                i.WindowState = FormWindowState.Maximized
                i.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DgvItem.KeyDown

        Try
            If e.KeyCode = 27 Then
                Me.btnCancel.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub FrmTransaksiPR_Cancel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.btnFind.Select()
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.KosongkanField()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '''Khusus [TAB] tambahan
    Dim SetTab As Boolean = False
    Private Sub btnExit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.GotFocus
        SetTab = True
    End Sub
    Private Sub btnFind_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFind.LostFocus
        SetTab = False
    End Sub
    Private Sub SplitContainer1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles SplitContainer1.GotFocus
        If SetTab = False Then
            btnExit.Focus()
        Else
            btnFind.Focus()
        End If
    End Sub
End Class