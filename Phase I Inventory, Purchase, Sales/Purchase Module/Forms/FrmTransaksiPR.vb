Option Explicit On
Option Strict Off
Public Class FrmTransaksiPR

    Public UserCode As String = ""
    Public BAckColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vTBLItemDelivery As New DseTransaksiPR.tbltrans_purchase_req_itemdeliveryDataTable
    Public vrowItemDelivery As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryRow

    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjPR As New ClsTransaksiPR
    Dim ObjPRItem As New ClsTransaksiPR_Item
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjRefParty As New Core_Module.ClsMasterRefParty
    Dim ObjPriorityLevel As New Core_Module.ClsMasterPriorityLevel

    Dim baru As Boolean
    Dim Buffer As String = ""
    Dim EraseRowIndex As Integer = 0
    Dim vReadyForPurchase As ClsTransaksiPR.EnumReadyForPurchase
    Dim vPrimaryUser As Double = -1
    Dim DefPrec As Integer = 0

    Private Sub AktifkanDetail()
        Try
            ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
             ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryPR.BackColor, False)
            ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryPR.BackColor, False)
            ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryPR.BackColor, False)

            Me.DgvItem.AllowUserToAddRows = True
            Me.DgvItem.AllowUserToDeleteRows = True

            Me.DgvItem.ReadOnly = False
            If Me.TxtRemarks.Enabled = True Then
                Me.ChkRFP.Enabled = False
            End If

            If baru = True Then
                Me.txtNoTransaksi.Text = Me.ObjPR.GetNoBaru(Me.DtpTglTransaksi.Value)
                Me.DtpTglTransaksi.Value = Now
                Me.DtpTglTransaksi.Focus()
                Me.ChkRFP.Checked = False
                Me.DgvItem.TabStop = True
            End If

            Me.SplitContainer1.TabStop = False
            Me.ObjCD.SetNumber(Me.DgvItem, False)
            Me.DtpTglTransaksi.Focus()
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.llbUpdateDeliveryDate.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)

            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, TxtReadOnly.BackColor, False)

            Me.DgvItem.ReadOnly = True
            Me.DgvItem.AllowUserToAddRows = False
            Me.DgvItem.AllowUserToDeleteRows = False
            Me.DgvItem.TabStop = False

            Me.ChkRFP.Enabled = True
            If Me.txtNoTransaksi.Text <> "" Then
                If Me.TxtRemarks.Enabled = False Then
                    Me.ChkRFP.Enabled = Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiPR_Otorisasi")
                Else
                    Me.ChkRFP.Enabled = False
                End If
            Else
                If Me.TxtPrimaryPR.Text = "-1" Then
                    Me.ChkRFP.Checked = False
                End If
            End If

            If Me.ChkRFP.Checked = False Then
                Me.LblOtorisasi.Text = "Authorize By : - "
                Me.TxtPrimaryOtorisasi.Text = -1
            End If
            Me.SplitContainer1.TabStop = True
            Me.CancelButton = Me.btnExit
            Me.llbUpdateDeliveryDate.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ObjCD.KosongkanControl(Me.gboDetail.Controls, True, False, True)
            ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, False, True)
            ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, False, True)
            Me.txtNoTransaksi.Text = ""
            Me.DtpTglTransaksi.Value = Today

            Me.TxtPrimaryPR.Text = "-1"
            Me.DgvItem.Rows.Clear()

            If baru = True Then
                Me.LblOtorisasi.Text = "Authorize By : - "
                Me.TxtPrimaryOtorisasi.Text = "-1"
            End If

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.vTBLItemDelivery.Rows.Clear()
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
            KoleksiNamaData.Add("Request By")
            KoleksiTextBox.Add(Me.TxtRequestBy)

            If Core_Module.ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '''pengecekan remarks jika mengandung tanda operator
            KoleksiNamaData1.Add("Remarks")
            KoleksiTextBox1.Add(Me.TxtRemarks)

            If Me.CekDataRemarks(KoleksiTextBox1, KoleksiNamaData1) = False Then
                Return False
            End If

            '''cek data yang salah, double dalam kode
            If Me.baru = True Then
                'If ObjPR.ApakahSudahAda(Me.txtNoTransaksi.Text) = True Then
                '    MsgBox("Duplication PR Number, please change your PR Number", _
                '    MsgBoxStyle.Information, "Information")
                '    Me.DtpTglTransaksi.Focus()
                '    Return False
                'End If
            Else
                'Dim ProductGroupLama As String = Me.ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                'If ProductGroupLama <> Me.txtNoTransaksi.Text Then
                '    If Me.ObjPR.ApakahSudahAda(Me.txtNoTransaksi.Text) = True Then
                '        MsgBox("Duplication PR Number, please change your PR Number", _
                '        MsgBoxStyle.Information, "Information")
                '        Me.DtpTglTransaksi.Focus()
                '        Return False
                '    End If
                'End If

                '''data salah jika nilai timestamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                '''proses edit
                Dim OldTimeStamp As Date = Me.ObjPR.GetTimestamp(Me.TxtPrimaryPR.Text)
                If Me.ObjPR.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjPR.CariDetail(Me.ObjPR.Primary, Me.DgvItem, Me.ObjCD, Me, Me.vTBLItemDelivery)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            '''data item jika nilai qty masih = 0 
            If Me.ApakahDataItemBenar = False Then
                Return False
            End If

            '''pengecekan jika dgv belum terisi
            If Me.DgvItem.Rows.Count - 1 = 0 Then
                MsgBox("Please input product code", MsgBoxStyle.Information, "Information ")
                Return False
            End If

            '''pengecekan ketika edit tanggal dan request by tidak boleh di edit jika sudah digunakan ditransaksi lain
            '''walaupun yang edit adalah 'USER'
            If baru = False Then
                Dim NoTrans As String = ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                If ObjPR.ApakahSudahAda2(NoTrans, False) = True Or ObjPR.ApakahSudahAda3(NoTrans, False) = True Then
                    If Me.DtpTglTransaksi.Value.Date <> ObjPR.GetTanggalTransaksi(Me.TxtPrimaryPR.Text).Date Then
                        MsgBox("Cannot edit dating because already other transaction.")
                        Me.DtpTglTransaksi.Value = ObjPR.GetTanggalTransaksi(Me.TxtPrimaryPR.Text)
                        Return False
                    End If
                    If Me.TxtPrimaryRefParty.Text <> ObjPR.GetPrimaryRefParty(Me.TxtPrimaryPR.Text) Then
                        MsgBox("Cannot edit request by because already other transaction.")
                        Me.TxtPrimaryRefParty.Text = ObjPR.GetPrimaryRefParty(Me.TxtPrimaryPR.Text)
                        Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                        Return False
                    End If
                End If
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.ObjPR
                .Primary = Me.TxtPrimaryPR.Text
                .NoTransaksi = Me.txtNoTransaksi.Text
                .Tanggal = Me.DtpTglTransaksi.Value

                If Me.ChkRFP.Checked = False Then
                    .RFP = ClsTransaksiPR.EnumReadyForPurchase.No
                Else
                    .RFP = ClsTransaksiPR.EnumReadyForPurchase.Yes
                End If

                .PrimaryUserOtorisasi = TxtPrimaryOtorisasi.Text

                If Me.ChkRFP.Checked = False Then
                    .TimeOtorisasi = #1/1/1950#
                    .PrimaryRefParty = Me.TxtPrimaryRefParty.Text
                Else
                    .TimeOtorisasi = .TimeOtorisasi
                    .PrimaryRefParty = Me.TxtPrimaryRefParty.Text
                End If
                .Remarks = Me.TxtRemarks.Text

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjPR
                Me.TxtPrimaryPR.Text = .Primary
                Me.DtpTglTransaksi.Text = .Tanggal
                Me.txtNoTransaksi.Text = .NoTransaksi
                If .RFP = False Then
                    Me.ChkRFP.Checked = False
                Else
                    Me.ChkRFP.Checked = True
                End If
                Me.TxtPrimaryOtorisasi.Text = .PrimaryUserOtorisasi
                If .PrimaryUserOtorisasi = -1 Then
                    Me.LblOtorisasi.Text = "Authorize By : - "
                Else
                    Me.LblOtorisasi.Text = "Authorize By : " & Me.ObjUser.GetKodeUser(.PrimaryUserOtorisasi) & " on " & _
                    Format(.TimeOtorisasi, "dd MMM yyyy - HH.mm") '+ "-" + .TimeOtorisasi.ToShortTimeString
                End If
                Me.TxtPrimaryRefParty.Text = .PrimaryRefParty
                Me.TxtRequestBy.Text = Me.ObjRefParty.GetCode(.PrimaryRefParty)
                Me.TxtRequestByName.Text = Me.ObjRefParty.GetDesc(.PrimaryRefParty)

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
                        If Me.ObjPR.SimpanBaru(Me.DgvItem, Me.UserCode, Me, Me.vTBLItemDelivery, Me.baru) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    '''disimpan atau tidak
                    If Me.ObjPR.ApakahBolehEdit(Me, Me.TxtPrimaryPR.Text, Me.UserCode, Me.txtNoTransaksi.Text, True) = True Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.ObjPR.SimpanPerubahan(Me.DgvItem, Me.UserCode, Me, Me.vTBLItemDelivery, Me.baru) Then
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
        Finally
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try

            If Me.DgvItem.Rows.Count >= 1 And Me.TxtRequestBy.Text <> "" Then
                ''' batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.NonAktifkanDetail()
                        Me.KosongkanField()
                    Else
                        Me.NonAktifkanDetail()
                        Me.ObjPR.CariDetail(Me.ObjPR.Primary, Me.DgvItem, Me.ObjCD, Me, Me.vTBLItemDelivery)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                    Me.NonAktifkanDetail()
                End If
            Else
                Me.NonAktifkanDetail()
                Me.KosongkanField()
            End If
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTransaksiPR_Cari
        Try
            hj.BackColorAll = Me.BAckColorAll
            hj.ShowDialog()

            If Not hj.PrimaryPR = -1 Then
                If Me.ObjPR.CariDetail(hj.PrimaryPR, Me.DgvItem, Me.ObjCD, Me, Me.vTBLItemDelivery) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
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
            Me.TxtRequestBy.Focus()
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
            If Me.TxtPrimaryPR.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.txtNoTransaksi.Text <> "" Then
                If ObjPR.HakKhususUser(Me.UserCode) = False Then
                    If Me.ChkRFP.Checked = True Then
                        MsgBox("This transaction have been authorize, please cancel authorization first.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    If Me.TxtPrimaryPR.Text <> "-1" And Me.ChkRFP.Checked = False Then
                        If Me.ObjPR.ApakahBolehEdit(Me, Me.TxtPrimaryPR.Text, UserCode, Me.txtNoTransaksi.Text, True) = True Then
                            Me.baru = False
                            Me.ObjPR.CariDetail(Me.ObjPR.Primary, Me.DgvItem, Me.ObjCD, Me, Me.vTBLItemDelivery)
                            Me.AktifkanDetail()
                            Me.SalinKeField()
                            Me.TxtRemarks.Focus()
                        End If
                    End If
                Else
                    If Me.ObjPR.ApakahBolehEdit(Me, Me.TxtPrimaryPR.Text, UserCode, Me.txtNoTransaksi.Text, True) = True Then
                        Me.baru = False
                        Me.ObjPR.CariDetail(Me.ObjPR.Primary, Me.DgvItem, Me.ObjCD, Me, Me.vTBLItemDelivery)
                        Me.AktifkanDetail()
                        Me.SalinKeField()
                        Me.TxtRemarks.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtNoTransaksi.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If ObjPR.HakKhususUser(Me.UserCode) = False Then
                If Me.ChkRFP.Checked = True Then
                    MsgBox("Please cancel the first authorization", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If Me.txtNoTransaksi.Text <> "" And Me.ChkRFP.Checked = False Then
                    If Me.ObjPR.ApakahBolehHapus(Me.TxtPrimaryPR.Text, UserCode, Me, True) = False Then
                        Me.btnFind.Focus()
                    Else
                        If ObjPR.ApakahSudahAda2(Me.txtNoTransaksi.Text, True) = True Then
                            Exit Sub
                        End If
                        If ObjPR.ApakahSudahAda3(Me.txtNoTransaksi.Text, True) = True Then
                            Exit Sub
                        End If

                        hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.ObjPR.Hapus(Me.TxtPrimaryPR.Text, Me.UserCode, Me) = True Then
                                Me.KosongkanField()
                                Me.NonAktifkanDetail()
                                Me.btnNew.Focus()
                                MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                            End If
                        End If
                    End If
                End If
            Else
                If Me.txtNoTransaksi.Text <> "" Then
                    If ObjPR.ApakahSudahAda2(Me.txtNoTransaksi.Text, True) = True Then
                        Exit Sub
                    End If
                    If ObjPR.ApakahSudahAda3(Me.txtNoTransaksi.Text, True) = True Then
                        Exit Sub
                    End If
                    If Me.ObjPR.ApakahBolehHapus(Me.TxtPrimaryPR.Text, UserCode, Me, True) = False Then
                        Me.btnFind.Focus()
                    Else
                        hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.ObjPR.Hapus(Me.TxtPrimaryPR.Text, Me.UserCode, Me) = True Then
                                Me.KosongkanField()
                                Me.NonAktifkanDetail()
                                Me.btnNew.Focus()
                                MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                            End If
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
            If Me.txtNoTransaksi.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            ''' simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPR.SimpanBaru(DgvItem, UserCode, Me, Me.vTBLItemDelivery, Me.baru) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            '''' simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPR.SimpanPerubahan(DgvItem, UserCode, Me, Me.vTBLItemDelivery, Me.baru) = True Then
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

    Private Sub TxtRequestBy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRequestBy.GotFocus, TxtRemarks.GotFocus
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

    Private Sub TxtRequestBy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRequestBy.LostFocus, TxtRemarks.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimaryPR.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '''Cari Data Keydown & Leave & Click 
    Private Sub TxtRequestBy_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRequestBy.KeyDown
        Dim ObjRefParty As New Core_Module.ClsMasterRefParty

               Try
            If e.KeyValue = 13 Then
                If Me.TxtRequestBy.Text <> "" Then
                    If ObjRefParty.ApakahSudahAda(Me.TxtRequestBy.Text, Core_Module.ClsMasterRefParty.EnumAktif.Yes) = True Then
                        Me.TxtPrimaryRefParty.Text = ObjRefParty.GetPrimary(Me.TxtRequestBy.Text)
                        Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                        Me.TxtRequestByName.Text = Me.ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                    Else
                        Me.BtnFindRefParty.PerformClick()
                    End If
                Else
                    Me.BtnFindRefParty.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub TxtRequestBy_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRequestBy.Leave
        Dim ObjRefParty As New Core_Module.ClsMasterRefParty
        Dim hj As New Core_Module.FrmMasterRefParty_Cari
        hj.Aktif = True
        hj.BackColorAll = Me.BAckColorAll

        Try
            If Me.TxtRequestBy.Text <> "" Then
                If ObjRefParty.ApakahSudahAda(Me.TxtRequestBy.Text, Core_Module.ClsMasterRefParty.EnumAktif.Yes) = True Then
                    Me.TxtPrimaryRefParty.Text = ObjRefParty.GetPrimary(Me.TxtRequestBy.Text)
                    Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                    Me.TxtRequestByName.Text = Me.ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                    Me.TxtRemarks.Focus()
                ElseIf Me.ObjRefParty.ApakahSudahAda(Me.TxtRequestBy.Text, Core_Module.ClsMasterRefParty.EnumAktif.Yes) = False And Me.TxtRequestBy.Enabled = True Then
                    If Me.TxtRequestBy.Text <> "" Then
                        hj.txtCari.Text = Me.TxtRequestBy.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryRefParty > 0 Then
                        Me.TxtPrimaryRefParty.Text = hj.PrimaryRefParty
                        Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                        Me.TxtRequestByName.Text = Me.ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                        Me.TxtRemarks.Focus()
                    Else
                        Me.TxtPrimaryRefParty.Text = "-1"
                        Me.TxtRequestBy.Clear()
                        Me.TxtRequestByName.Clear()
                        Me.TxtRequestBy.Select()
                    End If
                End If
            Else
                Me.TxtPrimaryRefParty.Text = "-1"
                Me.TxtRequestBy.Clear()
                Me.TxtRequestByName.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindRefParty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindRefParty.Click
        Dim ObjRefParty As New Core_Module.ClsMasterRefParty
        Dim hj As New Core_Module.FrmMasterRefParty_Cari
        hj.Aktif = True
        Try
            hj.BackColorAll = Me.BAckColorAll
            If Me.TxtRequestBy.Text <> "" Then
                hj.txtCari.Text = Me.TxtRequestBy.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryRefParty > 0 Then
                Me.TxtPrimaryRefParty.Text = hj.PrimaryRefParty
                Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                Me.TxtRequestByName.Text = Me.ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                Me.TxtRemarks.Focus()
            Else
                Me.TxtPrimaryRefParty.Text = "-1"
                Me.TxtRequestBy.Clear()
                Me.TxtRequestByName.Clear()
                Me.TxtRequestBy.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub


    Private Sub Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.btnFind.Focus()
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.KosongkanField()

            Me.TxtRemarks.Width = Me.SplitContainer1.Panel2.Width - Me.TxtRemarks.Left
            Me.TxtRequestByName.Width = Me.TxtRemarks.Left + Me.TxtRemarks.Width - Me.TxtRequestByName.Left

            'Me.gboDetail.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
            'Me.gboDetail.Left = (Me.Width - Me.gboDetail.Width) / 2
            Me.DtpTglTransaksi.Left = Me.SplitContainer1.Panel1.Width - 5 - Me.DtpTglTransaksi.Width
            Me.Label2.Left = Me.DtpTglTransaksi.Left - 5 - Me.Label2.Width
            Me.txtNoTransaksi.Width = Me.Label2.Left - 5 - Me.txtNoTransaksi.Left
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpTglTransaksi_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DtpTglTransaksi.ValueChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.baru = False Then
                    If Me.ObjPR.Tanggal.Month = Me.DtpTglTransaksi.Value.Month And _
                        Me.ObjPR.Tanggal.Year = Me.DtpTglTransaksi.Value.Year Then
                        Me.txtNoTransaksi.Text = Me.ObjPR.NoTransaksi
                    Else
                        Me.txtNoTransaksi.Text = Me.ObjPR.GetNoBaru(Me.DtpTglTransaksi.Value)
                    End If
                Else
                    Me.txtNoTransaksi.Text = Me.ObjPR.GetNoBaru(Me.DtpTglTransaksi.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Try
            'If baru = False Then
            '    Dim PrimaryCode As Double = Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value
            '    If e.ColumnIndex = Me.ColItem_CodeNo.Index Or e.ColumnIndex = Me.ColItem_Qty.Index Or _
            '                            e.ColumnIndex = Me.ColItem_Remarks.Index Or e.ColumnIndex = Me.ColItem_DeliveryDate.Index Or _
            '                           e.ColumnIndex = Me.ColItem_DeliveryDateControl.Index Then
            '        If ObjPR.ApakahBolehEditItem(Me.TxtPrimaryPR.Text, PrimaryCode, True) = False Then
            '            e.Cancel = True
            '            Exit Sub
            '        End If
            '    End If
            'End If

            If String.IsNullOrEmpty(Me.DgvItem(Me.ColItem_CodeNo.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColItem_CodeNo.Index Or e.ColumnIndex = Me.ColItem_Qty.Index Or _
                e.ColumnIndex = Me.ColItem_Remarks.Index Or e.ColumnIndex = Me.ColItem_DeliveryDate.Index Or _
                e.ColumnIndex = Me.ColItem_DeliveryDateControl.Index Or e.ColumnIndex = Me.ColItem_TrackUsing.Index Then
                    'jika jasa maka tidak boleh track using
                    If e.ColumnIndex = Me.ColItem_TrackUsing.Index Then
                        If Me.ObjProduct.ApakahGoods(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value) = False Then
                            e.Cancel = True
                        End If
                    End If
                    '''set header fungsi digunakan ketika edit code 
                    Me.DgvItem.Columns(Me.ColItem_CodeNo.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvItem.Columns(Me.ColItem_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvItem.Columns(Me.ColItem_Qty.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvItem.Columns(Me.ColItem_Unit.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvItem.Columns(Me.ColItem_DeliveryDate.Index).SortMode = DataGridViewColumnSortMode.NotSortable
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
                    Me.DgvItem.Columns(Me.ColItem_DeliveryDate.Index).SortMode = DataGridViewColumnSortMode.NotSortable
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

            ''' pastikan ada isinya
            With row
                If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    ''' Bila tidak ada perubahan data maka tidak perlu dilakukan cek
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


                    If e.ColumnIndex <> Me.ColItem_DeliveryDate.Index Then
                        If CStr(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToString = Me.Buffer Then
                            Exit Sub
                        End If
                    Else
                        If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Date = CDate(Me.Buffer).Date Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If e.ColumnIndex <> Me.ColItem_Remarks.Index Then
                        '''kalo bukan data remarks maka perlu dicek kebenaran data hasil edit
                        If e.ColumnIndex = Me.ColItem_Qty.Index Then
                            '''kontrol qty
                            If baru = True Then
                                Me.EndEditQty(sender, e, row)
                                Exit Sub
                            Else
                                Me.EndEditQty(sender, e, row)
                                Exit Sub
                                'Dim Code As String = Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value
                                'Dim PrimaryCode As Double = ObjProduct.GetPrimary(Code)
                                'If ObjPR.ApakahBolehEditItem(Me.TxtPrimaryPR.Text, PrimaryCode, True) = True Then
                                '    Me.EndEditQty(sender, e, row)
                                '    Exit Sub
                                'Else
                                '    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                '    Exit Sub
                                'End If
                            End If

                        ElseIf e.ColumnIndex = Me.ColItem_DeliveryDateControl.Index Then
                            '''kontrol delivery date
                            Me.EndEditDeliveryDateCont(sender, e, row)
                            Exit Sub
                        ElseIf e.ColumnIndex = Me.ColItem_CodeNo.Index Then
                            '''kontrol product code
                            If Me.baru = True Then
                                Me.EndEditProductCode(sender, e)
                                Exit Sub
                            Else
                                Dim PrimaryCode As Double = ObjProduct.GetPrimary(Buffer)
                                If ObjPR.ApakahBolehEditItem(Me.TxtPrimaryPR.Text, PrimaryCode, True) = True Then
                                    Me.EndEditProductCode(sender, e)
                                    Exit Sub
                                Else
                                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                            End If
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
            Me.DgvItem.Columns(Me.ColItem_DeliveryDate.Index).SortMode = DataGridViewColumnSortMode.Automatic
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
                    If RowDGV.Cells(Me.ColItem_PrimaryProduct.Index).Value = vPrimary Then
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
            Me.ObjCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellContentClick
        Dim ij As New FrmTransaksiPR_SetDelivery
        Dim DI As Integer = 0
        Dim rowD As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryRow
        Dim JmlDelivery As Integer = 0
        Dim DelDateD As Date = Today

        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColItem_DeliveryDate.Index Then
                If Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex).Value <> "-" And _
                Me.DgvItem.Item(Me.ColItem_DeliveryDateControl.Index, e.RowIndex).Value = Boolean.TrueString Then
                    ij.vPrimaryCodeNo = Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value
                    ij.VIdentifier = Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value
                    ij.FrmRef = Me
                    If Me.TxtRemarks.Enabled = False Then
                        ij.vReadOnly = True
                        ij.ShowDialog()
                        Exit Sub
                    End If
                    If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                        '''hapuskan semua item delivery yang identifiernya di edit
                        DI = 0
                        While DI <= Me.vTBLItemDelivery.Rows.Count - 1
                            Me.vrowItemDelivery = Me.vTBLItemDelivery.Rows(DI)
                            If Me.vrowItemDelivery.identifier = ij.VIdentifier Then
                                Me.vTBLItemDelivery.Rows.Remove(Me.vrowItemDelivery)
                            Else
                                DI += 1
                            End If
                        End While
                        '''masukkan data item delivery baru
                        ij.dgvDetail.AllowUserToAddRows = False
                        For Each rowDGV As DataGridViewRow In ij.dgvDetail.Rows
                            rowD = vTBLItemDelivery.Newtbltrans_purchase_req_itemdeliveryRow
                            With rowD
                                .delivery_date = CDate(rowDGV.Cells(ij.ColDeliveryDate.Index).Value)
                                .delivery_priority = rowDGV.Cells(ij.ColPrimaryPriority.Index).Value
                                .delivery_qty = rowDGV.Cells(ij.ColQty.Index).Value
                                .identifier = ij.VIdentifier
                                .primary = Me.GetPrimaryMax(vTBLItemDelivery) + 1
                                .primary_transaksi = Me.TxtPrimaryPR.Text
                                .remarks = rowDGV.Cells(ij.ColRemarks.Index).Value
                            End With
                            Me.vTBLItemDelivery.Rows.Add(rowD)
                            JmlDelivery += 1
                            DelDateD = CDate(rowDGV.Cells(ij.ColDeliveryDate.Index).Value)
                        Next

                        If JmlDelivery > 1 Then
                            Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex).Value = "Multiple"
                        ElseIf JmlDelivery = 1 Then
                            Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex).Value = Format(DelDateD, "dd MMM yyyy")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            rowD = Nothing
        End Try
    End Sub

    Private Sub DgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            Me.DgvItem.TabStop = True
            If Me.TxtRemarks.Enabled = True Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If

            rowDGV = Me.DgvItem.Rows(Me.DgvItem.RowCount - 1)
            With rowDGV
                .Cells(Me.ColItem_No.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColItem_PrimaryPR.Index).Value = "-1"
                .Cells(Me.ColItem_PrimaryProduct.Index).Value = "-1"
                .Cells(Me.ColItem_CodeNo.Index).Value = ""
                .Cells(Me.ColItem_Desc.Index).Value = ""
                .Cells(Me.ColItem_Qty.Index).Value = 0
                .Cells(Me.ColItem_Unit.Index).Value = ""
                .Cells(Me.ColItem_DeliveryDateControl.Index).Value = False
                .Cells(Me.ColItem_DeliveryDate.Index).Value = "-"
                .Cells(Me.ColItem_Remarks.Index).Value = ""
                .Cells(Me.ColItem_Status.Index).Value = ""
                .Cells(Me.ColItem_OB.Index).Value = 0
                .Cells(Me.ColItem_Identifier.Index).Value = Me.ObjPRItem.GetNextIdentifier(Me.DgvItem, Me.TxtPrimaryPR.Text, Me)
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            rowDGV = Nothing
        End Try

    End Sub

    Private Function GetPrimaryMax(ByVal vTBL As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryDataTable) As Double
        Dim Retval As Double = 0

        Try
            For Each Me.vrowItemDelivery In Me.vTBLItemDelivery.Rows
                If Retval < Me.vrowItemDelivery.primary Then
                    Retval = Me.vrowItemDelivery.primary
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vTBL.Dispose()
        End Try
    End Function

    '''up & down
    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.SelectionChanged
        Try
            If Me.DgvItem.ReadOnly = True Then Exit Sub
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

            If DgvItem.Rows.Count - 1 = 0 Then
                Me.DgvItem.AllowUserToAddRows = True
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

    ''' fungsi untuk product code
    Private Sub EndEditProductCode(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.DgvItem.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryProduct As Double = -1
        Dim RowIdx As Integer = e.RowIndex

        Try
            With row
                If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColItem_CodeNo.Index Then
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjProduct.ApakahSudahAda(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryProduct = Me.ObjProduct.GetPrimary(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItem(PrimaryProduct, e.RowIndex)
                            If index < 0 Then
                                Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value = PrimaryProduct
                                Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(PrimaryProduct)
                                Me.DgvItem.Item(Me.ColItem_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(PrimaryProduct)
                                Me.DgvItem.Item(Me.ColItem_Unit.Index, e.RowIndex).Value = Me.ObjProduct.GetUnitCode(PrimaryProduct)
                                Me.DgvItem.Item(Me.ColItem_TrackUsing.Index, e.RowIndex).Value = Me.ObjPR.Apakah_UseTracking2(PrimaryProduct)

                                'Me.ObjPR.SetOrderBalance(Me.TxtPrimaryPR.Text, vReadyForPurchase)
                                If ObjPR.vStatus = ClsTransaksiPR.EnumStatus.PR_Pending Then
                                    Me.DgvItem.Item(Me.ColItem_Status.Index, e.RowIndex).Value = "PR Pending"
                                ElseIf ObjPR.vStatus = ClsTransaksiPR.EnumStatus.Purchase_On_Progress Then
                                    Me.DgvItem.Item(Me.ColItem_Status.Index, e.RowIndex).Value = "Purchase On Progress"
                                Else
                                    Me.DgvItem.Item(Me.ColItem_Status.Index, e.RowIndex).Value = "Closed"
                                End If


                                If e.RowIndex > 0 Then
                                    If Me.DgvItem.Item(Me.ColItem_DeliveryDateControl.Index, e.RowIndex - 1).Value = Boolean.TrueString Then
                                        Me.DgvItem.Item(Me.ColItem_DeliveryDateControl.Index, e.RowIndex).Value = True
                                        'If IsDate(Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex - 1).Value) Then
                                        '    Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex).Value = Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex - 1).Value
                                        'Else
                                        '    Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex).Value = Me.DtpTglTransaksi.Value
                                        'End If
                                    Else
                                        Me.DgvItem.Item(Me.ColItem_DeliveryDateControl.Index, e.RowIndex).Value = False
                                    End If
                                    EndEditDeliveryDateCont(Me, e, Me.DgvItem.Rows(e.RowIndex))
                                End If

                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
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

                            Dim ij As New Core_Module.frmCariMasterProduct
                            ij.txtCari.Text = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BAckColorAll
                            ij.ShowDialog()
                            If ij.PrimaryProduct <> -1 Then

                                index = Me.GetIndexItem(ij.PrimaryProduct, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value = ij.PrimaryProduct
                                    Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                                    Me.DgvItem.Item(Me.ColItem_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                                    Me.DgvItem.Item(Me.ColItem_Unit.Index, e.RowIndex).Value = Me.ObjProduct.GetUnitCode(ij.PrimaryProduct)
                                    Me.DgvItem.Item(Me.ColItem_TrackUsing.Index, e.RowIndex).Value = Me.ObjPR.Apakah_UseTracking2(ij.PrimaryProduct)

                                    'Me.ObjPR.SetOrderBalance1(Me.TxtPrimaryPR.Text, vReadyForPurchase)
                                    If ObjPR.vStatus = ClsTransaksiPR.EnumStatus.PR_Pending Then
                                        Me.DgvItem.Item(Me.ColItem_Status.Index, e.RowIndex).Value = "PR Pending"
                                    ElseIf ObjPR.vStatus = ClsTransaksiPR.EnumStatus.Purchase_On_Progress Then
                                        Me.DgvItem.Item(Me.ColItem_Status.Index, e.RowIndex).Value = "Purchase On Progress"
                                    Else
                                        Me.DgvItem.Item(Me.ColItem_Status.Index, e.RowIndex).Value = "Closed"
                                    End If
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvItem)
                                        Exit Sub
                                    Else
                                        Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value = ij.PrimaryProduct
                                        Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                                        Me.DgvItem.Item(Me.ColItem_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                                        Me.DgvItem.Item(Me.ColItem_Unit.Index, e.RowIndex).Value = Me.ObjProduct.GetUnitCode(ij.PrimaryProduct)
                                        Me.DgvItem.Item(Me.ColItem_TrackUsing.Index, e.RowIndex).Value = Me.ObjPR.Apakah_UseTracking2(ij.PrimaryProduct)

                                        'Me.ObjPR.SetOrderBalance(Me.TxtPrimaryPR.Text, vReadyForPurchase)
                                        If ObjPR.vStatus = ClsTransaksiPR.EnumStatus.PR_Pending Then
                                            Me.DgvItem.Item(Me.ColItem_Status.Index, e.RowIndex).Value = "PR Pending"
                                        ElseIf ObjPR.vStatus = ClsTransaksiPR.EnumStatus.Purchase_On_Progress Then
                                            Me.DgvItem.Item(Me.ColItem_Status.Index, e.RowIndex).Value = "Purchase On Progress"
                                        Else
                                            Me.DgvItem.Item(Me.ColItem_Status.Index, e.RowIndex).Value = "Closed"
                                        End If
                                    End If
                                End If

                                If e.RowIndex > 0 Then
                                    If Me.DgvItem.Item(Me.ColItem_DeliveryDateControl.Index, e.RowIndex - 1).Value = Boolean.TrueString Then
                                        Me.DgvItem.Item(Me.ColItem_DeliveryDateControl.Index, e.RowIndex).Value = True
                                        'If IsDate(Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex - 1).Value) Then
                                        '    Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex).Value = Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex - 1).Value
                                        'Else
                                        '    Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex).Value = Me.DtpTglTransaksi.Value
                                        'End If
                                    Else
                                        Me.DgvItem.Item(Me.ColItem_DeliveryDateControl.Index, e.RowIndex).Value = False
                                    End If
                                    EndEditDeliveryDateCont(Me, e, Me.DgvItem.Rows(e.RowIndex))
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPreviousRowDeliveryDate(ByVal vPreviousRowIndex As Integer) As Date
        Try
            If vPreviousRowIndex > -1 Then
                If Me.DgvItem.Item(Me.ColItem_DeliveryDateControl.Index, vPreviousRowIndex).Value = True Then
                    If IsDate(Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, vPreviousRowIndex).Value) = True Then
                        Return CDate(Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, vPreviousRowIndex).Value)
                    Else
                        Me.vTBLItemDelivery.DefaultView.RowFilter = "Identifier = '" & Me.DgvItem.Item(Me.ColItem_Identifier.Index, vPreviousRowIndex).Value & "'"
                        If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                            Me.vTBLItemDelivery.DefaultView.Sort = "[delivery date] ASC"
                            Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row
                            Return Me.vrowItemDelivery.delivery_date
                        End If
                    End If
                End If
            End If
            Return Me.DtpTglTransaksi.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Me.DtpTglTransaksi.Value
        End Try
    End Function

    Private Sub EndEditDeliveryDateCont(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Dim rowD As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryRow
        Dim ij As Integer = 0
        Try
            With row
                If row.Cells(Me.ColItem_DeliveryDateControl.Index).Value = Boolean.TrueString Then
                    'jika jadi true maka delivery date jadi tanggal transaksi
                    row.Cells(Me.ColItem_DeliveryDate.Index).Value = Format(Me.GetPreviousRowDeliveryDate(e.RowIndex - 1), "dd MMM yyyy")

                    'tambahkan ke dalam item delivery date
                    rowD = Me.vTBLItemDelivery.Newtbltrans_purchase_req_itemdeliveryRow
                    With rowD
                        '.delivery_date = Me.DtpTglTransaksi.Value
                        .delivery_date = Me.GetPreviousRowDeliveryDate(e.RowIndex - 1)
                        '.delay = Me.ObjItem.GetPreparationDelay(row.Cells(Me.ColPrimaryCodeNo.Index).Value)
                        .delivery_qty = Me.DgvItem.Item(Me.ColItem_Qty.Index, e.RowIndex).Value
                        .identifier = Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value
                        .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                        .primary_transaksi = Me.TxtPrimaryPR.Text
                        .remarks = ""
                        .delivery_priority = Me.ObjPriorityLevel.GetPrimaryDefault
                    End With
                    Me.vTBLItemDelivery.Rows.Add(rowD)
                Else
                    'jika jadi false maka delivery date jadi "-"
                    row.Cells(Me.ColItem_DeliveryDate.Index).Value = "-"
                    'buang item delivery
                    While ij < Me.vTBLItemDelivery.Rows.Count
                        Me.vrowItemDelivery = Me.vTBLItemDelivery.Rows(ij)
                        If Me.vrowItemDelivery.identifier = Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value Then
                            Me.vTBLItemDelivery.Rows.Remove(Me.vrowItemDelivery)
                        Else
                            ij += 1
                        End If
                    End While
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            rowD = Nothing
        End Try
    End Sub

    Private Sub EndEditQty(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
                            ByVal row As DataGridViewRow)
        Dim JmlKurang As Decimal = 0
        Dim Unit As String
        Dim vPrimaryCodeNo As Double = -1
        Dim QtyPr As Decimal = 0
        Dim QtyCancel As Decimal = 0
        Dim QtyPickup As Decimal = 0
        Dim jmlQty As Decimal = 0

        vPrimaryCodeNo = Me.ObjProduct.GetPrimary(Me.DgvItem(Me.ColItem_CodeNo.Index, e.RowIndex).Value)
        DefPrec = Me.ObjProduct.GetAngkaDecimal(vPrimaryCodeNo)
        QtyPr = Me.DgvItem(e.ColumnIndex, e.RowIndex).Value
        QtyCancel = ObjPR.GetQtyCancel(Me.TxtPrimaryPR.Text, vPrimaryCodeNo)
        QtyPickup = ObjPR.GetQtyPickUp(Me.TxtPrimaryPR.Text, vPrimaryCodeNo)
        Unit = Me.DgvItem(Me.ColItem_Unit.Index, e.RowIndex).Value
        jmlQty = QtyCancel + QtyPickup
        Try
            With row
                If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    '''cek nilai Qty tidak lebih besar dari nilai Qtycancel + Qty Pickup
                    If baru = False Then
                        'If ObjPR.ApakahQtySudahBenar(Me, Me.TxtPrimaryPR.Text, QtyCancel, QtyPickup) = False Then
                        '    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        '    Exit Sub
                        'End If

                        If QtyPr < (QtyCancel + QtyPickup) Then
                            MsgBox("Minimum value you can edit is : " & ObjCD.DigitGrouping(jmlQty, DefPrec) & " " & Unit, MsgBoxStyle.Information, "Information")
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If

                    '''untuk nilai qty pastikan nilainya positif
                    If Me.ObjCD.CekBilanganPositif(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.OkOnly, _
                        "Information")
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    '''digit grouping
                    Me.DgvItem(e.ColumnIndex, e.RowIndex).Value = _
                    Me.ObjCD.DigitGrouping(Me.DgvItem(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                    'Me.ObjPR.SetOrderBalance(Me.TxtPrimaryPR.Text, vReadyForPurchase)

                    If ObjPR.vStatus = ClsTransaksiPR.EnumStatus.PR_Pending Then
                        Me.DgvItem.Item(Me.ColItem_Status.Index, e.RowIndex).Value = "PR Pending"
                    ElseIf ObjPR.vStatus = ClsTransaksiPR.EnumStatus.Purchase_On_Progress Then
                        Me.DgvItem.Item(Me.ColItem_Status.Index, e.RowIndex).Value = "Purchase On Progress"
                    Else
                        Me.DgvItem.Item(Me.ColItem_Status.Index, e.RowIndex).Value = "Closed"
                    End If

                    If Me.DgvItem.Item(Me.ColItem_DeliveryDateControl.Index, e.RowIndex).Value = Boolean.TrueString Then
                        If CDec(Me.Buffer) > CDec(Me.DgvItem.Item(Me.ColItem_Qty.Index, e.RowIndex).Value) Then
                            '''untuk item delivery, jika qty turun, maka hapuskan mulai dari baris paling bawah
                            JmlKurang = CDec(Me.Buffer) - CDec(Me.DgvItem.Item(Me.ColItem_Qty.Index, e.RowIndex).Value)
                            Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value + "'"
                            Me.vTBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                            While JmlKurang > 0
                                If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                                    Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row
                                    If Me.vrowItemDelivery.delivery_qty > JmlKurang Then
                                        Me.vrowItemDelivery.delivery_qty -= JmlKurang
                                        JmlKurang = 0
                                    Else
                                        JmlKurang -= Me.vrowItemDelivery.delivery_qty
                                        Me.vrowItemDelivery.delivery_qty = 0
                                    End If
                                    If Me.vrowItemDelivery.delivery_qty = 0 Then
                                        Me.vTBLItemDelivery.Rows.Remove(Me.vrowItemDelivery)
                                    End If
                                End If
                            End While

                        ElseIf CDec(Me.Buffer) < CDec(Me.DgvItem.Item(Me.ColItem_Qty.Index, e.RowIndex).Value) Then
                            '''untuk item delivery, jika qty naik, maka tambahkan sisanya pada baris paling bawah
                            JmlKurang = CDec(Me.DgvItem.Item(Me.ColItem_Qty.Index, e.RowIndex).Value) - CDec(Me.Buffer)
                            Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value + "'"
                            Me.vTBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                            If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                                Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row
                                Me.vrowItemDelivery.delivery_qty += JmlKurang
                            End If
                        End If
                    End If
                Else
                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
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

                ''cek data
                'If data.Length = 0 Then
                '    MsgBox(Convert.ToString(KoleksiNamaData.Item(i + 1)) + " cannot be an empty string.", MsgBoxStyle.Critical, "Error")
                '    KoleksiTextBox.Item(i + 1).focus()
                '    Return False
                'End If

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

    Private Sub LoadOtorisasi()
        Try
            If Me.TxtPrimaryOtorisasi.Text = "-1" Then
                Me.ChkRFP.Checked = True
                Me.ObjPR.CariDetail(Me.TxtPrimaryPR.Text, Me.DgvItem, Me.ObjCD, Me, Me.vTBLItemDelivery)
                Me.SalinKeField()
                Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadCancelOtorisasi()
        Try
            If Me.TxtPrimaryOtorisasi.Text <> "-1" Then
                Me.ChkRFP.Checked = False
                Me.ObjPR.CariDetail(Me.TxtPrimaryPR.Text, Me.DgvItem, Me.ObjCD, Me, Me.vTBLItemDelivery)
                Me.SalinKeField()
                Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkRFP_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkRFP.CheckedChanged
        Try
            If Me.ChkRFP.Checked = False Then
                vReadyForPurchase = ClsTransaksiPR.EnumReadyForPurchase.No
            Else
                vReadyForPurchase = ClsTransaksiPR.EnumReadyForPurchase.Yes
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkRFP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkRFP.Click

        Try
            vPrimaryUser = ObjUser.GetPrimary(Me.UserCode)

            If Me.ChkRFP.Checked = False Then

                Me.ChkRFP.Checked = True
            Else
                Me.ChkRFP.Checked = False
            End If

            '''kalo pr number kosong tidak bisa otorisasi
            If Me.txtNoTransaksi.Text = "" Then
                MsgBox("Please select a valid purchase request data first.", MsgBoxStyle.Information, "Information")
                Me.btnFind.Focus()
                Me.ChkRFP.Checked = False
            End If

            '''pada saat otorisasi
            If Me.txtNoTransaksi.Text <> "" And Me.TxtRemarks.Enabled = False Then
                If ObjPR.ApakahSudahDiOtorisasi(Me.TxtPrimaryPR.Text) = False Then
                    If MsgBox("Authorize?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Dim frm As New Core_Module.DialogPassword
                        frm.txtUserName.Text = Me.UserCode
                        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                            If Me.ObjPR.Authorize(Now, vPrimaryUser, Me.TxtPrimaryPR.Text, ClsTransaksiPR.EnumReadyForPurchase.Yes) = True Then
                                MsgBox("Authorization complete")
                                Me.LoadOtorisasi()
                            End If
                        Else
                            ChkRFP.Checked = False
                            Exit Sub
                        End If
                    Else
                        ChkRFP.Checked = False
                        Exit Sub
                    End If
                Else
                    '''pada saat cancel otorisasi
                    If ObjPR.ApakahSudahAdaDiDuaTransaksi(Me.txtNoTransaksi.Text) = True Then
                        MsgBox("The following purchase request data is reference by one or several purchase request cancel and " & _
                        "purchase request pickup order , you can’t cancel the authorization.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    If ObjPR.ApakahSudahAda2(Me.txtNoTransaksi.Text, True) = False Then
                        If ObjPR.ApakahSudahAda3(Me.txtNoTransaksi.Text, True) = False Then
                            '''tidak ada syarat khusus jika yang cancel atas nama "USER"
                            If Me.UserCode = "USER" Then
                                If MsgBox("Cancel authorization?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                    Dim frm As New Core_Module.DialogPassword
                                    frm.txtUserName.Text = Me.UserCode
                                    If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                                        If Me.ObjPR.CancelAuthorize(#1/1/1950#, -1, Me.TxtPrimaryPR.Text, ClsTransaksiPR.EnumReadyForPurchase.No) = True Then
                                            MsgBox("Cancellation complete")
                                            Me.LoadCancelOtorisasi()
                                        End If
                                    Else
                                        ChkRFP.Checked = True
                                        Exit Sub
                                    End If
                                Else
                                    ChkRFP.Checked = True
                                    Exit Sub
                                End If
                            Else

                                '''jika bukan atas nama "USER" 
                                '''yg boleh cancel otorisasi adalah yang pertama mengotorisasi
                                If ObjPR.ApakahUserSama(Me.TxtPrimaryPR.Text, Me.UserCode, True) = True Then
                                    If MsgBox("Cancel authorization?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                        Dim frm As New Core_Module.DialogPassword
                                        frm.txtUserName.Text = Me.UserCode
                                        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                                            If Me.ObjPR.CancelAuthorize(#1/1/1950#, -1, Me.TxtPrimaryPR.Text, ClsTransaksiPR.EnumReadyForPurchase.No) = True Then
                                                MsgBox("Cancellation complete")
                                                Me.LoadCancelOtorisasi()
                                            End If
                                        Else
                                            ChkRFP.Checked = True
                                            Exit Sub
                                        End If
                                    Else
                                        ChkRFP.Checked = True
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim i As New FrmTransaksiPR_Cetak
        Try
            '2. untuk smess
            If Me.TxtPrimaryPR.Text = -1 Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimaryPR.Text <> -1 Then
                If ObjPR.CetakPR(Me.TxtPrimaryPR.Text, i.crvReport) = True Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    i.WindowState = FormWindowState.Maximized
                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DgvItem.UserDeletingRow
        Dim RowDGV As DataGridViewRow = e.Row
        Try
            If Me.ObjPR.ApakahBolehHapusItem(Me.TxtPrimaryPR.Text, RowDGV.Cells(Me.ColItem_PrimaryProduct.Index).Value) = False Then
                e.Cancel = True
            Else
                'panggil fungsi set number
                Me.ObjCD.SetNumber(Me.DgvItem)
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

    '''Khusus [TAB] tambahan
    Dim SetTab As Boolean = False
    Private Sub btnExit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.GotFocus
        SetTab = True
    End Sub
    Private Sub btnFind_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.LostFocus
        SetTab = False
    End Sub
    Private Sub SplitContainer1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles SplitContainer1.GotFocus
        If SetTab = False Then
            btnExit.Focus()
        Else
            btnFind.Focus()
        End If
    End Sub

    Private Sub llbUpdateDeliveryDate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUpdateDeliveryDate.LinkClicked
        Dim f As New FrmTransaksiPO_SetDeliveryDate
        Dim rowD As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryRow
        Try
            If Me.DgvItem.RowCount = 1 Or Me.btnOK.Enabled = False Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            If f.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.vTBLItemDelivery.Clear()
                If f.RadioButton1.Checked = True Then
                    For Each r As DataGridViewRow In Me.DgvItem.Rows
                        If r.Cells(Me.ColItem_PrimaryProduct.Index).Value <> "-1" Then
                            r.Cells(Me.ColItem_DeliveryDateControl.Index).Value = True
                            r.Cells(Me.ColItem_DeliveryDate.Index).Value = Format(f.DateTimePicker1.Value.Date, "dd MMM yyyy")
                            rowD = Me.vTBLItemDelivery.Newtbltrans_purchase_req_itemdeliveryRow
                            With rowD
                                .delivery_date = f.DateTimePicker1.Value.Date
                                .delivery_priority = Me.ObjPriorityLevel.GetPrimaryDefault
                                .delivery_qty = r.Cells(Me.ColItem_Qty.Index).Value
                                .identifier = r.Cells(Me.ColItem_Identifier.Index).Value
                                .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                                .primary_transaksi = Me.TxtPrimaryPR.Text
                                .remarks = ""
                            End With
                            Me.vTBLItemDelivery.Rows.Add(rowD)
                        End If
                    Next
                Else
                    For Each r As DataGridViewRow In Me.DgvItem.Rows
                        r.Cells(Me.ColItem_DeliveryDateControl.Index).Value = False
                        r.Cells(Me.ColItem_DeliveryDate.Index).Value = "-"
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
End Class