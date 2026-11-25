Option Explicit On
Option Strict Off

Public Class frmMasterCategory
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow

    Dim baru As Boolean
    Dim objCategory As ClsMasterCategory
    Dim ObjCD As New ClsCekData

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimaryCategory.BackColor, False)
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
            Me.txtPrimaryCategory.Text = "-1"
            Me.txtPrimarySubBrand.Text = "-1"
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
            KoleksiNamaData.Add("category code")
            KoleksiTextBox.Add(Me.txtCategoryCode)
            KoleksiNamaData.Add("description")
            KoleksiTextBox.Add(Me.txtDescription)
            KoleksiNamaData.Add("sub brand code")
            KoleksiTextBox.Add(Me.txtSubBrandCode)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.objCategory.ApakahSudahAda(Me.txtCategoryCode.Text) = True Then
                    MsgBox("Data duplication, please change category code.", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtCategoryCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.objCategory.GetCategoryCode(Me.txtPrimaryCategory.Text)
                If ProductGroupLama <> Me.txtCategoryCode.Text Then
                    If Me.objCategory.ApakahSudahAda(Me.txtCategoryCode.Text) = True Then
                        MsgBox("Data duplication, please change category code.", _
                        MsgBoxStyle.Information, "Information")
                        Me.txtCategoryCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objCategory.GetTimeStamp(Me.txtPrimaryCategory.Text)
                If Me.objCategory.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objCategory.CariDetail(Me.objCategory.Primary)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

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
            With Me.objCategory
                .CategoryCode = Me.txtCategoryCode.Text
                .Description = Me.txtDescription.Text
                .PrimaryParent = Me.txtPrimarySubBrand.Text
                .Remarks = Me.txtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim i As New ClsMasterProductGroup

        Try
            With Me.objCategory
                Me.txtCategoryCode.Text = .CategoryCode
                Me.txtDescription.Text = .Description
                Me.txtPrimarySubBrand.Text = .PrimaryParent
                Me.txtPrimaryCategory.Text = .Primary
                If .PrimaryParent <> -1 Then
                    Me.txtSubBrandCode.Text = i.GetGroupCode(Me.txtPrimarySubBrand.Text)
                End If
                Me.txtRemarks.Text = .Remarks
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i = Nothing
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.objCategory = New ClsMasterCategory
            Me.baru = True
            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'cek data kosong
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objCategory.SimpanBaru(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'khusus untuk data lama, jika ada perubahan data kode category, maka akan berubah juga ke
                    'dalam data kode product
                    Dim KodeLama As String = Me.objCategory.GetCategoryCode(Me.txtPrimaryCategory.Text)
                    Dim samaKode, samaDesc As Boolean
                    If Me.txtCategoryCode.Text <> KodeLama Then
                        hasil = MsgBox("You changed category code, this will affect all product code refering to this category. Save?", _
                        MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        samaKode = False
                    Else
                        samaKode = True
                    End If
                    'cek perubahan data desc
                    If samaKode = True Then
                        Dim DescLama As String = Me.objCategory.GetDesc(Me.txtPrimaryCategory.Text)
                        If Me.txtDescription.Text <> DescLama Then
                            hasil = MsgBox("You changed category description, this will affect all product code refering to this category. Save?", _
                            MsgBoxStyle.YesNo, "Confirm")
                            If hasil = MsgBoxResult.No Then
                                Exit Sub
                            End If
                            samaDesc = False
                        Else
                            samaDesc = True
                        End If
                    Else
                        Dim DescLama As String = Me.objCategory.GetDesc(Me.txtPrimaryCategory.Text)
                        If Me.txtDescription.Text <> DescLama Then
                            samaDesc = False
                        Else
                            samaDesc = True
                        End If
                    End If

                    If samaKode = True And samaDesc = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    End If

                    'konfirmasi disimpan atau tidak
                    If hasil = MsgBoxResult.Yes Then
                        If samaKode = False Or samaDesc = False Then
                            Me.tspBar1.Visible = True
                            Me.tslModifiedBy.Visible = False
                            Me.tslModifiedDate.Visible = False
                        End If

                        If Me.objCategory.SimpanPerubahan(Me.UserCode, samaKode, samaDesc, Me.tspBar1) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()

                            If samaKode = False Or samaDesc = False Then
                                Me.tspBar1.Visible = False
                                Me.tslModifiedBy.Visible = True
                                Me.tslModifiedDate.Visible = True
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
            If Not Me.txtCategoryCode.Text = "" And Me.txtDescription.Text <> "" Then
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
            Me.txtCategoryCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New frmCariMasterCategory
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryCategory = -1 Then
                If Me.objCategory.CariDetail(hj.PrimaryCategory) = True Then
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
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If Me.txtCategoryCode.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.txtCategoryCode.Text <> "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.txtCategoryCode.Focus()
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtCategoryCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtCategoryCode.Text <> "" Then
                If Me.objCategory.ApakahBolehHapus(Me.txtPrimaryCategory.Text) = False Then
                    MsgBox("You cannot delete this data, there are products using this data.", _
                                        MsgBoxStyle.Information, "Information")
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objCategory.Hapus(Me.txtPrimaryCategory.Text) Then
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

    Private Sub txtSubBrandCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSubBrandCode.KeyDown
        Dim i As New FrmCariMasterProductGroup
        Dim Obj As New ClsMasterProductGroup

        Try
            i.BackColorAll = i.BackColorAll
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.txtSubBrandCode.Text, 1) = True Then
                    Me.txtPrimarySubBrand.Text = Obj.GetPrimary(Me.txtSubBrandCode.Text, 1)
                    Me.txtRemarks.Focus()
                Else
                    If Me.txtSubBrandCode.Text <> "" Then
                        i.txtCari.Text = Me.txtSubBrandCode.Text
                    End If

                    i.ProductGroupLevel = 1
                    i.ShowDialog()
                    If i.PrimaryProductGroup > 0 Then
                        Me.txtPrimarySubBrand.Text = i.PrimaryProductGroup
                        Me.txtSubBrandCode.Text = Obj.GetGroupCode(Me.txtPrimarySubBrand.Text)
                        Me.txtRemarks.Focus()
                    Else
                        Me.txtSubBrandCode.Clear()
                        Me.txtPrimarySubBrand.Text = "-1"
                        Me.txtSubBrandCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub txtSubBrandCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubBrandCode.Leave
        Dim i As New FrmCariMasterProductGroup
        Dim obj As New ClsMasterProductGroup

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.txtSubBrandCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.txtSubBrandCode.Text, 1) = True Then
                    Me.txtRemarks.Focus()
                    Me.txtPrimarySubBrand.Text = obj.GetPrimary(Me.txtSubBrandCode.Text, 1)
                Else
                    If Me.txtSubBrandCode.Text <> "" Then
                        i.txtCari.Text = Me.txtSubBrandCode.Text
                    End If

                    i.ProductGroupLevel = 1
                    i.ShowDialog()
                    If i.PrimaryProductGroup > 0 Then
                        Me.txtPrimarySubBrand.Text = i.PrimaryProductGroup
                        Me.txtSubBrandCode.Text = obj.GetGroupCode(Me.txtPrimarySubBrand.Text)
                        Me.txtRemarks.Focus()
                    Else
                        Me.txtSubBrandCode.Clear()
                        Me.txtPrimarySubBrand.Text = "-1"
                        Me.txtSubBrandCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimarySubBrand.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtCategoryCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objCategory.SimpanBaru(Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'khusus untuk data lama, jika ada perubahan data kode category, maka akan berubah juga ke
                            'dalam data kode product
                            Dim KodeLama As String = Me.objCategory.GetCategoryCode(Me.txtPrimaryCategory.Text)
                            Dim samaKode, samaDesc As Boolean
                            If Me.txtCategoryCode.Text <> KodeLama Then
                                samaKode = False
                            Else
                                samaKode = True
                            End If
                            'cek perubahan data desc
                            If samaKode = True Then
                                Dim DescLama As String = Me.objCategory.GetDesc(Me.txtPrimaryCategory.Text)
                                If Me.txtDescription.Text <> DescLama Then
                                    samaDesc = False
                                Else
                                    samaDesc = True
                                End If
                            End If

                            If samaKode = False Or samaDesc = False Then
                                Me.tspBar1.Visible = True
                                Me.tslModifiedBy.Visible = False
                                Me.tslModifiedDate.Visible = False
                            End If

                            If Me.objCategory.SimpanPerubahan(Me.UserCode, samaKode, samaDesc, Me.tspBar1) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()

                                If samaKode = False Or samaDesc = False Then
                                    Me.tspBar1.Visible = False
                                    Me.tslModifiedBy.Visible = True
                                    Me.tslModifiedDate.Visible = True
                                End If
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

    Private Sub btnFindSubBrand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSubBrand.Click
        Dim i As New FrmCariMasterProductGroup
        Dim obj As New ClsMasterProductGroup

        Try
            i.ProductGroupLevel = 1
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.txtSubBrandCode.Text <> "" Then
                i.txtCari.Text = Me.txtSubBrandCode.Text
            End If

            i.ProductGroupLevel = 1
            i.ShowDialog()
            If i.PrimaryProductGroup > 0 Then
                Me.txtPrimarySubBrand.Text = i.PrimaryProductGroup
                Me.txtSubBrandCode.Text = obj.GetGroupCode(Me.txtPrimarySubBrand.Text)
                Me.txtRemarks.Focus()
            Else
                Me.txtSubBrandCode.Clear()
                Me.txtSubBrandCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub txtCategoryCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCategoryCode.GotFocus
        txtCategoryCode.BackColor = BackColorAll
    End Sub

    Private Sub txtCategoryCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCategoryCode.LostFocus
        txtCategoryCode.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub txtDescription_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.GotFocus
        txtDescription.BackColor = BackColorAll
    End Sub

    Private Sub txtDescription_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.LostFocus
        txtDescription.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub txtSubBrandCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubBrandCode.GotFocus
        txtSubBrandCode.BackColor = BackColorAll
    End Sub

    Private Sub txtSubBrandCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubBrandCode.LostFocus
        txtSubBrandCode.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub txtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.Enter
        Try
            If Me.txtRemarks.ReadOnly = False And Me.txtRemarks.Enabled = True Then
                Me.txtRemarks.BackColor = BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.Leave
        Try
            If Me.txtRemarks.ReadOnly = False And Me.txtRemarks.Enabled = True Then
                Me.txtRemarks.BackColor = Me.txtPrimarySubBrand.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSubBrandCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSubBrandCode.TextChanged

    End Sub
End Class