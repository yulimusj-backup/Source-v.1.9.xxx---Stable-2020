Option Explicit On
Option Strict Off

Public Class FrmMasterSubBrand
    Public ProductGroupLevel As Integer = 1
    Public UserCode As String = ""
    Public BackcolorAll As Color = Color.LightGoldenrodYellow

    Dim baru As Boolean
    Dim objProductGroup As ClsMasterProductGroup
    Dim ObjCD As New ClsCekData

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimaryGroup.BackColor, False)
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
            Me.txtPrimaryGroup.Text = "-1"
            Me.txtPrimaryParentGroup.Text = "-1"
            Me.txtGroupLevel.Text = Me.ProductGroupLevel
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
            KoleksiNamaData.Add("sub brand code")
            KoleksiTextBox.Add(Me.txtGroupCode)
            KoleksiNamaData.Add("description")
            KoleksiTextBox.Add(Me.txtDescription)
            KoleksiNamaData.Add("brand code")
            KoleksiTextBox.Add(Me.txtParentGroupCode)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data salah bila ada duplikasi kode
            If Me.baru = True Then
                If Me.objProductGroup.ApakahSudahAda(Me.txtGroupCode.Text, Me.ProductGroupLevel) = True Then
                    MsgBox("Duplication in product sub brand code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtGroupCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.objProductGroup.GetGroupCode(Me.txtPrimaryGroup.Text)
                If ProductGroupLama <> Me.txtGroupCode.Text Then
                    If Me.objProductGroup.ApakahSudahAda(Me.txtGroupCode.Text, Me.ProductGroupLevel) = True Then
                        MsgBox("Duplication in product sub brand code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.txtGroupCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objProductGroup.GetTimeStamp(Me.txtPrimaryGroup.Text)
                If Me.objProductGroup.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objProductGroup.CariDetail(Me.objProductGroup.Primary)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'cek data yang salah
            If Me.ProductGroupLevel = 3 Then
                'ga diitung karena levelnya paling tinggi
            ElseIf Me.ProductGroupLevel = 2 Then
                If Me.objProductGroup.ApakahSudahAda(Me.txtParentGroupCode.Text, Me.ProductGroupLevel + 1) = False Then
                    MsgBox("Wrong product group code, we cannot find it in our database.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf Me.ProductGroupLevel = 1 Then
                If Me.objProductGroup.ApakahSudahAda(Me.txtParentGroupCode.Text, Me.ProductGroupLevel + 1) = False Then
                    MsgBox("Wrong product brand code, we cannot find it in our database.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
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

    Private Sub SalinKeObjek()
        Try
            With Me.objProductGroup
                .ProductGroupCode = Me.txtGroupCode.Text
                .Description = Me.txtDescription.Text
                .ProductGroupLevel = Me.ProductGroupLevel
                .PrimaryParentProductGroup = Me.txtPrimaryParentGroup.Text
                .Remarks = Me.txtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objProductGroup
                Me.txtGroupCode.Text = .ProductGroupCode
                Me.txtDescription.Text = .Description
                Me.txtPrimaryParentGroup.Text = .PrimaryParentProductGroup
                Me.txtPrimaryGroup.Text = .Primary
                Me.txtRemarks.Text = .Remarks
                If .PrimaryParentProductGroup <> -1 Then
                    Me.txtParentGroupCode.Text = Me.objProductGroup.GetGroupCode(Me.txtPrimaryParentGroup.Text)
                End If
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
            Me.objProductGroup = New ClsMasterProductGroup
            Me.baru = True
            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'cek data benar
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objProductGroup.SimpanBaru(Me.UserCode) Then
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
                        If Me.objProductGroup.SimpanPerubahan(Me.UserCode) = True Then
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
            If Not Me.txtGroupCode.Text = "" And Me.txtDescription.Text <> "" Then
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
            Me.txtGroupCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariMasterProductGroup
        Try
            hj.ProductGroupLevel = Me.ProductGroupLevel
            hj.ShowDialog()
            If Not hj.PrimaryProductGroup = -1 Then
                If Me.objProductGroup.CariDetail(hj.PrimaryProductGroup) = True Then
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
        If Me.txtGroupCode.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.txtGroupCode.Text <> "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.txtGroupCode.Focus()
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtGroupCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtGroupCode.Text <> "" Then
                If Me.objProductGroup.ApakahBolehHapus(Me.txtPrimaryGroup.Text) = False Then
                    MsgBox("You cannot delete this data, because there are category using this data.", _
                    MsgBoxStyle.Information, "Information")
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objProductGroup.Hapus(Me.txtPrimaryGroup.Text) = True Then
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

    Private Sub txtParentGroupCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtParentGroupCode.KeyDown
        Dim i As New FrmCariMasterProductGroup
        Dim Obj As New ClsMasterProductGroup

        Try
            i.BackColorAll = Me.BackcolorAll
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.txtParentGroupCode.Text, Me.ProductGroupLevel + 1) = True Then
                    Me.txtPrimaryParentGroup.Text = Obj.GetPrimary(Me.txtParentGroupCode.Text, Me.ProductGroupLevel + 1)
                    Me.txtRemarks.Focus()
                Else
                    If Me.txtParentGroupCode.Text <> "" Then
                        i.txtCari.Text = Me.txtParentGroupCode.Text
                    End If

                    i.ProductGroupLevel = Me.ProductGroupLevel + 1
                    i.ShowDialog()

                    If i.PrimaryProductGroup > 0 Then
                        Me.txtPrimaryParentGroup.Text = i.PrimaryProductGroup
                        Me.txtParentGroupCode.Text = Obj.GetGroupCode(Me.txtPrimaryParentGroup.Text)
                        Me.txtRemarks.Focus()
                    Else
                        Me.txtParentGroupCode.Clear()
                        Me.txtPrimaryParentGroup.Text = "-1"
                        Me.txtParentGroupCode.Focus()
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

    Private Sub txtParentGroupCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtParentGroupCode.Leave
        Dim i As New FrmCariMasterProductGroup
        Dim obj As New ClsMasterProductGroup

        Try
            i.BackColorAll = Me.BackcolorAll
            If Me.txtParentGroupCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.txtParentGroupCode.Text, Me.ProductGroupLevel + 1) = True Then
                    Me.txtRemarks.Focus()
                    Me.txtPrimaryParentGroup.Text = obj.GetPrimary(Me.txtParentGroupCode.Text, Me.ProductGroupLevel + 1)
                Else
                    If Me.txtParentGroupCode.Text <> "" Then
                        i.txtCari.Text = Me.txtParentGroupCode.Text
                    End If

                    i.ProductGroupLevel = Me.ProductGroupLevel + 1
                    i.ShowDialog()
                    If i.PrimaryProductGroup > 0 Then
                        Me.txtPrimaryParentGroup.Text = i.PrimaryProductGroup
                        Me.txtParentGroupCode.Text = obj.GetGroupCode(Me.txtPrimaryParentGroup.Text)
                        Me.txtRemarks.Focus()
                    Else
                        Me.txtParentGroupCode.Clear()
                        Me.txtPrimaryParentGroup.Text = "-1"
                        Me.txtParentGroupCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryParentGroup.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindParentGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindParentGroup.Click
        Dim i As New FrmCariMasterProductGroup
        Dim obj As New ClsMasterProductGroup

        Try
            i.BackColorAll = Me.BackcolorAll
            i.cboCriteria.SelectedIndex = 0
            i.ProductGroupLevel = Me.ProductGroupLevel + 1
            If Me.txtParentGroupCode.Text <> "" Then
                i.txtCari.Text = Me.txtParentGroupCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryProductGroup > 0 Then
                Me.txtPrimaryParentGroup.Text = i.PrimaryProductGroup
                Me.txtParentGroupCode.Text = obj.GetGroupCode(Me.txtPrimaryParentGroup.Text)
                Me.txtRemarks.Focus()
            Else
                Me.txtParentGroupCode.Clear()
                Me.txtParentGroupCode.Focus()
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
            If Me.txtGroupCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objProductGroup.SimpanBaru(Me.UserCode) = True Then
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
                            If Me.objProductGroup.SimpanPerubahan(Me.UserCode) = True Then
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
    Private Sub txtGroupCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGroupCode.GotFocus
        txtGroupCode.BackColor = BackcolorAll
    End Sub

    Private Sub txtGroupCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGroupCode.LostFocus
        txtGroupCode.BackColor = Me.txtPrimaryGroup.BackColor
    End Sub

    Private Sub txtDescription_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.GotFocus
        txtDescription.BackColor = BackcolorAll
    End Sub

    Private Sub txtDescription_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.LostFocus
        txtDescription.BackColor = Me.txtPrimaryGroup.BackColor
    End Sub

    Private Sub txtParentGroupCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtParentGroupCode.GotFocus
        txtParentGroupCode.BackColor = BackcolorAll
    End Sub

    Private Sub txtParentGroupCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtParentGroupCode.LostFocus
        txtParentGroupCode.BackColor = Me.txtPrimaryGroup.BackColor
    End Sub

    Private Sub txtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.GotFocus
        Try
            If Me.txtRemarks.ReadOnly = False And Me.txtRemarks.Enabled = True Then
                Me.txtRemarks.BackColor = BackcolorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.LostFocus
        Try
            If Me.txtRemarks.ReadOnly = False And Me.txtRemarks.Enabled = True Then
                Me.txtRemarks.BackColor = Me.txtPrimaryGroup.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class