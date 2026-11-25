Option Explicit On
Option Strict Off

Public Class FrmMasterStorageArea
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow

    Dim baru As Boolean
    Dim objStorageArea As clsMasterStorageArea
    Dim ObjCD As New ClsCekData

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimaryStorageArea.BackColor, False)
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
            Me.txtPrimaryStorageArea.Text = "-1"
            Me.txtPrimaryWarehouse.Text = "-1"
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
            KoleksiNamaData.Add("storage area code")
            KoleksiTextBox.Add(Me.txtStorageAreaCode)
            KoleksiNamaData.Add("description")
            KoleksiTextBox.Add(Me.txtDescription)
            KoleksiNamaData.Add("warehouse code")
            KoleksiTextBox.Add(Me.txtWarehouseCode)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.objStorageArea.ApakahSudahAda(Me.txtStorageAreaCode.Text) = True Then
                    MsgBox("Duplication in storage area code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtStorageAreaCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.objStorageArea.GetStorageAreaCode(Me.txtPrimaryStorageArea.Text)
                If ProductGroupLama <> Me.txtStorageAreaCode.Text Then
                    If Me.objStorageArea.ApakahSudahAda(Me.txtStorageAreaCode.Text) = True Then
                        MsgBox("Duplication in storage area code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.txtStorageAreaCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objStorageArea.GetTimeStamp(Me.txtPrimaryStorageArea.Text)
                If Me.objStorageArea.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objStorageArea.CariDetail(Me.objStorageArea.Primary)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'data salah jika kode storage area ="-"
            If Me.txtStorageAreaCode.Text = "-" Then
                MsgBox("Please enter other code besides '-'.", MsgBoxStyle.Information, "Information")
                Return False
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
            With Me.objStorageArea
                .StorageAreaCode = Me.txtStorageAreaCode.Text
                .Description = Me.txtDescription.Text
                .PrimaryParent = Me.txtPrimaryWarehouse.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim i As New clsMasterWarehouse

        Try
            With Me.objStorageArea
                Me.txtStorageAreaCode.Text = .StorageAreaCode
                Me.txtDescription.Text = .Description
                Me.txtPrimaryWarehouse.Text = .PrimaryParent
                Me.txtPrimaryStorageArea.Text = .Primary
                If .PrimaryParent <> -1 Then
                    Me.txtWarehouseCode.Text = i.GetWarehouseCode(Me.txtPrimaryWarehouse.Text)
                End If
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
            Me.objStorageArea = New clsMasterStorageArea
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
                        If Me.objStorageArea.SimpanBaru(Me.UserCode) Then
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
                        If Me.objStorageArea.SimpanPerubahan(Me.UserCode) Then
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
            If Not Me.txtStorageAreaCode.Text = "" And Me.txtDescription.Text <> "" Then
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
            Me.txtStorageAreaCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariMasterStorageArea
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryStorageArea = -1 Then
                If Me.objStorageArea.CariDetail(hj.PrimaryStorageArea) = True Then
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
        If Me.txtStorageAreaCode.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.txtStorageAreaCode.Text <> "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.txtStorageAreaCode.Focus()
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtStorageAreaCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtStorageAreaCode.Text <> "" Then
                If Me.objStorageArea.ApakahBolehHapus(Me.txtPrimaryStorageArea.Text) = False Then
                    MsgBox("You cannot delete this data, there are product(s) stored in this area.", _
                                        MsgBoxStyle.Information, "Information")
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objStorageArea.Hapus(Me.txtPrimaryStorageArea.Text) Then
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

    Private Sub txtWarehouseCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtWarehouseCode.KeyDown
        Dim i As New frmCariMasterWarehouse
        Dim Obj As New clsMasterWarehouse

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.txtWarehouseCode.Text) = True Then
                    Me.txtPrimaryWarehouse.Text = Obj.GetPrimary(Me.txtWarehouseCode.Text)
                    Me.btnOK.Focus()
                Else
                    If Me.txtWarehouseCode.Text <> "" Then
                        i.txtCari.Text = Me.txtWarehouseCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryWarehouse > 0 Then
                        Me.txtPrimaryWarehouse.Text = i.PrimaryWarehouse
                        Me.txtWarehouseCode.Text = Obj.GetWarehouseCode(Me.txtPrimaryWarehouse.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.txtWarehouseCode.Clear()
                        Me.txtPrimaryWarehouse.Text = "-1"
                        Me.txtWarehouseCode.Focus()
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

    Private Sub txtWarehouseCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWarehouseCode.Leave
        Dim i As New frmCariMasterWarehouse
        Dim obj As New clsMasterWarehouse

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.txtWarehouseCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.txtWarehouseCode.Text) = True Then
                    Me.btnOK.Focus()
                    Me.txtPrimaryWarehouse.Text = obj.GetPrimary(Me.txtWarehouseCode.Text)
                Else
                    If Me.txtWarehouseCode.Text <> "" Then
                        i.txtCari.Text = Me.txtWarehouseCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryWarehouse > 0 Then
                        Me.txtPrimaryWarehouse.Text = i.PrimaryWarehouse
                        Me.txtWarehouseCode.Text = obj.GetWarehouseCode(Me.txtPrimaryWarehouse.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.txtWarehouseCode.Clear()
                        Me.txtPrimaryWarehouse.Text = "-1"
                        Me.txtWarehouseCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryWarehouse.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindWarehouse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindWarehouse.Click
        Dim i As New frmCariMasterWarehouse
        Dim obj As New clsMasterWarehouse

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.txtWarehouseCode.Text <> "" Then
                i.txtCari.Text = Me.txtWarehouseCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryWarehouse > 0 Then
                Me.txtPrimaryWarehouse.Text = i.PrimaryWarehouse
                Me.txtWarehouseCode.Text = obj.GetWarehouseCode(Me.txtPrimaryWarehouse.Text)
                Me.btnOK.Focus()
            Else
                Me.txtWarehouseCode.Clear()
                Me.txtWarehouseCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub txtStorageAreaCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStorageAreaCode.GotFocus
        txtStorageAreaCode.BackColor = BackColorAll
    End Sub

    Private Sub txtStorageAreaCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStorageAreaCode.LostFocus
        txtStorageAreaCode.BackColor = Me.txtPrimaryStorageArea.BackColor
    End Sub

    Private Sub txtDescription_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.GotFocus
        txtDescription.BackColor = BackColorAll
    End Sub

    Private Sub txtDescription_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.LostFocus
        txtDescription.BackColor = Me.txtPrimaryStorageArea.BackColor
    End Sub

    Private Sub txtWarehouseCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWarehouseCode.GotFocus
        txtWarehouseCode.BackColor = BackColorAll
    End Sub

    Private Sub txtWarehouseCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWarehouseCode.LostFocus
        txtWarehouseCode.BackColor = Me.txtPrimaryStorageArea.BackColor
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtStorageAreaCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objStorageArea.SimpanBaru(Me.UserCode) = True Then
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
                            If Me.objStorageArea.SimpanPerubahan(Me.UserCode) = True Then
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
End Class