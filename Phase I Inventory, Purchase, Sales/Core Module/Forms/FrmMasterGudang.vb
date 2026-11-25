Option Explicit On
Option Strict Off

Public Class FrmMasterGudang
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim objWarehouse As New clsMasterWarehouse
    Dim objStorageArea As New clsMasterStorageArea
    Dim ObjCD As New ClsCekData
    Dim ObjLicense As New ClsLicense

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryWarehouse.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.chkUseSpecificTransitLoc.Enabled = True
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.SetUseTransit_InForm()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.chkUseSpecificTransitLoc.Enabled = False
            Me.CancelButton = Me.btnExit
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            Me.TxtPrimaryWarehouse.Text = "-1"
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
            KoleksiNamaData.Add("warehouse code")
            KoleksiTextBox.Add(Me.TxtWarehouseCode)
            KoleksiNamaData.Add("description")
            KoleksiTextBox.Add(Me.TxtDescription)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.objWarehouse.ApakahSudahAda(Me.TxtWarehouseCode.Text) = True Then
                    MsgBox("Duplication in warehouse code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtWarehouseCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.objWarehouse.GetWarehouseCode(Me.TxtPrimaryWarehouse.Text)
                If ProductGroupLama <> Me.TxtWarehouseCode.Text Then
                    If Me.objWarehouse.ApakahSudahAda(Me.TxtWarehouseCode.Text) = True Then
                        MsgBox("Duplication in warehouse code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtWarehouseCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objWarehouse.GetTimeStamp(Me.TxtPrimaryWarehouse.Text)
                If Me.objWarehouse.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objWarehouse.CariDetail(Me.objWarehouse.Primary)
                        Me.SalinKeField()
                    End If
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
            With Me.objWarehouse
                .WarehouseCode = Me.TxtWarehouseCode.Text
                .Description = Me.TxtDescription.Text
                If Me.chkUseSpecificTransitLoc.Checked = False Then
                    .UseTransit = clsMasterWarehouse.EnumUseTransitLocation.NotUseTransitLoc
                Else
                    .UseTransit = clsMasterWarehouse.EnumUseTransitLocation.UseTransitLoc
                End If
                If String.IsNullOrEmpty(Me.TxtPrimaryTransitLoc.Text) = True Then
                    Me.TxtPrimaryTransitLoc.Text = -1
                End If
                .PrimaryTransitLocation = Me.TxtPrimaryTransitLoc.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objWarehouse
                Me.TxtWarehouseCode.Text = .WarehouseCode
                Me.TxtDescription.Text = .Description
                Me.TxtPrimaryWarehouse.Text = .Primary
                If .UseTransit = clsMasterWarehouse.EnumUseTransitLocation.NotUseTransitLoc Then
                    Me.chkUseSpecificTransitLoc.Checked = False
                Else
                    Me.chkUseSpecificTransitLoc.Checked = True
                End If
                Me.TxtPrimaryTransitLoc.Text = .PrimaryTransitLocation
                Me.TxtTransitLocCode.Text = Me.objStorageArea.GetStorageAreaCode(.PrimaryTransitLocation)
                Me.TxtTransitLocDesc.Text = Me.objStorageArea.GetDescription(.PrimaryTransitLocation)
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
            Me.objWarehouse = New clsMasterWarehouse
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
                        If Me.objWarehouse.SimpanBaru(Me.UserCode) Then
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
                        If Me.objWarehouse.SimpanPerubahan(Me.UserCode) Then
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
            If Not Me.TxtWarehouseCode.Text = "" And Me.TxtDescription.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.objWarehouse.CariDetail(Me.TxtPrimaryWarehouse.Text)
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
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtWarehouseCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New frmCariMasterWarehouse
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryWarehouse = -1 Then
                If Me.objWarehouse.CariDetail(hj.PrimaryWarehouse) = True Then
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
        If Me.TxtWarehouseCode.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.TxtWarehouseCode.Text <> "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.TxtWarehouseCode.Focus()
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtWarehouseCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtWarehouseCode.Text <> "" Then
                If Me.objWarehouse.ApakahBolehHapus(Me.TxtPrimaryWarehouse.Text) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objWarehouse.Hapus(Me.TxtPrimaryWarehouse.Text, Me.UserCode) = True Then
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
            If Me.TxtWarehouseCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objWarehouse.SimpanBaru(Me.UserCode) = True Then
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
                            If Me.objWarehouse.SimpanPerubahan(Me.UserCode) = True Then
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

    Private Sub TxtWarehouseCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWarehouseCode.GotFocus
        TxtWarehouseCode.BackColor = BackColorAll
    End Sub

    Private Sub TxtWarehouseCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWarehouseCode.LostFocus
        TxtWarehouseCode.BackColor = Me.TxtPrimaryWarehouse.BackColor
    End Sub

    Private Sub TxtDescription_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDescription.GotFocus
        TxtDescription.BackColor = BackColorAll
    End Sub

    Private Sub TxtDescription_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDescription.LostFocus
        TxtDescription.BackColor = Me.TxtPrimaryWarehouse.BackColor
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub chkUseSpecificTransitLoc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseSpecificTransitLoc.CheckedChanged
        Try
            Me.SetUseTransit_InForm()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTransitLocCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTransitLocCode.GotFocus
        TxtTransitLocCode.BackColor = BackColorAll
    End Sub

    Private Sub TxtTransitLocCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTransitLocCode.KeyDown
        Dim ObjS As New clsMasterStorageArea
        Try
            If e.KeyValue = 13 Then
                If ObjS.ApakahSudahAda(Me.TxtTransitLocCode.Text, clsMasterStorageArea.enumTransit.Transit) = True Then
                    Me.TxtPrimaryTransitLoc.Text = ObjS.GetPrimary(Me.TxtTransitLocCode.Text)
                    Me.TxtTransitLocDesc.Text = ObjS.GetDescription(Me.TxtPrimaryTransitLoc.Text)
                    Me.btnOK.Focus()
                Else
                    Me.btnFindTransitLoc.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTransitLocCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTransitLocCode.Leave
        Dim ObjS As New clsMasterStorageArea
        Try
            If Me.TxtTransitLocCode.Text <> "" Then
                If ObjS.ApakahSudahAda(Me.TxtTransitLocCode.Text, clsMasterStorageArea.enumTransit.Transit) = True Then
                    Me.TxtPrimaryTransitLoc.Text = ObjS.GetPrimary(Me.TxtTransitLocCode.Text)
                    Me.TxtTransitLocDesc.Text = ObjS.GetDescription(Me.TxtPrimaryTransitLoc.Text)
                Else
                    Me.btnFindTransitLoc.PerformClick()
                End If
            Else
                Me.TxtPrimaryTransitLoc.Text = -1
                Me.TxtTransitLocCode.Text = ""
                Me.TxtTransitLocDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindTransitLoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindTransitLoc.Click
        Dim hj As New FrmCariMasterStorageArea

        Try
            hj.BackColorAll = Me.BackColorAll
            hj.PencarianTransit = clsMasterStorageArea.enumTransit.Transit
            hj.txtCari.Text = Me.TxtTransitLocCode.Text
            hj.ShowDialog()
            If Not hj.PrimaryStorageArea = -1 Then
                Me.TxtPrimaryTransitLoc.Text = hj.PrimaryStorageArea
                Me.TxtTransitLocCode.Text = objStorageArea.GetStorageAreaCode(hj.PrimaryStorageArea)
                Me.TxtTransitLocDesc.Text = objStorageArea.GetDescription(hj.PrimaryStorageArea)
                Me.btnOK.Focus()
            Else
                Me.TxtPrimaryTransitLoc.Text = -1
                Me.TxtTransitLocCode.Text = ""
                Me.TxtTransitLocDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub SetUseTransit_InForm()
        Try
            If Me.btnNew.Enabled = False Then
                If Me.chkUseSpecificTransitLoc.Checked = True Then
                    Me.TxtTransitLocCode.Enabled = True
                    Me.TxtTransitLocDesc.Enabled = True
                    Me.TxtTransitLocDesc.ReadOnly = True
                    Me.btnFindTransitLoc.Enabled = True
                Else
                    Me.TxtPrimaryTransitLoc.Text = -1
                    Me.TxtTransitLocCode.Text = ""
                    Me.TxtTransitLocDesc.Text = ""
                    Me.TxtTransitLocCode.Enabled = False
                    Me.TxtTransitLocDesc.Enabled = False
                    Me.btnFindTransitLoc.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTransitLocCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTransitLocCode.LostFocus
        TxtTransitLocCode.BackColor = Me.TxtPrimaryWarehouse.BackColor
    End Sub
End Class