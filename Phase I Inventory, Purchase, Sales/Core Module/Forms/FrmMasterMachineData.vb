Option Explicit On
Option Strict Off

Public Class FrmMasterMachineData
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim objMachineData As New ClsMasterMachineData
    Dim objMachineGroup As New ClsMasterMachineGroup
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.GboDetail.Controls, Me.TxtPrimary.BackColor, False)
            Core_Module.ModCekData.KendalikanControl(Me.GboDetail.Controls, True, True, True, True, False, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.ChkActive.Enabled = True
            Me.CancelButton = Me.BtnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Core_Module.ModCekData.KendalikanControl(Me.GboDetail.Controls, True, False, True, False, False, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.ChkActive.Enabled = False
            Me.CancelButton = Me.BtnExit
            Me.ObjCD.SetBackColor(Me.GboDetail.Controls, Me.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Core_Module.ModCekData.KosongkanControl(Me.GboDetail.Controls, True, False, False)
            Me.TxtPrimary.Text = "-1"
            Me.TxtPrimaryMachineGroup.Text = "-1"
            Me.ChkActive.Checked = False
            Me.TslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.TslModifiedBy.Text = "Modified By: "
            Me.TslModifiedDate.Text = "Modified Date: "
            Me.ChkActive.Checked = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'trim
            Me.TxtMachineDataCode.Text = Me.TxtMachineDataCode.Text.Trim
            Me.TxtDesc.Text = Me.TxtDesc.Text.Trim

            'cek data yang harus diisi
            KoleksiNamaData.Add("Machine Group Code")
            KoleksiTextBox.Add(Me.TxtMachineDataCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtDesc)
            KoleksiNamaData.Add("Machine Group Code")
            KoleksiTextBox.Add(Me.TxtMachineGroupCode)
            If Core_Module.ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.objMachineData.ApakahSudahAda(Me.TxtMachineDataCode.Text) = True Then
                    MsgBox("Data duplication, please change machine group code.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtMachineDataCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.objMachineData.GetCode(Me.TxtPrimary.Text)
                If ProductGroupLama <> Me.TxtMachineDataCode.Text Then
                    If Me.objMachineData.ApakahSudahAda(Me.TxtMachineDataCode.Text) = True Then
                        MsgBox("Data duplication, please change machine group code.", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtMachineDataCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objMachineData.GetTimeStamp(Me.TxtPrimary.Text)
                If Me.objMachineData.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objMachineData.CariDetail(Me.objMachineData.Primary)
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
            With Me.objMachineData
                .MachineDataCode = Me.TxtMachineDataCode.Text
                .Description = Me.TxtDesc.Text
                .PrimaryMachineGroup = Me.TxtPrimaryMachineGroup.Text
                .Remarks = Me.TxtRemarks.Text
                If Me.ChkActive.Checked = False Then
                    .Active = ClsMasterMachineData.enumActive.NonActive
                Else
                    .Active = ClsMasterMachineData.enumActive.Active
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objMachineData
                Me.TxtPrimary.Text = .Primary
                Me.TxtMachineDataCode.Text = .MachineDataCode
                Me.TxtDesc.Text = .Description
                Me.TxtPrimaryMachineGroup.Text = .PrimaryMachineGroup
                Me.TxtMachineGroupCode.Text = objMachineGroup.GetCode(Me.TxtPrimaryMachineGroup.Text)
                Me.TxtMachineGroupDesc.Text = objMachineGroup.GetDesc(Me.TxtPrimaryMachineGroup.Text)
                Me.TxtRemarks.Text = .Remarks
                If .Active = ClsMasterMachineData.enumActive.NonActive Then
                    Me.ChkActive.Checked = ClsMasterMachineData.enumActive.NonActive
                Else
                    Me.ChkActive.Checked = ClsMasterMachineData.enumActive.Active
                End If

                Me.TslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.TslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.TslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.TslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterMachineGroup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.objMachineData = New ClsMasterMachineData
            Me.TxtMachineGroupDesc.ReadOnly = True
            Me.baru = True
            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOk.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objMachineData.SimpanBaru(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.BtnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objMachineData.SimpanPerubahan(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.BtnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtMachineDataCode.Text = "" And Me.TxtDesc.Text <> "" Then
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
            Me.BtnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnNew.Click
        Try
            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtMachineDataCode.Focus()
            Me.TslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.TslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.TslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.TslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFind.Click
        Dim hj As New FrmMasterMachineDataCari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.Status = "Show All"
            hj.ShowDialog()
            If Not hj.PrimaryMachineData = -1 Then
                If Me.objMachineData.CariDetail(hj.PrimaryMachineData) = True Then
                    Me.SalinKeField()
                    Me.BtnEdit.Focus()
                Else
                    Me.BtnFind.Focus()
                End If
                Me.BtnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnEdit.Click
        If Me.TxtMachineDataCode.Text = "" Then
            Me.BtnFind.PerformClick()
        End If

        If Me.TxtMachineDataCode.Text <> "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.TxtMachineDataCode.Focus()
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtMachineDataCode.Text = "" Then
                Me.BtnFind.PerformClick()
            End If

            If Me.TxtMachineDataCode.Text <> "" Then
                If Me.objMachineData.ApakahBolehHapus(Me.TxtPrimary.Text) = False Then
                    Me.BtnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objMachineData.Hapus(Me.TxtPrimary.Text) = True Then
                            Me.KosongkanField()
                            Me.BtnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtMachineDataCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objMachineData.SimpanBaru(Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.BtnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objMachineData.SimpanPerubahan(Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.BtnEdit.Focus()
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

    Private Sub TxtMachineDataCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMachineDataCode.GotFocus, TxtDesc.GotFocus, TxtRemarks.GotFocus, TxtMachineGroupCode.GotFocus, TxtMachineGroupDesc.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.BtnOk.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMachineDataCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMachineDataCode.LostFocus, TxtDesc.LostFocus, TxtRemarks.LostFocus, TxtMachineGroupCode.LostFocus, TxtMachineGroupDesc.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.BtnOk.Enabled = True Then
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

    Private Sub BtnFindMachineGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindMachineGroup.Click
        Try
            Dim hj As New FrmMasterMachineGroupCari
            Try
                hj.BackColorAll = Me.BackColorAll
                If Me.TxtMachineGroupCode.Text <> "" Then
                    hj.txtCari.Text = Me.TxtMachineGroupCode.Text
                End If

                hj.ShowDialog()
                If hj.PrimaryMachineGroup > 0 Then
                    Me.TxtPrimaryMachineGroup.Text = hj.PrimaryMachineGroup
                    Me.TxtMachineGroupCode.Text = Me.objMachineGroup.GetCode(Me.TxtPrimaryMachineGroup.Text)
                    Me.TxtMachineGroupDesc.Text = Me.objMachineGroup.GetDesc(Me.TxtPrimaryMachineGroup.Text)
                    Me.TxtRemarks.Focus()
                Else
                    Me.TxtPrimaryMachineGroup.Text = "-1"
                    Me.TxtMachineGroupDesc.Clear()
                    Me.TxtMachineGroupCode.Clear()
                    Me.TxtMachineGroupCode.Focus()
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Finally
                hj.Dispose()
            End Try
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMachineGroupCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMachineGroupCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtMachineGroupCode.Text <> "" Then
                    If Me.objMachineGroup.ApakahSudahAda_Active(Me.TxtMachineGroupCode.Text) = True Then
                        Me.TxtPrimaryMachineGroup.Text = Me.objMachineGroup.GetPrimary(Me.TxtMachineGroupCode.Text)
                        Me.TxtMachineGroupCode.Text = Me.objMachineGroup.GetCode(Me.TxtPrimaryMachineGroup.Text)
                        Me.TxtMachineGroupDesc.Text = Me.objMachineGroup.GetDesc(Me.TxtPrimaryMachineGroup.Text)
                        Me.TxtRemarks.Focus()
                    Else
                        Me.BtnFindMachineGroup.PerformClick()
                    End If
                Else
                    Me.BtnFindMachineGroup.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMachineGroupCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMachineGroupCode.Leave
        Dim hj As New FrmMasterMachineGroupCari
        hj.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtMachineGroupCode.Text <> "" Then
                If Me.objMachineGroup.ApakahSudahAda_Active(Me.TxtMachineGroupCode.Text) = True Then
                    Me.TxtPrimaryMachineGroup.Text = Me.objMachineGroup.GetPrimary(Me.TxtMachineGroupCode.Text)
                    Me.TxtMachineGroupCode.Text = Me.objMachineGroup.GetCode(Me.TxtPrimaryMachineGroup.Text)
                    Me.TxtMachineGroupDesc.Text = Me.objMachineGroup.GetDesc(Me.TxtPrimaryMachineGroup.Text)
                    Me.TxtRemarks.Focus()
                Else
                    If Me.TxtMachineGroupCode.Text <> "" Then
                        hj.txtCari.Text = Me.TxtMachineGroupCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryMachineGroup > 0 Then
                        Me.TxtPrimaryMachineGroup.Text = hj.PrimaryMachineGroup
                        Me.TxtMachineGroupCode.Text = Me.objMachineGroup.GetCode(Me.TxtPrimaryMachineGroup.Text)
                        Me.TxtMachineGroupDesc.Text = Me.objMachineGroup.GetDesc(Me.TxtPrimaryMachineGroup.Text)
                        Me.TxtRemarks.Focus()
                    Else
                        Me.TxtPrimaryMachineGroup.Text = "-1"
                        Me.TxtMachineGroupDesc.Clear()
                        Me.TxtMachineGroupCode.Clear()
                        Me.TxtMachineGroupCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryMachineGroup.Text = "-1"
                Me.TxtMachineGroupDesc.Clear()
                Me.TxtMachineGroupCode.Clear()
                'Me.TxtRemarks.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
End Class