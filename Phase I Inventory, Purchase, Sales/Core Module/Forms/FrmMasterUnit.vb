Option Explicit On
Option Strict Off

Public Class FrmMasterUnit
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim objUnit As ClsMasterUnit
    Dim ObjCD As New ClsCekData
    Dim ObjLicense As New ClsLicense
    Dim ObjTaxUnit As New ClsMasterTaxMeasurementReference

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryUnit.BackColor, False)
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
            Me.TxtPrimaryUnit.Text = "-1"
            Me.TxtPrimaryTaxUnit.Text = "-1" '@ahdar
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
            KoleksiNamaData.Add("unit code")
            KoleksiTextBox.Add(Me.TxtUnitCode)
            KoleksiNamaData.Add("description")
            KoleksiTextBox.Add(Me.TxtUnitName)
            KoleksiNamaData.Add("Tax Measurement Ref.") '@ahdar =>
            KoleksiTextBox.Add(Me.TxtTaxUnit)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.objUnit.ApakahSudahAda(Me.TxtUnitCode.Text) = True Then
                    MsgBox("Duplication in unit code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtUnitCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.objUnit.GetUnitCode(Me.TxtPrimaryUnit.Text)
                If ProductGroupLama <> Me.TxtUnitCode.Text Then
                    If Me.objUnit.ApakahSudahAda(Me.TxtUnitCode.Text) = True Then
                        MsgBox("Duplication in unit code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtUnitCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objUnit.GetTimeStamp(Me.TxtPrimaryUnit.Text)
                If Me.objUnit.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objUnit.CariDetail(Me.objUnit.Primary)
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
            With Me.objUnit
                .UnitCode = Me.TxtUnitCode.Text
                .Description = Me.TxtUnitName.Text
                .PrimaryTaxUnit = Me.TxtPrimaryTaxUnit.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()

        Try
            With Me.objUnit
                Me.TxtUnitCode.Text = .UnitCode
                Me.TxtUnitName.Text = .Description
                Me.TxtPrimaryUnit.Text = .Primary
                Me.TxtPrimaryTaxUnit.Text = .PrimaryTaxUnit
                Me.TxtTaxUnit.Text = Me.ObjTaxUnit.GetCode(.PrimaryTaxUnit)
                Me.TxtTaxUnitDesc.Text = Me.ObjTaxUnit.GetDesc(.PrimaryTaxUnit)
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
            Me.objUnit = New ClsMasterUnit
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
                        If Me.objUnit.SimpanBaru(Me.UserCode) Then
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
                        If Me.objUnit.SimpanPerubahan(Me.UserCode) Then
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
            If Not Me.TxtUnitCode.Text = "" And Me.TxtUnitName.Text <> "" Then
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
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtUnitCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariMasterUnit
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryUnit = -1 Then
                If Me.objUnit.CariDetail(hj.PrimaryUnit) = True Then
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
        If Me.TxtUnitCode.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.TxtUnitCode.Text <> "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.TxtUnitCode.Focus()
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtUnitCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtUnitCode.Text <> "" Then
                If Me.objUnit.ApakahBolehHapus(Me.TxtPrimaryUnit.Text) = False Then
                    MsgBox("You cannot delete this data, because there are products using this data.", _
                    MsgBoxStyle.Information, "Information")
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objUnit.Hapus(Me.TxtPrimaryUnit.Text) = True Then
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
            If Me.TxtUnitCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objUnit.SimpanBaru(Me.UserCode) = True Then
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
                            If Me.objUnit.SimpanPerubahan(Me.UserCode) = True Then
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

    Private Sub TxtUnitCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitCode.GotFocus
        TxtUnitCode.BackColor = BackColorAll
    End Sub

    Private Sub TxtUnitCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitCode.LostFocus
        TxtUnitCode.BackColor = Me.TxtPrimaryUnit.BackColor
    End Sub

    Private Sub TxtUnitName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitName.GotFocus
        TxtUnitName.BackColor = BackColorAll
    End Sub

    Private Sub TxtUnitName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitName.LostFocus
        TxtUnitName.BackColor = Me.TxtPrimaryUnit.BackColor
    End Sub

    '@ahdar=>
#Region "Update 2025"
    Private Sub BtnFindTaxUnit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindTaxUnit.Click
        Dim obj As New ClsMasterTaxMeasurementReference
        Dim hj As New FrmMasterTaxMeasurementReference_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            If Me.TxtTaxUnit.Text <> "" Then
                hj.txtCari.Text = Me.TxtTaxUnit.Text
            End If

            hj.ShowDialog()
            If hj.Primary > 0 Then
                Me.TxtPrimaryTaxUnit.Text = hj.Primary
                Me.TxtTaxUnit.Text = obj.GetCode(Me.TxtPrimaryTaxUnit.Text)
                Me.TxtTaxUnitDesc.Text = obj.GetDesc(Me.TxtPrimaryTaxUnit.Text)
                Me.btnOK.Focus()
            Else
                Me.TxtPrimaryTaxUnit.Text = "-1"
                Me.TxtTaxUnitDesc.Clear()
                Me.TxtTaxUnit.Clear()
                Me.TxtTaxUnit.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub TxtTaxUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxUnit.KeyDown
        Dim obj As New ClsMasterTaxMeasurementReference
        Try
            If e.KeyValue = 13 Then
                If Me.TxtTaxUnit.Text <> "" Then
                    If obj.ApakahSudahAda(Me.TxtTaxUnit.Text) = True Then
                        Me.TxtPrimaryTaxUnit.Text = obj.GetPrimary(Me.TxtTaxUnit.Text)
                        Me.TxtTaxUnit.Text = obj.GetCode(Me.TxtPrimaryTaxUnit.Text)
                        Me.TxtTaxUnitDesc.Text = obj.GetDesc(Me.TxtPrimaryTaxUnit.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.BtnFindTaxUnit.PerformClick()
                    End If
                Else
                    Me.BtnFindTaxUnit.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxUnit.Leave
        Dim obj As New ClsMasterTaxMeasurementReference
        Dim hj As New FrmMasterTaxMeasurementReference_Cari
        hj.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtTaxUnit.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtTaxUnit.Text) = True Then
                    Me.TxtPrimaryTaxUnit.Text = obj.GetPrimary(Me.TxtTaxUnit.Text)
                    Me.TxtTaxUnit.Text = obj.GetCode(Me.TxtPrimaryTaxUnit.Text)
                    Me.TxtTaxUnitDesc.Text = obj.GetDesc(Me.TxtPrimaryTaxUnit.Text)
                Else
                    If Me.TxtTaxUnit.Text <> "" Then
                        hj.txtCari.Text = Me.TxtTaxUnit.Text
                    End If
                    hj.ShowDialog()

                    If hj.Primary > 0 Then
                        Me.TxtPrimaryTaxUnit.Text = hj.Primary
                        Me.TxtTaxUnit.Text = obj.GetCode(Me.TxtPrimaryTaxUnit.Text)
                        Me.TxtTaxUnitDesc.Text = obj.GetDesc(Me.TxtPrimaryTaxUnit.Text)
                    Else
                        Me.TxtPrimaryTaxUnit.Text = "-1"
                        Me.TxtTaxUnitDesc.Clear()
                        Me.TxtTaxUnit.Clear()
                        Me.TxtTaxUnit.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryTaxUnit.Text = "-1"
                Me.TxtTaxUnitDesc.Clear()
                Me.TxtTaxUnit.Clear()
                'Me.TxtRemarks.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub TxtTaxUnit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxUnit.GotFocus
        TxtTaxUnit.BackColor = BackColorAll
    End Sub

    Private Sub TxtTaxUnit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxUnit.LostFocus
        TxtTaxUnit.BackColor = Me.TxtPrimaryTaxUnit.BackColor
    End Sub
#End Region
 
End Class