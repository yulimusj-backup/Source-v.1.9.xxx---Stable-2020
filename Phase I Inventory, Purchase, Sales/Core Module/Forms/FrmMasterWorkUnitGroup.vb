Option Explicit On
Option Strict Off
Public Class FrmMasterWorkUnitGroup
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public Baru As Boolean = False
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjWG As New ClsMasterWorkUnitGroup

    Public Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            KoleksiNamaData.Add("Code")
            KoleksiTextBox.Add(Me.txtWorkUnitGroupCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.txtWorkUnitGroupDesc)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            If Me.Baru = True Then
                If Me.ObjWG.Apakahsudahada(Me.txtWorkUnitGroupCode.Text) = True Then
                    MsgBox("Duplication code, please change your code", _
                   MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                Dim productGroupLama As String = Me.ObjWG.GetCode(Me.TxtPrimary.Text)
                If productGroupLama <> Me.txtWorkUnitGroupCode.Text Then
                    If Me.ObjWG.Apakahsudahada(Me.txtWorkUnitGroupCode.Text) = True Then
                        MsgBox("Duplication code, please change your code", _
                         MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If

                'data salah jika nilai timestamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim oldTimeStamp As Date = Me.ObjWG.GetTimestamp(Me.TxtPrimary.Text)
                If Me.ObjWG.TimeStamp <> oldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjWG.CariDetail(Me.ObjWG.Primary, Me, Me.ObjCD)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimary.BackColor, False)
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.CancelButton = Me.btnExit
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmManufacturing_Order_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.btnFind.Select()
            Me.Baru = True
            Me.NonAktifkanDetail()
            Me.KosongkanField()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWorkUnitGroupCode.GotFocus, txtWorkUnitGroupDesc.GotFocus, _
    txtRemarks.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadOnly.BackColor
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

    Private Sub txtcode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWorkUnitGroupCode.LostFocus, txtWorkUnitGroupDesc.LostFocus, _
    txtRemarks.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimary.BackColor
                        End If
                    End If
                End If
            End If
            Txt.Text = Trim(Txt.Text).ToString.Trim.Replace("%", "").Replace("*", "").Replace("'", "").Replace(" ", "")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, False, True)
            Me.txtWorkUnitGroupCode.Clear()
            Me.txtRemarks.Clear()
            Me.TxtPrimary.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.Baru = True
            Me.KosongkanField()
            Me.txtWorkUnitGroupCode.Focus()
            Me.AktifkanDetail()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            With Me.ObjWG
                .Primary = Me.TxtPrimary.Text
                .Code = Me.txtWorkUnitGroupCode.Text
                .Desc = Me.txtWorkUnitGroupDesc.Text
                .Remarks = Me.txtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Try
            With Me.ObjWG
                Me.TxtPrimary.Text = .Primary
                Me.txtWorkUnitGroupCode.Text = .Code
                Me.txtWorkUnitGroupDesc.Text = .Desc
                Me.txtRemarks.Text = .Remarks
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtWorkUnitGroupCode.Enabled = True And Me.txtWorkUnitGroupCode.Text <> "" Then
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.SalinKeField()
                        Me.ObjWG.CariDetail(Me.TxtPrimary.Text, Me, Me.ObjCD)
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

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmMasterWorkUnitGroup_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryWorkUnitGroup = -1 Then
                If Me.ObjWG.CariDetail(hj.PrimaryWorkUnitGroup, Me, Me.ObjCD) = True Then
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

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimary.Text <> "-1" Then
                Me.Baru = False
                Me.ObjWG.CariDetail(Me.TxtPrimary.Text, Me, Me.ObjCD)
                Me.AktifkanDetail()
                Me.SalinKeField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjWG.SimpanBaru(Me, Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjWG.SimpanPerubahan(Me.UserCode, Me) Then
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

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtWorkUnitGroupCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtWorkUnitGroupCode.Text <> "" Then
                If Me.ObjWG.ApakahBolehHapus(Me.txtWorkUnitGroupCode.Text, True) = False Then
                    MsgBox("You cannot delete this data, there are group code using this data.", MsgBoxStyle.Information, "Information")
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjWG.Hapus(Me.TxtPrimary.Text, Me.UserCode, Me, Me.txtWorkUnitGroupCode.Text) = True Then
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
End Class