Option Explicit On
Option Strict Off
Public Class FrmMasterMachineLine
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim baru As Boolean
    Dim obj As New ClsMasterMachineLine
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimary.BackColor, False)
            ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, False, False, True, True)
            ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, False, False, True, False)
            ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ObjCD.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            Me.TxtPrimary.Text = "-1"
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
            KoleksiNamaData.Add("Code")
            KoleksiTextBox.Add(Me.TxtLineCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtDescription)

            If ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            If Me.baru = True Then
                If Me.obj.ApakahSudahAda(Me.TxtLineCode.Text) = True Then
                    MsgBox("Code " & Me.TxtLineCode.Text & " already exist.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtLineCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.obj.GetCode(Me.TxtPrimary.Text)
                If ProductGroupLama <> Me.TxtLineCode.Text Then
                    If Me.obj.ApakahSudahAda(Me.TxtLineCode.Text) = True Then
                        MsgBox("Code " & Me.TxtLineCode.Text & " already exist.", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtLineCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.obj.GetTimestamp(Me.TxtPrimary.Text)
                If Me.obj.Timestamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                        Me.obj.CariDetail(Me.obj.Primary)
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
            With Me.obj
                .Primary = Me.TxtPrimary.Text
                .Code = Me.TxtLineCode.Text
                .Description = Me.TxtDescription.Text
                .Remarks = Me.txtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtLineCode.Text = .Code
                Me.TxtDescription.Text = .Description
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

    Private Sub FrmMasterKota_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtLineCode.Enabled = True And Me.TxtLineCode.Text <> "" Then
                Dim hasil As MsgBoxResult = MsgBox("Save data before close?", MsgBoxStyle.YesNoCancel, "Confirmation")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.obj.SimpanBaru(Me.UserCode) = True Then
                                MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.obj.SimpanPerubahan(Me.UserCode) = True Then
                                MsgBox("Data change saved.", MsgBoxStyle.Information, "Information")
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

    Private Sub FrmMasterKota_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Text = "Production Line"
            Me.baru = True
            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'Confirmation disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirmation")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.obj.SimpanBaru(Me.UserCode) Then
                            MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'Confirmation disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirmation")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.obj.SimpanPerubahan(Me.UserCode) Then
                            MsgBox("Data change saved.", MsgBoxStyle.Information, "Information")
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

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtLineCode.Text <> "" Or Me.TxtDescription.Text <> "" Then
                'Confirmation batal atau tidak
                hasil = MsgBox("Cancel and discard data change?", MsgBoxStyle.YesNo, "Confirmation")
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
            Me.TxtLineCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmMasterMachineLine_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.Primary = -1 Then
                If Me.obj.CariDetail(hj.Primary) = True Then
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
        Try
            If Me.TxtLineCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtLineCode.Text <> "" Then
                Me.baru = False
                Me.AktifkanDetail()
                Me.TxtLineCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtLineCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtLineCode.Text <> "" Then
                If Me.obj.ApakahBolehHapus(Me.TxtPrimary.Text) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirmation")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.obj.Hapus(Me.TxtPrimary.Text) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted.", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub TxtLineCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLineCode.GotFocus, TxtDescription.GotFocus, txtRemarks.GotFocus
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            If txt.ReadOnly = False Then
                txt.BackColor = BackColorAll
            Else
                txt.BackColor = Me.TxtPrimary.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLineCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLineCode.LostFocus, TxtDescription.LostFocus, txtRemarks.LostFocus
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            If txt.ReadOnly = False And txt.Enabled = True Then
                txt.BackColor = Me.TxtPrimary.BackColor
            Else
                txt.BackColor = Me.txtReadOnly.BackColor
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtLineCode.Leave, TxtDescription.Leave, txtRemarks.Leave
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            txt.Text = txt.Text.Trim
            txt.Text = txt.Text.Replace("'", "")
            txt.Text = txt.Text.Replace("%", "")
            txt.Text = txt.Text.Replace("*", "")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLineCode.KeyDown
        If e.KeyCode = 13 Then
            Me.TxtDescription.Focus()
        End If
    End Sub

    Private Sub TxtDescription_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDescription.KeyDown
        If e.KeyCode = 13 Then
            Me.txtRemarks.Focus()
        End If
    End Sub

    Private Sub txtRemarks_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemarks.KeyDown
        If e.KeyCode = 13 Then
            Me.btnOK.Focus()
        End If
    End Sub
End Class