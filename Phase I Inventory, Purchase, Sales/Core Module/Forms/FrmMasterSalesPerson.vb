Option Explicit On
Option Strict Off

Public Class FrmMasterSalesPerson
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim objSalesPerson As ClsMasterSalesPerson
    Dim ObjCD As New ClsCekData
    Dim ObjLicense As New ClsLicense
    Dim ObjGroup As New ClsMasterSalesPersonGroup

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimarySalesPerson.BackColor, False)
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
            Me.TxtPrimarySalesPerson.Text = "-1"
            Me.TxtPrimarySalesPersonGroup.Text = "-1"
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
            KoleksiNamaData.Add("Sales Person Code")
            KoleksiTextBox.Add(Me.TxtSalesPersonCode)
            KoleksiNamaData.Add("description")
            KoleksiTextBox.Add(Me.TxtSalesPersonName)
            KoleksiNamaData.Add("Group Code")
            KoleksiTextBox.Add(Me.TxtSalesPersonGroup)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.objSalesPerson.ApakahSudahAda(Me.TxtSalesPersonCode.Text) = True Then
                    MsgBox("Duplication in Sales Person code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtSalesPersonCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.objSalesPerson.GetCode(Me.TxtPrimarySalesPerson.Text)
                If ProductGroupLama <> Me.TxtSalesPersonCode.Text Then
                    If Me.objSalesPerson.ApakahSudahAda(Me.TxtSalesPersonCode.Text) = True Then
                        MsgBox("Duplication in Sales Person code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtSalesPersonCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objSalesPerson.GetTimeStamp(Me.TxtPrimarySalesPerson.Text)
                If Me.objSalesPerson.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objSalesPerson.CariDetail(Me.objSalesPerson.Primary)
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
            With Me.objSalesPerson
                .Code = Me.TxtSalesPersonCode.Text
                .Description = Me.TxtSalesPersonName.Text
                .PrimaryGroup = Me.TxtPrimarySalesPersonGroup.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objSalesPerson
                Me.TxtSalesPersonCode.Text = .Code
                Me.TxtSalesPersonName.Text = .Description
                Me.TxtPrimarySalesPerson.Text = .Primary
                Me.TxtPrimarySalesPersonGroup.Text = .PrimaryGroup
                Me.TxtSalesPersonGroup.Text = Me.ObjGroup.GetCode(.PrimaryGroup)
                Me.TxtSalesPersonGroupDesc.Text = Me.ObjGroup.GetDesc(.PrimaryGroup)

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
            Me.objSalesPerson = New ClsMasterSalesPerson
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
                        If Me.objSalesPerson.SimpanBaru(Me.UserCode) Then
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
                        If Me.objSalesPerson.SimpanPerubahan(Me.UserCode) Then
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
            If Not Me.TxtSalesPersonCode.Text = "" And Me.TxtSalesPersonName.Text <> "" Then
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
            Me.TxtSalesPersonCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmMasterSalesPerson_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimarySalesPerson = -1 Then
                If Me.objSalesPerson.CariDetail(hj.PrimarySalesPerson) = True Then
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
        If Me.TxtSalesPersonCode.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.TxtSalesPersonCode.Text <> "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.TxtSalesPersonCode.Focus()
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtSalesPersonCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtSalesPersonCode.Text <> "" Then
                If Me.objSalesPerson.ApakahBolehHapus(Me.TxtPrimarySalesPerson.Text) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objSalesPerson.Hapus(Me.TxtPrimarySalesPerson.Text) = True Then
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
            If Me.TxtSalesPersonCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objSalesPerson.SimpanBaru(Me.UserCode) = True Then
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
                            If Me.objSalesPerson.SimpanPerubahan(Me.UserCode) = True Then
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

    Private Sub TxtSalesPersonCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPersonCode.GotFocus
        TxtSalesPersonCode.BackColor = BackColorAll
    End Sub

    Private Sub TxtSalesPersonCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPersonCode.LostFocus
        TxtSalesPersonCode.BackColor = Me.TxtPrimarySalesPerson.BackColor
    End Sub

    Private Sub TxtSalesPersonName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPersonName.GotFocus
        TxtSalesPersonName.BackColor = BackColorAll
    End Sub

    Private Sub TxtSalesPersonName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPersonName.LostFocus
        TxtSalesPersonName.BackColor = Me.TxtPrimarySalesPerson.BackColor
    End Sub

    Private Sub TxtSalesPersonGroup_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPersonGroup.GotFocus
        Try
            If Me.TxtSalesPersonGroup.Enabled = True Then
                If Me.TxtSalesPersonGroup.ReadOnly = False Then
                    Me.TxtSalesPersonGroup.BackColor = Me.BackColorAll
                Else
                    Me.TxtSalesPersonGroup.BackColor = Me.TxtSalesPersonGroupDesc.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesPersonGroup_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPersonGroup.LostFocus
        Try
            If Me.TxtSalesPersonGroup.Enabled = True Then
                If Me.TxtSalesPersonGroup.ReadOnly = False Then
                    Me.TxtSalesPersonGroup.BackColor = Me.TxtPrimarySalesPersonGroup.BackColor
                Else
                    Me.TxtSalesPersonGroup.BackColor = Me.TxtSalesPersonGroupDesc.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindSalesPersonGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindSalesPersonGroup.Click
        Dim FrmCariGroup As New FrmMasterSalesPersonGroup_Cari

        Try
            FrmCariGroup.BackColorAll = Me.BackColorAll
            If Me.TxtPrimarySalesPersonGroup.Text = -1 And Me.TxtSalesPersonGroup.Text = "" Then
                FrmCariGroup.ShowDialog()
            Else
                FrmCariGroup.txtCari.Text = Me.TxtSalesPersonGroup.Text
                FrmCariGroup.ShowDialog()
            End If
            If CDbl(FrmCariGroup.PrimarySalesPersonGroup) > 0 Then
                Me.TxtPrimarySalesPersonGroup.Text = FrmCariGroup.PrimarySalesPersonGroup
                Me.TxtSalesPersonGroup.Text = Me.ObjGroup.GetCode(FrmCariGroup.PrimarySalesPersonGroup)
                Me.TxtSalesPersonGroupDesc.Text = Me.ObjGroup.GetDesc(FrmCariGroup.PrimarySalesPersonGroup)
                Me.btnOK.Focus()
            Else
                Me.TxtPrimarySalesPersonGroup.Text = -1
                Me.TxtSalesPersonGroup.Text = ""
                Me.TxtSalesPersonGroupDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCariGroup.Dispose()
        End Try
    End Sub

    Private Sub TxtSalesPersonGroup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSalesPersonGroup.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjGroup.ApakahSudahAda(Me.TxtSalesPersonGroup.Text) = True Then
                    Me.TxtPrimarySalesPersonGroup.Text = Me.ObjGroup.GetPrimary(Me.TxtSalesPersonGroup.Text)
                    Me.TxtSalesPersonGroupDesc.Text = Me.ObjGroup.GetDesc(Me.TxtPrimarySalesPersonGroup.Text)
                    Me.btnOK.Focus()
                Else
                    Me.btnFindSalesPersonGroup.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesPersonGroup_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPersonGroup.Leave
        Try
            If Me.TxtSalesPersonGroup.Text <> "" Then
                If Me.ObjGroup.ApakahSudahAda(Me.TxtSalesPersonGroup.Text) = True Then
                    Me.TxtPrimarySalesPersonGroup.Text = Me.ObjGroup.GetPrimary(Me.TxtSalesPersonGroup.Text)
                    Me.TxtSalesPersonGroupDesc.Text = Me.ObjGroup.GetDesc(Me.TxtPrimarySalesPersonGroup.Text)
                Else
                    Me.btnFindSalesPersonGroup.PerformClick()
                End If
            Else
                Me.TxtPrimarySalesPersonGroup.Text = -1
                Me.TxtSalesPersonGroup.Text = ""
                Me.TxtSalesPersonGroupDesc.Text = Me.ObjGroup.GetDesc(Me.TxtPrimarySalesPersonGroup.Text)
            End If
            Me.TxtSalesPersonGroup.BackColor = Me.TxtPrimarySalesPersonGroup.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class