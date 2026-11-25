Option Explicit On
Option Strict Off

Public Class FrmMasterCompanyData

    Public UserCode As String = "USER"
    Public BackColorAll As Color = Color.LightSeaGreen

    Dim baru As Boolean
    Dim ObjPP As New ClsMasterCompanyData
    Dim buffer As String = ""
    Dim ObjCD As New ClsCekData
    Dim MrOk As Boolean = False
    Dim ObjComapanyData As New Core_Module.ClsMasterCompanyData
    Dim SedangLoad As Boolean = True

    'Comment: Perubahan NPWP 16 Digit
    Dim ObjSettingNPWP As New Core_Module.ClsTrans_SettingNPWPNo

    Private Sub PengaturanBarisAkhir()
        Try
            'tentukan nilai default untuk baris terakhir
            If Me.dgvAddress.RowCount > 0 Then
                With Me.dgvAddress.Rows(Me.dgvAddress.Rows.Count - 1)
                    .Cells(1).Value = "-"
                End With
                Me.ObjCD.SetNumber(Me.dgvAddress)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCompanyData_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Dim CompanyName As String = ""
        Dim ObjDB As New Core_Module.ClsDBVersion

        Try
            CompanyName = Me.ObjComapanyData.GetCompanyName
            If CompanyName <> "" Then
                CompanyName = " - " + CompanyName
            End If
            Me.ParentForm.Text = "SMESS - Ver. " + ObjDB.GetDBVersion + _
            CompanyName
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjDB = Nothing
        End Try
    End Sub

    Private Sub FrmMasterCompanyData_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As DataGridViewRow

        Try
            Me.KosongkanField()
            Me.dgvAddress.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            ''cari data
            If ObjPP.Pencarian(Me.dgvAddress, Me) = False Then
                Me.baru = True
                'Update PPN 12%
                Me.TxtTKU_ID.Text = "000000"
                Me.TxtCompanyName.Focus()
            Else
                Me.baru = False
                Me.dgvAddress.AllowUserToAddRows = False
                Me.SalinkeField()
                ObjPP.Pencarian(Me.dgvAddress, Me)
                Me.btnOK.Focus()
                Me.dgvAddress.AllowUserToAddRows = True
            End If

            Me.TxtCompanyName.Focus()
            Me.PengaturanBarisAkhir()
            Me.SetButtonUpDown()
            Me.SedangLoad = False

            rowDGV = Me.dgvAddress.Rows(Me.dgvAddress.Rows.Count - 1)
            rowDGV.Cells(Me.ColAddressIdentifier.Index).Value = Me.GetIdentifierBaru
            rowDGV.Cells(Me.Col_Line1.Index).Value = "-"
            rowDGV.Cells(Me.Col_Line2.Index).Value = "-"
            rowDGV.Cells(Me.Col_Line3.Index).Value = "-"
            rowDGV.Cells(Me.Col_Line4.Index).Value = "-"
            rowDGV.Cells(Me.Col_Line5.Index).Value = "-"
            rowDGV.Cells(Me.ColRemarksAdd.Index).Value = "-"
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvAddress.Rows.Count

            'Comment: Update NPWP 16 Digit
            Me.LlbEditNPWPLama.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeObjek()
        Try
            With ObjPP
                .CompanyName = Me.TxtCompanyName.Text
                .AddLine1 = Me.TxtAddLine1.Text
                .AddLine2 = Me.TxtAddLine2.Text
                .AddLine3 = Me.TxtAddLine3.Text
                .AddLine4 = Me.TxtAddLine4.Text
                .AddLine5 = Me.TxtAddLine5.Text
                .City = Me.TxtCity.Text
                .NPWP = Me.TxtNPWPNumber.Text
                .PKPRegDate = Me.dtp_PKP_Reg_Date.Value
                .TaxAddLine1 = Me.TxtTaxAddLine1.Text
                .TaxAddLine2 = Me.TxtTaxAddLine2.Text

                'Comment: Update NPWP 16 Digit
                .NPWPLama = Me.TxtNPWPLama.Text

                'Comment: Perubahan PPN 12%
                .TKU_ID = Me.TxtTKU_ID.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeField()
        Try
            With ObjPP
                Me.TxtPrimaryCompany.Text = .Primary
                Me.TxtCompanyName.Text = .CompanyName
                Me.TxtAddLine1.Text = .AddLine1
                Me.TxtAddLine2.Text = .AddLine2
                Me.TxtAddLine3.Text = .AddLine3
                Me.TxtAddLine4.Text = .AddLine4
                Me.TxtAddLine5.Text = .AddLine5
                Me.TxtCity.Text = .City
                Me.TxtNPWPNumber.Text = .NPWP
                Me.dtp_PKP_Reg_Date.Value = .PKPRegDate
                Me.TxtTaxAddLine1.Text = .TaxAddLine1
                Me.TxtTaxAddLine2.Text = .TaxAddLine2

                'Comment: Update NPWP 16 Digit
                Me.TxtNPWPLama.Text = .NPWPLama

                'Comment: Perubahan PPN 12%
                If .TKU_ID = "" Then
                    Me.TxtTKU_ID.Text = "000000"
                Else
                    Me.TxtTKU_ID.Text = .TKU_ID
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + _
                "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + _
                "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.Controls, True, True, False)
            Me.TxtPrimaryCompany.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvAddress.Rows.Clear()
            Me.ObjCD.SetNumber(Me.dgvAddress)

            'Update PPN 12%
            Me.TxtTKU_ID.Text = "000000"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Company Name")
            KoleksiTextBox.Add(Me.TxtCompanyName)
            KoleksiNamaData.Add("TKU ID")
            KoleksiTextBox.Add(Me.TxtTKU_ID)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'Cek DGVItem
            If Me.dgvAddress.Rows.Count < 1 Then
                MsgBox("Please input item", MsgBoxStyle.Information, "Information")
                Me.dgvAddress.Focus()
                Return False
            End If

            'cek apakah ada account number yang debet dan kreditnya bernilai 0
            Dim row As DataGridViewRow
            For Each row In Me.dgvAddress.Rows
                If String.IsNullOrEmpty(Me.dgvAddress.Item(Me.ColAddressCode.Index, row.Index).Value) = True Then
                    MsgBox("Please input address code at row " + " " + CStr(row.Index + 1), _
                    MsgBoxStyle.Information, "Information")
                    Return False
                ElseIf String.IsNullOrEmpty(Me.dgvAddress.Item(Me.Col_Line1.Index, row.Index).Value) = True Then
                    MsgBox("Please input line 1 at row " + " " + CStr(row.Index + 1), _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Next

            'cek perubahan nomor transaksi
            ''If Me.baru = False Then
            ''If Me.ObjPP.TanggalTransaksi.Month = Me.dtpTanggalTransaksi.Value.Month And _
            ''Me.ObjPP.TanggalTransaksi.Year = Me.dtpTanggalTransaksi.Value.Year Then
            ''Me.txtNoTransaksi.Text = Me.ObjPP.NoTransaksi
            ''Else
            ''If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
            ''MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
            ''Me.dtpTanggalTransaksi.Value = Me.ObjPP.TanggalTransaksi
            ''Me.txtNoTransaksi.Text = Me.ObjPP.NoTransaksi
            ''Return False
            ''End If
            ''End If
            ''End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.ObjPP.ApakahSudahAda(Me.TxtCompanyName.Text) = True Then
                    MsgBox("Duplication in company name =" + Me.TxtCompanyName.Text + ", please change your company name", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtCompanyName.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.ObjPP.GetCompanyName(Me.TxtPrimaryCompany.Text)
                If TipeCodeLama <> Me.TxtCompanyName.Text Then
                    If Me.ObjPP.ApakahSudahAda(Me.TxtCompanyName.Text) = True Then
                        MsgBox("Duplication in company name =" + Me.TxtCompanyName.Text + ", please change your company name", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtCompanyName.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjPP.GetTimeStamp(Me.TxtPrimaryCompany.Text)
                If Me.ObjPP.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjPP.CariDetail(Me.dgvAddress, Me)
                        Me.SalinkeField()
                    End If
                    Return False
                End If
            End If

            'Comment: Perubahan NPWP 16 Digit
            If Me.TxtNPWPLama.Text = "" Or Me.TxtNPWPLama.Text = "-" Then
                Me.TxtNPWPLama.Text = Me.TxtNPWPNumber.Text
            End If
            If Today.Date >= Me.ObjSettingNPWP.Tglberlaku Then
                If Me.ObjSettingNPWP.ApakahDataSudahBenar_MasterCompany(Me.TxtNPWPNumber.Text, Me.TxtNPWPLama.Text, True) = False Then
                    Return False
                End If
            End If

            'Update PPN 12%
            If ObjCD.CekNumerik(Me.TxtTKU_ID.Text) = False Then
                MsgBox("Must input number format", MsgBoxStyle.Information, "Information")
                Me.TxtTKU_ID.Focus()
                Return False
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

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            Me.dgvAddress.AllowUserToAddRows = False
            Me.SalinkeObjek()
            If baru = True Then
                If ObjPP.ApakahDataBenar(baru) = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If Me.ApakahDataBenar() = True Then
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPP.SimpanBaru(Me.dgvAddress, Me.UserCode, Me) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.baru = False
                                Me.MrOk = True
                                Me.Close()
                                ''Me.NonAktifkanDetail()
                                ''Me.btnFind.Focus()
                            End If
                            Me.dgvAddress.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            Me.dgvAddress.AllowUserToAddRows = True
                            Me.PengaturanBarisAkhir()
                        End If
                    Else
                        Me.dgvAddress.AllowUserToAddRows = True
                        Me.PengaturanBarisAkhir()
                    End If
                Else
                    Me.dgvAddress.AllowUserToAddRows = True
                    Me.PengaturanBarisAkhir()
                End If
            Else
                'konfirmasi disimpan atau tidak
                If Me.ApakahDataBenar = True Then
                    If ObjPP.ApakahBolehEdit(Me.TxtPrimaryCompany.Text, UserCode, True, True) = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPP.SimpanPerubahan(Me.dgvAddress, Me.UserCode, Me) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.baru = False
                                Me.MrOk = True
                                Me.Close()
                                ''Me.NonAktifkanDetail()
                                ''Me.btnEdit.Focus()
                            End If
                            Me.dgvAddress.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            Me.dgvAddress.AllowUserToAddRows = True
                            Me.PengaturanBarisAkhir()
                        End If
                        Me.dgvAddress.AllowUserToAddRows = True
                        Me.PengaturanBarisAkhir()
                    End If
                Else
                    Me.dgvAddress.AllowUserToAddRows = True
                    Me.PengaturanBarisAkhir()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            Me.dgvAddress.AllowUserToAddRows = False
            If Me.dgvAddress.RowCount = 0 And Me.TxtCompanyName.Text = "" Then
                Me.MrOk = True
                Me.Close()
                ''Me.KosongkanField()
                ''Me.dgvAddress.AllowUserToAddRows = True
                ''ObjCD.SetNumber(Me.dgvAddress)
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinkeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPP.SimpanBaru(Me.dgvAddress, Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.baru = False
                                Me.MrOk = True
                                Me.Close()
                                Me.btnOK.Focus()
                            Else
                                ''e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPP.SimpanPerubahan(Me.dgvAddress, Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.baru = False
                                Me.MrOk = True
                                Me.Close()
                                Me.btnOK.Focus()
                            Else
                                ''e.Cancel = True
                            End If
                        End If
                    Else
                        ''e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.No Then
                    Me.MrOk = True
                    Me.Close()
                End If
            End If
            Me.dgvAddress.AllowUserToAddRows = True
            Me.ObjCD.SetNumber(Me.dgvAddress)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.dgvAddress.AllowUserToAddRows = False
            If Me.MrOk = False Then
                If Me.TxtCompanyName.Enabled = True Then
                    Me.BringToFront()
                    Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ApakahDataBenar() = True Then
                            Me.SalinkeObjek()
                            If baru = True Then
                                'simpan baru
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.ObjPP.SimpanBaru(Me.dgvAddress, Me.UserCode, Me) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.SalinkeField()
                                    ''Me.NonAktifkanDetail()
                                    ''Me.btnFind.Focus()
                                    Me.btnOK.Focus()
                                Else
                                    e.Cancel = True
                                End If
                            Else
                                'simpan perubahan
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.ObjPP.SimpanPerubahan(Me.dgvAddress, Me.UserCode, Me) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.SalinkeField()
                                    ''Me.NonAktifkanDetail()
                                    ''Me.btnEdit.Focus()
                                    Me.btnOK.Focus()
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
            End If
            Me.dgvAddress.AllowUserToAddRows = True
            Me.ObjCD.SetNumber(Me.dgvAddress)


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAddLine1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddLine1.GotFocus
        Try
            If Me.TxtAddLine1.Enabled = True Then
                If Me.TxtAddLine1.ReadOnly = False Then
                    Me.TxtAddLine1.BackColor = Me.BackColorAll
                Else
                    TxtAddLine1.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAddLine1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddLine1.LostFocus
        Try
            If Me.TxtAddLine1.Enabled = True Then
                If Me.TxtAddLine1.ReadOnly = False Then
                    Me.TxtAddLine1.BackColor = Me.TxtPrimaryCompany.BackColor
                Else
                    TxtAddLine1.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAddLine2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddLine2.GotFocus
        Try
            If Me.TxtAddLine2.Enabled = True Then
                If Me.TxtAddLine2.ReadOnly = False Then
                    Me.TxtAddLine2.BackColor = Me.BackColorAll
                Else
                    TxtAddLine2.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAddLine2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddLine2.LostFocus
        Try
            If Me.TxtAddLine2.Enabled = True Then
                If Me.TxtAddLine2.ReadOnly = False Then
                    Me.TxtAddLine2.BackColor = Me.TxtPrimaryCompany.BackColor
                Else
                    TxtAddLine2.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAddLine3_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddLine3.GotFocus
        Try
            If Me.TxtAddLine3.Enabled = True Then
                If Me.TxtAddLine3.ReadOnly = False Then
                    Me.TxtAddLine3.BackColor = Me.BackColorAll
                Else
                    TxtAddLine3.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAddLine3_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddLine3.LostFocus
        Try
            If Me.TxtAddLine3.Enabled = True Then
                If Me.TxtAddLine3.ReadOnly = False Then
                    Me.TxtAddLine3.BackColor = Me.TxtPrimaryCompany.BackColor
                Else
                    TxtAddLine3.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAddLine4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddLine4.GotFocus
        Try
            If Me.TxtAddLine4.Enabled = True Then
                If Me.TxtAddLine4.ReadOnly = False Then
                    Me.TxtAddLine4.BackColor = Me.BackColorAll
                Else
                    TxtAddLine4.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAddLine4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddLine4.LostFocus
        Try
            If Me.TxtAddLine4.Enabled = True Then
                If Me.TxtAddLine4.ReadOnly = False Then
                    Me.TxtAddLine4.BackColor = Me.TxtPrimaryCompany.BackColor
                Else
                    TxtAddLine4.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAddLine5_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddLine5.GotFocus
        Try
            If Me.TxtAddLine5.Enabled = True Then
                If Me.TxtAddLine5.ReadOnly = False Then
                    Me.TxtAddLine5.BackColor = Me.BackColorAll
                Else
                    TxtAddLine5.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAddLine5_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddLine5.LostFocus
        Try
            If Me.TxtAddLine5.Enabled = True Then
                If Me.TxtAddLine5.ReadOnly = False Then
                    Me.TxtAddLine5.BackColor = Me.TxtPrimaryCompany.BackColor
                Else
                    TxtAddLine5.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCompanyName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCompanyName.GotFocus
        Try
            If Me.TxtCompanyName.Enabled = True Then
                If Me.TxtCompanyName.ReadOnly = False Then
                    Me.TxtCompanyName.BackColor = Me.BackColorAll
                Else
                    TxtCompanyName.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCompanyName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCompanyName.LostFocus
        Try
            If Me.TxtCompanyName.Enabled = True Then
                If Me.TxtCompanyName.ReadOnly = False Then
                    Me.TxtCompanyName.BackColor = Me.TxtPrimaryCompany.BackColor
                Else
                    TxtCompanyName.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCity_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCity.GotFocus
        Try
            If Me.TxtCity.Enabled = True Then
                If Me.TxtCity.ReadOnly = False Then
                    Me.TxtCity.BackColor = Me.BackColorAll
                Else
                    TxtCity.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCity_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCity.LostFocus
        Try
            If Me.TxtCity.Enabled = True Then
                If Me.TxtCity.ReadOnly = False Then
                    Me.TxtCity.BackColor = Me.TxtPrimaryCompany.BackColor
                Else
                    TxtCity.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNPWPNumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWPNumber.GotFocus
        Try
            If Me.TxtNPWPNumber.Enabled = True Then
                If Me.TxtNPWPNumber.ReadOnly = False Then
                    Me.TxtNPWPNumber.BackColor = Me.BackColorAll
                Else
                    TxtNPWPNumber.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNPWPNumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWPNumber.LostFocus
        Try
            If Me.TxtNPWPNumber.Enabled = True Then
                If Me.TxtNPWPNumber.ReadOnly = False Then
                    Me.TxtNPWPNumber.BackColor = Me.TxtPrimaryCompany.BackColor
                Else
                    TxtNPWPNumber.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxAddLine1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxAddLine1.GotFocus
        Try
            If Me.TxtTaxAddLine1.Enabled = True Then
                If Me.TxtTaxAddLine1.ReadOnly = False Then
                    Me.TxtTaxAddLine1.BackColor = Me.BackColorAll
                Else
                    TxtTaxAddLine1.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxAddLine1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxAddLine1.LostFocus
        Try
            If Me.TxtTaxAddLine1.Enabled = True Then
                If Me.TxtTaxAddLine1.ReadOnly = False Then
                    Me.TxtTaxAddLine1.BackColor = Me.TxtPrimaryCompany.BackColor
                Else
                    TxtTaxAddLine1.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxAddLine2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxAddLine2.GotFocus
        Try
            If Me.TxtTaxAddLine2.Enabled = True Then
                If Me.TxtTaxAddLine2.ReadOnly = False Then
                    Me.TxtTaxAddLine2.BackColor = Me.BackColorAll
                Else
                    TxtTaxAddLine2.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxAddLine2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxAddLine2.LostFocus
        Try
            If Me.TxtTaxAddLine2.Enabled = True Then
                If Me.TxtTaxAddLine2.ReadOnly = False Then
                    Me.TxtTaxAddLine2.BackColor = Me.TxtPrimaryCompany.BackColor
                Else
                    TxtTaxAddLine2.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Pengaturan DGV Address
    Private Sub dgvAddress_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvAddress.CellBeginEdit
        Try
            If e.ColumnIndex = 0 Then
                e.Cancel = True
            Else
                Me.buffer = Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAddress_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAddress.CellEndEdit
        Try
            If IsNothing(Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColAddressCode.Index Then
                    Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = ""
                End If
            Else
                If e.ColumnIndex = 1 Then
                    'untuk address code tidak boleh kosong
                    If Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        MsgBox("Address code cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                    'harus huruf kapital semua
                    Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                    'kode alamat tidak ada yang boleh double
                    Dim row As DataGridViewRow
                    For Each row In Me.dgvAddress.Rows
                        If row.Cells(1).Value = Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value And row.Index <> e.RowIndex Then
                            MsgBox("Duplication in address code, please change your address code.", MsgBoxStyle.Information, "Information")
                            Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    Next
                ElseIf e.ColumnIndex = Me.Col_Line1.Index Or e.ColumnIndex = Me.Col_Line2.Index Or _
                e.ColumnIndex = Me.Col_Line3.Index Or e.ColumnIndex = Me.Col_Line4.Index Or _
                e.ColumnIndex = Me.Col_Line5.Index Then
                    If Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 40 Then
                        MsgBox("Address data maximum length is 40 character.", MsgBoxStyle.Information, "Information")
                        Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 40)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAddress_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvAddress.UserAddedRow
       
    End Sub

    Private Sub dgvAddress_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvAddress.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvAddress)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAddress_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvAddress.UserDeletingRow
        Try
            'tidak boleh hapus baris terakhir
            If e.Row.Index = Me.dgvAddress.Rows.Count - 1 Then
                e.Cancel = True
            End If

            'tidak boleh dihapus jika sudah ada yg merujuk
            If Me.ObjComapanyData.ApakahBolehHapusAddress(Me.dgvAddress.Item(Me.ColAddressIdentifier.Index, e.Row.Index).Value, True) = False Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetButtonUpDown()
        Try
            If Me.dgvAddress.SelectedRows.Count > 0 Then
                If Me.dgvAddress.SelectedRows(0).Index <> Me.dgvAddress.Rows.Count - 1 Then
                    If Me.dgvAddress.SelectedRows(0).Index = Me.dgvAddress.RowCount - 2 And Me.dgvAddress.SelectedRows(0).Index = 0 Then
                        Me.btnUP.Enabled = False
                        Me.btnDown.Enabled = False
                    ElseIf Me.dgvAddress.SelectedRows(0).Index = Me.dgvAddress.RowCount - 2 Then
                        Me.btnUP.Enabled = True
                        Me.btnDown.Enabled = False
                    ElseIf Me.dgvAddress.SelectedRows(0).Index = 0 Then
                        Me.btnUP.Enabled = False
                        Me.btnDown.Enabled = True
                    Else
                        Me.btnUP.Enabled = True
                        Me.btnDown.Enabled = True
                    End If
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            Else
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAddress_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAddress.SelectionChanged
        Me.SetButtonUpDown()
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.dgvAddress.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvAddress.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvAddress.Rows.Remove(row)
                    Me.dgvAddress.Rows.InsertRange(indexbaru, row)
                    Me.dgvAddress.Focus()
                    Me.dgvAddress.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvAddress)
                    If row.Displayed = False Then
                        Me.ShowRow(row, True)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvAddress.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvAddress.FirstDisplayedScrollingRowIndex = Me.dgvAddress.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvAddress.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvAddress.Rows(row.Index).Displayed = False
                    Me.dgvAddress.FirstDisplayedScrollingRowIndex = Me.dgvAddress.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvAddress.DisplayedRowCount(True) <> Me.dgvAddress.DisplayedRowCount(False) Then
                    Me.dgvAddress.FirstDisplayedScrollingRowIndex = Me.dgvAddress.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvAddress.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvAddress.SelectedRows(0)
                If row.Index <> Me.dgvAddress.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvAddress.Rows.Remove(row)
                    Me.dgvAddress.Rows.InsertRange(indexbaru, row)
                    Me.dgvAddress.Focus()
                    Me.dgvAddress.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvAddress)
                    Me.ShowRow(row, False)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIdentifierBaru() As String
        Dim rowDGV As DataGridViewRow
        Dim MaxIdentifier As Integer = 0
        Dim Index As Integer = 0
        Dim Identifier As Integer = 0
        Dim Batas As Integer = 0

        Try
            If Me.dgvAddress.AllowUserToAddRows = True Then
                Batas = Me.dgvAddress.Rows.Count - 2
            Else
                Batas = Me.dgvAddress.Rows.Count - 1
            End If
            For Each rowDGV In Me.dgvAddress.Rows
                If rowDGV.Index <= Batas Then
                    Index = Me.dgvAddress.Item(Me.ColAddressIdentifier.Index, rowDGV.Index).Value.ToString.IndexOf("-")
                    If Index = 0 Then
                        Me.dgvAddress.Item(Me.ColAddressIdentifier.Index, rowDGV.Index).Value = Me.dgvAddress.Item(Me.ColAddressIdentifier.Index, rowDGV.Index).Value.ToString.Substring(1)
                        Index = Me.dgvAddress.Item(Me.ColAddressIdentifier.Index, rowDGV.Index).Value.ToString.IndexOf("-")
                    End If
                    If Index <> -1 Then
                        Identifier = rowDGV.Cells(Me.ColAddressIdentifier.Index).Value.ToString.Substring(Index + 1)
                        If MaxIdentifier < Identifier Then
                            MaxIdentifier = Identifier
                        End If
                    End If
                End If
            Next

            Return Me.TxtPrimaryCompany.Text + "-" + (MaxIdentifier + 1).ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub dgvAddress_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvAddress.RowsAdded
        Try
            If Me.SedangLoad = True Then
                Exit Sub
            End If
            'tentukan nilai default untuk baris terakhir
            With Me.dgvAddress.Rows(Me.dgvAddress.Rows.Count - 1)
                .Cells(1).Value = "-"
                .Cells(Me.Col_Line1.Index).Value = "-"
                .Cells(Me.Col_Line2.Index).Value = "-"
                .Cells(Me.Col_Line3.Index).Value = "-"
                .Cells(Me.Col_Line4.Index).Value = "-"
                .Cells(Me.Col_Line5.Index).Value = "-"
                .Cells(Me.ColRemarksAdd.Index).Value = "-"
                .Cells(Me.ColAddressIdentifier.Index).Value = Me.GetIdentifierBaru
            End With
            Me.ObjCD.SetNumber(Me.dgvAddress)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbEditNPWPLama_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LlbEditNPWPLama.Click
        'Comment: Perubahan NPWP 16 Digit
        Dim hj As New FrmMaster_EditNPWPLama
        Try
            If Me.TxtNPWPLama.Text = "" Then
                Me.TxtNPWPLama.Text = Me.TxtNPWPNumber.Text
            End If

            hj.TxtNPWPNo.Text = Me.TxtNPWPLama.Text
            If hj.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TxtNPWPLama.Text = hj.TxtNPWPNo.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub TxtTKU_ID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTKU_ID.GotFocus
        Try
            If Me.TxtTKU_ID.Enabled = True Then
                If Me.TxtTKU_ID.ReadOnly = False Then
                    Me.TxtTKU_ID.BackColor = Me.BackColorAll
                Else
                    TxtTKU_ID.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTKU_ID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTKU_ID.LostFocus
        Try
            If Me.TxtTKU_ID.Enabled = True Then
                If Me.TxtTKU_ID.ReadOnly = False Then
                    Me.TxtTKU_ID.BackColor = Me.TxtPrimaryCompany.BackColor
                Else
                    TxtTKU_ID.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class