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

    Private Sub FrmMasterCompanyData_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.KosongkanField()
            Me.dgvAddress.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            ''cari data
            If ObjPP.Pencarian(Me.dgvAddress, Me) = False Then
                Me.baru = True
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
            Me.WindowState = FormWindowState.Maximized
            Me.PengaturanBarisAkhir()
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
                If e.ColumnIndex = 1 Then
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
        Try
            'tentukan nilai default untuk baris terakhir
            With Me.dgvAddress.Rows(Me.dgvAddress.Rows.Count - 1)
                .Cells(1).Value = "-"
            End With
            Me.ObjCD.SetNumber(Me.dgvAddress)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
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
            If e.Row.Index = Me.dgvAddress.Rows.Count - 1 Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class