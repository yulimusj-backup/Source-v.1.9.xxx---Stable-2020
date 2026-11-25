Option Explicit On
Option Strict Off
Public Class FrmMasterCustomer
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public FromCekEmailCP As Boolean = False
    Public vPrimaryCustCekEmailCP As Double = -1

    Public baru As Boolean
    Public objCustomer As New clsMasterCustomer
    Dim ObjCekData As New ClsCekData
    Dim ObjCustGroup As New clsMasterCustomerGroup
    Dim ObjCustType As New clsMasterCustomerType
    Dim ObjCustSalesArea As New clsMasterCustomerSalesArea
    Dim ObjCustSegmentation As New clsMasterCustomerMarketSegmentation
    Dim ObjCurrency As New clsMasterCurrency
    Dim ObjCD As New ClsCekData
    Dim ObjSalesP As New ClsMasterSalesPerson
    Dim ObjRef As New ClsMasterCustomerReference
    Dim OptionNIKPASS As clsMasterCustomer.EnumOptionNikPass = clsMasterCustomer.EnumOptionNikPass.ByNIK
    Dim buffer As String = ""
    Dim ObjLicense As New ClsLicense
    'Comment: Perubahan NPWP 16 Digit
    Dim ObjSettingNPWP As New Core_Module.ClsTrans_SettingNPWPNo

    Public Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryCustomer.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboGrouping.Controls, Me.TxtPrimaryCustomer.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboTax.Controls, Me.TxtPrimaryCustomer.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryCustomer.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryCustomer.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gbo_EFaktur.Controls, Me.TxtPrimaryCustomer.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TableLayoutPanel1.Controls, Me.TxtPrimaryCustomer.BackColor, False)

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.gboGrouping.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.gboTax.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.gbo_EFaktur.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.TableLayoutPanel1.Controls, True, True, True, True, False, False, False, False)
            Me.DtpJoinDate.Enabled = True
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.dgvAddress.ReadOnly = False
            Me.dgvBank.ReadOnly = False
            Me.dgvCP.ReadOnly = False
            Me.dgvAddress.AllowUserToAddRows = True
            Me.dgvBank.AllowUserToAddRows = True
            Me.dgvCP.AllowUserToAddRows = True

            '//TR : Tambahan baru Nik/Pas
            Me.cboNIKPASS.Enabled = True
            If Me.cboNIKPASS.SelectedIndex = -1 Then
                Me.cboNIKPASS.SelectedIndex = 0
            End If

            'Comment: Perubahan NPWP 16 Digit
            Me.LlbEditNPWPLama.Enabled = True
            Me.CboTipeNPWP.Enabled = True

            'tentukan nilai default untuk baris terakhir
            With Me.dgvAddress.Rows(Me.dgvAddress.Rows.Count - 1)
                .Cells(1).Value = "-"
            End With
            Me.ObjCekData.SetNumber(Me.dgvAddress)
            'tentukan nilai default untuk baris terakhir
            With Me.dgvBank.Rows(Me.dgvBank.Rows.Count - 1)
                .Cells(1).Value = "-"
                .Cells(3).Value = "IDR"
                .Cells(Me.ColPrimaryCurrency.Index).Value = Me.ObjCurrency.GetPrimary("IDR")
            End With
            Me.ObjCekData.SetNumber(Me.dgvBank)
            'tentukan nilai default untuk baris terakhir
            With Me.dgvCP.Rows(Me.dgvCP.Rows.Count - 1)
                .Cells(1).Value = "-"
            End With
            Me.ObjCekData.SetNumber(Me.dgvCP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.gboTax.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.gbo_EFaktur.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.TableLayoutPanel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.gboGrouping.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.DtpJoinDate.Enabled = False
            Me.CancelButton = Me.btnExit
            Me.dgvAddress.ReadOnly = True
            Me.dgvBank.ReadOnly = True
            Me.dgvCP.ReadOnly = True
            Me.dgvAddress.AllowUserToAddRows = False
            Me.dgvBank.AllowUserToAddRows = False
            Me.dgvCP.AllowUserToAddRows = False

            '//TR : Tambahan baru Nik/Pas
            Me.cboNIKPASS.Enabled = False

            'Comment: Perubahan NPWP 16 Digit
            Me.LlbEditNPWPLama.Enabled = False
            Me.CboTipeNPWP.Enabled = False

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboTax.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gbo_EFaktur.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TableLayoutPanel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboGrouping.Controls, Me.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.gboGrouping.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.gboTax.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.gbo_EFaktur.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
            Me.TxtPrimaryCustRef.Text = "-1"
            Me.TxtPrimaryCustomer.Text = "-1"
            Me.TxtPrimaryCustomerGroup.Text = "-1"
            Me.TxtPrimaryCustomerSegmentation.Text = "-1"
            Me.TxtPrimarySalesPerson.Text = "-1"
            Me.TxtPrimaryCustomerType.Text = "-1"
            Me.TxtPrimarySalesArea.Text = "-1"
            Me.txtpaymentterm.Text = "0"
            Me.TxtCreditLimit.Text = "0"
            '2025 Core Tax Tahap 2.1
            If Today.Date >= Me.ObjSettingNPWP.Tglberlaku Then
                Me.TxtNPWP.Text = "0000 0000 0000 0000"
            Else
                Me.TxtNPWP.Text = "00.000.000.0-000.000"
            End If
            Me.TxtTaxCode.Text = "010"

            Me.Txt_E_Faktur_Jalan.Text = "-"
            Me.Txt_E_Faktur_Blok.Text = "-"
            Me.Txt_E_Faktur_No.Text = "-"
            Me.Txt_E_Faktur_RT.Text = "0"
            Me.Txt_E_Faktur_RW.Text = "0"
            Me.Txt_E_Faktur_Kecamatan.Text = "-"
            Me.Txt_E_Faktur_Kelurahan.Text = "-"
            Me.Txt_E_Faktur_Propinsi.Text = "-"
            Me.Txt_E_Faktur_Kabupaten.Text = "-"
            Me.Txt_E_Faktur_KodePos.Text = "00000"
            Me.Txt_E_Faktur_NoTelepon.Text = "-"

            Me.dgvAddress.Rows.Clear()
            Me.dgvBank.Rows.Clear()
            Me.dgvCP.Rows.Clear()

            '//TR : Tambahan baru Nik/Pas
            Me.txtNIKPASS.Text = "-"
            Me.cboNIKPASS.SelectedIndex = -1

            'Comment: NPWP no 16 Digit
            Me.CboTipeNPWP.SelectedIndex = 0
            Me.TxtNPWPLama.Text = Me.TxtNPWP.Text

            'Update PPN 12%
            Me.TxtTKUID.Text = "000000"
            Me.TxtEmail.Text = ""

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
            KoleksiNamaData.Add("customer code")
            KoleksiTextBox.Add(Me.TxtCustomerCode)
            KoleksiNamaData.Add("customer name")
            KoleksiTextBox.Add(Me.TxtCustomerName)
            KoleksiNamaData.Add("customer group code")
            KoleksiTextBox.Add(Me.TxtCustomerGroupCode)
            KoleksiNamaData.Add("sales area code")
            KoleksiTextBox.Add(Me.TxtSalesAreaCode)
            KoleksiNamaData.Add("customer segmentation code")
            KoleksiTextBox.Add(Me.TxtCustomerSegmentation)
            KoleksiNamaData.Add("customer type code")
            KoleksiTextBox.Add(Me.TxtCustomerType)
            'Perubahan PPN 12%
            KoleksiNamaData.Add("TKU ID")
            KoleksiTextBox.Add(Me.TxtTKUID)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            Me.TxtSalesPerson.Text = Me.TxtSalesPerson.Text.Trim
            If Me.TxtSalesPerson.Text = "" Then
                If MsgBox("No default sales person for this customer?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Me.TxtSalesPerson.Focus()
                    Return False
                End If
            End If

            'berkaitan dengan data e-faktur yang tidak lengkap
            Me.Txt_E_Faktur_Blok.Text = Me.Txt_E_Faktur_Blok.Text.Trim
            Me.Txt_E_Faktur_Jalan.Text = Me.Txt_E_Faktur_Jalan.Text.Trim
            Me.Txt_E_Faktur_Kelurahan.Text = Me.Txt_E_Faktur_Kelurahan.Text.Trim
            Me.Txt_E_Faktur_KodePos.Text = Me.Txt_E_Faktur_KodePos.Text.Trim
            Me.Txt_E_Faktur_No.Text = Me.Txt_E_Faktur_No.Text.Trim
            Me.Txt_E_Faktur_NoTelepon.Text = Me.Txt_E_Faktur_NoTelepon.Text.Trim
            Me.Txt_E_Faktur_Propinsi.Text = Me.Txt_E_Faktur_Propinsi.Text.Trim
            Me.Txt_E_Faktur_RT.Text = Me.Txt_E_Faktur_RT.Text.Trim
            Me.Txt_E_Faktur_RW.Text = Me.Txt_E_Faktur_RW.Text.Trim
            Me.Txt_E_Faktur_Kabupaten.Text = Me.Txt_E_Faktur_Kabupaten.Text.Trim
            Me.Txt_E_Faktur_Kecamatan.Text = Me.Txt_E_Faktur_Kecamatan.Text.Trim

            If Me.Txt_E_Faktur_Blok.Text = "" Or _
            Me.Txt_E_Faktur_Jalan.Text = "" Or _
            Me.Txt_E_Faktur_Kelurahan.Text = "" Or _
            Me.Txt_E_Faktur_KodePos.Text = "" Or _
            Me.Txt_E_Faktur_No.Text = "" Or _
            Me.Txt_E_Faktur_NoTelepon.Text = "" Or _
            Me.Txt_E_Faktur_Propinsi.Text = "" Or _
            Me.Txt_E_Faktur_RT.Text = "" Or _
            Me.Txt_E_Faktur_RW.Text = "" Or _
            Me.Txt_E_Faktur_Kecamatan.Text = "" Or _
            Me.Txt_E_Faktur_Kabupaten.Text = "" Then
                Me.TabControl1.SelectedTab = Me.TabPage4
                If MsgBox("Incomplete e-Faktur data, continue save?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Return False
                End If
            End If

            'data salah bila:
            '1.tidak ada data address
            If Me.dgvAddress.RowCount = 1 Then
                MsgBox("Please input at least one address data.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 0
                Return False
            End If

            '2.cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.objCustomer.ApakahSudahAda(Me.TxtCustomerCode.Text) = True Then
                    MsgBox("Duplication in Customer code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtCustomerCode.Focus()
                    Return False
                End If
            Else
                Dim CustomerGroupLama As String = Me.objCustomer.GetCustomerCode(Me.TxtPrimaryCustomer.Text)
                If CustomerGroupLama <> Me.TxtCustomerCode.Text Then
                    If Me.objCustomer.ApakahSudahAda(Me.TxtCustomerCode.Text) = True Then
                        MsgBox("Duplication in Customer code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtCustomerCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objCustomer.GetTimeStamp(Me.TxtPrimaryCustomer.Text)
                If Me.objCustomer.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objCustomer.CariDetail(Me.objCustomer.Primary, Me.dgvAddress, Me.dgvCP, Me.dgvBank)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            '3. creadit limit/payment terms negatif
            If Me.ObjCekData.CekNonNegatif(Me.TxtCreditLimit.Text) = False Then
                MsgBox("Credit limit must be a non negatif number.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtCreditLimit.Focus()
                Return False
            End If
            If Me.ObjCekData.CekNonNegatif(Me.txtpaymentterm.Text) = False Then
                MsgBox("Payment term must be a non negatif number.", MsgBoxStyle.Information, _
                "Information")
                Me.txtpaymentterm.Focus()
                Return False
            End If

            Dim row As DataGridViewRow
            Dim cell As DataGridViewCell
            '4. untuk item address:
            '   a. untuk baris yang bukan baris terakhir, tidak boleh sampai data line 1-5 kosong
            For Each row In Me.dgvAddress.Rows
                With row
                    If row.Index <> Me.dgvAddress.Rows.Count - 1 Then
                        For Each cell In row.Cells
                            If IsNothing(cell.Value) = True Then
                                cell.Value = ""
                            End If
                        Next
                        'tambahan sales area
                        If .Cells(Me.Col_SalesArea.Index).Value = "" Then
                            Me.TabControl1.SelectedIndex = 0
                            MsgBox("Wrong address data, please fill in sales area data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvAddress.ClearSelection()
                            .Selected = True
                            Return False
                        End If
                        If .Cells(Me.Col_Line1.Index).Value = "" And .Cells(Me.Col_Line2.Index).Value = "" And _
                           .Cells(Me.Col_Line3.Index).Value = "" And .Cells(Me.Col_Line4.Index).Value = "" And _
                           .Cells(Me.Col_Line5.Index).Value = "" Then
                            Me.TabControl1.SelectedIndex = 0
                            MsgBox("Wrong address data, please fill in address data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvAddress.ClearSelection()
                            .Selected = True
                            Return False
                        End If
                        '-----
                    End If
                End With
            Next

            '5. untuk item bank tidak boleh ada nama, currency, dan account no yang kosong
            For Each row In Me.dgvBank.Rows
                With row
                    If row.Index <> Me.dgvBank.Rows.Count - 1 Then
                        For Each cell In row.Cells
                            If IsNothing(cell.Value) = True Then
                                cell.Value = ""
                            End If
                        Next
                        If .Cells(1).Value = "" Then
                            MsgBox("Wrong bank's name, please fill in bank's name data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvBank.ClearSelection()
                            Me.TabControl1.SelectedIndex = 2
                            .Selected = True
                            Return False
                        End If
                        If .Cells(2).Value = "" Then
                            MsgBox("Wrong account no, please fill in account no data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvBank.ClearSelection()
                            Me.TabControl1.SelectedIndex = 2
                            .Selected = True
                            Return False
                        End If
                        If .Cells(3).Value = "" Then
                            MsgBox("Wrong currency code, please fill in currency code data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvBank.ClearSelection()
                            Me.TabControl1.SelectedIndex = 2
                            .Selected = True
                            Return False
                        End If
                    End If
                End With
            Next

            '6. untuk item contact person, tidak boleh ada nama contact person yang kosong
            Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
            For Each row In Me.dgvCP.Rows
                With row
                    If row.Index <> Me.dgvCP.Rows.Count - 1 Then
                        For Each cell In row.Cells
                            If IsNothing(cell.Value) = True Then
                                cell.Value = ""
                            End If
                        Next
                        If .Cells(1).Value = "" Then
                            MsgBox("Wrong contact person's name, please fill in contact person's name data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvCP.ClearSelection()
                            Me.TabControl1.SelectedIndex = 1
                            .Selected = True
                            Return False
                        End If

                        'utnuk setiap item contact person, pastikan email yang diberikan tepat
                        If .Cells(Me.ColPersonSendCL.Index).Value = Boolean.TrueString Then
                            .Cells(Me.ColPersonEmail.Index).Value = .Cells(Me.ColPersonEmail.Index).Value.ToString.Trim
                            If Expression.IsMatch(.Cells(Me.ColPersonEmail.Index).Value.ToString) = False Then
                                MsgBox("The email address is not valid for the following contact person: " + .Cells(Me.ColPersonName.Index).Value.ToString + _
                                ". Please check again.", MsgBoxStyle.Information, "Information")
                                Me.TabControl1.SelectedIndex = 1
                                Me.dgvCP.ClearSelection()
                                row.Selected = True

                                Return False
                            End If

                        End If
                    End If
                End With
            Next
            'Update PPN 12%
            If ObjCD.CekNumerik(Me.TxtTKUID.Text) = False Then
                MsgBox("Must input number format", MsgBoxStyle.Information, "Information")
                Me.TxtTKUID.Focus()
                Return False
            End If
            If Me.TxtEmail.Text <> "" Then
                Me.TxtEmail.Text = Me.TxtEmail.Text.Trim
                If Expression.IsMatch(Me.TxtEmail.Text) = False Then
                    MsgBox("The email address is not valid. Please check again.", MsgBoxStyle.Information, "Information")
                    Me.TxtEmail.Focus()
                    Return False
                End If
                Expression = Nothing
            End If


            '//TR : Tambahan baru Nik/Pas
            Me.txtNIKPASS.Text = Me.txtNIKPASS.Text.Replace("%", "").Replace("`", "").Replace("*", "")

            'Comment: Perubahan NPWP 16 Digit
            If Me.TxtNPWPLama.Text = "" Or Me.TxtNPWPLama.Text = "-" Then
                Me.TxtNPWPLama.Text = Me.TxtNPWP.Text
            End If
            If Today.Date >= Me.ObjSettingNPWP.Tglberlaku Then
                '2025 Core Tax Tahap 2.1
                If Me.ObjSettingNPWP.ApakahDataSudahBenar_MasterCustomer_New(Me.TxtNPWP.Text, Me.txtNIKPASS.Text, Me.TxtNPWPLama.Text, _
                   Me.CboTipeNPWP, Me.cboNIKPASS, True) = False Then
                    Return False
                End If
            Else
                Me.TxtNPWP.Text = Me.TxtNPWP.Text.Trim
                If Me.objCustomer.ApakahFormatNPWPBenar(Me.TxtNPWP.Text, True) = False Then
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
            With Me.objCustomer
                .PrimaryCustomerGroup = Me.TxtPrimaryCustomerGroup.Text
                .PrimaryCustomerSegmentation = Me.TxtPrimaryCustomerSegmentation.Text
                .PrimaryCustomertype = Me.TxtPrimaryCustomerType.Text
                .PrimarysSalesArea = Me.TxtPrimarySalesArea.Text
                .CustomerCode = Me.TxtCustomerCode.Text
                .CustomerName = Me.TxtCustomerName.Text
                .city = Me.TxtCity.Text
                .country = Me.TxtCountry.Text
                .CreditLimit = Me.TxtCreditLimit.Text
                .DefaultTaxCode = Me.TxtTaxCode.Text
                '2025 Core Tax Tahap 2.1, di replace NPWP tanpa spasi
                If Today.Date >= Me.ObjSettingNPWP.Tglberlaku Then
                    .NPWP = Me.TxtNPWP.Text.ToString.Replace(" ", "")
                Else
                    .NPWP = Me.TxtNPWP.Text
                End If
                .NPWP_Name = Me.TxtNPWP_Name.Text
                .PaymentTerm = Me.txtpaymentterm.Text
                .TaxAddressLine1 = Me.TxtAddressPajak1.Text
                .TaxAddressLine2 = Me.TxtAddressPajak2.Text
                .ZipCode = Me.TxtZipCode.Text
                .Remarks = Me.TxtRemarks.Text
                .PrimarySalesPerson = Me.TxtPrimarySalesPerson.Text
                .PrimaryCustomerReference = Me.TxtPrimaryCustRef.Text
                .JoinDate = Me.DtpJoinDate.Value
                .EFak_Blok = Me.Txt_E_Faktur_Blok.Text
                .EFak_NamaJalan = Me.Txt_E_Faktur_Jalan.Text
                .EFak_Kelurahan = Me.Txt_E_Faktur_Kelurahan.Text
                .EFak_KodePos = Me.Txt_E_Faktur_KodePos.Text
                .EFak_No = Me.Txt_E_Faktur_No.Text
                .EFak_Telepon = Me.Txt_E_Faktur_NoTelepon.Text
                .EFak_Propinsi = Me.Txt_E_Faktur_Propinsi.Text
                .EFak_RT = Me.Txt_E_Faktur_RT.Text
                .EFak_RW = Me.Txt_E_Faktur_RW.Text
                .EFak_Kabupaten = Me.Txt_E_Faktur_Kabupaten.Text
                .EFak_Kecamatan = Me.Txt_E_Faktur_Kecamatan.Text

                '//TR : Tambahan baru Nik/Pas
                .OptionNIKPASS = OptionNIKPASS
                .NoNIKPASS = Me.txtNIKPASS.Text

                'Comment: NPWP 16 Digit
                .NPWPLama = Me.TxtNPWPLama.Text
                .NPWPType = Me.CboTipeNPWP.SelectedIndex

                'Update PPN 12%
                .TKU_ID = Me.TxtTKUID.Text
                .Email = Me.TxtEmail.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objCustomer
                Me.TxtPrimaryCustomerGroup.Text = .PrimaryCustomerGroup
                Me.TxtCustomerGroupCode.Text = Me.ObjCustGroup.GetGroupCode(Me.TxtPrimaryCustomerGroup.Text)
                Me.TxtCustomerGroupDesc.Text = Me.ObjCustGroup.GetGroupDesc(Me.TxtPrimaryCustomerGroup.Text)
                Me.TxtPrimaryCustomerSegmentation.Text = .PrimaryCustomerSegmentation
                Me.TxtCustomerSegmentation.Text = Me.ObjCustSegmentation.GetSegmentationCode(Me.TxtPrimaryCustomerSegmentation.Text)
                Me.TxtCustomerSegmentationName.Text = Me.ObjCustSegmentation.GetDesc(Me.TxtPrimaryCustomerSegmentation.Text)
                Me.TxtPrimaryCustomerType.Text = .PrimaryCustomertype
                Me.TxtCustomerType.Text = Me.ObjCustType.GetTypeCode(Me.TxtPrimaryCustomerType.Text)
                Me.TxtCustomerTypeDesc.Text = Me.ObjCustType.GetTypeDesc(Me.TxtPrimaryCustomerType.Text)
                Me.TxtPrimarySalesArea.Text = .PrimarysSalesArea
                Me.TxtSalesAreaCode.Text = Me.ObjCustSalesArea.GetSalesAreaCode(Me.TxtPrimarySalesArea.Text)
                Me.TxtSalesAreaDesc.Text = Me.ObjCustSalesArea.GetSalesAreaDesc(Me.TxtPrimarySalesArea.Text)
                Me.TxtPrimaryCustomer.Text = .Primary
                Me.TxtCustomerCode.Text = .CustomerCode
                Me.TxtCustomerName.Text = .CustomerName
                Me.TxtCity.Text = .city
                Me.TxtCountry.Text = .country
                Me.TxtCreditLimit.Text = .CreditLimit
                Dim e As New Windows.Forms.KeyEventArgs(Keys.A)
                Me.ObjCekData.Number_KeyUp(Me.TxtCreditLimit, Me, e)
                e = Nothing
                Me.TxtTaxCode.Text = .DefaultTaxCode
                '2025 Core Tax Tahap 2.1, di replace NPWP tanpa spasi, titik dan -
                If Today.Date >= Me.ObjSettingNPWP.Tglberlaku And .NPWP.Replace(" ", "").Replace("-", "").Replace(".", "").Length = 16 Then
                    Me.TxtNPWP.Text = Me.ObjSettingNPWP.format16DigitSpasi(.NPWP) 'diformat spasi tiap 4 digit
                Else
                    Me.TxtNPWP.Text = .NPWP
                End If
                Me.TxtNPWP_Name.Text = .NPWP_Name
                Me.txtpaymentterm.Text = .PaymentTerm
                Me.TxtAddressPajak1.Text = .TaxAddressLine1
                Me.TxtAddressPajak2.Text = .TaxAddressLine2
                Me.TxtZipCode.Text = .ZipCode
                Me.TxtRemarks.Text = .Remarks
                Me.TxtPrimarySalesPerson.Text = .PrimarySalesPerson
                Me.TxtSalesPerson.Text = Me.ObjSalesP.GetCode(.PrimarySalesPerson)
                Me.TxtSalesPersonDesc.Text = Me.ObjSalesP.GetName(.PrimarySalesPerson)
                Me.TxtPrimaryCustRef.Text = .PrimaryCustomerReference
                Me.TxtCodeRef.Text = Me.ObjRef.GetCode(.PrimaryCustomerReference)
                Me.DtpJoinDate.Value = .JoinDate
                Me.Txt_E_Faktur_Blok.Text = .EFak_Blok
                Me.Txt_E_Faktur_Jalan.Text = .EFak_NamaJalan
                Me.Txt_E_Faktur_Kelurahan.Text = .EFak_Kelurahan
                Me.Txt_E_Faktur_KodePos.Text = .EFak_KodePos
                Me.Txt_E_Faktur_No.Text = .EFak_No
                Me.Txt_E_Faktur_NoTelepon.Text = .EFak_Telepon
                Me.Txt_E_Faktur_Propinsi.Text = .EFak_Propinsi
                Me.Txt_E_Faktur_RT.Text = .EFak_RT
                Me.Txt_E_Faktur_RW.Text = .EFak_RW
                Me.Txt_E_Faktur_Kabupaten.Text = .EFak_Kabupaten
                Me.Txt_E_Faktur_Kecamatan.Text = .EFak_Kecamatan

                '//TR : Tambahan baru Nik/Pas
                Me.cboNIKPASS.SelectedIndex = .OptionNIKPASS
                Me.txtNIKPASS.Text = .NoNIKPASS

                'Comment: NPWP No 16 Digit
                Me.TxtNPWPLama.Text = .NPWPLama
                Me.CboTipeNPWP.SelectedIndex = .NPWPType

                'Update PPN 12%
                Me.TxtTKUID.Text = .TKU_ID
                Me.TxtEmail.Text = .Email

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCustomer_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            Me.ObjCekData = Nothing
            Me.ObjCustGroup = Nothing
            Me.objCustomer = Nothing
            Me.ObjCustSalesArea = Nothing
            Me.ObjCustSegmentation = Nothing
            Me.ObjCustType = Nothing
            Me.ObjCurrency = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvAddress.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvBank.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvCP.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.FromCekEmailCP Then
                Me.ShowDetail_CekEmailCP(vPrimaryCustCekEmailCP)
            End If
            'Me.TxtNPWP.Width = Me.SplitContainer1.Panel1.Width - Me.TxtNPWP.Left - 5
            'Me.TxtTaxCode.Width = Me.TxtNPWP.Width
            'Me.TxtAddressPajak1.Width = Me.TxtNPWP.Width
            'Me.TxtAddressPajak2.Width = Me.TxtNPWP.Width

            'Me.TableLayoutPanel1.ColumnStyles(3).Width = 150
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
                        If Me.objCustomer.SimpanBaru(Me.UserCode, Me.dgvCP, Me.dgvAddress, Me.dgvBank, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                            If Me.FromCekEmailCP = True Then
                                Me.Close()
                            End If
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objCustomer.SimpanPerubahan(Me.UserCode, Me.dgvCP, Me.dgvAddress, Me.dgvBank, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                            If Me.FromCekEmailCP = True Then
                                Me.Close()
                            End If
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
            If Not Me.TxtCustomerCode.Text = "" And Me.TxtCustomerName.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.objCustomer.CariDetail(Me.objCustomer.Primary, Me.dgvAddress, Me.dgvCP, Me.dgvBank)
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
            If Me.FromCekEmailCP = True Then
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            'If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
            '    MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
            '    MsgBoxStyle.Information, "Information")
            '    Exit Sub
            'End If

            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TxtCustomerCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Public Sub ShowDetail_CekEmailCP(ByVal vPrimaryCust As Double)
        Try
            If Me.FromCekEmailCP = False Then
                Exit Sub
            End If

            If Not vPrimaryCust = -1 Then
                If Me.objCustomer.CariDetail(vPrimaryCust, Me.dgvAddress, _
                Me.dgvCP, Me.dgvBank) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()

                Me.baru = False
                Me.AktifkanDetail()
                Me.dgvCP.Focus()
                Me.TabControl1.SelectedIndex = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariMasterCustomer
        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            If Not ij.PrimaryCustomer = -1 Then
                If Me.objCustomer.CariDetail(ij.PrimaryCustomer, Me.dgvAddress, _
                Me.dgvCP, Me.dgvBank) = True Then
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
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If Me.TxtCustomerCode.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.TxtCustomerCode.Text <> "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.TxtCustomerCode.Focus()
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtCustomerCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtCustomerCode.Text <> "" Then
                If Me.objCustomer.ApakahBolehHapus(Me.TxtPrimaryCustomer.Text) = False Then
                    'MsgBox("You cannot delete this data, there are other people using this data.", _
                    '                    MsgBoxStyle.Information, "Information")
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objCustomer.Hapus(Me.TxtPrimaryCustomer.Text, Me.UserCode) = True Then
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

    Private Sub FrmMasterCustomerGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtCustomerCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objCustomer.SimpanBaru(Me.UserCode, Me.dgvCP, Me.dgvAddress, Me.dgvBank, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                                If Me.FromCekEmailCP = True Then
                                    Me.Close()
                                End If
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objCustomer.SimpanPerubahan(Me.UserCode, Me.dgvCP, Me.dgvAddress, Me.dgvBank, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                                If Me.FromCekEmailCP = True Then
                                    Me.Close()
                                End If
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                Else
                    If Me.FromCekEmailCP = True Then
                        Me.Close()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCreditLimit_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCreditLimit.Enter
        Try
            Me.buffer = Me.TxtCreditLimit.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCreditLimit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCreditLimit.Leave
        Try
            If ModCekData.CekNonNegatif(Me.TxtCreditLimit.Text) = False Then
                MsgBox("Creadit limit value must be a number.", MsgBoxStyle.Information, "Infomation")
                Me.TxtCreditLimit.Text = Me.buffer
            Else
                Me.TxtCreditLimit.Text = Me.ObjCD.DigitGrouping(Me.TxtCreditLimit.Text, Me.ObjCurrency.GetDecimal_DefaultCurrency)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPaymentTerm_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtpaymentterm.Enter
        Try
            Me.buffer = Me.txtpaymentterm.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPaymentTerm_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtpaymentterm.Leave
        Try
            If ModCekData.CekNonNegatif(Me.txtpaymentterm.Text) = False Then
                MsgBox("Payment term value must be a number.", MsgBoxStyle.Information, "Infomation")
                Me.txtpaymentterm.Text = Me.buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerGroupCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCustomerGroupCode.KeyDown
        Dim i As New FrmCariMasterCustomerGroup

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjCustGroup.ApakahSudahAda(Me.TxtCustomerGroupCode.Text) = True Then
                    Me.TxtPrimaryCustomerGroup.Text = Me.ObjCustGroup.GetPrimary(Me.TxtCustomerGroupCode.Text)
                    Me.TxtCustomerGroupDesc.Text = Me.ObjCustGroup.GetGroupDesc(Me.TxtPrimaryCustomerGroup.Text)
                    Me.TxtSalesAreaCode.Focus()
                Else
                    If Me.TxtCustomerGroupCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCustomerGroupCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomerGroup > 0 Then
                        Me.TxtPrimaryCustomerGroup.Text = i.PrimaryCustomerGroup
                        Me.TxtCustomerGroupCode.Text = Me.ObjCustGroup.GetGroupCode(Me.TxtPrimaryCustomerGroup.Text)
                        Me.TxtCustomerGroupDesc.Text = Me.ObjCustGroup.GetGroupDesc(Me.TxtPrimaryCustomerGroup.Text)
                        Me.TxtSalesAreaCode.Focus()
                    Else
                        Me.TxtCustomerGroupCode.Clear()
                        Me.TxtCustomerGroupDesc.Clear()
                        Me.TxtPrimaryCustomerGroup.Text = "-1"
                        Me.TxtCustomerGroupCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCustomerGroupCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerGroupCode.Leave
        Dim i As New FrmCariMasterCustomerGroup

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtCustomerGroupCode.Text <> "" Then
                If Me.ObjCustGroup.ApakahSudahAda(Me.TxtCustomerGroupCode.Text) = True Then
                    Me.TxtPrimaryCustomerGroup.Text = Me.ObjCustGroup.GetPrimary(Me.TxtCustomerGroupCode.Text)
                    Me.TxtCustomerGroupDesc.Text = Me.ObjCustGroup.GetGroupDesc(Me.TxtPrimaryCustomerGroup.Text)
                Else
                    If Me.TxtCustomerGroupCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCustomerGroupCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomerGroup > 0 Then
                        Me.TxtPrimaryCustomerGroup.Text = i.PrimaryCustomerGroup
                        Me.TxtCustomerGroupCode.Text = Me.ObjCustGroup.GetGroupCode(Me.TxtPrimaryCustomerGroup.Text)
                        Me.TxtCustomerGroupDesc.Text = Me.ObjCustGroup.GetGroupDesc(Me.TxtPrimaryCustomerGroup.Text)
                        Me.TxtSalesAreaCode.Focus()
                    Else
                        Me.TxtCustomerGroupCode.Clear()
                        Me.TxtCustomerGroupDesc.Clear()
                        Me.TxtPrimaryCustomerGroup.Text = "-1"
                        Me.TxtCustomerGroupCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCustomerGroup.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCustomerGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCustomerGroup.Click
        Dim i As New FrmCariMasterCustomerGroup
        Dim obj As New clsMasterCustomerGroup

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtCustomerGroupCode.Text <> "" Then
                i.txtCari.Text = Me.TxtCustomerGroupCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomerGroup > 0 Then
                Me.TxtPrimaryCustomerGroup.Text = i.PrimaryCustomerGroup
                Me.TxtCustomerGroupCode.Text = obj.GetGroupCode(Me.TxtPrimaryCustomerGroup.Text)
                Me.TxtCustomerGroupDesc.Text = Me.ObjCustGroup.GetGroupDesc(Me.TxtPrimaryCustomerGroup.Text)
                Me.TxtSalesAreaCode.Focus()
            Else
                Me.TxtCustomerGroupCode.Clear()
                Me.TxtCustomerGroupDesc.Clear()
                Me.TxtCustomerGroupCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub txtSalesArea_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSalesAreaCode.KeyDown
        Dim i As New FrmCariMasterSalesArea

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjCustSalesArea.ApakahSudahAda(Me.TxtSalesAreaCode.Text) = True Then
                    Me.TxtPrimarySalesArea.Text = Me.ObjCustSalesArea.GetPrimary(Me.TxtSalesAreaCode.Text)
                    Me.TxtSalesAreaDesc.Text = Me.ObjCustSalesArea.GetSalesAreaDesc(Me.TxtPrimarySalesArea.Text)
                    Me.TxtCustomerSegmentation.Focus()
                Else
                    If Me.TxtSalesAreaCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtSalesAreaCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomerSalesArea > 0 Then
                        Me.TxtPrimarySalesArea.Text = i.PrimaryCustomerSalesArea
                        Me.TxtSalesAreaCode.Text = Me.ObjCustSalesArea.GetSalesAreaCode(Me.TxtPrimarySalesArea.Text)
                        Me.TxtSalesAreaDesc.Text = Me.ObjCustSalesArea.GetSalesAreaDesc(Me.TxtPrimarySalesArea.Text)
                        Me.TxtCustomerSegmentation.Focus()
                    Else
                        Me.TxtSalesAreaCode.Clear()
                        Me.TxtSalesAreaDesc.Clear()
                        Me.TxtPrimarySalesArea.Text = "-1"
                        Me.TxtSalesAreaCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSalesArea_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesAreaCode.Leave
        Dim i As New FrmCariMasterSalesArea

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtSalesAreaCode.Text <> "" Then
                If Me.ObjCustSalesArea.ApakahSudahAda(Me.TxtSalesAreaCode.Text) = True Then
                    Me.TxtPrimarySalesArea.Text = Me.ObjCustSalesArea.GetPrimary(Me.TxtSalesAreaCode.Text)
                    Me.TxtSalesAreaDesc.Text = Me.ObjCustSalesArea.GetSalesAreaDesc(Me.TxtPrimarySalesArea.Text)
                Else
                    If Me.TxtSalesAreaCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtSalesAreaCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomerSalesArea > 0 Then
                        Me.TxtPrimarySalesArea.Text = i.PrimaryCustomerSalesArea
                        Me.TxtSalesAreaCode.Text = Me.ObjCustSalesArea.GetSalesAreaCode(Me.TxtPrimarySalesArea.Text)
                        Me.TxtSalesAreaDesc.Text = Me.ObjCustSalesArea.GetSalesAreaDesc(Me.TxtPrimarySalesArea.Text)
                        Me.TxtCustomerSegmentation.Focus()
                    Else
                        Me.TxtSalesAreaCode.Clear()
                        Me.TxtSalesAreaDesc.Clear()
                        Me.TxtPrimarySalesArea.Text = "-1"
                        Me.TxtSalesAreaCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimarySalesArea.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindSalesArea_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindSalesArea.Click
        Dim i As New FrmCariMasterSalesArea

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtSalesAreaCode.Text <> "" Then
                i.txtCari.Text = Me.TxtSalesAreaCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomerSalesArea > 0 Then
                Me.TxtPrimarySalesArea.Text = i.PrimaryCustomerSalesArea
                Me.TxtSalesAreaCode.Text = Me.ObjCustSalesArea.GetSalesAreaCode(Me.TxtPrimarySalesArea.Text)
                Me.TxtSalesAreaDesc.Text = Me.ObjCustSalesArea.GetSalesAreaDesc(Me.TxtPrimarySalesArea.Text)
                Me.TxtCustomerSegmentation.Focus()
            Else
                Me.TxtSalesAreaCode.Clear()
                Me.TxtSalesAreaDesc.Clear()
                Me.TxtSalesAreaCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCustomerSegmentation_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCustomerSegmentation.KeyDown
        Dim i As New FrmCariMasterCustomerSegmentation

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjCustSegmentation.ApakahSudahAda(Me.TxtCustomerSegmentation.Text) = True Then
                    Me.TxtPrimaryCustomerSegmentation.Text = Me.ObjCustSegmentation.GetPrimary(Me.TxtCustomerSegmentation.Text)
                    Me.TxtCustomerSegmentationName.Text = Me.ObjCustSegmentation.GetDesc(Me.TxtPrimaryCustomerSegmentation.Text)
                    Me.TxtCustomerType.Focus()
                Else
                    If Me.TxtCustomerSegmentation.Text <> "" Then
                        i.txtCari.Text = Me.TxtCustomerSegmentation.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomerMarketSegmentation > 0 Then
                        Me.TxtPrimaryCustomerSegmentation.Text = i.PrimaryCustomerMarketSegmentation
                        Me.TxtCustomerSegmentation.Text = Me.ObjCustSegmentation.GetSegmentationCode(Me.TxtPrimaryCustomerSegmentation.Text)
                        Me.TxtCustomerSegmentationName.Text = Me.ObjCustSegmentation.GetDesc(Me.TxtPrimaryCustomerSegmentation.Text)
                        Me.TxtCustomerType.Focus()
                    Else
                        Me.TxtCustomerSegmentation.Clear()
                        Me.TxtCustomerSegmentationName.Clear()
                        Me.TxtPrimaryCustomerSegmentation.Text = "-1"
                        Me.TxtCustomerSegmentation.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCustomerSegmentation_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerSegmentation.Leave
        Dim i As New FrmCariMasterCustomerSegmentation

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtCustomerSegmentation.Text <> "" Then
                If Me.ObjCustSegmentation.ApakahSudahAda(Me.TxtCustomerSegmentation.Text) = True Then
                    Me.TxtPrimaryCustomerSegmentation.Text = Me.ObjCustSegmentation.GetPrimary(Me.TxtCustomerSegmentation.Text)
                    Me.TxtCustomerSegmentationName.Text = Me.ObjCustSegmentation.GetDesc(Me.TxtPrimaryCustomerSegmentation.Text)
                Else
                    If Me.TxtCustomerSegmentation.Text <> "" Then
                        i.txtCari.Text = Me.TxtCustomerSegmentation.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomerMarketSegmentation > 0 Then
                        Me.TxtPrimaryCustomerSegmentation.Text = i.PrimaryCustomerMarketSegmentation
                        Me.TxtCustomerSegmentation.Text = Me.ObjCustSegmentation.GetSegmentationCode(Me.TxtPrimaryCustomerSegmentation.Text)
                        Me.TxtCustomerSegmentationName.Text = Me.ObjCustSegmentation.GetDesc(Me.TxtPrimaryCustomerSegmentation.Text)
                        Me.TxtCustomerType.Focus()
                    Else
                        Me.TxtCustomerSegmentation.Clear()
                        Me.TxtCustomerSegmentationName.Clear()
                        Me.TxtPrimaryCustomerSegmentation.Text = "-1"
                        Me.TxtCustomerSegmentation.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCustomerSegmentation.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCustomerSegmentation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustSegmentation.Click
        Dim i As New FrmCariMasterCustomerSegmentation

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtCustomerSegmentation.Text <> "" Then
                i.txtCari.Text = Me.TxtCustomerSegmentation.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomerMarketSegmentation > 0 Then
                Me.TxtPrimaryCustomerSegmentation.Text = i.PrimaryCustomerMarketSegmentation
                Me.TxtCustomerSegmentation.Text = Me.ObjCustSegmentation.GetSegmentationCode(Me.TxtPrimaryCustomerSegmentation.Text)
                Me.TxtCustomerSegmentationName.Text = Me.ObjCustSegmentation.GetDesc(Me.TxtPrimaryCustomerSegmentation.Text)
                Me.TxtCustomerType.Focus()
            Else
                Me.TxtCustomerSegmentation.Clear()
                Me.TxtCustomerSegmentationName.Clear()
                Me.TxtCustomerSegmentation.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCustomerType_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCustomerType.KeyDown
        Dim i As New frmCariMasterCustomerType

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjCustType.ApakahSudahAda(Me.TxtCustomerType.Text) = True Then
                    Me.TxtPrimaryCustomerType.Text = Me.ObjCustType.GetPrimary(Me.TxtCustomerType.Text)
                    Me.TxtCustomerTypeDesc.Text = Me.ObjCustType.GetTypeDesc(Me.TxtPrimaryCustomerType.Text)
                    Me.TabControl1.SelectedIndex = 1
                    Me.TxtNPWP.Focus()
                Else
                    If Me.TxtCustomerType.Text <> "" Then
                        i.txtCari.Text = Me.TxtCustomerType.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomerType > 0 Then
                        Me.TxtPrimaryCustomerType.Text = i.PrimaryCustomerType
                        Me.TxtCustomerType.Text = Me.ObjCustType.GetTypeCode(Me.TxtPrimaryCustomerType.Text)
                        Me.TxtCustomerTypeDesc.Text = Me.ObjCustType.GetTypeDesc(Me.TxtPrimaryCustomerType.Text)
                        Me.TabControl1.SelectedIndex = 1
                        Me.TxtNPWP.Focus()
                    Else
                        Me.TxtCustomerType.Clear()
                        Me.TxtCustomerTypeDesc.Clear()
                        Me.TxtPrimaryCustomerType.Text = "-1"
                        Me.TxtCustomerType.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCustomerType_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerType.Leave
        Dim i As New frmCariMasterCustomerType

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtCustomerType.Text <> "" Then
                If Me.ObjCustType.ApakahSudahAda(Me.TxtCustomerType.Text) = True Then
                    Me.TxtPrimaryCustomerType.Text = Me.ObjCustType.GetPrimary(Me.TxtCustomerType.Text)
                    Me.TxtCustomerTypeDesc.Text = Me.ObjCustType.GetTypeDesc(Me.TxtPrimaryCustomerType.Text)
                Else
                    If Me.TxtCustomerType.Text <> "" Then
                        i.txtCari.Text = Me.TxtCustomerType.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomerType > 0 Then
                        Me.TxtPrimaryCustomerType.Text = i.PrimaryCustomerType
                        Me.TxtCustomerType.Text = Me.ObjCustType.GetTypeCode(Me.TxtPrimaryCustomerType.Text)
                        Me.TxtCustomerTypeDesc.Text = Me.ObjCustType.GetTypeDesc(Me.TxtPrimaryCustomerType.Text)
                        Me.TabControl1.SelectedIndex = 1
                        Me.TxtNPWP.Focus()
                    Else
                        Me.TxtCustomerType.Clear()
                        Me.TxtCustomerTypeDesc.Clear()
                        Me.TxtPrimaryCustomerType.Text = "-1"
                        Me.TxtCustomerType.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCustomerType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCustomertype.Click
        Dim i As New frmCariMasterCustomerType

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtCustomerType.Text <> "" Then
                i.txtCari.Text = Me.TxtCustomerType.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomerType > 0 Then
                Me.TxtPrimaryCustomerType.Text = i.PrimaryCustomerType
                Me.TxtCustomerType.Text = Me.ObjCustType.GetTypeCode(Me.TxtPrimaryCustomerType.Text)
                Me.TxtCustomerTypeDesc.Text = Me.ObjCustType.GetTypeDesc(Me.TxtPrimaryCustomerType.Text)
                Me.TabControl1.SelectedIndex = 1
                Me.TxtNPWP.Focus()
            Else
                Me.TxtCustomerType.Clear()
                Me.TxtCustomerTypeDesc.Clear()
                Me.TxtCustomerType.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
    Private Sub TxtCustomerCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCode.GotFocus
        TxtCustomerCode.BackColor = BackColorAll
    End Sub

    Private Sub TxtCustomerCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCode.LostFocus
        TxtCustomerCode.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtCustomerName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerName.GotFocus
        TxtCustomerName.BackColor = BackColorAll
    End Sub

    Private Sub TxtCustomerName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerName.LostFocus
        TxtCustomerName.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtCity_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCity.GotFocus
        TxtCity.BackColor = BackColorAll
    End Sub

    Private Sub TxtCity_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCity.LostFocus
        TxtCity.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtCountry_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCountry.GotFocus
        TxtCountry.BackColor = BackColorAll
    End Sub

    Private Sub TxtCountry_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCountry.LostFocus
        TxtCountry.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtZipCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtZipCode.GotFocus
        TxtZipCode.BackColor = BackColorAll
    End Sub

    Private Sub TxtZipCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtZipCode.LostFocus
        TxtZipCode.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtCreditLimit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCreditLimit.GotFocus
        TxtCreditLimit.BackColor = BackColorAll
    End Sub

    Private Sub TxtCreditLimit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCreditLimit.LostFocus
        TxtCreditLimit.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub txtPaymentTerm_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtpaymentterm.GotFocus
        txtpaymentterm.BackColor = BackColorAll
    End Sub

    Private Sub txtPaymentTerm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtpaymentterm.LostFocus
        txtpaymentterm.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtCustomerGroupCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerGroupCode.GotFocus
        TxtCustomerGroupCode.BackColor = BackColorAll
    End Sub

    Private Sub TxtCustomerGroupCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerGroupCode.LostFocus
        TxtCustomerGroupCode.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtSalesArea_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesAreaCode.GotFocus
        TxtSalesAreaCode.BackColor = BackColorAll
    End Sub

    Private Sub TxtSalesArea_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesAreaCode.LostFocus
        TxtSalesAreaCode.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtCostumerSegmentation_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerSegmentation.GotFocus
        TxtCustomerSegmentation.BackColor = BackColorAll
    End Sub

    Private Sub TxtCostumerSegmentation_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerSegmentation.LostFocus
        TxtCustomerSegmentation.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtCustomertype_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerType.GotFocus
        TxtCustomerType.BackColor = BackColorAll
    End Sub

    Private Sub TxtCustomertype_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerType.LostFocus
        TxtCustomerType.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtNPWP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWP.GotFocus
        If Me.TxtNPWP.ReadOnly = False Then
            TxtNPWP.BackColor = BackColorAll
        End If
    End Sub

    Private Sub TxtNPWP_Name_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWP_Name.GotFocus
        TxtNPWP_Name.BackColor = BackColorAll
    End Sub

    Private Sub TxtNPWP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWP.LostFocus
        If Me.TxtNPWP.ReadOnly = False Then
            TxtNPWP.BackColor = Me.TxtPrimaryCustomer.BackColor
        End If
    End Sub

    Private Sub TxtNPWP_Name_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWP_Name.Leave
        Try
            Me.TxtNPWP_Name.Text = Me.TxtNPWP_Name.Text.Replace("'", "").Replace("*", "").Replace("%", "").Trim
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNPWP_Name_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWP_Name.LostFocus
        TxtNPWP_Name.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtTaxCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxCode.GotFocus
        TxtTaxCode.BackColor = BackColorAll
    End Sub

    Private Sub TxtTaxCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxCode.LostFocus
        TxtTaxCode.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtAddressPajak1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddressPajak1.GotFocus
        TxtAddressPajak1.BackColor = BackColorAll
    End Sub

    Private Sub TxtAddressPajak1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddressPajak1.LostFocus
        TxtAddressPajak1.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtAddressPajak2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddressPajak2.GotFocus
        TxtAddressPajak2.BackColor = BackColorAll
    End Sub

    Private Sub TxtAddressPajak2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddressPajak2.LostFocus
        TxtAddressPajak2.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtCreditLimit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCreditLimit.KeyDown
        Try
            Me.ObjCekData.Number_KeyDown(Me.TxtCreditLimit, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCreditLimit_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCreditLimit.KeyUp
        Try
            Me.ObjCekData.Number_KeyUp(Me.TxtCreditLimit, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtpaymentterm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtpaymentterm.KeyDown
        Try
            Me.ObjCekData.Number_KeyDown(Me.txtpaymentterm, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtpaymentterm_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtpaymentterm.KeyUp
        Try
            Me.ObjCekData.Number_KeyUp(Me.txtpaymentterm, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAddress_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvAddress.CellBeginEdit
        Try
            If e.ColumnIndex = 0 Or e.ColumnIndex = Me.Col_SalesAreaDesc.Index Then
                e.Cancel = True
            Else
                Me.buffer = Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Function GetIndexSalesArea(ByVal vPrimarySalesArea As String, _
    'ByVal vRowIdxFocus As Integer) As Double
    '    Dim RowDGV As DataGridViewRow
    '    Dim Retval As Double = -1
    '    Try
    '        For ij As Integer = 0 To Me.dgvAddress.Rows.Count - 2
    '            RowDGV = Me.dgvAddress.Rows(ij)
    '            If vRowIdxFocus <> ij Then
    '                If RowDGV.Cells(Me.Col_PrimarySalesArea.Index).Value = vPrimarySalesArea Then
    '                    Retval = ij
    '                    Exit For
    '                End If
    '            End If
    '        Next

    '        Return Retval
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Private Sub dgvAddress_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAddress.CellEndEdit
        Dim index As Integer = -1
        'Dim i As Integer = 0
        Dim PrimarySalesArea As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            If IsNothing(Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = 1 Or e.ColumnIndex = Me.Col_SalesArea.Index Then
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
                ElseIf e.ColumnIndex = Me.Col_SalesArea.Index Then
                    'aul
                    Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ObjCustSalesArea.ApakahSudahAda(Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        ''1. kode akun yang diketik ada --> load data deskripsi + primary akun
                        'PrimarySalesArea = ObjCustSalesArea.GetPrimary(Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value)
                        'index = Me.GetIndexSalesArea(PrimarySalesArea, e.RowIndex)
                        'If index < 0 Then
                        Me.dgvAddress.Item(Me.Col_PrimarySalesArea.Index, e.RowIndex).Value = Me.ObjCustSalesArea.GetPrimary(Me.dgvAddress.Item(Me.Col_SalesArea.Index, e.RowIndex).Value)
                        Me.dgvAddress.Item(Me.Col_SalesArea.Index, e.RowIndex).Value = Me.dgvAddress.Item(Me.Col_SalesArea.Index, e.RowIndex).Value
                        Me.dgvAddress.Item(Me.Col_SalesAreaDesc.Index, e.RowIndex).Value = Me.ObjCustSalesArea.GetSalesAreaDesc(Me.dgvAddress.Item(Me.Col_PrimarySalesArea.Index, e.RowIndex).Value)
                        'Else
                        '    If e.RowIndex <> index Then
                        '        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                        '        Me.dgvAddress.Item(Me.Col_SalesArea.Index, e.RowIndex).Value = ""
                        '        If String.IsNullOrEmpty(Me.buffer) = True Then
                        '            'Me.EraseRowIndex = e.RowIndex
                        '            'Me.BgwUser.RunWorkerAsync()
                        '        Else
                        '            Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        '        End If
                        '        Exit Sub
                        '    End If
                        'End If
                    Else
                        Dim ij As New FrmCariMasterSalesArea
                        ij.TxtCari.Text = Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.backColorAll = Me.BackColorAll
                        ij.ShowDialog()
                        If ij.PrimaryCustomerSalesArea <> -1 Then
                            'index = Me.GetIndexSalesArea(ij.PrimaryCustomerSalesArea, e.RowIndex)
                            'If index < 0 Then
                            Me.dgvAddress.Item(Me.Col_PrimarySalesArea.Index, e.RowIndex).Value = ij.PrimaryCustomerSalesArea
                            Me.dgvAddress.Item(Me.Col_SalesArea.Index, e.RowIndex).Value = Me.ObjCustSalesArea.GetSalesAreaCode(ij.PrimaryCustomerSalesArea)
                            Me.dgvAddress.Item(Me.Col_SalesAreaDesc.Index, e.RowIndex).Value = Me.ObjCustSalesArea.GetSalesAreaDesc(ij.PrimaryCustomerSalesArea)
                            'Else
                            '    If e.RowIndex <> index Then
                            '        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                            '        If String.IsNullOrEmpty(Me.buffer) = True Then
                            '            'Me.EraseRowIndex = e.RowIndex
                            '            'Me.BgwUser.RunWorkerAsync()
                            '        Else
                            '            Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            '        End If
                            '        Exit Sub
                            '    Else
                            '        Me.dgvAddress.Item(Me.Col_PrimarySalesArea.Index, e.RowIndex).Value = ij.PrimaryCustomerSalesArea
                            '        Me.dgvAddress.Item(Me.Col_SalesArea.Index, e.RowIndex).Value = Me.ObjCustSalesArea.GetSalesAreaCode(ij.PrimaryCustomerSalesArea)
                            '        Me.dgvAddress.Item(Me.Col_SalesAreaDesc.Index, e.RowIndex).Value = Me.ObjCustSalesArea.GetSalesAreaDesc(ij.PrimaryCustomerSalesArea)
                            '    End If
                            'End If
                        ElseIf ij.PrimaryCustomerSalesArea = -1 Then
                            If e.ColumnIndex = Me.Col_SalesArea.Index Then
                                If String.IsNullOrEmpty(Me.buffer) = True Then
                                    If Me.dgvAddress.Rows.Count - 1 <> e.RowIndex Then
                                        Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                    End If
                                Else
                                    Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                                'Else
                                '    Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                        End If
                        ij = Nothing
                    End If

                ElseIf e.ColumnIndex = Me.Col_Line1.Index Or e.ColumnIndex = Me.Col_Line2.Index Or _
                e.ColumnIndex = Me.Col_Line3.Index Or e.ColumnIndex = Me.Col_Line4.Index Or _
                e.ColumnIndex = Me.Col_Line5.Index Or e.ColumnIndex = Me.ColPhoneAdd.Index Or _
                e.ColumnIndex = Me.ColRemarksAdd.Index Or e.ColumnIndex = Me.ColFaxAdd.Index Then
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

    Private Sub dgvItemCP_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCP.CellBeginEdit
        Try
            If e.ColumnIndex = 0 Then
                e.Cancel = True
            Else
                Me.buffer = Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemCP_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCP.CellEndEdit
        Try
            If IsNothing(Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = 1 Then
                    Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                ElseIf e.ColumnIndex = Me.ColPersonSendCL.Index Then
                    Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = CBool(Boolean.FalseString)
                Else
                    Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = ""
                End If
            Else
                Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                If e.ColumnIndex = 1 Then
                    'untuk ItemCP nama tidak boleh kosong
                    If Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        MsgBox("Contact person name cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                ElseIf e.ColumnIndex = Me.ColPersonEmail.Index Then
                    Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value <> "" Then
                        Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                        If Expression.IsMatch(Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("The email address is not valid.", MsgBoxStyle.Information, "Information")
                            Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        Expression = Nothing
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemCP_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvCP.UserAddedRow
        Try
            Me.ObjCekData.SetNumber(Me.dgvCP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemBank_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvBank.CellBeginEdit
        Try
            If e.ColumnIndex = 0 Then
                e.Cancel = True
            Else
                Me.buffer = Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan nilai true jika ada 2 baris dalam dgv bank yang double. Dianggap sama jika
    ''' Nama bank, bank account no sama
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ApakahAdaDoubleBank() As Boolean
        Try
            If Me.dgvBank.RowCount > 2 Then
                For i As Integer = 0 To Me.dgvBank.RowCount - 3
                    For j As Integer = i + 1 To Me.dgvBank.RowCount - 2
                        If Me.dgvBank.Item(1, i).Value = Me.dgvBank.Item(1, j).Value And Me.dgvBank.Item(2, i).Value = Me.dgvBank.Item(2, j).Value Then
                            Return True
                        End If
                    Next
                Next
            End If

            'kalo sampe sini maka tidak ada yang sama
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItemBank_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvBank.CellEndEdit
        Try
            If IsNothing(Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
            Else
                If e.ColumnIndex = 1 Then
                    'untuk ItemBank nama tidak boleh kosong
                    If Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        MsgBox("Bank's name cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                    'cek tidak terima ada data double
                    If Me.ApakahAdaDoubleBank = True Then
                        MsgBox("You cannot have 2 bank's data with the same bank's name and account no", _
                        MsgBoxStyle.Information, "Information")
                        Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                ElseIf e.ColumnIndex = 2 Then
                    'untuk account number juga tidak boleh kosong
                    If Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        MsgBox("Account number cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                    'cek tidak terima ada data double
                    If Me.ApakahAdaDoubleBank = True Then
                        MsgBox("You cannot have 2 bank's data with the same bank's name and account no", _
                        MsgBoxStyle.Information, "Information")
                        Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                ElseIf e.ColumnIndex = 3 Then   'untuk currency harus yang ada didalam data master currency
                    Dim Obj As New clsMasterCurrency
                    If Obj.ApakahSudahAda(Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Dim i As New FrmCariMasterCurrency
                        If Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value <> "" Then
                            i.txtCari.Text = Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value
                        End If

                        i.ShowDialog()
                        If i.PrimaryCurrency <> -1 Then
                            Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.ObjCurrency.GetCurrencyCode(i.PrimaryCurrency)).ToUpper
                            Me.dgvBank.Item(8, e.RowIndex).Value = i.PrimaryCurrency
                        Else
                            Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        i = Nothing
                    Else
                        Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                        Me.dgvBank.Item(8, e.RowIndex).Value = Me.ObjCurrency.GetPrimary(Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If
                    Obj = Nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvBank_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvBank.UserDeletedRow
        Try
            Me.ObjCekData.SetNumber(Me.dgvBank)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvBank_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvBank.UserDeletingRow
        Try
            If e.Row.Index = Me.dgvBank.Rows.Count - 1 Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvBank_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvBank.UserAddedRow
        Try
            'tentukan nilai default untuk baris terakhir
            With Me.dgvBank.Rows(Me.dgvBank.Rows.Count - 1)
                .Cells(1).Value = "-"
                .Cells(3).Value = "IDR"
                .Cells(Me.ColPrimaryCurrency.Index).Value = Me.ObjCurrency.GetPrimary("IDR")
            End With
            Me.ObjCekData.SetNumber(Me.dgvBank)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCP_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvCP.UserDeletedRow
        Try
            Me.ObjCekData.SetNumber(Me.dgvCP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCP_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvCP.UserDeletingRow
        Try
            If e.Row.Index = Me.dgvCP.Rows.Count - 1 Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCP_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvCP.UserAddedRow
        Try
            'tentukan nilai default untuk baris terakhir
            With Me.dgvCP.Rows(Me.dgvCP.Rows.Count - 1)
                .Cells(1).Value = "-"
            End With
            Me.ObjCekData.SetNumber(Me.dgvCP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAddress_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvAddress.UserDeletedRow
        Try
            Me.ObjCekData.SetNumber(Me.dgvAddress)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAddress_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvAddress.UserDeletingRow
        Try
            If e.Row.Index = Me.dgvAddress.Rows.Count - 1 Then
                e.Cancel = True
            Else
                If Me.baru = False Then
                    If String.IsNullOrEmpty(e.Row.Cells(Me.ColAddressIdentifier.Index).Value) = False Then
                        If Me.objCustomer.ApakahBolehHapus_Address(e.Row.Cells(Me.ColAddressIdentifier.Index).Value) = False Then
                            e.Cancel = True
                        End If
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
                .Cells(Me.ColAddressIdentifier.Index).Value = "-"
            End With
            Me.ObjCekData.SetNumber(Me.dgvAddress)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        TxtRemarks.BackColor = BackColorAll
    End Sub

    Private Sub TxtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        TxtRemarks.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtSalesPerson_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPerson.GotFocus
        Try
            If Me.TxtSalesPerson.Enabled = True Then
                If Me.TxtSalesPerson.ReadOnly = False Then
                    Me.TxtSalesPerson.BackColor = Me.BackColorAll
                Else
                    Me.TxtSalesPerson.BackColor = Me.TxtSalesPersonDesc.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesPerson_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPerson.LostFocus
        Try
            If Me.TxtSalesPerson.Enabled = True Then
                If Me.TxtSalesPerson.ReadOnly = False Then
                    Me.TxtSalesPerson.BackColor = Me.TxtPrimarySalesPerson.BackColor
                Else
                    Me.TxtSalesPerson.BackColor = Me.TxtSalesPersonDesc.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindSalesPerson_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindSalesPerson.Click
        Dim FrmCariGroup As New FrmMasterSalesPerson_Cari

        Try
            FrmCariGroup.BackColorAll = Me.BackColorAll
            If Me.TxtPrimarySalesPerson.Text = -1 And Me.TxtSalesPerson.Text = "" Then
                FrmCariGroup.ShowDialog()
            Else
                FrmCariGroup.txtCari.Text = Me.TxtSalesPerson.Text
                FrmCariGroup.ShowDialog()
            End If
            If CDbl(FrmCariGroup.PrimarySalesPerson) > 0 Then
                Me.TxtPrimarySalesPerson.Text = FrmCariGroup.PrimarySalesPerson
                Me.TxtSalesPerson.Text = Me.ObjSalesP.GetCode(FrmCariGroup.PrimarySalesPerson)
                Me.TxtSalesPersonDesc.Text = Me.ObjSalesP.GetName(FrmCariGroup.PrimarySalesPerson)
                Me.TxtCity.Focus()
            Else
                Me.TxtPrimarySalesPerson.Text = -1
                Me.TxtSalesPerson.Text = ""
                Me.TxtSalesPersonDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCariGroup.Dispose()
        End Try
    End Sub

    Private Sub TxtSalesPerson_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSalesPerson.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjSalesP.ApakahSudahAda(Me.TxtSalesPerson.Text) = True Then
                    Me.TxtPrimarySalesPerson.Text = Me.ObjSalesP.GetPrimary(Me.TxtSalesPerson.Text)
                    Me.TxtSalesPersonDesc.Text = Me.ObjSalesP.GetName(Me.TxtPrimarySalesPerson.Text)
                    Me.TxtCity.Focus()
                Else
                    Me.btnFindSalesPerson.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesPerson_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPerson.Leave
        Try
            If Me.TxtSalesPerson.Text <> "" Then
                If Me.ObjSalesP.ApakahSudahAda(Me.TxtSalesPerson.Text) = True Then
                    Me.TxtPrimarySalesPerson.Text = Me.ObjSalesP.GetPrimary(Me.TxtSalesPerson.Text)
                    Me.TxtSalesPersonDesc.Text = Me.ObjSalesP.GetName(Me.TxtPrimarySalesPerson.Text)
                Else
                    Me.btnFindSalesPerson.PerformClick()
                End If
            Else
                Me.TxtPrimarySalesPerson.Text = -1
                Me.TxtSalesPerson.Text = ""
                Me.TxtSalesPersonDesc.Text = Me.ObjSalesP.GetName(Me.TxtPrimarySalesPerson.Text)
            End If
            Me.TxtSalesPerson.BackColor = Me.TxtPrimarySalesPerson.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCP_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvCP.RowsAdded
        Try
            If e.RowIndex = e.RowCount - 1 And Me.TxtRemarks.Enabled = True Then
                Me.dgvCP.Item(Me.ColPersonSendCL.Index, e.RowIndex).Value = Boolean.FalseString
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCP_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvCP.DataError
        'abaikan saja
    End Sub

    Private Sub Txt_E_Faktur_Jalan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_E_Faktur_Jalan.GotFocus, Txt_E_Faktur_Blok.GotFocus, _
    Txt_E_Faktur_Jalan.GotFocus, Txt_E_Faktur_Kelurahan.GotFocus, Txt_E_Faktur_KodePos.GotFocus, Txt_E_Faktur_No.GotFocus, Txt_E_Faktur_NoTelepon.GotFocus, Txt_E_Faktur_Propinsi.GotFocus, _
    Txt_E_Faktur_RT.GotFocus, Txt_E_Faktur_RW.GotFocus, Txt_E_Faktur_Kabupaten.GotFocus, Txt_E_Faktur_Kecamatan.GotFocus, TxtTKUID.GotFocus, TxtEmail.GotFocus
        Dim Txt As TextBox

        Try
            If Me.btnOK.Enabled = True Then
                Txt = TryCast(sender, TextBox)
                If IsNothing(Txt) = False Then
                    If Txt.ReadOnly = True Then
                        Txt.BackColor = Me.TxtCustomerGroupDesc.BackColor
                    Else
                        Txt.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_E_Faktur_Jalan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_E_Faktur_Jalan.LostFocus, Txt_E_Faktur_Blok.LostFocus, _
    Txt_E_Faktur_Jalan.LostFocus, Txt_E_Faktur_Kelurahan.LostFocus, Txt_E_Faktur_KodePos.LostFocus, Txt_E_Faktur_No.LostFocus, Txt_E_Faktur_NoTelepon.LostFocus, Txt_E_Faktur_Propinsi.LostFocus, _
    Txt_E_Faktur_RT.LostFocus, Txt_E_Faktur_RW.LostFocus, Txt_E_Faktur_Kabupaten.LostFocus, Txt_E_Faktur_Kecamatan.LostFocus, TxtTKUID.LostFocus, TxtEmail.LostFocus
        Dim Txt As TextBox

        Try
            If Me.btnOK.Enabled = True Then
                Txt = TryCast(sender, TextBox)
                If IsNothing(Txt) = False Then
                    If Txt.ReadOnly = True Then
                        Txt.BackColor = Me.TxtCustomerGroupDesc.BackColor
                    Else
                        Txt.BackColor = Me.TxtPrimaryCustomer.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedTab.Name = Me.TabPage4.Name Then
                If Me.btnOK.Enabled = True Then
                    'set aktif
                    Me.ObjCD.SetBackColor(Me.gbo_EFaktur.Controls, Me.TxtPrimaryCustomer.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.TableLayoutPanel1.Controls, Me.TxtPrimaryCustomer.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.gboTax.Controls, Me.TxtPrimaryCustomer.BackColor, False)
                Else
                    'set non aktif
                    Me.ObjCD.SetBackColor(Me.gbo_EFaktur.Controls, Me.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.TableLayoutPanel1.Controls, Me.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.gboTax.Controls, Me.BackColor, False)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboNIKPASS_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboNIKPASS.SelectedIndexChanged
        '//TR : Tambahan baru Nik/Pas
        Try
            If Me.cboNIKPASS.SelectedIndex = 0 Then
                OptionNIKPASS = clsMasterCustomer.EnumOptionNikPass.ByNIK
            ElseIf Me.cboNIKPASS.SelectedIndex = 1 Then
                OptionNIKPASS = clsMasterCustomer.EnumOptionNikPass.ByPASSPORT
            Else
                '2025 Core Tax Tahap 2.1
                OptionNIKPASS = clsMasterCustomer.EnumOptionNikPass.ByOtherID
            End If

            '2025 Core Tax Tahap 2.1
            If Me.CboTipeNPWP.SelectedIndex = ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                If Me.cboNIKPASS.SelectedIndex = clsMasterCustomer.EnumOptionNikPass.ByNIK Then
                    Me.TxtNPWP.Text = Me.txtNIKPASS.Text
                Else
                    Me.TxtNPWP.Text = "0000 0000 0000 0000"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeRef_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeRef.LostFocus
        TxtCodeRef.BackColor = Me.TxtPrimaryCustRef.BackColor
    End Sub

    Private Sub TxtCodeRef_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeRef.GotFocus
        TxtCodeRef.BackColor = BackColorAll
    End Sub

    Private Sub BtnFindRef_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFindRef.Click
        Dim FrmCariGroup As New FrmMasterCustomer_Reference_Cari

        Try
            FrmCariGroup.BackColorAll = Me.BackColorAll
            If Me.TxtPrimaryCustRef.Text = -1 And Me.TxtCodeRef.Text = "" Then
                FrmCariGroup.ShowDialog()
            Else
                FrmCariGroup.txtCari.Text = Me.TxtCodeRef.Text
                FrmCariGroup.ShowDialog()
            End If
            If CDbl(FrmCariGroup.PrimaryCustRef) > 0 Then
                Me.TxtPrimaryCustRef.Text = FrmCariGroup.PrimaryCustRef
                Me.TxtCodeRef.Text = Me.ObjRef.GetCode(FrmCariGroup.PrimaryCustRef)
                Me.DtpJoinDate.Focus()
            Else
                Me.TxtPrimaryCustRef.Text = -1
                Me.TxtCodeRef.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCariGroup.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeRef_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCodeRef.Leave
        Try
            If Me.TxtCodeRef.Text <> "" Then
                If Me.ObjRef.ApakahSudahAda(Me.TxtCodeRef.Text) = True Then
                    Me.TxtPrimaryCustRef.Text = Me.ObjRef.GetPrimary(Me.TxtCodeRef.Text)
                    Me.TxtCodeRef.Text = Me.ObjRef.GetCode(Me.TxtPrimaryCustRef.Text)
                Else
                    Me.BtnFindRef.PerformClick()
                End If
            Else
                Me.TxtPrimaryCustRef.Text = -1
                Me.TxtCodeRef.Text = ""
            End If
            Me.TxtCodeRef.BackColor = Me.TxtPrimaryCustRef.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeRef_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeRef.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjRef.ApakahSudahAda(Me.TxtCodeRef.Text) = True Then
                    Me.TxtPrimaryCustRef.Text = Me.ObjRef.GetPrimary(Me.TxtCodeRef.Text)
                    Me.TxtCodeRef.Text = Me.ObjRef.GetCode(Me.TxtPrimaryCustRef.Text)
                    Me.DtpJoinDate.Focus()
                Else
                    Me.BtnFindRef.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbEditNPWPLama_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LlbEditNPWPLama.Click
        'Comment: Perubahan NPWP 16 Digit
        Dim hj As New FrmMaster_EditNPWPLama
        Try
            If Me.TxtNPWPLama.Text = "" Then
                Me.TxtNPWPLama.Text = Me.TxtNPWP.Text
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

    Private Sub TxtNPWP_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWP.Leave
        Try
            '2024-09-19 IM, untuk mengantisipasi user mengganti no npwp sebelum tanggal mulai npwp 16 digit
            If Today.Date < Me.ObjSettingNPWP.Tglberlaku Then
                Me.TxtNPWPLama.Text = Me.TxtNPWP.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim bufferText As String = ""
    Private Sub txtNIKPASS_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNIKPASS.Enter
        Try
            Me.bufferText = Me.txtNIKPASS.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    '2024-09-27 IM, Perubahan NPWP 16 Strict: Start
    Private Sub CboTipeNPWP_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboTipeNPWP.SelectedIndexChanged
        Try
            If Today.Date >= Me.ObjSettingNPWP.TglberlakuStrict And CboTipeNPWP.SelectedIndex = 1 Then
                Me.TxtNPWP.ReadOnly = True
                Me.TxtNPWP.BackColor = Me.TxtSalesPersonDesc.BackColor
                If String.IsNullOrEmpty(Me.txtNIKPASS.Text) = False And Me.txtNIKPASS.Text <> "-" Then
                    '2025 Core Tax Tahap 2.1
                    If Me.cboNIKPASS.SelectedIndex = clsMasterCustomer.EnumOptionNikPass.ByNIK Then
                        Me.TxtNPWP.Text = Me.txtNIKPASS.Text
                    Else
                        Me.TxtNPWP.Text = "0000 0000 0000 0000"
                    End If
                End If
            Else
                Me.TxtNPWP.ReadOnly = False
                Me.TxtNPWP.BackColor = Me.TxtPrimaryCustomer.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNIKPASS_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNIKPASS.GotFocus

        txtNIKPASS.BackColor = BackColorAll
    End Sub

    Private Sub txtNIKPASS_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNIKPASS.LostFocus
        txtNIKPASS.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub txtNIKPASS_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNIKPASS.Leave
        Try
            If Today.Date >= Me.ObjSettingNPWP.TglberlakuStrict Then
                If Me.CboTipeNPWP.SelectedIndex = 1 Then
                    If Me.bufferText <> Me.txtNIKPASS.Text Then
                        If Me.txtNIKPASS.Text <> "-" And String.IsNullOrEmpty(Me.txtNIKPASS.Text) = False Then
                            If Me.cboNIKPASS.SelectedIndex = clsMasterCustomer.EnumOptionNikPass.ByNIK Then
                                Me.TxtNPWP.Text = Me.txtNIKPASS.Text
                            Else
                                Me.TxtNPWP.Text = "0000 0000 0000 0000"
                            End If
                        End If
                    Else
                        Me.txtNIKPASS.Text = Me.bufferText
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    '2024-09-27 IM, Perubahan NPWP 16 Strict: End

    'Update SMESS 2026
    Private Sub llbCopy_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCopy.LinkClicked
        Dim ij As New FrmCariMasterCustomer
        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            If Not ij.PrimaryCustomer = -1 Then
                Me.KosongkanField()
                If Me.objCustomer.CariDetail(ij.PrimaryCustomer, Me.dgvAddress, _
                Me.dgvCP, Me.dgvBank) = True Then
                    Me.SalinKeField()
                    Me.baru = True
                    Me.AktifkanDetail()

                    Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
                    Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                    Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
                    Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

                    Me.TxtCustomerCode.Text = "COPY OF " & Me.TxtCustomerCode.Text
                    Me.TxtCustomerName.Text = "COPY OF " & Me.TxtCustomerName.Text
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub
End Class
