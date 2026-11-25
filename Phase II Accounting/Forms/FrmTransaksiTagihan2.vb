''@rosadinnn
''201409091434
Option Explicit On
Option Strict Off
Imports Microsoft.Win32
Imports System.Net
Public Class FrmTransaksiTagihan2
    Dim frmTagihan As New AccModule.FrmTransaksiTagihan
    Dim PrimaryEmailAccount As Double = -1
    Dim rowEmailAccount As DseTransaksiTagihan.DataTable5Row
    Dim Primary As Double = -1
    Dim Notransaksi As String = ""
    Dim iPrimaryCust As Double = -1
    Dim PrimaryCustAcuan As Double = -1
    Dim OnProgressCollectionLetter As Boolean = False
    Dim PrinterNameAcuan As String = ""

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public listOfCust As New List(Of Double)
    Public vPrimaryCust As Double = -1
    Public UserCode As String = "USER"
    Public StopProses As Boolean = False
    Dim objCust As New Core_Module.clsMasterCustomer
    Dim ObjSuratTagihanCustom As New CustomSMESS.ClsTransaksiTagihan
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser
    Dim tblEmail As New DseTransaksiTagihan.DataTable5DataTable
    Dim obj As New ClsTransaksiTagihan
    Dim objUser As New Core_Module.ClsUser
    Dim bufferTime As String = Format(Now, "hh:mm")

    Dim objEmailSettings As New Core_Module.ClsEmailSettings
    Dim listAddr As List(Of String)
    Dim listAttc As List(Of String) = New List(Of String)
    Dim eFrom As String = ""
    Dim eReplyTo As String = ""
    Dim eUsername As String = ""
    Dim ePassword As String = ""
    Dim eSmtpServer As String = ""
    Dim eSmtpPort As String = ""
    Dim eCc As String = ""
    Dim eSubject As String = ""
    Dim eMessage As String = ""
    Dim eAttachmentFilename As String = ""
    Dim eUseSSL As Boolean = False

    Private Sub FrmTransaksiTagihan2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'cek llb settings
            If Me.objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmEmailSettings") = False Then
                Me.llbSettingEmail.Enabled = False
            End If

            If Me.listOfCust.Count = 1 Then
                Me.LblCust.Text = Me.objCust.GetCustName(Me.listOfCust(0))
            Else
                Me.LblCust.Visible = False
                Me.LblCustomer.Visible = False
            End If

            Me.txtSubject.Text = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SuratTagihan, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.SuratTagihan_Subject)
            Me.txtSignByPosition.Text = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SuratTagihan, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.SuratTagihan_SignedByPosition)
            Me.txtSignByName.Text = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SuratTagihan, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.SuratTagihan_SignedByName)
            Me.txtRemarks.Text = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SuratTagihan, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.SuratTagihan_Remarks)
            If Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.CollectionLetter, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.CollectionLetter_IsSendEmail, Me.UserCode) = "1" Then
                chkSendEmail.Checked = True
            Else
                chkSendEmail.Checked = False
            End If

            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SuratTagihan, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.SuratTagihan_BuatCustom) = "1" Then
                If Me.ObjSuratTagihanCustom.ApakahAutoNumber = False Then
                    Me.chkDirectlySave.Visible = False
                    Me.chkDirectlySave.TabStop = False
                End If
            End If

            'rnew
            Me.txtTime.Text = bufferTime
            If chkSendEmail.Checked = True Then
                chkSendEmail.Checked = False
                chkSendEmail.Checked = True
            Else
                chkSendEmail.Checked = True
                chkSendEmail.Checked = False
            End If
            Me.LoadEmailAccount(Me.UserCode, Me.tblEmail)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CheckEmailCustomer() As Boolean
        Dim DTA As New Core_Module.dseCustomerTableAdapters.tblmaster_customer_itemcp1TableAdapter
        Dim TBL As New Core_Module.dseCustomer.tblmaster_customer_itemcp1DataTable
        Dim row As Core_Module.dseCustomer.tblmaster_customer_itemcp1Row
        Dim ListCustMasalah As New List(Of Double)
        Dim Betul As Boolean = True
        Dim DV As DataRowView
        Dim AdaSendColl As Boolean = False

        Try
            Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
            Me.objCust.SetConnectionString(DTA)
            DTA.Fill(TBL)
            For Each vPrimaryCust As Double In Me.listOfCust
                Betul = True
                TBL.DefaultView.RowFilter = "[Primary Customer] = " + vPrimaryCust.ToString
                'cek:
                '1. minimal ada 1 baris
                If TBL.DefaultView.Count = 0 Then
                    Betul = False
                End If
                '2. minimal ada 1 baris yang send collection letter email
                If Betul = True Then
                    AdaSendColl = False
                    For Each DV In TBL.DefaultView
                        row = DV.Row
                        If row.Send_Collection_Letter = 1 Then
                            AdaSendColl = True
                        End If
                    Next
                    If AdaSendColl = False Then
                        Betul = False
                    End If
                End If
                '3. untuk yang send collection letter email, pastikan alamat emailnya valid
                If Betul = True Then
                    For Each DV In TBL.DefaultView
                        row = DV.Row
                        If row.Send_Collection_Letter = 1 Then
                            If Expression.IsMatch(row._E_mail_Address) = False Then
                                Betul = False
                                Exit For
                            End If
                        End If
                    Next
                End If
                If Betul = False Then
                    ListCustMasalah.Add(vPrimaryCust)
                End If
            Next
            If ListCustMasalah.Count > 0 Then
                If Me.objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterCustomer") = True Then
                    'tampilkan form perbaikan
                    MsgBox("Please fix customer data to send collection letter.", MsgBoxStyle.Information, "Information")
                    Dim ijk As New Core_Module.FrmMasterCustomer_CekEmailCP
                    ijk.ListPrimaryCust = ListCustMasalah
                    ijk.BackColorAll = Me.BackColorAll
                    ijk.ShowDialog()
                Else
                    'berikan peringatan
                    If MsgBox("The following customer have invalid data for sending collection letter email. Ingore this problem and continue sending email for customer with valid data?", _
                    MsgBoxStyle.Information, "Information") <> MsgBoxResult.Yes Then
                        Return False
                    End If
                End If
            End If
            Expression = Nothing

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ListCustMasalah = Nothing
        End Try
    End Function

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Dim objCd As New Core_Module.ClsCekData
        Dim namaData As New Collection
        Dim isiData As New Collection
        Dim printerName As String = ""

        Try
            'Dim siteResponds As Boolean = False
            'siteResponds = My.Computer.Network.Ping("http://www.google.com/index.html")
            'If siteResponds = False Then
            '    If Me.chkSendEmail.Checked = True Then
            '        MsgBox("Failed to connect to the internet, you cannot send emails.", MsgBoxStyle.Information, "Information")
            '        Me.chkSendEmail.Checked = False
            '    End If
            'End If

            Me.Cursor = Cursors.WaitCursor

            ''validasi semua parameter:
            namaData.Add("Subject")
            isiData.Add(Me.txtSubject)
            namaData.Add("Signed By Name")
            isiData.Add(Me.txtSignByName)
            namaData.Add("Signed By Position")
            isiData.Add(Me.txtSignByPosition)
            namaData.Add("Remarks")
            isiData.Add(Me.txtRemarks)
            If objCd.CekDataKosong(isiData, namaData) = False Then
                Exit Sub
            End If

            If chkSendEmail.Checked = True And Me.tblEmail.Count = 0 Then
                MsgBox("You cannot send email, because there isn't account associated with your username.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If chkSendEmail.Checked = True And Me.cboEmailList.SelectedIndex = -1 Then
                MsgBox("Please select an email account.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            ''get printer name
            If Me.chkDirectlyPrint.Checked = True Then
                If Me.PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    printerName = Me.PrintDialog1.PrinterSettings.PrinterName
                Else
                    MsgBox("You have to select a printer", MsgBoxStyle.Exclamation, "Warning")
                    Exit Sub
                End If
            End If

            ''do it
            If Me.listOfCust.Count > 0 Then
                If Me.chkSendEmail.Checked = True Then
                    rowEmailAccount = tblEmail.Rows(Me.cboEmailList.SelectedIndex)
                    Me.PrimaryEmailAccount = rowEmailAccount.Primary

                    'cek dan perbaiki email customer jika ada yang salah
                    If Me.CheckEmailCustomer = False Then
                        Exit Sub
                    End If
                End If
                If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SuratTagihan, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.SuratTagihan_BuatCustom) = "1" Then
                    Me.RunAoutoCollectionLetters_Custom(printerName)
                Else
                    Me.RunAoutoCollectionLetters(printerName)
                End If
            Else
                MsgBox("No customer data", MsgBoxStyle.Information, "Information")
                Me.SimpanDataSettingDanTutupForm()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        Dim frm As New Core_Module.FrmMasterNumbering
        Try
            frm.ShowDialog()
        Catch ex As Exception
            frm.Dispose()
        End Try
    End Sub

    Private Sub txtTime_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTime.Enter
        Try
            ToolTip1.Show("Type valid time, e.g. 01:00", Me.txtTime)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTime_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTime.GotFocus
        Try
            Me.txtTime.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTime_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTime.KeyPress
        Try
            If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) And Not e.KeyChar.ToString = ":" Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTime_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTime.Leave
        Try
            If IsValidTime(Me.txtTime.Text) = False Then
                MsgBox("Time format is not valid", MsgBoxStyle.Information, "Information")
                Me.txtTime.Focus()
                Me.txtTime.SelectAll()
                Me.ToolTip1.Hide(Me.txtTime)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTime_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTime.LostFocus
        Try
            Me.txtTime.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSubject_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubject.GotFocus
        Try
            If Me.txtSubject.ReadOnly = False Then
                Me.txtSubject.BackColor = Me.BackColorAll
            Else
                Me.txtSubject.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSubject_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubject.LostFocus
        Try
            If Me.txtSubject.ReadOnly = False Then
                Me.txtSubject.BackColor = Me.txtPrimary.BackColor
            Else
                Me.txtSubject.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSignByName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSignByName.GotFocus
        Try
            If Me.txtSignByName.ReadOnly = False Then
                Me.txtSignByName.BackColor = Me.BackColorAll
            Else
                Me.txtSignByName.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSignByName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSignByName.LostFocus
        Try
            If Me.txtSignByName.ReadOnly = False Then
                Me.txtSignByName.BackColor = Me.txtPrimary.BackColor
            Else
                Me.txtSignByName.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSignByPosition_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSignByPosition.GotFocus
        Try
            If Me.txtSignByPosition.ReadOnly = False Then
                Me.txtSignByPosition.BackColor = Me.BackColorAll
            Else
                Me.txtSignByPosition.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSignByPosition_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSignByPosition.LostFocus
        Try
            If Me.txtSignByPosition.ReadOnly = False Then
                Me.txtSignByPosition.BackColor = Me.txtPrimary.BackColor
            Else
                Me.txtSignByPosition.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.GotFocus
        Try
            If Me.txtRemarks.ReadOnly = False Then
                Me.txtRemarks.BackColor = Me.BackColorAll
            Else
                Me.txtRemarks.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.LostFocus
        Try
            If Me.txtRemarks.ReadOnly = False Then
                Me.txtRemarks.BackColor = Me.txtPrimary.BackColor
            Else
                Me.txtRemarks.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Cancel creating collection letter?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.StopProses = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkSendEmail_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSendEmail.CheckedChanged
        Try
            If Me.chkSendEmail.Checked = True Then
                Me.cboEmailList.Enabled = True
                If Me.objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmEmailSettings") = True Then
                    Me.llbSettingEmail.Enabled = True
                Else
                    Me.llbSettingEmail.Enabled = False
                End If
                If Me.cboEmailList.Items.Count > 0 Then
                    Me.cboEmailList.SelectedIndex = 0
                End If
            Else
                Me.cboEmailList.Enabled = False
                Me.llbSettingEmail.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#Region "function-sub"
    Private Sub LoadEmailAccount(ByVal vUserCode As String, ByVal vTbl As DseTransaksiTagihan.DataTable5DataTable)
        Dim rowEmail As DseTransaksiTagihan.DataTable5Row
        Dim cboIndex As Integer = -1
        Dim lastEmail As String = ""

        Try
            lastEmail = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.CollectionLetter, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.CollectionLetter_LastEmailUsed, Me.UserCode)
            Me.cboEmailList.Items.Clear()
            obj.LoadEmailAccount_CollectionLetter(vUserCode, vTbl)
            If vTbl.Count > 0 Then
                For Each rowEmail In Me.tblEmail
                    cboEmailList.Items.Add(rowEmail.Account_Code & " - " & rowEmail.Email_From)
                    If lastEmail = rowEmail.Account_Code & " - " & rowEmail.Email_From Then
                        cboIndex = Me.tblEmail.Rows.IndexOf(rowEmail)
                    End If
                Next
                Me.cboEmailList.SelectedIndex = cboIndex
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function AdaKarakterAnehKah(ByVal textbox As TextBox) As Boolean
        Try
            If textbox.Text.Contains("#") Or textbox.Text.Contains("%") Or textbox.Text.Contains("*") Or textbox.Text.Contains("'") Then
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub RunAoutoCollectionLetters_Custom(ByVal vPrinterName As String)
        Dim ShowProses As Boolean = False

        Try
            Me.ControlBox = False
            Me.btnStart.Enabled = False
            Me.prgAll.Maximum = Me.listOfCust.Count
            Me.btnCancel.Enabled = True

            For Each primaryCust As Double In Me.listOfCust
                If Me.StopProses = False Then
                    Me.prgAll.Value = Me.listOfCust.IndexOf(primaryCust) + 1
                    Me.lblProgress.Text = "On Process: " & Me.prgAll.Value.ToString + " of " + Me.prgAll.Maximum.ToString
                    If Me.listOfCust.Count > 1 Then
                        ShowProses = True
                    End If

                    If Me.ObjSuratTagihanCustom.ApakahSudahAda(primaryCust, Me.dtpEx.Value) = False Then
                        Me.OnProgressCollectionLetter = True
                        If Me.ObjSuratTagihanCustom.BuatSuratTagihan(Me.dtpDue.Value, primaryCust, Me.chkDirectlySave.Checked, Me.chkDirectlyPrint.Checked, vPrinterName, Me.UserCode, _
                         Me.txtSubject.Text, Me.txtSignByName.Text, Me.txtSignByPosition.Text, Me.txtRemarks.Text, ShowProses, Me.lblProgress.Text, Me.prgAll.Maximum, Me.prgAll.Value, _
                         Me.chkSendEmail.Checked, Me.PrimaryEmailAccount) = False Then
                            Me.StopProses = True
                        End If
                        Me.OnProgressCollectionLetter = False
                    End If
                Else
                    Me.StopProses = False
                    Exit Sub
                End If
            Next

            Me.SimpanDataSettingDanTutupForm()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        Finally
            Me.lblProgress.Text = "Ready"
            Me.prgAll.Value = 0
            Me.prgAll.Maximum = 0
            Me.btnStart.Enabled = True
            Me.ControlBox = True
            'objCust = Nothing
            Me.btnCancel.Enabled = False
        End Try
    End Sub

    Private Sub RunAoutoCollectionLetters(ByVal vPrinterName As String)
        
        Try
            Me.ControlBox = False
            Me.btnStart.Enabled = False
            Me.btnCancel.Enabled = True
            Me.prgAll.Maximum = Me.listOfCust.Count

            frmTagihan = New AccModule.FrmTransaksiTagihan
            frmTagihan.Show()
            frmTagihan.Hide()
            frmTagihan.Top = (frmTagihan.Height * -1) - 1
            frmTagihan.Left = (frmTagihan.Width * -1) - 1

            If Me.chkDirectlySave.Checked = False Then
                Me.NonAktifkanForm()
                For Each primaryCust As Double In Me.listOfCust
                    If Me.StopProses = False Then
                        Me.prgAll.Value = Me.listOfCust.IndexOf(primaryCust) + 1
                        Me.lblProgress.Text = "On Process: " & Me.prgAll.Value.ToString + " of " + Me.prgAll.Maximum.ToString

                        If Me.obj.ApakahSudahAda(primaryCust, Me.dtpEx.Value) = False Then
                            Me.BuatSuratTagihan(Me.dtpDue.Value, primaryCust, Me.chkDirectlySave.Checked, Me.chkDirectlyPrint.Checked, vPrinterName, Me.UserCode, _
                                Me.txtSubject.Text, Me.txtSignByName.Text, Me.txtSignByPosition.Text, Me.txtRemarks.Text, _
                                Me.chkSendEmail.Checked, Me.PrimaryEmailAccount, Me.frmTagihan)
                        End If
                    Else
                        Me.AktifkanForm()
                        Me.StopProses = False
                        Exit Sub
                    End If
                Next

                Me.SimpanDataSettingDanTutupForm()
            Else
                Me.PrinterNameAcuan = vPrinterName
                Me.BackgroundWorker2.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        Finally
            If Me.chkDirectlySave.Checked = False Then
                Me.lblProgress.Text = "Ready"
                Me.prgAll.Value = 0
                Me.prgAll.Maximum = 0
                Me.btnStart.Enabled = True
                Me.ControlBox = True
                'objCust = Nothing
                Me.btnCancel.Enabled = False
            End If
        End Try
    End Sub

    Private Sub SimpanDataSettingDanTutupForm()
        Try
            'simpan data setting
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SuratTagihan, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.SuratTagihan_Remarks, Me.txtRemarks.Text.Trim)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SuratTagihan, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.SuratTagihan_SignedByName, Me.txtSignByName.Text.Trim)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SuratTagihan, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.SuratTagihan_SignedByPosition, Me.txtSignByPosition.Text.Trim)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SuratTagihan, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.SuratTagihan_Subject, Me.txtSubject.Text.Trim)
            If Me.chkSendEmail.Checked = True Then
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.CollectionLetter, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.CollectionLetter_IsSendEmail, "1", Me.UserCode)
            Else
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.CollectionLetter, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.CollectionLetter_IsSendEmail, "0", Me.UserCode)
            End If
            Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.CollectionLetter, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.CollectionLetter_LastEmailUsed, Me.cboEmailList.Text, Me.UserCode)
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function BuatSuratTagihan(ByVal vDueDate As Date, ByVal vPrimaryCust As Double, _
            ByVal vDirectSave As Boolean, ByVal vDirectPrint As Boolean, ByVal vPrinterName As String, ByVal vUserCode As String, _
            ByVal vSubject As String, ByVal vSignedByName As String, ByVal vSignedByPosition As String, ByVal vRemarks As String, _
            ByVal vSendToEmail As Boolean, ByVal vPrimaryAccountCode As String, ByVal frmTagihan As FrmTransaksiTagihan) As Boolean
        Dim dtaAttnAddr As New DseTransaksiTagihanTableAdapters.getAttAndAddrTableAdapter
        Dim tblAttnAddr As New DseTransaksiTagihan.getAttAndAddrDataTable
        Dim rowAttnAddr As DseTransaksiTagihan.getAttAndAddrRow

        Try
            Me.OnProgressCollectionLetter = True
            ''set default property
            If Me.listOfCust.Count > 1 Then
                frmTagihan.ShowStopProcess = True
            End If
            frmTagihan.QuickInput = True
            'frmTagihan.Show()
            'frmTagihan.Hide()
            frmTagihan.UserCode = vUserCode
            frmTagihan.Top = (frmTagihan.Height * -1) - 1
            frmTagihan.Left = (frmTagihan.Width * -1) - 1
            frmTagihan.DirectlyPrint = vDirectPrint
            frmTagihan.PrinterName = vPrinterName
            frmTagihan.lblProgress.Text = Me.lblProgress.Text
            frmTagihan.prgAll.Value = 0
            frmTagihan.prgAll.Maximum = Me.prgAll.Maximum
            frmTagihan.prgAll.Value = Me.prgAll.Value

            ''set form parameters
            frmTagihan.NewTagihan()
            frmTagihan.dtpDate.Value = vDueDate
            frmTagihan.TxtSubject.Text = vSubject
            frmTagihan.TxtPrimaryCust.Text = vPrimaryCust
            frmTagihan.TxtCustCode.Text = objCust.GetCustomerCode(vPrimaryCust)
            frmTagihan.TxtCustName.Text = objCust.GetCustName(vPrimaryCust)
            objCust.IsiCboAddress(frmTagihan.cboAdd, vPrimaryCust)
            frmTagihan.TxtSignedBy.Text = vSignedByName
            frmTagihan.TxtPosition.Text = vSignedByPosition
            frmTagihan.TxtRemark.Text = vRemarks
            frmTagihan.IsiDGVOtomatis()
            ''lewati customer yang tidak punya hutang (invoice)
            If frmTagihan.dgvItem.Rows.Count <= 0 Then
                'MsgBox("Tidak ada Invoice untuk Customer: " & objCust.GetCustName(vPrimaryCust), MsgBoxStyle.Information, "Informasi")
                Me.OnProgressCollectionLetter = False
                Return True
            End If

            ''get attn and address
            dtaAttnAddr.Fill(tblAttnAddr, vPrimaryCust)
            If tblAttnAddr.Rows.Count > 0 Then

                ''jika attn dan addr ditemukan, salin
                rowAttnAddr = tblAttnAddr.Rows(tblAttnAddr.Rows.Count - 1)
                frmTagihan.TxtAttn.Text = rowAttnAddr.Attn
                frmTagihan.TxtCustAddIdentifier.Text = rowAttnAddr.Identifier_Alamat_Cust
                frmTagihan.cboAdd.Text = Me.objCust.GetKodeAddress(frmTagihan.TxtCustAddIdentifier.Text)

                ''asumsi data sudah lengkap,
                ''lakukan proses dibelakang jika directly save to db
                If vDirectSave = True Then
                    frmTagihan.SimpanBaru()
                Else
                    frmTagihan.Visible = False
                    frmTagihan.ShowDialog()
                End If
            Else
                'isikan nama attn berdasarkan data contact person pertama yang dikirimkan collection letter
                frmTagihan.TxtAttn.Text = Me.objCust.GetCustContactPerson_ForCollectionLetter(vPrimaryCust)
                ''jika attn dan addr tidak ditemukan, paksa user untuk mengisi
                If frmTagihan.TxtAttn.Text = "" Or frmTagihan.cboAdd.Items.Count = 0 Then
                    frmTagihan.Visible = False
                    If frmTagihan.cboAdd.Items.Count > 0 Then
                        frmTagihan.cboAdd.SelectedIndex = 0
                    End If
                    frmTagihan.ShowDialog()
                Else
                    If vDirectSave = True Then
                        frmTagihan.SimpanBaru()
                    Else
                        frmTagihan.Visible = False
                        frmTagihan.ShowDialog()
                    End If
                End If
            End If

            ''print
            'If vDirectPrint = True Then
            '    frmTagihan.PrintToReal(vPrinterName)
            'End If

            'rnew
            'If Me.chkDirectlySave.Checked = False Then
            '    frmTagihan.ShowDialog()
            '    'Else
            '    '    frmTagihan.SimpanBaru()
            'End If
            If vSendToEmail = True Then
                PrimaryEmailAccount = vPrimaryAccountCode
                Primary = frmTagihan.TxtPrimary.Text ''belum set betul
                Notransaksi = frmTagihan.TxtNoTrans.Text
                Me.iPrimaryCust = frmTagihan.TxtPrimaryCust.Text
                frmTagihan.PrintAndExport(vPrimaryAccountCode)
                Me.objCust.GetEmailCpList(frmTagihan.TxtPrimaryCust.Text, Me.listAddr)
                Me.objEmailSettings.GetAllDetail(vPrimaryAccountCode, Core_Module.ClsEmailSettings.EnumTipeTransaksiEmail.CollectionLetter, _
                eFrom, eReplyTo, eUsername, ePassword, eSmtpServer, eSmtpPort, eCc, eSubject, eMessage, eusessl)
                Me.eMessage = Me.eMessage.Replace("<<Date>>", Format(frmTagihan.dtpInvoiceDate.Value, "dd MMMM yyyy"))
                If frmTagihan.TxtSubject.Text <> "" Then
                    Me.eMessage = Me.eMessage.Replace("<<Subject>>", frmTagihan.TxtSubject.Text)
                End If
                If frmTagihan.TxtAttn.Text <> "" Then
                    Me.eMessage = Me.eMessage.Replace("<<Attn>>", frmTagihan.TxtAttn.Text)
                End If
                If frmTagihan.TxtCustCode.Text <> "" Then
                    Me.eMessage = Me.eMessage.Replace("<<CustCode>>", frmTagihan.TxtCustCode.Text)
                End If
                If frmTagihan.TxtCustName.Text <> "" Then
                    Me.eMessage = Me.eMessage.Replace("<<CustName>>", frmTagihan.TxtCustName.Text)
                End If
                If frmTagihan.TxtSignedBy.Text <> "" Then
                    Me.eMessage = Me.eMessage.Replace("<<SignBy>>", frmTagihan.TxtSignedBy.Text)
                End If
                If frmTagihan.TxtPosition.Text <> "" Then
                    Me.eMessage = Me.eMessage.Replace("<<Position>>", frmTagihan.TxtPosition.Text)
                End If
                If listAddr.Count > 0 Then
                    If My.Computer.FileSystem.FileExists(Application.StartupPath & "\CL\" & Me.objEmailSettings.GetPrefxFilename(vPrimaryAccountCode, Core_Module.ClsEmailSettings.EnumTipeTransaksiEmail.CollectionLetter) & frmTagihan.TxtNoTrans.Text.Replace("/", "-") & ".pdf") = True Then
                        Me.listAttc.Clear()
                        Me.listAttc.Add(Application.StartupPath & "\CL\" & Me.objEmailSettings.GetPrefxFilename(vPrimaryAccountCode, Core_Module.ClsEmailSettings.EnumTipeTransaksiEmail.CollectionLetter) & frmTagihan.TxtNoTrans.Text.Replace("/", "-") & ".pdf")
                        Me.SendEmail(Me.listAddr, Me.eCc, Me.eSubject, Me.eMessage, Me.listAttc, Me.eFrom, Me.eReplyTo, Me.eSmtpServer, Me.eSmtpPort, _
                        Me.eUsername, Me.ePassword, Me.eUseSSL)
                    End If
                End If
            End If

            If frmTagihan.StopProcess = True Then
                Me.OnProgressCollectionLetter = False
                Me.StopProses = True
                Return False
            End If

            Me.OnProgressCollectionLetter = False
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'frmTagihan.Dispose()
            Me.OnProgressCollectionLetter = False
        End Try
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Function IsValidTime(ByVal vStr As String) As Boolean
        Dim time As String()
        Try
            '1 jika tidak ada tanda :
            If Not Me.txtTime.Text.Contains(":") Then
                Return False
            End If
            '2 jika tanda : lebih dari 1

            '3 jika tidak ada jam dan/atau menit
            If Me.txtTime.Text.StartsWith(":") Then
                Return False
            End If
            If Me.txtTime.Text.EndsWith(":") Then
                Return False
            End If
            '4 jam lebih dari 23
            time = vStr.Split(":")
            If CInt(time(0)) > 23 Then
                Return False
            End If
            '5 menit lebih dari 59
            If CInt(time(1)) > 59 Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Function
#End Region

    Private Sub llbSettingEmail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSettingEmail.LinkClicked
        Dim frm As New Core_Module.FrmEmailSettings
        Try
            frm.UserCode = Me.UserCode
            frm.BackColorAll = Me.BackColorAll
            frm.ShowDialog()
            LoadEmailAccount(Me.UserCode, Me.tblEmail)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

#Region "email control"
    Public Sub SendEmail(ByVal vListAddress As List(Of String), ByVal vCc As String, _
            ByVal vSubject As String, ByVal vMessage As String, ByVal vListFileName As List(Of String), _
            ByVal vFrom As String, ByVal vReplyTo As String, ByVal vSmtpServer As String, ByVal vSmtpPort As String, ByVal vUsername As String, _
            ByVal vPassword As String, ByVal vUseSSL As Boolean)
        Dim AddTo, vFileName As String
        Dim Index1 As Integer = 0
        Try
            Dim MyMailMessage As New Net.Mail.MailMessage()
            For Each AddTo In vListAddress
                MyMailMessage.To.Add(New Mail.MailAddress(AddTo))
            Next
            MyMailMessage.CC.Add(New Mail.MailAddress(vCc))
            MyMailMessage.Subject = vSubject
            MyMailMessage.Body = vMessage
            MyMailMessage.From = New Mail.MailAddress(vFrom)
            MyMailMessage.ReplyTo = New Mail.MailAddress(vReplyTo)
            MyMailMessage.DeliveryNotificationOptions = Mail.DeliveryNotificationOptions.OnSuccess
            For Each vFileName In vListFileName
                If vFileName <> "" Then
                    Dim fileSize As Long
                    fileSize = My.Computer.FileSystem.GetFileInfo(vFileName).Length
                    If fileSize > 10485760 Then
                        MsgBox("Maximum allowable file size is 10 mb.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    MyMailMessage.Attachments.Add(New Mail.Attachment(vFileName))
                    Me.eAttachmentFilename = vFileName
                End If
            Next
            Dim SMTPServer As New Mail.SmtpClient
            SMTPServer.Credentials = New Net.NetworkCredential(vUsername, vPassword)
            SMTPServer.Host = vSmtpServer
            SMTPServer.Port = vSmtpPort
            SMTPServer.EnableSsl = vUseSSL
            Dim UserState As Object = MyMailMessage
            AddHandler SMTPServer.SendCompleted, AddressOf SMTPServer_OnCompleted
            'SMTPServer.SendAsync(MyMailMessage, UserState)
            Me.lblProgress.Text = Me.lblProgress.Text + " - sending email"
            SMTPServer.Send(MyMailMessage)
            Me.lblProgress.Text = Me.lblProgress.Text.Replace(" - sending email", "")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SMTPServer_OnCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        Dim sent As Integer = 0
        Dim _listAddrTo As String = ""
        Try
            Dim CompletedMailMessage As Mail.MailMessage = CType(e.UserState, Mail.MailMessage)
            Dim EmailFrom As String = CompletedMailMessage.From.Address
            Dim EmailSubject As String = CompletedMailMessage.Subject
            If e.Cancelled Then
                MsgBox("SendAsync email from " + EmailFrom.ToString + " regarding " + EmailSubject.ToString + " was canceled.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If e.Error Is Nothing Then
                sent = 1
            Else
                sent = 0
            End If
            For Each str As String In Me.listAddr
                _listAddrTo += str + ";"
            Next
            Me.objEmailSettings.SaveEmailLog(Me.PrimaryEmailAccount, Core_Module.ClsEmailSettings.EnumTipeTransaksiEmail.CollectionLetter, Me.Primary, Me.Notransaksi, Me.eFrom, _listAddrTo, Me.eCc, Me.eSubject, Me.eMessage, sent, Me.UserCode, Now, Me.iPrimaryCust, Me.eAttachmentFilename)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Me.NonAktifkanForm()
            Me.prgAll.Value = 0
            Me.prgAll.Maximum = Me.listOfCust.Count

            For Each primaryCust As Double In Me.listOfCust
                Me.PrimaryCustAcuan = primaryCust

                If Me.StopProses = False Then
                    Me.prgAll.Value += 1
                    Me.lblProgress.Text = "On Process: " & Me.prgAll.Value.ToString + " of " + Me.prgAll.Maximum.ToString + " - " + Me.objCust.GetCustName(primaryCust)
                Else
                    'Me.BackgroundWorker2.ReportProgress(1)
                    Me.SimpanDataSettingDanTutupForm()
                    Exit Sub
                End If

                Me.BackgroundWorker2.ReportProgress(0)
                Threading.Thread.Sleep(1000)
                While Me.OnProgressCollectionLetter = True
                    Threading.Thread.Sleep(1000)
                End While
            Next

            Me.AktifkanForm()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanForm()
        Try
            Me.btnStart.Enabled = False
            Me.ControlBox = False
            Me.btnCancel.Enabled = True
            Me.txtRemarks.ReadOnly = True
            Me.txtSubject.ReadOnly = True
            Me.txtSignByPosition.ReadOnly = True
            Me.txtSignByName.ReadOnly = True
            Me.chkDirectlySave.Enabled = False
            Me.chkDirectlyPrint.Enabled = False
            Me.chkSendEmail.Enabled = False
            Me.cboEmailList.Enabled = False
            Me.llbSettingEmail.Enabled = False
            Me.dtpDue.Enabled = False
            Me.dtpEx.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanForm()
        Try
            Me.lblProgress.Text = "Ready"
            Me.prgAll.Value = 0
            Me.prgAll.Maximum = 0
            Me.btnStart.Enabled = True
            Me.ControlBox = True
            'objCust = Nothing
            Me.btnCancel.Enabled = False
            Me.txtRemarks.ReadOnly = False
            Me.txtSubject.ReadOnly = False
            Me.txtSignByPosition.ReadOnly = False
            Me.txtSignByName.ReadOnly = False
            Me.chkDirectlySave.Enabled = True
            Me.chkDirectlyPrint.Enabled = True
            Me.chkSendEmail.Enabled = True
            Me.cboEmailList.Enabled = True
            If Me.objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmEmailSettings") = True Then
                Me.llbSettingEmail.Enabled = True
            Else
                Me.llbSettingEmail.Enabled = False
            End If
            Me.dtpDue.Enabled = True
            Me.dtpEx.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker2.ProgressChanged
        Try
            If e.ProgressPercentage = 0 Then
                If Me.obj.ApakahSudahAda(Me.PrimaryCustAcuan, Me.dtpEx.Value) = False Then
                    If Me.BuatSuratTagihan(Me.dtpDue.Value, Me.PrimaryCustAcuan, Me.chkDirectlySave.Checked, Me.chkDirectlyPrint.Checked, Me.PrinterNameAcuan, Me.UserCode, _
                        Me.txtSubject.Text, Me.txtSignByName.Text, Me.txtSignByPosition.Text, Me.txtRemarks.Text, _
                        Me.chkSendEmail.Checked, Me.PrimaryEmailAccount, frmTagihan) = False Then
                        Exit Sub
                    End If
                End If
            Else
                MsgBox("Process stopped by user.", MsgBoxStyle.Information, "Information")
                Me.btnCancel.Visible = False
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.btnCancel.Visible = False
            Me.btnCancel.Enabled = False
            Me.btnCancel.TabStop = False

            If Me.StopProses = False Then
                MsgBox("DO creation finish.", MsgBoxStyle.Information, "Information")
            Else
                Me.StopProses = False
                MsgBox("Process stopped by user.", MsgBoxStyle.Information, "Information")
            End If

            Me.btnStart.Enabled = True
            Me.btnCancel.Enabled = True
            Me.ControlBox = True
            Me.AktifkanForm()
            Me.SimpanDataSettingDanTutupForm()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
