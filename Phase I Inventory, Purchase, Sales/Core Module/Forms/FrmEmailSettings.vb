''@rosadinnn
''2015-04-15

Option Explicit On
Option Strict Off
Public Class FrmEmailSettings
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Private Baru As Boolean
    Dim obj As New ClsEmailSettings
    Dim objCd As New ClsCekData
    Dim ObjLicense As New ClsLicense

    Private Sub AktifkanDetail()
        Try
            objCd.KendalikanControl(GroupBox1.Controls, True, True, True, True, False, False, False, False, True, True)
            objCd.KendalikanControl(GroupBox2.Controls, True, True, False, False, False, False, False, False, False, False)
            objCd.KendalikanControl(TabPage_ColLetter.Controls, True, True, False, False, False, False, False, False, False, False)
            objCd.KendalikanControl(TabPage_Attendance.Controls, True, True, False, False, False, False, False, False, False, False)
            objCd.KendalikanControl(TabPage_SS.Controls, True, True, False, False, False, False, False, False, False, False)
            objCd.KendalikanControl(TabPage_PickUp.Controls, True, True, False, False, False, False, False, False, False, False)
            objCd.KendalikanControl(TabPage_OTOrder.Controls, True, True, False, False, False, False, False, False, False, False)

            llbTestEmail1.Enabled = True
            llbTestEmail_Att.Enabled = True
            llbTestEmail_PickUp.Enabled = True
            llbTestEmail_SS.Enabled = True
            llbTestEmail_OTOrder.Enabled = True

            btnOK.Enabled = True
            btnCancel.Enabled = True
            btnFind.Enabled = False
            btnNew.Enabled = False
            btnEdit.Enabled = False
            btnDelete.Enabled = False
            btnExit.Enabled = False
            CancelButton = btnCancel
            chkShowPassword.Enabled = True
            chkShowPassword.Checked = False
            Me.chkUseSSL.Enabled = True

            Me.AturDetailCollectionLetter()
            Me.AturDetailSalarySlip()
            Me.AturDetailAttendance()
            Me.AturDetailPickupOrder()
            Me.AturDetailOTOrder()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AturDetailCollectionLetter()
        Try
            If Me.chkCollectionLetter.Checked = True Then
                Me.txtItemCCto.ReadOnly = False
                Me.txtItemCCto.BackColor = Me.txtPrimary.BackColor
                Me.txtItemFileNamePrefix.ReadOnly = False
                Me.txtItemFileNamePrefix.BackColor = Me.txtPrimary.BackColor
                Me.txtItemMessage.ReadOnly = False
                Me.txtItemMessage.BackColor = Me.txtPrimary.BackColor
                Me.txtItemSubject.ReadOnly = False
                Me.txtItemSubject.BackColor = Me.txtPrimary.BackColor
            Else
                Me.txtItemCCto.ReadOnly = True
                Me.txtItemCCto.BackColor = Me.txtReadonly.BackColor
                Me.txtItemCCto.Text = ""
                Me.txtItemFileNamePrefix.ReadOnly = True
                Me.txtItemFileNamePrefix.BackColor = Me.txtReadonly.BackColor
                Me.txtItemFileNamePrefix.Text = ""
                Me.txtItemMessage.ReadOnly = True
                Me.txtItemMessage.BackColor = Me.txtReadonly.BackColor
                Me.txtItemMessage.Text = ""
                Me.txtItemSubject.ReadOnly = True
                Me.txtItemSubject.BackColor = Me.txtReadonly.BackColor
                Me.txtItemSubject.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AturDetailAttendance()
        Try
            If Me.chkAttendance.Checked = True Then
                Me.txtItemCCto_Att.ReadOnly = False
                Me.txtItemCCto_Att.BackColor = Me.txtPrimary.BackColor
                Me.txtItemFileNamePrefix_Att.ReadOnly = False
                Me.txtItemFileNamePrefix_Att.BackColor = Me.txtPrimary.BackColor
                Me.txtItemMessage_Att.ReadOnly = False
                Me.txtItemMessage_Att.BackColor = Me.txtPrimary.BackColor
                Me.txtItemSubject_Att.ReadOnly = False
                Me.txtItemSubject_Att.BackColor = Me.txtPrimary.BackColor
            Else
                Me.txtItemCCto_Att.ReadOnly = True
                Me.txtItemCCto_Att.BackColor = Me.txtReadonly.BackColor
                Me.txtItemCCto_Att.Text = ""
                Me.txtItemFileNamePrefix_Att.ReadOnly = True
                Me.txtItemFileNamePrefix_Att.BackColor = Me.txtReadonly.BackColor
                Me.txtItemFileNamePrefix_Att.Text = ""
                Me.txtItemMessage_Att.ReadOnly = True
                Me.txtItemMessage_Att.BackColor = Me.txtReadonly.BackColor
                Me.txtItemMessage_Att.Text = ""
                Me.txtItemSubject_Att.ReadOnly = True
                Me.txtItemSubject_Att.BackColor = Me.txtReadonly.BackColor
                Me.txtItemSubject_Att.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AturDetailSalarySlip()
        Try
            If Me.chkSalarySlip.Checked = True Then
                Me.txtItemCCto_SS.ReadOnly = False
                Me.txtItemCCto_SS.BackColor = Me.txtPrimary.BackColor
                Me.txtItemFileNamePrefix_SS.ReadOnly = False
                Me.txtItemFileNamePrefix_SS.BackColor = Me.txtPrimary.BackColor
                Me.txtItemMessage_SS.ReadOnly = False
                Me.txtItemMessage_SS.BackColor = Me.txtPrimary.BackColor
                Me.txtItemSubject_SS.ReadOnly = False
                Me.txtItemSubject_SS.BackColor = Me.txtPrimary.BackColor
            Else
                Me.txtItemCCto_SS.ReadOnly = True
                Me.txtItemCCto_SS.BackColor = Me.txtReadonly.BackColor
                Me.txtItemCCto_SS.Text = ""
                Me.txtItemFileNamePrefix_SS.ReadOnly = True
                Me.txtItemFileNamePrefix_SS.BackColor = Me.txtReadonly.BackColor
                Me.txtItemFileNamePrefix_SS.Text = ""
                Me.txtItemMessage_SS.ReadOnly = True
                Me.txtItemMessage_SS.BackColor = Me.txtReadonly.BackColor
                Me.txtItemMessage_SS.Text = ""
                Me.txtItemSubject_SS.ReadOnly = True
                Me.txtItemSubject_SS.BackColor = Me.txtReadonly.BackColor
                Me.txtItemSubject_SS.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AturDetailPickupOrder()
        Try
            If Me.chkPickUpOrder.Checked = True Then
                Me.txtItemCCto_PickUp.ReadOnly = False
                Me.txtItemCCto_PickUp.BackColor = Me.txtPrimary.BackColor
                Me.txtItemFileNamePrefix_PickUp.ReadOnly = False
                Me.txtItemFileNamePrefix_PickUp.BackColor = Me.txtPrimary.BackColor
                Me.txtItemMessage_PickUp.ReadOnly = False
                Me.txtItemMessage_PickUp.BackColor = Me.txtPrimary.BackColor
                Me.txtItemSubject_PickUp.ReadOnly = False
                Me.txtItemSubject_PickUp.BackColor = Me.txtPrimary.BackColor
            Else
                Me.txtItemCCto_PickUp.ReadOnly = True
                Me.txtItemCCto_PickUp.BackColor = Me.txtReadonly.BackColor
                Me.txtItemCCto_PickUp.Text = ""
                Me.txtItemFileNamePrefix_PickUp.ReadOnly = True
                Me.txtItemFileNamePrefix_PickUp.BackColor = Me.txtReadonly.BackColor
                Me.txtItemFileNamePrefix_PickUp.Text = ""
                Me.txtItemMessage_PickUp.ReadOnly = True
                Me.txtItemMessage_PickUp.BackColor = Me.txtReadonly.BackColor
                Me.txtItemMessage_PickUp.Text = ""
                Me.txtItemSubject_PickUp.ReadOnly = True
                Me.txtItemSubject_PickUp.BackColor = Me.txtReadonly.BackColor
                Me.txtItemSubject_PickUp.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AturDetailOTOrder()
        Try
            If Me.chkOTOrder.Checked = True Then
                Me.txtItemCCto_OTOrder.ReadOnly = False
                Me.txtItemCCto_OTOrder.BackColor = Me.txtPrimary.BackColor
                Me.txtItemFileNamePrefix_OTOrder.ReadOnly = False
                Me.txtItemFileNamePrefix_OTOrder.BackColor = Me.txtPrimary.BackColor
                Me.txtItemMessage_OTOrder.ReadOnly = False
                Me.txtItemMessage_OTOrder.BackColor = Me.txtPrimary.BackColor
                Me.txtItemSubject_OTOrder.ReadOnly = False
                Me.txtItemSubject_OTOrder.BackColor = Me.txtPrimary.BackColor
            Else
                Me.txtItemCCto_OTOrder.ReadOnly = True
                Me.txtItemCCto_OTOrder.BackColor = Me.txtReadonly.BackColor
                Me.txtItemCCto_OTOrder.Text = ""
                Me.txtItemFileNamePrefix_OTOrder.ReadOnly = True
                Me.txtItemFileNamePrefix_OTOrder.BackColor = Me.txtReadonly.BackColor
                Me.txtItemFileNamePrefix_OTOrder.Text = ""
                Me.txtItemMessage_OTOrder.ReadOnly = True
                Me.txtItemMessage_OTOrder.BackColor = Me.txtReadonly.BackColor
                Me.txtItemMessage_OTOrder.Text = ""
                Me.txtItemSubject_OTOrder.ReadOnly = True
                Me.txtItemSubject_OTOrder.BackColor = Me.txtReadonly.BackColor
                Me.txtItemSubject_OTOrder.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.chkUseSSL.Enabled = False
            objCd.KendalikanControl(GroupBox1.Controls, True, False, True, False, False, False, False, False, True, False)
            objCd.KendalikanControl(GroupBox2.Controls, True, False, False, False, False, False, False, False, False, False)
            objCd.KendalikanControl(TabPage_ColLetter.Controls, True, False, False, False, False, False, False, False, False, False)
            objCd.KendalikanControl(TabPage_Attendance.Controls, True, False, False, False, False, False, False, False, False, False)
            objCd.KendalikanControl(TabPage_SS.Controls, True, False, False, False, False, False, False, False, False, False)
            objCd.KendalikanControl(TabPage_PickUp.Controls, True, False, False, False, False, False, False, False, False, False)
            objCd.KendalikanControl(TabPage_OTOrder.Controls, True, False, False, False, False, False, False, False, False, False)

            llbTestEmail1.Enabled = False
            llbTestEmail_Att.Enabled = False
            llbTestEmail_PickUp.Enabled = False
            llbTestEmail_SS.Enabled = False
            llbTestEmail_OTOrder.Enabled = False

            btnOK.Enabled = False
            btnCancel.Enabled = False
            btnFind.Enabled = True
            btnNew.Enabled = True
            btnEdit.Enabled = True
            btnDelete.Enabled = True
            btnExit.Enabled = True
            CancelButton = btnExit
            chkShowPassword.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            objCd.KosongkanControl(GroupBox1.Controls, True, False, False)
            objCd.KosongkanControl(GroupBox2.Controls, True, False, False)
            chkShowPassword.CheckState = CheckState.Unchecked
            objCd.KosongkanControl(TabPage_ColLetter.Controls, True, False, False)
            objCd.KosongkanControl(TabPage_Attendance.Controls, True, False, False)
            objCd.KosongkanControl(TabPage_SS.Controls, True, False, False)
            objCd.KosongkanControl(TabPage_PickUp.Controls, True, False, False)
            objCd.KosongkanControl(TabPage_OTOrder.Controls, True, False, False)

            txtPrimary.Text = "-1"
            txtPrimaryUser.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.chkCollectionLetter.Checked = True
            Me.chkPickUpOrder.Checked = True
            Me.chkSalarySlip.Checked = True
            Me.chkAttendance.Checked = True
            Me.chkOTOrder.Checked = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            With obj
                .Primary = txtPrimary.Text
                .AccountCode = txtAccountCode.Text
                .PrimaryUser = txtPrimaryUser.Text
                .Description = txtDescription.Text
                .EmailFrom = txtEmailFrom.Text
                .ReplyTo = txtReplyTo.Text
                .Username = txtUsername.Text
                .Password = txtPassword.Text
                .SmtpServer = txtSmtpServer.Text
                .SmtpPort = txtSmtpPort.Text
                .CreatedBy = UserCode
                .CreatedDate = Now

                If Me.chkAttendance.Checked = True Then
                    .UseAttendance = 1
                Else
                    .UseAttendance = 0
                End If
                If Me.chkCollectionLetter.Checked = True Then
                    .UseCollectionLetter = 1
                Else
                    .UseCollectionLetter = 0
                End If
                If Me.chkSalarySlip.Checked = True Then
                    .UseSalarySlip = 1
                Else
                    .UseSalarySlip = 0
                End If
                If Me.chkPickUpOrder.Checked = True Then
                    .UsePickUpOrder = 1
                Else
                    .UsePickUpOrder = 0
                End If
                If Me.chkOTOrder.Checked = True Then
                    .UseOTOrder = 1
                Else
                    .UseOTOrder = 0
                End If

                If Baru = True Then
                    .ModifiedBy = UserCode
                    .ModifiedDate = Now
                End If
                .UseSSL = Me.chkUseSSL.Checked
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim objUser As New ClsUser
        Try
            With obj
                txtPrimary.Text = .Primary
                txtAccountCode.Text = .AccountCode
                txtPrimaryUser.Text = .PrimaryUser
                txtUserCode.Text = objUser.GetKodeUser(.PrimaryUser)
                txtDescription.Text = .Description
                txtEmailFrom.Text = .EmailFrom
                txtReplyTo.Text = .ReplyTo
                txtUsername.Text = .Username
                txtPassword.Text = .Password
                txtSmtpServer.Text = .SmtpServer
                txtSmtpPort.Text = .SmtpPort

                If .UseAttendance = 1 Then
                    Me.chkAttendance.Checked = True
                Else
                    Me.chkAttendance.Checked = False
                End If
                If .UseCollectionLetter = 1 Then
                    Me.chkCollectionLetter.Checked = True
                Else
                    Me.chkCollectionLetter.Checked = False
                End If
                If .UseSalarySlip = 1 Then
                    Me.chkSalarySlip.Checked = True
                Else
                    Me.chkSalarySlip.Checked = False
                End If
                If .UsePickUpOrder = 1 Then
                    Me.chkPickUpOrder.Checked = True
                Else
                    Me.chkPickUpOrder.Checked = False
                End If
                If .UseOTOrder = 1 Then
                    Me.chkOTOrder.Checked = True
                Else
                    Me.chkOTOrder.Checked = False
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
                Me.chkUseSSL.Checked = .UseSSL
            End With
            Me.chkShowPassword.Checked = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection

        Try
            KoleksiNamaData.Add("Account Code")
            KoleksiTextBox.Add(txtAccountCode)
            KoleksiNamaData.Add("User Code")
            KoleksiTextBox.Add(txtUserCode)
            KoleksiNamaData.Add("Email from")
            KoleksiTextBox.Add(txtEmailFrom)
            KoleksiNamaData.Add("Reply to")
            KoleksiTextBox.Add(txtReplyTo)
            KoleksiNamaData.Add("Username")
            KoleksiTextBox.Add(txtUsername)
            KoleksiNamaData.Add("Password")
            KoleksiTextBox.Add(txtPassword)
            KoleksiNamaData.Add("SMTP Server")
            KoleksiTextBox.Add(txtSmtpServer)
            KoleksiNamaData.Add("SMTP Port")
            KoleksiTextBox.Add(txtSmtpPort)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            Me.txtItemCCto.Text = Me.txtItemCCto.Text.Trim
            Me.txtItemSubject.Text = Me.txtItemSubject.Text.Trim
            Me.txtItemMessage.Text = Me.txtItemMessage.Text.Trim
            If Me.chkCollectionLetter.Checked = True Then
                'pastikan tidak kosong
                If Me.txtItemSubject.Text = "" Then
                    MsgBox("Please specify email subject.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtItemSubject.Focus()
                    Return False
                End If
                If Me.txtItemCCto.Text <> "" Then
                    Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                    If Expression.IsMatch(Me.txtItemCCto.Text) = False Then
                        MsgBox("The cc email address is not valid.", MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedTab = Me.TabPage_PickUp
                        Me.txtItemCCto.Focus()
                        Expression = Nothing
                        Return False
                    End If
                    Expression = Nothing
                End If
                If Me.txtItemMessage.Text = "" Then
                    MsgBox("Please specify email message.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtItemMessage.Focus()
                    Return False
                End If
            Else
                If Me.txtItemSubject.Text <> "" Or Me.txtItemMessage.Text <> "" Or Me.txtItemCCto.Text <> "" Then
                    MsgBox("Please tick 'Collection Letter' check box if you want to allow this account to send collection letter emails." + vbCrLf + _
                    "If this account is not allowed to send collection letter, please make sure email subject and message text box are empty.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 0
                    If Me.txtItemSubject.Text <> "" Then
                        Me.txtItemSubject.Focus()
                    Else
                        Me.txtItemMessage.Focus()
                    End If
                    Return False
                End If
            End If

            Me.txtItemCCto_Att.Text = Me.txtItemCCto_Att.Text.Trim
            Me.txtItemSubject_Att.Text = Me.txtItemSubject_Att.Text.Trim
            Me.txtItemMessage_Att.Text = Me.txtItemMessage_Att.Text.Trim
            If Me.chkAttendance.Checked = True Then
                'pastikan tidak kosong
                If Me.txtItemSubject_Att.Text = "" Then
                    MsgBox("Please specify email subject.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtItemSubject_Att.Focus()
                    Return False
                End If
                If Me.txtItemCCto_Att.Text <> "" Then
                    Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                    If Expression.IsMatch(Me.txtItemCCto_Att.Text) = False Then
                        MsgBox("The cc email address is not valid.", MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedTab = Me.TabPage_Attendance
                        Me.txtItemCCto_Att.Focus()
                        Expression = Nothing
                        Return False
                    End If
                    Expression = Nothing
                End If
                If Me.txtItemMessage_Att.Text = "" Then
                    MsgBox("Please specify email message.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtItemMessage_Att.Focus()
                    Return False
                End If
            Else
                If Me.txtItemSubject_Att.Text <> "" Or Me.txtItemMessage_Att.Text <> "" Or Me.txtItemCCto_Att.Text <> "" Then
                    MsgBox("Please tick 'Attendance' check box if you want to allow this account to send attendance report emails." + vbCrLf + _
                    "If this account is not allowed to send attendance report, please make sure email subject and message text box are empty.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 1
                    If Me.txtItemSubject_Att.Text <> "" Then
                        Me.txtItemSubject_Att.Focus()
                    Else
                        Me.txtItemMessage_Att.Focus()
                    End If
                    Return False
                End If
            End If

            Me.txtItemCCto_SS.Text = Me.txtItemCCto_SS.Text.Trim
            Me.txtItemSubject_SS.Text = Me.txtItemSubject_SS.Text.Trim
            Me.txtItemMessage_SS.Text = Me.txtItemMessage_SS.Text.Trim
            If Me.chkSalarySlip.Checked = True Then
                'pastikan tidak kosong
                If Me.txtItemSubject_SS.Text = "" Then
                    MsgBox("Please specify email subject.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 2
                    Me.txtItemSubject_SS.Focus()
                    Return False
                End If
                If Me.txtItemMessage_SS.Text = "" Then
                    MsgBox("Please specify email message.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtItemMessage_SS.Focus()
                    Return False
                End If
                If Me.txtItemCCto_SS.Text <> "" Then
                    Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                    If Expression.IsMatch(Me.txtItemCCto_SS.Text) = False Then
                        MsgBox("The cc email address is not valid.", MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedTab = Me.TabPage_SS
                        Me.txtItemCCto_SS.Focus()
                        Expression = Nothing
                        Return False
                    End If
                    Expression = Nothing
                End If
            Else
                If Me.txtItemSubject_SS.Text <> "" Or Me.txtItemMessage_SS.Text <> "" Or Me.txtItemMessage_SS.Text <> "" Then
                    MsgBox("Please tick 'Salary Slip' check box if you want to allow this account to send salary slip emails." + vbCrLf + _
                    "If this account is not allowed to send salary slip, please make sure email subject and message text box are empty.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 1
                    If Me.txtItemSubject_SS.Text <> "" Then
                        Me.txtItemSubject_SS.Focus()
                    Else
                        Me.txtItemMessage_SS.Focus()
                    End If
                    Return False
                End If
            End If

            Me.txtItemCCto_PickUp.Text = Me.txtItemCCto_PickUp.Text.Trim
            Me.txtItemSubject_PickUp.Text = Me.txtItemSubject_PickUp.Text.Trim
            Me.txtItemMessage_PickUp.Text = Me.txtItemMessage_PickUp.Text.Trim
            If Me.chkPickUpOrder.Checked = True Then
                'pastikan tidak kosong
                If Me.txtItemSubject_PickUp.Text = "" Then
                    MsgBox("Please specify email subject.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPage_PickUp
                    Me.txtItemSubject_PickUp.Focus()
                    Return False
                End If
                If Me.txtItemMessage_PickUp.Text = "" Then
                    MsgBox("Please specify email message.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtItemMessage_PickUp.Focus()
                    Return False
                End If
                If Me.txtItemCCto_PickUp.Text <> "" Then
                    Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                    If Expression.IsMatch(Me.txtItemCCto_PickUp.Text) = False Then
                        MsgBox("The cc email address is not valid.", MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedTab = Me.TabPage_PickUp
                        Me.txtItemCCto_PickUp.Focus()
                        Expression = Nothing
                        Return False
                    End If
                    Expression = Nothing
                End If
            Else
                If Me.txtItemSubject_PickUp.Text <> "" Or Me.txtItemMessage_PickUp.Text <> "" Or Me.txtItemCCto_PickUp.Text <> "" Then
                    MsgBox("Please tick 'Salary Slip' check box if you want to allow this account to send salary slip emails." + vbCrLf + _
                    "If this account is not allowed to send salary slip, please make sure email subject and message text box are empty.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 1
                    If Me.txtItemSubject_PickUp.Text <> "" Then
                        Me.txtItemSubject_PickUp.Focus()
                    Else
                        Me.txtItemMessage_PickUp.Focus()
                    End If
                    Return False
                End If
            End If

            Me.txtItemCCto_OTOrder.Text = Me.txtItemCCto_OTOrder.Text.Trim
            Me.txtItemSubject_OTOrder.Text = Me.txtItemSubject_OTOrder.Text.Trim
            Me.txtItemMessage_OTOrder.Text = Me.txtItemMessage_OTOrder.Text.Trim
            If Me.chkOTOrder.Checked = True Then
                'pastikan tidak kosong
                If Me.txtItemSubject_OTOrder.Text = "" Then
                    MsgBox("Please specify email subject.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPage_OTOrder
                    Me.txtItemSubject_OTOrder.Focus()
                    Return False
                End If
                If Me.txtItemMessage_OTOrder.Text = "" Then
                    MsgBox("Please specify email message.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtItemMessage_OTOrder.Focus()
                    Return False
                End If
                If Me.txtItemCCto_OTOrder.Text <> "" Then
                    Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                    If Expression.IsMatch(Me.txtItemCCto_OTOrder.Text) = False Then
                        MsgBox("The cc email address is not valid.", MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedTab = Me.TabPage_OTOrder
                        Me.txtItemCCto_OTOrder.Focus()
                        Expression = Nothing
                        Return False
                    End If
                    Expression = Nothing
                End If
            Else
                If Me.txtItemSubject_OTOrder.Text <> "" Or Me.txtItemMessage_OTOrder.Text <> "" Or Me.txtItemCCto_OTOrder.Text <> "" Then
                    MsgBox("Please tick 'Salary Slip' check box if you want to allow this account to send salary slip emails." + vbCrLf + _
                    "If this account is not allowed to send salary slip, please make sure email subject and message text box are empty.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 1
                    If Me.txtItemSubject_OTOrder.Text <> "" Then
                        Me.txtItemSubject_OTOrder.Focus()
                    Else
                        Me.txtItemMessage_OTOrder.Focus()
                    End If
                    Return False
                End If
            End If

            'minimal salah satu harus cek list
            If Me.chkAttendance.Checked = False And Me.chkCollectionLetter.Checked = False And Me.chkSalarySlip.Checked = False And Me.chkPickUpOrder.Checked = False And Me.chkOTOrder.Checked = False Then
                MsgBox("Please specify at least 1 type of email to be sent.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Baru = False Then
                If Me.txtAccountCode.Text <> Me.obj.AccountCode Then
                    If Me.obj.ApakahSudahAda(Me.txtAccountCode.Text) = True Then
                        MsgBox("Account Code: " & Me.txtAccountCode.Text & " is already exist", MsgBoxStyle.Information, "Information")
                        Me.txtAccountCode.Focus()
                        Return False
                    End If
                End If
            Else
                If Me.obj.ApakahSudahAda(Me.txtAccountCode.Text) = True Then
                    MsgBox("Account Code: " & Me.txtAccountCode.Text & " is already exist", MsgBoxStyle.Information, "Information")
                    Me.txtAccountCode.Focus()
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function ApakahBolehEdit() As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahBolehHapus() As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FrmEmailSettings_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.txtItemMessage_OTOrder.MaxLength = 1000
            Baru = True
            btnNew.Focus()
            NonAktifkanDetail()
            KosongkanField()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub btnFindUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindUser.Click
        Dim frm As New FrmCariMasterUser
        Dim objUser As New ClsUser
        Try
            Me.Cursor = Cursors.WaitCursor
            frm.TxtCari.Text = txtUserCode.Text
            frm.ShowDialog()
            If frm.Kodeuser <> "" Then
                txtPrimaryUser.Text = objUser.GetPrimary(frm.Kodeuser)
                txtUserCode.Text = frm.Kodeuser
            Else
                txtPrimaryUser.Text = "-1"
                txtUserCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub txtUserCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUserCode.KeyDown
        Dim objUser As New ClsUser
        Try
            Me.Cursor = Cursors.WaitCursor
            If e.KeyCode = 13 Then
                If txtUserCode.Text <> "" Then
                    If objUser.ApakahSudahAda(txtUserCode.Text) = True Then
                        txtPrimaryUser.Text = objUser.GetPrimary(txtUserCode.Text)
                        txtUserCode.Text = objUser.GetUserCode(txtPrimaryUser.Text)
                        txtDescription.Focus()
                    Else
                        btnFindUser.PerformClick()
                    End If
                Else
                    txtPrimaryUser.Text = "-1"
                    txtUserCode.Clear()
                    txtUserCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub txtUserCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUserCode.Leave
        Dim objUser As New ClsUser
        Try
            Me.Cursor = Cursors.WaitCursor
            If txtUserCode.Text <> "" Then
                If objUser.ApakahSudahAda(txtUserCode.Text) = True Then
                    txtPrimaryUser.Text = objUser.GetPrimary(txtUserCode.Text)
                    txtUserCode.Text = objUser.GetUserCode(txtPrimaryUser.Text)
                    txtDescription.Focus()
                Else
                    btnFindUser.PerformClick()
                End If
            Else
                txtPrimaryUser.Text = "-1"
                txtUserCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub clbShowPassword_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowPassword.CheckedChanged
        Try
            If chkShowPassword.Checked = True Then
                txtPassword.PasswordChar = ""
            ElseIf chkShowPassword.Checked = False Then
                txtPassword.PasswordChar = "*"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PictureBox1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseEnter
        'Try
        '    Me.ToolTip1.Show("Type <<Date>> to insert Date on Message" & vbNewLine & _
        '    "Type <<Subject>> to insert Subject on Message" & vbNewLine & _
        '    "Type <<Attn>> to insert Attn on Message" & vbNewLine & _
        '    "Type <<CustCode>> to insert Customer Code on Message" & vbNewLine & _
        '    "Type <<CustName>> to insert Customer Name on Message" & vbNewLine & _
        '    "Type <<SignBy>> to insert SignBy on Message" & vbNewLine & _
        '    "Type <<Position>> to insert Position on Message", Me.PictureBox1)
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub PictureBox1_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseHover
        Try
            Me.ToolTip1.Show("Type <<Date>> to insert Date on Message" & vbNewLine & _
            "Type <<Subject>> to insert Subject on Message" & vbNewLine & _
            "Type <<Attn>> to insert Attn on Message" & vbNewLine & _
            "Type <<CustCode>> to insert Customer Code on Message" & vbNewLine & _
            "Type <<CustName>> to insert Customer Name on Message" & vbNewLine & _
            "Type <<SignBy>> to insert SignBy on Message" & vbNewLine & _
            "Type <<Position>> to insert Position on Message", Me.PictureBox1, 15000)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbTestEmail1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbTestEmail1.LinkClicked
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim frm As New FrmEmailSettings_Test

        Try
            If Me.chkCollectionLetter.Checked = False Then
                MsgBox("Please tick 'Collection Letter' check box if you want to allow this account to send collection letter emails.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.txtItemCCto.Text <> "" Then
                Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                If Expression.IsMatch(Me.txtItemCCto.Text) = False Then
                    MsgBox("The cc email address is not valid.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPage_ColLetter
                    Me.txtItemCCto.Focus()
                    Expression = Nothing
                    Exit Sub
                End If
                Expression = Nothing
            End If

            KoleksiTextBox.Add(txtUserCode)
            KoleksiNamaData.Add("Email from")
            KoleksiTextBox.Add(txtEmailFrom)
            KoleksiNamaData.Add("Reply to")
            KoleksiTextBox.Add(txtReplyTo)
            KoleksiNamaData.Add("Username")
            KoleksiTextBox.Add(txtUsername)
            KoleksiNamaData.Add("Password")
            KoleksiTextBox.Add(txtPassword)
            KoleksiNamaData.Add("SMTP Server")
            KoleksiTextBox.Add(txtSmtpServer)
            KoleksiNamaData.Add("SMTP Port")
            KoleksiTextBox.Add(txtSmtpPort)
            KoleksiNamaData.Add("Subject")
            KoleksiTextBox.Add(txtItemSubject)
            KoleksiNamaData.Add("Message")
            KoleksiTextBox.Add(txtItemMessage)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Exit Sub
            End If
            frm.BackColorAll = BackColorAll
            frm.EmailFrom = txtEmailFrom.Text
            frm.ReplyTo = txtReplyTo.Text
            frm.UserName = txtUsername.Text
            frm.Password = txtPassword.Text
            frm.SMTPServer = txtSmtpServer.Text
            frm.SMTPPort = txtSmtpPort.Text
            frm.CCto = txtItemCCto.Text
            frm.Subject = txtItemSubject.Text
            frm.Message = txtItemMessage.Text
            frm.Message = frm.Message.Replace("<<Date>>", Format(Today, "dd MMM yyyy"))
            frm.Message = frm.Message.Replace("<<Subject>>", txtItemSubject.Text)
            frm.Message = frm.Message.Replace("<<EmpName>>", "Mr. James. K")
            frm.Message = frm.Message.Replace("<<SignBy>>", "Stephanus. S")
            frm.Message = frm.Message.Replace("<<Position>>", "Collection")
            frm.Message = frm.Message.Replace("<<CustCode>>", "DE-001")
            frm.Message = frm.Message.Replace("<<CustName>>", "PT. ABC Jaya")
            frm.PrefixFilename = txtItemFileNamePrefix.Text
            frm.UseSSL = Me.chkUseSSL.Checked
            frm.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAccountCode.GotFocus, txtDescription.GotFocus, txtEmailFrom.GotFocus, _
    txtItemCCto.GotFocus, txtItemFileNamePrefix.GotFocus, txtItemMessage.GotFocus, txtItemSubject.GotFocus, txtPassword.GotFocus, txtReplyTo.GotFocus, _
    txtSmtpPort.GotFocus, txtSmtpServer.GotFocus, txtUserCode.GotFocus, txtUsername.GotFocus, _
    txtItemCCto_Att.GotFocus, txtItemFileNamePrefix_Att.GotFocus, txtItemMessage_Att.GotFocus, txtItemSubject_Att.GotFocus, _
    txtItemCCto_SS.GotFocus, txtItemFileNamePrefix_SS.GotFocus, txtItemMessage_SS.GotFocus, txtItemSubject_SS.GotFocus, _
    txtItemCCto_PickUp.GotFocus, txtItemFileNamePrefix_PickUp.GotFocus, txtItemMessage_PickUp.GotFocus, txtItemSubject_PickUp.GotFocus, _
    txtItemCCto_OTOrder.GotFocus, txtItemFileNamePrefix_OTOrder.GotFocus, txtItemMessage_OTOrder.GotFocus, txtItemSubject_OTOrder.GotFocus
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            If txt.Enabled = True Then
                If txt.ReadOnly = True Then
                    txt.BackColor = Me.txtReadonly.BackColor
                Else
                    txt.BackColor = BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAccountCode.LostFocus, txtDescription.LostFocus, _
    txtEmailFrom.LostFocus, txtItemCCto.LostFocus, txtItemFileNamePrefix.LostFocus, txtItemMessage.LostFocus, txtItemSubject.LostFocus, _
    txtPassword.LostFocus, txtReplyTo.LostFocus, txtSmtpPort.LostFocus, txtSmtpServer.LostFocus, txtUserCode.LostFocus, txtUsername.LostFocus, _
    txtItemCCto_Att.LostFocus, txtItemFileNamePrefix_Att.LostFocus, txtItemMessage_Att.LostFocus, txtItemSubject_Att.LostFocus, _
    txtItemCCto_SS.LostFocus, txtItemFileNamePrefix_SS.LostFocus, txtItemMessage_SS.LostFocus, txtItemSubject_SS.LostFocus, _
    txtItemCCto_PickUp.LostFocus, txtItemFileNamePrefix_PickUp.LostFocus, txtItemMessage_PickUp.LostFocus, txtItemSubject_PickUp.LostFocus, _
        txtItemCCto_OTOrder.LostFocus, txtItemFileNamePrefix_OTOrder.LostFocus, txtItemMessage_OTOrder.LostFocus, txtItemSubject_OTOrder.LostFocus
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            If txt.Enabled = True Then
                If txt.ReadOnly = False Then
                    txt.BackColor = Me.txtPrimary.BackColor
                Else
                    txt.BackColor = Me.txtReadonly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If ApakahDataBenar() = True Then
                SalinKeObjek()
                If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    Me.Cursor = Cursors.WaitCursor
                    If Baru = True Then
                        If obj.SimpanBaru(txtItemCCto.Text, txtItemSubject.Text, txtItemMessage.Text, txtItemFileNamePrefix.Text, _
                        Me.txtItemCCto_Att.Text, Me.txtItemSubject_Att.Text, Me.txtItemMessage_Att.Text, Me.txtItemFileNamePrefix_Att.Text, _
                        Me.txtItemCCto_SS.Text, Me.txtItemSubject_SS.Text, Me.txtItemMessage_SS.Text, Me.txtItemFileNamePrefix_SS.Text, _
                        Me.txtItemCCto_PickUp.Text, Me.txtItemSubject_PickUp.Text, Me.txtItemMessage_PickUp.Text, Me.txtItemFileNamePrefix_PickUp.Text, _
                        Me.txtItemCCto_OTOrder.Text, Me.txtItemSubject_OTOrder.Text, Me.txtItemMessage_OTOrder.Text, Me.txtItemFileNamePrefix_OTOrder.Text) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            SalinKeField()
                            NonAktifkanDetail()
                            btnFind.Focus()
                        Else
                            MsgBox("Save data failed", MsgBoxStyle.Information, "Information")
                            btnOK.Focus()
                        End If
                    Else
                        If obj.SimpanPerubahan(txtItemCCto.Text, txtItemSubject.Text, txtItemMessage.Text, txtItemFileNamePrefix.Text, _
                        Me.txtItemCCto_Att.Text, Me.txtItemSubject_Att.Text, Me.txtItemMessage_Att.Text, Me.txtItemMessage_Att.Text, _
                        Me.txtItemCCto_SS.Text, Me.txtItemSubject_SS.Text, Me.txtItemMessage_SS.Text, Me.txtItemFileNamePrefix_SS.Text, _
                        Me.txtItemCCto_PickUp.Text, Me.txtItemSubject_PickUp.Text, Me.txtItemMessage_PickUp.Text, Me.txtItemFileNamePrefix_PickUp.Text, _
                        Me.txtItemCCto_OTOrder.Text, Me.txtItemSubject_OTOrder.Text, Me.txtItemMessage_OTOrder.Text, Me.txtItemFileNamePrefix_OTOrder.Text) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            SalinKeField()
                            NonAktifkanDetail()
                            btnFind.Focus()
                        Else
                            MsgBox("Save data failed", MsgBoxStyle.Information, "Information")
                            btnOK.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If txtAccountCode.Text <> "" Or Me.txtUserCode.Text <> "" Or txtDescription.Text <> "" Then
                If MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.obj.CariDetail(Me.txtPrimary.Text, Me)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                If Baru = True Then
                    Me.KosongkanField()
                Else
                    Me.obj.CariDetail(Me.txtPrimary.Text, Me)
                    Me.SalinKeField()
                End If
            End If
            Me.NonAktifkanDetail()
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim frm As New FrmEmailSettings_Cari
        Try
            frm.BackColorAll = Me.BackColorAll
            frm.ShowDialog()
            If frm.primaryAccount > 0 Then
                If obj.CariDetail(frm.primaryAccount, Me) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm.Dispose()
        End Try
    End Sub
    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.Baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.txtAccountCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            txtAccountCode.Focus()
            Me.chkAttendance.Checked = True
            Me.chkSalarySlip.Checked = True
            Me.chkCollectionLetter.Checked = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If txtPrimary.Text = "-1" Then
                btnFind.PerformClick()
            End If
            If txtPrimary.Text <> "-1" Then
                Baru = False
                AktifkanDetail()
                txtAccountCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If txtPrimary.Text = "-1" Then
                btnFind.PerformClick()
            End If
            If txtPrimary.Text <> "-1" Then
                If ApakahBolehHapus() = True Then
                    If MsgBox("Are you sure?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                        Me.Cursor = Cursors.WaitCursor
                        If obj.Hapus(txtPrimary.Text) = True Then
                            MsgBox("Delete deleted", MsgBoxStyle.Information, "Information")
                            KosongkanField()
                            NonAktifkanDetail()
                            btnNew.Focus()
                        Else
                            MsgBox("Delete data failed", MsgBoxStyle.Information, "Information")
                            btnDelete.Focus()
                        End If
                    Else
                        btnDelete.Focus()
                    End If
                Else
                    MsgBox("You cannot delete this data", MsgBoxStyle.Information, "Information")
                    btnDelete.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub FrmEmailSettings_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If (Me.txtAccountCode.Enabled = True) And (Me.txtAccountCode.Text <> "" Or Me.txtUserCode.Text <> "" Or Me.txtDescription.Text <> "") Then
                If MsgBox("Save data before exit?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    If ApakahDataBenar() = True Then
                        SalinKeObjek()
                        Me.Cursor = Cursors.WaitCursor
                        If Baru = True Then
                            If obj.SimpanBaru(txtItemCCto.Text, txtItemSubject.Text, txtItemMessage.Text, txtItemFileNamePrefix.Text, _
                            Me.txtItemCCto_Att.Text, Me.txtItemSubject_Att.Text, Me.txtItemMessage_Att.Text, Me.txtItemMessage_Att.Text, _
                            Me.txtItemCCto_SS.Text, Me.txtItemSubject_SS.Text, Me.txtItemMessage_SS.Text, Me.txtItemFileNamePrefix_SS.Text, _
                            Me.txtItemCCto_PickUp.Text, Me.txtItemSubject_PickUp.Text, Me.txtItemMessage_PickUp.Text, Me.txtItemFileNamePrefix_PickUp.Text, _
                            Me.txtItemCCto_OTOrder.Text, Me.txtItemSubject_OTOrder.Text, Me.txtItemMessage_OTOrder.Text, Me.txtItemFileNamePrefix_OTOrder.Text) = True Then
                                SalinKeField()
                                NonAktifkanDetail()
                                btnFind.Focus()
                            Else
                                MsgBox("Save data failed", MsgBoxStyle.Information, "Information")
                                btnOK.Focus()
                                e.Cancel = True
                            End If
                        Else
                            If obj.SimpanPerubahan(txtItemCCto.Text, txtItemSubject.Text, txtItemMessage.Text, txtItemFileNamePrefix.Text, _
                            Me.txtItemCCto_Att.Text, Me.txtItemSubject_Att.Text, Me.txtItemMessage_Att.Text, Me.txtItemMessage_Att.Text, _
                            Me.txtItemCCto_SS.Text, Me.txtItemSubject_SS.Text, Me.txtItemMessage_SS.Text, Me.txtItemFileNamePrefix_SS.Text, _
                            Me.txtItemCCto_PickUp.Text, Me.txtItemSubject_PickUp.Text, Me.txtItemMessage_PickUp.Text, Me.txtItemFileNamePrefix_PickUp.Text, _
                            Me.txtItemCCto_OTOrder.Text, Me.txtItemSubject_OTOrder.Text, Me.txtItemMessage_OTOrder.Text, Me.txtItemFileNamePrefix_OTOrder.Text) = True Then
                                SalinKeField()
                                NonAktifkanDetail()
                                btnFind.Focus()
                            Else
                                MsgBox("Save data failed", MsgBoxStyle.Information, "Information")
                                btnOK.Focus()
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                        txtAccountCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtEmailFrom_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEmailFrom.Leave
        Try
            If txtEmailFrom.Text <> "" Then
                Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                If Expression.IsMatch(txtEmailFrom.Text) = False Then
                    MsgBox("The email address is not valid.", MsgBoxStyle.Information, "Information")
                    txtEmailFrom.Focus()
                    txtEmailFrom.SelectAll()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtReplyTo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtReplyTo.Leave
        Try
            If txtReplyTo.Text <> "" Then
                Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                If Expression.IsMatch(txtReplyTo.Text) = False Then
                    MsgBox("The email address is not valid.", MsgBoxStyle.Information, "Information")
                    txtReplyTo.Focus()
                    txtReplyTo.SelectAll()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub txtItemCCto_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtItemCCto.Leave
        Try
            If txtItemCCto.Text <> "" Then
                Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                If Expression.IsMatch(txtItemCCto.Text) = False Then
                    MsgBox("The email address is not valid.", MsgBoxStyle.Information, "Information")
                    txtItemCCto.Focus()
                    txtItemCCto.SelectAll()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub txtSmtpPort_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSmtpPort.KeyPress
        Try
            If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabPage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage_ColLetter.Click

    End Sub

    Private Sub chkCollectionLetter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCollectionLetter.CheckedChanged
        Try
            Me.AturDetailCollectionLetter()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAttendance_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAttendance.CheckedChanged
        Try
            Me.AturDetailAttendance()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkSalarySlip_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSalarySlip.CheckedChanged
        Try
            Me.AturDetailSalarySlip()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PictureBox_Att_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox_Att.MouseHover
        Try
            Me.ToolTip1.Show("Type <<Date>> to insert Date on Message" & vbNewLine & _
            "Type <<Subject>> to insert Subject on Message" & vbNewLine & _
            "Type <<EmpName>> to insert Employee Name on Message" & vbNewLine & _
            "Type <<SignBy>> to insert SignBy on Message" & vbNewLine & _
            "Type <<Position>> to insert Position on Message", Me.PictureBox_Att, 15000)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PictureBox_SS_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox_SS.MouseHover
        Try
            Me.ToolTip1.Show("Type <<Date>> to insert Date on Message" & vbNewLine & _
            "Type <<Subject>> to insert Subject on Message" & vbNewLine & _
            "Type <<EmpName>> to insert Employee Name on Message" & vbNewLine & _
            "Type <<SignBy>> to insert SignBy on Message" & vbNewLine & _
            "Type <<Position>> to insert Position on Message", Me.PictureBox_SS, 15000)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PictureBox1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseLeave
        Me.ToolTip1.Hide(Me.PictureBox1)
    End Sub

    Private Sub PictureBox_Att_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox_Att.MouseLeave
        Me.ToolTip1.Hide(Me.PictureBox_Att)
    End Sub

    Private Sub PictureBox_SS_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox_SS.MouseLeave
        Me.ToolTip1.Hide(Me.PictureBox_SS)
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub llbTestEmail_Att_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbTestEmail_Att.LinkClicked
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim frm As New FrmEmailSettings_Test

        Try
            If Me.chkAttendance.Checked = False Then
                MsgBox("Please tick 'Attendance' check box if you want to allow this account to send attendance report emails.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.txtItemCCto_Att.Text <> "" Then
                Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                If Expression.IsMatch(Me.txtItemCCto_Att.Text) = False Then
                    MsgBox("The cc email address is not valid.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPage_Attendance
                    Me.txtItemCCto_Att.Focus()
                    Expression = Nothing
                    Exit Sub
                End If
                Expression = Nothing
            End If

            KoleksiNamaData.Add("User Code")
            KoleksiTextBox.Add(txtUserCode)
            KoleksiNamaData.Add("Email from")
            KoleksiTextBox.Add(txtEmailFrom)
            KoleksiNamaData.Add("Reply to")
            KoleksiTextBox.Add(txtReplyTo)
            KoleksiNamaData.Add("Username")
            KoleksiTextBox.Add(txtUsername)
            KoleksiNamaData.Add("Password")
            KoleksiTextBox.Add(txtPassword)
            KoleksiNamaData.Add("SMTP Server")
            KoleksiTextBox.Add(txtSmtpServer)
            KoleksiNamaData.Add("SMTP Port")
            KoleksiTextBox.Add(txtSmtpPort)
            KoleksiNamaData.Add("Subject")
            KoleksiTextBox.Add(txtItemSubject_Att)
            KoleksiNamaData.Add("Message")
            KoleksiTextBox.Add(txtItemMessage_Att)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Exit Sub
            End If
            frm.BackColorAll = BackColorAll
            frm.EmailFrom = txtEmailFrom.Text
            frm.ReplyTo = txtReplyTo.Text
            frm.UserName = txtUsername.Text
            frm.Password = txtPassword.Text
            frm.SMTPServer = txtSmtpServer.Text
            frm.SMTPPort = txtSmtpPort.Text
            frm.CCto = txtItemCCto_Att.Text
            frm.Subject = txtItemSubject_Att.Text
            frm.Message = txtItemMessage_Att.Text
            'ganti kata2 yang merupakan kode dengan menggunakan asumsi sembarang
            frm.Message = frm.Message.Replace("<<Date>>", Format(Today, "dd MMM yyyy"))
            frm.Message = frm.Message.Replace("<<Subject>>", txtItemSubject_Att.Text)
            frm.Message = frm.Message.Replace("<<EmpName>>", "Mr. Dadang")
            frm.Message = frm.Message.Replace("<<SignBy>>", "Joko. S")
            frm.Message = frm.Message.Replace("<<Position>>", "HR Staff")
            frm.Message = frm.Message.Replace("<<CustCode>>", "DU-001")
            frm.Message = frm.Message.Replace("<<CustName>>", "PT. Djakartindo")
            frm.PrefixFilename = txtItemFileNamePrefix_Att.Text
            frm.UseSSL = Me.chkUseSSL.Checked
            frm.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbTestEmail_SS_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbTestEmail_SS.LinkClicked
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim frm As New FrmEmailSettings_Test

        Try
            If Me.chkSalarySlip.Checked = False Then
                MsgBox("Please tick 'Salary Slip' check box if you want to allow this account to send salary slip emails.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.txtItemCCto_SS.Text <> "" Then
                Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                If Expression.IsMatch(Me.txtItemCCto_SS.Text) = False Then
                    MsgBox("The cc email address is not valid.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPage_SS
                    Me.txtItemCCto_SS.Focus()
                    Expression = Nothing
                    Exit Sub
                End If
                Expression = Nothing
            End If

            KoleksiNamaData.Add("User Code")
            KoleksiTextBox.Add(txtUserCode)
            KoleksiNamaData.Add("Email from")
            KoleksiTextBox.Add(txtEmailFrom)
            KoleksiNamaData.Add("Reply to")
            KoleksiTextBox.Add(txtReplyTo)
            KoleksiNamaData.Add("Username")
            KoleksiTextBox.Add(txtUsername)
            KoleksiNamaData.Add("Password")
            KoleksiTextBox.Add(txtPassword)
            KoleksiNamaData.Add("SMTP Server")
            KoleksiTextBox.Add(txtSmtpServer)
            KoleksiNamaData.Add("SMTP Port")
            KoleksiTextBox.Add(txtSmtpPort)
            KoleksiNamaData.Add("Subject")
            KoleksiTextBox.Add(txtItemSubject_SS)
            KoleksiNamaData.Add("Message")
            KoleksiTextBox.Add(txtItemMessage_SS)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Exit Sub
            End If
            frm.BackColorAll = BackColorAll
            frm.EmailFrom = txtEmailFrom.Text
            frm.ReplyTo = txtReplyTo.Text
            frm.UserName = txtUsername.Text
            frm.Password = txtPassword.Text
            frm.SMTPServer = txtSmtpServer.Text
            frm.SMTPPort = txtSmtpPort.Text
            frm.CCto = txtItemCCto_SS.Text
            frm.Subject = txtItemSubject_SS.Text
            frm.Message = txtItemMessage_SS.Text
            'ganti kata2 yang merupakan kode dengan menggunakan asumsi sembarang
            frm.Message = frm.Message.Replace("<<Date>>", Format(Today, "dd MMM yyyy"))
            frm.Message = frm.Message.Replace("<<Subject>>", txtItemSubject_SS.Text)
            frm.Message = frm.Message.Replace("<<EmpName>>", "Mr. Dadang")
            frm.Message = frm.Message.Replace("<<SignBy>>", "Joko. S")
            frm.Message = frm.Message.Replace("<<Position>>", "HR Staff")
            frm.Message = frm.Message.Replace("<<CustCode>>", "DU-001")
            frm.Message = frm.Message.Replace("<<CustName>>", "PT. Djakartindo")
            frm.PrefixFilename = txtItemFileNamePrefix_SS.Text
            frm.UseSSL = Me.chkUseSSL.Checked
            frm.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkPickUpOrder_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPickUpOrder.CheckedChanged
        Try
            Me.AturDetailPickupOrder()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkOTOrder_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOTOrder.CheckedChanged
        Try
            Me.AturDetailOTOrder()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PictureBox_PickUp_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox_PickUp.MouseHover
        Try
            Me.ToolTip1.Show("Type <<Date>> to insert Date on Message" & vbNewLine & _
            "Type <<Subject>> to insert Subject on Message" & vbNewLine & _
            "Type <<Attn>> to insert Attn on Message" & vbNewLine & _
            "Type <<RefPartyCode>> to insert Reference Party Code on Message" & vbNewLine & _
            "Type <<RefPartyName>> to insert Reference Party Name on Message" & vbNewLine & _
            "Type <<SignBy>> to insert SignBy on Message" & vbNewLine & _
            "Type <<Position>> to insert Position on Message", Me.PictureBox_PickUp, 15000)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PictureBox_PickUpMouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox_PickUp.MouseLeave
        Me.ToolTip1.Hide(Me.PictureBox_PickUp)
    End Sub

    Private Sub PictureBox_OTOrder_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox_OTOrder.MouseHover
        Try
            Me.ToolTip1.Show("Type <<Date>> to insert Date on Message" & vbNewLine & _
            "Type <<Subject>> to insert Subject on Message" & vbNewLine & _
            "Type <<ApprovedBy>> to insert ApprovedBy on Message" & vbNewLine & _
            "Type <<SignBy>> to insert SignBy on Message" & vbNewLine & _
            "Type <<Position>> to insert Position on Message", Me.PictureBox_OTOrder, 15000)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PictureBox_OTOrder_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox_OTOrder.MouseLeave
        Me.ToolTip1.Hide(Me.PictureBox_OTOrder)
    End Sub

    Private Sub llbTestEmail_PickUp_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbTestEmail_PickUp.LinkClicked
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim frm As New FrmEmailSettings_Test

        Try
            If Me.chkPickUpOrder.Checked = False Then
                MsgBox("Please tick 'Pick Up Order' check box if you want to allow this account to send pick up order emails.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.txtItemCCto_PickUp.Text <> "" Then
                Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                If Expression.IsMatch(Me.txtItemCCto_PickUp.Text) = False Then
                    MsgBox("The cc email address is not valid.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPage_PickUp
                    Me.txtItemCCto_PickUp.Focus()
                    Expression = Nothing
                    Exit Sub
                End If
                Expression = Nothing
            End If

            KoleksiNamaData.Add("User Code")
            KoleksiTextBox.Add(txtUserCode)
            KoleksiNamaData.Add("Email from")
            KoleksiTextBox.Add(txtEmailFrom)
            KoleksiNamaData.Add("Reply to")
            KoleksiTextBox.Add(txtReplyTo)
            KoleksiNamaData.Add("Username")
            KoleksiTextBox.Add(txtUsername)
            KoleksiNamaData.Add("Password")
            KoleksiTextBox.Add(txtPassword)
            KoleksiNamaData.Add("SMTP Server")
            KoleksiTextBox.Add(txtSmtpServer)
            KoleksiNamaData.Add("SMTP Port")
            KoleksiTextBox.Add(txtSmtpPort)
            KoleksiNamaData.Add("Subject")
            KoleksiTextBox.Add(txtItemSubject_PickUp)
            KoleksiNamaData.Add("Message")
            KoleksiTextBox.Add(txtItemMessage_PickUp)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Exit Sub
            End If
            frm.BackColorAll = BackColorAll
            frm.EmailFrom = txtEmailFrom.Text
            frm.ReplyTo = txtReplyTo.Text
            frm.UserName = txtUsername.Text
            frm.Password = txtPassword.Text
            frm.SMTPServer = txtSmtpServer.Text
            frm.SMTPPort = txtSmtpPort.Text
            frm.CCto = txtItemCCto_PickUp.Text
            frm.Subject = txtItemSubject_PickUp.Text
            frm.Message = txtItemMessage_PickUp.Text
            'ganti kata2 yang merupakan kode dengan menggunakan asumsi sembarang
            frm.Message = frm.Message.Replace("<<Date>>", Format(Today, "dd MMM yyyy"))
            frm.Message = frm.Message.Replace("<<Subject>>", txtItemSubject_PickUp.Text)
            frm.Message = frm.Message.Replace("<<EmpName>>", "Mr. Dadang")
            frm.Message = frm.Message.Replace("<<SignBy>>", "Joko. S")
            frm.Message = frm.Message.Replace("<<Position>>", "Logistik Staff")
            frm.Message = frm.Message.Replace("<<CustCode>>", "DU-001")
            frm.Message = frm.Message.Replace("<<CustName>>", "PT. Djakartindo")
            frm.Message = frm.Message.Replace("<<RefPartyCode>>", "AA-001")
            frm.Message = frm.Message.Replace("<<RefPartyName>>", "Gudang Bahan Baku")

            frm.PrefixFilename = txtItemFileNamePrefix_PickUp.Text
            frm.UseSSL = Me.chkUseSSL.Checked
            frm.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbTestEmail_OTOrder_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbTestEmail_OTOrder.LinkClicked
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim frm As New FrmEmailSettings_Test

        Try
            If Me.chkOTOrder.Checked = False Then
                MsgBox("Please tick 'Overtime Order' check box if you want to allow this account to send overtime order emails.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.txtItemCCto_OTOrder.Text <> "" Then
                Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                If Expression.IsMatch(Me.txtItemCCto_OTOrder.Text) = False Then
                    MsgBox("The cc email address is not valid.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPage_OTOrder
                    Me.txtItemCCto_OTOrder.Focus()
                    Expression = Nothing
                    Exit Sub
                End If
                Expression = Nothing
            End If

            KoleksiNamaData.Add("User Code")
            KoleksiTextBox.Add(txtUserCode)
            KoleksiNamaData.Add("Email from")
            KoleksiTextBox.Add(txtEmailFrom)
            KoleksiNamaData.Add("Reply to")
            KoleksiTextBox.Add(txtReplyTo)
            KoleksiNamaData.Add("Username")
            KoleksiTextBox.Add(txtUsername)
            KoleksiNamaData.Add("Password")
            KoleksiTextBox.Add(txtPassword)
            KoleksiNamaData.Add("SMTP Server")
            KoleksiTextBox.Add(txtSmtpServer)
            KoleksiNamaData.Add("SMTP Port")
            KoleksiTextBox.Add(txtSmtpPort)
            KoleksiNamaData.Add("Subject")
            KoleksiTextBox.Add(txtItemSubject_OTOrder)
            KoleksiNamaData.Add("Message")
            KoleksiTextBox.Add(txtItemMessage_OTOrder)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Exit Sub
            End If
            frm.BackColorAll = BackColorAll
            frm.EmailFrom = txtEmailFrom.Text
            frm.ReplyTo = txtReplyTo.Text
            frm.UserName = txtUsername.Text
            frm.Password = txtPassword.Text
            frm.SMTPServer = txtSmtpServer.Text
            frm.SMTPPort = txtSmtpPort.Text
            frm.CCto = txtItemCCto_OTOrder.Text
            frm.Subject = txtItemSubject_OTOrder.Text
            frm.Message = txtItemMessage_OTOrder.Text
            'ganti kata2 yang merupakan kode dengan menggunakan asumsi sembarang
            frm.Message = frm.Message.Replace("<<Date>>", Format(Today, "dd MMM yyyy"))
            frm.Message = frm.Message.Replace("<<Subject>>", txtItemSubject_OTOrder.Text)
            frm.Message = frm.Message.Replace("<<ApprovedBy>>", "Mr. Dadang")
            frm.Message = frm.Message.Replace("<<SignBy>>", "Joko. S")
            frm.Message = frm.Message.Replace("<<Position>>", "HR Staff")
            'frm.Message = frm.Message.Replace("<<CustCode>>", "DU-001")
            'frm.Message = frm.Message.Replace("<<CustName>>", "PT. Djakartindo")
            frm.PrefixFilename = txtItemFileNamePrefix_OTOrder.Text
            frm.UseSSL = Me.chkUseSSL.Checked
            frm.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class