''@rosadinnn
''2015-04-15

Option Explicit On
Option Strict Off
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO
Imports Microsoft.Win32
Imports System.Net

Public Class ClsEmailSettings
    Dim ObjCD As New Core_Module.ClsCekData

    Dim passPhrase As String = "SMESS_EMAIL_Pas5pr@se"        ' can be any string
    Dim initVector As String = "@1B2S3D4M5F6g7H8" ' must be 16 bytes

    Public Enum EnumTipeTransaksiEmail As Integer
        CollectionLetter
        Invoice
        Attendance
        SalarySlip
        PickUpOrder
        OTOrder
    End Enum
    Public Enum EnumEmailSearchCriteria As Integer
        ByAccountCode
        ByUserCode
        ByEmailFrom
        ByDescription
    End Enum

    Private mPrimary As Double = -1
    Private mAccountCode As String = ""
    Private mPrimaryUser As Double = -1
    Private mDescription As String = ""
    Private mEmailFrom As String = ""
    Private mReplyTo As String = ""
    Private mUsername As String = ""
    Private mPassword As String = ""
    Private mSmtpServer As String = ""
    Private mSmtpPort As String = ""
    Private mCreatedBy As String = ""
    Private mCreatedDate As Date = Now
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date = Now
    Private mTimestamp As Date = Now
    Private mUseSSL As Integer = 0
    Private mUseCollectionLetter As Integer = 1
    Private mUseAttendance As Integer = 1
    Private mUseSalarySlip As Integer = 1
    Private mUsePickUpOrder As Integer = 1
    Private mUseOTOrder As Integer = 1

    Public Function GetPassPhrase() As String
        Return Me.passPhrase
    End Function
    Public Function GetIniVector() As String
        Return Me.initVector
    End Function
    Public Property UseCollectionLetter() As Integer
        Get
            Return Me.mUseCollectionLetter
        End Get
        Set(ByVal value As Integer)
            mUseCollectionLetter = value
        End Set
    End Property

    Public Property UseAttendance() As Integer
        Get
            Return Me.mUseAttendance
        End Get
        Set(ByVal value As Integer)
            mUseAttendance = value
        End Set
    End Property

    Public Property UseSalarySlip() As Integer
        Get
            Return Me.mUseSalarySlip
        End Get
        Set(ByVal value As Integer)
            mUseSalarySlip = value
        End Set
    End Property

    Public Property UsePickUpOrder() As Integer
        Get
            Return Me.mUsePickUpOrder
        End Get
        Set(ByVal value As Integer)
            mUsePickUpOrder = value
        End Set
    End Property

    Public Property UseOTOrder() As Integer
        Get
            Return Me.mUseOTOrder
        End Get
        Set(ByVal value As Integer)
            mUseOTOrder = value
        End Set
    End Property

    Public Property UseSSL() As Boolean
        Get
            If Me.mUseSSL = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                Me.mUseSSL = 1
            Else
                Me.mUseSSL = 0
            End If
        End Set
    End Property

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property
    Public Property AccountCode() As String
        Get
            Return mAccountCode
        End Get
        Set(ByVal value As String)
            mAccountCode = value
        End Set
    End Property
    Public Property PrimaryUser() As Double
        Get
            Return mPrimaryUser
        End Get
        Set(ByVal value As Double)
            mPrimaryUser = value
        End Set
    End Property
    Public Property Description() As String
        Get
            Return mDescription
        End Get
        Set(ByVal value As String)
            mDescription = value
        End Set
    End Property
    Public Property EmailFrom() As String
        Get
            Return mEmailFrom
        End Get
        Set(ByVal value As String)
            mEmailFrom = value
        End Set
    End Property
    Public Property ReplyTo() As String
        Get
            Return mReplyTo
        End Get
        Set(ByVal value As String)
            mReplyTo = value
        End Set
    End Property
    Public Property Username() As String
        Get
            Return mUsername
        End Get
        Set(ByVal value As String)
            mUsername = value
        End Set
    End Property
    Public Property Password() As String
        Get
            Return mPassword
        End Get
        Set(ByVal value As String)
            mPassword = value
        End Set
    End Property
    Public Property SmtpServer() As String
        Get
            Return mSmtpServer
        End Get
        Set(ByVal value As String)
            mSmtpServer = value
        End Set
    End Property
    Public Property SmtpPort() As String
        Get
            Return mSmtpPort
        End Get
        Set(ByVal value As String)
            mSmtpPort = value
        End Set
    End Property
    Public Property CreatedBy() As String
        Get
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property
    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property
    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property
    Public Property ModifiedDate() As Date
        Get
            Return mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property
    Public Property Timestamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property
    Public Function ApakahDataBenar() As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
    Public Function ApakahBolehEdit() As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
    Public Function ApakahBolehHapus() As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function SimpanBaru(ByVal vCcTo As String, ByVal vSubject As String, ByVal vMessage As String, ByVal vFileNamePrefix As String, _
    ByVal vCcTo_Attendance As String, ByVal vSubject_Attendance As String, ByVal vMessage_Attendance As String, ByVal vFileNamePrefix_Attendance As String, _
    ByVal vCcTo_SS As String, ByVal vSubject_SS As String, ByVal vMessage_SS As String, ByVal vFileNamePrefix_SS As String, _
    ByVal vCcTo_PickUp As String, ByVal vSubject_PickUp As String, ByVal vMessage_PickUp As String, ByVal vFileNamePrefix_PickUp As String, _
    ByVal vCcTo_OT As String, ByVal vSubject_OT As String, ByVal vMessage_OT As String, ByVal vFileNamePrefix_OT As String) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettingsTableAdapter
        Dim ObjEncrypt As New Core_Module.ClsEncrypt(Me.passPhrase + AccountCode.ToString + PrimaryUser.ToString, Me.initVector)
        Dim EncryptedPass As String = ""

        Try
            EncryptedPass = ObjEncrypt.Encrypt(Password)
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Insert(AccountCode, PrimaryUser, Description, EmailFrom, ReplyTo, Username, EncryptedPass, _
            SmtpServer, SmtpPort, CreatedBy, CreatedDate, ModifiedBy, ModifiedDate, Now, Me.mUseSSL, Me.mUseCollectionLetter, _
            Me.mUseAttendance, Me.mUseSalarySlip, Me.mUsePickUpOrder, Me.mUseOTOrder) > 0 Then
                Primary = GetPrimaryByCode(AccountCode)
                If SimpanItem(Primary, vCcTo, vSubject, vMessage, vFileNamePrefix, _
                vCcTo_Attendance, vSubject_Attendance, vMessage_Attendance, vFileNamePrefix_Attendance, _
                vCcTo_SS, vSubject_SS, vMessage_SS, vFileNamePrefix_SS, _
                vCcTo_PickUp, vSubject_PickUp, vMessage_PickUp, vFileNamePrefix_PickUp, _
                vCcTo_OT, vSubject_OT, vMessage_OT, vFileNamePrefix_OT) = True Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            ObjEncrypt = Nothing
        End Try
    End Function

    Private Function SimpanItem(ByVal vPrimaryTrans As Double, ByVal vCcTo As String, ByVal vSubject As String, ByVal vMessage As String, ByVal vFileNamePrefix As String, _
    ByVal vCcTo_Attendance As String, ByVal vSubject_Attendance As String, ByVal vMessage_Attendance As String, ByVal vFileNamePrefix_Attendance As String, _
    ByVal vCcTo_SS As String, ByVal vSubject_SS As String, ByVal vMessage_SS As String, ByVal vFileNamePrefix_SS As String, _
    ByVal vCcTo_PickUp As String, ByVal vSubject_PickUp As String, ByVal vMessage_PickUp As String, ByVal vFileNamePrefix_PickUp As String, _
    ByVal vCcTo_OT As String, ByVal vSubject_OT As String, ByVal vMessage_OT As String, ByVal vFileNamePrefix_OT As String) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettings_itemTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            'collection letter
            If dta.Insert(vPrimaryTrans, ClsEmailSettings.EnumTipeTransaksiEmail.CollectionLetter, vCcTo, vSubject, vMessage, vFileNamePrefix) = 0 Then
                Return False
            End If
            'attendance
            If dta.Insert(vPrimaryTrans, ClsEmailSettings.EnumTipeTransaksiEmail.Attendance, vCcTo_Attendance, vSubject_Attendance, vMessage_Attendance, vFileNamePrefix_Attendance) = 0 Then
                Return False
            End If
            'salary slip
            If dta.Insert(vPrimaryTrans, ClsEmailSettings.EnumTipeTransaksiEmail.SalarySlip, vCcTo_SS, vSubject_SS, vMessage_SS, vFileNamePrefix_SS) = 0 Then
                Return False
            End If
            'pick up order
            If dta.Insert(vPrimaryTrans, ClsEmailSettings.EnumTipeTransaksiEmail.PickUpOrder, vCcTo_PickUp, vSubject_PickUp, vMessage_PickUp, vFileNamePrefix_PickUp) = 0 Then
                Return False
            End If
            'overtime order
            If dta.Insert(vPrimaryTrans, ClsEmailSettings.EnumTipeTransaksiEmail.OTOrder, vCcTo_OT, vSubject_OT, vMessage_OT, vFileNamePrefix_OT) = 0 Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vCcTo As String, ByVal vSubject As String, ByVal vMessage As String, ByVal vFileNamePrefix As String, _
    ByVal vCcTo_Attendance As String, ByVal vSubject_Attendance As String, ByVal vMessage_Attendance As String, ByVal vFileNamePrefix_Attendance As String, _
    ByVal vCcTo_SS As String, ByVal vSubject_SS As String, ByVal vMessage_SS As String, ByVal vFileNamePrefix_SS As String, _
    ByVal vCcTo_PickUp As String, ByVal vSubject_PickUp As String, ByVal vMessage_PickUp As String, ByVal vFileNamePrefix_PickUp As String, _
    ByVal vCcTo_OT As String, ByVal vSubject_OT As String, ByVal vMessage_OT As String, ByVal vFileNamePrefix_OT As String) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettingsTableAdapter
        Dim ObjEncrypt As New Core_Module.ClsEncrypt(Me.passPhrase + AccountCode.ToString + PrimaryUser.ToString, Me.initVector)
        Dim EncryptedPass As String = ""

        Try
            EncryptedPass = ObjEncrypt.Encrypt(Password)

            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Update(AccountCode, PrimaryUser, Description, EmailFrom, ReplyTo, Username, EncryptedPass, SmtpServer, _
            SmtpPort, ModifiedBy, ModifiedDate, Me.mUseSSL, Me.mUseCollectionLetter, Me.mUseAttendance, Me.mUseSalarySlip, _
            Me.mUsePickUpOrder, Me.mUseOTOrder, Primary) > 0 Then
                HapusItem(Primary)
                If SimpanItem(Primary, vCcTo, vSubject, vMessage, vFileNamePrefix, _
                vCcTo_Attendance, vSubject_Attendance, vMessage_Attendance, vFileNamePrefix_Attendance, _
                vCcTo_SS, vSubject_SS, vMessage_SS, vFileNamePrefix_SS, _
                vCcTo_PickUp, vSubject_PickUp, vMessage_PickUp, vFileNamePrefix_PickUp, _
                vCcTo_OT, vSubject_OT, vMessage_OT, vFileNamePrefix_OT) = True Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettingsTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Delete(vPrimary) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
    Private Function HapusItem(ByVal vPrimaryTrans As Double) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettings_itemTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Delete(vPrimaryTrans) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vFrm As FrmEmailSettings) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettingsTableAdapter
        Dim tbl As New DseEmailSettings.tbltrans_emailsettingsDataTable
        Dim row As DseEmailSettings.tbltrans_emailsettingsRow
        Dim EncryptedPass As String = ""

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.FillBy_Primary(tbl, vPrimary) > 0 Then
                row = tbl.Rows(0)
                With row
                    Primary = .Primary
                    AccountCode = .Account_Code
                    PrimaryUser = .Primary_User
                    Description = .Description
                    EmailFrom = .Email_From
                    ReplyTo = .Reply_To
                    Username = .Username
                    Dim ObjEncrypt As New Core_Module.ClsEncrypt(Me.passPhrase + AccountCode.ToString + PrimaryUser.ToString, Me.initVector)
                    EncryptedPass = ObjEncrypt.Decrypt(.Password)
                    ObjEncrypt = Nothing
                    Password = EncryptedPass
                    SmtpServer = .SMTP_Server
                    SmtpPort = .SMTP_Port
                    CreatedBy = .Created_By
                    CreatedDate = .Created_Date
                    ModifiedBy = .Modified_By
                    ModifiedDate = .Modified_Date
                    UseSSL = .use_ssl
                    Me.mUseAttendance = .use_attendance
                    Me.mUsePickUpOrder = .use_pick_up_order
                    Me.mUseCollectionLetter = .use_collection_letter
                    Me.mUseSalarySlip = .use_salary_slip
                    Me.mUseOTOrder = .use_ot_order
                End With

                If CariItem(vPrimary, vFrm) = True Then
                    Return True
                Else
                    Return False
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function LoadEmailAccount(ByVal vUserCode As String, ByVal vTbl As DseEmailSettings.DataTable5DataTable, ByVal vTipeItem As ClsEmailSettings.EnumTipeTransaksiEmail) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.DataTable5TableAdapter

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If vTipeItem = EnumTipeTransaksiEmail.Attendance Then
                dta.FillBy_UserCode_Attendance(vTbl, vUserCode)
            ElseIf vTipeItem = EnumTipeTransaksiEmail.CollectionLetter Then
                dta.FillBy_UserCode_CollectionLetter(vTbl, vUserCode)
            ElseIf vTipeItem = EnumTipeTransaksiEmail.SalarySlip Then
                dta.FillBy_UserCode_SalaryFormula(vTbl, vUserCode)
            ElseIf vTipeItem = EnumTipeTransaksiEmail.PickUpOrder Then
                dta.FillBy_UserCodePickUpOrder(vTbl, vUserCode)
            ElseIf vTipeItem = EnumTipeTransaksiEmail.OTOrder Then
                dta.FillBy_UserCode_OTOrder(vTbl, vUserCode, EnumTipeTransaksiEmail.OTOrder)
            Else
                MsgBox("belum setting")
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Private Function CariItem(ByVal vPrimaryTrans As Double, ByVal vFrm As FrmEmailSettings) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettings_itemTableAdapter
        Dim tbl As New DseEmailSettings.tbltrans_emailsettings_itemDataTable
        Dim row As DseEmailSettings.tbltrans_emailsettings_itemRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.FillBy_PrimaryTrans(tbl, vPrimaryTrans) > 0 Then
                For Each row In tbl.Rows
                    With row
                        If row.Tipe_Transaksi = ClsEmailSettings.EnumTipeTransaksiEmail.CollectionLetter Then
                            vFrm.txtItemCCto.Text = .CC
                            vFrm.txtItemSubject.Text = .Subject
                            vFrm.txtItemMessage.Text = .Message
                            vFrm.txtItemFileNamePrefix.Text = .Prefix_Filename
                        ElseIf row.Tipe_Transaksi = ClsEmailSettings.EnumTipeTransaksiEmail.Attendance Then
                            vFrm.txtItemCCto_Att.Text = .CC
                            vFrm.txtItemSubject_Att.Text = .Subject
                            vFrm.txtItemMessage_Att.Text = .Message
                            vFrm.txtItemFileNamePrefix_Att.Text = .Prefix_Filename
                        ElseIf row.Tipe_Transaksi = ClsEmailSettings.EnumTipeTransaksiEmail.SalarySlip Then
                            vFrm.txtItemCCto_SS.Text = .CC
                            vFrm.txtItemSubject_SS.Text = .Subject
                            vFrm.txtItemMessage_SS.Text = .Message
                            vFrm.txtItemFileNamePrefix_SS.Text = .Prefix_Filename
                        ElseIf row.Tipe_Transaksi = ClsEmailSettings.EnumTipeTransaksiEmail.PickUpOrder Then
                            vFrm.txtItemCCto_PickUp.Text = .CC
                            vFrm.txtItemSubject_PickUp.Text = .Subject
                            vFrm.txtItemMessage_PickUp.Text = .Message
                            vFrm.txtItemFileNamePrefix_PickUp.Text = .Prefix_Filename
                        ElseIf row.Tipe_Transaksi = ClsEmailSettings.EnumTipeTransaksiEmail.OTOrder Then
                            vFrm.txtItemCCto_OTOrder.Text = .CC
                            vFrm.txtItemSubject_OTOrder.Text = .Subject
                            vFrm.txtItemMessage_OTOrder.Text = .Message
                            vFrm.txtItemFileNamePrefix_OTOrder.Text = .Prefix_Filename
                        End If
                    End With
                Next
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function CariDetail2(ByVal vCriteria As String, ByVal vDgv As DataGridView, ByVal vTbl As DseEmailSettings.DataTable1DataTable, _
            ByVal vSearchCriteria As ClsEmailSettings.EnumEmailSearchCriteria, ByVal vTsl As ToolStripLabel, _
            Optional ByVal vIsAccountWithUserCode As Boolean = False, Optional ByVal vUserCode As String = "") As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.DataTable1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If vSearchCriteria = EnumEmailSearchCriteria.ByAccountCode Then
                If vIsAccountWithUserCode = False Then
                    QueryResult = dta.FillBy_LIKEAccountCode(vTbl, vCriteria + "%")
                Else
                    QueryResult = dta.FillBy_LIKEAccountCodeAndUserCodeOwner(vTbl, vCriteria + "%", vUserCode)
                End If
            ElseIf vSearchCriteria = EnumEmailSearchCriteria.ByUserCode Then
                If vIsAccountWithUserCode = False Then
                    QueryResult = dta.FillBy_LIKEUserCode(vTbl, vCriteria + "%")
                Else
                    QueryResult = dta.FillBy_LIKEUserCodeAndUserCodeOwner(vTbl, vCriteria + "%", vUserCode)
                End If
            ElseIf vSearchCriteria = EnumEmailSearchCriteria.ByEmailFrom Then
                If vIsAccountWithUserCode = False Then
                    QueryResult = dta.FillBy_LIKEEmailFrom(vTbl, vCriteria + "%")
                Else
                    QueryResult = dta.FillBy_LIKEEmailFromAndUserCodeOwner(vTbl, vCriteria + "%", vUserCode)
                End If
            ElseIf vSearchCriteria = EnumEmailSearchCriteria.ByDescription Then
                If vIsAccountWithUserCode = False Then
                    QueryResult = dta.FillBy_LIKEDesc(vTbl, vCriteria + "%")
                Else
                    QueryResult = dta.FillBy_LIKEDescAndUserCodeOwner(vTbl, vCriteria + "%", vUserCode)
                End If
            End If
            vDgv.DataSource = vTbl.DefaultView
            vDgv.Refresh()
            If vDgv.ColumnCount > 0 Then
                vDgv.Columns(0).Visible = False
                vDgv.Columns(1).HeaderText = "Account Code"
                vDgv.Columns(0).Frozen = True
                vDgv.Columns(1).Frozen = True
                vDgv.Columns(2).HeaderText = "User Code"
                vDgv.Columns(3).HeaderText = "Email From"
                vDgv.Columns(4).HeaderText = "Description"
            End If
            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    vTsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    vTsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                vTsl.Text = "No result found"
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function GetPrimaryByCode(ByVal vCode As String) As Double
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettingsTableAdapter
        Dim _primary As Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            _primary = dta.ScalarQuery_GetPrimaryByCode(vCode)
            If _primary.HasValue Then
                Return _primary
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Public Function GetCodeByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettingsTableAdapter
        Dim code As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            code = dta.ScalarQuery_GetCodeByPrimary(vPrimary)
            Return code
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
    Public Function GetUserCodeByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettingsTableAdapter
        Dim userCode As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            userCode = dta.ScalarQuery_GetUserCodeByPrimary(vPrimary)
            Return userCode
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
    Public Function ApakahSudahAda(ByVal vAccountCode As String) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettingsTableAdapter
        Dim jml As Nullable(Of Long) = 0
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            jml = dta.ScalarQuery_GetCountByAccountCode(vAccountCode)
            If jml.HasValue Then
                If jml.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
    Public Function ApakahSudahAda(ByVal vAccountCode As String, ByVal vUserCodeOwner As String) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettingsTableAdapter
        Dim jml As Nullable(Of Long) = 0
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            jml = dta.ScalarQuery_GetCountByAccountCodeAndUserCodeOwner(vAccountCode, vUserCodeOwner)
            If jml.HasValue Then
                If jml.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function GetPrefxFilename(ByVal vPrimaryAccount As Double, ByVal vTipeTrans As ClsEmailSettings.EnumTipeTransaksiEmail) As String
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettings_itemTableAdapter
        Dim prefixFilename As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            prefixFilename = dta.ScalarQuery_GetPrefixFilenameByPrimaryAccount(vPrimaryAccount, vTipeTrans)
            Return prefixFilename
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetAllDetail(ByVal vPrimaryTrans As Double, ByVal vTipeTrans As Core_Module.ClsEmailSettings.EnumTipeTransaksiEmail, _
            ByRef vFrom As String, ByRef vReplyTo As String, ByRef vUsername As String, ByRef vPassword As String, ByRef vSmtpServer As String, ByRef vSmtpPort As String, _
            ByRef vCc As String, ByRef vSubject As String, ByRef vMessage As String, ByRef vUseSSL As Boolean) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.DataTable2TableAdapter
        Dim tbl As New DseEmailSettings.DataTable2DataTable
        Dim row As DseEmailSettings.DataTable2Row
        Dim EncryptedPass As String = ""

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.FillBy_PrimaryAccountAndTypeAccount(tbl, vPrimaryTrans, vTipeTrans) > 0 Then
                row = tbl.Rows(0)
                With row
                    vFrom = .Email_From
                    vReplyTo = .Reply_To
                    vUsername = .Username
                    Dim ObjEncrypt As New Core_Module.ClsEncrypt(Me.passPhrase + .Account_Code.ToString + .Primary_User.ToString, Me.initVector)
                    EncryptedPass = ObjEncrypt.Decrypt(.Password)
                    ObjEncrypt = Nothing
                    vPassword = EncryptedPass

                    vSmtpServer = .SMTP_Server
                    vSmtpPort = .SMTP_Port
                    vCc = .CC
                    vSubject = .Subject
                    vMessage = .Message
                    If .use_ssl = 0 Then
                        vUseSSL = False
                    Else
                        vUseSSL = True
                    End If
                End With
                Return True
            Else
                vFrom = ""
                vReplyTo = ""
                vUsername = ""
                vPassword = ""
                vSmtpServer = ""
                vSmtpPort = ""
                vCc = ""
                vSubject = ""
                vMessage = ""
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Public Function GetCc(ByVal vPrimaryAccount As Double, ByVal vTipeTrans As ClsEmailSettings.EnumTipeTransaksiEmail) As String
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettings_itemTableAdapter
        Dim cc As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            cc = dta.ScalarQuery_GetCCByPrimaryTransNTipeTrans(vPrimaryAccount, vTipeTrans)
            Return cc
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function SaveEmailLog(ByVal vPrimaryEmailSetting As Double, ByVal vTipeTransaksi As ClsEmailSettings.EnumTipeTransaksiEmail, ByVal vPrimaryTransaksiAcuan As Double, _
    ByVal vDocNo As String, ByVal vEmailSender As String, ByVal vEmailReceiver As String, ByVal vCc As String, ByVal vSubject As String, ByVal vMessage As String, ByVal vStatus As String, _
    ByVal vUserCode As String, ByVal vDate As Date, ByVal vPrimaryCustomer As Double, ByVal vAttachmentFilename As String) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettings_logTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Insert(vPrimaryEmailSetting, vTipeTransaksi, vPrimaryTransaksiAcuan, vDocNo, vEmailSender, vEmailReceiver, vCc, vSubject, vMessage, vStatus, vUserCode, vDate, vPrimaryCustomer, vAttachmentFilename) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function
    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettingsTableAdapter
        Dim str As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            str = dta.ScalarQuery_GetDescByPrimary(vPrimary)
            Return str
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function
    Public Function GetEmailFrom(ByVal vPrimary As Double) As String
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettingsTableAdapter
        Dim str As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            str = dta.ScalarQuery_GetEmailFromByPrimary(vPrimary)
            Return str
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Sub SMTPServer_OnCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        Dim CompletedMailMessage As Mail.MailMessage = CType(e.UserState, Mail.MailMessage)
        Dim EmailTo As String = CompletedMailMessage.To.Item(0).ToString
        Dim EmailFrom As String = CompletedMailMessage.From.Address
        Dim EmailSubject As String = CompletedMailMessage.Subject

        Try
            If e.Cancelled Then
                MsgBox("SendAsync email to " + EmailTo.ToString + " regarding " + EmailSubject.ToString + " was canceled.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If e.Error Is Nothing Then
                MsgBox("SendAsync email to " + EmailTo.ToString + " regarding " + EmailSubject.ToString + " was *SENT*.", MsgBoxStyle.Information, "Information")
            Else
                MsgBox("SendAsync email to " + EmailTo.ToString + " regarding " + EmailSubject.ToString + " was *NOT SENT*.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SendEmail(ByVal ListAddress As List(Of String), ByVal ListAddressCc As List(Of String), _
    ByVal vEmailSubject As String, ByVal vEmailMessage As String, ByVal vListFileName As List(Of String), _
    ByVal vMailFrom As String, ByVal vReplyTo As String, ByVal vUsername As String, ByVal vPassword As String, _
    ByVal vSMTPServer As String, ByVal vSMTPPort As String, ByVal vUseSSL As Boolean)
        Dim AddTo, AddToCc, vFileName As String
        Dim MyMailMessage As New Net.Mail.MailMessage()

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            For Each AddTo In ListAddress
                MyMailMessage.To.Add(New Mail.MailAddress(AddTo))
            Next
            For Each AddToCc In ListAddressCc
                MyMailMessage.CC.Add(New Mail.MailAddress(AddToCc))
            Next
            MyMailMessage.Subject = vEmailSubject
            MyMailMessage.Body = vEmailMessage
            MyMailMessage.From = New Mail.MailAddress(vMailFrom)
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
                End If
            Next
            Dim SMTPServer As New Mail.SmtpClient
            SMTPServer.Credentials = New Net.NetworkCredential(vUsername, vPassword)
            SMTPServer.Host = vSMTPServer
            SMTPServer.Port = vSMTPPort
            SMTPServer.EnableSsl = vUseSSL
            'Dim UserState As Object = MyMailMessage
            'AddHandler SMTPServer.SendCompleted, AddressOf SMTPServer_OnCompleted
            'SMTPServer.SendAsync(MyMailMessage, UserState)
            SMTPServer.Send(MyMailMessage)
            MyMailMessage.Attachments.Dispose()
            MyMailMessage.Dispose()
            SMTPServer = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            MyMailMessage.Attachments.Dispose()
            MyMailMessage.Dispose()
            SmtpServer = Nothing
        Finally
            MyMailMessage = Nothing
        End Try
    End Sub
End Class
