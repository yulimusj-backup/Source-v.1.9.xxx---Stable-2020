Option Explicit On
Option Strict Off

''' <summary>
''' kelas ini mengatur email - email yang keluar dan masuk
''' </summary>
''' <remarks></remarks>
Public Class ClsTransEmail
    Dim mPrimary As Double = -1
    Dim mSubject As String = ""
    Dim mContent As String = ""
    Dim mOwner As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mFolder As EnumFolder = EnumFolder.Inbox
    Dim mEmailStatus As EnumEmailStatus = EnumEmailStatus.NewEmail
    Dim mTipeEmail As EnumTipeEmail = EnumTipeEmail.StandarEmail

    Public Enum EnumTipeEmail As Integer
        StandarEmail
    End Enum

    Public Property TipeEmail() As EnumTipeEmail
        Get
            Return Me.mTipeEmail
        End Get
        Set(ByVal value As EnumTipeEmail)
            mTipeEmail = value
        End Set
    End Property

    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property Subject() As String
        Get
            Return Me.mSubject
        End Get
        Set(ByVal value As String)
            mSubject = value
        End Set
    End Property

    Public Property Email_Status() As EnumEmailStatus
        Get
            Return Me.mEmailStatus
        End Get
        Set(ByVal value As EnumEmailStatus)
            mEmailStatus = value
        End Set
    End Property

    Public Property Folder() As EnumFolder
        Get
            Return Me.mFolder
        End Get
        Set(ByVal value As EnumFolder)
            mFolder = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return Me.mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property CreatedBy() As String
        Get
            Return Me.mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property Owner() As String
        Get
            Return Me.mOwner
        End Get
        Set(ByVal value As String)
            mOwner = value
        End Set
    End Property

    Public Property Content() As String
        Get
            Return Me.mContent
        End Get
        Set(ByVal value As String)
            mContent = value
        End Set
    End Property

    Public Enum EnumFolder As Integer
        Inbox
        SentItem
    End Enum

    Public Enum EnumEmailStatus As Integer
        NewEmail
        AlreadyRead
        Replied
    End Enum

    Public Function SendEmail(ByVal UserCode As List(Of String)) As Boolean
        Dim DTA As New DsTransaksiEmailTableAdapters.tbltrans_emailTableAdapter
        Dim DTASendTo As New DsTransaksiEmailTableAdapters.tbltrans_email_itemsendtoTableAdapter

        Dim ObjEW As New ClsEmailWatcher
        Dim SentFolder As EnumFolder = EnumFolder.SentItem
        Dim StatusSent As EnumEmailStatus = EnumEmailStatus.AlreadyRead
        Dim StatusReplied As EnumEmailStatus = EnumEmailStatus.Replied
        Dim KodeUser As String = ""
        Dim dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTASendTo.Connection.ConnectionString = DTA.Connection.ConnectionString
            With Me
                'masukkan data ke sent items
                Dim Timestamp As Date = Now
                For Each KodeUser In UserCode
                    If dummy = "" Then
                        dummy = KodeUser
                    Else
                        dummy = dummy + ", " + KodeUser
                    End If
                Next
                StatusSent = EnumEmailStatus.AlreadyRead
                SentFolder = EnumFolder.SentItem
                If .Content.Length > 10000 Then
                    .Content = .Content.Substring(0, 9996) + "..."
                End If
                If DTA.Insert(.Subject, .Content, StatusSent, SentFolder, dummy, .CreatedBy, Timestamp, .TipeEmail) > 0 Then
                    .Primary = DTA.ScalarQuery_GetPrimary(.CreatedBy, Timestamp, .Subject).Value
                    .CreatedDate = Timestamp

                    'kirimkan data untuk setiap user + notify
                    For Each KodeUser In UserCode
                        StatusSent = EnumEmailStatus.NewEmail
                        SentFolder = EnumFolder.Inbox
                        DTA.Insert(.Subject, .Content, StatusSent, SentFolder, _
                        KodeUser, .CreatedBy, Timestamp, .TipeEmail)
                        'simpan data send to
                        DTASendTo.Insert(.Primary, KodeUser)
                    Next
                    ObjEW.NotifyUser(UserCode)

                End If
            End With

            'kalo sampe sini maka proses berhasil
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTASendTo.Dispose()
            ObjEW = Nothing
        End Try
    End Function

    Public Function SendEmail(ByVal clbTO As Windows.Forms.CheckedListBox, _
    Optional ByVal Reply As Boolean = False, Optional ByVal PrimaryEmailYgDireply As Boolean = -1) As Boolean
        Dim DTA As New DsTransaksiEmailTableAdapters.tbltrans_emailTableAdapter
        Dim DTASendTo As New DsTransaksiEmailTableAdapters.tbltrans_email_itemsendtoTableAdapter

        Dim ObjEW As New ClsEmailWatcher
        Dim row As Core_Module.DSUser.tbluser1Row
        Dim rowview As DataRowView
        Dim SentFolder As EnumFolder = EnumFolder.SentItem
        Dim StatusSent As EnumEmailStatus = EnumEmailStatus.AlreadyRead
        Dim StatusReplied As EnumEmailStatus = EnumEmailStatus.Replied
        Dim dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTASendTo.Connection.ConnectionString = DTA.Connection.ConnectionString 
            With Me
                'masukkan data ke sent items
                Dim Timestamp As Date = Now
                For Each rowview In clbTO.CheckedItems
                    row = TryCast(rowview.Row, Core_Module.DSUser.tbluser1Row)
                    If dummy = "" Then
                        dummy = row.Kode_User
                    Else
                        dummy = dummy + ", " + row.Kode_User
                    End If
                Next
                SentFolder = EnumFolder.SentItem
                StatusSent = EnumEmailStatus.AlreadyRead
                DTA.Insert(.Subject, .Content, StatusSent, SentFolder, dummy, .CreatedBy, Timestamp, .TipeEmail)
                .Primary = CDbl(DTA.ScalarQuery_GetPrimary(.CreatedBy, Timestamp, .Subject))
                .CreatedDate = Timestamp

                'kirimkan data untuk setiap user + notify
                For Each rowview In clbTO.CheckedItems
                    row = TryCast(rowview.Row, Core_Module.DSUser.tbluser1Row)
                    StatusSent = EnumEmailStatus.NewEmail
                    SentFolder = EnumFolder.Inbox
                    DTA.Insert(.Subject, .Content, StatusSent, SentFolder, _
                    row.Kode_User, .CreatedBy, Timestamp, .TipeEmail)
                    'simpan data send to
                    DTASendTo.Insert(.Primary, row.Kode_User)
                Next
                ObjEW.NotifyUser(clbTO)

                'untuk email yang direply update status email yang direply
                If Reply = True Then
                    DTA.UpdateQuery_Status(StatusReplied, CDec(PrimaryEmailYgDireply))
                End If
            End With

            'kalo sampe sini maka proses berhasil
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTASendTo.Dispose()
            ObjEW = Nothing
        End Try
    End Function

    Public Function ApakahAdaYgDikirimiDOReady() As Boolean
        Dim DTA As New E_mail_Module.DsTransaksiEmailTableAdapters.tbltrans_email_automated_sendtoTableAdapter
        Dim TBL As New E_mail_Module.DsTransaksiEmail.tbltrans_email_automated_sendtoDataTable

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryEmail(TBL, 1)

            If TBL.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Sub HapusEmail(ByVal vPrimaryEmail As Double)
        Dim DTA As New DsTransaksiEmailTableAdapters.tbltrans_emailTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery_ByPrimary(CDec(vPrimaryEmail))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetCreatedBy(ByVal vPrimaryEmail As Double) As String
        Dim DTA As New DsTransaksiEmailTableAdapters.tbltrans_emailTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(CDec(vPrimaryEmail))) = False Then
                Return DTA.ScalarQuery_GetCreatedBy(CDec(vPrimaryEmail))
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub HapusEmail(ByVal vUserCode As String, ByVal vFolder As EnumFolder)
        Dim DTA As New DsTransaksiEmailTableAdapters.tbltrans_emailTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vFolder = EnumFolder.Inbox Then
                DTA.DeleteQuery_ByOwner(vUserCode, vFolder)
            Else
                DTA.DeleteQuery_byCreatedBy(vFolder, vUserCode)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub HapusEmailSudahDibaca(ByVal vUserCode As String)
        Dim DTA As New DsTransaksiEmailTableAdapters.tbltrans_emailTableAdapter
        Dim vFolder As EnumFolder = EnumFolder.Inbox
        Dim vStatus As EnumEmailStatus = EnumEmailStatus.AlreadyRead

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery_ByStatus(vUserCode, vFolder, vStatus)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiDGV(ByVal vOwner As String, ByVal vFolder As EnumFolder, _
    ByVal vDgv As Windows.Forms.DataGridView, ByVal vTBL As DsTransaksiEmail.tbltrans_email1DataTable)
        Dim DTA As New DsTransaksiEmailTableAdapters.tbltrans_email1TableAdapter
        Dim imgCol As New System.Windows.Forms.DataGridViewImageColumn
        Dim ij As New Windows.Forms.DataGridViewCellStyle

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fill dari database
            If vFolder = EnumFolder.Inbox Then
                DTA.Fill(vTBL, vFolder, vOwner)
            Else
                DTA.Fill_SentItem(vTBL, vFolder, vOwner)
            End If
            'kosongkan datagridview
            While vDgv.Columns.Count > 0
                vDgv.Columns.RemoveAt(0)
            End While
            vDgv.DataSource = Nothing

            'isikan ke dgv
            If vTBL.Rows.Count > 0 Then
                vDgv.DataSource = vTBL.DefaultView
                'setting kolom
                vDgv.Columns(0).Visible = False
                vDgv.Columns(2).Visible = False
                If vFolder = EnumFolder.Inbox Then
                    vDgv.Columns(3).HeaderText = "From"
                Else
                    vDgv.Columns(3).HeaderText = "To"
                End If
                vDgv.Columns(5).Visible = False
                ij.Format = "dd MMM yyyy - HH:mm"
                vDgv.Columns(4).DefaultCellStyle = ij
                imgCol.Name = "Image"
                imgCol.HeaderText = ""
                imgCol.ImageLayout = Windows.Forms.DataGridViewImageCellLayout.Zoom
                vDgv.Columns.Insert(1, imgCol)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub ShowContent(ByVal txt As Windows.Forms.TextBox, ByVal vPrimaryEmail As Double)
        Dim DTA As New DsTransaksiEmailTableAdapters.tbltrans_emailTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetContent(CDec(vPrimaryEmail))) = False Then
                txt.Text = DTA.ScalarQuery_GetContent(CDec(vPrimaryEmail))
            Else
                txt.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub UpDateStatus(ByVal Status As EnumEmailStatus, ByVal PrimaryEmail As Double)
        Dim DTA As New DsTransaksiEmailTableAdapters.tbltrans_emailTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_Status(Status, CDec(PrimaryEmail))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
End Class
