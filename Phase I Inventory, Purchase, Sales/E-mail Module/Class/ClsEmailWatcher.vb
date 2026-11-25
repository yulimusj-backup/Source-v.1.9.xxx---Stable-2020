Public Class ClsEmailWatcher
    Inherits Windows.Forms.Timer
    Dim WithEvents Notifier As New Windows.Forms.NotifyIcon
    Dim WithEvents SoundPlayer As New System.Media.SoundPlayer
    Dim AlarmNotifyMenu As Windows.Forms.ToolStripMenuItem
    Dim UserCode As String = ""
    Dim ParentForm As Windows.Forms.Form
    Dim i As Integer = 1

    Public Property User() As String
        Get
            Return Me.UserCode
        End Get
        Set(ByVal value As String)
            Me.UserCode = value
        End Set
    End Property

    Private Sub ClsEmailWatcher_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Try
            Notifier.Visible = False
            Notifier.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ClsEmailWatcher_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Tick
        Dim DTA As New DseEmailNotifierTableAdapters.tbltrans_email_notifierTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cek apakah ada user yang mengirimkan email ke kita
            If DTA.ScalarQuery_GetJml(Me.UserCode).HasValue = True Then
                If DTA.ScalarQuery_GetJml(Me.UserCode).Value > 0 Then
                    'jika ada tampakkan notify icon
                    Me.Notifier.ShowBalloonTip(10000)
                    Me.SoundPlayer.Play()

                    'hapuskan peringatan yang ada di dalam database
                    DTA.Delete(Me.UserCode)
                    'prosedur first check tidak perlu dijalankan lagi
                    i = 0
                    Exit Sub
                End If
            End If

            'prosedur ini hanya dijalankan untuk pertama kali aplikasi menyala
            If Me.i = 1 Then
                i = 0
                Dim Status As ClsTransEmail.EnumEmailStatus = ClsTransEmail.EnumEmailStatus.NewEmail
                Dim Folder As ClsTransEmail.EnumFolder = ClsTransEmail.EnumFolder.Inbox
                Dim DTA1 As New DsTransaksiEmailTableAdapters.tbltrans_emailTableAdapter
                DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
                If DTA1.ScalarQuery_GetJml(Status, Folder, Me.UserCode).HasValue = True Then
                    If DTA1.ScalarQuery_GetJml(Status, Folder, Me.UserCode).Value > 0 Then
                        Me.Notifier.Icon = My.Resources.email
                        Me.Notifier.ShowBalloonTip(10000)
                        Me.SoundPlayer.Play()
                    End If
                End If
                DTA1.Dispose()
            End If
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub New()
        Me.Enabled = False
    End Sub

    Public Sub New(ByVal vUserCode As String, ByVal mParentForm As Windows.Forms.Form, _
    ByVal vNotifier As Windows.Forms.NotifyIcon, ByVal AlarmMenuItem As Windows.Forms.ToolStripMenuItem)
        Try
            'cek email setiap 5 detik
            Me.Interval = 5000
            Me.UserCode = vUserCode
            Me.ParentForm = mParentForm
            'set notifier
            Me.Notifier = vNotifier
            With Me.Notifier
                .BalloonTipIcon = Windows.Forms.ToolTipIcon.Info
                .BalloonTipText = "You have new incoming mail, go to email form to see details."
                .BalloonTipTitle = "New incoming mail"
                .Icon = My.Resources.email
                .Visible = True
            End With
            'set sound player
            With Me.SoundPlayer
                .Stream() = My.Resources.Classy
            End With
            'set menu item
            Me.AlarmNotifyMenu = AlarmMenuItem
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' memberitahukan user dari email yang dituju bahwa dia mendapatkan email baru
    ''' </summary>
    ''' <param name="clb"></param>
    ''' <remarks></remarks>
    Public Sub NotifyUser(ByVal clb As Windows.Forms.CheckedListBox)
        Dim DTA As New DseEmailNotifierTableAdapters.tbltrans_email_notifierTableAdapter
        Dim row As Core_Module.DSUser.tbluser1Row
        Dim rowview As DataRowView

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each rowview In clb.CheckedItems
                row = TryCast(rowview.Row, Core_Module.DSUser.tbluser1Row)
                If DTA.ScalarQuery_GetJml(row.Kode_User).HasValue = False Then
                    DTA.Insert(row.Kode_User, 1)
                Else
                    If DTA.ScalarQuery_GetJml(row.Kode_User).Value = 0 Then
                        DTA.Insert(row.Kode_User, 1)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub NotifyUser(ByVal KodeUser As List(Of String))
        Dim DTA As New DseEmailNotifierTableAdapters.tbltrans_email_notifierTableAdapter
        Dim x As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each x In KodeUser
                If DTA.ScalarQuery_GetJml(x).HasValue = False Then
                    DTA.Insert(x, 1)
                Else
                    If DTA.ScalarQuery_GetJml(x).Value = 0 Then
                        DTA.Insert(x, 1)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub Notifier_BalloonTipClicked(ByVal sender As Object, ByVal e As System.EventArgs) Handles Notifier.BalloonTipClicked
        Try
            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
            'start memantau email datang
            If Me.Enabled = False Then
                Me.Start()
                Me.AlarmNotifyMenu.Checked = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.Arrow
        End Try
    End Sub
End Class
