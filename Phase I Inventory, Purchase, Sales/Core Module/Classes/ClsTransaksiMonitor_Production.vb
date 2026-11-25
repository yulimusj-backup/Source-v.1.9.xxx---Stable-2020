Option Explicit On
Option Strict Off

Public Class ClsTransaksiMonitor_Production
    Private userCode As String = ""
    Private userIdentifier As Double = 0
    Private userType As ClsTransaksiMonitor_Production.enumUserType = enumUserType.READ_OR_FULLCONTROL
    Private objMySQL As New Core_Module.ClsMySQLReference
    Dim ObjUser As New Core_Module.ClsCekData

    Public Enum enumUserType As Integer
        READ_ONLY
        READ_OR_FULLCONTROL
    End Enum
    Public Enum enumUserControl As Integer
        FULL_CONTROL
        READ_ONLY
    End Enum
    Public Enum enumRequestType As Integer
        TAKE
        SEND
    End Enum
    Public Enum enumAnswer As Integer
        WAITING
        ACCEPT
        REJECT
        FAIL_CREATE_REQUEST
        USER_CANCEL
        TIMEOUT
    End Enum

    Public Sub New(ByVal vUserCode As String, ByVal vUserIdentifier As Double, ByVal vUserType As ClsTransaksiMonitor_Production.enumUserType)
        Me.userCode = vUserCode
        Me.userIdentifier = vUserIdentifier
        Me.userType = vUserType
    End Sub
    Public Sub New()
    End Sub

#Region "common function"
    Public Function GetUserType(ByVal vIdentifier As Double) As enumUserType
        Return enumUserType.READ_OR_FULLCONTROL
    End Function

    Public Function GetUserTypeStr(ByVal vType As ClsTransaksiMonitor_Production.enumUserType) As String
        Try
            Select Case vType
                Case enumUserType.READ_ONLY
                    Return "Read Only"
                Case enumUserType.READ_OR_FULLCONTROL
                    Return "Read Onlny & Full Control"
            End Select
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function getUserControl(ByVal vUserIdentifier As Double, ByVal vUserType As ClsTransaksiMonitor_Production.enumUserType) As ClsTransaksiMonitor_Production.enumUserControl
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_monitorTableAdapter
        Dim n As New Nullable(Of Double)
        Dim n2 As New Nullable(Of Double)
        Try
            If vUserType = enumUserType.READ_ONLY Then
                Return enumUserControl.READ_ONLY
            Else
                d.Connection.ConnectionString = Core_Module.GetConstring
                n = d.ScalarQuery_GetMinimumPrimary(CDbl(vUserType))
                If n.HasValue Then
                    n2 = d.ScalarQuery_GetPrimaryByUserIdentifier(vUserIdentifier)
                    If n2.HasValue Then
                        If n.Value = n2.Value Then
                            Return enumUserControl.FULL_CONTROL
                        End If
                    End If
                End If
                Return enumUserControl.READ_ONLY
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return enumUserControl.READ_ONLY
        End Try
    End Function

    Public Function getUserControl() As ClsTransaksiMonitor_Production.enumUserControl
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_monitorTableAdapter
        Dim n As New Nullable(Of Double)
        Dim n2 As New Nullable(Of Double)
        Try
            If Me.userType = enumUserType.READ_ONLY Then
                Return enumUserControl.READ_ONLY
            Else
                d.Connection.ConnectionString = Core_Module.GetConstring
                n = d.ScalarQuery_GetMinimumPrimary(CInt(enumUserType.READ_OR_FULLCONTROL))
                If n.HasValue Then
                    n2 = d.ScalarQuery_GetPrimaryByUserIdentifier(Me.userIdentifier)
                    If n2.HasValue Then
                        If n.Value = n2.Value Then
                            Return enumUserControl.FULL_CONTROL
                        End If
                    End If
                End If
                Return enumUserControl.READ_ONLY
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return enumUserControl.READ_ONLY
        End Try
    End Function

    Public Function getPrimary(ByVal vUserIdentifier As Double) As Double
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_monitorTableAdapter
        Dim n As New Nullable(Of Double)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            n = d.ScalarQuery_GetPrimaryByUserIdentifier(vUserIdentifier)
            If n.HasValue Then
                Return n.Value
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Public Function updatePrimary(ByVal vNewPrimary As Double, ByVal vUserIdentifier As Double) As Boolean
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_monitorTableAdapter
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            If d.Update_PrimaryMonitor(vNewPrimary, vUserIdentifier) <= 0 Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Sub Set_UserIdentifier(ByVal vUserIdentifier As String)
        Try
            Me.userIdentifier = vUserIdentifier
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "production monitor"
    Private WithEvents productionTimer As New Timer
    Private WithEvents flushTimer As New Timer
    Private monitorInterval As Double = 3 * 60000 ''update monitor every 3 minutes 
    Private flushInterval As Double = 5 * 60000 ''flush expired monitor every 5 minutes

    Public Function startProductionMonitor(ByVal vMonitorInterval As Double, ByVal vFlushMonitor As Double) As Boolean
        Me.monitorInterval = vMonitorInterval
        Me.flushInterval = vFlushMonitor
        Try
            flushMonitor_Tick()
            flushTimer.Interval = Me.flushInterval
            flushTimer.Start()

            startProductionMonitor_Tick()
            productionTimer.Interval = Me.monitorInterval
            productionTimer.Start()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function stopProductionMonitor() As Boolean
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_monitorTableAdapter
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            d.Delete_ByUserIdentifier(Me.userIdentifier)
            flushTimer.Stop()
            productionTimer.Stop()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub startProductionMonitor_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles productionTimer.Tick
        Try
            startProductionMonitor_Tick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub startProductionMonitor_Tick()
        Try
            If Me.isExistIdentifier(userIdentifier) Then
                updateMonitor(Me.userCode, Me.userIdentifier)
            Else
                addNewMonitor(Me.userCode, Me.userIdentifier, Me.userType)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub flushMonitor_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles flushTimer.Tick
        Try
            flushMonitor_Tick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub flushMonitor_Tick()
        Dim timeKeyToDelete As Date
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_monitorTableAdapter
        Try
            timeKeyToDelete = Me.objMySQL.GetServerTime.AddMilliseconds(Me.flushInterval * -1)
            d.Connection.ConnectionString = Core_Module.GetConstring
            If d.Delete_Flush(timeKeyToDelete) > 0 Then
                'Console.WriteLine("flushed, time deleted: " & Format(timeKeyToDelete, "hh:mm:ss"))
            End If
            Me.flushRequest()
            Console.WriteLine("flushed, time deleted: " & Format(timeKeyToDelete, "yyyy MM dd hh:mm:ss"))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function isExistIdentifier(ByVal vUserIdentifier As Double) As Boolean
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_monitorTableAdapter
        Dim n As New Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            n = d.ScalarQuery_GetJmlByUserIdentifier(vUserIdentifier)
            If n.HasValue Then
                If n.Value > 0 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function addNewMonitor(ByVal vUserCode As String, ByVal vIdentifier As Double, ByVal vUserType As ClsTransaksiMonitor_Production.enumUserType) As Boolean
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_monitorTableAdapter
        Dim vComputerName As String = ""
        Dim vIPAddess As String = ""
        Dim vTimeKey As Date
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            vComputerName = System.Net.Dns.GetHostName()
            vIPAddess = System.Net.Dns.GetHostByName(vComputerName).AddressList(0).ToString()
            vTimeKey = objMySQL.GetServerTime
            If d.InsertMonitor(vUserCode, vComputerName, vIPAddess, vTimeKey, vIdentifier, vUserType) <= 0 Then
                Return False
            End If
            Console.WriteLine("inserted: " & Me.userIdentifier.ToString & ", new time key: " & Format(vTimeKey, "hh:mm:ss"))
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function updateMonitor(ByVal vUserCode As String, ByVal vUserIdentifier As Double) As Boolean
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_monitorTableAdapter
        Dim vTimeKey As Date = Now
        Try
            vTimeKey = Me.objMySQL.GetServerTime
            d.Connection.ConnectionString = Core_Module.GetConstring
            If d.UpdateTimeKey(vTimeKey, vUserCode, vUserIdentifier) <= 0 Then
                Return False
            End If
            Console.WriteLine("updated: " & Me.userIdentifier.ToString & ", new time key: " & Format(vTimeKey, "hh:mm:ss"))
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function flushRequest() As Boolean
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.tbltrans_prod_handshakeTableAdapter
        Dim serverTime As Date
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            serverTime = Me.objMySQL.GetServerTime
            d.DeleteExpired(serverTime)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
#End Region

#Region "user control refresh"
    Private vListOfButton As List(Of Button)
    Private vListOfDtp As List(Of DateTimePicker)
    Private vUserControl As ClsTransaksiMonitor_Production.enumUserControl
    Private vToolStrip As ToolStripStatusLabel
    Private WithEvents userControlTimer As New Timer
    Private userControlRefreshInterval As Double = 1000 ''check status (full control or read only) every second

    Public Sub startUserControlMonitor(ByVal vUserControl As ClsTransaksiMonitor_Production.enumUserControl, ByVal vToolStrip As ToolStripStatusLabel, _
            ByVal vUserControlRefreshInterval As Double, ByVal vListOfButton As List(Of Button), ByVal vListOfDtp As List(Of DateTimePicker))
        Me.userControlRefreshInterval = vUserControlRefreshInterval
        Try
            Me.vUserControl = vUserControl
            Me.vToolStrip = vToolStrip
            Me.vListOfButton = vListOfButton
            Me.vListOfDtp = vListOfDtp
            Me.userControlTimer_Tick(vUserControl, vToolStrip)
            Me.userControlTimer.Interval = Me.userControlRefreshInterval
            Me.userControlTimer.Start()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub userControlTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles userControlTimer.Tick
        Try
            Me.userControlTimer_Tick(Me.vUserControl, Me.vToolStrip)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub userControlTimer_Tick(ByVal vUserControl As ClsTransaksiMonitor_Production.enumUserControl, ByVal vToolStrip As ToolStripStatusLabel)
        Try
            vUserControl = getUserControl()
            If vUserControl = ClsTransaksiMonitor_Production.enumUserControl.FULL_CONTROL Then
                vToolStrip.Text = "Control Status : Full Control"
                For i As Integer = 0 To vListOfButton.Count - 1
                    vListOfButton.Item(i).Enabled = True
                Next
                For i As Integer = 0 To vListOfDtp.Count - 1
                    vListOfDtp.Item(i).Enabled = True
                Next
            Else
                vToolStrip.Text = "Control Status  : Read Only"
                For i As Integer = 0 To vListOfButton.Count - 1
                    vListOfButton.Item(i).Enabled = False
                Next
                For i As Integer = 0 To vListOfDtp.Count - 1
                    vListOfDtp.Item(i).Enabled = False
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function stopUserControlMonitor() As Boolean
        Try
            userControlTimer.Stop()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
#End Region

#Region "notif monitor"
    Private WithEvents NotifyIcon1 As New System.Windows.Forms.NotifyIcon
    Private WithEvents notifTimer As New Timer
    Private notifInterval As Double = 1000 ''check notification every second
    Dim vFrmNotif As New FrmTransaksiMonitor_Production_Notif

    Public Function startNotifMonitor(ByVal vNotifInterval As Double) As Boolean
        Me.notifInterval = vNotifInterval
        Try
            Me.notifTimer_Tick()
            Me.notifTimer.Interval = Me.notifInterval
            Me.notifTimer.Start()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub notifTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles notifTimer.Tick
        Try
            Me.notifTimer_Tick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub notifTimer_Tick()
        Dim d As New DseTransaksiMonitor_ProductionTableAdapters.DataTable1TableAdapter
        Dim t As New DseTransaksiMonitor_Production.DataTable1DataTable
        Dim r As DseTransaksiMonitor_Production.DataTable1Row
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            If d.Fill(t, Me.userIdentifier, ClsTransaksiMonitor_Production.enumAnswer.WAITING) > 0 Then
                r = t.Rows(0)
                Me.vFrmNotif = New FrmTransaksiMonitor_Production_Notif
                If r.request_type = ClsTransaksiMonitor_Production.enumRequestType.TAKE Then
                    vFrmNotif.TextBox1.Text = r.user_code & " want to take full control access of production planning form from you, do you want to accept?"
                    Me.vFrmNotif.NotifyIcon1.BalloonTipText = r.user_code & " want to take full control access of production planning form from you. Click here to respons."
                ElseIf r.request_type = ClsTransaksiMonitor_Production.enumRequestType.SEND Then
                    vFrmNotif.TextBox1.Text = r.user_code & " want to send you full control access of production planning, do you want to accept??"
                    Me.vFrmNotif.NotifyIcon1.BalloonTipText = r.user_code & " want to send you full control access of production planning. Click here to respons."
                End If
                Me.stopNotifMonitor()
                Me.vFrmNotif.expiredTime = r.expired_time
                Me.vFrmNotif.userIdentifier = r.identifier
                Me.vFrmNotif.identifierTarget = r.identifier_target
                Me.vFrmNotif.requestType = r.request_type
                Me.vFrmNotif.ShowDialog()
                Me.notifTimer.Start()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub stopNotifMonitor()
        Try
            Me.notifTimer.Stop()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#End Region

End Class
