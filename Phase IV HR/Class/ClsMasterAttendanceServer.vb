''@rosadinnn
''201508111022

Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class ClsMasterAttendanceServer

    Dim ObjCd As New Core_Module.ClsCekData

    Public Enum EnumSearchCriteria As Integer
        BY_SERVERNAME
        BY_LOCATION
        BY_REMARKS
    End Enum

    Public Enum EnumMachineType As Integer
        NITGEN
        MAGIC
    End Enum
    
    Private mPrimary As Double = -1
    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Private mServerName As String = ""
    Public Property ServerName() As String
        Get
            Return mServerName
        End Get
        Set(ByVal value As String)
            mServerName = value
        End Set
    End Property

    Private mLocation As String = ""
    Public Property Location() As String
        Get
            Return mLocation
        End Get
        Set(ByVal value As String)
            mLocation = value
        End Set
    End Property

    Private mIsActive As Integer = 1
    Public Property IsActive() As Integer
        Get
            Return mIsActive
        End Get
        Set(ByVal value As Integer)
            mIsActive = value
        End Set
    End Property

    Private mRemarks As String = ""
    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property
    Private mDataSource As String = ""

    Public Property DataSource() As String
        Get
            Return mDataSource
        End Get
        Set(ByVal value As String)
            mDataSource = value
        End Set
    End Property

    Private mInitialCatalog As String = ""
    Public Property InitialCatalog() As String
        Get
            Return mInitialCatalog
        End Get
        Set(ByVal value As String)
            mInitialCatalog = value
        End Set
    End Property

    Private mUserId As String = ""
    Public Property UserId() As String
        Get
            Return mUserId
        End Get
        Set(ByVal value As String)
            mUserId = value
        End Set
    End Property

    Private mPassword As String = ""
    Public Property Password() As String
        Get
            Return mPassword
        End Get
        Set(ByVal value As String)
            mPassword = value
        End Set
    End Property

    Private mCreatedBy As String = ""
    Public Property CreatedBy() As String
        Get
            Return Me.mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Private mCreatedDate As Date = Now
    Public Property CreatedDate() As Date
        Get
            Return Me.mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Private mModifiedBy As String = ""
    Public Property ModifiedBy() As String
        Get
            Return Me.mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Private mModifiedDate As Date = Now
    Public Property ModifiedDate() As Date
        Get
            Return Me.mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Private mTimestamp As Date = Now
    Public Property TimeStamp() As Date
        Get
            Return Me.mTimestamp
        End Get
        Set(ByVal value As Date)
            Me.mTimestamp = value
        End Set
    End Property

    Private mMachineType As ClsMasterAttendanceServer.EnumMachineType
    Public Property MachineType() As ClsMasterAttendanceServer.EnumMachineType
        Get
            Return mMachineType
        End Get
        Set(ByVal value As ClsMasterAttendanceServer.EnumMachineType)
            mMachineType = value
        End Set
    End Property


    Public Function SimpanBaru(ByVal vFrm As FrmMasterAttendanceServer, ByVal vBaru As Boolean) As Boolean
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim DtaMachine As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_server_machinesTableAdapter
        Dim Remarks_ As String = ""
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(vBaru) = True Then
                With Me
                    If Dta.Insert(.ServerName, .Location, .IsActive, .Remarks, .DataSource, .InitialCatalog, .UserId, .Password, .CreatedBy, Now, .ModifiedBy, Now, Now, .MachineType) > 0 Then
                        .Primary = .GetPrimary(.ServerName)
                        .CreatedDate = .GetCreatedDate(.Primary)
                        .ModifiedDate = .GetModifiedDate(.Primary)
                        .TimeStamp = .GetTimestamp(.Primary)
                        If Me.ApakahDataMesinBenar = True Then
                            For Each DgvRow As DataGridViewRow In vFrm.DgvItem.Rows
                                If IsNothing(DgvRow.Cells(vFrm.ColMachineCode.Index)) = False Then
                                    If IsNothing(DgvRow.Cells(vFrm.ColTerminalID.Index)) = False Then
                                        If DgvRow.Cells(vFrm.ColMachineCode.Index).Value <> "" And DgvRow.Cells(vFrm.ColTerminalID.Index).Value <> "" Then
                                            Dim IsActive_ As Integer = 1
                                            If DgvRow.Cells(vFrm.ColIsActive.Index).Value = Boolean.FalseString Then
                                                IsActive_ = 0
                                            End If
                                            If IsNothing(DgvRow.Cells(vFrm.ColRemarks.Index).Value) = True Then
                                                Remarks_ = ""
                                            Else
                                                Remarks_ = DgvRow.Cells(vFrm.ColRemarks.Index).Value.ToString
                                            End If
                                            If DtaMachine.Insert(.Primary, DgvRow.Cells(vFrm.ColMachineCode.Index).Value.ToString, _
                                                    DgvRow.Cells(vFrm.ColMachineName.Index).Value.ToString, DgvRow.Cells(vFrm.ColTerminalID.Index).Value.ToString, _
                                                    IsActive_, Remarks_) <= 0 Then
                                                Return False
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            Return False
                        End If
                    Else
                        Return False
                    End If
                End With
            Else
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Dta.Dispose()
            DtaMachine.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vFrm As FrmMasterAttendanceServer, ByVal vBaru As Boolean) As Boolean
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim DtaMachine As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_server_machinesTableAdapter
        Dim Remarks_ As String = ""
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehEdit = True Then
                If Me.ApakahDataBenar(vBaru) = True Then
                    With Me
                        If Dta.Update(.ServerName, .Location, .IsActive, .Remarks, .DataSource, .InitialCatalog, .UserId, .Password, .ModifiedBy, Now, Now, .Primary) > 0 Then
                            .ModifiedDate = .GetModifiedDate(.Primary)
                            .TimeStamp = .GetTimestamp(.Primary)
                            If Me.ApakahDataMesinBenar = True Then
                                DtaMachine.Delete(.Primary)
                                For Each DgvRow As DataGridViewRow In vFrm.DgvItem.Rows
                                    If IsNothing(DgvRow.Cells(vFrm.ColMachineCode.Index)) = False Then
                                        If IsNothing(DgvRow.Cells(vFrm.ColTerminalID.Index)) = False Then
                                            If DgvRow.Cells(vFrm.ColMachineCode.Index).Value <> "" And DgvRow.Cells(vFrm.ColTerminalID.Index).Value.ToString <> "" Then
                                                Dim IsActive_ As Integer = 1
                                                If DgvRow.Cells(vFrm.ColIsActive.Index).Value = Boolean.FalseString Then
                                                    IsActive_ = 0
                                                End If
                                                If IsNothing(DgvRow.Cells(vFrm.ColRemarks.Index).Value) = True Then
                                                    Remarks_ = ""
                                                Else
                                                    Remarks_ = DgvRow.Cells(vFrm.ColRemarks.Index).Value.ToString
                                                End If
                                                If DtaMachine.Insert(.Primary, DgvRow.Cells(vFrm.ColMachineCode.Index).Value.ToString, _
                                                        DgvRow.Cells(vFrm.ColMachineName.Index).Value.ToString, DgvRow.Cells(vFrm.ColTerminalID.Index).Value.ToString, _
                                                        IsActive_, Remarks_) <= 0 Then
                                                    Return False
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                Return False
                            End If
                        Else
                            Return False
                        End If
                    End With
                Else
                    Return False
                End If
            Else
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Dta.Dispose()
            DtaMachine.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus() = True Then
                If Dta.Delete(vPrimary) <= 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vServerName As String) As Boolean
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim Count_ As Nullable(Of Long) = 0
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Count_ = Dta.ScalarQuery_GetCount(vServerName)
            If Count_.HasValue = True Then
                If Count_.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Try
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.ServerName) = True Then
                    MsgBox(Me.ServerName & " is already exist.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                If Me.ApakahSudahAda(Me.ServerName) = True Then
                    If Me.GetServerName(Me.Primary) <> Me.ServerName Then
                        MsgBox(Me.ServerName & " is already exist.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally

        End Try
    End Function

    Public Function ApakahDataMesinBenar() As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally

        End Try
    End Function

    Public Function ApakahBolehEdit() As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally

        End Try
    End Function

    Public Function ApakahBolehHapus() As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally

        End Try
    End Function

    Public Function GetPrimary(ByVal vServerName As String) As Double
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim Primary_ As Nullable(Of Double) = 0
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Primary_ = Dta.ScalarQuery_GetPrimary(vServerName)
            If Primary_.HasValue = True Then
                Return Primary_.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetCreatedDate(ByVal vPrimary As Double) As Date
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim CreatedDate_ As Nullable(Of Date) = Now
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            CreatedDate_ = Dta.ScalarQuery_GetCreatedDate(vPrimary)
            If CreatedDate_.HasValue = True Then
                Return CreatedDate_.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetModifiedDate(ByVal vPrimary As Double) As Date
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim ModifiedDate_ As Nullable(Of Date) = Now
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            ModifiedDate_ = Dta.ScalarQuery_GetModifiedDate(vPrimary)
            If ModifiedDate_.HasValue = True Then
                Return ModifiedDate_.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim Timestamp_ As Nullable(Of Date) = Now
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Timestamp_ = Dta.ScalarQuery_GetCreatedDate(vPrimary)
            If Timestamp_.HasValue = True Then
                Return Timestamp_.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetServerName(ByVal vPrimary As Double) As String
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim ServerName_ As String = ""
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            ServerName_ = Dta.ScalarQuery_GetServerName(vPrimary)
            If String.IsNullOrEmpty(ServerName_) = False Then
                Return ServerName_
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetIsActive(ByVal vPrimary As Double) As Boolean
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim IsActive_ As Nullable(Of Long) = 0
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            IsActive_ = Dta.ScalarQuery_GetIsActive(vPrimary)
            If IsActive_.HasValue = True Then
                If IsActive_.Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetServerLocation(ByVal vPrimary As Double) As String
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim ServerLocation_ As String = ""
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            ServerLocation_ = Dta.ScalarQuery_GetServerLocation(vPrimary)
            If String.IsNullOrEmpty(ServerLocation_) = False Then
                Return ServerLocation_
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetDataSource(ByVal vPrimary As Double) As String
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim DataSource_ As String = ""
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            DataSource_ = Dta.ScalarQuery_GetDataSource(vPrimary)
            If String.IsNullOrEmpty(DataSource_) = False Then
                Return DataSource_
            Else
                Return "localhost"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "localhost"
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetInitialCatalog(ByVal vPrimary As Double) As String
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim InitialCatalog_ As String = ""
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            InitialCatalog_ = Dta.ScalarQuery_GetInitialCatalog(vPrimary)
            If String.IsNullOrEmpty(InitialCatalog_) = False Then
                Return InitialCatalog_
            Else
                Return "NitgenAccessManager"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "NitgenAccessManager"
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetUserId(ByVal vPrimary As Double) As String
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim UserId_ As String = ""
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            UserId_ = Dta.ScalarQuery_GetUserId(vPrimary)
            If String.IsNullOrEmpty(UserId_) = False Then
                Return UserId_
            Else
                Return "sa"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "sa"
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetPassword(ByVal vPrimary As Double) As String
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim Password_ As String = ""
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Password_ = Dta.ScalarQuery_GetPassword(vPrimary)
            If String.IsNullOrEmpty(Password_) = False Then
                Return Password_
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetAcmConnectionString(ByVal vPrimary As Double) As String
        Try
            Return "Data Source=" & Me.GetDataSource(vPrimary) & ";Initial Catalog=" & Me.GetInitialCatalog(vPrimary) & ";Persist Security Info=True;User ID=" & Me.GetUserId(vPrimary) & ";Password=" & Me.GetPassword(vPrimary) & ";Connect Timeout=30"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Data Source=.\SQLEXPRESS;Initial Catalog=NitgenAccessManager;Persist Security Info=True;User ID=sa;Password=;Connect Timeout=30"
        Finally

        End Try
    End Function

    Public Function GetMagicConnectionString(ByVal vPrimary As Double) As String
        Try
            Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Me.GetDataSource(vPrimary) & ";Persist Security Info=True;User ID=" & Me.GetUserId(vPrimary) & ";Jet OLEDB:Database Password=" & Me.GetPassword(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="";Persist Security Info=True;User ID=Admin;Jet OLEDB:Database Password="
        Finally

        End Try
    End Function

    Public Function GetAllPrimary(ByVal vIsActiveOnly As Boolean) As List(Of Double)
        Dim AllPrimary As New List(Of Double)
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim Tbl As New DseMasterAttendanceServer.tblmaster_fingerprint_serverDataTable
        Dim Row As DseMasterAttendanceServer.tblmaster_fingerprint_serverRow
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If vIsActiveOnly = True Then
                Dta.FillBy_ActiveOnly(Tbl)
            Else
                Dta.Fill(Tbl)
            End If
            If Tbl.Count > 0 Then
                For Each Row In Tbl
                    AllPrimary.Add(Row.primary)
                Next
            End If
            Return AllPrimary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetAllConnectionString(ByVal vIsActiveOnly As Boolean) As List(Of String)
        Dim AllConnectionString As New List(Of String)
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim Tbl As New DseMasterAttendanceServer.tblmaster_fingerprint_serverDataTable
        Dim Row As DseMasterAttendanceServer.tblmaster_fingerprint_serverRow
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If vIsActiveOnly = True Then
                Dta.FillBy_ActiveOnly(Tbl)
            Else
                Dta.Fill(Tbl)
            End If
            If Tbl.Count > 0 Then
                For Each Row In Tbl
                    AllConnectionString.Add(Me.GetAcmConnectionString(Row.primary))
                Next
            End If
            Return AllConnectionString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetListActiveTerminalID(ByVal vPrimary As Double) As List(Of String)
        Dim ListActiveTerminal As New List(Of String)
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_server_machinesTableAdapter
        Dim Tbl As New DseMasterAttendanceServer.tblmaster_fingerprint_server_machinesDataTable
        Dim Row As DseMasterAttendanceServer.tblmaster_fingerprint_server_machinesRow
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Dta.FillBy_ActiveTerminalOnly(Tbl, vPrimary) > 0 Then
                For Each Row In Tbl
                    ListActiveTerminal.Add(Row.terminal_id.ToString)
                Next
            End If
            Return ListActiveTerminal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vFrm As FrmMasterAttendanceServer) As Boolean
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim Tbl As New DseMasterAttendanceServer.tblmaster_fingerprint_serverDataTable
        Dim Row As DseMasterAttendanceServer.tblmaster_fingerprint_serverRow
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Dta.FillBy_Primary(Tbl, vPrimary) > 0 Then
                Row = Tbl.Rows(0)
                With Me
                    .Primary = Row.primary
                    .ServerName = Row.server_name
                    .Location = Row.location
                    .IsActive = Row.is_active
                    .Remarks = Row.remarks
                    .DataSource = Row.data_source
                    .InitialCatalog = Row.initial_catalog
                    .UserId = Row.user_id
                    .Password = Row.password
                    .CreatedBy = Row.created_by
                    .CreatedDate = Row.created_date
                    .ModifiedBy = Row.modified_by
                    .ModifiedDate = Row.modified_date
                    .TimeStamp = Row.timestamp
                    .MachineType = Row.machine_type
                End With
                Me.IsiDgvMesin(vPrimary, vFrm)
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Private Function IsiDgvMesin(ByVal vPrimary As Double, ByVal vFrm As FrmMasterAttendanceServer) As Boolean
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_server_machinesTableAdapter
        Dim Tbl As New DseMasterAttendanceServer.tblmaster_fingerprint_server_machinesDataTable
        Dim Row As DseMasterAttendanceServer.tblmaster_fingerprint_server_machinesRow
        Dim DgvRow As New DataGridViewRow
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Dta.FillBy_PrimaryServer(Tbl, vPrimary) > 0 Then
                vFrm.DgvItem.Rows.Clear()
                vFrm.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                vFrm.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                For Each Row In Tbl
                    vFrm.DgvItem.Rows.Add()
                    DgvRow = vFrm.DgvItem.Rows(vFrm.DgvItem.Rows.Count - 1)
                    With Row
                        DgvRow.Cells(vFrm.ColNo.Index).Value = ""
                        DgvRow.Cells(vFrm.ColMachineCode.Index).Value = .machine_code
                        DgvRow.Cells(vFrm.ColMachineName.Index).Value = .machine_name
                        DgvRow.Cells(vFrm.ColTerminalID.Index).Value = .terminal_id
                        If .is_active = 0 Then
                            DgvRow.Cells(vFrm.ColIsActive.Index).Value = Boolean.FalseString
                        Else
                            DgvRow.Cells(vFrm.ColIsActive.Index).Value = Boolean.TrueString
                        End If
                        DgvRow.Cells(vFrm.ColRemarks.Index).Value = .remarks
                    End With
                Next
                Me.ObjCd.SetNumber(vFrm.DgvItem)
                vFrm.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                vFrm.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function IsiDgv(ByVal vCriteria As String, ByVal vSearchBy As ClsMasterAttendanceServer.EnumSearchCriteria, _
            ByVal vDgv As DataGridView, ByVal vTsl As ToolStripLabel) As Boolean
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_server1TableAdapter
        Dim Tbl As New DseMasterAttendanceServer.tblmaster_fingerprint_server1DataTable
        Dim Row As DseMasterAttendanceServer.tblmaster_fingerprint_server1Row
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If vSearchBy = EnumSearchCriteria.BY_SERVERNAME Then
                QueryResult = Dta.FillBy_LikeServerName(Tbl, vCriteria + "%")
            ElseIf vSearchBy = EnumSearchCriteria.BY_LOCATION Then
                QueryResult = Dta.FillBy_LikeLocation(Tbl, vCriteria + "%")
            ElseIf vSearchBy = EnumSearchCriteria.BY_REMARKS Then
                QueryResult = Dta.FillBy_LikeRemarks(Tbl, vCriteria + "%")
            End If
            vDgv.DataSource = Tbl.DefaultView
            vDgv.Refresh()
            If vDgv.ColumnCount > 0 Then
                vDgv.Columns(0).Visible = False
                vDgv.Columns(0).Frozen = True
                vDgv.Columns(1).HeaderText = "Server Name"
                vDgv.Columns(1).Frozen = True
                vDgv.Columns(2).HeaderText = "Location"
                vDgv.Columns(3).HeaderText = "Remarks"
            End If
            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > 10000000 Then
                    vTsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Seconds / 10000) + " seconds"
                Else
                    vTsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / 10000, 0, MidpointRounding.AwayFromZero)) + " miliseconds"
                End If
            Else
                vTsl.Text = "No result found"
            End If
            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Public Function GetListLocation(Optional ByVal vActiveOnly As Boolean = False) As List(Of String)
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim ListLocation As New List(Of String)
        Dim Tbl As New DseMasterAttendanceServer.tblmaster_fingerprint_serverDataTable
        Dim Row As DseMasterAttendanceServer.tblmaster_fingerprint_serverRow
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If vActiveOnly = True Then
                Dta.FillBy_ActiveOnly(Tbl)
            Else
                Dta.Fill(Tbl)
            End If
            For Each Row In Tbl
                If String.IsNullOrEmpty(Row.location) = False Then
                    ListLocation.Add(Row.location)
                End If
            Next
            Return ListLocation
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ListLocation
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetListLocationTerminal(Optional ByVal vActiveOnly As Boolean = False) As List(Of String)
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_server_machines1TableAdapter
        Dim ListLocation As New List(Of String)
        Dim Tbl As New DseMasterAttendanceServer.tblmaster_fingerprint_server_machines1DataTable
        Dim Row As DseMasterAttendanceServer.tblmaster_fingerprint_server_machines1Row
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If vActiveOnly = True Then
                Dta.FillBy_ActiveOnly(Tbl)
            Else
                Dta.Fill(Tbl)
            End If
            For Each Row In Tbl
                If String.IsNullOrEmpty(Row.remarks) = False Then
                    If ListLocation.Contains(Row.remarks) = False Then
                        ListLocation.Add(Row.remarks)
                    End If
                End If
            Next
            Return ListLocation
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ListLocation
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetMachineType(ByVal vPrimary As Double) As ClsMasterAttendanceServer.EnumMachineType
        Dim Dta As New DseMasterAttendanceServerTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim MachineType_ As Nullable(Of Long)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            MachineType_ = Dta.ScalarQuery_GetMachineType(vPrimary)
            If MachineType_.HasValue = True Then
                Return MachineType_.Value
            Else
                Return EnumMachineType.NITGEN
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumMachineType.NITGEN
        Finally
            Dta.Dispose()
        End Try
    End Function

End Class
