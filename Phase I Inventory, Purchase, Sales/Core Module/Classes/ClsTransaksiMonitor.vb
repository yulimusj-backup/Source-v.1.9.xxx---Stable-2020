Option Explicit On
Option Strict Off

Public Class ClsTransaksiMonitor
    Public MonitorInterval As Double = 3 * 60000
    Public DeleteInterval As Double = 5 * 60000

    Public Enum EnumTipeTransaksi As Integer
        SMESS_License
        MO_Edit
        Custom_ImportEfaktur
        Custom_ImportArt
    End Enum

    Public Function GetSMESSActiveLicense() As Double
        Dim DTA As New DseTransaksiMonitorTableAdapters.tbl_smessTableAdapter
        Dim TBL As New DseTransaksiMonitor.tbl_smessDataTable
        Dim row As DseTransaksiMonitor.tbl_smessRow
        Dim JmlLicense As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                JmlLicense = row.active_license

                Return JmlLicense
            End If

            Return JmlLicense
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlMonitor(ByVal vTipeMonitor As EnumTipeTransaksi) As Double
        Dim DTA As New DseTransaksiMonitorTableAdapters.tbltrans_monitorTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByTipeTrans(CInt(vTipeMonitor))
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahLicenseSMESSCukup(Optional ByVal vFromLogin As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiMonitorTableAdapters.tbl_smessTableAdapter
        Dim TBL As New DseTransaksiMonitor.tbl_smessDataTable
        Dim row As DseTransaksiMonitor.tbl_smessRow
        Dim JmlOnline As Double = 0
        Dim JmlLicense As Double = 1
        Dim DTA1 As New DseTransaksiMonitorTableAdapters.tbltrans_monitorTableAdapter
        Dim ObjMySQL As New Core_Module.ClsMySQLReference
        Dim TipeTrans As EnumTipeTransaksi = EnumTipeTransaksi.SMESS_License

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            'hapuskan item yang time key > 5 menit
            DTA1.DeleteQuery(ObjMySQL.GetServerTime - TimeSpan.FromMilliseconds(Me.DeleteInterval))

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                JmlLicense = row.active_license
                JmlOnline = Me.GetJmlMonitor(EnumTipeTransaksi.SMESS_License)
                If vFromLogin = True Then
                    JmlOnline += 1
                End If

                If JmlOnline > JmlLicense Then
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjMySQL = Nothing
        End Try
    End Function

    Public Function ApakahAdaMonitor(ByVal vTipeTransaksi As EnumTipeTransaksi) As Boolean
        Dim DTA As New Core_Module.DseTransaksiMonitorTableAdapters.tbltrans_monitorTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByTipeTrans(vTipeTransaksi)
            If RetVal.HasValue = True Then
                If RetVal.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahAdaMonitor_SelainTipeYgDiberikan(ByVal vTipeTransaksi As EnumTipeTransaksi) As Boolean
        Dim DTA As New Core_Module.DseTransaksiMonitorTableAdapters.tbltrans_monitorTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByTipeTransSelain(vTipeTransaksi)
            If RetVal.HasValue = True Then
                If RetVal.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub StartMonitor(ByVal vUserName As String, ByVal vNamaKomp As String, ByVal vTipeTrans As EnumTipeTransaksi, ByVal vIdentifier As Double, Optional ByVal vPrimaryReference As Double = -1)
        Dim DTA As New DseTransaksiMonitorTableAdapters.tbltrans_monitorTableAdapter
        Dim Jml As Nullable(Of Long)
        Dim ApakahInsert As Boolean = True
        Dim IpAdd As String = ""
        Dim Timestamp As Date = Now
        Dim ObjMySQL As New Core_Module.ClsMySQLReference

        Try
            Timestamp = ObjMySQL.GetServerTime
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Jml = DTA.ScalarQuery_GetJml(CInt(vTipeTrans), vIdentifier)
            If Jml.HasValue = True Then
                If Jml.Value = 0 Then
                    ApakahInsert = True
                Else
                    ApakahInsert = False
                End If
            Else
                ApakahInsert = True
            End If
            If ApakahInsert = True Then
                IpAdd = ""
                If System.Net.Dns.GetHostAddresses("").Length > 2 Then
                    'IpAdd = System.Net.Dns.GetHostAddresses("")(2).ToString
                    For xx As Integer = 0 To 2
                        If System.Net.Dns.GetHostAddresses("")(xx).ToString.StartsWith("192") = True Then
                            IpAdd = System.Net.Dns.GetHostAddresses("")(xx).ToString
                        End If
                    Next
                    If IpAdd = "" Then
                        IpAdd = "-"
                    End If
                Else
                    IpAdd = "-"
                End If

                DTA.Insert(vUserName, My.Computer.Name, vTipeTrans, Timestamp, vIdentifier, IpAdd, vPrimaryReference)
            Else
                DTA.Update(Timestamp, vIdentifier, vTipeTrans)
            End If

            'hapuskan item yang time key > 5 menit
            DTA.DeleteQuery(Timestamp - TimeSpan.FromMilliseconds(Me.DeleteInterval))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjMySQL = Nothing
        End Try
    End Sub

    'Public Sub EndMonitor(ByVal vNamaKomputer As String, ByVal vNamaUser As String, ByVal vIdentifier As Decimal)
    '    Dim DTA As New DseTransaksiMonitorTableAdapters.tbltrans_monitorTableAdapter

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA.DeleteQuery1(vNamaKomputer, vNamaUser, vIdentifier)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Sub

    Public Sub EndMonitor(ByVal vTipeTrans As EnumTipeTransaksi, ByVal vIdentifier As Double)
        Dim DTA As New DseTransaksiMonitorTableAdapters.tbltrans_monitorTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(vTipeTrans, vIdentifier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetTipeTransaksiStr(ByVal vTipeTrans As EnumTipeTransaksi) As String
        Try
            If vTipeTrans = EnumTipeTransaksi.SMESS_License Then
                Return "SMESS License"
            ElseIf vTipeTrans = EnumTipeTransaksi.MO_Edit Then
                Return "MO Edit"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub ShowMonitor(ByVal vDGV As DataGridView, ByVal vFrm As FrmTransaksiMonitor, ByVal vEnumTipeTransaksi As EnumTipeTransaksi)
        Dim DTA As New DseTransaksiMonitorTableAdapters.tbltrans_monitorTableAdapter
        Dim TBL As New DseTransaksiMonitor.tbltrans_monitorDataTable
        Dim row As DseTransaksiMonitor.tbltrans_monitorRow
        Dim rowDGV As DataGridViewRow

        Try
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDGV.Rows.Clear()
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_SortUser(TBL, CInt(vEnumTipeTransaksi))
            For Each row In TBL.Rows
                vDGV.Rows.Add()
                rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                rowDGV.Cells(vFrm.ColNo.Index).Value = rowDGV.Index + 1
                rowDGV.Cells(vFrm.ColIpAdd.Index).Value = row.ip
                rowDGV.Cells(vFrm.ColNamaKomp.Index).Value = row.nama_komp
                rowDGV.Cells(vFrm.ColTipeTrans.Index).Value = Me.GetTipeTransaksiStr(row.tipe_transaksi)
                rowDGV.Cells(vFrm.ColUserName.Index).Value = row.user_name
            Next

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function IsTheFirstOne(ByVal vTipeTrans As EnumTipeTransaksi, ByVal vIdentifier As Decimal, ByRef vFirstUserCode As String, _
    ByVal vPrimaryReference As Double) As Boolean
        Dim DTA As New DseTransaksiMonitorTableAdapters.tbltrans_monitorTableAdapter
        Dim TBL As New DseTransaksiMonitor.tbltrans_monitorDataTable
        Dim row As DseTransaksiMonitor.tbltrans_monitorRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_TipeTransaksi(TBL, vTipeTrans, vPrimaryReference)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                vFirstUserCode = row.user_name
                If row.identifier = vIdentifier Then
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

    Public Function ApakahAdaMonitor2(ByVal vTipeTransaksi As EnumTipeTransaksi, ByVal vUserName As String, ByVal vIdentifier As Integer) As Boolean
        Dim DTA As New Core_Module.DseTransaksiMonitorTableAdapters.tbltrans_monitorTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlBy1(vUserName, vTipeTransaksi, vIdentifier)
            If RetVal.HasValue = True Then
                If RetVal.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
