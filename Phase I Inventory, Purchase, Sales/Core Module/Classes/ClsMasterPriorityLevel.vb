Option Explicit On
Option Strict Off

Public Class ClsMasterPriorityLevel

    Public Enum enumDefault As Integer
        no
        yes
    End Enum

    Public Function GetPrimary_ByDesc(ByVal vDesc As String) As Double
        Dim DTA As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Dim Retval As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimaryByDesc(vDesc)
            If Retval.HasValue = True Then
                Return Retval.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetUrutan(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Dim Retval As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetUrutan(vPrimary)
            If Retval.HasValue = True Then
                Return Retval.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetDescPriorityLevel(vPrimary)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub GetListDescPriorityLevel(ByVal vList As List(Of String))
        Dim DTA As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Dim TBL As New DseMasterPriorityLevel1.tblmaster_priority_levelDataTable
        Dim row As DseMasterPriorityLevel1.tblmaster_priority_levelRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vList.Clear()
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            For Each row In TBL.Rows
                vList.Add(row.desc)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetPrimaryDefault() As Double
        Dim DTA As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryDefault
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function refreshAllData(ByVal vDgv As DataGridView, ByVal vIndexPrimary As Integer, ByVal vIndexPriority As Integer, _
    ByVal vIndexCode As Integer, ByVal vIndexDesc As Integer, ByVal vIndexDefault As Integer) As Boolean
        Dim dta As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Dim tbl As New DseMasterPriorityLevel1.tblmaster_priority_levelDataTable
        Dim row As DseMasterPriorityLevel1.tblmaster_priority_levelRow
        Dim rowDgv As New DataGridViewRow

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Fill(tbl) > 0 Then
                vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                vDgv.AllowUserToAddRows = False
                vDgv.Rows.Clear()
                For Each row In tbl
                    With row
                        vDgv.Rows.Add()
                        rowDgv = vDgv.Rows(vDgv.Rows.Count - 1)
                        rowDgv.Cells(vIndexPrimary).Value = .primary
                        rowDgv.Cells(vIndexPriority).Value = .urutan
                        rowDgv.Cells(vIndexCode).Value = .code
                        rowDgv.Cells(vIndexDesc).Value = .desc
                        If ._default = ClsMasterPriorityLevel.enumDefault.no Then
                            rowDgv.Cells(vIndexDefault).Value = False
                        Else
                            rowDgv.Cells(vIndexDefault).Value = True
                        End If
                    End With
                Next
                vDgv.AllowUserToAddRows = True
                vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dta.Dispose()
            tbl.Dispose()
            row = Nothing
            rowDgv.Dispose()
        End Try
    End Function

    Public Function refreshAllData(ByVal vFrm As FrmMasterPriorityLevel, ByVal vDgv As DataGridView) As Boolean
        Dim dta As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Dim tbl As New DseMasterPriorityLevel1.tblmaster_priority_levelDataTable
        Dim row As DseMasterPriorityLevel1.tblmaster_priority_levelRow
        Dim rowDgv As New DataGridViewRow

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Fill(tbl) > 0 Then
                vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                vDgv.AllowUserToAddRows = False
                vDgv.Rows.Clear()
                For Each row In tbl
                    With row
                        vDgv.Rows.Add()
                        rowDgv = vDgv.Rows(vDgv.Rows.Count - 1)
                        rowDgv.Cells(vFrm.ColPrimary.Index).Value = .primary
                        rowDgv.Cells(vFrm.ColPriority.Index).Value = .urutan
                        rowDgv.Cells(vFrm.ColCode.Index).Value = .code
                        rowDgv.Cells(vFrm.ColDescription.Index).Value = .desc
                        If ._default = ClsMasterPriorityLevel.enumDefault.no Then
                            rowDgv.Cells(vFrm.ColDefault.Index).Value = False
                        Else
                            rowDgv.Cells(vFrm.ColDefault.Index).Value = True
                        End If
                    End With
                Next
                vDgv.AllowUserToAddRows = True
                vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dta.Dispose()
            tbl.Dispose()
            row = Nothing
            rowDgv.Dispose()
        End Try
    End Function

    Public Function GetJml() As Double
        Dim dta As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Return dta.ScalarQuery_GetJml1.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function simpanBaru(ByVal vUrutan As Integer, ByVal vCode As String, ByVal vDesc As String, ByVal vUserCode As String) As Boolean
        Dim dta As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Insert(vUrutan, vCode, vDesc, ClsMasterPriorityLevel.enumDefault.no, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function apakahSudahAda(ByVal vCode As String) As Boolean
        Dim dta As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Dim temp As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            temp = dta.ScalarQuery_GetJmlByCode(vCode).Value
            If temp.HasValue = True Then
                If temp.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function getPrimary(ByVal vCode As String) As Double
        Dim dta As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Dim temp As Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            temp = dta.ScalarQuery_GetPrimary(vCode).Value
            If temp.HasValue = True Then
                Return temp
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function updateDefault(ByVal vCode As String, ByVal vDesc As String, ByVal vUser As String, ByVal vPrimary As Double) As Boolean
        Dim dta As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.UpdateQuery_Default(vCode, vDesc, vUser, Now, Now, vPrimary) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function setDefaultAll(ByVal vUserCode As String, ByVal vPrimary As Double) As Boolean
        Dim dta As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_SetDefaultAll(ClsMasterPriorityLevel.enumDefault.no)
            dta.UpdateQuery_SetDefaultByPrimary(ClsMasterPriorityLevel.enumDefault.yes, vUserCode, Now, Now, vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Dim Jml As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cek pada item so
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Jml = DTA.ScalarQuery_GetJml2(vPrimary)
            If Jml.HasValue = False Then
                Return True
            Else
                If Jml.Value > 0 Then
                    If vShowMsg = True Then
                        MsgBox("This data is used on sales order, cannot delete this data.", MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                Else
                    Jml = DTA.ScalarQuery_GetJml3(vPrimary)
                    If Jml.HasValue = False Then
                        Return True
                    Else
                        If Jml.Value > 0 Then
                            If vShowMsg = True Then
                                MsgBox("This data is used on purchase request, cannot delete this data.", MsgBoxStyle.Information, "Information")
                            End If

                            Return False
                        Else
                            Return True
                        End If
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function setDefaultByPrimary(ByVal vUserCode As String, ByVal vDef As ClsMasterPriorityLevel.enumDefault, ByVal vPrimary As Double) As Boolean
        Dim dta As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_SetDefaultByPrimary(vDef, vUserCode, Now, Now, vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function repairPriority(ByVal vDgv As DataGridView, ByVal vIndexPrimary As Integer, ByVal vIndexPriority As Integer) As Boolean
        Dim dta As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Dim rowDgv As DataGridViewRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            For Each rowDgv In vDgv.Rows
                If rowDgv.Cells(vIndexPrimary).Value.ToString <> "-1" Then
                    dta.UpdateQuery_SetUrutanByPrimary(CInt(rowDgv.Cells(vIndexPriority).Value), CDbl(rowDgv.Cells(vIndexPrimary).Value))
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function repairPriority(ByVal vFrm As FrmMasterPriorityLevel, ByVal vDgv As DataGridView) As Boolean
        Dim dta As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Dim rowDgv As DataGridViewRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            For Each rowDgv In vDgv.Rows
                If rowDgv.Cells(vFrm.ColPrimary.Index).Value.ToString <> "-1" Then
                    dta.UpdateQuery_SetUrutanByPrimary(CInt(rowDgv.Cells(vFrm.ColPriority.Index).Value), CDbl(rowDgv.Cells(vFrm.ColPrimary.Index).Value))
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function hapus(ByVal vPrimary As Double) As Boolean
        Dim dta As New DseMasterPriorityLevel1TableAdapters.tblmaster_priority_levelTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Delete(vPrimary) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dta.Dispose()
        End Try
    End Function
End Class
