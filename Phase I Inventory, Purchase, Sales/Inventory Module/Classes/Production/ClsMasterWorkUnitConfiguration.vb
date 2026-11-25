Option Explicit On
Option Strict Off

Public Class ClsMasterWorkUnitConfiguration

    Dim ObjCD As New Core_Module.ClsCekData

    Public Enum EnumSearchBy As Integer
        CONFIG_CODE
        PRODUCT_CODE
        PRODUCT_DESCRIPTION
    End Enum

    Public Function IsiDGV(ByVal Criteria As String, ByVal DGV As DataGridView, ByVal TBL As DseMasterWorkUnitConfiguration.DataTable1DataTable, _
        ByVal SearchCriteria As ClsMasterWorkUnitConfiguration.EnumSearchBy, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterWorkUnitConfigurationTableAdapters.DataTable1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchBy.CONFIG_CODE Then
                QueryResult = DTA.FillBy_CodeConfigLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchBy.PRODUCT_CODE Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchBy.PRODUCT_DESCRIPTION Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ClsMasterWorkUnitConfiguration.EnumSearchBy) As Double
        Dim DTA As New DseMasterWorkUnitConfigurationTableAdapters.DataTable1TableAdapter
        Dim v As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchBy.CONFIG_CODE Then
                v = DTA.ScalarQuery_GetJmlByCodeConfigLIKE(vKriteria + "%")
                If v.HasValue Then
                    Return v.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchBy.PRODUCT_CODE Then
                v = DTA.ScalarQuery_GetJmlByProductCodeLIKE(vKriteria + "%")
                If v.HasValue Then
                    Return v.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchBy.PRODUCT_DESCRIPTION Then
                v = DTA.ScalarQuery_GetJmlByProductDescLIKE(vKriteria + "%")
                If v.HasValue Then
                    Return v.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsExistCode(ByVal vConfigCode As String) As Boolean
        Dim d As New DseMasterWorkUnitConfigurationTableAdapters.DataTable1TableAdapter
        Dim v As New Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            v = d.ScalarQuery_GetCountByCode(vConfigCode)
            If v.HasValue Then
                If v.Value > 0 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
        End Try
    End Function

    Public Sub FillPropertyByCode(ByVal vMasterLineCode As String, ByRef vPrimaryMasterLine As Object, ByRef vMasterLineDesc As Object, ByRef vRemarks As Object)
        Dim d As New Core_Module.DseMasterMachineLineTableAdapters.tblmaster_machine_lineTableAdapter
        Dim t As New Core_Module.DseMasterMachineLine.tblmaster_machine_lineDataTable
        Dim r As Core_Module.DseMasterMachineLine.tblmaster_machine_lineRow
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring()
            If d.FillBy_Code(t, vMasterLineCode) > 0 Then
                r = t.Rows(0)
                vPrimaryMasterLine = r.primary
                vMasterLineDesc = r.description
                vRemarks = r.remarks
            Else
                vPrimaryMasterLine = String.Empty
                vMasterLineDesc = String.Empty
                vRemarks = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FillPropertyByPrimary(ByVal vPrimaryMasterLine As Double, ByRef vMasterLineCode As Object, ByRef vMasterLineDesc As Object, ByRef vRemarks As Object)
        Dim d As New Core_Module.DseMasterMachineLineTableAdapters.tblmaster_machine_lineTableAdapter
        Dim t As New Core_Module.DseMasterMachineLine.tblmaster_machine_lineDataTable
        Dim r As Core_Module.DseMasterMachineLine.tblmaster_machine_lineRow
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring()
            If d.Fill(t, vPrimaryMasterLine) > 0 Then
                r = t.Rows(0)
                vMasterLineCode = r.code.ToUpper
                vMasterLineDesc = r.description
                vRemarks = r.remarks
            Else
                vMasterLineCode = String.Empty
                vMasterLineDesc = String.Empty
                vRemarks = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
