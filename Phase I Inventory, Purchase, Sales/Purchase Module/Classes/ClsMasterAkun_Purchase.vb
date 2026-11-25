Option Explicit On
Option Strict Off

Public Class ClsMasterAkun_Purchase
    Dim ObjCD As New Core_Module.ClsCekData

    Public Enum EnumTipeAkun As Integer
        Asset
        Liabilities
        Equity
        Income
        Expense
        Cost_of_Goods_Sales
        Other_Income
        Other_Expense
        Tax
    End Enum

    Public Enum EnumTipeAkun2 As Integer
        Cash_And_Bank
        Account_Receivable
        Other_Receivable
        Inventory
        Fixed_Asset
        Accumulated_Depreciation
        Other_Asset
        Account_Payable
        Other_Current_Liability
        Long_Term_Liability
        Equity
        Income
        Cost_Of_Sales
        Expense
        Other_Income
        Other_Expense
        Tax
        Asset
        Liabilities
        Other_Payable
    End Enum

    Public Function GetTipeAkun2(ByVal vPrimary As Double) As EnumTipeAkun2
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetTipeAkun2(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTipeAkun2(vPrimary).Value
            Else
                Return EnumTipeAkun.Asset
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumTipeAkun2.Account_Receivable
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipeAkun(ByVal vPrimary As Double) As EnumTipeAkun
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetTipePerkiraan(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTipePerkiraan(vPrimary).Value
            Else
                Return EnumTipeAkun.Asset
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumTipeAkun.Asset
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda_Aktif_NonHeader(ByVal vAccountCode As String) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml_Aktif_NonHeader(vAccountCode).HasValue Then
                If DTA.ScalarQuery_GetJml_Aktif_NonHeader(vAccountCode).Value > 0 Then
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

    Public Function IsiDGV(ByVal Criteria As String, _
        ByVal DGV As DataGridView, ByVal TBL As DseMasterAkun.tbl_Acari_master_akunDataTable, _
        ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByAccountNo Then
                QueryResult = DTA.FillBy_AccountCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Account Code"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Description"
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ModGlobalVariable.EnumSearchCriteria) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByAccountNo Then
                Return DTA.ScalarQuery_GetJmlByAccountCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                Return DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAccountCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetAccountCodeByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetAccountCodeByPrimary(vPrimary)
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

    Public Function GetJml_NonHeader_Aktif(ByVal vKodeAkun As String) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml_Aktif_NonHeader(vKodeAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetJml_Aktif_NonHeader(vKodeAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDescription(ByVal vKodeAkun As String) As String
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDesc_ByAccCode(vKodeAkun)) = False Then
                Return DTA.ScalarQuery_GetDesc_ByAccCode(vKodeAkun)
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

    Public Function GetDescription2(ByVal vPrimaryAkun As Double) As String
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDesc(vPrimaryAkun)) = False Then
                Return DTA.ScalarQuery_GetDesc(vPrimaryAkun)
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

    Public Function GetPrimary(ByVal vAccountCode As String) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vAccountCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vAccountCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPrimaryCurrency(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryAkun).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryAkun).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function
End Class
