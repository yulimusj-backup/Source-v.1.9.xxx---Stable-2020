''Aul 9 nov 2015
Option Explicit On
Option Strict Off

Public Class ClsMasterKompetensi_Product_Item
    Public Function IsiDGV(ByVal vPrimaryKompetensi As Double, ByVal vDGVItem As DataGridView, _
          ByVal i As FrmMasterKompetensi_Product, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseMasterKompetensi_ProductTableAdapters.tblmaster_kompetensi_product_itemTableAdapter
        Dim TBLItem As New DseMasterKompetensi_Product.tblmaster_kompetensi_product_itemDataTable
        Dim rowItem As DseMasterKompetensi_Product.tblmaster_kompetensi_product_itemRow
        Dim DefaultPrecision As Integer = 0
        Try
            'update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item
            DTAItem.FillBy_PrimaryKompetensi(TBLItem, vPrimaryKompetensi)
            vDGVItem.Rows.Clear()

            vDGVItem.AllowUserToAddRows = False
            For Each rowItem In TBLItem.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColNo.Index).Value = ""
                    .Cells(i.ColPrimaryKompetensi.Index).Value = rowItem.Primary_Kompetensi
                    .Cells(i.ColIdentifier.Index).Value = rowItem.Identifier
                    .Cells(i.ColCode.Index).Value = rowItem.Code
                    .Cells(i.ColDesc.Index).Value = rowItem.Desc
                    .Cells(i.ColRemarks.Index).Value = rowItem.Remarks
                End With
            Next
            vDGVItem.AllowUserToAddRows = True

            'set number
            vObjCD.SetNumber(vDGVItem)

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterKompetensi_ProductTableAdapters.tblmaster_kompetensi_product_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmMasterKompetensi_Product) As Boolean
        Dim DTAItem As New DseMasterKompetensi_ProductTableAdapters.tblmaster_kompetensi_product_itemTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            vDGVItem.AllowUserToAddRows = False
            For Each rowDGV In vDGVItem.Rows
                With rowDGV
                    If rowDGV.Cells(i.ColRemarks.Index).Value = "" Then
                        rowDGV.Cells(i.ColRemarks.Index).Value = ""
                    End If
                    DTAItem.Insert(vPrimary, .Cells(i.ColIdentifier.Index).Value, _
                    .Cells(i.ColCode.Index).Value, .Cells(i.ColDesc.Index).Value, _
                    .Cells(i.ColRemarks.Index).Value)
                End With
            Next
            vDGVItem.AllowUserToAddRows = True
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetNextIdentifier(ByVal vDGVItem As DataGridView, _
    ByVal vPrimaryKompetensi_Product As String, ByVal Frm As FrmMasterKompetensi_Product) As String
        Dim rowDGV As DataGridViewRow
        Dim MaxVal As Double = 0
        Dim Index As Integer = -1
        Try
            For Each rowDGV In vDGVItem.Rows
                If rowDGV.Cells(Frm.ColIdentifier.Index).Value <> "" Then
                    Index = rowDGV.Cells(Frm.ColIdentifier.Index).Value.ToString.LastIndexOf("-")
                    If Index <> -1 Then
                        If MaxVal < CDbl(rowDGV.Cells(Frm.ColIdentifier.Index).Value.ToString.Substring(Index + 1)) Then
                            MaxVal = rowDGV.Cells(Frm.ColIdentifier.Index).Value.ToString.Substring(Index + 1)
                        End If
                    End If
                End If
            Next
            Return vPrimaryKompetensi_Product + "-" + (MaxVal + 1).ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "-1-0"
        End Try
    End Function

    Public Function GetJmlCustVal(ByVal vPrimaryKompetensi As Double, ByVal vKriteria As String, ByVal vSearchBy As EnumSearchCriteria) As Double
        Dim DTA As New DseMasterKompetensi_ProductTableAdapters.tblmaster_kompetensi_product_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                Return DTA.ScalarQuery_GetJmlByPrimaryKompetensiNCodeLike(vPrimaryKompetensi, vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                Return DTA.ScalarQuery_GetJmlByPrimaryKompetensiNDescLike(vPrimaryKompetensi, vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlByPrimaryKompetensiNRemarksLike(vPrimaryKompetensi, vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGVCustVal(ByVal vPrimaryKompetensi As Double, ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterKompetensi_Product.tblmaster_kompetensi_product_itemDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterKompetensi_ProductTableAdapters.tblmaster_kompetensi_product_itemTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date = Now

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_PrimaryKompetensiNCodeLike(TBL, vPrimaryKompetensi, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_PrimaryKompetensiNDescLike(TBL, vPrimaryKompetensi, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_PrimaryKompetensiNRemarksLike(TBL, vPrimaryKompetensi, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).Visible = False
                DGV.Columns(2).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(2).Frozen = True
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > TimeSpan.TicksPerSecond Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Seconds) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, _
                    0, MidpointRounding.AwayFromZero)) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vPrimaryCustVal As Double, ByVal vDataCode As String) As String
        Dim DTA As New DseMasterKompetensi_ProductTableAdapters.tblmaster_kompetensi_product_itemTableAdapter
        Dim retval As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retval = DTA.ScalarQuery_GetJmlByPrimaryKompetensiNCode(vPrimaryCustVal, vDataCode)
            If IsNothing(retval) = True Then
                Return -1
            Else
                Return retval
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDesc_ByIdentifier(ByVal vIdentifier As String) As String
        Dim DTA As New DseMasterKompetensi_ProductTableAdapters.tblmaster_kompetensi_product_itemTableAdapter
        Dim retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retval = DTA.ScalarQuery_GetDescByIdentifier(vIdentifier)
            If String.IsNullOrEmpty(retval) = True Then
                Return ""
            Else
                Return retval
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetIdentifier_ByCode(ByVal vPrimaryCustVal As Double, ByVal vDataCode As String) As String
        Dim DTA As New DseMasterKompetensi_ProductTableAdapters.tblmaster_kompetensi_product_itemTableAdapter
        Dim retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retval = DTA.ScalarQuery_GetIdentifierByCode(vDataCode, vPrimaryCustVal)
            If String.IsNullOrEmpty(retval) = True Then
                Return ""
            Else
                Return retval
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCode_ByIdentifier(ByVal vIdentifier As String) As String
        Dim DTA As New DseMasterKompetensi_ProductTableAdapters.tblmaster_kompetensi_product_itemTableAdapter
        Dim retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retval = DTA.ScalarQuery_GetCodeByIdentifier(vIdentifier)

            If String.IsNullOrEmpty(retval) = True Then
                Return ""
            Else
                Return retval
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
