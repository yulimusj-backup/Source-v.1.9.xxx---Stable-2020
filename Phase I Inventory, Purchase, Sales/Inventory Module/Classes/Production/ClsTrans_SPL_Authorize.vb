Option Explicit On
Option Strict Off
Public Class ClsTrans_SPL_Authorize
    Private ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
    Private ObjCD As New Core_Module.ClsCekData
    Private objSPLAuthAccess As New ClsTrans_SPL_Authorize_Access
    Private objSPLAuthAccessItem As New ClsTrans_SPL_Authorize_AccessItem
    Private objSetting As New Core_Module.ClsTransaksiSetting

    Dim filterPrimaryBundling As New List(Of Double)
    

    Private Sub FilterCreatedBy(ByRef vTBL As DseTrans_SPL_Authorize.tbltrans_splDataTable, ByVal vFrm As FrmTrans_SPL_Authorize)
        Dim vRow As DseTrans_SPL_Authorize.tbltrans_splRow
        Dim tblSPL As New DseTrans_SPL_Authorize.tbltrans_splDataTable
        Dim rowSPL As DseTrans_SPL_Authorize.tbltrans_splRow
        Dim filterUser As String = ""
        Try
            'pindahkan isi tbl dari DB ke data table
            For Each vRow In vTBL
                rowSPL = tblSPL.Newtbltrans_splRow
                For i As Integer = 0 To vTBL.Columns.Count - 1
                    rowSPL.Item(i) = vRow.Item(i)
                Next
                tblSPL.Rows.Add(rowSPL)
            Next

            For Each rd As DataGridViewRow In vFrm.DgvUser.Rows
                If rd.Cells(vFrm.ColCreatedBy.Index).Value <> "" Then
                    If filterUser = "" Then
                        filterUser = "[created by] = '" & rd.Cells(vFrm.ColCreatedBy.Index).Value & "'"
                    Else
                        filterUser = filterUser & " OR [created by] = '" & rd.Cells(vFrm.ColCreatedBy.Index).Value & "'"
                    End If
                End If
            Next
            tblSPL.DefaultView.RowFilter = filterUser
            vTBL.Rows.Clear()
            For i As Integer = 0 To tblSPL.DefaultView.Count - 1
                rowSPL = tblSPL.DefaultView(i).Row
                vRow = vTBL.Newtbltrans_splRow
                For i2 As Integer = 0 To tblSPL.Columns.Count - 1
                    vRow.Item(i2) = rowSPL.Item(i2)
                Next
                vTBL.Rows.Add(vRow)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub FilterCreatedBy(ByRef vTBL As DseTrans_SPL_Authorize.tbltrans_splDataTable, ByVal vList As List(Of String))
        Dim vRow As DseTrans_SPL_Authorize.tbltrans_splRow
        Dim tblSPL As New DseTrans_SPL_Authorize.tbltrans_splDataTable
        Dim rowSPL As DseTrans_SPL_Authorize.tbltrans_splRow
        Dim filterUser As String = ""
        Try
            '//buat data table spl
            For Each vRow In vTBL
                rowSPL = tblSPL.Newtbltrans_splRow
                For i As Integer = 0 To vTBL.Columns.Count - 1
                    rowSPL.Item(i) = vRow.Item(i)
                Next
                tblSPL.Rows.Add(rowSPL)
            Next
            '//Cek Filter berdasarkan user di list
            For Each user As String In vList
                If filterUser = "" Then
                    filterUser = "[created by] = '" & user & "'"
                Else
                    filterUser = filterUser & " OR [created by] = '" & user & "'"
                End If
            Next

            tblSPL.DefaultView.RowFilter = filterUser
            vTBL.Rows.Clear()
            For i As Integer = 0 To tblSPL.DefaultView.Count - 1
                rowSPL = tblSPL.DefaultView(i).Row
                vRow = vTBL.Newtbltrans_splRow
                For i2 As Integer = 0 To tblSPL.Columns.Count - 1
                    vRow.Item(i2) = rowSPL.Item(i2)
                Next
                vTBL.Rows.Add(vRow)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Public Function IsiDGVAuthorize(ByVal vPrimarySPL As Double, ByVal vUserCode As String, ByVal vDGVAuthorize0 As DataGridView, ByVal vDGVAuthorize1 As DataGridView, ByVal vDGVAuthorize2 As DataGridView, _
                                  ByVal vFrm As FrmTrans_SPL_Authorize) As Boolean


        Dim ObjCD As New Core_Module.ClsCekData
        Dim objUser As New Core_Module.ClsUser
        Dim objSPL As New ClsTrans_SPL
        Dim DTA As New DseTrans_SPL_AuthorizeTableAdapters.tbltrans_splTableAdapter
        Dim TBL As New DseTrans_SPL_Authorize.tbltrans_splDataTable
        Dim rowItem As DseTrans_SPL_Authorize.tbltrans_splRow
        Dim rowDGV As DataGridViewRow

        Dim vAuthStatus1, vAuthStatus2 As Integer
        Dim vFrom, vTo As Date
        Try

            vDGVAuthorize0.Rows.Clear()
            vDGVAuthorize1.Rows.Clear()
            vDGVAuthorize2.Rows.Clear()

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGVAuthorize0.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVAuthorize0.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVAuthorize1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVAuthorize1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVAuthorize2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVAuthorize2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            '//Cek Filter berdasarkan tanggal
            If vFrm.ChkSPLDate.Checked = True Then
                vFrom = vFrm.DtpFrom.Value.Date
                vTo = vFrm.DtpTo.Value.Date
            Else
                vFrom = Date.MinValue
                vTo = Date.MaxValue
            End If

            '//Cek Filter berdasarkan status authorize
            If vFrm.TboAuthorize.SelectedIndex = 0 Then
                vAuthStatus1 = 0
                vAuthStatus2 = 0
            ElseIf vFrm.TboAuthorize.SelectedIndex = 1 Then
                vAuthStatus1 = 1
                vAuthStatus2 = 1
            ElseIf vFrm.TboAuthorize.SelectedIndex = 2 Then
                vAuthStatus1 = 0
                vAuthStatus2 = 1
            End If

            '//Cel Filter berdasarkan combo box authorized
            If vFrm.CboStatus.Enabled = True Then
                If vFrm.CboStatus.SelectedIndex = 0 Then
                    vAuthStatus1 = 0
                    vAuthStatus2 = 1
                ElseIf vFrm.CboStatus.SelectedIndex = 1 Then
                    vAuthStatus1 = 1
                    vAuthStatus2 = 1
                ElseIf vFrm.CboStatus.SelectedIndex = 2 Then
                    vAuthStatus1 = 0
                    vAuthStatus2 = 0
                End If
            End If

            '//Cek Filter berdasarkan Bukan No Dokumen 
            If vFrm.TxtNoSPL.Text = "" Then
                DTA.Fill(TBL, vAuthStatus1, vAuthStatus2, vFrom, vTo)
            Else '//Cek Filter berdasarkan No dokumen
                DTA.FillBy_DocumentNo(TBL, vAuthStatus1, vAuthStatus2, vFrom, vTo, vPrimarySPL)
            End If

            '//Cek Filter berdasarkan user
            If vFrm.DgvUser.RowCount > 1 Then
                FilterCreatedBy(TBL, vFrm)
            End If

            '//Cek Setingan Restrict
            If Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Overtime_Order, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Activate_Access_Retriction) = "1" Then ' restric unchecked
                '//jika status restrict ceklis, maka cek hak akses
                If Me.objSPLAuthAccess.ApakahAda(vUserCode) = False Then
                    Return False
                Else
                    '//Cek Filter Hak akses berdasarkan selected user only
                    If Me.objSPLAuthAccess.GetAuthType(vUserCode) = 1 Then
                        FilterCreatedBy(TBL, Me.objSPLAuthAccess.GetListOfUser(vUserCode))
                    End If
                End If
            End If


            For Each rowItem In TBL

                If vFrm.TboAuthorize.SelectedIndex = 0 Then
                    vDGVAuthorize0.Rows.Add()
                    rowDGV = vDGVAuthorize0.Rows(vDGVAuthorize0.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.ColDgvAuth_Primary.Index).Value = rowItem.primary
                        .Cells(vFrm.ColDgvAuth_SPLNo.Index).Value = rowItem.document_no
                        .Cells(vFrm.ColDgvAuth_Date.Index).Value = Format(rowItem._date, "dd MMM yyyy")
                    End With
                ElseIf vFrm.TboAuthorize.SelectedIndex = 1 Then
                    If apakahSudahBundling(rowItem.primary) = True Then
                        If objUser.GetUserCode(rowItem.primary_user_auth) = vUserCode Then
                            vDGVAuthorize1.Rows.Add()
                            rowDGV = vDGVAuthorize1.Rows(vDGVAuthorize1.Rows.Count - 1)
                            With rowDGV
                                .Cells(vFrm.ColDgvUnAuth_Primary.Index).Value = rowItem.primary
                                .Cells(vFrm.ColDgvUnAuth_SPLNo.Index).Value = rowItem.document_no
                                .Cells(vFrm.ColDgvUnAuth_Date.Index).Value = Format(rowItem._date, "dd MMM yyyy")
                            End With
                        End If
                    End If

                ElseIf vFrm.TboAuthorize.SelectedIndex = 2 Then
                    vDGVAuthorize2.Rows.Add()
                    rowDGV = vDGVAuthorize2.Rows(vDGVAuthorize2.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.col3_Primary.Index).Value = rowItem.primary
                        .Cells(vFrm.col3_DocNo.Index).Value = rowItem.document_no
                        If rowItem.is_authorized = 0 Then
                            .Cells(vFrm.col3_Status.Index).Value = "Un-Authorize"
                        ElseIf rowItem.is_authorized = 1 Then
                            .Cells(vFrm.col3_Status.Index).Value = "Authorize by " & rowItem.Kode_User & ""
                        End If
                        .Cells(vFrm.col3_Date.Index).Value = Format(rowItem._date, "dd MMM yyyy")
                    End With
                End If
            Next


            vDGVAuthorize0.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVAuthorize0.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            vDGVAuthorize1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVAuthorize1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            vDGVAuthorize2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVAuthorize2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Function

    

    Public Function updateSPL(ByVal vAuthorized As Integer, ByVal vPrimaryAuth As Double, ByVal vNodokumen As String) As Boolean
        Dim DTA As New DseTrans_SPL_AuthorizeTableAdapters.tbltrans_splTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.UpdateQuery_SPLAuthorized(vAuthorized, vPrimaryAuth, vNodokumen) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function apakahSudahBundling(ByVal vPrimarySPL As Double) As Boolean
        Dim dtaBDL As New DseTrans_SPLTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim tblBDL As New DseTrans_SPL.tbltransaksi_stockbundlingDataTable
        Dim dtaItemLOG As New DseTrans_SPLTableAdapters.tbltrans_spl_item_logTableAdapter
        Dim tblItemLOG As New DseTrans_SPL.tbltrans_spl_item_logDataTable
        Dim rowItemLOG As DseTrans_SPL.tbltrans_spl_item_logRow


        Try
            dtaBDL.Connection.ConnectionString = Core_Module.GetConstring
            dtaItemLOG.Connection.ConnectionString = Core_Module.GetConstring

            dtaBDL.Fill(tblBDL)
            dtaItemLOG.Fill(tblItemLOG)

            tblItemLOG.DefaultView.RowFilter = "[primary spl] = " & vPrimarySPL & ""
            If tblItemLOG.DefaultView.Count > 0 Then
                For i As Integer = 0 To tblItemLOG.DefaultView.Count - 1
                    rowItemLOG = tblItemLOG.DefaultView(i).Row
                    tblBDL.DefaultView.RowFilter = "[primary work unit] = " & rowItemLOG.primary_wokunit & " AND [end time actual] >= #" & CStr(rowItemLOG._date) & "# AND [start time actual] <= #" & CStr(rowItemLOG._date) & "#"
                    If tblBDL.DefaultView.Count > 0 Then
                        Return False
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dtaBDL.Dispose()
            dtaItemLOG.Dispose()
            tblBDL = Nothing
            tblItemLOG = Nothing
        End Try
    End Function

End Class
