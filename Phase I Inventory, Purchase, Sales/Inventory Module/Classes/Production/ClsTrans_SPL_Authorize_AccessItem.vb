Option Explicit On
Option Strict Off
Public Class ClsTrans_SPL_Authorize_AccessItem
    Public Function SimpanItem(ByVal vFrm As FrmTrans_SPL_Authorize_Access) As Boolean
        Dim DTA As New DseTrans_SPL_Autorize_AccessTableAdapters.tbltrans_spl_authaccess_itemTableAdapter
        Dim objUser As New Core_Module.ClsUser
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For i As Integer = 0 To vFrm.tblItemDown.Rows.Count - 1
                vFrm.rowItemDown = vFrm.tblItemDown.Rows.Item(i)
                If vFrm.rowItemDown.Identifier_Header.ToString.StartsWith("-1") Then
                    vFrm.rowItemDown.Identifier_Header = objUser.GetPrimary(vFrm.rowItemUp.usercode) & vFrm.rowItemUp.Identifier.ToString.Substring(2)
                End If
                DTA.Insert(vFrm.rowItemDown.primary_user, vFrm.rowItemDown.Identifier_Header)
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGVItem(ByVal vDGVItem As DataGridView, _
                              ByVal vFrm As FrmTrans_SPL_Authorize_Access) As Boolean

        Dim ObjCD As New Core_Module.ClsCekData
        Dim objUser As New Core_Module.ClsUser
        Dim DTA As New DseTrans_SPL_Autorize_AccessTableAdapters.tbltrans_spl_authaccess_itemTableAdapter
        Dim TBL As New DseTrans_SPL_Autorize_Access.tbltrans_spl_authaccess_itemDataTable
        Dim rowItem As DseTrans_SPL_Autorize_Access.tbltrans_spl_authaccess_itemRow


        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItem.Rows.Clear()
            If DTA.Fill(TBL) > 0 Then
                For Each rowItem In TBL
                    vFrm.TambahBarisBawah(rowItem.primary, rowItem.primary_user, rowItem.Identifier_Header)
                Next
                vFrm.reloadUser(vFrm.dgvManager.Item(vFrm.ColIdentifier.Index, 0).Value)
            End If
            ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function HapusItem() As Boolean
        Dim DTA As New DseTrans_SPL_Autorize_AccessTableAdapters.tbltrans_spl_authaccess_itemTableAdapter
        Try
            DTA.Delete()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vUsercode As String) As Boolean
        Dim DTA As New DseTrans_SPL_Autorize_AccessTableAdapters.tbltrans_spl_authaccess_itemTableAdapter
        Dim dummy As New Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            dummy = DTA.ScalarQuery_GetJmlByUsercode(vUsercode)
            If dummy.HasValue Then
                If dummy.Value > 0 Then
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
