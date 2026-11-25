Option Explicit On
Option Strict Off
Public Class ClsTrans_SPL_Authorize_Access

    Private mPrimary As Double = -1
    Private mUsercode As String = ""
    Private mAuthtype As Integer = 0

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property
    Public Property Usercode() As String
        Get
            Return mUsercode
        End Get
        Set(ByVal value As String)
            mUsercode = value
        End Set
    End Property

    Public Property Authtype() As Integer
        Get
            Return mAuthtype
        End Get
        Set(ByVal value As Integer)
            mAuthtype = value
        End Set
    End Property

    Public Function ApakahAda(ByVal vUsercode As String) As Boolean
        Dim DTA As New DseTrans_SPL_Autorize_AccessTableAdapters.tbltrans_spl_authaccessTableAdapter
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

    Public Function GetPrimary(ByVal vKodeUser As String) As Double
        Dim DTA As New DseTrans_SPL_Autorize_AccessTableAdapters.tbluserTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimarybyCode(vKodeUser).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimarybyCode(vKodeUser).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetAuthType(ByVal vkodeUser As String) As Double
        Dim DTA As New DseTrans_SPL_Autorize_AccessTableAdapters.tbltrans_spl_authaccessTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetAuthTypebyCode(vkodeUser).HasValue = True Then
                Return DTA.ScalarQuery_GetAuthTypebyCode(vkodeUser).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Simpan(ByVal vFrm As FrmTrans_SPL_Authorize_Access) As Boolean
        Dim DTA As New DseTrans_SPL_Autorize_AccessTableAdapters.tbltrans_spl_authaccessTableAdapter
        Dim objAAitem As New ClsTrans_SPL_Authorize_AccessItem
        Dim objUser As New Core_Module.ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            HapusHeader()
            objAAitem.HapusItem()
            For i As Integer = 0 To vFrm.tblItemUp.Rows.Count - 1
                vFrm.rowItemUp = vFrm.tblItemUp.Rows.Item(i)
                
                If vFrm.rowItemUp.Identifier.ToString.StartsWith("-1") Then
                    vFrm.rowItemUp.Identifier = objUser.GetPrimary(vFrm.rowItemUp.usercode) & vFrm.rowItemUp.Identifier.ToString.Substring(2)
                End If
                DTA.Insert(vFrm.rowItemUp.usercode, vFrm.rowItemUp.auth_type, vFrm.rowItemUp.Identifier)
            Next
            objAAitem.SimpanItem(vFrm)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGVHeader(ByVal vDGVHeader As DataGridView, _
                              ByVal vFrm As FrmTrans_SPL_Authorize_Access) As Boolean

        Dim ObjCD As New Core_Module.ClsCekData
        Dim objAAitem As New ClsTrans_SPL_Authorize_AccessItem
        Dim objUser As New Core_Module.ClsUser
        Dim DTA As New DseTrans_SPL_Autorize_AccessTableAdapters.tbltrans_spl_authaccessTableAdapter
        Dim TBL As New DseTrans_SPL_Autorize_Access.tbltrans_spl_authaccessDataTable
        Dim rowItem As DseTrans_SPL_Autorize_Access.tbltrans_spl_authaccessRow
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGVHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVHeader.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVHeader.Rows.Clear()
            If DTA.Fill(TBL) > 0 Then
                For Each rowItem In TBL
                    vDGVHeader.Rows.Add()
                    rowDGV = vDGVHeader.Rows(vDGVHeader.Rows.Count - 2)
                    With rowDGV
                        .Cells(vFrm.colPrimary.Index).Value = rowItem.primary
                        .Cells(vFrm.colUserCode.Index).Value = rowItem.usercode
                        .Cells(vFrm.colUsername.Index).Value = objUser.GetNama(rowItem.usercode)
                        If rowItem.auth_type = 0 Then
                            .Cells(vFrm.colAccessType.Index).Value = vFrm.colAccessType.Items.Item(0)
                        Else
                            .Cells(vFrm.colAccessType.Index).Value = vFrm.colAccessType.Items.Item(1)
                        End If
                        If rowItem.Identifier.ToString.StartsWith("-1") Then
                            vFrm.rowItemUp.Identifier = objUser.GetPrimary(rowItem.usercode) & rowItem.Identifier.ToString.Substring(2)
                        End If
                        .Cells(vFrm.ColIdentifier.Index).Value = rowItem.Identifier
                    End With

                    vFrm.TambahBarisAtas(vFrm.dgvManager.Item(vFrm.colPrimary.Index, vDGVHeader.Rows.Count - 2).Value, _
                                vFrm.dgvManager.Item(vFrm.colUserCode.Index, vDGVHeader.Rows.Count - 2).Value, _
                                vFrm.dgvManager.Item(vFrm.colAccessType.Index, vDGVHeader.Rows.Count - 2).Value, _
                                vFrm.dgvManager.Item(vFrm.ColIdentifier.Index, vDGVHeader.Rows.Count - 2).Value)
                Next

            End If
            ObjCD.SetNumber(vDGVHeader)
            vDGVHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVHeader.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            objAAitem.IsiDGVItem(vFrm.dgvUser, vFrm)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function HapusHeader() As Boolean
        Dim DTA As New DseTrans_SPL_Autorize_AccessTableAdapters.tbltrans_spl_authaccessTableAdapter
        Try
            DTA.Delete()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetListOfUser(ByVal vUserCode As String) As List(Of String)
        '//Salin ke List_ dari database user item
        Dim dta As New DseTrans_SPL_Autorize_AccessTableAdapters.tbluser1TableAdapter
        Dim tbl As New DseTrans_SPL_Autorize_Access.tbluser1DataTable
        Dim row As DseTrans_SPL_Autorize_Access.tbluser1Row
        Dim list_ As New List(Of String)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(tbl, vUserCode)
            For Each row In tbl
                If list_.Contains(row.Kode_User) = False Then
                    list_.Add(row.Kode_User)
                End If
            Next
            list_.Add(vUserCode)
            Return list_
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

End Class
