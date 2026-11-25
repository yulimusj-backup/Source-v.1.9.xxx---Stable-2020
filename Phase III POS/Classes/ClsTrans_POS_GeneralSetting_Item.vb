Option Explicit On
Option Strict Off
Imports Core_Module

Public Class ClsTrans_POS_GeneralSetting_Item

    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307"
        End Try
    End Function

    Public Function IsiDGV(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTrans_POS_Settings, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_setting_itemTableAdapter
        Dim TBLItem As New DseTrans_POS_GeneralSetting.tbltrans_pos_general_setting_itemDataTable
        Dim rowItem As DseTrans_POS_GeneralSetting.tbltrans_pos_general_setting_itemRow
        Dim ObjUser As New ClsUser
        Try
            'update connection string
            DTAItem.Connection.ConnectionString = Me.GetConstring

            vDGVItem.AllowUserToAddRows = False
            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item
            DTAItem.FillBy_Primary(TBLItem, vPrimary)
            vDGVItem.Rows.Clear()
            For Each rowItem In TBLItem.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColNo.Index).Value = ""
                    .Cells(i.ColUserKode.Index).Value = rowItem.User_Code
                    .Cells(i.ColUserName.Index).Value = ObjUser.GetNama(rowItem.User_Code)
                    If rowItem.Item_Deletion = 1 Then
                        .Cells(i.Col_Autho_AllowDeleteItem.Index).Value = True
                    Else
                        .Cells(i.Col_Autho_AllowDeleteItem.Index).Value = False
                    End If
                    If rowItem.Whole_Transaction_Deletion = 1 Then
                        .Cells(i.Col_Autho_WholeTransDelete.Index).Value = True
                    Else
                        .Cells(i.Col_Autho_WholeTransDelete.Index).Value = False
                    End If
                    If rowItem.Transaction_Modification = 1 Then
                        .Cells(i.Col_Autho_AllowModify.Index).Value = True
                    Else
                        .Cells(i.Col_Autho_AllowModify.Index).Value = False
                    End If
                    If rowItem.Modify_Unit_Price = 1 Then
                        .Cells(i.Col_Autho_ModifyUnitPrice.Index).Value = True
                    Else
                        .Cells(i.Col_Autho_ModifyUnitPrice.Index).Value = False
                    End If
                    If rowItem.Modifiy_Discount = 1 Then
                        .Cells(i.Col_Autho_ModifyDiscount.Index).Value = True
                    Else
                        .Cells(i.Col_Autho_ModifyDiscount.Index).Value = False
                    End If
                    If rowItem.Allow_Receipt_Reprint = 1 Then
                        .Cells(i.Col_Autho_AllowReprint.Index).Value = True
                    Else
                        .Cells(i.Col_Autho_AllowReprint.Index).Value = False
                    End If
                    .Cells(i.ColRemarks.Index).Value = rowItem.Remarks
                End With
            Next

            'set number
            vObjCD.SetNumber(vDGVItem)

            vDGVItem.AllowUserToAddRows = True
            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjUser = Nothing
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_setting_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTrans_POS_Settings) As Boolean
        Dim DTAItem As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_setting_itemTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Dim AllowDeleteItem As Double = -1
        Dim WholeTransDeleteItem As Double = -1
        Dim AllowModify As Double = -1
        Dim AllowModifyUnitPrice As Double = -1
        Dim AllowModifyDisc As Double = -1
        Dim AllowReprint As Double = -1
        Try
            DTAItem.Connection.ConnectionString = Me.GetConstring
            vDGVItem.AllowUserToAddRows = False
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            For Each rowDGV In vDGVItem.Rows
                If rowDGV.Cells(i.Col_Autho_AllowDeleteItem.Index).Value = False Then
                    AllowDeleteItem = 0
                Else
                    AllowDeleteItem = 1
                End If
                If rowDGV.Cells(i.Col_Autho_WholeTransDelete.Index).Value = False Then
                    WholeTransDeleteItem = 0
                Else
                    WholeTransDeleteItem = 1
                End If
                If rowDGV.Cells(i.Col_Autho_AllowModify.Index).Value = False Then
                    AllowModify = 0
                Else
                    AllowModify = 1
                End If
                If rowDGV.Cells(i.Col_Autho_ModifyUnitPrice.Index).Value = False Then
                    AllowModifyUnitPrice = 0
                Else
                    AllowModifyUnitPrice = 1
                End If
                If rowDGV.Cells(i.Col_Autho_ModifyDiscount.Index).Value = False Then
                    AllowModifyDisc = 0
                Else
                    AllowModifyDisc = 1
                End If
                If rowDGV.Cells(i.Col_Autho_AllowReprint.Index).Value = False Then
                    AllowReprint = 0
                Else
                    AllowReprint = 1
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColRemarks.Index).Value) = True Then
                    rowDGV.Cells(i.ColRemarks.Index).Value = ""
                End If
                With rowDGV
                    DTAItem.Insert(vPrimary, .Cells(i.ColUserKode.Index).Value, _
                    AllowDeleteItem, WholeTransDeleteItem, AllowModify, AllowModifyUnitPrice, _
                    AllowModifyDisc, AllowReprint, .Cells(i.ColRemarks.Index).Value)
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
End Class
