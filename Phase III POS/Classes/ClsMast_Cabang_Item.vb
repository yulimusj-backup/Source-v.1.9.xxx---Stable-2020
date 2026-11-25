Option Explicit On
Option Strict Off
Imports Core_Module

Public Class ClsMast_Cabang_Item
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
    ByVal i As FrmMast_Cabang, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseMast_CabangTableAdapters.tblmaster_cabang_itemlocationTableAdapter
        Dim TBLItemP As New DseMast_Cabang.tblmaster_cabang_itemlocationDataTable
        Dim rowItemP As DseMast_Cabang.tblmaster_cabang_itemlocationRow
        Dim ObjLocation As New clsMasterStorageArea
        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItem.AllowUserToAddRows = False

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item
            DTAItemP.FillBy_PrimaryCabang(TBLItemP, vPrimary)
            vDGVItem.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColNo.Index).Value = ""
                    .Cells(i.ColLocationCode.Index).Value = ObjLocation.GetStorageAreaCode(rowItemP.primary_location)
                    .Cells(i.ColLocationDesc.Index).Value = ObjLocation.GetDescription(rowItemP.primary_location)
                    .Cells(i.ColRemarksItem.Index).Value = rowItemP.Remarks
                    .Cells(i.ColPrimaryLocation.Index).Value = rowItemP.primary_location
                End With
            Next

            vObjCD.SetNumber(vDGVItem)

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            vDGVItem.AllowUserToAddRows = True

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemP.Dispose()
            TBLItemP.Dispose()
            ObjLocation = Nothing
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimary As Double) As Boolean
        Dim DTAP As New DseMast_CabangTableAdapters.tblmaster_cabang_itemlocationTableAdapter
        Try
            DTAP.Connection.ConnectionString = Me.GetConstring
            DTAP.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAP.Dispose()
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, ByVal i As FrmMast_Cabang) As Boolean
        Dim DTAItemP As New DseMast_CabangTableAdapters.tblmaster_cabang_itemlocationTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Try
            DTAItemP.Connection.ConnectionString = Me.GetConstring
            vDGVItem.AllowUserToAddRows = False

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            For Each rowDGV In vDGVItem.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColRemarksItem.Index).Value) = True Then
                    rowDGV.Cells(i.ColRemarksItem.Index).Value = ""
                End If
                With rowDGV
                    DTAItemP.Insert(.Cells(i.ColPrimaryLocation.Index).Value, .Cells(i.ColRemarksItem.Index).Value, _
                    vPrimary)
                End With
            Next

            vDGVItem.AllowUserToAddRows = True
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemP.Dispose()
            ObjCD = Nothing
        End Try
    End Function
End Class
