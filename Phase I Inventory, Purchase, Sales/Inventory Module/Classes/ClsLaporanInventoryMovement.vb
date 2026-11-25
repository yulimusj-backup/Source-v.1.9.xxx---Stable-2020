Option Explicit On
Option Strict Off

Public Class ClsLaporanInventoryMovement
    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim objProduct As New Core_Module.clsMasterProduct

    'membernya
    Private mPrimary As Double = -1
    Private mUser As String = ""
    Private mGrouping As Integer
    Private mGroup1 As Integer
    Private mGroup2 As Integer
    Private mGroupLevel As Integer

    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property User() As String
        Get
            Return Me.mUser
        End Get
        Set(ByVal value As String)
            mUser = value
        End Set
    End Property

    Public Property Grouping() As Integer
        Get
            Return Me.mGrouping
        End Get
        Set(ByVal value As Integer)
            mGrouping = value
        End Set
    End Property

    Public Property Group1() As Integer
        Get
            Return Me.mGroup1
        End Get
        Set(ByVal value As Integer)
            mGroup1 = value
        End Set
    End Property

    Public Property Group2() As Integer
        Get
            Return Me.mGroup2
        End Get
        Set(ByVal value As Integer)
            mGroup2 = value
        End Set
    End Property

    Public Property GroupLevel() As Integer
        Get
            Return Me.mGroupLevel
        End Get
        Set(ByVal value As Integer)
            mGroupLevel = value
        End Set
    End Property

    Public Function Simpan(ByVal vDGV As DataGridView, ByVal vfrm As FrmLaporanInventoryMovement, _
    ByVal vUserCode As String) As Boolean
        Dim DTA As New dseLaporanInventoryMovementTableAdapters.tbltrans_setting_report_inventorymovementTableAdapter
        Dim TBL As New dseLaporanInventoryMovement.tbltrans_setting_report_inventorymovementDataTable
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_Primary(TBL, Me.mPrimary) > 0 Then
                DTA.Delete(Me.mPrimary)
            End If
            If DTA.Insert(Me.mUser, Me.mGrouping, Me.mGroup1, Me.mGroup2, Me.mGroupLevel) > 0 Then
                Me.mPrimary = Me.GetPrimary(Me.mUser)
                Me.SimpanItemDGVProductGroup(mPrimary, vDGV, vfrm)
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanItemDGVProductGroup(ByVal vPrimaryTransaksi As Double, ByVal vDGVItem As DataGridView, _
    ByVal vFrm As FrmLaporanInventoryMovement) As Boolean
        Dim DTAItem As New dseLaporanInventoryMovementTableAdapters.tbltrans_setting_report_inventorymovement_itemTableAdapter
        Dim TBLItem As New dseLaporanInventoryMovement.tbltrans_setting_report_inventorymovement_itemDataTable
        Dim row As DataGridViewRow
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimaryTransaksi) > 0 Then
                DTAItem.Delete(vPrimaryTransaksi)
            End If
            If vDGVItem.RowCount > 1 Then
                For i As Integer = 0 To vDGVItem.RowCount - 2
                    row = vDGVItem.Rows(i)
                    With row
                        DTAItem.Insert(vPrimaryTransaksi, .Cells(vFrm.ColGroup_Primary.Index).Value)
                    End With
                Next
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New dseLaporanInventoryMovementTableAdapters.tbltrans_setting_report_inventorymovementTableAdapter
        Dim Retval As Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Retval = DTA.ScalarQuery_GetPrimary(vCode)
            If Retval.HasValue = True Then
                Return Retval.Value
            Else
                Return -1
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vUserCode As String, _
    ByVal vFrm As FrmLaporanInventoryMovement) As Boolean

        Dim DTA As New dseLaporanInventoryMovementTableAdapters.tbltrans_setting_report_inventorymovementTableAdapter
        Dim TBL As New dseLaporanInventoryMovement.tbltrans_setting_report_inventorymovementDataTable
        Dim row As dseLaporanInventoryMovement.tbltrans_setting_report_inventorymovementRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.mPrimary = GetPrimary(vUserCode)
            DTA.FillBy_Primary(TBL, CDec(Me.mPrimary))
            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), dseLaporanInventoryMovement.tbltrans_setting_report_inventorymovementRow)
                With row
                    Me.mPrimary = .Primary
                    Me.mUser = .User_Code
                    Me.mGrouping = .Grouping
                    Me.mGroup1 = .Group_1
                    Me.mGroup2 = .Group_2
                    Me.mGroupLevel = .Group_Level
                End With

                'Me.IsiDGVProductGroup(vFrm, Primary, vDGV)
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function IsiDGVProductGroup(ByVal vFrm As FrmLaporanInventoryMovement, ByVal vPrimaryTransaksi As Double, _
                                 ByVal vDGV As DataGridView) As Boolean
        Dim DTA As New dseLaporanInventoryMovementTableAdapters.tbltrans_setting_report_inventorymovement_itemTableAdapter
        Dim TBL As New dseLaporanInventoryMovement.tbltrans_setting_report_inventorymovement_itemDataTable
        Dim row As dseLaporanInventoryMovement.tbltrans_setting_report_inventorymovement_itemRow
        Dim rowDGV As DataGridViewRow
        Dim CodeItem As String = ""
        Dim DescItem As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.Rows.Clear()
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                vDGV.AllowUserToAddRows = False
                For Each row In TBL
                    vDGV.Rows.Add()
                    rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.ColGroup_Primary.Index).Value = row.Primary_Acuan
                        If Me.mGroupLevel = 0 Or Me.mGroupLevel = 1 Or Me.mGroupLevel = 2 Then
                            CodeItem = objGroup.GetGroupCode(row.Primary_Acuan)
                            DescItem = objGroup.GetGroupDesc(row.Primary_Acuan)
                        ElseIf Me.mGroupLevel = 3 Then
                            CodeItem = objCategory.GetCategoryCode(row.Primary_Acuan)
                            DescItem = objCategory.GetDesc(row.Primary_Acuan)
                        ElseIf Me.mGroupLevel = 4 Then
                            CodeItem = objProduct.GetProductCode(row.Primary_Acuan)
                            DescItem = objProduct.GetDescription(row.Primary_Acuan)
                        End If
                        .Cells(vFrm.ColGroup_Code.Index).Value = CodeItem
                        .Cells(vFrm.ColGroup_Desc.Index).Value = DescItem
                    End With
                Next
                vDGV.AllowUserToAddRows = True
            End If
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function
End Class
