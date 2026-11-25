Option Explicit On
Option Strict Off
Public Class ClsMasterCommission_Type_Item

    Public Enum Enum_BaseCalculation
        Fixed_Amount_For_Each_Sold_Item
        Percentage_of_Sales_Amount
    End Enum

    Public Enum Enum_LimitType
        By_Qty
        By_Amount
    End Enum

    Public Enum Enum_ProgresiveCalculation
        NoProgresive
        Progresive
    End Enum

    Public Enum Enum_ConditionMin
        LebihBesar
        LebihBesar_SamaDengan
    End Enum

    Public Enum Enum_ConditionMax
        LebihKecil
        LebihKecil_SamaDengan
    End Enum

    Public Function SimpanItem(ByVal vPrimaryTransaksi As Double, _
                                  ByVal vDGVItem As DataGridView, _
                                  ByVal Frm As FrmMasterCommission) As Boolean
        Dim DTAItem As New DseMasterCommissionTableAdapters.tblmaster_sales_commission_itemTableAdapter
        Dim row As DataGridViewRow
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGVItem.AllowUserToAddRows = False
            For Each row In vDGVItem.Rows
                With row
                    DTAItem.Insert(vPrimaryTransaksi, Me.GetInt_ByConditionMin(.Cells(Frm.Col_ConditionMin.Index).Value), .Cells(Frm.ColMin.Index).Value, _
                    Me.GetInt_ByConditionMax(.Cells(Frm.Col_ConditionMax.Index).Value), .Cells(Frm.ColMax.Index).Value, .Cells(Frm.ColCommision.Index).Value)
                End With
            Next
            vDGVItem.AllowUserToAddRows = True
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTAItem As New DseMasterCommissionTableAdapters.tblmaster_sales_commission_itemTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Delete(vPrimaryTransaksi)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function IsiDGVItem(ByVal vPrimaryTransaksi As Double, ByVal vDGVItem As DataGridView, _
                               ByVal vFrm As FrmMasterCommission, ByVal vObjCD As Core_Module.ClsCekData, _
                               ByVal vDefCurr As Decimal, ByVal vLimitType As ClsMasterCommission_Type_Item.Enum_LimitType, _
    ByVal vBaseCalculation As ClsMasterCommission_Type_Item.Enum_BaseCalculation) As Boolean
        Dim DTAItem As New DseMasterCommissionTableAdapters.tblmaster_sales_commission_itemTableAdapter
        Dim TBLItem As New DseMasterCommission.tblmaster_sales_commission_itemDataTable
        Dim rowItem As DseMasterCommission.tblmaster_sales_commission_itemRow
        Dim rowDGV As DataGridViewRow
        Dim CboKondisi As New Windows.Forms.DataGridViewComboBoxCell
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGVItem.Rows.Clear()
            vDGVItem.AllowUserToAddRows = False

            If DTAItem.Fill(TBLItem, vPrimaryTransaksi) > 0 Then
                For Each rowItem In TBLItem
                    vDGVItem.Rows.Add()
                    rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.ColNo.Index).Value = ""
                        .Cells(vFrm.Col_ConditionMin.Index).Value = Me.GetStr_ByConditionMin(rowItem.Condition_Min)
                        .Cells(vFrm.Col_ConditionMax.Index).Value = Me.GetStr_ByConditionMax(rowItem.Condition_Max)

                        If vLimitType = Enum_LimitType.By_Amount Then
                            .Cells(vFrm.ColMin.Index).Value = vObjCD.DigitGrouping(rowItem.Min, vDefCurr)
                            .Cells(vFrm.ColMax.Index).Value = vObjCD.DigitGrouping(rowItem.Max, vDefCurr)
                        ElseIf vLimitType = Enum_LimitType.By_Qty Then
                            .Cells(vFrm.ColMin.Index).Value = vObjCD.DigitGrouping(rowItem.Min, 2)
                            .Cells(vFrm.ColMax.Index).Value = vObjCD.DigitGrouping(rowItem.Max, 2)
                        End If
                        If vBaseCalculation = Enum_BaseCalculation.Fixed_Amount_For_Each_Sold_Item Then
                            .Cells(vFrm.ColCommision.Index).Value = vObjCD.DigitGrouping(rowItem.Comm_Value, vDefCurr)
                        ElseIf vBaseCalculation = Enum_BaseCalculation.Percentage_of_Sales_Amount Then
                            .Cells(vFrm.ColCommision.Index).Value = vObjCD.DigitGrouping(rowItem.Comm_Value, 2)
                        End If
                    End With
                Next
            End If
            vDGVItem.AllowUserToAddRows = True
            vObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem = Nothing
        End Try
    End Function

    Public Function GetInt_ByConditionMax(ByVal vStr As String) As Integer
        Try
            If vStr = "<" Then
                Return ClsMasterCommission_Type_Item.Enum_ConditionMax.LebihKecil
            ElseIf vStr = "<=" Then
                Return ClsMasterCommission_Type_Item.Enum_ConditionMax.LebihKecil_SamaDengan
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ClsMasterCommission_Type_Item.Enum_ConditionMax.LebihKecil
        End Try
    End Function

    Public Function GetInt_ByConditionMin(ByVal vStr As String) As Integer
        Try
            If vStr = ">" Then
                Return ClsMasterCommission_Type_Item.Enum_ConditionMin.LebihBesar
            ElseIf vStr = ">=" Then
                Return ClsMasterCommission_Type_Item.Enum_ConditionMin.LebihBesar_SamaDengan
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ClsMasterCommission_Type_Item.Enum_ConditionMax.LebihKecil
        End Try
    End Function

    Public Function GetStr_ByConditionMax(ByVal vInt As ClsMasterCommission_Type_Item.Enum_ConditionMax) As String
        Dim retval As String = ""
        Try
            If vInt = Enum_ConditionMax.LebihKecil Then
                retval = "<"
            ElseIf vInt = Enum_ConditionMax.LebihKecil_SamaDengan Then
                retval = "<="
            End If
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetStr_ByConditionMin(ByVal vInt As ClsMasterCommission_Type_Item.Enum_ConditionMin) As String
        Dim retval As String = ""
        Try
            If vInt = Enum_ConditionMin.LebihBesar Then
                retval = ">"
            ElseIf vInt = Enum_ConditionMin.LebihBesar_SamaDengan Then
                retval = ">="
            End If
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
End Class
