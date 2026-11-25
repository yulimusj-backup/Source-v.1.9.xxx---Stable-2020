Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports System.Drawing
Imports system.io
Public Class ClsTransaksiQuick_Sorting
    Public Enum EnumTipeTransaksi As Integer
        QuickAttendance
        QuickSalary
    End Enum

    Public Enum EnumTipeSorting As Integer
        Ascending
        Descending
        UnSorted
    End Enum

    Public Enum EnumTipeDesc As Integer
        Employee_Code
        Employee_NIK
        Department_Code
        Job_Position_Code
        Employee_Name
        Transaction_No
        Transaction_Date
        Period_From
        Period_To
    End Enum

    Public Function SimpanSetting(ByVal vPrimaryUser As Double, ByVal vTipeTransaksi As EnumTipeTransaksi, ByVal i As FrmTransaksiQuick_Sorting) As Boolean
        Dim Dta As New DseTransaksiQuick_SortingTableAdapters.tbltrans_setting_user_sortingTableAdapter
        'Dim row As DseTransaksiQuick_Sorting.tbltrans_setting_user_sortingRow
        Dim rd As DataGridViewRow
        Dim Desc As String = ""
        Dim TipeSort As EnumTipeSorting
        Dim Priority As Integer
        Dim EnumKolom As EnumTipeDesc

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            'hapus dahulu
            Dta.Delete(vPrimaryUser, vTipeTransaksi)
            'simpan
            For Each rd In i.DgvItem.Rows
                Desc = rd.Cells(i.ColDesc.Index).Value
                TipeSort = Me.GetTipeSorting(rd.Cells(i.ColSortType.Index).Value)
                If IsNumeric(rd.Cells(i.ColPriority.Index).Value) = True Then
                    Priority = CInt(rd.Cells(i.ColPriority.Index).Value)
                Else
                    Priority = -1
                End If
                EnumKolom = Me.GetTipeDesc(rd.Cells(i.ColDesc.Index).Value)

                Dta.Insert(vPrimaryUser, vTipeTransaksi, Desc, TipeSort, Priority, EnumKolom)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub LoadSetting(ByVal vPrimaryUser As Double, ByVal vTipeTransaksi As EnumTipeTransaksi, ByVal i As FrmTransaksiQuick_Sorting)
        Dim Dta As New DseTransaksiQuick_SortingTableAdapters.tbltrans_setting_user_sortingTableAdapter
        Dim Tbl As New DseTransaksiQuick_Sorting.tbltrans_setting_user_sortingDataTable
        Dim row As DseTransaksiQuick_Sorting.tbltrans_setting_user_sortingRow
        Dim rd As DataGridViewRow

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill(Tbl, vPrimaryUser, vTipeTransaksi)
            If Tbl.Count > 0 Then
                i.DgvItem.Rows.Clear()
                For Each row In Tbl
                    i.DgvItem.Rows.Add()
                    rd = i.DgvItem.Rows(i.DgvItem.RowCount - 1)
                    With row
                        rd.Cells(i.ColPrimaryItem.Index).Value = CDbl(.Primary)
                        rd.Cells(i.ColDesc.Index).Value = .Description
                        rd.Cells(i.ColSortType.Index).Value = Me.GetTipeSortingStr(.Tipe_Sorting)
                        If .Priority <> -1 Then
                            rd.Cells(i.ColPriority.Index).Value = CInt(Math.Round(.Priority, 0))
                        Else
                            rd.Cells(i.ColPriority.Index).Value = "-"
                        End If
                    End With
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetTipeSorting(ByVal vTipeSortingStr As String) As EnumTipeSorting
        Try
            If vTipeSortingStr = "Ascending" Then
                Return EnumTipeSorting.Ascending
            ElseIf vTipeSortingStr = "Descending" Then
                Return EnumTipeSorting.Descending
            ElseIf vTipeSortingStr = "Unsorted" Then
                Return EnumTipeSorting.UnSorted
            End If

            Return EnumTipeSorting.Ascending
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumTipeSorting.Ascending
        End Try
    End Function

    Public Function GetTipeSortingStr(ByVal vTipeSorting As EnumTipeSorting) As String
        Try
            If vTipeSorting = EnumTipeSorting.Ascending Then
                Return "Ascending"
            ElseIf vTipeSorting = EnumTipeSorting.Descending Then
                Return "Descending"
            ElseIf vTipeSorting = EnumTipeSorting.UnSorted Then
                Return "Unsorted"
            End If

            Return "Ascending"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Ascending"
        End Try
    End Function

    Public Function GetTipeDesc(ByVal vTipeDescStr As String) As EnumTipeDesc
        Try
            If vTipeDescStr = "Employee Code" Then
                Return EnumTipeDesc.Employee_Code
            ElseIf vTipeDescStr = "Employee NIK" Then
                Return EnumTipeDesc.Employee_NIK
            ElseIf vTipeDescStr = "Department Code" Then
                Return EnumTipeDesc.Department_Code
            ElseIf vTipeDescStr = "Job Position Code" Then
                Return EnumTipeDesc.Job_Position_Code
            ElseIf vTipeDescStr = "Employee Name" Then
                Return EnumTipeDesc.Employee_Name
            ElseIf vTipeDescStr = "Transaction No" Then
                Return EnumTipeDesc.Transaction_No
            ElseIf vTipeDescStr = "Transaction Date" Then
                Return EnumTipeDesc.Transaction_Date
            ElseIf vTipeDescStr = "Period From" Then
                Return EnumTipeDesc.Period_From
            ElseIf vTipeDescStr = "Period To" Then
                Return EnumTipeDesc.Period_To
            End If

            Return EnumTipeDesc.Employee_Code
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumTipeDesc.Employee_Code
        End Try
    End Function

    Public Function GetTipeDescStr(ByVal vTipeDesc As EnumTipeDesc) As String
        Try
            If vTipeDesc = EnumTipeDesc.Employee_Code Then
                Return "Employee Code"
            ElseIf vTipeDesc = EnumTipeDesc.Employee_NIK Then
                Return "Employee NIK"
            ElseIf vTipeDesc = EnumTipeDesc.Department_Code Then
                Return "Department Code"
            ElseIf vTipeDesc = EnumTipeDesc.Job_Position_Code Then
                Return "Job Position Code"
            ElseIf vTipeDesc = EnumTipeDesc.Employee_Name Then
                Return "Employee Name"
            ElseIf vTipeDesc = EnumTipeDesc.Transaction_No Then
                Return "Transaction No"
            ElseIf vTipeDesc = EnumTipeDesc.Transaction_Date Then
                Return "Transaction Date"
            ElseIf vTipeDesc = EnumTipeDesc.Period_To Then
                Return "Period To"
            ElseIf vTipeDesc = EnumTipeDesc.Period_From Then
                Return "Period From"
            End If

            Return "Employee Code"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Employee Code"
        End Try
    End Function

    Public Function GetSetting(ByVal vPrimaryUser As Double, ByVal vTipeTransaksi As EnumTipeTransaksi) As String
        Dim Dta As New DseTransaksiQuick_SortingTableAdapters.tbltrans_setting_user_sortingTableAdapter
        Dim Tbl As New DseTransaksiQuick_Sorting.tbltrans_setting_user_sortingDataTable
        Dim row As DseTransaksiQuick_Sorting.tbltrans_setting_user_sortingRow
        Dim Retval As String = ""
        Dim TipeSortingStr As String = ""

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill(Tbl, vPrimaryUser, vTipeTransaksi)

            Tbl.DefaultView.RowFilter = ""
            Tbl.DefaultView.Sort = "[Priority]"
            For Each dv As DataRowView In Tbl.DefaultView
                row = dv.Row
                With row
                    If .Tipe_Sorting = EnumTipeSorting.Ascending Then
                        TipeSortingStr = "ASC"
                    ElseIf .Tipe_Sorting = EnumTipeSorting.Descending Then
                        TipeSortingStr = "DESC"
                    Else
                        TipeSortingStr = ""
                    End If

                    If TipeSortingStr <> "" Then
                        If Retval = "" Then
                            Retval = Me.GetSettingDesc(.Enum_Column, .Tipe_Transaksi, TipeSortingStr)
                        Else
                            Retval &= ", " & Me.GetSettingDesc(.Enum_Column, .Tipe_Transaksi, TipeSortingStr)
                        End If
                    End If
                End With
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetSettingDesc(ByVal vTipeDesc As EnumTipeDesc, ByVal vTipeTransaksi As EnumTipeTransaksi, ByVal vTipeSorting As String) As String
        Try
            If vTipeTransaksi = EnumTipeTransaksi.QuickAttendance Then
                If vTipeDesc = EnumTipeDesc.Employee_Code Then
                    Return "[Emp Code] " & vTipeSorting
                ElseIf vTipeDesc = EnumTipeDesc.Employee_NIK Then
                    Return "[Emp NIK] " & vTipeSorting '
                ElseIf vTipeDesc = EnumTipeDesc.Department_Code Then
                    Return "[Dept Code] " & vTipeSorting '
                ElseIf vTipeDesc = EnumTipeDesc.Job_Position_Code Then
                    Return "[Job Code] " & vTipeSorting '
                ElseIf vTipeDesc = EnumTipeDesc.Employee_Name Then
                    Return "[Emp Name]"
                ElseIf vTipeDesc = EnumTipeDesc.Transaction_No Then
                    Return "[no transaksi] " & vTipeSorting
                ElseIf vTipeDesc = EnumTipeDesc.Transaction_Date Then
                    Return "[tanggal] " & vTipeSorting
                ElseIf vTipeDesc = EnumTipeDesc.Period_From Then
                    Return "[start date] " & vTipeSorting
                ElseIf vTipeDesc = EnumTipeDesc.Period_To Then
                    Return "[end date] " & vTipeSorting
                End If
            Else
                If vTipeDesc = EnumTipeDesc.Employee_Code Then
                    Return "[Code] " & vTipeSorting
                ElseIf vTipeDesc = EnumTipeDesc.Employee_NIK Then
                    Return "[NIK] " & vTipeSorting
                ElseIf vTipeDesc = EnumTipeDesc.Department_Code Then
                    Return "[Dept Code] " & vTipeSorting '
                ElseIf vTipeDesc = EnumTipeDesc.Job_Position_Code Then
                    Return "[Position Code] " & vTipeSorting '
                ElseIf vTipeDesc = EnumTipeDesc.Employee_Name Then
                    Return "[Full Name] " & vTipeSorting
                ElseIf vTipeDesc = EnumTipeDesc.Transaction_No Then
                    Return "[No Transaksi] " & vTipeSorting
                ElseIf vTipeDesc = EnumTipeDesc.Transaction_Date Then
                    Return "[Tanggal Transaksi] " & vTipeSorting
                ElseIf vTipeDesc = EnumTipeDesc.Period_From Then
                    Return "[Start Date] " & vTipeSorting
                ElseIf vTipeDesc = EnumTipeDesc.Period_To Then
                    Return "[End Date] " & vTipeSorting
                End If
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    'Public Sub UpdateTbl(ByVal vPrimaryItem As Double, ByVal vTipeSorting As EnumTipeSorting, ByVal vPriority As Integer, _
    'ByRef Tbl As DseTransaksiQuick_Sorting.tbltrans_setting_user_sortingDataTable)
    '    Dim row As DseTransaksiQuick_Sorting.tbltrans_setting_user_sortingRow

    '    Try
    '        Tbl.DefaultView.RowFilter = ""
    '        If Tbl.DefaultView.Count > 0 Then
    '            row = Tbl.DefaultView(0).Row
    '            row.Tipe_Sorting = vTipeSorting
    '            row.Priority = vPriority
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Public Sub UpdateDgv(ByVal Tbl As DseTransaksiQuick_Sorting.tbltrans_setting_user_sortingDataTable, _
    'ByVal i As FrmTransaksiQuick_Sorting)
    '    Dim row As DseTransaksiQuick_Sorting.tbltrans_setting_user_sortingRow

    '    Try

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub
End Class
