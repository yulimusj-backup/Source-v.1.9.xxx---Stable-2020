Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class ClsTrans_SPL_Budget_Item
    Dim ObjCd As New Core_Module.ClsCekData

    Public Function IsiDGVItem(ByVal vDGVItem As Windows.Forms.DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTrans_SPL_Budget) As Boolean
        Dim DTA As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_itemTableAdapter
        Dim TBL As New DseTrans_SPL_Budget.tbltrans_spl_budget_itemDataTable
        Dim row As DseTrans_SPL_Budget.tbltrans_spl_budget_itemRow
        Dim rowDGV As DataGridViewRow
        Dim Total As Double = 0

        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()

            DTA.Connection.ConnectionString = Core_Module.GetConstring

            ''isikan(tabel)
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColNo.Index).Value = vDGVItem.Rows.Count
                        rowDGV.Cells(Frm.ColDeptPrimary.Index).Value = .Primary_Department
                        rowDGV.Cells(Frm.ColDeptCode.Index).Value = .Dept_Code
                        rowDGV.Cells(Frm.ColDeptDesc.Index).Value = .Dept_Desc
                        rowDGV.Cells(Frm.ColJan.Index).Value = Me.ObjCd.Convert_ToString(.Jan)
                        rowDGV.Cells(Frm.ColFeb.Index).Value = Me.ObjCd.Convert_ToString(.Feb)
                        rowDGV.Cells(Frm.ColMar.Index).Value = Me.ObjCd.Convert_ToString(.Mar)
                        rowDGV.Cells(Frm.ColApr.Index).Value = Me.ObjCd.Convert_ToString(.Apr)
                        rowDGV.Cells(Frm.ColMay.Index).Value = Me.ObjCd.Convert_ToString(.May)
                        rowDGV.Cells(Frm.ColJun.Index).Value = Me.ObjCd.Convert_ToString(.Jun)
                        rowDGV.Cells(Frm.ColJul.Index).Value = Me.ObjCd.Convert_ToString(.Jul)
                        rowDGV.Cells(Frm.ColAug.Index).Value = Me.ObjCd.Convert_ToString(.Aug)
                        rowDGV.Cells(Frm.ColSept.Index).Value = Me.ObjCd.Convert_ToString(.Sep)
                        rowDGV.Cells(Frm.ColOkt.Index).Value = Me.ObjCd.Convert_ToString(.Oct)
                        rowDGV.Cells(Frm.ColNov.Index).Value = Me.ObjCd.Convert_ToString(.Nov)
                        rowDGV.Cells(Frm.ColDec.Index).Value = Me.ObjCd.Convert_ToString(.Dec)
                        Total = .Jan + .Feb + .Mar + .Apr + .May + .Jun + .Jul + .Aug + .Sep + .Oct + .Nov + .Dec
                        rowDGV.Cells(Frm.ColTotal.Index).Value = Me.ObjCd.Convert_ToString(Total)
                        rowDGV.Cells(Frm.ColRemarks.Index).Value = .Remarks
                    End With
                Next
            End If

            'set nomer
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Function

    Public Overridable Function SimpanItem(ByVal vPrimaryTransaksi As Double, ByVal vDGV As DataGridView, _
    ByVal Frm As FrmTrans_SPL_Budget) As Boolean
        Dim DTA As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_itemTableAdapter
        Dim row As DataGridViewRow
        Dim Mnt_Jan, Mnt_Feb, Mnt_Mar, Mnt_Apr, Mnt_May, Mnt_Jun, Mnt_Jul, Mnt_Aug, Mnt_Sep, Mnt_Oct, Mnt_Nov, Mnt_Dec As Double

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            For Each row In vDGV.Rows
                If row.Cells(Frm.ColDeptPrimary.Index).Value <> -1 And String.IsNullOrEmpty(row.Cells(Frm.ColDeptPrimary.Index).Value) = False Then
                    Mnt_Jan = Me.ObjCd.Convert_ToMinutes(row.Cells(Frm.ColJan.Index).Value)
                    Mnt_Feb = Me.ObjCd.Convert_ToMinutes(row.Cells(Frm.ColFeb.Index).Value)
                    Mnt_Mar = Me.ObjCd.Convert_ToMinutes(row.Cells(Frm.ColMar.Index).Value)
                    Mnt_Apr = Me.ObjCd.Convert_ToMinutes(row.Cells(Frm.ColApr.Index).Value)
                    Mnt_May = Me.ObjCd.Convert_ToMinutes(row.Cells(Frm.ColMay.Index).Value)
                    Mnt_Jun = Me.ObjCd.Convert_ToMinutes(row.Cells(Frm.ColJun.Index).Value)
                    Mnt_Jul = Me.ObjCd.Convert_ToMinutes(row.Cells(Frm.ColJul.Index).Value)
                    Mnt_Aug = Me.ObjCd.Convert_ToMinutes(row.Cells(Frm.ColAug.Index).Value)
                    Mnt_Sep = Me.ObjCd.Convert_ToMinutes(row.Cells(Frm.ColSept.Index).Value)
                    Mnt_Oct = Me.ObjCd.Convert_ToMinutes(row.Cells(Frm.ColOkt.Index).Value)
                    Mnt_Nov = Me.ObjCd.Convert_ToMinutes(row.Cells(Frm.ColNov.Index).Value)
                    Mnt_Dec = Me.ObjCd.Convert_ToMinutes(row.Cells(Frm.ColDec.Index).Value)
                    Mnt_Dec = Me.ObjCd.Convert_ToMinutes(row.Cells(Frm.ColDec.Index).Value)
                    If String.IsNullOrEmpty(row.Cells(Frm.ColRemarks.Index).Value) = True Then
                        row.Cells(Frm.ColRemarks.Index).Value = ""
                    End If

                    DTA.Insert(vPrimaryTransaksi, row.Cells(Frm.ColDeptPrimary.Index).Value, _
                    Mnt_Jan, Mnt_Feb, Mnt_Mar, Mnt_Apr, Mnt_May, Mnt_Jun, Mnt_Jul, Mnt_Aug, _
                    Mnt_Sep, Mnt_Oct, Mnt_Nov, Mnt_Dec, row.Cells(Frm.ColRemarks.Index).Value)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_itemTableAdapter

        Try
            'hapus data item
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.Delete(vPrimaryTransaksi)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function
End Class
