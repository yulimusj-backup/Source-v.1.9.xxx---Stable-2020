Option Explicit On
Option Strict Off

Public Class FrmTransaksiMonthlyClosing2_CekMultiUnit
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TBLCek As New DseTransaksiMonthlyClosing_Overhead.tblCek_MultiUnitDataTable
    Public UserCode As String = ""
    Public RefPeriod As Date = #1/1/1950#

    Dim rowCek As DseTransaksiMonthlyClosing_Overhead.tblCek_MultiUnitRow

    Private Sub FrmTransaksiMonthlyClosing2_CekMultiUnit_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgv1.Rows.Clear()
            Me.dgv1.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            For Each rowCek In TBLCek.Rows
                Me.dgv1.Rows.Add()
                rowDGV = Me.dgv1.Rows(Me.dgv1.Rows.Count - 1)
                With rowCek
                    rowDGV.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                    rowDGV.Cells(Me.ColCostGroup.Index).Value = .Cost_Group
                    rowDGV.Cells(Me.ColDesc.Index).Value = .Product_Desc
                    rowDGV.Cells(Me.ColRefUnit.Index).Value = .Ref_Unit
                    rowDGV.Cells(Me.ColUpdate.Index).Value = Boolean.FalseString
                End With
            Next

            If Me.UserCode <> "USER" Then
                Me.ColUpdate.Visible = False
                Me.LinkLabel1.Visible = False
                Me.llbUncekAll.Visible = False
                Me.llbCekAll.Visible = False
            Else
                Me.dgv1.ReadOnly = False
                Me.ColCodeNo.ReadOnly = False
                Me.ColCostGroup.ReadOnly = False
                Me.ColDesc.ReadOnly = False
                Me.ColRefUnit.ReadOnly = False
                Me.ColUpdate.ReadOnly = False
            End If
            Me.dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub

    Private Function ApakahAdaUpdate() As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgv1.Rows
                If rowDGV.Cells(Me.ColUpdate.Index).Value = Boolean.TrueString Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim ListPrimaryBundling As New List(Of Double)
        Dim ListCodeNo As New List(Of String)
        Dim rowDGV As DataGridViewRow
        Dim DTA As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing_Overhead.tbltransaksi_stockbundlingDataTable
        Dim row As DseTransaksiMonthlyClosing_Overhead.tbltransaksi_stockbundlingRow
        Dim StartDate, EndDate As Date
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim PrimaryProduct As Double = -1

        Try

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartDate = ObjCD.GetFirstDateofTheMonth(Me.RefPeriod)
            EndDate = ObjCD.GetLasDateOfTheMonth(Me.RefPeriod)

            If Me.ApakahAdaUpdate = True Then
                Dim ij As New Core_Module.FrmMasterProsesCari
                ij.ShowDialog()
                If ij.mPrimary <> -1 Then
                    For Each rowDGV In Me.dgv1.Rows
                        If rowDGV.Cells(Me.ColUpdate.Index).Value = Boolean.TrueString Then
                            If ListCodeNo.Contains(rowDGV.Cells(Me.ColCodeNo.Index).Value) = False Then
                                ListCodeNo.Add(rowDGV.Cells(Me.ColCodeNo.Index).Value)
                                PrimaryProduct = ObjProduct.GetPrimary(rowDGV.Cells(Me.ColCodeNo.Index).Value)
                                DTA.Fill(TBL, EndDate, StartDate, PrimaryProduct)
                                For Each row In TBL.Rows
                                    If ListPrimaryBundling.Contains(row.Primary) = False Then
                                        ListPrimaryBundling.Add(row.Primary)
                                        'update ke bundling
                                        DTA.UpdateQuery_SetProcess(ij.mPrimary, row.Primary)
                                    End If
                                Next
                            End If
                        End If
                    Next

                    MsgBox("Data updated.", MsgBoxStyle.Information, "Information")
                End If
                ij = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryBundling = Nothing
            ListCodeNo = Nothing
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
            ObjProduct = Nothing
        End Try
    End Sub

    Private Sub dgv1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgv1.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColUpdate.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class