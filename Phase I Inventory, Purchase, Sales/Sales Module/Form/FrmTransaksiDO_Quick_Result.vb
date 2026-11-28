Option Explicit On
Option Strict Off

Public Class FrmTransaksiDO_Quick_Result
    '     ORIGINAL CODING
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public UserCode As String = ""
    Public vRboSO As Boolean = True
    Public tblResult As New DseTransaksiQuick_DO2.tblResultDataTable
    Dim rowResult As DseTransaksiQuick_DO2.tblResultRow
    Dim ObjUser As New Core_Module.ClsUser
    Dim objCD As New Core_Module.ClsCekData

    Private Sub FrmTransaksiDO_Quick_Result_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.vRboSO = False Then
                Me.DgvItem.Columns(Me.colSONo.Index).HeaderText = "Booking No"
            End If
            Me.loadData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.objCD.SetNumber(Me.DgvItem)
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        End Try
    End Sub

    Private Sub loadData()
        Dim rowdgv As DataGridViewRow
        Try
            Me.DgvItem.Rows.Clear()
            For Each rowResult In Me.tblResult.Rows
                Me.DgvItem.Rows.Add()
                rowdgv = Me.DgvItem.Rows(Me.DgvItem.Rows.Count - 1)
                rowdgv.Cells(Me.ColDONo.Index).Value = rowResult.No_DO
                rowdgv.Cells(Me.colSONo.Index).Value = rowResult.No_Transaksi
                rowdgv.Cells(Me.ColPrimaryDO.Index).Value = rowResult.Primary_DO
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit

        Try
            If e.ColumnIndex <> Me.ColPrint.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub DgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellClick
        Try
            If e.ColumnIndex = Me.ColPrint.Index Then
                Me.PrintDO(Me.DgvItem.Item(Me.ColPrimaryDO.Index, e.RowIndex).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PrintDO(ByVal vPrimary As Double)
        Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi
        Try
            If vPrimary <> -1 Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiDO_Cetak") = False Then
                    FrmO.AccessName = "FrmTransaksiDO_Cetak"
                    If FrmO.ShowDialog <> Windows.Forms.DialogResult.OK Then
                        MsgBox("You don't have permission to perform this action.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If

                Dim frmCetak As New FrmCetakDO
                frmCetak.Primary = vPrimary
                frmCetak.UserCode = Me.UserCode
                frmCetak.ShowDialog()
                frmCetak = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class