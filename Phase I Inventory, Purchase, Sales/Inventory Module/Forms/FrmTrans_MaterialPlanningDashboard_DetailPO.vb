Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_DetailPO
    Public vDT_ItemPOB As DseTransaksiMaterialPlanning2.DataTable3DataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim rowPOB As DseTransaksiMaterialPlanning2.DataTable3Row
    Dim rowDGV As DataGridViewRow
    Dim DTAPO As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_detailotorisasipoTableAdapter
    Dim TBLPO As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_detailotorisasipoDataTable
    Dim rowPO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_detailotorisasipoRow
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjPO As New ClsTransaksiPO

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_DetailPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Jml As Nullable(Of Long)
        Dim Insert As Boolean = False
        Dim Include As Integer = 0

        Try
            'simpan data yang ada
            For Each rowDGV In Me.dgvItem.Rows
                Jml = Me.DTAPO.ScalarQuery_GetJml(rowDGV.Cells(Me.ColPrimaryPO.Index).Value)

                If Jml.HasValue = True Then
                    If Jml.Value > 0 Then
                        Insert = False
                    Else
                        Insert = True
                    End If
                Else
                    Insert = True
                End If
                If rowDGV.Cells(Me.ColInclude.Index).Value = Boolean.FalseString Then
                    Include = 0
                Else
                    Include = 1
                End If

                If Insert = True Then
                    'masukkan baris baru
                    Me.DTAPO.Insert(rowDGV.Cells(Me.ColPrimaryPO.Index).Value, Include)
                Else
                    'update data saja
                    Me.DTAPO.Update(Include, rowDGV.Cells(Me.ColPrimaryPO.Index).Value)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_DetailPO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim PrimaryPOMin As Double = Double.MaxValue
        Dim ListSudah As New List(Of Double)

        Try
            DTAPO.Connection.ConnectionString = Core_Module.GetConstring
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.Rows.Clear()

            For Each rowPOB In vDT_ItemPOB.Rows
                'tampilkan ke dalam dgv
                If rowPOB.Authorization_Level <> rowPOB.Authorization_Level_Needed Then
                    If ListSudah.Contains(rowPOB.Primary_PO) = False Then
                        ListSudah.Add(rowPOB.Primary_PO)

                        Me.dgvItem.Rows.Add()
                        rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)

                        With rowDGV
                            .Cells(Me.ColInclude.Index).Value = Boolean.FalseString
                            .Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                            If rowPOB.Tanggal_Transaksi = #1/1/1950# Then
                                .Cells(Me.ColPODate.Index).Value = ObjPO.GetTanggalPO(rowPOB.Primary_PO)
                                .Cells(Me.ColPONumber.Index).Value = ObjPO.GetNoTransaksiPO(rowPOB.Primary_PO)
                                .Cells(Me.ColSupplier.Index).Value = ObjPO.GetSupplierCode_Name(rowPOB.Primary_PO)
                            Else
                                .Cells(Me.ColPODate.Index).Value = rowPOB.Tanggal_Transaksi
                                .Cells(Me.ColPONumber.Index).Value = rowPOB.No_Transaksi
                                .Cells(Me.ColSupplier.Index).Value = rowPOB.Supplier
                            End If
                            .Cells(Me.ColPrimaryPO.Index).Value = rowPOB.Primary_PO
                        End With

                        If PrimaryPOMin > rowPOB.Primary_PO Then
                            PrimaryPOMin = rowPOB.Primary_PO
                        End If
                    End If
                End If
            Next

            'ambil data dari database, sesuaikan data sebelumnya
            DTAPO.FillBy_PrimaryPO(TBLPO, PrimaryPOMin)
            For Each rowDGV In Me.dgvItem.Rows
                TBLPO.DefaultView.RowFilter = "[primary po] = " + rowDGV.Cells(Me.ColPrimaryPO.Index).Value.ToString
                If TBLPO.DefaultView.Count = 0 Then
                    rowDGV.Cells(Me.ColInclude.Index).Value = Boolean.FalseString
                Else
                    rowPO = TBLPO.DefaultView(0).Row
                    If rowPO.include_in_calc = 1 Then
                        rowDGV.Cells(Me.ColInclude.Index).Value = Boolean.TrueString
                    Else
                        rowDGV.Cells(Me.ColInclude.Index).Value = Boolean.FalseString
                    End If
                End If
            Next

            'format
            Me.ColPODate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.dgvItem.Sort(Me.ColPONumber, System.ComponentModel.ListSortDirection.Ascending)
            Me.dgvItem.Sort(Me.ColPODate, System.ComponentModel.ListSortDirection.Ascending)
            ObjCD.SetNumber(Me.dgvItem)

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListSudah = Nothing
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColInclude.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbMP_CekAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbMP_CekAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            Windows.Forms.Cursor.Current = Cursors.Arrow
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDGV In Me.dgvItem.Rows
                rowDGV.Cells(Me.ColInclude.Index).Value = Boolean.TrueString
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbMP_UnCekAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbMP_UnCekAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            Windows.Forms.Cursor.Current = Cursors.Arrow
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDGV In Me.dgvItem.Rows
                rowDGV.Cells(Me.ColInclude.Index).Value = Boolean.FalseString
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class