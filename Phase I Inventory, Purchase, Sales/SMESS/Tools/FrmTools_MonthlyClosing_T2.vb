Option Explicit On
Option Strict Off

Partial Public Class FrmTools_MonthlyClosing
    Private Sub T2_BtnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles T2_BtnFind.Click
        Dim DTABaru As New DseToolsTableAdapters.DataTable1TableAdapter
        Dim TBLBaru As New DseTools.DataTable1DataTable
        Dim rowBaru As DseTools.DataTable1Row
        Dim DTALama As New DseToolsTableAdapters.DataTable2TableAdapter
        Dim TBLLama As New DseTools.DataTable2DataTable
        Dim rowLama As DseTools.DataTable2Row
        Dim TBLSelisih As New DseTools.DataTable3DataTable
        Dim rowSelisih As DseTools.DataTable3Row
        Dim TipeJournal As AccModule.ClsTransaksiJournal.EnumTipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchaseJournal
        Dim ObjAOP As New Purchase_Module.ClsTransaksiAOP
        Dim rowDGV As DataGridViewRow
        Dim DefPrec As Integer = Me.ObjCurr.GetDecimal_DefaultCurrency

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'load semua data awal
            DTABaru.Connection.ConnectionString = Core_Module.GetConstring
            DTALama.Connection.ConnectionString = DTABaru.Connection.ConnectionString
            TBLBaru.Rows.Clear()
            TBLLama.Rows.Clear()
            TBLSelisih.Rows.Clear()

            'isikan dari sisi baru
            DTABaru.FillBy1(TBLBaru, Me.ObjCD.GetLasDateOfTheMonth(Me.T2_dtp.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.T2_dtp.Value))
            For Each rowBaru In TBLBaru.Rows
                TBLSelisih.DefaultView.RowFilter = "[Primary AOP] = " + rowBaru.Primary_AOP.ToString + " AND [Primary Akun] = " + rowBaru.Primary_Inventory_Account.ToString
                If TBLSelisih.DefaultView.Count = 0 Then
                    'tambah baris baru
                    rowSelisih = TBLSelisih.NewDataTable3Row
                    With rowSelisih
                        .Acc_Code = rowBaru.Account_Code
                        .Acc_Desc = rowBaru.Acc_Desc
                        .AOP_No = rowBaru.No_Transaksi
                        .Debet = 0
                        .Debet_New = rowBaru.At_Amount
                        .Primary_Akun = rowBaru.Primary_Inventory_Account
                        .Primary_AOP = rowBaru.Primary_AOP
                    End With
                    TBLSelisih.Rows.Add(rowSelisih)
                Else
                    'tambah nilai debet new
                    rowSelisih = TBLSelisih.DefaultView(0).Row
                    rowSelisih.Debet_New += rowBaru.At_Amount
                End If
            Next

            'isikan dari sisi lama
            DTALama.Fill(TBLLama, CInt(TipeJournal), Me.ObjCD.GetLasDateOfTheMonth(Me.T2_dtp.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.T2_dtp.Value))
            For Each rowLama In TBLLama.Rows
                If rowLama.Debet > 0 Then
                    TBLSelisih.DefaultView.RowFilter = "[Primary AOP] = " + rowLama.Primary_Document_Acuan.ToString + " AND [Primary Akun] = " + rowLama.Primary_Akun.ToString
                    If TBLSelisih.DefaultView.Count = 0 Then
                        'tambah baris baru
                        rowSelisih = TBLSelisih.NewDataTable3Row
                        With rowSelisih
                            .Acc_Code = rowLama.Account_Code
                            .Acc_Desc = rowLama.Acc_Desc
                            .AOP_No = ObjAOP.GetNoTransaksi(rowLama.Primary_Document_Acuan)
                            .Debet = rowLama.Debet
                            .Debet_New = 0
                            .Primary_Akun = rowLama.Primary_Akun
                            .Primary_AOP = rowLama.Primary_Document_Acuan
                        End With
                        TBLSelisih.Rows.Add(rowSelisih)
                    Else
                        'tambah nilai debet new
                        rowSelisih = TBLSelisih.DefaultView(0).Row
                        rowSelisih.Debet += rowLama.Debet
                    End If
                End If
            Next

            'buang yang tidak selisih
            For ij As Double = TBLSelisih.Rows.Count - 1 To 0 Step -1
                rowSelisih = TBLSelisih.Rows(ij)
                If rowSelisih.Debet = rowSelisih.Debet_New Then
                    TBLSelisih.Rows.Remove(rowSelisih)
                End If
            Next

            'urutkan
            TBLSelisih.DefaultView.RowFilter = ""
            TBLSelisih.DefaultView.Sort = "[AOP No], [Acc Code]"

            'tampilkan ke dalam dgv
            Me.T2_dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.T2_dgv.Rows.Clear()
            For Each DV As DataRowView In TBLSelisih.DefaultView
                rowSelisih = DV.Row
                If rowSelisih.Acc_Code <> "161-212" Then
                    Me.T2_dgv.Rows.Add()
                    rowDGV = Me.T2_dgv.Rows(Me.T2_dgv.Rows.Count - 1)
                    With rowSelisih
                        rowDGV.Cells(Me.ColT2_AccCode.Index).Value = .Acc_Code
                        rowDGV.Cells(Me.ColT2_AccDesc.Index).Value = .Acc_Desc
                        rowDGV.Cells(Me.ColT2_AOPNo.Index).Value = .AOP_No
                        rowDGV.Cells(Me.ColT2_Debet.Index).Value = Me.ObjCD.DigitGrouping(.Debet, DefPrec)
                        rowDGV.Cells(Me.ColT2_Debet_Baru.Index).Value = Me.ObjCD.DigitGrouping(.Debet_New, DefPrec)
                    End With
                End If
            Next
            Me.T2_dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTABaru.Dispose()
            TBLBaru.Dispose()
            DTALama.Dispose()
            TBLLama.Dispose()
            TBLSelisih.Dispose()
        End Try
    End Sub
End Class
