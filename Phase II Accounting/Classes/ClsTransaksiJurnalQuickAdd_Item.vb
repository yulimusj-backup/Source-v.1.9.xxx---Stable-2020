Option Explicit On
Option Strict Off

Public Class ClsTransaksiJurnalQuickAdd_Item

    Public Function IsiDGVItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
   ByVal Frm As FrmTransaksiJournal_quickAdd) As Boolean
        Dim DTA As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_add_itemTableAdapter
        Dim TBL As New DseTransaksiJurnalQuickAdd.tbltrans_jurnal_quick_add_itemDataTable
        Dim row As DseTransaksiJurnalQuickAdd.tbltrans_jurnal_quick_add_itemRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim ObjAkun As New ClsMasterAkun
        Dim DefaultPrecision As Integer = 0
        Dim SaldoNormal As String = ""

        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()

            'isikan tabel
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColPrimaryAkun.Index).Value = .Primary_Akun
                        rowDGV.Cells(Frm.ColAccount.Index).Value = ObjAkun.GetAccountCode(.Primary_Akun)
                        Dim a As String = ObjAkun.GetAccountCode(.Primary_Akun)
                        rowDGV.Cells(Frm.ColAccountDesc.Index).Value = ObjAkun.GetDescription(a)
                        If .Saldo_Normal = 0 Then
                            rowDGV.Cells(Frm.ColDebetKredit.Index).Value = "Debet"
                        Else
                            rowDGV.Cells(Frm.ColDebetKredit.Index).Value = "Credit"
                        End If
                    End With
                Next
            End If
            'set nomer
            ''ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjAkun = Nothing
        End Try
    End Function

    Public Function SimpanItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTransaksiJournal_quickAdd) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_add_itemTableAdapter
        Dim DefaultPrecision As Integer = 0
        Dim SaldoNormal As Integer = 0

        Try
            'simpan data Produk Code No
            For Each row In vDGVItem.Rows
                If row.Cells(Frm.ColAccount.Index).Value <> "" Then
                    If row.Cells(Frm.ColDebetKredit.Index).Value = "Debet" Then
                        SaldoNormal = 0
                    Else
                        SaldoNormal = 1
                    End If
                    'simpan ke basis data item
                    DTA.Insert(vPrimaryTransaksi, CDbl(row.Cells(Frm.ColPrimaryAkun.Index).Value), SaldoNormal)
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_add_itemTableAdapter
        Try
            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
