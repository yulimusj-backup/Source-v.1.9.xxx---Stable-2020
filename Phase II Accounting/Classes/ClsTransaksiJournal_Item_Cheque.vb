Option Explicit On
Option Strict Off

Public Class ClsTransaksiJournal_Item_Cheque
    Private mPaid As EnumPaid = EnumPaid.Bayar

    Public Enum EnumPaid As Integer
        Bayar
        Belum
    End Enum

    Public Property Paid() As EnumPaid
        Get
            Return Me.mPaid
        End Get
        Set(ByVal value As EnumPaid)
            mPaid = value
        End Set
    End Property

    Public Function GetIdentifierUrutan(ByVal vDGVItem As DataGridView, _
       ByVal vColIdentifier As Integer, ByVal baru As Boolean, _
       ByVal vPrimaryTransaksi As Double, ByVal vIndexCell As Integer) As String
        Dim Row As DataGridViewRow
        Try
            vDGVItem.AllowUserToAddRows = False
            Dim Identifier As String = ""
            Dim UrutanMax As Double = 0
            If baru = True Then
                If vDGVItem.Rows.Count > 0 Then
                    For i As Integer = 0 To vDGVItem.Rows.Count - 1
                        Row = vDGVItem.Rows(i)
                        If Row.Index <> vIndexCell Then
                            With Row
                                If i = 0 Then
                                    UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-", 1) + 1))
                                Else
                                    If CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-", 1) + 1)) > UrutanMax Then
                                        UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-", 1) + 1))
                                    End If
                                End If
                            End With
                        End If
                    Next
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + CStr(UrutanMax + 1))
                Else
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + "1")
                End If
            Else
                If vDGVItem.Rows.Count > 0 Then
                    For i As Integer = 0 To vDGVItem.Rows.Count - 1
                        Row = vDGVItem.Rows(i)
                        If Row.Index <> vIndexCell Then
                            With Row
                                If i = 0 Then
                                    UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-") + 1))
                                Else
                                    If CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-") + 1)) > UrutanMax Then
                                        UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-") + 1))
                                    End If
                                End If
                            End With
                        End If
                    Next
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + CStr(UrutanMax + 1))
                Else
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + "1")
                End If
            End If
            vDGVItem.AllowUserToAddRows = True

            Return Identifier
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Overridable Function SimpanCheque(ByVal vPrimaryJurnal As Double, ByVal vDGVCheque As DataGridView, _
    ByVal Frm As FrmTransaksiJournal_CashBank_Cheque) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_chequeTableAdapter
        Dim Paid As Integer = 0
        Dim type As Integer = 0
        Dim row As DataGridViewRow

        Try
            vDGVCheque.AllowUserToAddRows = False

            Dim IdentifierLama As String = ""
            Dim Identifier As String = ""
            Dim counter As Integer = 0
            'set identifier
            For Each row In vDGVCheque.Rows
                With row
                    'cek apakah data identifier perlu diubah
                    IdentifierLama = row.Cells(Frm.ColIdentifier.Index).Value.ToString
                    If IdentifierLama.StartsWith(vPrimaryJurnal.ToString) = False Then
                        'nilai start index = 1 agar tanda minus tidak diperhitungkan
                        counter = IdentifierLama.Substring(IdentifierLama.IndexOf("-", 1) + 1)
                        Identifier = vPrimaryJurnal.ToString + "-" + counter.ToString
                        'ganti data identifier lama dengan yang baru, untuk item
                        .Cells(Frm.ColIdentifier.Index).Value = Identifier
                    Else
                        Identifier = IdentifierLama
                    End If
                End With
            Next

            For Each row In vDGVCheque.Rows
                If row.Cells(Frm.ColSPaid.Index).Value = True Then
                    Paid = EnumPaid.Bayar
                Else
                    Paid = EnumPaid.Belum
                End If
                If row.Cells(Frm.ColTipeInOut.Index).Value = 0 Then
                    type = 0
                Else
                    type = 1
                End If
                If String.IsNullOrEmpty(row.Cells(Frm.ColRemarks.Index).Value) = True Then
                    row.Cells(Frm.ColRemarks.Index).Value = ""
                End If
                DTA.Insert(vPrimaryJurnal, row.Cells(Frm.ColChequeNo.Index).Value, row.Cells(Frm.ColBank.Index).Value, _
                 row.Cells(Frm.ColDate.Index).Value, row.Cells(Frm.ColDueDate.Index).Value, row.Cells(Frm.ColTerms.Index).Value, _
                 row.Cells(Frm.ColAmount.Index).Value, row.Cells(Frm.ColRemarks.Index).Value, Paid, row.Cells(Frm.ColIdentifier.Index).Value, _
                  type, row.Cells(Frm.ColPrimaryCurrency.Index).Value)
            Next

            vDGVCheque.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal vDGV As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTransaksiJournal_CashBank_Cheque) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_chequeTableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnal_chequeDataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnal_chequeRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            'kosongkan dgv
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.Rows.Clear()

            'isikan tabel
            If DTA.FillBy_PrimaryJurnal(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        DefPrecBase = ObjCur.GetDefaultPrecission(.Primary_Currency)
                        vDGV.Rows.Add("")
                        rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColChequeNo.Index).Value = .Cheque_No
                        rowDGV.Cells(Frm.ColBank.Index).Value = .Bank
                        rowDGV.Cells(Frm.ColDate.Index).Value = Format(.Cheque_date, "dd MMM yyyy")
                        rowDGV.Cells(Frm.ColDueDate.Index).Value = Format(.Due_Date, "dd MMM yyyy")
                        rowDGV.Cells(Frm.ColTerms.Index).Value = .Term
                        rowDGV.Cells(Frm.ColCurrency.Index).Value = ObjCur.GetCurrencyCode(.Primary_Currency)
                        rowDGV.Cells(Frm.ColRemarks.Index).Value = .Remarks
                        rowDGV.Cells(Frm.ColAmount.Index).Value = ObjCD.DigitGrouping(.Amount, DefPrecBase)
                        If .Paid = 0 Then
                            rowDGV.Cells(Frm.ColSPaid.Index).Value = True
                        Else
                            rowDGV.Cells(Frm.ColSPaid.Index).Value = False
                        End If
                        rowDGV.Cells(Frm.ColIdentifier.Index).Value = .Identifier
                        If .Type_InOut = 0 Then
                            rowDGV.Cells(Frm.ColTipeInOut.Index).Value = 0
                        Else
                            rowDGV.Cells(Frm.ColTipeInOut.Index).Value = 1
                        End If
                        rowDGV.Cells(Frm.ColPrimaryCurrency.Index).Value = .Primary_Currency
                    End With
                Next
            End If

            'set nomer
            ''ObjCD.SetNumber(vDGVItem)
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_chequeTableAdapter
        Try
            'hapus data item
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(vPrimaryTransaksi)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function UpdateCheque(ByVal vPrimaryJurnal As Double) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_chequeTableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnal_chequeDataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnal_chequeRow
        Try
            Dim IdentifierLama As String = ""
            Dim Identifier As String = ""
            Dim counter As Integer = 0
            'set identifier
            DTA.Fill(TBL)
            For Each row In TBL.Rows
                With row
                    If .Primary_Jurnal = -1 Then
                        'cek apakah data identifier perlu diubah
                        IdentifierLama = .Identifier
                        If IdentifierLama.StartsWith(vPrimaryJurnal.ToString) = False Then
                            'nilai start index = 1 agar tanda minus tidak diperhitungkan
                            counter = IdentifierLama.Substring(IdentifierLama.IndexOf("-", 1) + 1)
                            Identifier = vPrimaryJurnal.ToString + "-" + counter.ToString
                        Else
                            Identifier = IdentifierLama
                        End If
                        DTA.Update(vPrimaryJurnal, Identifier, -1)
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Fungsi laporan Cheque Aging
    Public Sub UpdatePaid_LapChequeAging(ByVal vPaid As ClsTransaksiJournal_Item_Cheque.EnumPaid, ByVal vIdenfifier As String)
        Dim DTA As New DseLaporanChequePerPeriod_AgingTableAdapters.tbltrans_jurnal_chequeTableAdapter
        Dim DTA1 As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim DTA2 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_chequeTableAdapter

        Dim PrimaryJurnal As Double = -1

        Try
            DTA.UpdateQuery_Paid(vPaid, vIdenfifier)
            'update timestamp untuk jurnal ybs
            If DTA2.ScalarQuery_GetPrimaryJurnal(vIdenfifier).HasValue = True Then
                PrimaryJurnal = DTA2.ScalarQuery_GetPrimaryJurnal(vIdenfifier).Value
                DTA1.UpdateQuery_UpdateTimestamp(Now, PrimaryJurnal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DTA2.Dispose()
        End Try
    End Sub
End Class
