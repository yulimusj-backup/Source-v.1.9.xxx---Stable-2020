Option Explicit On
Option Strict Off

Imports System.Windows.Forms

Public Class ClsTransaksiLoan_Item

    Dim ObjCd As Core_Module.ClsCekData

    Public Function GetIdentifierUrutan(ByVal vDGVItem As DataGridView, _
       ByVal vColIdentifier As Integer, ByVal baru As Boolean, _
       ByVal vPrimaryTransaksi As Double, ByVal vIndexCell As Integer) As String
        Dim Row As DataGridViewRow
        Try
            'vDGVItem.AllowUserToAddRows = False
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
            End If
            'vDGVItem.AllowUserToAddRows = True

            Return Identifier
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function IsiDGVItem(ByVal vDGVItem As Windows.Forms.DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTransaksiLoan, ByVal vDefPrec As Integer) As Boolean
        Dim DTA As New DseTransaksiLoanTableAdapters.tbltrans_loan_itemTableAdapter
        Dim TBL As New DseTransaksiLoan.tbltrans_loan_itemDataTable
        Dim row As DseTransaksiLoan.tbltrans_loan_itemRow
        Dim rowDGV As DataGridViewRow
        Dim ObjCekData As New Core_Module.ClsCekData
        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()

            ''isikan(tabel)
            If DTA.FillBy(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")

                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColIntallmentPeriod.Index).Value = .Installment_Period
                        'My.Application.ChangeCulture("id-ID")
                        'My.Application.ChangeUICulture("id-ID")
                        rowDGV.Cells(Frm.ColDateSchedule.Index).Value = Format(.Date_Schedule, "dd MMM yyyy")
                        'My.Application.ChangeCulture("en-US")
                        'My.Application.ChangeUICulture("en-US")
                        rowDGV.Cells(Frm.ColInterest.Index).Value = ObjCekData.DigitGrouping(.Interest, vDefPrec)
                        rowDGV.Cells(Frm.ColBase.Index).Value = ObjCekData.DigitGrouping(.Base, vDefPrec)
                        rowDGV.Cells(Frm.ColInstallmentAmount.Index).Value = ObjCekData.DigitGrouping(.Installment_Amount, vDefPrec)
                        rowDGV.Cells(Frm.ColBalance.Index).Value = ObjCekData.DigitGrouping(.Balance, vDefPrec)
                        My.Application.ChangeCulture("id-ID")
                        My.Application.ChangeUICulture("id-ID")
                        If .Payment_Date = #1/1/1950# Then
                            rowDGV.Cells(Frm.ColPaymentDate.Index).Value = Format("-")
                        Else
                            rowDGV.Cells(Frm.ColPaymentDate.Index).Value = Format(.Payment_Date, "dd MMM yyyy")
                        End If
                        My.Application.ChangeCulture("en-US")
                        My.Application.ChangeUICulture("en-US")
                        rowDGV.Cells(Frm.ColPaymentAmount.Index).Value = ObjCekData.DigitGrouping(.Payment_Amount, vDefPrec)
                        rowDGV.Cells(Frm.ColRemarks.Index).Value = .Remarks
                        rowDGV.Cells(Frm.ColIdentifier.Index).Value = .Identifier

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
            ObjCekData = Nothing
        End Try
    End Function

    Public Overridable Function SimpanItem(ByVal vPrimaryTransaksi As Double, ByVal vDGV As DataGridView, _
  ByVal Frm As FrmTransaksiLoan) As Boolean
        Dim DTA As New DseTransaksiLoanTableAdapters.tbltrans_loan_itemTableAdapter
        Dim row As DataGridViewRow
        Dim remark As String = ""
        Dim Date_ As Date = #1/1/1950#
        Dim Tgl As Date = #1/1/1950#
        Try
            Dim IdentifierLama As String = ""
            Dim Identifier As String = ""
            Dim counter As Integer = 0
            'set identifier
            For Each row In vDGV.Rows
                Tgl = #1/1/1950#
                With row
                    'cek apakah data identifier perlu diubah
                    IdentifierLama = row.Cells(Frm.ColIdentifier.Index).Value.ToString

                    If IdentifierLama.StartsWith(vPrimaryTransaksi.ToString) = False Then
                        'nilai start index = 1 agar tanda minus tidak diperhitungkan
                        counter = IdentifierLama.Substring(IdentifierLama.IndexOf("-", 1) + 1)
                        Identifier = vPrimaryTransaksi.ToString + "-" + counter.ToString
                        'ganti data identifier lama dengan yang baru, untuk item
                        .Cells(Frm.ColIdentifier.Index).Value = Identifier
                    Else
                        Identifier = IdentifierLama
                    End If
                End With
                If row.Cells(Frm.ColRemarks.Index).Value = Nothing Then
                    remark = ""
                Else
                    remark = row.Cells(Frm.ColRemarks.Index).Value
                End If
                My.Application.ChangeCulture("id-ID")
                My.Application.ChangeUICulture("id-ID")
                Date_ = row.Cells(Frm.ColDateSchedule.Index).Value
                If String.IsNullOrEmpty(row.Cells(Frm.ColPaymentDate.Index).Value) = False Then
                    If IsDate(row.Cells(Frm.ColPaymentDate.Index).Value) = True Then
                        Tgl = row.Cells(Frm.ColPaymentDate.Index).Value
                    End If
                End If

                My.Application.ChangeCulture("en-US")
                My.Application.ChangeUICulture("en-US")
                DTA.Insert(vPrimaryTransaksi, row.Cells(Frm.ColIntallmentPeriod.Index).Value, Date_, _
                row.Cells(Frm.ColInterest.Index).Value, row.Cells(Frm.ColBase.Index).Value, row.Cells(Frm.ColInstallmentAmount.Index).Value, _
                row.Cells(Frm.ColBalance.Index).Value, Tgl, row.Cells(Frm.ColPaymentAmount.Index).Value, _
                remark, Identifier)
            Next


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiLoanTableAdapters.tbltrans_loan_itemTableAdapter
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
