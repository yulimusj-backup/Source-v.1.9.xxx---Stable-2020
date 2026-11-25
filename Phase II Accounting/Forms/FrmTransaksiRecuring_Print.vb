''@rosadinnn
''20150826 1311
Option Explicit On
Option Strict Off
Public Class FrmTransaksiRecuring_Print

    Public Primary As Double = -1
    Dim Report As New CrpTransaksiRecuring_Print
    Dim Report2 As New CrpTransaksiRecuring_Print2
    Dim DefPrecIDR As Integer = 0
    Dim ObjCurrency As New Core_Module.clsMasterCurrency

    Private Sub FrmTransaksiRecuring_Print_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            DefPrecIDR = Me.ObjCurrency.GetDefaultPrecission(1)
            Me.Cetak(Me.IsIDROnly)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsIDROnly() As Boolean
        Dim Dta As New DseTransaksiRecuring_PrintTableAdapters.tblmaster_currencyTableAdapter
        Dim Tbl As New DseTransaksiRecuring_Print.tblmaster_currencyDataTable
        Dim Row As DseTransaksiRecuring_Print.tblmaster_currencyRow
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Dta.FillBy_RecurringPrimary(Tbl, Me.Primary) > 0 Then
                For Each Row In Tbl
                    If Row.Currency_Code <> "IDR" Then
                        Return False
                    End If
                Next
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function Cetak(ByVal vIDROnly As Boolean) As Boolean
        Dim Ds As New DseTransaksiRecuring_Print
        Dim DtaHeader As New DseTransaksiRecuring_PrintTableAdapters.tbltrans_recurring_journalTableAdapter
        Dim DtaAccount As New DseTransaksiRecuring_PrintTableAdapters.DataTable1TableAdapter
        Dim DtaJournal As New DseTransaksiRecuring_PrintTableAdapters.tbltrans_recurring_journal_transaction_entriesTableAdapter
        Dim TblHeader As DseTransaksiRecuring_Print.tbltrans_recurring_journalDataTable = Ds.tbltrans_recurring_journal
        Dim TblAccount As DseTransaksiRecuring_Print.DataTable1DataTable = Ds.DataTable1
        Dim TblJournal As DseTransaksiRecuring_Print.tbltrans_recurring_journal_transaction_entriesDataTable = Ds.tbltrans_recurring_journal_transaction_entries
        Dim rowJ As DseTransaksiRecuring_Print.tbltrans_recurring_journal_transaction_entriesRow
        Try
            If Me.Primary > -1 Then
                Me.Cursor = Cursors.WaitCursor
                DtaHeader.Connection.ConnectionString = Core_Module.GetConstring
                DtaAccount.Connection.ConnectionString = Core_Module.GetConstring
                DtaJournal.Connection.ConnectionString = Core_Module.GetConstring
                If DtaHeader.FillBy_Primary(TblHeader, Me.Primary) > 0 Then
                    If DtaAccount.FillBy_Primary(TblAccount, Me.Primary) > 0 Then
                        If DtaJournal.FillBy_Primary(TblJournal, Me.Primary) > 0 Then
                            Dim IntA As Integer = 0
                            Dim EndDate As Date = Now
                            For Each rowJ In TblJournal.Rows
                                With rowJ
                                    IntA += 1
                                    If IntA = TblJournal.Rows.Count Then
                                        EndDate = .Journal_Date
                                    End If
                                End With
                            Next
                            If vIDROnly = True Then
                                Me.Report.SetDataSource(Ds)
                                Me.Report.SetParameterValue("DefPrecIDR", DefPrecIDR)
                                Me.Report.SetParameterValue("EndDate", EndDate)
                                Me.CrvRecurring.ReportSource = Me.Report
                            Else
                                Me.Report2.SetDataSource(Ds)
                                Me.Report2.SetParameterValue("DefPrecIDR", DefPrecIDR)
                                Me.Report2.SetParameterValue("EndDate", EndDate)
                                Me.CrvRecurring.ReportSource = Me.Report2
                            End If
                            Me.CrvRecurring.Zoom(90)
                            Me.CrvRecurring.Refresh()
                            Me.CrvRecurring.Show()
                        Else
                            MsgBox("There is no transaction entries.", MsgBoxStyle.Information, "Information")
                            Me.Close()
                        End If
                    Else
                        MsgBox("There is no accounts entries.", MsgBoxStyle.Information, "Information")
                        Me.Close()
                    End If
                Else
                    MsgBox("Invalid primary of recurring journal.", MsgBoxStyle.Information, "Information")
                    Me.Close()
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Function

    Private Sub BtnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class