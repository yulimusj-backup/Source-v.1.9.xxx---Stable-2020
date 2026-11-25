Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesInvoice_SettingTaxInvoice
    Public BatasStartDate As Date = #4/1/2013#
    Public BatasStartEFaktur As Date = #7/1/2013#
    Public PanjangNoSeriFP As Double = 13 + 3   '+ 3 karena memfasilitasi format penulisan .900-13.

    Public Sub IsiCBOFakturPajak(ByVal vYear As Date, ByVal vTBL As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_setting1DataTable, _
    ByVal vCBO As ComboBox)
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_setting1TableAdapter
        Dim row As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_setting1Row
        Dim StartStr, EndStr As String

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(vTBL, vYear)

            For Each row In vTBL.Rows
                StartStr = ""
                EndStr = ""

                StartStr = row.starting_number
                While StartStr.Length < 8
                    StartStr = "0" + StartStr
                End While
                EndStr = row.ending_number
                While EndStr.Length < 8
                    EndStr = "0" + EndStr
                End While
                row.display = row.prefix + StartStr + " - " + row.prefix + EndStr
            Next

            vCBO.DataSource = vTBL.DefaultView
            vCBO.DisplayMember = "display"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetEndingNumber(ByVal vTanggal As Date, ByVal vNoSeriFakturPajak As String) As Double
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter
        Dim TBL As New DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingDataTable
        Dim row As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrefixLength As Integer = 0
        Dim Dummy As String = ""
        Dim Number As Double = 0

        Try
            DTA.FillBy_Year(TBL, ObjCD.GetFisrtDateOfTheYear(vTanggal))
            For Each row In TBL.Rows
                If vNoSeriFakturPajak.StartsWith(row.prefix) = True Then
                    PrefixLength = row.prefix.Length

                    If vNoSeriFakturPajak.Length > PrefixLength Then
                        Dummy = vNoSeriFakturPajak.Substring(PrefixLength)
                        If IsNumeric(Dummy) = True Then
                            Number = Dummy
                            If Number >= row.starting_number And Number <= row.ending_number Then
                                Return row.ending_number
                            End If
                        End If
                    End If
                End If
            Next

            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            TBL.Dispose()
        End Try
    End Function

    Public Function GetStartingNumber(ByVal vTanggal As Date, ByVal vNoSeriFakturPajak As String) As Double
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter
        Dim TBL As New DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingDataTable
        Dim row As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrefixLength As Integer = 0
        Dim Dummy As String = ""
        Dim Number As Double = 0

        Try
            DTA.FillBy_Year(TBL, ObjCD.GetFisrtDateOfTheYear(vTanggal))
            For Each row In TBL.Rows
                If vNoSeriFakturPajak.StartsWith(row.prefix) = True Then
                    PrefixLength = row.prefix.Length

                    If vNoSeriFakturPajak.Length > PrefixLength Then
                        Dummy = vNoSeriFakturPajak.Substring(PrefixLength)
                        If IsNumeric(Dummy) = True Then
                            Number = Dummy
                            If Number >= row.starting_number And Number <= row.ending_number Then
                                Return row.starting_number
                            End If
                        End If
                    End If
                End If
            Next

            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            TBL.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vPrefix As String, ByVal vStartingNumber As Double, _
    ByVal vEndingNumber As Double, ByVal vYear As Double) As Double
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter
        Dim YearDate As Date = Today

        Try
            YearDate = CDate("#1/1/" + vYear.ToString + "#")

            If DTA.ScalarQuery_GetPrimary(vPrefix, vStartingNumber, vEndingNumber, YearDate).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vPrefix, vStartingNumber, vEndingNumber, YearDate).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vTanggal As Date, ByVal vNoSeriFakturPajak As String) As Double
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter
        Dim TBL As New DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingDataTable
        Dim row As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrefixLength As Integer = 0
        Dim Dummy As String = ""
        Dim Number As Double = 0

        Try
            DTA.FillBy_Year(TBL, ObjCD.GetFisrtDateOfTheYear(vTanggal))
            For Each row In TBL.Rows
                If vNoSeriFakturPajak.StartsWith(row.prefix) = True Then
                    PrefixLength = row.prefix.Length

                    If vNoSeriFakturPajak.Length > PrefixLength Then
                        Dummy = vNoSeriFakturPajak.Substring(PrefixLength)
                        If IsNumeric(Dummy) = True Then
                            Number = Dummy
                            If Number >= row.starting_number And Number <= row.ending_number Then
                                Return row.primary
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            TBL.Dispose()
        End Try
    End Function

    Public Sub SetNoUrutMaksimal(ByVal vPrimaryFP As Double, ByVal vNoSeriFakturPajak As String)
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter
        Dim TBL As New DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingDataTable
        Dim row As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrefixLength As Integer = 0
        Dim Dummy As String = ""
        Dim Number As Double = 0

        Try
            If vNoSeriFakturPajak = "" Then
                DTA.UpdateQuery_NoUrutMax(0, vPrimaryFP)
            Else
                DTA.FillBy_Primary(TBL, vPrimaryFP)
                For Each row In TBL.Rows
                    If vNoSeriFakturPajak.StartsWith(row.prefix) = True Then
                        PrefixLength = row.prefix.Length

                        If vNoSeriFakturPajak.Length > PrefixLength Then
                            Dummy = vNoSeriFakturPajak.Substring(PrefixLength)
                            If IsNumeric(Dummy) = True Then
                                Number = Dummy
                                If Number >= row.starting_number And Number <= row.ending_number Then
                                    DTA.UpdateQuery_NoUrutMax(Number, row.primary)
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            TBL.Dispose()
        End Try
    End Sub

    Public Sub SetNoUrutMaksimal(ByVal vTanggal As Date, ByVal vNoSeriFakturPajak As String)
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter
        Dim TBL As New DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingDataTable
        Dim row As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrefixLength As Integer = 0
        Dim Dummy As String = ""
        Dim Number As Double = 0

        Try
            DTA.FillBy_Year(TBL, ObjCD.GetFisrtDateOfTheYear(vTanggal))
            For Each row In TBL.Rows
                If vNoSeriFakturPajak.StartsWith(row.prefix) = True Then
                    PrefixLength = row.prefix.Length

                    If vNoSeriFakturPajak.Length > PrefixLength Then
                        Dummy = vNoSeriFakturPajak.Substring(PrefixLength)
                        If IsNumeric(Dummy) = True Then
                            Number = Dummy
                            If Number >= row.starting_number And Number <= row.ending_number Then
                                DTA.UpdateQuery_NoUrutMax(Number, row.primary)
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetNoSeriFakturPajakBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter
        Dim TBL As New DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingDataTable
        Dim row As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingRow
        Dim Number As Double = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim Dummy As String = ""

        Try
            'isikan berdasarkan tahun, untuk yg aktif saja
            DTA.FillBy_Year(TBL, ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi))

            For Each row In TBL.Rows
                If row.no_urut_max < row.ending_number Then
                    If row.no_urut_max = 0 Then
                        row.no_urut_max = row.starting_number - 1
                    End If

                    Dummy = row.no_urut_max + 1

                    While PanjangNoSeriFP - row.prefix.Length > Dummy.Length
                        Dummy = "0" + Dummy
                    End While

                    Return row.prefix + Dummy
                End If
            Next

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            TBL.Dispose()
        End Try
    End Function

    Public Function GetUrutan(ByVal vSeriFakturPajak As String, ByVal vTanggal As Date) As Double
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter
        Dim TBL As New DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingDataTable
        Dim row As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingRow
        Dim PrefixLength As Integer = 0
        Dim Dummy As String = ""
        Dim ObjCD As New Core_Module.ClsCekData
        Dim Number As Double = 0

        Try
            DTA.FillBy_Year(TBL, ObjCD.GetFisrtDateOfTheYear(vTanggal))

            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    If vSeriFakturPajak.StartsWith(row.prefix) = True Then
                        PrefixLength = row.prefix.Length

                        If vSeriFakturPajak.Length > PrefixLength Then
                            Dummy = vSeriFakturPajak.Substring(PrefixLength)
                            If IsNumeric(Dummy) = True Then
                                Number = Dummy
                                If Number >= row.starting_number And Number <= row.ending_number Then
                                    Return Dummy
                                End If
                            End If
                        End If
                    End If
                Next

                'kalo sampe sini ga ada yg cocok
                Return 0
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vPrimary As Double, ByVal vYear As Double, ByVal Prefix As String, _
    ByVal Remarks As String, ByVal StartingNumber As Double, ByVal EndingNumber As Double, ByVal vActive As Integer) As Boolean
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter

        Try
            DTA.UpdateQuery(Prefix, StartingNumber, EndingNumber, Remarks, CDate("1/1/" + vYear.ToString), vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanBaru(ByVal vYear As Double, ByVal Prefix As String, _
    ByVal Remarks As String, ByVal StartingNumber As Double, ByVal EndingNumber As Double, ByVal vActive As Integer) As Boolean
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter

        Try
            DTA.InsertQuery(Prefix, StartingNumber, EndingNumber, Remarks, CDate("1/1/" + vYear.ToString), 0, 1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter

        Try
            If DTA.DeleteQuery(vPrimary) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiDGV(ByVal vDGV As DataGridView, ByVal i As FrmTransaksiSalesInvoiceSettingPajak)
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter
        Dim TBL As New DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingDataTable
        Dim row As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingRow
        Dim rowDGV As DataGridViewRow

        Try
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGV.Rows.Clear()

            DTA.FillBy_SortPrimary(TBL)
            For Each row In TBL.Rows
                vDGV.Rows.Add()
                rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)

                If row.active = 1 Then
                    rowDGV.Cells(i.ColActive.Index).Value = Boolean.TrueString
                Else
                    rowDGV.Cells(i.ColActive.Index).Value = Boolean.FalseString
                End If
                rowDGV.Cells(i.ColEndingNumber.Index).Value = row.ending_number
                rowDGV.Cells(i.ColPrefix.Index).Value = row.prefix
                rowDGV.Cells(i.ColPrimary.Index).Value = row.primary
                rowDGV.Cells(i.ColRemarks.Index).Value = row.remarks
                rowDGV.Cells(i.ColStartingNumbe.Index).Value = row.starting_number
                rowDGV.Cells(i.ColYear.Index).Value = row.year.Year.ToString
            Next

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub
End Class
