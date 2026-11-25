Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports System.Data
Imports System.Deployment
Imports System.Drawing


Public Class FrmTrans_Gaji_Override_Create
    Private StartDate1, EndDate1 As Date
    Public UserCode As String = "USER"

    Dim ObjCD As New Core_Module.ClsCekData
    Dim StartDate, EndDate As Date

    'menghapuskan data overide salary dan other payment, jika tidak digunakan pada salary
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim DTA As New DseTransaksiSalary_Override_CreateTableAdapters.tbltrans_salaryTableAdapter
        Dim DummyDate As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari start date dan end date pada transaksi gaji berdasarkan tanggal gajian
            DummyDate = DTA.ScalarQuery_GetStartMin(Me.dtpTO.Value.Date, Me.dtpFrom.Value.Date)
            If DummyDate.HasValue = False Then
                MsgBox("No salary data found.", MsgBoxStyle.Information, "Information")
                Exit Sub
            Else
                StartDate1 = DummyDate.Value
            End If
            EndDate1 = DTA.ScalarQuery_GetEndMax(Me.dtpTO.Value.Date, Me.dtpFrom.Value.Date).Value

            If MsgBox("Start data check for delete?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgw1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HapusOtherPayment(ByVal StartDate As Date, ByVal EndDate As Date)
        Dim DTA1 As New DseTransaksiSalary_Override_CreateTableAdapters.tbltrans_salary_override_amount_itemTableAdapter
        Dim TBL1 As New DseTransaksiSalary_Override_Create.tbltrans_salary_override_amount_itemDataTable
        Dim row1 As DseTransaksiSalary_Override_Create.tbltrans_salary_override_amount_itemRow
        Dim DTA2 As New DseTransaksiSalary_Override_CreateTableAdapters.tbltrans_salary_item_otherTableAdapter
        Dim TBL2 As New DseTransaksiSalary_Override_Create.tbltrans_salary_item_otherDataTable
        Dim row2 As DseTransaksiSalary_Override_Create.tbltrans_salary_item_otherRow
        Dim TipeOverride As ClsTransaksiSalary_Override.EnumType = ClsTransaksiSalary_Override.EnumType.OtherPayment
        Dim Ketemu As Boolean = False

        Try
            Me.Prgb1.Value = 0

            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = DTA1.Connection.ConnectionString

            'loading other income sesuai start date dan end date salary
            Me.LblProgress.Text = "Phase 1 of 3: loading override data"
            DTA1.Fill(TBL1, EndDate, StartDate, CInt(TipeOverride))

            'loading other payment dr transaksi salary, berdasarkan start date dan end date salary
            Me.LblProgress.Text = "Phase 1 of 3: loading item other payment"
            DTA2.Fill(TBL2, Me.dtpTO.Value.Date, Me.dtpFrom.Value.Date)

            'cek untuk setiap baris pada item transaksi other income, pastikan ada di item salary other payment
            'cek berdsarkan: date, identifier salary component, primary ref doc, primary employee, amount
            'jika tidak ditemukan pasangannya, maka langsung hapus dari db
            Me.Prgb1.Maximum = TBL1.Rows.Count
            While TBL1.Rows.Count
                row1 = TBL1.Rows(0)
                Ketemu = False
                Me.Prgb1.Value += 1
                Me.LblProgress.Text = "Phase 1 of 3: checking data " + Me.Prgb1.Value.ToString + " of " + Me.Prgb1.Maximum.ToString

                For Each row2 In TBL2.Rows
                    If row1.Primary_Transaksi < row2.Primary_Ref_Doc Then
                        Exit For
                    Else
                        If row1.Primary_Transaksi = row2.Primary_Ref_Doc Then
                            If row1.Primary_Employee = row2.Primary_Employee And _
                            row1._Date >= row2.Start_Date And row1._Date <= row2.End_Date And _
                            row1.Identifier_Formula = row2.Identifier_Salary_Component And _
                            row1.Override_Amount = row2.Amount Then
                                Ketemu = True
                                Exit For
                            End If
                        End If
                    End If
                Next

                If Ketemu = False Then
                    DTA1.DeleteQuery(row1.Primary)
                End If

                TBL1.Rows.Remove(row1)
            End While

            'cek override amount
            Me.HapusOvverideData(StartDate, EndDate)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Rows.Clear()
            TBL1.Dispose()
            DTA2.Dispose()
            TBL2.Rows.Clear()
            TBL2.Dispose()
        End Try
    End Sub

    Private Sub HapusOvverideData(ByVal StartDate As Date, ByVal EndDate As Date)
        Dim DTA1 As New DseTransaksiSalary_Override_CreateTableAdapters.tbltrans_salary_override_amount_itemTableAdapter
        Dim TBL1 As New DseTransaksiSalary_Override_Create.tbltrans_salary_override_amount_itemDataTable
        Dim row1 As DseTransaksiSalary_Override_Create.tbltrans_salary_override_amount_itemRow
        Dim DTA2 As New DseTransaksiSalary_Override_CreateTableAdapters.tbltrans_salary_itemTableAdapter
        Dim TBL2 As New DseTransaksiSalary_Override_Create.tbltrans_salary_itemDataTable
        Dim row2 As DseTransaksiSalary_Override_Create.tbltrans_salary_itemRow
        Dim TipeOverride As ClsTransaksiSalary_Override.EnumType = ClsTransaksiSalary_Override.EnumType.SalaryOverride
        Dim Ketemu As Boolean = False

        Try
            Me.Prgb1.Value = 0

            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = DTA1.Connection.ConnectionString

            'loading other income sesuai start date dan end date salary
            Me.LblProgress.Text = "Phase 2 of 3: loading override data"
            DTA1.FillBy1(TBL1, EndDate, StartDate, CInt(TipeOverride))

            'loading other payment dr transaksi salary, berdasarkan start date dan end date salary
            Me.LblProgress.Text = "Phase 2 of 3: salary item data"
            DTA2.Fill(TBL2, Me.dtpTO.Value.Date, Me.dtpFrom.Value.Date)

            'cek untuk setiap baris pada item transaksi other income, pastikan ada di item salary other payment
            'cek berdsarkan: date, identifier salary component, primary employee, amount
            'jika tidak ditemukan pasangannya, maka langsung hapus dari db
            Me.Prgb1.Maximum = TBL1.Rows.Count
            While TBL1.Rows.Count
                row1 = TBL1.Rows(0)
                Ketemu = False
                Me.Prgb1.Value += 1
                Me.LblProgress.Text = "Phase 2 of 3: checking data " + Me.Prgb1.Value.ToString + " of " + Me.Prgb1.Maximum.ToString

                For Each row2 In TBL2.Rows
                    If row1.Primary_Employee < row2.Primary_Employee Then
                        Exit For
                    ElseIf row1.Primary_Employee = row2.Primary_Employee Then
                        If row1._Date >= row2.Start_Date And row1._Date <= row2.End_Date And _
                            row1.Identifier_Formula = row2.Identifier_Salary_Component And _
                            row1.Override_Amount = row2.Amount Then
                            Ketemu = True
                            Exit For
                        End If
                    End If
                Next

                If Ketemu = False Then
                    DTA1.DeleteQuery(row1.Primary)
                End If

                TBL1.Rows.Remove(row1)
            End While

            'cek override amount
            Me.DeleteTransaksiOverrideWitnNoItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Rows.Clear()
            TBL1.Dispose()
            DTA2.Dispose()
            TBL2.Rows.Clear()
            TBL2.Dispose()
        End Try
    End Sub

    Private Sub DeleteTransaksiOverrideWitnNoItem()
        Dim DTA As New DseTransaksiSalary_Override_CreateTableAdapters.tbltrans_salary_override_amountTableAdapter
        Dim TBL As New DseTransaksiSalary_Override_Create.tbltrans_salary_override_amountDataTable
        Dim row As DseTransaksiSalary_Override_Create.tbltrans_salary_override_amountRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.Prgb1.Value = 0
            Me.LblProgress.Text = "Phase 3 of 3: loading data from db"
            DTA.Fill(TBL)

            Me.Prgb1.Maximum = TBL.Rows.Count
            For Each row In TBL.Rows
                Me.Prgb1.Value += 1
                Me.LblProgress.Text = "Phase 3 of 3: deleting data " + Me.Prgb1.Value.ToString + " of " + Me.Prgb1.Maximum.ToString

                DTA.DeleteQuery(row.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Rows.Clear()
            TBL.Dispose()
        End Try
    End Sub

    'membuat data other payment dan override salary, berdasarkan transaksi salary
    Private Sub btnCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreate.Click
        Dim DTA As New DseTransaksiSalary_Override_CreateTableAdapters.tbltrans_salaryTableAdapter
        Dim DummyDate As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari start date dan end date pada transaksi gaji berdasarkan tanggal gajian
            DummyDate = DTA.ScalarQuery_GetStartMin(Me.dtpTO.Value.Date, Me.dtpFrom.Value.Date)
            If DummyDate.HasValue = False Then
                MsgBox("No salary data found.", MsgBoxStyle.Information, "Information")
                Exit Sub
            Else
                StartDate1 = DummyDate.Value
            End If
            EndDate1 = DTA.ScalarQuery_GetEndMax(Me.dtpTO.Value.Date, Me.dtpFrom.Value.Date).Value

            If MsgBox("Start create override data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgw2.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_Override_Create_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Prgb1.Value = 0
            Me.LblProgress.Text = ""
            Me.dgvTaxCategory.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Try
            Me.ObjCD.UpdateThreading()
            Me.HapusOtherPayment(Me.StartDate1, Me.EndDate1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Try
            Me.Prgb1.Value = 0
            Me.LblProgress.Text = ""
            MsgBox("Data checked.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw2.DoWork
        Try
            Me.ObjCD.UpdateThreading()
            Me.CreateOverrideData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreateOverrideData()
        Dim DTA1 As New DseTransaksiSalary_Override_CreateTableAdapters.tbltrans_salary_override_amount_itemTableAdapter
        Dim TBL1 As New DseTransaksiSalary_Override_Create.tbltrans_salary_override_amount_itemDataTable
        Dim row1 As DseTransaksiSalary_Override_Create.tbltrans_salary_override_amount_itemRow
        Dim DTA2 As New DseTransaksiSalary_Override_CreateTableAdapters.tbltrans_salary_itemTableAdapter
        Dim TBL2 As New DseTransaksiSalary_Override_Create.tbltrans_salary_itemDataTable
        Dim row2 As DseTransaksiSalary_Override_Create.tbltrans_salary_itemRow
        Dim TipeOverride As ClsTransaksiSalary_Override.EnumType = ClsTransaksiSalary_Override.EnumType.SalaryOverride
        Dim Ketemu As Boolean = False
        Dim DTAH As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim DTAI As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount_itemTableAdapter
        Dim DocNo As String = ""
        Dim PrimaryRef As Double = -1
        Dim Obj As New ClsTransaksiSalary_Override
        Dim ObjCD As New Core_Module.ClsCekData
        Dim FormulaType As ClsTransaksiSalary_Override.EnumFormulaType = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula

        Try
            Me.Prgb1.Value = 0

            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = DTA1.Connection.ConnectionString
            DTAH.Connection.ConnectionString = DTA1.Connection.ConnectionString
            DTAI.Connection.ConnectionString = DTA1.Connection.ConnectionString

            'load data item salary, berdasarkan tanggal salary
            Me.LblProgress.Text = "Phase 1 of 2: loading item salary"
            DTA2.Fill(TBL2, Me.dtpTO.Value.Date, Me.dtpFrom.Value.Date)

            'load item override
            Me.LblProgress.Text = "Phase 1 of 2: loading override transaction"
            DTA1.FillBy1(TBL1, Me.EndDate1, Me.StartDate1, CInt(TipeOverride))

            'jika item salary tersebut override, maka pastikan sumber datanya ada, jika tidak ada, maka buat
            'cek berdsarkan: date, identifier salary component, primary employee, amount
            'jika tidak ditemukan, maka buat transaksi override
            Me.Prgb1.Maximum = TBL2.Rows.Count
            While TBL2.Rows.Count > 0
                Me.Prgb1.Value += 1
                Me.LblProgress.Text = "Phase 1 of 2: checking data " + Me.Prgb1.Value.ToString + " of " + Me.Prgb1.Maximum.ToString

                row2 = TBL2.Rows(0)
                Ketemu = False

                For Each row1 In TBL1.Rows
                    If row2.Primary_Employee < row1.Primary_Employee Then
                        Exit For
                    ElseIf row2.Primary_Employee = row1.Primary_Employee Then
                        If row1._Date >= row2.Start_Date And row1._Date <= row2.End_Date And _
                        row1.Identifier_Formula = row2.Identifier_Salary_Component And _
                        row1.Override_Amount = row2.Amount Then
                            Ketemu = True
                            Exit For
                        End If
                    End If
                Next

                If Ketemu = False Then
                    PrimaryRef = Me.GetPrimaryAcuan(row2.End_Date, TipeOverride, row2.Identifier_Salary_Component)
                    If PrimaryRef = -1 Then
                        'simpan data header
                        DocNo = Obj.GetNoBaru(row2.End_Date, ObjCD, TipeOverride)
                        DTAH.Insert(DocNo, row2.End_Date, CInt(FormulaType), Me.GetPrimary(row2.Identifier_Salary_Component), row2.Identifier_Salary_Component, _
                        row2.Remarks, Me.UserCode, Now, Me.UserCode, Now, Now, CDbl(DocNo.Substring(0, 6)), TipeOverride)
                        'simpan data item
                        PrimaryRef = Obj.GetPrimary(DocNo)
                        DTAI.Insert(PrimaryRef, row2.Primary_Employee, row2.Remarks, row2.Amount)
                    Else
                        'simpan data item 
                        DTAI.Insert(PrimaryRef, row2.Primary_Employee, row2.Remarks, row2.Amount)
                    End If
                End If

                TBL2.Rows.Remove(row2)
            End While

            Me.CreatedOtherPaymentData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Rows.Clear()
            TBL1.Dispose()
            DTA2.Dispose()
            TBL2.Rows.Clear()
            TBL2.Dispose()
            DTAH.Dispose()
            DTAI.Dispose()
            Obj = Nothing
            ObjCD = Nothing
        End Try
    End Sub

    Private Function GetPrimary(ByVal vIdentifierComp As String) As Double
        Try
            If vIdentifierComp.Contains("_") = True Then
                Return vIdentifierComp.Substring(0, vIdentifierComp.IndexOf("_"))
            Else
                Return 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 1
        End Try
    End Function

    Private Sub CreatedOtherPaymentData()
        Dim DTA1 As New DseTransaksiSalary_Override_CreateTableAdapters.tbltrans_salary_override_amount_itemTableAdapter
        Dim TBL1 As New DseTransaksiSalary_Override_Create.tbltrans_salary_override_amount_itemDataTable
        Dim row1 As DseTransaksiSalary_Override_Create.tbltrans_salary_override_amount_itemRow
        Dim DTA2 As New DseTransaksiSalary_Override_CreateTableAdapters.tbltrans_salary_item_otherTableAdapter
        Dim TBL2 As New DseTransaksiSalary_Override_Create.tbltrans_salary_item_otherDataTable
        Dim row2 As DseTransaksiSalary_Override_Create.tbltrans_salary_item_otherRow
        Dim TipeOverride As ClsTransaksiSalary_Override.EnumType = ClsTransaksiSalary_Override.EnumType.OtherPayment
        Dim Ketemu As Boolean = False
        Dim DTAH As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim DTAI As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount_itemTableAdapter
        Dim DocNo As String = ""
        Dim PrimaryRef As Double = -1
        Dim Obj As New ClsTransaksiSalary_Override
        Dim ObjCD As New Core_Module.ClsCekData
        Dim FormulaType As ClsTransaksiSalary_Override.EnumFormulaType = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula

        Try
            Me.Prgb1.Value = 0

            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = DTA1.Connection.ConnectionString
            DTAH.Connection.ConnectionString = DTA1.Connection.ConnectionString
            DTAI.Connection.ConnectionString = DTA1.Connection.ConnectionString

            'loading other income sesuai start date dan end date salary
            Me.LblProgress.Text = "Phase 2 of 2: loading override data"
            DTA1.Fill(TBL1, EndDate1, StartDate1, CInt(TipeOverride))

            'loading other payment dr transaksi salary, berdasarkan start date dan end date salary
            Me.LblProgress.Text = "Phase 2 of 2: loading item other payment"
            DTA2.Fill(TBL2, Me.dtpTO.Value.Date, Me.dtpFrom.Value.Date)

            'cek untuk setiap baris pada item transaksi other income salary, pastikan ada di item orher income
            'cek berdsarkan: date, identifier salary component, primary ref doc, primary employee, amount
            'jika tidak ditemukan pasangannya, maka langsung hapus dari db
            Me.Prgb1.Maximum = TBL2.Rows.Count
            While TBL2.Rows.Count > 0
                row2 = TBL2.Rows(0)
                Ketemu = False
                Me.LblProgress.Text = "Phase 2 of 2: checking item override - data: " + Me.Prgb1.Value.ToString + " of " + Me.Prgb1.Maximum.ToString

                For Each row1 In TBL1.Rows
                    If row2.Primary_Ref_Doc < row1.Primary_Transaksi Then
                        Exit For
                    ElseIf row2.Primary_Ref_Doc = row1.Primary_Transaksi Then
                        If row1.Primary_Employee = row2.Primary_Employee Then
                            If row1._Date >= row2.Start_Date And row1._Date <= row2.End_Date And _
                            row1.Identifier_Formula = row2.Identifier_Salary_Component And _
                            row1.Override_Amount = row2.Amount Then
                                Ketemu = True
                                Exit For
                            End If
                        End If
                    End If
                Next

                If Ketemu = False Then
                    PrimaryRef = Me.GetPrimaryAcuan(row2.End_Date, TipeOverride, row2.Identifier_Salary_Component)
                    If PrimaryRef = -1 Then
                        'simpan data header
                        DocNo = Obj.GetNoBaru(row2.End_Date, ObjCD, TipeOverride)
                        DTAH.Insert(DocNo, row2.End_Date, CInt(FormulaType), Me.GetPrimary(row2.Identifier_Salary_Component), row2.Identifier_Salary_Component, _
                        row2.Remarks, Me.UserCode, Now, Me.UserCode, Now, Now, CDbl(DocNo.Substring(0, 6)), TipeOverride)
                        'simpan data item
                        PrimaryRef = Obj.GetPrimary(DocNo)
                        DTAI.Insert(PrimaryRef, row2.Primary_Employee, row2.Remarks, row2.Amount)
                    Else
                        DTAI.Insert(PrimaryRef, row2.Primary_Employee, row2.Remarks, row2.Amount)
                    End If
                End If

                TBL2.Rows.Remove(row2)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Rows.Clear()
            TBL1.Dispose()
            DTA2.Dispose()
            TBL2.Rows.Clear()
            TBL2.Dispose()
            Obj = Nothing
            ObjCD = Nothing
            DTAH.Dispose()
            DTAI.Dispose()
        End Try
    End Sub

    Private Function GetPrimaryAcuan(ByVal vDate As Date, ByVal vTipeTrans As ClsTransaksiSalary_Override.EnumType, ByVal vIdentifierSalaryKomponen As String) As Double
        Dim DTA As New DseTransaksiSalary_Override_CreateTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim TBL As New DseTransaksiSalary_Override_Create.tbltrans_salary_override_amount1DataTable
        Dim row As DseTransaksiSalary_Override_Create.tbltrans_salary_override_amount1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vIdentifierSalaryKomponen, vDate, CInt(vTipeTrans))
            If TBL.Rows.Count = 0 Then
                Return -1
            Else
                row = TBL.Rows(0)
                Return row.Primary
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Private Sub bgw2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw2.RunWorkerCompleted
        Try
            Me.Prgb1.Value = 0
            Me.LblProgress.Text = ""
            MsgBox("Done.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Dim DTASalary As New DseTransaksiSalary_OverrideTableAdapters.DataTable6TableAdapter
        Dim TBLSalary As New DseTransaksiSalary_Override.DataTable6DataTable
        Dim rowSalary As DseTransaksiSalary_Override.DataTable6Row

        Try
            DTASalary.Connection.ConnectionString = Core_Module.GetConstring

            If MsgBox("Start creating override data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                '1. cari start date dan end date
                DTASalary.Fill(TBLSalary, Me.dtpSalaryDate.Value.Date)
                If TBLSalary.Rows.Count > 0 Then
                    rowSalary = TBLSalary.Rows(0)
                    Me.StartDate = rowSalary.Start_Date
                    Me.EndDate = rowSalary.End_Date

                    Control.CheckForIllegalCrossThreadCalls = False
                    Me.bgw3.RunWorkerAsync()
                Else
                    MsgBox("No salary data found.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASalary.Dispose()
            TBLSalary.Dispose()
        End Try
    End Sub

    Private Sub bgw3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw3.DoWork
        Dim Tipe As ClsTransaksiSalary_Override.EnumType = ClsTransaksiSalary_Override.EnumType.SalaryOverride
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter

        Dim DTAH As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim DTAI As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount_itemTableAdapter
        Dim DocNo As String = ""
        Dim PrimaryRef As Double = -1
        Dim Obj As New ClsTransaksiSalary_Override
        Dim ObjCD As New Core_Module.ClsCekData
        Dim FormulaType As ClsTransaksiSalary_Override.EnumFormulaType = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula
        Dim IdentifierDummy As String = "-" 
        Dim DTASalaryItem As New DseTransaksiSalary_OverrideTableAdapters.DataTable7TableAdapter
        Dim TBLSalaryItem As New DseTransaksiSalary_Override.DataTable7DataTable
        Dim rowSalaryItem As DseTransaksiSalary_Override.DataTable7Row

        Try
            Me.ObjCD.UpdateThreading()

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAH.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAI.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTASalaryItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            '1. hapus transaksi override yang berada di antara start date dan end date dan bertipe override data
            Me.LblProgress.Text = "Deleting override data."
            DTA.DeleteQuery1(Me.EndDate, Me.StartDate, CInt(Tipe))

            '2. fill item salary dan buat transaksi override berdasarkan item salary
            DTASalaryItem.Fill(TBLSalaryItem, Me.dtpSalaryDate.Value.Date)
            Me.LblProgress.Text = "Loading salary item data."
            Me.Prgb1.Value = 0
            Me.Prgb1.Maximum = TBLSalaryItem.Rows.Count
            For Each rowSalaryItem In TBLSalaryItem.Rows
                Me.Prgb1.Value += 1
                If Me.Prgb1.Value < 500 Then
                    Me.LblProgress.Text = "Creating override data: " + Me.Prgb1.Value.ToString + " - " + Me.Prgb1.Maximum.ToString
                Else
                    If Math.IEEERemainder(Me.Prgb1.Value, 50) = 0 Then
                        Me.LblProgress.Text = "Creating override data: " + Me.Prgb1.Value.ToString + " - " + Me.Prgb1.Maximum.ToString
                    End If
                End If

                If IdentifierDummy <> rowSalaryItem.Identifier_Salary_Component Then
                    IdentifierDummy = rowSalaryItem.Identifier_Salary_Component

                    'simpan data header baru
                    DocNo = Obj.GetNoBaru(Me.EndDate, ObjCD, Tipe)
                    DTAH.Insert(DocNo, Me.EndDate, CInt(FormulaType), Me.GetPrimary(IdentifierDummy), IdentifierDummy, _
                    "", Me.UserCode, Now, Me.UserCode, Now, Now, CDbl(DocNo.Substring(0, 6)), Tipe)
                    PrimaryRef = Obj.GetPrimary(DocNo)
                End If

                'simpan data item
                DTAI.Insert(PrimaryRef, rowSalaryItem.Primary_Employee, "", rowSalaryItem.Amount)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAH.Dispose()
            DTAI.Dispose()
            Obj = Nothing
            ObjCD = Nothing
            DTASalaryItem.Dispose()
            TBLSalaryItem.Dispose()
        End Try
    End Sub

    Private Sub bgw3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw3.RunWorkerCompleted
        Try
            Me.Prgb1.Value = 0
            Me.LblProgress.Text = ""
            MsgBox("Data created." + vbCrLf + _
            "Cek apakah perlu jalankan fungsi update kategory pajak. Jika sudah ada perhitungan gaji u/ jan bulan selanjutnya, kemungkinan perlu.", _
            MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStart2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart2.Click
        Try
            'confirm
            If MsgBox("Start data update?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgw4.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 

    Private Sub btnSHow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSHow1.Click
        Dim DTA As New DseTransaksiSalary_Override_CreateTableAdapters.tblmaster_karyawanTableAdapter
        Dim TBL As New DseTransaksiSalary_Override_Create.tblmaster_karyawanDataTable
        Dim row As DseTransaksiSalary_Override_Create.tblmaster_karyawanRow
        Dim rowDGV As DataGridViewRow
        Dim ObjKary As New Core_Module.ClsMasterKaryawan

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)

            Me.dgvTaxCategory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvTaxCategory.Rows.Clear()

            For Each row In TBL.Rows
                Me.dgvTaxCategory.Rows.Add()
                rowDGV = Me.dgvTaxCategory.Rows(Me.dgvTaxCategory.Rows.Count - 1)
                rowDGV.Cells(Me.ColCode.Index).Value = row.Code
                rowDGV.Cells(Me.ColName.Index).Value = row.Full_Name
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTaxCategory.Rows.Count
                rowDGV.Cells(Me.ColTaxCategory.Index).Value = ObjKary.GetNPWPCategoryStr(row.npwp_category)
                rowDGV.Cells(Me.ColResign.Index).Value = CBool(row.Aktif)
            Next

            Me.dgvTaxCategory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub bgw4_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw4.DoWork
        Dim DTA As New DseTransaksiSalary_Override_CreateTableAdapters.tbltrans_salary1TableAdapter
        Dim TBL As New DseTransaksiSalary_Override_Create.tbltrans_salary1DataTable
        Dim row As DseTransaksiSalary_Override_Create.tbltrans_salary1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Me.ObjCD.UpdateThreading()

            'load data salary sesuai tanggal yang diberikan
            Me.LblProgress.Text = "Loading salary data"
            DTA.Fill(TBL, Me.dtpSalaryDate2.Value.Date)
            Me.Prgb1.Value = 0
            Me.Prgb1.Maximum = TBL.Rows.Count

            For Each row In TBL.Rows
                Me.Prgb1.Value += 1
                Me.LblProgress.Text = "Updating data " + Me.Prgb1.Value.ToString + " of " + Me.Prgb1.Maximum.ToString

                DTA.UpdateQuery(row.A_Status_NPWP, row.Primary_Employee)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub bgw4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw4.RunWorkerCompleted
        Try
            MsgBox("Data updated" + vbCrLf + "Check:" + vbCrLf + _
            "1. Show1: make sure all tax category is true." + vbCrLf + _
            "2. Show2: make sure all employee in salary calculation, has basic salary.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' mengecek, pastikan semua employee pada gajian dengan tanggal dimaksud, masih memiliki nilai basic salary
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnShow2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow2.Click
        Dim DTASalary As New DseTransaksiSalary_Override_CreateTableAdapters.tbltrans_salary1TableAdapter
        Dim TBLSalary As New DseTransaksiSalary_Override_Create.tbltrans_salary1DataTable
        Dim rowSalary As DseTransaksiSalary_Override_Create.tbltrans_salary1Row
        Dim DTAFormula As New DseTransaksiSalary_Override_CreateTableAdapters.tblmaster_salary_formula_component_detailemployeeTableAdapter
        Dim TBLFormula As New DseTransaksiSalary_Override_Create.tblmaster_salary_formula_component_detailemployeeDataTable
        Dim rowFormula As DseTransaksiSalary_Override_Create.tblmaster_salary_formula_component_detailemployeeRow
        Dim idx1, idx2 As Double
        Dim ketemu As Boolean = False
        Dim rowDGV As DataGridViewRow
        Dim ObjKary As New Core_Module.ClsMasterKaryawan

        Try
            If MsgBox("Start data check?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                DTASalary.Connection.ConnectionString = Core_Module.GetConstring
                DTAFormula.Connection.ConnectionString = DTASalary.Connection.ConnectionString

                '1. fill tabel
                TBLSalary.Rows.Clear()
                TBLFormula.Rows.Clear()

                DTASalary.FillBy1(TBLSalary, Me.dtpSalaryDate2.Value.Date)
                DTAFormula.Fill(TBLFormula)

                '2. buang setiap baris di tabel salary, yang di tabel formula
                While idx1 <= TBLSalary.Rows.Count - 1
                    ketemu = False
                    idx2 = 0
                    rowSalary = TBLSalary.Rows(idx1)

                    While idx2 <= TBLFormula.Rows.Count - 1
                        rowFormula = TBLFormula.Rows(idx2)
                        If rowSalary.Primary_Employee = rowFormula.Primary_Karyawan Then
                            ketemu = True
                            Exit While
                        ElseIf rowSalary.Primary_Employee > rowFormula.Primary_Karyawan Then
                            TBLFormula.Rows.RemoveAt(idx2)
                        Else
                            Exit While
                        End If
                    End While

                    If ketemu = True Then
                        TBLSalary.Rows.RemoveAt(idx1)
                        TBLFormula.Rows.RemoveAt(idx2)
                    Else
                        idx1 += 1
                    End If
                End While

                Me.dgvTaxCategory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvTaxCategory.Rows.Clear()

                If TBLSalary.Rows.Count > 0 Then

                    For Each rowSalary In TBLSalary.Rows
                        Me.dgvTaxCategory.Rows.Add()
                        rowDGV = Me.dgvTaxCategory.Rows(Me.dgvTaxCategory.Rows.Count - 1)
                        rowDGV.Cells(Me.ColCode.Index).Value = ObjKary.GetCodeByPrimary(rowSalary.Primary_Employee)
                        rowDGV.Cells(Me.ColName.Index).Value = ObjKary.GetFullNameByPrimary(rowSalary.Primary_Employee)
                        rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTaxCategory.Rows.Count
                    Next


                    MsgBox("Please specify basic salary for the following employee(s).", MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Data checked.")
                End If

                Me.dgvTaxCategory.Sort(Me.ColCode, System.ComponentModel.ListSortDirection.Ascending)
                Me.ObjCD.SetNumber(Me.dgvTaxCategory)
                Me.dgvTaxCategory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASalary.Dispose()
            TBLSalary.Dispose()
            DTAFormula.Dispose()
            TBLFormula.Dispose()
            ObjKary = Nothing
        End Try
    End Sub
End Class