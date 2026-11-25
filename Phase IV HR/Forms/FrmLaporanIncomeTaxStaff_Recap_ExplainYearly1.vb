Option Explicit On
Option Strict Off

Imports System.Windows.Forms

Public Class FrmLaporanIncomeTaxStaff_Recap_ExplainYearly1
    Public mPrimaryEmployee As Double = -1
    Public mStartPeriod As Integer = 0
    Public mEndPeriod As Integer = 0
    Public mIsRoutineOnly As Boolean = False
    Public mComponentType As Core_Module.ClsMasterSalaryFormula_item.EnumTax2
    Public mRefPeriod As Date = Today
    Public mDefPrec As Integer = 0

    Dim ObjCD As New Core_Module.ClsCekData
    Dim TBL As New DseMonthlyTax.tbltrans_salary_item5DataTable
    Dim row As DseMonthlyTax.tbltrans_salary_item5Row
    Dim DTADetail As New DseMonthlyTaxTableAdapters.tbltrans_salary_item3TableAdapter
    Dim TBLDetail As New DseMonthlyTax.tbltrans_salary_item3DataTable
    Dim rowDetail As DseMonthlyTax.tbltrans_salary_item3Row
    Dim LoadDetail As Boolean = False
    Dim DTADetailOther As New DseMonthlyTaxTableAdapters.tbltrans_salary_item_other1TableAdapter
    Dim TBLDetailOther As New DseMonthlyTax.tbltrans_salary_item_other1DataTable
    Dim rowDetailOther As DseMonthlyTax.tbltrans_salary_item_other1Row

    Private Sub FrmLaporanIncomeTaxStaff_Recap_ExplainYearly1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim StartDate, EndDate, DummyDate As Date
        Dim Idx As Integer = 0
        Dim rowDGV As DataGridViewRow
        Dim Ketemu As Boolean = False
        Dim TipeKomponen As Core_Module.FrmMasterSalaryFormula_Component.EnumType = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Income
        Dim Jml As Decimal = 0

        Try
            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.ClearSelection()
            DTADetail.Connection.ConnectionString = Core_Module.GetConstring
            DTADetailOther.Connection.ConnectionString = Core_Module.GetConstring

            TBLDetail.DefaultView.Sort = "[Tanggal Transaksi], Description"
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(CDate("#" + Me.mStartPeriod.ToString + "/1/" + Me.mRefPeriod.Year.ToString + "#"))
            EndDate = Me.ObjCD.GetLasDateOfTheMonth(CDate("#" + Me.mEndPeriod.ToString + "/1/" + Me.mRefPeriod.Year.ToString + "#"))
            If EndDate > Me.ObjCD.GetLasDateOfTheMonth(Me.mRefPeriod) Then
                EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.mRefPeriod)
            End If

            'load data detail
            Me.DTADetail.FillBy_TaxKomponen(Me.TBLDetail, EndDate, StartDate, Me.mPrimaryEmployee, CInt(mComponentType))

            'tambahkan dengan item other
            Me.DTADetailOther.Fill(Me.TBLDetailOther, EndDate, StartDate, Me.mPrimaryEmployee, CInt(mComponentType))
            While Me.TBLDetailOther.Rows.Count > 0
                rowDetailOther = Me.TBLDetailOther.Rows(0)

                rowDetail = TBLDetail.Newtbltrans_salary_item3Row
                With rowDetailOther
                    rowDetail.Amount = .Amount
                    rowDetail.Description = .Description
                    rowDetail.Identifier_Salary_Component = .Identifier_Salary_Component
                    rowDetail.Is_Rutin = .Is_Rutin
                    rowDetail.No_Transaksi = .No_Transaksi
                    rowDetail.Primary = .Primary
                    rowDetail.Primary_Transaksi = .Primary_Transaksi
                    rowDetail.Tanggal_Transaksi = .Tanggal_Transaksi
                    rowDetail.Type = .Type
                End With
                TBLDetail.Rows.Add(rowDetail)

                Me.TBLDetailOther.Rows.Remove(rowDetailOther)
            End While

            'filter jika hanya rutin saja
            If Me.mIsRoutineOnly = True Then
                While Idx <= Me.TBLDetail.Rows.Count - 1
                    Me.rowDetail = Me.TBLDetail.Rows(Idx)
                    If Me.rowDetail.Is_Rutin = 0 Then
                        Me.TBLDetail.Rows.Remove(Me.rowDetail)
                    Else
                        Idx += 1
                    End If
                End While
            End If

            'tentukan positif, negatif
            For Each Me.rowDetail In Me.TBLDetail.Rows
                TipeKomponen = Me.rowDetail.Type
                If Me.mComponentType = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.gaji_pensiun Or _
                   Me.mComponentType = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.honorarium Or _
                   Me.mComponentType = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.penerimaan_dalam_bentuk_natura Or _
                   Me.mComponentType = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.Tanteiem_bonus_gratifikasi_jasa_produksi_thr Or _
                   Me.mComponentType = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.premi_asuransi Or _
                   Me.mComponentType = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.tunjangan_lainnya Then
                    If TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Income Or _
                    TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.CompanyExpense Then
                        Me.rowDetail.Amount = Me.rowDetail.Amount
                    ElseIf TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Deduction Or _
                    TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.EmployeeExpense Then
                        Me.rowDetail.Amount = Me.rowDetail.Amount * -1
                    Else
                        MsgBox("Logic error: component tipe not defined. Please contact your provider if you see this message.", MsgBoxStyle.Information, "Information")
                    End If
                ElseIf Me.mComponentType = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.iuran_pensiun Then
                    If TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Income Or _
                    TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.EmployeeExpense Then
                        Me.rowDetail.Amount = Me.rowDetail.Amount
                    ElseIf TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Deduction Or _
                    TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.CompanyExpense Then
                        Me.rowDetail.Amount = Me.rowDetail.Amount * -1
                    Else
                        MsgBox("Logic error: component tipe not defined. Please contact your provider if you see this message.", MsgBoxStyle.Information, "Information")
                    End If
                Else
                    MsgBox("Logic error: tax component not defined. Please contact your provider if you see this message.", MsgBoxStyle.Information, "Information")
                End If
            Next

            'isikan dgv
            DummyDate = StartDate
            While DummyDate <= EndDate
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                rowDGV.Cells(Me.ColItem_Period.Index).Value = DummyDate
                rowDGV.Cells(Me.ColItem_Amount.Index).Value = 0
                DummyDate = Me.ObjCD.GetFirstDateofTheMonth(DummyDate + TimeSpan.FromDays(32))
            End While

            'isikan dgv rekap
            Idx = 0
            While Idx <= Me.TBLDetail.Rows.Count - 1
                Me.rowDetail = Me.TBLDetail.Rows(Idx)
                Ketemu = False
                For Each rowDGV In Me.dgvItem.Rows
                    If Format(CDate(rowDGV.Cells(Me.ColItem_Period.Index).Value), "MMM yyyy") = Format(Me.rowDetail.Tanggal_Transaksi, "MMM yyyy") Then
                        rowDGV.Cells(Me.ColItem_Amount.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColItem_Amount.Index).Value) + Me.rowDetail.Amount, mDefPrec)
                        Jml += Me.rowDetail.Amount
                        Ketemu = True
                        Exit For
                    End If
                Next

                If Ketemu = False Then
                    MsgBox("Logic error: header row not found. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                Idx += 1
            End While

            Me.ColItem_Period.DefaultCellStyle.Format = "MMM yyyy"
            Me.ColItem_Amount.DefaultCellStyle.Format = "N" + Me.mDefPrec.ToString
            Me.ColDetail_Amount.DefaultCellStyle.Format = Me.ColItem_Amount.DefaultCellStyle.Format
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.LoadDetail = True
            Me.dgvItem.ClearSelection()
            If Me.dgvItem.Rows.Count > 0 Then
                Me.dgvItem.Rows(0).Selected = True
            End If

            Me.Label1.Text = "Start Period: " + Format(Me.dgvItem.Item(Me.ColItem_Period.Index, 0).Value, "MMM yyyy")
            Me.Label2.Text = "End Period: " + Format(Me.dgvItem.Item(Me.ColItem_Period.Index, Me.dgvItem.RowCount - 1).Value, "MMM yyyy")
            Me.Label3.Text = "Total Amount: " + Me.ObjCD.DigitGrouping(Jml, Me.mDefPrec)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL.Dispose()
            DTADetail.Dispose()
            TBLDetail.Dispose()
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Dim Idx As Integer = -1
        Dim RefPeriod1 As Date = Today
        Dim TglTrans As Date = Today
        Dim rowDGV As DataGridViewRow
        Dim Jml As Decimal = 0

        Try
            If LoadDetail = True Then
                Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
                If Me.dgvItem.SelectedCells.Count > 0 Then
                    Idx = Me.dgvItem.SelectedCells(0).RowIndex
                    RefPeriod1 = Me.ObjCD.GetFirstDateofTheMonth(Me.dgvItem.Item(Me.ColItem_Period.Index, Idx).Value)

                    Me.dgvItemDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Me.dgvItemDetail.Rows.Clear()

                    For Each DV As DataRowView In Me.TBLDetail.DefaultView
                        Me.rowDetail = DV.Row
                        TglTrans = Me.ObjCD.GetFirstDateofTheMonth(Me.rowDetail.Tanggal_Transaksi)
                        If TglTrans.Date > RefPeriod1.Date Then
                            Exit For
                        ElseIf TglTrans.Date = RefPeriod1.Date Then
                            If Me.rowDetail.Amount <> 0 Then
                                Me.dgvItemDetail.Rows.Add()
                                rowDGV = Me.dgvItemDetail.Rows(Me.dgvItemDetail.Rows.Count - 1)
                                rowDGV.Cells(Me.ColDetail_Amount.Index).Value = Me.rowDetail.Amount
                                Jml += Me.rowDetail.Amount
                                rowDGV.Cells(Me.ColDetail_Desc.Index).Value = Me.rowDetail.Description
                                If rowDGV.Cells(Me.ColDetail_Desc.Index).Value = "" Then
                                    rowDGV.Cells(Me.ColDetail_Desc.Index).Value = "Basic Salary"
                                End If

                                If Me.rowDetail.Is_Rutin = 1 Then
                                    rowDGV.Cells(Me.ColDetail_IsRutin.Index).Value = Boolean.TrueString
                                Else
                                    rowDGV.Cells(Me.ColDetail_IsRutin.Index).Value = Boolean.FalseString
                                End If
                                rowDGV.Cells(Me.ColDetail_No.Index).Value = Me.dgvItemDetail.Rows.Count
                            End If
                        End If
                    Next

                    If Me.dgvItemDetail.RowCount > 1 Then
                        Me.dgvItemDetail.Rows.Add()
                        rowDGV = Me.dgvItemDetail.Rows(Me.dgvItemDetail.Rows.Count - 1)
                        rowDGV.Cells(Me.ColDetail_Desc.Index).Value = "Total"
                        rowDGV.Cells(Me.ColDetail_Amount.Index).Value = Jml
                    End If
                    Me.dgvItemDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class