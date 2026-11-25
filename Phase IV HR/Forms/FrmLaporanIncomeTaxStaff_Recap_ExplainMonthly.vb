Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports System.Drawing

Public Class FrmLaporanIncomeTaxStaff_Recap_ExplainMonthly
    Public mIsRutin As Integer = 0
    Public mPrimaryTransaksi As Double = -1
    Public mTaxComponent As Core_Module.ClsMasterSalaryFormula_item.EnumTax2
    Public mDefPrec As Integer = 0

    Dim mBackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim mObjCD As New Core_Module.ClsCekData

    Private Sub FrmLaporanIncomeTaxStaff_Recap_ExplainMonthly_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DTA As New DseMonthlyTaxTableAdapters.tbltrans_salary_item1TableAdapter
        Dim TBL As New DseMonthlyTax.tbltrans_salary_item1DataTable
        Dim row As DseMonthlyTax.tbltrans_salary_item1Row
        Dim rowDGV As DataGridViewRow
        Dim Total As Decimal = 0
        Dim TipeKomponen As Core_Module.FrmMasterSalaryFormula_Component.EnumType = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Income
        Dim DTAOther As New DseMonthlyTaxTableAdapters.tbltrans_salary_item_otherTableAdapter
        Dim TBLOther As New DseMonthlyTax.tbltrans_salary_item_otherDataTable
        Dim rowOther As DseMonthlyTax.tbltrans_salary_item_otherRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAOther.Connection.ConnectionString = Core_Module.GetConstring

            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.mBackColorAll
            Me.dgvItem.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.Rows.Clear()
            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor

            DTA.FillBy1(TBL, CInt(Me.mTaxComponent), Me.mIsRutin, Me.mPrimaryTransaksi)

            'tambahkan item other
            DTAOther.Fill(TBLOther, CInt(Me.mTaxComponent), Me.mIsRutin, Me.mPrimaryTransaksi)
            While TBLOther.Rows.Count > 0
                rowOther = TBLOther.Rows(0)

                row = TBL.Newtbltrans_salary_item1Row
                With rowOther
                    row.Amount = .Amount
                    row.Description = .Description
                    row.Identifier_Salary_Component = .Identifier_Salary_Component
                    row.Is_Rutin = .Is_Rutin
                    row.Primary = .Primary
                    row.Primary_Transaksi = .Primary_Transaksi
                    row.Type = .Type
                End With
                TBL.Rows.Add(row)

                TBLOther.Rows.Remove(rowOther)
            End While

            For Each row In TBL.Rows
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                rowDGV.Cells(Me.ColDesc.Index).Value = row.Description
                If rowDGV.Cells(Me.ColDesc.Index).Value = "" And row.Identifier_Salary_Component = "" Then
                    rowDGV.Cells(Me.ColDesc.Index).Value = "Basic Salary"
                End If
                If row.Is_Rutin = 1 Then
                    rowDGV.Cells(Me.ColIsRutin.Index).Value = Boolean.TrueString
                Else
                    rowDGV.Cells(Me.ColIsRutin.Index).Value = Boolean.FalseString
                End If
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count

                TipeKomponen = row.Type
                If Me.mTaxComponent = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.gaji_pensiun Or _
                Me.mTaxComponent = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.honorarium Or _
                Me.mTaxComponent = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.penerimaan_dalam_bentuk_natura Or _
                Me.mTaxComponent = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.Tanteiem_bonus_gratifikasi_jasa_produksi_thr Or _
                Me.mTaxComponent = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.premi_asuransi Or _
                Me.mTaxComponent = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.tunjangan_lainnya Then
                    If TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Income Or _
                    TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.CompanyExpense Then
                        rowDGV.Cells(Me.ColAmount.Index).Value = Me.mObjCD.DigitGrouping(row.Amount, Me.mDefPrec)
                        Total += row.Amount
                    ElseIf TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Deduction Or _
                    TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.EmployeeExpense Then
                        rowDGV.Cells(Me.ColAmount.Index).Value = Me.mObjCD.DigitGrouping(row.Amount * -1, Me.mDefPrec)
                        Total -= row.Amount
                    Else
                        MsgBox("Logic error: component tipe not defined. Please contact your provider if you see this message.", MsgBoxStyle.Information, "Information")
                    End If
                ElseIf Me.mTaxComponent = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.iuran_pensiun Then
                    If TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Income Or _
                    TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.EmployeeExpense Then
                        rowDGV.Cells(Me.ColAmount.Index).Value = Me.mObjCD.DigitGrouping(row.Amount, Me.mDefPrec)
                        Total += row.Amount
                    ElseIf TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Deduction Or _
                    TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.CompanyExpense Then
                        rowDGV.Cells(Me.ColAmount.Index).Value = Me.mObjCD.DigitGrouping(row.Amount * -1, Me.mDefPrec)
                        Total -= row.Amount
                    Else
                        MsgBox("Logic error: component tipe not defined. Please contact your provider if you see this message.", MsgBoxStyle.Information, "Information")
                    End If
                Else
                    MsgBox("Logic error: tax component not defined. Please contact your provider if you see this message.", MsgBoxStyle.Information, "Information")
                End If
            Next

            If Me.dgvItem.RowCount > 1 Then
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                rowDGV.Cells(Me.ColDesc.Index).Value = "Total"
                rowDGV.Cells(Me.ColAmount.Index).Value = Me.mObjCD.DigitGrouping(Total, Me.mDefPrec)
            End If

            Me.dgvItem.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAOther.Dispose()
            TBLOther.Dispose()
        End Try
    End Sub
End Class