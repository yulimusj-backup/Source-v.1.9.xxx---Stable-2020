Option Explicit On
Option Strict Off

Public Class FrmTransaksiMonthlyClosing2
    Public UserCode As String = ""
    Public BAckColorAll As Color = BackColor

    Dim TblProductQty As New Inventory_Module.DseLaporanStockDetail2.TblCetakDataTable

    Dim ObjClosing As New ClsTransaksiMonthlyClosing
    Dim DefaultPrecision As Integer = 0
    Dim Buffer As String = ""
    Dim ThreadBuffer As String = ""
    Dim baru As Boolean = True
    Dim ObjCLosingItem As New ClsTransaksiMonthlyClosingItem
    Public ObjProduct As New Core_Module.clsMasterProduct
    Dim objCurrency As New Core_Module.clsMasterCurrency
    Dim ObjUser As New Core_Module.ClsUser
    Dim TBL As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable
    Dim TBLUnitPrice As New DseTransaksiMonthlyClosing.DataTable9DataTable
    Dim TBLItemBefore As New DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable
    Dim TBLItemAfter As New DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable
    Dim LoadClosing As Boolean = False
    Dim DTA_CL As New DseTransaksiMonthlyClosingTableAdapters.DataTable25TableAdapter
    Dim TBL_CL As New DseTransaksiMonthlyClosing.DataTable25DataTable
    Dim ObjLicense As New Core_Module.ClsLicense
    Dim ObjAccClosing As New AccModule.ClsTransaksiClosing

    Dim TBLBeginning As New DseTransaksiMonthlyClosing.tbltrans_purchase_closing_item1DataTable
    Dim DTABeginning As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closing_item1TableAdapter

    Dim ClosingPeriodeBuffer As Date

    Delegate Sub AddNewRow(ByVal BeginQty As Decimal, ByVal BeginAmount As Decimal, _
    ByVal vPrimaryProduct As Double, ByVal vCodeNo As String, ByVal Description As String, ByVal vPrimaryUnit As Double, _
    ByVal UnitCode As String, ByVal OutQty As Decimal, ByVal EndQty As Decimal, ByVal InQty As Decimal, ByVal InAmount As Decimal)
    Delegate Sub AddInAmount(ByVal vPrimaryCodeNo As Double, ByVal vAdditionalAmount As Decimal)
    Delegate Sub SetOutAmount(ByVal rowDGV As DataGridViewRow, ByVal OutAmount As Decimal)
    Delegate Sub SetEndAmount(ByVal rowDGV As DataGridViewRow, ByVal EndAmount As Decimal)

    Public ObjCD As New Core_Module.ClsCekData

    Dim DTAClosing1 As New DseTransaksiMonthlyClosingTableAdapters.DataTable21TableAdapter
    Dim TBLClosing1 As New DseTransaksiMonthlyClosing.DataTable21DataTable
    Dim Dum_ClosingPeriode As Date = #1/2/1950#

    Dim vTBL1 As New Core_Module.dseMasterProduct.tblmaster_product1DataTable
    Dim vDTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_product1TableAdapter

    'Dim DTABundling As New DseTransaksiMonthlyClosingTableAdapters.DataTable12TableAdapter
    'Dim TBLBundling As New DseTransaksiMonthlyClosing.DataTable12DataTable

    Dim ListPrimaryCodeNo_Bundling As New List(Of Double)
    Dim UpdateJurnal As Boolean = False

    Public PrimaryTemporary As Double = -1
    Public saveTemporaryAktif As Boolean = False

    Public TblA As New DseMonthlyClosingNew.Adjustment_CorrectionDataTable
    Public rowA As DseMonthlyClosingNew.Adjustment_CorrectionRow
    Public TblA2 As New DseMonthlyClosingNew.Adjustment_Correction2DataTable
    Public rowA2 As DseMonthlyClosingNew.Adjustment_Correction2Row

    Dim TblQty3 As New DseMonthlyClosingNew.tblmaster_productqty1DataTable
    Dim rowQty3 As DseMonthlyClosingNew.tblmaster_productqty1Row

    Public Sub CreateJurnal(ByVal vPrimary As Double)
        Try
            If Me.ObjClosing.CariDetail(vPrimary, Me.dgvItem, Me.dgvItemAdjust, _
            Me.TBLItemDetail, Me.TBLBundling1, Me, Me.dgvJurnal, Me.TBLSettingItem, _
            Me.dgvItemCost_CostGroup, Me.vTBLCostGroup_Account, Me.vTBLCostGroup_Scope, Me.vTBLCostGroup_CodeNo, Me, Me.vTBLCostGroup_CodeNo2, Me.TblProductQty) _
            = True Then
                Me.SalinKeField()
                'jika yang ditampilkan adalah data pertama, maka hide kolom kolom yang tidak perlu
                If Me.ObjClosing.ApakahClosingPalingAwal(Me.TxtPrimaryClosing.Text) = True Then
                    Me.ColInAmount.Visible = False
                    Me.ColInQty.Visible = False
                    Me.ColBeginQty.Visible = False
                    Me.ColBeginAmount.Visible = False
                    Me.ColOutQty.Visible = False
                    Me.ColOutAmount.Visible = False
                    'Me.ColDetailCalculation.Visible = False
                Else
                    Me.ColInAmount.Visible = True
                    Me.ColInQty.Visible = True
                    Me.ColBeginQty.Visible = True
                    Me.ColBeginAmount.Visible = True
                    Me.ColOutQty.Visible = True
                    Me.ColOutAmount.Visible = True
                    'sementara
                    'Me.ColDetailCalculation.Visible = False
                End If

                'create jurnal
                Me.ObjClosing.TransferJurnal(vPrimary, Me.dtpPeriod.Value.Date, Me.ObjClosing.CreatedBy, Me.dgvJurnal, Me, -1)

                Me.KosongkanField()
                Me.KosongkanField_Overhead()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetAngkaDecimal(ByVal vPrimaryProduct As Double) As Integer
        Dim row As Core_Module.dseMasterProduct.tblmaster_product1Row
        Dim filter As String = ""

        Try
            filter = "primary =" + vPrimaryProduct.ToString
            Me.vTBL1.DefaultView.RowFilter = filter
            If Me.vTBL1.DefaultView.Count > 0 Then
                row = TryCast(Me.vTBL1.DefaultView(0).Row, Core_Module.dseMasterProduct.tblmaster_product1Row)
                Return row._Decimal
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.BringToFront()

                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjClosing.SimpanBaru(Me.dgvItem, Me.dgvItemAdjust, Me.UserCode, Me, Me.TBLItemDetail, Me.TBLBundling1, _
                            Me.dgvJurnal, Me.TBLSetting, Me.TBLSettingItem, Me.tblPool, _
                            Me.dgvItemCost_CostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, vTBLCostGroup_CodeNo, Me.vTBLCostGroup_CodeNo2, Me.TblProductQty, False) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjClosing.SimpanPerubahan(Me.dgvItem, Me.dgvItemAdjust, Me.UserCode, Me, Me.TBLItemDetail, Me.TBLBundling1, _
                            Me.dgvJurnal, Me.TBLSetting, Me.TBLSettingItem, Me.tblPool, _
                            Me.dgvItemCost_CostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, vTBLCostGroup_CodeNo, Me.vTBLCostGroup_CodeNo2, Me.TblProductQty) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                ElseIf hasil = MsgBoxResult.No Then
                    If Me.saveTemporaryAktif = True Then
                        DeleteTemporaryDataHeader()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Purchase_Module.SetCulture()
            Me.ColJournal_Remarks.Visible = False
            Me.Icon = My.Resources.desktop_icon_3_fw
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.dgvItem2.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.dgvItemAdjust.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.dgvJurnal.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.dgvItemCost_AllocationRecap1.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.dgvItemCost_AllocationRecap2.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.dgvItemCost_CostAccountAllocation.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.dgvItemCost_CostGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.dgvItemCost_DetailTransaction.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.dgvItemCost_Scope.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll

            Me.ColBeginAmount.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColBeginQty.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColInAmount.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColInQty.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColOutAmount.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColOutQty.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColEndAmount.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColEndQty.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            Me.DefaultPrecision = Me.objCurrency.GetDecimal_DefaultCurrency

            'format dgv journal
            Me.ColJournal_Dept.Visible = False
            Me.ColJournal_Job.Visible = False
            Me.SetMulticurrency()

            'untuk sementara
            'Me.ColDetailCalculation.Visible = False
            'Me.btnDelete.Anchor = AnchorStyles.Bottom
            'Me.btnFind.Anchor = AnchorStyles.Bottom
            'Me.btnNew.Anchor = AnchorStyles.Bottom
            'Me.btnEdit.Anchor = AnchorStyles.Bottom
            'Me.btnExit.Anchor = AnchorStyles.Bottom

            Me.GroupBox2.SendToBack()
            Me.LblProgress.Text = ""
            Me.LblStep.Text = ""
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Visible = True
            Me.DTASetting.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTAItemClosing.Connection.ConnectionString = Me.DTASetting.Connection.ConnectionString

            'MsgBox("Harus ambil setting 'hitung otomatis dahulu'", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        TxtRemark.BackColor = Me.BAckColorAll
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        TxtRemark.BackColor = Me.TxtPrimaryClosing.BackColor
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            'msgbox("belum")
            If Me.TxtPrimaryClosing.Text <> "" Then
                If CDbl(Me.TxtPrimaryClosing.Text) >= 0 Then
                    ObjClosing.Cetak()
                End If
            Else
                Me.btnFind.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.Enter
        Try
            Me.Buffer = Me.dtpDate.Value.Date.ToShortDateString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim AngkaDecimal As Integer = 0

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                AngkaDecimal = Me.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value)
                'digit grouping: 
                row.Cells(Me.ColBeginQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColBeginQty.Index).Value, _
                AngkaDecimal)
                row.Cells(Me.ColBeginAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColBeginAmount.Index).Value, Me.DefaultPrecision)

                row.Cells(Me.ColInQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColInQty.Index).Value, _
                AngkaDecimal)
                row.Cells(Me.ColInAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColInAmount.Index).Value, Me.DefaultPrecision)

                row.Cells(Me.ColOutQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColOutQty.Index).Value, _
                AngkaDecimal)
                row.Cells(Me.ColOutAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColOutAmount.Index).Value, Me.DefaultPrecision)

                row.Cells(Me.ColEndQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColEndQty.Index).Value, _
                AngkaDecimal)
                row.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColEndAmount.Index).Value, Me.DefaultPrecision)
            Next

            'untuk in adjustment dari awal prosesnya sudah ok.
            'For Each row In Me.dgvItemAdjust.Rows
            'row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQty.Index).Value)
            'row.Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColUnitPriceAdjust.Index).Value, Me.DefaultPrecision))
            'Next

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'Jika sudah ada transaksi closing sebelumnya maka tidak bisa diedit
            If Me.ObjClosing.GetJml > 1 Or (Me.ObjClosing.GetJml = 1 And Me.baru = True) Then
                'If Me.UserCode = "USER" Then
                'kalo pake kode "USER" maka bisa edit semuanya
                'If e.ColumnIndex = Me.ColEndQty.Index Or e.ColumnIndex = Me.ColEndAmount.Index Or _
                ' e.ColumnIndex = Me.ColRemarks.Index Or e.ColumnIndex = Me.ColInQty.Index Or _
                ' e.ColumnIndex = Me.ColInAmount.Index Or e.ColumnIndex = Me.ColOutQty.Index Or _
                'e.ColumnIndex = Me.ColOutAmount.Index Then
                'kalo sampe sini maka boleh edit
                'Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                'Exit Sub
                'Else
                'e.Cancel = True
                'Exit Sub
                'End If
                'Else
                'yang bisa ediedit adalah remarks
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    'kalo sampe sini maka boleh edit
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                    Exit Sub
                End If
                'End If
            ElseIf Me.ObjClosing.GetJml = 1 And Me.baru = False And e.ColumnIndex = Me.ColEndQty.Index Then
                e.Cancel = True
                Exit Sub
            ElseIf Me.ObjClosing.GetJml = 0 And e.ColumnIndex = Me.ColEndQty.Index Then
                e.Cancel = True
                Exit Sub
            End If

            'yang bisa diedit adalah:
            '1. End Qty
            '2. End Amount
            '3. Remarks
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = Me.ColEndQty.Index Or e.ColumnIndex = Me.ColEndAmount.Index Or _
                e.ColumnIndex = Me.ColRemarks.Index Then
                    'kalo sampe sini maka boleh edit
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                'hanya bisa edit pada saat detail aktif
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Try
            If e.ColumnIndex = Me.ColDetailCalculation.Index And e.RowIndex <> -1 Then
                Dim i As New FrmTransaksiMonthlyClosing_DetailCalc2
                i.vPrimaryProduct = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value
                i.TBLItemDetail = Me.TBLItemDetail
                i.Period = Me.dtpPeriod.Value
                i.UserCode = Me.UserCode
                i.ShowDialog()
                i.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex = Me.ColEndQty.Index Or e.ColumnIndex = Me.ColEndAmount.Index Or _
                    e.ColumnIndex = Me.ColInQty.Index Or _
                    e.ColumnIndex = Me.ColInAmount.Index Or e.ColumnIndex = Me.ColOutQty.Index Or _
                    e.ColumnIndex = Me.ColOutAmount.Index Then
                        If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            If e.ColumnIndex = Me.ColOutAmount.Index Then
                                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "-" Then
                                    Me.dgvItem.Item(Me.ColOutAmount.Index, e.RowIndex).Value = _
                                    Me.ObjCD.DigitGrouping(CDbl(Me.dgvItem.Item(Me.ColBeginAmount.Index, e.RowIndex).Value) + _
                                    CDbl(Me.dgvItem.Item(Me.ColInAmount.Index, e.RowIndex).Value) - _
                                    CDbl(Me.dgvItem.Item(Me.ColEndAmount.Index, e.RowIndex).Value))
                                Else
                                    If IsNumeric(Me.Buffer) = True Then
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Else
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = 0
                                    End If
                                End If
                            Else
                                If IsNumeric(Me.Buffer) = True Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Else
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = 0
                                End If
                            End If
                        Else
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar2() As Boolean
        Dim row As DataGridViewRow
        'Dim x As Double
        Dim CodeNo As String = ""
        Dim CodeNoBegin As String = ""
        Dim CodeNoOut As String = ""
        Dim CodeNoEnd As String = ""
        Dim CodeNoEnd1 As String = ""
        Dim CodeNo_QtyBegin As String = ""
        Dim CodeNo_QtyIn As String = ""
        Dim CodeNo_QtyOut As String = ""
        Dim CodeNo_QtyEnd As String = ""
        Dim CodeUnbalanceAmount As String = ""
        Dim CodeUnbalanceQty As String = ""
        Dim WrongEndAmountCalculation As String = ""
        Dim WrongOutAmountCalculation As String = ""
        Dim BeginAmount As String = ""
        Dim InAmount As String = ""
        Dim OutAmount As String = ""
        Dim EndAmount As String = ""

        Dim Pass As Boolean = True
        Dim Dummy_Adjust As String = ""
        Dim CostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod
        Dim Dummy As String = "" 

        Try 

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.CostingMethod)
            If Dummy = "" Then
                CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average
            Else
                CostingMethod = Dummy
            End If

            If Me.ObjClosing.GetJml = 0 Then
                Return Me.CekEndAmountSaja()
            ElseIf Me.baru = False And Me.ObjClosing.GetJml = 1 Then
                Return Me.CekEndAmountSaja()
            ElseIf Me.CekEndAmountSaja() = False Then
                Return False
            End If

            'jika bukan closing pertama
            For Each row In Me.dgvItem.Rows
                ''#sementara, hanay cek yang termasuk dalam kategory
                'TBLProd.DefaultView.RowFilter = "Primary = " + row.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString
                'If TBLProd.DefaultView.Count > 0 Then
                If Math.Abs((CDec(row.Cells(Me.ColBeginAmount.Index).Value) + CDec(row.Cells(Me.ColInAmount.Index).Value)) - _
                                (CDec(row.Cells(Me.ColOutAmount.Index).Value) + CDec(row.Cells(Me.ColEndAmount.Index).Value))) > 2 Then
                    If CodeUnbalanceAmount = "" Then
                        CodeUnbalanceAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeUnbalanceAmount = CodeUnbalanceAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColBeginQty.Index).Value) + CDec(row.Cells(Me.ColInQty.Index).Value) <> _
                CDec(row.Cells(Me.ColOutQty.Index).Value) + CDec(row.Cells(Me.ColEndQty.Index).Value) Then
                    If CodeUnbalanceQty = "" Then
                        CodeUnbalanceQty = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeUnbalanceQty = CodeUnbalanceQty + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If 

                If row.Cells(Me.ColInQty.Index).Value > 0 And row.Cells(Me.ColInAmount.Index).Value = 0 Then
                    If CodeNo = "" Then
                        CodeNo = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo = CodeNo + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If row.Cells(Me.ColBeginQty.Index).Value > 0 And row.Cells(Me.ColBeginAmount.Index).Value = 0 Then
                    If CodeNoBegin = "" Then
                        CodeNoBegin = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoBegin = CodeNoBegin + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If row.Cells(Me.ColOutQty.Index).Value > 0 And row.Cells(Me.ColOutAmount.Index).Value = 0 Then
                    If CodeNoOut = "" Then
                        CodeNoOut = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoOut = CodeNoOut + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If row.Cells(Me.ColEndQty.Index).Value > 0 And row.Cells(Me.ColEndAmount.Index).Value = 0 Then
                    If CodeNoEnd = "" Then
                        CodeNoEnd = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoEnd = CodeNoEnd + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                ElseIf row.Cells(Me.ColEndAmount.Index).Value < 0 Then
                    If CodeNoEnd1 = "" Then
                        CodeNoEnd1 = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoEnd1 = CodeNoEnd1 + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColBeginQty.Index).Value) < 0 Then
                    If CodeNo_QtyBegin = "" Then
                        CodeNo_QtyBegin = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyBegin = CodeNo_QtyBegin + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColInQty.Index).Value) < 0 Then
                    If CodeNo_QtyIn = "" Then
                        CodeNo_QtyIn = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyIn = CodeNo_QtyIn + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColOutQty.Index).Value) < 0 Then
                    If CodeNo_QtyOut = "" Then
                        CodeNo_QtyOut = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyOut = CodeNo_QtyOut + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColEndQty.Index).Value) < 0 Then
                    If CodeNo_QtyEnd = "" Then
                        CodeNo_QtyEnd = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyEnd = CodeNo_QtyEnd + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColBeginAmount.Index).Value) > 0 And CDec(row.Cells(Me.ColBeginQty.Index).Value) = 0 Then
                    If BeginAmount = "" Then
                        BeginAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        BeginAmount = BeginAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColInAmount.Index).Value) > 0 And CDec(row.Cells(Me.ColInQty.Index).Value) = 0 Then
                    If InAmount = "" Then
                        InAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        InAmount = InAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColOutAmount.Index).Value) > 0 And CDec(row.Cells(Me.ColOutQty.Index).Value) = 0 Then
                    If OutAmount = "" Then
                        OutAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        OutAmount = OutAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If 
            Next

            If CodeUnbalanceAmount <> "" Then
                MsgBox("Unbalance amount for: " + CodeUnbalanceAmount, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If CodeUnbalanceQty <> "" Then
                MsgBox("Unbalance qty for: " + CodeUnbalanceQty, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If WrongEndAmountCalculation <> "" Then
                MsgBox("Wrong end amount calculation for: " + WrongEndAmountCalculation, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If WrongOutAmountCalculation <> "" Then
                MsgBox("Wrong out amount calculation for: " + WrongOutAmountCalculation, MsgBoxStyle.Information, "Information")
                Pass = False
            End If


            If CodeNo_QtyBegin <> "" Then
                MsgBox("The following code no, have negative beginning qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyBegin, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If CodeNo_QtyIn <> "" Then
                MsgBox("The following code no, have negative in qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyIn, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If CodeNo_QtyOut <> "" Then
                MsgBox("The following code no, have negative out qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyOut, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If CodeNo_QtyEnd <> "" Then
                MsgBox("The following code no, have negative end qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyEnd, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If BeginAmount <> "" Then
                If MsgBox("The following code no, have positive beginning amount, but zero beginning qty. Continue save anyway?" + _
                vbCrLf + BeginAmount, MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Pass = False
                End If
            End If
            If InAmount <> "" Then
                MsgBox("The following code no, have positive in amount, but zero in qty. Please make sure you are aware of this." + _
                vbCrLf + InAmount, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If OutAmount <> "" Then
                MsgBox("The following code no, have positive out amount, but zero out qty. Please make sure you are aware of this." + _
                vbCrLf + OutAmount, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If EndAmount <> "" Then
                MsgBox("The following code no, have positive end amount, but zero end qty. Please make sure you are aware of this." + _
                vbCrLf + EndAmount, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If CodeNoBegin <> "" Then
                MsgBox("The following code no, have positive beginning qty but zero beginning amount, please make sure you are aware of this." + _
                vbCrLf + CodeNoBegin, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If CodeNo <> "" Then
                MsgBox("The following code no, have positive in qty but zero in amount, please make sure you are aware of this." + _
                vbCrLf + CodeNo, MsgBoxStyle.Information, "Information")
                Return False
                Pass = False
            Else
                If CodeNoOut <> "" Then
                    MsgBox("The following code no, have positive out qty but zero out amount, please make sure you are aware of this." + _
                    vbCrLf + CodeNoOut, MsgBoxStyle.Information, "Information")
                    Return False
                    Pass = False
                End If

                If CodeNoEnd <> "" Then
                    MsgBox("The following code no, have positive out qty but zero end amount, please make sure you are aware of this." + _
                    vbCrLf + CodeNoEnd, MsgBoxStyle.Information, "Information")
                    Return False
                    Pass = False
                End If
                If CodeNoEnd1 <> "" Then
                    If MsgBox("The following code no, negative end amount, please make sure you are aware of this." + _
                    vbCrLf + CodeNoEnd1 + vbCrLf + "Continue save anyway?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                        Return False
                        Pass = False
                    End If
                End If
            End If

            For Each row In Me.dgvItemAdjust.Rows
                If CDec(row.Cells(Me.ColUnitPriceAdjust.Index).Value) = 0 Then
                    If Dummy_Adjust = "" Then
                        Dummy_Adjust = "Line number: " + row.Cells(Me.ColNoAdjust.Index).Value.ToString + " - Code No: " + row.Cells(Me.ColCodeNoAdjust.Index).Value.ToString
                    Else
                        Dummy_Adjust = Dummy_Adjust + vbCrLf + "Line number: " + row.Cells(Me.ColNoAdjust.Index).Value.ToString + " - Code No: " + row.Cells(Me.ColCodeNoAdjust.Index).Value.ToString
                    End If
                End If
            Next

            If Dummy_Adjust <> "" Then
                If MsgBox("The following adjustment have zero unit price:" + vbCrLf _
                + Dummy_Adjust, MsgBoxStyle.Information, "Confirm") <> MsgBoxResult.Yes Then
                    Pass = False
                End If
            End If

            'cek qty beginning, qty in, qty out, qty end, harus lebih besar dari nol, kalo ga salah jadinya
            If Pass = False Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim AmountDebet, AmountCredit As Decimal
        Dim rowDGV As DataGridViewRow
        Dim DTAAOP As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_aopTableAdapter
        Dim TBLAOP As New DseTransaksiMonthlyClosing.tbltrans_aopDataTable
        Dim rowAOP As DseTransaksiMonthlyClosing.tbltrans_aopRow
        Dim DTAPR As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchasereturn2TableAdapter
        Dim TBLPR As New DseTransaksiMonthlyClosing.tbltrans_purchasereturn2DataTable
        Dim rowPR As DseTransaksiMonthlyClosing.tbltrans_purchasereturn2Row
        Dim DTAInv As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TBLInv As New DseTransaksiMonthlyClosing.tbltrans_salesinvoiceDataTable
        Dim rowInv As DseTransaksiMonthlyClosing.tbltrans_salesinvoiceRow
        Dim DTASR As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_salesreturn1TableAdapter
        Dim TBLSR As New DseTransaksiMonthlyClosing.tbltrans_salesreturn1DataTable
        Dim rowSR As DseTransaksiMonthlyClosing.tbltrans_salesreturn1Row
        Dim Dummy1 As New List(Of String)
        Dim Dummy2 As New List(Of String)
        Dim Dummy As String = ""
        Dim ij As Integer = 0
        Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
        Dim TipeSR As Sales_Module.ClsTransaksiSalesReturn.EnumTipeSR = Sales_Module.ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim StartTime, EndTIme As Date
        Dim ObjClosingJur As New AccModule.ClsTransaksiClosing

        Try
            If Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date) <= ObjClosingJur.GetClosingDate.Date Then
                MsgBox("You cannot save this data because " + Format(Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), "dd MMM yyyy") + " <= journal lock date" _
                    , MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek apakah perlu hitung ulang
            If Me.RecalculateNeeded = True Then
                MsgBox("Please recalculate data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("Transaction number")
            KoleksiTextBox.Add(Me.TxtClosingNumber)

            If ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                If MsgBox("Saved data with no item?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                    'MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.ObjClosing.TimeStamp <> Me.ObjClosing.GetTimestamp(Me.ObjClosing.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjClosing.CariDetail(Me.ObjClosing.Primary, Me.dgvItem, Me.dgvItemAdjust, _
                        Me.TBLItemDetail, Me.TBLBundling1, Me, Me.dgvJurnal, Me.TBLSettingItem, _
                        Me.dgvItemCost_CostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, vTBLCostGroup_CodeNo, Me, Me.vTBLCostGroup_CodeNo2, Me.TblProductQty)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'buat ulang jurnal
            If Me.UpdateJurnal = True Then
                'Me.UpdateKeDGVItemUtama()
                Me.BuatJurnal(False)
                Me.UpdateJurnal = False
            End If
            If Me.ApakahDataBenar2 = False Then
                Return False
            End If

            For Each rowDGV In Me.dgvJurnal.Rows
                AmountDebet += CDec(rowDGV.Cells(Me.ColJournal_Debet.Index).Value)
                AmountCredit += CDec(rowDGV.Cells(Me.ColJournal_Kredit.Index).Value)
            Next
            If AmountDebet <> AmountCredit Then
                MsgBox("Unbalance account amount", MsgBoxStyle.Information, "Information")

                If Me.UserCode = "USER" Then
                    If MsgBox("Continue save anyway?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            End If

            'StartTime = Now
            'If Me.CekKonsistensi1 = False Then
            '    Return False
            'End If
            'EndTIme = Now
            'MsgBox("cek konsistensi1 " + Me.ObjCD.DigitGrouping((EndTIme - StartTime).TotalMinutes, 1) + "minutes")

            DTAAOP.Fill(TBLAOP, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value))
            Dummy1.Clear()
            Dummy2.Clear()
            For Each rowAOP In TBLAOP.Rows
                If rowAOP.Primary_Currency = 1 Then
                    'pastikan pada bulan ybs tidak ada aop IDR yang memiliki exchange rate <> 1
                    If rowAOP.Tax_Exchange_Rate <> 1 Then
                        Dummy1.Add(rowAOP.No_Transaksi)
                    End If
                ElseIf rowAOP.Primary_Currency > 0 Then
                    'pastikan pada bulan ybs tidak ada aop non IDR yang memiliki exchange rate = 1
                    If rowAOP.Tax_Exchange_Rate = 1 Then
                        Dummy2.Add(rowAOP.No_Transaksi)
                    End If
                End If
            Next
            If Dummy1.Count > 0 Then
                Dummy = ""
                ij = 1
                For Each NoTrans As String In Dummy1
                    If Dummy = "" Then
                        Dummy = ij.ToString + ". " + NoTrans
                    Else
                        Dummy = Dummy + vbCrLf + ij.ToString + ". " + NoTrans
                    End If
                    ij += 1
                Next
                MsgBox("The following aop have IDR currency dan exchange rate <> 1, please fix." + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Dummy2.Count > 0 Then
                Dummy = ""
                ij = 1
                For Each NoTrans As String In Dummy2
                    If Dummy = "" Then
                        Dummy = ij.ToString + ". " + NoTrans
                    Else
                        Dummy = Dummy + vbCrLf + ij.ToString + ". " + NoTrans
                    End If
                    ij += 1
                Next
                MsgBox("The following aop have non-IDR currency dan exchange rate = 1, please fix." + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan pada bulan ybs tidak ada pr IDR yang memiliki exchange rate <> 1
            DTAPR.Fill(TBLPR, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value), CInt(TipePR))
            Dummy1.Clear()
            Dummy2.Clear()
            For Each rowPR In TBLPR.Rows
                If rowPR.Primary_Currency = 1 Then
                    'pastikan pada bulan ybs tidak ada pr IDR yang memiliki exchange rate <> 1
                    If rowPR.Tax_Exchange_Rate <> 1 Then
                        Dummy1.Add(rowPR.No_Transaksi)
                    End If
                ElseIf rowPR.Primary_Currency > 0 Then
                    'pastikan pada bulan ybs tidak ada pr non IDR yang memiliki exchange rate = 1
                    If rowPR.Tax_Exchange_Rate = 1 Then
                        Dummy2.Add(rowPR.No_Transaksi)
                    End If
                End If
            Next
            If Dummy1.Count > 0 Then
                Dummy = ""
                ij = 1
                For Each NoTrans As String In Dummy1
                    If Dummy = "" Then
                        Dummy = ij.ToString + ". " + NoTrans
                    Else
                        Dummy = Dummy + vbCrLf + ij.ToString + ". " + NoTrans
                    End If
                    ij += 1
                Next
                MsgBox("The following purchase return have IDR currency dan exchange rate <> 1, please fix." + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Dummy2.Count > 0 Then
                Dummy = ""
                ij = 1
                For Each NoTrans As String In Dummy2
                    If Dummy = "" Then
                        Dummy = ij.ToString + ". " + NoTrans
                    Else
                        Dummy = Dummy + vbCrLf + ij.ToString + ". " + NoTrans
                    End If
                    ij += 1
                Next
                MsgBox("The following purchase return have non-IDR currency dan exchange rate = 1, please fix." + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan pada bulan ybs tidak ada inv non IDR yang memiliki exchange rate = 1
            DTAInv.Fill(TBLInv, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value))
            Dummy1.Clear()
            Dummy2.Clear()
            For Each rowInv In TBLInv.Rows
                If rowInv.Primary_Currency = 1 Then
                    'pastikan pada bulan ybs tidak ada aop inv yang memiliki exchange rate <> 1
                    If rowInv.Tax_Exchange_Rate <> 1 Then
                        Dummy1.Add(rowInv.No_Transaksi)
                    End If
                ElseIf rowInv.Primary_Currency > 0 Then
                    'pastikan pada bulan ybs tidak ada inv non IDR yang memiliki exchange rate = 1
                    If rowInv.Tax_Exchange_Rate = 1 Then
                        Dummy2.Add(rowInv.No_Transaksi)
                    End If
                End If
            Next
            If Dummy1.Count > 0 Then
                Dummy = ""
                ij = 1
                For Each NoTrans As String In Dummy1
                    If Dummy = "" Then
                        Dummy = ij.ToString + ". " + NoTrans
                    Else
                        Dummy = Dummy + vbCrLf + ij.ToString + ". " + NoTrans
                    End If
                    ij += 1
                Next
                MsgBox("The following invoice have IDR currency dan exchange rate <> 1, please fix." + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Dummy2.Count > 0 Then
                Dummy = ""
                ij = 1
                For Each NoTrans As String In Dummy2
                    If Dummy = "" Then
                        Dummy = ij.ToString + ". " + NoTrans
                    Else
                        Dummy = Dummy + vbCrLf + ij.ToString + ". " + NoTrans
                    End If
                    ij += 1
                Next
                MsgBox("The following invoice have non-IDR currency dan exchange rate = 1, please fix." + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek sales return
            DTASR.Fill(TBLSR, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value), CInt(TipeSR))
            Dummy1.Clear()
            Dummy2.Clear()
            For Each rowSR In TBLSR.Rows
                If rowSR.Primary_Currency = 1 Then
                    'pastikan pada bulan ybs tidak ada sales return yang memiliki exchange rate <> 1
                    If rowSR.Tax_Exchange_Rate <> 1 Then
                        Dummy1.Add(rowSR.No_Transaksi)
                    End If
                ElseIf rowSR.Primary_Currency > 0 Then
                    'pastikan pada bulan ybs tidak ada sales return non IDR yang memiliki exchange rate = 1
                    If rowSR.Tax_Exchange_Rate = 1 Then
                        If rowSR.Tax_Exchange_Rate <> 1 Then
                            Dummy2.Add(rowSR.No_Transaksi)
                        End If
                    End If
                End If
            Next
            If Dummy1.Count > 0 Then
                Dummy = ""
                ij = 1
                For Each NoTrans As String In Dummy1
                    If Dummy = "" Then
                        Dummy = ij.ToString + ". " + NoTrans
                    Else
                        Dummy = Dummy + vbCrLf + ij.ToString + ". " + NoTrans
                    End If
                Next
                MsgBox("The following sales return have IDR currency dan exchange rate <> 1, please fix." + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Dummy2.Count > 0 Then
                Dummy = ""
                ij = 1
                For Each NoTrans As String In Dummy2
                    If Dummy = "" Then
                        Dummy = ij.ToString + ". " + NoTrans
                    Else
                        Dummy = Dummy + vbCrLf + ij.ToString + ". " + NoTrans
                    End If
                Next
                MsgBox("The following sales return have non-IDR currency dan exchange rate = 1, please fix." + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan semua product account di verifikasi
            If Me.ObjProduct.ApakahAda_ProductNotVerified(Me.dtpPeriod.Value) = True Then
                MsgBox("There are several products with unverified account. Please verify it first", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek overhead data
            If Me.ApakahDataBenar_Overhead() = False Then
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            DTAAOP.Dispose()
            TBLAOP.Dispose()
            DTAPR.Dispose()
            TBLPR.Dispose()
            DTAInv.Dispose()
            TBLInv.Dispose()
            DTASR.Dispose()
            TBLSR.Dispose()
            Dummy1 = Nothing
            Dummy2 = Nothing
            ObjClosingJur = Nothing
        End Try
    End Function

    Private Sub CekUnbalance()
        Dim AmountDebet As Decimal = 0
        Dim AmountCredit As Decimal = 0

        Try
            For Each rowDGV As DataGridViewRow In Me.dgvJurnal.Rows
                AmountDebet += CDec(rowDGV.Cells(Me.ColJournal_Debet.Index).Value)
                AmountCredit += CDec(rowDGV.Cells(Me.ColJournal_Kredit.Index).Value)
            Next
            If AmountDebet <> AmountCredit Then
                MsgBox("Unbalance account amount", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjClosing
                Me.TxtPrimaryClosing.Text = .Primary
                Me.TxtClosingNumber.Text = .NoTransaksi
                Me.TxtRemark.Text = .Remarks
                Me.dtpDate.Value = .ClosingDate
                Me.dtpPeriod.Value = .ClosingPeriod

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
                Me.TampilkanOverheadDetail()
            End With
            'Me.FormatDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.ObjClosing
                .NoTransaksi = Me.TxtClosingNumber.Text
                .Remarks = Me.TxtRemark.Text
                .ClosingDate = Me.dtpDate.Value
                .ClosingPeriod = Me.dtpPeriod.Value
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.ReadOnly = True
            Me.dgvItemAdjust.ReadOnly = True
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.btnFindRow.Enabled = True
            Me.NonAktifkanDetail_Overhead()

            Me.llbEditInAdjustment.Visible = False
            Me.Button3.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.TabControl1.SelectedTab = Me.TabPage1
            Me.BgwAktifkanDetail.RunWorkerAsync()
            'Me.AktifkanDetail2()
            'Me.btnOK.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail2()
        Try
            Me.ObjCD.UpdateThreading()
            Me.BgwAktifkanDetail.ReportProgress(1)

            If Me.baru = True Then
                Me.Calc_CurrentStep = 0
                Me.Calc_TotalStep = 26

                Me.GroupBox2.BringToFront()
                'alokasikan overhead
                Me.RecalculateOverhead()
                Me.BgwAktifkanDetail.ReportProgress(5)

                Me.LoadItemDetail2()
                Me.HitungKeDGVUtama1()
                If Me.TBL_Pending1.Rows.Count <> 0 Then
                    Exit Sub
                End If

                'jurnal hanya dibuat selain untuk begining balance
                Me.Calc_CurrentStep += 1
                Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : creating item data summary"
                Me.ProgressBar1.Value = 0
                Me.LblProgress.Text = ""
                'Me.LblProgress.Visible = True
                Me.ProgressBar1.SendToBack()

                Me.ThreadBuffer = "Start"
                Me.BgwAktifkanDetail.ReportProgress(2)

                While Me.ThreadBuffer = "Start"
                    Threading.Thread.Sleep(20000)
                End While
                Me.ProgressBar1.BringToFront()

                If Me.ObjClosing.ApakahClosingPalingAwal(Me.TxtPrimaryClosing.Text) = False Then
                    Me.BuatJurnal(False)
                End If
            End If
            BgwAktifkanDetail_Progress_3()

            If Me.TBL_Pending1.Rows.Count = 0 Then
                Me.RecalculateNeeded = False
            Else
                'Me.BgwAktifkanDetail_Progress_4()
                Me.dgvJurnal.Rows.Clear()
                Me.dgvItem.Rows.Clear()
                BgwAktifkanDetail_Progress_3()
                Me.RecalculateNeeded = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDataPengecekanQtyMinus()
        Dim DTAQty As New DseMonthlyClosingNewTableAdapters.tblmaster_productqty2TableAdapter
        Dim TBLQty As New DseMonthlyClosingNew.tblmaster_productqty2DataTable
        Dim rowQty As DseMonthlyClosingNew.tblmaster_productqty2Row
        Dim DTASC As New DseMonthlyClosingNewTableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim TBLSC As New DseMonthlyClosingNew.tbltransaksi_stock_cardDataTable
        Dim rowSC As DseMonthlyClosingNew.tbltransaksi_stock_cardRow
        Dim TglAcuan1 As Date = (Me.ObjCD.GetLasDateOfTheMonth(dtpPeriod.Value).AddDays(1)).Date
        Dim TBLProduct As New DseMonthlyClosingNew.tblmaster_product1DataTable
        Dim DTAProduct As New DseMonthlyClosingNewTableAdapters.tblmaster_product1TableAdapter
        Dim rowProduct As DseMonthlyClosingNew.tblmaster_product1Row

        Try
            DTAQty.Connection.ConnectionString = Core_Module.GetConstring
            DTASC.Connection.ConnectionString = DTAQty.Connection.ConnectionString
            DTAProduct.Connection.ConnectionString = DTAQty.Connection.ConnectionString

            DTAQty.Fill(TBLQty)
            DTASC.Fill(TBLSC, TglAcuan1)

            'isikan data qty saat ini
            For Each rowSC In TBLSC.Rows
                'balikan berdasarkan stock card   
                If rowSC.Qty_In <> rowSC.Qty_Out Then

                    TBLQty.DefaultView.RowFilter = Nothing
                    TBLQty.DefaultView.RowFilter = "[Primary Product] = " + rowSC.Primary_Product.ToString
                    If TBLQty.DefaultView.Count = 0 Then
                        'tambahkan baris baru
                        rowQty = TBLQty.Newtblmaster_productqty2Row
                        With rowQty
                            .Primary_Product = rowSC.Primary_Product
                            .Qty = rowSC.Qty_Out - rowSC.Qty_In
                        End With
                        TBLQty.Rows.Add(rowQty)
                    Else
                        'set qty-nya
                        rowQty = TBLQty.DefaultView(0).Row
                        rowQty.Qty += (rowSC.Qty_Out - rowSC.Qty_In)
                    End If
                End If
            Next

            'pindahkan ke tabel qty 3
            Me.TblQty3.Rows.Clear()
            DTAProduct.Fill(TBLProduct)
            For Each rowQty In TBLQty.Rows
                If rowQty.Qty < 0 Then
                    TBLProduct.DefaultView.RowFilter = "Primary = " + rowQty.Primary_Product.ToString
                    If TBLProduct.DefaultView.Count > 0 Then
                        rowProduct = TBLProduct.DefaultView(0).Row
                        Me.rowQty3 = Me.TblQty3.Newtblmaster_productqty1Row
                        With Me.rowQty3
                            .Angka_Decimal = rowProduct._Decimal
                            .Code_No = rowProduct.Product_Code
                            .Primary_Product = rowQty.Primary_Product
                            .Qty = rowQty.Qty
                            .Description = rowProduct.Description
                            .Unit_Code = rowProduct.Unit_Code
                        End With
                        Me.TblQty3.Rows.Add(Me.rowQty3)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAQty.Dispose()
            TBLQty.Dispose()
            DTASC.Dispose()
            TBLSC.Dispose()
            TBLProduct.Dispose()
            DTAProduct.Dispose()
        End Try
    End Sub

    'Private Sub IsikanDataPengecekanQtyMinus_Obs()
    '    Dim TBL_Awal As New DseMonthlyClosingNew.DataTable2DataTable
    '    Dim row_Awal As DseMonthlyClosingNew.DataTable2Row
    '    Dim DTAProduct As New DseMonthlyClosingNewTableAdapters.tblmaster_productTableAdapter
    '    Dim TBLProduct As New DseMonthlyClosingNew.tblmaster_productDataTable
    '    Dim rowProduct As DseMonthlyClosingNew.tblmaster_productRow

    '    Dim rowDGV As DataGridViewRow
    '    Dim DTAQty As New DseMonthlyClosingNewTableAdapters.tblmaster_productqtyTableAdapter
    '    Dim TBLQty1 As New DseMonthlyClosingNew.tblmaster_productqtyDataTable
    '    Dim TblQty2 As New DseMonthlyClosingNew.tblmaster_productqtyDataTable
    '    Dim rowQty1, rowQty2 As DseMonthlyClosingNew.tblmaster_productqtyRow
    '    Dim DV As Data.DataRowView
    '    Dim DTASC As New DseMonthlyClosingNewTableAdapters.tbltransaksi_stock_cardTableAdapter
    '    Dim TBLSC As New DseMonthlyClosingNew.tbltransaksi_stock_cardDataTable
    '    Dim rowSC As DseMonthlyClosingNew.tbltransaksi_stock_cardRow
    '    Dim DTA_Transfer As New DseMonthlyClosingNewTableAdapters.tbltransaksi_stocktransfer_itemTableAdapter
    '    Dim TBL_Transfer As New DseMonthlyClosingNew.tbltransaksi_stocktransfer_itemDataTable
    '    Dim row_Transfer As DseMonthlyClosingNew.tbltransaksi_stocktransfer_itemRow
    '    Dim TglAcuan1 As Date = (Me.ObjCD.GetLasDateOfTheMonth(dtpPeriod.Value).AddDays(1)).Date
    '    Dim AdaMinus As Boolean = False
    '    Dim QtyMinus As Decimal = 0
    '    Dim AmountAllocated As Decimal = 0
    '    Dim QtyAllocated As Decimal = 0
    '    Dim DefPrec As Integer = Me.objCurrency.GetDecimal_DefaultCurrency
    '    Dim ProductCodeRef As String = ""
    '    Dim DTABegin As New DseMonthlyClosingNewTableAdapters.tbltrans_purchase_closing_itemTableAdapter
    '    Dim TblBegin As New DseMonthlyClosingNew.tbltrans_purchase_closing_itemDataTable
    '    Dim rowBegin As DseMonthlyClosingNew.tbltrans_purchase_closing_itemRow
    '    Dim objUnit As New Core_Module.ClsMasterUnit

    '    Try
    '        TBL_Awal.Rows.Clear()
    '        TBLQty1.Rows.Clear()
    '        TblQty3.Rows.Clear()
    '        TBLSC.Rows.Clear()
    '        TBL_Transfer.Rows.Clear()

    '        TBLProduct.Rows.Clear()
    '        Dim tgl As Date = Me.dtpPeriod.Value.AddMonths(-1)

    '        DTAProduct.Fill(TBLProduct)
    '        DTAQty.Fill(TBLQty1)
    '        DTASC.Fill(TBLSC, TglAcuan1)
    '        DTA_Transfer.Fill(TBL_Transfer, TglAcuan1)
    '        DTAWarehouse.Fill(TBLWarehouse)
    '        DTABegin.Fill(TblBegin, Me.ObjClosing.GetPrimary(tgl))

    '        For Each rowProduct In TBLProduct.Rows
    '            With rowProduct
    '                row_Awal = TBL_Awal.NewDataTable2Row
    '                With row_Awal
    '                    .Category_Code = rowProduct.Category_Code
    '                    .Code_No = rowProduct.Product_Code
    '                    .Desc = rowProduct.Description
    '                    .Primary_Code_No = rowProduct.Primary
    '                    .End_Amount = 0
    '                    .Qty_End = 0
    '                    .Unit = 0
    '                    TblBegin.DefaultView.RowFilter = Nothing
    '                    TblBegin.DefaultView.RowFilter = " [Primary Product] = " & .Primary_Code_No
    '                    If TblBegin.DefaultView.Count > 0 Then
    '                        For Each dvB As Data.DataRowView In TblBegin.DefaultView
    '                            rowBegin = dvB.Row
    '                            If IsNumeric(rowBegin.Ending_Qty) Then
    '                                .Qty_End = rowBegin.Ending_Qty
    '                                .End_Amount = rowBegin.Ending_Amount
    '                                .Unit = objUnit.GetUnitCode(rowBegin.Primary_Unit)
    '                            End If
    '                        Next
    '                    End If
    '                End With
    '                TBL_Awal.Rows.Add(row_Awal)
    '            End With
    '        Next


    '        'masukkan ke data tabel qty
    '        TblQty3.Rows.Clear()

    '        For Each row_Awal In TBL_Awal.Rows
    '            ProductCodeRef = row_Awal.Code_No
    '            TblQty2.Rows.Clear()

    '            TBLQty1.DefaultView.RowFilter = Nothing
    '            TBLQty1.DefaultView.RowFilter = "[Primary Product] = " + row_Awal.Primary_Code_No.ToString
    '            For Each DV In TBLQty1.DefaultView
    '                rowQty1 = DV.Row
    '                rowQty2 = TblQty2.Newtblmaster_productqtyRow
    '                For ij As Integer = 0 To TBLQty1.Columns.Count - 1
    '                    rowQty2.Item(ij) = rowQty1.Item(ij)
    '                Next
    '                TblQty2.Rows.Add(rowQty2)
    '            Next

    '            'balikkan berdasarkan stock card
    '            TBLSC.DefaultView.RowFilter = Nothing
    '            TBLSC.DefaultView.RowFilter = "[Primary Product] = " + row_Awal.Primary_Code_No.ToString
    '            For Each DV In TBLSC.DefaultView
    '                rowSC = DV.Row
    '                If rowSC.Qty_In <> rowSC.Qty_Out Then
    '                    TblQty2.DefaultView.RowFilter = Nothing
    '                    TblQty2.DefaultView.RowFilter = "[Primary Product] = " + rowSC.Primary_Product.ToString + " AND [Primary Storage Area] = " + rowSC.primary_location.ToString
    '                    If TblQty2.DefaultView.Count = 0 Then
    '                        'tambahkan baris baru
    '                        rowQty2 = TblQty2.Newtblmaster_productqtyRow
    '                        With rowQty2
    '                            .Primary_Product = rowSC.Primary_Product
    '                            .Primary_Storage_Area = rowSC.primary_location
    '                            .Qty = rowSC.Qty_Out - rowSC.Qty_In
    '                        End With
    '                        TblQty2.Rows.Add(rowQty2)
    '                    Else
    '                        'set qty-nya
    '                        rowQty2 = TblQty2.DefaultView(0).Row
    '                        rowQty2.Qty += (rowSC.Qty_Out - rowSC.Qty_In)
    '                    End If
    '                End If
    '            Next



    '            'balikkan berdasarkan location transfer
    '            TBL_Transfer.DefaultView.RowFilter = Nothing
    '            TBL_Transfer.DefaultView.RowFilter = "[Primary Product] = " + row_Awal.Primary_Code_No.ToString
    '            For Each DV In TBL_Transfer.DefaultView
    '                row_Transfer = DV.Row
    '                'kurangkan menggunakan primary location target
    '                TblQty2.DefaultView.RowFilter = Nothing
    '                TblQty2.DefaultView.RowFilter = "[Primary Product] = " + row_Transfer.Primary_Product.ToString + " AND [Primary Storage Area] = " + row_Transfer.primary_destination_area.ToString
    '                If TblQty2.DefaultView.Count = 0 Then
    '                    'tambahkan baris baru
    '                    rowQty2 = TblQty2.Newtblmaster_productqtyRow
    '                    With rowQty2
    '                        .Primary_Product = row_Transfer.Primary_Product
    '                        .Primary_Storage_Area = row_Transfer.primary_destination_area
    '                        .Qty = (row_Transfer.Qty * -1)
    '                    End With
    '                    TblQty2.Rows.Add(rowQty2)
    '                Else
    '                    'set qty-nya
    '                    rowQty2 = TblQty2.DefaultView(0).Row
    '                    rowQty2.Qty -= row_Transfer.Qty
    '                End If



    '                'tambahkan menggunakan primary location sumber
    '                TblQty2.DefaultView.RowFilter = Nothing
    '                TblQty2.DefaultView.RowFilter = "[Primary Product] = " + row_Transfer.Primary_Product.ToString + " AND [Primary Storage Area] = " + row_Transfer.Primary_Storage_Area.ToString
    '                If TblQty2.DefaultView.Count = 0 Then
    '                    'tambahkan baris baru
    '                    rowQty2 = TblQty2.Newtblmaster_productqtyRow
    '                    With rowQty2
    '                        .Primary_Product = row_Transfer.Primary_Product
    '                        .Primary_Storage_Area = row_Transfer.Primary_Storage_Area
    '                        .Qty = row_Transfer.Qty
    '                    End With
    '                    TblQty2.Rows.Add(rowQty2)
    '                Else
    '                    'set qty-nya
    '                    rowQty2 = TblQty2.DefaultView(0).Row
    '                    rowQty2.Qty += row_Transfer.Qty
    '                End If
    '            Next



    '            'bagi ke beberapa warehouse
    '            TblQty2.DefaultView.RowFilter = Nothing
    '            For Each rowQty2 In TblQty2.Rows
    '                TBLWarehouse.DefaultView.RowFilter = Nothing
    '                TBLWarehouse.DefaultView.RowFilter = "Primary = " + rowQty2.Primary_Storage_Area.ToString
    '                If TBLWarehouse.DefaultView.Count = 0 Then
    '                    TblQty3.DefaultView.RowFilter = Nothing
    '                    TblQty3.DefaultView.RowFilter = "[Primary Product] = " + rowQty2.Primary_Product.ToString + " AND [Primary Warehouse] = -1"
    '                    If TblQty3.DefaultView.Count = 0 Then
    '                        'tambahkan baris baru
    '                        rowQty3 = TblQty3.Newtblmaster_productqty1Row
    '                        With rowQty3
    '                            .Primary_Product = rowQty2.Primary_Product
    '                            .Primary_Warehouse = -1
    '                            .Qty = rowQty2.Qty
    '                            .Warehouse_Code = "N/A"
    '                            .Warehouse_Desc = "N/A"
    '                        End With
    '                        TblQty3.Rows.Add(rowQty3)
    '                    Else
    '                        'tambahkan qty saja
    '                        rowQty3 = TblQty3.DefaultView(0).Row
    '                        rowQty3.Qty += rowQty2.Qty
    '                    End If
    '                Else
    '                    rowWarehouse = TBLWarehouse.DefaultView(0).Row
    '                    TblQty3.DefaultView.RowFilter = Nothing
    '                    TblQty3.DefaultView.RowFilter = "[Primary Product] = " + rowQty2.Primary_Product.ToString + " AND [Primary Warehouse] = " + rowWarehouse.Primary_Warehouse.ToString
    '                    If TblQty3.DefaultView.Count = 0 Then
    '                        'tambahkan baris baru
    '                        rowQty3 = TblQty3.Newtblmaster_productqty1Row
    '                        With rowQty3
    '                            .Primary_Product = rowQty2.Primary_Product
    '                            .Primary_Warehouse = rowWarehouse.Primary_Warehouse
    '                            .Qty = rowQty2.Qty
    '                            .Warehouse_Code = rowWarehouse.Warehouse_Code
    '                            .Warehouse_Desc = rowWarehouse.Warehouse_Desc
    '                        End With
    '                        TblQty3.Rows.Add(rowQty3)
    '                    Else
    '                        'tambahkan qty saja
    '                        rowQty3 = TblQty3.DefaultView(0).Row
    '                        rowQty3.Qty += rowQty2.Qty
    '                    End If
    '                End If
    '            Next
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message + ": " + ProductCodeRef, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        TBL_Awal.Dispose()
    '        DTAProduct.Dispose()
    '        TBLProduct.Dispose()
    '        DTAQty.Dispose()
    '        TBLQty1.Dispose()
    '        TblQty2.Dispose()
    '        DTASC.Dispose()
    '        TBLSC.Dispose()
    '        DTA_Transfer.Dispose()
    '        TBL_Transfer.Dispose()
    '        DTAWarehouse.Dispose()
    '        TBLWarehouse.Dispose()
    '    End Try
    'End Sub

    Private Function ApakahAdaQtyMinus(ByVal vtbl As DseMonthlyClosingNew.tblmaster_productqty1DataTable, ByRef ListProductMinus As List(Of String), _
    ByVal vFrm As FrmTransaksiMonthlyClosing_CekMinus) As Boolean
        Dim vrow As DseMonthlyClosingNew.tblmaster_productqty1Row
        Dim rowDGV As DataGridViewRow

        Try
            vFrm.dgvItem.Rows.Clear()

            If vtbl.Rows.Count > 0 Then
                For Each vrow In vtbl.Rows
                    With vrow
                        If .Qty < 0 Then
                            If ListProductMinus.Contains(vrow.Primary_Product) = False Then
                                vFrm.dgvItem.Rows.Add()
                                rowDGV = vFrm.dgvItem.Rows(vFrm.dgvItem.Rows.Count - 1)

                                rowDGV.Cells(vFrm.ColCodeNo.Index).Value = .Code_No
                                rowDGV.Cells(vFrm.ColDesc.Index).Value = .Description
                                rowDGV.Cells(vFrm.ColNo.Index).Value = rowDGV.Index + 1
                                rowDGV.Cells(vFrm.ColQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty, .Angka_Decimal)
                                rowDGV.Cells(vFrm.ColUnit.Index).Value = .Unit_Code

                                ListProductMinus.Add(.Code_No & " - " & .Description & " = " & Me.ObjCD.DigitGrouping(.Qty, .Angka_Decimal) + _
                                " " + .Unit_Code)
                            End If
                        End If
                    End With
                Next
            End If
            If ListProductMinus.Count > 0 Then
                Return True
            End If

            Return False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub LoadClosingData()
    '    Dim row As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2Row
    '    Dim rowDGV As DataGridViewRow
    '    Dim OldText As String = Me.Text
    '    Dim TipeTransaksiStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
    '    'Dim TBLInAdjust As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item6DataTable
    '    'Dim DTAAdjustIn As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item6TableAdapter
    '    'Dim rowAdjust As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item6Row
    '    Dim TotalAdjustment As Decimal = 0
    '    Dim counter As Double = 0

    '    Dim AddRowDelegate As AddNewRow = AddressOf Me.AddNewRow1
    '    'Dim AddInAmountDelegate As AddInAmount = AddressOf Me.AddInAmount1
    '    'Dim SetEndAmountDelegate As SetEndAmount = AddressOf Me.SetEndAmount1
    '    'Dim SetOutAmountDelegate As SetOutAmount = AddressOf Me.SetOutAmount1

    '    'Dim DTATaking1 As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_stock_takingTableAdapter
    '    'Dim DTATaking2 As New DseTransaksiMonthlyClosingTableAdapters.laporan_stock_taking_view2TableAdapter
    '    'Dim TBLTaking1 As New DseTransaksiMonthlyClosing.tbltrans_stock_takingDataTable
    '    'Dim TBLTaking2 As New DseTransaksiMonthlyClosing.laporan_stock_taking_view2DataTable
    '    'Dim rowTaking1 As DseTransaksiMonthlyClosing.tbltrans_stock_takingRow
    '    'Dim rowTaking2 As DseTransaksiMonthlyClosing.laporan_stock_taking_view2Row

    '    'Dim DTABundling As New DseTransaksiMonthlyClosingTableAdapters.DataTable12TableAdapter
    '    'Dim TBLBundling As New DseTransaksiMonthlyClosing.DataTable12DataTable
    '    'Dim rowBundling As DseTransaksiMonthlyClosing.DataTable12Row

    '    Dim TBLSalesReturn As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item81DataTable
    '    Dim DTASalesReturn As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item81TableAdapter
    '    Dim rowSalesReturn As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item81Row
    '    Dim counter1 As Double = 0
    '    Dim Filter As String = ""

    '    'Dim TBL_In_SR As New DseTransaksiMonthlyClosing.DataTable20DataTable
    '    'Dim row_In_SR As DseTransaksiMonthlyClosing.DataTable20Row
    '    'Dim DTA_In_SR As New DseTransaksiMonthlyClosingTableAdapters.DataTable20TableAdapter

    '    'Dim DTA_HIstoricalPrice As New DseTransaksiMonthlyClosingTableAdapters.DataTable29TableAdapter
    '    'Dim TBL_HistoricalPrice As New DseTransaksiMonthlyClosing.DataTable29DataTable
    '    'Dim row_HistoricalPrice As DseTransaksiMonthlyClosing.DataTable29Row
    '    Dim cmd As String = ""

    '    Dim DTABefore As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_beforeTableAdapter '--> Update ke SMESS
    '    Dim DTAAfter As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter
    '    Dim TBLBefore As New DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable
    '    Dim TBLAfter As New DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable

    '    Dim TBLBefore1 As New DseTransaksiMonthlyClosing.DataTable16DataTable
    '    Dim DTABefore1 As New DseTransaksiMonthlyClosingTableAdapters.DataTable16TableAdapter
    '    Dim TBLIn As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable
    '    Dim DTAIn As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item2TableAdapter

    '    Dim TBLPurchase As New DseTransaksiMonthlyClosing.DataTable4DataTable
    '    Dim DTAPurchase As New DseTransaksiMonthlyClosingTableAdapters.DataTable4TableAdapter
    '    Dim TBLSR As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable
    '    Dim DTASR As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item9_1TableAdapter

    '    Dim Sort As Boolean = False
    '    Dim DTABundling1 As New Inventory_Module.dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_after1TableAdapter
    '    Dim DefPrec As Integer = 0

    '    Dim DTA_Prod As New DseTransaksiMonthlyClosingTableAdapters.tblmaster_productTableAdapter
    '    Dim TBL_Prod As New DseTransaksiMonthlyClosing.tblmaster_productDataTable
    '    Dim ijk As Double = 0

    '    Try
    '        DefPrec = Me.objCurrency.GetDecimal_DefaultCurrency

    '        If Me.baru = True Then
    '            'hapuskan data unit price sementara
    '            Me.TBLUnitPrice.Rows.Clear()
    '            Me.TBLItemAfter.Rows.Clear()
    '            Me.TBLItemBefore.Rows.Clear()

    '            Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

    '            Me.dgvItem.Rows.Clear()
    '            Me.dgvItemAdjust.Rows.Clear()

    '            'set disable
    '            If Me.ObjClosing.GetJml = 0 Then
    '                Me.ColInAmount.Visible = False
    '                Me.ColInQty.Visible = False
    '                Me.ColBeginQty.Visible = False
    '                Me.ColBeginAmount.Visible = False
    '                Me.ColOutQty.Visible = False
    '                Me.ColOutAmount.Visible = False
    '                'Me.dtpPeriod.Enabled = True
    '                Me.dtpPeriod.Enabled = False
    '                Me.dgvItem.ReadOnly = False
    '                Me.LoadClosing = False
    '                'Me.dtpPeriod.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
    '                Me.dtpPeriod.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.ObjCD.GetFirstDateofTheMonth(Me.ObjAccClosing.GetConvertionDate()) - TimeSpan.FromDays(5))
    '                Me.LoadClosing = True
    '                'Me.ColDetailCalculation.Visible = False
    '            Else
    '                Me.ColInAmount.Visible = True
    '                Me.ColInQty.Visible = True
    '                Me.ColBeginQty.Visible = True
    '                Me.ColBeginAmount.Visible = True
    '                Me.ColOutQty.Visible = True
    '                Me.ColOutAmount.Visible = True
    '                Me.dtpPeriod.Enabled = False
    '                Me.LoadClosing = False
    '                Me.dtpPeriod.Value = Me.ObjClosing.GetNextClosingPeriode
    '                Me.LoadClosing = True
    '                'sementara
    '                'Me.ColDetailCalculation.Visible = False
    '            End If

    '            'load data unit price 
    '            DTA_HIstoricalPrice.Fill(TBL_HistoricalPrice)

    '            Me.Text = OldText + " - Extracting product data from database"
    '            'isikan data awal
    '            Me.TBL.Rows.Clear()
    '            If DTABundling1.ScalarQuery(Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), Me.dtpPeriod.Value.Date).HasValue = True Then '--> Update ke SMESS, berkaitan dengan urutan
    '                If DTABundling1.ScalarQuery(Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), Me.dtpPeriod.Value.Date).Value > 25 Then
    '                    Sort = True
    '                Else
    '                    Sort = False
    '                End If
    '            Else
    '                Sort = False
    '            End If
    '            Me.ObjClosing.IsiTBL4Closing(Me.TBL, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
    '            Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value), Sort)

    '            Me.DTABeginning.Fill(Me.TBLBeginning)
    '            Me.SetBeginningBalance(Me.TBL)
    '            DTASalesReturn.Fill(TBLSalesReturn, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), _
    '            Me.dtpPeriod.Value.Date)
    '            counter1 = 0
    '            While counter1 < TBLSalesReturn.Rows.Count
    '                rowSalesReturn = TBLSalesReturn.Rows(counter1)
    '                If rowSalesReturn.Price_At_Closing = 0 Then
    '                    If rowSalesReturn.Closing_Period.Date <> Me.dtpPeriod.Value.Date Then
    '                        'tambahkan nilai in amount untuk produt yang qty end =0, ingat kalo qty end =0 
    '                        'query mengembalikan nilai in amount =0, jadi ditambahkan secara manual
    '                        'nilainya 0 karena stocknya keluar semua pada bulan ybs
    '                        Filter = "[Primary Product] = " + rowSalesReturn.Primary_Product.ToString
    '                        Me.TBL.DefaultView.RowFilter = Filter
    '                        If TBL.DefaultView.Count > 0 Then
    '                            row = TBL.DefaultView(0).Row
    '                            row.In_Amount += Me.GetUnitPrice(rowSalesReturn.Primary_Product, _
    '                            rowSalesReturn.Closing_Period) * rowSalesReturn.Total_Qty_Retur
    '                        End If
    '                    End If
    '                    counter1 += 1
    '                Else
    '                    TBLSalesReturn.Rows.Remove(rowSalesReturn)
    '                End If
    '            End While

    '            'jika ada bundling yang berputar tanpa batas maka ambil dari dgv 
    '            DTABundling.Fill(TBLBundling, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), _
    '            Me.dtpPeriod.Value.Date)

    '            DTABefore.FillBy_TanggalSampaiDari(TBLBefore, _
    '            Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value)

    '            'isikan data yang akan digunakan untuk perhitungan bundling pada fungsi GetUnitPrice 
    '            DTABefore1.FillBy_ClosingPeriod2(TBLBefore1, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date - TimeSpan.FromDays(1)))
    '            DTAIn.Fill(TBLIn, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), _
    '            Me.dtpPeriod.Value.Date)
    '            DTAPurchase.Fill(TBLPurchase, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), Me.dtpPeriod.Value.Date)
    '            DTASR.FillBy_Period(TBLSR, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), Me.dtpPeriod.Value.Date)
    '            Me.DTA_CL.FillBy_All(Me.TBL_CL)

    '            For Each row In Me.TBL.Rows
    '                Me.Text = OldText + "- Loading product data (" + CStr(Me.dgvItem.Rows.Count) _
    '                + " of " + Me.TBL.Rows.Count.ToString + ")"
    '                If row.Begin_Qty <> 0 Or row.In_Qty <> 0 Or row.Out_Qty <> 0 Or row.End_Qty <> 0 Or _
    '                row.Begin_Amount <> 0 Or row.In_Amount <> 0 Then
    '                    'tambahkan baris baru
    '                    AddRowDelegate.Invoke(row.Begin_Qty, row.Begin_Amount, row.Primary_Product, row.Product_Code, _
    '                    row.Description, row.Primary_Unit, row.Unit_Code, row.Out_Qty, row.End_Qty, row.In_Qty, _
    '                    row.In_Amount)
    '                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)

    '                End If
    '            Next

    '            'load detail item
    '            'Exit Sub
    '            Me.LoadItemDetail()

    '            ' Me.FormatDGV()
    '        Else
    '            Me.dtpPeriod.Enabled = False
    '            'update tabel product
    '            Me.UpdateTBLProduct()
    '        End If

    '        If Sort = True Then
    '            Me.dgvItem.Sort(Me.ColCodeNo, System.ComponentModel.ListSortDirection.Ascending) '--> update ke SMESS
    '            Me.ObjCD.SetNumber(Me.dgvItem, True)
    '        End If

    '        'buang yang jasa pada dgv adjustment 
    '        DTA_Prod.Fill(TBL_Prod)
    '        ijk = 0
    '        While ijk <= Me.dgvItemAdjust.Rows.Count - 1
    '            rowDGV = Me.dgvItemAdjust.Rows(ijk)
    '            TBL_Prod.DefaultView.RowFilter = "Primary = " + rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value.ToString

    '            If TBL_Prod.DefaultView.Count = 0 Then
    '                Me.dgvItemAdjust.Rows.Remove(rowDGV)
    '            Else
    '                ijk += 1
    '            End If
    '        End While

    '        'sementara
    '        'Me.ColDetailCalculation.Visible = False

    '        'Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '        'Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

    '        Me.Text = OldText
    '        Me.dgvItemAdjust.Refresh()

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTAAdjustIn.Dispose()
    '        TBLInAdjust.Dispose()
    '        DTATaking1.Dispose()
    '        DTATaking2.Dispose()
    '        TBLTaking1.Dispose()
    '        TBLTaking2.Dispose()
    '        TBLSalesReturn.Dispose()
    '        DTASalesReturn.Dispose()
    '        TBL_In_SR.Dispose()
    '        DTA_In_SR.Dispose()
    '        DTA_HIstoricalPrice.Dispose()
    '        TBL_HistoricalPrice.Dispose()

    '        DTABefore.Dispose()
    '        DTAAfter.Dispose()
    '        TBLBefore.Dispose()
    '        TBLAfter.Dispose()
    '        TBLBefore1.Dispose()
    '        DTABefore1.Dispose()
    '        TBLAfter.Dispose()
    '        DTAAfter.Dispose()
    '        TBLPurchase.Dispose()
    '        DTAPurchase.Dispose()
    '        DTASR.Dispose()
    '        TBLSR.Dispose()

    '        DTA_Prod.Dispose()
    '        TBL_Prod.Dispose()
    '    End Try
    'End Sub

    ''' <summary>
    ''' tidak memperhitungkan baris terakhir dalam dgv
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetInAmount_DiDGVItem(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index <> Me.dgvItem.Rows.Count - 1 Then
                    If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                        Return rowDGV.Cells(Me.ColInAmount.Index).Value
                    End If
                End If
            Next

            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahSudahAda(ByVal vPrimaryCodeNo As Double) As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index <> Me.dgvItem.Rows.Count - 1 Then
                    If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                        Return True
                    End If
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Public Function GetUnitPrice(ByVal vPrimaryCodeNo As Double, ByVal DTABefore As DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_beforeTableAdapter, _
    'ByVal DTAAfter As DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter, ByVal TBLBefore As DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable, _
    'ByVal TBLAfter As DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable, ByVal ObjClosingItem As ClsTransaksiMonthlyClosingItem, _
    'ByVal TBL1 As DseTransaksiMonthlyClosing.DataTable16DataTable, _
    'ByVal TBL2 As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable, _
    'ByVal TBLPurchase As DseTransaksiMonthlyClosing.DataTable4DataTable, _
    'ByVal TBLSR As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable) As Decimal '--> update ke smess

    '    Dim InAmount As Decimal = 0
    '    Dim rowDGVAdjust As DataGridViewRow
    '    Dim BeginQty, BeginAmount, InQty As Decimal

    '    Try
    '        'ambil begin qty, begin amount, in qty, in amount-step1
    '        ObjClosingItem.Get_BeginAmount_BeginQty_InQty_IAmount(vPrimaryCodeNo, _
    '        BeginAmount, BeginQty, InQty, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
    '        Me.dtpPeriod.Value, Me.ObjCD, TBL1, TBL2)

    '        'hitung unit price
    '        If BeginQty + InQty <> 0 Then
    '            If Me.ApakahSudahAda(vPrimaryCodeNo) = True Then
    '                InAmount = Me.GetInAmount_DiDGVItem(vPrimaryCodeNo)
    '            Else
    '                'tambahkan in amount dengan:-step2
    '                InAmount = ObjClosingItem.GetInAmount_Purchase_SalesReturnBedaPeriode(vPrimaryCodeNo, _
    '                Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
    '                Me.dtpPeriod.Value, Me.TBL_CL, Me.ObjCD, TBLPurchase, TBLSR, Me.DefaultPrecision)

    '                '1. bundling-step3
    '                InAmount = InAmount + Me.GetInAmountFromBundling(vPrimaryCodeNo, DTABefore, DTAAfter, TBLBefore, TBLAfter, TBL1, TBL2, TBLPurchase, TBLSR)

    '                '2. in adjustment dari stock in dan stock taking-step4
    '                For Each rowDGVAdjust In Me.dgvItemAdjust.Rows
    '                    If rowDGVAdjust.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = vPrimaryCodeNo Then
    '                        InAmount += Math.Round(rowDGVAdjust.Cells(Me.ColUnitPriceAdjust.Index).Value * _
    '                        rowDGVAdjust.Cells(Me.ColQty.Index).Value, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
    '                    End If
    '                Next
    '            End If

    '            'hitung unit price
    '            Return Math.Round((BeginAmount + InAmount) / _
    '            (BeginQty + InQty), Me.DefaultPrecision, MidpointRounding.AwayFromZero)
    '        Else
    '            'tidak ada unit price
    '            Return 0
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjClosingItem = Nothing
    '    End Try
    'End Function

    Public Function GetUnitPrice(ByVal vPrimaryProduct As Double, ByVal vClosingPeriod As Date) As Decimal
        Dim Filter As String = ""
        Dim rowClosing As DseTransaksiMonthlyClosing.DataTable21Row

        Try
            If vClosingPeriod <> Me.Dum_ClosingPeriode Then
                Me.Dum_ClosingPeriode = vClosingPeriod
                Me.DTAClosing1.Fill(Me.TBLClosing1, Me.Dum_ClosingPeriode)
            End If
            Filter = "[Primary Product] = " + vPrimaryProduct.ToString
            Me.TBLClosing1.DefaultView.RowFilter = Filter
            If Me.TBLClosing1.DefaultView.Count = 0 Then
                Return 0
            Else
                rowClosing = Me.TBLClosing1.DefaultView(0).Row
                If rowClosing.Ending_Qty = 0 Then
                    If rowClosing.Out_Qty = 0 Then
                        Return 0
                    Else
                        Return Math.Round(rowClosing.Out_Amount / rowClosing.Out_Qty, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                    End If
                Else
                    Return Math.Round(rowClosing.Ending_Amount / rowClosing.Ending_Qty, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Information")
        End Try
    End Function

    Private Sub SetBeginningBalance(ByVal vTBL As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable)
        Dim row As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2Row
        Dim rowBeginning As DseTransaksiMonthlyClosing.tbltrans_purchase_closing_item1Row
        Dim Filter As String = ""

        Try
            For Each row In vTBL.Rows
                Filter = "[Primary Product] =" + row.Primary_Product.ToString
                Me.TBLBeginning.DefaultView.RowFilter = Filter
                If Me.TBLBeginning.DefaultView.Count <> 0 Then
                    rowBeginning = Me.TBLBeginning.DefaultView(0).Row
                    row.Begin_Amount = rowBeginning.Ending_Amount
                    row.Begin_Qty = rowBeginning.Ending_Qty
                Else
                    row.Begin_Amount = 0
                    row.Begin_Qty = 0
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub SetOutAmount1(ByVal rowDGV As DataGridViewRow, ByVal OutAmount As Decimal)
    '    Try
    '        rowDGV.Cells(Me.ColOutAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(OutAmount, _
    '        Me.DefaultPrecision, MidpointRounding.AwayFromZero))
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub SetEndAmount1(ByVal rowDGV As DataGridViewRow, ByVal EndAmount As Decimal)
    '    Try
    '        rowDGV.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(EndAmount, _
    '        Me.DefaultPrecision, MidpointRounding.AwayFromZero))
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub AddInAmount1(ByVal vPrimaryCodeNo As Double, ByVal vAdditionalAmount As Decimal)
    '    Dim rowDGV As DataGridViewRow

    '    Try
    '        For Each rowDGV In Me.dgvItem.Rows
    '            If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
    '                rowDGV.Cells(Me.ColInAmount.Index).Value = CDec(rowDGV.Cells(Me.ColInAmount.Index).Value) _
    '                + vAdditionalAmount
    '                Exit Sub
    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Public Sub AddNewRow1(ByVal BeginQty As Decimal, ByVal BeginAmount As Decimal, _
    ByVal vPrimaryProduct As Double, ByVal vCodeNo As String, ByVal Description As String, ByVal vPrimaryUnit As Double, _
    ByVal UnitCode As String, ByVal OutQty As Decimal, ByVal EndQty As Decimal, ByVal InQty As Decimal, ByVal InAmount As Decimal)
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvItem.Rows.Add("")
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
            With rowDGV
                .Cells(Me.ColNo.Index).Value = CStr(Me.dgvItem.Rows.Count)
                .Cells(Me.ColBeginQty.Index).Value = BeginQty
                .Cells(Me.ColBeginAmount.Index).Value = BeginAmount
                .Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct
                .Cells(Me.ColCodeNo.Index).Value = vCodeNo
                .Cells(Me.ColDescription.Index).Value = Description
                .Cells(Me.ColPrimaryUnits.Index).Value = vPrimaryUnit
                .Cells(Me.ColUnits.Index).Value = UnitCode
                .Cells(Me.ColRemarks.Index).Value = ""
                .Cells(Me.ColDetailCalculation.Index).Value = "Show"
                .Cells(Me.ColOutAmount.Index).Value = 0
                .Cells(Me.ColOutQty.Index).Value = OutQty
                .Cells(Me.ColEndAmount.Index).Value = 0
                .Cells(Me.ColEndQty.Index).Value = EndQty
                .Cells(Me.ColInAmount.Index).Value = InAmount
                .Cells(Me.ColInQty.Index).Value = InQty
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                'kalo ada yang perlu dihitung ulang berkaitan dengan data bundling, maka lakukan perhitungan ulang dahulu
                'If Me.ListPrimaryCodeNo_Bundling.Count <> 0 Then
                '    Me.HitungUlangUnitPrice1()
                'End If

                If Me.saveTemporaryAktif = True Then
                    Me.DeleteTemporaryDataHeader()
                End If
                
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjClosing.SimpanBaru(Me.dgvItem, Me.dgvItemAdjust, Me.UserCode, Me, _
                        Me.TBLItemDetail, Me.TBLBundling1, Me.dgvJurnal, Me.TBLSetting, Me.TBLSettingItem, Me.tblPool, _
                            Me.dgvItemCost_CostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, vTBLCostGroup_CodeNo, vTBLCostGroup_CodeNo2, Me.TblProductQty) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjClosing.SimpanPerubahan(Me.dgvItem, Me.dgvItemAdjust, Me.UserCode, Me, _
                        Me.TBLItemDetail, Me.TBLBundling1, Me.dgvJurnal, Me.TBLSetting, Me.TBLSettingItem, Me.tblPool, _
                            Me.dgvItemCost_CostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, vTBLCostGroup_CodeNo, vTBLCostGroup_CodeNo2, Me.TblProductQty) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.LoadClosing = False
            ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.LoadClosing = True
            Me.TxtPrimaryClosing.Text = "-1"

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvItem.Rows.Clear()
            Me.dgvItemAdjust.Rows.Clear()
            Me.dgvJurnal.Rows.Clear()
            Me.KosongkanField_Overhead()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtRemark.Text = "" And Me.dgvItem.RowCount = 0 Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                If Me.saveTemporaryAktif = True Then
                    DeleteTemporaryDataHeader()
                End If
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        If Me.saveTemporaryAktif = True Then
                            DeleteTemporaryDataHeader()
                        End If
                        Me.KosongkanField()
                    Else
                        Me.ObjClosing.CariDetail(Me.TxtPrimaryClosing.Text, Me.dgvItem, Me.dgvItemAdjust, _
                        Me.TBLItemDetail, Me.TBLBundling1, Me, Me.dgvJurnal, Me.TBLSettingItem, _
                            Me.dgvItemCost_CostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, vTBLCostGroup_CodeNo, Me, Me.vTBLCostGroup_CodeNo2, Me.TblProductQty)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariTransaksiMonthlyClosing

        Try
            ij.BackColorAll = Me.BAckColorAll
            ij.ShowDialog()
            If Not ij.PrimaryClosing = -1 Then
                If Me.ObjClosing.CariDetail(ij.PrimaryClosing, Me.dgvItem, Me.dgvItemAdjust, _
                        Me.TBLItemDetail, Me.TBLBundling1, Me, Me.dgvJurnal, Me.TBLSettingItem, _
                            Me.dgvItemCost_CostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, vTBLCostGroup_CodeNo, Me, Me.vTBLCostGroup_CodeNo2, Me.TblProductQty) _
                = True Then
                    Me.SalinKeField()
                    'jika yang ditampilkan adalah data pertama, maka hide kolom kolom yang tidak perlu
                    If Me.ObjClosing.ApakahClosingPalingAwal(Me.TxtPrimaryClosing.Text) = True Then
                        Me.ColInAmount.Visible = False
                        Me.ColInQty.Visible = False
                        Me.ColBeginQty.Visible = False
                        Me.ColBeginAmount.Visible = False
                        Me.ColOutQty.Visible = False
                        Me.ColOutAmount.Visible = False
                        'Me.ColDetailCalculation.Visible = False
                    Else
                        Me.ColInAmount.Visible = True
                        Me.ColInQty.Visible = True
                        Me.ColBeginQty.Visible = True
                        Me.ColBeginAmount.Visible = True
                        Me.ColOutQty.Visible = True
                        Me.ColOutAmount.Visible = True
                        'sementara
                        'Me.ColDetailCalculation.Visible = False
                    End If
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If

            'Me.ColDetailCalculation.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub ShowProductVerificationForm()
        Try
            Dim ij As Core_Module.frmMasterProduct_Verifikasi
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.frmMasterProduct_Verifikasi)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.frmMasterProduct_Verifikasi
            ' Increment the caption counter.
            i.UserCode = Me.UserCode
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me.ParentForm
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim Dummy As String = ""
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim DTA_Unbalance_Journal As New AccModule.DseTransaksiJournalTableAdapters.UnBalanceJurnalTableAdapter
        Dim TBL_Unbalance_Journal As New AccModule.DseTransaksiJournal.UnBalanceJurnalDataTable
        Dim row_UnBalance_Journal As AccModule.DseTransaksiJournal.UnBalanceJurnalRow
        Dim Counter As Double = 0
        Dim DefPrec1 As Integer = Me.objCurrency.GetDecimal_DefaultCurrency
        Dim RefDate_Verify As Date = Me.ObjClosing.GetNextClosingPeriode()
        Dim FrmCek1 As New Core_Module.FrmCekSiklikBundling
        Dim NextClosingPeriod As Date
        Dim FrmCekMinus As New FrmTransaksiMonthlyClosing_CekMinus
        Dim ObjClosingJur As New AccModule.ClsTransaksiClosing

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            NextClosingPeriod = Me.ObjClosing.GetNextClosingPeriode()
            If NextClosingPeriod <> #1/1/1950# Then
                If Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod) <= ObjClosingJur.GetClosingDate.Date Then
                    MsgBox("You can't create new closing data because " + Format(Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod), "dd MMM yyyy") + " <= journal lock date" _
                        , MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Jurnal_VerifyProductAccount) = "1" Then
                If Me.ObjProduct.ApakahAda_ProductNotVerified(RefDate_Verify) = True Then
                    MsgBox("There are several products with unverified account. Please verify it first", MsgBoxStyle.Information, "Information")
                    'panggil form untuk verification
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "frmMasterProduct_Verifikasi") = True Then
                        Me.ShowProductVerificationForm()
                    End If
                    Exit Sub
                End If
            End If

            'isikan data tbl product untuk mendapatkan angka decimal
            Me.vDTA1.Fill(Me.vTBL1)

            'peringatan, untuk bulan yang baru, tidak boleh ada AOP dan transaksi stock taking yang
            'masih aktif pada bulan yang bersangkutan
            If Me.ObjClosing.GetJml > 0 Then

                Dim ObjGRN As New ClsTransaksiGRN
                If ObjGRN.GetJmlBelumAOP(Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod), NextClosingPeriod) > 0 Then
                    MsgBox("There are stock in purchase data in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " that don't have approval yet. Please create AOP for: " + _
                    ObjGRN.GetNoGRN_BelumAOP(Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod), NextClosingPeriod) + ".", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                ObjGRN = Nothing

                Dim ObjStockTaking As New Inventory_Module.clsTransaksiStockTaking
                If ObjStockTaking.GetJmlActiveStockTaking(Me.ObjCD.GetFirstDateofTheMonth(NextClosingPeriod), Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod)) > 0 Then
                    MsgBox("There is an active stock taking on progress, please be aware that stock taking processing date cannot be in the period that has been closed. Please finish the stock taking process first.", MsgBoxStyle.Information, "Confirm")
                    ObjStockTaking = Nothing
                    Exit Sub
                End If
                ObjStockTaking = Nothing

                'tidak ada GRN yang dibuat pada bulan ybs yang di AOP bukan pada bulan ybs.
                Dummy = Me.ObjClosing.GetGRN_AOP_BedaPeriode(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are GRN in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " enter in AOP on different month. Please fix following data before continuing closing transaction:" + Dummy, _
                    MsgBoxStyle.Information, "Confirmation")

                    Exit Sub
                End If

                Dummy = Me.ObjClosing.GetAOP_GRNBEdaPeriode(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are AOP in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " enter GRN from different month. Please fix following data before continuing closing transaction:" _
                    + Dummy, MsgBoxStyle.Information, "Confirmation")
                    Exit Sub
                End If

                'tidak boleh ada purchase return yang belum di stock out
                Dummy = Me.ObjClosing.Get_PR_BelumStockOut(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are purchase return in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " with no stock out purchase return. Please fix this data before continuing closing transaction: " + Dummy, _
                    MsgBoxStyle.Information, "Confirmation")
                    Exit Sub
                End If

                'tidak stock out sales yang muncul dari periode sebelumnya --> berkaitan dengan balance out amount
                Dummy = Me.ObjClosing.GetInvoice_StockOutSalesPreviousPeriode(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    If MsgBox("There are invoice in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " with stock out sales from previous month. Please fix this data before continuing closing transaction. Show detail data?", _
                    MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                        Dim ijk As New Sales_Module.FrmLaporanExplainClosingB
                        ijk.dtpFrom.Value = NextClosingPeriod
                        ijk.ShowDialog()
                        ijk = Nothing
                    End If
                    Exit Sub
                End If

                'tidak DO yang sudah stock out sales tapi belum diinvoice
                Dummy = Me.ObjClosing.Get_DO_BelumInvoice(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are delivery order having stock out sales in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " with no invoice. Please fix the following data before continue: " + Dummy, _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                'tidak boleh ada stock in sales return yang dibuatkan sales return pada beda periode
                'atau bahkan belum dibuatkan sales return 
                Dummy = Me.ObjClosing.GetStockInSalesReturn_SRBedaPeriode(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are stock in sales return in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " with out sales return transaction. Please fix this data before continuing closing transaction: " + Dummy, _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                'tidak boleh ada sales return yang stock in sales returnya berasal dari periode bukan 
                'bulan yang sama
                Dummy = Me.ObjClosing.GetSR_StockInSalesReturnBedaPeriode(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are sales return in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " with stock in sales return transaction on different month. Please fix this data before continuing closing transaction: " + Dummy, _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                'tidak boleh ada sales return revisi invoice yang belum dimasukkan ke dalam revisi invoice
                Dummy = Me.ObjClosing.Get_SR_RevInvoice_BelumRevisi(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are sales return(invoice revision) in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " but no invoice. Please fix this data before continuing closing transaction: " + Dummy, _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                'tidak boleh ada purchase return revisi AOP yang belum dimasukkan pada AOP
                Dummy = Me.ObjClosing.Get_PR_Rev_BelumMasukAOP(NextClosingPeriod, _
                                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are purchase return(AOP revision) in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " but no AOP. Please fix this data before continuing closing transaction: " + Dummy, _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            'cek apakah ada unbalance jurnal 
            DTA_Unbalance_Journal.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Unbalance_Journal.Fill(TBL_Unbalance_Journal)
            Counter = 1
            If TBL_Unbalance_Journal.Rows.Count > 0 Then
                For Each row_UnBalance_Journal In TBL_Unbalance_Journal.Rows
                    If Me.ObjCD.DigitGrouping(row_UnBalance_Journal.Debet, DefPrec1) <> Me.ObjCD.DigitGrouping(row_UnBalance_Journal.Kredit, DefPrec1) Then
                        If Dummy = "" Then
                            Dummy = Counter.ToString + ". " + row_UnBalance_Journal.Ref_Number
                        Else
                            Dummy = Dummy + vbCrLf + Counter.ToString + ". " + row_UnBalance_Journal.Ref_Number
                        End If
                        Counter += 1
                    End If
                Next
                If Counter > 1 Then
                    If MsgBox("The following journal is unbalance, please contact your vendor to solve this issue. Continue process anyway?" + vbCrLf + Dummy, _
                     MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Exit Sub
                    End If
                End If
            End If


            Me.baru = True
            Me.KosongkanField()
            Me.ObjClosing.ResetObject(Me.dgvItem)
            Form.CheckForIllegalCrossThreadCalls = False
            Me.dtpDate.Focus()
            Me.settingHeader()
            Me.TxtClosingNumber.Text = Me.ObjClosing.GetNoBaru(Me.dtpPeriod.Value)

            ' MsgBox("Pengecekan minus belum")
            Me.IsikanDataPengecekanQtyMinus()
            Dim ListProductMinus As New List(Of String)
            Me.TblQty3.DefaultView.RowFilter = Nothing
            If Me.ApakahAdaQtyMinus(Me.TblQty3, ListProductMinus, FrmCekMinus) = True Then
                'Dim strMsg As String = ""
                'For intA As Integer = 0 To ListProductMinus.Count - 1
                '    If strMsg = "" Then
                '        strMsg &= ListProductMinus.Item(intA)
                '    Else
                '        strMsg &= vbNewLine & ListProductMinus.Item(intA)
                '    End If
                'Next
                If FrmCekMinus.dgvItem.RowCount > 0 Then
                    MsgBox("The following product has qty minus, please fix this data before continuing closing transaction.", _
                    MsgBoxStyle.Information, "Information")
                    FrmCekMinus.ShowDialog()
                    Exit Sub
                End If
            End If

            'cek siklik bundling
            'MsgBox("Pengecekan siklik bundling harus dinyalain lagi")
            If Me.ObjClosing.GetJml > 0 Then
                'MsgBox("skip siklik bundling")
                FrmCek1.CekFromClosing = True
                FrmCek1.CheckPeriod = NextClosingPeriod

                If FrmCek1.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                    If FrmCek1.DataGridView1.RowCount > 0 Then
                        Exit Sub
                    End If
                End If
            End If

            Me.saveTemporaryHeaderData()
            'kosongkan tabel adjustment sementara untuk isi
            Me.TblA.Rows.Clear()
            Me.FillDatatableAdjustment(True)
            If Me.ObjClosing.ApakahClosingPalingAwal(Me.TxtPrimaryClosing.Text) = False Then
                If Me.ShowDataAdjustment() = False Then
                    If Me.saveTemporaryAktif = True Then
                        DeleteTemporaryDataHeader()
                    End If
                    Me.KosongkanField()
                    Me.NonAktifkanDetail()
                    Me.btnNew.Focus()

                    Exit Sub
                End If
            End If

            If Me.ObjClosing.ApakahClosingPalingAwal(Me.TxtPrimaryClosing.Text) = False Then
                If Me.CopyDateOverheadBulanSebelumnya = False Then
                    If Me.saveTemporaryAktif = True Then
                        DeleteTemporaryDataHeader()
                    End If
                    Me.KosongkanField()
                    Me.NonAktifkanDetail()
                    Me.btnNew.Focus()

                    Exit Sub
                Else
                    'dipindahkan ke bgw calculate dan bgw aktifkan detail
                    'Me.RecalculateOverhead()
                    'Me.TampilkanOverheadDetail()
                End If
            End If

            Me.IsiKeDGVAdjust()
            Me.AktifkanDetail()

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Unbalance_Journal.Dispose()
            TBL_Unbalance_Journal.Dispose()
            FrmCekMinus = Nothing
            ObjClosingJur = Nothing
        End Try
    End Sub

    Private Function CopyDateOverheadBulanSebelumnya() As Boolean
        Dim vFrm1 As New FrmTransaksiMonthlyClosing2_Overhead2
        Dim rowDGV, rowdgv1 As DataGridViewRow
        Dim Dummy1 As Boolean = False

        Try
            vFrm1.usercode = Me.UserCode
            vFrm1.vTBLCostGroup_Account = Me.vTBLCostGroup_Account
            vFrm1.vTBLCostGroup_CodeNo = Me.vTBLCostGroup_CodeNo
            vFrm1.vTBLCostGroup_Scope = Me.vTBLCostGroup_Scope
            vFrm1.ClosingPeriod = Me.dtpPeriod.Value.Date

            If vFrm1.ShowDialog = Windows.Forms.DialogResult.OK Then
                vFrm1.dgvItemCost_CostGroup.AllowUserToAddRows = False
                Dummy1 = Me.dgvItemCost_CostGroup.AllowUserToAddRows
                Me.dgvItemCost_CostGroup.AllowUserToAddRows = False
                Me.dgvItemCost_CostGroup.Rows.Clear()

                For Each rowDGV In vFrm1.dgvItemCost_CostGroup.Rows
                    Me.dgvItemCost_CostGroup.Rows.Add()
                    rowdgv1 = Me.dgvItemCost_CostGroup.Rows(Me.dgvItemCost_CostGroup.Rows.Count - 1)

                    For ij As Integer = 0 To Me.dgvItemCost_CostGroup.ColumnCount - 1
                        rowdgv1.Cells(ij).Value = rowDGV.Cells(ij).Value
                    Next
                Next

                Me.dgvItemCost_CostGroup.AllowUserToAddRows = Dummy1

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub settingHeader()
        Try
            Me.TblA.Clear()
            If Me.ObjClosing.GetJml = 0 Then
                Me.ColInAmount.Visible = False
                Me.ColInQty.Visible = False
                Me.ColBeginQty.Visible = False
                Me.ColBeginAmount.Visible = False
                Me.ColOutQty.Visible = False
                Me.ColOutAmount.Visible = False
                'Me.dtpPeriod.Enabled = True
                Me.dtpPeriod.Enabled = False
                Me.dgvItem.ReadOnly = False
                Me.LoadClosing = False
                'Me.dtpPeriod.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
                Me.dtpPeriod.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.ObjCD.GetFirstDateofTheMonth(Me.ObjAccClosing.GetConvertionDate()) - TimeSpan.FromDays(5))
                Me.LoadClosing = True
                'Me.ColDetailCalculation.Visible = False
            Else
                Me.ColInAmount.Visible = True
                Me.ColInQty.Visible = True
                Me.ColBeginQty.Visible = True
                Me.ColBeginAmount.Visible = True
                Me.ColOutQty.Visible = True
                Me.ColOutAmount.Visible = True
                Me.dtpPeriod.Enabled = False
                Me.LoadClosing = False
                Me.dtpPeriod.Value = Me.ObjClosing.GetNextClosingPeriode
                Me.LoadClosing = True
                'sementara
                'Me.ColDetailCalculation.Visible = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub saveTemporaryHeaderData()
        Dim dtLastClosing As Date
        Try
            dtLastClosing = Me.ObjClosing.GetLastClosingPeriod()
            If dtLastClosing <> ObjCD.GetFirstDateofTheMonth(Now) Then
                With Me.ObjClosing
                    .NoTransaksi = Me.TxtClosingNumber.Text
                    .Remarks = Me.TxtRemark.Text
                    .ClosingDate = Me.dtpDate.Value
                    .ClosingPeriod = ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
                End With

                If Me.ObjClosing.SimpanBaru(Me.dgvItem, Me.dgvItemAdjust, Me.UserCode, Me, _
                         Me.TBLItemDetail, Me.TBLBundling1, Me.dgvJurnal, Me.TBLSetting, Me.TBLSettingItem, Me.tblPool, _
                            Me.dgvItemCost_CostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, vTBLCostGroup_CodeNo, vTBLCostGroup_CodeNo2, Me.TblProductQty, True) Then
                    Me.saveTemporaryAktif = True
                    Me.PrimaryTemporary = ObjClosing.GetPrimary(Me.TxtClosingNumber.Text)
                    Me.TxtPrimaryClosing.Text = Me.PrimaryTemporary
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DeleteTemporaryDataHeader()
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.PrimaryTemporary > -1 Then
                DTA.Delete(Me.PrimaryTemporary)
                Me.PrimaryTemporary = -1
                Me.saveTemporaryAktif = False
                Me.TxtPrimaryClosing.Text = Me.PrimaryTemporary
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FillDatatableAdjustment(Optional ByVal vFillFromDB As Boolean = True)
        Dim StartDate, EndDate As Date
        Dim Filter As String = ""
        Dim OldText As String = "Monthly Inventory Closing"
        Dim TotalAdjustment As Double = 0
        Dim rowDGV As DataGridViewRow
        Dim cmd As String = ""
        Dim counter As Double = 0
        Dim DefPrec As Integer = 0
        Dim counter1 As Double = 0
        Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
        Dim TipeStockOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
        Dim ListProduct As New List(Of Double)
        Dim ListProductReject As New List(Of Double)
        Dim Lvl As Integer = 0
        Dim TipeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
        Dim PrimaryClosingSebelum As Double = -1
        Dim TipeItem As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
        Dim PrimaryStockTaking As Double = -1
        Dim TglTaking As Date = #1/1/1950#
        Dim IndexDGV As Double = 0
        Dim AmountDummy As Decimal = 0
        Dim QtyDummy As Decimal = 0
        Dim CostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average
        Dim DummyCostingMethod As String = ""
        Dim DetailAmountDummy As Decimal = 0
        Dim ObjAccClosing As New AccModule.ClsTransaksiClosing
        Dim QtySRDummy As Decimal = 0
        Dim QtyCost As Decimal = 0
        Dim InterimString As String = ":"
        Dim ClosingPertama As Boolean = False

        Try
            ClosingPertama = Me.ObjClosing.ApakahClosingPalingAwal(Me.TxtPrimaryClosing.Text)

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
            EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value)

            If vFillFromDB = False Then
                Me.TblA.Rows.Clear()

                'saliv data dari dgv adjustment
                Me.TblA.Rows.Clear()
                For Each rowDGV In Me.dgvItemAdjust.Rows
                    Me.rowA = Me.TblA.NewAdjustment_CorrectionRow
                    With rowA
                        .Primary = -1
                        .Primary_Transaksi = -1
                        .Document = rowDGV.Cells(Me.ColJenisDocumentAdjustStr.Index).Value
                        .Document_No = rowDGV.Cells(Me.ColDocumentNo.Index).Value
                        .Code_No = rowDGV.Cells(Me.ColCodeNoAdjust.Index).Value
                        .Description = rowDGV.Cells(Me.ColDescAdjust.Index).Value
                        .Qty_In_Adjust = rowDGV.Cells(Me.ColQty.Index).Value
                        .Units = rowDGV.Cells(Me.ColUnitAdjust.Index).Value
                        .Primary_Code_No = rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value
                        .Jenis_Dokumen = rowDGV.Cells(Me.ColJenisDokumen.Index).Value
                        .Primary_Unit = rowDGV.Cells(Me.ColPrimaryUnitAdjust.Index).Value
                        .Primary_Documen_Acuan = rowDGV.Cells(Me.ColPrimaryStockInAdjust.Index).Value
                        .Tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem

                        If .Jenis_Dokumen = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInAdjust Then
                            .Tipe_Transaksi_Asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
                        ElseIf .Jenis_Dokumen = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockTakingInAdjustment Then
                            .Tipe_Transaksi_Asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust
                        ElseIf .Jenis_Dokumen = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama Then
                            .Tipe_Transaksi_Asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR
                        ElseIf .Jenis_Dokumen = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockBundlingAdjustment Then
                            .Tipe_Transaksi_Asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling
                        End If

                        Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString + _
                            " AND [tipe transaksi asal] = " + .Tipe_Transaksi_Asal.ToString
                        Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row

                        .Tanggal_Documen_Acuan = Me.rowItemDetail.tgl_transaksi_asal
                        .Primary_Adjustment_type = Me.rowItemDetail.primary_adjustment_type
                        .Primary_Gain_Account = Me.rowItemDetail.primary_gain_account
                        .Primary_Loss_Account = Me.rowItemDetail.primary_loss_account
                        .Adjust_Type_Desc = Me.ObjStockInOut.GetDescriptionSettingAdjustment(.Primary_Adjustment_type)
                        .remarks_adjust = ""
                        .DeffPrec = Me.ObjProduct.GetAngkaDecimal(.Primary_Code_No)
                        .Remarks = rowDGV.Cells(Me.ColRemarksAdjust.Index).Value

                        .Unit_Price = rowDGV.Cells(Me.ColUnitPriceAdjust.Index).Value
                    End With
                    If Me.rowA.Remarks.Length > 250 Then
                        Me.rowA.Remarks = Me.rowA.Remarks.Substring(0, 250)
                    End If

                    Me.TblA.Rows.Add(Me.rowA)
                Next
            Else
                Me.TblA.Rows.Clear()
                'A. Adjust 
                'revisi: 2018-02-08, perubahan cost berdasarkan nilai cost yang tercatat dalam closing sebelumnya
                'perubahan disini di sini dan form adjustment
                'DTA_HIstoricalPrice.Fill(TBL_HistoricalPrice)
                TBL_HistoricalPrice.Rows.Clear()
                TotalAdjustment = DTAAdjustIn.Fill(TBLInAdjust, TipeTransaksiStockInOut, EndDate, StartDate)
                DetailAmountDummy = 0
                For Each rowAdjust In TBLInAdjust.Rows
                    'Me.Text = OldText + "- Loading adjustment data (" + CStr(Me.dgvItemAdjust.Rows.Count) _
                    '+ " of " + TotalAdjustment.ToString + ")"
                    Me.rowA = Me.TblA.NewAdjustment_CorrectionRow
                    With rowA
                        .Primary = -1
                        .Primary_Transaksi = -1
                        .Document = "Stok In Adjust"
                        .Document_No = rowAdjust.No_Transaksi_Stock_In_Adjust
                        .Code_No = rowAdjust.Product_Code
                        .Description = rowAdjust.Description
                        .Qty_In_Adjust = Me.ObjCD.DigitGrouping(rowAdjust.Qty, rowAdjust._Decimal)
                        .Units = rowAdjust.Unit_Code
                        .Primary_Code_No = rowAdjust.Primary_Product
                        .Jenis_Dokumen = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInAdjust
                        .Primary_Unit = rowAdjust.Primary_Units
                        .Primary_Documen_Acuan = rowAdjust.Primary_Stock_In_Adjust
                        .Tanggal_Documen_Acuan = rowAdjust.Tanggal_Transaksi
                        .Tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                        .Tipe_Transaksi_Asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust

                        .Primary_Adjustment_type = rowAdjust.PrimaryAdjType
                        .Primary_Gain_Account = rowAdjust.Primary_Gain_Account
                        .Primary_Loss_Account = rowAdjust.Primary_Loss_Account
                        .Adjust_Type_Desc = rowAdjust.Adjustment_Type_Desc
                        .remarks_adjust = rowAdjust.Remarks
                        .DeffPrec = rowAdjust._Decimal

                        If ClosingPertama = True Then
                            'closing pertama ambil data dari begining balance
                            .Amount = Me.ObjCD.DigitGrouping(rowAdjust.Amount, Me.DefaultPrecision)
                            .Remarks = "Begining Balance"
                            rowAdjust.Unit_Price = Me.ObjCD.DigitGrouping(rowAdjust.Amount / rowAdjust.Qty, Me.DefaultPrecision)
                            .Unit_Price = rowAdjust.Unit_Price
                        Else
                            .Unit_Price = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                            .Remarks = ""
                            rowAdjust.Unit_Price = 0
                            rowAdjust.Amount = Me.ObjCD.DigitGrouping(rowAdjust.Unit_Price * rowAdjust.Qty, DefPrec)

                            ''jika closing normal, ambil data dari po
                            'cmd = "[Primary Product] =" + rowAdjust.Primary_Product.ToString
                            'TBL_HistoricalPrice.DefaultView.RowFilter = cmd
                            'If TBL_HistoricalPrice.DefaultView.Count > 0 Then
                            '    row_HistoricalPrice = TBL_HistoricalPrice.DefaultView(0).Row
                            '    If Me.ObjCD.GetFirstDateofTheMonth(row_HistoricalPrice.Tanggal_Transaksi) = #2/1/2014# Then
                            '        ''#sementara, hanya untuk closing lmb, seharusnya nilainya 0
                            '        ''#
                            '        .Unit_Price = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                            '        .Remarks = ""
                            '        rowAdjust.Unit_Price = 0
                            '    Else
                            '        .Remarks = "refer to purchase price from: " _
                            '        + row_HistoricalPrice.Supplier_Name + "(" + row_HistoricalPrice.Supplier_Code + _
                            '        ") with aop number: " + row_HistoricalPrice.No_Transaksi + " on " + Format(row_HistoricalPrice.Tanggal_Transaksi, _
                            '        "dd MMM yyyy")
                            '        rowAdjust.Unit_Price = row_HistoricalPrice.At_Price
                            '        .Unit_Price = Me.ObjCD.DigitGrouping(rowAdjust.Unit_Price, Me.DefaultPrecision)
                            '    End If
                            'Else
                            '    ''#sementara, hanya untuk closing lmb, seharusnya nilainya 0
                            '    ''#
                            '    .Unit_Price = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                            '    .Remarks = ""
                            '    rowAdjust.Unit_Price = 0
                            'End If

                        End If
                    End With
                    If Me.rowA.Remarks.Length > 250 Then
                        Me.rowA.Remarks = Me.rowA.Remarks.Substring(0, 250)
                    End If
                    Me.TblA.Rows.Add(Me.rowA)
                Next

                'B. Stock Taking
                If Me.ObjCD.GetFirstDateofTheMonth(Me.ObjAccClosing.GetConvertionDate) <= Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value) Then
                    DTATaking1.Fill(TBLTaking1, EndDate, StartDate)
                    TotalAdjustment = 0
                    For Each rowTaking1 In TBLTaking1.Rows
                        TotalAdjustment += DTATaking2.ScalarQuery_GetJmlByPrimaryTaking(rowTaking1.Primary_Stock_Taking).Value
                    Next
                    'masukkan ke dalam item adjustment
                    counter = 0
                    For Each rowTaking1 In TBLTaking1.Rows
                        DTATaking2.Fill(TBLTaking2, rowTaking1.Primary_Stock_Taking)
                        For Each rowTaking2 In TBLTaking2.Rows
                            counter += 1
                            'Me.Text = OldText & "- Loading adjustment data (" & CStr(counter) & " of " & TotalAdjustment.ToString & ")"
                            Me.rowA = Me.TblA.NewAdjustment_CorrectionRow
                            With rowA
                                .Primary = -1
                                .Primary_Transaksi = -1
                                .Document = "Stock Taking In Adjustment"
                                .Document_No = rowTaking1.Stock_Taking_Number
                                .Code_No = rowTaking2.Product_Code
                                .Description = rowTaking2.Description
                                .Qty_In_Adjust = Me.ObjCD.DigitGrouping(rowTaking2.Qty, rowTaking2._Decimal)
                                .Units = rowTaking2.Unit_Code
                                .Primary_Code_No = rowTaking2.Primary_Product
                                .Primary_Unit = rowTaking2.Primary_Unit
                                .Primary_Documen_Acuan = rowTaking2.Primary_Stock_Taking
                                .Jenis_Dokumen = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockTakingInAdjustment
                                .Tanggal_Documen_Acuan = rowTaking1.Start_Date
                                .Tipe_Transaksi_Asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust

                                .Tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                                .Primary_Adjustment_type = -1
                                .Primary_Gain_Account = -1
                                .Primary_Loss_Account = -1
                                .Adjust_Type_Desc = ""
                                .remarks_adjust = ""
                                .DeffPrec = rowTaking2._Decimal

                                .Unit_Price = Me.ObjCD.DigitGrouping(0, DefPrec)
                                .Remarks = ""

                                'cmd = "[Primary Product] =" + rowTaking2.Primary_Product.ToString
                                'TBL_HistoricalPrice.DefaultView.RowFilter = cmd
                                'If TBL_HistoricalPrice.DefaultView.Count > 0 Then
                                '    row_HistoricalPrice = TBL_HistoricalPrice.DefaultView(0).Row
                                '    If Me.ObjCD.GetFirstDateofTheMonth(row_HistoricalPrice.Tanggal_Transaksi) = #2/1/2014# Then
                                '        .Unit_Price = Me.ObjCD.DigitGrouping(0, DefPrec)
                                '        .Remarks = ""
                                '    Else
                                '        .Remarks = "refer to purchase price from: " _
                                '        + row_HistoricalPrice.Supplier_Name + "(" + row_HistoricalPrice.Supplier_Code + _
                                '        ") with aop number: " + row_HistoricalPrice.No_Transaksi + " on " + Format(row_HistoricalPrice.Tanggal_Transaksi, _
                                '        "dd MMM yyyy")
                                '        .Unit_Price = Me.ObjCD.DigitGrouping(row_HistoricalPrice.At_Price, DefPrec)
                                '    End If
                                'Else
                                '    .Unit_Price = Me.ObjCD.DigitGrouping(0, DefPrec)
                                '    .Remarks = ""
                                'End If
                            End With
                            If Me.rowA.Remarks.Length > 250 Then
                                Me.rowA.Remarks = Me.rowA.Remarks.Substring(0, 250)
                            End If
                            Me.TblA.Rows.Add(Me.rowA)
                        Next
                    Next
                End If


                'C. Stock In Sales Return yang di jual pada bulan yang sama
                TotalAdjustment = DTA_In_SR.Fill(TBL_In_SR, EndDate, StartDate)
                counter = 0
                For Each row_In_SR In TBL_In_SR.Rows
                    'rekomendasikan closing price bulan kemarin
                    counter1 = Me.GetUnitPrice(row_In_SR.Primary_Product, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date _
                    - TimeSpan.FromDays(1)))
                    row_In_SR.Unit_Price = counter1
                Next
                For Each row_In_SR In TBL_In_SR.Rows
                    counter += 1
                    'Me.Text = OldText & "- Loading in sales return data (" + CStr(counter) & " of " & TotalAdjustment.ToString & ")"
                    Me.rowA = Me.TblA.NewAdjustment_CorrectionRow
                    With Me.rowA
                        .Primary = -1
                        .Primary_Transaksi = -1
                        .Document = "Stock In Sales Return"
                        .Document_No = row_In_SR.Stock_In_Sales_Return_Number
                        .Code_No = row_In_SR.Product_Code
                        .Description = row_In_SR.Description
                        .Qty_In_Adjust = Me.ObjCD.DigitGrouping(row_In_SR.Qty, row_In_SR._Decimal)
                        .Units = row_In_SR.Unit_Code
                        .Unit_Price = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Primary_Code_No = row_In_SR.Primary_Product
                        .Primary_Unit = row_In_SR.Primary_Units
                        .Primary_Documen_Acuan = row_In_SR.Primary
                        .Jenis_Dokumen = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama
                        .Tanggal_Documen_Acuan = row_In_SR.Stock_In_Date
                        .Tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                        .Tipe_Transaksi_Asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR

                        .Primary_Adjustment_type = -1
                        .Primary_Gain_Account = -1
                        .Primary_Loss_Account = -1
                        .Adjust_Type_Desc = ""
                        .remarks_adjust = ""
                        .DeffPrec = row_In_SR._Decimal
                        counter1 = row_In_SR.Unit_Price

                        If counter1 = 0 Then
                            .Remarks = ""
                        Else
                            .Unit_Price = Me.ObjCD.DigitGrouping(counter1, DefPrec)
                            .Remarks = "refer to unit price on " + _
                            Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date - TimeSpan.FromDays(1)), _
                            "MMM yyyy")
                        End If
                    End With
                    If Me.rowA.Remarks.Length > 250 Then
                        Me.rowA.Remarks = Me.rowA.Remarks.Substring(0, 250)
                    End If
                    Me.TblA.Rows.Add(Me.rowA)
                Next


                'C.1. Sales return atas transaksi yang dibuat sebelum convertion date
                TotalAdjustment = DTA_In_SR1.Fill(TBL_In_SR1, EndDate, StartDate, ObjAccClosing.GetConvertionDate - TimeSpan.FromDays(1))
                counter = 0
                For Each row_In_SR1 In TBL_In_SR1.Rows
                    'set unit price nol
                    row_In_SR1.Unit_Price = 0
                Next
                For Each row_In_SR1 In TBL_In_SR1.Rows
                    counter += 1
                    'Me.Text = OldText & "- Loading in sales return data (" & CStr(counter) & " of " & TotalAdjustment.ToString + ")"
                    Me.rowA = Me.TblA.NewAdjustment_CorrectionRow
                    With Me.rowA
                        .Primary = -1
                        .Primary_Transaksi = -1
                        .Document = "Stock In Sales Return"
                        .Document_No = row_In_SR1.Stock_In_Sales_Return_Number
                        .Code_No = row_In_SR1.Product_Code
                        .Description = row_In_SR1.Description
                        .Qty_In_Adjust = Me.ObjCD.DigitGrouping(row_In_SR1.Qty, row_In_SR1._Decimal)
                        .Units = row_In_SR1.Unit_Code
                        .Unit_Price = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Primary_Code_No = row_In_SR1.Primary_Product
                        .Primary_Unit = row_In_SR1.Primary_Unit
                        .Primary_Documen_Acuan = row_In_SR1.Primary_Stock_In_Out
                        .Jenis_Dokumen = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama
                        .Remarks = ""
                        .Tanggal_Documen_Acuan = row_In_SR1.Stock_In_Date
                        .Tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                        .Tipe_Transaksi_Asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR

                        .Primary_Adjustment_type = -1
                        .Primary_Gain_Account = -1
                        .Primary_Loss_Account = -1
                        .Adjust_Type_Desc = ""
                        .remarks_adjust = ""
                        .DeffPrec = row_In_SR1._Decimal
                        counter1 = row_In_SR1.Unit_Price
                    End With
                    If Me.rowA.Remarks.Length > 250 Then
                        Me.rowA.Remarks = Me.rowA.Remarks.Substring(0, 250)
                    End If
                    Me.TblA.Rows.Add(Me.rowA)
                Next


                'C.2. Quick Sales return 
                TBL_In_SR1.Rows.Clear()
                TotalAdjustment = DTA_In_SR1.FillBy_QuickSR(TBL_In_SR1, EndDate, StartDate)
                counter = 0
                For Each row_In_SR1 In TBL_In_SR1.Rows
                    'set unit price nol
                    row_In_SR1.Unit_Price = 0
                Next
                For Each row_In_SR1 In TBL_In_SR1.Rows
                    counter += 1
                    'Me.Text = OldText & "- Loading in sales return data (" & CStr(counter) & " of " & TotalAdjustment.ToString & ")"
                    Me.rowA = Me.TblA.NewAdjustment_CorrectionRow
                    With Me.rowA
                        .Primary = -1
                        .Primary_Transaksi = -1
                        .Document = "Stock In Sales Return"
                        .Document_No = row_In_SR1.Stock_In_Sales_Return_Number
                        .Code_No = row_In_SR1.Product_Code
                        .Description = row_In_SR1.Description
                        .Qty_In_Adjust = Me.ObjCD.DigitGrouping(row_In_SR1.Qty, row_In_SR1._Decimal)
                        .Units = row_In_SR1.Unit_Code
                        .Unit_Price = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Primary_Code_No = row_In_SR1.Primary_Product
                        .Primary_Unit = row_In_SR1.Primary_Unit
                        .Primary_Documen_Acuan = row_In_SR1.Primary_Stock_In_Out
                        .Jenis_Dokumen = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama
                        .Remarks = ""
                        .Tanggal_Documen_Acuan = row_In_SR1.Stock_In_Date
                        .Tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                        .Tipe_Transaksi_Asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR
                        .Primary_Adjustment_type = -1
                        .Primary_Gain_Account = -1
                        .Primary_Loss_Account = -1
                        .Adjust_Type_Desc = ""
                        .remarks_adjust = ""
                        .DeffPrec = row_In_SR1._Decimal
                    End With
                    If Me.rowA.Remarks.Length > 250 Then
                        Me.rowA.Remarks = Me.rowA.Remarks.Substring(0, 250)
                    End If
                    Me.TblA.Rows.Add(Me.rowA)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message + ":" + counter.ToString, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ShowDataAdjustment() As Boolean
        Dim I As New FrmTransaksiMonthlyClosing_correction
        Dim RowDGV As New DataGridViewRow
        Dim DeffPrecIDR As Integer = 0

        Try
            'DeffPrecIDR = Me.objCurrency.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency)
            I.dgvItemAdjust.Rows.Clear()
            I.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            I.dgvItemAdjust.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If Me.TblA.Rows.Count > 0 Then
                Me.TblA2.Rows.Clear()
                For Each Me.rowA In Me.TblA.Rows
                    I.dgvItemAdjust.Rows.Add()
                    RowDGV = I.dgvItemAdjust.Rows(I.dgvItemAdjust.Rows.Count - 1)
                    With RowDGV
                        .Cells(I.ColJenisDocumentAdjustStr.Index).Value = rowA.Document
                        .Cells(I.ColDocumentNo.Index).Value = rowA.Document_No
                        .Cells(I.ColCodeNoAdjust.Index).Value = rowA.Code_No
                        .Cells(I.ColDescAdjust.Index).Value = rowA.Description
                        .Cells(I.ColQty.Index).Value = Me.ObjCD.DigitGrouping(rowA.Qty_In_Adjust, rowA.DeffPrec)
                        .Cells(I.ColUnitAdjust.Index).Value = rowA.Units
                        .Cells(I.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(rowA.Unit_Price, Me.DefaultPrecision)
                        .Cells(I.ColPrimaryCodeNoAdjust.Index).Value = rowA.Primary_Code_No
                        .Cells(I.ColPrimaryUnitAdjust.Index).Value = rowA.Primary_Unit
                        .Cells(I.ColPrimaryStockInAdjust.Index).Value = rowA.Primary_Documen_Acuan
                        .Cells(I.ColJenisDokumen.Index).Value = rowA.Jenis_Dokumen
                        .Cells(I.ColTglDocAcuan.Index).Value = rowA.Tanggal_Documen_Acuan
                        .Cells(I.ColTipeItemDetail.Index).Value = rowA.Tipe_item_detail
                        .Cells(I.ColPrimaryGainAccount.Index).Value = rowA.Primary_Gain_Account
                        .Cells(I.ColPrimaryLossAccount.Index).Value = rowA.Primary_Loss_Account
                        .Cells(I.ColPrimaryAdjustmentType.Index).Value = rowA.Primary_Adjustment_type
                        .Cells(I.ColAdjustTypeDesc.Index).Value = rowA.Adjust_Type_Desc
                        .Cells(I.ColRemarksAdjust.Index).Value = rowA.Remarks
                        .Cells(I.ColRemarks_Adjust_detail.Index).Value = rowA.remarks_adjust
                        .Cells(I.ColDeffPrec.Index).Value = rowA.DeffPrec
                        .Cells(I.ColNoAdjust.Index).Value = RowDGV.Index + 1
                        .Cells(I.ColTipeTransaksiAsal.Index).Value = rowA.Tipe_Transaksi_Asal

                        Me.rowA2 = Me.TblA2.NewAdjustment_Correction2Row
                        Me.rowA2.Primary = Me.rowA.Primary
                        Me.rowA2.Primary_Transaksi = Me.rowA.Primary_Transaksi
                        Me.rowA2.Document = Me.rowA.Document
                        Me.rowA2.Document_No = Me.rowA.Document_No
                        Me.rowA2.Code_No = Me.rowA.Code_No
                        Me.rowA2.Description = Me.rowA.Description
                        Me.rowA2.Qty_In_Adjust = Me.rowA.Qty_In_Adjust
                        Me.rowA2.Units = Me.rowA.Units
                        Me.rowA2.Unit_Price = Me.rowA.Unit_Price
                        Me.rowA2.Primary_Code_No = Me.rowA.Primary_Code_No
                        Me.rowA2.Primary_Unit = Me.rowA.Primary_Unit
                        Me.rowA2.Primary_Documen_Acuan = Me.rowA.Primary_Documen_Acuan
                        Me.rowA2.Jenis_Dokumen = Me.rowA.Jenis_Dokumen
                        Me.rowA2.Remarks = Me.rowA.Remarks
                        Me.rowA2.Tanggal_Documen_Acuan = Me.rowA.Tanggal_Documen_Acuan
                        Me.rowA2.Tipe_item_detail = Me.rowA.Tipe_item_detail
                        Me.rowA2.Tipe_Transaksi_Asal = Me.rowA.Tipe_Transaksi_Asal
                        Me.rowA2.Primary_Adjustment_type = Me.rowA.Primary_Adjustment_type
                        Me.rowA2.Primary_Gain_Account = Me.rowA.Primary_Gain_Account
                        Me.rowA2.Primary_Loss_Account = Me.rowA.Primary_Loss_Account
                        Me.rowA2.Adjust_Type_Desc = Me.rowA.Adjust_Type_Desc
                        Me.rowA2.remarks_adjust = Me.rowA.remarks_adjust
                        Me.rowA2.DeffPrec = Me.rowA.DeffPrec
                        Me.TblA2.Rows.Add(Me.rowA2)
                    End With
                Next
                I.DefaultPrecision = Me.DefaultPrecision
                I.UserCode = Me.UserCode
                I.BAckColorAll = Me.BAckColorAll
                I.TblA = Me.TblA

                If Me.TxtRemark.Enabled = True OrElse Me.baru = True Then
                    I.IsReadOnly = False
                Else
                    I.IsReadOnly = True
                End If
                I.BAckColorAll = Me.BAckColorAll
                If Me.baru = True Then
                    I.IsikanDataUnitPrice = True
                Else
                    I.IsikanDataUnitPrice = False
                End If
                I.RefPeriod = Me.dtpPeriod.Value.Date

                If I.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    Me.TblA = I.TblA
                    Dim ObjCustom As New CustomSMESS.ClsIdentifier
                    If ObjCustom.GetIdentifier = "smess_lmb" Then
                        Me.IsiResetAmountLMB()
                    End If
                Else
                    Return False
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsiResetAmountLMB()
        Dim DTA As New Inventory_Module.DataSet2_EndAmountTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TBL As New Inventory_Module.DataSet2_EndAmount.tbltransaksi_stockinout_itemDataTable
        Dim row As Inventory_Module.DataSet2_EndAmount.tbltransaksi_stockinout_itemRow
        Dim Dummy As String = ""
        Dim Idx As Integer = 0
        Dim Amount1, Amount2 As Decimal
        Dim PrimaryStockIn As Double = 0
        Dim DV_Adj As DataRowView
        Dim Idx_Adj As Integer = 0
        Dim JmlAllocated, QtyAdj As Decimal

        Try
            Amount1 = 0
            Amount2 = 0

            'load data adjustment dari lmb, berdasarkan periode closing
            TBL.Rows.Clear()
            DTA.FillBy_DB(TBL, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date))

            'untuk setiap baris data adjustment, pastikan ada padanannya di tbl adjustment, masukan nilai amountnnya
            For zx As Double = TBL.Rows.Count - 1 To 0 Step -1
                row = TBL.Rows(zx)
                If zx = 0 Then
                    PrimaryStockIn = row.Primary_Stock_In
                End If
                Amount1 += row.Amount

                Me.TblA.DefaultView.RowFilter = "[Tipe Transaksi Asal] = " + CInt((ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust)).ToString + _
                " AND [Primary Documen Acuan] = " + row.Primary_Stock_In.ToString + _
                " AND [Primary Code No] = " + row.Primary_Product.ToString

                Idx_Adj = 0
                QtyAdj = 0
                JmlAllocated = 0

                If Me.TblA.DefaultView.Count > 0 Then
                    For Each DV_Adj In Me.TblA.DefaultView
                        Me.rowA = DV_Adj.Row
                        QtyAdj += Me.rowA.Qty_In_Adjust
                    Next
                    For Each DV_Adj In Me.TblA.DefaultView
                        Me.rowA = DV_Adj.Row
                        Idx_Adj += 1

                        If Idx_Adj = Me.TblA.DefaultView.Count Then
                            Me.rowA.Amount = row.Amount - JmlAllocated
                        Else
                            Me.rowA.Amount = Math.Round((row.Amount * Me.rowA.Qty_In_Adjust) / QtyAdj, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                            JmlAllocated += Me.rowA.Amount
                        End If
                        Me.rowA.Unit_Price = Math.Round(Me.rowA.Amount / Me.rowA.Qty_In_Adjust, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                    Next

                    'hapus baris yang sudah di update data amount-nya
                    TBL.Rows.RemoveAt(zx)
                Else
                    MsgBox("")
                End If
            Next

            For Each Me.rowA In Me.TblA.Rows
                If Me.rowA.Primary_Documen_Acuan = PrimaryStockIn And Me.rowA.Tipe_Transaksi_Asal = CInt((ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust)) Then
                    If Me.rowA.Amount = 0 Then
                        'MsgBox("Invalid amount", MsgBoxStyle.Information, "Information")
                    Else
                        Amount2 += Me.rowA.Amount
                    End If
                End If
            Next
            If Amount1 <> Amount2 Then
                MsgBox("Amount data is invalid", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'jika ada yang tidak ada padanannya, maka catat dan berikan peringatan pada akhir fungsi
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    Idx += 1
                    If Dummy = "" Then
                        Dummy = Idx.ToString + ". " + row.Product_Code
                    Else
                        Dummy = Dummy + vbCrLf + Idx.ToString + ". " + row.Product_Code
                    End If
                Next

                MsgBox("The following code no, don't have amount data, please re-check: " + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiKeDGVAdjust()
        Dim RowDGV As New DataGridViewRow
        Dim DeffPrecIDR As Integer = 0

        Try
            ' DeffPrecIDR = Me.objCurrency.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency)

            Me.dgvItemAdjust.Rows.Clear()
            Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemAdjust.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItemAdjust.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            Me.dgvItemAdjust.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing

            'Dim X As New FormTEST
            'X.DataGridView1.DataSource = Nothing
            'X.DataGridView1.DataSource = Me.TblA
            'X.ShowDialog()

            If Me.TblA.Rows.Count > 0 Then
                For Each Me.rowA In Me.TblA.Rows
                    Me.dgvItemAdjust.Rows.Add()
                    RowDGV = Me.dgvItemAdjust.Rows(Me.dgvItemAdjust.Rows.Count - 1)
                    With RowDGV
                        .Cells(Me.ColJenisDocumentAdjustStr.Index).Value = Me.rowA.Document
                        .Cells(Me.ColDocumentNo.Index).Value = Me.rowA.Document_No
                        .Cells(Me.ColCodeNoAdjust.Index).Value = Me.rowA.Code_No
                        .Cells(Me.ColDescAdjust.Index).Value = Me.rowA.Description
                        .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(Me.rowA.Qty_In_Adjust, Me.rowA.DeffPrec)
                        .Cells(Me.ColUnitAdjust.Index).Value = Me.rowA.Units
                        .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(Me.rowA.Unit_Price, Me.DefaultPrecision)
                        .Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = Me.rowA.Primary_Code_No
                        .Cells(Me.ColPrimaryUnitAdjust.Index).Value = Me.rowA.Primary_Unit
                        .Cells(Me.ColPrimaryStockInAdjust.Index).Value = Me.rowA.Primary_Documen_Acuan
                        .Cells(Me.ColJenisDokumen.Index).Value = Me.rowA.Jenis_Dokumen
                        .Cells(Me.ColRemarksAdjust.Index).Value = Me.rowA.Remarks
                    End With
                Next
                Me.ObjCD.SetNumber(Me.dgvItemAdjust)
            End If

            Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtClosingNumber.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtClosingNumber.Text <> "" Then
                If Me.ObjClosing.ApakahBolehEdit(Me.ObjClosing.Primary, UserCode, True, True) = True Then
                    Me.baru = False
                    Me.ObjClosing.CariDetail(Me.ObjClosing.Primary, Me.dgvItem, Me.dgvItemAdjust, _
                        Me.TBLItemDetail, Me.TBLBundling1, Me, Me.dgvJurnal, Me.TBLSettingItem, _
                            Me.dgvItemCost_CostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, vTBLCostGroup_CodeNo, Me, Me.vTBLCostGroup_CodeNo2, Me.TblProductQty)
                    Me.SalinKeField()
                    Me.AktifkanDetail_Edit()
                    Me.dgvItem.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtClosingNumber.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtClosingNumber.Text <> "" Then
                If Me.ObjClosing.ApakahBolehHapus(Me.TxtPrimaryClosing.Text, True, UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjClosing.Hapus(Me.TxtPrimaryClosing.Text, UserCode, True) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub dtpPeriod_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPeriod.Enter
        Try
            Me.ClosingPeriodeBuffer = Me.dtpPeriod.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub dgvItemAdjust_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemAdjust.CellBeginEdit
    '    Try
    '        'saat ini karena adjustment diinput didepan maka, tidak ada yg boleh diedit pada dgv ini
    '        e.Cancel = True

    '        ''yang bisa diedit hanya unit price dan remarks saja, untuk data yg bukan begining balance
    '        'If Me.dgvItemAdjust.Item(Me.ColRemarksAdjust.Index, e.RowIndex).Value = "Begining Balance" And _
    '        '((Me.ObjClosing.GetJml = 0 And Me.baru = True) Or (Me.ObjClosing.GetJml = 1 And Me.baru = False)) And _
    '        'e.ColumnIndex = Me.ColUnitPriceAdjust.Index Then
    '        '    e.Cancel = True
    '        'Else
    '        '    If e.ColumnIndex = Me.ColUnitPriceAdjust.Index Or e.ColumnIndex = Me.ColRemarksAdjust.Index Then
    '        '        Me.Buffer = Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value
    '        '        If e.ColumnIndex = Me.ColUnitPriceAdjust.Index Then
    '        '            Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value)
    '        '        End If
    '        '    Else
    '        '        e.Cancel = True
    '        '    End If
    '        'End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Erro")
    '    End Try
    'End Sub

    'Private Sub dgvItemAdjust_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemAdjust.CellEndEdit
    '    Dim JenisDocument As ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockBundlingAdjustment
    '    Dim vDefPrec As Integer = 0

    '    Try
    '        vDefPrec = Me.objCurrency.GetDecimal_DefaultCurrency

    '        'jika yang diedit unit price
    '        If e.ColumnIndex = Me.ColUnitPriceAdjust.Index Then
    '            'pastikan tidak nothing
    '            If IsNothing(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
    '                Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
    '                Exit Sub
    '            End If

    '            'pastikan nilainya tidak negatif
    '            If Me.ObjCD.CekNonNegatif(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
    '                Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
    '                Exit Sub
    '            End If

    '            'digit grouping
    '            Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = _
    '            Me.ObjCD.DigitGrouping(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefaultPrecision)

    '            'hitung ulang unit price untuk item yang bersangkutan
    '            Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = _
    '            Me.ObjCD.DigitGrouping(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value)
    '            If Me.dgvItemAdjust.Item(Me.ColJenisDokumen.Index, e.RowIndex).Value <> JenisDocument Then
    '                Me.HitungUnitPrice(Me.dgvItemAdjust.Item(Me.ColPrimaryCodeNoAdjust.Index, _
    '                e.RowIndex).Value)

    '                'hitung ulang juga semua unit price yang menggunakan item ybs sebagai
    '                'item bundling before
    '                Dim DTABundling As New DseTransaksiMonthlyClosingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
    '                Dim TBLBundling As New DseTransaksiMonthlyClosing.tbltransaksi_stockbundling_item_afterDataTable
    '                Dim rowBundling As DseTransaksiMonthlyClosing.tbltransaksi_stockbundling_item_afterRow

    '                DTABundling.Fill(TBLBundling, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value.Date, _
    '                Me.dgvItemAdjust.Item(Me.ColPrimaryCodeNoAdjust.Index, e.RowIndex).Value)

    '                For Each rowBundling In TBLBundling.Rows
    '                    'Me.HitungUnitPrice(rowBundling.Primary_Product_After)
    '                    If Me.ListPrimaryCodeNo_Bundling.Contains(rowBundling.Primary_Product_After) = False Then
    '                        Me.ListPrimaryCodeNo_Bundling.Add(rowBundling.Primary_Product_After)
    '                    End If
    '                Next

    '                DTABundling.Dispose()
    '                TBLBundling.Dispose()
    '            Else
    '                'hitung ulang semua product yang terpengaruh dengan transaksi bundling yang bersangkutan
    '                Dim DTABundling As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter
    '                Dim TBLBundling As New DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable
    '                Dim rowBundling As DseTransaksiMonthlyClosing.Tbltrans_bundling_afterRow

    '                DTABundling.FillBy_PrimaryBundling(TBLBundling, _
    '                Me.dgvItemAdjust.Item(Me.ColPrimaryStockInAdjust.Index, e.RowIndex).Value)

    '                For Each rowBundling In TBLBundling.Rows
    '                    'Me.HitungUnitPrice(rowBundling.Primary_Product)
    '                    If Me.ListPrimaryCodeNo_Bundling.Contains(rowBundling.Primary_Product) = False Then
    '                        Me.ListPrimaryCodeNo_Bundling.Add(rowBundling.Primary_Product)
    '                    End If
    '                Next

    '                DTABundling.Dispose()
    '                TBLBundling.Dispose()
    '            End If

    '            'hitung ulang untuk item detail 2
    '            If Me.dgvItemAdjust.Item(Me.ColJenisDokumen.Index, e.RowIndex).Value = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama Then
    '                Me.HitungUlangItem2(Me.dgvItemAdjust.Item(Me.ColPrimaryCodeNoAdjust.Index, e.RowIndex).Value, _
    '                Me.dgvItemAdjust.Item(Me.ColPrimaryStockInAdjust.Index, e.RowIndex).Value, _
    '                ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR, _
    '                Me.GetSubTotalItemAdjust(Me.dgvItemAdjust.Item(Me.ColPrimaryCodeNoAdjust.Index, e.RowIndex).Value, _
    '                Me.dgvItemAdjust.Item(Me.ColPrimaryStockInAdjust.Index, e.RowIndex).Value, _
    '                Me.dgvItemAdjust.Item(Me.ColJenisDokumen.Index, e.RowIndex).Value, vDefPrec))
    '            ElseIf Me.dgvItemAdjust.Item(Me.ColJenisDokumen.Index, e.RowIndex).Value = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInAdjust Then
    '                Me.HitungUlangItem2(Me.dgvItemAdjust.Item(Me.ColPrimaryCodeNoAdjust.Index, e.RowIndex).Value, _
    '                Me.dgvItemAdjust.Item(Me.ColPrimaryStockInAdjust.Index, e.RowIndex).Value, _
    '                ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust, _
    '                Me.GetSubTotalItemAdjust(Me.dgvItemAdjust.Item(Me.ColPrimaryCodeNoAdjust.Index, e.RowIndex).Value, _
    '                Me.dgvItemAdjust.Item(Me.ColPrimaryStockInAdjust.Index, e.RowIndex).Value, _
    '                Me.dgvItemAdjust.Item(Me.ColJenisDokumen.Index, e.RowIndex).Value, vDefPrec))
    '            ElseIf Me.dgvItemAdjust.Item(Me.ColJenisDokumen.Index, e.RowIndex).Value = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockTakingInAdjustment Then
    '                Me.HitungUlangItem2(Me.dgvItemAdjust.Item(Me.ColPrimaryCodeNoAdjust.Index, e.RowIndex).Value, _
    '                Me.dgvItemAdjust.Item(Me.ColPrimaryStockInAdjust.Index, e.RowIndex).Value, _
    '                ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust, _
    '                Me.GetSubTotalItemAdjust(Me.dgvItemAdjust.Item(Me.ColPrimaryCodeNoAdjust.Index, e.RowIndex).Value, _
    '                Me.dgvItemAdjust.Item(Me.ColPrimaryStockInAdjust.Index, e.RowIndex).Value, _
    '                Me.dgvItemAdjust.Item(Me.ColJenisDokumen.Index, e.RowIndex).Value, vDefPrec))
    '            Else
    '                MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
    '                Exit Sub
    '            End If


    '            If CDec(Me.Buffer) <> CDec(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value) Then
    '                Me.UpdateJurnal = True
    '            End If
    '            'Me.BuatJurnal(False)

    '            'cek kembali --> dilakukan ketika tab change/btn ok
    '            Me.UpdateKeDGVItemUtama(Me.dgvItemAdjust.Item(Me.ColPrimaryCodeNoAdjust.Index, e.RowIndex).Value)
    '            'Me.CekKonsistensi()
    '        Else
    '            If IsNothing(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
    '                Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = ""
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Function GetSubTotalItemAdjust(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryTransaksiAsal As Double, _
    ByVal vTipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan, ByVal vDefPrec As Integer) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItemAdjust.Rows
                If rowDGV.Cells(Me.ColJenisDokumen.Index).Value = vTipeTransaksiAsal And _
                rowDGV.Cells(Me.ColPrimaryStockInAdjust.Index).Value = vPrimaryTransaksiAsal And _
                rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = vPrimaryCodeNo Then
                    RetVal += Me.ObjCD.DigitGrouping((rowDGV.Cells(Me.ColQty.Index).Value * _
                    rowDGV.Cells(Me.ColUnitPriceAdjust.Index).Value), vDefPrec)
                End If
            Next
            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub HitungUlangItem2(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, _
    'ByVal vTipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal, ByVal vAmount As Decimal)
    '    Dim CostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average
    '    Dim DummyCostingMethod As String = ""
    '    Dim TBLDummy As New DseTransaksiMonthlyClosing2.TBLDummyDataTable
    '    Dim rowDummy As DseTransaksiMonthlyClosing2.TBLDummyRow
    '    Dim ListPrimary As New List(Of Double)
    '    Dim ListPrimary2 As New List(Of Double)
    '    Dim filter As String = ""
    '    Dim DV As DataRowView
    '    Dim PrimaryDummy As Double = -1
    '    Dim vDefPrec As Integer = 0

    '    Try
    '        vDefPrec = Me.objCurrency.GetDecimal_DefaultCurrency
    '        DummyCostingMethod = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
    '        Core_Module.ClsTransaksiSetting.EnumNamaParameter.CostingMethod)
    '        If DummyCostingMethod = "" Then
    '            CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average
    '        Else
    '            CostingMethod = DummyCostingMethod
    '        End If

    '        '0. update data cost untuk transaksi ybs
    '        filter = "[primary transaksi asal] = " + vPrimaryTransaksi.ToString + " AND [primary code no] = " + vPrimaryProduct.ToString + _
    '        " AND [tipe transaksi asal] = " + CInt(vTipeTransaksiAsal).ToString + " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
    '        Me.TBLItemDetail.DefaultView.RowFilter = filter
    '        If Me.TBLItemDetail.DefaultView.Count = 1 Then
    '            Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
    '            Me.rowItemDetail.amount = Me.ObjCD.DigitGrouping(vAmount, vDefPrec)
    '            Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(rowItemDetail.amount / rowItemDetail.qty, vDefPrec)
    '        Else
    '            MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
    '        End If

    '        '1. hitung ulang cost product ybs
    '        Me.HitungCost(vPrimaryProduct, vDefPrec, CostingMethod)

    '        '2. ambil product2 yang menggunakan item ybs dan turunannya
    '        filter = "[Primary Product Before] = " + vPrimaryProduct.ToString
    '        Me.TBLBundling1.DefaultView.RowFilter = filter
    '        For Each DV In Me.TBLBundling1.DefaultView
    '            Me.rowBundling1 = DV.Row
    '            rowDummy = TBLDummy.FindByPrimary_Product(Me.rowBundling1.Primary_Product_After)

    '            If IsNothing(rowDummy) = True Then
    '                TBLDummy.Rows.Add(Me.rowBundling1.Primary_Product_After)
    '            Else
    '                TBLDummy.Rows.Remove(rowDummy)
    '                TBLDummy.Rows.Add(Me.rowBundling1.Primary_Product_After)
    '            End If
    '            ListPrimary.Add(Me.rowBundling1.Primary_Product_After)
    '        Next

    '        While ListPrimary.Count > 0
    '            ListPrimary2.Clear()

    '            For Each PrimaryDummy In ListPrimary
    '                filter = "[Primary Product Before] = " + PrimaryDummy.ToString
    '                Me.TBLBundling1.DefaultView.RowFilter = filter

    '                For Each DV In Me.TBLBundling1.DefaultView
    '                    Me.rowBundling1 = DV.Row
    '                    rowDummy = TBLDummy.FindByPrimary_Product(Me.rowBundling1.Primary_Product_After)

    '                    If IsNothing(rowDummy) = True Then
    '                        TBLDummy.Rows.Add(Me.rowBundling1.Primary_Product_After)
    '                    Else
    '                        TBLDummy.Rows.Remove(rowDummy)
    '                        TBLDummy.Rows.Add(Me.rowBundling1.Primary_Product_After)
    '                    End If

    '                    ListPrimary2.Add(Me.rowBundling1.Primary_Product_After)
    '                Next
    '            Next

    '            'pindahkan item dari primary 2 ke primary 1
    '            ListPrimary.Clear()
    '            For Each PrimaryDummy In ListPrimary2
    '                ListPrimary.Add(PrimaryDummy)
    '            Next
    '        End While

    '        '3. secara berurutan, hitung ulang product2 turunannya
    '        For i As Double = 0 To TBLDummy.Rows.Count - 1
    '            rowDummy = TBLDummy.Rows(i)
    '            Me.HitungCost(rowDummy.Primary_Product, vDefPrec, CostingMethod)
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        TBLDummy = Nothing
    '        ListPrimary = Nothing
    '        ListPrimary2 = Nothing
    '    End Try
    'End Sub

    ''' <summary>
    ''' menghitung ulang unit price pada dgv item yang primarynya diberikan sesuai dengan 
    ''' yang dimaksud. HANYA BOLEH DIGUNAKAN OLEH FUNGSI --> HitungUlangUnitPrice
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <remarks></remarks>
    'Private Sub HitungUnitPrice(ByVal vPrimaryCodeNo As Double, ByVal DTABefore As DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_beforeTableAdapter, _
    'ByVal DTAAfter As DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter, ByVal TBLBefore As DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable, _
    'ByVal TBLAfter As DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable, ByVal TBL1 As DseTransaksiMonthlyClosing.DataTable16DataTable, _
    'ByVal TBL2 As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable, _
    'ByVal TBLPurchase As DseTransaksiMonthlyClosing.DataTable4DataTable, _
    'ByVal TBLSR As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable, ByVal ObjClosingItem As ClsTransaksiMonthlyClosingItem)
    '    Dim row, rowDGVAdjust As DataGridViewRow
    '    Dim InAmount As Decimal = 0
    '    Dim UnitPrice As Decimal = 0

    '    Try
    '        For Each row In Me.dgvItem.Rows
    '            If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
    '                'ambil in amount
    '                InAmount = ObjClosingItem.GetInAmount_Purchase_SalesReturnBedaPeriode2(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
    '                Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value, Me.TBL_CL, Me.ObjCD, Me.DefaultPrecision)
    '                'tambahkan in amount dengan:
    '                '1. bundling
    '                InAmount = InAmount + Me.GetInAmountFromBundling(vPrimaryCodeNo, DTABefore, DTAAfter, TBLBefore, _
    '                TBLAfter, TBL1, TBL2, TBLPurchase, TBLSR)

    '                '2. in adjustment dari stock in dan stock taking
    '                For Each rowDGVAdjust In Me.dgvItemAdjust.Rows
    '                    If rowDGVAdjust.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = vPrimaryCodeNo Then
    '                        InAmount += Math.Round(rowDGVAdjust.Cells(Me.ColUnitPriceAdjust.Index).Value * _
    '                        rowDGVAdjust.Cells(Me.ColQty.Index).Value, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
    '                    End If
    '                Next
    '                row.Cells(Me.ColInAmount.Index).Value = Me.ObjCD.DigitGrouping(InAmount, Me.DefaultPrecision)

    '                'hitung unit price
    '                If CDec(row.Cells(Me.ColInQty.Index).Value) + CDec(row.Cells(Me.ColBeginQty.Index).Value) <> 0 Then
    '                    'hitung out amount
    '                    row.Cells(Me.ColOutAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(CDbl(row.Cells(Me.ColOutQty.Index).Value) * _
    '                    (CDbl(row.Cells(Me.ColBeginAmount.Index).Value) + CDbl(row.Cells(Me.ColInAmount.Index).Value)) / _
    '                    (CDbl(row.Cells(Me.ColBeginQty.Index).Value) + CDbl(row.Cells(Me.ColInQty.Index).Value)), Me.DefaultPrecision, MidpointRounding.AwayFromZero))

    '                    'hitung end balance
    '                    row.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(CDbl(row.Cells(Me.ColEndQty.Index).Value) * _
    '                    (CDbl(row.Cells(Me.ColBeginAmount.Index).Value) + CDbl(row.Cells(Me.ColInAmount.Index).Value)) / _
    '                    (CDbl(row.Cells(Me.ColBeginQty.Index).Value) + CDbl(row.Cells(Me.ColInQty.Index).Value)), Me.DefaultPrecision, MidpointRounding.AwayFromZero))
    '                Else
    '                    row.Cells(Me.ColOutAmount.Index).Value = 0
    '                    row.Cells(Me.ColEndAmount.Index).Value = 0
    '                End If

    '                'slesai
    '                Exit Sub
    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    '''' <summary>
    '''' menghitung ulang unit price pada dgv item yang primarynya diberikan sesuai dengan 
    '''' yang dimaksud. HANYA BOLEH DIGUNAKAN OLEH FUNGSI --> dgvItemAdjust_CellEndEdit 
    '''' </summary>
    '''' <param name="vPrimaryCodeNo"></param>
    '''' <remarks></remarks>
    'Private Sub HitungUnitPrice(ByVal vPrimaryCodeNo As Double)
    '    Dim row, rowDGVAdjust As DataGridViewRow
    '    Dim InAmount As Decimal = 0
    '    Dim UnitPrice As Decimal = 0
    '    Dim ObjClosingItem As New ClsTransaksiMonthlyClosingItem

    '    'elemen2 yang digunakan untuk pehitungan bundling
    '    Dim DTABefore As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_beforeTableAdapter
    '    Dim DTAAfter As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter
    '    Dim TBLBefore As New DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable
    '    Dim TBLAfter As New DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable

    '    Dim TBLBefore1 As New DseTransaksiMonthlyClosing.DataTable16DataTable
    '    Dim DTABefore1 As New DseTransaksiMonthlyClosingTableAdapters.DataTable16TableAdapter
    '    Dim TBLIn As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable
    '    Dim DTAIn As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item2TableAdapter

    '    Dim TBLPurchase As New DseTransaksiMonthlyClosing.DataTable4DataTable
    '    Dim DTAPurchase As New DseTransaksiMonthlyClosingTableAdapters.DataTable4TableAdapter
    '    Dim TBLSR As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable
    '    Dim DTASR As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item9_1TableAdapter

    '    Try
    '        For Each row In Me.dgvItem.Rows
    '            If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
    '                'ambil in amount
    '                InAmount = ObjClosingItem.GetInAmount_Purchase_SalesReturnBedaPeriode2(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
    '                Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value, Me.TBL_CL, Me.ObjCD, Me.DefaultPrecision)
    '                'tambahkan in amount dengan:
    '                '1. bundling
    '                InAmount = InAmount + Me.GetInAmountFromBundling(vPrimaryCodeNo, DTABefore, DTAAfter, TBLBefore, _
    '                TBLAfter, TBLBefore1, TBLIn, TBLPurchase, TBLSR)

    '                '2. in adjustment dari stock in dan stock taking
    '                For Each rowDGVAdjust In Me.dgvItemAdjust.Rows
    '                    If rowDGVAdjust.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = vPrimaryCodeNo Then
    '                        InAmount += Math.Round(rowDGVAdjust.Cells(Me.ColUnitPriceAdjust.Index).Value * _
    '                        rowDGVAdjust.Cells(Me.ColQty.Index).Value, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
    '                    End If
    '                Next
    '                row.Cells(Me.ColInAmount.Index).Value = Me.ObjCD.DigitGrouping(InAmount, Me.DefaultPrecision)

    '                'hitung unit price
    '                If CDec(row.Cells(Me.ColInQty.Index).Value) + CDec(row.Cells(Me.ColBeginQty.Index).Value) <> 0 Then
    '                    'hitung out amount
    '                    row.Cells(Me.ColOutAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(CDbl(row.Cells(Me.ColOutQty.Index).Value) * _
    '                    (CDbl(row.Cells(Me.ColBeginAmount.Index).Value) + CDbl(row.Cells(Me.ColInAmount.Index).Value)) / _
    '                    (CDbl(row.Cells(Me.ColBeginQty.Index).Value) + CDbl(row.Cells(Me.ColInQty.Index).Value)), Me.DefaultPrecision, MidpointRounding.AwayFromZero))

    '                    'hitung end balance
    '                    row.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(CDbl(row.Cells(Me.ColEndQty.Index).Value) * _
    '                    (CDbl(row.Cells(Me.ColBeginAmount.Index).Value) + CDbl(row.Cells(Me.ColInAmount.Index).Value)) / _
    '                    (CDbl(row.Cells(Me.ColBeginQty.Index).Value) + CDbl(row.Cells(Me.ColInQty.Index).Value)), Me.DefaultPrecision, MidpointRounding.AwayFromZero))
    '                Else
    '                    row.Cells(Me.ColOutAmount.Index).Value = 0
    '                    row.Cells(Me.ColEndAmount.Index).Value = 0
    '                End If

    '                'slesai
    '                Exit Sub
    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjClosingItem = Nothing
    '        DTABefore.Dispose()
    '        DTAAfter.Dispose()
    '        TBLBefore.Dispose()
    '        TBLAfter.Dispose()

    '        TBLBefore1.Dispose()
    '        DTABefore1.Dispose()
    '        TBLIn.Dispose()
    '        DTAIn.Dispose()

    '        TBLPurchase.Dispose()
    '        DTAPurchase.Dispose()
    '        TBLSR.Dispose()
    '        DTASR.Dispose()
    '    End Try
    'End Sub

    'Private Function GetInAmountFromBundling(ByVal vPrimaryCodeNo As Double, _
    'ByVal DTABefore As DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_beforeTableAdapter, ByVal DTAAfter As DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter, ByVal TBLBefore As DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable, _
    'ByVal TBLAfter As DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable, ByVal TBL1 As DseTransaksiMonthlyClosing.DataTable16DataTable, _
    'ByVal TBL2 As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable, _
    'ByVal TBLPurchase As DseTransaksiMonthlyClosing.DataTable4DataTable, _
    'ByVal TBLSR As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable) As Decimal
    '    Dim rowBefore As DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeRow
    '    Dim rowAfter As DseTransaksiMonthlyClosing.Tbltrans_bundling_afterRow
    '    Dim Filter As String = ""
    '    Dim TotalInAmount, InAmountRow As Decimal

    '    Dim Filter2 As String = ""
    '    Dim counter As Double = 0

    '    Dim TBLBefore1 As New DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable
    '    Dim TBLAfter1 As New DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable
    '    DTABefore.FillBy_TanggalSampaiDari(TBLBefore1, _
    '    Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value)


    '    Try
    '        'isikan data dari database
    '        DTAAfter.FillBy_PrimaryProductAfter(TBLAfter, _
    '        vPrimaryCodeNo, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value)

    '        'jumlahkan (unit price * faktor pemberat) untuk setiap item after yang ada
    '        For Each rowAfter In TBLAfter.Rows
    '            Filter = "[Primary Bundling] = " + rowAfter.Primary_Bundling.ToString + ""
    '            TBLBefore.DefaultView.RowFilter = Filter
    '            InAmountRow = 0
    '            counter = 0

    '            While counter < TBLBefore.DefaultView.Count
    '                rowBefore = TBLBefore.DefaultView(counter).Row
    '                Filter2 = "[Primary Transaksi] =" + rowAfter.Primary_Bundling.ToString + _
    '                " AND [Primary Product] =" + rowBefore.Primary_Product.ToString + ""
    '                TBLBundling.DefaultView.RowFilter = Filter2
    '                If TBLBundling.DefaultView.Count = 0 Then
    '                    InAmountRow += Me.GetUnitPrice(rowBefore.Primary_Product, DTABefore, DTAAfter, _
    '                    TBLBefore1, TBLAfter1, Me.ObjCLosingItem, TBL1, TBL2, TBLPurchase, TBLSR) * rowBefore.Qty
    '                Else
    '                    InAmountRow += Me.GetUnitPriceBundling(rowBefore.Primary_Product, _
    '                    rowAfter.Primary_Bundling) * rowBefore.Qty
    '                End If

    '                counter += 1
    '            End While

    '            TotalInAmount += Math.Round(InAmountRow * rowAfter.Faktor_Pemberat / 100, _
    '            Me.DefaultPrecision, MidpointRounding.AwayFromZero)
    '        Next

    '        Return TotalInAmount
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        TBLBefore1.Dispose()
    '        TBLAfter1.Dispose()
    '    End Try
    'End Function

    Private Function GetInAmountFromBundling(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim DTABefore As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_beforeTableAdapter
        Dim DTAAfter As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter
        Dim TBLBefore As New DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable
        Dim TBLAfter As New DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable
        Dim rowBefore As DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeRow
        Dim rowAfter As DseTransaksiMonthlyClosing.Tbltrans_bundling_afterRow
        Dim Filter As String = ""
        Dim TotalInAmount, InAmountRow As Decimal

        Dim DTABundling As New DseTransaksiMonthlyClosingTableAdapters.DataTable12TableAdapter
        Dim TBLBundling As New DseTransaksiMonthlyClosing.DataTable12DataTable
        Dim Filter2 As String = ""
        Dim counter As Double = 0

        Try
            'jika ada bundling yang berputar tanpa batas maka ambil dari dgv
            DTABundling.Fill(TBLBundling, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), _
            Me.dtpPeriod.Value.Date)

            'isikan data dari database
            DTABefore.FillBy_TanggalSampaiDari(TBLBefore, _
            Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value)
            DTAAfter.FillBy_PrimaryProductAfter(TBLAfter, _
            vPrimaryCodeNo, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value)

            'jumlahkan (unit price * faktor pemberat) untuk setiap item after yang ada
            For Each rowAfter In TBLAfter.Rows
                Filter = "[Primary Bundling] = " + rowAfter.Primary_Bundling.ToString + ""
                TBLBefore.DefaultView.RowFilter = Filter
                InAmountRow = 0
                counter = 0

                While counter < TBLBefore.DefaultView.Count
                    rowBefore = TBLBefore.DefaultView(counter).Row
                    Filter2 = "[Primary Transaksi] =" + rowAfter.Primary_Bundling.ToString + _
                    " AND [Primary Product] =" + rowBefore.Primary_Product.ToString + ""
                    TBLBundling.DefaultView.RowFilter = Filter2
                    If TBLBundling.DefaultView.Count = 0 Then
                        InAmountRow += Me.GetUnitPrice(rowBefore.Primary_Product) * rowBefore.Qty
                    Else
                        InAmountRow += Me.GetUnitPriceBundling(rowBefore.Primary_Product, _
                        rowAfter.Primary_Bundling) * rowBefore.Qty
                    End If

                    counter += 1
                End While

                TotalInAmount += Math.Round((InAmountRow * rowAfter.Faktor_Pemberat) / 100, _
                Me.DefaultPrecision, MidpointRounding.AwayFromZero)
            Next

            Return TotalInAmount
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTABefore.Dispose()
            DTAAfter.Dispose()
            TBLBefore.Dispose()
            TBLAfter.Dispose()
            TBLBundling.Dispose()
            DTABundling.Dispose()
        End Try
    End Function

    Public Function GetUnitPriceBundling(ByVal vPrimaryCodeNo As Double, _
    ByVal vPrimaryTransaksiBundling As Double) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim JenisDocument As ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockBundlingAdjustment

        Try
            For Each rowDGV In Me.dgvItemAdjust.Rows
                If rowDGV.Cells(Me.ColJenisDokumen.Index).Value = JenisDocument And _
                rowDGV.Cells(Me.ColPrimaryStockInAdjust.Index).Value = vPrimaryTransaksiBundling And _
                rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = vPrimaryCodeNo Then
                    Return rowDGV.Cells(Me.ColUnitPriceAdjust.Index).Value
                End If
            Next

            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub dtpPeriod_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPeriod.ValueChanged
    '    Try
    '        If Me.ClosingPeriodeBuffer <> Me.dtpPeriod.Value And Me.TxtRemark.Enabled = True And Me.LoadClosing = True Then
    '            If MsgBox("Reload closing data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
    '                Me.LoadClosingData()
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    ''' <summary>
    ''' mencari baris dimana kode product ybs ada
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnFindRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindRow.Click
        Dim i As New Core_Module.frmCariMasterProduct
        Dim PrimaryProduct As Double = -1
        Dim rowdgv As DataGridViewRow

        Try
            'cari product yang akan diset posisinya
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BAckColorAll
            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                If Me.TabControl1.SelectedIndex = 1 Then
                    Me.SearchInAdjust(i.PrimaryProduct, Me.dgvItemAdjust.FirstDisplayedScrollingRowIndex)
                    Exit Sub
                End If

                PrimaryProduct = i.PrimaryProduct
                'cari baris yang berikan product tsb
                For Each rowdgv In Me.dgvItem.Rows
                    If rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value = PrimaryProduct Then
                        'fokus ke baris yang bersangkutan
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = rowdgv.Index
                        Me.dgvItem.ClearSelection()
                        rowdgv.Selected = True
                    End If
                Next

                If Me.TabControl1.SelectedIndex <> 0 Then
                    Me.TabControl1.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SearchInAdjust(ByVal vPrimaryProduct As Double, ByVal vStartIndex As Double)
        Me.dgvItemAdjust.ClearSelection()

        Try
            'cari baris yang berikan product tsb
            For Each rowdgv As DataGridViewRow In Me.dgvItemAdjust.Rows
                If rowdgv.Index >= vStartIndex Then
                    If rowdgv.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = vPrimaryProduct Then
                        'fokus ke baris yang bersangkutan
                        Me.dgvItemAdjust.FirstDisplayedScrollingRowIndex = rowdgv.Index
                        Me.dgvItemAdjust.ClearSelection()
                        rowdgv.Selected = True
                        Exit Sub
                    End If
                End If
            Next

            If Me.dgvItemAdjust.SelectedRows.Count > 0 And vStartIndex > 0 Then
                If MsgBox("Could not find: " + Me.ObjProduct.GetProductCode(vPrimaryProduct) + ". Search from the beginning?", _
                MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.SearchInAdjust(vPrimaryProduct, 0)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_TabIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.TabIndexChanged
        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                Me.btnFindRow.Visible = True
            Else
                Me.btnFindRow.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CekEndAmountSaja() As Boolean
        Dim row As DataGridViewRow
        Dim CodeNoEnd As String = ""
        Dim CodeNo_QtyEnd As String = ""
        Dim CodeNo_AmountEnd As String = ""

        Dim Pass As Boolean = True

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColEndQty.Index).Value > 0 And row.Cells(Me.ColEndAmount.Index).Value = 0 Then
                    If CodeNoEnd = "" Then
                        CodeNoEnd = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoEnd = CodeNoEnd + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDbl(row.Cells(Me.ColEndQty.Index).Value) < 0 Then
                    If CodeNo_QtyEnd = "" Then
                        CodeNo_QtyEnd = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyEnd = CodeNo_QtyEnd + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If row.Cells(Me.ColEndAmount.Index).Value > 0 And row.Cells(Me.ColEndQty.Index).Value = 0 Then
                    If CodeNo_AmountEnd = "" Then
                        CodeNo_AmountEnd = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_AmountEnd = CodeNo_AmountEnd + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If
            Next

            If CodeNo_QtyEnd <> "" Then
                MsgBox("The following code no, have negative end qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyEnd, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If CodeNo_AmountEnd <> "" Then
                MsgBox("The following code no, have positive end amount but zero end qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_AmountEnd, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If CodeNoEnd <> "" Then
                If MsgBox("The following code no, have positive end qty but zero end amount, please make sure you are aware of this." + _
                vbCrLf + CodeNoEnd + vbCrLf + "Continue with this data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                    Pass = False
                End If
            End If

            If Pass = True Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCekData.Click
        Dim row As DataGridViewRow
        'Dim x As Double
        Dim CodeNo As String = ""
        Dim CodeNoBegin As String = ""
        Dim CodeNoOut As String = ""
        Dim CodeNoEnd As String = ""
        Dim CodeNo_QtyBegin As String = ""
        Dim CodeNo_QtyIn As String = ""
        Dim CodeNo_QtyOut As String = ""
        Dim CodeNo_QtyEnd As String = ""
        Dim CodeUnbalanceAmount As String = ""
        Dim CodeUnbalanceQty As String = ""
        Dim WrongEndAmountCalculation As String = ""
        Dim WrongOutAmountCalculation As String = ""
        Dim BeginAmount As String = ""
        Dim InAmount As String = ""
        Dim OutAmount As String = ""
        Dim EndAmount As String = ""

        Dim Pass As Boolean = True
        Dim i As Double = -1

        Try
            If Me.ObjClosing.GetJml = 0 Then
                If Me.CekEndAmountSaja() = True Then
                    MsgBox("Data checked", MsgBoxStyle.Information, "Information")
                End If
                Exit Sub
            ElseIf Me.ObjClosing.GetJml = 1 And Me.baru = False Then
                If Me.CekEndAmountSaja() = True Then
                    MsgBox("Data checked", MsgBoxStyle.Information, "Information")
                End If
                Exit Sub
            End If

            For Each row In Me.dgvItem.Rows
                If Math.Abs((CDec(row.Cells(Me.ColBeginAmount.Index).Value) + CDec(row.Cells(Me.ColInAmount.Index).Value)) - _
                (CDec(row.Cells(Me.ColOutAmount.Index).Value) + CDec(row.Cells(Me.ColEndAmount.Index).Value))) > 2 Then
                    If CodeUnbalanceAmount = "" Then
                        CodeUnbalanceAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeUnbalanceAmount = CodeUnbalanceAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColBeginQty.Index).Value) + CDec(row.Cells(Me.ColInQty.Index).Value) <> _
                CDec(row.Cells(Me.ColOutQty.Index).Value) + CDec(row.Cells(Me.ColEndQty.Index).Value) Then
                    If CodeUnbalanceQty = "" Then
                        CodeUnbalanceQty = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeUnbalanceQty = CodeUnbalanceQty + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If 

                If row.Cells(Me.ColInQty.Index).Value <> 0 And row.Cells(Me.ColInAmount.Index).Value = 0 Then
                    If CodeNo = "" Then
                        CodeNo = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo = CodeNo + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If row.Cells(Me.ColBeginQty.Index).Value <> 0 And row.Cells(Me.ColBeginAmount.Index).Value = 0 Then
                    If CodeNoBegin = "" Then
                        CodeNoBegin = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoBegin = CodeNoBegin + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If row.Cells(Me.ColOutQty.Index).Value <> 0 And row.Cells(Me.ColOutAmount.Index).Value = 0 Then
                    If CodeNoOut = "" Then
                        CodeNoOut = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoOut = CodeNoOut + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If row.Cells(Me.ColEndQty.Index).Value <> 0 And row.Cells(Me.ColEndAmount.Index).Value = 0 Then
                    If CodeNoEnd = "" Then
                        CodeNoEnd = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoEnd = CodeNoEnd + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColBeginQty.Index).Value) < 0 Then
                    If CodeNo_QtyBegin = "" Then
                        CodeNo_QtyBegin = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyBegin = CodeNo_QtyBegin + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColInQty.Index).Value) < 0 Then
                    If CodeNo_QtyIn = "" Then
                        CodeNo_QtyIn = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyIn = CodeNo_QtyIn + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColOutQty.Index).Value) < 0 Then
                    If CodeNo_QtyOut = "" Then
                        CodeNo_QtyOut = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyOut = CodeNo_QtyOut + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColEndQty.Index).Value) < 0 Then
                    If CodeNo_QtyEnd = "" Then
                        CodeNo_QtyEnd = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyEnd = CodeNo_QtyEnd + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If
                i = row.Cells(Me.ColPrimaryCodeNo.Index).Value

                If CDec(row.Cells(Me.ColBeginAmount.Index).Value) <> 0 And CDec(row.Cells(Me.ColBeginQty.Index).Value) = 0 Then
                    If BeginAmount = "" Then
                        BeginAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        BeginAmount = BeginAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColInAmount.Index).Value) <> 0 And CDec(row.Cells(Me.ColInQty.Index).Value) = 0 Then
                    If InAmount = "" Then
                        InAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        InAmount = InAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColOutAmount.Index).Value) <> 0 And CDec(row.Cells(Me.ColOutQty.Index).Value) = 0 Then
                    If OutAmount = "" Then
                        OutAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        OutAmount = OutAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColEndAmount.Index).Value) <> 0 And CDec(row.Cells(Me.ColEndQty.Index).Value) = 0 Then
                    If EndAmount = "" Then
                        EndAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        EndAmount = EndAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If
            Next

            If CodeUnbalanceAmount <> "" Then
                MsgBox("Unbalance amount for: " + CodeUnbalanceAmount, MsgBoxStyle.Information, "Information")
                Pass = False
                Exit Sub
            End If

            If CodeUnbalanceQty <> "" Then
                MsgBox("Unbalance qty for: " + CodeUnbalanceQty, MsgBoxStyle.Information, "Information")
                Pass = False
                Exit Sub
            End If

            If WrongEndAmountCalculation <> "" Then
                MsgBox("Wrong end amount calculation for: " + WrongEndAmountCalculation, MsgBoxStyle.Information, "Information")
                Pass = False
                Exit Sub
            End If

            If WrongOutAmountCalculation <> "" Then
                MsgBox("Wrong out amount calculation for: " + WrongOutAmountCalculation, MsgBoxStyle.Information, "Information")
                Pass = False
                Exit Sub
            End If

            If BeginAmount <> "" Then
                MsgBox("The following code no, have positive beginning amount, but zero beginning qty. Please make sure you are aware of this." + _
                vbCrLf + BeginAmount, MsgBoxStyle.Information, "Information")
                Pass = False
                Exit Sub
            End If
            If InAmount <> "" Then
                MsgBox("The following code no, have positive in amount, but zero in qty. Please make sure you are aware of this." + _
                vbCrLf + InAmount, MsgBoxStyle.Information, "Information")
                Pass = False
                Exit Sub
            End If
            If OutAmount <> "" Then
                MsgBox("The following code no, have positive out amount, but zero out qty. Please make sure you are aware of this." + _
                vbCrLf + OutAmount, MsgBoxStyle.Information, "Information")
                Pass = False
                Exit Sub
            End If
            If EndAmount <> "" Then
                MsgBox("The following code no, have positive end amount, but zero end qty. Please make sure you are aware of this." + _
                vbCrLf + EndAmount, MsgBoxStyle.Information, "Information")
                Pass = False
                Exit Sub
            End If

            If CodeNo_QtyBegin <> "" Then
                MsgBox("The following code no, have negative beginning qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyBegin, MsgBoxStyle.Information, "Information")
                Pass = False
                Exit Sub
            End If
            If CodeNo_QtyIn <> "" Then
                MsgBox("The following code no, have negative in qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyIn, MsgBoxStyle.Information, "Information")
                Pass = False
                Exit Sub
            End If
            If CodeNo_QtyOut <> "" Then
                MsgBox("The following code no, have negative out qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyOut, MsgBoxStyle.Information, "Information")
                Pass = False
                Exit Sub
            End If
            If CodeNo_QtyEnd <> "" Then
                MsgBox("The following code no, have negative end qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyEnd, MsgBoxStyle.Information, "Information")
                Pass = False
                Exit Sub
            End If


            If CodeNoBegin <> "" Then
                MsgBox("The following code no, have positive beginning qty but zero beginning amount, please make sure you are aware of this." + _
                vbCrLf + CodeNoBegin, MsgBoxStyle.Information, "Information")
                Pass = False
                Exit Sub
            End If

            If CodeNo <> "" Then
                MsgBox("The following code no, have positive in qty but zero in amount, please make sure you are aware of this." + _
                vbCrLf + CodeNo, MsgBoxStyle.Information, "Information")
                Exit Sub
                Pass = False
            Else
                If CodeNoOut <> "" Then
                    MsgBox("The following code no, have positive out qty but zero out amount, please make sure you are aware of this." + _
                    vbCrLf + CodeNoOut, MsgBoxStyle.Information, "Information")
                    Pass = False
                    Exit Sub
                End If

                If CodeNoEnd <> "" Then
                    MsgBox("The following code no, have positive out qty but zero end amount, please make sure you are aware of this." + _
                    vbCrLf + CodeNoEnd, MsgBoxStyle.Information, "Information")
                    Pass = False
                    Exit Sub
                End If
            End If

            'If Me.CekKonsistensi1 = False Then
            '    Pass = False
            'End If

            'cek qty beginning, qty in, qty out, qty end, harus lebih besar dari nol, kalo ga salah jadinya

            If Pass = True Then
                MsgBox("Data checked", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' cek yang remarks = "" atau "0" atau "-" harusnya end qty dan qnd amount =0
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ListCodeNo As New List(Of String)
        Dim rowDGV As DataGridViewRow
        Dim Dummy As String = ""

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColRemarks.Index).Value = "" Or rowDGV.Cells(Me.ColRemarks.Index).Value = "0" Or _
                rowDGV.Cells(Me.ColRemarks.Index).Value = "-" Then
                    If CDbl(rowDGV.Cells(Me.ColEndQty.Index).Value) <> 0 Or _
                    CDbl(rowDGV.Cells(Me.ColEndAmount.Index).Value) <> 0 Then
                        ListCodeNo.Add(rowDGV.Cells(Me.ColCodeNo.Index).Value)
                    End If
                End If
            Next

            If ListCodeNo.Count > 0 Then
                For Each CodeNo As String In ListCodeNo
                    If Dummy = "" Then
                        Dummy = CodeNo
                    Else
                        Dummy = Dummy + "," + CodeNo
                    End If
                Next
                MsgBox("Please check: " + Dummy, MsgBoxStyle.Information, "Information")
            Else
                MsgBox("Selesai")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' fungsi ini mengembalikan unit price dari product yang diberikan primarynya
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetUnitPrice(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim ObjClosingItem As New ClsTransaksiMonthlyClosingItem
        Dim InAmount As Decimal = 0
        Dim rowDGVAdjust As DataGridViewRow
        Dim BeginQty, BeginAmount, InQty As Decimal

        Try
            'ambil begin qty, begin amount, in qty, in amount
            ObjClosingItem.Get_BeginAmount_BeginQty_InQty_IAmount(vPrimaryCodeNo, _
            BeginAmount, BeginQty, InQty, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
            Me.dtpPeriod.Value, Me.ObjCD)
            'tambahkan in amount dengan:
            InAmount = ObjClosingItem.GetInAmount_Purchase_SalesReturnBedaPeriode2(vPrimaryCodeNo, _
            Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
            Me.dtpPeriod.Value, Me.TBL_CL, Me.ObjCD, Me.DefaultPrecision)

            '1. bundling
            InAmount = InAmount + Me.GetInAmountFromBundling(vPrimaryCodeNo)
            '2. in adjustment dari stock in dan stock taking
            For Each rowDGVAdjust In Me.dgvItemAdjust.Rows
                If rowDGVAdjust.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = vPrimaryCodeNo Then
                    InAmount += Math.Round(rowDGVAdjust.Cells(Me.ColUnitPriceAdjust.Index).Value * _
                    rowDGVAdjust.Cells(Me.ColQty.Index).Value, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                End If
            Next

            'hitung unit price
            If BeginQty + InQty <> 0 Then
                'hitung unit price
                Return Math.Round((BeginAmount + InAmount) / _
                (BeginQty + InQty), Me.DefaultPrecision, MidpointRounding.AwayFromZero)
            Else
                'tidak ada unit price
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjClosingItem = Nothing
        End Try
    End Function

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim row As DataGridViewRow
        Dim RndNumber As New Random

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            For Each row In Me.dgvItem.Rows
                If CDec(row.Cells(Me.ColEndQty.Index).Value) <> 0 Then
                    row.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(RndNumber.Next(1000, 1000000), 0, MidpointRounding.AwayFromZero))
                End If
            Next

            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'button ini yang digunakan jika ada masalah dengan perhitungan inventory movement, pada "FrmTransaksiMonthlyClosing_DetailCalc",
    'ada coding yang harus diaktifkan pada fungsi form shown
    'Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '    Dim row As DataGridViewRow
    '    Dim TotalIn As Double = 0
    '    Dim TotalPurchase, TotalSalesReturn, TotalBundling, TotalInAdjustment As Decimal
    '    'Dim DTA12 As New DseTransaksiMonthlyClosingTableAdapters.tblcheckTableAdapter
    '    Dim DetailSR As String = ""

    '    Try
    '        For Each row In Me.dgvItem.Rows
    '            TotalIn += CDec(row.Cells(Me.ColInAmount.Index).Value)
    '        Next

    '        MsgBox("Total In: " + Me.ObjCD.DigitGrouping(TotalIn))

    '        For Each row In Me.dgvItem.Rows
    '            Dim i As New FrmTransaksiMonthlyClosing_DetailCalc
    '            i.i = Me
    '            i.PrimaryCodeNo = row.Cells(Me.ColPrimaryCodeNo.Index).Value
    '            i.TanggalDari = Me.dtpPeriod.Value.Date
    '            i.TanggalSampai = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date)

    '            i.ShowDialog()

    '            TotalPurchase += i.TotalP
    '            TotalSalesReturn += i.TotalRtn
    '            TotalBundling += i.TotalBdl
    '            TotalInAdjustment += i.TotalIn

    '            If i.DetailSR <> "" Then
    '                If DetailSR = "" Then
    '                    DetailSR = i.DetailSR
    '                Else
    '                    DetailSR = DetailSR + vbCrLf + i.DetailSR
    '                End If
    '            End If

    '            i.Dispose()

    '            'DTA12.Insert(i.PrimaryCodeNo, i.TotalP)
    '        Next

    '        MsgBox("Total Purchase: " + Me.ObjCD.DigitGrouping(TotalPurchase) + ", " + vbCrLf + _
    '        "Total Sales Return: " + Me.ObjCD.DigitGrouping(TotalSalesReturn) + ", " + vbCrLf + _
    '        "Total Bundling: " + Me.ObjCD.DigitGrouping(TotalBundling) + ", " + vbCrLf + _
    '        "Total In Adjust: " + Me.ObjCD.DigitGrouping(TotalInAdjustment))
    '        If DetailSR <> "" Then
    '            MsgBox(DetailSR)
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TabControl1.SelectedIndex = 0 Then
                    'Dim th As New Threading.Thread(AddressOf HitungUlangUnitPrice)
                    'Control.CheckForIllegalCrossThreadCalls = False
                    'th.Start()
                    'Me.HitungUlangUnitPrice()
                    If Me.UpdateJurnal = True Then
                        'Me.UpdateKeDGVItemUtama()
                        Me.BuatJurnal(False)
                        Me.UpdateJurnal = False
                    End If
                ElseIf Me.TabControl1.SelectedIndex = 2 Then
                    If Me.UpdateJurnal = True Then
                        'Me.UpdateKeDGVItemUtama()
                        Me.BuatJurnal(False)
                        Me.UpdateJurnal = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub HitungUlangUnitPrice()
    '    Try
    '        Me.ControlBox = False
    '        Me.gboDetail.Enabled = False
    '        Dim Th As New Threading.Thread(AddressOf HitungUlangUnitPrice1)

    '        Dim ij As New System.Globalization.DateTimeFormatInfo
    '        Dim jk As New System.Globalization.NumberFormatInfo

    '        Control.CheckForIllegalCrossThreadCalls = False
    '        Th.Start()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub ChangeCulture1()
        Try
            Dim ij As New System.Globalization.DateTimeFormatInfo
            Dim jk As New System.Globalization.NumberFormatInfo

            'cek ui culture dan keyboard input 
            My.Application.ChangeCulture("en-US")
            My.Application.ChangeUICulture("en-US")


            ij.LongDatePattern = "dddd, MMM dd yyyy"
            ij.ShortDatePattern = "MMM dd yyyy"
            jk.CurrencyDecimalSeparator = "."
            jk.CurrencyGroupSeparator = ","

            My.Application.UICulture.NumberFormat = jk
            My.Application.UICulture.DateTimeFormat = ij
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub HitungUlangUnitPrice1()
    '    Dim PrimaryCodeNo As Double = -1
    '    Dim ij As New FrmTransaksiMonthlyClosing_HitungUnitPrice
    '    Dim Dummy As String = ""
    '    Dim ObjClosingItem As New ClsTransaksiMonthlyClosingItem

    '    'elemen2 yang digunakan untuk pehitungan bundling
    '    Dim DTABefore As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_beforeTableAdapter
    '    Dim DTAAfter As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter
    '    Dim TBLBefore As New DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable
    '    Dim TBLAfter As New DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable

    '    Dim TBLBefore1 As New DseTransaksiMonthlyClosing.DataTable16DataTable
    '    Dim DTABefore1 As New DseTransaksiMonthlyClosingTableAdapters.DataTable16TableAdapter
    '    Dim TBLIn As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable
    '    Dim DTAIn As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item2TableAdapter

    '    Dim TBLPurchase As New DseTransaksiMonthlyClosing.DataTable4DataTable
    '    Dim DTAPurchase As New DseTransaksiMonthlyClosingTableAdapters.DataTable4TableAdapter
    '    Dim TBLSR As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable
    '    Dim DTASR As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item9_1TableAdapter

    '    Try

    '        Me.ControlBox = False
    '        Me.gboDetail.Enabled = False
    '        Me.ChangeCulture1()

    '        'isikan data yang akan digunakan untuk perhitungan bundling pada fungsi GetUnitPrice 
    '        DTABefore.FillBy_TanggalSampaiDari(TBLBefore, _
    '        Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value)
    '        DTABefore1.FillBy_ClosingPeriod2(TBLBefore1, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date - TimeSpan.FromDays(1)))
    '        DTAIn.Fill(TBLIn, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), _
    '        Me.dtpPeriod.Value.Date)
    '        DTAPurchase.Fill(TBLPurchase, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), Me.dtpPeriod.Value.Date)
    '        DTASR.FillBy_Period(TBLSR, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), Me.dtpPeriod.Value.Date)

    '        ij.ProgressBar1.Value = 0
    '        ij.ProgressBar1.Maximum = Me.ListPrimaryCodeNo_Bundling.Count
    '        ij.Show()

    '        Dummy = ij.Text

    '        While Me.ListPrimaryCodeNo_Bundling.Count > 0
    '            ij.ProgressBar1.Value += 1
    '            ij.Text = Dummy + " (" + ij.ProgressBar1.Value.ToString + " of " + ij.ProgressBar1.Maximum.ToString + ")"
    '            PrimaryCodeNo = Me.ListPrimaryCodeNo_Bundling.Item(0)
    '            Me.HitungUnitPrice(PrimaryCodeNo, DTABefore, DTAAfter, TBLBefore, TBLAfter, TBLBefore1, TBLIn, TBLPurchase, TBLSR, ObjClosingItem)
    '            Me.ListPrimaryCodeNo_Bundling.RemoveAt(0)
    '        End While

    '        ij.Hide()
    '        Me.ControlBox = True
    '        Me.gboDetail.Enabled = True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ij.Dispose()
    '        DTABefore.Dispose()
    '        DTAAfter.Dispose()
    '        TBLBefore.Dispose()
    '        TBLAfter.Dispose()

    '        TBLBefore1.Dispose()
    '        DTABefore1.Dispose()
    '        TBLIn.Dispose()
    '        DTAIn.Dispose()

    '        TBLPurchase.Dispose()
    '        DTAPurchase.Dispose()
    '        TBLSR.Dispose()
    '        DTASR.Dispose()

    '        ObjClosingItem = Nothing
    '    End Try
    'End Sub

    'Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
    '    Dim rowDGV As DataGridViewRow
    '    Dim DTABundling As New DseTransaksiMonthlyClosingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
    '    Dim TBLBundling As New DseTransaksiMonthlyClosing.tbltransaksi_stockbundling_item_afterDataTable
    '    Dim rowBundling As DseTransaksiMonthlyClosing.tbltransaksi_stockbundling_item_afterRow

    '    Try
    '        Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '        Me.dgvItemAdjust.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

    '        For Each rowDGV In Me.dgvItemAdjust.Rows
    '            If CDec(rowDGV.Cells(Me.ColUnitPriceAdjust.Index).Value) = 0 Then
    '                rowDGV.Cells(Me.ColUnitPriceAdjust.Index).Value = "5000.00"
    '                Me.HitungUnitPrice(rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value)

    '                'hitung ulang juga semua unit price yang menggunakan item ybs sebagai
    '                'item bundling before

    '                DTABundling.Fill(TBLBundling, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value.Date, _
    '                rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value)

    '                For Each rowBundling In TBLBundling.Rows
    '                    If Me.ListPrimaryCodeNo_Bundling.Contains(rowBundling.Primary_Product_After) = False Then
    '                        Me.ListPrimaryCodeNo_Bundling.Add(rowBundling.Primary_Product_After)
    '                    End If
    '                Next
    '            End If
    '        Next

    '        Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '        MsgBox("Start Update")
    '        'If Me.ListPrimaryCodeNo_Bundling.Count <> 0 Then
    '        '    Me.HitungUlangUnitPrice()
    '        'End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTABundling.Dispose()
    '        TBLBundling.Dispose()
    '    End Try
    'End Sub

    'Private Function CekKonsistensi1() As Boolean
    '    Dim rowDGV As DataGridViewRow
    '    Dim Filter As String = ""
    '    Dim TipeItem As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
    '    Dim DV As DataRowView
    '    Dim Qty, Amount As Decimal
    '    Dim DTAInv As New DseTransaksiMonthlyClosing2TableAdapters.TblTransInvItemTableAdapter
    '    Dim TBLInv As New DseTransaksiMonthlyClosing2.TblTransInvItemDataTable
    '    Dim rowInv As DseTransaksiMonthlyClosing2.TblTransInvItemRow
    '    Dim DTAInv_Return As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_salesinvoice_item_returnTableAdapter
    '    Dim TBLInv_Return As New DseTransaksiMonthlyClosing2.tbltrans_salesinvoice_item_returnDataTable
    '    Dim rowInv_Return As DseTransaksiMonthlyClosing2.tbltrans_salesinvoice_item_returnRow
    '    Dim StartDate, EndDate As Date
    '    Dim ij As Integer = 0
    '    Dim TBLInv_Salah As New DseTransaksiMonthlyClosing2.DataTable2DataTable
    '    Dim rowInv_Salah As DseTransaksiMonthlyClosing2.DataTable2Row
    '    Dim Dummy As String = ""

    '    Try
    '        Windows.Forms.Cursor.Current = Cursors.WaitCursor

    '        cek kembali -> ga perlu, yakin bener - iwan -2018-01-30
    '        For Each rowDGV In Me.dgvItem.Rows
    '            'a. begining balance
    '            TipeItem = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
    '            Filter = "[tipe item detail] = " + CInt(TipeItem).ToString + " AND [primary code no] = " + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString
    '            Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '            Qty = 0
    '            Amount = 0
    '            For Each DV In Me.TBLItemDetail.DefaultView
    '                Me.rowItemDetail = DV.Row
    '                Qty += Me.rowItemDetail.qty
    '                Amount += Me.rowItemDetail.amount
    '            Next
    '            If CDec(rowDGV.Cells(Me.ColBeginQty.Index).Value) <> Qty Then
    '                MsgBox("Inconsitent begining balance qty in " + rowDGV.Cells(Me.ColCodeNo.Index).Value, MsgBoxStyle.Information, "Information")
    '                Return False
    '                Exit Function
    '            End If
    '            If CDec(rowDGV.Cells(Me.ColBeginAmount.Index).Value) <> Amount Then
    '                MsgBox("Inconsitent begining balance amount in " + _
    '                rowDGV.Cells(Me.ColCodeNo.Index).Value, MsgBoxStyle.Information, "Information")
    '                Return False
    '                Exit Function
    '            End If
    '            'b. in 
    '            TipeItem = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '            Filter = "[tipe item detail] = " + CInt(TipeItem).ToString + " AND [primary code no] = " + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString
    '            Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '            Qty = 0
    '            Amount = 0
    '            For Each DV In Me.TBLItemDetail.DefaultView
    '                Me.rowItemDetail = DV.Row
    '                Qty += Me.rowItemDetail.qty
    '                Amount += Me.rowItemDetail.amount
    '            Next
    '            If CDec(rowDGV.Cells(Me.ColInQty.Index).Value) <> Qty Then
    '                MsgBox("Inconsitent in qty in " + rowDGV.Cells(Me.ColCodeNo.Index).Value, MsgBoxStyle.Information, "Information")
    '                Return False
    '                Exit Function
    '            End If
    '            If CDec(rowDGV.Cells(Me.ColInAmount.Index).Value) <> Amount Then
    '                MsgBox("Inconsitent in amount in " + _
    '                rowDGV.Cells(Me.ColCodeNo.Index).Value, MsgBoxStyle.Information, "Information")
    '                Return False
    '                Exit Function
    '            End If
    '            'c. out 
    '            TipeItem = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '            Filter = "[tipe item detail] = " + CInt(TipeItem).ToString + " AND [primary code no] = " + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString
    '            Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '            Qty = 0
    '            Amount = 0
    '            For Each DV In Me.TBLItemDetail.DefaultView
    '                Me.rowItemDetail = DV.Row
    '                Qty += Me.rowItemDetail.qty
    '                Amount += Me.rowItemDetail.amount
    '            Next
    '            If CDec(rowDGV.Cells(Me.ColOutQty.Index).Value) <> Qty Then
    '                MsgBox("Inconsitent out qty in " + rowDGV.Cells(Me.ColCodeNo.Index).Value, MsgBoxStyle.Information, "Information")
    '                Return False
    '                Exit Function
    '            End If
    '            If CDec(rowDGV.Cells(Me.ColOutAmount.Index).Value) <> Amount Then
    '                MsgBox("Inconsitent out amount in " + _
    '                rowDGV.Cells(Me.ColCodeNo.Index).Value, MsgBoxStyle.Information, "Information")
    '                Return False
    '                Exit Function
    '            End If
    '            'd. end 
    '            TipeItem = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
    '            Filter = "[tipe item detail] = " + CInt(TipeItem).ToString + " AND [primary code no] = " + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString
    '            Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '            Qty = 0
    '            Amount = 0
    '            For Each DV In Me.TBLItemDetail.DefaultView
    '                Me.rowItemDetail = DV.Row
    '                Qty += Me.rowItemDetail.qty
    '                Amount += Me.rowItemDetail.amount
    '            Next
    '            If CDec(rowDGV.Cells(Me.ColEndQty.Index).Value) <> Qty Then
    '                MsgBox("Inconsitent end qty in " + rowDGV.Cells(Me.ColCodeNo.Index).Value, MsgBoxStyle.Information, "Information")
    '                Return False
    '                Exit Function
    '            End If
    '            If CDec(rowDGV.Cells(Me.ColEndAmount.Index).Value) <> Amount Then
    '                MsgBox("Inconsitent end amount in " + _
    '                rowDGV.Cells(Me.ColCodeNo.Index).Value, MsgBoxStyle.Information, "Information")
    '                Return False
    '                Exit Function
    '            End If
    '        Next

    '        cek sales, dari tabel vs item detail
    '        If Me.ObjCD.GetFirstDateofTheMonth(Me.ObjAccClosing.GetConvertionDate) <= Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value) Then
    '            StartDate = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
    '            EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value)
    '            a. load data invoice
    '            DTAInv.Fill(TBLInv, EndDate, StartDate)
    '            b. load data invoice return
    '            DTAInv_Return.Fill(TBLInv_Return, EndDate, StartDate)
    '            c. kurangkan dgn return
    '            For Each rowInv_Return In TBLInv_Return.Rows
    '                TBLInv.DefaultView.RowFilter = "Primary = " + rowInv_Return.Primary_Sales_Invoice.ToString + " AND [Primary Product] = " + rowInv_Return.Primary_Product.ToString
    '                If TBLInv.DefaultView.Count = 0 Then
    '                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
    '                    Return False
    '                Else
    '                    rowInv = TBLInv.DefaultView(0).Row
    '                    rowInv.Total_Qty -= rowInv_Return.Total_Qty
    '                End If
    '            Next
    '            d. cocokan dengan item detail
    '            For Each rowInv In TBLInv.Rows
    '                rowInv.Qty_Detail = 0
    '            Next
    '            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales).ToString
    '            For Each DV In Me.TBLItemDetail.DefaultView
    '                Me.rowItemDetail = DV.Row
    '                TBLInv.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString + " AND [Primary Product] = " + Me.rowItemDetail.primary_code_no.ToString
    '                If TBLInv.DefaultView.Count > 0 Then
    '                    rowInv = TBLInv.DefaultView(0).Row
    '                    rowInv.Qty_Detail = Me.rowItemDetail.qty
    '                Else
    '                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
    '                    Return False
    '                End If
    '            Next
    '            e. buang item yang cocok atau qty =0
    '            ij = 0
    '            While ij <= TBLInv.Rows.Count - 1
    '                rowInv = TBLInv.Rows(ij)
    '                If rowInv.Total_Qty < 0 Then
    '                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
    '                    Return False
    '                Else
    '                    If rowInv.Total_Qty <> rowInv.Qty_Detail Then
    '                        TBLInv_Salah.DefaultView.RowFilter = "[Primary Inv] = " + rowInv.Primary.ToString
    '                        If TBLInv_Salah.DefaultView.Count = 0 Then
    '                            tambahkan baris baru
    '                            rowInv_Salah = TBLInv_Salah.NewDataTable2Row
    '                            rowInv_Salah.Primary_Inv = rowInv.Primary
    '                            rowInv_Salah.No_Trans = rowInv.No_Transaksi
    '                            TBLInv_Salah.Rows.Add(rowInv_Salah)
    '                        Else
    '                            ij += 1
    '                        End If
    '                    Else
    '                        ij += 1
    '                    End If
    '                End If
    '            End While
    '            f. komplain jika ada yang tidak cocok
    '            If TBLInv_Salah.Rows.Count > 0 Then
    '                For Each rowInv_Salah In TBLInv_Salah.Rows
    '                    If Dummy = "" Then
    '                        Dummy = rowInv_Salah.No_Trans
    '                    Else
    '                        Dummy = Dummy + ", " + rowInv_Salah.No_Trans
    '                    End If
    '                Next
    '                MsgBox("This invoice is consistent: " + Dummy, MsgBoxStyle.Information, "Information")
    '                Return False
    '            End If
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTAInv.Dispose()
    '        TBLInv.Dispose()
    '        DTAInv_Return.Dispose()
    '        TBLInv_Return.Dispose()
    '        TBLInv_Salah.Dispose()
    '    End Try
    'End Function

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbEditInAdjustment.LinkClicked

        Try
            If Me.dgvItemAdjust.Rows.Count > 0 Then
                Me.FillDatatableAdjustment(False)
                If Me.ShowDataAdjustment = True Then
                    Me.IsiKeDGVAdjust()
                    'reset data, supaya user hitung ulang kembali datanya
                    Me.dgvItem.Rows.Clear()
                    Me.dgvJurnal.Rows.Clear()
                    Me.RecalculateNeeded = True
                End If
            Else
                MsgBox("There's no adjustment data to edit.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnRecalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRecalculate.Click
        Try
            If Me.ApakahDataBenar_Overhead(False) = False Then
                Exit Sub
            End If
            If MsgBox("Recalculate cost?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.BgwRecalculate.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwRecalculate_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwRecalculate.DoWork
        Try
            Me.ObjCD.UpdateThreading()
            Me.BgwRecalculate.ReportProgress(1)

            Me.Calc_CurrentStep = 0
            Me.Calc_TotalStep = 25

            Me.GroupBox2.BringToFront()
            'alokasikan overhead
            Me.RecalculateOverhead()
            Me.BgwRecalculate.ReportProgress(5)

            Me.LoadItemDetail2()
            Me.HitungKeDGVUtama1()
            If Me.TBL_Pending1.Rows.Count <> 0 Then
                Exit Sub
            End If

            'jurnal hanya dibuat selain untuk begining balance
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : creating item data summary"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            Me.LblProgress.Visible = True
            Me.ProgressBar1.SendToBack()

            Me.ThreadBuffer = "Start"
            Me.BgwRecalculate.ReportProgress(2)

            While Me.ThreadBuffer = "Start"
                Threading.Thread.Sleep(20000)
            End While
            Me.ProgressBar1.BringToFront()

            If Me.ObjClosing.ApakahClosingPalingAwal(Me.TxtPrimaryClosing.Text) = False Then
                Me.BgwRecalculate.ReportProgress(6)
            End If

            Me.BgwAktifkanDetail_Progress_3()
            'MsgBox(Me.Calc_CurrentStep.ToString)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwRecalculate_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BgwRecalculate.ProgressChanged
        Try
            If e.ProgressPercentage = 1 Then
                Me.BgwAktifkanDetail_Progress_1()
            ElseIf e.ProgressPercentage = 2 Then
                Me.BgwAktifkanDetail_Progress_2_b()
            ElseIf e.ProgressPercentage = 3 Then
                Me.BgwAktifkanDetail_Progress_3()
            ElseIf e.ProgressPercentage = 5 Then
                Me.BgwAktifkanDetail_Progress_5()
            ElseIf e.ProgressPercentage = 6 Then
                Me.BuatJurnal(False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwRecalculate_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwRecalculate.RunWorkerCompleted
        Try
            If Me.TBL_Pending1.Rows.Count = 0 Then
                Me.RecalculateNeeded = False
            Else
                Me.BgwAktifkanDetail_Progress_4()
                Me.dgvJurnal.Rows.Clear()
                Me.dgvItem.Rows.Clear()
                BgwAktifkanDetail_Progress_3()
                Me.RecalculateNeeded = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim i As New FrmTransaksiMonthlyClosing_Export

        Try
            If Me.TxtPrimaryClosing.Text = "" Or Me.TxtPrimaryClosing.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPrimaryClosing.Text = "" Or Me.TxtPrimaryClosing.Text = "-1" Then
                Exit Sub
            End If

            i.PrimaryClosing = Me.TxtPrimaryClosing.Text
            i.ClosingDate = Me.dtpPeriod.Value.Date
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class