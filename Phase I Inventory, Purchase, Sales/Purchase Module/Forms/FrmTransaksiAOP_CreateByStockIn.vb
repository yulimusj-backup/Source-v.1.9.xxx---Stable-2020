Option Explicit On
Option Strict Off

Public Class FrmTransaksiAOP_CreateByStockIn
    Public UserCode As String = ""
    Public BackColorALL As Color = Core_Module.GetSMESSBackColorAll

    Dim DelIndex As Integer = -1
    Dim ObjSupplier As New Core_Module.clsMasterSupplier
    Dim ObjPO As New ClsTransaksiPO
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjLicense As New ClsLicense
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser
    Dim EraseRowIndex As Integer = -1

    Dim Buffer As String = ""

    Private Sub dgvAOP_FilterSupplier_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvAOP_FilterSupplier.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColAOP_SupplierCode.Index Then
                Me.Buffer = Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexSupplier(ByVal vPrimarySupplier As Double, ByVal vRefIndex As Integer) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvAOP_FilterSupplier.Rows
                If rowDGV.Index <> Me.dgvAOP_FilterSupplier.Rows.Count - 1 And rowDGV.Index <> vRefIndex Then
                    If rowDGV.Cells(Me.ColAOP_PrimarySupplier.Index).Value = vPrimarySupplier Then
                        Return rowDGV.Index
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub dgvAOP_FilterSupplier_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAOP_FilterSupplier.CellEndEdit
        Try
            If IsNothing(Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColAOP_SupplierCode.Index Then
                    'kalo kode akun dihapus, maka hapus satu bari semuanya
                    If e.RowIndex <> Me.dgvAOP_FilterSupplier.RowCount - 1 Then
                        Me.DelIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    End If
                    Exit Sub
                End If
            Else
                ''data account number
                If e.ColumnIndex = Me.ColAOP_SupplierCode.Index Then
                    Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ObjSupplier.ApakahSudahAda(Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                        Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_PrimarySupplier.Index, e.RowIndex).Value = Me.ObjSupplier.GetPrimary(Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_SupplierName.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierName(Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_PrimarySupplier.Index, e.RowIndex).Value)

                        If Me.GetIndexSupplier(Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_PrimarySupplier.Index, e.RowIndex).Value, e.RowIndex) <> -1 Then
                            MsgBox("Supplier code duplication, please choose another supplier.", MsgBoxStyle.Information, "Information")
                            Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Not String.IsNullOrEmpty(Me.Buffer) Then
                                Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_PrimarySupplier.Index, e.RowIndex).Value = Me.ObjSupplier.GetPrimary(Me.Buffer)
                                Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_SupplierName.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierName(Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_PrimarySupplier.Index, e.RowIndex).Value)
                            End If
                           
                            If Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.DelIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                                Exit Sub
                            End If
                        End If
                    Else
                        '2. kode akun yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New Core_Module.FrmCariMasterSupplier
                        ij.BackColorAll = Me.BackColorALL
                        ij.txtCari.Text = Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.PrimarySupplier <> -1 Then
                            '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                            Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_PrimarySupplier.Index, e.RowIndex).Value = ij.PrimarySupplier
                            Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjSupplier.GetSupplierCode(ij.PrimarySupplier)
                            Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_SupplierName.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierName(Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_PrimarySupplier.Index, e.RowIndex).Value)

                            If Me.GetIndexSupplier(Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_PrimarySupplier.Index, e.RowIndex).Value, e.RowIndex) <> -1 Then
                                MsgBox("Supplier code duplication, please choose another supplier.", MsgBoxStyle.Information, "Information")
                                Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer

                                If String.IsNullOrEmpty(Me.Buffer) OrElse Me.Buffer = "" Then
                                    Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_PrimarySupplier.Index, e.RowIndex).Value = -1
                                    Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_SupplierName.Index, e.RowIndex).Value = ""
                                Else
                                    Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_PrimarySupplier.Index, e.RowIndex).Value = Me.ObjSupplier.GetPrimary(Me.Buffer)
                                    Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_SupplierName.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierName(Me.dgvAOP_FilterSupplier.Item(Me.ColAOP_PrimarySupplier.Index, e.RowIndex).Value)
                                End If
                              
                                If Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                    Me.DelIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                            Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.dgvAOP_FilterSupplier.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.DelIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(200)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.DelIndex <> -1 Then
                Me.dgvAOP_FilterSupplier.Rows.RemoveAt(Me.DelIndex)
                Me.DelIndex = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiAOP_CreateByStockIn_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.Transaksi_AOP_CreateFromStockIn, _
             Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.Transaksi_AOP_CreateFromStockIn_ShowSupplierMode, _
             Me.cboAOP_ShowSupplierMode.SelectedIndex, Me.UserCode)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiAOP_CreateByStockIn_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Dummy As String = ""

        Try
            Me.dgvAOP_TransStockIn.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorALL
            Me.dgvAOP_FilterSupplier.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorALL
            Me.cboAOP_ShowSupplierMode.SelectedIndex = 0

            'load data setting
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.Transaksi_AOP_CreateFromStockIn, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.Transaksi_AOP_CreateFromStockIn_ShowSupplierMode, Me.UserCode)
            If Dummy <> "" Then
                Me.cboAOP_ShowSupplierMode.SelectedIndex = Dummy
            End If

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiAOP_CreateByStockIn_CreateAOP") = False Then
                Me.btnCreateAOP.Visible = False
                Me.btnClose.Location = New Point(Me.btnAOP_Filter.Location.X, Me.btnAOP_Filter.Location.Y)
                Me.btnClose.Size = New Size(Me.btnAOP_Filter.Size.Width, Me.btnAOP_Filter.Size.Height)
                Me.btnClose.SendToBack()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAOP_PONumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAOP_PONumber.GotFocus
        Try
            Me.TxtAOP_PONumber.BackColor = Me.BackColorALL
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAOP_PONumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAOP_PONumber.LostFocus
        Try
            Me.TxtAOP_PONumber.BackColor = Me.TxtAOP_PrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAOP_CodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAOP_CodeNo.GotFocus
        Try
            Me.TxtAOP_CodeNo.BackColor = Me.BackColorALL
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAOP_CodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAOP_CodeNo.LostFocus
        Try
            Me.TxtAOP_CodeNo.BackColor = Me.TxtAOP_PrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnAOP_FindPO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAOP_FindPO.Click
        Dim i As New FrmCariTransaksiPO
        i.BackColorAll = Me.BackColorALL

        Try
            'i.cboCriteria.SelectedIndex = 0
            If Me.TxtAOP_PONumber.Text <> "" Then
                i.txtCari.Text = Me.TxtAOP_PONumber.Text
            End If

            i.ShowDialog()
            If i.PrimaryPO > 0 Then
                Me.TxtAOP_PrimaryPO.Text = i.PrimaryPO
                Me.TxtAOP_PONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtAOP_PrimaryPO.Text)
            Else
                Me.TxtAOP_PrimaryPO.Text = "-1"
                Me.TxtAOP_PONumber.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtAOP_PONumber_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAOP_PONumber.KeyDown
        Try
            If e.KeyValue = 13 Then
                Dim Primary As Double = ObjPO.GetPrimary(Me.TxtAOP_PONumber.Text)
                If Primary <> -1 Then
                    Me.TxtAOP_PrimaryPO.Text = Primary
                    Me.TxtAOP_PONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtAOP_PrimaryPO.Text)
                Else
                    Me.btnAOP_FindPO.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAOP_PONumber_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtAOP_PONumber.Leave
        Try
            If Me.TxtAOP_PONumber.Text <> "" Then
                Dim Primary As Double = ObjPO.GetPrimary(Me.TxtAOP_PONumber.Text)
                If Primary <> -1 Then
                    Me.TxtAOP_PrimaryPO.Text = Primary
                    Me.TxtAOP_PONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtAOP_PrimaryPO.Text)
                Else
                    Me.btnAOP_FindPO.PerformClick()
                End If
            Else
                Me.TxtAOP_PrimaryPO.Text = "-1"
                Me.TxtAOP_PONumber.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnAOP_FindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAOP_FindCodeNo.Click
        Dim i As New Core_Module.frmCariMasterProduct
        i.BackColorAll = Me.BackColorALL

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtAOP_CodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtAOP_CodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtAOP_PrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtAOP_CodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtAOP_PrimaryCodeNo.Text)
                Me.TxtAOP_Desc.Text = Me.ObjProduct.GetDescription(Me.TxtAOP_PrimaryCodeNo.Text)
            Else
                Me.TxtAOP_PrimaryCodeNo.Text = "-1"
                Me.TxtAOP_CodeNo.Clear()
                Me.TxtAOP_Desc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtAOP_CodeNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAOP_CodeNo.KeyDown
        Try
            If e.KeyValue = 13 Then
                Dim Primary As Double = ObjProduct.GetPrimary(Me.TxtAOP_CodeNo.Text)
                If Primary <> -1 Then
                    Me.TxtAOP_PrimaryCodeNo.Text = Primary
                    Me.TxtAOP_CodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtAOP_PrimaryCodeNo.Text)
                    Me.TxtAOP_Desc.Text = Me.ObjProduct.GetDescription(Me.TxtAOP_PrimaryCodeNo.Text)
                Else
                    Me.btnAOP_FindCodeNo.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAOP_CodeNo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtAOP_CodeNo.Leave
        Try
            If Me.TxtAOP_CodeNo.Text <> "" Then
                Dim Primary As Double = ObjProduct.GetPrimary(Me.TxtAOP_CodeNo.Text)
                If Primary <> -1 Then
                    Me.TxtAOP_PrimaryCodeNo.Text = Primary
                    Me.TxtAOP_CodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtAOP_PrimaryCodeNo.Text)
                    Me.TxtAOP_Desc.Text = Me.ObjProduct.GetDescription(Me.TxtAOP_PrimaryCodeNo.Text)
                Else
                    Me.btnAOP_FindCodeNo.PerformClick()
                End If
            Else
                Me.TxtAOP_PrimaryCodeNo.Text = "-1"
                Me.TxtAOP_CodeNo.Clear()
                Me.TxtAOP_Desc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnAOP_Filter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAOP_Filter.Click
        Dim TBLStockIn As New DseTransaksiAOP_CreatByStockInPurchase.DataTable1DataTable
        Dim DTAStockIn As New DseTransaksiAOP_CreatByStockInPurchaseTableAdapters.DataTable1TableAdapter
        Dim rowStockIn As DseTransaksiAOP_CreatByStockInPurchase.DataTable1Row
        Dim DTAAOP As New DseTransaksiAOP_CreatByStockInPurchaseTableAdapters.tbltrans_aop_itemTableAdapter
        Dim TBLAOP As New DseTransaksiAOP_CreatByStockInPurchase.tbltrans_aop_itemDataTable
        Dim rowAOP As DseTransaksiAOP_CreatByStockInPurchase.tbltrans_aop_itemRow
        Dim rowDGV As DataGridViewRow
        Dim Filter As String = ""
        Dim DTAPO As New DseTransaksiAOP_CreatByStockInPurchaseTableAdapters.DataTable3TableAdapter
        Dim TBLPO As New DseTransaksiAOP_CreatByStockInPurchase.DataTable3DataTable
        Dim TBLPO_D As New DseTransaksiAOP_CreatByStockInPurchase.DataTable3DataTable
        Dim rowPO, rowPO_D As DseTransaksiAOP_CreatByStockInPurchase.DataTable3Row
        Dim Lanjut As Boolean = False

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'mengisikan dgv
            DTAAOP.Connection.ConnectionString = Core_Module.GetConstring
            DTAStockIn.Connection.ConnectionString = Core_Module.GetConstring
            DTAPO.Connection.ConnectionString = Core_Module.GetConstring

            DTAAOP.Fill(TBLAOP, CInt(ClsTransaksiAOP.EnumTipeAop.AOP), Me.ObjCD.GetLasDateOfTheMonth(Me.dtpAOP_To.Value.Date).Date, _
            Me.ObjCD.GetFirstDateofTheMonth(Me.dtpAOP_From.Value.Date).Date)

            DTAStockIn.Fill(TBLStockIn, CInt(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase), _
            Me.dtpAOP_To.Value.Date, Me.dtpAOP_From.Value.Date)

            Me.dgvAOP_TransStockIn.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvAOP_TransStockIn.Rows.Clear()

            Filter = "[Primary] = [Primary]"

            'filter supplier
            Dim FilterS As String = ""
            For Each rowDGV In Me.dgvAOP_FilterSupplier.Rows
                If rowDGV.Index <> Me.dgvAOP_FilterSupplier.Rows.Count - 1 Then
                    If FilterS = "" Then
                        FilterS &= " [Primary Supplier] = " + rowDGV.Cells(Me.ColAOP_PrimarySupplier.Index).Value.ToString
                    Else
                        FilterS &= " OR [Primary Supplier] = " + rowDGV.Cells(Me.ColAOP_PrimarySupplier.Index).Value.ToString
                    End If
                End If
            Next
            If FilterS <> "" Then
                Filter &= " And (" & FilterS & ")"
            End If


            'filter po created by
            Me.DgvUser.AllowUserToAddRows = False
            TBLPO.Rows.Clear()
            TBLPO_D.Rows.Clear()
            If Me.DgvUser.RowCount > 0 Then
                For Each rowdgv In Me.DgvUser.Rows
                    TBLPO_D.Rows.Clear()
                    DTAPO.Fill(TBLPO_D, CInt(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase), _
                    Me.dtpAOP_To.Value.Date, Me.dtpAOP_From.Value.Date, rowDGV.Cells(Me.ColCreatedBy.Index).Value.ToString)

                    For Each rowPO_D In TBLPO_D.Rows
                        rowPO = TBLPO.NewDataTable3Row
                        For ij As Integer = 0 To TBLPO.Columns.Count - 1
                            rowPO.Item(ij) = rowPO_D.Item(ij)
                        Next
                        TBLPO.Rows.Add(rowPO)
                    Next
                Next
            End If


            TBLStockIn.DefaultView.RowFilter = Nothing
            TBLStockIn.DefaultView.RowFilter = Filter
            For Each dv As DataRowView In TBLStockIn.DefaultView
                rowStockIn = dv.Row
                TBLAOP.DefaultView.RowFilter = "[Primary GRN] = " + rowStockIn.Primary_GRN.ToString
                If TBLAOP.DefaultView.Count = 0 Then
                    Lanjut = False
                    If Me.DgvUser.Rows.Count > 0 Then
                        TBLPO.DefaultView.RowFilter = "Primary = " + rowStockIn.Primary.ToString
                        If TBLPO.DefaultView.Count > 0 Then
                            Lanjut = True
                        End If
                    Else
                        Lanjut = True
                    End If

                    'belum ada aop, masukkan ke dalam dgv
                    If Lanjut = True Then
                        Me.dgvAOP_TransStockIn.Rows.Add()
                        rowDGV = Me.dgvAOP_TransStockIn.Rows(Me.dgvAOP_TransStockIn.Rows.Count - 1)
                        With rowDGV
                            .Cells(Me.ColItemStockIn_CreateAOP.Index).Value = Boolean.FalseString
                            .Cells(Me.ColItemStockIn_Date.Index).Value = rowStockIn.Tanggal_Transaksi
                            .Cells(Me.ColItemStockIn_No.Index).Value = Me.dgvAOP_TransStockIn.Rows.Count
                            .Cells(Me.ColItemStockIn_Remarks.Index).Value = rowStockIn.Remarks
                            .Cells(Me.ColItemStockIn_Primary.Index).Value = rowStockIn.Primary
                            .Cells(Me.ColItemStockIn_PrimaryGRN.Index).Value = rowStockIn.Primary_GRN
                            .Cells(Me.ColItemStockIn_PrimarySupplier.Index).Value = rowStockIn.Primary_Supplier
                            If Me.cboAOP_ShowSupplierMode.SelectedIndex = 0 Then
                                .Cells(Me.ColItemStockIn_Supplier.Index).Value = rowStockIn.Supplier_Code
                            ElseIf Me.cboAOP_ShowSupplierMode.SelectedIndex = 1 Then
                                .Cells(Me.ColItemStockIn_Supplier.Index).Value = rowStockIn.Supplier_Name
                            Else
                                .Cells(Me.ColItemStockIn_Supplier.Index).Value = rowStockIn.Supplier_Code + " - " + rowStockIn.Supplier_Name
                            End If
                            .Cells(Me.ColItemStockIn_SupplierDONumber.Index).Value = rowStockIn._Supplier_s_DO_Number
                            .Cells(Me.ColItemStockIn_TransNo.Index).Value = rowStockIn.No_Transaksi
                            .Cells(Me.ColItemStockIn_CreatedBy.Index).Value = rowStockIn.Created_By
                        End With
                    End If
                End If
            Next

            Me.ColItemStockIn_Date.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.dgvAOP_TransStockIn.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvUser.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLStockIn.Dispose()
            DTAStockIn.Dispose()
            DTAAOP.Dispose()
            TBLAOP.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
            DTAPO.Dispose()
            TBLPO.Dispose()
        End Try
    End Sub

    Private Sub llbAOP_CheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAOP_CheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            If MsgBox("Check all data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.dgvAOP_TransStockIn.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                For Each rowDGV In Me.dgvAOP_TransStockIn.Rows
                    rowDGV.Cells(Me.ColItemStockIn_CreateAOP.Index).Value = Boolean.TrueString
                Next
                Me.dgvAOP_TransStockIn.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAOP_UncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAOP_UncheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            If MsgBox("Uncheck all data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.dgvAOP_TransStockIn.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                For Each rowDGV In Me.dgvAOP_TransStockIn.Rows
                    rowDGV.Cells(Me.ColItemStockIn_CreateAOP.Index).Value = Boolean.FalseString
                Next
                Me.dgvAOP_TransStockIn.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAOP_TransStockIn_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAOP_TransStockIn.CellClick
        Dim vPrimary As Double = -1

        Try
            If e.ColumnIndex = -1 Or e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColItemStockIn_TransNo.Index Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockInPurchase") = False Then
                    MsgBox("Insufficient acces.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                Else
                    vPrimary = Me.dgvAOP_TransStockIn.Item(Me.ColItemStockIn_Primary.Index, e.RowIndex).Value

                    Dim ij As New Inventory_Module.FrmTransaksiStockInPurchase2
                    Dim j As New Windows.Forms.Form

                    For Each j In Me.ParentForm.MdiChildren
                        ij = TryCast(j, Inventory_Module.FrmTransaksiStockInPurchase2)
                        If IsNothing(ij) = False Then
                            ij.BringToFront()
                            ij.Focus()
                            ij.ShowTransaction(vPrimary)
                            Exit Sub
                        End If
                    Next

                    Dim i As New Inventory_Module.FrmTransaksiStockInPurchase2
                    ' Increment the caption counter.
                    i.UserCode = Me.UserCode
                    ' Set the caption to be unique.
                    i.BackColorAll = Core_Module.GetSMESSBackColorAll
                    i.MdiParent = Me.ParentForm

                    i.Show()
                    i.ShowTransaction(vPrimary)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCreateAOP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateAOP.Click
        Dim rowDGV As DataGridViewRow
        Dim ApakahCreateAOP As Boolean = False
        Dim TBLRef As New DseTransaksiAOP_CreatByStockInPurchase.DataTable2DataTable
        Dim rowRef, rowRef1 As DseTransaksiAOP_CreatByStockInPurchase.DataTable2Row
        Dim DTA1 As New DseTransaksiAOP_CreatByStockInPurchaseTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TBL1 As New DseTransaksiAOP_CreatByStockInPurchase.tbltransaksi_stockinout_itemDataTable
        Dim row1 As DseTransaksiAOP_CreatByStockInPurchase.tbltransaksi_stockinout_itemRow
        Dim FrmAOP As New FrmTransaksiAOP2
        Dim ApakahLokal As Boolean = False
        Dim IndexStr, IndexStr1, Index2 As Integer
        Dim IndexGRN As Integer = 0
        Dim RefDateCocok As Boolean = False

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            For Each rowDGV In Me.dgvAOP_TransStockIn.Rows
                If rowDGV.Cells(Me.ColItemStockIn_CreateAOP.Index).Value = Boolean.TrueString Then
                    ApakahCreateAOP = True
                    Exit For
                End If
            Next
            If ApakahCreateAOP = False Then
                MsgBox("Please specify at least 1 stock in transaction to create Approval of Purchase.", MsgBoxStyle.Information, "Information")
                Exit Sub
            Else
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                DTA1.Connection.ConnectionString = Core_Module.GetConstring
                'isikan tabel untuk acuan primary currency
                DTA1.Fill(TBL1, Me.dtpAOP_To.Value.Date, Me.dtpAOP_From.Value.Date, _
                CInt(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase))

                'kelompokkan berdasarkan kombinasi supplier, vat exclude, tgl ref dan primary currency
                For Each rowDGV In Me.dgvAOP_TransStockIn.Rows
                    If rowDGV.Cells(Me.ColItemStockIn_CreateAOP.Index).Value = Boolean.TrueString Then
                        TBL1.DefaultView.RowFilter = "[Primary Transaksi] = " + rowDGV.Cells(Me.ColItemStockIn_Primary.Index).Value.ToString
                        If TBL1.DefaultView.Count = 0 Then
                            MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        Else
                            row1 = TBL1.DefaultView(0).Row
                            rowDGV.Cells(Me.ColItemStockIn_PrimaryCurrency.Index).Value = row1.Primary_Currency
                            If row1.Tax_Exclude = True Then
                                TBLRef.DefaultView.RowFilter = "[Primary Supplier] = " + rowDGV.Cells(Me.ColItemStockIn_PrimarySupplier.Index).Value.ToString + _
                                " AND [Primary Currency] = " + row1.Primary_Currency.ToString + " AND [Use PPN] = 1"
                            Else
                                TBLRef.DefaultView.RowFilter = "[Primary Supplier] = " + rowDGV.Cells(Me.ColItemStockIn_PrimarySupplier.Index).Value.ToString + _
                                " AND [Primary Currency] = " + row1.Primary_Currency.ToString + " AND [Use PPN] = 0"
                            End If

                            If TBLRef.DefaultView.Count = 0 Then
                                'tambahkan baris baru
                                rowRef = TBLRef.NewDataTable2Row
                                rowRef.Primary_Currency = row1.Primary_Currency
                                rowRef.Primary_Supplier = rowDGV.Cells(Me.ColItemStockIn_PrimarySupplier.Index).Value
                                If row1.Tax_Exclude = True Then
                                    rowRef.Use_PPN = 1
                                Else
                                    rowRef.Use_PPN = 0
                                End If
                                If row1.Foreign_Supplier = 1 Then
                                    rowRef.Foreign_Supplier = 1
                                Else
                                    rowRef.Foreign_Supplier = 0
                                End If
                                'cek ref date
                                rowRef.RefDate = Today
                                If Format(rowRef.RefDate, "MM yyyy") <> Format(rowDGV.Cells(Me.ColItemStockIn_Date.Index).Value, "MM yyyy") Then
                                    rowRef.RefDate = rowDGV.Cells(Me.ColItemStockIn_Date.Index).Value
                                End If
                                TBLRef.Rows.Add(rowRef)
                            Else
                                'pastikan ada salah satu yang ref datenya pada bulan yang cocok
                                RefDateCocok = False
                                For Each DV As DataRowView In TBLRef.DefaultView
                                    rowRef = DV.Row
                                    If Format(rowRef.RefDate, "MM yyyy") = Format(rowDGV.Cells(Me.ColItemStockIn_Date.Index).Value, "MM yyyy") Then
                                        RefDateCocok = True
                                        If rowRef.RefDate.Date < CDate(rowDGV.Cells(Me.ColItemStockIn_Date.Index).Value).Date Then
                                            'jika ada 2 stock in pada bulan yang sama, asumsi invoice dilakukan pada 
                                            'stock in yang tanggalnya paling bontot
                                            rowRef.RefDate = CDate(rowDGV.Cells(Me.ColItemStockIn_Date.Index).Value).Date
                                        End If
                                        Exit For
                                    End If
                                Next
                                'tambahkan baris baru sesuai ref date baru, jika tidak ada yg cocok
                                If RefDateCocok = False Then
                                    rowRef1 = TBLRef.NewDataTable2Row
                                    rowRef = TBLRef.DefaultView(0).Row
                                    For ij As Integer = 0 To TBLRef.Columns.Count - 1
                                        rowRef1.Item(ij) = rowRef.Item(ij)
                                    Next
                                    rowRef1.RefDate = CDate(rowDGV.Cells(Me.ColItemStockIn_Date.Index).Value).Date
                                    TBLRef.Rows.Add(rowRef1)
                                End If
                            End If
                        End If
                    End If
                Next

                'untuk setiap kombinasi, tampilkan form aop
                Index2 = 0
                For Each rowRef In TBLRef.Rows
                    If FrmAOP.LanjutCreateAOP = False Then
                        Exit For
                    End If
                    FrmAOP.LangsungClose = False

                    Index2 += 1
                    FrmAOP.prg1.Value = Index2
                    FrmAOP.prg1.Maximum = TBLRef.Rows.Count
                    FrmAOP.LblProgress.Text = "Data " + FrmAOP.prg1.Value.ToString + " of " + FrmAOP.prg1.Maximum.ToString

                    FrmAOP.UserCode = Me.UserCode
                    ' Set the caption to be unique.
                    FrmAOP.BackColorAll = Core_Module.GetSMESSBackColorAll
                    If rowRef.Foreign_Supplier = 1 Then
                        FrmAOP.Lokal = True
                    Else
                        FrmAOP.Lokal = False
                    End If

                    FrmAOP.vQuickCreateStockIn = True
                    FrmAOP.Show()
                    FrmAOP.LoadProcedure()
                    FrmAOP.NewData_Procedure()
                    'FrmAOP.Location = New Point(FrmAOP.Size.Width * -1, FrmAOP.Size.Height * -1)

                    'isikan data header
                    FrmAOP.txtPrimarySupplier.Text = rowRef.Primary_Supplier
                    FrmAOP.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(FrmAOP.txtPrimarySupplier.Text)
                    FrmAOP.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(FrmAOP.txtPrimarySupplier.Text)
                    FrmAOP.Objsupplier.IsiCboAddress(FrmAOP.cboSupplierAddress, FrmAOP.txtPrimarySupplier.Text)
                    FrmAOP.cboSupplierAddress.SelectedIndex = 0
                    FrmAOP.IsiCboDeliveryOrder(FrmAOP.txtPrimarySupplier.Text)
                    'secara default tanggal-nya akan dipilih tanggal today
                    FrmAOP.dtpDate.Value = rowRef.RefDate
                    FrmAOP.cboSupplierAddress.Focus()

                    'ganti no transaksi
                    If FrmAOP.baru = True Then
                        FrmAOP.TxtAOP.Text = FrmAOP.ObjAOP.GetNoBaru(FrmAOP.dtpDate.Value, FrmAOP.txtPrimarySupplier.Text)
                    Else
                        ApakahLokal = FrmAOP.Objsupplier.ApakahLokal(FrmAOP.txtPrimarySupplier.Text)
                        If ApakahLokal = True Then
                            IndexStr = FrmAOP.TxtAOP.Text.IndexOf("L")
                            IndexStr1 = FrmAOP.TxtAOP.Text.IndexOf("I", 12)
                        Else
                            IndexStr = FrmAOP.TxtAOP.Text.IndexOf("I", 12)
                            IndexStr1 = FrmAOP.TxtAOP.Text.IndexOf("L")
                        End If
                        If IndexStr = -1 Then
                            FrmAOP.TxtAOP.Text = FrmAOP.ObjAOP.GetNoBaru(FrmAOP.dtpDate.Value, FrmAOP.txtPrimarySupplier.Text)
                        End If
                    End If
                    FrmAOP.TxtNPWPNo.Text = FrmAOP.Objsupplier.GetNPWP(FrmAOP.txtPrimarySupplier.Text)
                    FrmAOP.txtPrimaryCurrency.Text = rowRef.Primary_Currency
                    FrmAOP.txtCurrencyCode.Text = FrmAOP.ObjCurrency.GetCurrencyCode(FrmAOP.txtPrimaryCurrency.Text)
                    FrmAOP.SetTaxRate()
                    FrmAOP.SetDefaultPrecision()
                    'tampilan exchange rate
                    FrmAOP.SetDGVExchangeRate(True)
                    FrmAOP.FilterDP(FrmAOP.txtCurrencyCode.Text)
                    FrmAOP.IsiCBOCurrencyTax()
                    FrmAOP.IsiCBOArAkunByTipePerkiraan1()

                    If rowRef.Use_PPN = 0 Then
                        FrmAOP.chkTax.Checked = False
                    End If


                    'masukkan item
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    For Each rowDGV In Me.dgvAOP_TransStockIn.Rows
                        If rowDGV.Cells(Me.ColItemStockIn_CreateAOP.Index).Value = Boolean.TrueString Then
                            If rowDGV.Cells(Me.ColItemStockIn_PrimarySupplier.Index).Value = rowRef.Primary_Supplier And _
                            rowDGV.Cells(Me.ColItemStockIn_PrimaryCurrency.Index).Value = rowRef.Primary_Currency And _
                            Format(rowDGV.Cells(Me.ColItemStockIn_Date.Index).Value, "MM yyyy") = Format(rowRef.RefDate, "MM yyyy") Then
                                'masukkan ke dalam dgv item
                                IndexGRN = 0
                                While IndexGRN <= FrmAOP.TBL.DefaultView.Count - 1
                                    FrmAOP.rowGRN1 = FrmAOP.TBL.DefaultView(IndexGRN).Row
                                    If FrmAOP.rowGRN1.PrimaryGRN = rowDGV.Cells(Me.ColItemStockIn_PrimaryGRN.Index).Value Then
                                        FrmAOP.cboDO.SelectedIndex = IndexGRN
                                        FrmAOP.InputProcedure()
                                        Exit While
                                    End If
                                    IndexGRN += 1
                                End While

                                'For Each DV As DataRowView In FrmAOP.TBL.DefaultView
                                '    FrmAOP.rowGRN1 = DV.Row
                                '    If FrmAOP.rowGRN1.PrimaryGRN = rowDGV.Cells(Me.ColItemStockIn_PrimaryGRN.Index).Value Then
                                '        FrmAOP.cboDO.SelectedIndex = IndexGRN
                                '        FrmAOP.InputProcedure()
                                '    End If
                                '    IndexGRN += 1
                                'Next
                            End If
                        End If
                    Next

                    'masukkan return
                    FrmAOP.IsiDGVReturn()
                    'masukkan dp
                    FrmAOP.Add_DP_Procedure()
                    'payment term
                    FrmAOP.TxtPaymentsTerms.Text = FrmAOP.GetPaymentDue()

                    Windows.Forms.Cursor.Current = Cursors.Arrow
                    FrmAOP.Hide()
                    FrmAOP.ShowDialog()
                Next
                'jika direct save jalan, maka langsung simpan
            End If

            'load ulang datanya
            Me.btnAOP_Filter.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLRef.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
            FrmAOP = Nothing
        End Try
    End Sub

    Private Sub dgvAOP_TransStockIn_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvAOP_TransStockIn.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvAOP_TransStockIn)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvUser.CellBeginEdit
        Try
            If Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value <> -1 Then
                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                    Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                    Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vKodeUser As String, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvUser.Rows.Count - 2
                RowDGV = Me.DgvUser.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCreatedBy.Index).Value = vKodeUser Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvUser_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvUser.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvUser.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim KodeUser As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCreatedBy.Index Then
                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjUser.ApakahSudahAda(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            KodeUser = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            index = Me.GetIndexItem(KodeUser, e.RowIndex)
                            If index < 0 Then
                                Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value)
                                Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value
                                Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwUser.RunWorkerAsync()
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterUser
                            ij.TxtCari.Text = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.backColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.Kodeuser <> "" Then
                                index = Me.GetIndexItem(ij.Kodeuser, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ij.Kodeuser
                                    Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                    Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(ij.Kodeuser)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwUser.RunWorkerAsync()
                                        Else
                                            Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Exit Sub
                                    Else
                                        Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ij.Kodeuser
                                        Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                        Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(ij.Kodeuser)
                                    End If
                                End If
                            ElseIf ij.Kodeuser = "" Then
                                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        If Me.DgvUser.Rows.Count - 1 <> e.RowIndex Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwUser.RunWorkerAsync()
                                        End If
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                Else
                                    Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.DgvUser)
                    End If
                Else
                    Me.EraseRowIndex = e.RowIndex
                    Me.BgwUser.RunWorkerAsync()
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvUser.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            ' If Me.TxtNoPO.Enabled = True Then
            rowDGV = Me.DgvUser.Rows(Me.DgvUser.RowCount - 1)
            With rowDGV
                ' .Cells(Me.ColN.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColCreatedBy.Index).Value = ""
                .Cells(Me.ColCreatedBy_Name.Index).Value = ""
                .Cells(Me.ColCreatedBy_Primary.Index).Value = "-1"
            End With
            ' End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUser_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwUser.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUser_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwUser.RunWorkerCompleted
        Try
            Me.DgvUser.Rows.RemoveAt(Me.EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class