Option Explicit On
Option Strict Off

Public Class frmCariMasterProduct
    Public PrimaryProduct As Double = -1
    Public TipePencarian As enumTipePencarian = enumTipePencarian.All

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""
    Public TipeProduct As Core_Module.clsMasterProduct.EnumTipeProduct = clsMasterProduct.EnumTipeProduct.Goods

    Dim Obj As New clsMasterProduct
    Dim ObjCD As New Core_Module.ClsCekData
    Dim TBL As New dseMasterProduct.tblmaster_product_view0DataTable
    Dim TBLPOS As New dseMasterProduct.tblmaster_product_view01DataTable
    Dim KriteriaCari As ModGlobalVariable.EnumSearchCriteria = EnumSearchCriteria.ByCode
    Dim DTA As New dseMasterProductTableAdapters.tblmaster_product_view0TableAdapter
    Dim DTAPOS As New dseMasterProductTableAdapters.tblmaster_product_view01TableAdapter
    'Dim Tblmaster_product_view0TableAdapter As New dseMasterProductTableAdapters.tblmaster_product_view0TableAdapter
    'Dim tblmaster_product_view0 As New dseMasterProduct.tblmaster_product_view0DataTable
    Dim Tblmaster_product_view01TableAdapter As New dseMasterProductTableAdapters.tblmaster_product_view01TableAdapter
    Dim tblmaster_product_view01 As New dseMasterProduct.tblmaster_product_view01DataTable
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser

    'column size acuan
    Dim TBLColSize As New DseMasterColSize.tblmaster_columns_sizeDataTable
    Dim TBLColSize1 As New DseMasterColSize.tblmaster_columns_sizeDataTable
    Dim rowColSize As DseMasterColSize.tblmaster_columns_sizeRow
    Dim ObjColSize As New Core_Module.ClsMasterColSize
    Dim SedangIsiDb As Boolean = False
    Dim ColumnChanged As Boolean = False

    Dim Tblmaster_product_view0TableAdapter As New dseMasterProduct2TableAdapters.tblmaster_product_view0TableAdapter
    Dim tblmaster_product_view0 As New dseMasterProduct2.tblmaster_product_view0DataTable

    Public Enum enumTipePencarian As Integer
        All
        Goods_Only
        Location_Control_Only
    End Enum

    Private Sub Pilih()
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.PrimaryProduct = Me.dgvHasil.SelectedRows(0).Cells(0).Value
            Else
                Me.PrimaryProduct = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmCariMasterCustomer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'TODO: This line of code loads data into the 'DseMasterProduct.tblmaster_product_view0' table. You can move, or remove it, as needed.
        Dim Dummy As String = ""

        Try
            Me.ShowIcon = False
            'Me.dgvHasil.Size = Me.DGVHasil1.Size
            'Me.dgvHasil.Anchor = Me.DGVHasil1.Anchor

            If Me.UserCode = "" Then
                Me.UserCode = My.Settings.user_code
            End If
            Me.ObjColSize.IsiTBL(Me.UserCode, ClsMasterColSize.EnumTipeTransaksi.MasterProduct, Me.TBLColSize)
            Me.ObjColSize.IsiTBL(Me.UserCode, ClsMasterColSize.EnumTipeTransaksi.MasterProduct1, Me.TBLColSize1)

            Me.txtCari.Focus()
            'Me.cboCriteria.SelectedIndex = 0
            'Me.KriteriaCari = EnumSearchCriteria.ByCode
            Dummy = Me.ObjSettingUser.GetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.MasterProduct, _
            ClsTransaksiSettingUser.EnumNamaParameter.MasterProduct_DefaultFindBy, Me.UserCode)
            If Dummy = "" Or Dummy = CInt(Core_Module.clsMasterProduct.EnumDefaultFindBy_SettingUser.ByGlobalSetting).ToString Then
                Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.MasterProduct, ClsTransaksiSetting.EnumNamaParameter.MasterProduct_DefaultFindBy)
                If Dummy = "" Then
                    Me.cboCriteria.SelectedIndex = 0
                Else
                    If CInt(Dummy) = CInt(Core_Module.clsMasterProduct.EnumDefaultFindBy.ByDesc) Then
                        Me.cboCriteria.SelectedIndex = 1
                    Else
                        Me.cboCriteria.SelectedIndex = 0
                    End If
                End If
            Else
                If CInt(Dummy) = CInt(Core_Module.clsMasterProduct.EnumDefaultFindBy_SettingUser.ByDesc) Then
                    Me.cboCriteria.SelectedIndex = 1
                Else
                    Me.cboCriteria.SelectedIndex = 0
                End If
            End If

            If Me.txtCari.Text <> "" Then
                Me.btnFind.PerformClick()
            End If
            Me.dgvHasil.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCari.KeyDown
        Try
            If e.KeyValue = 13 And Me.txtCari.Text <> "" Then
                Me.IsiDatabase(Me.TipePencarian)
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.ClearSelection()
                End If
            ElseIf e.KeyValue = 40 And Me.dgvHasil.RowCount > 0 Then
                Me.dgvHasil.Focus()
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.Rows.Item(0).Selected = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCari_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Try
            Me.IsiDatabase(Me.TipePencarian)
            If Me.dgvHasil.Rows.Count > 0 Then
                Me.dgvHasil.ClearSelection()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHasil.CellContentDoubleClick
        Try
            Me.Pilih()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvHasil.ColumnWidthChanged
        Try
            If Me.SedangIsiDb = False Then
                Me.ColumnChanged = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil1_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles DGVHasil1.ColumnWidthChanged
        Try
            If Me.SedangIsiDb = False Then
                Me.ColumnChanged = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvHasil.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.Pilih()
                Me.Close()
            ElseIf e.KeyValue = 38 Then
                If Me.dgvHasil.SelectedRows.Count > 0 And Me.dgvHasil.SelectedRows(0).Index = 0 Then
                    Me.txtCari.Focus()
                End If
            ElseIf e.KeyValue = 9 Then
                Me.txtCari.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DGVHasil1.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.Pilih()
                Me.Close()
            ElseIf e.KeyValue = 38 Then
                If Me.dgvHasil.SelectedRows.Count > 0 And Me.dgvHasil.SelectedRows(0).Index = 0 Then
                    Me.txtCari.Focus()
                End If
            ElseIf e.KeyValue = 9 Then
                Me.txtCari.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Me.Pilih()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnKeluar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.PrimaryProduct = -1
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCari.TextChanged
        Try
            If Me.Obj.GetJml(Me.txtCari.Text, Me.KriteriaCari) < 150 Then
                Me.IsiDatabase(Me.TipePencarian)
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.ClearSelection()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCari_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCari.GotFocus
        txtCari.BackColor = BackColorAll
    End Sub

    Private Sub TxtCari_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCari.LostFocus
        txtCari.BackColor = Me.txtprimarygroup.BackColor
    End Sub

    Private Sub cboCriteria_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCriteria.SelectedIndexChanged
        Try
            If Me.cboCriteria.SelectedIndex = 0 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.KriteriaCari = EnumSearchCriteria.ByCode
            ElseIf Me.cboCriteria.SelectedIndex = 1 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.KriteriaCari = EnumSearchCriteria.ByDescription
            ElseIf Me.cboCriteria.SelectedIndex = 2 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.KriteriaCari = EnumSearchCriteria.ByProductSpec
            ElseIf Me.cboCriteria.SelectedIndex = 3 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.KriteriaCari = EnumSearchCriteria.ByBarcode
            ElseIf Me.cboCriteria.SelectedIndex = 4 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.KriteriaCari = EnumSearchCriteria.ByTaxProductCode
            ElseIf Me.cboCriteria.SelectedIndex = 5 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.KriteriaCari = EnumSearchCriteria.ByTaxProductDesc
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvHasil.MouseDoubleClick
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.Pilih()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DGVHasil1.MouseDoubleClick
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.Pilih()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_RowHeaderMouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvHasil.RowHeaderMouseDoubleClick
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.Pilih()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil1_RowHeaderMouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGVHasil1.RowHeaderMouseDoubleClick
        Try
            If Me.DGVHasil1.SelectedRows.Count > 0 Then
                Me.Pilih()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiDatabase(ByVal vTipePencarian As enumTipePencarian)
        Dim QueryResult As Double = 0
        Dim StartT As Date = Now

        Try
            SedangIsiDb = True
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvHasil.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DGVHasil1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If Me.TipePencarian = enumTipePencarian.All Then
                If Me.KriteriaCari = EnumSearchCriteria.ByCode Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_CodeLIKE(tblmaster_product_view0, Me.txtCari.Text + "%")
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByDescription Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_DESCLike(tblmaster_product_view0, Me.txtCari.Text + "%")
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByProductSpec Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_ProductSpecLIKE(tblmaster_product_view0, Me.txtCari.Text + "%")
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByBarcode Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_BarcodeLIKE(tblmaster_product_view0, Me.txtCari.Text + "%")
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByTaxProductCode Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_TaxProductRefCodeLIKE(tblmaster_product_view0, Me.txtCari.Text + "%")
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByTaxProductDesc Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_TaxProductRefDescLIKE(tblmaster_product_view0, Me.txtCari.Text + "%")
                End If
                Me.dgvHasil.DataSource = Me.tblmaster_product_view0.DefaultView
                Me.dgvHasil.BringToFront()
            ElseIf Me.TipePencarian = enumTipePencarian.Location_Control_Only Then
                If Me.KriteriaCari = EnumSearchCriteria.ByCode Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_CodeLIKE_LocationControlled(tblmaster_product_view0, Me.txtCari.Text + "%")
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByDescription Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_DescLIKE_LocationControlled(tblmaster_product_view0, Me.txtCari.Text + "%")
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByProductSpec Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_SpecLIKE_LocationControlled(tblmaster_product_view0, Me.txtCari.Text + "%")
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByBarcode Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_BarcodeLocationControl(tblmaster_product_view0, Me.txtCari.Text + "%")
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByTaxProductCode Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_TaxProductRefCodeLIKE_LocationControlled(tblmaster_product_view0, Me.txtCari.Text + "%")
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByTaxProductDesc Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_TaxProductRefDescLIKE_LocationControlled(tblmaster_product_view0, Me.txtCari.Text + "%")
                End If
                Me.dgvHasil.DataSource = Me.tblmaster_product_view0.DefaultView
                Me.dgvHasil.BringToFront()
            ElseIf Me.TipePencarian = enumTipePencarian.Goods_Only Then
                If Me.KriteriaCari = EnumSearchCriteria.ByCode Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_CodeLIKE_Type(tblmaster_product_view0, Me.txtCari.Text + "%", CInt(Me.TipeProduct))
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByDescription Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_DescLIKE_Type(tblmaster_product_view0, Me.txtCari.Text + "%", CInt(Me.TipeProduct))
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByProductSpec Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_SpecLIKE_Type(tblmaster_product_view0, Me.txtCari.Text + "%", CInt(Me.TipeProduct))
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByBarcode Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_BarcodeLIKE_Type(tblmaster_product_view0, Me.txtCari.Text + "%", CInt(Me.TipeProduct))
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByTaxProductCode Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_TaxProductRefCodeLiKE_Type(tblmaster_product_view0, Me.txtCari.Text + "%", CInt(Me.TipeProduct))
                ElseIf Me.KriteriaCari = EnumSearchCriteria.ByTaxProductDesc Then
                    QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_taxProductRefDescLIKE_Type(tblmaster_product_view0, Me.txtCari.Text + "%", CInt(Me.TipeProduct))
                End If
                Me.dgvHasil.DataSource = Me.tblmaster_product_view0.DefaultView
                Me.dgvHasil.BringToFront()
            End If

            'update kolom
            Me.dgvHasil.Columns(0).Visible = False
            Me.dgvHasil.Columns(0).Frozen = True
            Me.dgvHasil.Columns(1).Frozen = True
            If Me.TBLColSize.Rows.Count = 0 Then
                Me.dgvHasil.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.dgvHasil.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            End If

            Dim TS As TimeSpan = Now - StartT
            Me.SetAutoSize1(TS)
            TS = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    Me.tslQuery.Text = CStr(QueryResult) + " result(s) in " + Me.ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    Me.tslQuery.Text = CStr(QueryResult) + " result(s) in " + Me.ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                Me.tslQuery.Text = "No result found"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            SedangIsiDb = False
        End Try
    End Sub

    Private Sub SetAutoSize1(ByVal TS As TimeSpan)

        Try
            'cek masalah waktu, 
            If TS.TotalMilliseconds > Me.ObjColSize.MaxWaitTime Then
                If Me.TBLColSize.Rows.Count = 0 Then
                    'jika lebih dari waktu max dan tidak ada acuan, maka simpan data acuan, masih tetap autosize
                    Me.SetAutoSizeKolom()
                    Me.CopyDanSimpanDataKolom(Me.dgvHasil, Me.TBLColSize, ClsMasterColSize.EnumTipeTransaksi.MasterProduct)
                Else
                    'jika lebih dari waktu max dan ada waktu acuan 
                    If Me.TBLColSize.Rows.Count = Me.dgvHasil.Columns.Count Then
                        'jika jumlah kolomnya tepat, maka load berdasarkan acuan, tidak autosize
                        Me.SetSesuaiKolomAcuan()
                    Else
                        'jika jumlah kolomnya berbeda dengan acuan, maka simpan data acuan, masih tetap autozise
                        Me.SetAutoSizeKolom()
                        Me.CopyDanSimpanDataKolom(Me.dgvHasil, Me.TBLColSize, ClsMasterColSize.EnumTipeTransaksi.MasterProduct)
                    End If
                End If
            Else
                If Me.TBLColSize.Rows.Count = 0 Then
                    'jika tidak lebih dari acuan dan jumlah kolom acuan = 0, maka autosize
                    Me.SetAutoSizeKolom()
                Else
                    If Me.dgvHasil.RowCount <= 250 Then
                        Me.SetAutoSizeKolom()
                    Else
                        'jika lebih dari waktu max dan ada waktu acuan 
                        If Me.TBLColSize.Rows.Count = Me.dgvHasil.Columns.Count Then
                            'jika jumlah kolomnya tepat, maka load berdasarkan acuan, tidak autosize
                            Me.SetSesuaiKolomAcuan()
                        Else
                            'jika jumlah kolomnya berbeda dengan acuan, maka simpan data acuan, masih tetap autozise
                            Me.SetAutoSizeKolom()
                            Me.CopyDanSimpanDataKolom(Me.dgvHasil, Me.TBLColSize, ClsMasterColSize.EnumTipeTransaksi.MasterProduct)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CopyDanSimpanDataKolom(ByVal vDGV As DataGridView, ByVal vTBL As DseMasterColSize.tblmaster_columns_sizeDataTable, _
    ByVal vTipeTransaksi As Core_Module.ClsMasterColSize.EnumTipeTransaksi)
        Dim Col As DataGridViewColumn

        Try
            vTBL.Rows.Clear()
            For Each Col In vDGV.Columns
                Me.rowColSize = vTBL.Newtblmaster_columns_sizeRow
                Me.rowColSize.col_index = Col.Index
                Me.rowColSize.col_size = Col.Width
                Me.rowColSize.tipe_transaksi = vTipeTransaksi
                Me.rowColSize.user = Me.UserCode
                Me.rowColSize.primary = vTBL.Rows.Count + 1
                vTBL.Rows.Add(Me.rowColSize)
            Next
            Me.ObjColSize.SimpanTBL(vTBL)
            Me.ColumnChanged = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetAutoSizeKolom()
        Try
            If Me.dgvHasil.AutoSizeColumnsMode <> DataGridViewAutoSizeColumnsMode.AllCells Then
                Me.dgvHasil.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.dgvHasil.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetSesuaiKolomAcuan()
        Dim ijk As Integer = 0

        Try
            Me.dgvHasil.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvHasil.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing
            For Each Me.rowColSize In Me.TBLColSize.Rows
                Me.dgvHasil.Columns(ijk).Width = Me.rowColSize.col_size
                ijk += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmCariMasterProduct_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.ColumnChanged = True Then
                If Me.TBLColSize.Rows.Count > 0 Then
                    Me.CopyDanSimpanDataKolom(Me.dgvHasil, Me.TBLColSize, ClsMasterColSize.EnumTipeTransaksi.MasterProduct)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class