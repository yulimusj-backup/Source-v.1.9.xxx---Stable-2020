Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmMast_Discount
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris
    Public Ok As Boolean
    Public Baru As Boolean

    Dim iz As FrmMast_Discount_Main
    Dim Buffer As String = ""
    Dim DefaultClock As DateTime = "1950/01/01" + " " + Now.Hour.ToString + ":" + Now.Minute.ToString + ":" + Now.Second.ToString
    Dim DefPrec As Double = 0
    Dim EraseRowIndex_Prod As Integer = 0
    Dim EraseRowIndex_Branch As Integer = 0

    Public Obj As New ClsMast_Diskon
    Dim ObjItem As New ClsMast_Diskon_Item
    Dim ObjCekData As New ClsCekData
    Dim objProduct As New clsMasterProduct
    Dim ObjBranch As New ClsMast_Cabang

    Dim DTAGetProd As New DseMast_DiskonTableAdapters.DataTable4TableAdapter

    Private Sub FrmMast_Discount_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'Me.cboMember.SelectedIndex = 0
            'Me.cboBranch.SelectedIndex = 0
            'Me.cboProduct.SelectedIndex = 0
            'Me.cboTime_Advance.SelectedIndex = 0
            'Me.cboTimeFilter.SelectedIndex = 0
            'Me.cboStatus.SelectedIndex = 0
            'Me.cboRounding.SelectedIndex = 0
            'Me.cboDisc_Type.SelectedIndex = 0
            If Me.Baru = True Then
                Me.BersihkanObjekTime()
                Me.GBOTimeFilterEnabled(False)
                Me.GBOGimmickTabStop(False)
            Else
                If Me.cboDisc_Type.SelectedIndex = 0 Then
                    'Nominal
                    Me.TxtDiscountRate.Enabled = True
                    Me.ChkAllowToSell.Enabled = True
                    Me.TxtMaxQty1.Enabled = True
                    Me.TxtMinQty1.Enabled = True
                    Me.cboRounding.Enabled = True
                    'Gimmick
                    Me.TxtBaseQty.Enabled = False
                    Me.TxtFreeQty.Enabled = False
                    Me.TxtMaxQty2.Enabled = False
                    Me.TxtMinQty2.Enabled = False
                    Me.ChkMultiply.Enabled = False
                Else
                    'Nominal
                    Me.TxtDiscountRate.Enabled = False
                    Me.ChkAllowToSell.Enabled = False
                    Me.TxtMaxQty1.Enabled = False
                    Me.TxtMinQty1.Enabled = False
                    Me.cboRounding.Enabled = False
                    'Gimmick
                    Me.TxtBaseQty.Enabled = True
                    Me.TxtFreeQty.Enabled = True
                    Me.TxtMaxQty2.Enabled = True
                    Me.TxtMinQty2.Enabled = True
                    Me.ChkMultiply.Enabled = True
                End If
                If Me.cboTimeFilter.SelectedIndex = 0 Then
                    Me.DtpDateFrom.Enabled = False
                    Me.DtpDateTo.Enabled = False
                    Me.DtpTimeFrom.Enabled = False
                    Me.DtpTimeTo.Enabled = False
                    Me.cboTime_Advance.Enabled = False
                    Me.cboTimeFilter.SelectedIndex = 0
                End If
            End If
            Me.dgvItemProd.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvBranch.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BersihkanObjekGimmick()
        Try
            Me.TxtBaseQty.Text = "0"
            Me.TxtFreeQty.Text = "0"
            Me.TxtMaxQty2.Text = "0"
            Me.TxtMinQty2.Text = "0"
            Me.ChkMultiply.Checked = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BersihkanObjekNominal()
        Try
            Me.TxtDiscountRate.Text = "0"
            Me.ChkAllowToSell.Checked = False
            Me.TxtMaxQty1.Text = "0"
            Me.TxtMinQty1.Text = "0"
            Me.cboRounding.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub GBOGimmickTabStop(ByVal Tab As Boolean)
        Try
            If Tab = False Then
                Me.TxtBaseQty.Enabled = False
                Me.TxtFreeQty.Enabled = False
                Me.TxtMaxQty2.Enabled = False
                Me.TxtMinQty2.Enabled = False
                Me.ChkMultiply.Enabled = False
            End If
            If Tab = True Then
                Me.TxtBaseQty.Enabled = True
                Me.TxtFreeQty.Enabled = True
                Me.TxtMaxQty2.Enabled = True
                Me.TxtMinQty2.Enabled = True
                Me.ChkMultiply.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub GBONominalTabStop(ByVal Tab As Boolean)
        Try
            If Tab = False Then
                Me.TxtDiscountRate.Enabled = False
                Me.ChkAllowToSell.Enabled = False
                Me.TxtMaxQty1.Enabled = False
                Me.TxtMinQty1.Enabled = False
                Me.cboRounding.Enabled = False
            End If
            If Tab = True Then
                Me.TxtDiscountRate.Enabled = True
                Me.ChkAllowToSell.Enabled = True
                Me.TxtMaxQty1.Enabled = True
                Me.TxtMinQty1.Enabled = True
                Me.cboRounding.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboDisc_Type_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDisc_Type.SelectedIndexChanged
        Try
            If Me.cboDisc_Type.SelectedIndex = 0 Then
                Me.gboNominal.BringToFront()
                Me.BersihkanObjekNominal()
                Me.GBONominalTabStop(True)
                Me.GBOGimmickTabStop(False)
            Else
                Me.gboGimmick.BringToFront()
                Me.BersihkanObjekGimmick()
                Me.GBOGimmickTabStop(True)
                Me.GBONominalTabStop(False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboProduct_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboProduct.SelectedIndexChanged
        Try
            If Me.cboProduct.SelectedIndex = 0 Then
                Me.dgvItemProd.Rows.Clear()
                Me.cboQuickAdd.Visible = False
                Me.cboQuickAdd.TabStop = False
                Me.TxtQuickAdd.Visible = False
                Me.TxtQuickAdd.TabStop = False
                Me.btnFindQuickAdd.Visible = False
                Me.btnFindQuickAdd.TabStop = False
                Me.BtnInputItem.Visible = False
                Me.BtnInputItem.TabStop = False
                Me.TxtPrimaryQuickAdd.Text = "-1"
                Me.TxtQuickAdd.Text = ""
                Me.dgvItemProd.Rows.Clear()
                Me.dgvItemProd.Enabled = False
            Else
                Me.cboQuickAdd.Visible = True
                Me.cboQuickAdd.SelectedIndex = 0
                Me.cboQuickAdd.TabStop = True
                Me.TxtQuickAdd.Visible = True
                Me.TxtQuickAdd.TabStop = True
                Me.btnFindQuickAdd.Visible = True
                Me.btnFindQuickAdd.TabStop = True
                Me.BtnInputItem.Visible = True
                Me.BtnInputItem.TabStop = True
                Me.TxtPrimaryQuickAdd.Text = "-1"
                Me.TxtQuickAdd.Text = ""
                Me.dgvItemProd.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboTime_Advance_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTime_Advance.SelectedIndexChanged
        Try
            If Me.cboTime_Advance.SelectedIndex = 0 Then
                Me.clbDate.Enabled = False
                Me.clbDayOfWeek.Enabled = False
                Me.clbDayOfWeek.BringToFront()
            ElseIf Me.cboTime_Advance.SelectedIndex = 1 Then
                Me.clbDate.Enabled = False
                Me.clbDayOfWeek.Enabled = True
                Me.clbDayOfWeek.BringToFront()
            ElseIf Me.cboTime_Advance.SelectedIndex = 2 Then
                Me.clbDate.Enabled = True
                Me.clbDayOfWeek.Enabled = False
                Me.clbDate.BringToFront()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindQuickAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindQuickAdd.Click
        Dim i As New FrmCariMasterProductGroup
        Dim j As New frmCariMasterCategory
        Dim ObjPG As New ClsMasterProductGroup
        Dim ObjCT As New ClsMasterCategory
        i.BackColorAll = Me.BackColorAll
        j.BackColorAll = Me.BackColorAll
        Try
            If Me.cboQuickAdd.SelectedIndex = 3 Then
                If Me.TxtQuickAdd.Text <> "" Then
                    j.txtCari.Text = Me.TxtQuickAdd.Text
                End If

                j.ShowDialog()
                If j.PrimaryCategory > 0 Then
                    Me.TxtPrimaryQuickAdd.Text = j.PrimaryCategory
                    Me.TxtQuickAdd.Text = ObjCT.GetCategoryCode(Me.TxtPrimaryQuickAdd.Text)
                    Me.BtnInputItem.Focus()
                Else
                    Me.TxtQuickAdd.Clear()
                    Me.TxtPrimaryQuickAdd.Text = "-1"
                    Me.TxtQuickAdd.Focus()
                End If
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.cboQuickAdd.SelectedIndex = 0 Then
                    i.ProductGroupLevel = 3
                ElseIf Me.cboQuickAdd.SelectedIndex = 1 Then
                    i.ProductGroupLevel = 2
                ElseIf Me.cboQuickAdd.SelectedIndex = 2 Then
                    i.ProductGroupLevel = 1
                End If

                If Me.TxtQuickAdd.Text <> "" Then
                    i.txtCari.Text = Me.TxtQuickAdd.Text
                End If

                i.ShowDialog()
                If i.PrimaryProductGroup > 0 Then
                    Me.TxtPrimaryQuickAdd.Text = i.PrimaryProductGroup
                    Me.TxtQuickAdd.Text = ObjPG.GetGroupCode(Me.TxtPrimaryQuickAdd.Text)
                    Me.BtnInputItem.Focus()
                Else
                    Me.TxtQuickAdd.Clear()
                    Me.TxtPrimaryQuickAdd.Text = "-1"
                    Me.TxtQuickAdd.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            j.Dispose()
            ObjPG = Nothing
            ObjCT = Nothing
        End Try
    End Sub

    Private Sub TxtQuickAdd_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtQuickAdd.KeyDown
        Dim i As New FrmCariMasterProductGroup
        Dim j As New frmCariMasterCategory
        Dim ObjPG As New ClsMasterProductGroup
        Dim ObjCT As New ClsMasterCategory
        i.BackColorAll = Me.BackColorAll
        j.BackColorAll = Me.BackColorAll
        Try
            If e.KeyValue = 13 Then
                If Me.TxtQuickAdd.Text <> "" Then
                    If Me.cboQuickAdd.SelectedIndex = 3 Then
                        If ObjCT.ApakahSudahAda(Me.TxtPrimaryQuickAdd.Text) = True Then
                            Me.TxtPrimaryQuickAdd.Text = ObjCT.GetPrimary(Me.TxtQuickAdd.Text)
                            Me.TxtQuickAdd.Text = ObjCT.GetCategoryCode(Me.TxtPrimaryQuickAdd.Text)
                            Me.BtnInputItem.Focus()
                        Else
                            Me.btnFindQuickAdd.PerformClick()
                        End If
                    Else
                        i.cboCriteria.SelectedIndex = 0
                        If Me.cboQuickAdd.SelectedIndex = 0 Then
                            i.ProductGroupLevel = 3
                        ElseIf Me.cboQuickAdd.SelectedIndex = 1 Then
                            i.ProductGroupLevel = 2
                        ElseIf Me.cboQuickAdd.SelectedIndex = 2 Then
                            i.ProductGroupLevel = 1
                        End If
                        If ObjPG.ApakahSudahAda(Me.TxtPrimaryQuickAdd.Text, i.ProductGroupLevel) = True Then
                            Me.TxtPrimaryQuickAdd.Text = ObjPG.GetPrimary(Me.TxtQuickAdd.Text, i.ProductGroupLevel)
                            Me.TxtQuickAdd.Text = ObjPG.GetGroupCode(Me.TxtPrimaryQuickAdd.Text)
                            Me.BtnInputItem.Focus()
                        Else
                            Me.btnFindQuickAdd.PerformClick()
                        End If
                    End If
                Else
                    Me.btnFindQuickAdd.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            j.Dispose()
            ObjPG = Nothing
            ObjCT = Nothing
        End Try
    End Sub

    Private Sub TxtQuickAdd_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtQuickAdd.Leave
        Dim i As New FrmCariMasterProductGroup
        Dim j As New frmCariMasterCategory
        Dim ObjPG As New ClsMasterProductGroup
        Dim ObjCT As New ClsMasterCategory
        i.BackColorAll = Me.BackColorAll
        j.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtQuickAdd.Text <> "" Then
                If Me.cboQuickAdd.SelectedIndex = 3 Then
                    If ObjCT.ApakahSudahAda(Me.TxtQuickAdd.Text) = True Then
                        Me.TxtPrimaryQuickAdd.Text = ObjCT.GetPrimary(Me.TxtQuickAdd.Text)
                        Me.TxtQuickAdd.Text = ObjCT.GetCategoryCode(Me.TxtPrimaryQuickAdd.Text)
                    Else
                        If Me.TxtQuickAdd.Text <> "" Then
                            j.txtCari.Text = Me.TxtQuickAdd.Text
                        End If

                        j.ShowDialog()
                        If j.PrimaryCategory > 0 Then
                            Me.TxtPrimaryQuickAdd.Text = j.PrimaryCategory
                            Me.TxtQuickAdd.Text = ObjCT.GetCategoryCode(Me.TxtPrimaryQuickAdd.Text)
                            Me.BtnInputItem.Focus()
                        Else
                            Me.TxtQuickAdd.Clear()
                            Me.TxtPrimaryQuickAdd.Text = "-1"
                            Me.TxtQuickAdd.Focus()
                        End If
                    End If
                Else
                    i.cboCriteria.SelectedIndex = 0
                    If Me.cboQuickAdd.SelectedIndex = 0 Then
                        i.ProductGroupLevel = 3
                    ElseIf Me.cboQuickAdd.SelectedIndex = 1 Then
                        i.ProductGroupLevel = 2
                    ElseIf Me.cboQuickAdd.SelectedIndex = 2 Then
                        i.ProductGroupLevel = 1
                    End If
                    If ObjPG.ApakahSudahAda(Me.TxtQuickAdd.Text, i.ProductGroupLevel) = True Then
                        Me.TxtPrimaryQuickAdd.Text = ObjPG.GetPrimary(Me.TxtQuickAdd.Text, i.ProductGroupLevel)
                        Me.TxtQuickAdd.Text = ObjPG.GetGroupCode(Me.TxtPrimaryQuickAdd.Text)
                    Else
                        If Me.TxtQuickAdd.Text <> "" Then
                            i.txtCari.Text = Me.TxtQuickAdd.Text
                        End If

                        i.ShowDialog()
                        If i.PrimaryProductGroup > 0 Then
                            Me.TxtPrimaryQuickAdd.Text = i.PrimaryProductGroup
                            Me.TxtQuickAdd.Text = ObjPG.GetGroupCode(Me.TxtPrimaryQuickAdd.Text)
                            Me.BtnInputItem.Focus()
                        Else
                            Me.TxtQuickAdd.Clear()
                            Me.TxtPrimaryQuickAdd.Text = "-1"
                            Me.TxtQuickAdd.Focus()
                        End If
                    End If
                End If
            Else
                Me.TxtQuickAdd.Clear()
                Me.TxtPrimaryQuickAdd.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            j.Dispose()
            ObjPG = Nothing
            ObjCT = Nothing
        End Try
    End Sub

    Private Sub cboQuickAdd_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboQuickAdd.SelectedIndexChanged
        Try
            Me.TxtPrimaryQuickAdd.Text = "-1"
            Me.TxtQuickAdd.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProd_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemProd.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCodeNo.Index Or _
            e.ColumnIndex = Me.ColRemarks.Index Then
                Me.Buffer = Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemInput(ByVal vPrimaryProductDG As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvItemProd.Rows
                If rowDGV.Cells(Me.ColPrimaryProduct.Index).Value = vPrimaryProductDG Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Function GetIndexItemInputBranch(ByVal vPrimaryBranch As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvBranch.Rows
                If rowDGV.Cells(Me.ColPrimaryProduct.Index).Value = vPrimaryBranch Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub dgvItemProd_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemProd.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvItemProd.Rows(e.RowIndex)
        Dim Index As Integer = -1
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    'Cek string keterangan
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If ObjCekData.CekDataKosong(Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                            MsgBox("Data keterangan" _
                            + " tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Information")
                            Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    ''Cell kode 
                    If e.ColumnIndex = Me.ColCodeNo.Index Then
                        Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.objProduct.ApakahSudahAdaBarcode(Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Index = Me.GetIndexItemInput(Me.objProduct.GetPrimary(Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value))
                            If Index < 0 Then
                                Me.dgvItemProd.Item(Me.ColPrimaryProduct.Index, e.RowIndex).Value = Me.objProduct.GetPrimary(Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvItemProd.Item(Me.ColDesc.Index, e.RowIndex).Value = Me.objProduct.GetDescription(Me.dgvItemProd.Item(Me.ColPrimaryProduct.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> Index Then
                                    If Bahasa = EnumBahasa.Inggris Then
                                        MsgBox("You cannot input the same product code", MsgBoxStyle.Information, "Information")
                                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                                        MsgBox("Anda tidak dapat menginput product code yang sama", MsgBoxStyle.Information, "Informasi")
                                    End If
                                    Me.dgvItemProd.Item(Me.ColCodeNo.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex_Prod = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCekData.SetNumber(Me.dgvItemProd)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New frmCariMasterProduct
                            ij.UntukPOS = True
                            ij.txtCari.Text = Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            If ij.PrimaryProduct <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Index = Me.GetIndexItemInput(ij.PrimaryProduct)
                                If Index < 0 Then
                                    Me.dgvItemProd.Item(Me.ColCodeNo.Index, e.RowIndex).Value = Me.objProduct.GetProductCode(ij.PrimaryProduct)
                                    Me.dgvItemProd.Item(Me.ColDesc.Index, e.RowIndex).Value = Me.objProduct.GetDescription(ij.PrimaryProduct)
                                    Me.dgvItemProd.Item(Me.ColPrimaryProduct.Index, e.RowIndex).Value = ij.PrimaryProduct
                                Else
                                    If e.RowIndex <> Index Then
                                        If Bahasa = EnumBahasa.Inggris Then
                                            MsgBox("You cannot input the same product code", MsgBoxStyle.Information, "Information")
                                        ElseIf Bahasa = EnumBahasa.Indonesia Then
                                            MsgBox("Anda tidak dapat menginput product code yang sama", MsgBoxStyle.Information, "Informasi")
                                        End If
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex_Prod = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCekData.SetNumber(Me.dgvItemProd)
                                        Exit Sub
                                    Else
                                        Me.dgvItemProd.Item(Me.ColCodeNo.Index, e.RowIndex).Value = Me.objProduct.GetProductCode(ij.PrimaryProduct)
                                        Me.dgvItemProd.Item(Me.ColDesc.Index, e.RowIndex).Value = Me.objProduct.GetDescription(ij.PrimaryProduct)
                                        Me.dgvItemProd.Item(Me.ColPrimaryProduct.Index, e.RowIndex).Value = ij.PrimaryProduct
                                    End If
                                End If
                            ElseIf ij.PrimaryProduct = -1 Then
                                If Me.dgvItemProd.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex_Prod = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCekData.SetNumber(Me.dgvItemProd)
                    End If
                    'Hapus Baris Jika Code No DownGrade Kosong
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If String.IsNullOrEmpty(Me.dgvItemProd.Item(Me.ColCodeNo.Index, e.RowIndex).Value) = True Then
                            If Me.dgvItemProd.Rows.Count - 1 <> e.RowIndex Then
                                Me.EraseRowIndex_Prod = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        If e.ColumnIndex = Me.ColCodeNo.Index Then
                            If String.IsNullOrEmpty(Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                If Me.dgvItemProd.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex_Prod = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                End If
                            End If
                        Else
                            Me.dgvItemProd.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.dgvItemProd.Rows.RemoveAt(Me.EraseRowIndex_Prod)
            Me.ObjCekData.SetNumber(Me.dgvItemProd)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.dgvBranch.Rows.RemoveAt(Me.EraseRowIndex_Branch)
            Me.ObjCekData.SetNumber(Me.dgvBranch)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvBranch_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvBranch.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColBranch_Code.Index Or _
            e.ColumnIndex = Me.ColBranc_Remarks.Index Then
                Me.Buffer = Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvBranch_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvBranch.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvBranch.Rows(e.RowIndex)
        Dim Index As Integer = -1
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    'Cek string keterangan
                    If e.ColumnIndex = Me.ColBranc_Remarks.Index Then
                        If ObjCekData.CekDataKosong(Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                            MsgBox("Data keterangan" _
                            + " tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Information")
                            Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    ''Cell kode 
                    If e.ColumnIndex = Me.ColBranch_Code.Index Then
                        Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjBranch.ApakahSudahAda(Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Index = Me.GetIndexItemInputBranch(Me.ObjBranch.GetPrimary(Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value))
                            If Index < 0 Then
                                Me.dgvBranch.Item(Me.ColBranch_Primary.Index, e.RowIndex).Value = Me.ObjBranch.GetPrimary(Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvBranch.Item(Me.ColBranch_Desc.Index, e.RowIndex).Value = Me.ObjBranch.GetDesc(Me.dgvBranch.Item(Me.ColBranch_Primary.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> Index Then
                                    If Bahasa = EnumBahasa.Inggris Then
                                        MsgBox("You cannot input the same branch code", MsgBoxStyle.Information, "Information")
                                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                                        MsgBox("Anda tidak dapat menginput branch code yang sama", MsgBoxStyle.Information, "Informasi")
                                    End If
                                    Me.dgvBranch.Item(Me.ColBranch_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex_Branch = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCekData.SetNumber(Me.dgvBranch)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New FrmMast_Cabang_Cari
                            ij.txtCari.Text = Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            If ij.mPrimary <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Index = Me.GetIndexItemInputBranch(ij.mPrimary)
                                If Index < 0 Then
                                    Me.dgvBranch.Item(Me.ColBranch_Code.Index, e.RowIndex).Value = Me.ObjBranch.GetCode(ij.mPrimary)
                                    Me.dgvBranch.Item(Me.ColBranch_Desc.Index, e.RowIndex).Value = Me.ObjBranch.GetDesc(ij.mPrimary)
                                    Me.dgvBranch.Item(Me.ColBranch_Primary.Index, e.RowIndex).Value = ij.mPrimary
                                Else
                                    If e.RowIndex <> Index Then
                                        If Bahasa = EnumBahasa.Inggris Then
                                            MsgBox("You cannot input the same branch code", MsgBoxStyle.Information, "Information")
                                        ElseIf Bahasa = EnumBahasa.Indonesia Then
                                            MsgBox("Anda tidak dapat menginput branch code yang sama", MsgBoxStyle.Information, "Informasi")
                                        End If
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex_Branch = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCekData.SetNumber(Me.dgvBranch)
                                        Exit Sub
                                    Else
                                        Me.dgvBranch.Item(Me.ColBranch_Code.Index, e.RowIndex).Value = Me.ObjBranch.GetCode(ij.mPrimary)
                                        Me.dgvBranch.Item(Me.ColBranch_Desc.Index, e.RowIndex).Value = Me.ObjBranch.GetDesc(ij.mPrimary)
                                        Me.dgvBranch.Item(Me.ColBranch_Primary.Index, e.RowIndex).Value = ij.mPrimary
                                    End If
                                End If
                            ElseIf ij.mPrimary = -1 Then
                                If Me.dgvBranch.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex_Branch = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCekData.SetNumber(Me.dgvBranch)
                    End If
                    'Hapus Baris Jika Code No DownGrade Kosong
                    If e.ColumnIndex = Me.ColBranc_Remarks.Index Then
                        If String.IsNullOrEmpty(Me.dgvBranch.Item(Me.ColBranch_Code.Index, e.RowIndex).Value) = True Then
                            If Me.dgvBranch.Rows.Count - 1 <> e.RowIndex Then
                                Me.EraseRowIndex_Branch = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            End If
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColBranc_Remarks.Index Then
                        Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        If e.ColumnIndex = Me.ColBranch_Code.Index Then
                            If String.IsNullOrEmpty(Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                If Me.dgvBranch.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex_Branch = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                End If
                            End If
                        Else
                            Me.dgvBranch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnInputItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInputItem.Click
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If Me.cboQuickAdd.SelectedIndex <> 3 Then
                Obj.FungsiInputProductCode(Me.dgvItemProd, DTAGetProd, _
                Me.TxtPrimaryQuickAdd.Text, -1, Me.cboQuickAdd.SelectedIndex, ObjCekData, Me)
            Else
                Obj.FungsiInputProductCode(Me.dgvItemProd, DTAGetProd, _
                 -1, Me.TxtPrimaryQuickAdd.Text, Me.cboQuickAdd.SelectedIndex, ObjCekData, Me)
            End If

            'fokuskan ke baris terakhir
            While Me.dgvItemProd.Rows(Me.dgvItemProd.Rows.Count - 1).Displayed = False _
            Or Me.dgvItemProd.DisplayedRowCount(True) <> Me.dgvItemProd.DisplayedRowCount(False)
                Me.dgvItemProd.FirstDisplayedScrollingRowIndex = Me.dgvItemProd.FirstDisplayedScrollingRowIndex + 1
            End While

            Me.TxtPrimaryQuickAdd.Text = "-1"
            Me.TxtQuickAdd.Clear()
            Me.TxtQuickAdd.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboBranch_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBranch.SelectedIndexChanged
        Try
            If Me.cboBranch.SelectedIndex = 0 Then
                Me.dgvBranch.Rows.Clear()
                Me.dgvBranch.Enabled = False
            Else
                Me.dgvBranch.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BersihkanObjekTime()
        Try
            Me.DtpDateFrom.Value = Now
            Me.DtpDateTo.Value = Now
            Me.DtpTimeFrom.Value = Me.DefaultClock
            Me.DtpTimeTo.Value = Me.DefaultClock
            Me.cboTime_Advance.SelectedIndex = 0
            For x As Integer = 0 To Me.clbDayOfWeek.Items.Count - 1
                Me.clbDayOfWeek.SetItemCheckState(x, CheckState.Unchecked)
            Next
            For y As Integer = 0 To Me.clbDate.Items.Count - 1
                Me.clbDate.SetItemCheckState(y, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GBOTimeFilterEnabled(ByVal Tab As Boolean)
        Try
            If Tab = False Then
                Me.DtpDateFrom.Enabled = False
                Me.DtpDateTo.Enabled = False
                Me.DtpTimeFrom.Enabled = False
                Me.DtpTimeTo.Enabled = False
                Me.cboTime_Advance.Enabled = False
                Me.clbDayOfWeek.Enabled = False
                Me.clbDate.Enabled = False
            Else
                Me.DtpDateFrom.Enabled = True
                Me.DtpDateTo.Enabled = True
                Me.DtpTimeFrom.Enabled = True
                Me.DtpTimeTo.Enabled = True
                Me.cboTime_Advance.Enabled = True
                Me.clbDayOfWeek.Enabled = False
                Me.clbDate.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboTimeFilter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboTimeFilter.SelectedIndexChanged
        Try
            If Me.cboTimeFilter.SelectedIndex = 0 Then
                Me.BersihkanObjekTime()
                Me.GBOTimeFilterEnabled(False)
            Else
                Me.BersihkanObjekTime()
                Me.GBOTimeFilterEnabled(True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCekData, Me.TxtPrimary.Text, Me.Bahasa, Me.dgvItemProd, _
                        Me.dgvBranch, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            ''Me.NonAktifkanDetail()
                            Me.DialogResult = Windows.Forms.DialogResult.OK
                            Me.Ok = True
                            Me.Close()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCekData, Me.TxtPrimary.Text, Me.Bahasa, Me.dgvItemProd, _
                        Me.dgvBranch, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.DialogResult = Windows.Forms.DialogResult.OK
                            Me.Ok = True
                            Me.Close()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtCode.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnCancel.Focus()
            Me.Ok = True
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterAkunDetail_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.Ok = False Then
                If Me.TxtCode.Enabled = True Then
                    Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ApakahDataBenar() = True Then
                            Me.SalinKeObjek()
                            If Baru = True Then
                                'simpan baru
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.Obj.SimpanBaru(Me.UserCode, ObjCekData, Me.TxtPrimary.Text, Me.Bahasa, Me.dgvItemProd, _
                                 Me.dgvBranch, Me) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.DialogResult = Windows.Forms.DialogResult.OK

                                Else
                                    e.Cancel = True
                                End If
                            Else
                                'simpan perubahan
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.Obj.SimpanPerubahan(Me.UserCode, ObjCekData, Me.TxtPrimary.Text, Me.Bahasa, Me.dgvItemProd, _
                                Me.dgvBranch, Me) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.DialogResult = Windows.Forms.DialogResult.OK

                                Else
                                    e.Cancel = True
                                End If
                            End If
                        Else
                            e.Cancel = True
                        End If
                    ElseIf hasil = MsgBoxResult.Cancel Then
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, True)
            Me.TxtPrimary.Text = "-1"
            Me.cboDisc_Type.SelectedIndex = 0
            Me.cboMember.SelectedIndex = 0
            Me.cboStatus.SelectedIndex = 0
            Me.ChkMultidiskon.Checked = False
            Me.ChkShowDiscount.Checked = False
            Me.cboProduct.SelectedIndex = 0
            Me.cboBranch.SelectedIndex = 0
            Me.BersihkanObjekNominal()
            Me.BersihkanObjekGimmick()
            Me.BersihkanObjekTime()
            Me.cboQuickAdd.SelectedIndex = 0
            Me.TxtPrimaryQuickAdd.Text = "-1"
            Me.cboTimeFilter.SelectedIndex = 0
            Me.cboTime_Advance.SelectedIndex = 0
            Me.TxtUrutanDiskon.Text = "-1"

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvItemProd.Rows.Clear()
            Me.dgvBranch.Rows.Clear()

            Me.dgvItemProd.Enabled = False
            Me.dgvBranch.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Code")
            KoleksiTextBox.Add(Me.TxtCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtDesc)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'cek data yang salah, double dalam kode
            If Me.Baru = True Then
                If Obj.ApakahSudahAda(Me.TxtCode.Text) = True Then
                    MsgBox("Duplication in code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtCode.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.Obj.GetCode(Me.TxtPrimary.Text)
                If TipeCodeLama <> Me.TxtCode.Text Then
                    If Me.Obj.ApakahSudahAda(Me.TxtCode.Text) = True Then
                        MsgBox("Duplication in code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Obj.GetTimestamp(Me.TxtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Obj.CariDetail(Me.Obj.Primary, ObjCekData, Me.dgvItemProd, Me.dgvBranch, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            Me.dgvItemProd.AllowUserToAddRows = False
            If Me.cboProduct.SelectedIndex = 1 Then
                If Me.dgvItemProd.RowCount <= 0 Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please input item product", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan data product", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.dgvItemProd.AllowUserToAddRows = True
                    Me.tcoProperties.SelectedIndex = 0
                    Return False
                End If
            End If
            Me.dgvItemProd.AllowUserToAddRows = True

            Me.dgvBranch.AllowUserToAddRows = False
            If Me.cboBranch.SelectedIndex = 1 Then
                If Me.dgvBranch.RowCount <= 0 Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please input item branch", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan data branch", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.dgvBranch.AllowUserToAddRows = True
                    Me.tcoProperties.SelectedIndex = 1
                    Return False
                End If
            End If
            Me.dgvBranch.AllowUserToAddRows = True

            'Cek Number Positif 
            If Me.cboDisc_Type.SelectedIndex = 0 Then
                'If ObjCekData.CekBilanganPositif(Me.TxtDiscountRate.Text) = False Then
                'If Bahasa = EnumBahasa.Inggris Then
                'MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                'ElseIf Bahasa = EnumBahasa.Indonesia Then
                'MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                'End If
                'Me.TxtDiscountRate.Focus()
                'Return False
                'End If

                If ObjCekData.CekBilanganPositif(Me.TxtMaxQty1.Text) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtMaxQty1.Focus()
                    Return False
                End If

                If ObjCekData.CekBilanganPositif(Me.TxtMinQty1.Text) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtMinQty1.Focus()
                    Return False
                End If
            End If

            If Me.cboDisc_Type.SelectedIndex = 1 Then
                If ObjCekData.CekBilanganPositif(Me.TxtBaseQty.Text) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtBaseQty.Focus()
                    Return False
                End If
                If ObjCekData.CekBilanganPositif(Me.TxtFreeQty.Text) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtFreeQty.Focus()
                    Return False
                End If

                If ObjCekData.CekBilanganPositif(Me.TxtMaxQty2.Text) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtMaxQty2.Focus()
                    Return False
                End If

                If ObjCekData.CekBilanganPositif(Me.TxtMinQty2.Text) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtMinQty2.Focus()
                    Return False
                End If
            End If
            'batasan qty maximum dan minimum
            If Me.cboDisc_Type.SelectedIndex = 0 Then
                If CDec(Me.TxtMaxQty1.Text) < CDec(Me.TxtMinQty1.Text) Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("you can not input min qty grather than max qty", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("anda tidak dapat menginput min qty lebih besar dari max qty", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            Else
                If CDec(Me.TxtMaxQty2.Text) < CDec(Me.TxtMinQty2.Text) Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("you can not input min qty grather than max qty", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("anda tidak dapat menginput min qty lebih besar dari max qty", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            If String.IsNullOrEmpty(Me.TxtRemarks.Text) = False Then
                If Me.ObjCekData.CekDataKosong(Me.TxtRemarks.Text, False) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Data tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtRemarks.Focus()
                    Return False
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                .Primary = Me.TxtPrimary.Text
                .Code = Me.TxtCode.Text
                .Desc = Me.TxtDesc.Text
                .UrutanDisk = Me.TxtUrutanDiskon.Text
                If Me.cboDisc_Type.SelectedIndex = 0 Then
                    .Type = ClsMast_Diskon.EnumType.Nominal_Percentage
                Else
                    .Type = ClsMast_Diskon.EnumType.Gimmick
                End If
                If Me.cboMember.SelectedIndex = 0 Then
                    .Membership = ClsMast_Diskon.EnumMembership.Apply_For_All
                Else
                    .Membership = ClsMast_Diskon.EnumMembership.Apply_For_Member_Only
                End If
                If Me.cboStatus.SelectedIndex = 0 Then
                    .Status = ClsMast_Diskon.EnumStatus.Active
                Else
                    .Status = ClsMast_Diskon.EnumStatus.Not_Active
                End If
                If Me.ChkMultidiskon.Checked = False Then
                    .Multidiskon = 0
                Else
                    .Multidiskon = 1
                End If
                If Me.ChkShowDiscount.Checked = False Then
                    .ShowDiskonOnReceipt = 0
                Else
                    .ShowDiskonOnReceipt = 1
                End If
                .Remarks = Me.TxtRemarks.Text
                If Me.ChkAllowToSell.Checked = False Then
                    .AllowToSellMoreThanMaxQty = 0
                Else
                    .AllowToSellMoreThanMaxQty = 1
                End If
                .DiskonRate = Me.TxtDiscountRate.Text
                .BaseQty = Me.TxtBaseQty.Text
                .FreeQty = Me.TxtFreeQty.Text
                If Me.ChkMultiply.Checked = False Then
                    .Multiply = 0
                Else
                    .Multiply = 1
                End If
                If Me.cboRounding.SelectedIndex = 0 Then
                    .Rounding = ClsMast_Diskon.EnumRounding.Up
                ElseIf Me.cboRounding.SelectedIndex = 1 Then
                    .Rounding = ClsMast_Diskon.EnumRounding.Down
                ElseIf Me.cboRounding.SelectedIndex = 2 Then
                    .Rounding = ClsMast_Diskon.EnumRounding.Middle
                End If
                If Me.cboDisc_Type.SelectedIndex = 0 Then
                    .MinQty = Me.TxtMinQty1.Text
                    .MaxQty = Me.TxtMaxQty1.Text
                Else
                    .MinQty = Me.TxtMinQty2.Text
                    .MaxQty = Me.TxtMaxQty2.Text
                End If

                If Me.cboProduct.SelectedIndex = 0 Then
                    .ProductDataValidity = ClsMast_Diskon.EnumProductDataValidity.ApplyForAllItem
                Else
                    .ProductDataValidity = ClsMast_Diskon.EnumProductDataValidity.ApplyForSpesificItemOnly
                End If
                If Me.cboBranch.SelectedIndex = 0 Then
                    .BranchDataValidity = ClsMast_Diskon.EnumBranchDataValidity.ApplyForAllBranch
                Else
                    .BranchDataValidity = ClsMast_Diskon.EnumBranchDataValidity.ApplyForSpecificBranch
                End If
                If Me.cboTimeFilter.SelectedIndex = 0 Then
                    .Period = ClsMast_Diskon.EnumPeriod.AlwaysValid
                Else
                    .Period = ClsMast_Diskon.EnumPeriod.ValidForCertainPeriod
                End If
                .FromDate = Me.DtpDateFrom.Value
                .ToDate = Me.DtpDateTo.Value
                .FromTime = Me.DtpTimeFrom.Value
                .ToTime = Me.DtpTimeTo.Value
                If Me.cboTime_Advance.SelectedIndex = 0 Then
                    .AdvanceFiltering = ClsMast_Diskon.EnumAdvanceFilter.ApplyAlways
                ElseIf Me.cboTime_Advance.SelectedIndex = 1 Then
                    .AdvanceFiltering = ClsMast_Diskon.EnumAdvanceFilter.ApplyOnSpecificDayOfTheWeek
                ElseIf Me.cboTime_Advance.SelectedIndex = 2 Then
                    .AdvanceFiltering = ClsMast_Diskon.EnumAdvanceFilter.ApplyOnSpecificDateOfTheMonth
                End If
                For x As Integer = 0 To Me.clbDayOfWeek.Items.Count - 1
                    If Me.clbDayOfWeek.GetItemCheckState(x) = CheckState.Checked Then
                        If x = 0 Then
                            .IsMonday = 1
                        End If
                        If x = 1 Then
                            .IsTuesday = 1
                        End If
                        If x = 2 Then
                            .IsWednesday = 1
                        End If
                        If x = 3 Then
                            .IsThursday = 1
                        End If
                        If x = 4 Then
                            .IsFriday = 1
                        End If
                        If x = 5 Then
                            .IsSaturday = 1
                        End If
                        If x = 6 Then
                            .IsSunday = 1
                        End If
                    ElseIf Me.clbDayOfWeek.GetItemCheckState(x) = CheckState.Unchecked Then
                        If x = 0 Then
                            .IsMonday = 0
                        End If
                        If x = 1 Then
                            .IsTuesday = 0
                        End If
                        If x = 2 Then
                            .IsWednesday = 0
                        End If
                        If x = 3 Then
                            .IsThursday = 0
                        End If
                        If x = 4 Then
                            .IsFriday = 0
                        End If
                        If x = 5 Then
                            .IsSaturday = 0
                        End If
                        If x = 6 Then
                            .IsSunday = 0
                        End If
                    End If
                Next
                For y As Integer = 0 To Me.clbDate.Items.Count - 1
                    If Me.clbDate.GetItemCheckState(y) = CheckState.Checked Then
                        If y = 0 Then
                            .Is1 = 1
                        End If
                        If y = 1 Then
                            .Is2 = 1
                        End If
                        If y = 2 Then
                            .Is3 = 1
                        End If
                        If y = 3 Then
                            .Is4 = 1
                        End If
                        If y = 4 Then
                            .Is5 = 1
                        End If
                        If y = 5 Then
                            .Is6 = 1
                        End If
                        If y = 6 Then
                            .Is7 = 1
                        End If
                        If y = 7 Then
                            .Is8 = 1
                        End If
                        If y = 8 Then
                            .Is9 = 1
                        End If
                        If y = 9 Then
                            .Is10 = 1
                        End If
                        If y = 10 Then
                            .Is11 = 1
                        End If
                        If y = 11 Then
                            .Is12 = 1
                        End If
                        If y = 12 Then
                            .Is13 = 1
                        End If
                        If y = 13 Then
                            .Is14 = 1
                        End If
                        If y = 14 Then
                            .Is15 = 1
                        End If
                        If y = 15 Then
                            .Is16 = 1
                        End If
                        If y = 16 Then
                            .Is17 = 1
                        End If
                        If y = 17 Then
                            .Is18 = 1
                        End If
                        If y = 18 Then
                            .Is19 = 1
                        End If
                        If y = 19 Then
                            .Is20 = 1
                        End If
                        If y = 20 Then
                            .Is21 = 1
                        End If
                        If y = 21 Then
                            .Is22 = 1
                        End If
                        If y = 22 Then
                            .Is23 = 1
                        End If
                        If y = 23 Then
                            .Is24 = 1
                        End If
                        If y = 24 Then
                            .Is25 = 1
                        End If
                        If y = 25 Then
                            .Is26 = 1
                        End If
                        If y = 26 Then
                            .Is27 = 1
                        End If
                        If y = 27 Then
                            .Is28 = 1
                        End If
                        If y = 28 Then
                            .Is29 = 1
                        End If
                        If y = 29 Then
                            .Is30 = 1
                        End If
                        If y = 30 Then
                            .Is31 = 1
                        End If
                    ElseIf Me.clbDate.GetItemCheckState(y) = CheckState.Unchecked Then
                        If y = 0 Then
                            .Is1 = 0
                        End If
                        If y = 1 Then
                            .Is2 = 0
                        End If
                        If y = 2 Then
                            .Is3 = 0
                        End If
                        If y = 3 Then
                            .Is4 = 0
                        End If
                        If y = 4 Then
                            .Is5 = 0
                        End If
                        If y = 5 Then
                            .Is6 = 0
                        End If
                        If y = 6 Then
                            .Is7 = 0
                        End If
                        If y = 7 Then
                            .Is8 = 0
                        End If
                        If y = 8 Then
                            .Is9 = 0
                        End If
                        If y = 9 Then
                            .Is10 = 0
                        End If
                        If y = 10 Then
                            .Is11 = 0
                        End If
                        If y = 11 Then
                            .Is12 = 0
                        End If
                        If y = 12 Then
                            .Is13 = 0
                        End If
                        If y = 13 Then
                            .Is14 = 0
                        End If
                        If y = 14 Then
                            .Is15 = 0
                        End If
                        If y = 15 Then
                            .Is16 = 0
                        End If
                        If y = 16 Then
                            .Is17 = 0
                        End If
                        If y = 17 Then
                            .Is18 = 0
                        End If
                        If y = 18 Then
                            .Is19 = 0
                        End If
                        If y = 19 Then
                            .Is20 = 0
                        End If
                        If y = 20 Then
                            .Is21 = 0
                        End If
                        If y = 21 Then
                            .Is22 = 0
                        End If
                        If y = 22 Then
                            .Is23 = 0
                        End If
                        If y = 23 Then
                            .Is24 = 0
                        End If
                        If y = 24 Then
                            .Is25 = 0
                        End If
                        If y = 25 Then
                            .Is26 = 0
                        End If
                        If y = 26 Then
                            .Is27 = 0
                        End If
                        If y = 27 Then
                            .Is28 = 0
                        End If
                        If y = 28 Then
                            .Is29 = 0
                        End If
                        If y = 29 Then
                            .Is30 = 0
                        End If
                        If y = 30 Then
                            .Is31 = 0
                        End If
                    End If
                Next
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtCode.Text = .Code
                Me.TxtDesc.Text = .Desc
                Me.TxtUrutanDiskon.Text = .UrutanDisk
                If .Type = ClsMast_Diskon.EnumType.Nominal_Percentage Then
                    Me.cboDisc_Type.SelectedIndex = 0
                Else
                    Me.cboDisc_Type.SelectedIndex = 1
                End If
                If .Membership = ClsMast_Diskon.EnumMembership.Apply_For_All Then
                    Me.cboMember.SelectedIndex = 0
                Else
                    Me.cboMember.SelectedIndex = 1
                End If
                If .Status = ClsMast_Diskon.EnumStatus.Active Then
                    Me.cboStatus.SelectedIndex = 0
                Else
                    Me.cboStatus.SelectedIndex = 1
                End If
                If .Multidiskon = 0 Then
                    Me.ChkMultidiskon.Checked = False
                Else
                    Me.ChkMultidiskon.Checked = True
                End If
                If .ShowDiskonOnReceipt = 0 Then
                    Me.ChkShowDiscount.Checked = False
                Else
                    Me.ChkShowDiscount.Checked = True
                End If
                Me.TxtRemarks.Text = .Remarks
                If .AllowToSellMoreThanMaxQty = 0 Then
                    Me.ChkAllowToSell.Checked = False
                Else
                    Me.ChkAllowToSell.Checked = True
                End If
                Me.TxtDiscountRate.Text = .DiskonRate
                Me.TxtBaseQty.Text = .BaseQty
                Me.TxtFreeQty.Text = .FreeQty
                If .Multiply = 0 Then
                    Me.ChkMultiply.Checked = False
                Else
                    Me.ChkMultiply.Checked = True
                End If
                If .Rounding = ClsMast_Diskon.EnumRounding.Up Then
                    Me.cboRounding.SelectedIndex = 0
                ElseIf .Rounding = ClsMast_Diskon.EnumRounding.Down Then
                    Me.cboRounding.SelectedIndex = 1
                ElseIf .Rounding = ClsMast_Diskon.EnumRounding.Middle Then
                    Me.cboRounding.SelectedIndex = 2
                End If
                If Me.cboDisc_Type.SelectedIndex = 0 Then
                    Me.TxtMinQty1.Text = .MinQty
                    Me.TxtMaxQty1.Text = .MaxQty
                    Me.TxtMinQty2.Text = "0"
                    Me.TxtMaxQty2.Text = "0"
                Else
                    Me.TxtMinQty2.Text = .MinQty
                    Me.TxtMaxQty2.Text = .MaxQty
                    Me.TxtMinQty1.Text = "0"
                    Me.TxtMaxQty1.Text = "0"
                End If

                If .ProductDataValidity = ClsMast_Diskon.EnumProductDataValidity.ApplyForAllItem Then
                    Me.cboProduct.SelectedIndex = 0
                Else
                    Me.cboProduct.SelectedIndex = 1
                End If
                If .BranchDataValidity = ClsMast_Diskon.EnumBranchDataValidity.ApplyForAllBranch Then
                    Me.cboBranch.SelectedIndex = 0
                Else
                    Me.cboBranch.SelectedIndex = 1
                End If
                If .Period = ClsMast_Diskon.EnumPeriod.AlwaysValid Then
                    Me.cboTimeFilter.SelectedIndex = 0
                Else
                    Me.cboTimeFilter.SelectedIndex = 1
                End If
                Me.DtpDateFrom.Value = .FromDate
                Me.DtpDateTo.Value = .ToDate
                Me.DtpTimeFrom.Value = .FromTime
                Me.DtpTimeTo.Value = .ToTime
                If .AdvanceFiltering = ClsMast_Diskon.EnumAdvanceFilter.ApplyAlways Then
                    Me.cboTime_Advance.SelectedIndex = 0
                ElseIf .AdvanceFiltering = ClsMast_Diskon.EnumAdvanceFilter.ApplyOnSpecificDayOfTheWeek Then
                    Me.cboTime_Advance.SelectedIndex = 1
                ElseIf .AdvanceFiltering = ClsMast_Diskon.EnumAdvanceFilter.ApplyOnSpecificDateOfTheMonth Then
                    Me.cboTime_Advance.SelectedIndex = 2
                End If
                If .IsMonday = 1 Then
                    Me.clbDayOfWeek.SetItemCheckState(0, CheckState.Checked)
                Else
                    Me.clbDayOfWeek.SetItemCheckState(0, CheckState.Unchecked)
                End If
                If .IsTuesday = 1 Then
                    Me.clbDayOfWeek.SetItemCheckState(1, CheckState.Checked)
                Else
                    Me.clbDayOfWeek.SetItemCheckState(1, CheckState.Unchecked)
                End If
                If .IsWednesday = 1 Then
                    Me.clbDayOfWeek.SetItemCheckState(2, CheckState.Checked)
                Else
                    Me.clbDayOfWeek.SetItemCheckState(2, CheckState.Unchecked)
                End If
                If .IsThursday = 1 Then
                    Me.clbDayOfWeek.SetItemCheckState(3, CheckState.Checked)
                Else
                    Me.clbDayOfWeek.SetItemCheckState(3, CheckState.Unchecked)
                End If
                If .IsFriday = 1 Then
                    Me.clbDayOfWeek.SetItemCheckState(4, CheckState.Checked)
                Else
                    Me.clbDayOfWeek.SetItemCheckState(4, CheckState.Unchecked)
                End If
                If .IsSaturday = 1 Then
                    Me.clbDayOfWeek.SetItemCheckState(5, CheckState.Checked)
                Else
                    Me.clbDayOfWeek.SetItemCheckState(5, CheckState.Unchecked)
                End If
                If .IsSunday = 1 Then
                    Me.clbDayOfWeek.SetItemCheckState(6, CheckState.Checked)
                Else
                    Me.clbDayOfWeek.SetItemCheckState(6, CheckState.Unchecked)
                End If
                If .Is1 = 1 Then
                    Me.clbDate.SetItemCheckState(0, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(0, CheckState.Unchecked)
                End If
                If .Is2 = 1 Then
                    Me.clbDate.SetItemCheckState(1, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(1, CheckState.Unchecked)
                End If
                If .Is3 = 1 Then
                    Me.clbDate.SetItemCheckState(2, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(2, CheckState.Unchecked)
                End If
                If .Is4 = 1 Then
                    Me.clbDate.SetItemCheckState(3, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(3, CheckState.Unchecked)
                End If
                If .Is5 = 1 Then
                    Me.clbDate.SetItemCheckState(4, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(4, CheckState.Unchecked)
                End If
                If .Is6 = 1 Then
                    Me.clbDate.SetItemCheckState(5, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(5, CheckState.Unchecked)
                End If
                If .Is7 = 1 Then
                    Me.clbDate.SetItemCheckState(6, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(6, CheckState.Unchecked)
                End If
                If .Is8 = 1 Then
                    Me.clbDate.SetItemCheckState(7, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(7, CheckState.Unchecked)
                End If
                If .Is9 = 1 Then
                    Me.clbDate.SetItemCheckState(8, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(8, CheckState.Unchecked)
                End If
                If .Is10 = 1 Then
                    Me.clbDate.SetItemCheckState(9, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(9, CheckState.Unchecked)
                End If
                If .Is11 = 1 Then
                    Me.clbDate.SetItemCheckState(10, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(10, CheckState.Unchecked)
                End If
                If .Is12 = 1 Then
                    Me.clbDate.SetItemCheckState(11, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(11, CheckState.Unchecked)
                End If
                If .Is13 = 1 Then
                    Me.clbDate.SetItemCheckState(12, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(12, CheckState.Unchecked)
                End If
                If .Is14 = 1 Then
                    Me.clbDate.SetItemCheckState(13, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(13, CheckState.Unchecked)
                End If
                If .Is15 = 1 Then
                    Me.clbDate.SetItemCheckState(14, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(14, CheckState.Unchecked)
                End If
                If .Is16 = 1 Then
                    Me.clbDate.SetItemCheckState(15, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(15, CheckState.Unchecked)
                End If
                If .Is17 = 1 Then
                    Me.clbDate.SetItemCheckState(16, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(16, CheckState.Unchecked)
                End If
                If .Is18 = 1 Then
                    Me.clbDate.SetItemCheckState(17, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(17, CheckState.Unchecked)
                End If
                If .Is19 = 1 Then
                    Me.clbDate.SetItemCheckState(18, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(18, CheckState.Unchecked)
                End If
                If .Is20 = 1 Then
                    Me.clbDate.SetItemCheckState(19, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(19, CheckState.Unchecked)
                End If
                If .Is21 = 1 Then
                    Me.clbDate.SetItemCheckState(20, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(20, CheckState.Unchecked)
                End If
                If .Is22 = 1 Then
                    Me.clbDate.SetItemCheckState(21, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(21, CheckState.Unchecked)
                End If
                If .Is23 = 1 Then
                    Me.clbDate.SetItemCheckState(22, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(22, CheckState.Unchecked)
                End If
                If .Is24 = 1 Then
                    Me.clbDate.SetItemCheckState(23, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(23, CheckState.Unchecked)
                End If
                If .Is25 = 1 Then
                    Me.clbDate.SetItemCheckState(24, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(24, CheckState.Unchecked)
                End If
                If .Is26 = 1 Then
                    Me.clbDate.SetItemCheckState(25, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(25, CheckState.Unchecked)
                End If
                If .Is27 = 1 Then
                    Me.clbDate.SetItemCheckState(26, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(26, CheckState.Unchecked)
                End If
                If .Is28 = 1 Then
                    Me.clbDate.SetItemCheckState(27, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(27, CheckState.Unchecked)
                End If
                If .Is29 = 1 Then
                    Me.clbDate.SetItemCheckState(28, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(28, CheckState.Unchecked)
                End If
                If .Is30 = 1 Then
                    Me.clbDate.SetItemCheckState(29, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(29, CheckState.Unchecked)
                End If
                If .Is31 = 1 Then
                    Me.clbDate.SetItemCheckState(30, CheckState.Checked)
                Else
                    Me.clbDate.SetItemCheckState(30, CheckState.Unchecked)
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub AktifkanDetail()
        Try
            ''Me.ObjCekData.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimary.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gboNominal.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gboGimmick.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.GroupBox1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.tbp_Branch.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.tbp_Produk.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.tcoProperties.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, True, True)

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub NonAktifkanDetail()
        Try
            ''Me.ObjCekData.SetBackColor(Me.gboDetail.Controls, Me.BackColor, True)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.gboNominal.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.gboGimmick.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.GroupBox1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.TabPage1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.tbp_Branch.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.tbp_Produk.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.tcoProperties.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, True, False)

            Me.CancelButton = Me.btnCancel
            Me.BackColorDefault()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.ObjCekData.SetBackColor(Me.gboDetail.Controls, Color.White, False)
            Me.ObjCekData.SetBackColor(Me.GroupBox1.Controls, Color.White, False)
            Me.ObjCekData.SetBackColor(Me.gboNominal.Controls, Color.White, False)
            Me.ObjCekData.SetBackColor(Me.gboGimmick.Controls, Color.White, False)
            Me.ObjCekData.SetBackColor(Me.TabPage1.Controls, Color.White, False)
            Me.ObjCekData.SetBackColor(Me.tbp_Branch.Controls, Color.White, False)
            Me.ObjCekData.SetBackColor(Me.tbp_Produk.Controls, Color.White, False)
            Me.ObjCekData.SetBackColor(Me.tcoProperties.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.ObjCekData.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCekData.SetBackColor(Me.GroupBox1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCekData.SetBackColor(Me.gboNominal.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCekData.SetBackColor(Me.gboGimmick.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCekData.SetBackColor(Me.TabPage1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCekData.SetBackColor(Me.tbp_Branch.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCekData.SetBackColor(Me.tbp_Produk.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCekData.SetBackColor(Me.tcoProperties.Controls, Me.TxtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCode.Enabled = True Then
                    If Me.TxtCode.ReadOnly = False Then
                        TxtCode.BackColor = Me.BackColorAll
                    Else
                        TxtCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCode.Enabled = True Then
                    If Me.TxtCode.ReadOnly = False Then
                        TxtCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDesc.Enabled = True Then
                    If Me.TxtDesc.ReadOnly = False Then
                        TxtDesc.BackColor = Me.BackColorAll
                    Else
                        TxtDesc.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDesc.Enabled = True Then
                    If Me.TxtDesc.ReadOnly = False Then
                        TxtDesc.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtDesc.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.Enabled = True Then
                    If Me.TxtRemarks.ReadOnly = False Then
                        TxtRemarks.BackColor = Me.BackColorAll
                    Else
                        TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.Enabled = True Then
                    If Me.TxtRemarks.ReadOnly = False Then
                        TxtRemarks.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscountRate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscountRate.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDiscountRate.Enabled = True Then
                    If Me.TxtDiscountRate.ReadOnly = False Then
                        TxtDiscountRate.BackColor = Me.BackColorAll
                    Else
                        TxtDiscountRate.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscountRate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscountRate.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDiscountRate.Enabled = True Then
                    If Me.TxtDiscountRate.ReadOnly = False Then
                        TxtDiscountRate.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtDiscountRate.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMaxQty1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMaxQty1.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtMaxQty1.Enabled = True Then
                    If Me.TxtMaxQty1.ReadOnly = False Then
                        TxtMaxQty1.BackColor = Me.BackColorAll
                    Else
                        TxtMaxQty1.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMaxQty1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMaxQty1.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtMaxQty1.Enabled = True Then
                    If Me.TxtMaxQty1.ReadOnly = False Then
                        TxtMaxQty1.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtMaxQty1.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMinQty1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMinQty1.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtMinQty1.Enabled = True Then
                    If Me.TxtMinQty1.ReadOnly = False Then
                        TxtMinQty1.BackColor = Me.BackColorAll
                    Else
                        TxtMinQty1.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMinQty1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMinQty1.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtMinQty1.Enabled = True Then
                    If Me.TxtMinQty1.ReadOnly = False Then
                        TxtMinQty1.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtMinQty1.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMaxQty2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMaxQty2.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtMaxQty2.Enabled = True Then
                    If Me.TxtMaxQty2.ReadOnly = False Then
                        TxtMaxQty2.BackColor = Me.BackColorAll
                    Else
                        TxtMaxQty2.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMaxQty2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMaxQty2.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtMaxQty2.Enabled = True Then
                    If Me.TxtMaxQty2.ReadOnly = False Then
                        TxtMaxQty2.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtMaxQty2.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMinQty2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMinQty2.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtMinQty2.Enabled = True Then
                    If Me.TxtMinQty2.ReadOnly = False Then
                        TxtMinQty2.BackColor = Me.BackColorAll
                    Else
                        TxtMinQty2.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMinQty2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMinQty2.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtMinQty2.Enabled = True Then
                    If Me.TxtMinQty2.ReadOnly = False Then
                        TxtMinQty2.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtMinQty2.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBaseQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBaseQty.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBaseQty.Enabled = True Then
                    If Me.TxtBaseQty.ReadOnly = False Then
                        TxtBaseQty.BackColor = Me.BackColorAll
                    Else
                        TxtBaseQty.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBaseQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBaseQty.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBaseQty.Enabled = True Then
                    If Me.TxtBaseQty.ReadOnly = False Then
                        TxtBaseQty.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtBaseQty.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFreeQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFreeQty.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtFreeQty.Enabled = True Then
                    If Me.TxtFreeQty.ReadOnly = False Then
                        TxtFreeQty.BackColor = Me.BackColorAll
                    Else
                        TxtFreeQty.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFreeQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFreeQty.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtFreeQty.Enabled = True Then
                    If Me.TxtFreeQty.ReadOnly = False Then
                        TxtFreeQty.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtFreeQty.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQuickAdd_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQuickAdd.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtQuickAdd.Enabled = True Then
                    If Me.TxtQuickAdd.ReadOnly = False Then
                        TxtQuickAdd.BackColor = Me.BackColorAll
                    Else
                        TxtQuickAdd.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQuickAdd_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQuickAdd.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtQuickAdd.Enabled = True Then
                    If Me.TxtQuickAdd.ReadOnly = False Then
                        TxtQuickAdd.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtQuickAdd.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvBranch_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvBranch.UserDeletedRow
        Try
            ObjCekData.SetNumber(Me.dgvBranch, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProd_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemProd.UserDeletedRow
        Try
            ObjCekData.SetNumber(Me.dgvItemProd, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Number
    Private Sub TxtDiscountRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscountRate.Enter
        Try
            Me.Buffer = Me.TxtDiscountRate.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscountRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscountRate.KeyDown
        Try
            ObjCekData.Disc_KeyDown(Me.TxtDiscountRate, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscountRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscountRate.KeyUp
        Try
            ObjCekData.Disc_KeyUp(Me.TxtDiscountRate, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscountRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscountRate.Leave
        Try
            ObjCekData.Disc_Leave(Me.TxtDiscountRate, Me.TxtDiscountRate.Text, sender, e, Me.DefPrec, ObjCekData, Me.Buffer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMaxQty1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMaxQty1.Enter
        Try
            ObjCekData.Number_Enter(Me.TxtMaxQty1, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMaxQty1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMaxQty1.KeyDown
        Try
            ObjCekData.Number_KeyDown(Me.TxtMaxQty1, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMaxQty1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMaxQty1.KeyUp
        Try
            ObjCekData.Number_KeyUp(Me.TxtMaxQty1, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMaxQty1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMaxQty1.Leave
        Try
            ObjCekData.Number_Leave(Me.TxtMaxQty1, sender, e, Me.DefPrec, ObjCekData)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMinQty1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMinQty1.Enter
        Try
            ObjCekData.Number_Enter(Me.TxtMinQty1, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMinQty1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMinQty1.KeyDown
        Try
            ObjCekData.Number_KeyDown(Me.TxtMinQty1, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMinQty1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMinQty1.KeyUp
        Try
            ObjCekData.Number_KeyUp(Me.TxtMinQty1, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMinQty1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMinQty1.Leave
        Try
            ObjCekData.Number_Leave(Me.TxtMinQty1, sender, e, Me.DefPrec, ObjCekData)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFreeQty_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFreeQty.Enter
        Try
            ObjCekData.Number_Enter(Me.TxtFreeQty, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFreeQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFreeQty.KeyDown
        Try
            ObjCekData.Number_KeyDown(Me.TxtFreeQty, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFreeQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFreeQty.KeyUp
        Try
            ObjCekData.Number_KeyUp(Me.TxtFreeQty, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFreeQty_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFreeQty.Leave
        Try
            ObjCekData.Number_Leave(Me.TxtFreeQty, sender, e, Me.DefPrec, ObjCekData)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBaseQty_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBaseQty.Enter
        Try
            ObjCekData.Number_Enter(Me.TxtBaseQty, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBaseQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBaseQty.KeyDown
        Try
            ObjCekData.Number_KeyDown(Me.TxtBaseQty, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBaseQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBaseQty.KeyUp
        Try
            ObjCekData.Number_KeyUp(Me.TxtBaseQty, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBaseQty_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBaseQty.Leave
        Try
            ObjCekData.Number_Leave(Me.TxtBaseQty, sender, e, Me.DefPrec, ObjCekData)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscountRate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtDiscountRate.TextChanged

    End Sub
End Class