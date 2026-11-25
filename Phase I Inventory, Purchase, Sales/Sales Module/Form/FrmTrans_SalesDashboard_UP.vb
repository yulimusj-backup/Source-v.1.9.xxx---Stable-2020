Option Explicit On
Option Strict Off

Imports Core_Module

Partial Public Class FrmTrans_SalesDashboard
    Dim vTBLDODummy As New DseTransaksiDO.DataTable13DataTable
    Dim vrowDODummy As DseTransaksiDO.DataTable13Row
    Dim vDTADODummy As New DseTransaksiDOTableAdapters.DataTable13TableAdapter

    Private Sub TxtUP_UserCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUP_UserCode.KeyDown
        Dim i As New FrmCariMasterUser
        i.backColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjUser.ApakahSudahAda(Me.TxtUP_UserCode.Text) = True Then
                    Me.TxtUP_UserName.Text = Me.ObjUser.GetNama(Me.TxtUP_UserCode.Text)
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.TxtUP_UserCode.Text <> "" Then
                        i.TxtCari.Text = Me.TxtUP_UserCode.Text
                    End If

                    i.ShowDialog()
                    If i.Kodeuser <> "" Then
                        Me.TxtUP_UserCode.Text = i.Kodeuser
                        Me.TxtUP_UserName.Text = Me.ObjUser.GetNama(Me.TxtUP_UserCode.Text)
                    Else
                        Me.TxtUP_UserCode.Clear()
                        Me.TxtUP_UserName.Clear()
                        Me.TxtUP_UserCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtUP_UserCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUP_UserCode.Leave
        Dim i As New FrmCariMasterUser
        i.backColorAll = Me.BackColorAll

        Try
            If Me.TxtUP_UserCode.Text <> "" Then
                If Me.ObjUser.ApakahSudahAda(Me.TxtUP_UserCode.Text) Then
                    Me.TxtUP_UserName.Text = Me.ObjUser.GetNama(Me.TxtUP_UserCode.Text)
                Else
                    If Me.TxtUP_UserCode.Text <> "" Then
                        i.TxtCari.Text = Me.TxtUP_UserCode.Text
                    End If

                    i.ShowDialog()
                    If i.Kodeuser <> "" Then
                        Me.TxtUP_UserCode.Text = i.Kodeuser
                        Me.TxtUP_UserName.Text = Me.ObjUser.GetNama(Me.TxtUP_UserCode.Text)
                    Else
                        Me.TxtUP_UserCode.Clear()
                        Me.TxtUP_UserName.Clear()
                        Me.TxtUP_UserCode.Focus()
                    End If
                End If
            Else
                Me.TxtUP_UserCode.Clear()
                Me.TxtUP_UserName.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnUP_FindUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP_FindUser.Click
        Dim i As New FrmCariMasterUser
        i.backColorAll = Me.BackColorAll

        Try
            If Me.TxtUP_UserCode.Text <> "" Then
                i.TxtCari.Text = Me.TxtUP_UserCode.Text
            End If

            i.ShowDialog()
            If i.Kodeuser <> "" Then
                Me.TxtUP_UserCode.Text = i.Kodeuser
                Me.TxtUP_UserName.Text = Me.ObjUser.GetNama(Me.TxtUP_UserCode.Text)
            Else
                Me.TxtUP_UserCode.Clear()
                Me.TxtUP_UserName.Clear()
                Me.TxtUP_UserCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtUP_CustCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUP_CustCode.KeyDown
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtUP_CustCode.Text) = True Then
                    Me.TxtUP_PrimaryCust.Text = Me.ObjCustomer.GetPrimary(Me.TxtUP_CustCode.Text)
                    Me.TxtUP_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtUP_PrimaryCust.Text)
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.TxtUP_CustCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtUP_CustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtUP_PrimaryCust.Text = i.PrimaryCustomer
                        Me.TxtUP_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtUP_PrimaryCust.Text)
                        Me.TxtUP_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtUP_PrimaryCust.Text)
                    Else
                        Me.TxtUP_CustCode.Clear()
                        Me.TxtUP_TxtCustName.Clear()
                        Me.TxtUP_PrimaryCust.Text = "-1"
                        Me.TxtUP_CustCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtUP_CustCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUP_CustCode.Leave
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtUP_CustCode.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtUP_CustCode.Text) Then
                    Dim PrimaryBaru As Decimal = Me.ObjCustomer.GetPrimary(Me.TxtUP_CustCode.Text)
                    If PrimaryBaru <> Me.TxtUP_PrimaryCust.Text Then
                        Me.TxtUP_PrimaryCust.Text = PrimaryBaru
                        Me.TxtUP_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtUP_PrimaryCust.Text)
                    End If
                Else
                    If Me.TxtUP_CustCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtUP_CustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtUP_PrimaryCust.Text = i.PrimaryCustomer
                        Me.TxtUP_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtUP_PrimaryCust.Text)
                        Me.TxtUP_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtUP_PrimaryCust.Text)
                    Else
                        Me.TxtUP_CustCode.Clear()
                        Me.TxtUP_TxtCustName.Clear()
                        Me.TxtUP_PrimaryCust.Text = "-1"
                        Me.TxtUP_CustCode.Focus()
                    End If
                End If
            Else
                Me.TxtUP_CustCode.Clear()
                Me.TxtUP_TxtCustName.Clear()
                Me.TxtUP_PrimaryCust.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnUP_FindCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP_FindCust.Click
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtUP_CustCode.Text <> "" Then
                i.txtCari.Text = Me.TxtUP_CustCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.TxtUP_PrimaryCust.Text = i.PrimaryCustomer
                Me.TxtUP_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtUP_PrimaryCust.Text)
                Me.TxtUP_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtUP_PrimaryCust.Text)
            Else
                Me.TxtUP_CustCode.Clear()
                Me.TxtUP_TxtCustName.Clear()
                Me.TxtUP_PrimaryCust.Text = "-1"
                Me.TxtUP_CustCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub GotFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAR_CustCode.GotFocus, TxtAR_UserCode.GotFocus, _
    TxtSF_CustCode.GotFocus, TxtSF_ProductCode.GotFocus, TxtSF_UserCode.GotFocus, TxtUP_CustCode.GotFocus, TxtUP_UserCode.GotFocus, _
    TxtMA_CustCode.GotFocus, TxtMA_UserCode.GotFocus
        Dim txt As TextBox = TryCast(sender, TextBox)
        txt.BackColor = Me.BackColorAll
    End Sub

    Private Sub LostFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAR_CustCode.LostFocus, TxtAR_UserCode.LostFocus, _
    TxtSF_CustCode.LostFocus, TxtSF_ProductCode.LostFocus, TxtSF_UserCode.LostFocus, TxtUP_CustCode.LostFocus, TxtUP_UserCode.LostFocus, _
    TxtMA_CustCode.LostFocus, TxtMA_UserCode.LostFocus
        Dim txt As TextBox = TryCast(sender, TextBox)
        txt.BackColor = Color.White
    End Sub

    Private Sub llbUP_Hide_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbUP_Hide.Click
        Try
            If Me.llbUP_Hide.Text = "Hide Detail" Then
                'hide detail
                Me.llbUP_Hide.Text = "Show Detail"
                Me.SplitContainer2.Panel2Collapsed = True
            Else
                'hide detail
                Me.llbUP_Hide.Text = "Hide Detail"
                Me.SplitContainer2.Panel2Collapsed = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Filter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUP_Filter.Click
        Dim ListPrimary As New List(Of Double)
        Dim Lanjut As Boolean = False

        Try
            'isikan data do, yang belum dibuatkan item pengirimannya
            Me.vDTADODummy.Connection.ConnectionString = Core_Module.GetConstring
            Me.vDTADODummy.Fill(Me.vTBLDODummy)

            Me.dgvUP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvUP.Rows.Clear()

            'filter atas (do yang dibuat oleh user ybs atau so-nya dibuat oleh user yang bersangkutan) dan kode customer
            For Each Me.vrowDODummy In Me.vTBLDODummy.Rows
                If ListPrimary.Contains(Me.vrowDODummy.Primary_DO) = False Then
                    ListPrimary.Add(Me.vrowDODummy.Primary_DO)
                    Lanjut = False

                    If Me.TxtUP_CustCode.Text.Trim <> "" Then
                        If Me.TxtUP_UserCode.Text = "" Then
                            'filter berdasarkan cust code
                            If Me.TxtUP_CustCode.Text.Trim = Me.vrowDODummy.Customer_Code Then
                                Lanjut = True
                            End If
                        Else
                            'filter berdasarkan cust code dan user code
                            If Me.TxtUP_CustCode.Text.Trim = Me.vrowDODummy.Customer_Code Then
                                If Me.TxtUP_UserCode.Text.Trim = Me.vrowDODummy.Created_By Or _
                                Me.TxtUP_UserCode.Text.Trim = Me.vrowDODummy.SO_Created_By Then
                                    Lanjut = True
                                End If
                            End If
                        End If
                    Else
                        If Me.TxtUP_UserCode.Text.Trim <> "" Then
                            'filter berdsarkaan user code
                            If Me.TxtUP_UserCode.Text.Trim = Me.vrowDODummy.Created_By Or _
                            Me.TxtUP_UserCode.Text.Trim = Me.vrowDODummy.SO_Created_By Then
                                Lanjut = True
                            End If
                        Else
                            'tidak filter
                            Lanjut = True
                        End If
                    End If
                    If Lanjut = True Then
                        Me.dgvUP.Rows.Add()

                        Me.rowDGV = Me.dgvUP.Rows(Me.dgvUP.Rows.Count - 1)
                        With Me.rowDGV
                            .Cells(Me.ColUP_CustCode.Index).Value = Me.vrowDODummy.Customer_Code
                            .Cells(Me.ColUp_Customer.Index).Value = Me.vrowDODummy.Customer_Name
                            .Cells(Me.ColUP_Delay.Index).Value = (Today - Me.vrowDODummy.Tanggal_Transaksi.Date).TotalDays
                            If CDec(.Cells(Me.ColUP_Delay.Index).Value) < 0 Then
                                .Cells(Me.ColUP_Delay.Index).Value = "-"
                            End If
                            .Cells(Me.ColUP_DODate.Index).Value = Me.vrowDODummy.Tanggal_Transaksi
                            .Cells(Me.ColUP_DONumber.Index).Value = Me.vrowDODummy.No_Transaksi
                            .Cells(Me.ColUP_No.Index).Value = Me.dgvUP.Rows.Count
                        End With
                    End If
                End If
            Next

            Me.ColUP_DODate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.dgvUP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvUP.ClearSelection()
            If Me.dgvUP.Rows.Count > 0 Then
                Me.dgvUP.Rows(0).Selected = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
        End Try
    End Sub


    Private Sub dgvUP_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvUP.SelectionChanged
        Dim DONumber As String = ""
        Dim ExpiryStr As String = ""
        Dim AdaBatch As Boolean = False
        Dim AdaExpiry As Boolean = False
        Dim DecMulti As Integer = 0

        Try
            Me.dgvDOItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvDOItem.Rows.Clear()

            If Me.dgvUP.SelectedCells.Count > 0 Then
                Me.rowDGV = Me.dgvUP.Rows(Me.dgvUP.SelectedCells(0).RowIndex)
                DONumber = Me.rowDGV.Cells(Me.ColUP_DONumber.Index).Value

                Me.vTBLDODummy.DefaultView.RowFilter = "[No Transaksi] = '" + DONumber + "'"
                For Each DV As DataRowView In Me.vTBLDODummy.DefaultView
                    Me.vrowDODummy = DV.Row
                    Me.dgvDOItem.Rows.Add()
                    Me.rowDGV = Me.dgvDOItem.Rows(Me.dgvDOItem.Rows.Count - 1)
                    DecMulti = Me.ObjProduct.GetAngkaDecimal(Me.vrowDODummy.Primary_Product, Me.vrowDODummy.Primary_Units_Multi)
                    If vrowDODummy.Batch_No.Trim <> "" Then
                        AdaBatch = True
                    End If
                    If vrowDODummy.Expiry_Date <> #1/1/1950# Then
                        AdaExpiry = True
                    End If
                    With Me.rowDGV
                        .Cells(Me.ColUPItem_BatchNo.Index).Value = Me.vrowDODummy.Batch_No
                        .Cells(Me.ColUPItem_CodeNo.Index).Value = Me.vrowDODummy.Product_Code
                        .Cells(Me.ColUPItem_Desc.Index).Value = Me.vrowDODummy.Product_Desc
                        If Me.vrowDODummy.Expiry_Date = #1/1/1950# Then
                            ExpiryStr = "-"
                        Else
                            ExpiryStr = Format(Me.vrowDODummy.Expiry_Date, "dd MMM yyyy")
                        End If
                        .Cells(Me.ColUPItem_ExpiryDate.Index).Value = ExpiryStr
                        .Cells(Me.ColUPItem_No.Index).Value = Me.dgvDOItem.Rows.Count
                        .Cells(Me.ColUPItem_Remarks.Index).Value = Me.vrowDODummy.Remarks
                        .Cells(Me.ColUPItem_SONO.Index).Value = Me.vrowDODummy.So_Number
                        .Cells(Me.ColUPItem_TotalQty.Index).Value = Me.ObjCd.DigitGrouping(Me.vrowDODummy.Total_Qty_Multi, DecMulti)
                        .Cells(Me.ColUPItem_UnitCode.Index).Value = Me.vrowDODummy.Unit_Code_Multi
                        .Cells(Me.ColUPItem_TotalQty_Default.Index).Value = Me.ObjCd.DigitGrouping(Me.vrowDODummy.Total_Qty, Me.vrowDODummy._Decimal)
                        .Cells(Me.ColUPItem_UnitDefault.Index).Value = Me.vrowDODummy.Unit_Code
                    End With
                Next
            End If

            If AdaBatch = True Then
                Me.ColUPItem_BatchNo.Visible = True
            Else
                Me.ColUPItem_BatchNo.Visible = False
            End If
            If AdaExpiry = True Then
                Me.ColUPItem_ExpiryDate.Visible = True
            Else
                Me.ColUPItem_ExpiryDate.Visible = False
            End If
            Me.dgvDOItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDOItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDOItem.Sorted
        Try
            Me.ObjCd.SetNumber(Me.dgvDOItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvUP_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvUP.Sorted
        Try
            Me.ObjCd.SetNumber(Me.dgvUP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Update SMESS 2026
    'Start
    Private Sub dgvUP_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvUP.CellClick
        Try
            If e.ColumnIndex = Me.ColUP_ChkDelete.Index Then
                If Me.dgvUP.Item(e.ColumnIndex, e.RowIndex).Value = False Then
                    Me.dgvUP.Item(e.ColumnIndex, e.RowIndex).Value = True
                Else
                    Me.dgvUP.Item(e.ColumnIndex, e.RowIndex).Value = False
                End If
                Me.KontrolLinkLabel(e.RowIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub KontrolLinkLabel(ByVal vIndex As Integer)
        Dim ij As Integer = 0
        Dim RowDGV As DataGridViewRow
        Dim Ketemu As Boolean = False
        Try
            While ij <= Me.dgvUP.Rows.Count - 1
                RowDGV = Me.dgvUP.Rows(ij)
                If vIndex <> ij Then
                    If RowDGV.Cells(Me.ColUP_ChkDelete.Index).Value = False Then
                        Ketemu = True
                        Exit While
                    End If
                End If
                ij += 1
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUP_AllDelete_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUP_CheckedAllDelete.LinkClicked
        Dim RowDGV As DataGridViewRow
        Try
            For Each RowDGV In Me.dgvUP.Rows
                RowDGV.Cells(Me.ColUP_ChkDelete.Index).Value = True
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUP_UncheckedAllDelete_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUP_UncheckedAllDelete.LinkClicked
        Dim RowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvUP.Rows
                rowDGV.Cells(Me.ColUP_ChkDelete.Index).Value = False
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnUP_Delete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnUP_Delete.Click
        Dim listNoDO As New List(Of String)
        Dim i As Integer = 0
        Dim rowdgvUP As DataGridViewRow
        Dim ObjDO As New ClsTransaksiDO

        'Update SMESS 2026
        Dim ObjDO_C As New CustomSMESS.ClsTransaksiDO
        Try
            For Each rowdgvUP In Me.dgvUP.Rows
                With rowdgvUP
                    If .Cells(Me.ColUP_ChkDelete.Index).Value = True Then
                        If listNoDO.Contains(.Cells(Me.ColUP_DONumber.Index).Value) = False Then
                            listNoDO.Add(.Cells(Me.ColUP_DONumber.Index).Value)
                        End If
                    End If
                End With
            Next

            If listNoDO.Count > 0 Then
                'Update SMESS 2026, tambah otorisasi
                If ObjDO_C.ApakahBolehHapus_Custom_UP(listNoDO, Me.UserCode) = False Then
                    Exit Sub
                End If

                If MsgBox("Delete Data?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                    Exit Sub
                End If
                Dim j As Integer = 0
                For i = 0 To listNoDO.Count - 1
                    Me.vTBLDODummy.DefaultView.RowFilter = "[No Transaksi] = '" + listNoDO(i).ToString + "'"
                    If Me.vTBLDODummy.DefaultView.Count > 0 Then
                        Me.vrowDODummy = Me.vTBLDODummy.DefaultView(0).Row
                        'delete data
                        If ObjDO.ApakahBolehHapus(Me.vrowDODummy.Primary_DO, Me.UserCode, False, False, True) Then
                            ObjDO.Hapus(Me.vrowDODummy.Primary_DO, Me.UserCode, False, False, True)
                            j += 1
                        End If
                    End If
                Next
                If j >= 1 Then
                    MsgBox("Data deleted.", MsgBoxStyle.Information, "Information")
                End If
                Me.btnUP_Filter.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            listNoDO.Clear()
            ObjDO = Nothing
        End Try
    End Sub

    'End

End Class
