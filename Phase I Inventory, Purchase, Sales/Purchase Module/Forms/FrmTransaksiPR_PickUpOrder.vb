Option Explicit On
Option Strict Off
Public Class FrmTransaksiPR_PickUpOrder

    Public UserCode As String = ""
    Public BAckColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjPR As New ClsTransaksiPR
    Dim ObjPR_PickUp As New ClsTransaksiPR_PickUpOrder
    Dim ObjPR_PickUp_Item As New ClsTransaksiPR_PickUpOrder_Item
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjRefParty As New Core_Module.ClsMasterRefParty
    Dim SetTab As Boolean
    Dim baru As Boolean
    Dim BufferQty As Decimal = 0
    Dim Buffer As String = ""
    Dim EraseRowIndex As Integer = 0
    Dim vReadyForPurchase As ClsTransaksiPR.EnumReadyForPurchase
    Dim vPrimaryUser As Double = -1
    Dim DefPrec As Integer = 0
    Dim vTBLPRItem As New DseTransaksiPR_PickUpOrder.TblPRDataTable
    Dim vRowPRItem As DseTransaksiPR_PickUpOrder.TblPRRow
    Dim vTBLPR As New DseTransaksiPR_PickUpOrder.TblPR1DataTable
    Dim vRowPR As DseTransaksiPR_PickUpOrder.TblPR1Row

    Private Sub AktifkanDetail()
        Try
            ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)
            ObjCD.KendalikanControl(Me.TLP1.Controls, True, True, True, True, True, True, True, True, True, True)
            ObjCD.KendalikanControl(Me.TLP2.Controls, True, True, True, True, True, True, True, True, True, True)
            ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimary.BackColor, False)
            ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimary.BackColor, False)
            ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimary.BackColor, False)
            ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimary.BackColor, False)
            ObjCD.SetBackColor(Me.TLP1.Controls, Me.TxtPrimary.BackColor, False)
            ObjCD.SetBackColor(Me.TLP2.Controls, Me.TxtPrimary.BackColor, False)
            ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimary.BackColor, False)
            ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryPR.BackColor, False)
            Me.DgvItem.AllowUserToDeleteRows = True
            Me.ObjCD.SetNumber(Me.DgvItem, False)
            Me.DtpTglTransaksi.Focus()
            Me.SplitContainer1.TabStop = False
            Me.SplitContainer2.TabStop = False
            Me.DgvItem.ReadOnly = False
            Me.TxtQty.Text = "0"
            Me.TxtRemarksPR.Clear()
            If Me.baru = True Then
                Me.TxtBalance.Text = "0"
                Me.TxtRequestBy.ReadOnly = False
                Me.TxtRequestBy.BackColor = Me.TxtRemarks.BackColor
                Me.BtnFindRefParty.Enabled = True
            Else
                If Me.DgvItem.RowCount = 0 Then
                    Me.TxtRequestBy.ReadOnly = False
                    Me.TxtRequestBy.BackColor = Me.TxtRemarks.BackColor
                    Me.BtnFindRefParty.Enabled = True
                Else
                    Me.TxtRequestBy.ReadOnly = True
                    Me.TxtRequestBy.BackColor = Me.TxtReadOnly.BackColor
                    Me.BtnFindRefParty.Enabled = False
                End If
            End If
            If Me.TxtRequestBy.ReadOnly = False Then
                Me.TxtRequestBy.TabStop = True
            Else
                Me.TxtRequestBy.TabStop = False
            End If
            If Me.DgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            If baru = True Then
                Me.DtpTglTransaksi.Value = Now
                Me.TxtNoTransaksi.Text = ObjPR_PickUp.GetNoBaru(Me.DtpTglTransaksi.Value)
            End If
            If Me.TxtPrimaryRefParty.Text <> "-1" And Me.TxtPrimaryRefParty.Text <> "" Then
                Me.IsiCboPRNo()
            Else
                Me.CboPRNo.DataSource = Nothing
                Me.CboPRNo.Items.Clear()
            End If
            'If Me.TxtBalance.Text = "" Then
            '    Me.TxtBalance.Text = "0"
            'End If
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.llbSendEmail.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            ObjCD.KendalikanControl(Me.TLP1.Controls, True, False, True, False, True, False, True, False)
            ObjCD.KendalikanControl(Me.TLP2.Controls, True, False, True, False, True, False, True, False)
            ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)

            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TLP1.Controls, TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TLP2.Controls, TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, TxtReadOnly.BackColor, False)
            Me.SplitContainer1.TabStop = True
            Me.SplitContainer2.TabStop = True

            Me.CboPRNo.DataSource = Nothing
            Me.CboPRNo.Items.Clear()
            Me.DgvItem.ReadOnly = True
            Me.DgvItem.AllowUserToAddRows = False
            Me.DgvItem.AllowUserToDeleteRows = False
            Me.DgvItem.TabStop = False
            Me.CancelButton = Me.btnExit

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockOutAdj") = False Then
                Me.btnGenerateStockOutAdjust.Enabled = False
                Me.btnGenerateStockOutAdjust.TabStop = False
            Else
                Me.btnGenerateStockOutAdjust.Enabled = True
                Me.btnGenerateStockOutAdjust.TabStop = True
            End If
            Me.llbSendEmail.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '''validasi
    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim KoleksiNamaData1 As New Collection
        Dim KoleksiTextBox1 As New Collection
        Try
            Me.IsiCboPRNo()

            '''cek data yang harus diisi
            KoleksiNamaData.Add("Request By")
            KoleksiTextBox.Add(Me.TxtRequestBy)
            If Core_Module.ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            '''khusus untuk remarks saja
            KoleksiNamaData1.Add("Remarks")
            KoleksiTextBox1.Add(Me.TxtRemarks)
            If Me.CekDataRemarks(KoleksiTextBox1, KoleksiNamaData1) = False Then
                Return False
            End If
            '''cek data yang salah, double dalam kode
            If Me.baru = True Then
                'If ObjPR_PickUp.ApakahSudahAda(Me.TxtNoTransaksi.Text) = True Then
                '    MsgBox("Duplication in document no, please change your document no", _
                '    MsgBoxStyle.Information, "Information")
                '    Me.DtpTglTransaksi.Focus()
                '    Return False
                'End If
            Else
                'Dim ProductGroupLama As String = Me.ObjPR_PickUp.GetNoTransaksi(Me.TxtPrimary.Text)
                'If ProductGroupLama <> Me.TxtNoTransaksi.Text Then
                '    If Me.ObjPR_PickUp.ApakahSudahAda(Me.TxtNoTransaksi.Text) = True Then
                '        MsgBox("Duplication in document no, please change your document no", _
                '        MsgBoxStyle.Information, "Information")
                '        Me.DtpTglTransaksi.Focus()
                '        Return False
                '    End If
                'End If

                '''data salah jika nilai timestamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                '''proses edit
                Dim OldTimeStamp As Date = Me.ObjPR_PickUp.GetTimestamp(Me.TxtPrimary.Text)
                If Me.ObjPR_PickUp.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjPR_PickUp.CariDetail(Me.ObjPR_PickUp.Primary, Me.DgvItem, Me.ObjCD, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            If Me.DgvItem.Rows.Count = 0 Then
                MsgBox("Please input product code", MsgBoxStyle.Information, "Information ")
                Return False
            End If

            ''cek kesesuain data pr pickup dengan pr 
            Dim row As DataGridViewRow
            For Each row In Me.DgvItem.Rows
                If ObjPR.GetJmlByPrimaryCode(CDec(row.Cells(Me.ColItem_PrimaryPR.Index).Value), CDec(row.Cells(Me.ColItem_PrimaryCodeNo.Index).Value)) <= 0 Then
                    MsgBox("Product code is nothing this PR Number", MsgBoxStyle.Information, "Information ")
                    Return False
                End If
            Next

            If Me.ObjPR_PickUp.ApakahQtySudahBenar(Me, Me.TxtPrimary.Text) = False Then
                Return False
            End If

            If ObjPR_PickUp.ApakahTanggalSudahBenar(Me.TxtPrimaryPR.Text, Me.DtpTglTransaksi.Value) = False Then
                Me.DtpTglTransaksi.Focus()
                Return False
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

    Private Function ApakahDataItemBenar() As Boolean
        Dim KoleksiNamaData1 As New Collection
        Dim KoleksiTextBox1 As New Collection
        Try
            'Me.IsiCboPRNo()
            If Me.TxtBalance.Text = "" Or Me.TxtBalance.Text = "0" Then
                MsgBox("Can't item for input", MsgBoxStyle.OkOnly, "Information")
                Return False
            End If

            DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCode.Text)
            '''cek qty hrs positive
            If Me.ObjCD.CekBilanganPositif(Me.TxtQty.Text) = False Or Me.TxtQty.Text = 0 Then
                MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.OkOnly, _
                                           "Information")
                Me.TxtQty.Focus()
                Me.TxtQty.Text = Me.ObjCD.DigitGrouping(0, DefPrec)
                Return False
            End If
            If ObjPR_PickUp.ApakahBalanceSudahBenar(CDec(Me.GetOB_ViaTable(Me.TxtPrimaryPR.Text, Me.TxtPrimaryCode.Text)), Me.TxtPrimaryPR.Text, Me.TxtPrimaryCode.Text, Me) = False Then
                Me.IsiCboPRNo()
                Me.CboCode.Focus()
                Return False
            End If
            If CDec(Me.TxtBalance.Text) < CDec(Me.TxtQty.Text) Then
                MsgBox("Maximum value you can enter is" & " " & Me.TxtBalance.Text & " " & Me.LblUnits_Hidden.Text & ".", MsgBoxStyle.OkOnly, "Information")
                Me.TxtQty.Focus()
                Return False
            End If
            '''cek remarks
            KoleksiNamaData1.Add("Remarks")
            KoleksiTextBox1.Add(Me.TxtRemarksPR)

            If Me.CekDataRemarks(KoleksiTextBox1, KoleksiNamaData1) = False Then
                Return False
                Me.TxtRemarks.Focus()
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem(True)
            ObjCD.KosongkanControl(Me.gboDetail.Controls, True, False, True)
            ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, False, True)
            ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, False, True)
            ObjCD.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, False, True)
            ObjCD.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, False, True)

            Me.DgvItem.Rows.Clear()
            Me.TxtPrimary.Text = "-1"
            Me.TxtPrimaryPR.Text = "-1"
            Me.TxtPrimaryCode.Text = "-1"
            Me.TxtPrimaryRefParty.Text = "-1"
            Me.TxtPrimaryStockInPurchase.Text = "-1"
            Me.TxtPrimaryStockOutAdjust.Text = "-1"

            Me.LblUnit.Text = "Qty (Units)"

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem(ByVal vKosongkanCBO As Boolean)
        Try
            If vKosongkanCBO = True Then
                Me.CboPRNo.DataSource = Nothing
                Me.CboCode.Items.Clear()
                Me.TxtDesc.Clear()
                Me.LblUnit.Text = "Qty (Units)"
                Me.TxtBalance.Text = "0"
            End If
            Me.TxtRemarksPR.Clear()
            If Me.TxtPrimaryCode.Text = "" Or Me.TxtPrimaryCode.Text = "-1" Then
                Me.TxtQty.Text = "0"
            Else
                Me.TxtQty.Text = Me.ObjCD.DigitGrouping(0, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCode.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.btnFind.Select()
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.KosongkanField()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.DtpTglTransaksi.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.SetStockOutAdjust()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.DgvItem.Rows.Count >= 1 And Me.TxtRequestBy.Text <> "" Then
                '''batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjPR_PickUp.CariDetail(Me.ObjPR_PickUp.Primary, Me.DgvItem, Me.ObjCD, Me)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnNew.Focus()
            Me.SetStockOutAdjust()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRequestBy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRequestBy.GotFocus, TxtRemarks.GotFocus, TxtQty.GotFocus, TxtRemarksPR.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.BAckColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TxtQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQty.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtQty, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQty.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtQty, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty.Leave
        Try
            If Me.CboCode.SelectedIndex <> -1 Then
                Dim DefPrec As Decimal = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCode.Text)
            End If
            Me.ObjCD.Number_Leave(Me.TxtQty, sender, e, DefPrec, ObjCD)
            If ObjCD.CekNonNegatif(CDec(Me.TxtQty.Text)) = False Then
                MsgBox("Quantity value can not be negative.", MsgBoxStyle.OkOnly, _
                                           "Information")
                Me.TxtQty.Text = "0"
                Me.TxtQty.Focus()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQty, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty.Click
        Try
            Me.ObjCD.Number_Click(Me.TxtQty)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRequestBy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRequestBy.LostFocus, TxtRemarks.LostFocus, TxtQty.LostFocus, TxtRemarksPR.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimaryPR.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            With Me.ObjPR_PickUp
                .Primary = Me.TxtPrimary.Text
                .PrimaryRefParty = Me.TxtPrimaryRefParty.Text
                .NoTransaksi = Me.TxtNoTransaksi.Text
                .Tanggal = Me.DtpTglTransaksi.Value
                .Remarks = Me.TxtRemarks.Text
                .PrimaryStockOutAdjust = Me.TxtPrimaryStockOutAdjust.Text
                .PrimaryStockInPurchase = Me.TxtPrimaryStockInPurchase.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjPR_PickUp
                Me.TxtPrimary.Text = .Primary
                Me.TxtPrimaryRefParty.Text = .PrimaryRefParty
                Me.TxtRequestBy.Text = Me.ObjRefParty.GetCode(.PrimaryRefParty)
                Me.DtpTglTransaksi.Text = Format(.Tanggal, "dd MMM yyyy")
                Me.TxtNoTransaksi.Text = .NoTransaksi
                Me.TxtRequestBy.Text = Me.ObjRefParty.GetCode(.PrimaryRefParty)
                Me.TxtRequestByName.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                Me.TxtRemarks.Text = .Remarks
                Me.TxtPrimaryStockOutAdjust.Text = .PrimaryStockOutAdjust
                Me.TxtStockOutAdjustNo.Text = Me.ObjStockInOut.GetNoTransaksi(Me.TxtPrimaryStockOutAdjust.Text)
                Me.TxtPrimaryStockInPurchase.Text = .PrimaryStockInPurchase
                Me.TxtStockInPurchaseNo.Text = Me.ObjStockInOut.GetNoTransaksi(Me.TxtPrimaryStockInPurchase.Text)

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '''Cari Data Keydown & Leave & Click (master reff party)
    Private Sub TxtRequestBy_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRequestBy.KeyDown
        Dim ObjRefParty As New Core_Module.ClsMasterRefParty
        Try
            If e.KeyValue = 13 Then
                If Me.TxtRequestBy.Text <> "" Then
                    If ObjRefParty.ApakahSudahAda(Me.TxtRequestBy.Text, 1) = True Then
                        Me.TxtPrimaryRefParty.Text = ObjRefParty.GetPrimary(Me.TxtRequestBy.Text)
                        Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                        Me.TxtRequestByName.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                        If Me.TxtRequestBy.ReadOnly = False Then
                            Me.ResetCbo()
                            Me.IsiCboPRNo()
                        End If
                        Me.TxtRemarks.Focus()
                    Else
                        Me.BtnFindRefParty.PerformClick()
                    End If
                Else
                    Me.BtnFindRefParty.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRequestBy_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRequestBy.Leave
        Dim ObjRefParty As New Core_Module.ClsMasterRefParty
        Dim hj As New Core_Module.FrmMasterRefParty_Cari
        hj.Aktif = True
        hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.TxtRequestBy.Text <> "" Then
                If ObjRefParty.ApakahSudahAda(Me.TxtRequestBy.Text, 1) = True Then
                    Me.TxtPrimaryRefParty.Text = ObjRefParty.GetPrimary(Me.TxtRequestBy.Text)
                    Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                    Me.TxtRequestByName.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                    If Me.TxtRequestBy.ReadOnly = False Then
                        Me.ResetCbo()
                        Me.IsiCboPRNo()
                    End If
                    Me.TxtRemarks.Focus()
                ElseIf Me.ObjRefParty.ApakahSudahAda(Me.TxtRequestBy.Text, 1) = False And Me.TxtRequestBy.Enabled = True Then
                    If Me.TxtRequestBy.Text <> "" Then
                        hj.txtCari.Text = Me.TxtRequestBy.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryRefParty > 0 Then
                        Me.TxtPrimaryRefParty.Text = hj.PrimaryRefParty
                        Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                        Me.TxtRequestByName.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                        If Me.TxtRequestBy.ReadOnly = False Then
                            Me.ResetCbo()
                            Me.IsiCboPRNo()
                        End If
                        Me.TxtRemarks.Focus()
                    Else
                        Me.TxtPrimaryRefParty.Text = "-1"
                        Me.TxtRequestBy.Clear()
                        Me.TxtRequestByName.Clear()
                        Me.TxtRequestBy.Focus()
                        Me.ResetCbo()
                    End If
                End If
            Else
                Me.TxtPrimaryRefParty.Text = "-1"
                Me.TxtRequestBy.Clear()
                Me.TxtRequestByName.Clear()
                Me.ResetCbo()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindRefParty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindRefParty.Click
        Dim ObjRefParty As New Core_Module.ClsMasterRefParty
        Dim hj As New Core_Module.FrmMasterRefParty_Cari
        hj.Aktif = True

        Try
            hj.BackColorAll = Me.BAckColorAll
            If Me.TxtRequestBy.Text <> "" Then
                hj.txtCari.Text = Me.TxtRequestBy.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryRefParty > 0 Then
                Me.TxtPrimaryRefParty.Text = hj.PrimaryRefParty
                Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                Me.TxtRequestByName.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                If Me.TxtRequestBy.ReadOnly = False Then
                    Me.ResetCbo()
                    Me.IsiCboPRNo()
                End If
                Me.TxtRemarks.Focus()
            Else
                Me.TxtPrimaryRefParty.Text = "-1"
                Me.ResetCbo()
                Me.TxtRequestByName.Clear()
                Me.TxtRequestBy.Clear()
                Me.TxtRequestBy.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    '''disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjPR_PickUp.SimpanBaru(Me.DgvItem, Me.UserCode, Me, Me.baru) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    '''disimpan atau tidak
                    If Me.ObjPR_PickUp.ApakahBolehEdit(Me, Me.TxtPrimary.Text, Me.UserCode, Me.TxtNoTransaksi.Text, True) = True Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.ObjPR_PickUp.SimpanPerubahan(Me.DgvItem, Me.UserCode, Me, Me.baru) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            End If
                        End If
                    End If
                End If
            End If
            Me.SetStockOutAdjust()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTransaksiPR_PickUpOrder_Cari
        Try
            hj.BackColorAll = Me.BAckColorAll
            hj.ShowDialog()
            If Not hj.PrimaryPR_PickupOrder = -1 Then
                If Me.ObjPR_PickUp.CariDetail(hj.PrimaryPR_PickupOrder, Me.DgvItem, Me.ObjCD, Me) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
                Me.NonAktifkanDetail()
            End If
            Me.SetStockOutAdjust()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimary.Text <> "-1" Then
                If Me.ObjPR_PickUp.ApakahBolehEdit(Me, Me.TxtPrimary.Text, UserCode, Me.TxtNoTransaksi.Text, True) = True Then
                    Me.baru = False
                    Me.ObjPR_PickUp.CariDetail(Me.TxtPrimary.Text, Me.DgvItem, Me.ObjCD, Me)
                    Me.AktifkanDetail()
                    Me.SalinKeField()
                    Me.DgvItem.ClearSelection()
                    Me.DtpTglTransaksi.Focus()
                End If
            End If
            Me.SetStockOutAdjust()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimary.Text <> "-1" Then
                If Me.ObjPR_PickUp.ApakahBolehHapus(Me.TxtPrimary.Text, UserCode, Me, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjPR_PickUp.Hapus(Me.TxtPrimary.Text, Me.UserCode, Me) = True Then
                            Me.KosongkanField()
                            Me.NonAktifkanDetail()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
            Me.SetStockOutAdjust()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmMasterLoad_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtNoTransaksi.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            ''' simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPR_PickUp.SimpanBaru(Me.DgvItem, Me.UserCode, Me, Me.baru) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            '''' simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPR_PickUp.SimpanPerubahan(Me.DgvItem, Me.UserCode, Me, Me.baru) = True Then
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
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Function HapusSpasi(ByVal Data As String) As String
        Try
            Do While Data.EndsWith(" ") And Data.Length > 0
                Data = Data.Remove(Data.Length - 1, 1)
            Loop
            Return Data
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function CekDataRemarks(ByVal KoleksiTextBox As Collection, ByVal KoleksiNamaData As Collection) As Boolean
        Dim i As Int16
        Dim data As String
        Dim Txt As Windows.Forms.TextBox

        Try
            Do While i < KoleksiTextBox.Count
                '''trim kanan kiri
                Txt = TryCast(KoleksiTextBox.Item(i + 1), Windows.Forms.TextBox)
                If IsNothing(Txt) = False Then
                    Txt.Text = Txt.Text.Trim
                End If
                '''alin data
                data = KoleksiTextBox.Item(i + 1).text
                data = HapusSpasi(data)

                ''cek data
                'If data.Length = 0 Then
                '    MsgBox(Convert.ToString(KoleksiNamaData.Item(i + 1)) + " cannot be an empty string.", MsgBoxStyle.Critical, "Error")
                '    KoleksiTextBox.Item(i + 1).focus()
                '    Return False
                'End If

                '''cek berkaitan dengan karakter yang tidak diingikan --> ', *, %
                If data.Contains("'") = True Or data.Contains("*") = True Or data.Contains("%") = True Then
                    MsgBox("" + KoleksiNamaData.Item(i + 1) _
                        + " data can not contain the following characters: ' , % , or * .", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                i += 1
            Loop

            '''sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

   
    'Public Sub IsiCboPRNo()

    '    Try
    '        Me.CboPRNo.DataSource = Nothing
    '        Me.CboPRNo.Items.Clear()
    '        If Me.TxtPrimaryRefParty.Text <> "-1" Then
    '            Me.ObjPR_PickUp.IsiTblPR(Me.vTBLPR, Me.TxtPrimaryRefParty.Text)
    '            Me.CboPRNo.DataSource = Me.vTBLPR.DefaultView
    '            Me.vTBLPR.DefaultView.Sort = "urutan"
    '            Me.CboPRNo.DisplayMember = "no transaksi"
    '        End If
    '        Me.CboPRNo.SelectedIndex = -1
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Public Sub IsiCboPRNo()
        Try
            Me.CboPRNo.DataSource = Nothing
            Me.CboPRNo.Items.Clear()
            If Me.TxtPrimaryRefParty.Text <> "-1" Then
                Me.ObjPR_PickUp.IsiTblPR(Me.vTBLPRItem, Me.TxtPrimaryRefParty.Text, Me.TxtPrimary.Text, Me.DgvItem, Me, Me.vTBLPR)
                Me.CboPRNo.DataSource = Me.vTBLPR.DefaultView
                Me.vTBLPR.DefaultView.Sort = "tgl, [no transaksi]"
                Me.CboPRNo.DisplayMember = "no transaksi"
            End If
            'pilih item pertama PR
            If Me.CboPRNo.Items.Count > 0 Then
                Me.CboPRNo.SelectedIndex = 0
            End If
            'Me.CboPRNo.SelectedIndex = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiCboCode()
        Try
            Me.CboCode.DataSource = Nothing
            Me.CboCode.Items.Clear()
            If Me.TxtPrimaryPR.Text <> "-1" Then
                Me.vTBLPRItem.DefaultView.RowFilter = "primary = " + Me.TxtPrimaryPR.Text
                Me.vTBLPRItem.DefaultView.Sort = "[Product Code]"
                Me.CboCode.DataSource = Me.vTBLPRItem.DefaultView
                Me.CboCode.DisplayMember = "Product Code"
            End If
            If Me.CboCode.Items.Count > 0 Then
                Me.CboCode.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub CboCode_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboCode.DropDownClosed
    '    Try
    '        If Me.CboCode.SelectedIndex <> -1 Then
    '            Me.TxtQty.Focus()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Public Sub ResetCbo()
        Try
            Me.CboPRNo.DataSource = Nothing
            Me.CboPRNo.Items.Add("item kosong")
            Me.CboPRNo.Items.Clear()
            Me.CboPRNo.SelectedIndex = -1
            Me.CboCode.DataSource = Nothing
            Me.CboCode.Items.Add("item kosong")
            Me.CboCode.Items.Clear()
            Me.CboCode.SelectedIndex = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '''up & down
    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.SelectionChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.DgvItem.SelectedRows.Count > 0 Then
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowRowProcess(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.DgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.DgvItem.FirstDisplayedScrollingRowIndex = Me.DgvItem.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.DgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.DgvItem.Rows(row.Index).Displayed = False
                    Me.DgvItem.FirstDisplayedScrollingRowIndex = Me.DgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.DgvItem.DisplayedRowCount(True) <> Me.DgvItem.DisplayedRowCount(False) Then
                    Me.DgvItem.FirstDisplayedScrollingRowIndex = Me.DgvItem.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.DgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.DgvItem.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.DgvItem.Rows.Remove(row)
                    Me.DgvItem.Rows.InsertRange(indexbaru, row)
                    Me.DgvItem.Focus()
                    Me.DgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.DgvItem)
                    If row.Displayed = False Then
                        Me.ShowRowProcess(row, True)
                    End If
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.DgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.DgvItem.SelectedRows(0)

                If row.Index <> Me.DgvItem.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.DgvItem.Rows.Remove(row)
                    Me.DgvItem.Rows.InsertRange(indexbaru, row)
                    Me.DgvItem.Focus()
                    Me.DgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.DgvItem)
                    Me.ShowRowProcess(row, True)
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvItem.RowsRemoved
        Try
            If Me.DgvItem.Rows.Count = 0 Then
                Me.TxtRequestBy.ReadOnly = False
                Me.TxtRequestBy.BackColor = Me.TxtRemarks.BackColor
                Me.BtnFindRefParty.Enabled = True
                Me.DgvItem.TabStop = False
                If Me.TxtRemarks.Enabled = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
                Me.TxtRequestBy.Focus()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Get Index & CekDataygSama & btnInput
    Private Function GetIndex(ByVal vPrimaryPR As Double, ByVal vPrimaryProduct As Double) As Integer
        Dim row As New DataGridViewRow
        Try
            For Each row In Me.DgvItem.Rows
                If row.Cells(Me.ColItem_PrimaryPR.Index).Value = CStr(vPrimaryPR) And _
                    row.Cells(Me.ColItem_PrimaryCodeNo.Index).Value = CStr(vPrimaryProduct) Then
                    Return row.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function CekItemYgSama(ByVal vPrimaryPR As Double, ByVal vPrimaryProduct As Double, _
                                   ByVal vColoumnIdxFocus As Integer, _
                                   ByVal vRowIdxFocus As Integer, ByRef vIndexGabungan As Integer) As Boolean

        Dim Row As DataGridViewRow
        Dim Retval As Double = -1
        Try
            'Cek index
            For ij As Integer = 0 To Me.DgvItem.Rows.Count - 1
                Row = Me.DgvItem.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If Row.Cells(Me.ColItem_PrimaryPR.Index).Value = CStr(vPrimaryPR) And _
                    Row.Cells(Me.ColItem_PrimaryCodeNo.Index).Value = CStr(vPrimaryProduct) And _
                        Retval = ij Then
                        Exit For
                    End If
                End If
            Next

            If Retval >= 0 Then
                'MsgBox("Duplication data, you can't edit this data.", MsgBoxStyle.Information, "Information")
                'Me.dgvItem.Item(vColoumnIdxFocus, vRowIdxFocus).Value = Me.Buffer
                vIndexGabungan = Retval
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub GabungkanBarisYangSama(ByVal vIndexGabungan As Integer, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            Me.DgvItem.Item(Me.ColItem_Qty.Index, vIndexGabungan).Value += CDbl(Me.DgvItem.Item(Me.ColItem_Qty.Index, e.RowIndex).Value)
            Me.DgvItem.Item(Me.ColItem_Qty.Index, vIndexGabungan).Value = Me.ObjCD.DigitGrouping(Me.DgvItem.Item(Me.ColItem_Qty.Index, vIndexGabungan).Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        Dim index As Integer = -1
        Dim row As DataGridViewRow

        Try
            If Me.ApakahDataItemBenar() = True Then
                If CDbl(Me.TxtPrimaryCode.Text) <> -1 Then
                    index = Me.GetIndex(Me.TxtPrimaryPR.Text, Me.TxtPrimaryCode.Text)
                Else
                    index = Me.GetIndex(Me.TxtPrimaryPR.Text, Me.TxtPrimaryCode.Text)
                End If

                If index <> -1 Then
                    row = Me.DgvItem.Rows(index)
                    DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCode.Text)
                    row.Cells(Me.ColItem_Qty.Index).Value = ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColItem_Qty.Index).Value _
                    + CDbl(Me.TxtQty.Text)))
                Else
                    Me.DgvItem.Rows.Add()
                    row = Me.DgvItem.Rows(Me.DgvItem.RowCount - 1)
                    row.Cells(Me.ColItem_PrimaryPR.Index).Value = Me.TxtPrimaryPR.Text
                    row.Cells(Me.ColItem_PrimaryCodeNo.Index).Value = Me.TxtPrimaryCode.Text
                    row.Cells(Me.ColItem_PRNo.Index).Value = Me.CboPRNo.Text
                    row.Cells(Me.ColItem_CodeNo.Index).Value = Me.CboCode.Text
                    row.Cells(Me.ColItem_Desc.Index).Value = Me.TxtDesc.Text
                    Dim DefPrec As Decimal = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCode.Text)
                    row.Cells(Me.ColItem_Qty.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtQty.Text, DefPrec)
                    row.Cells(Me.ColItem_Unit.Index).Value = ObjProduct.GetUnitCode(Me.TxtPrimaryCode.Text)
                    row.Cells(Me.ColItem_Remarks.Index).Value = Me.TxtRemarksPR.Text
                    row.Cells(Me.ColItem_PrimaryUnit.Index).Value = Me.ObjProduct.GetPrimaryUnit(Me.TxtPrimaryCode.Text)
                    Me.ObjCD.SetNumber(Me.DgvItem)
                End If

                'kurangkan di dalam tabel item PR
                For Each Me.vRowPRItem In Me.vTBLPRItem.Rows
                    If Me.vRowPRItem.primary_code_no = Me.TxtPrimaryCode.Text And Me.TxtPrimaryPR.Text = Me.vRowPRItem.primary Then
                        Me.vRowPRItem.OB -= CDec(Me.TxtQty.Text)

                        If vRowPRItem.OB > 0 Then
                            'jika tidak habis
                            'cek apakah dia index terakhir
                            If Me.CboCode.SelectedIndex = Me.CboCode.Items.Count - 1 Then
                                'jika index terakhir, maka diem aja
                                'fokus ke cbo pr
                                Me.CboPRNo.Focus()
                            Else
                                'jika bukan index terakhir, maka majukan index-nya
                                Me.CboCode.SelectedIndex += 1
                                Me.CboCode.Focus()
                            End If
                        Else
                            'jika habis
                            'buang baris
                            Me.vTBLPRItem.Rows.Remove(Me.vRowPRItem)
                            Me.RefreshDataCodeNo()
                            'cek apakah dia item terakhir
                            If Me.CboCode.Items.Count = 0 Then
                                'jika item terakhir, maka ganti PR, buang baris pr yang primary = me.txtprimary pr
                                For Each Me.vRowPR In Me.vTBLPR.Rows
                                    If Me.vRowPR.Primary_PR = Me.TxtPrimaryPR.Text Then
                                        Me.vTBLPR.Rows.Remove(Me.vRowPR)
                                        Me.RefreshDataPR()
                                        If Me.CboCode.SelectedIndex = -1 Then
                                            Me.RefreshDataCodeNo()
                                        End If
                                        Me.CboPRNo.Focus()
                                        Exit For
                                    End If
                                Next
                            Else
                                'jika bukan item terakhir, tidak perlu majukan index
                                'diem aja
                                Me.CboCode.Focus()
                            End If
                        End If
                        Exit For
                    End If
                Next
            End If

            'tampilkan DGV
            If Me.DgvItem.Rows.Count > 0 Then
                While Me.DgvItem.Rows(Me.DgvItem.Rows.Count - 1).Displayed = False _
                          Or Me.DgvItem.DisplayedRowCount(True) <> Me.DgvItem.DisplayedRowCount(False)
                    Me.DgvItem.FirstDisplayedScrollingRowIndex = Me.DgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
            End If


            'set read only
            If Me.DgvItem.RowCount > 0 Then
                Me.TxtRequestBy.ReadOnly = True
                Me.TxtRequestBy.BackColor = Me.TxtReadOnly.BackColor
                Me.BtnFindRefParty.Enabled = False
                Me.TxtRequestBy.TabStop = False
            End If

            Me.KosongkanItem(False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '''Dgvitem
    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvItem.RowsAdded
        Try
            Me.DgvItem.TabStop = True
            'If Me.TxtRemarks.Enabled = True Then
            '    Me.btnUP.Enabled = True
            '    Me.btnDown.Enabled = True
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvItem, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvItem.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Try

            If e.ColumnIndex = Me.ColItem_Qty.Index Or e.ColumnIndex = Me.ColItem_Remarks.Index Then
                Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                If e.ColumnIndex = Me.ColItem_Qty.Index Then
                    Me.IsiCboPRNo()
                    Me.BufferQty = Me.DgvItem.Item(Me.ColItem_Qty.Index, e.RowIndex).Value
                End If
            Else
                e.Cancel = True
            End If
            'Else
            '    If e.ColumnIndex = Me.ColItem_CodeNo.Index Then
            '        Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            '    Else
            '        e.Cancel = True
            '    End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpTglTransaksi_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpTglTransaksi.ValueChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.baru = False Then
                    If Me.ObjPR_PickUp.Tanggal.Month = Me.DtpTglTransaksi.Value.Month And _
                        Me.ObjPR_PickUp.Tanggal.Year = Me.DtpTglTransaksi.Value.Year Then
                        Me.TxtNoTransaksi.Text = Me.ObjPR_PickUp.NoTransaksi
                    Else
                        Me.TxtNoTransaksi.Text = Me.ObjPR_PickUp.GetNoBaru(Me.DtpTglTransaksi.Value)
                    End If
                Else
                    Me.TxtNoTransaksi.Text = Me.ObjPR_PickUp.GetNoBaru(Me.DtpTglTransaksi.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCBO_EditMakinKecil(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim Ketemu As Boolean = False
        Dim Ketemu1 As Boolean = False
        Dim rowdgv1 As DataGridViewRow

        Try
            For Each Me.vRowPRItem In Me.vTBLPRItem.Rows
                If Me.vRowPRItem.primary_code_no = Me.DgvItem.Item(Me.ColItem_PrimaryCodeNo.Index, e.RowIndex).Value And _
                Me.vRowPRItem.primary = Me.DgvItem.Item(Me.ColItem_PrimaryPR.Index, e.RowIndex).Value Then
                    'ketemu
                    Ketemu = True

                    Me.vRowPRItem.OB -= (CDec(Me.DgvItem.Item(Me.ColItem_Qty.Index, e.RowIndex).Value) - Me.BufferQty)
                    'refresh code no update nilai balance saja
                    Me.RefreshDataCodeNo()

                    Exit For
                End If
            Next

            If Ketemu = False Then
                rowdgv1 = Me.DgvItem.Rows(e.RowIndex)

                'tambah baris baru di item
                Me.vRowPRItem = Me.vTBLPRItem.NewTblPRRow
                With Me.vRowPRItem
                    .no_transaksi = rowdgv1.Cells(Me.ColItem_PRNo.Index).Value
                    .OB = Me.BufferQty - CDec(rowdgv1.Cells(Me.ColItem_Qty.Index).Value)
                    .primary = rowdgv1.Cells(Me.ColItem_PrimaryPR.Index).Value
                    .primary_code_no = rowdgv1.Cells(Me.ColItem_PrimaryCodeNo.Index).Value
                    .primary_unit = rowdgv1.Cells(Me.ColItem_PrimaryUnit.Index).Value
                    .Product_Code = rowdgv1.Cells(Me.ColItem_CodeNo.Index).Value
                    .Product_Desc = rowdgv1.Cells(Me.ColItem_Desc.Index).Value
                    .tanggal = Me.ObjPR.GetTanggalTransaksi(rowdgv1.Cells(Me.ColItem_PrimaryPR.Index).Value)
                    .Unit_Code = rowdgv1.Cells(Me.ColItem_Unit.Index).Value
                    ._Decimal = Me.ObjProduct.GetAngkaDecimal(rowdgv1.Cells(Me.ColItem_PrimaryCodeNo.Index).Value)
                End With
                Me.vTBLPRItem.Rows.Add(Me.vRowPRItem)

                'kalo di tbl pr tidak ada, maka tambahkan juga baris baru
                For Each Me.vRowPR In Me.vTBLPR.Rows
                    If Me.vRowPR.Primary_PR = rowdgv1.Cells(Me.ColItem_PrimaryPR.Index).Value Then
                        Ketemu1 = True
                        Exit For
                    End If
                Next
                If Ketemu1 = False Then
                    'tambah baris di tabel pr
                    Me.vRowPR = Me.vTBLPR.NewTblPR1Row
                    Me.vRowPR.No_Transaksi = rowdgv1.Cells(Me.ColItem_PRNo.Index).Value
                    Me.vRowPR.Primary_PR = rowdgv1.Cells(Me.ColItem_PrimaryPR.Index).Value
                    Me.vRowPR.Tgl = Me.ObjPR.GetTanggalTransaksi(rowdgv1.Cells(Me.ColItem_PrimaryPR.Index).Value)
                    Me.vTBLPR.Rows.Add(Me.vRowPR)
                End If
                'load ulang cbo pr
                Me.CboPRNo.SelectedIndex = -1
                Me.CboPRNo.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub SetCBO_EditMakinBesar(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            For Each Me.vRowPRItem In Me.vTBLPRItem.Rows
                If Me.vRowPRItem.primary_code_no = Me.DgvItem.Item(Me.ColItem_PrimaryCodeNo.Index, e.RowIndex).Value And _
                Me.vRowPRItem.primary = Me.DgvItem.Item(Me.ColItem_PrimaryPR.Index, e.RowIndex).Value Then
                    Me.vRowPRItem.OB -= (CDec(Me.DgvItem.Item(Me.ColItem_Qty.Index, e.RowIndex).Value) - Me.BufferQty)
                    'If Me.vRowPRItem.OB = 0 Then
                    '    Me.vTBLPRItem.Rows.Remove(Me.vRowPRItem)
                    'End If

                    If vRowPRItem.OB > 0 Then
                        'jika tidak habis, refresh code no update nilai balance saja
                        Me.RefreshDataCodeNo()
                    Else
                        'jika habis
                        'buang baris
                        Me.vTBLPRItem.Rows.Remove(Me.vRowPRItem)
                        Me.RefreshDataCodeNo()
                        'cek apakah dia item terakhir
                        If Me.CboCode.Items.Count = 0 Then
                            'jika item terakhir, maka ganti PR, buang baris pr yang primary = me.txtprimary pr
                            For Each Me.vRowPR In Me.vTBLPR.Rows
                                If Me.vRowPR.Primary_PR = Me.TxtPrimaryPR.Text Then
                                    Me.vTBLPR.Rows.Remove(Me.vRowPR)
                                    Me.RefreshDataPR()
                                    If Me.CboCode.SelectedIndex = -1 Then
                                        Me.RefreshDataCodeNo()
                                    End If
                                    Exit For
                                End If
                            Next
                        Else
                            'jika bukan item terakhir, tidak perlu majukan index
                            'diem aja
                        End If
                    End If

                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try

    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Dim index As Integer = -1
        Try
            Dim row As DataGridViewRow = Me.DgvItem.Rows(e.RowIndex)
            With row
                If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColItem_Qty.Index Then
                        If IsNumeric(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            If IsNumeric(Me.Buffer) = True Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Else
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = 0
                            End If
                        End If
                    End If

                    If e.ColumnIndex <> Me.ColItem_Remarks.Index Then
                        If e.ColumnIndex = Me.ColItem_Qty.Index Then
                            Me.EndEditQty(sender, e, row)

                            'update nilai ob-nya
                            If Me.BufferQty > Me.DgvItem.Item(Me.ColItem_Qty.Index, e.RowIndex).Value Then
                                'diedit jadi makin kecil
                                Me.SetCBO_EditMakinKecil(e)
                            Else
                                'diedit jadi makin besar
                                Me.SetCBO_EditMakinBesar(e)
                            End If

                            Exit Sub
                        Else
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColItem_Remarks.Index Then
                        Me.DgvItem(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
                If e.ColumnIndex = Me.ColItem_Remarks.Index Then
                    If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> "" Then
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Trim(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQty(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
     ByVal row As DataGridViewRow)
        Dim vPrimaryCodeNo As Double = -1
        Dim vPrimaryPR As Double = -1
        Dim vUnit As String = ""
        vPrimaryCodeNo = Me.DgvItem.Item(Me.ColItem_PrimaryCodeNo.Index, e.RowIndex).Value
        vPrimaryPR = Me.DgvItem.Item(Me.ColItem_PrimaryPR.Index, e.RowIndex).Value
        vUnit = Me.DgvItem.Item(Me.ColItem_Unit.Index, e.RowIndex).Value
        Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(vPrimaryCodeNo)
        Try
            'Me.RefreshDataCodeNo()
            'Me.RefreshDataPR()
            'Me.IsiCboPRNo()
            'Me.IsiCboCode()

            With row
                '''untuk nilai qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.OkOnly, _
                    "Information")
                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'Dim OB As Decimal = ObjPR.GetOB(Me.DgvItem.Item(Me.ColItem_PrimaryPR.Index, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_PrimaryCodeNo.Index, e.RowIndex).Value)
                'Dim ItemDB As Decimal = ObjPR_PickUp.GetQty(Me.DgvItem.Item(Me.ColItem_PrimaryPR.Index, e.RowIndex).Value, Me.TxtPrimary.Text, Me.DgvItem.Item(Me.ColItem_PrimaryCodeNo.Index, e.RowIndex).Value)
                'Dim DefPrec As Decimal = Me.ObjProduct.GetAngkaDecimal(Me.DgvItem.Item(Me.ColItem_PrimaryCodeNo.Index, e.RowIndex).Value)
                'Dim Maximumnya As Decimal = 0
                'Dim OB_StlhEdit As Decimal = 0

                'OB_StlhEdit = OB + ItemDB - CDec(row.Cells(Me.ColItem_Qty.Index).Value)
                'If OB_StlhEdit < 0 Then
                '    Maximumnya = CDec(row.Cells(Me.ColItem_Qty.Index).Value) + OB_StlhEdit
                '    MsgBox("Maximum value you can enter is" & " " & Me.ObjCD.DigitGrouping(Maximumnya, DefPrec) & _
                '    " " & Me.DgvItem.Item(Me.ColItem_Unit.Index, e.RowIndex).Value & " ", MsgBoxStyle.OkOnly, "Information")
                '    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                '    Exit Sub
                'End If

                Dim DGVQty As Decimal = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Dim OB As Decimal = Me.GetOB_ViaTable(vPrimaryPR, vPrimaryCodeNo)
                Dim ItemDB As Decimal = Me.ObjPR_PickUp.GetQty(vPrimaryPR, Me.TxtPrimary.Text, vPrimaryCodeNo)
                Dim Maximumnya As Decimal = 0
                Dim OB_StlhEdit As Decimal = 0

                OB_StlhEdit = OB - (CDec(Me.DgvItem.Item(Me.ColItem_Qty.Index, e.RowIndex).Value) - Me.BufferQty)
                If OB_StlhEdit < 0 Then
                    Maximumnya = DGVQty + OB_StlhEdit
                    MsgBox("Maximum value you can enter is" & " " & Me.ObjCD.DigitGrouping(Maximumnya, DefPrec) & _
                    " " & vUnit.ToLower & " ", MsgBoxStyle.OkOnly, "Information")
                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                '''digit grouping
                Me.DgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.DgvItem(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetOB_ViaTable(ByVal vPrimaryPR As Double, ByVal vPrimaryCodeNo As Double) As Decimal
        Dim rowD As DseTransaksiPR_PickUpOrder.TblPRRow
        Try
            For Each rowD In Me.vTBLPRItem.Rows
                If rowD.primary_code_no = vPrimaryCodeNo And rowD.primary = vPrimaryPR Then
                    Return rowD.OB
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Function

    Private Sub CboPRNo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboPRNo.SelectedValueChanged
        Try
            Me.RefreshDataPR()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub RefreshDataPR()
        Try
            If Me.CboPRNo.SelectedIndex = -1 Then
                Me.TxtPrimaryPR.Text = "-1"
                Me.CboCode.DataSource = Nothing
                Me.CboCode.Items.Clear()
            Else
                vRowPR = Me.vTBLPR.DefaultView(Me.CboPRNo.SelectedIndex).Row
                Me.TxtPrimaryPR.Text = vRowPR.Primary_PR
                Me.IsiCboCode()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboCode_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboCode.SelectedValueChanged
        Try
            Me.RefreshDataCodeNo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub RefreshDataCodeNo()
        Try
            If Me.CboCode.SelectedIndex = -1 Then
                Me.TxtPrimaryCode.Text = "-1"
                Me.LblUnit.Text = "Qty (Units)"
                Me.TxtDesc.Text = ""
                Me.TxtBalance.Text = "0"
                Me.LblUnits_Hidden.Text = "Units"
                Me.TxtRemarksPR.Text = ""
            Else
                Me.vRowPRItem = Me.vTBLPRItem.DefaultView(Me.CboCode.SelectedIndex).Row
                Me.TxtPrimaryCode.Text = Me.vRowPRItem.primary_code_no
                Me.TxtDesc.Text = Me.vRowPRItem.Product_Desc
                Dim vUnit As String = Me.vRowPRItem.Unit_Code
                DefPrec = Me.vRowPRItem._Decimal
                Me.TxtBalance.Text = ObjCD.DigitGrouping(Me.vRowPRItem.OB, DefPrec)
                Me.LblUnit.Text = "Qty" & " " & "(" & vUnit & ")"
                Me.LblUnits_Hidden.Text = vUnit
                If Me.TxtPrimaryCode.Text = "" Or Me.TxtPrimaryCode.Text = "-1" Or Me.TxtPrimaryPR.Text = "" Or Me.TxtPrimaryPR.Text = "-1" Then
                    Me.TxtRemarksPR.Text = ""
                Else
                    Me.TxtRemarksPR.Text = Me.ObjPR.GetRemarksItem(Me.TxtPrimaryPR.Text, Me.TxtPrimaryCode.Text)
                End If
            End If
            'Me.TxtQty.Text = "0"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DgvItem.UserDeletingRow
        Dim Ketemu As Boolean = False
        Dim Ketemu1 As Boolean = False
        Dim rowdgv1 As DataGridViewRow

        Try
            For Each Me.vRowPRItem In Me.vTBLPRItem.Rows
                If Me.vRowPRItem.primary_code_no = Me.DgvItem.Item(Me.ColItem_PrimaryCodeNo.Index, e.Row.Index).Value And _
                Me.vRowPRItem.primary = Me.DgvItem.Item(Me.ColItem_PrimaryPR.Index, e.Row.Index).Value Then
                    'ketemu
                    Ketemu = True

                    Me.vRowPRItem.OB += CDec(Me.DgvItem.Item(Me.ColItem_Qty.Index, e.Row.Index).Value)
                    'refresh code no update nilai balance saja
                    Me.RefreshDataCodeNo()
                    Exit For
                End If
            Next

            If Ketemu = False Then
                rowdgv1 = Me.DgvItem.Rows(e.Row.Index)

                'tambah baris baru di item
                Me.vRowPRItem = Me.vTBLPRItem.NewTblPRRow
                With Me.vRowPRItem
                    .no_transaksi = rowdgv1.Cells(Me.ColItem_PRNo.Index).Value
                    .OB = CDec(rowdgv1.Cells(Me.ColItem_Qty.Index).Value)
                    .primary = rowdgv1.Cells(Me.ColItem_PrimaryPR.Index).Value
                    .primary_code_no = rowdgv1.Cells(Me.ColItem_PrimaryCodeNo.Index).Value
                    .primary_unit = rowdgv1.Cells(Me.ColItem_PrimaryUnit.Index).Value
                    .Product_Code = rowdgv1.Cells(Me.ColItem_CodeNo.Index).Value
                    .Product_Desc = rowdgv1.Cells(Me.ColItem_Desc.Index).Value
                    .tanggal = Me.ObjPR.GetTanggalTransaksi(rowdgv1.Cells(Me.ColItem_PrimaryPR.Index).Value)
                    .Unit_Code = rowdgv1.Cells(Me.ColItem_Unit.Index).Value
                    ._Decimal = Me.ObjProduct.GetAngkaDecimal(rowdgv1.Cells(Me.ColItem_PrimaryCodeNo.Index).Value)
                End With
                Me.vTBLPRItem.Rows.Add(Me.vRowPRItem)

                'kalo di tbl pr tidak ada, maka tambahkan juga baris baru
                For Each Me.vRowPR In Me.vTBLPR.Rows
                    If Me.vRowPR.Primary_PR = rowdgv1.Cells(Me.ColItem_PrimaryPR.Index).Value Then
                        Ketemu1 = True
                        Exit For
                    End If
                Next
                If Ketemu1 = False Then
                    'tambah baris di tabel pr
                    Me.vRowPR = Me.vTBLPR.NewTblPR1Row
                    Me.vRowPR.No_Transaksi = rowdgv1.Cells(Me.ColItem_PRNo.Index).Value
                    Me.vRowPR.Primary_PR = rowdgv1.Cells(Me.ColItem_PrimaryPR.Index).Value
                    Me.vRowPR.Tgl = Me.ObjPR.GetTanggalTransaksi(rowdgv1.Cells(Me.ColItem_PrimaryPR.Index).Value)
                    Me.vTBLPR.Rows.Add(Me.vRowPR)
                End If
                'load ulang cbo pr
                Me.CboPRNo.SelectedIndex = -1
                Me.CboPRNo.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim i As New FrmTransaksiPR_PickUpOrder_Cetak
        Try
            If Me.TxtPrimary.Text = -1 Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimary.Text <> -1 Then
                ObjPR_PickUp.CetakPR_PickUp_Order(Me.TxtPrimary.Text, i.crvReport)
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                'i.WindowState = FormWindowState.Maximized
                i.ShowDialog()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '''Khusus [TAB] tambahan
    Private Sub btnExit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.GotFocus
        SetTab = True
    End Sub
    Private Sub btnFind_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.LostFocus
        SetTab = False
    End Sub
    Private Sub SplitContainer1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles SplitContainer1.GotFocus
        If SetTab = False Then
            btnExit.Focus()
        Else
            btnFind.Focus()
        End If
    End Sub
    Private Sub SplitContainer2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles SplitContainer2.GotFocus
        If SetTab = False Then
            btnExit.Focus()
        Else
            btnFind.Focus()
        End If
    End Sub

    Public Sub SetStockOutAdjust()
        Try
            'cek apakah sudah ada primary stock out untuk pick up order ini
            If Me.TxtPrimaryStockOutAdjust.Text = "" Or Me.TxtPrimaryStockOutAdjust.Text = "-1" Then
                'llb show stock out visible false
                'btn generate stock out text -> create stock out
                Me.llbShowStockOut.Visible = False
                Me.btnGenerateStockOutAdjust.Text = "&Create Stock Out"
            Else
                'jika sudah ada:
                'llb show visible true
                'btn generate stock out text-> delete stock out adjust
                Me.llbShowStockOut.Visible = True
                Me.btnGenerateStockOutAdjust.Text = "&Delete Stock Out"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnGenerateStockOutAdjust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerateStockOutAdjust.Click
        Dim ij As New Inventory_Module.FrmTransaksiStockOutAdj
        Dim ListPrimary As New List(Of Double)
        Dim rowDGV As DataGridViewRow

        Try
            If Me.TxtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimary.Text <> "-1" Then
                Me.TxtPrimaryStockOutAdjust.Text = Me.ObjPR_PickUp.GetPrimaryStockOutAdjust(Me.TxtPrimary.Text)
                Me.TxtStockOutAdjustNo.Text = Me.ObjStockInOut.GetNoTransaksi(Me.TxtPrimaryStockOutAdjust.Text)
                Me.SetStockOutAdjust()

                If Me.TxtStockOutAdjustNo.Text = "" Then
                    ij.vPrimaryPickUpOrder = Me.TxtPrimary.Text
                    ij.UserCode = Me.UserCode
                    ij.BackColorAll = Me.BAckColorAll
                    ij.vFromPickUpOrder = True
                    ij.ShowDialog()

                    'update primary stock out adjust
                    Me.TxtPrimaryStockOutAdjust.Text = Me.ObjPR_PickUp.GetPrimaryStockOutAdjust(Me.TxtPrimary.Text)
                    Me.TxtStockOutAdjustNo.Text = Me.ObjStockInOut.GetNoTransaksi(Me.TxtPrimaryStockOutAdjust.Text)
                    Me.ObjPR_PickUp.TimeStamp = Me.ObjPR_PickUp.GetTimestamp(Me.TxtPrimary.Text)

                    'set tampilan
                    Me.SetStockOutAdjust()
                    For Each rowDGV In Me.DgvItem.Rows
                        If ListPrimary.Contains(rowDGV.Cells(Me.ColItem_PrimaryPR.Index).Value) = False Then
                            ListPrimary.Add(rowDGV.Cells(Me.ColItem_PrimaryPR.Index).Value)
                        End If
                    Next
                    For Each PrimaryPR As Double In ListPrimary
                        Me.ObjPR.SetOrderBalance1(PrimaryPR, ClsTransaksiPR.EnumReadyForPurchase.Yes)
                    Next
                Else
                    ObjStockInOut.Primary = Me.TxtPrimaryStockOutAdjust.Text
                    ObjStockInOut.CreatedBy = Me.UserCode
                    ObjStockInOut.TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
                    If ObjStockInOut.ApakahBolehHapus(Me.UserCode, True, True, True) = True Then
                        If MsgBox("Delete stock out adjust:" + Me.TxtStockOutAdjustNo.Text + " refering to this transaction?", _
                        MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            If Me.ObjStockInOut.Hapus(Me.TxtPrimaryStockOutAdjust.Text, Me.UserCode, ij.TblItemCustom, Me.TxtPrimary.Text) = True Then
                                Me.TxtPrimaryStockOutAdjust.Text = Me.ObjPR_PickUp.GetPrimaryStockOutAdjust(Me.TxtPrimary.Text)
                                Me.TxtStockOutAdjustNo.Text = Me.ObjStockInOut.GetNoTransaksi(Me.TxtPrimaryStockOutAdjust.Text)
                                Me.ObjPR_PickUp.TimeStamp = Me.ObjPR_PickUp.GetTimestamp(Me.TxtPrimary.Text)

                                'set tampilan
                                Me.SetStockOutAdjust()
                                For Each rowDGV In Me.DgvItem.Rows
                                    If ListPrimary.Contains(rowDGV.Cells(Me.ColItem_PrimaryPR.Index).Value) = False Then
                                        ListPrimary.Add(rowDGV.Cells(Me.ColItem_PrimaryPR.Index).Value)
                                    End If
                                Next
                                For Each PrimaryPR As Double In ListPrimary
                                    Me.ObjPR.SetOrderBalance1(PrimaryPR, ClsTransaksiPR.EnumReadyForPurchase.Yes)
                                Next

                                MsgBox("Data deleted.", MsgBoxStyle.Information, "Information")
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub llbShowStockOut_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbShowStockOut.LinkClicked
        Try
            Me.TxtPrimaryStockOutAdjust.Text = Me.ObjPR_PickUp.GetPrimaryStockOutAdjust(Me.TxtPrimary.Text)
            Me.TxtStockOutAdjustNo.Text = Me.ObjStockInOut.GetNoTransaksi(Me.TxtPrimaryStockOutAdjust.Text)
            Me.SetStockOutAdjust()

            If Me.TxtStockOutAdjustNo.Text <> "" Then
                If MsgBox("Show stock out data?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    Dim ij As Inventory_Module.FrmTransaksiStockOutAdj
                    Dim j As Windows.Forms.Form

                    For Each j In Me.ParentForm.MdiChildren
                        ij = TryCast(j, Inventory_Module.FrmTransaksiStockOutAdj)
                        If IsNothing(ij) = False Then
                            ij.BringToFront()
                            ij.Focus()
                            ij.ShowData(Me.TxtPrimaryStockOutAdjust.Text)
                            Exit Sub
                        End If
                    Next

                    Dim i As New Inventory_Module.FrmTransaksiStockOutAdj
                    ' Increment the caption counter.
                    i.UserCode = Me.UserCode
                    i.MdiParent = Me.ParentForm
                    i.vPrimaryLoadStockOut = Me.TxtPrimaryStockOutAdjust.Text

                    i.Show()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbSendEmail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSendEmail.LinkClicked
        Dim vEmail As String = ""

        Try
            If Me.TxtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimary.Text <> "-1" Then
                'cek apakah email dari reference party valid
                vEmail = Me.ObjRefParty.GetEmail(Me.TxtPrimaryRefParty.Text)
                If vEmail = "" Then
                    MsgBox("Please specify a valid email for the reference party.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If Me.ObjCD.ApakahEmailValid(vEmail) = False Then
                    MsgBox("Please specify a valid email for the reference party.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If Me.TxtPrimaryStockOutAdjust.Text <> "" And Me.TxtPrimaryStockOutAdjust.Text <> "-1" Then
                    If MsgBox("There stock out adjust referring to this document, this indicate the goods have been handed over to the ref party. Send email anyway?", _
                    MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Exit Sub
                    End If
                End If

                If MsgBox("Send this pick up order via email to the reference party?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.SendEmail()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SendEmail()
        Try
            Dim ij As New FrmTransaksiPR_PickUpOrder_SendEmail
            Dim ListPrimary As New List(Of Double)

            Try
                ListPrimary.Add(Me.TxtPrimary.Text)

                If ListPrimary.Count > 0 Then
                    ij.UserCode = Me.UserCode
                    ij.BackColorAll = Me.BAckColorAll
                    ij.ListPrimary_PickUpOrder = ListPrimary
                    ij.ShowDialog()
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Finally
                ij.Dispose()
                ListPrimary = Nothing
            End Try
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindPR_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindPR.Click
        Dim ObjRefParty As New ClsTransaksiPR
        Dim hj As New FrmTransaksiPR_Cari
        Dim ListNoPR As New List(Of String)
        Dim rowPR As DseTransaksiPR_PickUpOrder.TblPR1Row

        Try
            hj.BackColorAll = Me.BAckColorAll
            'If Me.TxtRequestBy.Text <> "" Then
            '    hj.txtCari.Text = Me.TxtRequestBy.Text
            'End If
            hj.txtCari.Text = Me.TxtRequestBy.Text
            hj.IsPickupOrder = True
            hj.btnFind.PerformClick()
            hj.ShowDialog()



            Dim ada As Boolean = False
            Dim index As Integer = 0
            ListNoPR.Clear()

            If hj.PrimaryPR > 0 Then
                If ObjRefParty.GetNoTransaksi(hj.PrimaryPR) <> "" Then
                    If Me.vTBLPR.DefaultView.Count > 0 Then
                        Dim dv2 As Data.DataRowView
                        For Each dv2 In vTBLPR.DefaultView
                            rowPR = dv2.Row
                            If rowPR.Primary_PR = hj.PrimaryPR Then
                                ada = True
                                Exit For
                            End If
                        Next
                    End If

                    If ada = True Then
                        Me.CboPRNo.Text = ObjRefParty.GetNoTransaksi(hj.PrimaryPR)
                    Else
                        MsgBox("The purchase request number " & ObjRefParty.GetNoTransaksi(hj.PrimaryPR) & _
                                " is not on the list, please choose another purchase request. ", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                Else
                    MsgBox("The purchase request number " & ObjRefParty.GetNoTransaksi(hj.PrimaryPR) & _
                            " is not on the list, please choose another purchase request. ", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If


            'If hj.PrimaryRefParty > 0 Then
            '    If ObjRefParty.GetNoTransaksi(hj.PrimaryRefParty) <> "" Then
            '        Me.CboPRNo.SelectedText = ObjRefParty.GetNoTransaksi(hj.PrimaryRefParty)
            '    Else
            '        MsgBox("The purchase request number " & ObjRefParty.GetNoTransaksi(hj.PrimaryRefParty) & _
            '                " is not on the list, please choose another purchase request. ", MsgBoxStyle.Information, "Information")
            '        Exit Sub
            '    End If
            'Else
            '    Me.TxtPrimaryRefParty.Text = "-1"
            '    Me.ResetCbo()
            '    Me.TxtRequestByName.Clear()
            '    Me.TxtRequestBy.Clear()
            '    Me.TxtRequestBy.Focus()
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
End Class