Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiSalesBalance_Import
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightSeaGreen

    Dim DefaultPrecision As Integer = 0
    Dim Buffer As String = ""

    Dim baru As Boolean = True
    Dim ObjPP As New ClsTransaksiSalesBalance_Import
    Dim ObjPPItem As New ClsTransaksiSalesBalance_Import_Item
    Dim ObjCD As New ClsCekData
    Dim ObjINV As New ClsTransaksiSalesInvoice
    Dim ObjJurnal As New ClsTransaksiJournal

    Private Sub FrmTransaksiSalesBalance_Import_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.ObjPP = New ClsTransaksiSalesBalance_Import
            Me.baru = True
            Me.KosongkanField()
            Me.TxtNoTransaksi.Clear()

            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.GroupBox2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.GroupBox3.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, False, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, False, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, False, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, False, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, False, True, True)
            ModCekData.KendalikanControl(Me.GroupBox2.Controls, True, True, True, True, True, False, True, True)
            ModCekData.KendalikanControl(Me.GroupBox3.Controls, True, True, True, True, True, False, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, True, True)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False

            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If

            Me.TxtNoTransaksi.ReadOnly = True
            Me.TxtNoTransaksi.BackColor = Me.txtReadOnly.BackColor

            ''ga standar
            Me.NonAktifkanTxtRecieved()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.GroupBox2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.GroupBox3.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, True, False, True, False)
            Me.CancelButton = Me.btnExit

            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True

            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.GroupBox2.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.GroupBox3.Controls, Me.BackColor, True)

            ''ga standar 
            Me.NonAktifkanTxtRecieved()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            Me.TxtPrimaryStockIn.Text = "-1"
            Me.TxtTotalSalesIDR.Text = "0"
            Me.TxtTotalDutyStamp.Text = "0"
            Me.TxtTotalVAT.Text = "0"
            Me.TxtGrandTotal.Text = "0"
            Me.TxtTotalBalance.Text = "0"

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvItem.Rows.Clear()
            Me.ObjCD.SetNumber(Me.dgvItem)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Try
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, False)
            Me.TxtVAT.Text = "0"
            Me.TxtDutystamp.Text = "0"
            Me.TxtSales.Text = "0"
            Me.TxtTotal.Text = "0"
            Me.TxtReceived.Text = "0"
            Me.TxtBalance.Text = "0"
            Me.DtpDueDate.Value = Now
            Me.dtpInvoiceDate.Value = Now
            ''Me.lblUnits.Text = "Units"
            ''Me.LblSecUnits.Text = "Units"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeObjek()
        Try
            With ObjPP
                .NoTransaksi = Me.TxtNoTransaksi.Text
                .TotalSales = Me.TxtTotalSalesIDR.Text
                .TotalDutystamp = Me.TxtTotalDutyStamp.Text
                .TotalVAT = Me.TxtTotalVAT.Text
                .GrandTotal = Me.TxtGrandTotal.Text
                .TotalRecieved = Me.TxtTotalBalance.Text
                .Remarks = Me.TxtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeField()
        Try
            With ObjPP
                Me.TxtPrimaryStockIn.Text = .Primary
                Me.TxtNoTransaksi.Text = .NoTransaksi
                Me.TxtTotalSalesIDR.Text = ObjCD.DigitGrouping(.TotalSales)
                Me.TxtTotalVAT.Text = ObjCD.DigitGrouping(.TotalVAT)
                Me.TxtTotalDutyStamp.Text = ObjCD.DigitGrouping(.TotalDutystamp)
                Me.TxtGrandTotal.Text = ObjCD.DigitGrouping(.GrandTotal)
                Me.TxtTotalBalance.Text = ObjCD.DigitGrouping(.TotalRecieved)
                Me.TxtRemarks.Text = .Remarks

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + _
                "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + _
                "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("No Transaksi")
            KoleksiTextBox.Add(Me.TxtNoTransaksi)
            ''KoleksiNamaData.Add("Schedule No")
            ''KoleksiTextBox.Add(Me.txtNoSchedule)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'Cek DGVItem
            If Me.dgvItem.Rows.Count <= 0 Then
                MsgBox("Please input item", MsgBoxStyle.Information, "Information")
                Me.TxtInvoiceNumber.Focus()
                Return False
            End If

            'cek perubahan nomor transaksi
            'If Me.baru = False Then
            'If Me.ObjPP.TanggalTransaksi.Month = Me.dtpTanggalTransaksi.Value.Month And _
            'Me.ObjPP.TanggalTransaksi.Year = Me.dtpTanggalTransaksi.Value.Year Then
            'Me.TxtNoTransaksi.Text = Me.ObjPP.NoTransaksi
            'Else
            'If MsgBox("You just changed transaction date, this will change your transaction " _
            '+ "number and there's a possibility your transaction no cannot be change back." _
            '+ "Continue and accept this change?", _
            'MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
            'Me.dtpTanggalTransaksi.Value = Me.ObjPP.TanggalTransaksi
            'Me.TxtNoTransaksi.Text = Me.ObjPP.NoTransaksi
            'Return False
            'End If
            'End If
            'End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.ObjPP.ApakahSudahAda(Me.TxtNoTransaksi.Text) = True Then
                    MsgBox("Duplication in no transaksi, please change your number", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtNoTransaksi.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.ObjPP.GetNotransaksi(Me.TxtPrimaryStockIn.Text)
                If TipeCodeLama <> Me.TxtNoTransaksi.Text Then
                    If Me.ObjPP.ApakahSudahAda(Me.TxtNoTransaksi.Text) = True Then
                        MsgBox("Duplication in no transaksi, please change your number", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtNoTransaksi.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjPP.GetTimestamp(Me.TxtPrimaryStockIn.Text)
                If Me.ObjPP.Timestamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data." _
                    + "Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjPP.CariDetail(Me.ObjPP.Primary, Me.dgvItem, Me)
                        Me.SalinkeField()
                    End If
                    Return False
                End If
            End If

            ''1 invoice hanya dapat di input 1x
            ''If ObjPP.CekInvoiceNumber(Me.dgvItem, Me.TxtPrimaryStockIn.Text, Me) <> "" Then
            ''MsgBox("The following invoice number have been recorded on the following Sales Balance Import(s):" _
            ''+ vbCrLf _
            ''+ ObjPP.CekInvoiceNumber(Me.dgvItem, Me.TxtPrimaryStockIn.Text, Me), MsgBoxStyle.Information, "Information")
            ''Return False
            ''End If

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
        Dim vClosingdate As Date = Now
        Dim ObjClosing As New ClsTransaksiClosing
        Try
            ''Tanggal < Closing Date
            vClosingdate = ObjClosing.GetClosingDate()
            If Me.dtpInvoiceDate.Value.Date > vClosingdate Then
                MsgBox("This invoice cannot input because invoice date > closing date", MsgBoxStyle.Information, "Information")
                Return False
            End If

            ''Balance > 0 
            If CDbl(Me.TxtBalance.Text) <= 0 Then
                MsgBox("Balance <= 0, please adjusted received's value", MsgBoxStyle.Information, "Information")
                Return False
            End If

            ''1 invoice hanya dapat di input 1x
            If Me.baru = True Then
                If ObjPP.CekJmlInvoiceNumber(Me.TxtPrimaryInvoice.Text) > 0 Then
                    MsgBox(ObjPP.CekInvoiceNumber_Item(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryInvoice.Text, Me.TxtInvoiceNumber.Text, Me), _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                If ObjPP.CekJmlInvoiceNumber_EDIT(Me.TxtPrimaryInvoice.Text, Me.TxtPrimaryStockIn.Text) > 0 Then
                    If ObjPP.CekJmlInvoiceNumber(Me.TxtPrimaryInvoice.Text) > 0 Then
                        MsgBox(ObjPP.CekInvoiceNumber_Item(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryInvoice.Text, Me.TxtInvoiceNumber.Text, Me), _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjClosing = Nothing
        End Try
    End Function

    Private Sub txtRemarks_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.Enter
        Try
            Me.Buffer = Me.TxtRemarks.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReceived_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceived.Enter
        Try
            Me.Buffer = Me.TxtReceived.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtReceived.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtReceived, sender, e)
            If IsNumeric(Me.TxtReceived.Text) = False Then
                If IsNumeric(Me.Buffer) = True Then
                    Me.TxtReceived.Text = Me.Buffer
                Else
                    Me.TxtReceived.Text = 0
                End If
            Else
                Me.TxtBalance.Text = ObjCD.DigitGrouping(ObjPP.HitungBalance(CDbl(Me.TxtSales.Text), CDbl(Me.TxtVAT.Text), _
                CDbl(Me.TxtDutystamp.Text), CDbl(Me.TxtReceived.Text)).ToString)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReceived_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceived.Leave
        Try
            If IsNumeric(Me.TxtReceived.Text) = False Then
                If IsNumeric(Me.Buffer) = True Then
                    Me.TxtReceived.Text = Me.Buffer
                Else
                    Me.TxtReceived.Text = 0
                End If
            Else
                Me.TxtBalance.Text = ObjCD.DigitGrouping(ObjPP.HitungBalance(CDbl(Me.TxtSales.Text), CDbl(Me.TxtVAT.Text), _
                CDbl(Me.TxtDutystamp.Text), CDbl(Me.TxtReceived.Text)).ToString)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtReceived.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtReceived, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.BackColorAll
                Else
                    TxtRemarks.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.TxtPrimaryStockIn.BackColor
                Else
                    TxtRemarks.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtScheduleNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInvoiceNumber.GotFocus
        Try
            If Me.TxtInvoiceNumber.Enabled = True Then
                If Me.TxtInvoiceNumber.ReadOnly = False Then
                    Me.TxtInvoiceNumber.BackColor = Me.BackColorAll
                Else
                    TxtInvoiceNumber.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtScheduleNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInvoiceNumber.LostFocus
        Try
            If Me.TxtInvoiceNumber.Enabled = True Then
                If Me.TxtInvoiceNumber.ReadOnly = False Then
                    Me.TxtInvoiceNumber.BackColor = Me.TxtPrimaryStockIn.BackColor
                Else
                    TxtInvoiceNumber.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLainnya_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceived.GotFocus
        Try
            If Me.TxtReceived.Enabled = True Then
                If Me.TxtReceived.ReadOnly = False Then
                    Me.TxtReceived.BackColor = Me.BackColorAll
                Else
                    TxtReceived.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLainnya_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceived.LostFocus
        Try
            If Me.TxtReceived.Enabled = True Then
                If Me.TxtReceived.ReadOnly = False Then
                    Me.TxtReceived.BackColor = Me.TxtPrimaryStockIn.BackColor
                Else
                    TxtReceived.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNoTransaksi_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoTransaksi.GotFocus
        Try
            If Me.TxtNoTransaksi.Enabled = True Then
                If Me.TxtNoTransaksi.ReadOnly = False Then
                    Me.TxtNoTransaksi.BackColor = Me.BackColorAll
                Else
                    TxtNoTransaksi.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNoTransaksi_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoTransaksi.LostFocus
        Try
            If Me.TxtNoTransaksi.Enabled = True Then
                If Me.TxtNoTransaksi.ReadOnly = False Then
                    Me.TxtNoTransaksi.BackColor = Me.TxtPrimaryStockIn.BackColor
                Else
                    TxtNoTransaksi.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Button dan event
    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmTransaksiSalesBalance_Import_Cari
        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            If Not ij.PrimaryPP = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjPP.CariDetail(ij.PrimaryPP, Me.dgvItem, Me) = True Then
                    Me.SalinkeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            'standar
            'Me.txtNoSchedule.ReadOnly = False
            'Me.txtNoSchedule.BackColor = Me.txtPrimarySchedule.BackColor
            'Me.btnFindPdoNo.Enabled = True
            'Me.txtNoSchedule.TabStop = True
            Me.ObjPP.ResetObject(Me.dgvItem)
            Me.TxtNoTransaksi.Text = Me.ObjPP.GetNoBaru()
            Me.TxtRemarks.Focus()

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtNoTransaksi.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtNoTransaksi.Text <> "" Then
                If Me.ObjPP.ApakahBolehEdit(Me.ObjPP.Primary, UserCode, True, True) = True Then
                    Me.baru = False
                    Me.ObjPP.CariDetail(Me.ObjPP.Primary, Me.dgvItem, Me)
                    Me.SalinkeField()
                    Me.AktifkanDetail()
                    'Me.IsiCBOProduct()
                    'ga standar
                    'If Me.dgvItem.RowCount > 0 Then
                    'Me.txtNoSchedule.ReadOnly = True
                    'Me.txtNoSchedule.BackColor = Me.txtReadOnly.BackColor
                    'Me.btnFindPdoNo.Enabled = False
                    'Me.txtNoSchedule.TabStop = False
                    'End If
                    ''Me.dgvItem.Focus()
                    Me.TxtRemarks.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtNoTransaksi.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtNoTransaksi.Text <> "" Then
                If Me.ObjPP.ApakahBolehHapus(Me.TxtPrimaryStockIn.Text, True, UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjPP.Hapus(Me.TxtPrimaryStockIn.Text, UserCode, True) = True Then
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

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            Me.KosongkanItem()
            Me.SalinkeObjek()
            If baru = True Then
                If ObjPP.ApakahDataBenar(baru) = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If Me.ApakahDataBenar() = True Then
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPP.SimpanBaru(Me.dgvItem, Me.UserCode, Me) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            End If
                        End If
                    End If
                End If
            Else
                'konfirmasi disimpan atau tidak
                If Me.ApakahDataBenar = True Then
                    If ObjPP.ApakahBolehEdit(Me.TxtPrimaryStockIn.Text, UserCode, True, True) = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPP.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            End If
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

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtRemarks.Text = "" And Me.dgvItem.RowCount = 0 And _
            Me.TxtNoTransaksi.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.KosongkanItem()
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjPP.CariDetail(Me.TxtPrimaryStockIn.Text, Me.dgvItem, Me)
                        Me.SalinkeField()
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

    Private Sub FrmTransaksiPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinkeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPP.SimpanBaru(Me.dgvItem, Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPP.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
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
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.ObjPP = New ClsTransaksiSalesBalance_Import
            Me.baru = True
            Me.KosongkanField()

            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Pengaturan DGV Item
    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_RowsAdded(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Try
            Me.dgvItem.TabStop = True
            If Me.TxtRemarks.Enabled = True Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_RowsRemoved(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Me.dgvItem.TabStop = False
                If Me.TxtRemarks.Enabled = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub dgvItemProcess_UserDeletedRow(ByVal sender As Object, ByVal e _
    As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Dim vTotalSales As Double = 0
        Dim vTotalVAT As Double = 0
        Dim vTotalDutystamp As Double = 0
        Dim vTotalBalance As Double = 0
        Try
            If Me.TxtRemarks.Enabled = True Then
                'If Me.dgvItem.RowCount = 0 Then
                'Me.txtNoSchedule.ReadOnly = False
                'Me.txtNoSchedule.BackColor = Me.txtPrimarySchedule.BackColor
                'Me.btnFindPdoNo.Enabled = True
                'Me.txtNoSchedule.TabStop = True
                'End If
                ObjPP.GetTotalHeader(Me.dgvItem, Me, vTotalSales, vTotalVAT, vTotalBalance, vTotalDutystamp)
                Me.TxtTotalSalesIDR.Text = ObjCD.DigitGrouping(vTotalSales.ToString)
                Me.TxtTotalVAT.Text = ObjCD.DigitGrouping(vTotalVAT.ToString)
                Me.TxtTotalBalance.Text = ObjCD.DigitGrouping(vTotalBalance.ToString)
                Me.TxtTotalDutyStamp.Text = ObjCD.DigitGrouping(vTotalDutystamp.ToString)
                Me.TxtGrandTotal.Text = ObjCD.DigitGrouping(ObjPP.GetGrandTotal(Me, Me.dgvItem).ToString)

            End If
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRowProcess(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvItem.Rows(row.Index).Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False) Then
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)

                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
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

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False

                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> Me.dgvItem.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    Me.ShowRowProcess(row, False)
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles dgvItem.SelectionChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.dgvItem.SelectedRows.Count > 0 Then
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

    Private Sub AktifkanTxtRecieved()
        Try
            Me.TxtReceived.Enabled = True
            Me.TxtReceived.TabStop = True
            Me.TxtReceived.ReadOnly = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanTxtRecieved()
        Try
            Me.TxtReceived.Enabled = False
            Me.TxtReceived.TabStop = False
            Me.TxtReceived.ReadOnly = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Find Code No
    Private Sub TxtNoJadwalProduksi_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtInvoiceNumber.KeyDown
        Dim objCode As New ClsTransaksiSalesInvoice
        Try
            If e.KeyValue = 13 Then
                If Me.TxtInvoiceNumber.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtInvoiceNumber.Text) = True Then
                        Me.TxtPrimaryInvoice.Text = objCode.GetPrimary(Me.TxtInvoiceNumber.Text)
                        Me.TxtInvoiceNumber.Text = objCode.GetInvoiceNumber(Me.TxtPrimaryInvoice.Text)
                        Me.TxtCustCode.Text = objCode.GetCustomerCode(Me.TxtPrimaryInvoice.Text)
                        Me.TxtCustName.Text = objCode.GetCustomerName(Me.TxtPrimaryInvoice.Text)
                        ''Sesuaikan Tanggal
                        Me.dtpInvoiceDate.Value = ObjINV.GetInvoiceDate(Me.TxtPrimaryInvoice.Text)
                        Me.DtpDueDate.Value = ObjPP.Get_DUE_DATE(Me.dtpInvoiceDate.Value.Date, Me.TxtPrimaryInvoice.Text)
                        ''Tarik data sales
                        Dim Sales As Double = ObjPP.GetSales(Me.TxtPrimaryInvoice.Text)
                        Dim Currency As Double = ObjPP.GetPrimaryCurrency(Me.TxtPrimaryInvoice.Text)
                        Me.TxtSales.Text = ObjCD.DigitGrouping(ObjPP.GetNilaiSalesInvoice(Sales, Currency, Me.TxtPrimaryInvoice.Text).ToString)
                        ''Tark data VAT, Dutystamp
                        Me.TxtVAT.Text = ObjCD.DigitGrouping(ObjPP.GetVAT(Me.TxtPrimaryInvoice.Text).ToString)
                        Me.TxtDutystamp.Text = ObjCD.DigitGrouping(ObjPP.GetDutystamp(Me.TxtPrimaryInvoice.Text).ToString)
                        Me.TxtTotal.Text = ObjCD.DigitGrouping(Math.Round(CDbl(Me.TxtSales.Text) + CDbl(Me.TxtVAT.Text) + _
                        CDbl(Me.TxtDutystamp.Text), 0, MidpointRounding.AwayFromZero).ToString)
                        ''hitung Balance
                        Me.TxtBalance.Text = ObjCD.DigitGrouping(ObjPP.HitungBalance(CDbl(Me.TxtSales.Text), CDbl(Me.TxtVAT.Text), _
                        CDbl(Me.TxtDutystamp.Text), CDbl(Me.TxtReceived.Text)).ToString)
                        Me.AktifkanTxtRecieved()
                        Me.TxtReceived.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.BtnFindINV.PerformClick()
                    End If
                Else
                    Me.BtnFindINV.PerformClick()
                End If
            End If
            ''Me.IsiCBOProduct()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
        End Try
    End Sub

    Private Sub TxtNoJadwalProduksi_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtInvoiceNumber.Leave
        Dim i As New FrmCariTransaksiSalesInvoice
        ''i.vTipeTransaksi = ClsTransaksiJadwal.EnumTipe.Schedule
        Dim objCode As New ClsTransaksiSalesInvoice
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtInvoiceNumber.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtInvoiceNumber.Text) = True Then
                    Me.TxtPrimaryInvoice.Text = objCode.GetPrimary(Me.TxtInvoiceNumber.Text)
                    Me.TxtInvoiceNumber.Text = objCode.GetInvoiceNumber(Me.TxtPrimaryInvoice.Text)
                    Me.TxtCustCode.Text = objCode.GetCustomerCode(Me.TxtPrimaryInvoice.Text)
                    Me.TxtCustName.Text = objCode.GetCustomerName(Me.TxtPrimaryInvoice.Text)
                    ''Sesuaikan Tanggal
                    Me.dtpInvoiceDate.Value = ObjINV.GetInvoiceDate(Me.TxtPrimaryInvoice.Text)
                    Me.DtpDueDate.Value = ObjPP.Get_DUE_DATE(Me.dtpInvoiceDate.Value.Date, Me.TxtPrimaryInvoice.Text)
                    ''Tarik data sales
                    Dim Sales As Double = ObjPP.GetSales(Me.TxtPrimaryInvoice.Text)
                    Dim Currency As Double = ObjPP.GetPrimaryCurrency(Me.TxtPrimaryInvoice.Text)
                    Me.TxtSales.Text = ObjCD.DigitGrouping(ObjPP.GetNilaiSalesInvoice(Sales, Currency, Me.TxtPrimaryInvoice.Text).ToString)
                    ''Tark data VAT, Dutystamp
                    Me.TxtVAT.Text = ObjCD.DigitGrouping(ObjPP.GetVAT(Me.TxtPrimaryInvoice.Text).ToString)
                    Me.TxtDutystamp.Text = ObjCD.DigitGrouping(ObjPP.GetDutystamp(Me.TxtPrimaryInvoice.Text).ToString)
                    Me.TxtTotal.Text = ObjCD.DigitGrouping(Math.Round(CDbl(Me.TxtSales.Text) + CDbl(Me.TxtVAT.Text) + _
                    CDbl(Me.TxtDutystamp.Text), 0, MidpointRounding.AwayFromZero).ToString)
                    ''hitung Balance
                    Me.TxtBalance.Text = ObjCD.DigitGrouping(ObjPP.HitungBalance(CDbl(Me.TxtSales.Text), CDbl(Me.TxtVAT.Text), _
                    CDbl(Me.TxtDutystamp.Text), CDbl(Me.TxtReceived.Text)).ToString)
                    Me.AktifkanTxtRecieved()
                Else
                    If Me.TxtInvoiceNumber.Text <> "" Then
                        i.txtCari.Text = Me.TxtInvoiceNumber.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksi > 0 Then
                        Me.TxtPrimaryInvoice.Text = i.PrimaryTransaksi
                        Me.TxtInvoiceNumber.Text = objCode.GetInvoiceNumber(Me.TxtPrimaryInvoice.Text)
                        Me.TxtCustCode.Text = objCode.GetCustomerCode(Me.TxtPrimaryInvoice.Text)
                        Me.TxtCustName.Text = objCode.GetCustomerName(Me.TxtPrimaryInvoice.Text)
                        ''Sesuaikan Tanggal
                        Me.dtpInvoiceDate.Value = ObjINV.GetInvoiceDate(Me.TxtPrimaryInvoice.Text)
                        Me.DtpDueDate.Value = ObjPP.Get_DUE_DATE(Me.dtpInvoiceDate.Value.Date, Me.TxtPrimaryInvoice.Text)
                        ''Tarik data sales
                        Dim Sales As Double = ObjPP.GetSales(Me.TxtPrimaryInvoice.Text)
                        Dim Currency As Double = ObjPP.GetPrimaryCurrency(Me.TxtPrimaryInvoice.Text)
                        Me.TxtSales.Text = ObjCD.DigitGrouping(ObjPP.GetNilaiSalesInvoice(Sales, Currency, Me.TxtPrimaryInvoice.Text).ToString)
                        ''Tark data VAT, Dutystamp
                        Me.TxtVAT.Text = ObjCD.DigitGrouping(ObjPP.GetVAT(Me.TxtPrimaryInvoice.Text).ToString)
                        Me.TxtDutystamp.Text = ObjCD.DigitGrouping(ObjPP.GetDutystamp(Me.TxtPrimaryInvoice.Text).ToString)
                        Me.TxtTotal.Text = ObjCD.DigitGrouping(Math.Round(CDbl(Me.TxtSales.Text) + CDbl(Me.TxtVAT.Text) + _
                        CDbl(Me.TxtDutystamp.Text), 0, MidpointRounding.AwayFromZero).ToString)
                        ''hitung Balance
                        Me.TxtBalance.Text = ObjCD.DigitGrouping(ObjPP.HitungBalance(CDbl(Me.TxtSales.Text), CDbl(Me.TxtVAT.Text), _
                        CDbl(Me.TxtDutystamp.Text), CDbl(Me.TxtReceived.Text)).ToString)
                        Me.AktifkanTxtRecieved()
                    Else
                        Me.TxtInvoiceNumber.Clear()
                        Me.TxtPrimaryInvoice.Text = "-1"
                        Me.TxtInvoiceNumber.Focus()
                        Me.TxtCustName.Clear()
                        Me.TxtCustCode.Clear()
                        Me.TxtSales.Text = "0"
                        Me.TxtVAT.Text = "0"
                        Me.TxtDutystamp.Text = "0"
                        Me.dtpInvoiceDate.Value = Now
                        Me.DtpDueDate.Value = Now
                        Me.TxtTotal.Text = "0"
                        Me.TxtBalance.Text = "0"
                        Me.TxtReceived.Text = "0"
                        Me.NonAktifkanTxtRecieved()
                        ''Me.KosongkanItem()
                    End If
                End If
            Else
                Me.TxtInvoiceNumber.Clear()
                Me.TxtPrimaryInvoice.Text = "-1"
                Me.TxtCustCode.Clear()
                Me.TxtCustName.Clear()
                Me.TxtSales.Text = "0"
                Me.TxtVAT.Text = "0"
                Me.TxtDutystamp.Text = "0"
                Me.dtpInvoiceDate.Value = Now
                Me.DtpDueDate.Value = Now
                Me.TxtTotal.Text = "0"
                Me.TxtBalance.Text = "0"
                Me.TxtReceived.Text = "0"
                Me.NonAktifkanTxtRecieved()
                ''Me.btnInput.Focus()
                ''Me.KosongkanItem()
            End If
            '' Me.IsiCBOProduct()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindSchedule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles BtnFindINV.Click
        Dim i As New FrmCariTransaksiSalesInvoice
        ''i.vTipeTransaksi = ClsTransaksiJadwal.EnumTipe.Schedule
        Dim objCode As New ClsTransaksiSalesInvoice
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtInvoiceNumber.Text <> "" Then
                i.txtCari.Text = Me.TxtInvoiceNumber.Text
            End If

            i.ShowDialog()
            If i.PrimaryTransaksi > 0 Then
                Me.TxtPrimaryInvoice.Text = i.PrimaryTransaksi
                Me.TxtInvoiceNumber.Text = objCode.GetInvoiceNumber(Me.TxtPrimaryInvoice.Text)
                Me.TxtCustCode.Text = objCode.GetCustomerCode(Me.TxtPrimaryInvoice.Text)
                Me.TxtCustName.Text = objCode.GetCustomerName(Me.TxtPrimaryInvoice.Text)
                ''Sesuaikan Tanggal
                Me.dtpInvoiceDate.Value = ObjINV.GetInvoiceDate(Me.TxtPrimaryInvoice.Text)
                Me.DtpDueDate.Value = ObjPP.Get_DUE_DATE(Me.dtpInvoiceDate.Value.Date, Me.TxtPrimaryInvoice.Text)
                ''Tarik data sales
                Dim Sales As Double = ObjPP.GetSales(Me.TxtPrimaryInvoice.Text)
                Dim Currency As Double = ObjPP.GetPrimaryCurrency(Me.TxtPrimaryInvoice.Text)
                Me.TxtSales.Text = ObjCD.DigitGrouping(ObjPP.GetNilaiSalesInvoice(Sales, Currency, Me.TxtPrimaryInvoice.Text).ToString)
                ''Tark data VAT, Dutystamp
                Me.TxtVAT.Text = ObjCD.DigitGrouping(ObjPP.GetVAT(Me.TxtPrimaryInvoice.Text).ToString)
                Me.TxtDutystamp.Text = ObjCD.DigitGrouping(ObjPP.GetDutystamp(Me.TxtPrimaryInvoice.Text).ToString)
                Me.TxtTotal.Text = ObjCD.DigitGrouping(Math.Round(CDbl(Me.TxtSales.Text) + CDbl(Me.TxtVAT.Text) + _
                CDbl(Me.TxtDutystamp.Text), 0, MidpointRounding.AwayFromZero).ToString)
                ''hitung Balance
                Me.TxtBalance.Text = ObjCD.DigitGrouping(ObjPP.HitungBalance(CDbl(Me.TxtSales.Text), CDbl(Me.TxtVAT.Text), _
                CDbl(Me.TxtDutystamp.Text), CDbl(Me.TxtReceived.Text)).ToString)
                Me.AktifkanTxtRecieved()
                Me.TxtReceived.Focus()
            Else
                Me.TxtInvoiceNumber.Clear()
                Me.TxtPrimaryInvoice.Text = "-1"
                Me.TxtInvoiceNumber.Focus()
                Me.TxtCustCode.Clear()
                Me.TxtCustName.Clear()
                Me.TxtSales.Text = "0"
                Me.TxtVAT.Text = "0"
                Me.TxtDutystamp.Text = "0"
                Me.dtpInvoiceDate.Value = Now
                Me.DtpDueDate.Value = Now
                Me.TxtTotal.Text = "0"
                Me.TxtBalance.Text = "0"
                Me.TxtReceived.Text = "0"
                Me.NonAktifkanTxtRecieved()
                ''Me.KosongkanItem()
            End If
            ''Me.IsiCBOProduct()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInput.Click
        Dim Index1 As Integer = -1
        Dim row As DataGridViewRow
        Dim ObjCust As New clsMasterCustomer
        Dim vTotalSales As Double = 0
        Dim vTotalVAT As Double = 0
        Dim vTotalDutystamp As Double = 0
        Dim vTotalBalance As Double = 0
        Try
            If Me.ApakahDataItemBenar = True Then
                Index1 = Me.GetIndex(Me.TxtPrimaryInvoice.Text, Me.TxtInvoiceNumber.Text)

                'cari data yang sama persis
                If Index1 <> -1 Then
                    'tambahkan nilai qty
                    row = Me.dgvItem.Rows(Index1)
                    MsgBox("Invoice No" + " " + Me.TxtInvoiceNumber.Text + " " + "has been recorded at line" + " " + _
                    CStr(row.Index + 1), MsgBoxStyle.Information, "Information")
                Else
                    'tambahkan baris baru
                    Me.dgvItem.Rows.Add()
                    row = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                    row.Cells(Me.ColInv_Number.Index).Value = Me.TxtInvoiceNumber.Text
                    row.Cells(Me.ColInv_Date.Index).Value = Format(Me.dtpInvoiceDate.Value, "dd MMM yyyy")
                    row.Cells(Me.Col_Inv_DueDate.Index).Value = Format(Me.DtpDueDate.Value, "dd MMM yyyy")
                    row.Cells(Me.ColCustCode.Index).Value = Me.TxtCustCode.Text
                    row.Cells(Me.ColCustName.Index).Value = Me.TxtCustName.Text
                    row.Cells(Me.ColSales.Index).Value = ObjCD.DigitGrouping(CDbl(Me.TxtSales.Text))
                    row.Cells(Me.ColVAT.Index).Value = ObjCD.DigitGrouping(CDbl(Me.TxtVAT.Text))
                    row.Cells(Me.ColDutystamp.Index).Value = ObjCD.DigitGrouping(CDbl(Me.TxtDutystamp.Text))
                    row.Cells(Me.ColTotal.Index).Value = ObjCD.DigitGrouping(CDbl(Me.TxtTotal.Text))
                    row.Cells(Me.ColReceived.Index).Value = ObjCD.DigitGrouping(CDbl(Me.TxtReceived.Text))
                    row.Cells(Me.ColBalance.Index).Value = ObjCD.DigitGrouping(CDbl(Me.TxtBalance.Text))
                    row.Cells(Me.ColRemarks.Index).Value = ""
                    row.Cells(Me.ColPrimaryInv.Index).Value = Me.TxtPrimaryInvoice.Text
                    row.Cells(Me.ColPrimaryCust.Index).Value = ObjCust.GetPrimary(Me.TxtCustCode.Text)
                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItem)
                End If
                'kosongkan gbo item
                Me.KosongkanItem()

                ''Hitung Total
                ObjPP.GetTotalHeader(Me.dgvItem, Me, vTotalSales, vTotalVAT, vTotalBalance, vTotalDutystamp)
                Me.TxtTotalSalesIDR.Text = ObjCD.DigitGrouping(vTotalSales.ToString)
                Me.TxtTotalVAT.Text = ObjCD.DigitGrouping(vTotalVAT.ToString)
                Me.TxtTotalBalance.Text = ObjCD.DigitGrouping(vTotalBalance.ToString)
                Me.TxtTotalDutyStamp.Text = ObjCD.DigitGrouping(vTotalDutystamp.ToString)
                Me.TxtGrandTotal.Text = ObjCD.DigitGrouping(ObjPP.GetGrandTotal(Me, Me.dgvItem).ToString)

                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                ''If Me.dgvItem.RowCount > 0 Then
                ''Me.txtNoSchedule.ReadOnly = True
                ''Me.txtNoSchedule.BackColor = Me.txtReadOnly.BackColor
                ''Me.btnFindPdoNo.Enabled = False
                ''Me.txtNoSchedule.TabStop = False
                ''End If
                Me.TxtInvoiceNumber.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCust = Nothing
        End Try
    End Sub

    Private Function GetIndex(ByVal vPrimaryInvoice As Double, ByVal vInvoiceNumber As String) As Integer
        Dim row As DataGridViewRow
        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryInv.Index).Value = CStr(vPrimaryInvoice) And _
                row.Cells(Me.ColInv_Number.Index).Value = vInvoiceNumber Then
                    Return row.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItem_CellBeginEdit(ByVal sender As System.Object, ByVal e _
    As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Dim Jurnal As String = ""

        Try
            If e.ColumnIndex = Me.ColReceived.Index Then
                Jurnal = Me.ObjJurnal.Get_SalesBalanceTransfer_Jurnal_Item(Me.dgvItem.Item(Me.ColPrimaryInv.Index, e.RowIndex).Value)
                If Jurnal <> "" Then
                    MsgBox("This invoice referred by sales receivement journal: " + Jurnal + ".", _
                    MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                Else
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            ElseIf e.ColumnIndex = Me.ColRemarks.Index Then
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Index As Integer = -1
        Dim vTotalSales As Double = 0
        Dim vTotalVAT As Double = 0
        Dim vTotalDutystamp As Double = 0
        Dim vTotalBalance As Double = 0
        Dim vHitungBalance As Double = 0
        Try
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex = Me.ColReceived.Index Then
                        If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            If IsNumeric(Me.Buffer) = True Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Else
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = 0
                            End If
                        End If
                    End If

                    If e.ColumnIndex <> Me.ColRemarks.Index Then
                        'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit
                        If e.ColumnIndex = Me.ColReceived.Index Then
                            If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, _
                            e.RowIndex).Value) = False Then
                                MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                                "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If

                            ''hitung Balance
                            row.Cells(Me.ColBalance.Index).Value = ObjCD.DigitGrouping(ObjPP.HitungBalance(CDbl(row.Cells(Me.ColSales.Index).Value), _
                            CDbl(row.Cells(Me.ColVAT.Index).Value), CDbl(row.Cells(Me.ColDutystamp.Index).Value), CDbl(row.Cells(Me.ColReceived.Index).Value)).ToString)
                            If row.Cells(Me.ColBalance.Index).Value <= 0 Then
                                MsgBox("Balance <= 0, please adjusted received's value", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer)
                                ''Hitung lagi balancenya 
                                row.Cells(Me.ColBalance.Index).Value = ObjCD.DigitGrouping(ObjPP.HitungBalance(CDbl(row.Cells(Me.ColSales.Index).Value), _
                                CDbl(row.Cells(Me.ColVAT.Index).Value), CDbl(row.Cells(Me.ColDutystamp.Index).Value), CDbl(row.Cells(Me.ColReceived.Index).Value)).ToString)
                            End If
                            row.Cells(Me.ColReceived.Index).Value = ObjCD.DigitGrouping(row.Cells(Me.ColReceived.Index).Value)

                            ''Hitung Total
                            ObjPP.GetTotalHeader(Me.dgvItem, Me, vTotalSales, vTotalVAT, vTotalBalance, vTotalDutystamp)
                            Me.TxtTotalSalesIDR.Text = ObjCD.DigitGrouping(vTotalSales.ToString)
                            Me.TxtTotalVAT.Text = ObjCD.DigitGrouping(vTotalVAT.ToString)
                            Me.TxtTotalBalance.Text = ObjCD.DigitGrouping(vTotalBalance.ToString)
                            Me.TxtTotalDutyStamp.Text = ObjCD.DigitGrouping(vTotalDutystamp.ToString)
                            Me.TxtGrandTotal.Text = ObjCD.DigitGrouping(ObjPP.GetGrandTotal(Me, Me.dgvItem).ToString)
                        End If
                    End If
                    ''If e.ColumnIndex = Me.ColTanggalMulai.Index Then
                    'cek apakah formatnya tanggal
                    ''If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) <> True Then
                    'kalo format bukan tanggal maka balik ke nilai semula
                    ''Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    ''Else
                    'kalo tanggal, maka sesuaikan formatnya
                    ''Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format _
                    ''(CDate(Me.dgvItem.Item(e.ColumnIndex, _
                    ''e.RowIndex).Value), "dd MMM yyyy")
                    ''End If
                    ''End If
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
End Class