Public Class FrmTrans_Data_Transfer1
    Dim PrimaryRujukanTransfer As Double    'variabel ini berfungsi sebagai primary rujukan yang akan digunakan
    Dim ObjCabang As New ClsMast_Cabang
    Dim TBLCabang As New DseMast_Cabang.tblmaster_cabang2DataTable
    Dim rowCabang As DseMast_Cabang.tblmaster_cabang2Row
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim DefaultCur As Decimal = 1
    Dim DefaultCustomer As Decimal = 1
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim ObjSession As New ClsMast_Session

    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow

    'melanjutkan proses transfer setelah proses tersebut dipause

    Public Enum EnumTranferProgress As Integer
        Update_Transfer_Start   ' set status transfer pada database cabang
        SQLScriptGeneration 'mulai dari bikin sql script
        SQLScriptInsertion 'masukkan data ke dalam tabel sementara
        SessionHeaderTransfer 'memasukkan data session header
        Session_Item_PayOutDropCash 'memasukkan data pay out drop cash, lengkap dengan jurnal2-nya
        Session_item_Void
        POS
        Update_Finish 'set status transfer pada database cabang sudah selesai
        Finish
    End Enum

    Private Sub SimpanInvoice(ByVal vPrimarySession As Double)
        Dim IJ As New Sales_Module.FrmTransaksiSalesInvoiceQuick
        Dim TBLItem As New dseTrans_Data_Transfer.tbltrans_pos_cabang_itemDataTable
        Dim DTAItem As New dseTrans_Data_TransferTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Dim FrmDO As New Sales_Module.FrmTransaksiSalesInvoiceQuick_DetailNumber

        Try
            'kumpulkan semua transaksi pos session tersebut
            DTAItem.Fill(TBLItem, vPrimarySession)

            'pindahkan ke dalam 1 invoice
            'set form invoice
            IJ.Show()
            IJ.WindowState = FormWindowState.Normal
            IJ.MinimumSize = New System.Drawing.Size(100, 100)
            IJ.Size = New System.Drawing.Size(100, 100)
            IJ.BringToFront()
            IJ.Location = New Point(-100, -100)
            IJ.UserCode = "USER"
            IJ.CheckLicense = False
            IJ.btnNew.PerformClick()

            'set data header
            IJ.chkTax.Checked = False
            IJ.txtPrimaryCustomer.Text = Me.DefaultCustomer
            IJ.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(IJ.txtPrimaryCustomer.Text)
            IJ.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(IJ.txtPrimaryCustomer.Text)
            IJ.dtpDate.Value = Me.ObjSession.GetEndDateSession(vPrimarySession)
            Me.ObjCustomer.IsiCboAddress(IJ.cboCustomerAddress, IJ.txtPrimaryCustomer.Text)
            IJ.cboCustomerAddress.SelectedIndex = 0
            IJ.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(IJ.txtPrimaryCustomer.Text)
            IJ.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(IJ.txtPrimaryCustomer.Text)
            'set data item
            Me.MasukkanDataItemInvoice(IJ, TBLItem)

            'simpan baru
            If IJ.Simpan(FrmDO) = True Then
                MsgBox("Data saved")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            IJ.Dispose()
            FrmDO.Dispose()
        End Try
    End Sub

    Private Sub MasukkanDataItemInvoice(ByVal ij As Sales_Module.FrmTransaksiSalesInvoiceQuick, _
    ByVal TBLItem As dseTrans_Data_Transfer.tbltrans_pos_cabang_itemDataTable)
        Dim indexSama As Integer = -1
        Dim row As DataGridViewRow
        Dim AngkaDecimal As Double = 0
        Dim AngkaDecimalCurrency As Double = 0
        Dim PrimaryCurrenct As Double = -1
        Dim rowItem As dseTrans_Data_Transfer.tbltrans_pos_cabang_itemRow

        Try
            For Each rowItem In TBLItem.Rows
                indexSama = ij.GetIndex(rowItem.nett_price, rowItem.batch_number, 0, rowItem.primary_location, rowItem.primary_product)

                If indexSama = -1 Then
                    'tambahkan baris baru
                    ij.dgvItem.Rows.Add()
                    row = ij.dgvItem.Rows(ij.dgvItem.Rows.Count - 1)
                    AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(rowItem.primary_product)
                    AngkaDecimalCurrency = Me.ObjCurrency.GetDefaultPrecission(Me.DefaultCur)

                    With row
                        .Cells(Me.ColNo0.Index).Value = ij.dgvItem.Rows.Count
                        .Cells(Me.ColCodeNo.Index).Value = ""
                        .Cells(Me.ColDescription.Index).Value = ""
                        .Cells(Me.ColCustCodeNo.Index).Value = ""
                        .Cells(Me.ColCustDesc.Index).Value = ""
                        .Cells(Me.ColSONO.Index).Value = ""
                        .Cells(Me.ColCustPONo.Index).Value = ""
                        .Cells(Me.ColDONO.Index).Value = ""
                        .Cells(Me.ColStockOutNo.Index).Value = ""
                        .Cells(Me.ColLotNumber.Index).Value = rowItem.batch_number
                        .Cells(Me.ColExpiry.Index).Value = #1/1/1950#
                        .Cells(Me.ColLocationCode.Index).Value = ""
                        .Cells(Me.ColQtyPerMaster.Index).Value = 0
                        .Cells(Me.ColQtyMaster.Index).Value = 0
                        .Cells(Me.ColMasterUnits.Index).Value = ""
                        .Cells(Me.ColQty.Index).Value = 0
                        .Cells(Me.ColTotalQty.Index).Value = rowItem.qty
                        .Cells(Me.ColUnit.Index).Value = ""
                        .Cells(Me.ColSecQty.Index).Value = -1
                        .Cells(Me.ColSecUnit.Index).Value = ""
                        .Cells(Me.ColCurrency.Index).Value = ""
                        .Cells(Me.ColUnitPrice.Index).Value = rowItem.nett_price
                        .Cells(Me.ColDiscount.Index).Value = 0
                        .Cells(Me.ColNettPrice.Index).Value = rowItem.nett_price
                        .Cells(Me.ColAmount.Index).Value = rowItem.sub_total
                        .Cells(Me.ColRemarksItem.Index).Value = ""
                        .Cells(Me.ColPrimaryLocation.Index).Value = rowItem.primary_location
                        .Cells(Me.ColPrimaryDO.Index).Value = -1
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = rowItem.primary_product
                        .Cells(Me.ColPrimarySO.Index).Value = -1
                        .Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                        .Cells(Me.ColPrimaryUnit.Index).Value = rowItem.primary_unit
                        .Cells(Me.ColPrimarySecUnit.Index).Value = -1
                        .Cells(Me.ColPrimaryCurrency.Index).Value = Me.DefaultCur
                    End With
                Else
                    row = ij.dgvItem.Rows(indexSama)
                    'tambahkan qty, 
                    row.Cells(Me.ColTotalQty.Index).Value = CDec(row.Cells(Me.ColTotalQty.Index).Value) + CDec(rowItem.qty)
                    'hitung ulang amount
                    row = ij.dgvItem.Rows(indexSama)
                    ij.Hitung_Amount(row)
                End If
            Next

            ij.Hitung_Total()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateStatusTansfer(ByVal vBranchConstring As String, ByVal vPrimarySession As Double, ByVal vStatus As EnumTranferProgress)
        Dim DTA As New dseTrans_Data_TransferTableAdapters.tbltrans_session_cabangTableAdapter

        Try
            DTA.Connection.ConnectionString = vBranchConstring
            DTA.UpdateQuery_SetStatusTransfer(vStatus, vPrimarySession)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub TransferSession(ByVal vBranchConstring As String, ByVal vServerConstring As String, ByVal vPrimarySession As Double, ByVal vPrimaryBranch As Double)
        Try
            'perhatikan bahwa proses ini bisa di pause dan kemudian di lanjutkan kembali, setelah button start ditekan, 
            'dia akan berubah menjadi button pause. ketika pause, harus update statusnya ke dalam tabel di database, begitu juga ketika lanjut lagi
            'data transfer selalu dimulai dari session yang lebih dulu dilakukan --> tanggalnya paling kecil
            MsgBox("Fitur pause proses belum dicoding")

            '1. tarik semua data jumlah untuk session dan transaksi pos yang akan ditransfer, untuk digunakan melacak progress kerja
            'jumlah terdiri atas:

            '2. untuk setiap data session,
            '   a. tandai bawah session ybs, sedang dalam proses transfer(mulai)
            Me.UpdateStatusTansfer(vBranchConstring, vPrimarySession, EnumTranferProgress.Update_Transfer_Start)
            '   b. copy semua data session ybs ke dalam script my sql (tbltrans_session_cabang, tbltrans_session_cabang_itemvoid, 
            '      tbltrans_session_cabang_itempayoutdropcash, tbltrans_pos_cabang, tbltrans_pos_cabang_item, tbltrans_pos_cabang_item_payment
            Me.PindahkanKeSessionTransfer(vPrimaryBranch, vPrimarySession, vBranchConstring, vServerConstring)
            Me.UpdateStatusTansfer(vBranchConstring, vPrimarySession, EnumTranferProgress.SQLScriptGeneration)
            '   c. jalankan script tersebut dan masukkan ke dalam tabel sementara

            Me.UpdateStatusTansfer(vBranchConstring, vPrimarySession, EnumTranferProgress.SQLScriptInsertion)
            '   d. tandai session ybs, start proses transfer-on progress

            '   e. cek validitas stock, jangan sampai transfer mengakibatkan data minus
            '   d. pindahkan data header session, tandai tbltrans_session_cabang, pada kolom primary trans session server, 
            '      buat jurnal open session session dan jurnal close session
            'pindahkan data header
            Me.HeaderTransfer(vPrimaryBranch, vPrimarySession, vBranchConstring, vServerConstring)
            Me.UpdateStatusTansfer(vBranchConstring, vPrimarySession, EnumTranferProgress.SessionHeaderTransfer)
            '   e. pindahkan data item pay out drop cash session, buat juga jurnal yang bersesuai untuk setiap tipe pay out drop cash (bukan untuk setiap item)
            '   f. pindahkan data transaksi pos, satu per satu, lengkap dengan transferan invoice dan data pembayarannya, untuk jurnal-nya 1 pos 
            '      minimal buat 2 jurnal yaitu, jurnal sales, pada saat invoice, dan jurnal sales receivement pada saat pembayaran, semua transaksi dianggap
            '      sudah dibayarkan. Jika tidak bisa simpan ditengah jalan karena mengakibatkan stock minus, maka dilewat dulu. Untuk setiap transaksi pos, yang 
            '      dilakukan adalah:
            '           i.   buat quick invoice untuk transaksi
            '           ii.  buat sales payment jurnal untuk invoice ybs
            '           iii. buat jurnal cash withdrawal
            '           iv.  buat jurnal payment disbursment untuk yang delay-nya = 0 dan untuk yang sudah lewat masa pembayarannya

            'masalah yang kemungkinan timbul ketika melakukan data transfer:
            '1. Ada stock minus, Ada stock yang harusnya location, batch number, atau expiry control tetapi tidak dimasukkan datanya ketika transaksi kasir, 
            '   (tidak location control, batch atau expiry, kalo ada salah satu lihat ke bawah): 
            '   tergantung setting yg digunakan oleh user apakah mau melakukan penyesuaian sendiri atau dilakukan stock in adjust secara otomatis
            '2. Ada code no yang memiliki lokasi, ada beberapa alternatif:
            '   a. setiap item yang memiliki barcode tidak boleh memiliki batch number control, expiry control, dan location control (bisa pilih misalnya,
            '      boleh location control, tapi ga boleh batch number control atau kombinasi lainnya.)
            '   b. user melakukan input ketika di dalam transaksi pos, kemudian dicek kembali ketika transfer
            '   c. user menentukan sendiri untuk setiap item pos yang ada location, expiry, dan batch number yang digunakan
            '   d. software akan memilih sendiri item mana yang akan digunakan untuk transfer dengan ketentuan:
            '       - location      : setiap cabang memiliki location yang terhubung dengan cabang tersebut, bisa beberapa lokasi dengan urutan prioritas yang 
            '                         jelas.
            '       - expiry        : setelah location, jika expiry control, maka pilih item yang ada, 
            '                         berdasarkan expiry yang paling dekat, jadi yang keluar adalah item dengan expiry yang paling dekat
            '       - batch number  : setelah expiry, jika bathc number kontrol, maka dipilih bacth number secara ASC/DESC (tergantung setting)

            'berikut ini cara penomoran transaksi yang digenerate secara otomatis oleh software
            '1. invoice --> ikutin ketentuan penomoran standar
            '2. sales jurnal --> ikutin ketentuan sales jurnal
            '3. sales payment jurnal --> 
            '4. open session jurnal -->
            '5. close session jurnal -->
            '6. payout drop cash jurnal --> 

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HeaderTransfer(ByVal vPrimaryBranch As Double, ByVal vPrimarySession As Double, ByVal vConstringCabang As String, _
    ByVal vConnStringUtama As String)
        Dim DTAHQ As New dseTrans_Data_TransferTableAdapters.tbltrans_sessionTableAdapter
        Dim DTABranch As New dseTrans_Data_TransferTableAdapters.tbltrans_session_cabang1TableAdapter
        Dim TBLBranch As New dseTrans_Data_Transfer.tbltrans_session_cabang1DataTable
        Dim rowBranch As dseTrans_Data_Transfer.tbltrans_session_cabang1Row

        Try
            'hapus data session di cabang utama
            DTAHQ.Connection.ConnectionString = vConnStringUtama
            DTAHQ.DeleteQuery_ByPrimaryCabang(vPrimarySession, vPrimaryBranch)

            'masukkan data header session di cabang utama
            DTABranch.Connection.ConnectionString = vConstringCabang
            DTABranch.FillBy_Primary(TBLBranch, vPrimarySession)
            If TBLBranch.Rows.Count > 0 Then
                rowBranch = TBLBranch.Rows(0)
                With rowBranch
                    DTAHQ.Insert(.session_id, .primary, .status, .start_date, .end_date, .Cashier_Code, _
                    .primary_branch, .remarks, .beginng_cash, .cash_sales, .drop_cash, .pay_out, .balance, _
                    .primary_begining_cash_account, .primary_cashier_cash_account, .primary_balance_cash_account, _
                    .primary_jurnal_open_session, .primary_jurnal_close_session, .created_by, .created_date, .modified_by, .modified_date, _
                    .timestamp, -1)

                    'update primary session server di cabang
                    If DTAHQ.ScalarQuery_GetPrimary(.primary, .primary_branch).HasValue = True Then
                        DTABranch.UpdateQuery_SetPrimaryTransaksiServer(DTAHQ.ScalarQuery_GetPrimary(.primary, .primary_branch), .primary)
                    End If
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAHQ.Dispose()
            TBLBranch.Dispose()
            DTABranch.Dispose()
        End Try
    End Sub

    Private Sub PindahkanKeSessionTransfer(ByVal vPrimaryBranch As Double, ByVal vPrimarySession As Double, ByVal vConstringCabang As String, _
    ByVal vConnStringUtama As String)
        Dim ListScript As New List(Of String)

        Try
            'hapus data diserver utama untuk primary session dan primary branch ybs

            'Masukkan data header session

            'item session, item void

            'item session, item payout drop cash

            'header pos

            'item pos 

            'item pos payment

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListScript = Nothing
        End Try
    End Sub

    Private Sub BuatOpenSessionJurnal()

    End Sub

    Private Sub BuatQuickInvoice()

    End Sub

    Private Sub BuatPaymentJurnal()

    End Sub

    Private Sub BuatPaymentDisbursmentJurnal()

    End Sub

    Private Sub BuatPayOutDropCashJurnal()

    End Sub

    Private Sub BuatCloseSessionJurnal()

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim DV As DataRowView
        Dim ConstringCabang As String = ""
        Dim ConstringPusat As String = ""
        Dim TBLSession As New dseTrans_Data_Transfer.tbltrans_session_cabangDataTable
        Dim DTASession As New dseTrans_Data_TransferTableAdapters.tbltrans_session_cabangTableAdapter
        Dim RowSession As dseTrans_Data_Transfer.tbltrans_session_cabangRow

        Try
            If MsgBox("Start data transfer to main branch?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.SimpanInvoice(1)
                Me.SimpanInvoice(2)
                Me.SimpanInvoice(3)
                Me.SimpanInvoice(4)
                Me.SimpanInvoice(5)
                Me.SimpanInvoice(6)
                Me.SimpanInvoice(7)
                Me.SimpanInvoice(8)
                Me.SimpanInvoice(9)

                Exit Sub

                ConstringPusat = Me.GetConstringPusat

                For Each DV In Me.clbBranch.CheckedItems
                    Me.rowCabang = DV.Row
                    ConstringCabang = Me.GetConstring(Me.rowCabang.Server_Name, Me.rowCabang.User_Id, Me.rowCabang.Password, _
                    Me.rowCabang.Database_Name, Me.rowCabang.Port)

                    'kumpulkan data session yang sudah selesai, belum ditransfer, sesuai dengan tanggal yg dispesifikasikan 
                    DTASession.Connection.ConnectionString = ConstringCabang
                    DTASession.Fill(TBLSession, Me.dtpTo.Value, Me.dtpFrom.Value)
                    For Each RowSession In TBLSession.Rows
                        Me.TransferSession(ConstringCabang, ConstringPusat, RowSession.primary, -1)
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLSession.Dispose()
            DTASession.Dispose()
        End Try
    End Sub

    Private Function GetConstringPusat() As String
        Try
            For Each Me.rowCabang In Me.TBLCabang.Rows
                If Me.rowCabang.Main_Branch = 1 Then
                    Return Me.GetConstring(Me.rowCabang.Server_Name, Me.rowCabang.User_Id, Me.rowCabang.Password, _
                    Me.rowCabang.Database_Name, Me.rowCabang.Port)
                End If
            Next

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307;default command timeout=600000"
        End Try
    End Function

    Private Function GetConstring(ByVal vNamaServer As String, ByVal vUserId As String, ByVal vPass As String, _
    ByVal vDatabaseName As String, ByVal vPortNumber As String) As String
        Try
            Return "server=" + vNamaServer + ";user id=" + vUserId + ";Password=" + vPass + _
            ";persist security info=True;database=" + vDatabaseName + ";port=" + vPortNumber + ";default command timeout=600000"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307;default command timeout=600000"
        End Try
    End Function

    Private Sub FrmTrans_Data_Transfer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'load semua data cabang
            Me.ObjCabang.IsiCheckListBox(Me.clbBranch, Me.TBLCabang)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class