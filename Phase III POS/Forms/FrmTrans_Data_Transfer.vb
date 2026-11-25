Option Explicit On
Option Strict Off
Imports Core_Module, AccModule

Public Class FrmTrans_Data_Transfer
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris

    Dim Obj As New ClsTrans_Data_Transfer
    Dim ObjCD As New ClsCekData
    Dim ObjCabang As New ClsMast_Cabang
    Dim ObjSession As New ClsTrans_Session
    Dim ObjSessionItem As New ClsTrans_Session_Item
    Dim ObjAcc As New ClsMasterAkun
    Dim ObjUser As New ClsUser
    Dim ObjGeneralSet As New ClsTrans_POS_GeneralSetting
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjAntrianTunggal As New Core_Module.ClsAntrian
    Dim ObjData As New ClsMySQL

    Public DefPrec As Double = 0
    Public RoundingTipe As Decimal
    Public RoundingValue As Decimal

    Dim TBL As New dseTrans_Data_Transfer.tblmaster_cabangDataTable
    Dim TBLItemPaymentDetails As New DseTrans_Session.DGVPaymentDetailsDataTable

    Private Sub BtnBeginSessionTransfer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBeginSessionTransfer.Click
        Dim ij As New FrmTrans_Data_Transfer1

        Try
            '1. session transfer hanya bisa dilakukan 1 user pada saat yang bersamaan, gunakan transaksi antrian untuk hal ini, 
            'antrian ini langsung menembak ke database server utama, sehingga dari cabang mana pun hanya 1 orang yang melakukan transfer

            'belum set antrian ke server utama
            MsgBox("belum set antrian k server utama, sudah coding")

            '2. transfer data bisa di pause, dan kemudian di lanjutkan kembali. untuk data yang dipause, harus diselesaikan dahulu ketika 
            '   transfernya akan dilanjutkan, atau ada proses transfer baru. data yang setengah transfer ini juga bisa dicancel proses transfernya.
            MsgBox("fitur pause transfer belum dicoding")

            '3. alternatif cara transfer: 
            '   a. langsung melalui server mysql, via WAN atau VPN --> hanya ini yang dicoding untuk kali ini
            '   b. via email
            MsgBox("fitur transfer via email juga belum dicoding")
            '

            'Tahapan proses transfer (lihat frmTrans_Data_Tranfer1, untuk detailnya):
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkProduct.CheckedChanged
        Try
            ''MsgBox("Untuk product, artinya termasuk juga tbl_master_product_qty")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnCancelSessionTransfer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancelSessionTransfer.Click
        Try
            'balikan proses data transfer, sehingga tidak ada jejak sama sekali bahwa data session tersebut pernah di transfer
            '1. cek apakah periode tersebut sudah dikunci sebelumnya, jika sudah termasuk dalam periode yang sudah dikunci, tidak bisa cancel
            '2. cek apakah sudah ada jurnal payment disbursment untuk setiap item invoice yang diberikan, jika ada maka sebutkan nomor jurnalnya
            '   proses tidak bisa dilanjutkan
            '3. session yang ditransfer hanya boleh dicancel ketika proses transfernya sudah selesai atau sedang di pause
            '4. proses cancel dan proses transfer hanya bisa terjadi salah satu pada saat yang bersamaa, jadi tidak mungkin ada orang yang melakukan
            '   data transfer dan ada orang lain yang melakukan transfer cancellation
            '5. hapuskan semua dokumen yang muncul ketika ada proses transfer:
            '   a. jurnal open session
            '   b. jurnal close session
            '   c. jurnal sales receivement
            '   d. semua invoice yang merujuk pada pos
            '   e. pos + itemnya pada server utama
            '   f. session pada server utama
            '   g. update kembali session pada server cabang, bahwa status transfernya kembali = 0

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function BuatSQLInsertionScript(ByVal vNamaTabel As String, ByVal conStr As String, _
    ByVal Conn As MySql.Data.MySqlClient.MySqlConnection) As String
        Dim sql As String = ""
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader
        Dim Command1 As New MySql.Data.MySqlClient.MySqlCommand()
        Dim TBL As DataTable
        Dim InsertCommand As String = ""
        Dim NamaKolom As String = ""
        Dim Row As DataRow

        Try
            sql = "SELECT * FROM " + vNamaTabel + ";"
            InsertCommand = "INSERT INTO `" + vNamaTabel + "`"
            Command1.CommandText = sql
            Command1.Connection = Conn

            Conn.Open()
            Reader = Command1.ExecuteReader
            TBL = Reader.GetSchemaTable


            For Each Row In TBL.Rows
                If NamaKolom = "" Then
                    NamaKolom = "`" + Row(0).ToString + "`"
                Else
                    NamaKolom = NamaKolom + ", `" + Row(0).ToString + "`"
                End If
            Next

            InsertCommand = InsertCommand + " (" + NamaKolom + ") VALUES "

            Conn.Close()
            TBL.Dispose()

            Return InsertCommand
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            Command1.Dispose()
        End Try
    End Function

    Public Function MengecekStringKarakterPetik(ByVal karakter As String) As String
        Dim Retval As String = ""
        Dim i As Integer = 0
        Dim Index As Double = 0
        Dim KarakterDummy As String = ""
        Dim PenandaString As String = ""
        Dim ListString As New List(Of String)
        Dim ParamString As String = ""
        Try
            If karakter.IndexOf("'") > 0 Then
                While i <= karakter.Length - 1
                    If i = 0 Then
                        KarakterDummy = karakter.Substring(0, karakter.IndexOf("'") + 1)
                        ListString.Add(KarakterDummy + "'")
                        i = KarakterDummy.Length
                    Else
                        PenandaString = karakter.Substring(i)
                        If PenandaString.IndexOf("'") > 0 Then
                            KarakterDummy = PenandaString.Substring(i, PenandaString.IndexOf("'") + 1)
                            ListString.Add(KarakterDummy + "'")
                            i = KarakterDummy.Length
                        Else
                            ListString.Add(karakter.Substring(i))
                            Exit While
                        End If
                    End If
                End While
                For Each ParamString In ListString
                    Retval += ParamString
                Next
            Else
                Retval = karakter
            End If

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ListString = Nothing
        End Try
    End Function

    Private Sub HapusDataCabang(ByVal vNamaTabel As String, ByVal ConStrCabang As String)
        Dim conStr As String = ConStrCabang
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection()
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript()

        Try
            'open connection
            Conn.ConnectionString = conStr
            script.Connection = Conn
            script.Query = sql
            Conn.Open()

            sql = "DELETE FROM " + vNamaTabel
            script.Query = sql
            'execute script
            script.Execute()

            Conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Sub

    Private Sub TransMasterProduct(ByVal vListPrimaryCabang As List(Of Double), ByVal ListConnStringCab As List(Of String), _
    ByVal vConnStringUtama As String)
        Dim vPrimary As Double = -1
        Dim vConnStringCab As String = ""
        Dim ListScriptProd As New List(Of String)
        Dim ListScriptProdPrice As New List(Of String)
        Dim ListScriptProdQty As New List(Of String)

        Try
            For i As Integer = 0 To vListPrimaryCabang.Count - 1
                vPrimary = vListPrimaryCabang(i)
                vConnStringCab = ListConnStringCab(i)
                'Hapus data cabang
                Me.HapusDataCabang("tblmaster_product", vConnStringCab)
                Me.HapusDataCabang("tblmaster_productqty", vConnStringCab)
                'Buat List Data Cabang
                ListScriptProd = Me.ObjData.BuatSqlScript("tblmaster_product", vConnStringUtama)
                'Insert Data Cabang 
                Me.ObjData.ExecuteSqlScript(vConnStringCab, ListScriptProd, Me.ProgressBar2, Me.LblInsertStatus)
                'Buat List Data Cabang
                ListScriptProdPrice = Me.ObjData.BuatSqlScript("tblmaster_product_pricing", vConnStringUtama)
                'Insert Data Cabang 
                Me.ObjData.ExecuteSqlScript(vConnStringCab, ListScriptProdPrice, Me.ProgressBar2, Me.LblInsertStatus)
                'Buat List Data Cabang
                ListScriptProdQty = Me.ObjData.BuatSqlScript("tblmaster_productqty", vConnStringUtama)
                'Insert Data Cabang 
                Me.ObjData.ExecuteSqlScript(vConnStringCab, ListScriptProdQty, Me.ProgressBar2, Me.LblInsertStatus)
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListScriptProd = Nothing
            ListScriptProdPrice = Nothing
            ListScriptProdQty = Nothing
        End Try
    End Sub

    Private Sub TransMasterDiscount(ByVal vListPrimaryCabang As List(Of Double), ByVal ListConnStringCab As List(Of String), _
    ByVal vConnStringUtama As String)
        Dim vPrimary As Double = -1
        Dim vConnStringCab As String = ""
        Dim ListScriptDisk As New List(Of String)
        Dim ListScriptDiskItemBranch As New List(Of String)
        Dim ListScriptDiscItemProd As New List(Of String)
        Try
            '2. master discount --> tblmaster_discount, tblmaster_discount_itembranch, tblmaster_discount_itemproduct
            For i As Integer = 0 To vListPrimaryCabang.Count - 1
                vPrimary = vListPrimaryCabang(i)
                vConnStringCab = ListConnStringCab(i)
                'Hapus data cabang
                Me.HapusDataCabang("tblmaster_discount", vConnStringCab)
                'Buat List Data Cabang
                ListScriptDisk = Me.ObjData.BuatSqlScript("tblmaster_discount", vConnStringUtama)
                'Insert Data Cabang 
                Me.ObjData.ExecuteSqlScript(vConnStringCab, ListScriptDisk, Me.ProgressBar2, Me.LblStatus)
                'Buat List Data Cabang
                ListScriptDiskItemBranch = Me.ObjData.BuatSqlScript("tblmaster_discount_itembranch", vConnStringUtama)
                'Insert Data Cabang 
                Me.ObjData.ExecuteSqlScript(vConnStringCab, ListScriptDiskItemBranch, Me.ProgressBar2, Me.LblStatus)
                'Buat List Data Cabang
                ListScriptDiscItemProd = Me.ObjData.BuatSqlScript("tblmaster_discount_itemproduct", vConnStringUtama)
                'Insert Data Cabang 
                Me.ObjData.ExecuteSqlScript(vConnStringCab, ListScriptDiscItemProd, Me.ProgressBar2, Me.LblStatus)
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListScriptDisk = Nothing
            ListScriptDiskItemBranch = Nothing
            ListScriptDiscItemProd = Nothing
        End Try
    End Sub

    Private Sub TransMasterPaymentOption(ByVal vListPrimaryCabang As List(Of Double), ByVal ListConnStringCab As List(Of String), _
    ByVal vConnStringUtama As String)
        Dim vPrimary As Double = -1
        Dim vConnStringCab As String = ""
        Dim ListScriptDisk As New List(Of String)
        Try
            For i As Integer = 0 To vListPrimaryCabang.Count - 1
                vPrimary = vListPrimaryCabang(i)
                vConnStringCab = ListConnStringCab(i)
                'Hapus data cabang
                Me.HapusDataCabang("tblmaster_payment_option", vConnStringCab)
                'Buat List Data Cabang
                ListScriptDisk = Me.ObjData.BuatSqlScript("tblmaster_payment_option", vConnStringUtama)
                'Insert Data Cabang 
                Me.ObjData.ExecuteSqlScript(vConnStringCab, ListScriptDisk, Me.ProgressBar2, Me.LblStatus)
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListScriptDisk = Nothing
        End Try
    End Sub

    Private Sub TransMasterCabang(ByVal vListPrimaryCabang As List(Of Double), ByVal ListConnStringCab As List(Of String), _
    ByVal vConnStringUtama As String)
        Dim vPrimary As Double = -1
        Dim vConnStringCab As String = ""
        Dim ListScriptDisk As New List(Of String)
        Dim ListScriptDiskItemBranch As New List(Of String)
        Try
            '4. master branch --> tblmaster_cabang, tblmaster_cabang_itemlocation
            For i As Integer = 0 To vListPrimaryCabang.Count - 1
                vPrimary = vListPrimaryCabang(i)
                vConnStringCab = ListConnStringCab(i)
                'Hapus data cabang
                Me.HapusDataCabang("tblmaster_cabang", vConnStringCab)
                'Buat List Data Cabang
                ListScriptDisk = Me.ObjData.BuatSqlScript("tblmaster_cabang", vConnStringUtama)
                'Insert Data Cabang 
                Me.ObjData.ExecuteSqlScript(vConnStringCab, ListScriptDisk, Me.ProgressBar2, Me.LblStatus)
                'Buat List Data Cabang
                ListScriptDiskItemBranch = Me.ObjData.BuatSqlScript("tblmaster_cabang_itemlocation", vConnStringUtama)
                'Insert Data Cabang 
                Me.ObjData.ExecuteSqlScript(vConnStringCab, ListScriptDiskItemBranch, Me.ProgressBar2, Me.LblStatus)
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListScriptDisk = Nothing
            ListScriptDiskItemBranch = Nothing
        End Try
    End Sub

    Private Sub TransMasterMember(ByVal vListPrimaryCabang As List(Of Double), ByVal ListConnStringCab As List(Of String), _
    ByVal vConnStringUtama As String)
        Dim vPrimary As Double = -1
        Dim vConnStringCab As String = ""
        Dim ListScriptDisk As New List(Of String)
        Try
            '5. master member --> tblmaster_member
            For i As Integer = 0 To vListPrimaryCabang.Count - 1
                vPrimary = vListPrimaryCabang(i)
                vConnStringCab = ListConnStringCab(i)
                'Hapus data cabang
                Me.HapusDataCabang("tblmaster_member", vConnStringCab)
                'Buat List Data Cabang
                ListScriptDisk = Me.ObjData.BuatSqlScript("tblmaster_member", vConnStringUtama)
                'Insert Data Cabang 
                Me.ObjData.ExecuteSqlScript(vConnStringCab, ListScriptDisk, Me.ProgressBar2, Me.LblStatus)
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListScriptDisk = Nothing
        End Try
    End Sub

    Private Sub TransMasterPayout(ByVal vListPrimaryCabang As List(Of Double), ByVal ListConnStringCab As List(Of String), _
    ByVal vConnStringUtama As String)
        Dim vPrimary As Double = -1
        Dim vConnStringCab As String = ""
        Dim ListScriptDisk As New List(Of String)
        Try
            '6. master pay out drop cash --> tblmaster_payout_dropcash
            For i As Integer = 0 To vListPrimaryCabang.Count - 1
                vPrimary = vListPrimaryCabang(i)
                vConnStringCab = ListConnStringCab(i)
                'Hapus data cabang
                Me.HapusDataCabang("tblmaster_payout_dropcash", vConnStringCab)
                'Buat List Data Cabang
                ListScriptDisk = Me.ObjData.BuatSqlScript("tblmaster_payout_dropcash", vConnStringUtama)
                'Insert Data Cabang 
                Me.ObjData.ExecuteSqlScript(vConnStringCab, ListScriptDisk, Me.ProgressBar2, Me.LblStatus)
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListScriptDisk = Nothing
        End Try
    End Sub

    Private Sub TransMasterGeneralSetting(ByVal vListPrimaryCabang As List(Of Double), ByVal ListConnStringCab As List(Of String), _
    ByVal vConnStringUtama As String)
        Dim vPrimary As Double = -1
        Dim vConnStringCab As String = ""
        Dim ListScriptDisk As New List(Of String)
        Dim ListScriptDiskItemBranch As New List(Of String)
        Try
            '7. master general setting --> tbltrans_pos_general_setting, tbltrans_pos_general_setting_item
            For i As Integer = 0 To vListPrimaryCabang.Count - 1
                vPrimary = vListPrimaryCabang(i)
                vConnStringCab = ListConnStringCab(i)
                'Hapus data cabang
                Me.HapusDataCabang("tbltrans_pos_general_setting", vConnStringCab)
                'Buat List Data Cabang
                ListScriptDisk = Me.ObjData.BuatSqlScript("tbltrans_pos_general_setting", vConnStringUtama)
                'Insert Data Cabang 
                Me.ObjData.ExecuteSqlScript(vConnStringCab, ListScriptDisk, Me.ProgressBar2, Me.LblStatus)
                'Buat List Data Cabang
                ListScriptDiskItemBranch = Me.ObjData.BuatSqlScript("tbltrans_pos_general_setting_item", vConnStringUtama)
                'Insert Data Cabang 
                Me.ObjData.ExecuteSqlScript(vConnStringCab, ListScriptDiskItemBranch, Me.ProgressBar2, Me.LblStatus)
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListScriptDisk = Nothing
            ListScriptDiskItemBranch = Nothing
        End Try
    End Sub

    Private Function ApakahBolehTransferMaster() As Boolean
        Try
            If Me.clbBranch.CheckedItems.Count <= 0 Then
                MsgBox("Please choose branch to transfer", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.ChkProduct.Checked = False Then
                If Me.ChkDiscount.Checked = False Then
                    If Me.ChkPaymentOption.Checked = False Then
                        If Me.ChkBranch.Checked = False Then
                            If Me.ChkMemberData.Checked = False Then
                                If Me.ChkPayoutDropCash.Checked = False Then
                                    If ChkGeneralSetting.Checked = False Then
                                        MsgBox("Please choose data to transfer", MsgBoxStyle.Information, "Information")
                                        Return False
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnTransferMaster_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTransferMaster.Click
        Dim ConnStringUtama As String = ""
        Dim ListCabang As New List(Of Double)
        Dim ListConn As New List(Of String)
        Dim vDefaultCommandTimeOut As Double = 6000
        Dim Row As dseTrans_Data_Transfer.tblmaster_cabangRow
        Dim Hasil As MsgBoxResult
        Try
            Me.ProgressBar1.Visible = True
            Me.ProgressBar2.Visible = True
            Me.LblInsertStatus.Visible = True
            Me.LblStatus.Visible = True

            If Me.ApakahBolehTransferMaster() = True Then
                '0. Jika hanya ada 1 cabang (yaitu cabang utama), maka proses data transfer ini tidak bisa dilakukan
                If ObjCabang.GetJmlUntukCabangSaja <= 0 Then
                    MsgBox("Data transfer cannot be done because there is no branch data in the master branch", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                Hasil = MsgBox("Transfer data ?", MsgBoxStyle.YesNo, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    '1. untuk setiap tabel master yang di check list, transfer data master yang dicheck list
                    For i As Integer = 0 To Me.clbBranch.Items.Count - 1
                        Row = Me.TBL.Rows(i)
                        If Me.clbBranch.GetItemChecked(i) = True Then
                            If ListCabang.Contains(Row.Primary) = False Then
                                ListCabang.Add(Row.Primary)
                                ListConn.Add("server=" + ObjCabang.GetServer(Row.Primary) + ";user id=" + ObjCabang.GetUserID(Row.Primary) _
                                + ";Password=" + ObjCabang.GetPassword(Row.Primary) + ";persist security info=True;database=" + _
                                ObjCabang.GetDatabaseName(Row.Primary) + ";port=" + ObjCabang.GetPort(Row.Primary) + _
                                ";default command timeout=" + vDefaultCommandTimeOut.ToString)
                            End If
                        End If
                    Next
                    '2. ambil constring cabang utama, siapkan sql script berdasarkan data pada cabang utama 
                    ConnStringUtama = "server=" + ObjCabang.GetServerNameUtama() + ";user id=" + ObjCabang.GetUserIDUtama() _
                                    + ";Password=" + ObjCabang.GetPasswordUtama() + ";persist security info=True;database=" + _
                                    ObjCabang.GetDatabaseNameUtama() + ";port=" + ObjCabang.GetPortUtama() + _
                                    ";default command timeout=" + vDefaultCommandTimeOut.ToString
                    '3. Untuk setiap cabang yang diupdate
                    '   a. ambil constring cabang
                    '   b. hapuskan data cabang
                    '   c. insert data cabang
                    'nama tabel yang ditransfer:
                    '1. master product --> tblmaster_product, tblmaster_product_pricing, tblmaster_productqty
                    If Me.ChkProduct.Checked = True Then
                        Me.TransMasterProduct(ListCabang, ListConn, ConnStringUtama)
                    End If
                    '2. master discount --> tblmaster_discount, tblmaster_discount_itembranch, tblmaster_discount_itemproduct
                    If Me.ChkDiscount.Checked = True Then
                        Me.TransMasterDiscount(ListCabang, ListConn, ConnStringUtama)
                    End If
                    '3. master payment option --> tblmaster_payment_option
                    If Me.ChkPaymentOption.Checked = True Then
                        Me.TransMasterPaymentOption(ListCabang, ListConn, ConnStringUtama)
                    End If
                    '4. master branch --> tblmaster_cabang, tblmaster_cabang_itemlocation
                    If Me.ChkBranch.Checked = True Then
                        Me.TransMasterCabang(ListCabang, ListConn, ConnStringUtama)
                    End If
                    '5. master member --> tblmaster_member
                    If Me.ChkMemberData.Checked = True Then
                        Me.TransMasterMember(ListCabang, ListConn, ConnStringUtama)
                    End If
                    '6. master pay out drop cash --> tblmaster_payout_dropcash
                    If Me.ChkPayoutDropCash.Checked = True Then
                        Me.TransMasterPayout(ListCabang, ListConn, ConnStringUtama)
                    End If
                    '7. master general setting --> tbltrans_pos_general_setting, tbltrans_pos_general_setting_item
                    If Me.ChkGeneralSetting.Checked = True Then
                        Me.TransMasterGeneralSetting(ListCabang, ListConn, ConnStringUtama)
                    End If
                    'progress berjalan dengan pembagian, kalo ada 2 cabang, maka ada 14 progress (asumsi semua di update)
                    'statunya di tuliskan "Updating Master Product on Cabang Kepatihan (Step 1 of 14)"
                    MsgBox("Finished", MsgBoxStyle.Information, "Information")
                End If
            End If

            Me.ProgressBar1.Visible = False
            Me.ProgressBar2.Visible = False
            Me.LblInsertStatus.Visible = False
            Me.LblStatus.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListCabang = Nothing
            ListConn = Nothing
        End Try
    End Sub

    Private Sub FrmTrans_Data_Transfer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim NamaUserLain, NamaKomputerLain As String

        Try
            NamaUserLain = ""
            NamaKomputerLain = ""
            If Me.ObjAntrianTunggal.StartAntrianTunggal(Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransferDataPOS, _
            Me.UserCode, NamaUserLain, NamaKomputerLain) = False Then
                Me.Close()
            End If

            'pada clb branch, isikan data cabang, kecuali cabang utama
            Me.Obj.IsiListBoxBrand(Me.clbBranch, Me.TBL)
            '-----
            Me.chk_PaymentS.Checked = False
            Me.dgvAccount_OpSes.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvAccount_CloseSes.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgv_DetailTrans.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgv_cancelled_Trans.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgv_Payment_Sum.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgv_PaymentD.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgv_PayOutS.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgv_CancelledItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVPendingTrans.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DefPrec = ObjCur.GetDefaultPrecission(ObjGeneralSet.GetPrimaryCurr())
            If ObjGeneralSet.GetAllowRounding = 0 Then
                Me.RoundingTipe = ClsCekData.EnumTipeRounding.Middlepoint
                Me.RoundingValue = 1
            ElseIf ObjGeneralSet.GetAllowRounding = 1 Then
                Me.RoundingTipe = ObjGeneralSet.GetRoundingTipe
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.A_Seribu Then
                    Me.RoundingValue = 1000
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.B_LimaRatus Then
                    Me.RoundingValue = 500
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.C_Seratus Then
                    Me.RoundingValue = 100
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.D_LimaPuluh Then
                    Me.RoundingValue = 50
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.E_Sepuluh Then
                    Me.RoundingValue = 10
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.F_Lima Then
                    Me.RoundingValue = 5
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.G_Satu Then
                    Me.RoundingValue = 1
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.H_NolKomaLima Then
                    Me.RoundingValue = 0.5
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.I_NolKomaSatu Then
                    Me.RoundingValue = 0.1
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.J_NolKomaNolLima Then
                    Me.RoundingValue = 0.05
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.K_NolKomaNolSatu Then
                    Me.RoundingValue = 0.01
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.L_NolKomaNolNolLima Then
                    Me.RoundingValue = 0.005
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.M_NolKomaNolNolSatu Then
                    Me.RoundingValue = 0.001
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            For i As Integer = 0 To Me.clbBranch.Items.Count - 1
                Me.clbBranch.SetItemChecked(i, True)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            For i As Integer = 0 To Me.clbBranch.Items.Count - 1
                Me.clbBranch.SetItemChecked(i, False)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindSession_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindSession.Click
        Dim ijk As New FrmTransaksiSessionCari
        Try
            ijk.BackColorAll = Me.BackColorAll
            ijk.cboCriteria.Items.Clear()
            ijk.txtCari.Text = ""
            ijk.dgvHasil.DataSource = ""
            ijk.ShowDialog()

            If Not ijk.mPrimary = -1 Then
                If Me.ObjSession.CariDetailDataTransfer(ijk.mPrimary, Me.ObjCD, Me, Me.TBLItemPaymentDetails) = True Then
                    Me.SalinKeField()
                    Me.BtnBeginSessionTransfer.Focus()
                Else
                    Me.BtnFindSession.Focus()
                End If
                Me.BtnFindSession.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijk.Dispose()
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjSession
                Me.TxtPrimary.Text = .Primary
                Me.TxtSessionID.Text = .SessionID
                Me.dtpSessionPeriod.Value = Format(.StartDate, "dd MMM yyyy") + " " + .StartDate.Hour.ToString + ":" + .StartDate.Minute.ToString
                Me.DtpSessionThrough.Value = Format(.EndDate, "dd MMM yyyy") + " " + .EndDate.Hour.ToString + ":" + .EndDate.Minute.ToString
                Me.TxtRemarks.Text = .Remarks
                If .Status = 0 Then
                    Me.cboStatus.SelectedIndex = 0
                Else
                    Me.cboStatus.SelectedIndex = 1
                End If
                Me.TxtPrimaryCashier.Text = "-1"
                Me.TxtCashierCode.Text = .CashierCode
                Me.TxtCashierName.Text = ObjUser.GetNama(.CashierCode)
                Me.TxtPrimaryBranch.Text = .PrimaryBranch
                Me.TxtKodeBranch.Text = ObjCabang.GetCode(.PrimaryBranch)
                Me.TxtKodeBranchDesc.Text = ObjCabang.GetDesc(.PrimaryBranch)
                Me.TxtCash_Beginning.Text = ObjCD.DigitGrouping(.BeginningCash, Me.DefPrec, Me.RoundingTipe)
                Me.TxtPrimaryBeginingCashAccount.Text = .PrimaryBeginningCA
                Me.TxtBeginingCashAccountCode.Text = ObjAcc.GetAccountCode(.PrimaryBeginningCA)
                Me.TxtBeginingCashAccountDesc.Text = ObjAcc.GetDescription(Me.TxtBeginingCashAccountCode.Text)
                Me.TxtPrimaryCashierCashAccount.Text = .PrimaryCashierCA
                Me.TxtCashierCashAccountCode.Text = ObjAcc.GetAccountCode(.PrimaryCashierCA)
                Me.TxtCashierCashAccountDesc.Text = ObjAcc.GetDescription(Me.TxtCashierCashAccountCode.Text)
                Me.TxtPrimaryBalanceCashAccount.Text = .PrimaryBalanceCA
                Me.TxtBalanceCashAccountCode.Text = ObjAcc.GetAccountCode(.PrimaryBalanceCA)
                Me.TxtBalanceCashAccountDesc.Text = ObjAcc.GetDescription(Me.TxtBalanceCashAccountCode.Text)

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chk_PaymentS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_PaymentS.CheckedChanged
        Try
            If Me.chk_PaymentS.Checked = True Then
                Me.spc_Payment.Panel2Collapsed = False
            Else
                Me.spc_Payment.Panel2Collapsed = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgv_Payment_Sum_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_Payment_Sum.CellClick
        Try
            If Me.dgv_Payment_Sum.Rows.Count > 0 Then
                ObjSessionItem.IsiDGVPaymentDetailsDataTransfer(Me.dgv_Payment_Sum.SelectedRows(0).Cells(Me.Col_PaymentS_PrimaryPayment.Index).Value, _
                Me, Me.ObjCD, Me.TBLItemPaymentDetails)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgv_DetailTrans_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgv_DetailTrans.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.col_TransD_ShowDetail.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgv_DetailTrans_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_DetailTrans.CellClick
        Dim iFrm As New FrmTrans_Session_View
        Try
            If Me.TxtRemarks.Enabled = True Then
                If e.ColumnIndex = Me.col_TransD_ShowDetail.Index Then
                    iFrm.mPrimaryPOS = Me.dgv_DetailTrans.Item(Me.Col_TransD_PrimaryPOS.Index, e.RowIndex).Value
                    iFrm.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            iFrm.Dispose()
        End Try
    End Sub

    Private Sub dgv_cancelled_Trans_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgv_cancelled_Trans.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.Col_CancelledTrans_ShowDetails.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgv_cancelled_Trans_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_cancelled_Trans.CellClick
        Dim iFrm As New FrmTrans_Session_View
        Try
            If Me.TxtRemarks.Enabled = True Then
                If e.ColumnIndex = Me.Col_CancelledTrans_ShowDetails.Index Then
                    iFrm.mPrimaryPOS = Me.dgv_cancelled_Trans.Item(Me.Col_CancelledTrans_PrimaryPOS.Index, e.RowIndex).Value
                    iFrm.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            iFrm.Dispose()
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVPendingTrans_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVPendingTrans.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.Col_CancelledTrans_ShowDetails.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVPendingTrans_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVPendingTrans.CellClick
        Dim iFrm As New FrmTrans_Session_View
        Try
            If Me.TxtRemarks.Enabled = True Then
                If e.ColumnIndex = Me.Col_CancelledTrans_ShowDetails.Index Then
                    iFrm.mPrimaryPOS = Me.dgv_cancelled_Trans.Item(Me.Col_CancelledTrans_PrimaryPOS.Index, e.RowIndex).Value
                    iFrm.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            iFrm.Dispose()
        End Try
    End Sub
End Class