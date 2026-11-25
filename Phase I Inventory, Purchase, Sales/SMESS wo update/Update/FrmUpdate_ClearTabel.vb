Option Explicit On
Option Strict Off

Imports MySql.Data.MySqlClient
Imports MySql.Data

Partial Public Class FrmUpdate
    Dim TBL_CT As New DseUpdate_ClearTable.tbltrans_clear_dataDataTable
    Dim row_CT As DseUpdate_ClearTable.tbltrans_clear_dataRow
    Dim CT_ErrorReason As String = "Invalid connection string"

    Private Sub SalinKeField()
        Try
            If Me.cboCT_SchemeName.SelectedIndex <> -1 Then
                Me.row_CT = Me.TBL_CT.DefaultView(Me.cboCT_SchemeName.SelectedIndex).Row

                Me.tslCT_CreatedBy.Text = "Created by: " + row_CT.created_by
                Me.tslCT_CreatedDate.Text = "Created date: " + Format(row_CT.created_date, "dd MMM yyyy HH:mm")
                Me.tslCT_ModifiedBy.Text = "Modified by: " + row_CT.modified_by
                Me.tslCT_ModifiedDate.Text = "Modified date: " + Format(row_CT.modified_date, "dd MMM yyyy HH:mm")

                Me.TxtCT_ConnectionString.Text = Me.row_CT.target_constring
                Me.TxtCT_DBName.Text = Me.row_CT.database_name
                Me.TxtCT_Primary.Text = Me.row_CT.primary
                Me.TxtCT_NamaScheme.Text = Me.row_CT.description
            Else
                Me.tslCT_CreatedBy.Text = "Created by: "
                Me.tslCT_CreatedDate.Text = "Created date: "
                Me.tslCT_ModifiedBy.Text = "Modified by: "
                Me.tslCT_ModifiedDate.Text = "Modified date: "

                Me.TxtCT_ConnectionString.Text = ""
                Me.TxtCT_DBName.Text = ""
                Me.TxtCT_Primary.Text = "-1"
                Me.TxtCT_NamaScheme.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadCBO()
        Dim DTA As New DseUpdate_ClearTableTableAdapters.tbltrans_clear_dataTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Desc(Me.TBL_CT)
            Me.cboCT_SchemeName.DataSource = Me.TBL_CT.DefaultView
            Me.cboCT_SchemeName.DisplayMember = "description"

            If Me.cboCT_SchemeName.Items.Count > 0 Then
                Me.cboCT_SchemeName.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub LoadTableName()
        Dim DTA As New DseUpdate_ClearTableTableAdapters.tbltrans_clear_dataTableAdapter
        Dim DTAItem As New DseUpdate_ClearTableTableAdapters.tbltrans_clear_data_itemTableAdapter
        Dim TBLItem As New DseUpdate_ClearTable.tbltrans_clear_data_itemDataTable
        'Dim rowItem As DseUpdate_ClearTable.tbltrans_clear_data_itemRow
        Dim rowDGV As DataGridViewRow
        Dim DTAUser As New Core_Module.DSUserTableAdapters.tblmasterhakaksesTableAdapter
        Dim TBLDummy As New Core_Module.DSUser.tblmasterhakaksesDataTable

        Try
            Me.CT_ErrorReason = "Invalid connection string"
            'pastikan connection string valid
            'Me.TxtCT_ConnectionString.Text = Core_Module.GetConstring
            'DTAUser.Connection.ConnectionString = Me.TxtCT_ConnectionString.Text
            DTAUser.Fill(TBLDummy)
            Me.CT_ErrorReason = ""

            'load semua nama tabel ke dalam tabel 
            Me.dgvCT_DetailTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            Me.Load1()

            'tandai mana saja yang diclear berdasarkan data yang ada
            If Me.TxtCT_Primary.Text <> "-1" And Me.TxtCT_Primary.Text <> "" Then
                DTAItem.FillBy_PrimaryTransaksi(TBLItem, Me.TxtCT_Primary.Text)
                For Each rowDGV In Me.dgvCT_DetailTable.Rows
                    TBLItem.DefaultView.RowFilter = "[tabel name] = '" + rowDGV.Cells(Me.ColCT_TableName.Index).Value.ToString + "'"
                    If TBLItem.DefaultView.Count > 0 Then
                        rowDGV.Cells(Me.ColCT_Clear.Index).Value = Boolean.TrueString
                    Else
                        rowDGV.Cells(Me.ColCT_Clear.Index).Value = Boolean.FalseString
                    End If
                Next
            End If

            Me.dgvCT_DetailTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            If ex.Message = "Unable to connect to any of the specified MySQL hosts." And Me.CT_ErrorReason = "Invalid connection string" Then
                MsgBox(Me.CT_ErrorReason, MsgBoxStyle.Critical, "Error")
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            TBLDummy.Dispose()
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Sub

    Private Sub Load1()
        Dim connection As New MySqlConnection(Me.TxtCT_ConnectionString.Text)
        Dim tablesCommand As MySqlCommand = connection.CreateCommand()
        Dim rowsCommand As MySqlCommand = connection.CreateCommand()
        Dim tablesReader As MySqlDataReader
        Dim intA As Integer = 0
        Dim JmlBaris As Double = 0
        Dim tablesArray As New ArrayList()

        Try
            Me.dgvCT_DetailTable.Rows.Clear()

            connection.Open()
            tablesCommand.CommandText = "SHOW FULL TABLES FROM " & Me.TxtCT_DBName.Text & " WHERE Table_type <> 'VIEW'"
            tablesReader = tablesCommand.ExecuteReader()
            While tablesReader.Read()
                tablesArray.Add(tablesReader.GetValue(0).ToString())
            End While
            connection.Close()

            For intA = 0 To tablesArray.Count - 1
                connection.Open()

                tablesCommand.CommandText = "SELECT COUNT(*) FROM " + tablesArray(intA) + ";"
                tablesReader = tablesCommand.ExecuteReader()
                While tablesReader.Read()
                    JmlBaris = CDbl(tablesReader.GetValue(0).ToString())
                End While

                Me.dgvCT_DetailTable.Rows.Add(Me.dgvCT_DetailTable.Rows.Count + 1, tablesArray(intA), JmlBaris, Boolean.FalseString)

                connection.Close()
            Next

            tablesReader.Close()
            tablesCommand.Dispose()
            'connection.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCT_LoadTableName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCT_LoadTableName.Click
        Try
            If MsgBox("Reload table name?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.LoadTableName()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCT_Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCT_Clear.Click
        Try
            If Me.dgvCT_DetailTable.RowCount > 0 Then
                If MsgBox("Start clear table?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.ClearTable()
                    Me.LoadTableName()
                    MsgBox("Table cleared.", MsgBoxStyle.Information, "Information")
                End If
            Else
                MsgBox("Please specify table to clear", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub ClearTable()
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection()
        Dim script As New MySql.Data.MySqlClient.MySqlScript()
        Dim ObjCurrency As New Core_Module.clsMasterCurrency

        Try
            'open connection
            Conn.ConnectionString = Me.TxtCT_ConnectionString.Text
            script.Connection = Conn
            script.Query = ""
            Conn.Open()

            For Each rowDGV As DataGridViewRow In Me.dgvCT_DetailTable.Rows
                If rowDGV.Cells(Me.ColCT_Clear.Index).Value = Boolean.TrueString Then
                    script.Query = "DELETE FROM " + rowDGV.Cells(Me.ColCT_TableName.Index).Value.ToString + ";" + vbCrLf + "COMMIT;"

                    If rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tblmaster_akun" Then
                        'execute script
                        script.Execute()
                        'masukkan data awal akun
                        Me.InsertMasterAkun_Awal()
                    ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tblmaster_currency" Then
                        'kecualikan IDR
                        script.Query = "DELETE FROM " + rowDGV.Cells(Me.ColCT_TableName.Index).Value.ToString + " WHERE `primary` <> 1;" + vbCrLf + "COMMIT;"
                        'execute script
                        script.Execute()
                    ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tblmaster_usergroup" Then
                        script.Query = "DELETE FROM " + rowDGV.Cells(Me.ColCT_TableName.Index).Value.ToString + " WHERE `Primary` <> 1;" + vbCrLf + "COMMIT;"
                        'execute script
                        script.Execute()
                    ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tbluser" Then
                        script.Query = "DELETE FROM " + rowDGV.Cells(Me.ColCT_TableName.Index).Value.ToString + " WHERE `Kode User` <> 'USER';" + vbCrLf + "COMMIT;"
                        'execute script
                        script.Execute()
                    ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tblantarauserhakakses" Then
                        script.Query = "DELETE FROM " + rowDGV.Cells(Me.ColCT_TableName.Index).Value.ToString + " WHERE `Kode User` <> 'USER';" + vbCrLf + "COMMIT;"
                        'execute script
                        script.Execute()
                    Else
                        'execute script
                        script.Execute()
                    End If
                End If
            Next

            Conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Sub

    'Private Sub InsertMasterCurrency_Awal(ByVal vPrimaryAkunAP As Double, ByVal vPrimaryAkunAR As Double, ByVal vPrimaryAkunRealizedGainLoss As Double, _
    'ByVal vprimaryUnRealizedGainnLoss As Double)
    '    Dim ObjAkun As New Core_Module.ClsMasterAkun1
    '    Dim Conn As New MySql.Data.MySqlClient.MySqlConnection()
    '    Dim script As New MySql.Data.MySqlClient.MySqlScript()
    '    Dim ObjCurrency As New Core_Module.clsMasterCurrency

    '    Try
    '        'cek apakah data akunnya ada, jika tidak ada maka:
    '        If ObjAkun.ApakahSudahAda_ByPrimary(vPrimaryAkunAP) = False Then
    '            '1. insert akunnya
    '            Conn.ConnectionString = Me.TxtCT_ConnectionString.Text
    '            script.Connection = Conn
    '            script.Query = ""
    '            Conn.Open()

    '            script.Query = "INSERT INTO `tblmaster_akun` (`Primary`,`Account Code`,`Description`,`Account Nature`,`Aktif`,`Header`,`Beginning Balance`,`Tipe Perkiraan`,`Level Perkiraan`,`Tampilan`,`Saldo Normal`,`Urutan`,`Primary Lvl 1`,`Primary Lvl 2`,`Primary Lvl 3`,`Created By`,`Created Date`,`Modified By`,`Modified Date`,`Timestamp`,`primary currency`,`Beggining Balance Original Cur`,`Tipe Perkiraan2`,`Primary Lvl 4`,`Primary Lvl 5`,`Primary Lvl 6`,`Primary Lvl 7`) VALUES " + _
    '            "(10,'1.0','PIUTANG IDR','',1,0,0.000000,0,2,'     1.0 - PIUTANG IDR',0,1,1,-1,-1,'USER','2015-12-03 21:46:47','USER','2015-12-03 21:47:08','2015-12-03 21:47:08',1,0.000000,1,-1,-1,-1,-1)," + _
    '            "(11,'2.0','HUTANG IDR','',1,0,0.000000,1,2,'     2.0 - HUTANG IDR',1,3,2,-1,-1,'USER','2015-12-03 21:47:22','USER','2015-12-03 21:47:22','2015-12-03 21:47:22',1,0.000000,7,-1,-1,-1,-1)," + _
    '            "(12,'6.0','BEBAN SELISIH KURS','',1,0,0.000000,4,2,'     6.0 - BEBAN SELISIH KURS',0,8,6,-1,-1,'USER','2015-12-03 21:47:43','USER','2015-12-03 21:47:43','2015-12-03 21:47:43',1,0.000000,13,-1,-1,-1,-1);" + _
    '            "COMMIT;"

    '            'execute script
    '            script.Execute()

    '            Conn.Close()

    '            '2. update primary
    '            ObjCurrency.CariDetail(1)
    '            ObjCurrency.PrimaryPayableAccount = 10
    '            ObjCurrency.PrimaryReceivableAccount = 11
    '            ObjCurrency.PrimaryUnrealizedExchangeGainLossAccount = 12
    '            ObjCurrency.PrimaryRealizedExchangeGainLossAccount = 12
    '            ObjCurrency.DefaultPrecision = 0
    '            ObjCurrency.SimpanPerubahan(ObjCurrency.CreatedBy)
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjAkun = Nothing
    '    End Try
    'End Sub

    Private Sub InsertMasterAkun_Awal()
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection()
        Dim script As New MySql.Data.MySqlClient.MySqlScript()
        Dim ObjCurrency As New Core_Module.clsMasterCurrency

        Try
            'open connection
            Conn.ConnectionString = Me.TxtCT_ConnectionString.Text
            script.Connection = Conn
            script.Query = ""
            Conn.Open()

            script.Query = "INSERT INTO `tblmaster_akun` (`Primary`, `Account Code`, `Description`, `Account Nature`, `Aktif`, `Header`, `Beginning Balance`, `Tipe Perkiraan`, `Level Perkiraan`, `Tampilan`, `Saldo Normal`, `Urutan`, `Primary Lvl 1`, `Primary Lvl 2`, `Primary Lvl 3`, `Created By`, `Created Date`, `Modified By`, `Modified Date`, `Timestamp`, `primary currency`, `Beggining Balance Original Cur`, `Tipe Perkiraan2`, `Primary Lvl 4`, `Primary Lvl 5`, `Primary Lvl 6`, `Primary Lvl 7`) VALUES" + _
            "(1,'1','ASSET','Berisikan semua data asset perusahaan',1,1,0,0,1,'1 - ASSET',0,0,-1,-1,-1,'USER','2010-07-01 12:15:40','USER','2013-04-23 14:15:16','2013-04-24 04:15:16',1,0,17,-1,-1,-1,-1)," + _
            "(2,'2','KEWAJIBAN','',1,1,0,1,1,'2 - KEWAJIBAN',1,2,-1,-1,-1,'USER','2010-07-01 12:16:56','USER','2013-04-23 15:01:33','2015-12-04 21:52:32',1,0,18,-1,-1,-1,-1)," + _
            "(3,'3','EKUITAS','',1,1,0,2,1,'3 - EKUITAS',1,4,-1,-1,-1,'USER','2010-07-01 12:24:46','USER','2013-04-23 15:22:39','2015-12-04 21:52:36',1,0,10,-1,-1,-1,-1)," + _
            "(4,'4','PENDAPATAN USAHA','',1,1,0,3,1,'4 - PENDAPATAN USAHA',1,5,-1,-1,-1,'USER','2010-07-01 12:25:45','USER','2013-04-23 15:25:09','2015-12-04 21:52:37',1,0,11,-1,-1,-1,-1)," + _
            "(5,'5','BEBAN POKOK PENJUALAN','',1,1,0,5,1,'5 - BEBAN POKOK PENJUALAN',0,6,-1,-1,-1,'USER','2010-07-01 12:27:27','USER','2013-04-23 15:30:26','2015-12-04 21:52:38',1,0,12,-1,-1,-1,-1)," + _
            "(6,'6','BEBAN OPERASIONAL','',1,1,0,4,1,'6 - BEBAN OPERASIONAL',0,7,-1,-1,-1,'USER','2010-07-01 12:29:19','USER','2013-04-23 15:56:53','2015-12-04 21:52:40',1,0,13,-1,-1,-1,-1)," + _
            "(7,'7','PENDAPATAN LAIN - LAIN','',1,1,0,6,1,'7 - PENDAPATAN LAIN - LAIN',1,9,-1,-1,-1,'USER','2010-07-01 12:30:23','USER','2013-04-23 16:22:34','2015-12-04 21:52:42',1,0,14,-1,-1,-1,-1)," + _
            "(8,'8','PENGELUARAN LAIN - LAIN','',1,1,0,7,1,'8 - PENGELUARAN LAIN - LAIN',0,10,-1,-1,-1,'USER','2010-07-01 12:31:26','USER','2013-04-23 16:24:36','2015-12-04 21:52:43',1,0,15,-1,-1,-1,-1)," + _
            "(9,'9.000.000','TAX','-',1,1,0,8,1,'9.000.000 - TAX',0,11,-1,-1,-1,'USER','2010-07-13 10:41:32','ILHAM','2015-05-15 15:15:37','2015-12-04 21:52:47',1,0,16,-1,-1,-1,-1)," + _
            "(10,'1.0','PIUTANG IDR','',1,0,0,0,2,'     1.0 - PIUTANG IDR',0,1,1,-1,-1,'USER','2015-12-03 21:46:47','USER','2015-12-03 21:47:08','2015-12-03 21:47:08',1,0,1,-1,-1,-1,-1)," + _
            "(11,'2.0','HUTANG IDR','',1,0,0,1,2,'     2.0 - HUTANG IDR',1,3,2,-1,-1,'USER','2015-12-03 21:47:22','USER','2015-12-03 21:47:22','2015-12-03 21:47:22',1,0,7,-1,-1,-1,-1)," + _
            "(12,'6.0','BEBAN SELISIH KURS','',1,0,0,4,2,'     6.0 - BEBAN SELISIH KURS',0,8,6,-1,-1,'USER','2015-12-03 21:47:43','USER','2015-12-03 21:47:43','2015-12-03 21:47:43',1,0,13,-1,-1,-1,-1);" + _
            "COMMIT;"

            'execute script
            script.Execute()

            Conn.Close()

            '2.update primary akun bersesuai untuk currency IDR
            ObjCurrency.CariDetail(1)
            ObjCurrency.PrimaryPayableAccount = 10
            ObjCurrency.PrimaryReceivableAccount = 11
            ObjCurrency.PrimaryUnrealizedExchangeGainLossAccount = 12
            ObjCurrency.PrimaryRealizedExchangeGainLossAccount = 12
            ObjCurrency.DefaultPrecision = 0
            ObjCurrency.SimpanPerubahan(ObjCurrency.CreatedBy)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
            ObjCurrency = Nothing
        End Try
    End Sub

    Private Sub btnCT_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCT_Save.Click
        Try
            If Me.cboCT_SchemeName.SelectedIndex = -1 Then
                Dim ij As New FrmUpdate_ClearTabel_new
                ij.vBaru = True
                ij.vPrimaryAcuan = -1

                If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.TxtCT_NamaScheme.Text = ij.TextBox1.Text
                    Me.SimpanCT_Baru()
                    Me.LoadCBO()
                End If
            Else
                If Me.chkCT_New.Checked = True Then
                    Dim ij As New FrmUpdate_ClearTabel_new
                    ij.vBaru = True
                    ij.vPrimaryAcuan = -1

                    If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Me.TxtCT_NamaScheme.Text = ij.TextBox1.Text
                        Me.SimpanCT_Baru()
                        Me.LoadCBO()
                    End If
                Else
                    Dim ij As New FrmUpdate_ClearTabel_new
                    ij.vBaru = False
                    ij.vPrimaryAcuan = Me.TxtCT_Primary.Text
                    ij.TextBox1.Text = Me.cboCT_SchemeName.Text

                    If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Me.TxtCT_NamaScheme.Text = ij.TextBox1.Text
                        Me.SimpanCT_Perubahan()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim AdaCheck As Boolean = False
        Dim DTA As New Core_Module.DSUserTableAdapters.tblmasterhakaksesTableAdapter
        Dim TBLDummy As New Core_Module.DSUser.tblmasterhakaksesDataTable

        Try
            'pastikan connection string valid
            Me.CT_ErrorReason = "Invalid connection string"
            'pastikan connection string valid
            'Me.TxtCT_ConnectionString.Text = Core_Module.GetConstring
            'DTA.Connection.ConnectionString = Me.TxtCT_ConnectionString.Text
            DTA.Fill(TBLDummy)
            Me.CT_ErrorReason = ""

            'pastikan nama target database diisikan
            Me.TxtCT_DBName.Text = Me.TxtCT_DBName.Text.Trim
            If Me.TxtCT_DBName.Text = "" Then
                MsgBox("Please specify target database", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan minimal ada 1 baris di dalam dgv
            If Me.dgvCT_DetailTable.Rows.Count = 0 Then
                MsgBox("Please load table's name first.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan minimal ada 1 baris yang di check untuk clear
            For Each rowDGV In Me.dgvCT_DetailTable.Rows
                If rowDGV.Cells(Me.ColCT_Clear.Index).Value = Boolean.TrueString Then
                    AdaCheck = True
                    Exit For
                End If
            Next
            If AdaCheck = False Then
                MsgBox("Please specify at least 1 table in database to clear.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            If ex.Message = "Unable to connect to any of the specified MySQL hosts." And Me.CT_ErrorReason = "Invalid connection string" Then
                MsgBox(Me.CT_ErrorReason, MsgBoxStyle.Critical, "Error")
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            TBLDummy.Dispose()
        End Try
    End Function

    Private Sub SimpanCT_Perubahan()
        Dim DTA As New DseUpdate_ClearTableTableAdapters.tbltrans_clear_dataTableAdapter
        Dim DTAItem As New DseUpdate_ClearTableTableAdapters.tbltrans_clear_data_itemTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim vPrimary As Double = -1

        Try
            If Me.ApakahDataBenar(True) = True Then
                'update header
                DTA.Update(Me.TxtCT_NamaScheme.Text, "USER", Now, Now, Me.TxtCT_ConnectionString.Text, Me.TxtCT_DBName.Text, Me.TxtCT_Primary.Text)

                'hapus item
                DTAItem.Delete(Me.TxtCT_Primary.Text)

                'simpan item
                For Each rowDGV In Me.dgvCT_DetailTable.Rows
                    If rowDGV.Cells(Me.ColCT_Clear.Index).Value = Boolean.TrueString Then
                        DTAItem.Insert(Me.TxtCT_Primary.Text, rowDGV.Cells(Me.ColCT_TableName.Index).Value.ToString)
                    End If
                Next

                'konfirmasi
                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanCT_Baru()
        Dim DTA As New DseUpdate_ClearTableTableAdapters.tbltrans_clear_dataTableAdapter
        Dim DTAItem As New DseUpdate_ClearTableTableAdapters.tbltrans_clear_data_itemTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim vPrimary As Double = -1

        Try
            If Me.ApakahDataBenar(True) = True Then
                'simpan header
                DTA.Insert(Me.TxtCT_NamaScheme.Text, "USER", Now, "USER", Now, Now, Me.TxtCT_ConnectionString.Text, Me.TxtCT_DBName.Text)
                vPrimary = DTA.ScalarQuery_GetPrimary(Me.TxtCT_NamaScheme.Text).Value

                'simpan item
                For Each rowDGV In Me.dgvCT_DetailTable.Rows
                    If rowDGV.Cells(Me.ColCT_Clear.Index).Value = Boolean.TrueString Then
                        DTAItem.Insert(vPrimary, rowDGV.Cells(Me.ColCT_TableName.Index).Value.ToString)
                    End If
                Next

                'konfirmasi
                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanFieldCT()
        Try
            Me.TxtCT_Primary.Text = "-1"
            Me.TxtCT_NamaScheme.Text = ""
            Me.TxtCT_DBName.Text = ""
            Me.TxtCT_ConnectionString.Text = ""
            Me.dgvCT_DetailTable.Rows.Clear()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvCT_DetailTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDGV In Me.dgvCT_DetailTable.Rows
                rowDGV.Cells(Me.ColCT_Clear.Index).Value = CBool(Boolean.TrueString)
                If rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tblmasterhakakses" Then
                    rowDGV.Cells(Me.ColCT_Clear.Index).Value = CBool(Boolean.FalseString)
                ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tbltrans_accounting_closing" Then
                    rowDGV.Cells(Me.ColCT_Clear.Index).Value = CBool(Boolean.FalseString)
                ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tbllistedsoftware" Then
                    rowDGV.Cells(Me.ColCT_Clear.Index).Value = CBool(Boolean.FalseString)
                ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tblmaster_tarif_pph" Then
                    rowDGV.Cells(Me.ColCT_Clear.Index).Value = CBool(Boolean.FalseString)
                ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tblmaster_salary_formula_parameter" Then
                    rowDGV.Cells(Me.ColCT_Clear.Index).Value = CBool(Boolean.FalseString)
                ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tbltrans_company_data" Then
                    rowDGV.Cells(Me.ColCT_Clear.Index).Value = CBool(Boolean.FalseString)
                ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tblmaster_kode_negara" Then
                    rowDGV.Cells(Me.ColCT_Clear.Index).Value = CBool(Boolean.FalseString)
                ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tblmaster_ptkp" Then
                    rowDGV.Cells(Me.ColCT_Clear.Index).Value = CBool(Boolean.FalseString)
                ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tbltrans_clear_data" Then
                    rowDGV.Cells(Me.ColCT_Clear.Index).Value = CBool(Boolean.FalseString)
                ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tbltrans_clear_data_item" Then
                    rowDGV.Cells(Me.ColCT_Clear.Index).Value = CBool(Boolean.FalseString)
                ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tbltrans_dbversion" Then
                    rowDGV.Cells(Me.ColCT_Clear.Index).Value = CBool(Boolean.FalseString)
                ElseIf rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tbltrans_email_automated" Then
                    rowDGV.Cells(Me.ColCT_Clear.Index).Value = CBool(Boolean.FalseString)
                End If
            Next
            Me.dgvCT_DetailTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvCT_DetailTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDGV In Me.dgvCT_DetailTable.Rows
                rowDGV.Cells(Me.ColCT_Clear.Index).Value = CBool(Boolean.FalseString)
            Next
            Me.dgvCT_DetailTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedTab.Name = Me.TabPage5.Name Then
                Me.LoadCBO()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCT_SchemeName_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCT_SchemeName.SelectedValueChanged
        Try
            If Me.cboCT_SchemeName.SelectedIndex = -1 Then
                Me.KosongkanFieldCT()
            Else
                Me.SalinKeField()
                Me.LoadTableName()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCT_DetailTable_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCT_DetailTable.CellEndEdit
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tblmaster_akun" Then
                If CBool(Me.dgvCT_DetailTable.Item(Me.ColCT_Clear.Index, e.RowIndex).Value) = CBool(Boolean.TrueString) Then
                    'jika hapus master akun, maka table-table di bawah ini juga harus dihapuskan
                    'table setting adj
                    'tbl_master_currency
                    'tblmaster_budget
                    For Each rowDGV In Me.dgvCT_DetailTable.Rows
                        If rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tbl_setting_adjustment" Or _
                        rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tblmaster_currency" Or _
                        rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tblmaster_budget" Or _
                        rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tbltrans_jurnal" Or _
                        rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tbltrans_jurnal_quick_add" Or _
                        rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tbltrans_recurring_journal" Then
                            rowDGV.Cells(Me.ColCT_Clear.Index).Value = Boolean.TrueString
                        End If
                    Next
                End If
            ElseIf Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tblantarauserhakakses" Or _
            Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tbltrans_setting_user" Or _
            Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tblmaster_usergroup" Or _
            Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tbluser" Then
                'table user -> jika baris user dihapuskan, maka tabel berikut juga dihapuskan
                '1. user group
                '2. antara hak akses
                '3. setting user
                '4. user group item
                If CBool(Me.dgvCT_DetailTable.Item(Me.ColCT_Clear.Index, e.RowIndex).Value) = CBool(Boolean.TrueString) Then
                    For Each rowDGV In Me.dgvCT_DetailTable.Rows
                        If rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tblantarauserhakakses" Or _
                        rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tbltrans_setting_user" Or _
                        rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tblmaster_usergroup" Or _
                        rowDGV.Cells(Me.ColCT_TableName.Index).Value = "tbluser" Then
                            rowDGV.Cells(Me.ColCT_Clear.Index).Value = Boolean.TrueString
                        End If
                    Next
                End If
            ElseIf Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tblmasterhakakses" Then
                Me.dgvCT_DetailTable.Item(Me.ColCT_Clear.Index, e.RowIndex).Value = CBool(Boolean.FalseString)
            ElseIf Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tbltrans_accounting_closing" Then
                Me.dgvCT_DetailTable.Item(Me.ColCT_Clear.Index, e.RowIndex).Value = CBool(Boolean.FalseString)
            ElseIf Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tbllistedsoftware" Then
                Me.dgvCT_DetailTable.Item(Me.ColCT_Clear.Index, e.RowIndex).Value = CBool(Boolean.FalseString)
            ElseIf Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tblmaster_tarif_pph" Then
                Me.dgvCT_DetailTable.Item(Me.ColCT_Clear.Index, e.RowIndex).Value = CBool(Boolean.FalseString)
            ElseIf Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tblmaster_salary_formula_parameter" Then
                Me.dgvCT_DetailTable.Item(Me.ColCT_Clear.Index, e.RowIndex).Value = CBool(Boolean.FalseString)
            ElseIf Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tbltrans_company_data" Then
                Me.dgvCT_DetailTable.Item(Me.ColCT_Clear.Index, e.RowIndex).Value = CBool(Boolean.FalseString)
            ElseIf Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tblmaster_kode_negara" Then
                Me.dgvCT_DetailTable.Item(Me.ColCT_Clear.Index, e.RowIndex).Value = CBool(Boolean.FalseString)
            ElseIf Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tblmaster_ptkp" Then
                Me.dgvCT_DetailTable.Item(Me.ColCT_Clear.Index, e.RowIndex).Value = CBool(Boolean.FalseString)
            ElseIf Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tbltrans_clear_data" Then
                Me.dgvCT_DetailTable.Item(Me.ColCT_Clear.Index, e.RowIndex).Value = CBool(Boolean.FalseString)
            ElseIf Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tbltrans_clear_data_item" Then
                Me.dgvCT_DetailTable.Item(Me.ColCT_Clear.Index, e.RowIndex).Value = CBool(Boolean.FalseString)
            ElseIf Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tbltrans_dbversion" Then
                Me.dgvCT_DetailTable.Item(Me.ColCT_Clear.Index, e.RowIndex).Value = CBool(Boolean.FalseString)
            ElseIf Me.dgvCT_DetailTable.Item(Me.ColCT_TableName.Index, e.RowIndex).Value = "tbltrans_email_automated" Then
                Me.dgvCT_DetailTable.Item(Me.ColCT_Clear.Index, e.RowIndex).Value = CBool(Boolean.FalseString)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InsertNumbering()

    End Sub

    Private Sub InsertPriorityLevel()

    End Sub
End Class
