Option Explicit On
Option Strict Off

Partial Public Class FrmUpdate
    Private Sub Update_v_1_6_0()

    End Sub

    Private Sub UpdateNumberingMaterialPlanning(ByVal vLblProgress As Label)
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim DTANumbering As New DseUpdate_v160TableAdapters.tblmaster_numberingTableAdapter
        Dim Jml As Nullable(Of Long)
        Dim DTAMP As New DseUpdate_v160TableAdapters.tbltrans_material_planningTableAdapter

        Try
            DTANumbering.Connection.ConnectionString = Core_Module.GetConstring
            DTAMP.Connection.ConnectionString = DTANumbering.Connection.ConnectionString

            vLblProgress.Text = "Updating"
            Jml = DTANumbering.ScalarQuery_GetJmlByDesc("Collection Letter")
            If Jml.Value = 0 Then
                'auto number
                Conn.Open()

                Jml = DTAMP.ScalarQuery_GetJml
                If Jml.HasValue = True Then
                    If Jml.Value = 0 Then
                        sql = "INSERT INTO `tblmaster_numbering` (`primary`, `deskripsi`, `tipe transaksi`, `auto number`, `reset numbering`, `created by`, `created date`, `modified by`, `modified date`, `timestamp`, `tabel rujukan`) VALUES " + vbCrLf + _
                                       "(7,'Material Planning',14,1,0,'USER', '" + Format(Now, "yyyy-MM-dd hh:mm:ss") + "','USER','" + Format(Now, "yyyy-MM-dd hh:mm:ss") + "','" + Format(Now, "yyyy-MM-dd hh:mm:ss") + "','tbltrans_material_planning');" + vbCrLf + _
                                       "COMMIT;"
                        script.Query = sql
                        'execute script
                        script.Execute()

                        sql = "INSERT INTO `tblmaster_numbering_item` (`primary transaksi`, `tipe komponen`, `format`, `value`) VALUES " + vbCrLf + _
                        "(7,3,13,'000001')," + vbCrLf + _
                        "(7,0,18,'/MP/')," + vbCrLf + _
                        "(7,1,2,'03')," + vbCrLf + _
                        "(7,0,18,'/')," + vbCrLf + _
                        "(7,2,8,'2012');" + vbCrLf + _
                        "COMMIT;"
                        script.Query = sql
                        'execute script
                        script.Execute()
                    Else
                        sql = "INSERT INTO `tblmaster_numbering` (`primary`, `deskripsi`, `tipe transaksi`, `auto number`, `reset numbering`, `created by`, `created date`, `modified by`, `modified date`, `timestamp`, `tabel rujukan`) VALUES " + vbCrLf + _
                                       "(7,'Material Planning',14,0,0,'USER', '" + Format(Now, "yyyy-MM-dd hh:mm:ss") + "','USER','" + Format(Now, "yyyy-MM-dd hh:mm:ss") + "','" + Format(Now, "yyyy-MM-dd hh:mm:ss") + "','tbltrans_material_planning');" + vbCrLf + _
                                       "COMMIT;"
                        script.Query = sql
                        'execute script
                        script.Execute()
                    End If
                End If



                Conn.Close()

                'manual numbering
            End If
            vLblProgress.Text = "done"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTANumbering.Dispose()
            DTAMP.Dispose()
        End Try
    End Sub

    Private Sub UpdateNumberingCollectionLetter(ByVal vLblProgress As Label)
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim DTANumbering As New DseUpdate_v160TableAdapters.tblmaster_numberingTableAdapter
        Dim Jml As Nullable(Of Long)
        Dim DTATagihan As New DseUpdate_v160TableAdapters.tbltrans_surattagihanTableAdapter

        Try
            DTANumbering.Connection.ConnectionString = Core_Module.GetConstring
            DTATagihan.Connection.ConnectionString = DTANumbering.Connection.ConnectionString

            vLblProgress.Text = "Updating"
            Jml = DTANumbering.ScalarQuery_GetJmlByDesc("Collection Letter")
            If Jml.Value = 0 Then
                'auto number
                Conn.Open()

                Jml = DTATagihan.ScalarQuery_GetJml
                If Jml.HasValue = True Then
                    If Jml.Value = 0 Then
                        sql = "INSERT INTO `tblmaster_numbering` (`primary`, `deskripsi`, `tipe transaksi`, `auto number`, `reset numbering`, `created by`, `created date`, `modified by`, `modified date`, `timestamp`, `tabel rujukan`) VALUES " + vbCrLf + _
                                       "(6,'Collection Letter',13,1,0,'USER', '" + Format(Now, "yyyy-MM-dd hh:mm:ss") + "','USER','" + Format(Now, "yyyy-MM-dd hh:mm:ss") + "','" + Format(Now, "yyyy-MM-dd hh:mm:ss") + "','tbltrans_surattagihan');" + vbCrLf + _
                                       "COMMIT;"
                        script.Query = sql
                        'execute script
                        script.Execute()

                        sql = "INSERT INTO `tblmaster_numbering_item` (`primary transaksi`, `tipe komponen`, `format`, `value`) VALUES " + vbCrLf + _
                        "(6,3,13,'000001')," + vbCrLf + _
                        "(6,0,18,'/COL/')," + vbCrLf + _
                        "(6,1,2,'03')," + vbCrLf + _
                        "(6,0,18,'/')," + vbCrLf + _
                        "(6,2,8,'2012');" + vbCrLf + _
                        "COMMIT;"
                        script.Query = sql
                        'execute script
                        script.Execute()
                    Else
                        sql = "INSERT INTO `tblmaster_numbering` (`primary`, `deskripsi`, `tipe transaksi`, `auto number`, `reset numbering`, `created by`, `created date`, `modified by`, `modified date`, `timestamp`, `tabel rujukan`) VALUES " + vbCrLf + _
                                       "(6,'Collection Letter',13,0,0,'USER', '" + Format(Now, "yyyy-MM-dd hh:mm:ss") + "','USER','" + Format(Now, "yyyy-MM-dd hh:mm:ss") + "','" + Format(Now, "yyyy-MM-dd hh:mm:ss") + "','tbltrans_surattagihan');" + vbCrLf + _
                                       "COMMIT;"
                        script.Query = sql
                        'execute script
                        script.Execute()
                    End If
                End If



                Conn.Close()

                'manual numbering
            End If
            vLblProgress.Text = "done"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTANumbering.Dispose()
            DTATagihan.Dispose()
        End Try
    End Sub

    Private Sub ExecuteSqlScript(ByVal sql As String, ByVal vShowMsg As Boolean)
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)

        Try
            'alter database: tambahkan kolom batch number, expiry, primary location
            Conn.Open()
            script.Query = sql
            'execute script
            script.Execute()
            Conn.Close()
        Catch ex As Exception
            If vShowMsg = True Then
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            Conn.Dispose()
            script = Nothing
        End Try
    End Sub

    'fungsi ini membuat stock card bisa menyimpan data batch number, expiry dan location
    Private Sub UpdateStockCard()
        Dim sql As String = ""

        Try
            If MsgBox("Update table structure?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.lblPrgStockCard.Text = "Updating tabel structure: tbltransaksi_stock_card"
                'alter database: drop kolom batch number, expiry, primary location
                sql = "alter table tbltransaksi_stock_card drop column `Expiry`;"
                Me.ExecuteSqlScript(sql, False)
                sql = "alter table tbltransaksi_stock_card drop column `Batch Number`;"
                Me.ExecuteSqlScript(sql, False)
                sql = "alter table tbltransaksi_stock_card drop column `primary location`;"
                Me.ExecuteSqlScript(sql, False)
                'alter database: add kolom batch number, expiry, primary location
                sql = "alter table tbltransaksi_stock_card add column `Batch Number` varchar(250) not null default '';"
                Me.ExecuteSqlScript(sql, False)
                sql = "alter table tbltransaksi_stock_card add column `Expiry` date not null default '1950-01-01';"
                Me.ExecuteSqlScript(sql, False)
                sql = "alter table tbltransaksi_stock_card add column `primary location` double(15,0) not null default -1;"
                Me.ExecuteSqlScript(sql, False)
            End If

            'update data
            '0. hapus data, kecuali yang berasal dari stock taking
            Me.HapusStockCard_v160()
            '1. stock in out
            Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut
            ObjStockInOut.UpdateStockCard_TransaksiStockInOut(Me.prgStockCard, Me.lblPrgStockCard)
            ObjStockInOut = Nothing
            '2. bundling
            Dim ObjStockBundling As New Inventory_Module.ClsTransaksiStockBundling
            ObjStockBundling.UpdateStockCard_TransaksiStockBundling(Me.prgStockCard, Me.lblPrgStockCard)
            ObjStockBundling = Nothing

            Me.lblPrgStockCard.Text = "done"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'fungsi ini menghapuskan semua stock card kecuali yang berasal dari stock taking
    Public Sub HapusStockCard_v160()
        Dim DTA As New DseUpdate_v160TableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim TipeTransaksiStockTaking As Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal = _
        Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Taking

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(TipeTransaksiStockTaking)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub PindahkanSalesPrice_v160()
        Dim DTALama As New DseUpdate_v160TableAdapters.tblmaster_product_pricingTableAdapter
        Dim TBLLama As New DseUpdate_v160.tblmaster_product_pricingDataTable
        Dim rowLama As DseUpdate_v160.tblmaster_product_pricingRow
        Dim DTABaru As New DseUpdate_v160TableAdapters.tblmaster_product_pricelistTableAdapter

        Try
            DTALama.Connection.ConnectionString = Core_Module.GetConstring
            DTABaru.Connection.ConnectionString = DTALama.Connection.ConnectionString
            DTALama.Fill(TBLLama)
            DTABaru.DeleteQuery()

            Me.prgUpdateSalesPrice.Value = 0
            Me.prgUpdateSalesPrice.Maximum = TBLLama.Rows.Count

            For Each rowLama In TBLLama.Rows
                Me.prgUpdateSalesPrice.Value += 1
                Me.LblProgres_PindahSalesPrice.Text = "Transfering data " + Me.prgUpdateSalesPrice.Value.ToString + " of " + Me.prgUpdateSalesPrice.Maximum.ToString + " - " + rowLama.Product_Code
                With rowLama
                    DTABaru.Insert(.Primary_Product, .Unit_Price, .Discount, 0, .Primary_Currency, .Primary_Unit, "")
                End With
            Next

            Me.LblProgres_PindahSalesPrice.Text = "done"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTALama.Dispose()
            TBLLama.Dispose()
            DTABaru.Dispose()
        End Try
    End Sub

    Private Sub btnPindahkanSalesPrice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPindahkanSalesPrice.Click
        Try
            Me.PindahkanSalesPrice_v160()
            MsgBox("Done")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateIdentifierSO()
        Dim ListPrimary As New List(Of Double)
        Dim DTAItemSO As New DseUpdate_v160TableAdapters.tbltrans_so_itemTableAdapter
        Dim TBLItemSO As New DseUpdate_v160.tbltrans_so_itemDataTable
        Dim rowItemSO As DseUpdate_v160.tbltrans_so_itemRow
        Dim Urutan As Integer = 1
        Dim DTAItemDelivery As New DseUpdate_v160TableAdapters.tbltrans_so_itemdeliveryTableAdapter

        Try
            'kosongkan item delivery
            DTAItemDelivery.Delete()

            'load semua item so
            DTAItemSO.Fill(TBLItemSO)

            'masukkan primary transaksi ke dalam list
            Me.LblIdentifierSO.Text = "Registering SO transaction"
            Me.prgIdentifierSO.Value = 0
            Me.prgIdentifierSO.Maximum = TBLItemSO.Rows.Count
            For Each rowItemSO In TBLItemSO.Rows
                If ListPrimary.Contains(rowItemSO.Primary_Transaksi) = False Then
                    ListPrimary.Add(rowItemSO.Primary_Transaksi)
                End If
                Me.prgIdentifierSO.Value += 1
            Next

            'update identifier menurut list primary transaksi
            Me.prgIdentifierSO.Value = 0
            Me.prgIdentifierSO.Maximum = ListPrimary.Count
            Me.LblIdentifierSO.Text = "Updating identifier"
            For Each PrimarySO As Double In ListPrimary
                Me.prgIdentifierSO.Value += 1
                TBLItemSO.DefaultView.RowFilter = "[Primary Transaksi] = " + PrimarySO.ToString
                TBLItemSO.DefaultView.Sort = "Primary"
                Urutan = 1
                For Each DV As DataRowView In TBLItemSO.DefaultView
                    Me.LblIdentifierSO.Text = "Updating identifier - data " + Me.prgIdentifierSO.Value.ToString + " of " + Me.prgIdentifierSO.Maximum.ToString
                    rowItemSO = DV.Row
                    DTAItemSO.UpdateQuery1(rowItemSO.Primary_Transaksi.ToString + "-" + Urutan.ToString, rowItemSO.Primary)

                    'delivery date
                    If rowItemSO.Delivery_Date_Cont = True Then
                        DTAItemDelivery.Insert(rowItemSO.Primary_Transaksi, rowItemSO.Primary_Transaksi.ToString + "-" + Urutan.ToString, rowItemSO.Qty, rowItemSO.Delivery_Date, "", 1)
                    End If
                    Urutan += 1
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
            DTAItemSO.Dispose()
            TBLItemSO.Dispose()
        End Try
    End Sub

    Private Sub btnUpdate_IdentifierSO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_IdentifierSO.Click
        Try
            'Me.UpdateIdentifierSO()
            'Me.LblIdentifierSO.Text = "Done"
            'MsgBox("Done")
            Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker9.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker9_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker9.DoWork
        Me.UpdateIdentifierSO()
    End Sub

    Private Sub BackgroundWorker9_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker9.RunWorkerCompleted
        Me.LblIdentifierSO.Text = "Done"
        MsgBox("Done")
    End Sub

    Private Sub UpdateIdentifierPO()
        Dim ListPrimary As New List(Of Double)
        Dim DTAItemPO As New DseUpdate_v160TableAdapters.tbltrans_po_itemTableAdapter
        Dim TBLItemPO As New DseUpdate_v160.tbltrans_PO_itemDataTable
        Dim rowItemPO As DseUpdate_v160.tbltrans_PO_itemRow
        Dim Urutan As Integer = 1
        Dim DTAItemDelivery As New DseUpdate_v160TableAdapters.tbltrans_PO_itemdeliveryTableAdapter
        Dim Sql As String = ""

        Try
            'update tabel 
            'alter database: drop kolom identifier
            If MsgBox("Update tabel structure?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.lblProgressPO.Text = "Updating tabel structure: tbltrans_po_item"
                Sql = "alter table tbltrans_po_item drop column `identifier`;"
                Me.ExecuteSqlScript(Sql, False)
                Sql = "alter table tbltrans_po_item add column `identifier` varchar(250) not null default '';"
                Me.ExecuteSqlScript(Sql, False)

                'tambahkan tabel item delivery
                Me.lblProgressPO.Text = "Updating tabel structure: tbltrans_po_itemdelivery"
                Sql = "DROP TABLE IF EXISTS `tbltrans_po_itemdelivery`;" + _
                "CREATE TABLE `tbltrans_po_itemdelivery` (" + _
                "`primary` double(15,0) NOT NULL auto_increment," + _
                "`primary transaksi` double(15,0) NOT NULL," + _
                "`identifier` varchar(250) NOT NULL," + _
                "`delivery qty` decimal(21,6) NOT NULL default '0.000000'," + _
                "`delivery date` date NOT NULL," + _
                "`remarks` varchar(250) NOT NULL," + _
                "`delivery priority` double(15,0) NOT NULL default '-1'," + _
                "`delay` int(3) NOT NULL default '0'," + _
                "PRIMARY KEY  (`primary`)," + _
                "KEY `primary transaksi` (`primary transaksi`)," + _
                "CONSTRAINT `tbltrans_po_itemdelivery_fk` FOREIGN KEY (`primary transaksi`) REFERENCES `tbltrans_po` (`Primary`) ON DELETE CASCADE ON UPDATE NO ACTION" + _
                ") ENGINE=InnoDB DEFAULT CHARSET=latin1;"
                Me.ExecuteSqlScript(Sql, False)
            End If

            'koPOngkan item delivery
            DTAItemDelivery.Delete()

            'load semua item PO
            DTAItemPO.Fill(TBLItemPO)

            'masukkan primary transaksi ke dalam list
            Me.lblProgressPO.Text = "Registering PO transaction"
            Me.prgIdentifierPO.Value = 0
            Me.prgIdentifierPO.Maximum = TBLItemPO.Rows.Count
            For Each rowItemPO In TBLItemPO.Rows
                If ListPrimary.Contains(rowItemPO.Primary_Transaksi) = False Then
                    ListPrimary.Add(rowItemPO.Primary_Transaksi)
                End If
                Me.prgIdentifierPO.Value += 1
            Next

            'update identifier menurut list primary transaksi
            Me.prgIdentifierPO.Value = 0
            Me.prgIdentifierPO.Maximum = ListPrimary.Count
            Me.lblProgressPO.Text = "Updating identifier & item delivery"
            For Each PrimaryPO As Double In ListPrimary
                Me.prgIdentifierPO.Value += 1
                TBLItemPO.DefaultView.RowFilter = "[Primary Transaksi] = " + PrimaryPO.ToString
                TBLItemPO.DefaultView.Sort = "Primary"
                Urutan = 1
                For Each DV As DataRowView In TBLItemPO.DefaultView
                    Me.lblProgressPO.Text = "Updating identifier - data " + Me.prgIdentifierPO.Value.ToString + " of " + Me.prgIdentifierPO.Maximum.ToString
                    rowItemPO = DV.Row
                    DTAItemPO.UpdateQuery(rowItemPO.Primary_Transaksi.ToString + "-" + Urutan.ToString, rowItemPO.Primary)
                    If rowItemPO.Delivery_Date_Cont = True Then
                        DTAItemDelivery.Insert(rowItemPO.Primary_Transaksi, rowItemPO.Primary_Transaksi.ToString + "-" + Urutan.ToString, rowItemPO.Qty, rowItemPO.Delivery_Date, "", 1)
                    End If
                    Urutan += 1
                Next
            Next

            Me.lblProgressPO.Text = "Done"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
            DTAItemPO.Dispose()
            TBLItemPO.Dispose()
        End Try
    End Sub

    Private Sub btnUpdatev160_TBLMasterProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdatev160_TBLMasterProduct.Click
        Dim sql As String = ""

        Try
            MsgBox("tolong dicek codingnya", MsgBoxStyle.Information, "informasi")
            Exit Sub
            Me.lblProgressPO.Text = "Updating tabel structure: tblmaster_product"
            sql = "alter table tblmaster_product drop column `procurement type`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `procurement type` int(3) not null default '0';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `procurement strategy`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `procurement strategy` int(3) not null default '0';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `supplier selection mode`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `supplier selection mode` int(3) not null default '0';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `primary prefered supplier`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `primary prefered supplier` double(15,0) not null default '-1';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `recomended order qty type`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `recomended order qty type` int(3) not null default '0';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `optimum stock`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `optimum stock` decimal(21,6) NOT NULL default '0.000000';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `maximum stock`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `maximum stock` decimal(21,6) NOT NULL default '0.000000';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `primary prefered purchase unit`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `primary prefered purchase unit` double(15,0) not null default '-1';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `optimum order qty`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `optimum order qty` decimal(21,6) NOT NULL default '0.000000';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `maximum order qty`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `maximum order qty` decimal(21,6) NOT NULL default '0.000000';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `avg montly demand`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `avg montly demand` decimal(21,6) NOT NULL default '0.000000';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `start period`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `start period` date NOT NULL default '1950-01-01';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `end period`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `end period` date NOT NULL default '1950-01-01';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `include sales`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `include sales` int(3) not null default '1';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `include bundling`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `include bundling` int(3) not null default '1';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `include adjustment`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `include adjustment` int(3) not null default '1';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `safety stock strategy`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `safety stock strategy` int(3) not null default '0';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `safety stock days`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `safety stock days` int(3) not null default '0';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `safety stock qty`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `safety stock qty` decimal(21,6) NOT NULL default '0.000000';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `lead time strategy`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `lead time strategy` int(3) not null default '0';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `lead time days`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `lead time days` int(3) not null default '0';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `lead time qty`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `lead time qty` decimal(21,6) NOT NULL default '0.000000';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `exclude on material planning`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `exclude on material planning` int(3) not null default '0';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `explain avg qty`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `explain avg qty` varchar(1000) NOT NULL default '';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `explain safety stock qty`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `explain safety stock qty` varchar(250) NOT NULL default '';"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product drop column `explain lead time qty`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "alter table tblmaster_product add column `explain lead time qty` varchar(250) NOT NULL default '';"
            Me.ExecuteSqlScript(sql, False)
            Me.lblProgressPO.Text = "done"
            MsgBox("Update finish")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnTambahTabelBaru_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdatev160_TambahTabelBaru.Click
        Dim sql As String = ""

        Try
            MsgBox("tolong dicek codingnya, cek apakah sudah lengkap semua tabelnya.", MsgBoxStyle.Information, "informasi")
            Exit Sub

            MsgBox("Set database version jadi 1.6.0")
            '1. tabel priority level + tambahkan baris untuk reguler priority, high priority, dan low priority
            Me.lblUpdatev160_TambahTabelBaru.Text = "Updating tblmaster_priority_level"
            sql = "DROP TABLE IF EXISTS `tblmaster_priority_level`;" + _
            "CREATE TABLE `tblmaster_priority_level` (" + _
            "`primary` double(15,0) NOT NULL auto_increment," + _
            "`urutan` int(3) NOT NULL default '1'," + _
            "`code` varchar(250) NOT NULL," + _
            "`desc` varchar(250) NOT NULL," + _
            "`default` int(3) NOT NULL default '0'," + _
            "`created by` varchar(250) NOT NULL," + _
            "`created date` datetime NOT NULL," + _
            "`modified by` varchar(250) NOT NULL," + _
            "`modified date` datetime NOT NULL," + _
            "`timestamp` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP," + _
            "PRIMARY KEY  (`primary`)" + _
            ") ENGINE=InnoDB DEFAULT CHARSET=latin1;"
            Me.ExecuteSqlScript(sql, False)

            '2. tbltrans_material_planning, tbltrans_material_planning_detailotorisasipo, tbltrans_material_planning_item,
            '   tbltrans_material_planning_itemdo, tbltrans_material_planning_itemgit, tbltrans_material_planning_itempob,
            '   tbltrans_material_planning_itemsob
            Me.lblUpdatev160_TambahTabelBaru.Text = "Updating tbltrans_material_planning"
            sql = "DROP TABLE IF EXISTS `tbltrans_material_planning`;" + _
                       "CREATE TABLE `tbltrans_material_planning` (" + _
                       "`primary` double(15,0) NOT NULL auto_increment," + _
                       "`doc no` varchar(250) NOT NULL," + _
                       "`urutan` double(15,0) NOT NULL," + _
                       "`date` date NOT NULL," + _
                       "`remarks` varchar(250) NOT NULL," + _
                       "`created by` varchar(250) NOT NULL," + _
                       "`created date` datetime NOT NULL," + _
                       "`modified by` varchar(250) NOT NULL," + _
                       "`modified date` datetime NOT NULL," + _
                       "`timestamp` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP," + _
                       "`ref date` date NOT NULL default '1950-01-01'," + _
                       "`days per month` double(15,0) NOT NULL default '0'," + _
                       "PRIMARY KEY  (`primary`)" + _
                       ") ENGINE=InnoDB DEFAULT CHARSET=latin1;"
            Me.ExecuteSqlScript(sql, False)

            Me.lblUpdatev160_TambahTabelBaru.Text = "Updating tbltrans_material_planning_detailotorisasipo"
            sql = "DROP TABLE IF EXISTS `tbltrans_material_planning_detailotorisasipo`;" + _
            "CREATE TABLE `tbltrans_material_planning_detailotorisasipo` (" + _
            "`primary` double(15,0) NOT NULL auto_increment," + _
            "`primary po` double(15,0) NOT NULL default '-1'," + _
            "`include in calc` int(3) NOT NULL default '0'," + _
            "PRIMARY KEY  (`primary`)" + _
            ") ENGINE=InnoDB DEFAULT CHARSET=latin1;"
            Me.ExecuteSqlScript(sql, False)

            Me.lblUpdatev160_TambahTabelBaru.Text = "Updating tbltrans_material_planning_item"
            sql = "DROP TABLE IF EXISTS `tbltrans_material_planning_item`;" + _
            "CREATE TABLE `tbltrans_material_planning_item` (" + _
            "`primary` double(15,0) NOT NULL auto_increment," + _
            "`primary transaksi` double(15,0) NOT NULL," + _
            "`primary code no` double(15,0) NOT NULL," + _
            "`primary unit` double(15,0) NOT NULL," + _
            "`qty on hand` decimal(21,6) NOT NULL default '0.000000'," + _
            "`qty pob` decimal(21,6) NOT NULL default '0.000000'," + _
            "`qty git` decimal(21,6) NOT NULL default '0.000000'," + _
            "`qty inquiry` decimal(21,6) NOT NULL default '0.000000'," + _
            "`qty pending do` decimal(21,6) NOT NULL default '0.000000'," + _
            "`qty sob` decimal(21,6) NOT NULL default '0.000000'," + _
            "`qty safety stock` decimal(21,6) NOT NULL default '0.000000'," + _
            "`qty lead time` decimal(21,6) NOT NULL default '0.000000'," + _
            "`qty required` decimal(21,6) NOT NULL default '0.000000'," + _
            "`default prec unit` int(3) NOT NULL default '0'," + _
            "`qty order maksimum` decimal(21,6) NOT NULL default '0.000000'," + _
            "`qty order minimum` decimal(21,6) NOT NULL default '0.000000'," + _
            "`qty order optimum` decimal(21,6) NOT NULL default '0.000000'," + _
            "`qty max` decimal(21,6) NOT NULL default '0.000000'," + _
            "`qty min` decimal(21,6) NOT NULL default '0.000000'," + _
            "`qty optimum` decimal(21,6) NOT NULL default '0.000000'," + _
            "`qty order` decimal(21,6) NOT NULL default '0.000000'," + _
            "`primary order unit` double(15,0) NOT NULL default '0'," + _
            "`def prec order unit` int(3) NOT NULL default '0'," + _
            "`order point` date NOT NULL," + _
            "`procurement type` int(3) NOT NULL," + _
            "`primary prefered supplier` double(15,0) NOT NULL default '-1'," + _
            "`primary currency` double(15,0) NOT NULL default '-1'," + _
             "`unit price` decimal(21,6) NOT NULL default '0.000000'," + _
            "`discount` decimal(21,6) NOT NULL default '0.000000'," + _
            "`nett price` decimal(21,6) NOT NULL default '0.000000'," + _
            "`sub total` decimal(21,6) NOT NULL default '0.000000'," + _
            "`def prec currency` decimal(15,0) NOT NULL default '-1'," + _
            "`order strategy` int(3) NOT NULL default '0'," + _
            "`next order point` date NOT NULL," + _
            "`qty_order_multiunit` decimal(21,6) NOT NULL default '0.000000'," + _
            "`lead time days` decimal(21,6) NOT NULL default '0.000000'," + _
            "`order point delay` double(15,0) NOT NULL default '0'," + _
            "`next order point delay` double(15,0) NOT NULL default '0'," + _
            "`qty avg montly demand` decimal(21,6) NOT NULL default '0.000000'," + _
            "`supplier selection mode` int(3) NOT NULL default '0'," + _
            "`ExplainAvgQty` varchar(1000) NOT NULL," + _
            "`Explain Safety Stock` varchar(250) NOT NULL," + _
            "`Explain Lead Time` varchar(250) NOT NULL," + _
            "`SafetyStockStrategy` int(3) NOT NULL default '0'," + _
            "`Lead Time Strategy` int(3) NOT NULL default '3'," + _
            "`modify` int(3) NOT NULL default '0'," + _
            "`procurement strategy` int(3) NOT NULL default '0'," + _
            "PRIMARY KEY  (`primary`)," + _
            "KEY `primary transaksi` (`primary transaksi`)," + _
            "CONSTRAINT `tbltrans_material_planning_item_fk` FOREIGN KEY (`primary transaksi`) REFERENCES `tbltrans_material_planning` (`primary`) ON DELETE CASCADE ON UPDATE NO ACTION" + _
            ") ENGINE=InnoDB DEFAULT CHARSET=latin1;"
            Me.ExecuteSqlScript(sql, False)

            Me.lblUpdatev160_TambahTabelBaru.Text = "Updating tbltrans_material_planning_itemdo"
            sql = "DROP TABLE IF EXISTS `tbltrans_material_planning_itemdo`;" + _
            "CREATE TABLE `tbltrans_material_planning_itemdo` (" + _
            "`primary` double(15,0) NOT NULL auto_increment," + _
            "`primary transaksi` double(15,0) NOT NULL default '-1'," + _
            "`primary product` double(15,0) NOT NULL default '-1'," + _
            "`qty` decimal(21,6) NOT NULL default '0.000000'," + _
            "`primary do rec` double(15,0) NOT NULL default '-1'," + _
            "`primary customer rec` double(15,0) NOT NULL default '-1'," + _
            "`do date rec` date NOT NULL," + _
            "`do no rec` varchar(250) NOT NULL," + _
            "PRIMARY KEY  (`primary`)," + _
            "KEY `primary transaksi` (`primary transaksi`)," + _
            "CONSTRAINT `tbltrans_material_planning_itemdo_fk` FOREIGN KEY (`primary transaksi`) REFERENCES `tbltrans_material_planning` (`primary`) ON DELETE CASCADE ON UPDATE NO ACTION" + _
            ") ENGINE=InnoDB DEFAULT CHARSET=latin1;"
            Me.ExecuteSqlScript(sql, False)

            Me.lblUpdatev160_TambahTabelBaru.Text = "Updating tbltrans_material_planning_itemgit"
            sql = "DROP TABLE IF EXISTS `tbltrans_material_planning_itemgit`;" + _
            "CREATE TABLE `tbltrans_material_planning_itemgit` (" + _
            "`primary` double(15,0) NOT NULL auto_increment," + _
            "`primary transaksi` double(15,0) NOT NULL default '-1'," + _
            "`primary product` double(15,0) NOT NULL default '-1'," + _
            "`qty` decimal(21,6) NOT NULL default '0.000000'," + _
            "`primary packing list rec` double(15,0) NOT NULL default '-1'," + _
            "`primary suppier rec` double(15,0) NOT NULL default '-1'," + _
            "`packing list date rec` date NOT NULL," + _
            "`packing list no rec` varchar(250) NOT NULL," + _
            "`supplier packing list no rec` varchar(250) NOT NULL," + _
            "PRIMARY KEY  (`primary`)," + _
            "KEY `primary transaksi` (`primary transaksi`)," + _
            "CONSTRAINT `tbltrans_material_planning_itemgit_fk` FOREIGN KEY (`primary transaksi`) REFERENCES `tbltrans_material_planning` (`primary`) ON DELETE CASCADE ON UPDATE NO ACTION" + _
            ") ENGINE=InnoDB DEFAULT CHARSET=latin1;"
            Me.ExecuteSqlScript(sql, False)

            Me.lblUpdatev160_TambahTabelBaru.Text = "Updating tbltrans_material_planning_itempob"
            sql = "DROP TABLE IF EXISTS `tbltrans_material_planning_itempob`;" + _
            "CREATE TABLE `tbltrans_material_planning_itempob` (" + _
            "`primary` double(15,0) NOT NULL auto_increment," + _
            "`primary transaksi` double(15,0) NOT NULL default '-1'," + _
            "`primary product` double(15,0) NOT NULL default '-1'," + _
            "`qty` decimal(21,6) NOT NULL default '0.000000'," + _
            "`primary po rec` double(15,0) NOT NULL default '-1'," + _
            "`primary supplier rec` double(15,0) NOT NULL default '-1'," + _
            "`po date rec` date NOT NULL," + _
            "`po no rec` varchar(250) NOT NULL," + _
            "PRIMARY KEY  (`primary`)," + _
            "KEY `primary transaksi` (`primary transaksi`)," + _
            "CONSTRAINT `tbltrans_material_planning_itempo_fk` FOREIGN KEY (`primary transaksi`) REFERENCES `tbltrans_material_planning` (`primary`) ON DELETE CASCADE ON UPDATE NO ACTION" + _
            ") ENGINE=InnoDB DEFAULT CHARSET=latin1;"
            Me.ExecuteSqlScript(sql, False)

            Me.lblUpdatev160_TambahTabelBaru.Text = "Updating tbltrans_material_planning_itemsob"
            sql = "DROP TABLE IF EXISTS `tbltrans_material_planning_itemsob`;" + _
            "CREATE TABLE `tbltrans_material_planning_itemsob` (" + _
            "`primary` double(15,0) NOT NULL auto_increment," + _
            "`primary transaksi` double(15,0) NOT NULL default '-1'," + _
            "`primary product` double(15,0) NOT NULL default '-1'," + _
            "`qty` decimal(21,6) NOT NULL default '0.000000'," + _
            "`primary so rec` double(15,0) NOT NULL default '-1'," + _
            "`primary customer rec` double(15,0) NOT NULL default '-1'," + _
            "`so date rec` date NOT NULL," + _
            "`so no rec` varchar(250) NOT NULL," + _
            "PRIMARY KEY  (`primary`)," + _
            "KEY `primary transaksi` (`primary transaksi`)," + _
            "CONSTRAINT `tbltrans_material_planning_itemsob_fk` FOREIGN KEY (`primary transaksi`) REFERENCES `tbltrans_material_planning` (`primary`) ON DELETE CASCADE ON UPDATE NO ACTION" + _
            ") ENGINE=InnoDB DEFAULT CHARSET=latin1;"
            Me.ExecuteSqlScript(sql, False)

            MsgBox("belum untuk: tbltrans_shipping_forecast")
            'tabel tbltrans_shipping_forecast + semua anaknya

            Me.lblUpdatev160_TambahTabelBaru.Text = "done"
            MsgBox("Finish")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker10_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker10.DoWork
        Me.UpdateIdentifierPO()
    End Sub

    Private Sub btnUpdateIdentifierPO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateIdentifierPO.Click
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker10.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker10_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker10.RunWorkerCompleted
        MsgBox("Finish")
    End Sub

    Private Sub btnUpdataHR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdataHR.Click
        Dim sql As String = ""

        Try
            '1. tambahkan tabel sejumlah yang tepat
            '2. tambahkan view
            '3. untuk tabel2 salary parameter, kode negara, ptkp , data tarif pph 21, sesuai standar awal
            '   - (iniled) tblmaster_salary_formula_parameter 
            '   - (iniled) tblmaster_kode_negara
            '   - (iniled) tblmaster_ptkp
            '   - (inilied) tblmaster_tarif_pph

            Me.lblUpdatev160_HR.Text = "Updating tabel for HR parameter"
            sql = "DROP TABLE IF EXISTS `tblmaster_salary_formula_parameter`;"
            Me.ExecuteSqlScript(sql, False)

            sql = "CREATE TABLE `tblmaster_salary_formula_parameter` (" + _
              "`Primary` double(18,0) NOT NULL auto_increment," + _
              "`Code` varchar(250) NOT NULL," + _
  "`Description` varchar(250) NOT NULL," + _
  "`Remarks` varchar(250) NOT NULL," + _
  "`Unit` varchar(250) NOT NULL," + _
  "`Tipe` int(11) NOT NULL default '0'," + _
  "`Is Variable` int(11) NOT NULL default '0'," + _
  "`Is Constraint` int(11) NOT NULL default '0'," + _
  "PRIMARY KEY  (`Primary`)," + _
  "UNIQUE KEY `Primary_2` (`Primary`)" + _
") ENGINE=InnoDB DEFAULT CHARSET=latin1;"
            Me.ExecuteSqlScript(sql, False)

            sql = "INSERT INTO `tblmaster_salary_formula_parameter` (`Primary`, `Code`, `Description`, `Remarks`, `Unit`, `Tipe`, `Is Variable`, `Is Constraint`) VALUES" + _
              "(1,'Hari','Hari kerja: Monday, tuesday, ect','','day of week',0,0,1)," + _
              "(2,'Sch In','Jadwal Masuk Kerja','','Time',0,0,1)," + _
              "(3,'Sch Out','Jadwal masuk pulang , format HH:mm','','Time',0,0,1)," + _
              "(4,'Actual In','Aktual masuk kerja sesuai absensi','','Time',0,0,1)," + _
              "(5,'Actual Out','Aktual pulang kerja sesuai absensi','','Time',0,0,1)," + _
              "(6,'Late Duration','Durasi telat','','Duration',0,1,1)," + _
              "(7,'Early In Duration','Durasi masuk lebih cepat dari jadwal','','Duration',0,1,1)," + _
              "(8,'Early Out Duration','Durasi pulang lebih cepat dari jadwal','','Duration',0,1,1)," + _
              "(9,'Late OT Duration','Durasi lembur setelah jam kerja',' ','Duration',0,1,1)," + _
              "(10,'Early OT Duration','Durasi lembur sebelum jam kerja',' ','Duration',0,1,1)," + _
              "(11,'Late','nilai yes jika hari itu telat',' ','yes/no',0,0,1)," + _
              "(12,'Early In','nilai yes jika hari itu masuk lebih cepat dari jadwal',' ','yes/no',0,0,1)," + _
              "(13,'Early Out','nilai yes jika hari itu pulang lebih cepat dari jadwal',' ','yes/no',0,0,1)," + _
              "(14,'Late OT','nilai yes jika hari itu ada lembur setelah masuk',' ','yes/no',0,0,1)," + _
              "(15,'Early OT','nilai yes jika hari itu ada lembur sebelum masuk',' ','yes/no',0,0,1)," + _
              "(16,'Schedule Work Duration','Total Jam kerja menurut waktu',' ','Duration',0,1,1)," + _
              "(17,'Actual Work Duration','Total jam masuk kerja',' ','Duration',0,1,1)," + _
              "(18,'Is Present','Masuk kerja',' ','0/1/0.5',0,0,1)," + _
              "(19,'Is Sick','Sakit',' ','0/1/0.5',0,0,1)," + _
              "(20,'Is Leave','Cuti',' ','0/1/0.5',0,0,1)," + _
      "(21,'Is Alpha','Alfa',' ','0/1/0.5',0,0,1)," + _
      "(22,'Is Permit','Permit',' ','0/1/0.5',0,0,1)," + _
      "(23,'Sch Break Out','Jadwal mulai istirahat',' ','Time',0,0,1)," + _
      "(24,'Sch Break In','Jadwal selesai istirahat',' ','Time',0,0,1)," + _
      "(25,'Actual Break Out','Aktual Mulai Istirahat',' ','Time',0,0,1)," + _
      "(26,'Actual Break In','Aktual selesai Istirahat',' ','Time',0,0,1)," + _
      "(27,'Actual Break Time','Durasi istirahat aktual',' ','Duration',0,1,1)," + _
      "(28,'Late Break In Duration','Durasi telat selesai istirahat',' ','Duration',0,1,1)," + _
      "(31,'Early Break Out Duration','Durasi keluar istirahat lebih cepat dari jadwal',' ','Duration',0,1,1)," + _
      "(32,'Weekday OT','nilainya yes = jika ada lembur dihari kerja','','yes/no',0,0,1)," + _
      "(34,'Holiday OT','nilainya yes = jika ada lembur di holiday',' ','yes/no',0,0,1)," + _
      "(35,'Weekday','nilainya yes jika masuk di hari kerja','','yes/no',0,0,1)," + _
      "(37,'Holiday','nilainya no jika masuk di holiday',' ','yes/no',0,0,1)," + _
      "(40,'Total Late Duration Monthly','Total durasi telat selama 1 bulan',' ','Duration',1,1,1)," + _
      "(41,'Total Early In Duration Monthly','Total durasi early in selama 1 bulan',' ','Duration',1,1,1)," + _
      "(42,'Total Early Out Duration Monthly','Total durasi early out selama 1 bulan',' ','Duration',1,1,1)," + _
      "(43,'Total Late Overtime Duration Monthly','Total durasi lembur akhir selama 1 bulan',' ','Duration',1,1,1)," + _
      "(44,'Total Early Overtime Duration Monthly','Total durasi lembur awal selama 1 bulan',' ','Duration',1,1,1)," + _
      "(45,'Total Late Days Monthly','Total hari telat selama 1 bulan',' ','Days',1,1,1)," + _
      "(46,'Total Early In Days Monthly','Total early in selama 1 bulan dalam hitungan hari',' ','Days',1,1,1)," + _
      "(47,'Total Early Out Days Monthly','Total early out selama 1 bulan dalam hitungan hari',' ','Days',1,1,1)," + _
      "(48,'Total Late Overtime Days Monthly','Total lembur akhir selama 1 bulan dalam hitungan hari',' ','Days',1,1,1)," + _
      "(49,'Total Early Overtime Days Monthly','Total lembur awal selama 1 bulan dalam hitungan hari',' ','Days',1,1,1)," + _
      "(50,'Total Actual Work Days Monthly','Total jam kerja selama 1 bulan',' ','Duration',1,1,1)," + _
      "(51,'Total Present Monthly','Total hari masuk kerja',' ','Days',1,1,1)," + _
      "(52,'Total Sick Monthly','Total hari sakit','','Days',1,1,1)," + _
      "(53,'Total Leave Monthly','Total hari leave',' ','Days',1,1,1)," + _
      "(54,'Total Alpha  Monthly','Total hari tidak masuk tanpa keterangan','','Days',1,1,1)," + _
      "(55,'Total Permit Monthly','Total hari permit','','Days',1,1,1)," + _
      "(56,'Total Actual Break Time  Monthly','Total actual istirahat selama 1 bulan',' ','Duration',1,1,1)," + _
      "(57,'Total Late Break in Duration Monthly','Total durasi telat selesai istirahat',' ','Duration',1,1,1)," + _
      "(60,'Total Early Break Out Duration  Monthly','Total durasi keluar istirahat lebih cepat dari jadwal',' ','Duration',1,1,1)," + _
      "(61,'Total Late Break In Days  Monthly','Total hari telat selesai istirahat',' ','Days',1,1,1)," + _
      "(64,'Total Early Break Out  days  Monthly','Total hari keluar istirahat lebih cepat dari jadwal','','Days',1,1,1)," + _
      "(65,'Total Weekday OT Monthly','Total hari lembur selama hari biasa dalam sebulan',' ','Days',1,1,1)," + _
      "(67,'Total Holidays OT Monthly','Total hari lembur selama hari libur dalam sebulan',' ','Days',1,1,1)," + _
      "(68,'Total Weekdays Monthly','Total hari kerja selama 1 bulan',' ','Days',1,1,1)," + _
      "(70,'Total holidays Monthly','Total hari libur selama 1 bulan',' ','Days',1,1,1)," + _
      "(73,'Sex','Sex',' ','male/female',3,0,1)," + _
      "(74,'Basic Salary','Basic Salary',' ','numeric',3,1,1)," + _
      "(75,'Mariage Status','Marriage Status',' ','Single/Married',3,1,1)," + _
      "(76,'Total Late Duration Yearly','Total durasi telat selama 1 tahun',' ','Duration',2,1,1)," + _
      "(77,'Total Early In Duration Yearly','Total durasi early in selama 1 tahun',' ','Duration',2,1,1)," + _
      "(78,'Total Early Out Duration Yearly','Total early out selama 1 tahun',' ','Duration',2,1,1)," + _
      "(79,'Total Late Overtime Duration Yearly','Total durasi lembur akhir selama 1 tahun',' ','Duration',2,1,1)," + _
      "(80,'Total Early Overtime Duration Yearly','Total durasi lembur awal selama 1 tahun',' ','Duration',2,1,1)," + _
      "(81,'Total Late Days Yearly','Total hari telat selama 1 tahun ',' ','Days',2,1,1)," + _
      "(82,'Total Early In Days Yearly','Total early in selama 1 tahun dalam hitungan hari ',' ','Days',2,1,1)," + _
      "(83,'Total Early Out Days Yearly','Total early out selama 1 tahun dalam hitungan hari',' ','Days',2,1,1)," + _
      "(84,'Total Late Overtime Days Yearly','Total lembur akhir selama 1 tahun dalam hitungan hari',' ','Days',2,1,1)," + _
      "(85,'Total Early Overtime Days Yearly','Total lembur awal selama 1 tahun dalam hitungan hari',' ','Days',2,1,1)," + _
      "(86,'Total Actual Work Time Yearly','Total jamkerja selama 1 tahun',' ','Duration',2,1,1)," + _
      "(87,'Total present Yearly','Total hari masuk kerja',' ','Days',2,1,1)," + _
      "(88,'Total Sick Yearly','Total hari sakit',' ','Days',2,1,1)," + _
      "(89,'Total Leave Yearly','Total hari leave',' ','Days',2,1,1)," + _
      "(90,'Total Alpha Yearly','Total hari tidak masuk tanpa keterangan',' ','Days',2,1,1)," + _
      "(91,'Total Permit Yearly','Total hari permit',' ','Days',2,1,1)," + _
      "(92,'Total Actual Break Time Yearly','Total actual istirahat selama 1 tahun',' ','Duration',2,1,1)," + _
      "(93,'Total Late Break In Duration Yearly','Total durasi telat selesai istirahat',' ','Duration',2,1,1)," + _
      "(96,'Total Early Break Out Duration Yearly','Total durasi keluar istirahat lebih cepat dari jadwal',' ','Duration',2,1,1)," + _
      "(97,'Total Late Break In Days Yearly','Total hari telat selesai istirahat ',' ','Days',2,1,1)," + _
      "(100,'Total Early Break Out Days Yearly','Total hari keluar istirahat lebih cepat dari jadwal',' ','Days',2,1,1)," + _
      "(101,'Total Weekday OT Yearly','Total hari lembur selama hari biasa dalam setahun',' ','Days',2,1,1)," + _
      "(103,'Total Holiday OT Yearly','Total hari lembur selama hari libur dalam setahun',' ','Days',2,1,1)," + _
      "(104,'Total Weekday Yearly','Total hari kerja selama setahun',' ','Days',2,1,1)," + _
      "(106,'Total Holiday Yearly','Total hari libur selama setahun',' ','Days',2,1,1)," + _
      "(109,'Schedule Break Time','Durasi istirahat schedule','','Duration',0,1,1)," + _
      "(110,'Total Schedule Break Time Monthly','Total schedule istirahat selama 1 bulan','','Duration',1,1,1)," + _
      "(111,'Total Schedule Break Time Yearly','Total schedule istirahat selama 1 tahun',' ','Duration',2,1,1)," + _
      "(112,'Permit Duration','Permit Duration',' ','Duration',0,1,1)," + _
      "(113,'Total Permit Duration Monthly','Total permit duration selama 1 bulan',' ','Duration',1,1,1)," + _
      "(114,'Total Permit Duration Yearly','Total permit duration selama 1 tahun',' ','Duration',2,1,1)," + _
      "(115,'Jumlah Anak','Jumlah Anak',' ','Numeric',3,1,1)," + _
      "(116,'OT Duration','Total durasi lembur jam kerja early tambah late overtime','','Duration',0,1,1)," + _
      "(117,'OT','nilai yes jika hari itu ada lembur early maupun late ot','','yes/no',0,0,1)," + _
      "(118,'Total Overtime Duration Monthly','Total durasi lembur selama 1 bulan','','Duration',1,1,1)," + _
      "(119,'Total Overtime Days Monthly','Total lembur selama 1 bulan dalam hitungan hari','','Days',1,1,1)," + _
      "(120,'Total Overtime Duration Yearly','Total durasi lembur selama 1 tahun','','Duration',2,1,1)," + _
      "(121,'Total Overtime Days Yearly','Total lembur selama 1 tahun dalam hitungan hari','','Days',2,1,1)," + _
      "(122,'Total Overtime Duration Weekday Monthly','Total durasi lembur tipe weekday selama 1 bulan','','Duration',1,1,1)," + _
      "(123,'Total Overtime Duration Holiday Monthly','Total durasi lembur tipe holiday selama 1 bulan','','Duration',1,1,1)," + _
      "(124,'Total Overtime Duration Weekday Yearly','Total durasi lembur tipe weekday selama 1 tahun','','Duration',2,1,1)," + _
      "(125,'Total Overtime Duration Holiday Yearly','Total durasi lembur tipe holiday selama 1 tahun','','Duration',2,1,1)," + _
      "(126,'Total Overtime Day Weekday Monthly','Total lembur tipe weekday selama 1 bulan dalam hitungan hari','','Days',1,1,1)," + _
      "(127,'Total Overtime Day Holiday Monthly','Total lembur tipe holiday selama 1 bulan dalam hitungan hari','','Days',1,1,1)," + _
      "(128,'Total Overtime Day Weekday Yearly','Total lembur tipe weekday selama 1 tahun dalam hitungan hari','','Days',2,1,1)," + _
      "(129,'Total Overtime Day Holiday Yearly','Total lembur tipe holiday selama 1 tahun dalam hitungan hari','','Days',2,1,1)," + _
      "(130,'Employee Status','Status pekerja','','probation/contract/jobholder',3,0,1)," + _
      "(131,'Job Position','Jabatan','','From master job position',3,0,1)," + _
      "(132,'Total Leave Balance Yearly','Total sisa cuti','','Days',2,1,0)," + _
      "(133,'Leave code','Jenis cuti','','From master leave',3,0,1);" + _
        "COMMIT;"
            Me.ExecuteSqlScript(sql, False)

            sql = "drop table if exists `tblmaster_ptkp`;"
            Me.ExecuteSqlScript(sql, False)
            sql = "CREATE TABLE `tblmaster_ptkp` (" + _
              "`primary` double(15,0) NOT NULL auto_increment," + _
          "`enum golongan pajak` int(3) NOT NULL default '0'," + _
          "`PTKP Tahunan` decimal(21,6) NOT NULL default '0.000000'," + _
          "`remarks` varchar(250) NOT NULL," + _
          "PRIMARY KEY  (`primary`)" + _
          ") ENGINE=InnoDB DEFAULT CHARSET=latin1;"
            Me.ExecuteSqlScript(sql, False)

            sql = "INSERT INTO `tblmaster_ptkp` (`primary`, `enum golongan pajak`, `PTKP Tahunan`, `remarks`) VALUES" + _
                "(1,0,24300000,'')," + _
                "(2,1,26325000,'')," + _
                "(3,2,28350000,'')," + _
                "(4,3,30375000,'')," + _
                "(5,4,26325000,'')," + _
                "(6,5,28350000,'')," + _
                "(7,6,30375000,'')," + _
                "(8,7,32400000,'');" + _
                "COMMIT;"
            Me.ExecuteSqlScript(sql, False)

            sql = "drop table if exists `tblmaster_tarif_pph`;"
            Me.ExecuteSqlScript(sql, False)

            sql = "CREATE TABLE `tblmaster_tarif_pph` (" + _
              "`primary` double(15,0) NOT NULL auto_increment," + _
              "`maximum income` decimal(21,6) NOT NULL default '0.000000'," + _
              "`tarif` decimal(21,6) NOT NULL default '0.000000'," + _
              "`remarks` varchar(250) NOT NULL," + _
              "PRIMARY KEY  (`primary`)" + _
              ") ENGINE=InnoDB DEFAULT CHARSET=latin1;"
            Me.ExecuteSqlScript(sql, False)

            sql = "INSERT INTO `tblmaster_tarif_pph` (`primary`, `maximum income`, `tarif`, `remarks`) VALUES" + _
            "(9,50000000,5,'')," + _
            "(10,250000000,15,'')," + _
            "(11,500000000,25,'')," + _
            "(12,999999999999,30,'');" + _
            "COMMIT;"
            Me.ExecuteSqlScript(sql, False)


            sql = "drop table if exists `tblmaster_kode_negara`;"
            Me.ExecuteSqlScript(sql, False)

            sql = "CREATE TABLE `tblmaster_kode_negara` (" + _
        "`Primary` double(18,0) NOT NULL auto_increment," + _
        "`Kode Negara` varchar(250) NOT NULL," + _
        "`Nama Negara` varchar(250) NOT NULL," + _
        "PRIMARY KEY  (`Primary`)," + _
        "UNIQUE KEY `Primary_2` (`Primary`)" + _
        ") ENGINE=InnoDB DEFAULT CHARSET=latin1;"
            Me.ExecuteSqlScript(sql, False)

            sql = "INSERT INTO `tblmaster_kode_negara` (`Primary`, `Kode Negara`, `Nama Negara`) VALUES" + _
              "(1,'AFG','Afganistan')," + _
              "(2,'AGO','Angola')," + _
              "(3,'ALB','Albania')," + _
              "(4,'AND','Andora')," + _
              "(5,'ARE','Uni Emirat Arab')," + _
              "(6,'ARG','Argentina')," + _
              "(7,'ARM','Armenia')," + _
              "(8,'ATG','Antigua dan Barbuda')," + _
              "(9,'AUS','Australia')," + _
              "(10,'AUT','Austria')," + _
              "(11,'AZE','Azebaijan')," + _
              "(12,'BDI','Burundi')," + _
              "(13,'BEL','Belgia')," + _
              "(14,'BEN','Benin')," + _
              "(15,'BFA','Burkina Faso')," + _
              "(16,'BGD','Bangladesh')," + _
              "(17,'BGR','Bulgaria')," + _
              "(18,'BHR','Bahrain')," + _
              "(19,'BHS','Bahama')," + _
              "(20,'BIH','Bosnia dan Herzegovina')," + _
              "(21,'BLR','Belarus')," + _
              "(22,'BLZ','Belize')," + _
              "(23,'BOL','Bolivia')," + _
              "(24,'BRA','Brasil')," + _
              "(25,'BRB','Barbados')," + _
              "(26,'BRN','Brunei Darussalam')," + _
              "(27,'BTN','Bhutan')," + _
              "(28,'BWA','Bostwana')," + _
              "(29,'CAF','Afrika Tengah')," + _
              "(30,'CAN','Canada')," + _
              "(31,'CHE','Swiss')," + _
              "(32,'CHL','Chili')," + _
              "(33,'CHN','China')," + _
              "(34,'CIV','Pantai Gading')," + _
              "(35,'CMR','Kamerun')," + _
              "(36,'COD','Kongo - Republik Demokratif Kongo')," + _
              "(37,'COG','Kongo - Republik Kongo')," + _
              "(38,'COL','Kolombia')," + _
              "(39,'COM','Komoro')," + _
              "(40,'CPV','Tanjung Verde')," + _
              "(41,'CRI','Kosta Rika')," + _
              "(42,'CUB','Kuba')," + _
              "(43,'CYP','Siprus')," + _
              "(44,'CZE','Republik Ceko')," + _
              "(45,'DEU','Jerman')," + _
              "(46,'DJI','Djiboti')," + _
              "(47,'DMA','Dominika - Persemakmuran Dominika')," + _
              "(48,'DNK','Denmark')," + _
              "(49,'DOM','Republik Dominika')," + _
              "(50,'DSA','Aljazair')," + _
              "(51,'ECU','Ekuador')," + _
              "(52,'EGY','Mesir')," + _
              "(53,'ERI','Eritrea')," + _
              "(54,'ESP','Spanyol')," + _
              "(55,'EST','Estonia')," + _
              "(56,'ETH','Ethiopia')," + _
              "(57,'FIN','Finlandia')," + _
              "(58,'FJI','Fiji')," + _
              "(59,'FRA','Perancis')," + _
              "(60,'FSM','Mikronesia')," + _
              "(61,'GAB','Gabon')," + _
              "(62,'GBR','Inggris')," + _
              "(63,'GEO','Georgia')," + _
              "(64,'GHA','Ghana')," + _
              "(65,'GIN','Guinea')," + _
              "(66,'GMB','Gambia')," + _
              "(67,'GNB','Guinea-Bissau')," + _
              "(68,'GNQ','Guinea Khatulistiwa')," + _
              "(69,'GRC','Yunani')," + _
              "(70,'GRD','Grenada')," + _
              "(71,'GTM','Guatemala')," + _
              "(72,'GUY','Guyana')," + _
              "(73,'HKG','Hong Kong')," + _
              "(74,'HND','Honduras')," + _
              "(75,'HRV','Kroasia')," + _
              "(76,'HTI','Haiti')," + _
              "(77,'HUN','Hungaria')," + _
              "(78,'IND','India')," + _
              "(79,'IRL','Irlandia')," + _
              "(80,'IRN','Iran')," + _
              "(81,'IRQ','Irak')," + _
              "(82,'ISL','Islandia')," + _
              "(83,'ISR','Israel')," + _
              "(84,'ITA','Italia')," + _
              "(85,'JAM','Jamaika')," + _
              "(86,'JOR','Yordania')," + _
              "(87,'JPN','Jepang')," + _
              "(88,'KAZ','Kazakhstan')," + _
              "(89,'KEN','Kenya')," + _
              "(90,'KGZ','Kirgiztan')," + _
              "(91,'KHM','Kamboja')," + _
              "(92,'KIR','Kiribati')," + _
              "(93,'KNA','Saint Kitts dan Nevis')," + _
              "(94,'KOR','Korea Selatan')," + _
              "(95,'KWT','Kuwait')," + _
              "(96,'LAO','Laos')," + _
              "(97,'LBN','Lebanon')," + _
              "(98,'LBR','Liberia')," + _
              "(99,'LBY','Libya')," + _
              "(100,'LCA','Saint Lucia')," + _
              "(101,'LIE','Liechtenstein')," + _
              "(102,'LKA','Sri Lanka')," + _
              "(103,'LSO','Lesotho')," + _
              "(104,'LTU','Lituania')," + _
              "(105,'LUX','Luxembourg')," + _
              "(106,'LVA','Latvia')," + _
              "(107,'MAR','Maroko')," + _
              "(108,'MCO','Monako')," + _
              "(109,'MDA','Moldova')," + _
              "(110,'MDG','Madagaskar')," + _
              "(111,'MDV','Maladewa')," + _
              "(112,'MEX','Mexico')," + _
              "(113,'MHL','Marshall')," + _
              "(114,'MKD','Makedonia')," + _
              "(115,'MLI','Malia')," + _
              "(116,'MLT','Malta')," + _
              "(117,'MMR','Myanmar')," + _
              "(118,'MNE','Montenegro')," + _
              "(119,'MNG','Mongolia')," + _
              "(120,'MOZ','Mozambik')," + _
              "(121,'MRT','Mauritania')," + _
              "(122,'MUS','Mauritius')," + _
              "(123,'MWI','Malawi')," + _
              "(124,'MYS','Malaysia')," + _
              "(125,'NAM','Namibia')," + _
              "(126,'NER','Niger')," + _
              "(127,'NGA','Nigeria')," + _
              "(128,'NIC','Nikaragua')," + _
              "(129,'NLD','Belanda')," + _
              "(130,'NOR','Norwegia')," + _
              "(131,'NPL','Nepal')," + _
              "(132,'NRU','Nauru')," + _
              "(133,'NZL','Selandia Baru')," + _
              "(134,'OMN','Oman')," + _
              "(135,'PAK','Pakistan')," + _
              "(136,'PAN','Panama')," + _
              "(137,'PER','Peru')," + _
              "(138,'PHL','Philipina')," + _
              "(139,'PLW','Palau')," + _
              "(140,'PNG','Papua Nugini')," + _
              "(141,'POL','Polandia')," + _
              "(142,'PRK','Korea Utara')," + _
              "(143,'PRT','Portugal')," + _
              "(144,'PRY','Paraguay')," + _
              "(145,'QAT','Qatar')," + _
              "(146,'ROU','Romania')," + _
              "(147,'RUS','Rusia')," + _
              "(148,'RWA','Rwanda')," + _
              "(149,'SAU','Saudi Arabia')," + _
              "(150,'SDN','Sudan')," + _
              "(151,'SEN','Senegal')," + _
              "(152,'SGP','Singapura')," + _
              "(153,'SLB','Solomon')," + _
              "(154,'SLE','Sierra Leone')," + _
              "(155,'SLV','El Salvador')," + _
              "(156,'SMR','San Marino')," + _
              "(157,'SOM','Somalia')," + _
              "(158,'SRB','Serbia')," + _
              "(159,'STP','Sao Tome dan Principe')," + _
              "(160,'SUR','Suriname')," + _
              "(161,'SVK','Slovakia')," + _
              "(162,'SVN','Slovenia')," + _
              "(163,'SWE','Swedia')," + _
              "(164,'SWZ','Swaziland')," + _
              "(165,'SYC','Seychelles')," + _
              "(166,'SYR','Suriah')," + _
              "(167,'TCS','Chad')," + _
              "(168,'TGO','Togo')," + _
              "(169,'THA','Thailand')," + _
              "(170,'TJK','Tajikistan')," + _
              "(171,'TKM','Turkmenistan')," + _
              "(172,'TLS','Timor Leste')," + _
              "(173,'TON','Tonga')," + _
              "(174,'TTO','Trinidad dan Tobago')," + _
              "(175,'TUN','Tunisia')," + _
              "(176,'TUR','Turki')," + _
              "(177,'TUV','Tuvalu')," + _
              "(178,'TWN','Taiwan')," + _
              "(179,'TZA','Tanzania')," + _
              "(180,'UGA','Uganda')," + _
              "(181,'UKR','Ukraina')," + _
              "(182,'URY','Uruguay')," + _
              "(183,'USA','Amerika Serikat')," + _
              "(184,'UZB','Uzbekistan')," + _
              "(185,'VAT','Vatikan')," + _
              "(186,'VCT','Saint Vincent dan Grenadines')," + _
              "(187,'VEN','Venezuela')," + _
              "(188,'VNM','Vietnam')," + _
              "(189,'VUT','Vanuatu')," + _
              "(190,'WSM','Samoa')," + _
              "(191,'YEM','Yaman')," + _
              "(192,'ZAF','Afrika Selatan')," + _
              "(193,'ZMB','Zambia')," + _
              "(194,'ZWE','Zimbabwe');" + _
            "COMMIT;"
            Me.ExecuteSqlScript(sql, False)
            Me.lblUpdatev160_HR.Text = "done"
            MsgBox("Finish")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdateDiscAmount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateDiscAmount.Click
        Try
            If MsgBox("Update disc amount data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker15.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker15_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker15.DoWork
        Dim DTASO As New DseUpdate_v160TableAdapters.tbltrans_so_itemTableAdapter
        Dim TBLSO As New DseUpdate_v160.tbltrans_so_itemDataTable
        Dim rowSO As DseUpdate_v160.tbltrans_so_itemRow
        Dim DTAInv As New DseUpdate_v160TableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Dim TBLInv As New DseUpdate_v160.tbltrans_salesinvoice_itemDataTable
        Dim rowInv As DseUpdate_v160.tbltrans_salesinvoice_itemRow
        Dim DTAInv_Obs As New DseUpdate_v160TableAdapters.tbltrans_salesinvoiceobselete_itemTableAdapter
        Dim TBLInv_Obs As New DseUpdate_v160.tbltrans_salesinvoiceobselete_itemDataTable
        Dim rowInv_Obs As DseUpdate_v160.tbltrans_salesinvoiceobselete_itemRow
        Dim DTASR As New DseUpdate_v160TableAdapters.tbltrans_salesreturn_itemTableAdapter
        Dim TBLSR As New DseUpdate_v160.tbltrans_salesreturn_itemDataTable
        Dim rowSR As DseUpdate_v160.tbltrans_salesreturn_itemRow

        Dim DTAPO As New DseUpdate_v160TableAdapters.tbltrans_po_itemTableAdapter
        Dim TBLPO As New DseUpdate_v160.tbltrans_po_itemDataTable
        Dim rowPO As DseUpdate_v160.tbltrans_po_itemRow
        Dim DTAAOP As New DseUpdate_v160TableAdapters.tbltrans_aop_itemTableAdapter
        Dim TBLAOP As New DseUpdate_v160.tbltrans_aop_itemDataTable
        Dim rowAOP As DseUpdate_v160.tbltrans_aop_itemRow
        Dim DTAPR As New DseUpdate_v160TableAdapters.tbltrans_purchasereturn_itemTableAdapter
        Dim TBLPR As New DseUpdate_v160.tbltrans_purchasereturn_itemDataTable
        Dim rowPR As DseUpdate_v160.tbltrans_purchasereturn_itemRow

        Try
            'so
            Me.Prg_UpdateDiscAmount.Value = 0
            DTASO.Fill(TBLSO)
            Me.Prg_UpdateDiscAmount.Maximum = TBLSO.Rows.Count
            For Each rowSO In TBLSO.Rows
                Me.Prg_UpdateDiscAmount.Value += 1
                rowSO.Disc_Amount = rowSO.Unit_Price - rowSO.Nett_Price
                DTASO.UpdateQuery(rowSO.Disc_Amount, rowSO.Disc_From_Perct, rowSO.Primary)
            Next
            'invoice
            Me.Prg_UpdateDiscAmount.Value = 0
            DTAInv.Fill(TBLInv)
            Me.Prg_UpdateDiscAmount.Maximum = TBLInv.Rows.Count
            For Each rowInv In TBLInv.Rows
                Me.Prg_UpdateDiscAmount.Value += 1
                rowInv.Disc_Amount = rowInv.Unit_Price - rowInv.Nett_Price
                DTAInv.UpdateQuery(rowInv.Disc_Amount, rowInv.DiscFromPerc, rowInv.Primary)
            Next
            'obselete invoice
            Me.Prg_UpdateDiscAmount.Value = 0
            DTAInv_Obs.Fill(TBLInv_Obs)
            Me.Prg_UpdateDiscAmount.Maximum = TBLInv_Obs.Rows.Count
            For Each rowInv_Obs In TBLInv_Obs.Rows
                Me.Prg_UpdateDiscAmount.Value += 1
                rowInv_Obs.Disc_Amount = rowInv_Obs.Unit_Price - rowInv_Obs.Nett_Price
                DTAInv_Obs.UpdateQuery(rowInv_Obs.Disc_Amount, rowInv_Obs.Disc_From_Perct, rowInv_Obs.Primary)
            Next

            'sales return
            Me.Prg_UpdateDiscAmount.Value = 0
            DTASR.Fill(TBLSR)
            Me.Prg_UpdateDiscAmount.Maximum = TBLSR.Rows.Count
            For Each rowSR In TBLSR.Rows
                Me.Prg_UpdateDiscAmount.Value += 1
                rowSR.Disc_Amount = rowSR.Unit_Price - rowSR.Nett_Price
                DTASR.UpdateQuery(rowSR.Disc_Amount, rowSR.Disc_From_Perct, rowSR.Primary)
            Next

            'po
            Me.Prg_UpdateDiscAmount.Value = 0
            DTAPO.Fill(TBLPO)
            Me.Prg_UpdateDiscAmount.Maximum = TBLPO.Rows.Count
            For Each rowPO In TBLPO.Rows
                Me.Prg_UpdateDiscAmount.Value += 1
                rowPO.Disc_Amount = rowPO.Unit_Price - rowPO.Nett_Price
                DTAPO.UpdateQuery1(rowPO.Disc_Amount, rowPO.Disc_From_Perct, rowPO.Primary)
            Next

            'aop
            Me.Prg_UpdateDiscAmount.Value = 0
            DTAAOP.Fill(TBLAOP)
            Me.Prg_UpdateDiscAmount.Maximum = TBLAOP.Rows.Count
            For Each rowAOP In TBLAOP.Rows
                Me.Prg_UpdateDiscAmount.Value += 1
                rowAOP.Disc_Amount = rowAOP.Unit_Price - rowAOP.Nett_Price
                DTAAOP.UpdateQuery(rowAOP.Disc_Amount, rowAOP.Disc_From_Perct, rowAOP.Primary)
            Next

            'purchase return
            Me.Prg_UpdateDiscAmount.Value = 0
            DTAPR.Fill(TBLPR)
            Me.Prg_UpdateDiscAmount.Maximum = TBLPR.Rows.Count
            For Each rowPR In TBLPR.Rows
                Me.Prg_UpdateDiscAmount.Value += 1
                rowPR.Disc_Amount = rowPR.Unit_Price - rowPR.Nett_Price
                DTAPR.UpdateQuery(rowPR.Disc_Amount, rowPR.Disc_From_Perct, rowPR.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASO.Dispose()
            TBLSO.Dispose()
            DTAInv.Dispose()
            TBLInv.Dispose()
            DTAInv_Obs.Dispose()
            TBLInv_Obs.Dispose()
            DTASR.Dispose()
            TBLSR.Dispose()
            DTAPO.Dispose()
            TBLPO.Dispose()
            DTAAOP.Dispose()
            TBLAOP.Dispose()
            DTAPR.Dispose()
            TBLPR.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker15_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker15.RunWorkerCompleted
        Try
            MsgBox("Finish.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
