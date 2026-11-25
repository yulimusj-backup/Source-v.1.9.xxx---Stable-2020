Option Explicit On
Option Strict Off
Public Class FrmMasterProductionRate_Apply
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public mProductionRate As Double = -1
    Public mPrimaryProcess As Double = -1
    Public UserCode As String = ""
    Public TipeDependency As clsMasterProduct.EnumTipeDependency = clsMasterProduct.EnumTipeDependency.End_Start1
    Public DelayDependency As Double = 0
    Dim objCd As New Core_Module.ClsCekData
    Dim obj As New ClsMasterProductionRate

    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim objProduct As New Core_Module.clsMasterProduct

    Private Sub FrmMasterProductionRate_Apply_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ControlBox = False
            FrmMasterProductionRate.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker1.RunWorkerAsync()
            'Me.Apply()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Me.objCd.UpdateThreading()
            Me.Apply()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Apply()
        Dim dtaOtherTrans As New DseMasterProductionRateTableAdapters.DataTable3TableAdapter
        Dim tblOtherTrans As New DseMasterProductionRate.DataTable3DataTable
        Dim rowOtherTrans As DseMasterProductionRate.DataTable3Row

        Dim dtaCurrentTrans As New DseMasterProductionRateTableAdapters.tbltrans_production_rate_item_productTableAdapter
        Dim tblCurrentTrans As New DseMasterProductionRate.tbltrans_production_rate_item_productDataTable
        Dim rowCurrentTrans As DseMasterProductionRate.tbltrans_production_rate_item_productRow

        Dim dtaP As New DseMasterProductionRateTableAdapters.DataTable4TableAdapter
        Dim tblP As New DseMasterProductionRate.DataTable4DataTable
        Dim rowP As DseMasterProductionRate.DataTable4Row
        Dim dtaC As New DseMasterProductionRateTableAdapters.DataTable5TableAdapter
        Dim tblC As New DseMasterProductionRate.DataTable5DataTable
        Dim rowC As DseMasterProductionRate.DataTable5Row
        Dim dtaSB As New DseMasterProductionRateTableAdapters.DataTable6TableAdapter
        Dim tblSB As New DseMasterProductionRate.DataTable6DataTable
        Dim rowSB As DseMasterProductionRate.DataTable6Row
        Dim dtaB As New DseMasterProductionRateTableAdapters.DataTable7TableAdapter
        Dim tblB As New DseMasterProductionRate.DataTable7DataTable
        Dim rowB As DseMasterProductionRate.DataTable7Row

        Dim dtaC2 As New DseMasterProductionRateTableAdapters.DataTable8TableAdapter
        Dim tblC2 As New DseMasterProductionRate.DataTable8DataTable
        Dim rowC2 As DseMasterProductionRate.DataTable8Row
        Dim dtaSB2 As New DseMasterProductionRateTableAdapters.DataTable9TableAdapter
        Dim tblSB2 As New DseMasterProductionRate.DataTable9DataTable
        Dim rowSB2 As DseMasterProductionRate.DataTable9Row
        Dim dtaB2 As New DseMasterProductionRateTableAdapters.DataTable10TableAdapter
        Dim tblB2 As New DseMasterProductionRate.DataTable10DataTable
        Dim rowB2 As DseMasterProductionRate.DataTable10Row
        Dim dtaG2 As New DseMasterProductionRateTableAdapters.DataTable11TableAdapter
        Dim tblG2 As New DseMasterProductionRate.DataTable11DataTable
        Dim rowG2 As DseMasterProductionRate.DataTable11Row

        Dim tblUpdate As New DseMasterProductionRate.DataTable12DataTable
        Dim rowUpdate As DseMasterProductionRate.DataTable12Row
        Dim listOfPrimaryProductExclude As New List(Of Double)

        Dim dtaUpdate As New DseMasterProductionRateTableAdapters.tblmaster_product_item_workunitTableAdapter
        Dim dtaWU As New DseMasterProductionRateTableAdapters.tblmsater_production_rate_item_workunitTableAdapter
        Dim tblWU As New DseMasterProductionRate.tblmsater_production_rate_item_workunitDataTable
        Dim rowWU As DseMasterProductionRate.tblmsater_production_rate_item_workunitRow

        Dim DTA_PredefinedWU_ItemProses As New DseMasterProductionRateTableAdapters.tbltrans_mo_master_predefine_workunit_itemprocessTableAdapter
        Dim TBL_PreDefinedWU_ItemProses As New DseMasterProductionRate.tbltrans_mo_master_predefine_workunit_itemprocessDataTable
        Dim row_PredefinedWU_ItemProses As DseMasterProductionRate.tbltrans_mo_master_predefine_workunit_itemprocessRow
        Dim counter As Double = 0

        Try
            dtaP.Connection.ConnectionString = Core_Module.GetConstring
            DTA_PredefinedWU_ItemProses.Connection.ConnectionString = dtaP.Connection.ConnectionString
            dtaC.Connection.ConnectionString = Core_Module.GetConstring
            dtaSB.Connection.ConnectionString = Core_Module.GetConstring
            dtaB.Connection.ConnectionString = Core_Module.GetConstring

            dtaC2.Connection.ConnectionString = Core_Module.GetConstring
            dtaSB2.Connection.ConnectionString = Core_Module.GetConstring
            dtaB2.Connection.ConnectionString = Core_Module.GetConstring

            dtaOtherTrans.Connection.ConnectionString = Core_Module.GetConstring
            dtaCurrentTrans.Connection.ConnectionString = Core_Module.GetConstring

            dtaUpdate.Connection.ConnectionString = Core_Module.GetConstring
            dtaWU.Connection.ConnectionString = Core_Module.GetConstring

            dtaOtherTrans.Fill(tblOtherTrans, Me.mProductionRate, Me.obj.getMinimalGroupLevel(Me.mProductionRate), Me.mPrimaryProcess)
            dtaCurrentTrans.Fill(tblCurrentTrans, Me.mProductionRate)

            ''kumpulkan produk produk dari transaksi lain, jika punya group level yang lebih rendah (spesifik) dibanding transaksi ini dg proses yang sama
            ''nanti jangan update...
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = tblOtherTrans.Count
            For Each rowOtherTrans In tblOtherTrans
                Me.ProgressBar1.Value += 1
                Me.Label1.Text = "Checking other transcation to exclude...item " & Me.ProgressBar1.Value & " of " & Me.ProgressBar1.Maximum & "."
                With rowOtherTrans
                    Select Case .group_level
                        Case 4
                            dtaP.Fill(tblP, .primary_ref_party)
                            rowP = tblP.Rows(0)
                            .primary_category = rowP.primary_category
                            .primary_sub_brand = rowP.primary_sub_brand
                            .primary_brand = rowP.primary_brand
                            .primary_group = rowP.primary_group
                        Case 3
                            dtaC.Fill(tblC, .primary_ref_party)
                            rowC = tblC.Rows(0)
                            .primary_sub_brand = rowC.primary_sub_brand
                            .primary_brand = rowC.primary_brand
                            .primary_group = rowC.primary_group
                        Case 2
                            dtaSB.Fill(tblSB, .primary_ref_party)
                            rowSB = tblSB.Rows(0)
                            .primary_brand = rowSB.primary_brand
                            .primary_group = rowSB.primary_group
                        Case 1
                            dtaB.Fill(tblB, .primary_ref_party)
                            rowB = tblB.Rows(0)
                            .primary_group = rowB.primary_group
                    End Select
                    tblCurrentTrans.DefaultView.RowFilter = "[group level] < " & .group_level.ToString
                    If tblCurrentTrans.DefaultView.Count > 0 Then
                        For i As Integer = 0 To tblCurrentTrans.DefaultView.Count - 1
                            rowCurrentTrans = tblCurrentTrans.DefaultView(i).Row
                            If rowCurrentTrans.group_level = 3 And rowCurrentTrans.primary_ref_party = .primary_category Or _
                                rowCurrentTrans.group_level = 2 And rowCurrentTrans.primary_ref_party = .primary_sub_brand Or _
                                rowCurrentTrans.group_level = 1 And rowCurrentTrans.primary_ref_party = .primary_brand Or _
                                rowCurrentTrans.group_level = 0 And rowCurrentTrans.primary_ref_party = .primary_group Then
                                If .group_level = 4 Then
                                    If Not listOfPrimaryProductExclude.Contains(.primary_ref_party) Then
                                        listOfPrimaryProductExclude.Add(.primary_ref_party)
                                    End If
                                ElseIf .group_level = 3 Then
                                    dtaC2.Fill(tblC2, .primary_ref_party, Me.mPrimaryProcess)
                                    For Each rowC2 In tblC2
                                        If Not listOfPrimaryProductExclude.Contains(rowC2.Primary) Then
                                            listOfPrimaryProductExclude.Add(rowC2.Primary)
                                        End If
                                    Next
                                ElseIf .group_level = 2 Then
                                    dtaSB2.Fill(tblSB2, .primary_ref_party, Me.mPrimaryProcess)
                                    For Each rowSB2 In tblSB2
                                        If Not listOfPrimaryProductExclude.Contains(rowSB2.Primary) Then
                                            listOfPrimaryProductExclude.Add(rowSB2.Primary)
                                        End If
                                    Next
                                ElseIf .group_level = 1 Then
                                    dtaB2.Fill(tblB2, .primary_ref_party, Me.mPrimaryProcess)
                                    For Each rowB2 In tblB2
                                        If Not listOfPrimaryProductExclude.Contains(rowB2.Primary) Then
                                            listOfPrimaryProductExclude.Add(rowB2.Primary)
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    End If
                End With
            Next

            ''kumpulkan daftar produk dari transaksi ini
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = tblCurrentTrans.Count
            For Each rowCurrentTrans In tblCurrentTrans.Rows
                Me.ProgressBar1.Value += 1
                Me.Label1.Text = "Collecting product list to update...item " & Me.ProgressBar1.Value & " of " & Me.ProgressBar1.Maximum & "."
                Select Case rowCurrentTrans.group_level
                    Case 0
                        dtaG2.Fill(tblG2, rowCurrentTrans.primary_ref_party, Me.mPrimaryProcess)
                        For Each rowG2 In tblG2.Rows
                            Me.MasukanListUpdate(tblUpdate, rowG2)
                        Next
                    Case 1
                        dtaB2.Fill(tblB2, rowCurrentTrans.primary_ref_party, Me.mPrimaryProcess)
                        For Each rowB2 In tblB2.Rows
                            Me.MasukanListUpdate(tblUpdate, rowB2)
                        Next
                    Case 2
                        dtaSB2.Fill(tblSB2, rowCurrentTrans.primary_ref_party, Me.mPrimaryProcess)
                        For Each rowSB2 In tblSB2.Rows
                            Me.MasukanListUpdate(tblUpdate, rowSB2)
                        Next
                    Case 3
                        dtaC2.Fill(tblC2, rowCurrentTrans.primary_ref_party, Me.mPrimaryProcess)
                        For Each rowC2 In tblC2.Rows
                            Me.MasukanListUpdate(tblUpdate, rowC2)
                        Next
                    Case 4
                        tblUpdate.DefaultView.RowFilter = "[Primary Product] = " & rowCurrentTrans.primary_ref_party.ToString
                        If tblUpdate.DefaultView.Count = 0 Then
                            rowUpdate = tblUpdate.NewDataTable12Row
                            rowUpdate.Primary_Product = rowCurrentTrans.primary_ref_party
                            rowUpdate._Decimal = Me.objProduct.GetAngkaDecimal(rowCurrentTrans.primary_ref_party)
                            tblUpdate.Rows.Add(rowUpdate)
                        End If
                End Select
            Next

            ''update semua produk pada transaksi ini, kecuali jika ada pada daftar produk pengecualian
            ''jangan lupa hanya produk dg proses ybs
            dtaWU.Fill(tblWU, Me.mProductionRate)
            DTA_PredefinedWU_ItemProses.Fill(TBL_PreDefinedWU_ItemProses, Me.mPrimaryProcess)
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = tblUpdate.Rows.Count

            Dim insertQuery As String = ""
            Dim values As String = ""
            Dim IdentifierProses As String = ""

            For Each rowUpdate In tblUpdate
                Me.ProgressBar1.Value += 1
                If Math.IEEERemainder(Me.ProgressBar1.Value, 5) = 0 Then
                    Me.Label1.Text = "Updating product...item " & Me.ProgressBar1.Value & " of " & Me.ProgressBar1.Maximum & "."
                ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                    Me.Label1.Text = "Updating product...item " & Me.ProgressBar1.Value & " of " & Me.ProgressBar1.Maximum & "."
                End If
                If Not listOfPrimaryProductExclude.Contains(rowUpdate.Primary_Product) Then
                    dtaUpdate.DeleteQuery(rowUpdate.Primary_Product)
                    TBL_PreDefinedWU_ItemProses.DefaultView.RowFilter = "[primary product] = " + rowUpdate.Primary_Product.ToString

                    For Each rowWU In tblWU
                        values = values & "(" & rowUpdate.Primary_Product.ToString & "," & rowWU.primary_work_unit.ToString & "," & rowWU.eficiency & ",'" & rowWU.remarks.Replace("\", "\\").Replace(vbCrLf, "\r\n").Replace("'", "''") & "'," & Me.objCd.DigitGrouping(rowWU.rate, rowUpdate._Decimal) & "),"
                        'dtaUpdate.InsertQuery(rowUpdate.Primary_Product, rowWU.primary_work_unit, rowWU.eficiency, rowWU.remarks, Me.objCd.DigitGrouping(rowWU.rate, rowUpdate._Decimal))

                        For Each DV As Data.DataRowView In TBL_PreDefinedWU_ItemProses.DefaultView
                            row_PredefinedWU_ItemProses = DV.Row
                            If row_PredefinedWU_ItemProses.primary_work_unit = rowWU.primary_work_unit Then
                                DTA_PredefinedWU_ItemProses.UpdateQuery(CDec(Me.objCd.DigitGrouping(rowWU.rate, rowUpdate._Decimal)), CDec(Me.objCd.DigitGrouping(rowWU.eficiency, 2)), _
                                CDec(Me.objCd.DigitGrouping(rowWU.effective_rate, rowUpdate._Decimal)), row_PredefinedWU_ItemProses.identifier, row_PredefinedWU_ItemProses.primary_work_unit)
                                '''Exit For
                            End If
                        Next
                    Next
                    dtaUpdate.UpdateQuery(Me.UserCode, Now, Me.TipeDependency, Me.DelayDependency, rowUpdate.Primary_Product)
                    counter += 1

                    If counter > 500 Then
                        counter -= 500
                        If values.Length > 0 Then
                            insertQuery = "INSERT INTO tblmaster_product_item_workunit (`primary product`, `primary work unit`, efficiency, remarks, `rate per hour`) VALUES " & values.Remove(values.Length - 1, 1) & ";"
                            Dim mysqlCommand As New MySql.Data.MySqlClient.MySqlCommand(insertQuery, New MySql.Data.MySqlClient.MySqlConnection(Core_Module.GetConstring))
                            mysqlCommand.Connection.Open()
                            mysqlCommand.ExecuteNonQuery()
                            mysqlCommand.Connection.Close()

                            mysqlCommand.Dispose()
                        End If
                        values = ""
                    End If
                End If
            Next

            If values.Length > 0 Then
                insertQuery = "INSERT INTO tblmaster_product_item_workunit (`primary product`, `primary work unit`, efficiency, remarks, `rate per hour`) VALUES " & values.Remove(values.Length - 1, 1) & ";"
                Dim mysqlCommand As New MySql.Data.MySqlClient.MySqlCommand(insertQuery, New MySql.Data.MySqlClient.MySqlConnection(Core_Module.GetConstring))
                mysqlCommand.Connection.Open()
                mysqlCommand.ExecuteNonQuery()
                mysqlCommand.Connection.Close()

                Me.ProgressBar1.Value = 0
                Me.ProgressBar1.Maximum = 0
                mysqlCommand.Dispose()
            End If

            'update rate produksi u/ item predefine wu untuk proses ybs

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_PredefinedWU_ItemProses.Dispose()
            TBL_PreDefinedWU_ItemProses.Dispose()
            dtaOtherTrans.Dispose()
            tblOtherTrans.Dispose()
            dtaCurrentTrans.Dispose()
            tblCurrentTrans.Dispose()
            dtaP.Dispose()
            tblP.Dispose()
            dtaC.Dispose()
            tblC.Dispose()
            dtaSB.Dispose()
            tblSB.Dispose()
            dtaB.Dispose()
            tblB.Dispose()
            dtaC2.Dispose()
            tblC2.Dispose()
            dtaSB2.Dispose()
            tblSB2.Dispose()
            dtaB2.Dispose()
            tblB2.Dispose()
            dtaG2.Dispose()
            tblG2.Dispose()
            tblUpdate.Dispose()
            listOfPrimaryProductExclude = Nothing

            dtaUpdate.Dispose()
            dtaWU.Dispose()
            tblWU.Dispose()
        End Try
    End Sub

    Private Sub MasukanListUpdate(ByVal tblUpdate As DseMasterProductionRate.DataTable12DataTable, ByVal row As DataRow)
        Dim rowUpdate As DseMasterProductionRate.DataTable12Row
        Try
            tblUpdate.DefaultView.RowFilter = "[Primary Product] = " & row.Item("Primary").ToString
            If tblUpdate.DefaultView.Count = 0 Then
                rowUpdate = tblUpdate.NewDataTable12Row
                rowUpdate.Primary_Product = row.Item("Primary")
                rowUpdate._Decimal = row.Item("Decimal")
                tblUpdate.Rows.Add(rowUpdate)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getPrimaryProcess(ByVal vprimaryProduct As Double) As Double
        Dim dta As New DseMasterProductionRateTableAdapters.tblmaster_product_item_workunitTableAdapter
        Dim v As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            v = dta.ScalarQuery_GetPrimaryProcess(vprimaryProduct)
            If v.HasValue Then
                Return v.Value
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.Label1.Text = "Finish..."
            MsgBox("Rate setting applied.", MsgBoxStyle.Information, "Information")
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 
End Class