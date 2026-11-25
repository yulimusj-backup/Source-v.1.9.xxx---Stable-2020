Public Class FrmTransaksiSalesInvoiceSettingPajak
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Color.LightGoldenrodYellow

    Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice
    Dim ObjInv As New ClsTransaksiSalesInvoice
    Dim Buffer As String = ""
    Dim SedangLoading As Boolean = False
    Dim SeriFPLength As Integer = Me.Obj.PanjangNoSeriFP
    Dim KodeFakturPajak As String = "010"

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Dim Dummy As String = ""

        Try
            If e.RowIndex <> Me.dgvItem.RowCount - 1 Then
                'yang boleh diedit:
                '1. prefix, 
                '2. starting number
                '3. ending number
                '4. active
                If e.ColumnIndex = Me.ColPrefix.Index Or _
                e.ColumnIndex = Me.ColStartingNumbe.Index Or _
                e.ColumnIndex = Me.ColEndingNumber.Index Or _
                e.ColumnIndex = Me.ColActive.Index Or _
                e.ColumnIndex = Me.ColYear.Index Then
                    'jika kolomnya adalah tahun atau prefix, maka tidak boleh sudah dirujuk oleh invoice
                    If e.ColumnIndex = Me.ColPrefix.Index Or e.ColumnIndex = Me.ColYear.Index Then
                        If Me.ObjInv.ApakahMenggunakanSettingFP(Me.dgvItem.Item(Me.ColPrimary.Index, e.RowIndex).Value) = True Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If

                    If e.ColumnIndex = Me.ColPrefix.Index Then
                        Dummy = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        Dummy = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace(".", "")
                        Dummy = Dummy.Replace("-", "")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Dummy
                    End If

                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Dummy As String = ""
        Dim DummyVal As Double = 0
        Dim rowDGV As DataGridViewRow

        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If

            If e.ColumnIndex = Me.ColPrefix.Index Then
                Dummy = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Dummy = Dummy.Trim
                Dummy = Dummy.Replace(".", "")
                Dummy = Dummy.Replace("-", "")
                'pastikan ada isinya
                If Dummy = "" Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'formatnya adalah:
                '.XXX-XX.
                'length = 5
                If Dummy.Length <> 5 Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
                'kembalikan ke format defaultnya
                Dummy = "." + Dummy.Substring(0, 3) + "-" + Dummy.Substring(3) + "."
                'pasangkan datanya dalam dgv
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Dummy.ToUpper
                Me.SetExample(e.RowIndex)
            ElseIf e.ColumnIndex = Me.ColStartingNumbe.Index Then
                'harus angka non negatif
                If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'harus integer
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Math.Round(CDbl(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), 0)

                'harus lebih kecil dari ending number
                If CDbl(Me.dgvItem.Item(Me.ColEndingNumber.Index, e.RowIndex).Value) < _
                CDbl(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'minimum invoice harus lebih besar dari starting number
                DummyVal = CDbl(Me.ObjInv.GetMinimumUrutanFP(Me.dgvItem.Item(Me.ColPrimary.Index, e.RowIndex).Value))
                If DummyVal <> 0 Then
                    If DummyVal < _
                    CDbl(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                        MsgBox("Maximum starting number is " + DummyVal.ToString, MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If

                'nilai max = 8 digit
                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 8 Then
                    MsgBox("Please enter a number with maximum value 99999999.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'update remarks
                Me.SetExample(e.RowIndex)
            ElseIf e.ColumnIndex = Me.ColEndingNumber.Index Then
                'harus angka non negatif
                If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'harus integer
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Math.Round(CDbl(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), 0)

                'nilai max = 8 digit
                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 8 Then
                    MsgBox("Please enter a number with maximum value 99999999.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'harus lebih besar dari starting number
                If CDbl(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) < CDbl(Me.dgvItem.Item(Me.ColStartingNumbe.Index, e.RowIndex).Value) Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'jika kolom ending number maka maximum invoice harus lebih kecil dari ending number
                DummyVal = CDbl(Me.ObjInv.GetMaximumUrutanFP(Me.dgvItem.Item(Me.ColPrimary.Index, e.RowIndex).Value))
                If DummyVal <> 0 Then
                    If DummyVal > _
                    CDbl(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                        MsgBox("Minimum ending number is " + DummyVal.ToString, MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If

                'update remarks
                Me.SetExample(e.RowIndex)
            End If

            'cek apakah datanya benar, jika sudah benar, maka simpan baru atau simpan perubahan langsung ke dalam basis data
            rowDGV = Me.dgvItem.Rows(e.RowIndex)
            If Me.ApakahDataBenar(e.RowIndex, False) = True Then
                Me.Simpan(rowDGV)
                Me.dgvItem.Item(Me.ColSave.Index, e.RowIndex).Value = ""
            Else
                Me.dgvItem.Item(Me.ColSave.Index, e.RowIndex).Value = "Save"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Simpan(ByVal rowDGV As DataGridViewRow)
        Try
            If Me.dgvItem.Item(Me.ColPrimary.Index, rowDGV.Index).Value < 0 Then
                'simpan baru
                Me.Obj.SimpanBaru(rowDGV.Cells(Me.ColYear.Index).Value.ToString, _
                rowDGV.Cells(Me.ColPrefix.Index).Value, _
                rowDGV.Cells(Me.ColRemarks.Index).Value, rowDGV.Cells(Me.ColStartingNumbe.Index).Value, _
                rowDGV.Cells(Me.ColEndingNumber.Index).Value, 1)

                rowDGV.Cells(Me.ColPrimary.Index).Value = Me.Obj.GetPrimary(rowDGV.Cells(Me.ColPrefix.Index).Value, _
                rowDGV.Cells(Me.ColStartingNumbe.Index).Value, rowDGV.Cells(Me.ColEndingNumber.Index).Value, _
                rowDGV.Cells(Me.ColYear.Index).Value)
            Else
                'simpan perubahan
                Me.Obj.SimpanPerubahan(rowDGV.Cells(Me.ColPrimary.Index).Value, rowDGV.Cells(Me.ColYear.Index).Value.ToString, _
                rowDGV.Cells(Me.ColPrefix.Index).Value, _
                rowDGV.Cells(Me.ColRemarks.Index).Value, rowDGV.Cells(Me.ColStartingNumbe.Index).Value, _
                rowDGV.Cells(Me.ColEndingNumber.Index).Value, 1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar(ByVal vRowIndex As Integer, ByVal vShwMsg As Boolean) As Boolean
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter
        Dim TBL As New DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingDataTable
        Dim row As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingRow
        Dim rowDGV As DataGridViewRow
        Dim NoDari, NoSampai As Decimal
        Dim Prefix As String = ""
        Dim Year As Date = Today

        Try
            '1. tidak boleh ada nomor yg konflik untuk 1 tahun, prefiks yg sama
            rowDGV = Me.dgvItem.Rows(vRowIndex)
            NoDari = rowDGV.Cells(Me.ColStartingNumbe.Index).Value
            NoSampai = rowDGV.Cells(Me.ColEndingNumber.Index).Value
            Prefix = rowDGV.Cells(Me.ColPrefix.Index).Value
            Year = CDate("1/1/" + rowDGV.Cells(Me.ColYear.Index).Value)

            DTA.Fill(TBL)
            'buang dirinya sendiri, sehingga tidak dibandingkan dengan dirinya sendiri
            For Each row In TBL.Rows
                If row.primary = rowDGV.Cells(Me.ColPrimary.Index).Value Then
                    TBL.Rows.Remove(row)
                    Exit For
                End If
            Next
            For Each row In TBL.Rows
                If row.prefix = Prefix And row.year = Year Then

                    '1. no sampai antara no sampai dari
                    If NoDari <= row.ending_number And NoDari >= row.starting_number Then
                        If vShwMsg = True Then
                            MsgBox("Numbering conflict, please check your data.", MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                    '2. no dari antara no sampai dari
                    If NoSampai <= row.ending_number And NoSampai >= row.starting_number Then
                        If vShwMsg = True Then
                            MsgBox("Numbering conflict, please check your data.", MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If

                    '3. no dari kurang dari no dari, dan no sampai lebih besar dr no sampai
                    If NoSampai >= row.ending_number And NoDari <= row.starting_number Then
                        If vShwMsg = True Then
                            MsgBox("Numbering conflict, please check your data.", MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            Next
            '2. 

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Private Sub SetExample()
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                Me.SetExample(rowDGV.Index)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetExample(ByVal vRowIndex As Integer)
        Dim rowDGV As DataGridViewRow
        Dim StartingNumber, EndingNumber As String
        Dim Prefix As String = ""

        Try
            rowDGV = Me.dgvItem.Rows(vRowIndex)
            StartingNumber = rowDGV.Cells(Me.ColStartingNumbe.Index).Value
            EndingNumber = rowDGV.Cells(Me.ColEndingNumber.Index).Value
            Prefix = rowDGV.Cells(Me.ColPrefix.Index).Value

            While StartingNumber.Length + Prefix.Length < Me.SeriFPLength
                StartingNumber = "0" + StartingNumber
            End While
            StartingNumber = Prefix + StartingNumber

            While EndingNumber.Length + Prefix.Length < Me.SeriFPLength
                EndingNumber = "0" + EndingNumber
            End While
            EndingNumber = Prefix + EndingNumber

            rowDGV.Cells(Me.ColRemarks.Index).Value = Me.KodeFakturPajak + StartingNumber + " - " + Me.KodeFakturPajak + EndingNumber
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvItem.CellMouseClick
        Try
            If e.ColumnIndex = Me.ColSave.Index And e.RowIndex <> Me.dgvItem.Rows.Count - 1 Then
                If Me.dgvItem.Item(Me.ColSave.Index, e.RowIndex).Value.ToString = "Save" Then
                    If Me.ApakahDataBenar(e.RowIndex, True) = True Then
                        Me.Simpan(Me.dgvItem.Rows(e.RowIndex))
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError

    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim rowdGV As DataGridViewRow

        Try
            If Me.SedangLoading = True Then
                Exit Sub
            End If

            rowdGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)

            With rowdGV
                If Today.Year <= 2017 And Today.Year >= 2013 Then
                    .Cells(Me.ColYear.Index).Value = Today.Year.ToString
                Else
                    .Cells(Me.ColYear.Index).Value = 2013
                End If
                .Cells(Me.ColStartingNumbe.Index).Value = 0
                .Cells(Me.ColEndingNumber.Index).Value = 1
                .Cells(Me.ColPrefix.Index).Value = ".900-13."
                .Cells(Me.ColPrimary.Index).Value = -1
                .Cells(Me.ColActive.Index).Value = Boolean.TrueString
                .Cells(Me.ColSave.Index).Value = "Save"

                Me.SetExample(rowdGV.Index)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmSalesInvoiceSettingPajak_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.IsiDGV()
            Me.KosongkanField()
            Me.NonAktifkanDetail()
            Me.baru = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiDGV()
        Dim i As DataGridViewButtonCell
        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.SedangLoading = True
            Obj.IsiDGV(Me.dgvItem, Me)
            'set semua button save pd disable 
            For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                i = rowdgv.Cells(Me.ColSave.Index)
                i.Value = ""
            Next
            Me.SedangLoading = False
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.SedangLoading = False
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Dim rowDGV As DataGridViewRow

        Try
            'syarat tidak boleh hapus jalan jika primary <> -1
            rowDGV = e.Row
            If rowDGV.Cells(Me.ColPrimary.Index).Value <> -1 Then
                '1. jika sudah ada invoice yg merujuk untuk setting fp ybs
                If Me.ObjInv.ApakahMenggunakanSettingFP(rowDGV.Cells(Me.ColPrimary.Index).Value) = True Then
                    MsgBox("There are invoices using this data, you cannot delete this data.", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                Else
                    'hapus tax invoice setting
                    If Me.Obj.Hapus(rowDGV.Cells(Me.ColPrimary.Index).Value) = True Then
                        MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class