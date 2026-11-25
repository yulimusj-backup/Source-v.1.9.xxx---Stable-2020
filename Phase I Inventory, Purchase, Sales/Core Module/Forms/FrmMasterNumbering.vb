Option Explicit On
Option Strict Off

Public Class FrmMasterNumbering
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim AdaPerubahan As Boolean = False
    Dim Obj As New ClsMasterTransactionNumbering
    Dim TBL As New DseMasterNumbering.tblmaster_numberingDataTable
    Dim SedangLoad As Boolean = False
    Dim ReferenceDate As Date = "3/1/2012"
    Dim ReferenceUrutan As Double = 1
    Dim Buffer As String = ""
    Dim DefaultFixedText As String = "/"

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            If Me.AdaPerubahan = True Then
                If MsgBox("Save before exit?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    If Me.SimpanPerubahan() = True Then
                        Me.Close()
                    End If
                End If
            Else
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObject()
        Try
            Me.Obj.ApakahAutoNumber = Me.chkAutoNumber.Checked = True
            Me.Obj.Reset = Me.cboReset.SelectedIndex
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function SimpanPerubahan() As Boolean
        Try
            If MsgBox("Save data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                If Me.ApakahDataBenar = True Then
                    Me.SalinKeObject()
                    If Me.Obj.SimpanPerubahan(Me.dgvItem, Me, Me.UserCode) = True Then
                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Return True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Dim TipeKomponen As ClsMasterTransactionNumbering.EnumTipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.FixedText

        Try
            If e.ColumnIndex = Me.ColFormat.Index Or e.ColumnIndex = Me.ColTipeKomponen.Index Or e.ColumnIndex = Me.ColValue.Index Then
                'jik tipenya fixed text, maka format tidak bisa diedit
                TipeKomponen = Me.dgvItem.Item(Me.ColTipe_Int.Index, e.RowIndex).Value
                If TipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.FixedText And e.ColumnIndex = Me.ColFormat.Index Then
                    e.Cancel = True
                    Exit Sub
                End If

                'jika value yg di edit, maka simpan dalam buffer, hanya berlaku untuk tipe komponen fixed text
                If e.ColumnIndex = Me.ColValue.Index Then
                    If TipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.FixedText Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColValue.Index Then
                'fixed text tidak boleh kosong
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If

                'maksimum length 50
                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 50 Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 50)
                End If

                'UPPER CASE
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper

                'set result text
                Me.SetResultText()
            ElseIf e.ColumnIndex = Me.ColTipeKomponen.Index Then
                'komponen running number hanya boleh ada 1 saja
                If Me.dgvItem.Item(Me.ColTipe_Int.Index, e.RowIndex).Value = ClsMasterTransactionNumbering.EnumTipeKomponen.RunningNumber Then
                    If Me.GetJmlRow(ClsMasterTransactionNumbering.EnumTipeKomponen.RunningNumber) > 1 Then
                        'ubah jadi fixed text
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "Fixed Text"
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellValueChanged
        Dim rowDGV As DataGridViewRow
        Dim TipeKomponen As ClsMasterTransactionNumbering.EnumTipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.FixedText

        Try
            If e.ColumnIndex = Me.ColTipeKomponen.Index Then
                If e.RowIndex = -1 Then
                    Exit Sub
                End If
                rowDGV = Me.dgvItem.Rows(e.RowIndex)

                'isikan tipe komponen item
                rowDGV.Cells(Me.ColTipe_Int.Index).Value = CInt(Me.Obj.GetTipeKomponenInt(Me.dgvItem.Item(Me.ColTipeKomponen.Index, e.RowIndex).Value))

                'isikan tipe format
                TipeKomponen = rowDGV.Cells(Me.ColTipe_Int.Index).Value
                If TipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.FixedText Then
                    Me.IsikanFormatFixedText(e.RowIndex)
                ElseIf TipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.Month Then
                    Me.IsikanCboFormatMonth(e.RowIndex)
                ElseIf TipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.RunningNumber Then
                    Me.IsikanCboFormatRunningNumber(e.RowIndex)
                ElseIf TipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.Year Then
                    Me.IsikanCboFormatYear(e.RowIndex)
                End If

                'default fixed text
                TipeKomponen = rowDGV.Cells(Me.ColTipe_Int.Index).Value
                If TipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.FixedText Then
                    If IsNothing(Me.dgvItem.Item(Me.ColValue.Index, e.RowIndex).Value) = True Then
                        Me.dgvItem.Item(Me.ColValue.Index, e.RowIndex).Value = Me.DefaultFixedText
                    ElseIf String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColValue.Index, e.RowIndex).Value) = True Then
                        Me.dgvItem.Item(Me.ColValue.Index, e.RowIndex).Value = Me.DefaultFixedText
                    End If
                End If

                Me.SetResultText()
            ElseIf e.ColumnIndex = Me.ColFormat.Index Then
                If e.RowIndex = -1 Then
                    Exit Sub
                End If
                rowDGV = Me.dgvItem.Rows(e.RowIndex)
                rowDGV.Cells(Me.ColFormatInt.Index).Value = CInt(Me.Obj.GetFormatInt(rowDGV.Cells(Me.ColFormat.Index).Value, rowDGV.Cells(Me.ColTipe_Int.Index).Value))
                rowDGV.Cells(Me.ColValue.Index).Value = Me.Obj.GetValue(rowDGV.Cells(Me.ColTipe_Int.Index).Value, _
                rowDGV.Cells(Me.ColFormatInt.Index).Value, Me.ReferenceDate, Me.ReferenceUrutan)
                TipeKomponen = rowDGV.Cells(Me.ColTipe_Int.Index).Value
                If TipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.FixedText Then
                    If IsNothing(Me.dgvItem.Item(Me.ColValue.Index, e.RowIndex).Value) = True Then
                        Me.dgvItem.Item(Me.ColValue.Index, e.RowIndex).Value = Me.DefaultFixedText
                    ElseIf String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColValue.Index, e.RowIndex).Value) = True Then
                        Me.dgvItem.Item(Me.ColValue.Index, e.RowIndex).Value = Me.DefaultFixedText
                    End If
                End If

                Me.SetResultText()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanFormatFixedText(ByVal vRowIndex As Integer)
        Dim rowDGV As DataGridViewRow
        Dim i As DataGridViewComboBoxCell

        Try
            rowDGV = Me.dgvItem.Rows(vRowIndex)
            i = rowDGV.Cells(Me.ColFormat.Index)
            i.Items.Clear()
            i.Items.Add("")
            i.Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanCboFormatRunningNumber(ByVal vRowIndex As Integer)
        Dim rowDGV As DataGridViewRow
        Dim i As DataGridViewComboBoxCell

        Try
            rowDGV = Me.dgvItem.Rows(vRowIndex)
            i = rowDGV.Cells(Me.ColFormat.Index)
            i.Items.Clear()
            i.Items.Add("01, 02, 03, 04, ...")
            i.Items.Add("001, 002, 003, 004, ...")
            i.Items.Add("0001, 0002, 0003, 0004, ...")
            i.Items.Add("00001, 00002, 00003, 00004, ...")
            i.Items.Add("000001, 000002, 000003, 000004, ...")
            i.Items.Add("0000001, 0000002, 0000003, 0000004, ...")
            i.Items.Add("00000001, 00000002, 00000003, 00000004, ...")
            i.Items.Add("000000001, 000000002, 000000003, 000000004, ...")
            i.Items.Add("0000000001, 0000000002, 0000000003, 0000000004, ...")
            i.Value = "01, 02, 03, 04, ..."
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanCboFormatMonth(ByVal vRowIndex As Integer)
        Dim rowDGV As DataGridViewRow
        Dim i As DataGridViewComboBoxCell

        Try
            rowDGV = Me.dgvItem.Rows(vRowIndex)
            i = rowDGV.Cells(Me.ColFormat.Index)
            i.Items.Clear()

            i.Items.Add("1, 2, 3, 4, ...")
            i.Items.Add("01, 02, 03, 04, ...")
            i.Items.Add("Jan, Feb, Mar, Apr, Mei, ...")
            i.Items.Add("Januari, Februari, Maret, ...")
            i.Items.Add("Jan, Feb, Mar, Apr, May, ...")
            i.Items.Add("January, February, March, ...")
            i.Items.Add("I, II, III, IV, ...")

            i.Value = "1, 2, 3, 4, ..."
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanCboFormatYear(ByVal vRowIndex As Integer)
        Dim rowDGV As DataGridViewRow
        Dim i As DataGridViewComboBoxCell

        Try
            rowDGV = Me.dgvItem.Rows(vRowIndex)
            i = rowDGV.Cells(Me.ColFormat.Index)
            i.Items.Clear()
            i.Items.Add("01, 02, 03, 04, ...")
            i.Items.Add("2001, 2002, 2003, 2004, ...")
            i.Value = "01, 02, 03, 04, ..."
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError
        Try
            If e.ColumnIndex = Me.ColFormat.Index Then
                'do nothing

            Else
                MsgBox("Data error please contact your vendor")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Fixed Text"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'data salah jika:
    '1. jika auto number, maka minimal harus ada 1 running number, 1 month, 1 year: supaya ada running number dan tidak terjadi double
    '2. jika bukan auto number, maka komponen tidak boleh ada 1 pun
    '3. tiap komponen yg tipenya fixed text harus ada isinya
    '4. hanya boleh memiliki 1 komponen running number
    '5. untuk yang reset monthly, harus ada komponen monthlynya
    Public Function ApakahDataBenar() As Boolean
        Dim Jml As Integer = 0

        Try
            '1. jika auto number, maka minimal harus ada 1 running number, 1 month, 1 year: supaya ada running number dan tidak terjadi double
            If Me.chkAutoNumber.Checked = True Then
                Jml = Me.GetJmlRow(ClsMasterTransactionNumbering.EnumTipeKomponen.RunningNumber)
                If Jml <> 1 Then                    '4. hanya boleh memiliki 1 komponen running number
                    MsgBox("Please specify 1 'running number' component", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If Me.GetJmlRow(ClsMasterTransactionNumbering.EnumTipeKomponen.Year) = 0 Then
                    MsgBox("Please specify at least 1 'year' component", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                '2. jika bukan auto number, maka komponen tidak boleh ada 1 pun
                If Me.dgvItem.RowCount <> 0 Then
                    MsgBox("No numbering component is allowed for manual numbering. Please check your component data.", MsgBoxStyle.Information, "Information")
                End If
            End If

            '5. untuk yang reset monthly, harus ada komponen monthlynya
            If Me.chkAutoNumber.Checked = True Then
                If Me.cboReset.SelectedIndex = 1 Then
                    If Me.GetJmlRow(ClsMasterTransactionNumbering.EnumTipeKomponen.Month) = 0 Then
                        MsgBox("Please specify at least 1 'month' component", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FrmMasterNumbering_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.chkAutoNumber.Checked = True
            Me.chkAutoNumber.Checked = False
            Me.SedangLoad = True
            Me.Obj.IsiCbo(Me.cboTransaction, Me.TBL)
            Me.cboTransaction.SelectedIndex = -1
            Me.cboTransaction.SelectedIndex = 0
            Me.SedangLoad = False
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboTransaction_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboTransaction.Enter
        Try
            Me.Buffer = Me.cboTransaction.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboTransaction_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboTransaction.SelectedValueChanged
        'komponen ini memastikan agar parameter "sedang load", selalu up to date 
        'dengan proses load yg sedang berjalan
        Dim SedangLoad2 As Boolean = False

        Try
            Me.dgvItem.AllowUserToAddRows = False

            If Me.cboTransaction.SelectedIndex = -1 Then
                'kosongkan field
                Me.KosongkanField()
            Else
                'cek data lama, apakah cocok dengan data baru
                If Me.SedangLoad = False Then
                    If Me.ApakahCocok = False And Me.Buffer <> Me.cboTransaction.Text Then
                        'simpan perubahan data
                        Me.SimpanPerubahan()
                    End If
                End If

                'cari detail, tampilkan data baru
                Me.TxtPrimary.Text = Me.Obj.GetPrimary(Me.cboTransaction.Text)
                If Me.SedangLoad = True Then
                    SedangLoad2 = True
                Else
                    SedangLoad = True
                    SedangLoad2 = False
                End If
                Me.dgvItem.AllowUserToAddRows = False
                Me.Obj.CariDetail(Me.TxtPrimary.Text, Me.dgvItem, Me)
                'If Me.dgvItem.Enabled = True Then
                'Me.dgvItem.AllowUserToAddRows = True
                'End If

                'salin ke field
                Me.SalinKeField()
                If SedangLoad2 = False Then
                    Me.SedangLoad = False
                End If
            End If

            If Me.chkAutoNumber.Checked = True And Me.dgvItem.Enabled = True Then
                Me.dgvItem.AllowUserToAddRows = True
            End If

            'update buffer
            Me.Buffer = Me.cboTransaction.Text

            'matikan fungsi edit jika memang sudah tidak bisa diedit kembali
            If Me.TxtPrimary.Text <> "-1" Then
                If Me.Obj.ApakahBolehEdit(Me.TxtPrimary.Text) = False Then
                    Me.chkAutoNumber.Enabled = False
                    Me.cboReset.Enabled = False
                    Me.GroupBox1.Enabled = False
                Else
                    Me.chkAutoNumber.Enabled = True
                    If Me.chkAutoNumber.Checked = False Then
                        Me.cboReset.Enabled = False
                    Else
                        Me.cboReset.Enabled = True
                    End If
                    Me.GroupBox1.Enabled = True
                    If Me.chkAutoNumber.Checked = True And Me.dgvItem.Enabled = True Then
                        Me.dgvItem.AllowUserToAddRows = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahCocok() As Boolean
        Try
            If Me.chkAutoNumber.Checked <> Me.Obj.ApakahAutoNumber Then
                Return False
            End If
            If Me.cboReset.SelectedIndex <> Me.Obj.Reset Then
                Return False
            End If
            Me.dgvItem.AllowUserToAddRows = False
            If Me.Obj.ApakahItemCocok(Me.dgvItem, Me, Me.TxtPrimary.Text) = False Then
                Return False
            End If
            If Me.chkAutoNumber.Checked = True And Me.dgvItem.Enabled = True Then
                Me.dgvItem.AllowUserToAddRows = True
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeField()
        Try
            If Me.Obj.ApakahAutoNumber = True Then
                Me.chkAutoNumber.Checked = True
            Else
                Me.chkAutoNumber.Checked = False
            End If
            Me.cboReset.SelectedIndex = Me.Obj.Reset
            Me.SetResultText()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.dgvItem.Rows.Clear()
            Me.chkAutoNumber.Checked = False
            Me.cboReset.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAutoNumber_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAutoNumber.CheckedChanged
        Try
            If Me.chkAutoNumber.Checked = True Then
                If Me.SedangLoad = False Then
                    'tambahkan baris pertama running number dengan nilai default
                    Me.TambahkanBarisDefault()
                End If

                'aktifkan item
                Me.dgvItem.AllowUserToAddRows = True
                Me.dgvItem.AllowUserToDeleteRows = True
                Me.dgvItem.ReadOnly = False
                Me.GroupBox1.Enabled = True
                Me.cboReset.Enabled = True
            Else
                Me.dgvItem.AllowUserToAddRows = False
                Me.dgvItem.AllowUserToDeleteRows = False
                Me.dgvItem.ReadOnly = True
                Me.GroupBox1.Enabled = False
                Me.cboReset.Enabled = False
                Me.TxtResult.Text = ""
                Me.dgvItem.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahkanBarisDefault_DO()
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Running Number"
            rowDGV.Cells(Me.ColFormat.Index).Value = "000001, 000002, 000003, 000004, ..."
            rowDGV.Cells(Me.ColValue.Index).Value = "000001"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Fixed Text"
            rowDGV.Cells(Me.ColFormat.Index).Value = ""
            rowDGV.Cells(Me.ColValue.Index).Value = "/DO/"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Month"
            rowDGV.Cells(Me.ColFormat.Index).Value = "01, 02, 03, 04, ..."
            rowDGV.Cells(Me.ColValue.Index).Value = "03"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Fixed Text"
            rowDGV.Cells(Me.ColFormat.Index).Value = ""
            rowDGV.Cells(Me.ColValue.Index).Value = "/"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Year"
            rowDGV.Cells(Me.ColFormat.Index).Value = "2001, 2002, 2003, 2004, ..."
            rowDGV.Cells(Me.ColValue.Index).Value = "2012"
            Me.SetResultText()
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahkanBarisDefault_SalesInvoice()
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Running Number"
            rowDGV.Cells(Me.ColFormat.Index).Value = "000001, 000002, 000003, 000004, ..."
            rowDGV.Cells(Me.ColValue.Index).Value = "000001"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Fixed Text"
            rowDGV.Cells(Me.ColFormat.Index).Value = ""
            rowDGV.Cells(Me.ColValue.Index).Value = "/INV/"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Month"
            rowDGV.Cells(Me.ColFormat.Index).Value = "01, 02, 03, 04, ..."
            rowDGV.Cells(Me.ColValue.Index).Value = "03"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Fixed Text"
            rowDGV.Cells(Me.ColFormat.Index).Value = ""
            rowDGV.Cells(Me.ColValue.Index).Value = "/"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Year"
            rowDGV.Cells(Me.ColFormat.Index).Value = "2001, 2002, 2003, 2004, ..."
            rowDGV.Cells(Me.ColValue.Index).Value = "2012"
            Me.SetResultText()
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahkanBarisDefault_POImpor()
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Running Number"
            rowDGV.Cells(Me.ColFormat.Index).Value = "000001, 000002, 000003, 000004, ..."
            rowDGV.Cells(Me.ColValue.Index).Value = "000001"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Fixed Text"
            rowDGV.Cells(Me.ColFormat.Index).Value = ""
            rowDGV.Cells(Me.ColValue.Index).Value = "/PO/I/"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Month"
            rowDGV.Cells(Me.ColFormat.Index).Value = "01, 02, 03, 04, ..."
            rowDGV.Cells(Me.ColValue.Index).Value = "03"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Fixed Text"
            rowDGV.Cells(Me.ColFormat.Index).Value = ""
            rowDGV.Cells(Me.ColValue.Index).Value = "/"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Year"
            rowDGV.Cells(Me.ColFormat.Index).Value = "2001, 2002, 2003, 2004, ..."
            rowDGV.Cells(Me.ColValue.Index).Value = "2012"
            Me.SetResultText()
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahkanBarisDefault_POLokal()
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Running Number"
            rowDGV.Cells(Me.ColFormat.Index).Value = "000001, 000002, 000003, 000004, ..."
            rowDGV.Cells(Me.ColValue.Index).Value = "000001"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Fixed Text"
            rowDGV.Cells(Me.ColFormat.Index).Value = ""
            rowDGV.Cells(Me.ColValue.Index).Value = "/PO/L/"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Month"
            rowDGV.Cells(Me.ColFormat.Index).Value = "01, 02, 03, 04, ..."
            rowDGV.Cells(Me.ColValue.Index).Value = "03"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Fixed Text"
            rowDGV.Cells(Me.ColFormat.Index).Value = ""
            rowDGV.Cells(Me.ColValue.Index).Value = "/"
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Year"
            rowDGV.Cells(Me.ColFormat.Index).Value = "2001, 2002, 2003, 2004, ..."
            rowDGV.Cells(Me.ColValue.Index).Value = "2012"
            Me.SetResultText()
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahkanBarisDefault()
        Dim rowDGV As DataGridViewRow

        Try
            If Me.Obj.TipeTransaksi = ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal Then
                Me.TambahkanBarisDefault_POLokal()
            ElseIf Me.Obj.TipeTransaksi = ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor Then
                Me.TambahkanBarisDefault_POImpor()
            ElseIf Me.Obj.TipeTransaksi = ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice Then
                Me.TambahkanBarisDefault_SalesInvoice()
            ElseIf Me.Obj.TipeTransaksi = ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO Then
                Me.TambahkanBarisDefault_DO()
            Else
                Me.dgvItem.AllowUserToAddRows = False
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Running Number"
                rowDGV.Cells(Me.ColFormat.Index).Value = "001, 002, 003, 004, ..."
                rowDGV.Cells(Me.ColValue.Index).Value = "001"
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Fixed Text"
                rowDGV.Cells(Me.ColFormat.Index).Value = ""
                rowDGV.Cells(Me.ColValue.Index).Value = "/"
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Month"
                rowDGV.Cells(Me.ColFormat.Index).Value = "01, 02, 03, 04, ..."
                rowDGV.Cells(Me.ColValue.Index).Value = "03"
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Fixed Text"
                rowDGV.Cells(Me.ColFormat.Index).Value = ""
                rowDGV.Cells(Me.ColValue.Index).Value = "/"
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                rowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Year"
                rowDGV.Cells(Me.ColFormat.Index).Value = "01, 02, 03, 04, ..."
                rowDGV.Cells(Me.ColValue.Index).Value = "12"
                Me.SetResultText()
                Me.dgvItem.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetResultText()
        Dim Dummy As String = ""
        Dim rowDGV As DataGridViewRow
        Dim TipeKomponen As ClsMasterTransactionNumbering.EnumTipeKomponen
        Dim Format As ClsMasterTransactionNumbering.EnumFormat

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If Me.dgvItem.AllowUserToAddRows = True And rowDGV.Index = Me.dgvItem.Rows.Count - 1 Then
                    Exit For
                End If

                TipeKomponen = rowDGV.Cells(Me.ColTipe_Int.Index).Value
                Format = rowDGV.Cells(Me.ColFormatInt.Index).Value

                If TipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.FixedText Then
                    If IsNothing(rowDGV.Cells(Me.ColValue.Index).Value) = True Then
                        rowDGV.Cells(Me.ColValue.Index).Value = ""
                    End If
                    If Dummy = "" Then
                        Dummy = rowDGV.Cells(Me.ColValue.Index).Value.ToString
                    Else
                        Dummy = Dummy + rowDGV.Cells(Me.ColValue.Index).Value.ToString
                    End If
                Else
                    If Dummy = "" Then
                        Dummy = Me.Obj.GetValue(TipeKomponen, Format, Me.ReferenceDate, Me.ReferenceUrutan)
                    Else
                        Dummy = Dummy + Me.Obj.GetValue(TipeKomponen, Format, Me.ReferenceDate, Me.ReferenceUrutan)
                    End If
                End If
            Next

            Me.TxtResult.Text = Dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSimpan.Click
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.ApakahCocok = False Then
                Me.SimpanPerubahan()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            If Me.chkAutoNumber.Checked = True Then
                Me.dgvItem.AllowUserToAddRows = False
            End If
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Me.SetResultText()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'yg tidak boleh didelete:
    '1. running number, jika hanya ada 1 running number
    '2. month, jika hanya ada 1 month
    '3. year jika hanya ada 1 year
    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Dim TipeKomponen As ClsMasterTransactionNumbering.EnumTipeKomponen

        Try
            TipeKomponen = Me.dgvItem.Item(Me.ColTipe_Int.Index, e.Row.Index).Value
            If TipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.Year Or _
            TipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.RunningNumber Then
                If Me.GetJmlRow(TipeKomponen) <= 1 Then
                    e.Cancel = True
                End If
            ElseIf TipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.Month Then
                If Me.cboReset.SelectedIndex = 1 Then
                    'monthly
                    If Me.GetJmlRow(TipeKomponen) <= 1 Then
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJmlRow(ByVal vTipeKomponen As ClsMasterTransactionNumbering.EnumTipeKomponen) As Double
        Dim Jml As Double = 0
        Dim rowDGV As DataGridViewRow
        Dim Batas As Double = 0

        Try
            If Me.dgvItem.AllowUserToAddRows = True Then
                Batas = Me.dgvItem.Rows.Count - 2
            Else
                Batas = Me.dgvItem.Rows.Count - 1
            End If
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index <= Batas Then
                    If rowDGV.Cells(Me.ColTipe_Int.Index).Value = vTipeKomponen Then
                        Jml += 1
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SetButtonUpDown()
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                If Me.dgvItem.SelectedRows(0).Index <> Me.dgvItem.Rows.Count - 1 Then
                    If Me.dgvItem.SelectedRows(0).Index = Me.dgvItem.RowCount - 2 And Me.dgvItem.SelectedRows(0).Index = 0 Then
                        Me.btnUP.Enabled = False
                        Me.btnDown.Enabled = False
                    ElseIf Me.dgvItem.SelectedRows(0).Index = Me.dgvItem.RowCount - 2 Then
                        Me.btnUP.Enabled = True
                        Me.btnDown.Enabled = False
                    ElseIf Me.dgvItem.SelectedRows(0).Index = 0 Then
                        Me.btnUP.Enabled = False
                        Me.btnDown.Enabled = True
                    Else
                        Me.btnUP.Enabled = True
                        Me.btnDown.Enabled = True
                    End If
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            Else
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Me.SetButtonUpDown()
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
                    If row.Displayed = False Then
                        Me.ShowRow(row, True)
                    End If
                End If
                Me.SetResultText()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean)
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
                    Me.ShowRow(row, False)
                End If
                Me.SetResultText()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboTransaction_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTransaction.SelectedIndexChanged

    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick

    End Sub
End Class