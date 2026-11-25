Option Explicit On
Option Strict Off

Public Class FrmTransaksiDOView
    Public PrimaryDO As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjDO As New ClsTransaksiDOItem

    Private Sub FrmTransaksiDOView_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ObjDO.IsiDGV_Default(Me.dgvItemDO2, PrimaryDO)
            Me.FormatDGVDO()
            Me.dgvItemDO2.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGVDO()
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True

        Try
            'yang diformat dgv-nya:
            For Each row In Me.dgvItemDO2.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(7).Value) = False Then
                    row.Cells(7).Value = "-"
                Else
                    If CDate(row.Cells(7).Value).Date = #1/1/1950# Then
                        row.Cells(7).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(7).Value = Format(CDate(row.Cells(7).Value), "dd MMM yyyy")
                    End If
                End If

                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If IsNumeric(row.Cells(14).Value) = True Then
                    If row.Cells(14).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                If row.Cells(14).Value = "-1" Then
                    row.Cells(14).Value = "-"
                    row.Cells(15).Value = "-"
                End If

                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(6).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
            Next
            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                Me.dgvItemDO2.Columns(14).Visible = False
                Me.dgvItemDO2.Columns(15).Visible = False
            Else
                Me.dgvItemDO2.Columns(14).Visible = True
                Me.dgvItemDO2.Columns(15).Visible = True
            End If
            '3. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.dgvItemDO2.Columns(6).Visible = False
            Else
                Me.dgvItemDO2.Columns(6).Visible = True
            End If
            '4. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.dgvItemDO2.Columns(7).Visible = False
            Else
                Me.dgvItemDO2.Columns(7).Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class