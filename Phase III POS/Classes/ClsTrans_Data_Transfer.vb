Public Class ClsTrans_Data_Transfer

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransferDataPOS
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New DseTransaksiAntrianTableAdapters.tbltrans_antrianTableAdapter

    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307"
        End Try
    End Function

    ''Coding standarnya 
    Private Sub GetNoAntrian()
        Try
            DTAAntrian.Connection.ConnectionString = Me.GetConstring
            Me.vIdentifier = Me.ObjRnd.Next(1000000)
            'jangan sampai sama dengan identifier yang sudah ada
            While Me.DTAAntrian.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 0
                Me.vIdentifier = Me.ObjRnd.Next(1000000)
            End While
            'masukan data baru ke antrian
            Me.DTAAntrian.Insert(Me.vTipeAntrian, Me.vIdentifier)

            'cek jangan sampai ada 2 data yang sama
            While Me.DTAAntrian.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 1
                'bila ada yang sama maka harus masukkan data baru
                DTAAntrian.DeleteQuery(Me.vTipeAntrian, Me.vIdentifier)
                'minta identifier baru
                Me.vIdentifier = Me.ObjRnd.Next(1000000)
                'jangan sampai sama dengan identifier yang sudah ada
                While Me.DTAAntrian.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 0
                    Me.vIdentifier = Me.ObjRnd.Next(1000000)
                    'masukan data baru ke antrian
                    Me.DTAAntrian.Insert(Me.vTipeAntrian, Me.vIdentifier)
                End While
            End While

            'ambil no antrian
            Me.vPrimary_Antrian = DTAAntrian.ScalarQuery_GetPrimary(Me.vIdentifier, Me.vTipeAntrian).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function IsiListBoxBrand(ByVal lbo As CheckedListBox, ByVal TBL As dseTrans_Data_Transfer.tblmaster_cabangDataTable) As Boolean
        Dim DTA As New dseTrans_Data_TransferTableAdapters.tblmaster_cabangTableAdapter
        Dim row As dseTrans_Data_Transfer.tblmaster_cabangRow
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            DTA.Fill(TBL)
            lbo.DataSource = TBL.DefaultView
            For Each row In TBL.Rows
                row.Combination = row.Code + " - " + row.Description
            Next
            lbo.DisplayMember = "Combination"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
