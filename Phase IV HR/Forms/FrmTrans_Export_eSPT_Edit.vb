Public Class FrmTrans_Export_eSPT_Edit
    Public vPrimarySalary As Double = -1
    Public vUrutanBuktiPotong As String = ""
    Public vPrefixBuktiPotong As String = ""
    Public vStartDoc As EnumStartDoc = EnumStartDoc.PPH21Pesangon
    Public vTglSalary As Date = #1/1/1950#

    Dim ObjSalary As New ClsTransaksiSalary
    Dim ObjEmployee As New Core_Module.ClsMasterKaryawan
    Dim BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCD As New Core_Module.ClsCekData
    Dim RefDocNo As String = ""

    Public Enum EnumStartDoc As Integer
        PPH21
        PPH21Pesangon
    End Enum

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim DummyPrimary As Nullable(Of Double)
        Dim DTA As New DseTrans_Export_eSPTTableAdapters.tbltrans_salary_monthlyTableAdapter
        Dim Dummy As String = ""

        Try
            'cek semua karakter sudah diisikan
            If Me.TxtNoUrut.Text.Length <> 7 Then
                MsgBox("Please specify a valid document number.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If CDec(Me.TxtNoUrut.Text) = 0 Then
                MsgBox("Please specify a valid document number.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'cek apakah sudah digunakan oleh pihak lain
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.vStartDoc = EnumStartDoc.PPH21Pesangon Then
                DummyPrimary = DTA.ScalarQuery_GetPrimarySalary(Me.TxtPrefix.Text, Me.TxtNoUrut.Text, Me.vPrimarySalary)
                If DummyPrimary.HasValue = True Then
                    If DummyPrimary.Value <> Me.vPrimarySalary Then
                        MsgBox("Salary document: " + Me.ObjSalary.GetNoTransaksi(DummyPrimary.Value) + " use the same tax document no, please change the specified document no.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If
            Else
                Dummy = DTA.ScalarQuery_GetNoTrans_ByUrutanFormA1(Me.TxtNoUrut.Text, Me.ObjCD.GetLastDateOfTheYear(Me.vTglSalary), _
                Me.ObjCD.GetFisrtDateOfTheYear(Me.vTglSalary))
                If String.IsNullOrEmpty(Dummy) = False Then
                    RefDocNo = Me.ObjSalary.GetNoTransaksi(Me.vPrimarySalary)

                    If RefDocNo <> Dummy Then
                        MsgBox("Salary document: " + Dummy + " (" + Me.ObjEmployee.GetFullNameByPrimary(Me.ObjSalary.GetPrimaryEmployee(Dummy)) + ") use the same A1 Form no, please change the specified document no.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If
            End If

            'konfirmasi
            If MsgBox("Save data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                'simpan perubahan data
                Me.vUrutanBuktiPotong = Me.TxtNoUrut.Text
                If Me.vStartDoc = EnumStartDoc.PPH21Pesangon Then
                    DTA.UpdateQuery(Me.vUrutanBuktiPotong, Me.vPrefixBuktiPotong, Me.vPrimarySalary)
                Else
                    DTA.UpdateQuery1(Me.vPrefixBuktiPotong + "-" + Me.vUrutanBuktiPotong, Me.vUrutanBuktiPotong, Me.vPrimarySalary)
                End If
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Export_eSPT_Edit_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.TxtNoUrut.Text = Me.vUrutanBuktiPotong
            Me.TxtPrefix.Text = Me.vPrefixBuktiPotong
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNoUrut_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoUrut.GotFocus
        Try
            Me.TxtNoUrut.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNoUrut_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoUrut.LostFocus
        Try
            Me.TxtNoUrut.BackColor = Me.TxtRef.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class