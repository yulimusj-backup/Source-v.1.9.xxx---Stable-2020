Option Explicit On
Option Strict Off

Public Class ClsTransaksiSettingTransf_AOP
    Public Enum Enum_Tipe_Setting
        None
        Debet
        Kredit
    End Enum

    Public Function GetJml_ByPrimaryAkun(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseTransaksiSettingTransf_AOPTableAdapters.tbltrans_setting_transf_aopTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlByPrimaryAkun(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByPrimaryAkun(vPrimaryAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanData(ByVal vFrm As FrmTransaksiSettingAOP, ByVal vdgv As DataGridView, _
   ByVal vUserCode As String, ByVal vtsl_CreatedBy As ToolStripStatusLabel, ByVal vtsl_CreatedDate As ToolStripStatusLabel) As Boolean
        Dim DTA As New DseTransaksiSettingTransf_AOPTableAdapters.tbltrans_setting_transf_aopTableAdapter
        Dim row_dgv As DataGridViewRow
        Dim Purchase_Int, VAT_Int, Dutystamp_Int, TotalCost_Int, DP_Int As Enum_Tipe_Setting
        Dim Timestamp As Date

        Try
            'hapus data yang sudah ada terlebih dahulu
            DTA.Delete()
            'simpan data baru
            Timestamp = Now

            For Each row_dgv In vdgv.Rows
                If row_dgv.Index <> vdgv.Rows.Count - 1 Then
                    With row_dgv
                        If .Cells(vFrm.ColPurchase.Index).Value = "-" Then
                            Purchase_Int = Enum_Tipe_Setting.None
                        ElseIf .Cells(vFrm.ColPurchase.Index).Value = "Debet" Then
                            Purchase_Int = Enum_Tipe_Setting.Debet
                        ElseIf .Cells(vFrm.ColPurchase.Index).Value = "Credit" Then
                            Purchase_Int = Enum_Tipe_Setting.Kredit
                        End If
                        If .Cells(vFrm.ColVAT.Index).Value = "-" Then
                            VAT_Int = Enum_Tipe_Setting.None
                        ElseIf .Cells(vFrm.ColVAT.Index).Value = "Debet" Then
                            VAT_Int = Enum_Tipe_Setting.Debet
                        ElseIf .Cells(vFrm.ColVAT.Index).Value = "Credit" Then
                            VAT_Int = Enum_Tipe_Setting.Kredit
                        End If
                        If .Cells(vFrm.ColDutyStamp.Index).Value = "-" Then
                            Dutystamp_Int = Enum_Tipe_Setting.None
                        ElseIf .Cells(vFrm.ColDutyStamp.Index).Value = "Debet" Then
                            Dutystamp_Int = Enum_Tipe_Setting.Debet
                        ElseIf .Cells(vFrm.ColDutyStamp.Index).Value = "Credit" Then
                            Dutystamp_Int = Enum_Tipe_Setting.Kredit
                        End If
                        If .Cells(vFrm.ColTotalCost.Index).Value = "-" Then
                            TotalCost_Int = Enum_Tipe_Setting.None
                        ElseIf .Cells(vFrm.ColTotalCost.Index).Value = "Debet" Then
                            TotalCost_Int = Enum_Tipe_Setting.Debet
                        ElseIf .Cells(vFrm.ColTotalCost.Index).Value = "Credit" Then
                            TotalCost_Int = Enum_Tipe_Setting.Kredit
                        End If
                        If .Cells(vFrm.ColDownpayment.Index).Value = "-" Then
                            DP_Int = Enum_Tipe_Setting.None
                        ElseIf .Cells(vFrm.ColDownpayment.Index).Value = "Debet" Then
                            DP_Int = Enum_Tipe_Setting.Debet
                        ElseIf .Cells(vFrm.ColDownpayment.Index).Value = "Credit" Then
                            DP_Int = Enum_Tipe_Setting.Kredit
                        End If

                        DTA.Insert(.Cells(vFrm.ColPrimaryAkun.Index).Value, Purchase_Int, VAT_Int, Dutystamp_Int, TotalCost_Int, _
                        .Cells(vFrm.ColRemarks.Index).Value, vUserCode, Timestamp, vUserCode, Timestamp, Timestamp, DP_Int)
                    End With
                End If
            Next
            'update data tool strip status label
            vtsl_CreatedBy.Text = "Created By: " + vUserCode
            vtsl_CreatedDate.Text = "Created Date: " + Format(Timestamp, "dd-MMM-yyyy") + "-" + Timestamp.ToShortTimeString

            'sampai sini kembalikan nilai true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function LoadData(ByVal vFrm As FrmTransaksiSettingAOP, ByVal vdgv As DataGridView, _
    ByVal vtsl_CreatedBy As ToolStripStatusLabel, ByVal vtsl_CreatedDate As ToolStripStatusLabel) As Boolean
        Dim DTA As New DseTransaksiSettingTransf_AOPTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseTransaksiSettingTransf_AOP.DataTable1DataTable
        Dim row As DseTransaksiSettingTransf_AOP.DataTable1Row
        Dim row_dgv As DataGridViewRow
        Dim Purchase_Int, VAT_Int, Dutystamp_Int, TotalCost_Int, DP_Int As Enum_Tipe_Setting

        Try
            'isikan data
            DTA.Fill(TBL)

            'kosongkan dgv
            vdgv.Rows.Clear()

            'isikan dgv
            vdgv.AllowUserToAddRows = False
            For Each row In TBL.Rows
                vdgv.Rows.Add()
                row_dgv = vdgv.Rows(vdgv.Rows.Count - 1)
                With row_dgv
                    .Cells(vFrm.ColAccount.Index).Value = row.Account_Code
                    .Cells(vFrm.ColAccount_Desc.Index).Value = row.Description

                    Purchase_Int = row.Purchase
                    VAT_Int = row.VAT
                    Dutystamp_Int = row.Duty_Stamp
                    TotalCost_Int = row.Total_Cost
                    DP_Int = row.DownPayment

                    If Purchase_Int = Enum_Tipe_Setting.Debet Then
                        .Cells(vFrm.ColPurchase.Index).Value = "Debet"
                    ElseIf Purchase_Int = Enum_Tipe_Setting.Kredit Then
                        .Cells(vFrm.ColPurchase.Index).Value = "Credit"
                    Else
                        .Cells(vFrm.ColPurchase.Index).Value = "-"
                    End If
                    If VAT_Int = Enum_Tipe_Setting.Debet Then
                        .Cells(vFrm.ColVAT.Index).Value = "Debet"
                    ElseIf VAT_Int = Enum_Tipe_Setting.Kredit Then
                        .Cells(vFrm.ColVAT.Index).Value = "Credit"
                    Else
                        .Cells(vFrm.ColVAT.Index).Value = "-"
                    End If
                    If Dutystamp_Int = Enum_Tipe_Setting.Debet Then
                        .Cells(vFrm.ColDutyStamp.Index).Value = "Debet"
                    ElseIf Dutystamp_Int = Enum_Tipe_Setting.Kredit Then
                        .Cells(vFrm.ColDutyStamp.Index).Value = "Credit"
                    Else
                        .Cells(vFrm.ColDutyStamp.Index).Value = "-"
                    End If
                    If TotalCost_Int = Enum_Tipe_Setting.Debet Then
                        .Cells(vFrm.ColTotalCost.Index).Value = "Debet"
                    ElseIf TotalCost_Int = Enum_Tipe_Setting.Kredit Then
                        .Cells(vFrm.ColTotalCost.Index).Value = "Credit"
                    Else
                        .Cells(vFrm.ColTotalCost.Index).Value = "-"
                    End If
                    If DP_Int = Enum_Tipe_Setting.Debet Then
                        .Cells(vFrm.ColDownpayment.Index).Value = "Debet"
                    ElseIf DP_Int = Enum_Tipe_Setting.Kredit Then
                        .Cells(vFrm.ColDownpayment.Index).Value = "Credit"
                    Else
                        .Cells(vFrm.ColDownpayment.Index).Value = "-"
                    End If

                    .Cells(vFrm.ColRemarks.Index).Value = row.Remarks
                    .Cells(vFrm.ColPrimaryAkun.Index).Value = row.Primary_Akun
                End With
            Next
            vdgv.AllowUserToAddRows = True

            'isikan data created by
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy) = True Then
                vtsl_CreatedBy.Text = "Created By: "
            Else
                vtsl_CreatedBy.Text = "Created By: " + DTA.ScalarQuery_GetCreatedBy
            End If

            'isikan created date
            If DTA.ScalarQuery_GetCreatedDate.HasValue = True Then
                vtsl_CreatedDate.Text = "Created Date: " + Format(DTA.ScalarQuery_GetCreatedDate.Value, "dd-MMM-yyyy") + "-" + DTA.ScalarQuery_GetCreatedDate.Value.ToShortTimeString
            Else
                vtsl_CreatedDate.Text = "Created Date: "
            End If

            'isikan vdgv_show
            vFrm.Tampilkan_DGV_Akun()

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function
End Class
