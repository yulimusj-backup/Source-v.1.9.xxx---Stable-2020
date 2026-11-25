Public Class FrmLapAsset_DepresiasiPerPeriod
    Public BackcolorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim TipeJurnalDisposal As ClsTransaksiJournal.EnumTipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.AssetDisposal
    Dim report1 As New CrpLapAsset_DepresiasiPerPeriod
    Dim mDefPrec As Integer = 0
    Dim ObCur As New Core_Module.clsMasterCurrency

    Dim DS As New DseLapAsset_DepresiasiPerPeriod
    Dim TBLLaporan As New DseLapAsset_DepresiasiPerPeriod.TblLaporanDataTable
    Dim rowLaporan As DseLapAsset_DepresiasiPerPeriod.TblLaporanRow
    Dim TBLRevMaterial As New DseLapAsset_DepresiasiPerPeriod.TblRev_MaterialUsageDataTable
    Dim TBLRevOther As New DseLapAsset_DepresiasiPerPeriod.TblRev_OtherDataTable
    Dim TBLDisposal As New DseLapAsset_DepresiasiPerPeriod.TblDisposalDataTable
    Dim TBLDepresiasi As New DseLapAsset_DepresiasiPerPeriod._TblDepresiasiDataTable
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Public Sub CreateDepreciationReport2()
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan datanya manual
            '1. TblLaporan
            Me.rowLaporan = Me.TBLLaporan.NewTblLaporanRow
            With Me.rowLaporan
                .Primary_Asset = 1
                .Primary_Asset_Group = 1
                .Depreciation = 75000
            End With
            Me.TBLLaporan.Rows.Add(Me.rowLaporan)
            '2. TblDepresiasi

            '3. TblRevaluasiMaterialUsage
            '4. TblRevaluasiOther
            '5. TblDisposal

            'tampilkan laporan
            Me.report1.SetParameterValue(report1.Parameter_Decimal.ParameterFieldName, Me.mDefPrec)
            Me.crv1.ReportSource = Me.report1

            'tampilkan dalam dgv

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'ini bagian iwan
    Public Sub CreateDepreciationReport()
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'pastikan periode from to bener

            'isikan daftar asset

            'buang asset yang tidak disusutkan

            'buang asset yang sudah disposed pada periode sebelum awal periode

            'isikan property master asset

            'isikan akum dep

            'isikan nilai depresiasi

            'isikan nilai revaluasi

            'isikan nilai disposal

            'isikan nilai buku akhir

            'tampilkan laporan
            Me.report1.SetParameterValue(report1.Parameter_Decimal.ParameterFieldName, Me.mDefPrec)
            Me.crv1.ReportSource = Me.report1

            'tampilkan dalam dgv

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            Me.CreateDepreciationReport2()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLapAsset_DepresiasiPerPeriod_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.mDefPrec = Me.ObCur.GetDecimal_DefaultCurrency
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowGroupTreeButton = False
            Me.crv1.ShowExportButton = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Click
        Dim ijk As New FrmLapAsset_DepresiasiPerPeriod_Explain
        ijk.ShowDialog()
    End Sub
End Class