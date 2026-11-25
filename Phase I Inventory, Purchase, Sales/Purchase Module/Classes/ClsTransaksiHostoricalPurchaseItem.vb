
Option Explicit On
Option Strict Off
Public Class ClsTransaksiHostoricalPurchaseItem
    Private mPrimaryTransaksi As Double
    Private mPrimaryProduct As Double
    Private mQty As Double
    Private mAmount As Double
    Private mRemarks As String
    Public Property PrimaryTransaksi() As Double
        Get
            Return mPrimaryTransaksi
        End Get
        Set(ByVal value As Double)
            mPrimaryTransaksi = value
        End Set
    End Property
    Private Property PrimaryProduct() As Double
        Get
            Return mPrimaryProduct
        End Get
        Set(ByVal value As Double)
            mPrimaryProduct = value
        End Set
    End Property
    Private Property Qty() As Double
        Get
            Return mQty
        End Get
        Set(ByVal value As Double)
            mQty = value
        End Set
    End Property
    Private Property Amount() As Double
        Get
            Return mAmount
        End Get
        Set(ByVal value As Double)
            mAmount = value
        End Set
    End Property
    Private Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property
    Public Function Simpan(ByVal dgv As DataGridView, ByVal vprimaryTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchase_itemTableAdapter
        Dim dgvrow As DataGridViewRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            For Each dgvrow In dgv.Rows
                DTA.Insert(vprimaryTransaksi, dgvrow.Cells("ColPrimaryCodeNo").Value, dgvrow.Cells("ColQty").Value, dgvrow.Cells("ColAmount").Value, dgvrow.Cells("ColRemark").Value)
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiHistoricalPurchaseTableAdapters.tbltrans_historicalpurchase_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Delete(vPrimaryTransaksi) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus() As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
End Class
