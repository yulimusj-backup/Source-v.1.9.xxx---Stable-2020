Public Class FrmLaporanQuotationRealisasiChart
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Enum EnumTipeData As Integer
        Quotation_Qty
        Perct_Realized
        Perct_Unrealized
        Perct_Realized_Unrealized_Combined
        Realized_Amount
        Unrealized_Amount
        Total_Amount
        Realized_Unrealized_Amount_Combined
    End Enum

    Private Enum EnumTipeChart As Integer
        Bar
        Stacked_Bar
        Pie
    End Enum

    Private Enum EnumGroupBy As Integer
        Group_By_Sales_Person
        Group_By_Sales_Person_Group
        Group_By_Product_Group
        Group_By_Brand
        Group_By_Sub_Brand
        Group_By_Category
        Group_By_Product
    End Enum

    Private Sub FrmLaporanQuotationRealisasiChart_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.CboData.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTipeData(ByVal vParameter As String) As EnumTipeData
        Try
            If vParameter = "Quotation Qty" Then
                Return EnumTipeData.Quotation_Qty
            ElseIf vParameter = "% Realized" Then
                Return EnumTipeData.Perct_Realized
            ElseIf vParameter = "% Unrealized" Then
                Return EnumTipeData.Perct_Unrealized
            ElseIf vParameter = "% Realized - % Unrealized Combined" Then
                Return EnumTipeData.Perct_Realized_Unrealized_Combined
            ElseIf vParameter = "Realized Amount" Then
                Return EnumTipeData.Realized_Amount
            ElseIf vParameter = "Unrealized Amount" Then
                Return EnumTipeData.Unrealized_Amount
            ElseIf vParameter = "Total Amount" Then
                Return EnumTipeData.Total_Amount
            ElseIf vParameter = "Realized - Unrealized Amount Combined" Then
                Return EnumTipeData.Realized_Unrealized_Amount_Combined
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetGroupBy(ByVal vParameter As String) As EnumGroupBy
        Try
            If vParameter = "Sales Person" Then
                Return EnumGroupBy.Group_By_Sales_Person
            ElseIf vParameter = "Sales Person Group" Then
                Return EnumGroupBy.Group_By_Sales_Person_Group
            ElseIf vParameter = "Product Group" Then
                Return EnumGroupBy.Group_By_Product_Group
            ElseIf vParameter = "Poduct Brand" Then
                Return EnumGroupBy.Group_By_Brand
            ElseIf vParameter = "Product Sub Brand" Then
                Return EnumGroupBy.Group_By_Sub_Brand
            ElseIf vParameter = "Product Category" Then
                Return EnumGroupBy.Group_By_Category
            ElseIf vParameter = "Product Code" Then
                Return EnumGroupBy.Group_By_Product
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetGroupBy(ByVal vParameter As EnumGroupBy) As String
        Try
            If vParameter = EnumGroupBy.Group_By_Sales_Person Then
                Return "Sales Person"
            ElseIf vParameter = EnumGroupBy.Group_By_Sales_Person_Group Then
                Return "Sales Person Group"
            ElseIf vParameter = EnumGroupBy.Group_By_Product_Group Then
                Return "Product Group"
            ElseIf vParameter = EnumGroupBy.Group_By_Brand Then
                Return "Poduct Brand"
            ElseIf vParameter = EnumGroupBy.Group_By_Sub_Brand Then
                Return "Product Sub Brand"
            ElseIf vParameter = EnumGroupBy.Group_By_Category Then
                Return "Product Category"
            ElseIf vParameter = EnumGroupBy.Group_By_Product Then
                Return "Product Code"
            End If

            Return "Sales Person"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Sales Person"
        End Try
    End Function

    Private Function GetTipeChart(ByVal vParameter As String) As EnumTipeChart
        Try
            If vParameter = "Bar" Then
                Return EnumTipeChart.Bar
            ElseIf vParameter = "Stacked Bar" Then
                Return EnumTipeChart.Stacked_Bar
            ElseIf vParameter = "Pie Chart" Then
                Return EnumTipeChart.Stacked_Bar
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetTipeChart(ByVal vParameter As EnumTipeChart) As String
        Try
            If vParameter = EnumTipeChart.Bar Then
                Return "Bar"
            ElseIf vParameter = EnumTipeChart.Stacked_Bar Then
                Return "Stacked Bar"
            ElseIf vParameter = EnumTipeChart.Pie Then
                Return "Pie Chart"
            End If

            Return "Bar"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Bar"
        End Try
    End Function

    Private Sub CboData_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboData.SelectedIndexChanged
        Dim TipeData As EnumTipeData = EnumTipeData.Perct_Realized
        Dim TipeGroupBy As EnumGroupBy = EnumGroupBy.Group_By_Brand
        Dim TipeChart As EnumTipeChart = EnumTipeChart.Bar

        Try
            Me.cboGroupBy.Items.Clear()
            Me.cboChart.Items.Clear()

            TipeData = Me.GetTipeData(Me.CboData.Text)
            If TipeData = EnumTipeData.Quotation_Qty Then
                TipeGroupBy = EnumGroupBy.Group_By_Sales_Person
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Sales_Person_Group
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                Me.cboGroupBy.SelectedIndex = 0

                TipeChart = EnumTipeChart.Bar
                Me.cboChart.Items.Add(Me.GetTipeChart(TipeChart))
                TipeChart = EnumTipeChart.Pie
                Me.cboChart.Items.Add(Me.GetTipeChart(TipeChart))
                Me.cboChart.SelectedIndex = 0
            ElseIf TipeData = EnumTipeData.Perct_Realized Or _
            TipeData = EnumTipeData.Perct_Unrealized Then
                TipeGroupBy = EnumGroupBy.Group_By_Sales_Person
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Sales_Person_Group
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Product_Group
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Brand
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Sub_Brand
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Category
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Product
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                Me.cboGroupBy.SelectedIndex = 0

                TipeChart = EnumTipeChart.Bar
                Me.cboChart.Items.Add(Me.GetTipeChart(TipeChart))
                Me.cboChart.SelectedIndex = 0
            ElseIf TipeData = EnumTipeData.Realized_Amount Or _
            TipeData = EnumTipeData.Unrealized_Amount Or _
            TipeData = EnumTipeData.Total_Amount Then
                TipeGroupBy = EnumGroupBy.Group_By_Sales_Person
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Sales_Person_Group
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Product_Group
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Brand
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Sub_Brand
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Category
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Product
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                Me.cboGroupBy.SelectedIndex = 0

                TipeChart = EnumTipeChart.Bar
                Me.cboChart.Items.Add(Me.GetTipeChart(TipeChart))
                TipeChart = EnumTipeChart.Pie
                Me.cboChart.Items.Add(Me.GetTipeChart(TipeChart))
                Me.cboChart.SelectedIndex = 0
            ElseIf TipeData = EnumTipeData.Perct_Realized_Unrealized_Combined Or _
            TipeData = EnumTipeData.Realized_Unrealized_Amount_Combined Then
                TipeGroupBy = EnumGroupBy.Group_By_Sales_Person
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Sales_Person_Group
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Product_Group
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Brand
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Sub_Brand
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Category
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                TipeGroupBy = EnumGroupBy.Group_By_Product
                Me.cboGroupBy.Items.Add(Me.GetGroupBy(TipeGroupBy))
                Me.cboGroupBy.SelectedIndex = 0

                TipeChart = EnumTipeChart.Bar
                Me.cboChart.Items.Add(Me.GetTipeChart(TipeChart))
                TipeChart = EnumTipeChart.Stacked_Bar
                Me.cboChart.Items.Add(Me.GetTipeChart(TipeChart))
                Me.cboChart.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class