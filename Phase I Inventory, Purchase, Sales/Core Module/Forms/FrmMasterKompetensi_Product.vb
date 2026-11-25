Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmMasterKompetensi_Product
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim ObjCD As New ClsCekData
    Dim Obj As New ClsMasterKompetensi_Product
    Dim ObjItem As New ClsMasterKompetensi_Product_Item
    Dim bufferCode As String = ""
    Dim buffer As String = ""
    Dim EraseRowIndex As Integer

    Private Sub FrmMasterCustomData_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.cboDataType.SelectedIndex = 0
            Me.cboDefValue.SelectedIndex = 0
            Me.txtDefPrec.Enabled = False
            Me.TxtDefValueDigit.BringToFront()
            'Me.Size = New System.Drawing.Size(510, 335)
            'Me.GboItemValue.Visible = False
            Me.DgvItemValue.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ClearObjectDefaultValue(ByVal Semua As Boolean, _
    Optional ByVal IndexCBOData As Integer = -1)
        Try
            If Semua = True Then
                Me.txtDefPrec.Text = "0"
                Me.TxtDefValueDigit.Text = "0"
                Me.TxtDefValueText.Clear()
                Me.DtpDefValue.Value = Now
                Me.cboDefValue.SelectedIndex = 0
                Me.cboAlfabetic.SelectedIndex = 0
                Me.TxtIdentifier.Text = "-1"
                Me.CboDropDownList.Items.Clear()
                Me.CboDropDownList.Items.Add("n/a")
            Else
                If IndexCBOData = 0 Then
                    Me.TxtDefValueDigit.Text = "0"
                End If
                If IndexCBOData = 1 Then
                    Me.DtpDefValue.Value = Now
                End If
                If IndexCBOData = 2 Then
                    Me.TxtDefValueText.Clear()
                End If
                If IndexCBOData = 3 Then
                    Me.cboDefValue.SelectedIndex = 0
                End If
                If IndexCBOData = 4 Then
                    Me.cboAlfabetic.SelectedIndex = 0
                End If
                If IndexCBOData = 5 Then
                    Me.CboDropDownList.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboDataType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDataType.SelectedIndexChanged
        Try
            If cboDataType.SelectedIndex = CInt(ClsMasterKompetensi_Product.EnumDataType.Number) Then
                Me.txtDefPrec.Enabled = True
                Me.txtDefPrec.TabStop = True
                Me.ClearObjectDefaultValue(True)
                Me.TxtDefValueDigit.BringToFront()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = True
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = False
                Me.cboAlfabetic.TabStop = False
                Me.CboDropDownList.TabStop = False
                Me.DgvItemValue.TabStop = False

                Me.txtDefPrec.Visible = True
                Me.TxtDefValueDigit.Visible = True
                Me.LblDecimal.Visible = True
                Me.TxtDefValueText.Visible = False
                Me.cboDefValue.Visible = False
                Me.DtpDefValue.Visible = False
                Me.cboAlfabetic.Visible = False

                Me.CboDropDownList.Visible = False
                'Me.Size = New System.Drawing.Size(510, 335)
                'Me.GboItemValue.Visible = False
                Me.DgvItemValue.Rows.Clear()
                Me.DgvItemValue.ReadOnly = True
                Me.DgvItemValue.AllowUserToAddRows = False

                Exit Sub
            End If
            If cboDataType.SelectedIndex = CInt(ClsMasterKompetensi_Product.EnumDataType.Tanggal) Then
                Me.txtDefPrec.Enabled = False
                Me.txtDefPrec.TabStop = False
                Me.ClearObjectDefaultValue(True)
                Me.DtpDefValue.BringToFront()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = True
                Me.cboDefValue.TabStop = False
                Me.cboAlfabetic.TabStop = False
                Me.CboDropDownList.TabStop = False
                Me.DgvItemValue.TabStop = False

                Me.txtDefPrec.Visible = False
                Me.TxtDefValueDigit.Visible = False
                Me.LblDecimal.Visible = False
                Me.TxtDefValueText.Visible = False
                Me.cboDefValue.Visible = False
                Me.DtpDefValue.Visible = True
                Me.cboAlfabetic.Visible = False

                Me.CboDropDownList.Visible = False
                'Me.Size = New System.Drawing.Size(510, 335)
                'Me.GboItemValue.Visible = False
                Me.DgvItemValue.Rows.Clear()
                Me.DgvItemValue.ReadOnly = True
                Me.DgvItemValue.AllowUserToAddRows = False

                Exit Sub
            End If
            If cboDataType.SelectedIndex = CInt(ClsMasterKompetensi_Product.EnumDataType.Text) Then
                Me.txtDefPrec.Enabled = False
                Me.txtDefPrec.TabStop = False
                Me.ClearObjectDefaultValue(True)
                Me.TxtDefValueText.BringToFront()
                Me.TxtDefValueText.TabStop = True
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = False
                Me.cboAlfabetic.TabStop = False
                Me.CboDropDownList.TabStop = False
                Me.DgvItemValue.TabStop = False

                Me.txtDefPrec.Visible = False
                Me.TxtDefValueDigit.Visible = False
                Me.LblDecimal.Visible = False
                Me.TxtDefValueText.Visible = True
                Me.cboDefValue.Visible = False
                Me.DtpDefValue.Visible = False
                Me.cboAlfabetic.Visible = False

                Me.CboDropDownList.Visible = False
                'Me.Size = New System.Drawing.Size(510, 335)
                'Me.GboItemValue.Visible = False
                Me.DgvItemValue.Rows.Clear()
                Me.DgvItemValue.ReadOnly = True
                Me.DgvItemValue.AllowUserToAddRows = False

                Exit Sub
            End If
            If cboDataType.SelectedIndex = CInt(ClsMasterKompetensi_Product.EnumDataType.YN) Then
                Me.txtDefPrec.Enabled = False
                Me.txtDefPrec.TabStop = False
                Me.ClearObjectDefaultValue(True)
                Me.cboDefValue.BringToFront()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = True
                Me.cboAlfabetic.TabStop = False
                Me.CboDropDownList.TabStop = False
                Me.DgvItemValue.TabStop = False

                Me.txtDefPrec.Visible = False
                Me.TxtDefValueDigit.Visible = False
                Me.LblDecimal.Visible = False
                Me.TxtDefValueText.Visible = False
                Me.cboDefValue.Visible = True
                Me.DtpDefValue.Visible = False
                Me.cboAlfabetic.Visible = False

                Me.CboDropDownList.Visible = False
                'Me.Size = New System.Drawing.Size(510, 335)
                'Me.GboItemValue.Visible = False
                Me.DgvItemValue.Rows.Clear()
                Me.DgvItemValue.ReadOnly = True
                Me.DgvItemValue.AllowUserToAddRows = False
                Exit Sub
            End If

            If cboDataType.SelectedIndex = CInt(ClsMasterKompetensi_Product.EnumDataType.Alfabetic) Then
                Me.txtDefPrec.Enabled = False
                Me.txtDefPrec.TabStop = False
                Me.ClearObjectDefaultValue(True)
                Me.cboDefValue.BringToFront()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = False
                Me.cboAlfabetic.TabStop = True
                Me.CboDropDownList.TabStop = False
                Me.DgvItemValue.TabStop = False

                Me.txtDefPrec.Visible = False
                Me.TxtDefValueDigit.Visible = False
                Me.LblDecimal.Visible = False
                Me.TxtDefValueText.Visible = False
                Me.cboDefValue.Visible = False
                Me.DtpDefValue.Visible = False
                Me.cboAlfabetic.Visible = True

                Me.CboDropDownList.Visible = False
                'Me.Size = New System.Drawing.Size(510, 335)
                'Me.GboItemValue.Visible = False
                Me.DgvItemValue.Rows.Clear()
                Me.DgvItemValue.ReadOnly = True
                Me.DgvItemValue.AllowUserToAddRows = False

                Exit Sub
            End If

            If cboDataType.SelectedIndex = CInt(ClsMasterKompetensi_Product.EnumDataType.DropdownList) Then
                Me.txtDefPrec.Enabled = False
                Me.txtDefPrec.TabStop = False
                Me.ClearObjectDefaultValue(True)
                Me.CboDropDownList.BringToFront()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = False
                Me.cboAlfabetic.TabStop = False
                Me.CboDropDownList.TabStop = True
                Me.DgvItemValue.TabStop = True

                Me.txtDefPrec.Visible = False
                Me.TxtDefValueDigit.Visible = False
                Me.LblDecimal.Visible = False
                Me.TxtDefValueText.Visible = False
                Me.cboDefValue.Visible = False
                Me.DtpDefValue.Visible = False
                Me.cboAlfabetic.Visible = False
                Me.CboDropDownList.Visible = True
                Me.CboDropDownList.SelectedIndex = 0

                'Me.Size = New System.Drawing.Size(510, 475)
                Me.GboItemValue.Visible = True
                If Me.btnOK.Enabled = True Then
                    Me.DgvItemValue.ReadOnly = False
                    Me.DgvItemValue.AllowUserToAddRows = True
                End If

                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub GetValueCBO(ByVal CBOIndex As Integer)
        Try
            If CBOIndex = 0 Then
                Me.txtDefPrec.Enabled = True
                Me.txtDefPrec.TabStop = True
                Me.ClearObjectDefaultValue(True)
                Me.TxtDefValueDigit.BringToFront()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = True
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = False
            End If
            If CBOIndex = 1 Then
                Me.txtDefPrec.Enabled = False
                Me.txtDefPrec.TabStop = False
                Me.ClearObjectDefaultValue(True)
                Me.DtpDefValue.BringToFront()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = True
                Me.cboDefValue.TabStop = False
            End If
            If CBOIndex = 2 Then
                Me.txtDefPrec.Enabled = False
                Me.txtDefPrec.TabStop = False
                Me.ClearObjectDefaultValue(True)
                Me.TxtDefValueText.BringToFront()
                Me.TxtDefValueText.TabStop = True
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = False
            End If
            If CBOIndex = 3 Then
                Me.txtDefPrec.Enabled = False
                Me.txtDefPrec.TabStop = False
                Me.ClearObjectDefaultValue(True)
                Me.cboDefValue.BringToFront()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=cien-asus;user id=root;Password=dichangputsa_s!+4T*r@_cien_manjushri;persist security info=True;database=smess_lmb"
        End Try
    End Function

    Private Sub AktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            'ModCekData.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            If Me.cboDataType.SelectedIndex = 0 Then
                Me.txtDefPrec.Enabled = True
            Else
                Me.txtDefPrec.Enabled = False
            End If

            'jika sudah pernah digunakan maka data type tidak boleh di edit
            If Me.Obj.ApakahSudahDigunakan(Me.txtPrimaryStockIn.Text, False) = True Then
                Me.cboDataType.Enabled = False
            Else
                Me.cboDataType.Enabled = True
            End If
            'If Me.dgvItem.Rows.Count = 0 Then
            'Me.btnUP.Enabled = False
            'Me.btnDown.Enabled = False
            'End If
            'Me.dgvItem.ReadOnly = False
            'Me.dgvInput.ReadOnly = False

            ''@rosadinnn
            If Me.baru = False Then
                If Me.Obj.IsAlreadyUse(Me.txtPrimaryStockIn.Text) = True Then
                    Me.cboDataType.Enabled = False
                End If
            End If

            If cboDataType.SelectedIndex = CInt(ClsMasterKompetensi_Product.EnumDataType.DropdownList) Then
                Me.DgvItemValue.ReadOnly = False
                Me.DgvItemValue.AllowUserToAddRows = True
            Else
                Me.DgvItemValue.ReadOnly = True
                Me.DgvItemValue.AllowUserToAddRows = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.CancelButton = Me.btnExit
            Me.btnCancel.DialogResult = Nothing
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            'ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, True, False, True, True, False, False, False, False)
            'Me.dgvItem.ReadOnly = True
            'Me.dgvInput.ReadOnly = True
            Me.DgvItemValue.AllowUserToAddRows = False
            Me.DgvItemValue.ReadOnly = True
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            Me.txtPrimaryStockIn.Text = "-1"
            Me.DtpDefValue.Value = Now
            Me.cboDataType.SelectedIndex = 0
            Me.cboDefValue.SelectedIndex = 0
            Me.txtDefPrec.Text = "0"
            Me.TxtDefValueDigit.Text = "0"
            Me.TxtIdentifier.Text = "-1"

            'secara default, untuk versi awal, semua data competence dinilai dengan alfabetic
            'Me.cboDataType.SelectedIndex = CInt(ClsMasterKompetensi_Product.EnumDataType.Alfabetic)
            Me.CboDropDownList.Items.Clear()
            Me.CboDropDownList.Items.Add("n/a")
            Me.DgvItemValue.Rows.Clear()
            Me.ObjCD.SetNumber(Me.DgvItemValue)

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim ParamStr As String
        Dim ParamInt As Integer
        Dim CekAngka As String
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Data Code")
            KoleksiTextBox.Add(Me.txtDataCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.txtDescription)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'tidak boleh ada kode yang sama
            If Me.baru = True Then
                If Me.Obj.ApakahSudahAda(Me.txtDataCode.Text) = True Then
                    MsgBox("Code duplication, please change code data.", _
                        MsgBoxStyle.Information, "Information")
                    Me.txtDataCode.Focus()
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.Obj.GetDataCode(Me.txtPrimaryStockIn.Text)
                If CodeLama <> Me.txtDataCode.Text Then
                    If Me.Obj.ApakahSudahAda(Me.txtDataCode.Text) = True Then
                        MsgBox("Code duplication, please change code data.", _
                            MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            If String.IsNullOrEmpty(Me.txtRemarks.Text) = False Then
                If ObjCD.CekDataKosong(Me.txtRemarks.Text) = False Then
                    MsgBox("Data remarks" _
                    + " tidak boleh memuat karakter: ' , % , atau * .", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'Periksa length
            If Me.baru = True Then
                If Me.cboDataType.SelectedIndex = 0 Then
                    If Me.TxtDefValueDigit.Text.IndexOf(".") > 0 Then
                        'cek length
                        If Me.TxtDefValueDigit.Text.Substring(Me.TxtDefValueDigit.Text.IndexOf(".")).Length - 1 > Me.txtDefPrec.Text Then
                            MsgBox("Data value's precision isn't appropriate with data default precision", MsgBoxStyle.Information, "Information")
                            Me.txtDefPrec.Focus()
                            Return False
                        End If
                    End If
                End If
            End If

            If Me.baru = False Then
                If Me.cboDataType.SelectedIndex = 0 Then
                    If Obj.GetDataType(Me.txtPrimaryStockIn.Text) = 0 Then
                        If CDbl(Obj.GetDefPrec(Me.txtPrimaryStockIn.Text)) > CDbl(Me.txtDefPrec.Text) Then
                            ParamStr = Obj.GetValue(Me.txtPrimaryStockIn.Text)
                            ParamInt = ParamStr.Substring(ParamStr.IndexOf(".")).Length
                            CekAngka = ParamStr.Substring(ParamStr.IndexOf("."), ParamInt)
                            Dim str As String = ""
                            Dim IdxAkhir As Integer = -1
                            str = CekAngka.Substring(1, ParamInt - 1)
                            IdxAkhir = str.Substring(str.Length - 1)
                            If CInt(IdxAkhir) = 0 Then
                                If Me.txtDefPrec.Text < str.Length - 1 Then
                                    MsgBox("Data default precision can't edit smaller than " + CDbl(Obj.GetDefPrec(Me.txtPrimaryStockIn.Text) - 1).ToString + ".", MsgBoxStyle.Information, "Information")
                                    Me.txtDefPrec.Focus()
                                    Return False
                                End If
                            Else
                                If Me.txtDefPrec.Text < str.Length Then
                                    MsgBox("Data default precision can't edit smaller than " + Obj.GetDefPrec(Me.txtPrimaryStockIn.Text).ToString + ".", MsgBoxStyle.Information, "Information")
                                    Me.txtDefPrec.Focus()
                                    Return False
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            ''Aul 11/05/2015
            If Me.cboDataType.SelectedIndex = CInt(ClsMasterKompetensi_Product.EnumDataType.DropdownList) Then
                'Item tidak boleh kosong 
                If Me.DgvItemValue.Rows.Count <= 1 Then
                    MsgBox("n/a", MsgBoxStyle.Information, "Information ")
                    Return False
                End If

                'Dim RowDGV As DataGridViewRow
                'Dim Retval As Double = -1
                'For ij As Integer = 0 To Me.DgvItemValue.Rows.Count - 2
                '    RowDGV = Me.DgvItemValue.Rows(ij)
                '    If RowDGV.Cells(Me.ColCode.Index).Value = "" Then
                '        MsgBox("Code Item cannot be an empty string.", MsgBoxStyle.Information, "Information")
                '        Return False
                '    ElseIf RowDGV.Cells(Me.ColDesc.Index).Value = "" Then
                '        MsgBox("Description Item cannot be an empty string.", MsgBoxStyle.Information, "Information")
                '        Return False
                '    End If
                '    RowDGV.Selected = True
                'Next

                Dim rowDgv2 As DataGridViewRow
                For Each rowDgv2 In Me.DgvItemValue.Rows
                    If rowDgv2.Cells(Me.ColCode.Index).Value <> "" Then
                        If String.IsNullOrEmpty(Trim(rowDgv2.Cells(Me.ColDesc.Index).Value)) = True Then
                            Me.DgvItemValue.ClearSelection()
                            rowDgv2.Selected = True
                            MsgBox("Description Item cannot empty string.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                Next
            End If

            ''----------------
            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.Obj
                .DataCode = Me.txtDataCode.Text
                .Description = Me.txtDescription.Text
                If Me.cboDataType.SelectedIndex = 0 Then
                    .DataType = ClsMasterKompetensi_Product.EnumDataType.Number
                    .DefaultPrecision = ObjCD.DigitGrouping(Me.txtDefPrec.Text, 0)
                    .Value = ObjCD.DigitGrouping(Me.TxtDefValueDigit.Text, Me.txtDefPrec.Text).ToString
                ElseIf Me.cboDataType.SelectedIndex = 1 Then
                    .DataType = ClsMasterKompetensi_Product.EnumDataType.Tanggal
                    .DefaultPrecision = 0
                    .Value = Format(Me.DtpDefValue.Value, "MM/dd/yyyy").ToString
                ElseIf Me.cboDataType.SelectedIndex = 2 Then
                    .DataType = ClsMasterKompetensi_Product.EnumDataType.Text
                    .DefaultPrecision = 0
                    .Value = Me.TxtDefValueText.Text
                ElseIf Me.cboDataType.SelectedIndex = 3 Then
                    .DataType = ClsMasterKompetensi_Product.EnumDataType.YN
                    .DefaultPrecision = 0
                    .Value = Me.cboDefValue.SelectedIndex
                ElseIf Me.cboDataType.SelectedIndex = 4 Then
                    .DataType = ClsMasterKompetensi_Product.EnumDataType.Alfabetic
                    .DefaultPrecision = 0
                    .Value = Me.cboAlfabetic.SelectedIndex
                ElseIf Me.cboDataType.SelectedIndex = 5 Then
                    .DataType = ClsMasterKompetensi_Product.EnumDataType.DropdownList
                    .DefaultPrecision = 0
                    .Value = Me.TxtIdentifier.Text
                End If

                .Remarks = Me.txtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.txtDataCode.Text = .DataCode
                Me.txtDescription.Text = .Description
                Me.txtRemarks.Text = .Remarks
                Me.txtPrimaryStockIn.Text = .Primary
                If .DataType = ClsMasterKompetensi_Product.EnumDataType.Number Then
                    Me.cboDataType.SelectedIndex = 0
                ElseIf .DataType = ClsMasterKompetensi_Product.EnumDataType.Tanggal Then
                    Me.cboDataType.SelectedIndex = 1
                ElseIf .DataType = ClsMasterKompetensi_Product.EnumDataType.Text Then
                    Me.cboDataType.SelectedIndex = 2
                ElseIf .DataType = ClsMasterKompetensi_Product.EnumDataType.YN Then
                    Me.cboDataType.SelectedIndex = 3
                ElseIf .DataType = ClsMasterKompetensi_Product.EnumDataType.Alfabetic Then
                    Me.cboDataType.SelectedIndex = 4
                ElseIf .DataType = ClsMasterKompetensi_Product.EnumDataType.DropdownList Then
                    Me.cboDataType.SelectedIndex = 5
                End If

                If .DataType = ClsMasterKompetensi_Product.EnumDataType.Number Then
                    Me.GetValueCBO(Me.cboDataType.SelectedIndex)
                    Me.txtDefPrec.Text = ObjCD.DigitGrouping(.DefaultPrecision, 0)
                    Me.TxtDefValueDigit.Text = ObjCD.DigitGrouping(.Value, .DefaultPrecision)
                ElseIf .DataType = ClsMasterKompetensi_Product.EnumDataType.Tanggal Then
                    Me.GetValueCBO(Me.cboDataType.SelectedIndex)
                    Me.txtDefPrec.Text = "0"
                    Me.DtpDefValue.Value = CDate(.Value)
                ElseIf .DataType = ClsMasterKompetensi_Product.EnumDataType.Text Then
                    Me.GetValueCBO(Me.cboDataType.SelectedIndex)
                    Me.txtDefPrec.Text = "0"
                    Me.TxtDefValueText.Text = .Value
                ElseIf .DataType = ClsMasterKompetensi_Product.EnumDataType.YN Then
                    Me.GetValueCBO(Me.cboDataType.SelectedIndex)
                    Me.txtDefPrec.Text = "0"
                    Me.cboDefValue.SelectedIndex = .Value
                ElseIf .DataType = ClsMasterKompetensi_Product.EnumDataType.Alfabetic Then
                    Me.GetValueCBO(Me.cboDataType.SelectedIndex)
                    Me.txtDefPrec.Text = "0"
                    Me.cboAlfabetic.SelectedIndex = .Value
                ElseIf .DataType = ClsMasterKompetensi_Product.EnumDataType.DropdownList Then
                    Me.GetValueCBO(Me.cboDataType.SelectedIndex)
                    Me.txtDefPrec.Text = "0"
                    Me.TxtIdentifier.Text = .Value
                    ListDropDownList()
                    'Me.CboDropDownList.SelectedIndex = .Value
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                'Me.KosongkanItem()
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.DgvItemValue, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Select()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.DgvItemValue, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Select()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtDataCode.Text = "" And Me.txtDescription.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Select()
                Me.KosongkanField()
                ''Tambahan
                Me.txtDefPrec.Enabled = False
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                        Me.ClearObjectDefaultValue(True)
                    Else
                        Me.ClearObjectDefaultValue(True)
                        Me.Obj.CariDetail(Me.txtPrimaryStockIn.Text, Me.DgvItemValue, Me, Me.ObjCD)
                        Me.SalinKeField()
                        'Me.KosongkanItem()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Select()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.txtDataCode.Focus()
            Me.txtDataCode.BackColor = Me.BackColorAll

            'secara default, untuk versi awal, semua data competence dinilai dengan alfabetic
            Me.cboDataType.SelectedIndex = CInt(ClsMasterKompetensi_Product.EnumDataType.Alfabetic)

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim i As New FrmMasterKompetensi_Product_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            i.ShowDialog()
            If Not i.mPrimary = -1 Then
                If Me.Obj.CariDetail(i.mPrimary, Me.DgvItemValue, Me, Me.ObjCD) = True Then
                    Me.SalinKeField()
                    Me.DgvItemValue.AllowUserToAddRows = False
                    Me.txtDefPrec.Enabled = False
                    Me.btnEdit.Select()
                Else
                    Me.btnFind.Select()
                End If
                Me.btnFind.Select()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.txtDataCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtDataCode.Text <> "" Then
                If Me.Obj.ApakahBolehEdit(Me.txtPrimaryStockIn.Text, Me.UserCode, True, True) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    'Me.dgvItem.Focus()
                    'nilai min juga harus diatas tanggal pengkuncian data terakhir
                    'msgbox("belum")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtDataCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtDataCode.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.txtPrimaryStockIn.Text, Me.UserCode, True, True) = False Then
                    'MsgBox("You cannot delete this data, because this will cause negative quantity data.", _
                    'MsgBoxStyle.Information, "Information")
                    Me.btnFind.Select()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.Hapus(Me.txtPrimaryStockIn.Text, Me.UserCode, True, True) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Select()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtRemarks.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me.UserCode, Me.DgvItemValue, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Select()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me.DgvItemValue, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Select()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDataCode_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDataCode.GotFocus
        Try
            Me.txtDataCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDataCode_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDataCode.LostFocus
        Try
            Me.txtDataCode.BackColor = Me.txtPrimaryStockIn.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefValueDigit_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueDigit.Enter
        Try
            If Me.txtRemarks.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.TxtDefValueDigit, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefValueDigit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueDigit.GotFocus
        Try
            Me.TxtDefValueDigit.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefValueDigit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueDigit.LostFocus
        Try
            Me.TxtDefValueDigit.BackColor = Me.txtPrimaryStockIn.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueText.GotFocus
        Try
            Me.TxtDefValueText.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueText.LostFocus
        Try
            Me.TxtDefValueText.BackColor = Me.txtPrimaryStockIn.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRollNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDefPrec.GotFocus
        Try
            Me.txtDefPrec.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRollNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDefPrec.LostFocus
        Try
            Me.txtDefPrec.BackColor = Me.txtPrimaryStockIn.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtKodeLokasi_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.GotFocus
        Try
            Me.txtDescription.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtKodeLokasi_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.LostFocus
        Try
            Me.txtDescription.BackColor = Me.txtPrimaryStockIn.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.GotFocus
        Try
            Me.txtRemarks.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.LostFocus
        Try
            Me.txtRemarks.BackColor = Me.txtPrimaryStockIn.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDefPrec_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDefPrec.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.txtDefPrec, sender, e, 5)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDefPrec_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDefPrec.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.txtDefPrec, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefValueDigit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDefValueDigit.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.TxtDefValueDigit, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefValueDigit_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDefValueDigit.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.TxtDefValueDigit, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDefPrec_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDefPrec.Leave

        Try
            If Me.txtRemarks.Enabled = True Then
                Me.TxtDefValueDigit.Text = Me.ObjCD.DigitGrouping(Me.TxtDefValueDigit.Text, Me.txtDefPrec.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefValueDigit_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtDefValueDigit.TextChanged

    End Sub

    Private Sub TxtDefValueDigit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueDigit.Leave
        Try
            If Me.txtRemarks.Enabled = True Then
                Me.TxtDefValueDigit.Text = Me.ObjCD.DigitGrouping(Me.TxtDefValueDigit.Text, Me.txtDefPrec.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    ''Aul 9 nov 2015
    Private Sub ListDropDownList()
        Dim Row As DataGridViewRow
        Dim ListRowDgv As New List(Of String)
        Dim ListRowDgv3 As New List(Of String)
        Dim Baris As Integer = -1

        Try
            For ij As Integer = 0 To Me.DgvItemValue.Rows.Count - 1
                Row = Me.DgvItemValue.Rows(ij)
                If Me.DgvItemValue.Rows.Count - 1 Then
                    If Row.Cells(Me.ColCode.Index).Value <> "" Then
                        ListRowDgv.Add(Row.Cells(Me.ColCode.Index).Value)
                        ListRowDgv3.Add(Row.Cells(Me.ColIdentifier.Index).Value)
                    End If
                End If
            Next

            If ListRowDgv.Count > 0 Then
                Dim ab As String = ""
                Me.CboDropDownList.Items.Clear()
                For i As Integer = 0 To ListRowDgv.Count - 1
                    ab = ListRowDgv.Item(i)
                    Me.CboDropDownList.Items.Add(ab)
                Next
            End If

            If ListRowDgv3.Count > 0 Then
                Dim ac As String = ""
                For i As Integer = 0 To ListRowDgv.Count - 1
                    ac = ListRowDgv3.Item(i)
                    Baris += 1
                    If Me.TxtIdentifier.Text = ac Then
                        Me.CboDropDownList.SelectedIndex = Baris
                    End If
                Next
            End If

            If ListRowDgv.Count = 0 Then
                Me.CboDropDownList.Items.Clear()
                Me.CboDropDownList.Items.Add("n/a")
                Me.CboDropDownList.SelectedIndex = 0
            ElseIf ListRowDgv.Count = 1 Then
                Me.CboDropDownList.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboDropDownList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboDropDownList.SelectedIndexChanged
        Dim Row As DataGridViewRow
        Dim ListRowDgv2 As New List(Of String)
        Dim Baris As Integer = 0
        Dim Jml As Integer = -1
        Try
            For ij As Integer = 0 To Me.DgvItemValue.Rows.Count - 1
                Row = Me.DgvItemValue.Rows(ij)
                If Me.DgvItemValue.Rows.Count - 1 Then
                    If Row.Cells(Me.ColCode.Index).Value <> "" Then
                        ListRowDgv2.Add(Row.Cells(Me.ColIdentifier.Index).Value)
                    End If
                End If
            Next
            Baris = Me.CboDropDownList.SelectedIndex
            If ListRowDgv2.Count > 0 Then
                Dim cd As String = ""
                For i As Integer = 0 To ListRowDgv2.Count - 1
                    Jml += 1
                    cd = ListRowDgv2.Item(i)
                    Me.TxtIdentifier.Text = cd
                    If Baris = Jml Then
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemCode(ByVal vCode As String, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim Row As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvItemValue.Rows.Count - 2
                Row = Me.DgvItemValue.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If Row.Cells(Me.ColCode.Index).Value = vCode Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvItemValue_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItemValue.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvItemValue(Me.ColCode.Index, e.RowIndex).Value) = False Then
                If (e.ColumnIndex = Me.ColCode.Index Or e.ColumnIndex = Me.ColDesc.Index Or e.ColumnIndex = Me.ColRemarks.Index) Then
                    Me.buffer = Me.DgvItemValue.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = False
                End If
            Else
                If e.ColumnIndex = Me.ColCode.Index Then
                    Me.buffer = Me.DgvItemValue.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemValueValue_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItemValue.CellEndEdit
        Dim index As Integer = -1
        Try
            Dim row As DataGridViewRow = Me.DgvItemValue.Rows(e.RowIndex)
            With row
                If IsNothing(Me.DgvItemValue.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCode.Index Then
                        If e.ColumnIndex = Me.ColCode.Index Then
                            Me.DgvItemValue.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvItemValue.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        End If
                        index = GetIndexItemCode(Me.DgvItemValue.Item(Me.ColCode.Index, e.RowIndex).Value, e.RowIndex)
                        If index >= 0 Then
                            MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                            Me.DgvItemValue.Item(Me.ColCode.Index, e.RowIndex).Value = ""
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            Else
                                Me.DgvItemValue.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Exit Sub
                        End If
                    End If
                    If e.ColumnIndex = Me.ColCode.Index Or e.ColumnIndex = Me.ColDesc.Index Then
                        If String.IsNullOrEmpty(Trim(Me.DgvItemValue.Item(e.ColumnIndex, e.RowIndex).Value)) = True Then
                            Me.DgvItemValue.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColCode.Index Then
                        If row.Cells(Me.ColDesc.Index).Value <> "" Then
                            Me.DgvItemValue.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.EraseRowIndex = row.Index
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
                    If e.ColumnIndex = Me.ColDesc.Index Then
                        Me.DgvItemValue.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End With

            ListDropDownList()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemValue_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvItemValue.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            If Me.txtDescription.Enabled = True Then
                rowDGV = Me.DgvItemValue.Rows(Me.DgvItemValue.RowCount - 1)
                With rowDGV
                    .Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColCode.Index).Value = ""
                    .Cells(Me.ColDesc.Index).Value = ""
                    .Cells(Me.ColPrimaryKompetensi.Index).Value = "-1"
                    .Cells(Me.ColIdentifier.Index).Value = Me.ObjItem.GetNextIdentifier(Me.DgvItemValue, Me.txtPrimaryStockIn.Text, Me)
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemValue_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItemValue.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvItemValue)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemValue_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvItemValue.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.DgvItemValue)
            ListDropDownList()
            If Me.CboDropDownList.Text = "" Then
                Me.CboDropDownList.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemValue_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DgvItemValue.UserDeletingRow
        Dim RowDGV As DataGridViewRow = e.Row
        Try
            If Me.Obj.ApakahBolehHapusItem(Me.txtPrimaryStockIn.Text, RowDGV.Cells(Me.ColIdentifier.Index).Value, Me.UserCode, True, True) = False Then
                e.Cancel = True
                Exit Sub
            End If
            'ListDropDownList()
            'Me.ObjCD.SetNumber(Me.DgvItemValue)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.DgvItemValue.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.DgvItemValue)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    ''----------------------------------
End Class