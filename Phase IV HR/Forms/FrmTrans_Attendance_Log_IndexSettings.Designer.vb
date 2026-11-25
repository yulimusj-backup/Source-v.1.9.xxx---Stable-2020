<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Attendance_Log_IndexSettings
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.year_index = New System.Windows.Forms.TextBox
        Me.year_lenght = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.month_length = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.month_index = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.day_length = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.day_index = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.hour_length = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.hour_index = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.minute_length = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.minute_index = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.user_length = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.user_index = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Button1 = New System.Windows.Forms.Button
        Me.remove_zero = New System.Windows.Forms.CheckBox
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Year Index"
        '
        'year_index
        '
        Me.year_index.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.year_index.Location = New System.Drawing.Point(90, 3)
        Me.year_index.MaxLength = 150
        Me.year_index.Name = "year_index"
        Me.year_index.Size = New System.Drawing.Size(95, 20)
        Me.year_index.TabIndex = 1
        Me.year_index.Text = "0"
        Me.year_index.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'year_lenght
        '
        Me.year_lenght.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.year_lenght.Location = New System.Drawing.Point(90, 29)
        Me.year_lenght.MaxLength = 150
        Me.year_lenght.Name = "year_lenght"
        Me.year_lenght.Size = New System.Drawing.Size(95, 20)
        Me.year_lenght.TabIndex = 3
        Me.year_lenght.Text = "0"
        Me.year_lenght.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Year Length"
        '
        'month_length
        '
        Me.month_length.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.month_length.Location = New System.Drawing.Point(89, 81)
        Me.month_length.MaxLength = 150
        Me.month_length.Name = "month_length"
        Me.month_length.Size = New System.Drawing.Size(95, 20)
        Me.month_length.TabIndex = 7
        Me.month_length.Text = "0"
        Me.month_length.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 84)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Month Length"
        '
        'month_index
        '
        Me.month_index.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.month_index.Location = New System.Drawing.Point(89, 55)
        Me.month_index.MaxLength = 150
        Me.month_index.Name = "month_index"
        Me.month_index.Size = New System.Drawing.Size(95, 20)
        Me.month_index.TabIndex = 5
        Me.month_index.Text = "0"
        Me.month_index.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 58)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(66, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Month Index"
        '
        'day_length
        '
        Me.day_length.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.day_length.Location = New System.Drawing.Point(89, 133)
        Me.day_length.MaxLength = 150
        Me.day_length.Name = "day_length"
        Me.day_length.Size = New System.Drawing.Size(95, 20)
        Me.day_length.TabIndex = 11
        Me.day_length.Text = "0"
        Me.day_length.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(7, 136)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(62, 13)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Day Length"
        '
        'day_index
        '
        Me.day_index.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.day_index.Location = New System.Drawing.Point(89, 107)
        Me.day_index.MaxLength = 150
        Me.day_index.Name = "day_index"
        Me.day_index.Size = New System.Drawing.Size(95, 20)
        Me.day_index.TabIndex = 9
        Me.day_index.Text = "0"
        Me.day_index.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(7, 110)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(55, 13)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "Day Index"
        '
        'hour_length
        '
        Me.hour_length.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.hour_length.Location = New System.Drawing.Point(97, 32)
        Me.hour_length.MaxLength = 150
        Me.hour_length.Name = "hour_length"
        Me.hour_length.Size = New System.Drawing.Size(76, 20)
        Me.hour_length.TabIndex = 15
        Me.hour_length.Text = "0"
        Me.hour_length.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(15, 35)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(66, 13)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Hour Length"
        '
        'hour_index
        '
        Me.hour_index.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.hour_index.Location = New System.Drawing.Point(97, 6)
        Me.hour_index.MaxLength = 150
        Me.hour_index.Name = "hour_index"
        Me.hour_index.Size = New System.Drawing.Size(76, 20)
        Me.hour_index.TabIndex = 13
        Me.hour_index.Text = "0"
        Me.hour_index.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(15, 9)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(59, 13)
        Me.Label8.TabIndex = 12
        Me.Label8.Text = "Hour Index"
        '
        'minute_length
        '
        Me.minute_length.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.minute_length.Location = New System.Drawing.Point(97, 84)
        Me.minute_length.MaxLength = 150
        Me.minute_length.Name = "minute_length"
        Me.minute_length.Size = New System.Drawing.Size(76, 20)
        Me.minute_length.TabIndex = 19
        Me.minute_length.Text = "0"
        Me.minute_length.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(15, 87)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(75, 13)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "Minute Length"
        '
        'minute_index
        '
        Me.minute_index.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.minute_index.Location = New System.Drawing.Point(97, 58)
        Me.minute_index.MaxLength = 150
        Me.minute_index.Name = "minute_index"
        Me.minute_index.Size = New System.Drawing.Size(76, 20)
        Me.minute_index.TabIndex = 17
        Me.minute_index.Text = "0"
        Me.minute_index.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(15, 61)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(68, 13)
        Me.Label10.TabIndex = 16
        Me.Label10.Text = "Minute Index"
        '
        'user_length
        '
        Me.user_length.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.user_length.Location = New System.Drawing.Point(97, 136)
        Me.user_length.MaxLength = 150
        Me.user_length.Name = "user_length"
        Me.user_length.Size = New System.Drawing.Size(76, 20)
        Me.user_length.TabIndex = 23
        Me.user_length.Text = "0"
        Me.user_length.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(15, 139)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(79, 13)
        Me.Label11.TabIndex = 22
        Me.Label11.Text = "User ID Length"
        '
        'user_index
        '
        Me.user_index.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.user_index.Location = New System.Drawing.Point(97, 110)
        Me.user_index.MaxLength = 150
        Me.user_index.Name = "user_index"
        Me.user_index.Size = New System.Drawing.Size(76, 20)
        Me.user_index.TabIndex = 21
        Me.user_index.Text = "0"
        Me.user_index.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(15, 113)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(72, 13)
        Me.Label12.TabIndex = 20
        Me.Label12.Text = "User ID Index"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.year_index)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.year_lenght)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.month_index)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.month_length)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.day_index)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.day_length)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.hour_index)
        Me.SplitContainer1.Panel2.Controls.Add(Me.user_length)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label11)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.user_index)
        Me.SplitContainer1.Panel2.Controls.Add(Me.hour_length)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label12)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label10)
        Me.SplitContainer1.Panel2.Controls.Add(Me.minute_length)
        Me.SplitContainer1.Panel2.Controls.Add(Me.minute_index)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer1.Size = New System.Drawing.Size(373, 159)
        Me.SplitContainer1.SplitterDistance = 193
        Me.SplitContainer1.TabIndex = 24
        Me.SplitContainer1.TabStop = False
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(316, 165)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(58, 21)
        Me.Button1.TabIndex = 25
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'remove_zero
        '
        Me.remove_zero.AutoSize = True
        Me.remove_zero.Location = New System.Drawing.Point(3, 168)
        Me.remove_zero.Name = "remove_zero"
        Me.remove_zero.Size = New System.Drawing.Size(164, 17)
        Me.remove_zero.TabIndex = 26
        Me.remove_zero.Text = "Remove Zero Before User ID"
        Me.remove_zero.UseVisualStyleBackColor = True
        '
        'FrmTrans_Attendance_Log_IndexSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(379, 190)
        Me.Controls.Add(Me.remove_zero)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(395, 229)
        Me.Name = "FrmTrans_Attendance_Log_IndexSettings"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents year_index As System.Windows.Forms.TextBox
    Friend WithEvents year_lenght As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents month_length As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents month_index As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents day_length As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents day_index As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents hour_length As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents hour_index As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents minute_length As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents minute_index As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents user_length As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents user_index As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents remove_zero As System.Windows.Forms.CheckBox
End Class
