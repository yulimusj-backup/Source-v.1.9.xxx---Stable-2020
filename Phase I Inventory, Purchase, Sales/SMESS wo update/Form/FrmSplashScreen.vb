Option Explicit On
Option Strict Off

Public NotInheritable Class FrmSplashScreen
    Dim WithEvents T1 As New Timer

    'TODO: This form can easily be set as the splash screen for the application by going to the "Application" tab
    '  of the Project Designer ("Properties" under the "Project" menu).


    Private Sub FrmSplashScreen_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Set up the dialog text at runtime according to the application's assembly information.  

        'TODO: Customize the application's assembly information in the "Application" pane of the project 
        '  properties dialog (under the "Project" menu).

        'Application title
        'If My.Application.Info.Title <> "" Then
        '    ApplicationTitle.Text = My.Application.Info.Title
        'Else
        '    'If the application title is missing, use the application name, without the extension
        '    ApplicationTitle.Text = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        'End If

        'Format the version information using the text set into the Version control at design time as the
        '  formatting string.  This allows for effective localization if desired.
        '  Build and revision information could be included by using the following code and changing the 
        '  Version control's designtime text to "Version {0}.{1:00}.{2}.{3}" or something similar.  See
        '  String.Format() in Help for more information.
        '
        '    Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build, My.Application.Info.Version.Revision)

        'Version.Text = "SMESS - Ver. " +  My.Settings.db_version

        ''Copyright info
        'Copyright.Text = My.Application.Info.Copyright
        Me.LogoPictureBox.Anchor = AnchorStyles.Top
        Me.Width -= 40

        Me.BackColor = Core_Module.GetSMESSColor
        Me.T1.Interval = 4050
        Me.T1.Enabled = True
    End Sub

    Private Sub T1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles T1.Tick
        Try
            Me.Close()
            Me.T1.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
