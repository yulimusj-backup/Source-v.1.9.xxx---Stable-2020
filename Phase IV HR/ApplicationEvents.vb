Namespace My

    ' The following events are availble for MyApplication:
    ' 
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication
        'Private Sub MyApplication_Startup(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup
        '    Dim ij As New System.Globalization.DateTimeFormatInfo
        '    Dim jk As New System.Globalization.NumberFormatInfo
        '    My.Application.ChangeCulture("en-US")
        '    My.Application.ChangeUICulture("en-US")
        '    Me.ChangeCulture("en-US")
        '    Me.ChangeUICulture("en-US")
        '    ij.LongDatePattern = "dddd, MMM dd yyyy"
        '    ij.ShortDatePattern = "MMM dd yyyy"
        '    jk.CurrencyDecimalSeparator = "."
        '    jk.CurrencyGroupSeparator = ","
        '    Me.Culture.NumberFormat = jk
        '    Me.Culture.DateTimeFormat = ij
        'End Sub
    End Class

End Namespace

