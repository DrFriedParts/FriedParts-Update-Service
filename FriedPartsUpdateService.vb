Public Class FriedPartsUpdateService
    Const fpUS_Version = "0.89"
    Const MyLog As String = "FriedParts"
    Const MySource As String = "fpUpdateService"
    Const TimerIntervalInSeconds = 10

    ' To access the constructor in Visual Basic, select New from the
    ' method name drop-down list. 
    Public Sub New()
        MyBase.New()
        InitializeComponent()

        'System.Diagnostics.EventLog.DeleteEventSource("UpdateService")
        If Not System.Diagnostics.EventLog.SourceExists(MySource) Then
            System.Diagnostics.EventLog.CreateEventSource(MySource,
            MyLog)
        End If
        EventLog1.Source = MySource
        EventLog1.Log = MyLog

        Timer1 = New System.Timers.Timer(1000 * TimerIntervalInSeconds)
        Timer1.AutoReset = True
        Timer1.Start()

    End Sub

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        'Timer1.AutoReset = True
        'Timer1.Interval = 1000 * TimerIntervalInSeconds 'Value must be in milliseconds
        'Timer1.Start()
        Msg("FP Update Service Version " & fpUS_Version & " STARTED!")
    End Sub


    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        'Should we call a notify function in the Web App? Hmmm...
        'Timer1.Stop() 'Not really necessary since Timer1 is about to get poofed from memory
        Msg("FP Update Service STOPPED!")
    End Sub

    Protected Sub Timer1_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles Timer1.Elapsed
        'Call the worker thread via the WebService -- we do this so that the web app handles 
        'this within its own local scope -- where is it easy to access all the functions and libs
        Msg("Tick!")
        'Call the WEB SERVICE!
    End Sub

    Protected Sub Msg(ByRef theMessageToLog As String)
        EventLog1.WriteEntry(theMessageToLog)
    End Sub
End Class
