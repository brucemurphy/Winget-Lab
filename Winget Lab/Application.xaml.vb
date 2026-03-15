Imports System.Threading.Tasks

Partial Public Class Application
    Inherits System.Windows.Application

    Protected Overrides Async Sub OnStartup(e As StartupEventArgs)
        MyBase.OnStartup(e)

        ShutdownMode = ShutdownMode.OnExplicitShutdown

        Dim splash = New SplashWindow()
        splash.Show()

        Await Task.Delay(TimeSpan.FromSeconds(1.5))

        splash.Close()

        Dim mainWindow = New MainWindow()
        mainWindow = mainWindow
        mainWindow.Show()

        ShutdownMode = ShutdownMode.OnMainWindowClose
    End Sub

End Class
