Imports System.IO
Imports System.Threading
Imports System.Windows.Threading
Imports System.Security.Principal
Imports System.IO.Compression
Imports System.Net

Class MainWindow
    ''' <summary>
    ''' RCC Downloader
    ''' </summary>
    ''' <remarks></remarks>
    Friend WithEvents RCCDL As New WebClient()
    ''' <summary>
    ''' Has the Shown() event already been raised once?
    ''' </summary>
    ''' <remarks></remarks>
    Private Shown As Boolean = False
    ''' <summary>
    ''' Equals to Shown() in WinForm programming.
    ''' </summary>
    ''' <remarks></remarks>
    Private Event OnWindowShown()
    ''' <summary>
    ''' Equals to OnShown() in WinForm programming.
    ''' </summary>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub OnContentRendered(e As EventArgs)
        MyBase.OnContentRendered(e)
        If Shown Then
            Exit Sub
        Else
            RaiseEvent OnWindowShown()
            Shown = True
        End If
    End Sub
    ''' <summary>
    ''' Wait a minute... emmm...
    ''' via Stack Overflow.
    ''' </summary>
    ''' <param name="seconds"></param>
    ''' <remarks></remarks>
    Private Sub Wait(ByVal Seconds As Double)
        Dim Frame = New DispatcherFrame()
        Dim Thr As New Thread(CType((Sub()
                                         Thread.Sleep(TimeSpan.FromSeconds(Seconds))
                                         Frame.[Continue] = False
                                     End Sub), ThreadStart))
        Thr.Start()
        Dispatcher.PushFrame(Frame)
    End Sub
    Public Shared AppPath As String = (New FileInfo(Reflection.Assembly.GetExecutingAssembly.Location)).Directory.FullName

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        TabItem_Revoking.Visibility = Windows.Visibility.Hidden
        TabItem_About.Visibility = Windows.Visibility.Hidden
        Grid_Status.Visibility = Windows.Visibility.Hidden
        ProgressRing_Revoke.Visibility = Windows.Visibility.Hidden
        ProgressBar_Revoke.Visibility = Windows.Visibility.Hidden
        ProgressBar_Main.Visibility = Windows.Visibility.Hidden
        Button_DownloadRCC.Visibility = Windows.Visibility.Hidden
        TextBlock_Status_Revoke.Visibility = Windows.Visibility.Hidden
    End Sub

    Private Sub MainWindow_OnWindowShown() Handles Me.OnWindowShown
        Try
            'Detect permissions
            TextBlock_Status.Text = "Detecting privileges..."
            If Not New WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator) Then
                TextBlock_Status.Text = "Acquiring Administrative privileges..."
                Dim proc As New Process
                With proc.StartInfo
                    .UseShellExecute = True
                    .Verb = "runas"
                    .FileName = Reflection.Assembly.GetExecutingAssembly.Location
                    .Arguments = Process.GetCurrentProcess.StartInfo.Arguments
                End With
                proc.Start()
                Application.Current.Shutdown(0)
            End If
            TextBlock_Status.Text = "Detecting RevokeChinaCerts Repo..."
            If Directory.Exists(Path.Combine(AppPath, "RevokeChinaCerts")) Then
                TextBlock_RCCRepo.Text = "Yes"
            Else
                TextBlock_RCCRepo.Text = "No"
                GoTo NoRccRepo
            End If
            TextBlock_Status.Text = "Counting Certificates..."
            'Directories:
            'RevokeChinaCerts/Windows/Certificates/CodeSigning
            'RevokeChinaCerts/Windows/Certificates/Organization
            'Other
            Dim Quantity As Integer = 0
            Quantity += (New DirectoryInfo(Path.Combine(AppPath, "RevokeChinaCerts", "Windows", "Certificates", "CodeSigning"))).GetFiles().Count
            Quantity += (New DirectoryInfo(Path.Combine(AppPath, "RevokeChinaCerts", "Windows", "Certificates", "Organization"))).GetFiles().Count
            TextBlock_RCC_Certificates.Text = Quantity - 2

            'NO RCC REPO
NoRccRepo:
            'Post-start options
            ProgressRing_Default.Visibility = Windows.Visibility.Hidden
            TextBlock_Status.Visibility = Windows.Visibility.Hidden
            Grid_Status.Visibility = Windows.Visibility.Visible
            TabItem_Revoking.Visibility = Windows.Visibility.Visible
            TabItem_About.Visibility = Windows.Visibility.Visible
            Button_DownloadRCC.Visibility = Windows.Visibility.Visible
        Catch ex As Exception
            TextBlock_Status.Text = "Error while loading: " & ex.Message
        End Try


    End Sub

    Private Sub Pre_Revoke()
        ToggleSwitch_CodeSigning.IsEnabled = False
        ToggleSwitch_Organization.IsEnabled = False
        ToggleSwitch_Other.IsEnabled = False
        Button_Certmgr.IsEnabled = False
        Button_Revoke.IsEnabled = False
        Button_Update.IsEnabled = False
        ProgressRing_Revoke.Visibility = Windows.Visibility.Visible
        ProgressBar_Revoke.Visibility = Windows.Visibility.Visible
        TextBlock_Status_Revoke.Visibility = Windows.Visibility.Visible
    End Sub
    Private Sub Post_Revoke()
        ToggleSwitch_CodeSigning.IsEnabled = True
        ToggleSwitch_Organization.IsEnabled = True
ToggleSwitch_Other.IsEnabled = False
        Button_Certmgr.IsEnabled = True
        Button_Revoke.IsEnabled = True
        Button_Update.IsEnabled = True
        ProgressRing_Revoke.Visibility = Windows.Visibility.Hidden
        ProgressBar_Revoke.Visibility = Windows.Visibility.Hidden
        TextBlock_Status_Revoke.Visibility = Windows.Visibility.Hidden
    End Sub

    Private Sub Button_Revoke_Click(sender As Object, e As RoutedEventArgs) Handles Button_Revoke.Click
        Pre_Revoke()
        TextBlock_Status_Revoke.Text = "Prepared."
        Try
            If ToggleSwitch_CodeSigning.IsChecked Then
                Dim Index As Integer = 0
                TextBlock_Status_Revoke.Text = "Revoking CodeSigning..."
                ProgressBar_Revoke.Maximum = (New DirectoryInfo(Path.Combine(AppPath, "RevokeChinaCerts", "Windows", "Certificates", "CodeSigning"))).GetFiles().Count
                For Each Cert As FileInfo In (New DirectoryInfo(Path.Combine(AppPath, "RevokeChinaCerts", "Windows", "Certificates", "CodeSigning"))).GetFiles()
                    Dim CertMgr As New Process()
                    With CertMgr.StartInfo
                        .FileName = Path.Combine(AppPath, "RevokeChinaCerts", "Windows", "Tools", "CertMgr.exe")
                        .Arguments = "-add -c " & Cert.FullName & " -s -r localMachine Disallowed"
                        .CreateNoWindow = True
                        .UseShellExecute = False
                        .WindowStyle = ProcessWindowStyle.Hidden
                    End With
                    CertMgr.Start()
                    While (CertMgr.HasExited = False)
                        Wait(0.1)
                    End While
                    Index += 1
                    ProgressBar_Revoke.Value = Index
                Next
            End If
            If ToggleSwitch_Organization.IsChecked Then
                Dim Index As Integer = 0
                TextBlock_Status_Revoke.Text = "Revoking Organization..."
                ProgressBar_Revoke.Maximum = (New DirectoryInfo(Path.Combine(AppPath, "RevokeChinaCerts", "Windows", "Certificates", "Organization"))).GetFiles().Count
                For Each Cert As FileInfo In (New DirectoryInfo(Path.Combine(AppPath, "RevokeChinaCerts", "Windows", "Certificates", "Organization"))).GetFiles()
                    Dim CertMgr As New Process()
                    With CertMgr.StartInfo
                        .FileName = Path.Combine(AppPath, "RevokeChinaCerts", "Windows", "Tools", "CertMgr.exe")
                        .Arguments = "-add -c " & Cert.FullName & " -s -r localMachine Disallowed"
                        .CreateNoWindow = True
                        .UseShellExecute = False
                        .WindowStyle = ProcessWindowStyle.Hidden
                    End With
                    CertMgr.Start()
                    While (CertMgr.HasExited = False)
                        Wait(0.1)
                    End While
                    Index += 1
                    ProgressBar_Revoke.Value = Index
                Next
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.ToString, "RevokeChinaCertsGUI", MessageBoxButton.OK, MessageBoxImage.Error)
        Finally
            Post_Revoke()
        End Try

    End Sub

    Private Sub Button_Certmgr_Click(sender As Object, e As RoutedEventArgs) Handles Button_Certmgr.Click
        Try
            Dim proc As New Process()
            With proc.StartInfo
                .FileName = "certmgr.msc"
                .UseShellExecute = True
            End With
            proc.Start()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button_Update_Click(sender As Object, e As RoutedEventArgs) Handles Button_Update.Click
        Pre_Revoke()
        Try
            TextBlock_Status_Revoke.Text = "Preparing..."
            Dim SWUObj As DirectoryInfo = Directory.CreateDirectory(Path.Combine(AppPath, "RevokeChinaCerts", "Windows", "Certificates", "Other", "SyncWithWU"))
            Dim GWUObj As DirectoryInfo = Directory.CreateDirectory(Path.Combine(AppPath, "RevokeChinaCerts", "Windows", "Certificates", "Other", "GenerateSSTFromWU"))

            TextBlock_Status_Revoke.Text = "Syncing certificates with Windows Update..."
            Dim SWUProc As New Process()
            With SWUProc.StartInfo
                .FileName = "certutil.exe"
                .Arguments = "-syncWithWU " & Convert.ToChar(34) & SWUObj.FullName & Convert.ToChar(34)
                .CreateNoWindow = True
                .UseShellExecute = False
                .WindowStyle = ProcessWindowStyle.Hidden
            End With
            SWUProc.Start()
            While (SWUProc.HasExited = False)
                Wait(0.1)
            End While

            TextBlock_Status_Revoke.Text = "Generating SST from Windows Update..."
            Dim GWUProc As New Process()
            With GWUProc.StartInfo
                .FileName = "certutil.exe"
                .Arguments = "-generateSSTFromWU " & Convert.ToChar(34) & Path.Combine(GWUObj.FullName, "AuthRoot.sst") & Convert.ToChar(34)
                .CreateNoWindow = True
                .UseShellExecute = False
                .WindowStyle = ProcessWindowStyle.Hidden
            End With
            GWUProc.Start()
            While (GWUProc.HasExited = False)
                Wait(0.1)
            End While

            TextBlock_Status_Revoke.Text = "Preparing..."
            Dim AutoRootSST As New FileInfo(Path.Combine(GWUObj.FullName, "AuthRoot.sst"))
            Dim PinRulesSST As String = Path.Combine(SWUObj.FullName, "pinrules.sst")

            TextBlock_Status_Revoke.Text = "Applying AutoRoot..."
            Dim ARSSTProc As New Process()
            With ARSSTProc.StartInfo
                .FileName = Path.Combine(AppPath, "RevokeChinaCerts", "Windows", "Tools", "CertMgr.exe")
                .Arguments = "-add -all -c " & Convert.ToChar(34) & AutoRootSST.FullName & Convert.ToChar(34) & " -s -r localMachine AuthRoot"
                .CreateNoWindow = True
                .UseShellExecute = False
                .WindowStyle = ProcessWindowStyle.Hidden
            End With
            ARSSTProc.Start()
            While (ARSSTProc.HasExited = False)
                Wait(0.1)
            End While

            TextBlock_Status_Revoke.Text = "Applying PinRules..."
            Dim PRProc As New Process()
            With PRProc.StartInfo
                .FileName = Path.Combine(AppPath, "RevokeChinaCerts", "Windows", "Tools", "CertMgr.exe")
                .Arguments = "-add -all -c " & Convert.ToChar(34) & PinRulesSST & Convert.ToChar(34) & " -s -r localMachine AuthRoot"
                .CreateNoWindow = True
                .UseShellExecute = False
                .WindowStyle = ProcessWindowStyle.Hidden
            End With
            PRProc.Start()
            While (PRProc.HasExited = False)
                Wait(0.1)
            End While

            TextBlock_Status_Revoke.Text = "Cleaning up..."
            Try
                My.Computer.FileSystem.DeleteDirectory(SWUObj.FullName, FileIO.DeleteDirectoryOption.DeleteAllContents)
                My.Computer.FileSystem.DeleteDirectory(GWUObj.FullName, FileIO.DeleteDirectoryOption.DeleteAllContents)
            Catch ex As Exception
            End Try
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.ToString, "RevokeChinaCertsGUI", MessageBoxButton.OK, MessageBoxImage.Error)
        Finally
            Post_Revoke()
        End Try
    End Sub

    Private Sub Button_DownloadRCC_Click(sender As Object, e As RoutedEventArgs) Handles Button_DownloadRCC.Click
        'Pre-disable
        TabItem_Revoking.Visibility = Windows.Visibility.Hidden
        Button_DownloadRCC.IsEnabled = False
        TextBlock_Status.Visibility = Windows.Visibility.Visible
        ProgressRing_Default.Visibility = Windows.Visibility.Visible
        ProgressBar_Main.Visibility = Windows.Visibility.Visible
        Try
            TextBlock_Status.Text = "Removing old RevokeChinaCerts Repository..."
            Try
                My.Computer.FileSystem.DeleteDirectory(Path.Combine(AppPath, "RevokeChinaCerts"), FileIO.DeleteDirectoryOption.DeleteAllContents)
            Catch ex As Exception
            End Try
            TextBlock_Status.Text = "Downloading RevokeChinaCerts..."
            RCCDL.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 10.0; WOW64) MSNET_WebClient/4.5 RevokeChinaCertGUI/1.0")
            RCCDL.DownloadFileAsync(New Uri("https://github.com/chengr28/RevokeChinaCerts/archive/master.zip"), Path.Combine(AppPath, "RCC.zip"))
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.ToString, "RevokeChinaCertsGUI", MessageBoxButton.OK, MessageBoxImage.Error)
            TabItem_Revoking.Visibility = Windows.Visibility.Visible
            Button_DownloadRCC.IsEnabled = True
            TextBlock_Status.Visibility = Windows.Visibility.Hidden
            ProgressRing_Default.Visibility = Windows.Visibility.Hidden
            ProgressBar_Main.Visibility = Windows.Visibility.Hidden
        End Try
    End Sub

    Private Sub RCCDL_DownloadFileCompleted(sender As Object, e As ComponentModel.AsyncCompletedEventArgs) Handles RCCDL.DownloadFileCompleted
        Try
            TextBlock_Status.Text = "Extracting..."
            Wait(0.5)
            ZipFile.ExtractToDirectory(Path.Combine(AppPath, "RCC.zip"), AppPath)
            TextBlock_Status.Text = "Finalizing..."
            Wait(0.5)
            Directory.Move(Path.Combine(AppPath, "RevokeChinaCerts-master"), Path.Combine(AppPath, "RevokeChinaCerts"))
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.ToString, "RevokeChinaCertsGUI", MessageBoxButton.OK, MessageBoxImage.Error)
        Finally
            Dim proc As New Process
            With proc.StartInfo
                .UseShellExecute = True
                .Verb = "runas"
                .FileName = Reflection.Assembly.GetExecutingAssembly.Location
                .Arguments = Process.GetCurrentProcess.StartInfo.Arguments
            End With
            proc.Start()
            Application.Current.Shutdown(0)
        End Try
    End Sub

    Private Sub RCCDL_DownloadProgressChanged(sender As Object, e As DownloadProgressChangedEventArgs) Handles RCCDL.DownloadProgressChanged
        TextBlock_Status.Text = "Downloading RevokeChinaCerts(" & Math.Round((e.BytesReceived / e.TotalBytesToReceive) * 100, 2) & "%)..."
        ProgressBar_Main.Maximum = CInt(e.TotalBytesToReceive / 10000)
        ProgressBar_Main.Value = CInt(e.BytesReceived / 10000)
    End Sub
End Class
